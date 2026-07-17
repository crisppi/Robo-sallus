#!/usr/bin/env python3
"""Lanca uma evolucao clinica no Salus usando a API da aba logada.

Requer Chrome com DevTools remoto na porta 9222 e sessao Salus autenticada.
Por padrao salva apenas o rascunho; use --confirmar para enviar as respostas.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import Any

from etapa2_lancar_evolucao_salus import (
    read_clinical,
    read_queue,
    value_to_text,
    write_report,
    PatientResult,
)
from salus_cdp import DEFAULT_CDP, SALUS_ORIGIN, evaluate_js


USER_KEY = "49"


def normalize(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    return text.casefold()


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def salus_fetch(endpoint: str, method: str = "GET", body: Any | None = None, cdp_url: str = DEFAULT_CDP) -> Any:
    url = endpoint if endpoint.startswith("http") else f"{SALUS_ORIGIN}{endpoint}"
    js = f"""
    (async () => {{
      const options = {{
        method: {json.dumps(method)},
        credentials: 'include',
        headers: {{ 'Accept': 'application/json' }}
      }};
      if ({json.dumps(body is not None)}) {{
        options.headers['Content-Type'] = 'application/json';
        options.body = JSON.stringify({json.dumps(body, ensure_ascii=False)});
      }}
      const response = await fetch({json.dumps(url)}, options);
      const text = await response.text();
      let parsed;
      try {{ parsed = JSON.parse(text); }} catch (_) {{ parsed = text; }}
      return {{
        ok: response.ok,
        status: response.status,
        statusText: response.statusText,
        url: response.url,
        body: parsed
      }};
    }})()
    """
    return evaluate_js(js, cdp_url=cdp_url)


def split_multi(value: Any) -> list[Any]:
    if isinstance(value, (list, tuple)):
        return [item for item in value if not is_blank(item)]
    text = value_to_text(value)
    if ";" in text:
        return [part.strip() for part in text.split(";") if part.strip()]
    return [value]


def normalize_date(value: Any) -> str:
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, dt.date):
        return value.strftime("%Y-%m-%d")
    text = value_to_text(value)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(text[:19], fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return text


def option_for(question: dict[str, Any], value: Any) -> dict[str, Any] | None:
    text = value_to_text(value)
    options = question.get("opcoes", {}).get("itens") or []
    for option in options:
        candidates = [
            option.get("idOpcao"),
            option.get("codigo"),
            option.get("descricao"),
            str(option.get("descricao", "")).strip(),
        ]
        if any(normalize(candidate) == normalize(text) for candidate in candidates):
            answer = {"idOpcao": int(option["idOpcao"]), "texto": option["descricao"]}
            codigo = str(option.get("codigo") or "").strip()
            if codigo:
                answer["codigo"] = codigo
            return answer
    return None


def answer_value(question: dict[str, Any], raw_value: Any) -> list[dict[str, Any]]:
    if is_blank(raw_value):
        return []

    question_type = question.get("tipo", {})
    code = question_type.get("codigo")
    is_multi = bool(question_type.get("multiselecao")) or code == "LISTA_MULTIPLA"
    options = question.get("opcoes", {}).get("itens") or []

    value = normalize_date(raw_value) if code == "DATA" else raw_value
    parts = split_multi(value) if is_multi else [value]

    answers: list[dict[str, Any]] = []
    for part in parts:
        if is_blank(part):
            continue
        matched = option_for(question, part) if options else None
        if matched:
            answers.append(matched)
        else:
            answers.append({"idOpcao": 0, "texto": value_to_text(part)})
    return answers


def build_payload(definition: dict[str, Any], row_values: dict[str, Any], id_to_column: dict[int, str]) -> dict[str, Any]:
    sections = []
    for section in definition.get("secoes", []):
        if normalize(section.get("nome")) == "resumo":
            continue
        respostas = []
        for question in section.get("perguntas", []):
            question_id = int(question["idFormularioSecaoPergunta"])
            column = id_to_column.get(question_id)
            raw_value = row_values.get(column) if column else None
            respostas.append(
                {
                    "idFormularioSecaoPergunta": question_id,
                    "valorResposta": answer_value(question, raw_value),
                }
            )
        sections.append({"idSecao": section["idSecao"], "resposta": respostas})

    return {
        "idAuditor": int(USER_KEY),
        "idFormularioVersao": definition["idFormularioVersao"],
        "usuario": USER_KEY,
        "secoes": sections,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Salva/confirma evolucao clinica no Salus via API.")
    parser.add_argument("--fila", required=True)
    parser.add_argument("--clinica", required=True)
    parser.add_argument("--saida", required=True)
    parser.add_argument("--senha", required=True)
    parser.add_argument("--confirmar", action="store_true", help="Faz o POST final de respostas.")
    parser.add_argument("--cdp-url", default=DEFAULT_CDP)
    args = parser.parse_args()

    queue = {patient.senha: patient for patient in read_queue(Path(args.fila))}
    clinical, meta, _ = read_clinical(Path(args.clinica))
    if args.senha not in queue:
        raise RuntimeError(f"Senha {args.senha} nao encontrada na fila.")
    if args.senha not in clinical or len(clinical[args.senha]) != 1:
        raise RuntimeError(f"Senha {args.senha} nao encontrada de forma unica na planilha clinica.")

    patient = queue[args.senha]
    clinical_patient = clinical[args.senha][0]
    if not clinical_patient.id_internacao:
        raise RuntimeError(f"Paciente {args.senha} sem ID internacao na planilha clinica.")

    id_to_column = {
        int(field.pergunta_id): column
        for column, field in meta.items()
        if str(field.pergunta_id).strip().isdigit()
    }

    form_response = salus_fetch(
        f"/api/formularios/auditoria/definicao?user_key={USER_KEY}&idInternacao={clinical_patient.id_internacao}",
        cdp_url=args.cdp_url,
    )
    if not form_response.get("ok"):
        raise RuntimeError(f"Falha ao baixar definicao: {form_response}")
    payload = build_payload(form_response["body"], clinical_patient.values, id_to_column)

    for section in payload["secoes"]:
        draft_response = salus_fetch(
            f"/api/formularios/auditoria/internacoes/{clinical_patient.id_internacao}/rascunho?user_key={USER_KEY}",
            method="PUT",
            body={
                "idSecao": section["idSecao"],
                "idFormularioVersao": payload["idFormularioVersao"],
                "usuario": USER_KEY,
                "resposta": section["resposta"],
            },
            cdp_url=args.cdp_url,
        )
        print(f"Rascunho secao {section['idSecao']}:", json.dumps(draft_response, ensure_ascii=False))
        if not draft_response.get("ok"):
            raise RuntimeError(f"Falha ao salvar rascunho da secao {section['idSecao']}.")

    status = "RASCUNHO_SALVO"
    message = "Rascunho salvo no Salus; envio final nao confirmado."
    if args.confirmar:
        submit_response = salus_fetch(
            f"/api/formularios/auditoria/internacoes/{clinical_patient.id_internacao}/respostas?user_key={USER_KEY}",
            method="POST",
            body=payload,
            cdp_url=args.cdp_url,
        )
        print("Confirmacao:", json.dumps(submit_response, ensure_ascii=False))
        if not submit_response.get("ok"):
            status = "ERRO"
            message = f"Falha no POST final: {submit_response}"
            write_report(
                [
                    PatientResult(
                        senha=patient.senha,
                        nome=patient.nome,
                        iniciais=patient.iniciais,
                        status=status,
                        mensagem=message,
                        campos_preenchidos=[
                            column for column, value in clinical_patient.values.items() if not is_blank(value)
                        ],
                    )
                ],
                Path(args.saida),
            )
            raise RuntimeError(message)
        status = "SUCESSO"
        message = "Evolucao salva com sucesso no Salus."

    result = PatientResult(
        senha=patient.senha,
        nome=patient.nome,
        iniciais=patient.iniciais,
        status=status,
        mensagem=message,
        campos_preenchidos=[
            column for column, value in clinical_patient.values.items() if not is_blank(value)
        ],
    )
    write_report([result], Path(args.saida))
    print(message)
    print(f"Relatorio gerado: {args.saida}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        raise SystemExit(1)

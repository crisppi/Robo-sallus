#!/usr/bin/env python3
"""Sugestão conservadora de CID-10 a partir do diagnóstico da evolução.

A ordem é intencional: código escrito pelo profissional, diagnóstico principal
explícito e somente depois regras de patologia suficientemente específicas.
"""

from __future__ import annotations

import re
import unicodedata


def normalized(value: str) -> str:
    text = unicodedata.normalize("NFD", str(value or "").lower())
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", text).strip()


def infer_cid_suggestion(text: str) -> tuple[str, str]:
    """Retorna ``(cid, origem)``; vazio quando não há base clínica suficiente."""
    raw = str(text or "")
    explicit_cids = re.findall(
        r"(?i)\bcid\s*[:\-]?\s*([A-Z]\d{2}(?:\.\d{1,2})?)\b",
        raw,
    )
    if explicit_cids:
        return explicit_cids[-1].upper(), "CID explícito na evolução"

    value = normalized(raw)
    rules: tuple[tuple[tuple[str, ...], str, str], ...] = (
        (("isquemia mesenterica", "perfuracao de alca"), "K55.0", "isquemia mesentérica aguda"),
        (("choque septico refratario",), "R57.2", "choque séptico"),
        (("apendicectomia", "apendicite aguda"), "K35.9", "apendicite tratada por apendicectomia"),
        (("broncoaspir", "bcp aspirativa", "pneumonia aspirativa"), "J69.0", "pneumonia aspirativa"),
        (("pneumonia bacteriana", "broncopneumonia", "opacidades pulmonares"), "J18.9", "pneumonia não especificada"),
        (("virus sincicial respiratorio", "vsr positivos", "vsr positivo"), "J21.0", "bronquiolite por VSR"),
        (("sindrome coronariana aguda", " sca "), "I24.9", "síndrome coronariana aguda"),
        (("aterosclerose coronariana", "lesao coronariana", "doenca aterosclerotica"), "I25.1", "doença aterosclerótica do coração"),
        (("estenose aortica", "tavi"), "I35.0", "estenose aórtica"),
        (("ablacao de trn", "taquicardia por reentrada nodal"), "I47.1", "taquicardia supraventricular"),
        (("avc isquemico", "avci", "multiplos focos isquemicos", "infarto cerebral"), "I63.9", "infarto cerebral"),
        (("fratura trans trocanterica", "fratura transtroc"), "S72.1", "fratura pertrocantérica"),
        (("fratura do osso nasal", "fratura do septo nasal"), "S02.2", "fratura nasal"),
        (("sincope",), "R55", "síncope e colapso"),
        (("anemia ferropriva secundaria a menorragia",), "D50.0", "anemia ferropriva por perda sanguínea"),
        (("neutropenia febril",), "D70", "neutropenia"),
        (("mastocitose", "elevacao da triptase"), "D47.0", "neoplasia de comportamento incerto de mastócitos"),
        (("carcinoma espinocelular", "cec?"), "C44.7", "neoplasia maligna da pele do membro inferior"),
        (("adenocarcinoma de pulmao",), "C34.9", "neoplasia maligna de pulmão"),
        (("adenocarcinoma de cabeca de pancreas",), "C25.0", "neoplasia maligna da cabeça do pâncreas"),
        (("neoplasia colorretal", "ca de intestino"), "C19", "neoplasia colorretal"),
        (("colangite",), "K83.0", "colangite"),
        (("pancreatite cronica agudizada", "pancreatite aguda", "pancreatite leve"), "K85.9", "pancreatite aguda"),
        (("migracao de calculo", "coledocolitiase"), "K80.5", "cálculo de via biliar"),
        (("fistula liquorica", "hipotensao liquorica"), "G96.0", "fístula liquórica"),
        (("abstinencia alcoolica",), "F10.3", "abstinência alcoólica"),
        (("linfonodomegalias", "esplenomegalia sintomatica"), "R59.1", "linfonodomegalia generalizada"),
        (("dispneia", "dessaturacao"), "R06.0", "dispneia"),
        (("diarreia",), "R19.7", "diarreia"),
        (("nauseas", "vomitos"), "R11", "náuseas e vômitos"),
        (("sintomas gripais", "quadro gripal"), "J06.9", "infecção aguda de vias aéreas superiores"),
        (("broncopatia inflamatoria",), "J40", "bronquite não especificada"),
        (("edema em mie", "edema de membro inferior"), "R60.0", "edema localizado"),
        (("hematuria",), "R31", "hematúria"),
        (("hiperemia na regiao tibial", "teicoplanina"), "L03.1", "celulite de membro"),
        (("malformacao arteriovenosa", " mav "), "Q28.2", "malformação arteriovenosa"),
        (("alteracao de habito intestinal", "alteracao do habito intestinal"), "R19.4", "alteração do hábito intestinal"),
        (("abscesso dentario",), "K04.7", "abscesso dentário"),
        (("gastroenterocolite",), "A09", "gastroenterocolite"),
        (("cansaco", "fadiga", "mal estar inespecifico"), "R53", "mal-estar e fadiga"),
    )
    padded = f" {value} "
    for terms, cid, reason in rules:
        if any(term in padded for term in terms):
            return cid, reason
    return "", ""


def infer_cid_from_evolution(text: str) -> str:
    return infer_cid_suggestion(text)[0]


def infer_adjusted_cid_from_evolution(text: str) -> str:
    value = normalized(text)
    if any(term in value for term in ("area de isquemia recente", "avc isquemico", "avci", "infarto cerebral")):
        return "I63.9"
    return ""

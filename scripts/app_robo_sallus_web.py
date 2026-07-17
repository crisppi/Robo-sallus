#!/usr/bin/env python3
"""Tela web local do Robo Sallus.

Este app nao depende de tkinter. Ele abre uma pagina local no navegador com:
- botoes Etapa 1 e Etapa 2
- cards de contagem
- paciente/senha em processamento
- log de execucao
"""

from __future__ import annotations

import datetime as dt
import json
import subprocess
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from etapa2_lancar_evolucao_salus import (
    PatientResult,
    QueuePatient,
    process_patients,
    read_clinical,
    read_queue,
    read_successful_passwords,
    write_report,
)


ROOT = Path(__file__).resolve().parents[1]
EXPORTS = ROOT / "exports"
HOST = "127.0.0.1"
PORT = 8765


def newest(patterns: str | tuple[str, ...], fallback: str) -> Path:
    if isinstance(patterns, str):
        patterns = (patterns,)
    matches = []
    for pattern in patterns:
        matches.extend(EXPORTS.glob(pattern))
    matches = sorted(set(matches), key=lambda p: p.stat().st_mtime, reverse=True)
    return matches[0] if matches else EXPORTS / fallback


def default_state() -> dict:
    return {
        "running": False,
        "status": "Pronto",
        "current_patient": "Nenhum paciente em execucao",
        "current_password": "-",
        "processed_now": 0,
        "logs": [],
        "files": {
            "fila": str(newest(("fila_salus_*.xlsx", "pacientes_sirio_libanes_*.xlsx"), "fila_salus_DATA.xlsx")),
            "clinica": str(newest(("data_base_lancamento_*.xlsx", "data_base_lantamento_*.xlsx", "preenchimento_evolucao_clinica_*_colorido.xlsx"), "data_base_lancamento_DATA.xlsx")),
            "relatorio": str(EXPORTS / "data_base_lancamento_16_07_2026.xlsx"),
        },
        "cards": {
            "salus": "-",
            "excel": "-",
            "encontrados": "-",
            "faltam": "-",
        },
    }


STATE = default_state()
LOCK = threading.Lock()


def log(message: str) -> None:
    timestamp = dt.datetime.now().strftime("%H:%M:%S")
    with LOCK:
        STATE["logs"].append(f"[{timestamp}] {message}")
        STATE["logs"] = STATE["logs"][-300:]


def set_state(**kwargs) -> None:
    with LOCK:
        STATE.update(kwargs)


def calculate_cards() -> dict:
    files = STATE["files"]
    fila = Path(files["fila"])
    clinica = Path(files["clinica"])
    relatorio = Path(files["relatorio"])

    queue_patients = read_queue(fila) if fila.exists() else []
    clinical_by_password, _, _ = read_clinical(clinica) if clinica.exists() else ({}, {}, [])
    successful = read_successful_passwords(relatorio)

    clinical_unique = {senha for senha, rows in clinical_by_password.items() if len(rows) == 1}
    queue_passwords = [patient.senha for patient in queue_patients]
    found = sum(1 for senha in queue_passwords if senha in clinical_unique)
    missing_to_launch = sum(1 for senha in queue_passwords if senha not in successful)

    return {
        "salus": len(queue_patients),
        "excel": sum(len(rows) for rows in clinical_by_password.values()),
        "encontrados": found,
        "faltam": missing_to_launch,
    }


def refresh_cards() -> dict:
    cards = calculate_cards()
    with LOCK:
        STATE["cards"] = cards
        STATE["status"] = "Cards atualizados"
    return cards


def run_etapa1_worker() -> None:
    try:
        output = EXPORTS / f"pacientes_sirio_libanes_{dt.date.today().isoformat()}.xlsx"
        cmd = [sys.executable, str(ROOT / "scripts" / "gerar_lista_pacientes.py"), "--saida", str(output)]
        log("Etapa 1 iniciada: baixar fila Salus.")
        completed = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, check=False)
        if completed.stdout.strip():
            log(completed.stdout.strip())
        if completed.stderr.strip():
            log(completed.stderr.strip())
        if completed.returncode != 0:
            set_state(status=f"Etapa 1 falhou com codigo {completed.returncode}")
            log(f"ERRO: Etapa 1 falhou com codigo {completed.returncode}.")
        else:
            with LOCK:
                STATE["files"]["fila"] = str(output)
            refresh_cards()
            set_state(status="Etapa 1 concluida")
            log(f"Etapa 1 concluida. Arquivo: {output}")
    except Exception as exc:
        set_state(status="Erro na Etapa 1")
        log(f"ERRO: {exc}")
    finally:
        set_state(running=False)


def run_etapa2_worker(real_run: bool, only_password: str | None) -> None:
    try:
        files = STATE["files"].copy()
        fila = Path(files["fila"])
        clinica = Path(files["clinica"])
        relatorio = Path(files["relatorio"])
        mode = "REAL" if real_run else "DRY-RUN"
        log(f"Etapa 2 iniciada em modo {mode}.")

        queue_patients = read_queue(fila)
        clinical_by_password, field_meta, field_headers = read_clinical(clinica)
        successful_passwords = read_successful_passwords(relatorio)
        processed_count = 0

        def progress(event: str, patient: QueuePatient, result: PatientResult | None) -> None:
            nonlocal processed_count
            if event == "inicio":
                set_state(
                    current_patient=patient.nome or "-",
                    current_password=patient.senha,
                )
            elif event == "fim" and result:
                processed_count += 1
                set_state(processed_now=processed_count)
                log(f"{result.senha} - {result.nome} - {result.status}: {result.mensagem}")

        results = process_patients(
            queue_patients=queue_patients,
            clinical_by_password=clinical_by_password,
            field_meta=field_meta,
            field_headers=field_headers,
            successful_passwords=successful_passwords,
            dry_run=not real_run,
            only_password=only_password,
            progress_callback=progress,
        )
        write_report(results, relatorio)
        refresh_cards()
        set_state(status=f"Etapa 2 concluida. Relatorio: {relatorio}")
        log(f"Etapa 2 concluida. Relatorio: {relatorio}")
    except Exception as exc:
        set_state(status="Erro na Etapa 2")
        log(f"ERRO: {exc}")
    finally:
        set_state(running=False)


HTML = r"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Robo Sallus</title>
  <style>
    :root {
      --bg: #f4f6f8;
      --panel: #ffffff;
      --ink: #1f2937;
      --muted: #667085;
      --blue: #1f4e78;
      --line: #d8dee6;
      --danger: #9b1c1c;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      color: var(--ink);
      background: var(--bg);
    }
    main { max-width: 1220px; margin: 0 auto; padding: 24px; }
    h1 { margin: 0; font-size: 32px; letter-spacing: 0; }
    .sub { color: var(--muted); margin: 6px 0 20px; }
    .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
      margin-bottom: 16px;
    }
    .files {
      display: grid;
      grid-template-columns: 150px 1fr;
      gap: 10px 12px;
      align-items: center;
    }
    label { font-weight: 650; color: #374151; }
    input[type="text"] {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 10px 12px;
      font-size: 14px;
      background: #fff;
    }
    .actions { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
    button {
      border: 0;
      border-radius: 6px;
      padding: 11px 16px;
      font-size: 15px;
      font-weight: 750;
      cursor: pointer;
      color: #fff;
      background: var(--blue);
    }
    button.secondary { background: #4b5563; }
    button:disabled { opacity: .55; cursor: wait; }
    .check { color: var(--danger); font-weight: 750; display: flex; gap: 8px; align-items: center; }
    .cards { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; }
    .card {
      background: #fff;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px;
      min-height: 92px;
    }
    .card .title { color: var(--muted); font-size: 13px; }
    .card .value { color: var(--blue); font-weight: 800; font-size: 30px; margin-top: 8px; }
    .current {
      display: grid;
      grid-template-columns: 1fr 180px;
      gap: 16px;
      background: #eaf2f8;
      border: 1px solid #c7d9e8;
      border-radius: 8px;
      padding: 16px;
      margin-bottom: 16px;
    }
    .current .title { color: #506070; font-size: 13px; }
    .current .value { color: #17324d; font-size: 24px; font-weight: 800; margin-top: 4px; }
    .status { color: var(--muted); margin-bottom: 8px; }
    pre {
      margin: 0;
      min-height: 300px;
      max-height: 420px;
      overflow: auto;
      background: #111827;
      color: #e5e7eb;
      border-radius: 8px;
      padding: 14px;
      white-space: pre-wrap;
      font: 13px/1.45 Menlo, Consolas, monospace;
    }
    @media (max-width: 900px) {
      .cards { grid-template-columns: repeat(2, 1fr); }
      .current { grid-template-columns: 1fr; }
      .files { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <main>
    <h1>Robo Sallus</h1>
    <p class="sub">Etapa 1 baixa a fila. Etapa 2 confere senha por senha e prepara ou lança a evolução.</p>

    <section class="panel files">
      <label>Fila Salus</label>
      <input id="fila" type="text">
      <label>Planilha Clínica</label>
      <input id="clinica" type="text">
      <label>Relatório</label>
      <input id="relatorio" type="text">
    </section>

    <section class="panel actions">
      <button id="etapa1" onclick="startEtapa1()">Etapa 1: Baixar Fila Salus</button>
      <button id="etapa2" onclick="startEtapa2()">Etapa 2: Lançar Evolução</button>
      <button class="secondary" onclick="refreshCards()">Atualizar Cards</button>
      <label style="margin-left: 8px;">Somente senha</label>
      <input id="senha" type="text" style="width: 140px;">
      <label class="check"><input id="real" type="checkbox"> Executar no Salus real</label>
    </section>

    <section class="cards">
      <div class="card"><div class="title">Pacientes no Salus</div><div id="salus" class="value">-</div></div>
      <div class="card"><div class="title">Pacientes no Excel</div><div id="excel" class="value">-</div></div>
      <div class="card"><div class="title">Senhas encontradas</div><div id="encontrados" class="value">-</div></div>
      <div class="card"><div class="title">Faltam lançar</div><div id="faltam" class="value">-</div></div>
      <div class="card"><div class="title">Processados agora</div><div id="processed" class="value">0</div></div>
    </section>

    <section class="current">
      <div>
        <div class="title">Paciente atual</div>
        <div id="current_patient" class="value">Nenhum paciente em execução</div>
      </div>
      <div>
        <div class="title">Senha</div>
        <div id="current_password" class="value">-</div>
      </div>
    </section>

    <div id="status" class="status">Pronto</div>
    <pre id="logs"></pre>
  </main>

  <script>
    async function api(path, options = {}) {
      const response = await fetch(path, options);
      if (!response.ok) throw new Error(await response.text());
      return await response.json();
    }

    function readFiles() {
      return {
        fila: document.getElementById('fila').value,
        clinica: document.getElementById('clinica').value,
        relatorio: document.getElementById('relatorio').value
      };
    }

    async function saveFiles() {
      await api('/api/files', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(readFiles())
      });
    }

    async function refreshCards() {
      await saveFiles();
      await api('/api/refresh', {method: 'POST'});
      await poll();
    }

    async function startEtapa1() {
      await saveFiles();
      await api('/api/etapa1', {method: 'POST'});
      await poll();
    }

    async function startEtapa2() {
      await saveFiles();
      const real = document.getElementById('real').checked;
      if (real && !confirm('Confirma execução real no Salus?')) return;
      await api('/api/etapa2', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
          real_run: real,
          only_password: document.getElementById('senha').value.trim()
        })
      });
      await poll();
    }

    async function poll() {
      const state = await api('/api/status');
      document.getElementById('fila').value = state.files.fila;
      document.getElementById('clinica').value = state.files.clinica;
      document.getElementById('relatorio').value = state.files.relatorio;
      document.getElementById('salus').textContent = state.cards.salus;
      document.getElementById('excel').textContent = state.cards.excel;
      document.getElementById('encontrados').textContent = state.cards.encontrados;
      document.getElementById('faltam').textContent = state.cards.faltam;
      document.getElementById('processed').textContent = state.processed_now;
      document.getElementById('current_patient').textContent = state.current_patient;
      document.getElementById('current_password').textContent = state.current_password;
      document.getElementById('status').textContent = state.status;
      document.getElementById('logs').textContent = state.logs.join('\n');
      document.getElementById('etapa1').disabled = state.running;
      document.getElementById('etapa2').disabled = state.running;
    }

    setInterval(poll, 1200);
    poll();
  </script>
</body>
</html>
"""


class Handler(BaseHTTPRequestHandler):
    def _send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path == "/":
            body = HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if path == "/api/status":
            with LOCK:
                payload = json.loads(json.dumps(STATE, ensure_ascii=False))
            self._send_json(payload)
            return
        self._send_json({"error": "Nao encontrado"}, 404)

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        try:
            if path == "/api/files":
                data = self._read_json()
                with LOCK:
                    for key in ("fila", "clinica", "relatorio"):
                        if data.get(key):
                            STATE["files"][key] = data[key]
                self._send_json({"ok": True})
                return

            if path == "/api/refresh":
                cards = refresh_cards()
                self._send_json({"ok": True, "cards": cards})
                return

            if path == "/api/etapa1":
                with LOCK:
                    if STATE["running"]:
                        self._send_json({"error": "Ja existe uma etapa em execucao."}, 409)
                        return
                    STATE["running"] = True
                    STATE["processed_now"] = 0
                    STATE["status"] = "Executando Etapa 1"
                threading.Thread(target=run_etapa1_worker, daemon=True).start()
                self._send_json({"ok": True})
                return

            if path == "/api/etapa2":
                data = self._read_json()
                with LOCK:
                    if STATE["running"]:
                        self._send_json({"error": "Ja existe uma etapa em execucao."}, 409)
                        return
                    STATE["running"] = True
                    STATE["processed_now"] = 0
                    STATE["status"] = "Executando Etapa 2"
                    STATE["current_patient"] = "Iniciando..."
                    STATE["current_password"] = "-"
                threading.Thread(
                    target=run_etapa2_worker,
                    args=(bool(data.get("real_run")), data.get("only_password") or None),
                    daemon=True,
                ).start()
                self._send_json({"ok": True})
                return

            self._send_json({"error": "Nao encontrado"}, 404)
        except Exception as exc:
            self._send_json({"error": str(exc)}, 500)

    def log_message(self, format: str, *args) -> None:
        return


def main() -> int:
    EXPORTS.mkdir(parents=True, exist_ok=True)
    try:
        refresh_cards()
    except Exception as exc:
        log(f"Aviso ao carregar cards: {exc}")

    server = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}"
    print(f"Robo Sallus aberto em {url}")
    print("Mantenha esta janela aberta enquanto usa a tela.")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nEncerrando Robo Sallus.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

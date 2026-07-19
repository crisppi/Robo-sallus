#!/usr/bin/env python3
r"""Utilitarios para chamar APIs do Salus usando um Chrome ja logado.

Requisito: Chrome aberto com DevTools remoto na porta 9222.
Exemplo de abertura:
  /Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome \
    --remote-debugging-port=9222 --user-data-dir=/private/tmp/chrome-salus-profile
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import websocket


DEFAULT_CDP = "http://127.0.0.1:9222"
SALUS_ORIGIN = "https://salus.orizon.com.br"
SALUS_LOGIN = "https://www.orizonbrasil.com.br/acesso-restrito.html"


class SalusCdpError(RuntimeError):
    pass


@dataclass
class BrowserTab:
    title: str
    url: str
    websocket_debugger_url: str


def _get_json(url: str) -> Any:
    try:
        with urllib.request.urlopen(url, timeout=10) as response:
            return json.loads(response.read().decode("utf-8"))
    except (OSError, urllib.error.URLError, json.JSONDecodeError) as exc:
        raise SalusCdpError(
            "O Chrome do Salus nao esta disponivel. Feche-o e abra novamente "
            "pelo RoboSallus.command; depois faca login no Salus."
        ) from exc


def start_salus_chrome(cdp_url: str = DEFAULT_CDP, wait_seconds: float = 12.0) -> bool:
    """Garante um Chrome separado, com DevTools remoto, pronto para o Salus."""
    try:
        _get_json(f"{cdp_url.rstrip('/')}/json/version")
        return False
    except SalusCdpError:
        pass

    if sys.platform != "darwin":
        raise SalusCdpError(
            "Inicializacao automatica disponivel no macOS. Abra o Chrome com "
            "--remote-debugging-port=9222 e acesse o Salus."
        )

    candidates = (
        Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
        Path.home() / "Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    )
    chrome = next((path for path in candidates if path.exists()), None)
    if chrome is None:
        raise SalusCdpError("Google Chrome nao encontrado. Instale o Chrome e abra o Robo novamente.")

    profile = Path.home() / ".robo-sallus" / "chrome-profile"
    profile.mkdir(parents=True, exist_ok=True)
    with open(os.devnull, "wb") as devnull:
        subprocess.Popen(
            [
                str(chrome),
                "--remote-debugging-port=9222",
                f"--user-data-dir={profile}",
                SALUS_LOGIN,
            ],
            stdout=devnull,
            stderr=devnull,
            start_new_session=True,
        )

    deadline = time.time() + wait_seconds
    while time.time() < deadline:
        try:
            _get_json(f"{cdp_url.rstrip('/')}/json/version")
            return True
        except SalusCdpError:
            time.sleep(0.4)
    raise SalusCdpError("O Chrome foi aberto, mas a conexao com ele nao ficou pronta.")


def find_salus_tab(cdp_url: str = DEFAULT_CDP, url_contains: str | None = None) -> BrowserTab:
    tabs = _get_json(f"{cdp_url.rstrip('/')}/json")
    candidates = [
        tab
        for tab in tabs
        if "salus.orizon.com.br" in tab.get("url", "")
        and tab.get("webSocketDebuggerUrl")
    ]
    if url_contains:
        preferred = [tab for tab in candidates if url_contains in tab.get("url", "")]
        if preferred:
            candidates = preferred

    if not candidates:
        raise SalusCdpError(
            "Nao encontrei uma aba do Salus no Chrome remoto. "
            "Abra o Salus logado no Chrome com --remote-debugging-port=9222."
        )
    candidates.sort(
        key=lambda tab: (
            "Evolução clínica" in tab.get("title", "") or "Evolucao clinica" in tab.get("title", ""),
            "avaliacao-internacao" in tab.get("url", ""),
            "gestao-internacao" in tab.get("url", ""),
            tab.get("url", "").rstrip("/") != SALUS_ORIGIN,
        ),
        reverse=True,
    )
    tab = candidates[0]
    return BrowserTab(
        title=tab.get("title", ""),
        url=tab.get("url", ""),
        websocket_debugger_url=tab["webSocketDebuggerUrl"],
    )


def evaluate_js(expression: str, cdp_url: str = DEFAULT_CDP, url_contains: str | None = None) -> Any:
    tab = find_salus_tab(cdp_url, url_contains=url_contains)
    ws = websocket.create_connection(
        tab.websocket_debugger_url,
        timeout=120,
        suppress_origin=True,
    )
    try:
        message_id = int(time.time() * 1000) % 1_000_000
        payload = {
            "id": message_id,
            "method": "Runtime.evaluate",
            "params": {
                "expression": expression,
                "awaitPromise": True,
                "returnByValue": True,
                "timeout": 120000,
            },
        }
        ws.send(json.dumps(payload))
        while True:
            message = json.loads(ws.recv())
            if message.get("id") != message_id:
                continue
            if "error" in message:
                raise SalusCdpError(str(message["error"]))
            result = message.get("result", {}).get("result", {})
            if "exceptionDetails" in message.get("result", {}):
                raise SalusCdpError(str(message["result"]["exceptionDetails"]))
            if result.get("subtype") == "error":
                raise SalusCdpError(result.get("description", "Erro JavaScript"))
            return result.get("value")
    finally:
        ws.close()


def navigate_salus(url: str, cdp_url: str = DEFAULT_CDP, url_contains: str | None = None) -> None:
    """Navega a aba Salus escolhida via CDP Page.navigate."""
    tab = find_salus_tab(cdp_url, url_contains=url_contains)
    ws = websocket.create_connection(
        tab.websocket_debugger_url,
        timeout=30,
        suppress_origin=True,
    )
    try:
        message_id = int(time.time() * 1000) % 1_000_000
        ws.send(json.dumps({"id": message_id, "method": "Page.enable"}))
        ws.send(json.dumps({"id": message_id + 1, "method": "Page.navigate", "params": {"url": url}}))
        deadline = time.time() + 30
        while time.time() < deadline:
            message = json.loads(ws.recv())
            if message.get("id") == message_id + 1:
                if "error" in message:
                    raise SalusCdpError(str(message["error"]))
                return
    finally:
        ws.close()


def call_salus_api(endpoint: str, cdp_url: str = DEFAULT_CDP) -> Any:
    """Chama uma URL/endpoint do Salus com os cookies da aba logada."""
    if endpoint.startswith("http"):
        url = endpoint
    else:
        url = f"{SALUS_ORIGIN}{endpoint if endpoint.startswith('/') else '/' + endpoint}"

    js = f"""
    (async () => {{
      const response = await fetch({json.dumps(url)}, {{ credentials: 'include' }});
      const text = await response.text();
      let body;
      try {{ body = JSON.parse(text); }} catch (_) {{ body = text; }}
      return {{
        ok: response.ok,
        status: response.status,
        url: response.url,
        body
      }};
    }})()
    """
    result = evaluate_js(js, cdp_url=cdp_url)
    if not result or not result.get("ok"):
        raise SalusCdpError(f"Falha na API {url}: {result}")
    return result["body"]


def main() -> int:
    parser = argparse.ArgumentParser(description="Chama uma API do Salus usando a sessao do Chrome.")
    parser.add_argument("endpoint", help="Endpoint relativo ou URL completa do Salus")
    parser.add_argument("--cdp-url", default=DEFAULT_CDP)
    args = parser.parse_args()
    body = call_salus_api(args.endpoint, cdp_url=args.cdp_url)
    print(json.dumps(body, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SalusCdpError as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        raise SystemExit(1)

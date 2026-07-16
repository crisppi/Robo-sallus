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
import sys
import time
import urllib.request
from dataclasses import dataclass
from typing import Any

import websocket


DEFAULT_CDP = "http://127.0.0.1:9222"
SALUS_ORIGIN = "https://salus.orizon.com.br"


class SalusCdpError(RuntimeError):
    pass


@dataclass
class BrowserTab:
    title: str
    url: str
    websocket_debugger_url: str


def _get_json(url: str) -> Any:
    with urllib.request.urlopen(url, timeout=10) as response:
        return json.loads(response.read().decode("utf-8"))


def find_salus_tab(cdp_url: str = DEFAULT_CDP) -> BrowserTab:
    tabs = _get_json(f"{cdp_url.rstrip('/')}/json")
    candidates = [
        tab
        for tab in tabs
        if "salus.orizon.com.br" in tab.get("url", "")
        and tab.get("webSocketDebuggerUrl")
    ]
    if not candidates:
        raise SalusCdpError(
            "Nao encontrei uma aba do Salus no Chrome remoto. "
            "Abra o Salus logado no Chrome com --remote-debugging-port=9222."
        )
    tab = candidates[0]
    return BrowserTab(
        title=tab.get("title", ""),
        url=tab.get("url", ""),
        websocket_debugger_url=tab["webSocketDebuggerUrl"],
    )


def evaluate_js(expression: str, cdp_url: str = DEFAULT_CDP) -> Any:
    tab = find_salus_tab(cdp_url)
    ws = websocket.create_connection(
        tab.websocket_debugger_url,
        timeout=30,
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
                "timeout": 30000,
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

#!/bin/zsh
cd "$(dirname "$0")"

PYTHON_BIN="$(command -v python3)"
if [ -z "$PYTHON_BIN" ]; then
  echo "Python 3 nao encontrado. Instale o Python 3 e tente novamente."
  read -k 1 "?Pressione qualquer tecla para fechar."
  exit 1
fi

if [ ! -x .venv/bin/python ]; then
  echo "Primeiro uso: preparando o Robo Sallus..."
  "$PYTHON_BIN" -m venv .venv || exit 1
fi

if ! .venv/bin/python -c "import openpyxl, websocket" >/dev/null 2>&1; then
  echo "Instalando dependencias necessarias..."
  .venv/bin/python -m pip install -r requirements.txt || {
    echo "Nao foi possivel instalar as dependencias. Verifique a internet."
    read -k 1 "?Pressione qualquer tecla para fechar."
    exit 1
  }
fi

.venv/bin/python scripts/app_robo_sallus_web.py

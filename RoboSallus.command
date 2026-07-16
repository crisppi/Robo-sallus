#!/bin/zsh
cd "$(dirname "$0")"
if [ -x /opt/homebrew/bin/python3 ]; then
  /opt/homebrew/bin/python3 scripts/app_robo_sallus_web.py
else
  python3 scripts/app_robo_sallus_web.py
fi

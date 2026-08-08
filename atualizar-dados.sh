#!/usr/bin/env bash
# ============================================================
#  Mapa Zelar — atualiza o fallback js/data.js a partir da
#  planilha "Mapa - Zelar.xlsx" (na raiz do projeto).
#  Rodar:  ./atualizar-dados.sh
#  (se faltar dependencia, instala automaticamente)
# ============================================================
set -e
cd "$(dirname "$0")"

PLANILHA="Mapa - Zelar.xlsx"
if [ ! -f "$PLANILHA" ]; then
  echo "ERRO: nao encontrei '$PLANILHA' na raiz do projeto."
  exit 1
fi

if ! node -e "require('xlsx'); require('open-location-code')" 2>/dev/null; then
  echo "-> instalando dependencias (xlsx, open-location-code)..."
  npm install --no-save xlsx open-location-code
fi

node tools/build.mjs
echo "-> fallback js/data.js atualizado a partir de '$PLANILHA'."

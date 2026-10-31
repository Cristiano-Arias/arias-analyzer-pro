#!/usr/bin/env bash
# Cria um banco efêmero, aplica todas as migrações em ordem e roda os testes SQL.
# Uso: PGHOST=... PGUSER=... ./db/run-tests.sh   (requer psql e um superusuário)
set -euo pipefail
cd "$(dirname "$0")"

DB="garagem_test_$$"
cleanup() { dropdb --if-exists "$DB" >/dev/null 2>&1 || true; }
trap cleanup EXIT

createdb "$DB"
for f in migrations/*.sql; do
  echo "→ aplicando $f"
  psql -X -q -v ON_ERROR_STOP=1 -d "$DB" -f "$f"
done
for f in tests/*.sql; do
  echo "→ testando $f"
  psql -X -q -v ON_ERROR_STOP=1 -d "$DB" -f "$f"
done

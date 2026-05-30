#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$(cd "$SCRIPT_DIR/../.." && pwd)"
ENV_FILE="$SCRIPT_DIR/.env"
BACKUP_DIR="${TAKSKLAD_BACKUP_DIR:-/opt/taksklad/backups/postgres}"
RETENTION_DAYS="${TAKSKLAD_BACKUP_RETENTION_DAYS:-14}"

if [[ ! -f "$ENV_FILE" ]]; then
  echo "Missing env file: $ENV_FILE" >&2
  exit 1
fi

set -a
source "$ENV_FILE"
set +a

mkdir -p "$BACKUP_DIR"
chmod 700 "$BACKUP_DIR"

timestamp="$(date -u +%Y%m%dT%H%M%SZ)"
backup_file="$BACKUP_DIR/taksklad-postgres-$timestamp.sql.gz"

cd "$APP_DIR"
docker compose --env-file "$ENV_FILE" -f "$SCRIPT_DIR/docker-compose.yml" exec -T postgres \
  pg_dump -U "$POSTGRES_USER" "$POSTGRES_DB" | gzip -9 > "$backup_file"
chmod 600 "$backup_file"

find "$BACKUP_DIR" -type f -name 'taksklad-postgres-*.sql.gz' -mtime "+$RETENTION_DAYS" -delete

echo "$backup_file"

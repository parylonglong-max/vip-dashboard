#!/usr/bin/env bash
# 前端修改前强制备份：复制关键源文件并记录 Git 基线、工作区状态及 SHA-256。
# 用法：./scripts/backup_before_frontend_change.sh [修改目的]
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
STAMP="$(date +%Y%m%d-%H%M%S)"
BACKUP="backups/frontend/$STAMP"
PURPOSE="${1:-未填写修改目的}"
mkdir -p "$BACKUP"
git status --short > "$BACKUP/GIT_STATUS_BEFORE.txt"
git rev-parse HEAD > "$BACKUP/BASE_COMMIT.txt"
printf '%s\n' "$PURPOSE" > "$BACKUP/CHANGE_PURPOSE.txt"
FILES=(app.js styles.css index.html)
for file in "${FILES[@]}"; do
  [[ -f "$file" ]] && cp "$file" "$BACKUP/"
done
(
  cd "$BACKUP"
  sha256sum app.js styles.css index.html > SHA256SUMS.txt
)
printf 'BACKUP_CREATED=%s\nBASE_COMMIT=%s\n' "$BACKUP" "$(cat "$BACKUP/BASE_COMMIT.txt")"

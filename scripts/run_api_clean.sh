#!/usr/bin/env bash
set -euo pipefail

PORT=8000
RELOAD=0

while [[ $# -gt 0 ]]; do
  case "$1" in
    --port)
      PORT="$2"
      shift 2
      ;;
    --reload)
      RELOAD=1
      shift
      ;;
    *)
      echo "Unknown argument: $1" >&2
      exit 1
      ;;
  esac
done

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WORKSPACE_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$WORKSPACE_ROOT"

stop_uvicorn_processes() {
  local target_port="$1"

  if command -v lsof >/dev/null 2>&1; then
    mapfile -t pids < <(lsof -tiTCP:"$target_port" -sTCP:LISTEN 2>/dev/null || true)
    if ((${#pids[@]} > 0)); then
      kill -9 "${pids[@]}" 2>/dev/null || true
    fi
  elif command -v fuser >/dev/null 2>&1; then
    fuser -k "${target_port}/tcp" 2>/dev/null || true
  fi

  pkill -f "uvicorn app.main:app" 2>/dev/null || true
}

stop_uvicorn_processes "$PORT"
sleep 0.5

PYTHON_EXE="$WORKSPACE_ROOT/.venv/bin/python"
if [[ ! -x "$PYTHON_EXE" ]]; then
  echo "Python executable not found: $PYTHON_EXE" >&2
  echo "Create virtualenv first: python3.11 -m venv .venv" >&2
  exit 1
fi

UVICORN_ARGS=(
  -m
  uvicorn
  app.main:app
  --host
  127.0.0.1
  --port
  "$PORT"
)

if [[ "$RELOAD" -eq 1 ]]; then
  UVICORN_ARGS+=(--reload)
fi

exec "$PYTHON_EXE" "${UVICORN_ARGS[@]}"

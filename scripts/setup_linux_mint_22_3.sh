#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WORKSPACE_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$WORKSPACE_ROOT"

echo "[1/4] Updating apt cache"
sudo apt-get update

echo "[2/4] Installing system dependencies for Linux Mint 22.3"
sudo apt-get install -y \
  python3.11 \
  python3.11-venv \
  python3-pip \
  build-essential \
  libsqlite3-dev \
  git \
  curl \
  jq \
  tesseract-ocr \
  tesseract-ocr-rus \
  tesseract-ocr-eng \
  poppler-utils \
  imagemagick

echo "[3/4] Creating virtual environment"
if [[ ! -d .venv ]]; then
  python3.11 -m venv .venv
fi

source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel

echo "[4/4] Installing Python requirements"
pip install -r requirements.txt

echo "Done. Next commands:"
echo "  source .venv/bin/activate"
echo "  bash scripts/run_api_clean.sh --port 8000 --reload"

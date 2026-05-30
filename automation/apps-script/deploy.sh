#!/usr/bin/env bash
# Deploy the RHM Weekly Report Generator Apps Script.
# Usage: ./deploy.sh
set -euo pipefail

cd "$(dirname "$0")"

if [[ ! -f .clasp.json ]]; then
  echo "ERROR: .clasp.json not found. Run one-time setup first (see README.md)." >&2
  exit 1
fi

echo "Pushing to Apps Script project..."
clasp push --force

echo
echo "Done. Open the project with: clasp open"

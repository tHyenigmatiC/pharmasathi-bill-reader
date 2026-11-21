#!/usr/bin/env bash
# Reads a folder of bill photos into PharmaSathi import files. See README.md.
here="$(cd "$(dirname "$0")" && pwd)"
exec "$here/.venv/bin/python" -I "$here/bill_reader.py" "$@"

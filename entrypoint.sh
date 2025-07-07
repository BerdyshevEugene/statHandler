#!/bin/sh
set -e

uv run python main.py

exec "$@"
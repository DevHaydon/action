#!/usr/bin/env bash
set -euo pipefail

if ! command -v uv >/dev/null 2>&1; then
    echo "uv not found. Installing..."
    curl -Ls https://astral.sh/uv/install.sh | bash
fi

uv sync
uv tool install crewai
uv tool upgrade crewai

echo "Remember to create the .env file with your API keys as described in setup/SETUP-linux.md"

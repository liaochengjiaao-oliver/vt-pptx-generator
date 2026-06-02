#!/bin/bash
set -e

INPUT="$1"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

if [ -z "$INPUT" ]; then
    echo "Usage: ./generate.sh <content.md>"
    echo ""
    echo "Examples:"
    echo "  ./generate.sh weekly/2026-06-02-literature.md"
    echo "  ./generate.sh weekly/2026-06-02-progress.md"
    exit 1
fi

if [ ! -f "$INPUT" ]; then
    echo "Error: file not found - $INPUT"
    exit 1
fi

python3 -c "import pptx, yaml" 2>/dev/null || {
    echo "Installing dependencies..."
    pip3 install python-pptx pyyaml --quiet
}

python3 "$SCRIPT_DIR/generate_pptx.py" "$INPUT"

#!/bin/bash
set -e

TYPE="$1"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DATE=$(date +%Y-%m-%d)

if [ -z "$TYPE" ]; then
    echo "Usage: ./new.sh <literature|progress>"
    echo ""
    echo "Examples:"
    echo "  ./new.sh literature   # Create literature report"
    echo "  ./new.sh progress     # Create progress report"
    exit 1
fi

case "$TYPE" in
    literature|lit)
        TEMPLATE="$SCRIPT_DIR/templates/literature_report.md"
        OUTPUT="$SCRIPT_DIR/weekly/${DATE}-literature.md"
        ;;
    progress|prog)
        TEMPLATE="$SCRIPT_DIR/templates/progress_report.md"
        OUTPUT="$SCRIPT_DIR/weekly/${DATE}-progress.md"
        ;;
    *)
        echo "Error: type must be 'literature' or 'progress'"
        exit 1
        ;;
esac

if [ -f "$OUTPUT" ]; then
    echo "File exists: $OUTPUT"
    exit 0
fi

cp "$TEMPLATE" "$OUTPUT"
sed -i '' "s/YYYY-MM-DD/$DATE/g" "$OUTPUT"

echo "Created: $OUTPUT"
echo "Edit it, then run: ./generate.sh $OUTPUT"

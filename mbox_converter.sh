#!/bin/bash
# MBOX Converter - Unix/Linux/macOS Shell Runner
# This script ensures Python is available and runs the converter

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CONVERTER="$SCRIPT_DIR/mbox_converter.py"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Check for Python
get_python() {
    if command -v python3 &> /dev/null; then
        echo "python3"
    elif command -v python &> /dev/null; then
        echo "python"
    else
        echo ""
    fi
}

PYTHON=$(get_python)

if [ -z "$PYTHON" ]; then
    echo -e "${RED}Error: Python is not installed.${NC}"
    echo "Please install Python 3.8+ using your package manager:"
    echo "  Ubuntu/Debian: sudo apt install python3 python3-pip"
    echo "  macOS: brew install python3"
    echo "  Fedora: sudo dnf install python3"
    exit 1
fi

# Check Python version
PY_VERSION=$($PYTHON -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
PY_MAJOR=$($PYTHON -c 'import sys; print(sys.version_info.major)')
PY_MINOR=$($PYTHON -c 'import sys; print(sys.version_info.minor)')

if [ "$PY_MAJOR" -lt 3 ] || ([ "$PY_MAJOR" -eq 3 ] && [ "$PY_MINOR" -lt 8 ]); then
    echo -e "${RED}Error: Python 3.8+ required. Found: $PY_VERSION${NC}"
    exit 1
fi

# Check/install dependencies
check_deps() {
    $PYTHON -c "import tqdm" 2>/dev/null || {
        echo -e "${YELLOW}Installing dependencies...${NC}"
        $PYTHON -m pip install -r "$SCRIPT_DIR/requirements.txt" --quiet
    }
}

# Show help if no arguments
if [ $# -eq 0 ]; then
    echo -e "${GREEN}MBOX Converter v2.0${NC}"
    echo ""
    $PYTHON "$CONVERTER" --help
    exit 0
fi

# Install deps and run
check_deps
exec $PYTHON "$CONVERTER" "$@"

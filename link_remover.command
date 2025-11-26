#!/bin/bash

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Change to the script directory
cd "$SCRIPT_DIR"

# The Python script will handle venv creation and package installation automatically
# Just run it - it will set up everything needed
python3 link_remover.py

# Keep terminal open to see results
echo ""
echo "Press any key to close this window..."
read -n 1 -s


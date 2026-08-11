#!/usr/bin/env bg-sh
#!/bin/bash
# Rexinux's Tally Analyser Launcher Script

DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

echo "=================================================="
echo "🚀 Starting Rexinux's Tally Analyser..."
echo "=================================================="

python3 server.py 8000

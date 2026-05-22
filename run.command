#!/bin/bash
cd "$(dirname "$0")"
source venv/bin/activate 2>/dev/null
python3 metadata_cleaner.py &
disown
sleep 1
osascript -e 'tell application "Terminal" to close (first window whose name contains "run.command")' 2>/dev/null &

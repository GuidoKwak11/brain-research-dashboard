#!/bin/bash
# Double-click this file (or run it) to preview the dashboard locally.
# It starts a small web server and opens the page in your browser.
cd "$(dirname "$0")" || exit 1
PORT=8000
echo "Serving USP Brain Research Dashboard on http://localhost:$PORT"
echo "Press Ctrl+C to stop."
( sleep 1 && open "http://localhost:$PORT" ) &
python3 -m http.server "$PORT"

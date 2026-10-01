#!/usr/bin/env bash
# Keeps the report logger running: if the app stops (for example to load an update),
# it starts again. Press Ctrl+C to stop for good.
cd "$(dirname "$0")"
trap 'echo; echo "Stopped."; exit 0' INT
while true; do
  python3 app.py
  echo "App stopped -- restarting in 1 second (Ctrl+C to quit)..."
  sleep 1
done

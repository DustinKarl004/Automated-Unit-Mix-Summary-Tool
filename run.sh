#!/bin/bash
# Start the Unit Mix Summary Tool

cd "$(dirname "$0")"

# Create venv if missing
if [ ! -d ".venv" ]; then
  echo "Setting up environment..."
  python3 -m venv .venv
  .venv/bin/pip install -r requirements.txt -q
fi

echo "Starting Unit Mix Summary Tool at http://localhost:8888"
echo "Press Ctrl+C to stop."
.venv/bin/uvicorn app.main:app --host 0.0.0.0 --port 8888 --reload

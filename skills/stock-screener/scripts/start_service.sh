#!/bin/bash
# Auto-start script for stock screener service

SCREENER_ROOT="/Users/harryhuang/Algo Trading/knowledge-base/screener_v2"
PID_FILE="/tmp/screener_service.pid"
LOG_FILE="/tmp/screener_service.log"

# Check if service is already running
if [ -f "$PID_FILE" ]; then
    PID=$(cat "$PID_FILE")
    if ps -p "$PID" > /dev/null 2>&1; then
        echo "Service already running (PID: $PID)"
        exit 0
    fi
fi

# Start the service
cd "$SCREENER_ROOT"
source venv/bin/activate
nohup uvicorn main:app --host 0.0.0.0 --port 8000 > "$LOG_FILE" 2>&1 &
echo $! > "$PID_FILE"

echo "Screener service started (PID: $(cat $PID_FILE))"

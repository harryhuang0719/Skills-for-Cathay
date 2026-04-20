#!/bin/bash
# Stop script for stock screener service

PID_FILE="/tmp/screener_service.pid"

if [ -f "$PID_FILE" ]; then
    PID=$(cat "$PID_FILE")
    if ps -p "$PID" > /dev/null 2>&1; then
        kill "$PID"
        rm "$PID_FILE"
        echo "Service stopped"
    else
        rm "$PID_FILE"
        echo "Service not running"
    fi
else
    echo "Service not running"
fi

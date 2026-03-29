#!/bin/bash
# Start Web AI Tool as a persistent background process
# Usage: ./start.sh        — start in background
#        ./start.sh stop   — stop the server
#        ./start.sh status — check if running
#        ./start.sh log    — tail the log

DIR="$(cd "$(dirname "$0")" && pwd)"
PID_FILE="$DIR/.webai.pid"
LOG_FILE="$DIR/.webai.log"

case "${1:-start}" in
    start)
        if [ -f "$PID_FILE" ] && kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
            echo "Already running (PID $(cat "$PID_FILE")). Use './start.sh stop' first."
            exit 1
        fi
        echo "Starting Web AI Tool..."
        cd "$DIR"
        nohup python app.py >> "$LOG_FILE" 2>&1 &
        echo $! > "$PID_FILE"
        echo "Started (PID $!). Log: $LOG_FILE"
        echo "Open http://localhost:7860"
        ;;
    stop)
        if [ -f "$PID_FILE" ]; then
            PID=$(cat "$PID_FILE")
            if kill -0 "$PID" 2>/dev/null; then
                kill "$PID"
                rm -f "$PID_FILE"
                echo "Stopped (PID $PID)"
            else
                rm -f "$PID_FILE"
                echo "Process not running (stale PID file removed)"
            fi
        else
            echo "Not running"
        fi
        ;;
    status)
        if [ -f "$PID_FILE" ] && kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
            echo "Running (PID $(cat "$PID_FILE"))"
        else
            echo "Not running"
        fi
        ;;
    log)
        tail -f "$LOG_FILE"
        ;;
    *)
        echo "Usage: $0 {start|stop|status|log}"
        ;;
esac

#!/bin/bash
# ===========================================================================
# run_timesheet.sh — Wrapper for unattended scheduled runs
#
# Called by the macOS LaunchAgent (com.darren.timesheet-bot.plist).
# Cleans up stale browser locks, runs the fill+submit script, and
# writes timestamped output to logs/.
# ===========================================================================

set -euo pipefail

# ---- Resolve paths --------------------------------------------------------
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
VENV_PYTHON="${PROJECT_DIR}/.venv/bin/python"
LOG_DIR="${PROJECT_DIR}/logs"
PROFILE_DIR="${PROJECT_DIR}/browser_state/profile"

# ---- Create log directory -------------------------------------------------
mkdir -p "$LOG_DIR"

TIMESTAMP="$(date +%Y%m%d_%H%M%S)"
LOG_FILE="${LOG_DIR}/timesheet_${TIMESTAMP}.log"

# ---- Helper: log with timestamp -------------------------------------------
log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $*"
}

# ---- Start logging --------------------------------------------------------
{
    log "=========================================="
    log "SharePoint Timesheet Bot — Scheduled Run"
    log "=========================================="
    log "Project dir : $PROJECT_DIR"
    log "Python      : $VENV_PYTHON"
    log "Log file    : $LOG_FILE"
    log ""

    # ---- Pre-flight: kill stale Chromium & remove lock --------------------
    pkill -f "chromium" 2>/dev/null || true
    rm -f "${PROFILE_DIR}/SingletonLock" 2>/dev/null || true
    log "Pre-flight cleanup done"

    # ---- Verify venv exists -----------------------------------------------
    if [[ ! -x "$VENV_PYTHON" ]]; then
        log "ERROR: Python venv not found at $VENV_PYTHON"
        log "Run: python3 -m venv .venv && pip install -r requirements.txt"
        exit 1
    fi

    # ---- Run the fill script ----------------------------------------------
    log "Starting timesheet fill + submit..."
    log ""

    cd "$PROJECT_DIR"
    "$VENV_PYTHON" scripts/test_fill_timesheet.py --submit
    EXIT_CODE=$?

    log ""
    log "Script exited with code: $EXIT_CODE"

    # ---- Cleanup ----------------------------------------------------------
    pkill -f "chromium" 2>/dev/null || true
    rm -f "${PROFILE_DIR}/SingletonLock" 2>/dev/null || true

    log "Post-run cleanup done"
    log "=========================================="
    log "Finished at $(date '+%Y-%m-%d %H:%M:%S')"
    log "=========================================="

    exit $EXIT_CODE

} 2>&1 | tee "$LOG_FILE"

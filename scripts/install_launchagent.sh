#!/bin/bash
# ===========================================================================
# install_launchagent.sh — Install the macOS LaunchAgent for scheduled runs
# ===========================================================================

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
PLIST_NAME="com.darren.timesheet-bot.plist"
PLIST_SRC="${PROJECT_DIR}/${PLIST_NAME}"
PLIST_DST="$HOME/Library/LaunchAgents/${PLIST_NAME}"
LOG_DIR="${PROJECT_DIR}/logs"

echo "📦 Installing SharePoint Timesheet Bot LaunchAgent..."
echo ""

# ---- Ensure the plist exists ----------------------------------------------
if [[ ! -f "$PLIST_SRC" ]]; then
    echo "❌ Plist not found: $PLIST_SRC"
    exit 1
fi

# ---- Ensure wrapper script is executable ----------------------------------
chmod +x "${PROJECT_DIR}/scripts/run_timesheet.sh"

# ---- Ensure logs directory exists -----------------------------------------
mkdir -p "$LOG_DIR"

# ---- Unload existing agent if loaded -------------------------------------
if launchctl list | grep -q "$PLIST_NAME" 2>/dev/null; then
    echo "⏳ Unloading existing agent..."
    launchctl unload "$PLIST_DST" 2>/dev/null || true
fi

# ---- Remove old symlink/file if present -----------------------------------
rm -f "$PLIST_DST"

# ---- Create symlink -------------------------------------------------------
ln -sf "$PLIST_SRC" "$PLIST_DST"
echo "🔗 Symlinked: $PLIST_DST → $PLIST_SRC"

# ---- Load the agent -------------------------------------------------------
launchctl load "$PLIST_DST"
echo "✅ LaunchAgent loaded"

echo ""
echo "ℹ️  Schedule: Every Friday at 09:00"
echo "ℹ️  Logs:     $LOG_DIR/"
echo ""
echo "📋 Useful commands:"
echo "   Check status:   launchctl list | grep timesheet"
echo "   Unload:         launchctl unload $PLIST_DST"
echo "   View logs:      ls -lt $LOG_DIR/ | head"
echo ""
echo "Done! The bot will run automatically every Friday at 9 AM. 🎉"

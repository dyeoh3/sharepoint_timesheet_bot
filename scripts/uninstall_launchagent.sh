#!/bin/bash
# ===========================================================================
# uninstall_launchagent.sh — Remove the macOS LaunchAgent
# ===========================================================================

set -euo pipefail

PLIST_NAME="com.darren.timesheet-bot.plist"
PLIST_DST="$HOME/Library/LaunchAgents/${PLIST_NAME}"

echo "🗑️  Uninstalling SharePoint Timesheet Bot LaunchAgent..."
echo ""

# ---- Unload the agent -----------------------------------------------------
if launchctl list | grep -q "timesheet-bot" 2>/dev/null; then
    echo "⏳ Unloading agent..."
    launchctl unload "$PLIST_DST" 2>/dev/null || true
    echo "✅ Agent unloaded"
else
    echo "ℹ️  Agent was not loaded"
fi

# ---- Remove symlink/file --------------------------------------------------
if [[ -e "$PLIST_DST" ]]; then
    rm -f "$PLIST_DST"
    echo "✅ Removed: $PLIST_DST"
else
    echo "ℹ️  Plist not found at $PLIST_DST"
fi

echo ""
echo "Done! The scheduled bot has been removed. 🏁"

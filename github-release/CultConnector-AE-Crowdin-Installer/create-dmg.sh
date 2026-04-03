#!/bin/bash
# Create a DMG containing the .pkg (run after build.sh)
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"
hdiutil create -volname "Cult Connector (AE ↔ Crowdin)" -srcfolder "CultConnector-AE-Crowdin-Installer.pkg" -ov -format UDZO "CultConnector-AE-Crowdin-Installer.dmg"
echo "Created: $SCRIPT_DIR/CultConnector-AE-Crowdin-Installer.dmg"

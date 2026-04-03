#!/bin/bash
# Build the Cult Connector AE x Crowdin installer package (.pkg)
# Run this on macOS: chmod +x build.sh && ./build.sh

set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

PKG_ID="com.cultextensions.cult-connector-ae-crowdin"
PKG_VERSION="1.0"
OUTPUT_NAME="CultConnector-AE-Crowdin-Installer"
PAYLOAD_ROOT="$SCRIPT_DIR/payload"
SCRIPTS_DIR="$SCRIPT_DIR/scripts"
OUTPUT_PKG="$SCRIPT_DIR/${OUTPUT_NAME}.pkg"

# Ensure postinstall is executable
chmod 755 "$SCRIPTS_DIR/postinstall"

# Check payload exists
PAYLOAD_SCRIPT="Cult Connector (AE ↔ Crowdin).jsxbin"
if [ ! -f "$PAYLOAD_ROOT/tmp/CultConnectorInstall/$PAYLOAD_SCRIPT" ]; then
    echo "Error: Payload not found. Put '$PAYLOAD_SCRIPT' in payload/tmp/CultConnectorInstall/"
    exit 1
fi

echo "Building package..."
pkgbuild \
    --identifier "$PKG_ID" \
    --version "$PKG_VERSION" \
    --scripts "$SCRIPTS_DIR" \
    --root "$PAYLOAD_ROOT" \
    --install-location "/" \
    "$OUTPUT_PKG"

echo "Created: $OUTPUT_PKG"
echo ""
echo "To create a DMG with the installer:"
echo "  hdiutil create -volname \"Cult Connector (AE ↔ Crowdin)\" -srcfolder \"$OUTPUT_PKG\" -ov -format UDZO \"$SCRIPT_DIR/${OUTPUT_NAME}.dmg\""

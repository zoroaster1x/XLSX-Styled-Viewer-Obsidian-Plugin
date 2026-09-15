#!/usr/bin/env bash
# Installs the XLSX Styled Viewer into an Obsidian vault.
#
# Usage:
#   ./install.sh /path/to/YourVault
#   ./install.sh /path/to/YourVault --no-build
#
# The vault must already exist and contain a .obsidian folder.
# Open the vault once in Obsidian if it does not.

set -euo pipefail

VAULT="${1:-}"
NO_BUILD="${2:-}"

if [ -z "$VAULT" ]; then
	echo "Usage: $0 /path/to/YourVault [--no-build]"
	exit 1
fi

if [ ! -d "$VAULT/.obsidian" ]; then
	echo "Error: $VAULT does not look like an Obsidian vault (no .obsidian folder found)."
	echo "Open the folder once in Obsidian, then run this script again."
	exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

build_with() {
	if [ "$1" = "bun" ]; then
		bun install --silent
		bun run build
	else
		npm install --silent
		npm run build
	fi
}

if [ "$NO_BUILD" != "--no-build" ]; then
	if command -v bun >/dev/null 2>&1; then
		echo "Building the plugin with bun and esbuild..."
		build_with bun
	elif command -v npm >/dev/null 2>&1; then
		echo "Building the plugin with npm and esbuild..."
		build_with npm
	else
		echo "Neither bun nor npm was found."
		if [ ! -f main.js ]; then
			echo "Error: no build tool and no prebuilt main.js found."
			exit 1
		fi
		echo "Using the existing prebuilt main.js."
	fi
fi

DEST="$VAULT/.obsidian/plugins/xlsx-styled-viewer"
mkdir -p "$DEST"
cp main.js manifest.json styles.css "$DEST/"
if [ -f versions.json ]; then
	cp versions.json "$DEST/"
fi

echo
echo "Installed XLSX Styled Viewer to:"
echo "  $DEST"
echo
echo "Next steps in Obsidian:"
echo "  1. Settings -> Community plugins -> make sure Restricted mode is off."
echo "  2. Click the reload icon next to Installed plugins."
echo "  3. Enable 'XLSX Styled Viewer' and open any .xlsx file."

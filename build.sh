#!/bin/sh
# Packs the extension in src/ into dist/ics-extract-<version>.xpi
set -e
cd "$(dirname "$0")"
VERSION=$(sed -n 's/^ *"version": *"\([^"]*\)".*/\1/p' src/manifest.json)
mkdir -p dist
rm -f "dist/ics-extract-$VERSION.xpi"
(cd src && zip -r -X -q "../dist/ics-extract-$VERSION.xpi" . -x '.*')
echo "Created dist/ics-extract-$VERSION.xpi"

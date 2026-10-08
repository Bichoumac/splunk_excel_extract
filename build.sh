#!/usr/bin/env bash
#
# Packages the Splunk app from src/ into dist/splunk_excel_extract-<version>.tgz
#
#   ./build.sh
#
set -euo pipefail

APP=splunk_excel_extract
ROOT="$(cd "$(dirname "$0")" && pwd)"
VERSION="$(sed -n 's/^version *= *//p' "$ROOT/src/default/app.conf" | head -1)"
STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE"' EXIT

mkdir -p "$STAGE/$APP" "$ROOT/dist"
cp -R "$ROOT/src/." "$STAGE/$APP/"
cp "$ROOT/README.md" "$ROOT/LICENSE.txt" "$STAGE/$APP/"

# No local/ settings, compiled Python or hidden files in a package.
rm -rf "$STAGE/$APP/local" "$STAGE/$APP/metadata/local.meta"
find "$STAGE/$APP" \( -name '__pycache__' -o -name '*.pyc' -o -name '.*' \) -prune -exec rm -rf {} +
find "$STAGE/$APP" -type d -exec chmod 755 {} +
find "$STAGE/$APP" -type f -exec chmod 644 {} +

OUT="$ROOT/dist/$APP-$VERSION.tgz"
COPYFILE_DISABLE=1 tar --owner=0 --group=0 -C "$STAGE" -czf "$OUT" "$APP"
echo "Built $OUT"

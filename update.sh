#!/usr/bin/env bash
set -e
cd "$(dirname "${BASH_SOURCE[0]}")"
echo "pulling latest..."
git pull --ff-only
bash install.sh

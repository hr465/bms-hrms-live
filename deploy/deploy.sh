#!/usr/bin/env bash
# Pull the latest code and restart the portal. Run from anywhere:  bash ~/hrms/deploy/deploy.sh
set -euo pipefail
cd "$(dirname "$0")/.."
git pull --ff-only
npm ci --omit=dev
pm2 startOrReload deploy/ecosystem.config.js --update-env
pm2 save
echo "Deployed: $(git log --oneline -1)"

#!/usr/bin/env bash
# Activates HTTPS for a company's own web address on this server.
# Run after the company has added its DNS record (Type A, value = this server's IP):
#   sudo bash ~/hrms/deploy/add-domain.sh hr.company.com
set -euo pipefail
DOMAIN="${1:?usage: sudo bash add-domain.sh hr.company.com}"
[[ "$DOMAIN" =~ ^[a-z0-9.-]+\.[a-z]{2,}$ ]] || { echo "Not a valid domain: $DOMAIN"; exit 1; }
CONF="/etc/nginx/conf.d/hrms-$DOMAIN.conf"
cat > "$CONF" <<NGINX
server {
    listen 80;
    server_name $DOMAIN;
    client_max_body_size 20m;
    location / {
        proxy_pass http://127.0.0.1:3100;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
        proxy_read_timeout 120s;
    }
}
NGINX
nginx -t
systemctl reload nginx
certbot --nginx -d "$DOMAIN" --non-interactive --agree-tos -m "${CERT_EMAIL:-hello@mewareconnect.com}" --redirect
echo "Done. https://$DOMAIN now serves the portal."

#!/bin/sh\n\
# Install certbot if it's not already in the image\n\
apt-get update && apt-get install -y certbot\n\
# Stop any existing nginx process that might be using port 80\n\
pkill -9 nginx\n\
# Request the certificate using standalone mode (uses port 80)\n\
certbot certonly --standalone --non-interactive --agree-tos --email karasev.a@vezu.ru -d vpn.vezu.ru\n\
# Copy the certificate and key to a shared folder\n\
cp /etc/letsencrypt/live/vpn.vezu.ru/fullchain.pem /mnt/router/cert.pem\n\
cp /etc/letsencrypt/live/vpn.vezu.ru/privkey.pem /mnt/router/key.pem\n\
echo \"Certificates obtained and copied successfully!\"\n\
\""
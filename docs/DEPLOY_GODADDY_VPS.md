# Deploy to GoDaddy VPS (GitHub Actions)

This guide configures automatic deployment on every push to `main`.

## 1) Prepare VPS

SSH into VPS and create app directory:

```bash
mkdir -p /var/www/mom-app
```

Install Node.js 20+ and rsync if not already installed.

Create production `.env` on VPS (never commit this file):

```bash
nano /var/www/mom-app/.env
```

Example production values:

```env
NODE_ENV=production
PORT=3000
APP_HOST=127.0.0.1
APP_BASE_URL=https://your-domain.com

ZOHO_USE_MOCK=false
ZOHO_BASE_URL=https://projectsapi.zoho.in/restapi
ZOHO_ACCOUNTS_BASE_URL=https://accounts.zoho.in
ZOHO_PORTAL_ID=your_portal_id
ZOHO_CLIENT_ID=your_client_id
ZOHO_CLIENT_SECRET=your_client_secret
ZOHO_REFRESH_TOKEN=your_refresh_token
ORG_USER_EMAIL_DOMAIN=elegrow.com
```

## 2) Create systemd service (recommended)

Create `/etc/systemd/system/mom-app.service`:

```ini
[Unit]
Description=M.O.M App
After=network.target

[Service]
Type=simple
User=www-data
WorkingDirectory=/var/www/mom-app
EnvironmentFile=/var/www/mom-app/.env
ExecStart=/usr/bin/node /var/www/mom-app/src/server.js
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

Enable service:

```bash
sudo systemctl daemon-reload
sudo systemctl enable mom-app
sudo systemctl start mom-app
```

## 3) Add repository secrets in GitHub

In GitHub: `Settings -> Secrets and variables -> Actions -> New repository secret`

- `DEPLOY_SSH_PRIVATE_KEY` = private SSH key for deployment user
- `VPS_HOST` = VPS public IP / hostname
- `VPS_PORT` = SSH port (usually `22`)
- `VPS_USER` = SSH username
- `VPS_DEPLOY_PATH` = app path on VPS (example `/var/www/mom-app`)
- `VPS_SERVICE_NAME` = `mom-app` (if using systemd)
- Optional: `VPS_KNOWN_HOSTS` for strict host key pinning

## 4) Public key authorization on VPS

Add the corresponding **public** key to:

```bash
~/.ssh/authorized_keys
```

## 5) Trigger deployment

Push to `main`:

```bash
git push origin main
```

Workflow file:

- `.github/workflows/deploy-vps.yml`

## 6) DNS in GoDaddy

Point your domain/subdomain `A` record to the VPS IP.

If using reverse proxy (Nginx), route domain traffic to app port (example `3000`) and enable HTTPS (Let's Encrypt).

# PIH Proposal Tool - Production Deployment

This repo now includes a production Docker deployment for a Flask app using:
- `gunicorn` (app server, internal port `8000`)
- `caddy` (reverse proxy + automatic HTTPS on `80/443`)
- persistent Docker volumes for uploads and SQLite data

## Added deployment files
- `Dockerfile`
- `docker-compose.yml`
- `Caddyfile`
- `.env.example`

## Environment variables
Copy `.env.example` to `.env` and set real values:

```bash
cp .env.example .env
```

- `DOMAIN`: your public domain pointing to the VPS
- `SESSION_SECRET`: long random secret for Flask sessions
- `DATABASE_URL`: defaults to SQLite on a persistent volume (`sqlite:////data/pih_data.db`)

## Runtime architecture
- Client requests hit `caddy` on ports `80` and `443`
- Caddy reverse proxies to `app:8000`
- Flask/Gunicorn runs only on the internal Docker network (port `8000`)
- Persistent volumes:
  - uploads: mounted to `/app/static/uploads`
  - SQLite data: mounted to `/data`

## Health endpoint
A simple health endpoint is available at:
- `GET /healthz`

---

## Ubuntu 22.04 deployment guide

### 1. Install Docker Engine + Compose plugin

```bash
sudo apt-get update
sudo apt-get install -y ca-certificates curl gnupg

sudo install -m 0755 -d /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /etc/apt/keyrings/docker.gpg
sudo chmod a+r /etc/apt/keyrings/docker.gpg

echo \
  "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.gpg] https://download.docker.com/linux/ubuntu \
  $(. /etc/os-release && echo \"$VERSION_CODENAME\") stable" | \
  sudo tee /etc/apt/sources.list.d/docker.list > /dev/null

sudo apt-get update
sudo apt-get install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin

sudo systemctl enable docker
sudo systemctl start docker
```

Optional (avoid typing `sudo` for Docker commands):

```bash
sudo usermod -aG docker $USER
newgrp docker
```

### 2. Clone the repo

```bash
git clone <your-repo-url>
cd PIHProposalTool
```

### 3. Create `.env`

```bash
cp .env.example .env
```

Edit `.env`:

```env
DOMAIN=yourdomain.com
SESSION_SECRET=<long-random-secret>
DATABASE_URL=sqlite:////data/pih_data.db
```

### 4. Point your domain to the server
Create DNS records before starting:
- `A` record for `yourdomain.com` -> your VPS public IPv4
- optional `AAAA` if using IPv6

### 5. Start the stack

```bash
docker compose up -d --build
```

### 6. Verify deployment

```bash
docker compose ps
docker compose logs -f app
docker compose logs -f caddy
curl -I http://yourdomain.com/healthz
curl -I https://yourdomain.com/healthz
```

If DNS is correct and ports `80/443` are open, Caddy will automatically provision and renew TLS certificates.

## Updating later

```bash
git pull
docker compose up -d --build
```

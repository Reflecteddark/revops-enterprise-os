#!/usr/bin/env bash
# ========================================================
# AI-ROP MAX Messenger Bot - 1-Command Deployment on Linux VPS
# ========================================================
set -e

echo "[1/4] Checking environment and Python3..."
command -v python3 >/dev/null 2>&1 || {
    echo "Installing python3..."
    apt-get update -y && apt-get install -y python3 git
}

APP_DIR="/opt/revops-enterprise-os"
REPO_URL="https://github.com/Reflecteddark/revops-enterprise-os.git"

echo "[2/4] Syncing application code to $APP_DIR..."
if [ -d "$APP_DIR/.git" ]; then
    cd "$APP_DIR"
    git pull origin main
else
    mkdir -p "$APP_DIR"
    git clone "$REPO_URL" "$APP_DIR"
    cd "$APP_DIR"
fi

echo "[3/4] Creating systemd service (ai-rop-max-bot)..."
cat << 'EOF' > /etc/systemd/system/ai-rop-max-bot.service
[Unit]
Description=AI-ROP MAX Messenger Bot Service (@se14526668_bot)
After=network.target

[Service]
Type=simple
User=root
WorkingDirectory=/opt/revops-enterprise-os
ExecStart=/usr/bin/python3 /opt/revops-enterprise-os/services/max_bot.py
Restart=always
RestartSec=5
Environment=PYTHONUNBUFFERED=1

[Install]
WantedBy=multi-user.target
EOF

echo "[4/4] Enabling and starting service..."
systemctl daemon-reload
systemctl enable ai-rop-max-bot
systemctl restart ai-rop-max-bot

echo ""
echo "========================================================"
echo "AI-ROP MAX Bot successfully deployed and running 24/7!"
echo "Check status: systemctl status ai-rop-max-bot"
echo "Check logs:   journalctl -u ai-rop-max-bot -f"
echo "========================================================"

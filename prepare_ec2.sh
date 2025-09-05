#!/bin/bash

set -e

echo "Updating package list..."
sudo apt-get update

echo "Installing Python3, pip, venv, and git..."
sudo apt-get install -y python3 python3-pip python3.12-venv git

echo "Activating virtual environment and upgrading pip..."
python3 -m venv venv
source venv/bin/activate
python3 -m pip install --upgrade pip

echo "Installing Python dependencies from requirements.txt..."
python3 -m pip install -r requirements.txt

echo "Creating systemd service for BCA_Streamlit..."

SERVICE_FILE="/etc/systemd/system/bca_streamlit.service"
APP_DIR="/home/ubuntu/BCA_Streamlit"
VENV_PATH="$APP_DIR/venv/bin/streamlit"
APP_FILE="app.py"
USER="ubuntu"

sudo bash -c "cat > $SERVICE_FILE" <<EOL
[Unit]
Description=BCA Streamlit App
After=network.target

[Service]
User=$USER
WorkingDirectory=$APP_DIR
ExecStart=$VENV_PATH run $APP_FILE --server.port=8501 --server.headless true
Restart=always

[Install]
WantedBy=multi-user.target
EOL

echo "Reloading systemd daemon and enabling service..."
sudo systemctl daemon-reload
sudo systemctl enable bca_streamlit.service

echo "Systemd service for BCA_Streamlit prepared. You can start it with: sudo systemctl start bca_streamlit"
echo "EC2 preparation
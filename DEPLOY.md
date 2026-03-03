# Deploying Word-to-PDF on Azure Ubuntu VM

## 1. Install LibreOffice

```bash
sudo apt update
sudo apt install -y libreoffice-core libreoffice-writer
```

Verify it works:

```bash
libreoffice --headless --version
```

## 2. Install Node.js (v20 LTS)

```bash
curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
sudo apt install -y nodejs
```

Verify:

```bash
node -v
npm -v
```

## 3. Copy the project to the VM

Use `scp` or `rsync` to copy the `word-to-pdf/` folder to the VM:

```bash
scp -r ./word-to-pdf user@your-vm-ip:~/word-to-pdf
```

Then on the VM:

```bash
cd ~/word-to-pdf
npm install
```

## 4. Test it manually

```bash
node server.js
```

Open `http://your-vm-ip:3000` in a browser (make sure port 3000 is open — see step 6).

## 5. Run with PM2 (production)

Install PM2 globally:

```bash
sudo npm install -g pm2
```

Start the service:

```bash
cd ~/word-to-pdf
pm2 start server.js --name word-to-pdf
```

Enable auto-start on reboot:

```bash
pm2 startup
pm2 save
```

Useful PM2 commands:

```bash
pm2 status          # check status
pm2 logs word-to-pdf  # view logs
pm2 restart word-to-pdf  # restart
```

## 6. Open the port in Azure

In the Azure Portal:

1. Go to your VM → **Networking** → **Network Settings**
2. Click **Add inbound port rule**
3. Set: Source = Any, Destination port = 3000, Protocol = TCP, Action = Allow
4. Save

Now access the service at `http://your-vm-public-ip:3000`

## 7. (Optional) Nginx reverse proxy on port 80

If you want to serve on port 80 instead of 3000:

```bash
sudo apt install -y nginx
```

Create config:

```bash
sudo tee /etc/nginx/sites-available/word-to-pdf << 'EOF'
server {
    listen 80;
    server_name _;

    client_max_body_size 50M;

    location / {
        proxy_pass http://127.0.0.1:3000;
        proxy_http_version 1.1;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    }
}
EOF
```

Enable and restart:

```bash
sudo ln -sf /etc/nginx/sites-available/word-to-pdf /etc/nginx/sites-enabled/
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t
sudo systemctl restart nginx
```

Then open port 80 in Azure NSG (same steps as step 6, but port 80).

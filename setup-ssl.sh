#!/bin/bash

# 安装Certbot
apt-get update
apt-get install -y certbot python3-certbot-nginx

# 获取SSL证书
certbot --nginx -d tool.jarvismedical.asia --non-interactive --agree-tos --email your-email@example.com

# 设置自动续期
echo "0 0,12 * * * root python -c 'import random; import time; time.sleep(random.random() * 3600)' && certbot renew -q" | sudo tee -a /etc/crontab > /dev/null

# 重启Nginx
systemctl restart nginx

echo "SSL证书设置完成！"

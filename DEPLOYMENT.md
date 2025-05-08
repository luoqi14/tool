# Jarvis工具集部署指南 (OpenCloudOS Server 9)

本文档提供将Jarvis工具集部署到 https://tool.jarvismedical.asia 的详细步骤，适用于OpenCloudOS Server 9系统。当前工具集包含邮件发送模块，可通过 https://tool.jarvismedical.asia/email 访问。

## 准备工作

1. 一台运行OpenCloudOS Server 9的服务器
2. 已注册的域名 tool.jarvismedical.asia，并将DNS记录指向您的服务器IP
3. 确保80和443端口已开放（用于HTTP和HTTPS）

## 部署步骤

### 1. 安装必要的软件

```bash
# 更新系统
sudo dnf update -y

# 安装Docker和Docker Compose
sudo dnf config-manager --add-repo=https://download.docker.com/linux/centos/docker-ce.repo
sudo dnf install -y docker-ce docker-ce-cli containerd.io docker-compose-plugin
sudo systemctl enable --now docker

# 安装Nginx
sudo dnf install -y nginx
sudo systemctl enable --now nginx

# 安装EPEL仓库（用于Certbot）
sudo dnf install -y epel-release

# 安装Certbot
sudo dnf install -y certbot python3-certbot-nginx
```

### 2. 配置Nginx

将项目中的`nginx.conf`文件复制到Nginx配置目录：

```bash
sudo cp nginx.conf /etc/nginx/conf.d/tool.jarvismedical.asia.conf
sudo mv /etc/nginx/nginx.conf /etc/nginx/nginx.conf.backup  # 备份原始配置
sudo nginx -t  # 测试配置是否有效
sudo systemctl restart nginx
```

### 3. 获取SSL证书

运行Certbot获取SSL证书：

```bash
sudo certbot --nginx -d tool.jarvismedical.asia --non-interactive --agree-tos --email your-email@example.com
```

或者使用项目中的脚本：

```bash
sudo chmod +x setup-ssl.sh
sudo ./setup-ssl.sh
```

### 4. 使用Docker Compose部署应用

项目中包含了`docker-compose.yml`文件，可以使用Docker Compose进行部署：

```bash
# 构建并启动容器
docker compose up -d
```

或者使用项目中的更新脚本：

```bash
chmod +x update.sh
./update.sh
```

### 5. 验证部署

访问以下地址确认应用是否正常运行：

- 工具集主页：https://tool.jarvismedical.asia
- 邮件发送工具：https://tool.jarvismedical.asia/email

## 更新应用

当需要更新应用时，可以使用项目中的更新脚本：

```bash
# 拉取最新代码
git pull

# 运行更新脚本
./update.sh
```

更新脚本会自动执行以下操作：

1. 拉取最新代码
2. 停止并删除旧容器
3. 重新构建Docker镜像
4. 启动新容器

## 故障排除

### 1. 检查应用日志

```bash
# 查看容器日志
docker compose logs

# 查看特定容器的日志
docker logs email-app
```

### 2. 检查Nginx日志

```bash
sudo tail -f /var/log/nginx/access.log
sudo tail -f /var/log/nginx/error.log
```

### 3. 检查SSL证书

```bash
sudo certbot certificates
```

### 4. 重启服务

```bash
# 使用Docker Compose重启所有服务
docker compose restart

# 或者重启特定容器
docker restart email-app

# 重启Nginx
sudo systemctl restart nginx

# 如果遇到防火墙问题，可以使用以下命令开放端口
sudo firewall-cmd --permanent --add-service=http
sudo firewall-cmd --permanent --add-service=https
sudo firewall-cmd --permanent --add-port=3001/tcp
sudo firewall-cmd --reload
```

## 安全考虑

1. **环境变量与敏感信息**：将邮箱凭证和其他敏感信息存储在环境变量或外部配置文件中，而不是硬编码在代码中
2. **定期更新**：定期更新系统、Docker镜像和应用依赖，以修复安全漏洞
3. **网络安全**：
   - 配置防火墙，只开放必要的端口（80「443、和内部3001端口）
   - 启用HTTPS并定期更新SSL证书（Certbot会自动处理）
4. **访问控制**：考虑添加用户认证或IP限制，以限制对工具的访问
5. **日志和监控**：设置日志监控，以及时检测异常行为

## OpenCloudOS特有的注意事项

1. SELinux可能会阻止Nginx访问某些文件或端口，如果遇到权限问题，可以使用以下命令调整：
   ```bash
   sudo setsebool -P httpd_can_network_connect 1
   ```

2. 如果使用的是DNF模块流，可能需要使用以下命令启用模块：
   ```bash
   sudo dnf module enable nginx:1.20
   sudo dnf module enable python39
   ```

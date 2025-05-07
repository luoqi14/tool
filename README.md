# 邮件发送工具

这是一个基于Flask的邮件发送Web应用，提供用户友好的界面来发送邮件，支持HTML/纯文本格式和文件附件。

## 功能特点

- 美观的Bootstrap用户界面
- 支持多个收件人（用逗号分隔）
- 支持HTML和纯文本格式的邮件内容
- 支持添加多个文件附件
- 实时表单验证和用户反馈
- 使用企业微信邮箱SMTP服务发送邮件

## 项目结构

```
email/
├── app.py              # Flask后端API
├── index.html          # 前端HTML页面
├── static/             # 静态资源目录
│   └── js/             
│       └── main.js     # 前端JavaScript逻辑
└── requirements.txt    # 项目依赖
```

## 安装与运行

1. 安装依赖：

```bash
# 使用pip
pip install -r requirements.txt

# 或使用uv
uv pip install -r requirements.txt
```

2. 运行应用：

```bash
python app.py
```

3. 在浏览器中访问：

```
http://localhost:8080
```

## 部署指南

### 本地部署

应用默认在8080端口运行。如需更改端口，请修改`app.py`文件中的相应配置。

### 服务器部署

#### 使用Gunicorn（推荐用于生产环境）

1. 安装Gunicorn：

```bash
pip install gunicorn
```

2. 启动应用：

```bash
gunicorn -w 4 -b 0.0.0.0:8080 app:app
```

#### 使用Docker部署

1. 创建Dockerfile：

```dockerfile
FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install gunicorn

COPY . .

EXPOSE 8080

CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:8080", "app:app"]
```

2. 构建并运行Docker镜像：

```bash
docker build -t email-sender .
docker run -p 8080:8080 email-sender
```

## 安全注意事项

- 应用中包含邮箱授权码，生产环境中应使用环境变量或配置文件存储
- 建议在生产环境中启用HTTPS
- 考虑添加用户认证以限制访问

# Jarvis工具集

这是一个基于Flask的模块化工具集应用，当前包含邮件发送工具，并支持扩展更多工具模块。

## 功能特点

- 模块化架构，易于扩展新功能
- 美观的Bootstrap用户界面
- 当前包含邮件发送工具，支持：
  - 多个收件人（用逗号分隔）
  - HTML和纯文本格式的邮件内容
  - 多个文件附件
  - 实时表单验证和用户反馈
- 使用企业微信邮箱SMTP服务发送邮件

## 项目结构

```
/
├─ app.py                # 主应用入口
├─ static/                # 全局静态文件
├─ modules/               # 模块目录
│   └─ email/             # 邮件发送模块
│       ├─ __init__.py      # 模块初始化文件
│       ├─ routes.py        # 邮件模块路由和逻辑
│       ├─ static/          # 邮件模块静态文件
│       │   └─ js/
│       │       └─ main.js  # 前端JavaScript逻辑
│       └─ templates/       # 邮件模块模板
│           └─ index.html   # 邮件发送页面
├─ docker-compose.yml      # Docker Compose配置
├─ Dockerfile              # Docker构建文件
├─ nginx.conf              # Nginx配置
├─ update.sh               # 更新脚本
└─ requirements.txt        # 项目依赖
```

## 安装与运行

1. 安装依赖：

```bash
pip install -r requirements.txt
```

2. 运行应用：

```bash
python app.py
```

3. 在浏览器中访问：

```
# 工具列表
http://localhost:3000

# 邮件发送工具
http://localhost:3000/email
```

## 添加新工具模块

要添加新的工具模块，请按照以下步骤操作：

1. 在`modules`目录下创建新的模块目录，例如`modules/file_processor/`
2. 创建必要的文件：`__init__.py`、`routes.py`、`templates/`和`static/`
3. 在`routes.py`中定义新的蓝图和路由：

```python
from flask import Blueprint

file_bp = Blueprint('file', __name__, 
                   url_prefix='/file',
                   static_folder='static',
                   template_folder='templates')

@file_bp.route('/')
def index():
    return "File Processing Tool"
```

4. 在主`app.py`中导入并注册新的蓝图：

```python
from modules.file_processor.routes import file_bp
app.register_blueprint(file_bp)
```

5. 在根路径API的工具列表中添加新工具信息

## 部署指南

### 使用Docker Compose部署（推荐）

1. 确保服务器已安装Docker和Docker Compose

2. 部署应用：

```bash
docker compose up -d
```

3. 更新应用：

```bash
./update.sh
```

### 使用Nginx和SSL

项目包含了Nginx配置文件和SSL设置脚本，详细部署步骤请参考`DEPLOYMENT.md`文件。

## 安全注意事项

- 应用中包含邮箱授权码，生产环境中应使用环境变量或配置文件存储
- 建议在生产环境中启用HTTPS

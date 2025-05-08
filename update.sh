#!/bin/bash

# 设置颜色输出
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${YELLOW}开始更新邮件发送工具...${NC}"

# 确保logs目录存在
mkdir -p logs

# 拉取最新代码
echo -e "${YELLOW}正在拉取最新代码...${NC}"
git pull

# 使用Docker Compose重新构建并启动服务
echo -e "${YELLOW}正在重新构建并启动服务...${NC}"
docker compose down
docker compose build --no-cache
docker compose up -d

# 检查容器是否成功启动
if [ "$(docker ps -q -f name=email-app)" ]; then
    echo -e "${GREEN}更新成功！应用已重新部署。${NC}"
    echo -e "${GREEN}可通过以下地址访问：${NC}"
    echo -e "${GREEN}http://localhost:3000/email${NC}"
    echo -e "${GREEN}或者${NC}"
    echo -e "${GREEN}https://tool.jarvismedical.asia/email${NC}"
else
    echo -e "${YELLOW}警告：容器可能未成功启动，请检查日志。${NC}"
    docker compose logs
fi

# 显示容器状态
echo -e "${YELLOW}容器状态：${NC}"
docker ps | grep email-app

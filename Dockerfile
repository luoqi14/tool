FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
RUN pip install gunicorn

COPY . .

EXPOSE 3001

CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:3001", "--timeout", "600", "app:app"]

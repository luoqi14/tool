FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install gunicorn

COPY . .

EXPOSE 3001

CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:3001", "app:app"]

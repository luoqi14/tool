FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install gunicorn

COPY . .

EXPOSE 6000

CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:6000", "app:app"]

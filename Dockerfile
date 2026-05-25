FROM python:3.11-slim

WORKDIR /app

# Dependências do sistema (mínimas)
RUN apt-get update && apt-get install -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PYTHONUNBUFFERED=1
ENV PORT=10000
EXPOSE 10000

CMD gunicorn --bind 0.0.0.0:${PORT:-10000} --timeout 300 --workers 1 --threads 4 app:app

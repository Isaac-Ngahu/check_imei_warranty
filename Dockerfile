FROM python:3.11-slim

WORKDIR /app

COPY . .

RUN pip install --no-cache-dir -r requirements.txt
RUN pip install playwright

# 🔥 THIS is the key fix (handles Debian changes automatically)
RUN playwright install --with-deps chromium

CMD ["gunicorn", "app:app", "-b", "0.0.0.0:8080"]
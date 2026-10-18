FROM python:3.11-slim

WORKDIR /app

COPY . .

# RUN ls*-la /app

RUN pip install --no-cache-dir -r requirements.txt
# RUN pip install playwright
RUN apt-get update && apt-get install -y chromium chromium-driver
# 🔥 THIS is the key fix (handles Debian changes automatically)
# RUN playwright install --with-deps chromium

CMD ["gunicorn", "check_warranty:app", "-b", "0.0.0.0:8080"]
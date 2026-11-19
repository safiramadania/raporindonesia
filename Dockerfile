FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    APP_ENV=public

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir --index-url https://download.pytorch.org/whl/cpu torch==2.10.0+cpu \
    && pip install --no-cache-dir -r requirements.txt

COPY app.py input_dictionary.json ./
COPY model_bundle ./model_bundle
COPY static ./static
COPY templates ./templates

EXPOSE 8080
CMD ["sh", "-c", "exec gunicorn app:app --bind 0.0.0.0:${PORT:-8080} --workers 1 --threads 4 --timeout 180"]

FROM python:3.11-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

ARG PIP_INDEX_URL=https://pypi.org/simple

RUN apt-get update \
    && apt-get install --no-install-recommends -y \
        libgl1 \
        libglib2.0-0 \
        libgomp1 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN python -m pip install --index-url "${PIP_INDEX_URL}" --upgrade pip \
    && python -m pip install --index-url "${PIP_INDEX_URL}" -r requirements.txt

COPY . .

RUN mkdir -p /app/storage \
    && chown -R 1000:1000 /app

USER 1000:1000

EXPOSE 5000

CMD ["gunicorn", "--bind=0.0.0.0:5000", "--workers=1", "--worker-class=gthread", "--threads=4", "--timeout=600", "--graceful-timeout=60", "--access-logfile=-", "--error-logfile=-", "app:app"]

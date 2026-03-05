FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice \
    && rm -rf /var/lib/apt/lists/*

COPY pyproject.toml ./
RUN python - <<'PY' > /tmp/requirements.txt
import tomllib
from pathlib import Path

data = tomllib.loads(Path("pyproject.toml").read_text())
for dep in data.get("project", {}).get("dependencies", []):
    print(dep)
PY
RUN pip install --no-cache-dir -r /tmp/requirements.txt

COPY . .

RUN mkdir -p /app/static/uploads /data

EXPOSE 8000

CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--workers", "2", "--threads", "4", "--timeout", "120", "main:app"]

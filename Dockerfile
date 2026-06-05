# WGHT Vehicle Mass — single service: FastAPI serves the API AND the web front
# (same origin, so no CORS / no __WGHT_API_BASE__ needed).
FROM python:3.12-slim

WORKDIR /app

# curl is used to fetch the browser vendor bundles at build time (kept out of git).
RUN apt-get update \
    && apt-get install -y --no-install-recommends curl ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Python deps first for better layer caching.
COPY api/requirements.txt api/requirements.txt
RUN pip install --no-cache-dir -r api/requirements.txt

# App code + front-end.
COPY api api
COPY web web

# Vue + HyperFormula are gitignored (web/vendor/). Download the exact versions the
# app expects (see scripts/setup-web-vendor.ps1) so the front boots on the server.
RUN mkdir -p web/vendor \
    && curl -fsSL https://cdn.jsdelivr.net/npm/vue@3.5.13/dist/vue.esm-browser.prod.js \
        -o web/vendor/vue.esm-browser.prod.js \
    && curl -fsSL https://cdn.jsdelivr.net/npm/hyperformula@3.0.0/dist/hyperformula.full.min.js \
        -o web/vendor/hyperformula.full.min.js

ENV PORT=8000
EXPOSE 8000

# main.py mounts ../../../web relative to api/app/main.py, so run from api/.
WORKDIR /app/api
CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-8000}"]

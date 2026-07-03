FROM node:20-bookworm-slim

WORKDIR /app

ARG NPM_REGISTRY=
ARG PIP_INDEX_URL=

RUN apt-get update \
  && apt-get install -y --no-install-recommends ca-certificates python3 python3-venv \
  && rm -rf /var/lib/apt/lists/*

RUN python3 -m venv /opt/product-huizhen-venv

ENV NODE_ENV=production \
  PORT=4000 \
  PYTHON_BIN=/opt/product-huizhen-venv/bin/python \
  CHATBI_STORAGE_DIR=/app/storage \
  PATH=/opt/product-huizhen-venv/bin:$PATH

COPY package*.json ./

RUN if [ -n "$NPM_REGISTRY" ]; then npm config set registry "$NPM_REGISTRY"; fi \
  && npm ci --omit=dev --ignore-scripts \
  && npm cache clean --force

RUN if [ -n "$PIP_INDEX_URL" ]; then pip config set global.index-url "$PIP_INDEX_URL"; fi \
  && python -m pip install --no-cache-dir --upgrade pip \
  && python -m pip install --no-cache-dir openpyxl

COPY public ./public
COPY scripts/analyze_delisting.py ./scripts/analyze_delisting.py
COPY server.js ./server.js

RUN mkdir -p /app/storage \
  && chown -R node:node /app

USER node

EXPOSE 4000

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:' + (process.env.PORT || 4000) + '/api/status').then((res) => process.exit(res.ok ? 0 : 1)).catch(() => process.exit(1))"

CMD ["npm", "start"]

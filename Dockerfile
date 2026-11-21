FROM node:24.22.3-alpine3.24@sha256:116b800c3754b440ee0aa3ed2f20f5bf3d99a554bfd9adeea09ea8017f6e85f5 AS deps
WORKDIR /app
# pnpm-workspace.yaml carries the pnpm overrides recorded in the lockfile; without it --frozen-lockfile fails.
COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./
RUN npm install -g pnpm@12.21.0 && pnpm install --prod --frozen-lockfile

FROM node:24.22.3-alpine3.24@sha256:116b800c3754b440ee0aa3ed2f20f5bf3d99a554bfd9adeea09ea8017f6e85f5
ENV NODE_ENV=production DATA_DIR=/data APP_PORT=8790 CONTENT_PORT=8791
WORKDIR /app
COPY --from=deps /app/node_modules ./node_modules
COPY package.json ./
COPY src ./src
COPY public ./public
COPY content ./content
# The server runs plain node; the base image's bundled npm and corepack are never used at runtime,
# and npm's own dependencies (undici, brace-expansion) were every HIGH finding Trivy reported here.
RUN rm -rf /usr/local/lib/node_modules/npm /usr/local/lib/node_modules/corepack \
      /usr/local/bin/npm /usr/local/bin/npx /usr/local/bin/corepack \
  && mkdir -p /data && chown node:node /data
USER node
VOLUME ["/data"]
EXPOSE 8790 8791
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s CMD wget -qO- http://127.0.0.1:8790/healthz >/dev/null || exit 1
CMD ["node", "src/server.ts"]

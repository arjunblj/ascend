FROM oven/bun:1 AS base
WORKDIR /app

FROM base AS deps
COPY package.json bun.lock ./
COPY packages/*/package.json ./packages/
COPY apps/*/package.json ./apps/
RUN for f in packages/*/package.json apps/*/package.json; do \
      mkdir -p "$(dirname "$f")" && mv "$f" "$(dirname "$f")/"; \
    done
RUN bun install --frozen-lockfile

FROM base AS build
COPY --from=deps /app/node_modules ./node_modules
COPY . .
RUN bunx tsc --build

FROM base AS runtime
COPY --from=deps /app/node_modules ./node_modules
COPY . .



ARG NODE_IMAGE=node:22-alpine
ARG NGINX_IMAGE=nginxinc/nginx-unprivileged:1.29-alpine

# ---------- 1. build ----------
FROM --platform=$BUILDPLATFORM ${NODE_IMAGE} AS build
WORKDIR /src
ENV NODE_OPTIONS=--max-old-space-size=3072 \
    npm_config_fund=false \
    npm_config_audit=false \
    npm_config_update_notifier=false

# Dependencies first, for layer caching. --ignore-scripts: no package may run install hooks.
COPY frontend/package.json frontend/package-lock.json ./
RUN npm ci --ignore-scripts --include=dev

COPY frontend/ ./
RUN npm run build \
 && find dist -type f \( -name '*.js' -o -name '*.css' -o -name '*.html' -o -name '*.svg' \
        -o -name '*.wasm' -o -name '*.json' -o -name '*.ttf' -o -name '*.otf' \) \
        -size +1k -exec gzip -9 -k -n {} +

# ---------- 2. runtime ----------
FROM ${NGINX_IMAGE}

LABEL org.opencontainers.image.title="Greek Office" \
      org.opencontainers.image.description="In-browser Word (.docx) and Excel (.xlsx) editors; documents never leave the browser." \
      org.opencontainers.image.licenses="Apache-2.0"

USER root

RUN rm -rf /etc/nginx/conf.d/* /docker-entrypoint.d/* /usr/share/nginx/html/* \
 && mkdir -p /etc/nginx/snippets \
 && (apk del --no-cache apk-tools >/dev/null 2>&1 || true)

COPY docker/nginx.conf             /etc/nginx/nginx.conf
COPY docker/security-headers.conf  /etc/nginx/snippets/security-headers.conf
COPY docker/entrypoint.sh          /usr/local/bin/greek-entrypoint.sh
COPY --from=build /src/dist        /usr/share/nginx/html

# Content, config and the start-up script are root-owned and read-only for the nginx user.
RUN chown -R root:root /usr/share/nginx/html /etc/nginx /usr/local/bin/greek-entrypoint.sh \
 && chmod -R a=rX,u+w /usr/share/nginx/html /etc/nginx \
 && chmod 755 /usr/local/bin/greek-entrypoint.sh

USER 101:101
EXPOSE 8080
STOPSIGNAL SIGQUIT
HEALTHCHECK --interval=30s --timeout=3s --start-period=10s --retries=3 \
  CMD wget -q -O /dev/null http://127.0.0.1:8080/healthz || exit 1


ENTRYPOINT ["/usr/local/bin/greek-entrypoint.sh"]
CMD ["nginx", "-g", "daemon off;"]

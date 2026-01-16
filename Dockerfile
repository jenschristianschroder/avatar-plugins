# Build stage: install dependencies and bundle the front-end assets
FROM node:20-alpine AS build
WORKDIR /app

# Install dependencies required to build the static bundle
COPY package*.json ./
RUN npm install

# Copy the rest of the source and build using webpack
COPY . .
RUN npx webpack --config webpack.config.js

# Runtime stage: serve static assets and run proxy services
FROM node:20-alpine AS runtime
WORKDIR /app

ARG CONFIG_FILE=settings.json

ENV NODE_ENV=production \
	STATIC_ASSETS_DIR=/app/static \
	SERVICES_PROXY_PORT=8080 \
	AGENT_PROXY_PORT=4000 \
	CONFIG_FILE=${CONFIG_FILE}

# Copy application source and build artifacts
COPY --from=build /app/services-proxy-server ./services-proxy-server
COPY --from=build /app/agent-proxy-server ./agent-proxy-server
COPY --from=build /app/common ./common
COPY --from=build /app/config ./config
COPY --from=build /app/scripts ./scripts
COPY --from=build /app/shared ./shared
COPY --from=build /app/plugins ./plugins
COPY --from=build /app/dist ./static/dist
COPY --from=build /app/css ./static/css
COPY --from=build /app/js ./static/js
COPY --from=build /app/image ./static/image
COPY --from=build /app/video ./static/video
COPY --from=build /app/*.html ./static/
COPY --from=build /app/favicon.ico ./static/

# Install production dependencies for the proxy services
RUN npm install --omit=dev --prefix agent-proxy-server && \
	npm install --omit=dev --prefix services-proxy-server

# Align primary settings file with selected configuration source
RUN if [ "$CONFIG_FILE" != "settings.json" ]; then \
	cp "./config/${CONFIG_FILE}" ./config/settings.json ; \
fi

# Expose ports for the services proxy (static + config relay) and agent proxy
EXPOSE 80
EXPOSE 4000

# Launch both proxy servers inside the container
CMD ["node", "scripts/start-container.js"]

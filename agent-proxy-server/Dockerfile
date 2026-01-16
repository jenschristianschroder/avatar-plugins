FROM node:20-alpine
WORKDIR /app
COPY package*.json ./
RUN npm ci --omit=dev
COPY src ./src
ENV NODE_ENV=production \
    PORT=4000
EXPOSE 4000
CMD ["node", "src/index.js"]
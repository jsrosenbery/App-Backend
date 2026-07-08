FROM node:20-bookworm-slim

ENV NODE_ENV=production
ENV LIBREOFFICE_PATH=/usr/bin/soffice

RUN apt-get update \
  && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    fonts-dejavu \
    fonts-liberation \
    fontconfig \
  && apt-get clean \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm install --omit=dev

COPY . .

EXPOSE 3000

CMD ["node", "server.js"]

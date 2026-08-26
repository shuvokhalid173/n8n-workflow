FROM node:24-alpine

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY play-with-n8n.js ./

CMD ["npm", "start"]

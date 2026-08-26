FROM node:24-alpine

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY n8n-outlook.js ./

CMD ["npm", "start"]
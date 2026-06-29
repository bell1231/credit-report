FROM node:22-alpine

RUN apk add --no-cache ca-certificates

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm install --production

COPY . .

EXPOSE 3000

CMD ["node", "server.js"]

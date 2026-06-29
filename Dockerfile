FROM node:22-alpine

RUN apk add --no-cache ca-certificates

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm install --production

# 安装 CloudBase SDK（容器内自动鉴权，用于线上缓存）
RUN npm install @cloudbase/node-sdk

COPY . .

EXPOSE 3000

CMD ["node", "server.js"]

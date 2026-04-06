FROM node:20-alpine
WORKDIR /app
COPY package*.json ./
RUN npm install --omit=dev
RUN npm install -g supergateway
COPY . .
EXPOSE 8000
CMD ["supergateway", "--stdio", "node index.js", "--port", "8000", "--baseUrl", "/", "--ssePath", "/sse", "--messagePath", "/message"]
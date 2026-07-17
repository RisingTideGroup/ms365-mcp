FROM node:22-alpine
WORKDIR /app
COPY package*.json tsconfig.json ./
RUN npm ci --no-audit --no-fund
COPY src ./src
RUN npx esbuild src/*.ts src/http/*.ts --outdir=build --outbase=src --format=esm --platform=node --target=es2022
ENV NODE_ENV=production PORT=3000
EXPOSE 3000
USER node
CMD ["node", "build/http/http-main.js"]

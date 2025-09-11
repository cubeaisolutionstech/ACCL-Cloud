# Step 1: Build React app
FROM node:20-alpine as build

WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm install --legacy-peer-deps
COPY . .
RUN npm run build

# Step 2: Serve with Nginx
FROM nginx:alpine
COPY --from=build /app/build /usr/share/nginx/html

# nginx.conf will be mounted from root (via docker-compose volume)
EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]

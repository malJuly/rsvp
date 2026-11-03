# Stage 1: build the static site with Node
FROM node:20-alpine AS build
WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY . .
# Vite bakes VITE_* values into the JS at build time, so they are passed in here
ARG VITE_RSVP_SCRIPT_URL
ARG VITE_SITE_URL
ENV VITE_RSVP_USE_MOCK=false
ENV VITE_RSVP_SCRIPT_URL=$VITE_RSVP_SCRIPT_URL
# The live address, for the WhatsApp/Facebook preview image (needs a full https:// address)
ENV VITE_SITE_URL=$VITE_SITE_URL
RUN npm run build

# Stage 2: serve the built files with nginx (runs as non-root)
FROM nginxinc/nginx-unprivileged:1.27-alpine
COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=build /app/dist /usr/share/nginx/html
EXPOSE 9100

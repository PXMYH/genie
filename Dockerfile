FROM node:13.3-alpine

RUN mkdir -p /app
WORKDIR /app
# COPY '/bin' '/app'
# COPY '/public' '/app'
# COPY '/routes' '/app'
# COPY '/views' '/app'
# COPY 'app.js', 'package.json', 'package-lock.json' '/app'
COPY . /app

CMD ['npm', 'start']
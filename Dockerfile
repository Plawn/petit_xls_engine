FROM node:lts-alpine

WORKDIR /api

COPY package.json yarn.lock ./

RUN yarn install

RUN yarn build

EXPOSE 3000

ENTRYPOINT [ "node build/start.js" ]
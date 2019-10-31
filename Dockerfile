FROM node:lts-alpine

COPY package.json yarn.lock ./

RUN yarn install

RUN yarn build

EXPOSE 3000
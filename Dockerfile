FROM node:lts-alpine

WORKDIR /api

COPY package.json yarn.lock ./

RUN yarn install

COPY . .

RUN yarn build

COPY package.json build/package.json
COPY yarn.lock build/yarn.lock

EXPOSE 3000

ENTRYPOINT [ "node build/start.js 3000 config.yaml" ]
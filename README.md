# Excel publiposting service

**Do NOT update jszip**

This service will be behind the api-doc service behind the excel-publiposting lib

## To start the service
 
```sh
yarn start
```
or
```sh
yarn build && node build/start.js <port> <config_file>
```

## TODO

- make better object copy in order to reduce latency and cpu usage
- make Dockerfile
- make docker-compose

This app should run in the end in the same container as the main api-doc service and should be hidden from the other applications
This app should only be queried by the main api-doc app


## Endpoints

POST /publipost

This endpoint will be used to publipost the template:
You need to send :
```json
{
    "data": {},
    "template_name":"filename",
    "output_name":"<name>",
    "bucket_name":"<name>"
}
```
### POST /documents

This endpoint will be used to get documents infos
You need to send :
```json
{
    "name":"<name>"
}
```

### POST /load_template

This service will be controlled by the global publiposting service, so we have to tell him what to load and how
You need to send:
```json
{
    "template_name":"<name>",
    "bucket_name":"<name>"
}
```

It will respond with an object containing the fails and successes
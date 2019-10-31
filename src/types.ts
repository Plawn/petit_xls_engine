import {Client} from 'minio';

export type reqPubli = {
    template_name:string;
    data:any;
    output_name:string;
    output_bucket:string;
}


export type template = {
    bucket_name: string;
    template_name: string;
    output: string;
};


export type configType = {
    port?: number;
    minio?: Client;
}


export type minioInfosType = {
    endpoint:string;
    passkey:string;
    access_key:string;
}
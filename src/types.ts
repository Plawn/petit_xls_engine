import {Client} from 'minio';

export type reqPubli = {
    template_name:string;
    data:any;
    output_name:string;
    output_bucket:string;
}

export type reqPullTemplate = {
    template_name:string;
    bucket_name:string;
    exposed_as:string;
}[];

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
    host:string;
    pass_key:string;
    access_key:string;
    secure:boolean;
}
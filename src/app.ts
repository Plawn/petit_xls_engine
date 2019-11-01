import express from 'express';
import bodyParser from 'body-parser';
import { Client } from 'minio';
import { asyncMiddleware, streamToBuffer } from './utils';
import templateDB from './TemplateDB';
import { reqPubli, configType, minioInfosType } from './types';

const app = express();

const db = new templateDB();
const config: configType = {
    port: null,
    minio: null,
}



app.use(bodyParser.json());

app.post('/publipost', asyncMiddleware(async (req, res) => {
    const data: reqPubli = req.body;
    const rendered = db.renderTemplate(data.template_name, data.data);
    await config.minio.putObject(data.output_bucket, data.output_name, rendered);
    res.send({ error: false });
}));


app.post('/documents', (req, res) => res.send(db.getPlaceholder(req.body.name)));


app.post('/load_templates', asyncMiddleware(async (req, res) => {
    const success = [];
    const failed = [];
    for (const element of req.body) {
        try {
            const stream = await config.minio.getObject(element.bucket_name, element.template_name);
            const b = await streamToBuffer(stream);
            await db.addTemplate(element.template_name, b);
            success.push(element.template_name)
        } catch (e) {
            console.warn(e);
            failed.push(element.template_name)
        }
    }
    res.send({ success: success, failed: failed });
}));


export default (port: number, minioInfos: minioInfosType, afterStart?: Function) => {
    config.port = port;
    app.listen(port, '127.0.0.1', async () => {
        config.minio = new Client({
            endPoint: minioInfos.endpoint,
            port: 443,
            useSSL: true,
            accessKey: minioInfos.access_key,
            secretKey: minioInfos.passkey,
        });
        console.log(`started on port ${port}`);
        if (afterStart) {
            afterStart();
        }
    });
}

import express from 'express';
import bodyParser from 'body-parser';
import { Client as MinioClient } from 'minio';
import { asyncMiddleware, streamToBuffer, portFromUrl } from './utils';
import templateDB from './TemplateDB';
import { reqPubli, configType, minioInfosType } from './types';

const app = express();
app.use(bodyParser.json());

let configured = false;
const db = new templateDB();
const config: configType = {
    minio: null,
}


app.post('/publipost', asyncMiddleware(async (req, res) => {
    if (!configured) {
        return res.send({ error: true });
    }
    const data: reqPubli = req.body;
    const rendered = db.renderTemplate(data.template_name, data.data);
    await config.minio.putObject(data.output_bucket, data.output_name, rendered);
    res.send({ error: false });
}));


app.post('/get_placeholders', (req, res) => {
    if (!configured) {
        return res.send({ error: true });
    }
    res.send(db.getPlaceholder(req.body.name))
});


app.post('/load_templates', asyncMiddleware(async (req, res) => {
    if (!configured) {
        return res.send({ error: true });
    }
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

// using this endpoint the app will be configured
app.post('/configure', asyncMiddleware(async (req, res) => {
    config.minio = new MinioClient({
        endPoint: req.body.endpoint,
        port: portFromUrl(req.body.endpoint),
        useSSL: req.body.secure,
        accessKey: req.body.access_key,
        secretKey: req.body.passkey,
    });
    await config.minio.listBuckets();
    configured = true;
    res.send({ error: false });
    console.log('Successfuly configured');
}));


export default async (port: number) => {
    app.listen(port, async () => {
        console.log(`started on port ${port}`);
    });
}

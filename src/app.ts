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
    minio: undefined,
}


app.post('/publipost', asyncMiddleware(async (req, res) => {
    if (!configured) {
        return res.status(402).send({ error: true });
    }
    const data: reqPubli = req.body;
    const rendered = db.renderTemplate(data.template_name, data.data);
    await config.minio.putObject(data.output_bucket, data.output_name, rendered);
    res.send({ error: false });
}));


app.post('/get_placeholders', (req, res) => {
    if (!configured) {
        return res.status(402).send({ error: true });
    }
    res.send(db.getPlaceholder(req.body.name))
});


app.post('/load_templates', asyncMiddleware(async (req, res) => {
    if (!configured) {
        return res.status(402).send({ error: true });
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
    try {
        const data:minioInfosType = req.body;
        config.minio = new MinioClient({
            endPoint: data.host.split(':')[0],
            port: portFromUrl(data.host),
            useSSL: data.secure,
            accessKey: data.access_key,
            secretKey: data.pass_key,
        });
        await config.minio.listBuckets();
        configured = true;
        console.log('Successfuly configured');
        res.status(200).send({ error: false });
    } catch (e) {
        configured = false;
        config.minio = undefined;
        res.status(402).send({ error: true });
        console.error(e);
    }
}));

app.get('/live', (_, res) => {
    if (configured) {
        res.status(200).send('OK');
    } else {
        res.status(402).send('KO');
    }
});

export default async (port: number) => {
    app.listen(port, async () => {
        console.log(`started on port ${port}`);
    });
}

import express from 'express';
import bodyParser from 'body-parser';
import { Client as MinioClient } from 'minio';
import { asyncMiddleware, streamToBuffer, portFromUrl } from './utils';
import templateDB from './TemplateDB';
import { reqPubli, configType, minioInfosType, reqPullTemplate } from './types';

const app = express();
app.use(bodyParser.json());

const db = new templateDB();
const config: configType = {
    minio: undefined,
}


app.post('/publipost', asyncMiddleware(async (req, res) => {
    if (!config.minio) {
        return res.status(402).send({ error: true });
    }
    const data: reqPubli = req.body;
    const rendered = db.renderTemplate(data.template_name, data.data);
    // could have some abstraction here
    await config.minio.putObject(data.output_bucket, data.output_name, rendered);
    res.send({ error: false });
}));


app.post('/get_placeholders', (req, res) => {
    if (!config.minio) {
        return res.status(402).send({ error: true });
    }
    res.send(db.getPlaceholder(req.body.name))
});

app.get('/list', (req, res) => {
    const result = {};
    db.templates.forEach((value, key) => {
        result[key] = value.pulled_at;
    });
    res.send(result);
});

app.delete('/remove_template', (req, res) => {
    try {
        db.removeTemplate(req.body.template_name);
        res.status(200).send({ error: false });
    } catch (e) {
        res.status(400).send({ error: true });
    }
});

app.post('/load_templates', asyncMiddleware(async (req, res) => {
    if (!config.minio) {
        return res.status(402).send({ error: true });
    }
    const success = [];
    const failed = [];
    const body: reqPullTemplate = req.body;
    for (const element of body) {
        try {
            const stream = await config.minio.getObject(element.bucket_name, element.template_name);
            const b = await streamToBuffer(stream);
            await db.addTemplate(element.exposed_as, b);
            success.push({
                template_name: element.exposed_as,
                fields: db.getPlaceholder(element.exposed_as)
            });
        } catch (e) {
            console.warn(e);
            failed.push({ template_name: element.exposed_as });
        }
    }
    console.log('loaded templates', { success, failed });
    res.send({ success, failed });
}));

// using this endpoint the app will be configured
app.post('/configure', asyncMiddleware(async (req, res) => {
    try {
        const data: minioInfosType = req.body;
        config.minio = new MinioClient({
            endPoint: data.host.split(':')[0],
            port: portFromUrl(data.host),
            useSSL: data.secure,
            accessKey: data.access_key,
            secretKey: data.pass_key,
        });
        await config.minio.listBuckets();
        console.log('Successfuly configured');
        res.status(200).send({ error: false });
    } catch (e) {
        config.minio = undefined;
        res.status(402).send({ error: true });
        console.error(e);
    }
}));

app.get('/live', (_, res) => {
    if (config.minio) {
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

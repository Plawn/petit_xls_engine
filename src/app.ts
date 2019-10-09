import express from 'express';
import bodyParser from 'body-parser';
import { Client } from 'minio';
import { asyncMiddleware, streamToBuffer } from './utils';
import templateDB from './TemplateDB';
import { reqPubli } from './requestTypes'

const db = new templateDB();

const port = 3001;
const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));


// should use env variables of conf file
const minio = new Client({
    endPoint: 'documents.juniorisep.com',
    port: 443,
    useSSL: true,
    accessKey: 'adminadmin',
    secretKey: 'adminadmin'
});


app.post('/publipost', asyncMiddleware(async (req, res) => {
    const data: reqPubli = req.body;
    const rendered = await db.renderTemplate(data.template_name, data.data);
    await minio.putObject(data.output_bucket, data.output_name, rendered);
    res.send({ error: false });
}));


app.post('/documents', (req, res) => res.send(db.getPlaceholder(req.body.name)));


app.post('/load_templates', asyncMiddleware(async (req, res) => {
    let success = [];
    let failed = [];
    for (const element of req.body) {
        try {
            const stream = await minio.getObject(element.bucket_name, element.template_path);
            const b = await streamToBuffer(stream);
            db.addTemplate(element.template_name, b);
            success.push(element.template_name)
        } catch (e) {
            failed.push(element.template_name)
        }
    }
    res.send({ success: success, failed: failed });
}));


app.listen(port, async () => {

    console.log(`started on port ${port}`);

});

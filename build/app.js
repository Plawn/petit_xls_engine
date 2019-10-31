"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const body_parser_1 = __importDefault(require("body-parser"));
const minio_1 = require("minio");
const utils_1 = require("./utils");
const TemplateDB_1 = __importDefault(require("./TemplateDB"));
const db = new TemplateDB_1.default();
const config = {
    port: null,
    minio: null,
};
const app = express_1.default();
app.use(body_parser_1.default.json());
app.use(body_parser_1.default.urlencoded({ extended: true }));
app.post('/publipost', utils_1.asyncMiddleware((req, res) => __awaiter(void 0, void 0, void 0, function* () {
    const data = req.body;
    const rendered = yield db.renderTemplate(data.template_name, data.data);
    yield config.minio.putObject(data.output_bucket, data.output_name, rendered);
    res.send({ error: false });
})));
app.post('/documents', (req, res) => res.send(db.getPlaceholder(req.body.name)));
app.post('/load_templates', utils_1.asyncMiddleware((req, res) => __awaiter(void 0, void 0, void 0, function* () {
    const success = [];
    const failed = [];
    for (const element of req.body) {
        try {
            const stream = yield config.minio.getObject(element.bucket_name, element.template_name);
            const b = yield utils_1.streamToBuffer(stream);
            db.addTemplate(element.template_name, b);
            success.push(element.template_name);
        }
        catch (e) {
            console.warn(e);
            failed.push(element.template_name);
        }
    }
    res.send({ success: success, failed: failed });
})));
exports.default = (port, minioInfos) => {
    config.port = port;
    app.listen(port, '127.0.0.1', () => __awaiter(void 0, void 0, void 0, function* () {
        config.minio = new minio_1.Client({
            endPoint: minioInfos.endpoint,
            port: 443,
            useSSL: true,
            accessKey: minioInfos.access_key,
            secretKey: minioInfos.passkey,
        });
        console.log(`started on port ${port}`);
    }));
};

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
const excel_module_1 = __importDefault(require("./excel_module"));
const express_1 = __importDefault(require("express"));
const body_parser_1 = __importDefault(require("body-parser"));
const minio_1 = require("minio");
const to_array_buffer_1 = __importDefault(require("to-array-buffer"));
function streamToBuffer(stream) {
    return new Promise((resolve, reject) => {
        let buffers = [];
        stream.on('error', reject);
        stream.on('data', (data) => buffers.push(data));
        stream.on('end', () => resolve(Buffer.concat(buffers)));
    });
}
const minio = new minio_1.Client({
    endPoint: 'documents.juniorisep.com',
    port: 443,
    useSSL: true,
    accessKey: 'adminadmin',
    secretKey: 'adminadmin'
});
class templateDB {
    constructor() {
        this.addTemplate = (name, data) => {
            this.db[name] = data;
            this.loadedDB[name] = new excel_module_1.default(data);
        };
        this.renderTemplate = (filename, data) => __awaiter(this, void 0, void 0, function* () {
            const template = new excel_module_1.default(this.db[filename]);
            template.sheets.forEach((sheet) => template.substitute(sheet.id, data));
            return Buffer.from(to_array_buffer_1.default(template.generate()));
        });
        this.getPlaceholder = (name) => this.loadedDB[name].getAllPlaceholders();
        this.db = {};
        this.loadedDB = {};
    }
}
const db = new templateDB();
const port = 3001;
const app = express_1.default();
app.use(body_parser_1.default.json());
app.use(body_parser_1.default.urlencoded({ extended: true }));
app.post('/publipost', (req, res) => __awaiter(void 0, void 0, void 0, function* () {
    const data = req.body;
    const generated = yield db.renderTemplate(data.template_name, data.data);
    const re = yield minio.putObject(data.output_bucket, data.output_name, generated);
    res.send(re);
}));
// post for now will be get later on
app.post('/documents', (req, res) => {
    res.send(db.getPlaceholder(req.body.name));
});
app.post('/load_templates', (req, res) => __awaiter(void 0, void 0, void 0, function* () {
    let success = [];
    let failed = [];
    for (const element of req.body) {
        try {
            const stream = yield minio.getObject(element.bucket_name, element.template_name);
            const b = yield streamToBuffer(stream);
            // das working ok
            // await fs.writeFile('dl.xlsx', b);
            db.addTemplate(element.template_name, b);
            success.push(element.template_name);
        }
        catch (e) {
            failed.push(element.template_name);
        }
    }
    res.send({ success: success, failed: failed });
}));
// const filename = './ndf.xlsx';
// const values = {
//     prix:{
//         deb:15.27
//     },
//     date:'HEBE'
// };
// (async () => {
//     const t = new Date().getTime();
//     console.log('started')
//     const res = await publipost(filename, values);
//     console.log('done', (new Date().getTime() - t));
//     await save(res, 'res.xlsx');
//     const deb = await getVariables(filename);
//     console.log(deb);
//     console.log('end of test');
// })();
// dats a test
app.listen(port, () => __awaiter(void 0, void 0, void 0, function* () {
    console.log('started');
}));

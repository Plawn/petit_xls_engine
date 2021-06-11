import makeConnector from 'petit_nodejs_publipost_connector';
import { ExcelTemplate } from './TemplateDB';

// Main wrapped for asyncness
const main = () => {
    const port = Number(process.argv[2]);
    if (isNaN(port) || port > 65535) {
        throw new Error(`invalid port ${port}`)
    }
    const app = makeConnector(ExcelTemplate);
    app.listen(3000, () => {
        console.log(`Connector started on port ${port}`);
    });
};

main();



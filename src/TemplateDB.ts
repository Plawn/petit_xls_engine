import XlsxTemplate from './excel_module';
import ab from 'to-array-buffer';
import { SafeMap } from './utils';

class templateDB {
    db: SafeMap<string, Buffer>;
    loadedDB: SafeMap<string, XlsxTemplate>;

    constructor() {
        this.db = new SafeMap<string, Buffer>();
        this.loadedDB = new SafeMap<string, XlsxTemplate>();
    }

    addTemplate = (name: string, data: Buffer) => {
        this.db.set(name, data);
        this.loadedDB.set(name, new XlsxTemplate(data));
    }

    renderTemplate = (filename: string, data: any) => {
        const template = new XlsxTemplate(this.db.safeGet(filename));
        template.sheets.forEach((sheet: { id: number; }) => template.substitute(sheet.id, data));
        return Buffer.from(ab(template.generate()));
    }

    getPlaceholder = (name: string) => this.loadedDB.safeGet(name).getAllPlaceholders();
}

export default templateDB;
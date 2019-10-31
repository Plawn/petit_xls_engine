import XlsxTemplate from './excel_module';
import ab from 'to-array-buffer';
import { SafeMap } from './utils';

class templateDB {
    db: SafeMap<string, Buffer>;
    loadedDB: SafeMap<string, XlsxTemplate>;
    placeHolders: SafeMap<string, any>;

    constructor() {
        this.db = new SafeMap();
        this.loadedDB = new SafeMap();
        this.placeHolders = new SafeMap();
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

    getPlaceholder = (name: string) => {
        if (!this.placeHolders.has(name)) {
            const res = this.loadedDB.safeGet(name).getAllPlaceholders();
            this.placeHolders.set(name, res);
            return res;
        }
        return this.placeHolders.get(name);

    };
}

export default templateDB;
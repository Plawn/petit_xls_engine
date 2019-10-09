import XlsxTemplate from './excel_module';
import ab from 'to-array-buffer';

class templateDB {
    db: { [s: string]: Buffer; };
    loadedDB: { [s: string]: XlsxTemplate; };
    
    constructor() {
        this.db = {};
        this.loadedDB = {};
    }

    addTemplate = (name: string, data: Buffer) => {
        this.db[name] = data;
        this.loadedDB[name] = new XlsxTemplate(data);
    }

    renderTemplate = async (filename: string, data: any) => {
        const template: XlsxTemplate = new XlsxTemplate(this.db[filename]);
        template.sheets.forEach((sheet: { id: number; }) => template.substitute(sheet.id, data));
        return Buffer.from(ab(template.generate()));
    }

    getPlaceholder = (name: string) => this.loadedDB[name].getAllPlaceholders();
}

export default templateDB;
import XlsxTemplate from './excel_module';
import ab from 'to-array-buffer';
import { SafeMap } from './utils';

const delimiters = { start: '{{', end: '}}' };

class TemplateContainer {
    pulled_at: number;
    template: XlsxTemplate;
    placeHolders: any;
    data: Buffer;

    constructor(pulled_at: number, data: Buffer) {
        this.data = data;
        this.pulled_at = pulled_at;
        this.template = new XlsxTemplate(data, delimiters);
        this.placeHolders = this.template.getAllPlaceholders();
    }
};


class templateDB {
    templates: SafeMap<string, TemplateContainer>;

    constructor() {
        this.templates = new SafeMap();
    }

    addTemplate = (name: string, data: Buffer) => {
        const now = new Date().getTime() / 1000; // we get the time in millis and we want it in seconds
        this.templates.set(name, new TemplateContainer(now, data));
    }

    removeTemplate = (name: string) => {
        this.templates.delete(name);
    }

    renderTemplate = (filename: string, data: any) => {
        const template = new XlsxTemplate(this.templates.get(filename).data, delimiters);
        template.sheets.forEach((sheet: { id: number | string; }) => template.substitute(sheet.id, data));
        return Buffer.from(ab(template.generate()));
    }

    getPlaceholder = (name: string) => {
        return this.templates.get(name).placeHolders;
    };
}

export default templateDB;
import fs from 'mz/fs';
import path from 'path';
import save from 'save-file';
import XlsxTemplate from './excel_module';


const filename = './ndf.xlsx'

const values = {
    'prix': 15.27,
    "date": 'HEBEB'
}

const publipost = async (filename: string, data: any): Promise<Blob> => {
    const filedata = await fs.readFile(filename);
    //placeholder for now
    const sheetNumber = 1;
    const template = new XlsxTemplate(filedata);
    template.substitute(sheetNumber, data);
    return template.generate();
}


const getVariables = async (filename: string):Promise<string[]> => {
    const filedata = await fs.readFile(filename);
    // placeholder for now
    const sheetNumber = 1;
    const template = new XlsxTemplate(filedata);
    return template.getAllPlaceholder(sheetNumber);
}

(async () => {
    const res = await publipost(filename, values);
    const deb = await getVariables(filename);
    console.log(deb);
    console.log('end of test');
})();



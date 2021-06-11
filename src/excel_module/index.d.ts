declare class Workbook {
    sheets: any[];
    substitute: (sheetId: number | string, data: any) => void;
    generate: (options?: any) => any;
    getAllPlaceholders: () => string[];
    constructor(data: any, delimiters: any);
}

export default Workbook;
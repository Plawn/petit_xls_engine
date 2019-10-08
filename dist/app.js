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
const fs_1 = __importDefault(require("mz/fs"));
const excel_module_1 = __importDefault(require("./excel_module"));
const filename = './ndf.xlsx';
const values = {
    'prix': 15.27,
    "date": 'HEBEB'
};
const publipost = (filename, data) => __awaiter(void 0, void 0, void 0, function* () {
    const filedata = yield fs_1.default.readFile(filename);
    //placeholder for now
    const sheetNumber = 1;
    const template = new excel_module_1.default(filedata);
    template.substitute(sheetNumber, data);
    return template.generate();
});
const getVariables = (filename) => __awaiter(void 0, void 0, void 0, function* () {
    const filedata = yield fs_1.default.readFile(filename);
    // placeholder for now
    const sheetNumber = 1;
    const template = new excel_module_1.default(filedata);
    return template.getAllPlaceholder(sheetNumber);
});
(() => __awaiter(void 0, void 0, void 0, function* () {
    const res = yield publipost(filename, values);
    const deb = yield getVariables(filename);
    console.log(deb);
    console.log('end of test');
}))();

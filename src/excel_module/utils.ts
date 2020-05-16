"use strict";

const _get_simple = (obj: { [x: string]: any; }, desc: string) => {
    if (desc.indexOf("[") >= 0) {
        const specification = desc.split(/[[[\]]/);
        const property = specification[0];
        const index = specification[1];
        return obj[property][index];
    }
    return obj[desc];
}


/**
* Based on http://stackoverflow.com/questions/8051975
* Mimic https://lodash.com/docs#get
*/
export const _get = (obj: { [x: string]: any; }, desc: string, defaultValue: string) => {
    var arr = desc.split('.');
    try {
        while (arr.length) {
            obj = _get_simple(obj, arr.shift());
        }
    } catch (ex) {
        /* invalid chain */
        obj = undefined;
    }
    return obj === undefined ? defaultValue : obj;
}

// Split a reference into an object with keys `row` and `col` and,
// optionally, `table`, `rowAbsolute` and `colAbsolute`.
export const splitRef = function (ref: string) {
    const match = ref.match(/(?:(.+)!)?(\$)?([A-Z]+)(\$)?([0-9]+)/);
    return {
        table: match && match[1] || null,
        colAbsolute: Boolean(match && match[2]),
        col: match && match[3],
        rowAbsolute: Boolean(match && match[4]),
        row: parseInt(match && match[5], 10)
    };
};

// Join an object with keys `row` and `col` into a single reference string
export const joinRef = ref =>
    (ref.table ? ref.table + "!" : "") +
    (ref.colAbsolute ? "$" : "") +
    ref.col.toUpperCase() +
    (ref.rowAbsolute ? "$" : "") +
    Number(ref.row).toString();

// Get the next row's cell reference given a reference like "B2".
export const nextRow = (ref: string) => {
    ref = ref.toUpperCase();
    return ref.replace(/[0-9]+/, match => {
        return (parseInt(match, 10) + 1).toString();
    });
};

// Turn a reference like "AA" into a number like 27
export const charToNum = (str: string) => {
    let num = 0;
    for (let idx = str.length - 1, iteration = 0; idx >= 0; --idx, ++iteration) {
        const thisChar = str.charCodeAt(idx) - 64 // A -> 1; B -> 2; ... Z->26
        const multiplier = Math.pow(26, iteration);
        num += multiplier * thisChar;
    }
    return num;
};

// Is ref a range?
export const isRange = (ref: string) => ref.indexOf(':') !== -1;


export const joinRange = (range: { start: string; end: string; }) => range.start + ":" + range.end;



export const splitRange = (range: string) => {
    const split = range.split(":");
    return {
        start: split[0],
        end: split[1]
    };
};

// Replace all children of `parent` with the nodes in the list `children`
export const replaceChildren = (parent: { delSlice: (arg0: number, arg1: any) => void; len: () => void; append: (arg0: any) => void; }, children: any[]) => {
    parent.delSlice(0, parent.len());
    children.forEach(function (child) {
        parent.append(child);
    });
};

// Calculate the current row based on a source row and a number of new rows
// that have been inserted above
export const getCurrentRow = function (row: { attrib: { r: string; }; }, rowsInserted: number) {
    return parseInt(row.attrib.r, 10) + rowsInserted;
};


export const quoteRegex = function (str: string) {
    return str.replace(/([.?*+^$[\]\\(){}|-])/g, "\\$1");
};
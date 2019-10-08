"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const _get_simple = (obj, desc) => {
    if (desc.indexOf("[") >= 0) {
        var specification = desc.split(/[[[\]]/);
        var property = specification[0];
        var index = specification[1];
        return obj[property][index];
    }
    return obj[desc];
};
/**
* Based on http://stackoverflow.com/questions/8051975
* Mimic https://lodash.com/docs#get
*/
exports._get = (obj, desc, defaultValue) => {
    var arr = desc.split('.');
    try {
        while (arr.length) {
            obj = _get_simple(obj, arr.shift());
        }
    }
    catch (ex) {
        /* invalid chain */
        obj = undefined;
    }
    return obj === undefined ? defaultValue : obj;
};
// Split a reference into an object with keys `row` and `col` and,
// optionally, `table`, `rowAbsolute` and `colAbsolute`.
exports.splitRef = function (ref) {
    var match = ref.match(/(?:(.+)!)?(\$)?([A-Z]+)(\$)?([0-9]+)/);
    return {
        table: match && match[1] || null,
        colAbsolute: Boolean(match && match[2]),
        col: match && match[3],
        rowAbsolute: Boolean(match && match[4]),
        row: parseInt(match && match[5], 10)
    };
};
// Join an object with keys `row` and `col` into a single reference string
exports.joinRef = ref => (ref.table ? ref.table + "!" : "") +
    (ref.colAbsolute ? "$" : "") +
    ref.col.toUpperCase() +
    (ref.rowAbsolute ? "$" : "") +
    Number(ref.row).toString();
// Get the next row's cell reference given a reference like "B2".
exports.nextRow = ref => {
    ref = ref.toUpperCase();
    return ref.replace(/[0-9]+/, function (match) {
        return (parseInt(match, 10) + 1).toString();
    });
};
// Turn a reference like "AA" into a number like 27
exports.charToNum = (str) => {
    var num = 0;
    for (var idx = str.length - 1, iteration = 0; idx >= 0; --idx, ++iteration) {
        var thisChar = str.charCodeAt(idx) - 64, // A -> 1; B -> 2; ... Z->26
        multiplier = Math.pow(26, iteration);
        num += multiplier * thisChar;
    }
    return num;
};
// Is ref a range?
exports.isRange = (ref) => ref.indexOf(':') !== -1;
exports.joinRange = (range) => range.start + ":" + range.end;
exports.splitRange = (range) => {
    var split = range.split(":");
    return {
        start: split[0],
        end: split[1]
    };
};
// Replace all children of `parent` with the nodes in the list `children`
exports.replaceChildren = (parent, children) => {
    parent.delSlice(0, parent.len());
    children.forEach(function (child) {
        parent.append(child);
    });
};
// Calculate the current row based on a source row and a number of new rows
// that have been inserted above
exports.getCurrentRow = function (row, rowsInserted) {
    return parseInt(row.attrib.r, 10) + rowsInserted;
};

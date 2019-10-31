/*jshint globalstrict:true, devel:true */
/*eslint no-var:0 */
/*global require, module, Buffer */
"use strict";

import path from 'path';
import zip from 'jszip';
import etree from 'elementtree';

import {
    _get,
    splitRef,
    joinRef,
    nextRow,
    charToNum,
    isRange,
    joinRange,
    splitRange,
    replaceChildren,
    getCurrentRow
} from "./utils";

const DOCUMENT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const CALC_CHAIN_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
const SHARED_STRINGS_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const HYPERLINK_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";


const extractPlaceholders = string => {
    // Yes, that's right. It's a bunch of brackets and question marks and stuff.
    const re = /\${(?:(.+?):)?(.+?)(?:\.(.+?))?}/g;

    let match = null
    let matches = [];
    while ((match = re.exec(string)) !== null) {
        matches.push({
            placeholder: match[0],
            type: match[1] || 'normal',
            name: match[2],
            key: match[3],
            full: match[0].length === string.length
        });
    }

    return matches;
};




// Turn a number like 27 into a reference like "AA"
const numToColumnIdentifier = num => {
    let str = "";

    for (let i = 0; num > 0; ++i) {
        const remainder = num % 26;
        const charCode = remainder + 64;
        num = (num - remainder) / 26;
        // Compensate for the fact that we don't represent zero, e.g. A = 1, Z = 26, but AA = 27
        if (remainder === 0) { // 26 -> Z
            charCode = 90;
            --num;
        }
        str = String.fromCharCode(charCode) + str;
    }
    return str;
};

// Adjust the row `spans` attribute by `cellsInserted`
const updateRowSpan = (row, cellsInserted) => {
    if (cellsInserted !== 0 && row.attrib.spans) {
        let rowSpan = row.attrib.spans.split(':').map(f => parseInt(f, 10));
        rowSpan[1] += cellsInserted;
        row.attrib.spans = rowSpan.join(":");
    }
};

// Get a list of sheet ids, names and filenames
const loadSheets = (prefix, workbook, workbookRels) => {
    const sheets = [];

    workbook.findall("sheets/sheet").forEach(sheet => {
        const sheetId = sheet.attrib.sheetId;
        const relId = sheet.attrib['r:id'];
        const relationship = workbookRels.find("Relationship[@Id='" + relId + "']");
        const filename = prefix + "/" + relationship.attrib.Target;

        sheets.push({
            id: parseInt(sheetId, 10),
            name: sheet.attrib.name,
            filename: filename
        });
    });

    return sheets;
};

/**
     * Create a new workbook. Either pass the raw data of a .xlsx file,
     * or call `loadTemplate()` later.
     */
export default class Workbook {
    constructor(data) {

        this.archive = null;
        this.sharedStrings = [];
        this.sharedStringsLookup = {};
        this.sheets = [];
        this.allPlaceholders = [];
        this.readPlaceholders = false;
        if (data) {
            this.loadTemplate(data);
        }
    }
    /**
        * Delete unused sheets if needed
        */
    deleteSheet(sheetName) {
        // var self = this;
        const sheet = this.loadSheet(sheetName);
        const sh = this.workbook.find("sheets/sheet[@sheetId='" + sheet.id + "']");
        this.workbook.find("sheets").remove(sh);
        const rel = this.workbookRels.find("Relationship[@Id='" + sh.attrib['r:id'] + "']");
        this.workbookRels.remove(rel);
        this._rebuild();
        return this;
    }
    /**
        * Clone sheets in current workbook template
        */
    copySheet(sheetName, copyName) {
        const sheet = this.loadSheet(sheetName);
        const newSheetIndex = (this.workbook.findall("sheets/sheet").length + 1).toString();
        const fileName = 'worksheets' + '/' + 'sheet' + newSheetIndex + '.xml';
        const arcName = this.prefix + '/' + fileName;
        this.archive.file(arcName, etree.tostring(sheet.root));
        this.archive.files[arcName].options.binary = true;
        const newSheet = etree.SubElement(this.workbook.find('sheets'), 'sheet');
        newSheet.attrib.name = copyName || 'Sheet' + newSheetIndex;
        newSheet.attrib.sheetId = newSheetIndex;
        newSheet.attrib['r:id'] = 'rId' + newSheetIndex;
        const newRel = etree.SubElement(this.workbookRels, 'Relationship');
        newRel.attrib.Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
        newRel.attrib.Target = fileName;
        this._rebuild();
        //    TODO: work with "definedNames" 
        //    var defn = etree.SubElement(self.workbook.find('definedNames'), 'definedName');
        //
        return this;
    }
    /**
        *  Partially rebuild after copy/delete sheets
        */
    _rebuild() {
        //each <sheet> 'r:id' attribute in '\xl\workbook.xml'
        //must point to correct <Relationship> 'Id' in xl\_rels\workbook.xml.rels
        let self = this;
        const order = ['worksheet', 'theme', 'styles', 'sharedStrings'];
        self.workbookRels.findall("*")
            .sort(function (rel1, rel2) {
                const index1 = order.indexOf(path.basename(rel1.attrib.Type));
                const index2 = order.indexOf(path.basename(rel2.attrib.Type));
                if ((index1 + index2) == 0) {
                    if (rel1.attrib.Id && rel2.attrib.Id)
                        return rel1.attrib.Id.substring(3) - rel2.attrib.Id.substring(3);
                    return rel1._id - rel2._id;
                }
                return index1 - index2;
            })
            .forEach(function (item, index) {
                item.attrib.Id = 'rId' + (index + 1);
            });
        self.workbook.findall("sheets/sheet").forEach(function (item, index) {
            item.attrib['r:id'] = 'rId' + (index + 1);
            item.attrib.sheetId = (index + 1).toString();
        });
        self.archive.file(self.prefix + '/' + '_rels' + '/' + path.basename(self.workbookPath) + '.rels', etree.tostring(self.workbookRels));
        self.archive.file(self.workbookPath, etree.tostring(self.workbook));
        self.sheets = loadSheets(self.prefix, self.workbook, self.workbookRels);
    }
    /**
         * Load a .xlsx file from a byte array.
         */
    loadTemplate(data) {
        var self = this;
        if (Buffer.isBuffer(data)) {
            data = data.toString('binary');
        }
        const t = new zip(data, { base64: false, checkCRC32: true });
        // const t = await zip.loadAsync(data);
        this.archive = t;
        // Load relationships
        const rels = etree.parse(self.archive.file("_rels/.rels").asText()).getroot(), ;
        const workbookPath = rels.find("Relationship[@Type='" + DOCUMENT_RELATIONSHIP + "']").attrib.Target;

        self.workbookPath = workbookPath;
        self.prefix = path.dirname(workbookPath);
        self.workbook = etree.parse(self.archive.file(workbookPath).asText()).getroot();
        self.workbookRels = etree.parse(self.archive.file(self.prefix + "/" + '_rels' + "/" + path.basename(workbookPath) + '.rels').asText()).getroot();
        self.sheets = loadSheets(self.prefix, self.workbook, self.workbookRels);
        self.calChainRel = self.workbookRels.find("Relationship[@Type='" + CALC_CHAIN_RELATIONSHIP + "']");
        if (self.calChainRel) {
            self.calcChainPath = self.prefix + "/" + self.calChainRel.attrib.Target;
        }
        self.sharedStringsPath = self.prefix + "/" + self.workbookRels.find("Relationship[@Type='" + SHARED_STRINGS_RELATIONSHIP + "']").attrib.Target;
        self.sharedStrings = [];
        etree.parse(self.archive.file(self.sharedStringsPath).asText()).getroot().findall('si').forEach(function (si) {
            const t = { text: '' };
            si.findall('t').forEach(function (tmp) {
                t.text += tmp.text;
            });
            si.findall('r/t').forEach(function (tmp) {
                t.text += tmp.text;
            });
            self.sharedStrings.push(t.text);
            self.sharedStringsLookup[t.text] = self.sharedStrings.length - 1;
        });
    }


    // not finished
    clone() {
        const n = new Workbook();


        n.workbookPath = this.workbookPath; //immuit izi
        n.prefix = this.prefix; // immut izi


        self.workbook = etree.parse(self.archive.file(workbookPath).asText()).getroot();
        self.workbookRels = etree.parse(self.archive.file(self.prefix + "/" + '_rels' + "/" + path.basename(workbookPath) + '.rels').asText()).getroot();
        self.sheets = loadSheets(self.prefix, self.workbook, self.workbookRels);
        self.calChainRel = self.workbookRels.find("Relationship[@Type='" + CALC_CHAIN_RELATIONSHIP + "']");
        if (self.calChainRel) {
            self.calcChainPath = self.prefix + "/" + self.calChainRel.attrib.Target;
        }
        self.sharedStringsPath = self.prefix + "/" + self.workbookRels.find("Relationship[@Type='" + SHARED_STRINGS_RELATIONSHIP + "']").attrib.Target;
        self.sharedStrings = [];
        etree.parse(self.archive.file(self.sharedStringsPath).asText()).getroot().findall('si').forEach(function (si) {
            var t = { text: '' };
            si.findall('t').forEach(function (tmp) {
                t.text += tmp.text;
            });
            si.findall('r/t').forEach(function (tmp) {
                t.text += tmp.text;
            });
            self.sharedStrings.push(t.text);
            self.sharedStringsLookup[t.text] = self.sharedStrings.length - 1;
        });
    }


    /**
         * Interpolate values for the sheet with the given number (1-based) or
         * name (if a string) using the given substitutions (an object).
         */
    substitute(sheetName, substitutions) {
        var self = this;
        var sheet = self.loadSheet(sheetName);
        var dimension = sheet.root.find("dimension"), sheetData = sheet.root.find("sheetData"), currentRow = null, totalRowsInserted = 0, totalColumnsInserted = 0, namedTables = self.loadTables(sheet.root, sheet.filename), rows = [];
        sheetData.findall("row").forEach(function (row) {
            row.attrib.r = currentRow = getCurrentRow(row, totalRowsInserted);
            rows.push(row);
            var cells = [], cellsInserted = 0, newTableRows = [];
            row.findall("c").forEach(function (cell) {
                var appendCell = true;
                cell.attrib.r = self.getCurrentCell(cell, currentRow, cellsInserted);
                // If c[@t="s"] (string column), look up /c/v@text as integer in
                // `this.sharedStrings`
                if (cell.attrib.t === "s") {
                    // Look for a shared string that may contain placeholders
                    var cellValue = cell.find("v"), stringIndex = parseInt(cellValue.text, 10), string = self.sharedStrings[stringIndex];
                    if (string === undefined) {
                        return;
                    }
                    // Loop over placeholders
                    extractPlaceholders(string).forEach(function (placeholder) {
                        // Only substitute things for which we have a substitution
                        var substitution = _get(substitutions, placeholder.name, ''), newCellsInserted = 0;
                        if (placeholder.full && placeholder.type === "table" && substitution instanceof Array) {
                            newCellsInserted = self.substituteTable(row, newTableRows, cells, cell, namedTables, substitution, placeholder.key);
                            // don't double-insert cells
                            // this applies to arrays only, incorrectly applies to object arrays when there a single row, thus not rendering single row
                            if (newCellsInserted !== 0 || substitution.length) {
                                if (substitution.length === 1) {
                                    appendCell = true;
                                }
                                if (substitution[0][placeholder.key] instanceof Array) {
                                    appendCell = false;
                                }
                            }
                            // Did we insert new columns (array values)?
                            if (newCellsInserted !== 0) {
                                cellsInserted += newCellsInserted;
                                self.pushRight(self.workbook, sheet.root, cell.attrib.r, newCellsInserted);
                            }
                        }
                        else if (placeholder.full && placeholder.type === "normal" && substitution instanceof Array) {
                            appendCell = false; // don't double-insert cells
                            newCellsInserted = self.substituteArray(cells, cell, substitution);
                            if (newCellsInserted !== 0) {
                                cellsInserted += newCellsInserted;
                                self.pushRight(self.workbook, sheet.root, cell.attrib.r, newCellsInserted);
                            }
                        }
                        else {
                            if (placeholder.key) {
                                substitution = _get(substitutions, placeholder.name + '.' + placeholder.key);
                            }
                            string = self.substituteScalar(cell, string, placeholder, substitution);
                        }
                    });
                }
                // if we are inserting columns, we may not want to keep the original cell anymore
                if (appendCell) {
                    cells.push(cell);
                }
            }); // cells loop
            // We may have inserted columns, so re-build the children of the row
            replaceChildren(row, cells);
            // Update row spans attribute
            if (cellsInserted !== 0) {
                updateRowSpan(row, cellsInserted);
                if (cellsInserted > totalColumnsInserted) {
                    totalColumnsInserted = cellsInserted;
                }
            }
            // Add newly inserted rows
            if (newTableRows.length > 0) {
                newTableRows.forEach(function (row) {
                    rows.push(row);
                    ++totalRowsInserted;
                });
                self.pushDown(self.workbook, sheet.root, namedTables, currentRow, newTableRows.length);
            }
        }); // rows loop
        // We may have inserted rows, so re-build the children of the sheetData
        replaceChildren(sheetData, rows);
        // Update placeholders in table column headers
        self.substituteTableColumnHeaders(namedTables, substitutions);
        // Update placeholders in hyperlinks
        self.substituteHyperlinks(sheet.filename, substitutions);
        // Update <dimension /> if we added rows or columns
        if (dimension) {
            if (totalRowsInserted > 0 || totalColumnsInserted > 0) {
                var dimensionRange = splitRange(dimension.attrib.ref), dimensionEndRef = splitRef(dimensionRange.end);
                dimensionEndRef.row += totalRowsInserted;
                dimensionEndRef.col = numToColumnIdentifier(charToNum(dimensionEndRef.col) + totalColumnsInserted);
                dimensionRange.end = joinRef(dimensionEndRef);
                dimension.attrib.ref = joinRange(dimensionRange);
            }
        }
        //Here we are forcing the values in formulas to be recalculated
        // existing as well as just substituted
        sheetData.findall("row").forEach(function (row) {
            row.findall("c").forEach(function (cell) {
                var formulas = cell.findall('f');
                if (formulas && formulas.length > 0) {
                    cell.findall('v').forEach(function (v) {
                        cell.remove(v);
                    });
                }
            });
        });
        // Write back the modified XML trees
        self.archive.file(sheet.filename, etree.tostring(sheet.root));
        self.archive.file(self.workbookPath, etree.tostring(self.workbook));
        // Remove calc chain - Excel will re-build, and we may have moved some formulae
        if (self.calcChainPath && self.archive.file(self.calcChainPath)) {
            self.archive.remove(self.calcChainPath);
        }
        self.writeSharedStrings();
        self.writeTables(namedTables);
    }
    /**
         * Generate a new binary .xlsx file
         */
    generate(options) {
        if (!options) {
            options = {
                base64: false
            };
        }
        // console.log(this.archive.generate.toString());
        return this.archive.generate(options);
    }
    // Helpers
    // Write back the new shared strings list
    writeSharedStrings() {
        var self = this;
        var root = etree.parse(self.archive.file(self.sharedStringsPath).asText()).getroot(), children = root.getchildren();
        root.delSlice(0, children.length);
        self.sharedStrings.forEach(function (string) {
            var si = new etree.Element("si"), t = new etree.Element("t");
            t.text = string;
            si.append(t);
            root.append(si);
        });
        root.attrib.count = self.sharedStrings.length;
        root.attrib.uniqueCount = self.sharedStrings.length;
        self.archive.file(self.sharedStringsPath, etree.tostring(root));
    }
    // Add a new shared string
    addSharedString(s) {
        const idx = this.sharedStrings.length;
        this.sharedStrings.push(s);
        this.sharedStringsLookup[s] = idx;
        return idx;
    }
    // Get the number of a shared string, adding a new one if necessary.
    stringIndex(s) {
        var self = this;
        var idx = self.sharedStringsLookup[s];
        if (idx === undefined) {
            idx = self.addSharedString(s);
        }
        return idx;
    }
    // Replace a shared string with a new one at the same index. Return the
    // index.
    replaceString(oldString, newString) {
        // var self = this;
        var idx = this.sharedStringsLookup[oldString];
        if (idx === undefined) {
            idx = this.addSharedString(newString);
        }
        else {
            this.sharedStrings[idx] = newString;
            delete this.sharedStringsLookup[oldString];
            this.sharedStringsLookup[newString] = idx;
        }
        return idx;
    }
    // Get sheet a sheet, including filename and name
    loadSheet(sheet) {
        // var self = this;
        let info = null;
        for (var i = 0; i < this.sheets.length; ++i) {
            if ((typeof (sheet) === "number" && this.sheets[i].id === sheet) || (this.sheets[i].name === sheet)) {
                info = this.sheets[i];
                break;
            }
        }
        if (info === null && (typeof (sheet) === "number")) {
            //Get the sheet that corresponds to the 0 based index if the id does not work
            info = this.sheets[sheet - 1];
        }
        if (info === null) {
            throw new Error(`Sheet ${sheet} not found`);
        }
        return {
            filename: info.filename,
            name: info.name,
            id: info.id,
            root: etree.parse(this.archive.file(info.filename).asText()).getroot()
        };
    }
    // Load tables for a given sheet
    loadTables(sheet, sheetFilename) {
        const sheetDirectory = path.dirname(sheetFilename)
        const sheetName = path.basename(sheetFilename)
        const relsFilename = sheetDirectory + "/" + '_rels' + "/" + sheetName + '.rels'
        const relsFile = this.archive.file(relsFilename)
        let tables = [];
        if (relsFile === null) {
            return tables;
        }
        const rels = etree.parse(relsFile.asText()).getroot();
        sheet.findall("tableParts/tablePart").forEach(tablePart => {
            const relationshipId = tablePart.attrib['r:id']
            const target = rels.find("Relationship[@Id='" + relationshipId + "']").attrib.Target
            const tableFilename = target.replace('..', this.prefix)
            const tableTree = etree.parse(this.archive.file(tableFilename).asText());
            tables.push({
                filename: tableFilename,
                root: tableTree.getroot()
            });
        });
        return tables;
    }
    // Write back possibly-modified tables
    writeTables(tables) {
        tables.forEach(namedTable => {
            this.archive.file(namedTable.filename, etree.tostring(namedTable.root));
        });
    }
    //Perform substitution in hyperlinks
    substituteHyperlinks(sheetFilename, substitutions) {
        const sheetDirectory = path.dirname(sheetFilename);
        const sheetName = path.basename(sheetFilename);
        const relsFilename = sheetDirectory + "/" + '_rels' + "/" + sheetName + '.rels';
        const relsFile = this.archive.file(relsFilename);
        etree.parse(this.archive.file(this.sharedStringsPath).asText()).getroot();
        if (relsFile === null) {
            return;
        }
        const rels = etree.parse(relsFile.asText()).getroot();
        const relationships = rels._children;
        const newRelationships = [];
        relationships.forEach(relationship => {
            newRelationships.push(relationship);
            if (relationship.attrib.Type === HYPERLINK_RELATIONSHIP) {
                let target = relationship.attrib.Target;
                //Double-decode due to excel double encoding url placeholders
                target = decodeURI(decodeURI(target));
                extractPlaceholders(target).forEach(placeholder => {
                    const substitution = substitutions[placeholder.name];
                    if (substitution === undefined) {
                        return;
                    }
                    target = target.replace(placeholder.placeholder, this.stringify(substitution));
                    relationship.attrib.Target = encodeURI(target);
                });
            }
        });
        replaceChildren(rels, newRelationships);
        this.archive.file(relsFilename, etree.tostring(rels));
    }
    getPlaceholdersOneSheet(sheetName) {

        const placeholders = [];
        const sheet = this.loadSheet(sheetName);
        let cellsInserted = 0;
        let sheetData = sheet.root.find("sheetData")
        let currentRow = null
        let totalRowsInserted = 0
        let rows = [];
        sheetData.findall("row").forEach(row => {
            row.attrib.r = currentRow = getCurrentRow(row, totalRowsInserted);
            rows.push(row);
            row.findall("c").forEach(cell => {
                let appendCell = true;
                cell.attrib.r = this.getCurrentCell(cell, currentRow, cellsInserted);
                // If c[@t="s"] (string column), look up /c/v@text as integer in
                // `this.sharedStrings`
                if (cell.attrib.t === "s") {
                    // Look for a shared string that may contain placeholders
                    const cellValue = cell.find("v")
                    const stringIndex = parseInt(cellValue.text, 10);
                    const s = this.sharedStrings[stringIndex];
                    if (s === undefined) {
                        return;
                    }
                    // Loop over placeholders
                    const res = extractPlaceholders(s).map(p => p.placeholder.slice(2, -1));
                    if (res.length > 0) {
                        placeholders.push(...res);
                    }
                }
            });
        });
        return placeholders;
    }

    getAllPlaceholders() {
        if (this.readPlaceholders) return this.allPlaceholders;
        const placeholders = [];
        this.sheets
            .forEach(sheet => placeholders.push(...this.getPlaceholdersOneSheet(sheet.id)));
        this.readPlaceholders = true;
        this.allPlaceholders = placeholders;
        return this.allPlaceholders;
    }

    // Perform substitution in table headers
    substituteTableColumnHeaders(tables, substitutions) {
        tables.forEach(table => {
            const root = table.root, columns = root.find("tableColumns");
            let autoFilter = root.find("autoFilter");
            let tableRange = splitRange(root.attrib.ref);
            let idx = 0;
            let inserted = 0;
            let newColumns = [];
            columns.findall("tableColumn").forEach(col => {
                ++idx;
                col.attrib.id = Number(idx).toString();
                newColumns.push(col);
                const name = col.attrib.name;
                extractPlaceholders(name).forEach(placeholder => {
                    var substitution = substitutions[placeholder.name];
                    if (substitution === undefined) {
                        return;
                    }
                    // Array -> new columns
                    if (placeholder.full && placeholder.type === "normal" && substitution instanceof Array) {
                        substitution.forEach((element, i) => {
                            var newCol = col;
                            if (i > 0) {
                                newCol = this.cloneElement(newCol);
                                newCol.attrib.id = Number(++idx).toString();
                                newColumns.push(newCol);
                                ++inserted;
                                tableRange.end = this.nextCol(tableRange.end);
                            }
                            newCol.attrib.name = this.stringify(element);
                        });
                        // Normal placeholder
                    }
                    else {
                        name = name.replace(placeholder.placeholder, this.stringify(substitution));
                        col.attrib.name = name;
                    }
                });
            });
            replaceChildren(columns, newColumns);
            // Update range if we inserted columns
            if (inserted > 0) {
                columns.attrib.count = Number(idx).toString();
                root.attrib.ref = joinRange(tableRange);
                if (autoFilter !== null) {
                    // XXX: This is a simplification that may stomp on some configurations
                    autoFilter.attrib.ref = joinRange(tableRange);
                }
            }
            //update ranges for totalsRowCount
            const tableRoot = table.root;
            tableRange = splitRange(tableRoot.attrib.ref);
            tableStart = splitRef(tableRange.start);
            tableEnd = splitRef(tableRange.end);
            if (tableRoot.attrib.totalsRowCount) {
                autoFilter = tableRoot.find("autoFilter");
                if (autoFilter !== null) {
                    autoFilter.attrib.ref = joinRange({
                        start: joinRef(tableStart),
                        end: joinRef(tableEnd),
                    });
                }
                ++tableEnd.row;
                tableRoot.attrib.ref = joinRange({
                    start: joinRef(tableStart),
                    end: joinRef(tableEnd),
                });
            }
        });
    }
    // Return a list of tokens that may exist in the string.
    // Keys are: `placeholder` (the full placeholder, including the `${}`
    // delineators), `name` (the name part of the token), `key` (the object key
    // for `table` tokens), `full` (boolean indicating whether this placeholder
    // is the entirety of the string) and `type` (one of `table` or `cell`)
    // Get the next column's cell reference given a reference like "B2".
    nextCol(ref) {
        var self = this;
        ref = ref.toUpperCase();
        return ref.replace(/[A-Z]+/, function (match) {
            return numToColumnIdentifier(charToNum(match) + 1);
        });
    }
    // Is ref inside the table defined by startRef and endRef?
    isWithin(ref, startRef, endRef) {
        var self = this;
        var start = splitRef(startRef), end = splitRef(endRef), target = splitRef(ref);
        start.col = charToNum(start.col);
        end.col = charToNum(end.col);
        target.col = charToNum(target.col);
        return (start.row <= target.row && target.row <= end.row &&
            start.col <= target.col && target.col <= end.col);
    }
    // Turn a value of any type into a string
    stringify(value) {
        if (value instanceof Date) {
            //In Excel date is a number of days since 01/01/1900
            //           timestamp in ms    to days      + number of days from 1900 to 1970
            return Number((value.getTime() / (1000 * 60 * 60 * 24)) + 25569);
        }
        else if (typeof (value) === "number" || typeof (value) === "boolean") {
            return Number(value).toString();
        }
        else if (typeof (value) === "string") {
            return String(value).toString();
        }
        return "";
    }
    // Insert a substitution value into a cell (c tag)
    insertCellValue(cell, substitution) {
        var self = this;
        var cellValue = cell.find("v"), stringified = self.stringify(substitution);
        if (typeof substitution === 'string' && substitution[0] === '=') {
            //substitution, started with '=' is a formula substitution
            var formula = new etree.Element("f");
            formula.text = substitution.substr(1);
            cell.insert(1, formula);
            delete cell.attrib.t; //cellValue will be deleted later
            return formula.text;
        }
        if (typeof (substitution) === "number" || substitution instanceof Date) {
            delete cell.attrib.t;
            cellValue.text = stringified;
        }
        else if (typeof (substitution) === "boolean") {
            cell.attrib.t = "b";
            cellValue.text = stringified;
        }
        else {
            cell.attrib.t = "s";
            cellValue.text = Number(self.stringIndex(stringified)).toString();
        }
        return stringified;
    }
    // Perform substitution of a single value
    substituteScalar(cell, string, placeholder, substitution) {
        var self = this;
        if (placeholder.full) {
            return self.insertCellValue(cell, substitution);
        }
        else {
            var newString = string.replace(placeholder.placeholder, self.stringify(substitution));
            cell.attrib.t = "s";
            return self.insertCellValue(cell, newString);
        }
    }
    // Perform a columns substitution from an array
    substituteArray(cells, cell, substitution) {
        var self = this;
        var newCellsInserted = -1, // we technically delete one before we start adding back
            currentCell = cell.attrib.r;
        // add a cell for each element in the list
        substitution.forEach(function (element) {
            ++newCellsInserted;
            if (newCellsInserted > 0) {
                currentCell = self.nextCol(currentCell);
            }
            var newCell = self.cloneElement(cell);
            self.insertCellValue(newCell, element);
            newCell.attrib.r = currentCell;
            cells.push(newCell);
        });
        return newCellsInserted;
    }
    // Perform a table substitution. May update `newTableRows` and `cells` and change `cell`.
    // Returns total number of new cells inserted on the original row.
    substituteTable(row, newTableRows, cells, cell, namedTables, substitution, key) {
        var self = this, newCellsInserted = 0; // on the original row
        // if no elements, blank the cell, but don't delete it
        if (substitution.length === 0) {
            delete cell.attrib.t;
            replaceChildren(cell, []);
        }
        else {
            var parentTables = namedTables.filter(function (namedTable) {
                var range = splitRange(namedTable.root.attrib.ref);
                return self.isWithin(cell.attrib.r, range.start, range.end);
            });
            substitution.forEach(function (element, idx) {
                var newRow, newCell, newCellsInsertedOnNewRow = 0, newCells = [], value = _get(element, key, '');
                if (idx === 0) { // insert in the row where the placeholders are
                    if (value instanceof Array) {
                        newCellsInserted = self.substituteArray(cells, cell, value);
                    }
                    else {
                        self.insertCellValue(cell, value);
                    }
                }
                else { // insert new rows (or reuse rows just inserted)
                    // Do we have an existing row to use? If not, create one.
                    if ((idx - 1) < newTableRows.length) {
                        newRow = newTableRows[idx - 1];
                    }
                    else {
                        newRow = self.cloneElement(row, false);
                        newRow.attrib.r = getCurrentRow(row, newTableRows.length + 1);
                        newTableRows.push(newRow);
                    }
                    // Create a new cell
                    newCell = self.cloneElement(cell);
                    newCell.attrib.r = joinRef({
                        row: newRow.attrib.r,
                        col: splitRef(newCell.attrib.r).col
                    });
                    if (value instanceof Array) {
                        newCellsInsertedOnNewRow = self.substituteArray(newCells, newCell, value);
                        // Add each of the new cells created by substituteArray()
                        newCells.forEach(function (newCell) {
                            newRow.append(newCell);
                        });
                        updateRowSpan(newRow, newCellsInsertedOnNewRow);
                    }
                    else {
                        self.insertCellValue(newCell, value);
                        // Add the cell that previously held the placeholder
                        newRow.append(newCell);
                    }
                    // expand named table range if necessary
                    parentTables.forEach(function (namedTable) {
                        var tableRoot = namedTable.root, autoFilter = tableRoot.find("autoFilter"), range = splitRange(tableRoot.attrib.ref);
                        if (!self.isWithin(newCell.attrib.r, range.start, range.end)) {
                            range.end = nextRow(range.end);
                            tableRoot.attrib.ref = joinRange(range);
                            if (autoFilter !== null) {
                                // XXX: This is a simplification that may stomp on some configurations
                                autoFilter.attrib.ref = tableRoot.attrib.ref;
                            }
                        }
                    });
                }
            });
        }
        return newCellsInserted;
    }
    // Clone an element. If `deep` is true, recursively clone children
    cloneElement(element, deep) {
        var self = this;
        var newElement = etree.Element(element.tag, element.attrib);
        newElement.text = element.text;
        newElement.tail = element.tail;
        if (deep !== false) {
            element.getchildren().forEach(function (child) {
                newElement.append(self.cloneElement(child, deep));
            });
        }
        return newElement;
    }
    // Calculate the current cell based on asource cell, the current row index,
    // and a number of new cells that have been inserted so far
    getCurrentCell(cell, currentRow, cellsInserted) {
        var self = this;
        var colRef = splitRef(cell.attrib.r).col, colNum = charToNum(colRef);
        return joinRef({
            row: currentRow,
            col: numToColumnIdentifier(colNum + cellsInserted)
        });
    }
    // Split a range like "A1:B1" into {start: "A1", end: "B1"}
    // Join into a a range like "A1:B1" an object like {start: "A1", end: "B1"}
    // Look for any merged cell or named range definitions to the right of
    // `currentCell` and push right by `numCols`.
    pushRight(workbook, sheet, currentCell, numCols) {
        var self = this;
        var cellRef = splitRef(currentCell), currentRow = cellRef.row, currentCol = charToNum(cellRef.col);
        // Update merged cells on the same row, at a higher column
        sheet.findall("mergeCells/mergeCell").forEach(function (mergeCell) {
            var mergeRange = splitRange(mergeCell.attrib.ref), mergeStart = splitRef(mergeRange.start), mergeStartCol = charToNum(mergeStart.col), mergeEnd = splitRef(mergeRange.end), mergeEndCol = charToNum(mergeEnd.col);
            if (mergeStart.row === currentRow && currentCol < mergeStartCol) {
                mergeStart.col = numToColumnIdentifier(mergeStartCol + numCols);
                mergeEnd.col = numToColumnIdentifier(mergeEndCol + numCols);
                mergeCell.attrib.ref = joinRange({
                    start: joinRef(mergeStart),
                    end: joinRef(mergeEnd),
                });
            }
        });
        // Named cells/ranges
        workbook.findall("definedNames/definedName").forEach(function (name) {
            var ref = name.text;
            if (isRange(ref)) {
                var namedRange = splitRange(ref), namedStart = splitRef(namedRange.start), namedStartCol = charToNum(namedStart.col), namedEnd = splitRef(namedRange.end), namedEndCol = charToNum(namedEnd.col);
                if (namedStart.row === currentRow && currentCol < namedStartCol) {
                    namedStart.col = numToColumnIdentifier(namedStartCol + numCols);
                    namedEnd.col = numToColumnIdentifier(namedEndCol + numCols);
                    name.text = joinRange({
                        start: joinRef(namedStart),
                        end: joinRef(namedEnd),
                    });
                }
            }
            else {
                var namedRef = splitRef(ref), namedCol = charToNum(namedRef.col);
                if (namedRef.row === currentRow && currentCol < namedCol) {
                    namedRef.col = numToColumnIdentifier(namedCol + numCols);
                    name.text = joinRef(namedRef);
                }
            }
        });
    }
    // Look for any merged cell, named table or named range definitions below
    // `currentRow` and push down by `numRows` (used when rows are inserted).
    pushDown(workbook, sheet, tables, currentRow, numRows) {
        var self = this;
        var mergeCells = sheet.find("mergeCells");
        // Update merged cells below this row
        sheet.findall("mergeCells/mergeCell").forEach(function (mergeCell) {
            var mergeRange = splitRange(mergeCell.attrib.ref), mergeStart = splitRef(mergeRange.start), mergeEnd = splitRef(mergeRange.end);
            if (mergeStart.row > currentRow) {
                mergeStart.row += numRows;
                mergeEnd.row += numRows;
                mergeCell.attrib.ref = joinRange({
                    start: joinRef(mergeStart),
                    end: joinRef(mergeEnd),
                });
            }
            //add new merge cell
            if (mergeStart.row == currentRow) {
                for (var i = 1; i <= numRows; i++) {
                    var newMergeCell = self.cloneElement(mergeCell);
                    mergeStart.row += 1;
                    mergeEnd.row += 1;
                    newMergeCell.attrib.ref = joinRange({
                        start: joinRef(mergeStart),
                        end: joinRef(mergeEnd)
                    });
                    mergeCells.attrib.count += 1;
                    mergeCells._children.push(newMergeCell);
                }
            }
        });
        // Update named tables below this row
        tables.forEach(function (table) {
            var tableRoot = table.root, tableRange = splitRange(tableRoot.attrib.ref), tableStart = splitRef(tableRange.start), tableEnd = splitRef(tableRange.end);
            if (tableStart.row > currentRow) {
                tableStart.row += numRows;
                tableEnd.row += numRows;
                tableRoot.attrib.ref = joinRange({
                    start: joinRef(tableStart),
                    end: joinRef(tableEnd),
                });
                var autoFilter = tableRoot.find("autoFilter");
                if (autoFilter !== null) {
                    // XXX: This is a simplification that may stomp on some configurations
                    autoFilter.attrib.ref = tableRoot.attrib.ref;
                }
            }
        });
        // Named cells/ranges
        workbook.findall("definedNames/definedName").forEach(function (name) {
            var ref = name.text;
            if (isRange(ref)) {
                var namedRange = splitRange(ref), namedStart = splitRef(namedRange.start), namedEnd = splitRef(namedRange.end);
                if (namedStart) {
                    if (namedStart.row > currentRow) {
                        namedStart.row += numRows;
                        namedEnd.row += numRows;
                        name.text = joinRange({
                            start: joinRef(namedStart),
                            end: joinRef(namedEnd),
                        });
                    }
                }
            }
            else {
                var namedRef = splitRef(ref);
                if (namedRef.row > currentRow) {
                    namedRef.row += numRows;
                    name.text = joinRef(namedRef);
                }
            }
        });
    }
}


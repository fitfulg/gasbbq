// Models
class SheetModel {
    constructor(sheet) {
        this.sheet = sheet;
    }

    getSheet() {
        return this.sheet;
    }

    getRange(range) {
        return this.sheet.getRange(range);
    }

    setHeaders(headers) {
        headers.forEach((header, index) => {
            this.sheet.getRange(1, index + 1)
                .setValue(header)
                .setFontWeight('bold')
                .setBorder(true, true, true, true, true, true);
        });
    }

    setColumnWidths(columnWidths) {
        columnWidths.forEach((width, index) => {
            this.sheet.setColumnWidth(index + 1, width);
        });
    }
}


// Controllers
class SheetController {
    constructor(sheet) {
        this.sheetService = new SheetService(sheet);
        this.validationService = new ValidationService(sheet);
        this.protectionService = new ProtectionService(sheet);
        this.wordCountService = new WordCountService(sheet);
    }

    setupSheet() {
        this.sheetService.ensureRowCount(45);
        this.sheetService.setupHeaders();
        this.sheetService.setColumnWidths();
        this.sheetService.applyFormatting();
        this.sheetService.applyBackgroundColors();

        this.sheetService.applyTextColorToRange('F2:G45', COLORS.lightGray());

        this.validationService.applyConfirmationValidation();
        this.protectionService.protectColumns(['F2:F45', 'G2:G45']);

        this.wordCountService.countWords('C', 'F');
        this.wordCountService.countWords('D', 'G');
    }
}


class EventController {
    constructor(sheet) {
        this.sheetController = new SheetController(sheet);
    }

    onOpen() {
        this.sheetController.setupSheet();
    }

    onEdit(e) {
        const range = e.range;
        const sheet = e.source.getActiveSheet();

        if (range.getColumn() === 3 || range.getColumn() === 4) {
            const value = range.getValue().trim();
            if (value && value.indexOf(',') === -1 && value.indexOf(' ') !== -1) {
                range.setValue(value.replace(/\s+/g, '-'));
            }

            this.sheetController.wordCountService.countWords('C', 'F');
            this.sheetController.wordCountService.countWords('D', 'G');
        }
    }
}


// Services
class SheetService {
    constructor(sheet) {
        this.sheet = sheet;
    }

    ensureRowCount(count) {
        const currentRows = this.sheet.getMaxRows();
        if (currentRows > count) {
            this.sheet.deleteRows(count + 1, currentRows - count);
        } else if (currentRows < count) {
            this.sheet.insertRowsAfter(currentRows, count - currentRows);
        }
    }

    setupHeaders() {
        for (const config of COLUMN_CONFIG) {
            const columnIndex = this.sheet.getRange(config.column + '1').getColumn();
            this.sheet.getRange(1, columnIndex)
                .setValue(config.name)
                .setFontWeight('bold')
                .setBorder(true, true, true, true, true, true);
        }
        this.sheet.getRange('A1:E1').setFontColor(COLORS.white());
    }

    setColumnWidths() {
        for (const config of COLUMN_CONFIG) {
            const columnIndex = this.sheet.getRange(config.column + '1').getColumn();
            this.sheet.setColumnWidth(columnIndex, config.width);
        }
    }

    applyFormatting() {
        this.sheet.getRange('B1:G45').setWrap(true)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
        this.sheet.getRange('A1:A45').setWrap(true)
            .setHorizontalAlignment("left")
            .setVerticalAlignment("middle");
    }

    applyBackgroundColors() {
        this.sheet.getRange('A1:E1').setBackground(COLORS.darkGray());
        this.sheet.getRange('F1:G1').setBackground(COLORS.white()).setFontColor(COLORS.lightGray());
        this.sheet.getRange('A2:A45').setBackground(COLORS.lightGray());
        this.sheet.getRange('C2:C45').setBackground(COLORS.lightYellow());
        this.sheet.getRange('D2:D45').setBackground(COLORS.lightBlue());
    }

    applyTextColorToRange(range, color) {
        this.sheet.getRange(range).setFontColor(color);
    }
}



class ValidationService {
    constructor(sheet) {
        this.sheet = sheet;
    }

    applyConfirmationValidation() {
        const confirmRange = this.sheet.getRange('B2:B45');
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Sí', 'No'], true).build();
        confirmRange.setDataValidation(rule);
        this.sheet.getRange('B2:B45').setBorder(true, true, true, true, true, true);
    }
}


class ProtectionService {
    constructor(sheet) {
        this.sheet = sheet;
    }

    protectColumns(ranges) {
        ranges.forEach(range => {
            const protection = this.sheet.getRange(range).protect();
            protection.setDescription('Automatic count protection');
            protection.removeEditors(protection.getEditors());
            if (protection.canDomainEdit()) {
                protection.setDomainEdit(false);
            }
        });
    }
}


class WordCountService {
    constructor(sheet) {
        this.sheet = sheet;
    }

    countWords(sourceColumn, targetColumn) {
        const dataRange = this.sheet.getRange(`${sourceColumn}2:${sourceColumn}45`);
        const dataValues = dataRange.getValues().flat();

        const wordCount = dataValues.reduce((count, value) => {
            if (value) {
                const words = value.toString().toLowerCase().split(/[\s,]+/);
                words.forEach(word => {
                    count[word] = (count[word] || 0) + 1;
                });
            }
            return count;
        }, {});

        const sortedWordCount = Object.entries(wordCount).sort(([a], [b]) => a.localeCompare(b));
        const resultValues = sortedWordCount.map(([word, count]) => [`${word}: ${count}`]);

        const resultRange = this.sheet.getRange(`${targetColumn}2:${targetColumn}${sortedWordCount.length + 1}`);
        resultRange.clearContent();
        resultRange.setValues(resultValues);  // Aquí se pasa un array 2D correctamente
    }
}


// Utils
const COLORS = {
    darkGray: () => '#4d4d4d',
    lightGray: () => '#d9d9d9',
    white: () => '#ffffff',
    lightYellow: () => '#ffffe6',
    lightBlue: () => '#e6f2ff',
};

const COLUMN_CONFIG = [
    { column: 'A', name: 'Nom', width: 150 },
    { column: 'B', name: 'Confirmació', width: 150 },
    { column: 'C', name: 'Preferència menjars', width: 300 },
    { column: 'D', name: 'Preferència begudes', width: 300 },
    { column: 'E', name: 'Al·lèrgies', width: 100 },
    { column: 'F', name: 'C-counter (no editar)', width: 200 },
    { column: 'G', name: 'D-counter (no editar)', width: 200 }
];




// Main Script
function onOpen() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const eventController = new EventController(sheet);
    eventController.onOpen();
}

function onEdit(e) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const eventController = new EventController(sheet);
    eventController.onEdit(e);
}



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
        const headers = ['Nom', 'Confirmació', 'Preferència menjars', 'Preferència begudes', 'Al·lèrgies', 'C-counter (no editar)', 'D-counter (no editar)'];
        headers.forEach((header, index) => {
            this.sheet.getRange(1, index + 1)
                .setValue(header)
                .setFontWeight('bold')
                .setBorder(true, true, true, true, true, true);
        });
    }

    setColumnWidths() {
        const columnWidths = [150, 150, 300, 300, 100, 200, 200];
        columnWidths.forEach((width, index) => {
            this.sheet.setColumnWidth(index + 1, width);
        });
    }

    applyFormatting() {
        this.sheet.getRange('B1:G45').setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle");
        this.sheet.getRange('A1:A45').setWrap(true).setHorizontalAlignment("left").setVerticalAlignment("middle");
    }

    applyBackgroundColors() {
        this.sheet.getRange('A1:E1').setBackground(ColorUtils.darkGray());
        this.sheet.getRange('F1:G1').setBackground(ColorUtils.white()).setFontColor(ColorUtils.lightGray());
        this.sheet.getRange('A2:A45').setBackground(ColorUtils.lightGray());
        this.sheet.getRange('C2:C45').setBackground(ColorUtils.lightYellow());
        this.sheet.getRange('D2:D45').setBackground(ColorUtils.lightBlue());
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
const ColorUtils = {
    darkGray: () => '#4d4d4d',
    lightGray: () => '#d9d9d9',
    white: () => '#ffffff',
    lightYellow: () => '#ffffe6',
    lightBlue: () => '#e6f2ff',
};



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



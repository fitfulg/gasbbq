// Models


// Controllers
class SheetController {
    constructor(sheet) {
        Logger.log('SheetController constructor called');
        this.sheetService = new SheetService(sheet);
        this.validationService = new ValidationService(sheet);
        this.protectionService = new ProtectionService(sheet);
        this.wordCountService = new WordCountService(sheet);
    }

    setupSheet() {
        Logger.log('setupSheet called');
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
        Logger.log('EventController constructor called');
        this.sheetController = new SheetController(sheet);
    }

    onOpen() {
        Logger.log('onOpen called');
        this.sheetController.setupSheet();
    }

    onEdit(e) {
        Logger.log('onEdit called');
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
        Logger.log('SheetService constructor called');
        this.sheet = sheet;
    }

    ensureRowCount(count) {
        Logger.log('ensureRowCount called with count: ' + count);
        const currentRows = this.sheet.getMaxRows();
        if (currentRows > count) {
            this.sheet.deleteRows(count + 1, currentRows - count);
        } else if (currentRows < count) {
            this.sheet.insertRowsAfter(currentRows, count - currentRows);
        }
    }

    setupHeaders() {
        Logger.log('setupHeaders called');
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
        Logger.log('setColumnWidths called');
        for (const config of COLUMN_CONFIG) {
            const columnIndex = this.sheet.getRange(config.column + '1').getColumn();
            this.sheet.setColumnWidth(columnIndex, config.width);
        }
    }

    applyFormatting() {
        Logger.log('applyFormatting called');
        this.sheet.getRange('B1:G45').setWrap(true)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
        this.sheet.getRange('A1:A45').setWrap(true)
            .setHorizontalAlignment("left")
            .setVerticalAlignment("middle");
    }

    applyBackgroundColors() {
        Logger.log('applyBackgroundColors called');
        this.sheet.getRange('A1:E1').setBackground(COLORS.darkGray());
        this.sheet.getRange('F1:G1').setBackground(COLORS.white()).setFontColor(COLORS.lightGray());
        this.sheet.getRange('A2:A45').setBackground(COLORS.lightGray());
        this.sheet.getRange('C2:C45').setBackground(COLORS.lightYellow());
        this.sheet.getRange('D2:D45').setBackground(COLORS.lightBlue());
    }

    applyTextColorToRange(range, color) {
        Logger.log('applyTextColorToRange called with range: ' + range + ', color: ' + color);
        this.sheet.getRange(range).setFontColor(color);
    }
}



class ValidationService {
    constructor(sheet) {
        Logger.log('ValidationService constructor called');
        this.sheet = sheet;
    }

    applyConfirmationValidation() {
        Logger.log('applyConfirmationValidation called');
        const confirmRange = this.sheet.getRange('B2:B45');
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Sí', 'No'], true).build();
        confirmRange.setDataValidation(rule);
        this.sheet.getRange('B2:B45').setBorder(true, true, true, true, true, true);
    }
}


class ProtectionService {
    constructor(sheet) {
        Logger.log('ProtectionService constructor called');
        this.sheet = sheet;
    }

    protectColumns(ranges) {
        Logger.log('protectColumns called with ranges: ' + ranges);
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
        Logger.log('WordCountService constructor called');
        this.sheet = sheet;
    }

    countWords(sourceColumn, targetColumn) {
        Logger.log('countWords called with sourceColumn: ' + sourceColumn + ', targetColumn: ' + targetColumn);
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
    Logger.log('onOpen called');
    Logger.log("Hello World");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const eventController = new EventController(sheet);
    eventController.onOpen();
}

function onEdit(e) {
    Logger.log('onEdit called');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const eventController = new EventController(sheet);
    eventController.onEdit(e);
}



// Controllers
class SheetController {
    constructor(sheet, languageService) {
        Logger.log('SheetController constructor called');
        this.sheetService = new SheetService(sheet, languageService);
        this.dropdownService = new DropdownService(sheet);
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
        this.sheetService.applyTextColorToRange('A1:E1', COLORS.white());
        this.sheetService.applyTextColorToRange('F2:G45', COLORS.lightGray());

        this.dropdownService.applyConfirmationValidation();

        this.protectionService.protectColumns(['F2:F45', 'G2:G45']);

        this.wordCountService.countWords('C', 'F');
        this.wordCountService.countWords('D', 'G');
    }
}


class EventController {
    constructor(sheet) {
        Logger.log('EventController constructor called');
        this.languageService = new LanguageService(sheet);
        this.sheetController = new SheetController(sheet, this.languageService); // Inject LanguageService into SheetController
        this.menuService = new MenuService(this.languageService); // Inject LanguageService into MenuService
    }

    /**
     * Handles the onOpen event to set up the sheet and menu.
     */
    onOpen() {
        Logger.log('onOpen called');
        this.sheetController.setupSheet();
        this.menuService.createMenu(this); // Create the language change menu
    }

    /**
     * Handles the onEdit event to update word counts when cells are edited.
     * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object.
     */
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

    /**
     * Changes the language when the menu item is selected.
     * @param {string} languageCode - The code of the selected language.
     */
    changeLanguage(languageCode) {
        Logger.log(`EventController: changeLanguage to ${languageCode}`);
        this.languageService.changeLanguage(languageCode);
        this.sheetController.setupSheet(); // Reapply headers with the new language
    }
}


// Services
class SheetService {
    constructor(sheet, languageService) {
        Logger.log('SheetService constructor called');
        this.sheet = sheet;
        this.languageService = languageService;
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
        const headers = this.languageService.getHeaders();
        headers.forEach((header, index) => {
            this.sheet.getRange(1, index + 1)
                .setValue(header)
                .setFontWeight('bold')
                .setBorder(true, true, true, true, true, true);
        });
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



class DropdownService {
    constructor(sheet) {
        Logger.log('DropdownService  constructor called');
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


class LanguageService {
    constructor(sheet) {
        Logger.log('LanguageService constructor called');
        this.sheet = sheet;
        this.currentLanguage = 'ca'; // Default to English
    }

    /**
     * Changes the language of the sheet based on the selected language code.
     * @param {string} languageCode - The code of the language to switch to.
     */
    changeLanguage(languageCode) {
        Logger.log(`Changing language to: ${languageCode}`);
        const selectedLanguage = LANGUAGES.find(lang => lang.code === languageCode);

        if (selectedLanguage) {
            const headers = selectedLanguage.headers;
            for (let i = 0; i < headers.length; i++) {
                this.sheet.getRange(1, i + 1).setValue(headers[i]);
            }
            this.currentLanguage = languageCode;
        } else {
            Logger.log(`Language code: ${languageCode} not found.`);
        }
        SpreadsheetApp.flush();// Force changes to be written to the sheet
    }

    getHeaders() {
        const selectedLanguage = LANGUAGES.find(lang => lang.code === this.currentLanguage);
        return selectedLanguage ? selectedLanguage.headers : [];
    }

    /**
     * Retrieves the current language code.
     * @returns {string} - The current language code.
     */
    getCurrentLanguage() {
        return this.currentLanguage;
    }

    /**
     * Returns the name of the 'Language' menu based on the current language.
     * @returns {string} - The localized name for the menu.
     */
    getMenuName() {
        const currentLanguage = LANGUAGES.find(lang => lang.code === this.currentLanguage);
        return currentLanguage ? currentLanguage.menuName : 'Language';
    }
}

class MenuService {
    constructor(languageService) {
        Logger.log('MenuService constructor called');
        this.languageService = languageService; // Inject LanguageService dependency
    }

    /**
    * Creates a custom menu in the Google Sheets UI to change languages.
    * @param {EventController} eventController - The event controller to handle menu actions.
    */
    createMenu(eventController) {
        const ui = SpreadsheetApp.getUi();
        const menuName = this.languageService.getMenuName();
        const menu = ui.createMenu(menuName);

        // Add language options to the menu, binding them to the common function
        LANGUAGES.forEach(language => {
            menu.addItem(language.name, `changeLanguage_${language.code}`);
        });

        menu.addToUi();
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

const LANGUAGES = [
    {
        code: 'en',
        name: 'English',
        menuName: 'Language',
        headers: ['Name', 'Confirmation', 'Food Preference', 'Drink Preference', 'Allergies', 'C-counter (do not edit)', 'D-counter (do not edit)']
    },
    {
        code: 'es',
        name: 'Castellano',
        menuName: 'Idioma',
        headers: ['Nombre', 'Confirmación', 'Preferencia de Comida', 'Preferencia de Bebida', 'Alergias', 'C-contador (no editar)', 'D-contador (no editar)']
    },
    {
        code: 'ca',
        name: 'Català',
        menuName: 'Idioma',
        headers: ['Nom', 'Confirmació', 'Preferència menjars', 'Preferència begudes', 'Al·lèrgies', 'C-counter (no editar)', 'D-counter (no editar)']
    }
];



// Main Script
function onOpen() {
    Logger.log('onOpen called');
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

function onChangeLanguage(languageCode) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const eventController = new EventController(sheet);
    eventController.changeLanguage(languageCode);
}

const changeLanguage_en = () => onChangeLanguage('en');
const changeLanguage_es = () => onChangeLanguage('es');
const changeLanguage_ca = () => onChangeLanguage('ca');




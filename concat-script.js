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
        this.sheetController.sheetService.applyCompletionFormatting();
        this.menuService.createLanguageMenu(this);
        this.menuService.createDeveloperMenu();
    }

    /**
     * Handles the onEdit event to update word counts when cells are edited.
     * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object.
     */
    onEdit(e) {
        Logger.log('onEdit called');
        const range = e.range;
        const sheet = e.source.getActiveSheet();

        if (range.getColumn() >= 1 && range.getColumn() <= 5) {
            this.sheetController.sheetService.applyCompletionFormatting(); // Check for completion on edit
        }

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
    * Changes the language when the menu item is selected and shows an alert to reload the page.
    * @param {string} languageCode - The code of the selected language.
    */
    changeLanguage(languageCode) {
        Logger.log(`EventController: changeLanguage to ${languageCode}`);
        this.languageService.changeLanguage(languageCode);

        const messages = this.languageService.getAlertMessages();
        const ui = SpreadsheetApp.getUi();
        ui.alert(messages.languageChanged, messages.reloadPage, ui.ButtonSet.OK);
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

    /**
    * Applies the default background colors to the entire sheet.
    */
    applyBackgroundColors() {
        Logger.log('applyBackgroundColors called');
        this.sheet.getRange('A1:E1').setBackground(COLORS.darkGray());
        this.sheet.getRange('F1:G1').setBackground(COLORS.white()).setFontColor(COLORS.lightGray());

        for (let row = 2; row <= 45; row++) {
            this.applyRowBackground(row); // Reuse the same logic for default backgrounds
        }
    }

    /**
    * Applies the background colors for a specific ROW based on the default layout.
    * @param {number} row - The row number where default backgrounds should be reapplied.
    */
    applyRowBackground(row) {
        Logger.log(`applyRowBackground called for row: ${row}`);
        this.sheet.getRange(row, 1).setBackground(COLORS.lightGray());
        this.sheet.getRange(row, 2).setBackground(COLORS.lightGray());
        this.sheet.getRange(row, 3).setBackground(COLORS.lightYellow());
        this.sheet.getRange(row, 4).setBackground(COLORS.lightBlue());
        this.sheet.getRange(row, 5).setBackground(COLORS.lightGray());
    }

    applyTextColorToRange(range, color) {
        Logger.log('applyTextColorToRange called with range: ' + range + ', color: ' + color);
        this.sheet.getRange(range).setFontColor(color);
    }

    /**
     * Checks for completed rows (A to E) and applies light green background to completed rows.
     * Incomplete rows maintain their default background colors.
     */
    applyCompletionFormatting() {
        Logger.log('applyCompletionFormatting called');
        const range = this.sheet.getRange('A2:E45'); // Range of rows to check (A2 to E45)
        const values = range.getValues();

        for (let i = 0; i < values.length; i++) {
            const rowValues = values[i];
            const isRowComplete = rowValues.every(cell => cell !== '');

            const rowRange = this.sheet.getRange(i + 2, 1, 1, 5);

            if (isRowComplete) {
                rowRange.setBackground(COLORS.lightGreen());
            } else {
                this.applyRowBackground(i + 2);
            }
        }
    }
}



class DropdownService {
    constructor(sheet) {
        Logger.log('DropdownService  constructor called');
        this.sheet = sheet;
    }

    /**
     * Apply dropdown validation to the range
     */
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
        this.currentLanguage = 'ca'; // Default language
        this.currentLanguage = this.getStoredLanguage() || this.defaultLanguage;
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
            this.storeLanguage(languageCode);
        } else {
            Logger.log(`Language code: ${languageCode} not found.`);
        }
        SpreadsheetApp.flush();// Force changes to be written to the sheet
    }

    /**
    * Stores the selected language in PropertiesService.
    * @param {string} languageCode - The language code to store.
    */
    storeLanguage(languageCode) {
        const userProperties = PropertiesService.getUserProperties();
        userProperties.setProperty('SELECTED_LANGUAGE', languageCode);
    }

    /**
     * Retrieves the stored language from PropertiesService.
     * @returns {string|null} - The stored language code or null if not set.
     */
    getStoredLanguage() {
        const userProperties = PropertiesService.getUserProperties();
        return userProperties.getProperty('SELECTED_LANGUAGE');
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

    /**
    * Gets the localized alert messages for the current language.
    * @returns {object} - The messages for alerts in the selected language.
    */
    getAlertMessages() {
        const selectedLanguage = LANGUAGES.find(lang => lang.code === this.currentLanguage);
        return selectedLanguage ? selectedLanguage.messages : { languageChanged: 'Language changed', reloadPage: 'Please reload the page to apply the changes.' };
    }
}

class MenuService {
    constructor(languageService) {
        Logger.log('MenuService constructor called');
        this.languageService = languageService; // Inject LanguageService dependency
    }

    /**
    * Creates the language menu, removing any old menu when the language changes.
    * @param {EventController} eventController - Controller to handle the menu actions.
    * @param {boolean} isLanguageChange - Indicates if the language is being changed to remove the previous menu.
    */
    createLanguageMenu(eventController) {
        Logger.log('Create menu called');
        const ui = SpreadsheetApp.getUi();
        ui.createMenu(this.languageService.getMenuName())
            .addItem('English', 'changeLanguage_en')
            .addItem('Castellano', 'changeLanguage_es')
            .addItem('Català', 'changeLanguage_ca')
            .addToUi();
    }

    /**
    * Creates a separate developer menu for handling properties.
    * @param {EventController} eventController - Controller to handle the menu actions.
    */
    createDeveloperMenu() {
        const ui = SpreadsheetApp.getUi();
        ui.createMenu('Developer')
            .addItem('GAS Console: List All Properties', 'listAllProperties')    // Action to list properties
            .addItem('GAS Console: Delete Property', 'promptDeleteProperty')     // Option to delete a property
            .addItem('GAS Console: Clear All Properties', 'clearAllProperties')  // Option to clear all properties
            .addToUi();
    }
}

class SheetPropertiesService {
    /**
     * Retrieves all properties stored in UserProperties.
     * @returns {Object} - An object containing all properties and their values.
     */
    static listProperties() {
        const userProperties = PropertiesService.getUserProperties();
        const properties = userProperties.getProperties();
        Logger.log('Listing all properties:');
        for (let key in properties) {
            Logger.log(`${key}: ${properties[key]}`);
        }
        return properties;
    }

    /**
     * Retrieves a specific property by key.
     * @param {string} key - The key of the property to retrieve.
     * @returns {string|null} - The value of the property or null if not found.
     */
    static getProperty(key) {
        const userProperties = PropertiesService.getUserProperties();
        return userProperties.getProperty(key);
    }

    /**
     * Deletes a specific property by key.
     * @param {string} key - The key of the property to delete.
     */
    static deleteProperty(key) {
        const userProperties = PropertiesService.getUserProperties();
        userProperties.deleteProperty(key);
        Logger.log(`Deleted property: ${key}`);
    }

    /**
     * Clears all properties.
     */
    static clearAllProperties() {
        const userProperties = PropertiesService.getUserProperties();
        userProperties.deleteAllProperties();
        Logger.log('All properties deleted.');
    }
}



// Utils
const COLORS = {
    darkGray: () => '#4d4d4d',
    lightGray: () => '#d9d9d9',
    white: () => '#ffffff',
    lightYellow: () => '#ffffe6',
    lightBlue: () => '#e6f2ff',
    lightGreen: () => '#e6ffe6'
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
        headers: ['Name', 'Confirmation', 'Food Preference', 'Drink Preference', 'Allergies', 'C-counter (do not edit)', 'D-counter (do not edit)'],
        messages: {
            languageChanged: 'Language changed',
            reloadPage: 'Please reload the page to apply the changes.'
        }
    },
    {
        code: 'es',
        name: 'Castellano',
        menuName: 'Idioma',
        headers: ['Nombre', 'Confirmación', 'Preferencia de Comida', 'Preferencia de Bebida', 'Alergias', 'C-contador (no editar)', 'D-contador (no editar)'],
        messages: {
            languageChanged: 'Idioma cambiado',
            reloadPage: 'Por favor, recargue la página para aplicar los cambios.'
        }
    },
    {
        code: 'ca',
        name: 'Català',
        menuName: 'Idioma',
        headers: ['Nom', 'Confirmació', 'Preferència menjars', 'Preferència begudes', 'Al·lèrgies', 'C-counter (no editar)', 'D-counter (no editar)'],
        messages: {
            languageChanged: 'Idioma canviat',
            reloadPage: 'Si us plau, recarregui la pàgina per aplicar els canvis.'
        }
    }
];




// Triggers
function handleEvent(callback, ...args) {
    Logger.log(`${callback.name} called`);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const eventController = new EventController(sheet);
    callback.apply(eventController, args);
}

function onOpen() {
    handleEvent(EventController.prototype.onOpen);
}

function onEdit(e) {
    handleEvent(EventController.prototype.onEdit, e);
}

function onChangeLanguage(languageCode) {
    handleEvent(EventController.prototype.changeLanguage, languageCode);
}


// Actions

const changeLanguage_en = () => onChangeLanguage('en');
const changeLanguage_es = () => onChangeLanguage('es');
const changeLanguage_ca = () => onChangeLanguage('ca');

/**
 * List all properties and their values.
 * @returns {Object} - An object containing all properties and their values.
 */
function listAllProperties() {
    Logger.log('ListAllProperties called. Listing all properties:');
    const properties = SheetPropertiesService.listProperties();
    return properties;
}

/**
 * Deletes a specific property by key.
 * @param {string} key - The key of the property to delete.
 */
function deleteProperty(key) {
    Logger.log(`Attempting to delete property: ${key}`);
    SheetPropertiesService.deleteProperty(key);
    SpreadsheetApp.getUi().alert(`Property ${key} has been deleted.`);
}

/**
 * Deletes all properties.
 */
function clearAllProperties() {
    Logger.log('Attempting to delete all properties.');
    SheetPropertiesService.clearAllProperties();
    SpreadsheetApp.getUi().alert('All properties have been deleted.');
}


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

            this.sheetController.wordCountService.countWords('C', 'H');
            this.sheetController.wordCountService.countWords('D', 'I');
        }
    }

    /**
    * Handles the change of language from the menu.
    * @param {string} languageCode - The code of the selected language.
    */
    handleLanguageChange(languageCode) {
        Logger.log(`EventController: handleLanguageChange to ${languageCode}`);
        this.sheetController.sheetService.clearRange(['F1', 'G1']); //(TEMPORARY SOLUTION)
        this.languageService.changeLanguage(languageCode);

        this.sheetController.dropdownService.updateDropdownValues();
        this.sheetController.dropdownService.applyConfirmationValidation();

        this.sheetController.sheetService.setupHeaders();

        const messages = this.languageService.getAlertMessages();
        const ui = SpreadsheetApp.getUi();
        ui.alert(messages.languageChanged, messages.reloadPage, ui.ButtonSet.OK);
    }

}
// module.exports = { EventController };
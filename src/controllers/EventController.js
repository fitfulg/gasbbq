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
// module.exports = { EventController };
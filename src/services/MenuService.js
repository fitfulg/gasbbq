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
    createMenu(eventController) {
        Logger.log('Create menu called');
        const ui = SpreadsheetApp.getUi();
        ui.createMenu(this.languageService.getMenuName())
            .addItem('English', 'changeLanguage_en')
            .addItem('Castellano', 'changeLanguage_es')
            .addItem('Català', 'changeLanguage_ca')
            .addToUi();
    }
}
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
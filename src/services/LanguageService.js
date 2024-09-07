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

    /**
    * Returns the localized dropdown options based on the current language.
    * @returns {Array<string>} - The dropdown options in the selected language.
    */
    getDropdownOptions() {
        const selectedLanguage = LANGUAGES.find(lang => lang.code === this.currentLanguage);
        return selectedLanguage ? selectedLanguage.dropdownOptions : ['Sí', 'No', 'NS/NR'];
    }
}
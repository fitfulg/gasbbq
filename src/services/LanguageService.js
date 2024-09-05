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
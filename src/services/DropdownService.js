class DropdownService {
    constructor(sheet, languageService) {
        Logger.log('DropdownService  constructor called');
        this.sheet = sheet;
        this.languageService = languageService;
    }

    /**
     * Apply dropdown validation to the range
     */
    applyConfirmationValidation() {
        Logger.log('applyConfirmationValidation called');
        const confirmRange = this.sheet.getRange('B2:B45');
        const dropdownOptions = this.languageService.getDropdownOptions(); // Obtener opciones traducidas

        const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(dropdownOptions, true)
            .build();

        confirmRange.setDataValidation(rule);
        this.sheet.getRange('B2:B45').setBorder(true, true, true, true, true, true);
    }
}
// module.exports = { DropdownService  };
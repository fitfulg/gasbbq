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
// module.exports = { DropdownService  };
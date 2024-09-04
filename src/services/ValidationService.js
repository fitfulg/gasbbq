class ValidationService {
    constructor(sheet) {
        this.sheet = sheet;
    }

    applyConfirmationValidation() {
        const confirmRange = this.sheet.getRange('B2:B45');
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Sí', 'No'], true).build();
        confirmRange.setDataValidation(rule);
        this.sheet.getRange('B2:B45').setBorder(true, true, true, true, true, true);
    }
}
module.exports = { ValidationService };
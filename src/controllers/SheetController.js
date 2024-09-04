class SheetController {
    constructor(sheet) {
        this.sheetService = new SheetService(sheet);
        this.validationService = new ValidationService(sheet);
        this.protectionService = new ProtectionService(sheet);
        this.wordCountService = new WordCountService(sheet);
    }

    setupSheet() {
        this.sheetService.ensureRowCount(45);
        this.sheetService.setupHeaders();
        this.sheetService.setColumnWidths();
        this.sheetService.applyFormatting();
        this.sheetService.applyBackgroundColors();

        this.validationService.applyConfirmationValidation();
        this.protectionService.protectColumns(['F2:F45', 'G2:G45']);

        this.wordCountService.countWords('C', 'F');
        this.wordCountService.countWords('D', 'G');
    }
}
module.exports = { SheetController };
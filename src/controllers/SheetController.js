class SheetController {
    constructor(sheet) {
        Logger.log('SheetController constructor called');
        this.sheetService = new SheetService(sheet);
        this.dropdownService = new DropdownService(sheet);
        this.protectionService = new ProtectionService(sheet);
        this.wordCountService = new WordCountService(sheet);
    }

    setupSheet() {
        Logger.log('setupSheet called');
        this.sheetService.ensureRowCount(45);
        this.sheetService.setupHeaders();
        this.sheetService.setColumnWidths();
        this.sheetService.applyFormatting();
        this.sheetService.applyBackgroundColors();
        this.sheetService.applyTextColorToRange('A1:E1', COLORS.white());
        this.sheetService.applyTextColorToRange('F2:G45', COLORS.lightGray());

        this.dropdownService.applyConfirmationValidation();

        this.protectionService.protectColumns(['F2:F45', 'G2:G45']);

        this.wordCountService.countWords('C', 'F');
        this.wordCountService.countWords('D', 'G');
    }
}
// module.exports = { SheetController };
class SheetController {
    constructor(sheet, languageService) {
        Logger.log('SheetController constructor called');
        this.sheetService = new SheetService(sheet, languageService);
        this.dropdownService = new DropdownService(sheet, languageService);
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
        this.sheetService.applyTextColorToRange('H2:I45', COLORS.lightGray());

        this.dropdownService.applyConfirmationValidation();

        this.protectionService.protectColumns(['H2:H45', 'I2:I45']);

        this.wordCountService.countWords('C', 'H');
        this.wordCountService.countWords('D', 'I');
    }
}
// module.exports = { SheetController };
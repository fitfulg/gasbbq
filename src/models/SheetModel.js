class SheetModel {
    constructor(sheet) {
        Logger.log('SheetModel constructor called');
        this.sheet = sheet;
    }

    getSheet() {
        Logger.log('getSheet called');
        return this.sheet;
    }

    getRange(range) {
        Logger.log('getRange called with range: ' + range);
        return this.sheet.getRange(range);
    }

    setHeaders(headers) {
        Logger.log('setHeaders called');
        headers.forEach((header, index) => {
            this.sheet.getRange(1, index + 1)
                .setValue(header)
                .setFontWeight('bold')
                .setBorder(true, true, true, true, true, true);
        });
    }

    setColumnWidths(columnWidths) {
        Logger.log('setColumnWidths called');
        columnWidths.forEach((width, index) => {
            this.sheet.setColumnWidth(index + 1, width);
        });
    }
}
module.exports = { SheetModel };
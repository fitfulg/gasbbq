class SheetModel {
    constructor(sheet) {
        this.sheet = sheet;
    }

    getSheet() {
        return this.sheet;
    }

    getRange(range) {
        return this.sheet.getRange(range);
    }

    setHeaders(headers) {
        headers.forEach((header, index) => {
            this.sheet.getRange(1, index + 1)
                .setValue(header)
                .setFontWeight('bold')
                .setBorder(true, true, true, true, true, true);
        });
    }

    setColumnWidths(columnWidths) {
        columnWidths.forEach((width, index) => {
            this.sheet.setColumnWidth(index + 1, width);
        });
    }
}
module.exports = { SheetModel };
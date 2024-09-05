class SheetService {
    constructor(sheet, languageService) {
        Logger.log('SheetService constructor called');
        this.sheet = sheet;
        this.languageService = languageService;
    }

    ensureRowCount(count) {
        Logger.log('ensureRowCount called with count: ' + count);
        const currentRows = this.sheet.getMaxRows();
        if (currentRows > count) {
            this.sheet.deleteRows(count + 1, currentRows - count);
        } else if (currentRows < count) {
            this.sheet.insertRowsAfter(currentRows, count - currentRows);
        }
    }

    setupHeaders() {
        Logger.log('setupHeaders called');
        const headers = this.languageService.getHeaders();
        headers.forEach((header, index) => {
            this.sheet.getRange(1, index + 1)
                .setValue(header)
                .setFontWeight('bold')
                .setBorder(true, true, true, true, true, true);
        });
    }

    setColumnWidths() {
        Logger.log('setColumnWidths called');
        for (const config of COLUMN_CONFIG) {
            const columnIndex = this.sheet.getRange(config.column + '1').getColumn();
            this.sheet.setColumnWidth(columnIndex, config.width);
        }
    }

    applyFormatting() {
        Logger.log('applyFormatting called');
        this.sheet.getRange('B1:G45').setWrap(true)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
        this.sheet.getRange('A1:A45').setWrap(true)
            .setHorizontalAlignment("left")
            .setVerticalAlignment("middle");
    }

    applyBackgroundColors() {
        Logger.log('applyBackgroundColors called');
        this.sheet.getRange('A1:E1').setBackground(COLORS.darkGray());
        this.sheet.getRange('F1:G1').setBackground(COLORS.white()).setFontColor(COLORS.lightGray());
        this.sheet.getRange('A2:A45').setBackground(COLORS.lightGray());
        this.sheet.getRange('C2:C45').setBackground(COLORS.lightYellow());
        this.sheet.getRange('D2:D45').setBackground(COLORS.lightBlue());
    }

    applyTextColorToRange(range, color) {
        Logger.log('applyTextColorToRange called with range: ' + range + ', color: ' + color);
        this.sheet.getRange(range).setFontColor(color);
    }
}

// module.exports = { SheetService };

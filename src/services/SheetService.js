class SheetService {
    constructor(sheet) {
        this.sheet = sheet;
    }

    ensureRowCount(count) {
        const currentRows = this.sheet.getMaxRows();
        if (currentRows > count) {
            this.sheet.deleteRows(count + 1, currentRows - count);
        } else if (currentRows < count) {
            this.sheet.insertRowsAfter(currentRows, count - currentRows);
        }
    }

    setupHeaders() {
        for (const config of COLUMN_CONFIG) {
            const columnIndex = this.sheet.getRange(config.column + '1').getColumn();
            this.sheet.getRange(1, columnIndex)
                .setValue(config.name)
                .setFontWeight('bold')
                .setBorder(true, true, true, true, true, true);
        }
        this.sheet.getRange('A1:E1').setFontColor(COLORS.white());
    }

    setColumnWidths() {
        for (const config of COLUMN_CONFIG) {
            const columnIndex = this.sheet.getRange(config.column + '1').getColumn();
            this.sheet.setColumnWidth(columnIndex, config.width);
        }
    }

    applyFormatting() {
        this.sheet.getRange('B1:G45').setWrap(true)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
        this.sheet.getRange('A1:A45').setWrap(true)
            .setHorizontalAlignment("left")
            .setVerticalAlignment("middle");
    }

    applyBackgroundColors() {
        this.sheet.getRange('A1:E1').setBackground(COLORS.darkGray());
        this.sheet.getRange('F1:G1').setBackground(COLORS.white()).setFontColor(COLORS.lightGray());
        this.sheet.getRange('A2:A45').setBackground(COLORS.lightGray());
        this.sheet.getRange('C2:C45').setBackground(COLORS.lightYellow());
        this.sheet.getRange('D2:D45').setBackground(COLORS.lightBlue());
    }

    applyTextColorToRange(range, color) {
        this.sheet.getRange(range).setFontColor(color);
    }
}

module.exports = { SheetService };

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

    /**
    * Applies the default background colors to the entire sheet.
    */
    applyBackgroundColors() {
        Logger.log('applyBackgroundColors called');
        this.sheet.getRange('A1:E1').setBackground(COLORS.darkGray());
        this.sheet.getRange('F1:G1').setBackground(COLORS.white()).setFontColor(COLORS.lightGray());

        for (let row = 2; row <= 45; row++) {
            this.applyRowBackground(row); // Reuse the same logic for default backgrounds
        }
    }

    /**
    * Applies the background colors for a specific ROW based on the default layout.
    * @param {number} row - The row number where default backgrounds should be reapplied.
    */
    applyRowBackground(row) {
        Logger.log(`applyRowBackground called for row: ${row}`);
        this.sheet.getRange(row, 1).setBackground(COLORS.lightGray());
        this.sheet.getRange(row, 2).setBackground(COLORS.lightGray());
        this.sheet.getRange(row, 3).setBackground(COLORS.lightYellow());
        this.sheet.getRange(row, 4).setBackground(COLORS.lightBlue());
        this.sheet.getRange(row, 5).setBackground(COLORS.lightGray());
    }

    applyTextColorToRange(range, color) {
        Logger.log('applyTextColorToRange called with range: ' + range + ', color: ' + color);
        this.sheet.getRange(range).setFontColor(color);
    }

    /**
     * Checks for completed rows (A to E) and applies light green background to completed rows.
     * Incomplete rows maintain their default background colors.
     */
    applyCompletionFormatting() {
        Logger.log('applyCompletionFormatting called');
        const range = this.sheet.getRange('A2:E45'); // Range of rows to check (A2 to E45)
        const values = range.getValues();

        for (let i = 0; i < values.length; i++) {
            const rowValues = values[i];
            const isRowComplete = rowValues.every(cell => cell !== '');

            const rowRange = this.sheet.getRange(i + 2, 1, 1, 5);

            if (isRowComplete) {
                rowRange.setBackground(COLORS.lightGreen());
            } else {
                this.applyRowBackground(i + 2);
            }
        }
    }
}

// module.exports = { SheetService };

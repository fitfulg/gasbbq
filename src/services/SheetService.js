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

        // Explicitly clear the contents of columns F and G (TEMPORARY SOLUTION)
        this.sheet.getRange('F1').clearContent();
        this.sheet.getRange('G1').clearContent();

        // Assign the first 5 headers to columns A to E
        headers.slice(0, 5).forEach((header, index) => {
            this.sheet.getRange(1, index + 1).setValue(header)
                .setFontWeight('bold')
                .setBorder(true, true, true, true, true, true);
        });

        // Assign the last 2 headers to columns H and I
        this.sheet.getRange('H1').setValue(headers[5]) // C-counter
            .setFontWeight('bold')
            .setBorder(true, true, true, true, true, true);

        this.sheet.getRange('I1').setValue(headers[6]) // D-counter
            .setFontWeight('bold')
            .setBorder(true, true, true, true, true, true);

        SpreadsheetApp.flush(); // Force changes to be written to the sheet
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
        this.sheet.getRange('H1:I1').setBackground(COLORS.white());

        for (let row = 2; row <= 45; row++) {
            this.applyRowBackground(row);
        }
    }

    applyRowBackground(row) {
        Logger.log(`applyRowBackground called for row: ${row}`);
        this.sheet.getRange(row, 1).setBackground(COLORS.lightGray());
        this.sheet.getRange(row, 2).setBackground(COLORS.lightGray());
        this.sheet.getRange(row, 3).setBackground(COLORS.lightYellow());
        this.sheet.getRange(row, 4).setBackground(COLORS.lightBlue());
        this.sheet.getRange(row, 5).setBackground(COLORS.lightGray());
    }

    applyTextColorToRange(range, color) {
        try {
            Logger.log('applyTextColorToRange called with range: ' + range + ', color: ' + color);
            if (range && typeof range.getA1Notation === 'function') {
                range.setFontColor(color);
            } else {
                Logger.log('Invalid range passed to applyTextColorToRange: ' + range);
            }
        } catch (error) {
            Logger.log('Error applying text color to range: ' + error);
            throw error; // rethrow to allow visibility of the error in logs
        }
    }


    /**
     * Applies text colors to columns C and D in a completed row.
     * @param {number} row - The row number to apply the colors.
     */
    applyCompletionTextColor(row) {
        const rangeC = this.sheet.getRange(row, 3); // Get Range for column C
        const rangeD = this.sheet.getRange(row, 4); // Get Range for column D
        this.applyTextColorToRange(rangeC, COLORS.brown()); // Apply brown color to column C
        this.applyTextColorToRange(rangeD, COLORS.blue());  // Apply blue color to column D
    }

    /**
     * Checks for completed rows (A to E) and applies light green background to completed rows.
     * It also changes the text color in columns C and D if the row is complete.
     */
    applyCompletionFormatting() {
        Logger.log('applyCompletionFormatting called');
        const range = this.sheet.getRange('A2:E45'); // Range from columns A to E
        const values = range.getValues();
        const confirmRange = this.sheet.getRange('B2:B45'); // Range for the "Confirmation" column (B)
        const confirmValues = confirmRange.getValues();

        for (let i = 0; i < values.length; i++) {
            const rowValues = values[i];
            const confirmationValue = confirmValues[i][0]; // Value from column B (Yes/No)
            const rowNumber = i + 2; // Row number on the sheet
            const isRowComplete = rowValues.every(cell => cell !== '');

            if (confirmationValue === 'No') {
                this.sheet.getRange(rowNumber, 1, 1, 5).setBackground(COLORS.lightRed());
            } else if (isRowComplete) {
                this.sheet.getRange(rowNumber, 1, 1, 5).setBackground(COLORS.lightGreen());
                this.applyCompletionTextColor(rowNumber); // Apply text color to columns C and D
            } else {
                this.applyRowBackground(rowNumber);
            }

            // Always apply red color to column E
            const rangeE = this.sheet.getRange(rowNumber, 5);
            this.applyTextColorToRange(rangeE, COLORS.red());
        }
    }

    clearRange(cells) {
        Logger.log('clearRange called');
        cells.forEach(cell => {
            this.sheet.getRange(cell).clearContent();
        });
    }

}

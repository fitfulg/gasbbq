class EventController {
    constructor(sheet) {
        this.sheetController = new SheetController(sheet);
    }

    onOpen() {
        this.sheetController.setupSheet();
    }

    onEdit(e) {
        const range = e.range;
        const sheet = e.source.getActiveSheet();

        if (range.getColumn() === 3 || range.getColumn() === 4) {
            const value = range.getValue().trim();
            if (value && value.indexOf(',') === -1 && value.indexOf(' ') !== -1) {
                range.setValue(value.replace(/\s+/g, '-'));
            }

            this.sheetController.wordCountService.countWords('C', 'F');
            this.sheetController.wordCountService.countWords('D', 'G');
        }
    }
}
module.exports = { EventController };
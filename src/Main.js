function onOpen() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const eventController = new EventController(sheet);
    eventController.onOpen();
}

function onEdit(e) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const eventController = new EventController(sheet);
    eventController.onEdit(e);
}
module.exports = { onOpen, onEdit };
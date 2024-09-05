function onOpen() {
    Logger.log('onOpen called');
    Logger.log("Hello World");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const eventController = new EventController(sheet);
    eventController.onOpen();
}

function onEdit(e) {
    Logger.log('onEdit called');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const eventController = new EventController(sheet);
    eventController.onEdit(e);
}
module.exports = { onOpen, onEdit };
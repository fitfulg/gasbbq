function handleEvent(callback, ...args) {
    Logger.log(`${callback.name} called`);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const eventController = new EventController(sheet);
    callback.apply(eventController, args);
}

function onOpen() {
    handleEvent(EventController.prototype.onOpen);
}

function onEdit(e) {
    handleEvent(EventController.prototype.onEdit, e);
}

function onChangeLanguage(languageCode) {
    handleEvent(EventController.prototype.changeLanguage, languageCode);
}

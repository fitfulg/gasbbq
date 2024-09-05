function onOpen() {
    Logger.log('onOpen called');
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

function onChangeLanguage(languageCode) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const eventController = new EventController(sheet);
    eventController.changeLanguage(languageCode);
}

const changeLanguage_en = () => onChangeLanguage('en');
const changeLanguage_es = () => onChangeLanguage('es');
const changeLanguage_ca = () => onChangeLanguage('ca');

// module.exports = { onOpen, onEdit };
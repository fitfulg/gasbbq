// Simulates the global scope of the Google Apps Script environment

global.SheetController = require('./src/controllers/SheetController').SheetController;
global.EventController = require('./src/controllers/EventController').EventController;
global.SheetService = require('./src/services/SheetService').SheetService;
global.ValidationService = require('./src/services/ValidationService').ValidationService;
global.ProtectionService = require('./src/services/ProtectionService').ProtectionService;
global.WordCountService = require('./src/services/WordCountService').WordCountService;
global.Utils = require('./src/utils/Utils').Utils;

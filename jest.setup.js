// Simula la concatenación de archivos y coloca las clases en el entorno global

global.SheetModel = require('./src/models/SheetModel').SheetModel;
global.SheetController = require('./src/controllers/SheetController').SheetController;
global.EventController = require('./src/controllers/EventController').EventController;
global.SheetService = require('./src/services/SheetService').SheetService;
global.ValidationService = require('./src/services/ValidationService').ValidationService;
global.ProtectionService = require('./src/services/ProtectionService').ProtectionService;
global.WordCountService = require('./src/services/WordCountService').WordCountService;
global.ColorUtils = require('./src/utils/Utils'); 

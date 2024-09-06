/**
 * List all properties and their values.
 * @returns {Object} - An object containing all properties and their values.
 */
function listAllProperties() {
    Logger.log('ListAllProperties called. Listing all properties:');
    const properties = SheetPropertiesService.listProperties();
    return properties;
}

/**
 * Deletes a specific property by key.
 * @param {string} key - The key of the property to delete.
 */
function deleteProperty(key) {
    Logger.log(`Attempting to delete property: ${key}`);
    SheetPropertiesService.deleteProperty(key);
    SpreadsheetApp.getUi().alert(`Property ${key} has been deleted.`);
}

/**
 * Deletes all properties.
 */
function clearAllProperties() {
    Logger.log('Attempting to delete all properties.');
    SheetPropertiesService.clearAllProperties();
    SpreadsheetApp.getUi().alert('All properties have been deleted.');
}
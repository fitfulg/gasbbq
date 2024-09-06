class SheetPropertiesService {
    /**
     * Retrieves all properties stored in UserProperties.
     * @returns {Object} - An object containing all properties and their values.
     */
    static listProperties() {
        const userProperties = PropertiesService.getUserProperties();
        const properties = userProperties.getProperties();
        Logger.log('Listing all properties:');
        for (let key in properties) {
            Logger.log(`${key}: ${properties[key]}`);
        }
        return properties;
    }

    /**
     * Retrieves a specific property by key.
     * @param {string} key - The key of the property to retrieve.
     * @returns {string|null} - The value of the property or null if not found.
     */
    static getProperty(key) {
        const userProperties = PropertiesService.getUserProperties();
        return userProperties.getProperty(key);
    }

    /**
     * Deletes a specific property by key.
     * @param {string} key - The key of the property to delete.
     */
    static deleteProperty(key) {
        const userProperties = PropertiesService.getUserProperties();
        userProperties.deleteProperty(key);
        Logger.log(`Deleted property: ${key}`);
    }

    /**
     * Clears all properties.
     */
    static clearAllProperties() {
        const userProperties = PropertiesService.getUserProperties();
        userProperties.deleteAllProperties();
        Logger.log('All properties deleted.');
    }
}


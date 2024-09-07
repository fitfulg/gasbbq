class WordCountService {
    constructor(sheet) {
        Logger.log('WordCountService constructor called');
        this.sheet = sheet;
    }

    countWords(sourceColumn, targetColumn) {
        Logger.log('countWords called with sourceColumn: ' + sourceColumn + ', targetColumn: ' + targetColumn);
        const dataRange = this.sheet.getRange(`${sourceColumn}2:${sourceColumn}45`);
        const dataValues = dataRange.getValues().flat();

        const wordCount = dataValues.reduce((count, value) => {
            if (value) {
                // Removes whitespace + validates that it does not contain symbols exclusively
                const trimmedValue = value.toString().trim();
                // Detect if the string has at least one alphanumeric character
                const hasAlphanumeric = /[a-zA-Z0-9]/.test(trimmedValue);
                if (hasAlphanumeric) {
                    const words = trimmedValue.toLowerCase().split(/[\s,]+/);
                    words.forEach(word => {
                        count[word] = (count[word] || 0) + 1;
                    });
                }
            }
            return count;
        }, {});

        const sortedWordCount = Object.entries(wordCount).sort(([a], [b]) => a.localeCompare(b));
        const resultValues = sortedWordCount.map(([word, count]) => [`${word}: ${count}`]);

        const resultRange = this.sheet.getRange(`${targetColumn}2:${targetColumn}${sortedWordCount.length + 1}`);
        resultRange.clearContent();
        resultRange.setValues(resultValues);
    }
}
// module.exports = { WordCountService };
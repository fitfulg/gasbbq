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
                const words = value.toString().toLowerCase().split(/[\s,]+/);
                words.forEach(word => {
                    count[word] = (count[word] || 0) + 1;
                });
            }
            return count;
        }, {});

        const sortedWordCount = Object.entries(wordCount).sort(([a], [b]) => a.localeCompare(b));
        const resultValues = sortedWordCount.map(([word, count]) => [`${word}: ${count}`]);

        const resultRange = this.sheet.getRange(`${targetColumn}2:${targetColumn}${sortedWordCount.length + 1}`);
        resultRange.clearContent();
        resultRange.setValues(resultValues);  // Aquí se pasa un array 2D correctamente
    }
}
module.exports = { WordCountService };
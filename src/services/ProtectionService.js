class ProtectionService {
    constructor(sheet) {
        this.sheet = sheet;
    }

    protectColumns(ranges) {
        ranges.forEach(range => {
            const protection = this.sheet.getRange(range).protect();
            protection.setDescription('Automatic count protection');
            protection.removeEditors(protection.getEditors());
            if (protection.canDomainEdit()) {
                protection.setDomainEdit(false);
            }
        });
    }
}
module.exports = { ProtectionService };
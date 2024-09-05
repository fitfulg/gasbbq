class ProtectionService {
    constructor(sheet) {
        Logger.log('ProtectionService constructor called');
        this.sheet = sheet;
    }

    protectColumns(ranges) {
        Logger.log('protectColumns called with ranges: ' + ranges);
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
// module.exports = { ProtectionService };
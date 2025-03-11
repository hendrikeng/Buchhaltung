// src/helpers.js

const Helpers = {
    parseDate(value) {
        if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
        if (typeof value === "string") {
            const d = new Date(value);
            return isNaN(d.getTime()) ? null : d;
        }
        return null;
    },

    parseCurrency(value) {
        if (typeof value === "number") return value;
        const str = value.toString().replace(/[^\d,.-]/g, "").replace(",", ".");
        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    },

    parseMwstRate(value) {
        // Hier kannst du zum Beispiel einen Default-Wert festlegen, falls config.tax.defaultMwst nicht verfügbar ist.
        if (typeof value === "number") return value < 1 ? value * 100 : value;
        const rate = parseFloat(value.toString().replace("%", "").replace(",", ".").trim());
        return isNaN(rate) ? 19 : rate;
    },

    getFolderByName(parent, name) {
        const folderIter = parent.getFoldersByName(name);
        return folderIter.hasNext() ? folderIter.next() : null;
    },

    extractDateFromFilename(filename) {
        const nameWithoutExtension = filename.replace(/\.[^/.]+$/, "");
        const match = nameWithoutExtension.match(/RE-(\d{4}-\d{2}-\d{2})/);
        if (match?.[1]) {
            const [year, month, day] = match[1].split("-");
            return `${day}.${month}.${year}`;
        }
        return "";
    },

    setConditionalFormattingForColumn(sheet, column, conditions) {
        const lastRow = sheet.getLastRow();
        const range = sheet.getRange(`${column}2:${column}${lastRow}`);
        const rules = conditions.map(({ value, background, fontColor }) =>
            SpreadsheetApp.newConditionalFormatRule()
                .whenTextEqualTo(value)
                .setBackground(background)
                .setFontColor(fontColor)
                .setRanges([range])
                .build()
        );
        sheet.setConditionalFormatRules(rules);
    },

    getMonthFromRow(row, colIndex = 13) {
        const d = Helpers.parseDate(row[colIndex]); // Aufruf über das Helpers-Objekt
        // Hier als Beispiel: Falls das Jahr nicht 2021 ist, gib 0 zurück
        if (!d || d.getFullYear() !== 2021) return 0;
        return d.getMonth() + 1;
    }
};

export default Helpers;

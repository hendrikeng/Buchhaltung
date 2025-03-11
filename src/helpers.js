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

        // Verschiedene Formate erkennen

        // Format: RE-YYYY-MM-DD oder ähnliches mit Trennzeichen
        let match = nameWithoutExtension.match(/[^0-9](\d{4}[-_.\/]\d{2}[-_.\/]\d{2})[^0-9]/);
        if (match?.[1]) {
            const dateParts = match[1].split(/[-_.\/]/);
            if (dateParts.length === 3) {
                const [year, month, day] = dateParts;
                return `${day}.${month}.${year}`;
            }
        }

        // Format: YYYY-MM-DD am Anfang oder Ende
        match = nameWithoutExtension.match(/(^|[^0-9])(\d{4}[-_.\/]\d{2}[-_.\/]\d{2})($|[^0-9])/);
        if (match?.[2]) {
            const dateParts = match[2].split(/[-_.\/]/);
            if (dateParts.length === 3) {
                const [year, month, day] = dateParts;
                return `${day}.${month}.${year}`;
            }
        }

        // Format: DD.MM.YYYY im Dateinamen
        match = nameWithoutExtension.match(/(\d{2}[.]\d{2}[.]\d{4})/);
        if (match?.[1]) {
            return match[1];
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
        const d = Helpers.parseDate(row[colIndex]);

        // Auf das Jahr aus der Konfiguration prüfen oder das aktuelle Jahr verwenden
        const targetYear = config.tax?.year || new Date().getFullYear();

        // Wenn kein Datum oder das Jahr nicht übereinstimmt
        if (!d || d.getFullYear() !== targetYear) return 0;

        return d.getMonth() + 1;
    }
};

export default Helpers;

// src/helpers.js
import config from "./config.js";

/**
 * Hilfsmodule für verschiedene häufig benötigte Funktionen
 */
const Helpers = {
    /**
     * Konvertiert verschiedene Datumsformate in ein gültiges Date-Objekt
     * @param {Date|string} value - Das zu parsende Datum
     * @returns {Date|null} - Das geparste Datum oder null bei ungültigem Format
     */
    parseDate(value) {
        // Wenn bereits ein Date-Objekt
        if (value instanceof Date) return isNaN(value.getTime()) ? null : value;

        // Wenn String
        if (typeof value === "string") {
            // Deutsche Datumsformate (DD.MM.YYYY) unterstützen
            if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(value)) {
                const [day, month, year] = value.split('.').map(Number);
                const date = new Date(year, month - 1, day);
                return isNaN(date.getTime()) ? null : date;
            }

            const d = new Date(value);
            return isNaN(d.getTime()) ? null : d;
        }

        return null;
    },

    /**
     * Konvertiert einen String oder eine Zahl in einen numerischen Währungswert
     * @param {number|string} value - Der zu parsende Wert
     * @returns {number} - Der geparste Währungswert oder 0 bei ungültigem Format
     */
    parseCurrency(value) {
        if (value === null || value === undefined || value === "") return 0;
        if (typeof value === "number") return value;

        // Entferne alle Zeichen außer Ziffern, Komma, Punkt und Minus
        const str = value.toString()
            .replace(/[^\d,.-]/g, "")
            .replace(/,/g, "."); // Alle Kommas durch Punkte ersetzen

        // Bei mehreren Punkten nur den letzten als Dezimaltrenner behandeln
        const parts = str.split('.');
        if (parts.length > 2) {
            const last = parts.pop();
            return parseFloat(parts.join('') + '.' + last);
        }

        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    },

    /**
     * Parst einen MwSt-Satz und normalisiert ihn
     * @param {number|string} value - Der zu parsende MwSt-Satz
     * @returns {number} - Der normalisierte MwSt-Satz (0-100)
     */
    parseMwstRate(value) {
        if (value === null || value === undefined || value === "") {
            // Verwende den Standard-MwSt-Satz aus der Konfiguration oder fallback auf 19%
            return config?.tax?.defaultMwst || 19;
        }

        if (typeof value === "number") {
            // Wenn der Wert < 1 ist, nehmen wir an, dass es sich um einen Dezimalwert handelt (z.B. 0.19)
            return value < 1 ? value * 100 : value;
        }

        // String-Wert parsen und bereinigen
        const rate = parseFloat(
            value.toString()
                .replace(/%/g, "")
                .replace(/,/g, ".")
                .trim()
        );

        // Wenn der geparste Wert ungültig ist, Standardwert zurückgeben
        if (isNaN(rate)) {
            return config?.tax?.defaultMwst || 19;
        }

        // Normalisieren: Werte < 1 werden als Dezimalwerte interpretiert (z.B. 0.19 -> 19)
        return rate < 1 ? rate * 100 : rate;
    },

    /**
     * Sucht nach einem Ordner mit bestimmtem Namen innerhalb eines übergeordneten Ordners
     * @param {Folder} parent - Der übergeordnete Ordner
     * @param {string} name - Der gesuchte Ordnername
     * @returns {Folder|null} - Der gefundene Ordner oder null
     */
    getFolderByName(parent, name) {
        if (!parent) return null;

        try {
            const folderIter = parent.getFoldersByName(name);
            return folderIter.hasNext() ? folderIter.next() : null;
        } catch (e) {
            console.error("Fehler beim Suchen des Ordners:", e);
            return null;
        }
    },

    /**
     * Extrahiert ein Datum aus einem Dateinamen in verschiedenen Formaten
     * @param {string} filename - Der Dateiname, aus dem das Datum extrahiert werden soll
     * @returns {string} - Das extrahierte Datum im Format DD.MM.YYYY oder leer
     */
    extractDateFromFilename(filename) {
        if (!filename) return "";

        const nameWithoutExtension = filename.replace(/\.[^/.]+$/, "");

        // Verschiedene Formate erkennen (vom spezifischsten zum allgemeinsten)

        // 1. Format: DD.MM.YYYY im Dateinamen (deutsches Format)
        let match = nameWithoutExtension.match(/(\d{2}[.]\d{2}[.]\d{4})/);
        if (match?.[1]) {
            return match[1];
        }

        // 2. Format: RE-YYYY-MM-DD oder ähnliches mit Trennzeichen
        match = nameWithoutExtension.match(/[^0-9](\d{4}[-_.\/]\d{2}[-_.\/]\d{2})[^0-9]/);
        if (match?.[1]) {
            const dateParts = match[1].split(/[-_.\/]/);
            if (dateParts.length === 3) {
                const [year, month, day] = dateParts;
                return `${day}.${month}.${year}`;
            }
        }

        // 3. Format: YYYY-MM-DD am Anfang oder Ende
        match = nameWithoutExtension.match(/(^|[^0-9])(\d{4}[-_.\/]\d{2}[-_.\/]\d{2})($|[^0-9])/);
        if (match?.[2]) {
            const dateParts = match[2].split(/[-_.\/]/);
            if (dateParts.length === 3) {
                const [year, month, day] = dateParts;
                return `${day}.${month}.${year}`;
            }
        }

        // 4. Format: DD-MM-YYYY mit verschiedenen Trennzeichen
        match = nameWithoutExtension.match(/(\d{2})[-_.\/](\d{2})[-_.\/](\d{4})/);
        if (match) {
            const [_, day, month, year] = match;
            return `${day}.${month}.${year}`;
        }

        // 5. Aktuelles Datum als Fallback (optional)
        const today = new Date();
        const day = String(today.getDate()).padStart(2, '0');
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const year = today.getFullYear();

        // Als Kommentar belassen, da es möglicherweise besser ist, ein leeres Datum zurückzugeben
        // return `${day}.${month}.${year}`;

        return "";
    },

    /**
     * Setzt bedingte Formatierung für eine Spalte
     * @param {Sheet} sheet - Das zu formatierende Sheet
     * @param {string} column - Die zu formatierende Spalte (z.B. "A")
     * @param {Array<Object>} conditions - Array mit Bedingungen ({value, background, fontColor})
     */
    setConditionalFormattingForColumn(sheet, column, conditions) {
        if (!sheet || !column || !conditions || !conditions.length) return;

        try {
            const lastRow = Math.max(sheet.getLastRow(), 2);
            const range = sheet.getRange(`${column}2:${column}${lastRow}`);

            // Bestehende Regeln für die Spalte löschen
            const existingRules = sheet.getConditionalFormatRules();
            const newRules = existingRules.filter(rule => {
                const ranges = rule.getRanges();
                return !ranges.some(r =>
                    r.getColumn() === range.getColumn() &&
                    r.getRow() === range.getRow() &&
                    r.getNumColumns() === range.getNumColumns()
                );
            });

            // Neue Regeln erstellen
            const formatRules = conditions.map(({ value, background, fontColor }) =>
                SpreadsheetApp.newConditionalFormatRule()
                    .whenTextEqualTo(value)
                    .setBackground(background || "#ffffff")
                    .setFontColor(fontColor || "#000000")
                    .setRanges([range])
                    .build()
            );

            // Regeln anwenden
            sheet.setConditionalFormatRules([...newRules, ...formatRules]);
        } catch (e) {
            console.error("Fehler beim Setzen der bedingten Formatierung:", e);
        }
    },

    /**
     * Extrahiert den Monat aus einem Datum in einer Zeile
     * @param {Array} row - Die Zeile mit dem Datum
     * @param {string} sheetName - Der Name des Sheets (für Spaltenkonfiguration)
     * @returns {number} - Die Monatsnummer (1-12) oder 0 bei Fehler
     */
    getMonthFromRow(row, sheetName = null) {
        if (!row) return 0;

        // Spalte für Zeitstempel basierend auf Sheet-Typ bestimmen
        let timestampColumn;

        if (sheetName) {
            // Spaltenkonfiguration aus dem Sheetnamen bestimmen
            const sheetConfig = config.sheets[sheetName.toLowerCase()]?.columns;
            if (sheetConfig && sheetConfig.zeitstempel) {
                timestampColumn = sheetConfig.zeitstempel - 1; // 0-basiert
            } else {
                // Fallback auf Standardposition, falls keine Konfiguration gefunden
                timestampColumn = 15; // Standard: Spalte P (16. Spalte, 0-basiert: 15)
            }
        } else {
            // Wenn kein Sheetname angegeben, Fallback auf Position 13
            // (entspricht dem ursprünglichen Wert im Code)
            timestampColumn = 13;
        }

        // Sicherstellen, dass die Zeile lang genug ist
        if (row.length <= timestampColumn) return 0;

        const d = this.parseDate(row[timestampColumn]);

        // Auf das Jahr aus der Konfiguration prüfen oder das aktuelle Jahr verwenden
        const targetYear = config?.tax?.year || new Date().getFullYear();

        // Wenn kein Datum oder das Jahr nicht übereinstimmt
        if (!d || d.getFullYear() !== targetYear) return 0;

        return d.getMonth() + 1; // JavaScript Monate sind 0-basiert, wir geben 1-12 zurück
    },

    /**
     * Formatiert ein Datum im deutschen Format (DD.MM.YYYY)
     * @param {Date|string} date - Das zu formatierende Datum
     * @returns {string} - Das formatierte Datum oder leer bei ungültigem Datum
     */
    formatDate(date) {
        const d = this.parseDate(date);
        if (!d) return "";

        const day = String(d.getDate()).padStart(2, '0');
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const year = d.getFullYear();

        return `${day}.${month}.${year}`;
    },

    /**
     * Formatiert einen Währungsbetrag im deutschen Format
     * @param {number|string} amount - Der zu formatierende Betrag
     * @param {string} currency - Das Währungssymbol (Standard: "€")
     * @returns {string} - Der formatierte Betrag
     */
    formatCurrency(amount, currency = "€") {
        const value = this.parseCurrency(amount);
        return value.toLocaleString('de-DE', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }) + " " + currency;
    },

    /**
     * Generiert eine eindeutige ID (für Referenzzwecke)
     * @param {string} prefix - Optional ein Präfix für die ID
     * @returns {string} - Eine eindeutige ID
     */
    generateUniqueId(prefix = "") {
        const timestamp = new Date().getTime();
        const random = Math.floor(Math.random() * 10000);
        return `${prefix}${timestamp}${random}`;
    },

    /**
     * Konvertiert einen Spaltenindex (1-basiert) in einen Spaltenbuchstaben (A, B, C, ...)
     * @param {number} columnIndex - 1-basierter Spaltenindex
     * @returns {string} - Spaltenbuchstabe(n)
     */
    getColumnLetter(columnIndex) {
        let letter = '';
        while (columnIndex > 0) {
            const modulo = (columnIndex - 1) % 26;
            letter = String.fromCharCode(65 + modulo) + letter;
            columnIndex = Math.floor((columnIndex - modulo) / 26);
        }
        return letter;
    }
};

export default Helpers;
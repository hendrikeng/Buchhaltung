// src/helpers.js
import config from "./config.js";

/**
 * Hilfsmodule für verschiedene häufig benötigte Funktionen
 */
const Helpers = {
    /**
     * Cache für häufig verwendete Berechnungen
     * Verbessert die Performance bei wiederholten Aufrufen
     */
    _cache: {
        dates: new Map(),
        currency: new Map(),
        mwstRates: new Map(),
        columnLetters: new Map()
    },

    /**
     * Cache leeren
     */
    clearCache() {
        this._cache.dates.clear();
        this._cache.currency.clear();
        this._cache.mwstRates.clear();
        this._cache.columnLetters.clear();
    },

    /**
     * Konvertiert verschiedene Datumsformate in ein gültiges Date-Objekt
     * @param {Date|string} value - Das zu parsende Datum
     * @returns {Date|null} - Das geparste Datum oder null bei ungültigem Format
     */
    parseDate(value) {
        // Cache-Lookup für häufig verwendete Werte
        const cacheKey = value instanceof Date
            ? value.getTime().toString()
            : value ? value.toString() : '';

        if (this._cache.dates.has(cacheKey)) {
            return this._cache.dates.get(cacheKey);
        }

        let result = null;

        // Wenn bereits ein Date-Objekt
        if (value instanceof Date) {
            result = isNaN(value.getTime()) ? null : value;
        }
        // Wenn String
        else if (typeof value === "string") {
            // Deutsche Datumsformate (DD.MM.YYYY) unterstützen
            if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(value)) {
                const [day, month, year] = value.split('.').map(Number);
                const date = new Date(year, month - 1, day);
                result = isNaN(date.getTime()) ? null : date;
            } else {
                const d = new Date(value);
                result = isNaN(d.getTime()) ? null : d;
            }
        }

        // Ergebnis cachen
        this._cache.dates.set(cacheKey, result);
        return result;
    },

    /**
     * Konvertiert einen String oder eine Zahl in einen numerischen Währungswert
     * @param {number|string} value - Der zu parsende Wert
     * @returns {number} - Der geparste Währungswert oder 0 bei ungültigem Format
     */
    parseCurrency(value) {
        if (value === null || value === undefined || value === "") return 0;
        if (typeof value === "number") return value;

        // Cache-Lookup für String-Werte
        const stringValue = value.toString();
        if (this._cache.currency.has(stringValue)) {
            return this._cache.currency.get(stringValue);
        }

        // Entferne alle Zeichen außer Ziffern, Komma, Punkt und Minus
        const str = stringValue
            .replace(/[^\d,.-]/g, "")
            .replace(/,/g, "."); // Alle Kommas durch Punkte ersetzen

        // Bei mehreren Punkten nur den letzten als Dezimaltrenner behandeln
        const parts = str.split('.');
        let result;

        if (parts.length > 2) {
            const last = parts.pop();
            result = parseFloat(parts.join('') + '.' + last);
        } else {
            result = parseFloat(str);
        }

        result = isNaN(result) ? 0 : result;

        // Ergebnis cachen
        this._cache.currency.set(stringValue, result);
        return result;
    },

    /**
     * Parst einen MwSt-Satz und normalisiert ihn
     * @param {number|string} value - Der zu parsende MwSt-Satz
     * @returns {number} - Der normalisierte MwSt-Satz (0-100)
     */
    parseMwstRate(value) {
        const defaultMwst = config?.tax?.defaultMwst || 19;

        if (value === null || value === undefined || value === "") {
            return defaultMwst;
        }

        // Cache-Lookup für häufig verwendete Werte
        const cacheKey = value.toString();
        if (this._cache.mwstRates.has(cacheKey)) {
            return this._cache.mwstRates.get(cacheKey);
        }

        let result;

        if (typeof value === "number") {
            // Wenn der Wert < 1 ist, nehmen wir an, dass es sich um einen Dezimalwert handelt (z.B. 0.19)
            result = value < 1 ? value * 100 : value;
        } else {
            // String-Wert parsen und bereinigen
            const rateStr = value.toString()
                .replace(/%/g, "")
                .replace(/,/g, ".")
                .trim();

            const rate = parseFloat(rateStr);

            // Wenn der geparste Wert ungültig ist, Standardwert zurückgeben
            if (isNaN(rate)) {
                result = defaultMwst;
            } else {
                // Normalisieren: Werte < 1 werden als Dezimalwerte interpretiert (z.B. 0.19 -> 19)
                result = rate < 1 ? rate * 100 : rate;
            }
        }

        // Ergebnis cachen
        this._cache.mwstRates.set(cacheKey, result);
        return result;
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

        // Cache-Lookup
        if (this._cache.dates.has(`filename_${filename}`)) {
            return this._cache.dates.get(`filename_${filename}`);
        }

        const nameWithoutExtension = filename.replace(/\.[^/.]+$/, "");
        let result = "";

        // Verschiedene Formate erkennen (vom spezifischsten zum allgemeinsten)

        // 1. Format: DD.MM.YYYY im Dateinamen (deutsches Format)
        let match = nameWithoutExtension.match(/(\d{2}[.]\d{2}[.]\d{4})/);
        if (match?.[1]) {
            result = match[1];
        } else {
            // 2. Format: RE-YYYY-MM-DD oder ähnliches mit Trennzeichen
            match = nameWithoutExtension.match(/[^0-9](\d{4}[-_.\/]\d{2}[-_.\/]\d{2})[^0-9]/);
            if (match?.[1]) {
                const dateParts = match[1].split(/[-_.\/]/);
                if (dateParts.length === 3) {
                    const [year, month, day] = dateParts;
                    result = `${day}.${month}.${year}`;
                }
            } else {
                // 3. Format: YYYY-MM-DD am Anfang oder Ende
                match = nameWithoutExtension.match(/(^|[^0-9])(\d{4}[-_.\/]\d{2}[-_.\/]\d{2})($|[^0-9])/);
                if (match?.[2]) {
                    const dateParts = match[2].split(/[-_.\/]/);
                    if (dateParts.length === 3) {
                        const [year, month, day] = dateParts;
                        result = `${day}.${month}.${year}`;
                    }
                } else {
                    // 4. Format: DD-MM-YYYY mit verschiedenen Trennzeichen
                    match = nameWithoutExtension.match(/(\d{2})[-_.\/](\d{2})[-_.\/](\d{4})/);
                    if (match) {
                        const [_, day, month, year] = match;
                        result = `${day}.${month}.${year}`;
                    }
                }
            }
        }

        // Ergebnis cachen
        this._cache.dates.set(`filename_${filename}`, result);
        return result;
    },

    /**
     * Setzt bedingte Formatierung für eine Spalte
     * @param {Sheet} sheet - Das zu formatierende Sheet
     * @param {string} column - Die zu formatierende Spalte (z.B. "A")
     * @param {Array<Object>} conditions - Array mit Bedingungen ({value, background, fontColor, pattern})
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
            const formatRules = conditions.map(({ value, background, fontColor, pattern }) => {
                let rule;

                if (pattern === "beginsWith") {
                    rule = SpreadsheetApp.newConditionalFormatRule()
                        .whenTextStartsWith(value);
                } else if (pattern === "contains") {
                    rule = SpreadsheetApp.newConditionalFormatRule()
                        .whenTextContains(value);
                } else {
                    rule = SpreadsheetApp.newConditionalFormatRule()
                        .whenTextEqualTo(value);
                }

                return rule
                    .setBackground(background || "#ffffff")
                    .setFontColor(fontColor || "#000000")
                    .setRanges([range])
                    .build();
            });

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
            const sheetConfig = config[sheetName.toLowerCase()]?.columns;
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
        // Cache-Lookup für häufig verwendete Indizes
        if (this._cache.columnLetters.has(columnIndex)) {
            return this._cache.columnLetters.get(columnIndex);
        }

        let letter = '';
        let colIndex = columnIndex;

        while (colIndex > 0) {
            const modulo = (colIndex - 1) % 26;
            letter = String.fromCharCode(65 + modulo) + letter;
            colIndex = Math.floor((colIndex - modulo) / 26);
        }

        // Ergebnis cachen
        this._cache.columnLetters.set(columnIndex, letter);
        return letter;
    },

    /**
     * Prüft, ob zwei Zahlenwerte im Rahmen einer bestimmten Toleranz gleich sind
     * @param {number} a - Erster Wert
     * @param {number} b - Zweiter Wert
     * @param {number} tolerance - Toleranzwert (Standard: 0.01)
     * @returns {boolean} - true wenn Werte innerhalb der Toleranz gleich sind
     */
    isApproximatelyEqual(a, b, tolerance = 0.01) {
        return Math.abs(a - b) <= tolerance;
    },

    /**
     * Sicheres Runden eines Werts auf n Dezimalstellen
     * @param {number} value - Der zu rundende Wert
     * @param {number} decimals - Anzahl der Dezimalstellen (Standard: 2)
     * @returns {number} - Gerundeter Wert
     */
    round(value, decimals = 2) {
        const factor = Math.pow(10, decimals);
        return Math.round((value + Number.EPSILON) * factor) / factor;
    },

    /**
     * Prüft, ob ein Wert leer oder undefiniert ist
     * @param {*} value - Der zu prüfende Wert
     * @returns {boolean} - true wenn der Wert leer ist
     */
    isEmpty(value) {
        return value === null || value === undefined || value.toString().trim() === "";
    },

    /**
     * Bereinigt einen Text von Sonderzeichen und macht ihn vergleichbar
     * @param {string} text - Der zu bereinigende Text
     * @returns {string} - Der bereinigte Text
     */
    normalizeText(text) {
        if (!text) return "";
        return text.toString()
            .toLowerCase()
            .replace(/[äöüß]/g, match => {
                return {
                    'ä': 'ae',
                    'ö': 'oe',
                    'ü': 'ue',
                    'ß': 'ss'
                }[match];
            })
            .replace(/[^a-z0-9]/g, '');
    },

    /**
     * Optimierte Batch-Verarbeitung für Google Sheets API-Calls
     * Vermeidet häufige API-Calls, die zur Drosselung führen können
     * @param {Sheet} sheet - Das zu aktualisierende Sheet
     * @param {Array} data - Array mit Daten-Zeilen
     * @param {number} startRow - Startzeile (1-basiert)
     * @param {number} startCol - Startspalte (1-basiert)
     */
    batchWriteToSheet(sheet, data, startRow, startCol) {
        if (!sheet || !data || !data.length || !data[0].length) return;

        try {
            // Schreibe alle Daten in einem API-Call
            sheet.getRange(
                startRow,
                startCol,
                data.length,
                data[0].length
            ).setValues(data);
        } catch (e) {
            console.error("Fehler beim Batch-Schreiben in das Sheet:", e);

            // Fallback: Schreibe in kleineren Blöcken, falls der ursprüngliche Call fehlschlägt
            const BATCH_SIZE = 50; // Kleinere Batch-Größe für Fallback

            for (let i = 0; i < data.length; i += BATCH_SIZE) {
                const batchData = data.slice(i, i + BATCH_SIZE);
                try {
                    sheet.getRange(
                        startRow + i,
                        startCol,
                        batchData.length,
                        batchData[0].length
                    ).setValues(batchData);

                    // Kurze Pause, um API-Drosselung zu vermeiden
                    if (i + BATCH_SIZE < data.length) {
                        Utilities.sleep(100);
                    }
                } catch (innerError) {
                    console.error(`Fehler beim Schreiben von Batch ${i / BATCH_SIZE}:`, innerError);
                }
            }
        }
    }
};

export default Helpers;
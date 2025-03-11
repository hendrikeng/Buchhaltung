/**
 * helperModule.js - Hilfsmodul mit allgemeinen Funktionen
 *
 * Stellt übergreifende Hilfsfunktionen bereit, die von allen anderen Modulen
 * verwendet werden können, z.B. für Datumformatierung, Währungsberechnungen,
 * Validierungen usw.
 */

const HelperModule = (function() {
    // Private Variablen und Funktionen

    /**
     * Formatiert einen Betrag als Währung (€)
     * @param {number} amount - Der zu formatierende Betrag
     * @param {number} decimals - Anzahl der Dezimalstellen (Standard: 2)
     * @returns {string} Formatierter Betrag mit €-Symbol
     */
    function formatCurrency(amount, decimals = 2) {
        if (typeof amount !== 'number') {
            amount = parseFloat(amount) || 0;
        }
        return amount.toLocaleString('de-DE', {
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals
        }) + ' €';
    }

    /**
     * Parst einen Währungsbetrag aus einem String
     * @param {string|number} value - Der zu parsende Wert
     * @returns {number} Geparster Betrag als Zahl
     */
    function parseCurrency(value) {
        if (typeof value === 'number') return value;

        // String bereinigen: Entferne Währungssymbole, Tausendertrennzeichen, etc.
        const cleanedValue = value.toString().replace(/[^\d,.-]/g, "").replace(",", ".");

        // In Zahl umwandeln
        return parseFloat(cleanedValue) || 0;
    }

    /**
     * Parst einen Mehrwertsteuersatz (z.B. "19%" zu 0.19)
     * @param {string|number} value - Der zu parsende Steuersatz
     * @returns {number} Steuersatz als Dezimalzahl
     */
    function parseMwstRate(value) {
        if (typeof value === 'number') {
            // Wenn der Wert bereits eine Zahl ist und < 1,
            // wird er als Dezimalwert (z.B. 0.19) interpretiert
            if (value < 1) return value;
            // Sonst als Prozent (z.B. 19 ⇒ 0.19)
            return value / 100;
        }

        // String bereinigen und in Zahl umwandeln
        const cleanedValue = value.toString().replace("%", "").replace(",", ".").trim();
        const rate = parseFloat(cleanedValue) || 0;

        // Wenn der Wert > 1 ist, wird er als Prozent interpretiert
        return rate >= 1 ? rate / 100 : rate;
    }

    /**
     * Parst ein Datum aus verschiedenen Formaten
     * @param {Date|string} value - Das zu parsende Datum
     * @returns {Date|null} Geparste Date oder null bei Fehler
     */
    function parseDate(value) {
        if (value instanceof Date) {
            // Bereits ein Date-Objekt
            return isNaN(value.getTime()) ? null : value;
        }

        if (typeof value === 'string') {
            // Verschiedene Datumsformate behandeln
            const germanFormat = /^\d{1,2}\.\d{1,2}\.(\d{2}|\d{4})$/;

            if (germanFormat.test(value)) {
                // Deutsches Format (DD.MM.YYYY) selbst parsen
                const parts = value.split('.');
                const day = parseInt(parts[0], 10);
                const month = parseInt(parts[1], 10) - 1; // Monate sind 0-basiert
                const year = parseInt(parts[2], 10);

                // Zweistelliges Jahr auf 2000er aufstocken
                const fullYear = year < 100 ? 2000 + year : year;

                const date = new Date(fullYear, month, day);
                return isNaN(date.getTime()) ? null : date;
            } else {
                // Standard-Parsing für andere Formate
                const date = new Date(value);
                return isNaN(date.getTime()) ? null : date;
            }
        }

        return null;
    }

    /**
     * Formatiert ein Datum als deutsches Datumsformat
     * @param {Date|string} date - Das zu formatierende Datum
     * @param {boolean} includeTime - Zeit mit anzeigen?
     * @returns {string} Formatiertes Datum (z.B. "31.12.2023")
     */
    function formatDate(date, includeTime = false) {
        const parsedDate = parseDate(date);
        if (!parsedDate) return '';

        const day = String(parsedDate.getDate()).padStart(2, '0');
        const month = String(parsedDate.getMonth() + 1).padStart(2, '0');
        const year = parsedDate.getFullYear();

        let result = `${day}.${month}.${year}`;

        if (includeTime) {
            const hours = String(parsedDate.getHours()).padStart(2, '0');
            const minutes = String(parsedDate.getMinutes()).padStart(2, '0');
            result += ` ${hours}:${minutes}`;
        }

        return result;
    }

    /**
     * Ermittelt das Quartal für ein Datum
     * @param {Date|string} date - Das Datum
     * @returns {number} Quartal (1-4)
     */
    function getQuartal(date) {
        const parsedDate = parseDate(date);
        if (!parsedDate) return 0;

        const month = parsedDate.getMonth(); // 0-basiert
        return Math.ceil((month + 1) / 3);
    }

    /**
     * Ermittelt den Monat aus einer Zeile basierend auf dem Datum
     * @param {Array} row - Zeile mit Daten
     * @param {number} colIndex - Spaltenindex des Datums
     * @returns {number} Monat (1-12) oder 0 bei Fehler
     */
    function getMonthFromRow(row, colIndex = 0) {
        const datum = row[colIndex];
        const parsedDate = parseDate(datum);
        if (!parsedDate) return 0;

        // Optional: Nur Datumsangaben aus dem aktuellen Jahr berücksichtigen
        if (CONFIG.BENUTZER?.AKTUELLES_JAHR &&
            parsedDate.getFullYear() !== CONFIG.BENUTZER.AKTUELLES_JAHR) {
            return 0;
        }

        return parsedDate.getMonth() + 1; // 1-basiert zurückgeben
    }

    /**
     * Sucht nach einem Ordner mit bestimmtem Namen
     * @param {Object} parent - Das übergeordnete Verzeichnis
     * @param {string} name - Name des gesuchten Ordners
     * @returns {Object|null} Gefundener Ordner oder null
     */
    function getFolderByName(parent, name) {
        const folderIter = parent.getFoldersByName(name);
        return folderIter.hasNext() ? folderIter.next() : null;
    }

    /**
     * Extrahiert ein Datum aus einem Dateinamen
     * @param {string} filename - Dateiname
     * @returns {string} Extrahiertes Datum im Format "DD.MM.YYYY" oder Leerstring
     */
    function extractDateFromFilename(filename) {
        const nameWithoutExtension = filename.replace(/\.[^/.]+$/, "");

        // Suche nach dem Muster RE-YYYY-MM-DD (Rechnungsnummer)
        const match = nameWithoutExtension.match(/RE-(\d{4})-(\d{2})-(\d{2})/);
        if (match) {
            const year = match[1];
            const month = match[2];
            const day = match[3];
            return `${day}.${month}.${year}`;
        }

        // Alternative Muster könnte hier implementiert werden

        return "";
    }

    /**
     * Wandelt einen String in einen boolean-Wert um
     * @param {string|boolean} value - Der umzuwandelnde Wert
     * @returns {boolean} Ergebnis der Umwandlung
     */
    function parseBoolean(value) {
        if (typeof value === 'boolean') return value;
        if (typeof value === 'number') return value !== 0;

        if (typeof value === 'string') {
            const lowerValue = value.toLowerCase().trim();
            return lowerValue === 'true' ||
                lowerValue === 'ja' ||
                lowerValue === 'yes' ||
                lowerValue === '1';
        }

        return false;
    }

    /**
     * Zeigt eine Toast-Nachricht im Spreadsheet an
     * @param {string} message - Anzuzeigende Nachricht
     * @param {string} type - Typ der Nachricht (Info, Erfolg, Warnung, Fehler)
     */
    function showToast(message, type = 'Info') {
        // Farben/Icons je nach Meldungstyp anpassen
        let title;
        switch(type.toLowerCase()) {
            case 'erfolg':
                title = '✅ Erfolg';
                break;
            case 'warnung':
                title = '⚠️ Warnung';
                break;
            case 'fehler':
                title = '❌ Fehler';
                break;
            default:
                title = 'ℹ️ Info';
        }

        SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 5);
    }

    /**
     * Bedingte Formatierung für eine Spalte anwenden
     * @param {Object} sheet - Sheet-Objekt
     * @param {string} column - Spaltenbezeichnung (z.B. "E")
     * @param {Array} conditions - Array mit Bedingungen {value, background, fontColor}
     */
    function setConditionalFormattingForColumn(sheet, column, conditions) {
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return; // Nur Header

        const range = sheet.getRange(`${column}2:${column}${lastRow}`);

        // Regeln erstellen
        const rules = conditions.map(({ value, background, fontColor }) =>
            SpreadsheetApp.newConditionalFormatRule()
                .whenTextEqualTo(value)
                .setBackground(background)
                .setFontColor(fontColor)
                .setRanges([range])
                .build()
        );

        // Regeln anwenden
        sheet.setConditionalFormatRules(rules);
    }

    /**
     * Validiert, ob ein Wert numerisch ist
     * @param {*} value - Zu prüfender Wert
     * @returns {boolean} true, wenn der Wert numerisch ist
     */
    function isNumeric(value) {
        if (typeof value === 'number') return true;
        if (!value) return false;

        return !isNaN(parseFloat(value)) && isFinite(value.toString().trim().replace(',', '.'));
    }

    /**
     * Prüft, ob ein Wert leer ist
     * @param {*} value - Zu prüfender Wert
     * @returns {boolean} true, wenn der Wert leer ist
     */
    function isEmpty(value) {
        if (value === null || value === undefined) return true;
        if (typeof value === 'string') return value.trim() === '';
        return false;
    }

    /**
     * Generiert eine eindeutige ID
     * @param {string} prefix - Optionales Präfix
     * @returns {string} Eindeutige ID
     */
    function generateUniqueId(prefix = '') {
        const timestamp = new Date().getTime();
        const random = Math.floor(Math.random() * 10000);
        return `${prefix}${timestamp}-${random}`;
    }

    /**
     * Ermittelt die Anzahl der Tage zwischen zwei Daten
     * @param {Date|string} date1 - Erstes Datum
     * @param {Date|string} date2 - Zweites Datum
     * @returns {number} Anzahl der Tage
     */
    function daysBetween(date1, date2) {
        const d1 = parseDate(date1);
        const d2 = parseDate(date2);

        if (!d1 || !d2) return 0;

        // Zeitzone ignorieren und nur auf Tagesbasis berechnen
        const utc1 = Date.UTC(d1.getFullYear(), d1.getMonth(), d1.getDate());
        const utc2 = Date.UTC(d2.getFullYear(), d2.getMonth(), d2.getDate());

        // Millisekunden zu Tagen
        return Math.floor((utc2 - utc1) / (1000 * 60 * 60 * 24));
    }

    /**
     * Erzeugt eine Dropdown-Validierung für eine Zelle oder einen Bereich
     * @param {Object} sheet - Sheet-Objekt
     * @param {number} row - Startzeile
     * @param {number} col - Startspalte
     * @param {number} numRows - Anzahl der Zeilen
     * @param {number} numCols - Anzahl der Spalten
     * @param {Array} values - Erlaubte Werte für das Dropdown
     */
    function createDropdownValidation(sheet, row, col, numRows, numCols, values) {
        // Validierungsregel erstellen
        const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(values, true)
            .build();

        // Regel auf den gewünschten Bereich anwenden
        sheet.getRange(row, col, numRows, numCols).setDataValidation(rule);
    }

    // Öffentliche API
    return {
        /**
         * Parst ein Datum aus verschiedenen Formaten
         * @param {Date|string} value - Das zu parsende Datum
         * @returns {Date|null} Geparste Date oder null bei Fehler
         */
        parseDate: parseDate,

        /**
         * Formatiert ein Datum als deutsches Datumsformat
         * @param {Date|string} date - Das zu formatierende Datum
         * @param {boolean} includeTime - Zeit mit anzeigen?
         * @returns {string} Formatiertes Datum (z.B. "31.12.2023")
         */
        formatDate: formatDate,

        /**
         * Parst einen Währungsbetrag aus einem String
         * @param {string|number} value - Der zu parsende Wert
         * @returns {number} Geparster Betrag als Zahl
         */
        parseCurrency: parseCurrency,

        /**
         * Formatiert einen Betrag als Währung (€)
         * @param {number} amount - Der zu formatierende Betrag
         * @param {number} decimals - Anzahl der Dezimalstellen (Standard: 2)
         * @returns {string} Formatierter Betrag mit €-Symbol
         */
        formatCurrency: formatCurrency,

        /**
         * Parst einen Mehrwertsteuersatz (z.B. "19%" zu 0.19)
         * @param {string|number} value - Der zu parsende Steuersatz
         * @returns {number} Steuersatz als Dezimalzahl
         */
        parseMwstRate: parseMwstRate,

        /**
         * Ermittelt das Quartal für ein Datum
         * @param {Date|string} date - Das Datum
         * @returns {number} Quartal (1-4)
         */
        getQuartal: getQuartal,

        /**
         * Ermittelt den Monat aus einer Zeile basierend auf dem Datum
         * @param {Array} row - Zeile mit Daten
         * @param {number} colIndex - Spaltenindex des Datums
         * @returns {number} Monat (1-12) oder 0 bei Fehler
         */
        getMonthFromRow: getMonthFromRow,

        /**
         * Sucht nach einem Ordner mit bestimmtem Namen
         * @param {Object} parent - Das übergeordnete Verzeichnis
         * @param {string} name - Name des gesuchten Ordners
         * @returns {Object|null} Gefundener Ordner oder null
         */
        getFolderByName: getFolderByName,

        /**
         * Extrahiert ein Datum aus einem Dateinamen
         * @param {string} filename - Dateiname
         * @returns {string} Extrahiertes Datum im Format "DD.MM.YYYY" oder Leerstring
         */
        extractDateFromFilename: extractDateFromFilename,

        /**
         * Zeigt eine Toast-Nachricht im Spreadsheet an
         * @param {string} message - Anzuzeigende Nachricht
         * @param {string} type - Typ der Nachricht (Info, Erfolg, Warnung, Fehler)
         */
        showToast: showToast,

        /**
         * Bedingte Formatierung für eine Spalte anwenden
         * @param {Object} sheet - Sheet-Objekt
         * @param {string} column - Spaltenbezeichnung (z.B. "E")
         * @param {Array} conditions - Array mit Bedingungen {value, background, fontColor}
         */
        setConditionalFormattingForColumn: setConditionalFormattingForColumn,

        /**
         * Erzeugt eine Dropdown-Validierung für eine Zelle oder einen Bereich
         * @param {Object} sheet - Sheet-Objekt
         * @param {number} row - Startzeile
         * @param {number} col - Startspalte
         * @param {number} numRows - Anzahl der Zeilen
         * @param {number} numCols - Anzahl der Spalten
         * @param {Array} values - Erlaubte Werte für das Dropdown
         */
        createDropdownValidation: createDropdownValidation,

        /**
         * Validiert, ob ein Wert numerisch ist
         * @param {*} value - Zu prüfender Wert
         * @returns {boolean} true, wenn der Wert numerisch ist
         */
        isNumeric: isNumeric,

        /**
         * Prüft, ob ein Wert leer ist
         * @param {*} value - Zu prüfender Wert
         * @returns {boolean} true, wenn der Wert leer ist
         */
        isEmpty: isEmpty,

        /**
         * Generiert eine eindeutige ID
         * @param {string} prefix - Optionales Präfix
         * @returns {string} Eindeutige ID
         */
        generateUniqueId: generateUniqueId,

        /**
         * Ermittelt die Anzahl der Tage zwischen zwei Daten
         * @param {Date|string} date1 - Erstes Datum
         * @param {Date|string} date2 - Zweites Datum
         * @returns {number} Anzahl der Tage
         */
        daysBetween: daysBetween,

        /**
         * Wandelt einen String in einen boolean-Wert um
         * @param {string|boolean} value - Der umzuwandelnde Wert
         * @returns {boolean} Ergebnis der Umwandlung
         */
        parseBoolean: parseBoolean
    };
})();
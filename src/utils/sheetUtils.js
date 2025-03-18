// utils/sheetUtils.js
/**
 * Funktionen für die Arbeit mit Google Sheets
 */

/**
 * Optimierte Batch-Verarbeitung für Google Sheets API-Calls
 * Vermeidet häufige API-Calls, die zur Drosselung führen können
 * @param {Sheet} sheet - Das zu aktualisierende Sheet
 * @param {Array} data - Array mit Daten-Zeilen
 * @param {number} startRow - Startzeile (1-basiert)
 * @param {number} startCol - Startspalte (1-basiert)
 * @returns {boolean} - True bei Erfolg, false bei Fehler
 */
function batchWriteToSheet(sheet, data, startRow, startCol) {
    if (!sheet || !data || !data.length || !data[0].length) return false;

    try {
        // Optimierung: Prüfen, ob die Range existiert und genug Zeilen/Spalten hat
        const existingRows = sheet.getMaxRows();
        const existingCols = sheet.getMaxColumns();
        const requiredRows = startRow + data.length - 1;
        const requiredCols = startCol + data[0].length - 1;

        // Range automatisch erweitern wenn nötig
        if (requiredRows > existingRows) {
            sheet.insertRowsAfter(existingRows, requiredRows - existingRows);
        }

        if (requiredCols > existingCols) {
            sheet.insertColumnsAfter(existingCols, requiredCols - existingCols);
        }

        // Schreibe alle Daten in einem API-Call
        sheet.getRange(
            startRow,
            startCol,
            data.length,
            data[0].length,
        ).setValues(data);

        return true;
    } catch (e) {
        console.error('Fehler beim Batch-Schreiben in das Sheet:', e);

        // Fallback: Schreibe in kleineren Blöcken
        try {
            const BATCH_SIZE = 50; // Kleinere Batch-Größe für Fallback
            let success = true;

            for (let i = 0; i < data.length; i += BATCH_SIZE) {
                const batchData = data.slice(i, i + BATCH_SIZE);
                try {
                    sheet.getRange(
                        startRow + i,
                        startCol,
                        batchData.length,
                        batchData[0].length,
                    ).setValues(batchData);

                    // Kurze Pause, um API-Drosselung zu vermeiden
                    if (i + BATCH_SIZE < data.length) {
                        Utilities.sleep(100);
                    }
                } catch (innerError) {
                    console.error(`Fehler beim Schreiben von Batch ${i / BATCH_SIZE}:`, innerError);
                    success = false;
                }
            }

            return success;
        } catch (fallbackError) {
            console.error('Fallback-Fehler beim Batch-Schreiben:', fallbackError);
            return false;
        }
    }
}

/**
 * Setzt bedingte Formatierung für eine Spalte mit optimierter Regelanwendung
 * @param {Sheet} sheet - Das zu formatierende Sheet
 * @param {string} column - Die zu formatierende Spalte (z.B. "A")
 * @param {Array<Object>} conditions - Array mit Bedingungen ({value, background, fontColor, pattern})
 * @returns {boolean} - True bei Erfolg, false bei Fehler
 */
function setConditionalFormattingForColumn(sheet, column, conditions) {
    if (!sheet || !column || !conditions || !conditions.length) return false;

    try {
        const lastRow = Math.max(sheet.getLastRow(), 2);
        const range = sheet.getRange(`${column}2:${column}${lastRow}`);

        // Bestehende Regeln für die Spalte abrufen und filtern
        const existingRules = sheet.getConditionalFormatRules();
        const newRules = existingRules.filter(rule => {
            const ranges = rule.getRanges();
            return !ranges.some(r =>
                r.getColumn() === range.getColumn() &&
                r.getRow() === range.getRow() &&
                r.getNumColumns() === range.getNumColumns(),
            );
        });

        // Neue Regeln in einem Batch erstellen
        const formatRules = conditions.map(({ value, background, fontColor, pattern }) => {
            let rule;

            if (pattern === 'beginsWith') {
                rule = SpreadsheetApp.newConditionalFormatRule()
                    .whenTextStartsWith(value);
            } else if (pattern === 'contains') {
                rule = SpreadsheetApp.newConditionalFormatRule()
                    .whenTextContains(value);
            } else {
                rule = SpreadsheetApp.newConditionalFormatRule()
                    .whenTextEqualTo(value);
            }

            return rule
                .setBackground(background || '#ffffff')
                .setFontColor(fontColor || '#000000')
                .setRanges([range])
                .build();
        });

        // Regeln in einem API-Call anwenden
        sheet.setConditionalFormatRules([...newRules, ...formatRules]);
        return true;
    } catch (e) {
        console.error('Fehler beim Setzen der bedingten Formatierung:', e);
        return false;
    }
}

/**
 * Sucht nach einem Ordner mit bestimmtem Namen innerhalb eines übergeordneten Ordners
 * @param {Folder} parent - Der übergeordnete Ordner
 * @param {string} name - Der gesuchte Ordnername
 * @returns {Folder|null} - Der gefundene Ordner oder null
 */
function getFolderByName(parent, name) {
    if (!parent) return null;

    try {
        const folderIter = parent.getFoldersByName(name);
        return folderIter.hasNext() ? folderIter.next() : null;
    } catch (e) {
        console.error('Fehler beim Suchen des Ordners:', e);
        return null;
    }
}

/**
 * Holt ein Sheet aus dem Spreadsheet oder erstellt es, wenn es nicht existiert
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {string} sheetName - Name des Sheets
 * @returns {Sheet} - Das vorhandene oder neu erstellte Sheet
 */
function getOrCreateSheet(ss, sheetName) {
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
    }

    return sheet;
}

export default {
    batchWriteToSheet,
    setConditionalFormattingForColumn,
    getFolderByName,
    getOrCreateSheet,
};
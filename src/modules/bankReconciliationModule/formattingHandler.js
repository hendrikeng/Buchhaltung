// src/modules/bankReconciliationModule/formattingHandler.js
import stringUtils from '../../utils/stringUtils.js';

/**
 * Formatiert Zeilen basierend auf dem Match-Typ
 * @param {Sheet} sheet - Das Sheet
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {Array} matchResults - Match-Ergebnisse
 * @param {Object} columns - Spaltenkonfiguration
 */
function formatMatchedRows(sheet, firstDataRow, matchResults, columns) {
    // Performance-optimiertes Batch-Update vorbereiten
    const formatBatches = {
        'Einnahme': { rows: [], color: '#EAF1DD' },
        'Vollständige Zahlung (Einnahme)': { rows: [], color: '#C6EFCE' },
        'Teilzahlung (Einnahme)': { rows: [], color: '#FCE4D6' },
        'Ausgabe': { rows: [], color: '#FFCCCC' },
        'Vollständige Zahlung (Ausgabe)': { rows: [], color: '#FFC7CE' },
        'Teilzahlung (Ausgabe)': { rows: [], color: '#FCE4D6' },
        'Eigenbeleg': { rows: [], color: '#DDEBF7' },
        'Vollständige Zahlung (Eigenbeleg)': { rows: [], color: '#9BC2E6' },
        'Teilzahlung (Eigenbeleg)': { rows: [], color: '#FCE4D6' },
        'Gesellschafterkonto': { rows: [], color: '#E2EFDA' },
        'Holding Transfer': { rows: [], color: '#FFF2CC' },
        'Gutschrift': { rows: [], color: '#E6E0FF' },
        'Gesellschaftskonto/Holding': { rows: [], color: '#FFEB9C' },
    };

    // Zeilen nach Kategorien gruppieren
    matchResults.forEach((matchInfo, index) => {
        const rowIndex = firstDataRow + index;
        const matchText = (matchInfo && matchInfo[0]) ? matchInfo[0].toString() : '';

        if (!matchText) return; // Überspringe leere Matches

        if (matchText.includes('Einnahme')) {
            if (matchText.includes('Vollständige Zahlung')) {
                formatBatches['Vollständige Zahlung (Einnahme)'].rows.push(rowIndex);
            } else if (matchText.includes('Teilzahlung')) {
                formatBatches['Teilzahlung (Einnahme)'].rows.push(rowIndex);
            } else {
                formatBatches['Einnahme'].rows.push(rowIndex);
            }
        } else if (matchText.includes('Ausgabe')) {
            if (matchText.includes('Vollständige Zahlung')) {
                formatBatches['Vollständige Zahlung (Ausgabe)'].rows.push(rowIndex);
            } else if (matchText.includes('Teilzahlung')) {
                formatBatches['Teilzahlung (Ausgabe)'].rows.push(rowIndex);
            } else {
                formatBatches['Ausgabe'].rows.push(rowIndex);
            }
        } else if (matchText.includes('Eigenbeleg')) {
            if (matchText.includes('Vollständige Zahlung')) {
                formatBatches['Vollständige Zahlung (Eigenbeleg)'].rows.push(rowIndex);
            } else if (matchText.includes('Teilzahlung')) {
                formatBatches['Teilzahlung (Eigenbeleg)'].rows.push(rowIndex);
            } else {
                formatBatches['Eigenbeleg'].rows.push(rowIndex);
            }
        } else if (matchText.includes('Gesellschafterkonto')) {
            formatBatches['Gesellschafterkonto'].rows.push(rowIndex);
        } else if (matchText.includes('Holding Transfer')) {
            formatBatches['Holding Transfer'].rows.push(rowIndex);
        } else if (matchText.includes('Gutschrift')) {
            formatBatches['Gutschrift'].rows.push(rowIndex);
        } else if (matchText.includes('Gesellschaftskonto') || matchText.includes('Holding')) {
            formatBatches['Gesellschaftskonto/Holding'].rows.push(rowIndex);
        }
    });

    // Batch-Formatting anwenden
    Object.values(formatBatches).forEach(batch => {
        if (batch.rows.length > 0) {
            // Gruppen von maximal 20 Zeilen formatieren (API-Limits vermeiden)
            applyFormatBatches(sheet, batch.rows, batch.color, columns.matchInfo);
        }
    });
}

/**
 * Wendet Formatierung auf Batches von Zeilen an
 */
function applyFormatBatches(sheet, rows, color, maxColumn) {
    const chunkSize = 20;
    for (let i = 0; i < rows.length; i += chunkSize) {
        const chunk = rows.slice(i, i + chunkSize);
        chunk.forEach(rowIndex => {
            try {
                sheet.getRange(rowIndex, 1, 1, maxColumn)
                    .setBackground(color);
            } catch (e) {
                console.error(`Fehler beim Formatieren von Zeile ${rowIndex}:`, e);
            }
        });

        // Kurze Pause um API-Limits zu vermeiden
        if (i + chunkSize < rows.length) {
            Utilities.sleep(50);
        }
    }
}

/**
 * Setzt bedingte Formatierung für die Match-Spalte
 */
function setMatchColumnFormatting(sheet, columnLetter) {
    const conditions = [
        // Grundlegende Match-Typen
        {value: 'Einnahme', background: '#C6EFCE', fontColor: '#006100', pattern: 'beginsWith'},
        {value: 'Ausgabe', background: '#FFC7CE', fontColor: '#9C0006', pattern: 'beginsWith'},
        {value: 'Eigenbeleg', background: '#DDEBF7', fontColor: '#2F5597', pattern: 'beginsWith'},
        {value: 'Gesellschafterkonto', background: '#E2EFDA', fontColor: '#375623', pattern: 'beginsWith'},
        {value: 'Holding Transfer', background: '#FFF2CC', fontColor: '#7F6000', pattern: 'beginsWith'},
        {value: 'Gutschrift', background: '#E6E0FF', fontColor: '#5229A3', pattern: 'beginsWith'},

        // Zusätzliche Betragstypen
        {value: 'Vollständige Zahlung', background: '#C6EFCE', fontColor: '#006100', pattern: 'contains'},
        {value: 'Teilzahlung', background: '#FCE4D6', fontColor: '#974706', pattern: 'contains'},
        {value: 'Unsichere Zahlung', background: '#F8CBAD', fontColor: '#843C0C', pattern: 'contains'},
        {value: 'Vollständige Gutschrift', background: '#E6E0FF', fontColor: '#5229A3', pattern: 'contains'},
    ];

    // Bedingte Formatierung für die Match-Spalte setzen
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const range = sheet.getRange(`${columnLetter}2:${columnLetter}${lastRow}`);

    // Bestehende Regeln für die Spalte löschen
    const existingRules = sheet.getConditionalFormatRules();
    const newRules = existingRules.filter(rule => {
        const ranges = rule.getRanges();
        return !ranges.some(r =>
            r.getColumn() === range.getColumn() &&
            r.getRow() === range.getRow() &&
            r.getNumColumns() === range.getNumColumns(),
        );
    });

    // Neue Regeln erstellen
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
            .setBackground(background)
            .setFontColor(fontColor)
            .setRanges([range])
            .build();
    });

    // Regeln anwenden
    sheet.setConditionalFormatRules([...newRules, ...formatRules]);
}

export default {
    formatMatchedRows,
    setMatchColumnFormatting,
};
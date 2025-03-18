// src/modules/bankReconciliationModule/formattingHandler.js
import stringUtils from '../../utils/stringUtils.js';

/**
 * Formatiert Zeilen basierend auf dem Match-Typ mit optimierter Batch-Verarbeitung
 * @param {Sheet} sheet - Das Sheet
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {Array} matchResults - Match-Ergebnisse
 * @param {Object} columns - Spaltenkonfiguration
 */
function formatMatchedRows(sheet, firstDataRow, matchResults, columns) {
    // Performance-optimiertes Batch-Update - Kategorien definieren
    const matchCategories = {
        'Einnahme': { color: '#EAF1DD' },
        'Vollständige Zahlung (Einnahme)': { color: '#C6EFCE' },
        'Teilzahlung (Einnahme)': { color: '#FCE4D6' },
        'Ausgabe': { color: '#FFCCCC' },
        'Vollständige Zahlung (Ausgabe)': { color: '#FFC7CE' },
        'Teilzahlung (Ausgabe)': { color: '#FCE4D6' },
        'Eigenbeleg': { color: '#DDEBF7' },
        'Vollständige Zahlung (Eigenbeleg)': { color: '#9BC2E6' },
        'Teilzahlung (Eigenbeleg)': { color: '#FCE4D6' },
        'Gesellschafterkonto': { color: '#E2EFDA' },
        'Holding Transfer': { color: '#FFF2CC' },
        'Gutschrift': { color: '#E6E0FF' },
        'Gesellschaftskonto/Holding': { color: '#FFEB9C' },
    };

    // Optimiert: In einem Durchgang kategorisieren
    const rowsByCategory = {};

    // Initialisiere Kategorien
    Object.keys(matchCategories).forEach(category => {
        rowsByCategory[category] = [];
    });

    // Kategorisiere Zeilen in einem Durchgang
    matchResults.forEach((matchInfo, index) => {
        const rowIndex = firstDataRow + index;
        const matchText = (matchInfo && matchInfo[0]) ? matchInfo[0].toString() : '';

        if (!matchText) return; // Überspringe leere Matches

        // Optimierung: Mit einer einzigen if-Kette statt mehreren Prüfungen
        if (matchText.includes('Einnahme')) {
            if (matchText.includes('Vollständige Zahlung')) {
                rowsByCategory['Vollständige Zahlung (Einnahme)'].push(rowIndex);
            } else if (matchText.includes('Teilzahlung')) {
                rowsByCategory['Teilzahlung (Einnahme)'].push(rowIndex);
            } else {
                rowsByCategory['Einnahme'].push(rowIndex);
            }
        } else if (matchText.includes('Ausgabe')) {
            if (matchText.includes('Vollständige Zahlung')) {
                rowsByCategory['Vollständige Zahlung (Ausgabe)'].push(rowIndex);
            } else if (matchText.includes('Teilzahlung')) {
                rowsByCategory['Teilzahlung (Ausgabe)'].push(rowIndex);
            } else {
                rowsByCategory['Ausgabe'].push(rowIndex);
            }
        } else if (matchText.includes('Eigenbeleg')) {
            if (matchText.includes('Vollständige Zahlung')) {
                rowsByCategory['Vollständige Zahlung (Eigenbeleg)'].push(rowIndex);
            } else if (matchText.includes('Teilzahlung')) {
                rowsByCategory['Teilzahlung (Eigenbeleg)'].push(rowIndex);
            } else {
                rowsByCategory['Eigenbeleg'].push(rowIndex);
            }
        } else if (matchText.includes('Gesellschafterkonto')) {
            rowsByCategory['Gesellschafterkonto'].push(rowIndex);
        } else if (matchText.includes('Holding Transfer')) {
            rowsByCategory['Holding Transfer'].push(rowIndex);
        } else if (matchText.includes('Gutschrift')) {
            rowsByCategory['Gutschrift'].push(rowIndex);
        } else if (matchText.includes('Gesellschaftskonto') || matchText.includes('Holding')) {
            rowsByCategory['Gesellschaftskonto/Holding'].push(rowIndex);
        }
    });

    // Batch-Formatting für jede Kategorie anwenden
    Object.entries(rowsByCategory).forEach(([category, rows]) => {
        if (rows.length > 0) {
            applyFormatBatches(sheet, rows, matchCategories[category].color, columns.matchInfo);
        }
    });
}

/**
 * Wendet Formatierung auf Batches von Zeilen an - optimiert für weniger API-Calls
 * @param {Sheet} sheet - Das Sheet
 * @param {Array} rows - Die zu formatierenden Zeilen
 * @param {string} color - Die Hintergrundfarbe
 * @param {number} maxColumn - Die maximale Spalte
 */
function applyFormatBatches(sheet, rows, color, maxColumn) {
    // Optimierung: Größere Chunks verwenden, um API-Calls zu reduzieren
    const chunkSize = 50;

    // Rows nach Intervallen gruppieren für effizientere Formatierung
    const contiguousRanges = [];
    let currentRange = [];

    // Sortieren für effizientere Intervallverarbeitung
    rows.sort((a, b) => a - b);

    // Zusammenhängende Zeilen identifizieren
    for (let i = 0; i < rows.length; i++) {
        if (i === 0 || rows[i] !== rows[i-1] + 1) {
            // Neue Gruppe starten
            if (currentRange.length > 0) {
                contiguousRanges.push(currentRange);
            }
            currentRange = [rows[i]];
        } else {
            // Aktuelle Gruppe erweitern
            currentRange.push(rows[i]);
        }
    }

    // Letzte Gruppe hinzufügen
    if (currentRange.length > 0) {
        contiguousRanges.push(currentRange);
    }

    // Zusammenhängende Bereiche formatieren (reduziert API-Calls drastisch)
    contiguousRanges.forEach(range => {
        try {
            const startRow = range[0];
            const numRows = range.length;
            sheet.getRange(startRow, 1, numRows, maxColumn)
                .setBackground(color);
        } catch (e) {
            console.error(`Fehler beim Formatieren von Zeilenbereich ${range[0]}-${range[range.length-1]}:`, e);

            // Fallback: Einzelne Zeilen formatieren
            for (let i = 0; i < range.length; i += chunkSize) {
                const chunk = range.slice(i, i + chunkSize);
                try {
                    chunk.forEach(rowIndex => {
                        sheet.getRange(rowIndex, 1, 1, maxColumn)
                            .setBackground(color);
                    });
                } catch (innerError) {
                    console.error('Fehler beim Formatieren von Zeilen-Chunk:', innerError);
                }

                // Pause um API-Limits zu vermeiden
                if (i + chunkSize < range.length) {
                    Utilities.sleep(100);
                }
            }
        }
    });
}

/**
 * Setzt bedingte Formatierung für die Match-Spalte mit optimierten Regeln
 * @param {Sheet} sheet - Das Sheet
 * @param {string} columnLetter - Der Spaltenbuchstabe
 */
function setMatchColumnFormatting(sheet, columnLetter) {
    // Optimierter Array von Formatierungsbedingungen
    const conditions = [
        // Grundlegende Match-Typen (mit Formatierungs-Patterns)
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

    // Verwende die optimierte Methode für bedingte Formatierung
    try {
        // Automatisch vorhandene Regeln entfernen und neue hinzufügen
        const range = sheet.getRange(`${columnLetter}2:${columnLetter}${lastRow}`);

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

        // Neue Regeln erstellen mit korrektem Typ
        const formatRules = conditions.map(({ value, background, fontColor, pattern }) => {
            let rule;

            switch (pattern) {
            case 'beginsWith':
                rule = SpreadsheetApp.newConditionalFormatRule()
                    .whenTextStartsWith(value);
                break;
            case 'contains':
                rule = SpreadsheetApp.newConditionalFormatRule()
                    .whenTextContains(value);
                break;
            default:
                rule = SpreadsheetApp.newConditionalFormatRule()
                    .whenTextEqualTo(value);
            }

            return rule
                .setBackground(background)
                .setFontColor(fontColor)
                .setRanges([range])
                .build();
        });

        // Regeln in einem API-Call anwenden
        sheet.setConditionalFormatRules([...newRules, ...formatRules]);
    } catch (e) {
        console.error('Fehler beim Setzen der bedingten Formatierung für Match-Spalte:', e);
    }
}

export default {
    formatMatchedRows,
    setMatchColumnFormatting,
};
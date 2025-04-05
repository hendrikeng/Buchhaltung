// src/modules/refreshModule/dataSheetHandler.js
import stringUtils from '../../utils/stringUtils.js';
import formattingHandler from './formattingHandler.js';
import accountHandler from './accountHandler.js';

/**
 * Aktualisiert ein Datenblatt (Einnahmen, Ausgaben, Eigenbelege) mit optimierter Batch-Verarbeitung
 * @param {Sheet} sheet - Das zu aktualisierende Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} - true bei erfolgreicher Aktualisierung
 */
function refreshDataSheet(sheet, sheetName, config) {
    try {
        console.log(`Refreshing ${sheetName} sheet`);
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return true; // Keine Daten zum Aktualisieren

        const numRows = lastRow - 1;

        // Passende Spaltenkonfiguration für das entsprechende Sheet auswählen
        const configKeyMap = {
            'Einnahmen': 'einnahmen',
            'Ausgaben': 'ausgaben',
            'Eigenbelege': 'eigenbelege',
            'Gesellschafterkonto': 'gesellschafterkonto',
            'Holding Transfers': 'holdingTransfers',
        };

        const configKey = configKeyMap[sheetName];
        if (!configKey) return false; // Unbekanntes Sheet

        const columns = config[configKey].columns;
        console.log(`Found ${numRows} rows to process with columns configuration for ${configKey}`);

        // Spaltenbuchstaben für Formeln einmalig generieren
        const columnLetters = {};
        Object.entries(columns).forEach(([key, index]) => {
            columnLetters[key] = stringUtils.getColumnLetter(index);
        });

        // Optimierung: Formeln für alle Zeilen in Batches erstellen
        const formulaBatches = createFormulaBatches(numRows, columnLetters, columns, sheetName);

        // Formeln in Batches anwenden (weniger API-Calls)
        applyFormulaBatches(sheet, formulaBatches, columns);

        // Bezahlter Betrag - Leerzeichen durch 0 ersetzen
        updateEmptyPaymentValues(sheet, columns, numRows);

        // Dropdown-Validierungen je nach Sheet-Typ setzen
        formattingHandler.setDropdownValidations(sheet, sheetName, numRows, columns, config);

        // Bedingte Formatierung für Status-Spalte
        if (columns.zahlungsstatus) {
            applyStatusColumnFormatting(sheet, sheetName, columnLetters.zahlungsstatus);
        }

        // Buchungskonto basierend auf Kategorie aktualisieren
        accountHandler.updateBookingAccounts(sheet, configKey, config, false);

        // Spaltenbreiten anpassen
        sheet.autoResizeColumns(1, sheet.getLastColumn());

        console.log(`${sheetName} sheet refreshed successfully`);
        return true;
    } catch (e) {
        console.error(`Fehler beim Aktualisieren von ${sheetName}:`, e);
        return false;
    }
}

/**
 * Erstellt Formel-Batches für alle zu aktualisierenden Spalten
 * @param {number} numRows - Anzahl der Datenzeilen
 * @param {Object} columnLetters - Spaltenbuchstaben
 * @param {Object} columns - Spaltenkonfiguration
 * @param {string} sheetName - Name des Sheets
 * @returns {Object} Formel-Batches
 */
function createFormulaBatches(numRows, columnLetters, columns, sheetName) {
    const formulaBatches = {};

    // MwSt-Betrag mit Berücksichtigung von Ausland
    if (columns.mwstBetrag && columns.nettobetrag && columns.mwstSatz && columns.ausland) {
        formulaBatches[columns.mwstBetrag] = [];
        for (let i = 0; i < numRows; i++) {
            const row = i + 2;
            const formula = `=IF(OR(${columnLetters.ausland}${row}="EU-Ausland";${columnLetters.ausland}${row}="Nicht-EU-Ausland");0;${columnLetters.nettobetrag}${row}*${columnLetters.mwstSatz}${row})`;
            formulaBatches[columns.mwstBetrag].push([formula]);
        }
    }
    // Fallback für den Fall, dass keine Ausland-Spalte vorhanden ist
    else if (columns.mwstBetrag && columns.nettobetrag && columns.mwstSatz) {
        formulaBatches[columns.mwstBetrag] = Array.from(
            {length: numRows},
            (_, i) => [`=${columnLetters.nettobetrag}${i + 2}*${columnLetters.mwstSatz}${i + 2}`],
        );
    }

    // Brutto-Betrag
    if (columns.bruttoBetrag && columns.nettobetrag && columns.mwstBetrag) {
        formulaBatches[columns.bruttoBetrag] = Array.from(
            {length: numRows},
            (_, i) => [`=${columnLetters.nettobetrag}${i + 2}+${columnLetters.mwstBetrag}${i + 2}`],
        );
    }

    // Restbetrag Netto für Teilzahlungen
    if (columns.restbetragNetto && columns.bruttoBetrag && columns.bezahlt && columns.mwstSatz) {
        // Bei Ausland ist der MwSt-Satz 0, daher müssen wir hier eine Sonderbehandlung einfügen
        if (columns.ausland) {
            formulaBatches[columns.restbetragNetto] = [];
            for (let i = 0; i < numRows; i++) {
                const row = i + 2;
                const formula = `=IF(OR(${columnLetters.ausland}${row}="EU-Ausland";${columnLetters.ausland}${row}="Nicht-EU-Ausland");${columnLetters.bruttoBetrag}${row}-${columnLetters.bezahlt}${row};(${columnLetters.bruttoBetrag}${row}-${columnLetters.bezahlt}${row})/(1+${columnLetters.mwstSatz}${row}))`;
                formulaBatches[columns.restbetragNetto].push([formula]);
            }
        } else {
            formulaBatches[columns.restbetragNetto] = Array.from(
                {length: numRows},
                (_, i) => [`=(${columnLetters.bruttoBetrag}${i + 2}-${columnLetters.bezahlt}${i + 2})/(1+${columnLetters.mwstSatz}${i + 2})`],
            );
        }
    }

    // Quartal
    if (columns.quartal && columns.datum) {
        formulaBatches[columns.quartal] = Array.from(
            {length: numRows},
            (_, i) => [`=IF(${columnLetters.datum}${i + 2}="";"";ROUNDUP(MONTH(${columnLetters.datum}${i + 2})/3;0))`],
        );
    }

    // Zahlungsstatus
    if (columns.zahlungsstatus && columns.bezahlt && columns.bruttoBetrag) {
        if (sheetName !== 'Eigenbelege') {
            // Für Einnahmen und Ausgaben
            formulaBatches[columns.zahlungsstatus] = Array.from(
                {length: numRows},
                (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Bezahlt";"Teilbezahlt"))`],
            );
        } else {
            // Für Eigenbelege
            formulaBatches[columns.zahlungsstatus] = Array.from(
                {length: numRows},
                (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Erstattet";"Teilerstattet"))`],
            );
        }
    }

    return formulaBatches;
}

/**
 * Wendet Formel-Batches auf das Sheet an
 * @param {Sheet} sheet - Das Sheet
 * @param {Object} formulaBatches - Formel-Batches
 * @param {Object} columns - Spaltenkonfiguration
 */
function applyFormulaBatches(sheet, formulaBatches, columns) {
    // Optimierung: Formeln in optimierten Batches anwenden
    const batchSize = 500; // Größerer Batch für bessere Performance

    Object.entries(formulaBatches).forEach(([col, formulas]) => {
        try {
            // Für große Sheets in Batches arbeiten
            if (formulas.length > batchSize) {
                for (let i = 0; i < formulas.length; i += batchSize) {
                    const batchFormulas = formulas.slice(i, Math.min(i + batchSize, formulas.length));
                    sheet.getRange(2 + i, Number(col), batchFormulas.length, 1)
                        .setFormulas(batchFormulas);

                    // Kurze Pause zwischen den Batches
                    if (i + batchSize < formulas.length) {
                        Utilities.sleep(100);
                    }
                }
            } else {
                // Für kleinere Sheets alles in einem API-Call
                sheet.getRange(2, Number(col), formulas.length, 1)
                    .setFormulas(formulas);
            }
        } catch (e) {
            console.error(`Fehler beim Setzen der Formeln für Spalte ${col}:`, e);
        }
    });
}

/**
 * Aktualisiert leere Zahlungswerte auf 0
 * @param {Sheet} sheet - Das Sheet
 * @param {Object} columns - Spaltenkonfiguration
 * @param {number} numRows - Anzahl der Datenzeilen
 */
function updateEmptyPaymentValues(sheet, columns, numRows) {
    if (!columns.bezahlt) return;

    try {
        // Alle Zahlungswerte in einem API-Call laden
        const bezahltRange = sheet.getRange(2, columns.bezahlt, numRows, 1);
        const bezahltValues = bezahltRange.getValues();

        // Prüfen, ob Aktualisierungen nötig sind
        let needsUpdate = false;
        const updatedValues = bezahltValues.map(([val]) => {
            if (val === '' || val === null || val === undefined) {
                needsUpdate = true;
                return [0];
            }
            return [val];
        });

        // Nur aktualisieren, wenn es leere Werte gab
        if (needsUpdate) {
            bezahltRange.setValues(updatedValues);
        }
    } catch (e) {
        console.error('Fehler beim Aktualisieren leerer Zahlungswerte:', e);
    }
}

/**
 * Wendet bedingte Formatierung auf die Status-Spalte an
 * @param {Sheet} sheet - Das Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {string} statusColumn - Spalte für den Status
 */
function applyStatusColumnFormatting(sheet, sheetName, statusColumn) {
    try {
        if (sheetName !== 'Eigenbelege') {
            // Für Einnahmen und Ausgaben
            formattingHandler.setConditionalFormattingForStatusColumn(sheet, statusColumn, [
                {value: 'Offen', background: '#FFC7CE', fontColor: '#9C0006'},
                {value: 'Teilbezahlt', background: '#FFEB9C', fontColor: '#9C6500'},
                {value: 'Bezahlt', background: '#C6EFCE', fontColor: '#006100'},
            ]);
        } else {
            // Für Eigenbelege
            formattingHandler.setConditionalFormattingForStatusColumn(sheet, statusColumn, [
                {value: 'Offen', background: '#FFC7CE', fontColor: '#9C0006'},
                {value: 'Teilerstattet', background: '#FFEB9C', fontColor: '#9C6500'},
                {value: 'Erstattet', background: '#C6EFCE', fontColor: '#006100'},
            ]);
        }
    } catch (e) {
        console.error(`Fehler bei der Statusformatierung für ${sheetName}:`, e);
    }
}

export default {
    refreshDataSheet,
};
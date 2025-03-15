// src/modules/refreshModule/dataSheetHandler.js
import stringUtils from '../../utils/stringUtils.js';
import formattingHandler from './formattingHandler.js';
import cellValidator from '../validatorModule/cellValidator.js';

/**
 * Aktualisiert ein Datenblatt (Einnahmen, Ausgaben, Eigenbelege)
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
        let columns;
        if (sheetName === 'Einnahmen') {
            columns = config.einnahmen.columns;
        } else if (sheetName === 'Ausgaben') {
            columns = config.ausgaben.columns;
        } else if (sheetName === 'Eigenbelege') {
            columns = config.eigenbelege.columns;
        } else if (sheetName === 'Gesellschafterkonto') {
            columns = config.gesellschafterkonto.columns;
        } else if (sheetName === 'Holding Transfers') {
            columns = config.holdingTransfers.columns;
        } else {
            return false; // Unbekanntes Sheet
        }

        console.log(`Found ${numRows} rows to process with columns:`, columns);

        // Spaltenbuchstaben aus den Indizes generieren
        const columnLetters = {};
        Object.entries(columns).forEach(([key, index]) => {
            columnLetters[key] = stringUtils.getColumnLetter(index);
        });

        // Batch-Array für Formeln erstellen (effizienter als einzelne Range-Updates)
        const formulasBatch = {};

        // MwSt-Betrag
        if (columns.mwstBetrag && columns.nettobetrag && columns.mwstSatz) {
            formulasBatch[columns.mwstBetrag] = Array.from(
                {length: numRows},
                (_, i) => [`=${columnLetters.nettobetrag}${i + 2}*${columnLetters.mwstSatz}${i + 2}`],
            );
        }

        // Brutto-Betrag
        if (columns.bruttoBetrag && columns.nettobetrag && columns.mwstBetrag) {
            formulasBatch[columns.bruttoBetrag] = Array.from(
                {length: numRows},
                (_, i) => [`=${columnLetters.nettobetrag}${i + 2}+${columnLetters.mwstBetrag}${i + 2}`],
            );
        }

        // Restbetrag Netto für Teilzahlungen
        if (columns.restbetragNetto && columns.bruttoBetrag && columns.bezahlt && columns.mwstSatz) {
            formulasBatch[columns.restbetragNetto] = Array.from(
                {length: numRows},
                (_, i) => [`=(${columnLetters.bruttoBetrag}${i + 2}-${columnLetters.bezahlt}${i + 2})/(1+${columnLetters.mwstSatz}${i + 2})`],
            );
        }

        // Quartal
        if (columns.quartal && columns.datum) {
            formulasBatch[columns.quartal] = Array.from(
                {length: numRows},
                (_, i) => [`=IF(${columnLetters.datum}${i + 2}="";"";ROUNDUP(MONTH(${columnLetters.datum}${i + 2})/3;0))`],
            );
        }

        // Zahlungsstatus
        if (columns.zahlungsstatus && columns.bezahlt && columns.bruttoBetrag) {
            if (sheetName !== 'Eigenbelege') {
                // Für Einnahmen und Ausgaben: Zahlungsstatus
                formulasBatch[columns.zahlungsstatus] = Array.from(
                    {length: numRows},
                    (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Bezahlt";"Teilbezahlt"))`],
                );
            } else {
                // Für Eigenbelege: Zahlungsstatus
                formulasBatch[columns.zahlungsstatus] = Array.from(
                    {length: numRows},
                    (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Erstattet";"Teilerstattet"))`],
                );
            }
        }

        console.log('Formula batch keys:', Object.keys(formulasBatch));

        // Formeln in Batches anwenden (weniger API-Calls)
        Object.entries(formulasBatch).forEach(([col, formulas]) => {
            try {
                console.log(`Setting ${formulas.length} formulas for column ${col}`);
                sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulas);
            } catch (e) {
                console.error(`Fehler beim Setzen der Formeln für Spalte ${col}:`, e);
            }
        });

        // Bezahlter Betrag - Leerzeichen durch 0 ersetzen für Berechnungen
        if (columns.bezahlt) {
            const bezahltRange = sheet.getRange(2, columns.bezahlt, numRows, 1);
            const bezahltValues = bezahltRange.getValues();
            const updatedBezahltValues = bezahltValues.map(
                ([val]) => [stringUtils.isEmpty(val) ? 0 : val],
            );
            bezahltRange.setValues(updatedBezahltValues);
        }

        // Dropdown-Validierungen je nach Sheet-Typ setzen
        console.log('Setting dropdown validations');
        formattingHandler.setDropdownValidations(sheet, sheetName, numRows, columns, config);

        // Bedingte Formatierung für Status-Spalte
        if (columns.zahlungsstatus) {
            console.log('Setting conditional formatting for status column');
            if (sheetName !== 'Eigenbelege') {
                // Für Einnahmen und Ausgaben: Zahlungsstatus
                formattingHandler.setConditionalFormattingForStatusColumn(sheet, columnLetters.zahlungsstatus, [
                    {value: 'Offen', background: '#FFC7CE', fontColor: '#9C0006'},
                    {value: 'Teilbezahlt', background: '#FFEB9C', fontColor: '#9C6500'},
                    {value: 'Bezahlt', background: '#C6EFCE', fontColor: '#006100'},
                ]);
            } else {
                // Für Eigenbelege: Status
                formattingHandler.setConditionalFormattingForStatusColumn(sheet, columnLetters.zahlungsstatus, [
                    {value: 'Offen', background: '#FFC7CE', fontColor: '#9C0006'},
                    {value: 'Teilerstattet', background: '#FFEB9C', fontColor: '#9C6500'},
                    {value: 'Erstattet', background: '#C6EFCE', fontColor: '#006100'},
                ]);
            }
        }

        // Spaltenbreiten automatisch anpassen
        sheet.autoResizeColumns(1, sheet.getLastColumn());

        console.log(`${sheetName} sheet refreshed successfully`);
        return true;
    } catch (e) {
        console.error(`Fehler beim Aktualisieren von ${sheetName}:`, e);
        return false;
    }
}

export default {
    refreshDataSheet,
};
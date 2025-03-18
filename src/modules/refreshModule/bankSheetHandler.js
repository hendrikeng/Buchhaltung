// src/modules/refreshModule/bankSheetHandler.js
import matchingHandler from './matchingHandler.js';
import formattingHandler from './formattingHandler.js';
import accountHandler from './accountHandler.js';
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';

/**
 * Aktualisiert das Bankbewegungen-Sheet
 * @param {Sheet} sheet - Das Bankbewegungen-Sheet
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} - true bei erfolgreicher Aktualisierung
 */
function refreshBankSheet(sheet, config) {
    try {
        console.log('Refreshing bank sheet');
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const lastRow = sheet.getLastRow();
        if (lastRow < 3) return true; // Nicht genügend Daten zum Aktualisieren

        const firstDataRow = 3; // Erste Datenzeile (nach Header-Zeile)
        const numDataRows = lastRow - firstDataRow + 1;
        const transRows = lastRow - firstDataRow - 1; // Anzahl der Transaktionszeilen ohne die letzte Zeile

        console.log(`Found ${numDataRows} rows to process in bank sheet`);

        // Bankbewegungen-Konfiguration für Spalten
        const columns = config.bankbewegungen.columns;

        // Spaltenbuchstaben aus den Indizes generieren
        const columnLetters = {};
        Object.entries(columns).forEach(([key, index]) => {
            columnLetters[key] = stringUtils.getColumnLetter(index);
        });

        // 1. Saldo-Formeln setzen (jede Zeile verwendet den Saldo der vorherigen Zeile + aktuellen Betrag)
        if (transRows > 0) {
            console.log(`Setting ${transRows} saldo formulas`);
            sheet.getRange(firstDataRow, columns.saldo, transRows, 1).setFormulas(
                Array.from({length: transRows}, (_, i) =>
                    [`=${columnLetters.saldo}${firstDataRow + i - 1}+${columnLetters.betrag}${firstDataRow + i}`],
                ),
            );
        }

        // 2. Transaktionstyp basierend auf dem Betrag setzen (Einnahme/Ausgabe)
        const amounts = sheet.getRange(firstDataRow, columns.betrag, numDataRows, 1).getValues();
        const typeValues = amounts.map(([val]) => {
            const amt = numberUtils.parseCurrency(val);
            return [amt > 0 ? 'Einnahme' : amt < 0 ? 'Ausgabe' : ''];
        });
        sheet.getRange(firstDataRow, columns.transaktionstyp, numDataRows, 1).setValues(typeValues);

        // 3. Dropdown-Validierungen für Typ, Kategorie und Konten
        console.log('Applying bank sheet validations');
        formattingHandler.applyBankSheetValidations(sheet, firstDataRow, numDataRows, columns, config);

        // 4. Bedingte Formatierung für Transaktionstyp-Spalte
        formattingHandler.setConditionalFormattingForStatusColumn(sheet, columnLetters.transaktionstyp, [
            {value: 'Einnahme', background: '#C6EFCE', fontColor: '#006100'},
            {value: 'Ausgabe', background: '#FFC7CE', fontColor: '#9C0006'},
        ]);

        // 5. Buchungskonten (KontoSoll, KontoHaben) mit accountHandler aktualisieren
        accountHandler.updateBookingAccounts(sheet, 'bankbewegungen', config, true);

        // 6. REFERENZEN-MATCHING: Suche nach Referenzen in Einnahmen- und Ausgaben-Sheets
        console.log('Performing bank reference matching');
        matchingHandler.performBankReferenceMatching(ss, sheet, firstDataRow, numDataRows, lastRow, columns, columnLetters, config);

        // 7. Endsaldo-Zeile aktualisieren
        updateEndSaldoRow(sheet, lastRow, columns, columnLetters);

        // 8. Spaltenbreiten anpassen
        sheet.autoResizeColumns(1, sheet.getLastColumn());

        console.log('Bank sheet refreshed successfully');
        return true;
    } catch (e) {
        console.error('Fehler beim Aktualisieren des Bankbewegungen-Sheets:', e);
        return false;
    }
}

/**
 * Aktualisiert die Endsaldo-Zeile im Bankbewegungen-Sheet
 * @param {Sheet} sheet - Das Sheet
 * @param {number} lastRow - Letzte Zeile
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} columnLetters - Buchstaben für die Spalten
 */
function updateEndSaldoRow(sheet, lastRow, columns, columnLetters) {
    // Endsaldo-Zeile überprüfen
    const lastRowText = sheet.getRange(lastRow, columns.buchungstext).getDisplayValue().toString().trim().toLowerCase();
    const endSaldoDate = sheet.getRange(lastRow -1, columns.datum).getDisplayValue();

    if (lastRowText === 'endsaldo') {
        // Endsaldo-Zeile aktualisieren
        sheet.getRange(lastRow, columns.datum).setValue(endSaldoDate);
        sheet.getRange(lastRow, columns.saldo).setFormula(`=${columnLetters.saldo}${lastRow - 1}`);
    } else {
        // Neue Endsaldo-Zeile erstellen
        const newRow = Array(sheet.getLastColumn()).fill('');
        newRow[columns.datum - 1] = endSaldoDate;
        newRow[columns.buchungstext - 1] = 'Endsaldo';
        newRow[columns.saldo - 1] = sheet.getRange(lastRow, columns.saldo).getValue();

        sheet.appendRow(newRow);
    }
}

export default {
    refreshBankSheet,
};
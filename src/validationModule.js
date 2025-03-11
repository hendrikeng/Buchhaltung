/**
 * validationModule.js - Modul für die Datenvalidierung
 *
 * Prüft die Daten auf Vollständigkeit und Korrektheit
 * und gibt entsprechende Fehlermeldungen aus
 */

const ValidationModule = (function() {
    // Private Variablen und Funktionen

    /**
     * Prüft, ob eine Zeile leer oder gelöscht ist
     * @param {Array} row - Zeile aus dem Sheet
     * @returns {boolean} True, wenn die Zeile leer ist
     */
    function isRowEmpty(row) {
        return row.every(cell => cell === null || cell === '');
    }

    /**
     * Prüft, ob zwei Zahlen (mit Toleranz) übereinstimmen
     * @param {number} a - Erste Zahl
     * @param {number} b - Zweite Zahl
     * @param {number} tolerance - Toleranzgrenze (Standard: 0.01)
     * @returns {boolean} True, wenn die Zahlen innerhalb der Toleranz übereinstimmen
     */
    function numbersMatch(a, b, tolerance = 0.01) {
        return Math.abs(a - b) <= tolerance;
    }

    /**
     * Erzeugt eine Fehlermeldung für eine bestimmte Zelle
     * @param {number} row - Zeilennummer (1-basiert)
     * @param {number} col - Spaltennummer (1-basiert)
     * @param {string} message - Fehlermeldung
     * @param {string} severity - Schweregrad ('error', 'warning', 'info')
     * @returns {Object} Fehlermeldung als Objekt
     */
    function createError(row, col, message, severity = 'error') {
        return {
            row: row,
            col: col,
            message: message,
            severity: severity
        };
    }

    /**
     * Validiert die Einnahmen-Daten
     * @param {Array} data - Daten aus dem Einnahmen-Sheet
     * @returns {Array} Array mit Fehlermeldungen
     */
    function validateEinnahmen(data) {
        const errors = [];

        // Header überspringen
        for (let i = 1; i < data.length; i++) {
            const row = data[i];

            // Leere Zeilen überspringen
            if (isRowEmpty(row)) continue;

            // Zeilennummer (1-basiert)
            const rowNum = i + 1;

            // Pflichtfelder prüfen
            if (!row[CONFIG.SYSTEM.EINNAHMEN_COLS.DATUM]) {
                errors.push(createError(rowNum, CONFIG.SYSTEM.EINNAHMEN_COLS.DATUM + 1, 'Rechnungsdatum fehlt'));
            }

            if (!row[CONFIG.SYSTEM.EINNAHMEN_COLS.RECHNUNGSNUMMER]) {
                errors.push(createError(rowNum, CONFIG.SYSTEM.EINNAHMEN_COLS.RECHNUNGSNUMMER + 1, 'Rechnungsnummer fehlt'));
            }

            if (!row[CONFIG.SYSTEM.EINNAHMEN_COLS.KATEGORIE]) {
                errors.push(createError(rowNum, CONFIG.SYSTEM.EINNAHMEN_COLS.KATEGORIE + 1, 'Kategorie fehlt'));
            }

            // Beträge und Berechnungen prüfen
            const netto = parseFloat(row[CONFIG.SYSTEM.EINNAHMEN_COLS.NETTO]) || 0;
            const mwstSatz = parseFloat(row[CONFIG.SYSTEM.EINNAHMEN_COLS.MWST_SATZ]) || 0;
            const mwstBetrag = parseFloat(row[CONFIG.SYSTEM.EINNAHMEN_COLS.MWST_BETRAG]) || 0;
            const brutto = parseFloat(row[CONFIG.SYSTEM.EINNAHMEN_COLS.BRUTTO]) || 0;

            // MwSt-Betrag korrekt?
            const expectedMwSt = netto * mwstSatz;
            if (!numbersMatch(mwstBetrag, expectedMwSt)) {
                errors.push(createError(
                    rowNum,
                    CONFIG.SYSTEM.EINNAHMEN_COLS.MWST_BETRAG + 1,
                    `MwSt-Betrag falsch: ${mwstBetrag.toFixed(2)}€, erwartet: ${expectedMwSt.toFixed(2)}€`
                ));
            }

            // Bruttobetrag korrekt?
            const expectedBrutto = netto + mwstBetrag;
            if (!numbersMatch(brutto, expectedBrutto)) {
                errors.push(createError(
                    rowNum,
                    CONFIG.SYSTEM.EINNAHMEN_COLS.BRUTTO + 1,
                    `Bruttobetrag falsch: ${brutto.toFixed(2)}€, erwartet: ${expectedBrutto.toFixed(2)}€`
                ));
            }

            // Zahlungsstatus konsistent?
            const zahlungsstatus = row[CONFIG.SYSTEM.EINNAHMEN_COLS.ZAHLUNGSSTATUS];
            const bezahlt = parseFloat(row[CONFIG.SYSTEM.EINNAHMEN_COLS.BEZAHLT]) || 0;

            if (zahlungsstatus === 'Bezahlt' && !numbersMatch(bezahlt, brutto)) {
                errors.push(createError(
                    rowNum,
                    CONFIG.SYSTEM.EINNAHMEN_COLS.BEZAHLT + 1,
                    `Bezahlter Betrag (${bezahlt.toFixed(2)}€) stimmt nicht mit Brutto (${brutto.toFixed(2)}€) überein`,
                    'warning'
                ));
            }

            if (zahlungsstatus === 'Offen' && bezahlt > 0) {
                errors.push(createError(
                    rowNum,
                    CONFIG.SYSTEM.EINNAHMEN_COLS.ZAHLUNGSSTATUS + 1,
                    `Status ist 'Offen', aber bereits ${bezahlt.toFixed(2)}€ bezahlt`,
                    'warning'
                ));
            }

            // Bei Zahlungsdatum: Prüfen, ob Zahlungsstatus auch "Bezahlt" ist
            const zahlungsdatum = row[CONFIG.SYSTEM.EINNAHMEN_COLS.ZAHLUNGSDATUM];
            if (zahlungsdatum && zahlungsstatus !== 'Bezahlt') {
                errors.push(createError(
                    rowNum,
                    CONFIG.SYSTEM.EINNAHMEN_COLS.ZAHLUNGSDATUM + 1,
                    `Zahlungsdatum vorhanden, aber Status ist nicht 'Bezahlt'`,
                    'warning'
                ));
            }
        }

        return errors;
    }

    /**
     * Validiert die Ausgaben-Daten
     * @param {Array} data - Daten aus dem Ausgaben-Sheet
     * @returns {Array} Array mit Fehlermeldungen
     */
    function validateAusgaben(data) {
        const errors = [];

        // Header überspringen
        for (let i = 1; i < data.length; i++) {
            const row = data[i];

            // Leere Zeilen überspringen
            if (isRowEmpty(row)) continue;

            // Zeilennummer (1-basiert)
            const rowNum = i + 1;

            // Pflichtfelder prüfen
            if (!row[CONFIG.SYSTEM.AUSGABEN_COLS.DATUM]) {
                errors.push(createError(rowNum, CONFIG.SYSTEM.AUSGABEN_COLS.DATUM + 1, 'Rechnungsdatum fehlt'));
            }

            if (!row[CONFIG.SYSTEM.AUSGABEN_COLS.RECHNUNGSNUMMER]) {
                errors.push(createError(rowNum, CONFIG.SYSTEM.AUSGABEN_COLS.RECHNUNGSNUMMER + 1, 'Rechnungsnummer fehlt'));
            }

            if (!row[CONFIG.SYSTEM.AUSGABEN_COLS.KATEGORIE]) {
                errors.push(createError(rowNum, CONFIG.SYSTEM.AUSGABEN_COLS.KATEGORIE + 1, 'Kategorie fehlt'));
            }

            // Beträge und Berechnungen prüfen
            const netto = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.NETTO]) || 0;
            const mwstSatz = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.MWST_SATZ]) || 0;
            const mwstBetrag = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.MWST_BETRAG]) || 0;
            const brutto = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.BRUTTO]) || 0;

            // MwSt-Betrag korrekt?
            const expectedMwSt = netto * mwstSatz;
            if (!numbersMatch(mwstBetrag, expectedMwSt)) {
                errors.push(createError(
                    rowNum,
                    CONFIG.SYSTEM.AUSGABEN_COLS.MWST_BETRAG + 1,
                    `MwSt-Betrag falsch: ${mwstBetrag.toFixed(2)}€, erwartet: ${expectedMwSt.toFixed(2)}€`
                ));
            }

            // Bruttobetrag korrekt?
            const expectedBrutto = netto + mwstBetrag;
            if (!numbersMatch(brutto, expectedBrutto)) {
                errors.push(createError(
                    rowNum,
                    CONFIG.SYSTEM.AUSGABEN_COLS.BRUTTO + 1,
                    `Bruttobetrag falsch: ${brutto.toFixed(2)}€, erwartet: ${expectedBrutto.toFixed(2)}€`
                ));
            }

            // Zahlungsstatus konsistent?
            const zahlungsstatus = row[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS];
            const bezahlt = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.BEZAHLT]) || 0;

            if (zahlungsstatus === 'Bezahlt' && !numbersMatch(bezahlt, brutto)) {
                errors.push(createError(
                    rowNum,
                    CONFIG.SYSTEM.AUSGABEN_COLS.BEZAHLT + 1,
                    `Bezahlter Betrag (${bezahlt.toFixed(2)}€) stimmt nicht mit Brutto (${brutto.toFixed(2)}€) überein`,
                    'warning'
                ));
            }
        }

        return errors;
    }

    /**
     * Validiert die Bankbewegungen-Daten
     * @param {Array} data - Daten aus dem Bankbewegungen-Sheet
     * @returns {Array} Array mit Fehlermeldungen
     */
    function validateBankbewegungen(data) {
        const errors = [];

        // Header überspringen
        for (let i = 1; i < data.length; i++) {
            const row = data[i];

            // Leere Zeilen überspringen
            if (isRowEmpty(row)) continue;

            // Zeilennummer (1-basiert)
            const rowNum = i + 1;

            // Pflichtfelder prüfen
            if (!row[CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.DATUM]) {
                errors.push(createError(rowNum, CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.DATUM + 1, 'Buchungsdatum fehlt'));
            }

            // Betrag sollte vorhanden sein
            if (row[CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.BETRAG] === null || row[CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.BETRAG] === '') {
                errors.push(createError(rowNum, CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.BETRAG + 1, 'Betrag fehlt'));
            }

            // Typ sollte für Nicht-Anfangssaldo gesetzt sein
            if (row[CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.BUCHUNGSTEXT] !== 'Anfangssaldo' &&
                !row[CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.TYP]) {
                errors.push(createError(
                    rowNum,
                    CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.TYP + 1,
                    'Typ (Einnahme/Ausgabe) fehlt',
                    'warning'
                ));
            }

            // Prüfen, ob Soll- und Haben-Konten zugewiesen wurden, wenn die Buchung erfasst ist
            const erfasst = row[CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.ERFASST];
            const sollKonto = row[CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.KONTO_SOLL];
            const habenKonto = row[CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.KONTO_HABEN];

            if (erfasst && (!sollKonto || sollKonto === 'Manuell prüfen')) {
                errors.push(createError(
                    rowNum,
                    CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.KONTO_SOLL + 1,
                    'Soll-Konto nicht zugewiesen',
                    'warning'
                ));
            }

            if (erfasst && (!habenKonto || habenKonto === 'Manuell prüfen')) {
                errors.push(createError(
                    rowNum,
                    CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.KONTO_HABEN + 1,
                    'Haben-Konto nicht zugewiesen',
                    'warning'
                ));
            }
        }

        return errors;
    }

    // Öffentliche API
    return {
        /**
         * Validiert alle Daten im Spreadsheet
         * @param {boolean} showUI - Dialog mit Ergebnissen anzeigen?
         * @returns {Object} Validierungsergebnisse nach Sheets
         */
        validateAllData: function(showUI = true) {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();

                // Ergebnis-Objekt initialisieren
                const results = {
                    einnahmen: [],
                    ausgaben: [],
                    bankbewegungen: [],
                    summary: {
                        totalErrors: 0,
                        totalWarnings: 0
                    }
                };

                // Einnahmen validieren
                const einnahmenSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN);
                if (einnahmenSheet) {
                    const einnahmenData = einnahmenSheet.getDataRange().getValues();
                    results.einnahmen = validateEinnahmen(einnahmenData);
                }

                // Ausgaben validieren
                const ausgabenSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN);
                if (ausgabenSheet) {
                    const ausgabenData = ausgabenSheet.getDataRange().getValues();
                    results.ausgaben = validateAusgaben(ausgabenData);
                }

                // Bankbewegungen validieren
                const bankSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.BANKBEWEGUNGEN);
                if (bankSheet) {
                    const bankData = bankSheet.getDataRange().getValues();
                    results.bankbewegungen = validateBankbewegungen(bankData);
                }

                // Zusammenfassung erstellen
                Object.keys(results).forEach(key => {
                    if (key === 'summary') return;

                    results[key].forEach(error => {
                        if (error.severity === 'error') {
                            results.summary.totalErrors++;
                        } else if (error.severity === 'warning') {
                            results.summary.totalWarnings++;
                        }
                    });
                });

                // Bedingte Formatierung anwenden
                this.applyConditionalFormatting(results);

                // Dialog anzeigen, wenn gewünscht
                if (showUI) {
                    this.showValidationDialog(results);
                }

                return results;
            } catch (e) {
                Logger.log(`Fehler bei der Datenvalidierung: ${e.message}`);
                showToast(`Fehler bei der Datenvalidierung: ${e.message}`, 'Fehler');
                return { error: e.message };
            }
        },

        /**
         * Wendet bedingte Formatierung auf Zellen mit Fehlern an
         * @param {Object} results - Validierungsergebnisse
         */
        applyConditionalFormatting: function(results) {
            const ss = SpreadsheetApp.getActiveSpreadsheet();

            // Für jedes Sheet mit Fehlern
            Object.keys(results).forEach(sheetName => {
                if (sheetName === 'summary') return;

                const errors = results[sheetName];
                if (errors.length === 0) return;

                // Sheet-Namen-Mapping
                const sheetNameMapping = {
                    'einnahmen': CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN,
                    'ausgaben': CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN,
                    'bankbewegungen': CONFIG.SYSTEM.SHEET_NAMES.BANKBEWEGUNGEN
                };

                const sheet = ss.getSheetByName(sheetNameMapping[sheetName]);
                if (!sheet) return;

                // Bestehende bedingte Formatierungen löschen
                sheet.clearConditionalFormatRules();

                // Neue Regeln sammeln
                const errorRules = [];
                const warningRules = [];

                // Für jeden Fehler/Warnung eine Regel erstellen
                errors.forEach(error => {
                    const range = sheet.getRange(error.row, error.col);

                    let rule;
                    if (error.severity === 'error') {
                        // Rot für Fehler
                        rule = SpreadsheetApp.newConditionalFormatRule()
                            .whenFormulaSatisfied('=TRUE')
                            .setBackground('#FFD9D9')
                            .setRanges([range])
                            .build();
                        errorRules.push(rule);
                    } else if (error.severity === 'warning') {
                        // Gelb für Warnungen
                        rule = SpreadsheetApp.newConditionalFormatRule()
                            .whenFormulaSatisfied('=TRUE')
                            .setBackground('#FFF2CC')
                            .setRanges([range])
                            .build();
                        warningRules.push(rule);
                    }
                });

                // Regeln anwenden (erst Fehler, dann Warnungen)
                if (errorRules.length > 0 || warningRules.length > 0) {
                    sheet.setConditionalFormatRules([...errorRules, ...warningRules]);
                }
            });
        },

        /**
         * Zeigt einen Dialog mit den Validierungsergebnissen an
         * @param {Object} results - Validierungsergebnisse
         */
        showValidationDialog: function(results) {
            const ui = SpreadsheetApp.getUi();

            // Zusammenfassung
            const totalErrors = results.summary.totalErrors;
            const totalWarnings = results.summary.totalWarnings;

            if (totalErrors === 0 && totalWarnings === 0) {
                ui.alert(
                    '✅ Validierung erfolgreich',
                    'Alle Daten wurden erfolgreich validiert. Keine Fehler oder Warnungen gefunden.',
                    ui.ButtonSet.OK
                );
                return;
            }

            // HTML für den Dialog erstellen
            let html = '<div style="font-family: Arial, sans-serif; max-height: 400px; overflow-y: auto;">';
            html += `<h3>Validierungsergebnisse</h3>`;
            html += `<p>Gefunden: <b style="color: red;">${totalErrors} Fehler</b> und <b style="color: orange;">${totalWarnings} Warnungen</b></p>`;

            // Details für jedes Sheet
            Object.keys(results).forEach(sheetName => {
                if (sheetName === 'summary') return;

                const errors = results[sheetName];
                if (errors.length === 0) return;

                // Sheet-Namen-Mapping für die Anzeige
                const sheetDisplayName = {
                    'einnahmen': 'Einnahmen',
                    'ausgaben': 'Ausgaben',
                    'bankbewegungen': 'Bankbewegungen'
                };

                html += `<h4>${sheetDisplayName[sheetName]}</h4>`;
                html += '<ul>';

                errors.forEach(error => {
                    const color = error.severity === 'error' ? 'red' : 'orange';
                    html += `<li style="color: ${color};">Zeile ${error.row}, Spalte ${error.col}: ${error.message}</li>`;
                });

                html += '</ul>';
            });

            html += '</div>';

            // Dialog anzeigen
            const htmlOutput = HtmlService
                .createHtmlOutput(html)
                .setWidth(500)
                .setHeight(400);

            ui.showModalDialog(htmlOutput, 'Validierungsergebnisse');
        },

        /**
         * Validiert ein einzelnes Sheet
         * @param {string} sheetName - Name des zu validierenden Sheets
         * @param {boolean} showUI - Dialog mit Ergebnissen anzeigen?
         * @returns {Array} Array mit Fehlermeldungen
         */
        validateSheet: function(sheetName, showUI = true) {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const sheet = ss.getSheetByName(sheetName);

                if (!sheet) {
                    throw new Error(`Sheet "${sheetName}" nicht gefunden`);
                }

                const data = sheet.getDataRange().getValues();
                let errors = [];

                // Je nach Sheet-Typ die passende Validierungsfunktion aufrufen
                switch(sheetName) {
                    case CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN:
                        errors = validateEinnahmen(data);
                        break;
                    case CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN:
                        errors = validateAusgaben(data);
                        break;
                    case CONFIG.SYSTEM.SHEET_NAMES.BANKBEWEGUNGEN:
                        errors = validateBankbewegungen(data);
                        break;
                    default:
                        showToast(`Keine Validierungsfunktion für Sheet "${sheetName}" verfügbar`, 'Info');
                        return [];
                }

                // Bedingte Formatierung anwenden
                const results = { [sheetName]: errors, summary: { totalErrors: 0, totalWarnings: 0 } };

                // Zusammenfassung aktualisieren
                errors.forEach(error => {
                    if (error.severity === 'error') {
                        results.summary.totalErrors++;
                    } else if (error.severity === 'warning') {
                        results.summary.totalWarnings++;
                    }
                });

                this.applyConditionalFormatting(results);

                // Dialog anzeigen, wenn gewünscht
                if (showUI) {
                    this.showValidationDialog(results);
                }

                return errors;
            } catch (e) {
                Logger.log(`Fehler bei der Validierung von ${sheetName}: ${e.message}`);
                showToast(`Fehler bei der Validierung: ${e.message}`, 'Fehler');
                return [];
            }
        }
    };
})();
/**
 * ausgabenModule.js - Modul für die Verwaltung von Ausgaben
 *
 * Verwaltet alle Funktionen für Ausgaben und Rechnungen
 */

const AusgabenModule = (function() {
    // Private Variablen und Funktionen

    /**
     * Berechnet das Quartal für ein Datum
     * @param {Date|string} datum - Datum
     * @returns {number} Quartalszahl (1-4)
     */
    function getQuartal(datum) {
        return AccountingModule.getQuartal(datum);
    }

    /**
     * Berechnet die Mehrwertsteuer für einen Nettobetrag
     * @param {number} nettoBetrag - Nettobetrag
     * @param {number} steuersatz - Steuersatz (z.B. 0.19 für 19%)
     * @returns {number} Mehrwertsteuerbetrag
     */
    function calculateMwSt(nettoBetrag, steuersatz) {
        return Math.round(nettoBetrag * steuersatz * 100) / 100;
    }

    /**
     * Validiert die Ausgaben-Daten und aktualisiert Berechnungen
     * @param {Object} sheet - Sheet-Objekt (Ausgaben)
     */
    function validateAndUpdate(sheet) {
        // Daten laden
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) return; // Nur Header

        // Update-Operationen sammeln
        const updates = [];

        // Header überspringen
        for (let i = 1; i < data.length; i++) {
            const row = data[i];

            // Leere Zeilen überspringen
            if (!row[CONFIG.SYSTEM.AUSGABEN_COLS.DATUM] && !row[CONFIG.SYSTEM.AUSGABEN_COLS.RECHNUNGSNUMMER]) {
                continue;
            }

            // Zeilennummer (1-basiert)
            const rowNum = i + 1;

            // 1. Quartal aktualisieren, falls leer oder falsch
            if (!row[CONFIG.SYSTEM.AUSGABEN_COLS.QUARTAL] && row[CONFIG.SYSTEM.AUSGABEN_COLS.DATUM]) {
                const datum = new Date(row[CONFIG.SYSTEM.AUSGABEN_COLS.DATUM]);
                const quartal = getQuartal(datum);

                updates.push({
                    row: rowNum,
                    col: CONFIG.SYSTEM.AUSGABEN_COLS.QUARTAL + 1,
                    value: quartal
                });
            }

            // 2. MwSt-Betrag berechnen, falls nötig
            const netto = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.NETTO]) || 0;
            const mwstSatz = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.MWST_SATZ]) || 0;
            const mwstBetrag = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.MWST_BETRAG]) || 0;

            const calculatedMwSt = calculateMwSt(netto, mwstSatz);

            if (Math.abs(calculatedMwSt - mwstBetrag) > 0.01) {
                updates.push({
                    row: rowNum,
                    col: CONFIG.SYSTEM.AUSGABEN_COLS.MWST_BETRAG + 1,
                    value: calculatedMwSt
                });
            }

            // 3. Bruttobetrag berechnen, falls nötig
            const brutto = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.BRUTTO]) || 0;
            const calculatedBrutto = netto + (parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.MWST_BETRAG]) || calculatedMwSt);

            if (Math.abs(calculatedBrutto - brutto) > 0.01) {
                updates.push({
                    row: rowNum,
                    col: CONFIG.SYSTEM.AUSGABEN_COLS.BRUTTO + 1,
                    value: calculatedBrutto
                });
            }

            // 4. Restbetrag berechnen, falls nötig
            const bezahlt = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.BEZAHLT]) || 0;
            const restbetrag = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.RESTBETRAG]) || 0;

            const calculatedRestbetrag = Math.max(0, netto - bezahlt);

            if (Math.abs(calculatedRestbetrag - restbetrag) > 0.01) {
                updates.push({
                    row: rowNum,
                    col: CONFIG.SYSTEM.AUSGABEN_COLS.RESTBETRAG + 1,
                    value: calculatedRestbetrag
                });
            }

            // 5. Zahlungsstatus aktualisieren, falls nötig
            const zahlungsstatus = row[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS];
            let newStatus = zahlungsstatus;

            if (calculatedRestbetrag <= 0 && zahlungsstatus !== 'Bezahlt') {
                newStatus = 'Bezahlt';
            } else if (calculatedRestbetrag > 0 && calculatedRestbetrag < netto && zahlungsstatus !== 'Teilweise bezahlt') {
                newStatus = 'Teilweise bezahlt';
            } else if (calculatedRestbetrag === netto && zahlungsstatus !== 'Offen') {
                newStatus = 'Offen';
            }

            if (newStatus !== zahlungsstatus) {
                updates.push({
                    row: rowNum,
                    col: CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS + 1,
                    value: newStatus
                });
            }

            // 6. Letzte Aktualisierung setzen
            updates.push({
                row: rowNum,
                col: CONFIG.SYSTEM.AUSGABEN_COLS.LETZTE_AKTUALISIERUNG + 1,
                value: new Date()
            });
        }

        // Alle Updates in einem Batch durchführen
        updates.forEach(update => {
            sheet.getRange(update.row, update.col).setValue(update.value);
        });

        return updates.length;
    }

    /**
     * Importiert eine Rechnung aus Google Drive und fügt sie ins Ausgaben-Sheet ein
     * @param {string} fileId - Google Drive Datei-ID der Rechnung
     * @returns {Object} Ergebnis des Imports
     */
    function importInvoiceFromDrive(fileId) {
        try {
            // Google Drive Datei abrufen
            const file = DriveApp.getFileById(fileId);
            if (!file) {
                throw new Error(`Datei mit ID ${fileId} nicht gefunden`);
            }

            const fileName = file.getName();
            const fileUrl = file.getUrl();

            // Prüfen, ob die Datei bereits importiert wurde
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN);
            const data = sheet.getDataRange().getValues();

            // Nach dem Dateinamen oder der URL in den vorhandenen Einträgen suchen
            for (let i = 1; i < data.length; i++) {
                if (data[i][CONFIG.SYSTEM.AUSGABEN_COLS.DATEINAME] === fileName ||
                    data[i][CONFIG.SYSTEM.AUSGABEN_COLS.RECHNUNG_LINK] === fileUrl) {
                    return {
                        success: false,
                        message: `Die Rechnung "${fileName}" wurde bereits importiert`,
                        existingRow: i + 1
                    };
                }
            }

            // Neue Zeile für die Rechnung erstellen
            const newRow = Array(sheet.getLastColumn()).fill('');

            // Standardwerte setzen
            newRow[CONFIG.SYSTEM.AUSGABEN_COLS.DATUM] = new Date();
            newRow[CONFIG.SYSTEM.AUSGABEN_COLS.MWST_SATZ] = CONFIG.STEUER.MEHRWERTSTEUER_STANDARD;
            newRow[CONFIG.SYSTEM.AUSGABEN_COLS.QUARTAL] = getQuartal(new Date());
            newRow[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS] = 'Offen';
            newRow[CONFIG.SYSTEM.AUSGABEN_COLS.DATEINAME] = fileName;
            newRow[CONFIG.SYSTEM.AUSGABEN_COLS.RECHNUNG_LINK] = fileUrl;
            newRow[CONFIG.SYSTEM.AUSGABEN_COLS.LETZTE_AKTUALISIERUNG] = new Date();

            // Rechnungsnummer aus Dateinamen extrahieren (falls vorhanden)
            const rechnungsNrMatch = fileName.match(/RE-\d{4}-\d{2}-\d{2}-\d+/);
            if (rechnungsNrMatch) {
                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.RECHNUNGSNUMMER] = rechnungsNrMatch[0];
            }

            // Neue Zeile einfügen
            const lastRow = sheet.getLastRow();
            sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);

            // In Änderungshistorie eintragen
            logFileImport('Ausgabe', fileName, fileUrl);

            return {
                success: true,
                message: `Rechnung "${fileName}" erfolgreich importiert`,
                row: lastRow + 1
            };
        } catch (e) {
            Logger.log(`Fehler beim Import der Rechnung: ${e.message}`);
            return {
                success: false,
                message: `Fehler beim Import: ${e.message}`,
                error: e
            };
        }
    }

    /**
     * Protokolliert den Import einer Datei in der Änderungshistorie
     * @param {string} type - Typ der Rechnung ('Einnahme' oder 'Ausgabe')
     * @param {string} fileName - Name der Datei
     * @param {string} fileUrl - URL der Datei
     */
    function logFileImport(type, fileName, fileUrl) {
        try {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AENDERUNGSHISTORIE);
            if (!sheet) return;

            const newRow = [
                new Date(), // Datum
                type,       // Rechnungstyp
                fileName,   // Dateiname
                fileUrl     // Link zur Datei
            ];

            const lastRow = sheet.getLastRow();
            sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
        } catch (e) {
            Logger.log(`Fehler beim Protokollieren des Imports: ${e.message}`);
        }
    }

    // Öffentliche API
    return {
        /**
         * Aktualisiert das Ausgaben-Sheet
         * - Berechnet fehlende Werte
         * - Aktualisiert Status
         * @returns {number} Anzahl der durchgeführten Updates
         */
        refreshAusgaben: function() {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN);
            if (!sheet) {
                throw new Error(`Sheet "${CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN}" nicht gefunden`);
            }

            const updatesCount = validateAndUpdate(sheet);
            Logger.log(`Ausgaben aktualisiert: ${updatesCount} Änderungen`);

            return updatesCount;
        },

        /**
         * Zeigt einen Dialog zum Importieren einer Rechnung aus Google Drive
         */
        showImportDialog: function() {
            const ui = SpreadsheetApp.getUi();

            const response = ui.prompt(
                'Rechnung importieren',
                'Bitte geben Sie die Google Drive Datei-ID oder URL der Rechnung ein:',
                ui.ButtonSet.OK_CANCEL
            );

            if (response.getSelectedButton() !== ui.Button.OK) {
                return;
            }

            const input = response.getResponseText().trim();
            let fileId;

            // Prüfen, ob es sich um eine URL oder eine ID handelt
            if (input.startsWith('https://')) {
                // URL: ID extrahieren
                const match = input.match(/\/d\/([a-zA-Z0-9_-]+)/);
                if (match && match[1]) {
                    fileId = match[1];
                } else {
                    ui.alert('Fehler', 'Die eingegebene URL konnte nicht verarbeitet werden. Bitte geben Sie eine gültige Google Drive Datei-URL ein.', ui.ButtonSet.OK);
                    return;
                }
            } else {
                // Direkte ID
                fileId = input;
            }

            // Rechnung importieren
            const result = importInvoiceFromDrive(fileId);

            if (result.success) {
                ui.alert('Erfolg', result.message, ui.ButtonSet.OK);

                // Direkt zur importierten Zeile springen
                if (result.row) {
                    SpreadsheetApp.getActiveSpreadsheet()
                        .getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN)
                        .setActiveRange(SpreadsheetApp.getActiveSpreadsheet()
                            .getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN)
                            .getRange(result.row, 1));
                }
            } else {
                ui.alert('Fehler', result.message, ui.ButtonSet.OK);

                // Wenn die Datei bereits existiert, zur entsprechenden Zeile springen
                if (result.existingRow) {
                    SpreadsheetApp.getActiveSpreadsheet()
                        .getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN)
                        .setActiveRange(SpreadsheetApp.getActiveSpreadsheet()
                            .getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN)
                            .getRange(result.existingRow, 1));
                }
            }
        },

        /**
         * Erstellt eine neue Ausgabe/Rechnung
         * @param {Object} data - Rechnungsdaten
         * @returns {Object} Ergebnis der Erstellung
         */
        createAusgabe: function(data) {
            try {
                const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN);
                if (!sheet) {
                    throw new Error(`Sheet "${CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN}" nicht gefunden`);
                }

                // Pflichtfelder prüfen
                if (!data.datum) data.datum = new Date();
                if (!data.lieferant) throw new Error('Lieferant ist ein Pflichtfeld');
                if (!data.netto) throw new Error('Nettobetrag ist ein Pflichtfeld');

                // MwSt-Satz prüfen
                if (!data.mwstSatz && data.mwstSatz !== 0) {
                    data.mwstSatz = CONFIG.STEUER.MEHRWERTSTEUER_STANDARD;
                }

                // Rechnungsnummer generieren, falls nicht vorhanden
                if (!data.rechnungsnummer) {
                    const now = new Date();
                    const dateString = [
                        now.getFullYear(),
                        String(now.getMonth() + 1).padStart(2, '0'),
                        String(now.getDate()).padStart(2, '0')
                    ].join('-');
                    data.rechnungsnummer = `AUSGABE-${dateString}-${Math.floor(Math.random() * 1000)}`;
                }

                // Neue Zeile für die Rechnung erstellen
                const newRow = Array(sheet.getLastColumn()).fill('');

                // Daten setzen
                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.DATUM] = new Date(data.datum);
                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.RECHNUNGSNUMMER] = data.rechnungsnummer;
                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.KATEGORIE] = data.kategorie || 'Wareneinkauf 19% VSt';
                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.LIEFERANT] = data.lieferant;
                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.NETTO] = parseFloat(data.netto);
                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.MWST_SATZ] = parseFloat(data.mwstSatz);

                // Berechnete Felder
                const mwstBetrag = calculateMwSt(data.netto, data.mwstSatz);
                const brutto = parseFloat(data.netto) + mwstBetrag;

                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.MWST_BETRAG] = mwstBetrag;
                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.BRUTTO] = brutto;
                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.BEZAHLT] = data.bezahlt || 0;
                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.RESTBETRAG] = parseFloat(data.netto) - (data.bezahlt || 0);
                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.QUARTAL] = getQuartal(data.datum);

                // Status
                if (data.bezahlt && parseFloat(data.bezahlt) >= parseFloat(data.netto)) {
                    newRow[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS] = 'Bezahlt';
                    newRow[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSART] = data.zahlungsart || 'Überweisung';
                    newRow[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSDATUM] = data.zahlungsdatum || new Date();
                } else if (data.bezahlt && parseFloat(data.bezahlt) > 0) {
                    newRow[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS] = 'Teilweise bezahlt';
                    newRow[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSART] = data.zahlungsart || 'Überweisung';
                    newRow[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSDATUM] = data.zahlungsdatum;
                } else {
                    newRow[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS] = 'Offen';
                }

                // Metadaten
                if (data.dateiname) newRow[CONFIG.SYSTEM.AUSGABEN_COLS.DATEINAME] = data.dateiname;
                if (data.link) newRow[CONFIG.SYSTEM.AUSGABEN_COLS.RECHNUNG_LINK] = data.link;
                newRow[CONFIG.SYSTEM.AUSGABEN_COLS.LETZTE_AKTUALISIERUNG] = new Date();

                // Neue Zeile einfügen
                const lastRow = sheet.getLastRow();
                sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);

                return {
                    success: true,
                    message: `Ausgabe "${data.rechnungsnummer}" erfolgreich erstellt`,
                    row: lastRow + 1,
                    rechnungsnummer: data.rechnungsnummer
                };
            } catch (e) {
                Logger.log(`Fehler beim Erstellen der Ausgabe: ${e.message}`);
                return {
                    success: false,
                    message: `Fehler beim Erstellen: ${e.message}`,
                    error: e
                };
            }
        },

        /**
         * Markiert eine Ausgabe als bezahlt
         * @param {string} rechnungsnummer - Rechnungsnummer
         * @param {Object} paymentData - Zahlungsdaten (Betrag, Datum, Zahlungsart)
         * @returns {Object} Ergebnis der Aktualisierung
         */
        markAsPaid: function(rechnungsnummer, paymentData) {
            try {
                const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN);
                if (!sheet) {
                    throw new Error(`Sheet "${CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN}" nicht gefunden`);
                }

                // Daten laden
                const data = sheet.getDataRange().getValues();

                // Rechnung finden
                let rowIndex = -1;
                let invoiceData;

                for (let i = 1; i < data.length; i++) {
                    if (data[i][CONFIG.SYSTEM.AUSGABEN_COLS.RECHNUNGSNUMMER] === rechnungsnummer) {
                        rowIndex = i;
                        invoiceData = data[i];
                        break;
                    }
                }

                if (rowIndex === -1) {
                    throw new Error(`Rechnung mit Nummer "${rechnungsnummer}" nicht gefunden`);
                }

                // Prüfen, ob die Rechnung bereits vollständig bezahlt ist
                if (invoiceData[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS] === 'Bezahlt') {
                    return {
                        success: false,
                        message: `Die Rechnung "${rechnungsnummer}" ist bereits vollständig bezahlt`,
                        alreadyPaid: true
                    };
                }

                // Zahlungsdaten aktualisieren
                const brutto = parseFloat(invoiceData[CONFIG.SYSTEM.AUSGABEN_COLS.BRUTTO]) || 0;
                const bezahlt = parseFloat(paymentData.betrag || brutto);

                // Zeile aktualisieren (1-basierter Index)
                sheet.getRange(rowIndex + 1, CONFIG.SYSTEM.AUSGABEN_COLS.BEZAHLT + 1).setValue(bezahlt);
                sheet.getRange(rowIndex + 1, CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS + 1).setValue('Bezahlt');
                sheet.getRange(rowIndex + 1, CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSART + 1).setValue(paymentData.zahlungsart || 'Überweisung');
                sheet.getRange(rowIndex + 1, CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSDATUM + 1).setValue(paymentData.datum || new Date());
                sheet.getRange(rowIndex + 1, CONFIG.SYSTEM.AUSGABEN_COLS.RESTBETRAG + 1).setValue(0);
                sheet.getRange(rowIndex + 1, CONFIG.SYSTEM.AUSGABEN_COLS.LETZTE_AKTUALISIERUNG + 1).setValue(new Date());

                return {
                    success: true,
                    message: `Rechnung "${rechnungsnummer}" erfolgreich als bezahlt markiert`,
                    row: rowIndex + 1
                };
            } catch (e) {
                Logger.log(`Fehler beim Markieren als bezahlt: ${e.message}`);
                return {
                    success: false,
                    message: `Fehler beim Markieren als bezahlt: ${e.message}`,
                    error: e
                };
            }
        }
    };
})();
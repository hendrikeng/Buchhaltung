/**
 * bankModule.js - Modul für die Verwaltung von Bankbewegungen
 *
 * Verwaltet alle Funktionen für Bankbewegungen einschließlich
 * automatischer Abgleich mit Rechnungen
 */

const BankModule = (function() {
    // Private Variablen und Funktionen

    /**
     * Prüft, ob zwei Beträge innerhalb einer Toleranzgrenze übereinstimmen
     * @param {number} amount1 - Erster Betrag
     * @param {number} amount2 - Zweiter Betrag
     * @param {number} tolerance - Toleranz (Standard: 0.01 €)
     * @returns {boolean} True, wenn die Beträge innerhalb der Toleranz übereinstimmen
     */
    function amountsMatch(amount1, amount2, tolerance = 0.01) {
        return Math.abs(amount1 - amount2) <= tolerance;
    }

    /**
     * Sucht nach einer Rechnungsnummer im Buchungstext
     * @param {string} buchungstext - Text der Bankbuchung
     * @param {string} rechnungsnummer - Zu suchende Rechnungsnummer
     * @returns {boolean} True, wenn die Rechnungsnummer im Text gefunden wurde
     */
    function containsInvoiceNumber(buchungstext, rechnungsnummer) {
        if (!buchungstext || !rechnungsnummer) return false;

        // Bereinige Eingaben für besseren Vergleich
        const cleanedText = buchungstext.replace(/[\s-]/g, '').toUpperCase();
        const cleanedInvoice = rechnungsnummer.replace(/[\s-]/g, '').toUpperCase();

        return cleanedText.includes(cleanedInvoice);
    }

    /**
     * Berechnet einen Matching-Score zwischen einer Bankbewegung und einer Rechnung
     * @param {Object} bankRow - Daten der Bankbewegung
     * @param {Object} invoice - Daten der Rechnung
     * @returns {number} Score (0-10), wobei höhere Werte bessere Übereinstimmungen anzeigen
     */
    function calculateMatchScore(bankRow, invoice) {
        let score = 0;

        // 1. Exakte Betragsübereinstimmung (höchste Priorität)
        const bankAmount = Math.abs(bankRow[CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.BETRAG]);
        const invoiceAmount = Math.abs(invoice[getInvoiceColumns(invoice).BRUTTO]);

        if (amountsMatch(bankAmount, invoiceAmount)) {
            score += 5; // Hohe Punktzahl für Betragsmatch
        } else {
            // Bei falscher Betragsübereinstimmung frühzeitig aussteigen
            return 0;
        }

        // 2. Rechnungsnummer im Buchungstext
        const buchungstext = bankRow[CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.BUCHUNGSTEXT];
        const rechnungsnummer = invoice[getInvoiceColumns(invoice).RECHNUNGSNUMMER];

        if (containsInvoiceNumber(buchungstext, rechnungsnummer)) {
            score += 4; // Hohe Punktzahl für Rechnungsnummern-Match
        }

        // 3. Zeitliche Nähe zum Rechnungsdatum (optional)
        const bankDatum = new Date(bankRow[CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.DATUM]);
        const rechnungsDatum = new Date(invoice[getInvoiceColumns(invoice).DATUM]);

        // Differenz in Tagen
        const daysDiff = Math.abs((bankDatum - rechnungsDatum) / (1000 * 60 * 60 * 24));

        // Innerhalb von 30 Tagen: 1 Punkt
        if (daysDiff <= 30) {
            score += 1;
        }

        return score;
    }

    /**
     * Ermittelt die passenden Spaltenindizes basierend auf dem Typ der Transaktion
     * @param {Array} transaction - Transaktionsdaten (Einnahme oder Ausgabe)
     * @returns {Object} Objekt mit den relevanten Spaltenindizes
     */
    function getInvoiceColumns(transaction) {
        // Prüfe, ob es sich um Einnahmen oder Ausgaben handelt
        // Wir verwenden einen Heuristik basierend auf der Länge des Arrays und dem Vorhandensein bestimmter Werte

        // Wenn die "Lieferant"-Spalte gefüllt ist, handelt es sich vermutlich um eine Ausgabe
        if (transaction[CONFIG.SYSTEM.AUSGABEN_COLS.LIEFERANT]) {
            return CONFIG.SYSTEM.AUSGABEN_COLS;
        }

        // Ansonsten nehmen wir eine Einnahme an
        return CONFIG.SYSTEM.EINNAHMEN_COLS;
    }

    /**
     * Lädt und filtert Transaktionen (Einnahmen oder Ausgaben) nach Zahlungsstatus
     * @param {string} sheetName - Name des Sheets (Einnahmen oder Ausgaben)
     * @param {Array} statusFilter - Array mit zu berücksichtigenden Zahlungsstatus
     * @returns {Array} Array mit gefilterten Transaktionen
     */
    function loadTransactions(sheetName, statusFilter = ['Offen', 'Teilweise bezahlt']) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) return [];

        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) return []; // Nur Header vorhanden oder leeres Sheet

        // Header entfernen und nach Zahlungsstatus filtern
        return data.slice(1).filter(row => {
            const statusColumn = (sheetName === CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN)
                ? CONFIG.SYSTEM.EINNAHMEN_COLS.ZAHLUNGSSTATUS
                : CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS;

            return statusFilter.includes(row[statusColumn]);
        });
    }

    // Öffentliche API
    return {
        /**
         * Aktualisiert das Bankbewegungen-Sheet
         * Kann zusätzliche Logik wie automatische Kontozuordnung enthalten
         */
        refreshBankbewegungen: function() {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.BANKBEWEGUNGEN);
            if (!sheet) {
                throw new Error(`Sheet "${CONFIG.SYSTEM.SHEET_NAMES.BANKBEWEGUNGEN}" nicht gefunden`);
            }

            // Hier könnte zusätzliche Aktualisierungslogik stehen, wie:
            // - Automatische Kategorisierung
            // - Berechnung von Saldo
            // - Farbliche Hervorhebung unbelegter Buchungen
            // etc.

            Logger.log('Bankbewegungen aktualisiert');
            return true;
        },

        /**
         * Automatische Zuordnung von Bankbewegungen zu Rechnungen
         * Sucht nach unbelegten Bankbewegungen und versucht, passende Rechnungen zu finden
         * @returns {Object} Statistik über die Zuordnungen
         */
        matchBankbewegungToRechnung: function() {
            const ui = SpreadsheetApp.getUi();

            // Bestätigung vom Benutzer einholen
            const response = ui.alert(
                'Bankbewegungen automatisch zuordnen?',
                'Diese Funktion versucht, offene Rechnungen automatisch Bankbewegungen zuzuordnen. Fortfahren?',
                ui.ButtonSet.YES_NO
            );

            if (response !== ui.Button.YES) {
                return {aborted: true};
            }

            try {
                // 1. Bankbewegungen laden
                const bankSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.BANKBEWEGUNGEN);
                const bankData = bankSheet.getDataRange().getValues();

                if (bankData.length <= 1) {
                    showToast('Keine Bankbewegungen gefunden', 'Warnung');
                    return {error: 'Keine Bankbewegungen gefunden'};
                }

                // 2. Einnahmen und Ausgaben laden
                const einnahmen = loadTransactions(CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN);
                const ausgaben = loadTransactions(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN);

                // 3. Statistik initialisieren
                const stats = {
                    totalBankRows: bankData.length - 1,
                    unassigned: 0,
                    matched: {
                        einnahmen: 0,
                        ausgaben: 0
                    }
                };

                // 4. Jede Bankbewegung durchgehen (außer Header)
                for (let i = 1; i < bankData.length; i++) {
                    // Bereits zugeordnete Bankbewegungen überspringen
                    if (bankData[i][CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.ERFASST] === true) {
                        continue;
                    }

                    // Betrag und Typ bestimmen
                    const betrag = bankData[i][CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.BETRAG];
                    const isEinnahme = betrag > 0;

                    // Je nach Typ in Einnahmen oder Ausgaben suchen
                    const matches = [];

                    if (isEinnahme) {
                        einnahmen.forEach(einnahme => {
                            const score = calculateMatchScore(bankData[i], einnahme);
                            if (score >= 5) { // Minimaler Score für Berücksichtigung
                                matches.push({
                                    type: 'einnahme',
                                    data: einnahme,
                                    score: score
                                });
                            }
                        });
                    } else {
                        ausgaben.forEach(ausgabe => {
                            const score = calculateMatchScore(bankData[i], ausgabe);
                            if (score >= 5) { // Minimaler Score für Berücksichtigung
                                matches.push({
                                    type: 'ausgabe',
                                    data: ausgabe,
                                    score: score
                                });
                            }
                        });
                    }

                    // Beste Übereinstimmung finden
                    if (matches.length > 0) {
                        // Nach Score sortieren (höchster zuerst)
                        matches.sort((a, b) => b.score - a.score);
                        const bestMatch = matches[0];

                        // Automatische Zuordnung durchführen
                        const success = this.assignBankbewegungToRechnung(
                            bankData[i],
                            bestMatch.data,
                            bestMatch.type,
                            i, // Bankzeile (0-basiert)
                            bankSheet
                        );

                        if (success) {
                            stats.matched[bestMatch.type + 'n']++; // einnahmen oder ausgaben
                        }
                    } else {
                        stats.unassigned++;
                    }
                }

                // Erfolgsmeldung anzeigen
                const totalMatched = stats.matched.einnahmen + stats.matched.ausgaben;
                showToast(
                    `Automatischer Abgleich abgeschlossen: ${totalMatched} Zuordnungen gefunden ` +
                    `(${stats.matched.einnahmen} Einnahmen, ${stats.matched.ausgaben} Ausgaben)`,
                    'Erfolg'
                );

                return stats;
            } catch (e) {
                Logger.log(`Fehler beim automatischen Abgleich: ${e.message}`);
                showToast(`Fehler beim Abgleich: ${e.message}`, 'Fehler');
                return {error: e.message};
            }
        },

        /**
         * Ordnet eine Bankbewegung einer Rechnung zu und aktualisiert beide Datensätze
         * @param {Array} bankRow - Daten der Bankbewegung
         * @param {Array} invoice - Daten der Rechnung (Einnahme oder Ausgabe)
         * @param {string} type - Typ der Rechnung ('einnahme' oder 'ausgabe')
         * @param {number} bankRowIndex - Index der Bankbewegung im Sheet (0-basiert)
         * @param {Object} bankSheet - Sheet-Objekt der Bankbewegungen
         * @returns {boolean} True bei erfolgreicher Zuordnung
         */
        assignBankbewegungToRechnung: function(bankRow, invoice, type, bankRowIndex, bankSheet) {
            try {
                // 1. Ziel-Sheet bestimmen (Einnahmen oder Ausgaben)
                const targetSheetName = (type === 'einnahme')
                    ? CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN
                    : CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN;

                const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);

                // 2. Rechnungszeile im Sheet finden
                const cols = getInvoiceColumns(invoice);
                const rechnungsnummer = invoice[cols.RECHNUNGSNUMMER];

                // Alle Daten aus dem Sheet laden
                const sheetData = targetSheet.getDataRange().getValues();

                // Zeilenindex der Rechnung finden
                let invoiceRowIndex = -1;
                for (let i = 1; i < sheetData.length; i++) {
                    if (sheetData[i][cols.RECHNUNGSNUMMER] === rechnungsnummer) {
                        invoiceRowIndex = i;
                        break;
                    }
                }

                if (invoiceRowIndex === -1) {
                    throw new Error(`Rechnung ${rechnungsnummer} nicht im Sheet ${targetSheetName} gefunden`);
                }

                // 3. Rechnung als bezahlt markieren
                targetSheet.getRange(invoiceRowIndex + 1, cols.ZAHLUNGSSTATUS + 1).setValue('Bezahlt');
                targetSheet.getRange(invoiceRowIndex + 1, cols.ZAHLUNGSART + 1).setValue('Überweisung');
                targetSheet.getRange(invoiceRowIndex + 1, cols.ZAHLUNGSDATUM + 1).setValue(bankRow[CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.DATUM]);

                // Bezahlten Betrag aktualisieren
                const bruttoBetrag = invoice[cols.BRUTTO];
                targetSheet.getRange(invoiceRowIndex + 1, cols.BEZAHLT + 1).setValue(bruttoBetrag);
                targetSheet.getRange(invoiceRowIndex + 1, cols.RESTBETRAG + 1).setValue(0);

                // 4. Bankbewegung als erfasst markieren
                bankSheet.getRange(bankRowIndex + 1, CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.ERFASST + 1).setValue(true);
                bankSheet.getRange(bankRowIndex + 1, CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.REFERENZ_RECHNUNG + 1).setValue(rechnungsnummer);

                // 5. Kategorie aus der Rechnung übernehmen
                const kategorie = invoice[cols.KATEGORIE];
                if (kategorie) {
                    bankSheet.getRange(bankRowIndex + 1, CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.KATEGORIE + 1).setValue(kategorie);

                    // 6. SKR04-konforme Konten zuweisen basierend auf der Kategorie
                    // Diese Funktion müsste separat implementiert werden
                    const konten = AccountingModule.getSKR04Accounts(kategorie, type === 'einnahme');
                    if (konten && konten.soll && konten.haben) {
                        bankSheet.getRange(bankRowIndex + 1, CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.KONTO_SOLL + 1).setValue(konten.soll);
                        bankSheet.getRange(bankRowIndex + 1, CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.KONTO_HABEN + 1).setValue(konten.haben);
                    }
                }

                // Letzte Aktualisierung setzen
                const now = new Date();
                targetSheet.getRange(invoiceRowIndex + 1, cols.LETZTE_AKTUALISIERUNG + 1).setValue(now);
                bankSheet.getRange(bankRowIndex + 1, CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.LETZTE_AKTUALISIERUNG + 1).setValue(now);

                return true;
            } catch (e) {
                Logger.log(`Fehler bei der Zuordnung: ${e.message}`);
                return false;
            }
        },

        /**
         * Hilfsfunktion für Tests - gibt den Match-Score zwischen einer Bankbewegung und Rechnung zurück
         * @param {Array} bankRow - Daten der Bankbewegung
         * @param {Array} invoice - Daten der Rechnung
         * @returns {number} Berechneter Score
         */
        testMatchScore: function(bankRow, invoice) {
            return calculateMatchScore(bankRow, invoice);
        }
    };
})();
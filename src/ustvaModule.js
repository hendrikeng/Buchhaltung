// file: src/ustvaModule.js
/**
 * ustvaModule.js - Modul für die Umsatzsteuervoranmeldung (UStVA)
 *
 * Berechnet und verwaltet die monatliche/quartalsweise Umsatzsteuervoranmeldung
 * nach den deutschen Steuerregeln mit Berücksichtigung der Ist-Besteuerung
 */

import HelperModule from "./helperModule.js";
import CONFIG from "./config.js";

const UStVAModule = (function() {
    // Private Variablen und Funktionen

    /**
     * Berechnet die maßgeblichen Zeiträume für eine UStVA (Monat oder Quartal)
     * @param {string} period - Zeitraum ("Januar", "Februar", ..., "Q1", "Q2", ...)
     * @param {number} year - Jahr (z.B. 2023)
     * @returns {Object} Start- und Enddatum für den Zeitraum
     */
    function getDateRangeForPeriod(period, year) {
        // Jahr aus der Konfiguration nehmen, wenn nicht angegeben
        const targetYear = year || CONFIG.BENUTZER.AKTUELLES_JAHR;

        // Monatliche Zeiträume
        const months = {
            "Januar": 0, "Februar": 1, "März": 2, "April": 3, "Mai": 4, "Juni": 5,
            "Juli": 6, "August": 7, "September": 8, "Oktober": 9, "November": 10, "Dezember": 11
        };

        // Quartalszeiträume
        const quarters = {
            "Q1": [0, 2], "Q2": [3, 5], "Q3": [6, 8], "Q4": [9, 11],
            "Quartal 1": [0, 2], "Quartal 2": [3, 5], "Quartal 3": [6, 8], "Quartal 4": [9, 11]
        };

        // Jährlicher Zeitraum
        const annual = {
            "Jahr": [0, 11], "Gesamt": [0, 11]
        };

        let startMonth, endMonth;

        if (months[period] !== undefined) {
            // Monatlicher Zeitraum
            startMonth = endMonth = months[period];
        } else if (quarters[period] !== undefined) {
            // Quartalszeitraum
            [startMonth, endMonth] = quarters[period];
        } else if (annual[period] !== undefined) {
            // Jährlicher Zeitraum
            [startMonth, endMonth] = annual[period];
        } else {
            throw new Error(`Ungültiger Zeitraum: ${period}`);
        }

        // Start- und Enddatum erstellen
        const startDate = new Date(targetYear, startMonth, 1);

        // Enddatum ist der letzte Tag des Endmonats
        const endDate = new Date(targetYear, endMonth + 1, 0);

        return {
            startDate: startDate,
            endDate: endDate,
            periodType: months[period] !== undefined ? 'month' :
                quarters[period] !== undefined ? 'quarter' : 'year'
        };
    }

    /**
     * Lädt Transaktionen aus einem Sheet für einen bestimmten Zeitraum
     * @param {string} sheetName - Name des Sheets
     * @param {Date} startDate - Startdatum
     * @param {Date} endDate - Enddatum
     * @param {boolean} istBesteuerung - Bei true nur bezahlte Transaktionen berücksichtigen (Ist-Besteuerung)
     * @returns {Array} Array mit gefilterten Transaktionen
     */
    function loadTransactionsForPeriod(sheetName, startDate, endDate, istBesteuerung = true) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) return [];

        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) return []; // Nur Header vorhanden oder leeres Sheet

        // Header entfernen
        const transactions = data.slice(1);

        // Spaltenindizes bestimmen
        let datumCol, zahlungsstatusCol, zahlungsdatumCol;

        switch(sheetName) {
            case CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN:
                datumCol = CONFIG.SYSTEM.EINNAHMEN_COLS.DATUM;
                zahlungsstatusCol = CONFIG.SYSTEM.EINNAHMEN_COLS.ZAHLUNGSSTATUS;
                zahlungsdatumCol = CONFIG.SYSTEM.EINNAHMEN_COLS.ZAHLUNGSDATUM;
                break;
            case CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN:
                datumCol = CONFIG.SYSTEM.AUSGABEN_COLS.DATUM;
                zahlungsstatusCol = CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS;
                zahlungsdatumCol = CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSDATUM;
                break;
            case CONFIG.SYSTEM.SHEET_NAMES.EIGENBELEGE:
                datumCol = CONFIG.SYSTEM.EIGENBELEGE_COLS.DATUM;
                zahlungsstatusCol = CONFIG.SYSTEM.EIGENBELEGE_COLS.ZAHLUNGSSTATUS;
                zahlungsdatumCol = CONFIG.SYSTEM.EIGENBELEGE_COLS.ZAHLUNGSDATUM;
                break;
            default:
                throw new Error(`Unbekanntes Sheet: ${sheetName}`);
        }

        // Transaktionen filtern
        return transactions.filter(row => {
            // Leere Zeilen überspringen
            if (!row[datumCol]) return false;

            let transactionDate = new Date(row[datumCol]);

            // Bei Ist-Besteuerung: Nur bezahlte Transaktionen berücksichtigen und Zahlungsdatum prüfen
            if (istBesteuerung) {
                // Nur bezahlte Transaktionen berücksichtigen
                if (row[zahlungsstatusCol] !== 'Bezahlt') return false;

                // Zahlungsdatum verwenden statt Rechnungsdatum
                if (row[zahlungsdatumCol]) {
                    transactionDate = new Date(row[zahlungsdatumCol]);
                }
            }

            // Prüfen, ob das Datum im gewünschten Zeitraum liegt
            return transactionDate >= startDate && transactionDate <= endDate;
        });
    }

    /**
     * Kategorisiert Transaktionen nach steuerlichen Kriterien
     * @param {Array} transactions - Array mit Transaktionen
     * @param {string} sheetType - Typ des Sheets ('einnahmen', 'ausgaben', 'eigenbelege')
     * @returns {Object} Kategorisierte Summen für die Steuermeldung
     */
    function categorizeTransactions(transactions, sheetType) {
        // Ergebnis-Objekt initialisieren
        const result = {
            steuerpflichtig: {
                standard: 0,    // 19% MwSt
                reduziert: 0    // 7% MwSt
            },
            steuerfrei: {
                inland: 0,      // Steuerfreie Umsätze im Inland
                ausland: {
                    eu: 0,        // Innergemeinschaftliche Lieferungen
                    nonEu: 0      // Ausfuhrlieferungen
                }
            },
            mwst: {
                standard: 0,    // 19% MwSt-Betrag
                reduziert: 0    // 7% MwSt-Betrag
            },
            total: 0          // Gesamtsumme (für Prüfzwecke)
        };

        // Spaltenindizes je nach Sheet-Typ
        let kategorieCol, nettoCol, mwstSatzCol, mwstBetragCol;

        switch(sheetType.toLowerCase()) {
            case 'einnahmen':
                kategorieCol = CONFIG.SYSTEM.EINNAHMEN_COLS.KATEGORIE;
                nettoCol = CONFIG.SYSTEM.EINNAHMEN_COLS.NETTO;
                mwstSatzCol = CONFIG.SYSTEM.EINNAHMEN_COLS.MWST_SATZ;
                mwstBetragCol = CONFIG.SYSTEM.EINNAHMEN_COLS.MWST_BETRAG;
                break;
            case 'ausgaben':
                kategorieCol = CONFIG.SYSTEM.AUSGABEN_COLS.KATEGORIE;
                nettoCol = CONFIG.SYSTEM.AUSGABEN_COLS.NETTO;
                mwstSatzCol = CONFIG.SYSTEM.AUSGABEN_COLS.MWST_SATZ;
                mwstBetragCol = CONFIG.SYSTEM.AUSGABEN_COLS.MWST_BETRAG;
                break;
            case 'eigenbelege':
                kategorieCol = CONFIG.SYSTEM.EIGENBELEGE_COLS.KATEGORIE;
                nettoCol = CONFIG.SYSTEM.EIGENBELEGE_COLS.NETTO;
                mwstSatzCol = CONFIG.SYSTEM.EIGENBELEGE_COLS.MWST_SATZ;
                mwstBetragCol = CONFIG.SYSTEM.EIGENBELEGE_COLS.MWST_BETRAG;
                break;
            default:
                throw new Error(`Unbekannter Sheet-Typ: ${sheetType}`);
        }

        // Transaktionen durchgehen und kategorisieren
        transactions.forEach(row => {
            // Leere Zeilen oder Zeilen ohne Betrag überspringen
            if (!row[nettoCol]) return;

            const kategorie = row[kategorieCol] || '';
            const nettoBetrag = parseFloat(row[nettoCol]) || 0;
            const mwstSatz = parseFloat(row[mwstSatzCol]) || 0;
            const mwstBetrag = parseFloat(row[mwstBetragCol]) || 0;

            // Gesamtsumme aktualisieren
            result.total += nettoBetrag;

            // Steuerliche Kategorisierung
            if (isEUTaxFree(kategorie)) {
                // Innergemeinschaftliche Lieferung (§4 Nr. 1b UStG)
                result.steuerfrei.ausland.eu += nettoBetrag;
            } else if (isNonEUTaxFree(kategorie)) {
                // Ausfuhrlieferung (§4 Nr. 1a UStG)
                result.steuerfrei.ausland.nonEu += nettoBetrag;
            } else if (isDomesticTaxFree(kategorie)) {
                // Steuerfreie Inlandsleistung (§4 Nr. 8-28 UStG)
                result.steuerfrei.inland += nettoBetrag;
            } else if (Math.abs(mwstSatz - CONFIG.STEUER.MEHRWERTSTEUER_REDUZIERT) < 0.01) {
                // 7% MwSt
                result.steuerpflichtig.reduziert += nettoBetrag;
                result.mwst.reduziert += mwstBetrag;
            } else {
                // 19% MwSt (Standard) oder andere Steuersätze
                result.steuerpflichtig.standard += nettoBetrag;
                result.mwst.standard += mwstBetrag;
            }
        });

        return result;
    }

    /**
     * Prüft, ob eine Kategorie innergemeinschaftliche Lieferungen betrifft
     * @param {string} kategorie - Kategoriename
     * @returns {boolean} True, wenn es sich um innergemeinschaftliche Lieferungen handelt
     */
    function isEUTaxFree(kategorie) {
        return kategorie.includes('innergemeinschaftlich') ||
            kategorie.includes('§4 Nr. 1b') ||
            kategorie.includes('EU-Lieferung');
    }

    /**
     * Prüft, ob eine Kategorie Ausfuhrlieferungen betrifft
     * @param {string} kategorie - Kategoriename
     * @returns {boolean} True, wenn es sich um Ausfuhrlieferungen handelt
     */
    function isNonEUTaxFree(kategorie) {
        return kategorie.includes('Ausfuhrlieferung') ||
            kategorie.includes('§4 Nr. 1a') ||
            kategorie.includes('Export');
    }

    /**
     * Prüft, ob eine Kategorie steuerfreie Inlandsleistungen betrifft
     * @param {string} kategorie - Kategoriename
     * @returns {boolean} True, wenn es sich um steuerfreie Inlandsleistungen handelt
     */
    function isDomesticTaxFree(kategorie) {
        return kategorie.includes('steuerfrei') &&
            !isEUTaxFree(kategorie) &&
            !isNonEUTaxFree(kategorie);
    }

    /**
     * Zeigt eine Benachrichtigung als Toast an
     * @param {string} message - Nachricht
     * @param {string} title - Titel (optional)
     */
    function showToast(message, title = 'Info') {
        if (typeof HelperModule.showToast === 'function') {
            HelperModule.showToast(message, title);
        } else {
            SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
        }
    }

    // Öffentliche API
    return {
        /**
         * Berechnet die Umsatzsteuervoranmeldung für einen bestimmten Zeitraum
         * @param {string} period - Zeitraum ("Januar", "Februar", "Q1", etc.)
         * @param {number} year - Jahr (wenn nicht angegeben, wird das aktuelle Jahr verwendet)
         * @param {boolean} writeToSheet - Ergebnisse ins UStVA-Sheet schreiben?
         * @returns {Object} UStVA-Ergebnisse
         */
        calculateUStVA: function(period, year, writeToSheet = true) {
            try {
                const ui = SpreadsheetApp.getUi();

                // Zeitraum auswählen lassen, wenn nicht angegeben
                if (!period) {
                    const response = ui.prompt(
                        'UStVA berechnen',
                        'Bitte geben Sie den Zeitraum ein (z.B. "Januar", "Q1", "Gesamt"):',
                        ui.ButtonSet.OK_CANCEL
                    );

                    if (response.getSelectedButton() !== ui.Button.OK) {
                        return null;
                    }

                    period = response.getResponseText().trim();
                }

                // Jahr verwenden oder auswählen lassen
                if (!year) {
                    year = CONFIG.BENUTZER.AKTUELLES_JAHR;
                }

                // Zeitraum berechnen
                const { startDate, endDate, periodType } = getDateRangeForPeriod(period, year);

                // Status-Meldung
                showToast(`Berechne UStVA für ${period} ${year}...`, 'Info');

                // Daten laden (mit Ist-Besteuerung)
                const einnahmen = loadTransactionsForPeriod(
                    CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN, startDate, endDate, true
                );

                const ausgaben = loadTransactionsForPeriod(
                    CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN, startDate, endDate, true
                );

                const eigenbelege = loadTransactionsForPeriod(
                    CONFIG.SYSTEM.SHEET_NAMES.EIGENBELEGE, startDate, endDate, true
                );

                // Daten kategorisieren
                const einnahmenStats = categorizeTransactions(einnahmen, 'einnahmen');
                const ausgabenStats = categorizeTransactions(ausgaben, 'ausgaben');
                const eigenbelageStats = categorizeTransactions(eigenbelege, 'eigenbelege');

                // Bewirtungskosten speziell behandeln (nur 70% der Vorsteuer abzugsfähig)
                let nichtAbzugsfaehigeBewirtung = 0;

                ausgaben.forEach(row => {
                    const kategorie = row[CONFIG.SYSTEM.AUSGABEN_COLS.KATEGORIE] || '';
                    if (kategorie.toLowerCase().includes('bewirtung')) {
                        const mwstBetrag = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.MWST_BETRAG]) || 0;
                        nichtAbzugsfaehigeBewirtung += mwstBetrag * (1 - CONFIG.STEUER.BEWIRTUNG_ABZUGSFAEHIG);
                    }
                });

                // UStVA-Ergebnis zusammenstellen
                const result = {
                    zeitraum: period,
                    jahr: year,

                    // Einnahmen
                    einnahmen: {
                        steuerpflichtig: einnahmenStats.steuerpflichtig.standard + einnahmenStats.steuerpflichtig.reduziert,
                        steuerfrei: {
                            inland: einnahmenStats.steuerfrei.inland,
                            ausland: einnahmenStats.steuerfrei.ausland.eu + einnahmenStats.steuerfrei.ausland.nonEu
                        },
                        ust: {
                            standard: einnahmenStats.mwst.standard,
                            reduziert: einnahmenStats.mwst.reduziert
                        }
                    },

                    // Ausgaben
                    ausgaben: {
                        steuerpflichtig: ausgabenStats.steuerpflichtig.standard + ausgabenStats.steuerpflichtig.reduziert,
                        steuerfrei: {
                            inland: ausgabenStats.steuerfrei.inland,
                            ausland: ausgabenStats.steuerfrei.ausland.eu + ausgabenStats.steuerfrei.ausland.nonEu
                        },
                        vorsteuer: {
                            standard: ausgabenStats.mwst.standard,
                            reduziert: ausgabenStats.mwst.reduziert
                        }
                    },

                    // Eigenbelege
                    eigenbelege: {
                        steuerpflichtig: eigenbelageStats.steuerpflichtig.standard + eigenbelageStats.steuerpflichtig.reduziert,
                        steuerfrei: eigenbelageStats.steuerfrei.inland + eigenbelageStats.steuerfrei.ausland.eu + eigenbelageStats.steuerfrei.ausland.nonEu
                    },

                    // Bewirtung
                    nichtAbzugsfaehigeBewirtung: nichtAbzugsfaehigeBewirtung,

                    // Gesamtergebnis
                    ustZahlung: (einnahmenStats.mwst.standard + einnahmenStats.mwst.reduziert) -
                        (ausgabenStats.mwst.standard + ausgabenStats.mwst.reduziert - nichtAbzugsfaehigeBewirtung)
                };

                // In Sheet schreiben, wenn gewünscht
                if (writeToSheet) {
                    this.writeUStVAToSheet(result);
                }

                // Erfolgsbenachrichtigung
                showToast(`UStVA für ${period} ${year} erfolgreich berechnet`, 'Erfolg');

                return result;
            } catch (e) {
                Logger.log(`Fehler bei der UStVA-Berechnung: ${e.message}`);
                showToast(`Fehler bei der UStVA-Berechnung: ${e.message}`, 'Fehler');
                throw e;
            }
        },

        /**
         * Schreibt die UStVA-Ergebnisse in das UStVA-Sheet
         * @param {Object} ustvaData - UStVA-Ergebnisse
         */
        writeUStVAToSheet: function(ustvaData) {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.USTVA);
            if (!sheet) {
                throw new Error(`Sheet "${CONFIG.SYSTEM.SHEET_NAMES.USTVA}" nicht gefunden`);
            }

            // Vorhandene Daten laden
            const data = sheet.getDataRange().getValues();

            // Passende Zeile finden oder neue Zeile hinzufügen
            let rowIndex = -1;

            for (let i = 1; i < data.length; i++) {
                if (data[i][CONFIG.SYSTEM.USTVA_COLS.ZEITRAUM] === ustvaData.zeitraum) {
                    rowIndex = i + 1; // 1-basierter Index für Sheets
                    break;
                }
            }

            // Neue Zeile hinzufügen, wenn keine passende gefunden wurde
            if (rowIndex === -1) {
                rowIndex = data.length + 1;
            }

            // UStVA-Daten in das Sheet schreiben
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.ZEITRAUM + 1).setValue(ustvaData.zeitraum);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.STEUERPFLICHTIGE_EINNAHMEN + 1).setValue(ustvaData.einnahmen.steuerpflichtig);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.STEUERFREIE_INLAND_EINNAHMEN + 1).setValue(ustvaData.einnahmen.steuerfrei.inland);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.STEUERFREIE_AUSLAND_EINNAHMEN + 1).setValue(ustvaData.einnahmen.steuerfrei.ausland);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.STEUERPFLICHTIGE_AUSGABEN + 1).setValue(ustvaData.ausgaben.steuerpflichtig);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.STEUERFREIE_INLAND_AUSGABEN + 1).setValue(ustvaData.ausgaben.steuerfrei.inland);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.STEUERFREIE_AUSLAND_AUSGABEN + 1).setValue(ustvaData.ausgaben.steuerfrei.ausland);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.EIGENBELEGE_STEUERPFLICHTIG + 1).setValue(ustvaData.eigenbelege.steuerpflichtig);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.EIGENBELEGE_STEUERFREI + 1).setValue(ustvaData.eigenbelege.steuerfrei);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.NICHT_ABZUGSFAEHIGE_VST_BEWIRTUNG + 1).setValue(ustvaData.nichtAbzugsfaehigeBewirtung);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.UST_7 + 1).setValue(ustvaData.einnahmen.ust.reduziert);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.UST_19 + 1).setValue(ustvaData.einnahmen.ust.standard);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.VST_7 + 1).setValue(ustvaData.ausgaben.vorsteuer.reduziert);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.VST_19 + 1).setValue(ustvaData.ausgaben.vorsteuer.standard);
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.UST_ZAHLUNG + 1).setValue(ustvaData.ustZahlung);

            // Ergebnis (kann für spezielle Berechnungen angepasst werden)
            const ergebnis = ustvaData.einnahmen.steuerpflichtig - ustvaData.ausgaben.steuerpflichtig;
            sheet.getRange(rowIndex, CONFIG.SYSTEM.USTVA_COLS.ERGEBNIS + 1).setValue(ergebnis);
        },

        /**
         * Aktualisiert das UStVA-Sheet
         * - Berechnet alle fehlenden Zeiträume neu
         * - Aktualisiert Quartals- und Jahressummen
         */
        refreshUStVA: function() {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.USTVA);
            if (!sheet) {
                throw new Error(`Sheet "${CONFIG.SYSTEM.SHEET_NAMES.USTVA}" nicht gefunden`);
            }

            // Vorhandene Daten laden
            const data = sheet.getDataRange().getValues();
            if (data.length <= 1) {
                // Nur Header vorhanden, alle Monate berechnen
                this.calculateAllPeriods();
                return;
            }

            // Prüfen, welche Zeiträume vorhanden sind
            const existingPeriods = {};
            for (let i = 1; i < data.length; i++) {
                const period = data[i][CONFIG.SYSTEM.USTVA_COLS.ZEITRAUM];
                if (period) {
                    existingPeriods[period] = true;
                }
            }

            // Fehlende Monate berechnen
            const months = ["Januar", "Februar", "März", "April", "Mai", "Juni",
                "Juli", "August", "September", "Oktober", "November", "Dezember"];

            const currentYear = CONFIG.BENUTZER.AKTUELLES_JAHR;
            const currentMonth = new Date().getMonth(); // 0-basiert

            for (let i = 0; i <= currentMonth; i++) {
                const month = months[i];
                if (!existingPeriods[month]) {
                    this.calculateUStVA(month, currentYear);
                }
            }

            // Quartals- und Jahressummen aktualisieren
            this.updateQuarterlyAndAnnualTotals();

            showToast('UStVA-Sheet erfolgreich aktualisiert', 'Erfolg');
        },

        /**
         * Berechnet die UStVA für alle Monate des aktuellen Jahres
         */
        calculateAllPeriods: function() {
            const months = ["Januar", "Februar", "März", "April", "Mai", "Juni",
                "Juli", "August", "September", "Oktober", "November", "Dezember"];

            const currentYear = CONFIG.BENUTZER.AKTUELLES_JAHR;
            const currentMonth = new Date().getMonth(); // 0-basiert

            // Status-Meldung
            showToast(`Berechne UStVA für alle Monate ${currentYear}...`, 'Info');

            // Alle vergangenen Monate des aktuellen Jahres berechnen
            for (let i = 0; i <= currentMonth; i++) {
                this.calculateUStVA(months[i], currentYear);
            }

            // Quartals- und Jahressummen aktualisieren
            this.updateQuarterlyAndAnnualTotals();

            showToast(`UStVA für alle Monate ${currentYear} erfolgreich berechnet`, 'Erfolg');
        },

        /**
         * Schreibt Quartals- oder Jahressummen ins UStVA-Sheet
         * @param {Object} data - Daten für das Quartal oder Jahr
         */
        writeQuarterOrAnnualData: function(data) {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.USTVA);

            // Vorhandene Daten laden
            const sheetData = sheet.getDataRange().getValues();

            // Passende Zeile finden oder neue Zeile hinzufügen
            let rowIndex = -1;

            for (let i = 1; i < sheetData.length; i++) {
                if (sheetData[i][0] === data[0]) {
                    rowIndex = i + 1; // 1-basierter Index für Sheets
                    break;
                }
            }

            // Neue Zeile hinzufügen, wenn keine passende gefunden wurde
            if (rowIndex === -1) {
                rowIndex = sheetData.length + 1;
            }

            // Daten ins Sheet schreiben
            const range = sheet.getRange(rowIndex, 1, 1, Object.keys(data).length);
            const values = [];
            const row = [];

            for (let col = 0; col < Object.keys(data).length; col++) {
                row.push(data[col]);
            }

            values.push(row);
            range.setValues(values);
        },

        /**
         * Aktualisiert Quartals- und Jahressummen im UStVA-Sheet
         */
        updateQuarterlyAndAnnualTotals: function() {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.USTVA);
            if (!sheet) return;

            const data = sheet.getDataRange().getValues();
            if (data.length <= 1) return; // Nur Header

            const months = ["Januar", "Februar", "März", "April", "Mai", "Juni",
                "Juli", "August", "September", "Oktober", "November", "Dezember"];

            // Daten für jeden Monat sammeln
            const monthlyData = {};
            for (let i = 1; i < data.length; i++) {
                const period = data[i][CONFIG.SYSTEM.USTVA_COLS.ZEITRAUM];
                if (months.includes(period)) {
                    monthlyData[period] = {};

                    // Alle relevanten Spalten kopieren
                    for (let col = 1; col < data[i].length; col++) {
                        monthlyData[period][col] = data[i][col];
                    }
                }
            }

            // Quartalsdaten berechnen
            const quarters = [
                { name: "Q1", months: ["Januar", "Februar", "März"] },
                { name: "Q2", months: ["April", "Mai", "Juni"] },
                { name: "Q3", months: ["Juli", "August", "September"] },
                { name: "Q4", months: ["Oktober", "November", "Dezember"] }
            ];

            quarters.forEach(quarter => {
                // Prüfen, ob alle Monate des Quartals vorhanden sind
                const availableMonths = quarter.months.filter(m => monthlyData[m]);
                if (availableMonths.length === 0) return; // Keine Daten für dieses Quartal

                // Quartalsdaten ins Sheet schreiben
                this.writeQuarterOrAnnualData(quarterData);
            });

            // Jahressumme berechnen
            const availableMonths = Object.keys(monthlyData);
            if (availableMonths.length === 0) return; // Keine Daten

            const annualData = { 0: "Gesamt" }; // Zeitraum

            // Alle anderen Spalten summieren (außer Zeitraum)
            for (let col = 1; col < data[0].length; col++) {
                annualData[col] = 0;
                availableMonths.forEach(month => {
                    annualData[col] += (monthlyData[month][col] || 0);
                });
            }

            // Jahressumme ins Sheet schreiben
            this.writeQuarterOrAnnualData(annualData);
        },

        /**
         * Erstellt einen ELSTER-Export für die UStVA
         * @param {string} period - Zeitraum ("Januar", "Februar", "Q1", etc.)
         * @param {number} year - Jahr
         */
        createElsterExport: function(period, year) {
            // Platzhalter für zukünftige ELSTER-Export-Funktion
            const ui = SpreadsheetApp.getUi();

            ui.alert(
                'ELSTER-Export',
                'Diese Funktion ist noch in Entwicklung. In Zukunft wird hier ein Export für ELSTER erstellt.',
                ui.ButtonSet.OK
            );

            // TODO: Implementieren Sie hier den ELSTER-Export
            // Mögliche Formate: CSV, XML oder spezielle ELSTER-Formate
        }ssummen berechnen
        const quarterData = { 0: quarter.name }; // Zeitraum

        // Alle anderen Spalten summieren (außer Zeitraum)
        for (let col = 1; col < data[0].length; col++) {
        quarterData[col] = 0;
        availableMonths.forEach(month => {
            quarterData[col] += (monthlyData[month][col] || 0);
        });
    }

// Quartal
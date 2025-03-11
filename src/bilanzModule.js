/**
 * bilanzModule.js - Modul für die Erstellung der Bilanz
 *
 * Erstellt eine einfache Bilanz basierend auf den vorhandenen Daten
 * aus den verschiedenen Sheets (Bankbewegungen, BWA, etc.)
 */

const BilanzModule = (function() {
    // Private Variablen und Funktionen

    /**
     * Ermittelt das Bankguthaben aus dem Bankbewegungen-Sheet
     * @returns {number} Bankguthaben (Endsaldo) oder 0 bei Fehler
     */
    function getBankGuthaben() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bankSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.BANKBEWEGUNGEN);

        if (!bankSheet) return 0;

        const lastRow = bankSheet.getLastRow();
        if (lastRow < 2) return 0;

        const data = bankSheet.getRange(1, 1, lastRow, bankSheet.getLastColumn()).getValues();

        // Nach "Endsaldo" in der zweiten Spalte suchen
        for (let i = lastRow - 1; i >= 0; i--) {
            const label = data[i][1]?.toString().toLowerCase() || "";
            if (label === "endsaldo") {
                const saldo = parseFloat(data[i][3]) || 0;
                return saldo;
            }
        }

        // Alternativ: Falls kein expliziter Endsaldo vorhanden ist,
        // nehmen wir den letzten Saldo-Eintrag
        return parseFloat(data[lastRow - 1][3]) || 0;
    }

    /**
     * Ermittelt den Jahresüberschuss aus dem BWA-Sheet
     * @returns {number} Jahresüberschuss oder 0 bei Fehler
     */
    function getJahresUeberschuss() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bwaSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.BWA);

        if (!bwaSheet) return 0;

        const data = bwaSheet.getDataRange().getValues();
        if (data.length < 2) return 0;

        // Nach "Jahresüberschuss" in der ersten Spalte suchen
        for (let i = data.length - 1; i >= 0; i--) {
            const label = data[i][0]?.toString().toLowerCase() || "";
            if (label.includes("jahresüberschuss") || label.includes("jahresueberschuss")) {
                // Jahreswert (letzte Spalte)
                return parseFloat(data[i][data[i].length - 1]) || 0;
            }
        }

        return 0;
    }

    /**
     * Ermittelt offene Forderungen aus Einnahmen
     * @returns {number} Summe der offenen Forderungen
     */
    function getOpenReceivables() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN);

        if (!sheet) return 0;

        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) return 0; // Nur Header vorhanden

        let sumReceivables = 0;

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const status = row[CONFIG.SYSTEM.EINNAHMEN_COLS.ZAHLUNGSSTATUS];

            // Offene oder teilweise bezahlte Rechnungen
            if (status === "Offen" || status === "Teilweise bezahlt") {
                const restbetrag = parseFloat(row[CONFIG.SYSTEM.EINNAHMEN_COLS.RESTBETRAG]) || 0;
                sumReceivables += restbetrag;
            }
        }

        return sumReceivables;
    }

    /**
     * Ermittelt offene Verbindlichkeiten aus Ausgaben
     * @returns {number} Summe der offenen Verbindlichkeiten
     */
    function getOpenPayables() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN);

        if (!sheet) return 0;

        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) return 0; // Nur Header vorhanden

        let sumPayables = 0;

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const status = row[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS];

            // Offene oder teilweise bezahlte Rechnungen
            if (status === "Offen" || status === "Teilweise bezahlt") {
                const restbetrag = parseFloat(row[CONFIG.SYSTEM.AUSGABEN_COLS.RESTBETRAG]) || 0;
                sumPayables += restbetrag;
            }
        }

        return sumPayables;
    }

    /**
     * Ermittelt Gesellschafterdarlehen aus dem Gesellschafterkonto
     * @returns {number} Summe der Gesellschafterdarlehen
     */
    function getShareholderLoans() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.GESELLSCHAFTERKONTO);

        if (!sheet) return 0;

        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) return 0; // Nur Header vorhanden

        let sumLoans = 0;

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const art = row[CONFIG.SYSTEM.GESELLSCHAFTERKONTO_COLS.ART];

            if (art === "Darlehen") {
                const betrag = parseFloat(row[CONFIG.SYSTEM.GESELLSCHAFTERKONTO_COLS.BETRAG]) || 0;
                sumLoans += betrag;
            } else if (art === "Rückzahlung") {
                const betrag = parseFloat(row[CONFIG.SYSTEM.GESELLSCHAFTERKONTO_COLS.BETRAG]) || 0;
                sumLoans -= Math.abs(betrag);
            }
        }

        return sumLoans;
    }

    /**
     * Berechnet eine einfache Bilanz basierend auf den verfügbaren Daten
     * @returns {Object} Bilanzdaten (Aktiva und Passiva)
     */
    function calculateBilanzData() {
        // Werte ermitteln
        const bankGuthaben = getBankGuthaben();
        const jahresUeberschuss = getJahresUeberschuss();
        const openReceivables = getOpenReceivables();
        const openPayables = getOpenPayables();
        const shareholderLoans = getShareholderLoans();

        // Aktiva-Positionen (Vermögenswerte)
        const aktiva = [
            {
                label: "Aktiva (Vermögenswerte)",
                value: ""
            },
            {
                label: "1.1 Anlagevermögen",
                value: ""
            },
            {
                label: "Sachanlagen",
                value: 0  // Hier könnte man Werte aus einem Asset-Register einlesen
            },
            {
                label: "Immaterielle Vermögenswerte",
                value: 0
            },
            {
                label: "Finanzanlagen",
                value: 0
            },
            {
                label: "Zwischensumme Anlagevermögen",
                value: "=SUM(B3:B5)"
            },
            {
                label: "",
                value: ""
            },
            {
                label: "1.2 Umlaufvermögen",
                value: ""
            },
            {
                label: "Bankguthaben",
                value: bankGuthaben
            },
            {
                label: "Kasse",
                value: 0
            },
            {
                label: "Forderungen aus L&L",
                value: openReceivables
            },
            {
                label: "Vorräte",
                value: 0
            },
            {
                label: "Zwischensumme Umlaufvermögen",
                value: "=SUM(B9:B12)"
            },
            {
                label: "",
                value: ""
            },
            {
                label: "Gesamt Aktiva",
                value: "=B6+B13"
            }
        ];

        // Passiva-Positionen (Finanzierung)
        const passiva = [
            {
                label: "Passiva (Finanzierung & Schulden)",
                value: ""
            },
            {
                label: "2.1 Eigenkapital",
                value: ""
            },
            {
                label: "Stammkapital",
                value: CONFIG.STEUER?.stammkapital || 25000
            },
            {
                label: "Gewinn-/Verlustvortrag",
                value: 0  // Hier könnte der Vortrag aus dem Vorjahr eingelesen werden
            },
            {
                label: "Jahresüberschuss/-fehlbetrag",
                value: jahresUeberschuss
            },
            {
                label: "Zwischensumme Eigenkapital",
                value: "=SUM(F3:F5)"
            },
            {
                label: "",
                value: ""
            },
            {
                label: "2.2 Verbindlichkeiten",
                value: ""
            },
            {
                label: "Bankdarlehen",
                value: 0
            },
            {
                label: "Gesellschafterdarlehen",
                value: shareholderLoans
            },
            {
                label: "Verbindlichkeiten aus L&L",
                value: openPayables
            },
            {
                label: "Steuerrückstellungen",
                value: 0
            },
            {
                label: "Zwischensumme Verbindlichkeiten",
                value: "=SUM(F9:F12)"
            },
            {
                label: "",
                value: ""
            },
            {
                label: "Gesamt Passiva",
                value: "=F6+F13"
            }
        ];

        return { aktiva, passiva };
    }

    /**
     * Formatiert die Bilanz für eine bessere Darstellung
     * @param {Object} sheet - Das Bilanz-Sheet
     */
    function formatBilanzSheet(sheet) {
        // Spaltenbreiten anpassen
        sheet.autoResizeColumns(1, 6);

        // Überschriften formatieren
        sheet.getRange(1, 1).setFontWeight('bold');
        sheet.getRange(1, 5).setFontWeight('bold');
        sheet.getRange(2, 1).setFontWeight('bold');
        sheet.getRange(2, 5).setFontWeight('bold');
        sheet.getRange(8, 1).setFontWeight('bold');
        sheet.getRange(8, 5).setFontWeight('bold');
        sheet.getRange(15, 1).setFontWeight('bold');
        sheet.getRange(15, 5).setFontWeight('bold');

        // Zahlenformat für Beträge
        sheet.getRange("B3:B5").setNumberFormat("#,##0.00 €");
        sheet.getRange("B6").setNumberFormat("#,##0.00 €");
        sheet.getRange("B9:B12").setNumberFormat("#,##0.00 €");
        sheet.getRange("B13").setNumberFormat("#,##0.00 €");
        sheet.getRange("B15").setNumberFormat("#,##0.00 €");

        sheet.getRange("F3:F5").setNumberFormat("#,##0.00 €");
        sheet.getRange("F6").setNumberFormat("#,##0.00 €");
        sheet.getRange("F9:F12").setNumberFormat("#,##0.00 €");
        sheet.getRange("F13").setNumberFormat("#,##0.00 €");
        sheet.getRange("F15").setNumberFormat("#,##0.00 €");

        // Zwischensummen und Gesamtsummen fett formatieren
        sheet.getRange("A6:B6").setFontWeight("bold");
        sheet.getRange("A13:B13").setFontWeight("bold");
        sheet.getRange("E6:F6").setFontWeight("bold");
        sheet.getRange("E13:F13").setFontWeight("bold");

        // Hintergrundfarben für bessere Übersicht
        sheet.getRange("A1:B1").setBackground("#D9EAD3");
        sheet.getRange("E1:F1").setBackground("#D9EAD3");
        sheet.getRange("A15:B15").setBackground("#D9EAD3");
        sheet.getRange("E15:F15").setBackground("#D9EAD3");
    }

    // Öffentliche API
    return {
        /**
         * Erstellt die Bilanz und schreibt sie ins Bilanz-Sheet
         * @returns {boolean} true bei Erfolg, false bei Fehler
         */
        generateBilanz: function() {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();

                // Bilanzdaten berechnen
                const { aktiva, passiva } = calculateBilanzData();

                // Bilanz-Sheet erstellen oder aktualisieren
                const bilanzSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.BILANZ) ||
                    ss.insertSheet(CONFIG.SYSTEM.SHEET_NAMES.BILANZ);
                bilanzSheet.clearContents();

                // Aktiva und Passiva ins Sheet schreiben
                const aktivaValues = aktiva.map(item => [item.label, item.value]);
                const passivaValues = passiva.map(item => [item.label, item.value]);

                bilanzSheet.getRange(1, 1, aktivaValues.length, 2).setValues(aktivaValues);
                bilanzSheet.getRange(1, 5, passivaValues.length, 2).setValues(passivaValues);

                // Formatierung anwenden
                formatBilanzSheet(bilanzSheet);

                // Erfolgsmeldung
                SpreadsheetApp.getUi().alert("Bilanz wurde erstellt!");

                return true;
            } catch (e) {
                Logger.log(`Fehler bei der Bilanzerstellung: ${e.message}`);
                SpreadsheetApp.getUi().alert("Fehler bei der Bilanzerstellung: " + e.message);
                return false;
            }
        },

        /**
         * Aktualisiert das Bilanz-Sheet
         * @returns {boolean} true bei Erfolg, false bei Fehler
         */
        refreshBilanz: function() {
            return this.generateBilanz();
        },

        /**
         * Erstellt eine Bilanz zum Jahresabschluss
         * @param {number} year - Jahr des Abschlusses
         * @returns {boolean} true bei Erfolg, false bei Fehler
         */
        createYearEndBalance: function(year = null) {
            // Wenn kein Jahr angegeben, das aktuelle Jahr verwenden
            const targetYear = year || CONFIG.BENUTZER.AKTUELLES_JAHR;

            try {
                // Bestätigung vom Benutzer einholen
                const ui = SpreadsheetApp.getUi();
                const response = ui.alert(
                    `Jahresabschluss ${targetYear}`,
                    `Möchten Sie die Bilanz zum Jahresabschluss ${targetYear} erstellen? ` +
                    "Dies ist eine wichtige Funktion, die normalerweise zum Ende des Geschäftsjahres " +
                    "durchgeführt wird.",
                    ui.ButtonSet.YES_NO
                );

                if (response !== ui.Button.YES) {
                    return false;
                }

                // Normale Bilanz erstellen
                const result = this.generateBilanz();

                if (!result) {
                    throw new Error("Fehler bei der Erstellung der normalen Bilanz");
                }

                // Jahresabschluss-spezifische Informationen hier hinzufügen
                // ...

                // Hinweis anzeigen
                ui.alert(
                    `Jahresabschluss ${targetYear}`,
                    "Die Bilanz zum Jahresabschluss wurde erstellt. " +
                    "In einer vollständigen Version würden hier weitere Jahresabschluss-Funktionen ausgeführt werden, " +
                    "wie z.B. die Erstellung der GuV, Eröffnungsbilanz, etc.",
                    ui.ButtonSet.OK
                );

                return true;
            } catch (e) {
                Logger.log(`Fehler beim Jahresabschluss: ${e.message}`);
                SpreadsheetApp.getUi().alert("Fehler beim Jahresabschluss: " + e.message);
                return false;
            }
        },

        /**
         * Exportiert die Bilanz als CSV
         * @returns {Object} Ergebnis des Exports
         */
        exportCSV: function() {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const bilanzSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.BILANZ);

                if (!bilanzSheet) {
                    throw new Error(`Sheet "${CONFIG.SYSTEM.SHEET_NAMES.BILANZ}" nicht gefunden`);
                }

                // Bilanzdaten laden
                const data = bilanzSheet.getDataRange().getValues();
                if (data.length <= 1) {
                    throw new Error("Keine Bilanzdaten vorhanden");
                }

                // CSV erstellen
                let csv = "";
                data.forEach(row => {
                    // Zeile formatieren und mit Semicolon trennen
                    const formattedRow = row.map(cell => {
                        if (typeof cell === 'number') {
                            return cell.toFixed(2);
                        } else if (typeof cell === 'string' && (cell.includes(',') || cell.includes('\n') || cell.includes(';'))) {
                            return `"${cell.replace(/"/g, '""')}"`;
                        } else {
                            return cell || "";
                        }
                    }).join(';');

                    csv += formattedRow + '\n';
                });

                // CSV anzeigen oder speichern
                const ui = SpreadsheetApp.getUi();
                ui.alert(
                    'Bilanz CSV-Export',
                    'Die Bilanzdaten wurden erfolgreich als CSV exportiert.',
                    ui.ButtonSet.OK
                );

                return {
                    success: true,
                    csvData: csv
                };
            } catch (e) {
                Logger.log(`Fehler beim CSV-Export: ${e.message}`);
                SpreadsheetApp.getUi().alert("Fehler beim CSV-Export: " + e.message);
                return {
                    success: false,
                    error: e.message
                };
            }
        }
    };
})();
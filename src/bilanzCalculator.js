// file: src/bilanzCalculator.js
import Helpers from "./helpers.js";
import config from "./config.js";

/**
 * Modul zur Erstellung einer Bilanz nach SKR04
 * Erstellt eine standardkonforme Bilanz basierend auf den Daten aus anderen Sheets
 */
const BilanzCalculator = (() => {
    /**
     * Erstellt eine leere Bilanz-Datenstruktur
     * @returns {Object} Leere Bilanz-Datenstruktur
     */
    const createEmptyBilanz = () => ({
        // Aktiva (Vermögenswerte)
        aktiva: {
            // Anlagevermögen
            sachanlagen: 0,                 // SKR04: 0400-0699, Sachanlagen
            immaterielleVermoegen: 0,       // SKR04: 0100-0199, Immaterielle Vermögensgegenstände
            finanzanlagen: 0,               // SKR04: 0700-0899, Finanzanlagen
            summeAnlagevermoegen: 0,        // Summe Anlagevermögen

            // Umlaufvermögen
            bankguthaben: 0,                // SKR04: 1200, Bank
            kasse: 0,                       // SKR04: 1210, Kasse
            forderungenLuL: 0,              // SKR04: 1300-1370, Forderungen aus Lieferungen und Leistungen
            vorraete: 0,                    // SKR04: 1400-1590, Vorräte
            summeUmlaufvermoegen: 0,        // Summe Umlaufvermögen

            // Rechnungsabgrenzung
            rechnungsabgrenzung: 0,         // SKR04: 1900-1990, Aktiver Rechnungsabgrenzungsposten

            // Gesamtsumme
            summeAktiva: 0                  // Summe aller Aktiva
        },

        // Passiva (Kapital und Schulden)
        passiva: {
            // Eigenkapital
            stammkapital: 0,                // SKR04: 2000, Gezeichnetes Kapital
            kapitalruecklagen: 0,           // SKR04: 2100, Kapitalrücklage
            gewinnvortrag: 0,               // SKR04: 2970, Gewinnvortrag
            verlustvortrag: 0,              // SKR04: 2978, Verlustvortrag
            jahresueberschuss: 0,           // Jahresüberschuss aus BWA
            summeEigenkapital: 0,           // Summe Eigenkapital

            // Verbindlichkeiten
            bankdarlehen: 0,                // SKR04: 3150, Verbindlichkeiten gegenüber Kreditinstituten
            gesellschafterdarlehen: 0,      // SKR04: 3300, Verbindlichkeiten gegenüber Gesellschaftern
            verbindlichkeitenLuL: 0,        // SKR04: 3200, Verbindlichkeiten aus Lieferungen und Leistungen
            steuerrueckstellungen: 0,       // SKR04: 3060, Steuerrückstellungen
            summeVerbindlichkeiten: 0,      // Summe Verbindlichkeiten

            // Rechnungsabgrenzung
            rechnungsabgrenzung: 0,         // SKR04: 3800-3990, Passiver Rechnungsabgrenzungsposten

            // Gesamtsumme
            summePassiva: 0                 // Summe aller Passiva
        }
    });

    /**
     * Sammelt Daten aus verschiedenen Sheets für die Bilanz
     *
     * @returns {Object} Bilanz-Datenstruktur mit befüllten Werten
     */
    const aggregateBilanzData = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const bilanzData = createEmptyBilanz();

            // Spalten-Konfigurationen für die verschiedenen Sheets
            const bankCols = config.sheets.bankbewegungen.columns;
            const ausgabenCols = config.sheets.ausgaben.columns;
            const gesellschafterCols = config.sheets.gesellschafterkonto.columns;

            // 1. Banksaldo aus "Bankbewegungen" (Endsaldo)
            const bankSheet = ss.getSheetByName("Bankbewegungen");
            if (bankSheet) {
                const lastRow = bankSheet.getLastRow();
                if (lastRow >= 1) {
                    const label = bankSheet.getRange(lastRow, bankCols.buchungstext).getValue().toString().toLowerCase();
                    if (label === "endsaldo") {
                        bilanzData.aktiva.bankguthaben = Helpers.parseCurrency(
                            bankSheet.getRange(lastRow, bankCols.saldo).getValue()
                        );
                    }
                }
            }

            // 2. Jahresüberschuss aus "BWA" (Letzte Zeile, sofern dort "Jahresüberschuss" vorkommt)
            const bwaSheet = ss.getSheetByName("BWA");
            if (bwaSheet) {
                const data = bwaSheet.getDataRange().getValues();
                for (let i = data.length - 1; i >= 0; i--) {
                    const row = data[i];
                    if (row[0].toString().toLowerCase().includes("jahresüberschuss")) {
                        // Letzte Spalte enthält den Jahreswert
                        bilanzData.passiva.jahresueberschuss = Helpers.parseCurrency(row[row.length - 1]);
                        break;
                    }
                }
            }

            // 3. Stammkapital aus Konfiguration
            bilanzData.passiva.stammkapital = config.tax.stammkapital || 25000;

            // 4. Suche nach Gesellschafterdarlehen im Gesellschafterkonto-Sheet
            const gesellschafterSheet = ss.getSheetByName("Gesellschafterkonto");
            if (gesellschafterSheet) {
                let darlehenSumme = 0;
                const data = gesellschafterSheet.getDataRange().getValues();

                // Überschrift überspringen
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    // Prüfen, ob es sich um ein Gesellschafterdarlehen handelt
                    if (row[gesellschafterCols.kategorie - 1] &&
                        row[gesellschafterCols.kategorie - 1].toString().toLowerCase() === "gesellschafterdarlehen") {
                        darlehenSumme += Helpers.parseCurrency(row[gesellschafterCols.betrag - 1] || 0);
                    }
                }

                bilanzData.passiva.gesellschafterdarlehen = darlehenSumme;
            }

            // 5. Steuerrückstellungen aus BWA oder Ausgaben-Sheet
            const ausSheet = ss.getSheetByName("Ausgaben");
            if (ausSheet) {
                let steuerRueckstellungen = 0;
                const data = ausSheet.getDataRange().getValues();

                // Überschrift überspringen
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const category = row[ausgabenCols.kategorie - 1]?.toString().trim() || "";

                    if (["Gewerbesteuerrückstellungen", "Körperschaftsteuer", "Solidaritätszuschlag", "Sonstige Steuerrückstellungen"].includes(category)) {
                        steuerRueckstellungen += Helpers.parseCurrency(row[ausgabenCols.nettobetrag - 1] || 0);
                    }
                }

                bilanzData.passiva.steuerrueckstellungen = steuerRueckstellungen;
            }

            // 6. Berechnung der Summen
            bilanzData.aktiva.summeAnlagevermoegen =
                bilanzData.aktiva.sachanlagen +
                bilanzData.aktiva.immaterielleVermoegen +
                bilanzData.aktiva.finanzanlagen;

            bilanzData.aktiva.summeUmlaufvermoegen =
                bilanzData.aktiva.bankguthaben +
                bilanzData.aktiva.kasse +
                bilanzData.aktiva.forderungenLuL +
                bilanzData.aktiva.vorraete;

            bilanzData.aktiva.summeAktiva =
                bilanzData.aktiva.summeAnlagevermoegen +
                bilanzData.aktiva.summeUmlaufvermoegen +
                bilanzData.aktiva.rechnungsabgrenzung;

            bilanzData.passiva.summeEigenkapital =
                bilanzData.passiva.stammkapital +
                bilanzData.passiva.kapitalruecklagen +
                bilanzData.passiva.gewinnvortrag -
                bilanzData.passiva.verlustvortrag +
                bilanzData.passiva.jahresueberschuss;

            bilanzData.passiva.summeVerbindlichkeiten =
                bilanzData.passiva.bankdarlehen +
                bilanzData.passiva.gesellschafterdarlehen +
                bilanzData.passiva.verbindlichkeitenLuL +
                bilanzData.passiva.steuerrueckstellungen;

            bilanzData.passiva.summePassiva =
                bilanzData.passiva.summeEigenkapital +
                bilanzData.passiva.summeVerbindlichkeiten +
                bilanzData.passiva.rechnungsabgrenzung;

            return bilanzData;
        } catch (e) {
            console.error("Fehler bei der Sammlung der Bilanzdaten:", e);
            SpreadsheetApp.getUi().alert("Fehler bei der Bilanzerstellung: " + e.toString());
            return null;
        }
    };

    /**
     * Konvertiert einen Zahlenwert in eine Zellenformel oder direkte Zahl
     * @param {number} value - Der Wert
     * @param {boolean} useFormula - Ob eine Formel verwendet werden soll
     * @returns {string|number} - Formel als String oder direkter Wert
     */
    const valueOrFormula = (value, useFormula = false) => {
        if (value === 0 && useFormula) {
            return "";  // Leere Zelle für 0-Werte bei Formeln
        }
        return value;
    };

    /**
     * Erstellt ein Array für die Aktiva-Seite der Bilanz
     * @param {Object} bilanzData - Die Bilanzdaten
     * @returns {Array} Array mit Zeilen für die Aktiva-Seite
     */
    const createAktivaArray = (bilanzData) => {
        const { aktiva } = bilanzData;
        const year = config.tax.year || new Date().getFullYear();

        return [
            [`Bilanz ${year} - Aktiva (Vermögenswerte)`, ""],
            ["", ""],
            ["1. Anlagevermögen", ""],
            ["1.1 Sachanlagen", valueOrFormula(aktiva.sachanlagen)],
            ["1.2 Immaterielle Vermögensgegenstände", valueOrFormula(aktiva.immaterielleVermoegen)],
            ["1.3 Finanzanlagen", valueOrFormula(aktiva.finanzanlagen)],
            ["Summe Anlagevermögen", "=SUM(B4:B6)"],
            ["", ""],
            ["2. Umlaufvermögen", ""],
            ["2.1 Bankguthaben", valueOrFormula(aktiva.bankguthaben)],
            ["2.2 Kasse", valueOrFormula(aktiva.kasse)],
            ["2.3 Forderungen aus Lieferungen und Leistungen", valueOrFormula(aktiva.forderungenLuL)],
            ["2.4 Vorräte", valueOrFormula(aktiva.vorraete)],
            ["Summe Umlaufvermögen", "=SUM(B10:B13)"],
            ["", ""],
            ["3. Rechnungsabgrenzungsposten", valueOrFormula(aktiva.rechnungsabgrenzung)],
            ["", ""],
            ["Summe Aktiva", "=B7+B14+B16"]
        ];
    };

    /**
     * Erstellt ein Array für die Passiva-Seite der Bilanz
     * @param {Object} bilanzData - Die Bilanzdaten
     * @returns {Array} Array mit Zeilen für die Passiva-Seite
     */
    const createPassivaArray = (bilanzData) => {
        const { passiva } = bilanzData;
        const year = config.tax.year || new Date().getFullYear();

        return [
            [`Bilanz ${year} - Passiva (Kapital und Schulden)`, ""],
            ["", ""],
            ["1. Eigenkapital", ""],
            ["1.1 Gezeichnetes Kapital (Stammkapital)", valueOrFormula(passiva.stammkapital)],
            ["1.2 Kapitalrücklage", valueOrFormula(passiva.kapitalruecklagen)],
            ["1.3 Gewinnvortrag", valueOrFormula(passiva.gewinnvortrag)],
            ["1.4 Verlustvortrag (negativ)", valueOrFormula(passiva.verlustvortrag)],
            ["1.5 Jahresüberschuss/Jahresfehlbetrag", valueOrFormula(passiva.jahresueberschuss)],
            ["Summe Eigenkapital", "=SUM(F4:F8)"],
            ["", ""],
            ["2. Verbindlichkeiten", ""],
            ["2.1 Verbindlichkeiten gegenüber Kreditinstituten", valueOrFormula(passiva.bankdarlehen)],
            ["2.2 Verbindlichkeiten gegenüber Gesellschaftern", valueOrFormula(passiva.gesellschafterdarlehen)],
            ["2.3 Verbindlichkeiten aus Lieferungen und Leistungen", valueOrFormula(passiva.verbindlichkeitenLuL)],
            ["2.4 Steuerrückstellungen", valueOrFormula(passiva.steuerrueckstellungen)],
            ["Summe Verbindlichkeiten", "=SUM(F12:F15)"],
            ["", ""],
            ["3. Rechnungsabgrenzungsposten", valueOrFormula(passiva.rechnungsabgrenzung)],
            ["", ""],
            ["Summe Passiva", "=F9+F16+F18"]
        ];
    };

    /**
     * Hauptfunktion zur Erstellung der Bilanz
     * Sammelt Daten und erstellt ein Bilanz-Sheet
     */
    const calculateBilanz = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();

            // Bilanzdaten aggregieren
            const bilanzData = aggregateBilanzData();
            if (!bilanzData) return;

            // Bilanz-Arrays erstellen
            const aktivaArray = createAktivaArray(bilanzData);
            const passivaArray = createPassivaArray(bilanzData);

            // Prüfen, ob Aktiva = Passiva
            const aktivaSumme = bilanzData.aktiva.summeAktiva;
            const passivaSumme = bilanzData.passiva.summePassiva;
            const differenz = Math.abs(aktivaSumme - passivaSumme);

            if (differenz > 0.01) {
                // Bei Differenz die Bilanz trotzdem erstellen, aber warnen
                ui.alert(
                    "Bilanz ist nicht ausgeglichen",
                    `Die Bilanzsummen von Aktiva (${aktivaSumme} €) und Passiva (${passivaSumme} €) ` +
                    `stimmen nicht überein. Differenz: ${differenz.toFixed(2)} €. ` +
                    `Bitte überprüfen Sie Ihre Buchhaltungsdaten.`,
                    ui.ButtonSet.OK
                );
            }

            // Erstelle oder leere das Blatt "Bilanz"
            let bilanzSheet = ss.getSheetByName("Bilanz");
            if (!bilanzSheet) {
                bilanzSheet = ss.insertSheet("Bilanz");
            } else {
                bilanzSheet.clearContents();
            }

            // Schreibe Aktiva ab Zelle A1 und Passiva ab Zelle E1
            bilanzSheet.getRange(1, 1, aktivaArray.length, 2).setValues(aktivaArray);
            bilanzSheet.getRange(1, 5, passivaArray.length, 2).setValues(passivaArray);

            // Formatierung anwenden
            // Überschriften formatieren
            bilanzSheet.getRange("A1").setFontWeight("bold").setFontSize(12);
            bilanzSheet.getRange("E1").setFontWeight("bold").setFontSize(12);

            // Zwischensummen und Gesamtsummen formatieren
            const summenZeilenAktiva = [7, 14, 18]; // Zeilen mit Summen in Aktiva
            const summenZeilenPassiva = [9, 16, 20]; // Zeilen mit Summen in Passiva

            summenZeilenAktiva.forEach(row => {
                bilanzSheet.getRange(row, 1, 1, 2).setFontWeight("bold");
                if (row === 18) { // Gesamtsumme Aktiva
                    bilanzSheet.getRange(row, 1, 1, 2).setBackground("#e6f2ff");
                } else {
                    bilanzSheet.getRange(row, 1, 1, 2).setBackground("#f0f0f0");
                }
            });

            summenZeilenPassiva.forEach(row => {
                bilanzSheet.getRange(row, 5, 1, 2).setFontWeight("bold");
                if (row === 20) { // Gesamtsumme Passiva
                    bilanzSheet.getRange(row, 5, 1, 2).setBackground("#e6f2ff");
                } else {
                    bilanzSheet.getRange(row, 5, 1, 2).setBackground("#f0f0f0");
                }
            });

            // Abschnittsüberschriften formatieren
            [3, 9, 11, 16].forEach(row => {
                bilanzSheet.getRange(row, 1).setFontWeight("bold");
            });

            [3, 11, 17].forEach(row => {
                bilanzSheet.getRange(row, 5).setFontWeight("bold");
            });

            // Währungsformat für Beträge anwenden
            bilanzSheet.getRange("B4:B18").setNumberFormat("#,##0.00 €");
            bilanzSheet.getRange("F4:F20").setNumberFormat("#,##0.00 €");

            // Spaltenbreiten anpassen
            bilanzSheet.autoResizeColumns(1, 6);

            // Erfolgsmeldung
            ui.alert("Die Bilanz wurde erfolgreich erstellt!");
        } catch (e) {
            console.error("Fehler bei der Bilanzerstellung:", e);
            SpreadsheetApp.getUi().alert("Fehler bei der Bilanzerstellung: " + e.toString());
        }
    };

    // Öffentliche API des Moduls
    return {
        calculateBilanz,
        // Für Testzwecke könnten hier weitere Funktionen exportiert werden
        _internal: {
            createEmptyBilanz,
            aggregateBilanzData
        }
    };
})();

export default BilanzCalculator;
// modules/bilanzModule/formatter.js
import numberUtils from '../../utils/numberUtils.js';
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Konvertiert einen Zahlenwert in eine Zellenformel oder direkte Zahl
 * @param {number} value - Der Wert
 * @param {boolean} useFormula - Ob eine Formel verwendet werden soll
 * @returns {string|number} - Formel als String oder direkter Wert
 */
function valueOrFormula(value, useFormula = false) {
    if (numberUtils.isApproximatelyEqual(value, 0) && useFormula) {
        return "";  // Leere Zelle für 0-Werte bei Formeln
    }
    return value;
}

/**
 * Erstellt ein Array für die Aktiva-Seite der Bilanz
 * @param {Object} bilanzData - Die Bilanzdaten
 * @param {Object} config - Die Konfiguration
 * @returns {Array} Array mit Zeilen für die Aktiva-Seite
 */
function createAktivaArray(bilanzData, config) {
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
}

/**
 * Erstellt ein Array für die Passiva-Seite der Bilanz
 * @param {Object} bilanzData - Die Bilanzdaten
 * @param {Object} config - Die Konfiguration
 * @returns {Array} Array mit Zeilen für die Passiva-Seite
 */
function createPassivaArray(bilanzData, config) {
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
}

/**
 * Erstellt das Bilanz-Sheet
 * @param {Object} bilanzData - Die Bilanzdaten
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} true bei Erfolg, false bei Fehler
 */
function generateBilanzSheet(bilanzData, ss, config) {
    try {
        // Bilanz-Arrays erstellen
        const aktivaArray = createAktivaArray(bilanzData, config);
        const passivaArray = createPassivaArray(bilanzData, config);

        // Erstelle oder leere das Blatt "Bilanz"
        let bilanzSheet = sheetUtils.getOrCreateSheet(ss, "Bilanz");
        bilanzSheet.clearContents();

        // Batch-Write statt einzelner Zellen-Updates
        // Schreibe Aktiva ab Zelle A1 und Passiva ab Zelle E1
        bilanzSheet.getRange(1, 1, aktivaArray.length, 2).setValues(aktivaArray);
        bilanzSheet.getRange(1, 5, passivaArray.length, 2).setValues(passivaArray);

        // Formatierung anwenden
        formatBilanzSheet(bilanzSheet, aktivaArray.length, passivaArray.length);

        // Bilanz-Sheet aktivieren
        ss.setActiveSheet(bilanzSheet);

        return true;
    } catch (e) {
        console.error("Fehler bei der Generierung des Bilanz-Sheets:", e);
        return false;
    }
}

/**
 * Formatiert das Bilanz-Sheet
 * @param {Sheet} bilanzSheet - Das zu formatierende Sheet
 * @param {number} aktivaLength - Anzahl der Aktiva-Zeilen
 * @param {number} passivaLength - Anzahl der Passiva-Zeilen
 */
function formatBilanzSheet(bilanzSheet, aktivaLength, passivaLength) {
    // Überschriften formatieren
    bilanzSheet.getRange("A1").setFontWeight("bold").setFontSize(12);
    bilanzSheet.getRange("E1").setFontWeight("bold").setFontSize(12);

    // Zwischensummen und Gesamtsummen formatieren
    const summenZeilenAktiva = [7, 14, 18]; // Zeilen mit Summen in Aktiva
    const summenZeilenPassiva = [9, 16, 20]; // Zeilen mit Summen in Passiva

    summenZeilenAktiva.forEach(row => {
        if (row <= aktivaLength) {
            bilanzSheet.getRange(row, 1, 1, 2).setFontWeight("bold");
            if (row === 18) { // Gesamtsumme Aktiva
                bilanzSheet.getRange(row, 1, 1, 2).setBackground("#e6f2ff");
            } else {
                bilanzSheet.getRange(row, 1, 1, 2).setBackground("#f0f0f0");
            }
        }
    });

    summenZeilenPassiva.forEach(row => {
        if (row <= passivaLength) {
            bilanzSheet.getRange(row, 5, 1, 2).setFontWeight("bold");
            if (row === 20) { // Gesamtsumme Passiva
                bilanzSheet.getRange(row, 5, 1, 2).setBackground("#e6f2ff");
            } else {
                bilanzSheet.getRange(row, 5, 1, 2).setBackground("#f0f0f0");
            }
        }
    });

    // Abschnittsüberschriften formatieren
    [3, 9, 16].forEach(row => {
        if (row <= aktivaLength) {
            bilanzSheet.getRange(row, 1).setFontWeight("bold");
        }
    });

    [3, 11, 17].forEach(row => {
        if (row <= passivaLength) {
            bilanzSheet.getRange(row, 5).setFontWeight("bold");
        }
    });

    // Währungsformat für Beträge anwenden
    bilanzSheet.getRange("B4:B18").setNumberFormat("#,##0.00 €");
    bilanzSheet.getRange("F4:F20").setNumberFormat("#,##0.00 €");

    // Spaltenbreiten anpassen
    bilanzSheet.autoResizeColumns(1, 6);
}

export default {
    valueOrFormula,
    createAktivaArray,
    createPassivaArray,
    generateBilanzSheet,
    formatBilanzSheet
};
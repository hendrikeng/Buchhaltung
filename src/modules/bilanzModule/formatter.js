// modules/bilanzModule/formatter.js
import numberUtils from '../../utils/numberUtils.js';
import sheetUtils from '../../utils/sheetUtils.js';


/**
 * Konvertiert einen Zahlenwert in eine Zellenformel oder direkte Zahl
 * Optimiert für einheitliche Rückgabe
 * @param {number} value - Der Wert
 * @param {boolean} useFormula - Ob eine Formel verwendet werden soll
 * @returns {string|number} - Formel als String oder direkter Wert
 */
function valueOrFormula(value, useFormula = false) {
    // Wenn der Wert nahe 0 ist und Formeln verwendet werden sollen, leeren String zurückgeben
    if (numberUtils.isApproximatelyEqual(value, 0) && useFormula) {
        return '';  // Leere Zelle für 0-Werte bei Formeln
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

    // Optimierung: Alle Werte als optimierte Formeln vorbereiten
    return [
        [`Bilanz ${year} - Aktiva (Vermögenswerte)`, ''],
        ['', ''],
        ['1. Anlagevermögen', ''],
        ['1.1 Sachanlagen', valueOrFormula(aktiva.sachanlagen)],
        ['1.2 Immaterielle Vermögensgegenstände', valueOrFormula(aktiva.immaterielleVermoegen)],
        ['1.3 Finanzanlagen', valueOrFormula(aktiva.finanzanlagen)],
        ['Summe Anlagevermögen', '=SUM(B4:B6)'],
        ['', ''],
        ['2. Umlaufvermögen', ''],
        ['2.1 Bankguthaben', valueOrFormula(aktiva.bankguthaben)],
        ['2.2 Kasse', valueOrFormula(aktiva.kasse)],
        ['2.3 Forderungen aus Lieferungen und Leistungen', valueOrFormula(aktiva.forderungenLuL)],
        ['2.4 Vorräte', valueOrFormula(aktiva.vorraete)],
        ['Summe Umlaufvermögen', '=SUM(B10:B13)'],
        ['', ''],
        ['3. Rechnungsabgrenzungsposten', valueOrFormula(aktiva.rechnungsabgrenzung)],
        ['', ''],
        ['Summe Aktiva', '=B7+B14+B16'],
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

    // Optimierung: Alle Werte als optimierte Formeln vorbereiten
    return [
        [`Bilanz ${year} - Passiva (Kapital und Schulden)`, ''],
        ['', ''],
        ['1. Eigenkapital', ''],
        ['1.1 Gezeichnetes Kapital (Stammkapital)', valueOrFormula(passiva.stammkapital)],
        ['1.2 Kapitalrücklage', valueOrFormula(passiva.kapitalruecklagen)],
        ['1.3 Gewinnvortrag', valueOrFormula(passiva.gewinnvortrag)],
        ['1.4 Verlustvortrag (negativ)', valueOrFormula(passiva.verlustvortrag)],
        ['1.5 Jahresüberschuss/Jahresfehlbetrag', valueOrFormula(passiva.jahresueberschuss)],
        ['Summe Eigenkapital', '=SUM(F4:F8)'],
        ['', ''],
        ['2. Verbindlichkeiten', ''],
        ['2.1 Verbindlichkeiten gegenüber Kreditinstituten', valueOrFormula(passiva.bankdarlehen)],
        ['2.2 Verbindlichkeiten gegenüber Gesellschaftern', valueOrFormula(passiva.gesellschafterdarlehen)],
        ['2.3 Verbindlichkeiten aus Lieferungen und Leistungen', valueOrFormula(passiva.verbindlichkeitenLuL)],
        ['2.4 Steuerrückstellungen', valueOrFormula(passiva.steuerrueckstellungen)],
        ['Summe Verbindlichkeiten', '=SUM(F12:F15)'],
        ['', ''],
        ['3. Rechnungsabgrenzungsposten', valueOrFormula(passiva.rechnungsabgrenzung)],
        ['', ''],
        ['Summe Passiva', '=F9+F16+F18'],
    ];
}

/**
 * Erstellt das Bilanz-Sheet mit optimierter Batch-Verarbeitung
 * @param {Object} bilanzData - Die Bilanzdaten
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} true bei Erfolg, false bei Fehler
 */
function generateBilanzSheet(bilanzData, ss, config) {
    try {
        console.log('Generating bilanz sheet...');

        // Bilanz-Arrays erstellen
        const aktivaArray = createAktivaArray(bilanzData, config);
        const passivaArray = createPassivaArray(bilanzData, config);

        // Erstelle oder leere das Blatt "Bilanz"
        const bilanzSheet = sheetUtils.getOrCreateSheet(ss, 'Bilanz');
        bilanzSheet.clearContents();

        // Optimierung: Batch-Write statt einzelner Zellen-Updates
        bilanzSheet.getRange(1, 1, aktivaArray.length, 2).setValues(aktivaArray);
        bilanzSheet.getRange(1, 5, passivaArray.length, 2).setValues(passivaArray);

        // Formatierung in einem optimierten Batch anwenden
        formatBilanzSheet(bilanzSheet, aktivaArray.length, passivaArray.length);

        // Bilanz-Sheet aktivieren
        ss.setActiveSheet(bilanzSheet);
        console.log('Bilanz sheet generated successfully');

        return true;
    } catch (e) {
        console.error('Fehler bei der Generierung des Bilanz-Sheets:', e);
        return false;
    }
}

/**
 * Formatiert das Bilanz-Sheet mit optimierter Batch-Formatierung
 * @param {Sheet} bilanzSheet - Das zu formatierende Sheet
 * @param {number} aktivaLength - Anzahl der Aktiva-Zeilen
 * @param {number} passivaLength - Anzahl der Passiva-Zeilen
 */
function formatBilanzSheet(bilanzSheet, aktivaLength, passivaLength) {
    // Optimierung: Strukturierte Formatierungsgruppen für Batch-Verarbeitung
    const formatGroups = {
        // Überschriften
        headers: [
            { range: 'A1', type: 'header' },
            { range: 'E1', type: 'header' },
        ],

        // Summenzeilen Aktiva
        aktivaSums: [
            { row: 7, type: sumRowType(7 === aktivaLength) },
            { row: 14, type: sumRowType(14 === aktivaLength) },
            { row: 18, type: 'total' },
        ],

        // Summenzeilen Passiva
        passivaSums: [
            { row: 9, type: sumRowType(9 === passivaLength) },
            { row: 16, type: sumRowType(16 === passivaLength) },
            { row: 20, type: 'total' },
        ],

        // Abschnittsüberschriften Aktiva
        aktivaHeaders: [3, 9, 16],

        // Abschnittsüberschriften Passiva
        passivaHeaders: [3, 11, 17],
    };

    // Formatierungsbatches anwenden

    // 1. Überschriften formatieren
    formatGroups.headers.forEach(header => {
        bilanzSheet.getRange(header.range)
            .setFontWeight('bold')
            .setFontSize(12);
    });

    // 2. Summenzeilen Aktiva formatieren
    formatGroups.aktivaSums.forEach(sum => {
        if (sum.row <= aktivaLength) {
            const range = bilanzSheet.getRange(sum.row, 1, 1, 2);
            range.setFontWeight('bold');

            if (sum.type === 'total') {
                range.setBackground('#e6f2ff');
            } else if (sum.type === 'subtotal') {
                range.setBackground('#f0f0f0');
            }
        }
    });

    // 3. Summenzeilen Passiva formatieren
    formatGroups.passivaSums.forEach(sum => {
        if (sum.row <= passivaLength) {
            const range = bilanzSheet.getRange(sum.row, 5, 1, 2);
            range.setFontWeight('bold');

            if (sum.type === 'total') {
                range.setBackground('#e6f2ff');
            } else if (sum.type === 'subtotal') {
                range.setBackground('#f0f0f0');
            }
        }
    });

    // 4. Abschnittsüberschriften formatieren
    formatGroups.aktivaHeaders.forEach(row => {
        if (row <= aktivaLength) {
            bilanzSheet.getRange(row, 1).setFontWeight('bold');
        }
    });

    formatGroups.passivaHeaders.forEach(row => {
        if (row <= passivaLength) {
            bilanzSheet.getRange(row, 5).setFontWeight('bold');
        }
    });

    // 5. Währungsformat für Beträge anwenden
    bilanzSheet.getRange('B4:B18').setNumberFormat('#,##0.00 €');
    bilanzSheet.getRange('F4:F20').setNumberFormat('#,##0.00 €');

    // 6. Spaltenbreiten anpassen
    bilanzSheet.autoResizeColumns(1, 6);
}

/**
 * Hilfsfunktion zur Bestimmung des Summenzeilen-Typs
 * @param {boolean} isLastRow - Ist diese Zeile die letzte Zeile
 * @returns {string} - Zeilentyp
 */
function sumRowType(isLastRow) {
    return isLastRow ? 'total' : 'subtotal';
}

export default {
    valueOrFormula,
    createAktivaArray,
    createPassivaArray,
    generateBilanzSheet,
    formatBilanzSheet,
};
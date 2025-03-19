// modules/ustvaModule/formatter.js
import numberUtils from '../../utils/numberUtils.js';
import calculator from './calculator.js';
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Formatiert eine UStVA-Datenzeile für die Ausgabe
 *
 * @param {string} label - Bezeichnung der Zeile (z.B. Monat oder Quartal)
 * @param {Object} d - UStVA-Datenobjekt für den Zeitraum
 * @returns {Array} Formatierte Zeile für die Ausgabe
 */
function formatUStVARow(label, d) {
    // Optimierung: Werte einmal berechnen für mehrfache Verwendung
    const ust = d.ust_7 + d.ust_19;
    const vst = d.vst_7 + d.vst_19;
    const vstNetto = vst - d.nicht_abzugsfaehige_vst;

    // Berechnung der USt-Zahlung: USt minus VSt (abzüglich nicht abzugsfähiger VSt)
    const ustZahlung = numberUtils.round(ust - vstNetto, 2);

    // Optimierung: Einnahmen und Ausgaben separat berechnen
    const einnahmen = d.steuerpflichtige_einnahmen + d.steuerfreie_inland_einnahmen +
        d.steuerfreie_ausland_einnahmen;

    const ausgaben = d.steuerpflichtige_ausgaben + d.steuerfreie_inland_ausgaben +
        d.steuerfreie_ausland_ausgaben + d.eigenbelege_steuerpflichtig +
        d.eigenbelege_steuerfrei;

    // Berechnung des Gesamtergebnisses: Einnahmen minus Ausgaben (ohne Steueranteil)
    const ergebnis = numberUtils.round(einnahmen - ausgaben, 2);

    // Formatierte Zeile zurückgeben
    return [
        label,
        d.steuerpflichtige_einnahmen,
        d.steuerfreie_inland_einnahmen,
        d.steuerfreie_ausland_einnahmen,
        d.steuerpflichtige_ausgaben,
        d.steuerfreie_inland_ausgaben,
        d.steuerfreie_ausland_ausgaben,
        d.eigenbelege_steuerpflichtig,
        d.eigenbelege_steuerfrei,
        d.nicht_abzugsfaehige_vst,
        d.ust_7,
        d.ust_19,
        d.vst_7,
        d.vst_19,
        ustZahlung,
        ergebnis,
    ];
}

/**
 * Erstellt das UStVA-Sheet basierend auf den UStVA-Daten
 * Optimierte Version mit Batch-Operationen
 *
 * @param {Object} ustvaData - UStVA-Daten nach Monaten
 * @param {Spreadsheet} ss - Das aktive Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} true bei Erfolg, false bei Fehler
 */
function generateUStVASheet(ustvaData, ss, config) {
    try {
        console.log('Generating UStVA sheet...');

        // Ausgabe-Header erstellen
        const outputRows = [
            [
                'Zeitraum',
                'Steuerpflichtige Einnahmen',
                'Steuerfreie Inland-Einnahmen',
                'Steuerfreie Ausland-Einnahmen',
                'Steuerpflichtige Ausgaben',
                'Steuerfreie Inland-Ausgaben',
                'Steuerfreie Ausland-Ausgaben',
                'Eigenbelege steuerpflichtig',
                'Eigenbelege steuerfrei',
                'Nicht abzugsfähige VSt (Bewirtung)',
                'USt 7%',
                'USt 19%',
                'VSt 7%',
                'VSt 19%',
                'USt-Zahlung',
                'Ergebnis',
            ],
        ];

        // Optimierung: Prepare quarterly data in advance to avoid multiple recalculations
        const quarters = {};
        for (let q = 1; q <= 4; q++) {
            const start = (q - 1) * 3 + 1;
            const end = q * 3;
            quarters[q] = calculator.aggregateUStVA(ustvaData, start, end);
        }

        // Ganzjahresergebnis vorab berechnen
        const yearlyData = calculator.aggregateUStVA(ustvaData, 1, 12);

        // Monatliche Daten ausgeben und berechnete Quartale verwenden
        config.common.months.forEach((name, i) => {
            const month = i + 1;
            outputRows.push(formatUStVARow(name, ustvaData[month]));

            // Nach jedem Quartal eine Zusammenfassung einfügen
            if (month % 3 === 0) {
                const quartalsNummer = month / 3;
                outputRows.push(formatUStVARow(
                    `Quartal ${quartalsNummer}`,
                    quarters[quartalsNummer],
                ));
            }
        });

        // Jahresergebnis anfügen aus vorberechneten Daten
        outputRows.push(formatUStVARow('Gesamtjahr', yearlyData));

        // UStVA-Sheet erstellen oder aktualisieren
        const ustvaSheet = sheetUtils.getOrCreateSheet(ss, 'UStVA');
        ustvaSheet.clearContents();

        // Optimierung: Effizienteres Batch-Update
        const dataRange = ustvaSheet.getRange(1, 1, outputRows.length, outputRows[0].length);
        dataRange.setValues(outputRows);

        // Optimierte Batch-Formatierung
        applyUStVASheetFormatting(ustvaSheet, outputRows);

        // UStVA-Sheet aktivieren
        ss.setActiveSheet(ustvaSheet);
        console.log('UStVA sheet generated successfully');

        return true;
    } catch (e) {
        console.error('Fehler bei der Generierung des UStVA-Sheets:', e);
        return false;
    }
}

/**
 * Wendet Formatierungen auf das UStVA-Sheet an
 * @param {Sheet} sheet - Das UStVA-Sheet
 * @param {Array} outputRows - Die Ausgabedaten
 */
function applyUStVASheetFormatting(sheet, outputRows) {
    try {
        // Header formatieren in einem API-Call
        sheet.getRange(1, 1, 1, outputRows[0].length)
            .setFontWeight('bold')
            .setBackground('#f3f3f3');

        // Quartalszeilen formatieren in einem API-Call
        for (let i = 0; i < 4; i++) {
            const row = 3 * (i + 1) + 1 + i; // Position der Quartalszeile
            sheet.getRange(row, 1, 1, outputRows[0].length)
                .setBackground('#e6f2ff');
        }

        // Jahreszeile formatieren in einem API-Call
        sheet.getRange(outputRows.length, 1, 1, outputRows[0].length)
            .setBackground('#d9e6f2')
            .setFontWeight('bold');

        // Währungsformat für Beträge (Spalten 2-16) in einem API-Call
        sheet.getRange(2, 2, outputRows.length - 1, 15)
            .setNumberFormat('#,##0.00 €');

        // Spaltenbreiten optimal anpassen
        sheet.autoResizeColumns(1, outputRows[0].length);
    } catch (e) {
        console.error('Fehler bei der Formatierung des UStVA-Sheets:', e);
    }
}

export default {
    formatUStVARow,
    generateUStVASheet,
};
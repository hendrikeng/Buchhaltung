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
    // Berechnung der USt-Zahlung: USt minus VSt (abzüglich nicht abzugsfähiger VSt)
    const ustZahlung = numberUtils.round(
        (d.ust_7 + d.ust_19) - ((d.vst_7 + d.vst_19) - d.nicht_abzugsfaehige_vst),
        2
    );

    // Berechnung des Gesamtergebnisses: Einnahmen minus Ausgaben (ohne Steueranteil)
    const ergebnis = numberUtils.round(
        (d.steuerpflichtige_einnahmen + d.steuerfreie_inland_einnahmen + d.steuerfreie_ausland_einnahmen) -
        (d.steuerpflichtige_ausgaben + d.steuerfreie_inland_ausgaben + d.steuerfreie_ausland_ausgaben +
            d.eigenbelege_steuerpflichtig + d.eigenbelege_steuerfrei),
        2
    );

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
        ergebnis
    ];
}

/**
 * Erstellt das UStVA-Sheet basierend auf den UStVA-Daten
 *
 * @param {Object} ustvaData - UStVA-Daten nach Monaten
 * @param {Spreadsheet} ss - Das aktive Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} true bei Erfolg, false bei Fehler
 */
function generateUStVASheet(ustvaData, ss, config) {
    try {
        // Ausgabe-Header erstellen
        const outputRows = [
            [
                "Zeitraum",
                "Steuerpflichtige Einnahmen",
                "Steuerfreie Inland-Einnahmen",
                "Steuerfreie Ausland-Einnahmen",
                "Steuerpflichtige Ausgaben",
                "Steuerfreie Inland-Ausgaben",
                "Steuerfreie Ausland-Ausgaben",
                "Eigenbelege steuerpflichtig",
                "Eigenbelege steuerfrei",
                "Nicht abzugsfähige VSt (Bewirtung)",
                "USt 7%",
                "USt 19%",
                "VSt 7%",
                "VSt 19%",
                "USt-Zahlung",
                "Ergebnis"
            ]
        ];

        // Monatliche Daten ausgeben
        config.common.months.forEach((name, i) => {
            const month = i + 1;
            outputRows.push(formatUStVARow(name, ustvaData[month]));

            // Nach jedem Quartal eine Zusammenfassung einfügen
            if (month % 3 === 0) {
                const quartalsNummer = month / 3;
                const quartalsStart = month - 2;
                outputRows.push(formatUStVARow(
                    `Quartal ${quartalsNummer}`,
                    calculator.aggregateUStVA(ustvaData, quartalsStart, month)
                ));
            }
        });

        // Jahresergebnis anfügen
        outputRows.push(formatUStVARow("Gesamtjahr", calculator.aggregateUStVA(ustvaData, 1, 12)));

        // UStVA-Sheet erstellen oder aktualisieren
        let ustvaSheet = sheetUtils.getOrCreateSheet(ss, "UStVA");
        ustvaSheet.clearContents();

        // Daten in das Sheet schreiben mit optimiertem Batch-Update
        const dataRange = ustvaSheet.getRange(1, 1, outputRows.length, outputRows[0].length);
        dataRange.setValues(outputRows);

        // Header formatieren
        ustvaSheet.getRange(1, 1, 1, outputRows[0].length).setFontWeight("bold");

        // Quartalszellen formatieren
        for (let i = 0; i < 4; i++) {
            const row = 3 * (i + 1) + 1 + i; // Position der Quartalszeile
            ustvaSheet.getRange(row, 1, 1, outputRows[0].length).setBackground("#e6f2ff");
        }

        // Jahreszeile formatieren
        ustvaSheet.getRange(outputRows.length, 1, 1, outputRows[0].length)
            .setBackground("#d9e6f2")
            .setFontWeight("bold");

        // Währungsformat für Beträge (Spalten 2-16)
        ustvaSheet.getRange(2, 2, outputRows.length - 1, 15).setNumberFormat("#,##0.00 €");

        // Spaltenbreiten anpassen
        ustvaSheet.autoResizeColumns(1, outputRows[0].length);

        // UStVA-Sheet aktivieren
        ss.setActiveSheet(ustvaSheet);

        return true;
    } catch (e) {
        console.error("Fehler bei der Generierung des UStVA-Sheets:", e);
        return false;
    }
}

export default {
    formatUStVARow,
    generateUStVASheet
};
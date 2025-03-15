// modules/bwaModule/formatter.js
import numberUtils from '../../utils/numberUtils.js';
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Erstellt den Header für die BWA mit Monats- und Quartalsspalten
 * @param {Object} config - Die Konfiguration
 * @returns {Array} Header-Zeile
 */
function buildHeaderRow(config) {
    const headers = ["Kategorie"];
    for (let q = 0; q < 4; q++) {
        for (let m = q * 3; m < q * 3 + 3; m++) {
            headers.push(`${config.common.months[m]} (€)`);
        }
        headers.push(`Q${q + 1} (€)`);
    }
    headers.push("Jahr (€)");
    return headers;
}

/**
 * Erstellt eine Ausgabezeile für eine Position
 * @param {Object} pos - Position mit Label und Wert-Funktion
 * @param {Object} bwaData - BWA-Daten
 * @returns {Array} Formatierte Zeile
 */
function buildOutputRow(pos, bwaData) {
    const monthly = [];
    let yearly = 0;

    // Monatswerte berechnen
    for (let m = 1; m <= 12; m++) {
        const val = pos.get(bwaData[m]) || 0;
        monthly.push(val);
        yearly += val;
    }

    // Quartalswerte berechnen
    const quarters = [0, 0, 0, 0];
    for (let i = 0; i < 12; i++) {
        quarters[Math.floor(i / 3)] += monthly[i];
    }

    // Runde alle Werte für bessere Darstellung
    const roundedMonthly = monthly.map(val => numberUtils.round(val, 2));
    const roundedQuarters = quarters.map(val => numberUtils.round(val, 2));
    const roundedYearly = numberUtils.round(yearly, 2);

    // Zeile zusammenstellen
    return [pos.label,
        ...roundedMonthly.slice(0, 3), roundedQuarters[0],
        ...roundedMonthly.slice(3, 6), roundedQuarters[1],
        ...roundedMonthly.slice(6, 9), roundedQuarters[2],
        ...roundedMonthly.slice(9, 12), roundedQuarters[3],
        roundedYearly];
}

/**
 * Erstellt das BWA-Sheet basierend auf den BWA-Daten
 * @param {Object} bwaData - BWA-Daten nach Monaten
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} true bei Erfolg, false bei Fehler
 */
function generateBWASheet(bwaData, ss, config) {
    try {
        // Positionen für die BWA definieren
        const positions = [
            {label: "Erlöse aus Lieferungen und Leistungen", get: d => d.umsatzerloese || 0},
            {label: "Provisionserlöse", get: d => d.provisionserloese || 0},
            {label: "Steuerfreie Inland-Einnahmen", get: d => d.steuerfreieInlandEinnahmen || 0},
            {label: "Steuerfreie Ausland-Einnahmen", get: d => d.steuerfreieAuslandEinnahmen || 0},
            {label: "Sonstige betriebliche Erträge", get: d => d.sonstigeErtraege || 0},
            {label: "Erträge aus Vermietung/Verpachtung", get: d => d.vermietung || 0},
            {label: "Erträge aus Zuschüssen", get: d => d.zuschuesse || 0},
            {label: "Erträge aus Währungsgewinnen", get: d => d.waehrungsgewinne || 0},
            {label: "Erträge aus Anlagenabgängen", get: d => d.anlagenabgaenge || 0},
            {label: "Betriebserlöse", get: d => d.gesamtErloese || 0},
            {label: "Wareneinsatz", get: d => d.wareneinsatz || 0},
            {label: "Bezogene Leistungen", get: d => d.fremdleistungen || 0},
            {label: "Roh-, Hilfs- & Betriebsstoffe", get: d => d.rohHilfsBetriebsstoffe || 0},
            {label: "Gesamtkosten Material & Fremdleistungen", get: d => d.gesamtWareneinsatz || 0},
            {label: "Bruttolöhne & Gehälter", get: d => d.bruttoLoehne || 0},
            {label: "Soziale Abgaben & Arbeitgeberanteile", get: d => d.sozialeAbgaben || 0},
            {label: "Sonstige Personalkosten", get: d => d.sonstigePersonalkosten || 0},
            {label: "Werbung & Marketing", get: d => d.werbungMarketing || 0},
            {label: "Reisekosten", get: d => d.reisekosten || 0},
            {label: "Versicherungen", get: d => d.versicherungen || 0},
            {label: "Telefon & Internet", get: d => d.telefonInternet || 0},
            {label: "Bürokosten", get: d => d.buerokosten || 0},
            {label: "Fortbildungskosten", get: d => d.fortbildungskosten || 0},
            {label: "Kfz-Kosten", get: d => d.kfzKosten || 0},
            {label: "Sonstige betriebliche Aufwendungen", get: d => d.sonstigeAufwendungen || 0},
            {label: "Abschreibungen Maschinen", get: d => d.abschreibungenMaschinen || 0},
            {label: "Abschreibungen Büroausstattung", get: d => d.abschreibungenBueromaterial || 0},
            {label: "Abschreibungen immaterielle Wirtschaftsgüter", get: d => d.abschreibungenImmateriell || 0},
            {label: "Zinsen auf Bankdarlehen", get: d => d.zinsenBank || 0},
            {label: "Zinsen auf Gesellschafterdarlehen", get: d => d.zinsenGesellschafter || 0},
            {label: "Leasingkosten", get: d => d.leasingkosten || 0},
            {label: "Gesamt Abschreibungen & Zinsen", get: d => d.gesamtAbschreibungenZinsen || 0},
            {label: "Eigenkapitalveränderungen", get: d => d.eigenkapitalveraenderungen || 0},
            {label: "Gesellschafterdarlehen", get: d => d.gesellschafterdarlehen || 0},
            {label: "Ausschüttungen an Gesellschafter", get: d => d.ausschuettungen || 0},
            {label: "Steuerrückstellungen", get: d => d.steuerrueckstellungen || 0},
            {label: "Rückstellungen sonstige", get: d => d.rueckstellungenSonstige || 0},
            {label: "Betriebsergebnis vor Steuern (EBIT)", get: d => d.ebit || 0},
            {label: "Umsatzsteuer (abzuführen)", get: d => d.umsatzsteuer || 0},
            {label: "Vorsteuer", get: d => d.vorsteuer || 0},
            {label: "Nicht abzugsfähige VSt (Bewirtung)", get: d => d.nichtAbzugsfaehigeVSt || 0},
            {label: "Körperschaftsteuer", get: d => d.koerperschaftsteuer || 0},
            {label: "Solidaritätszuschlag", get: d => d.solidaritaetszuschlag || 0},
            {label: "Gewerbesteuer", get: d => d.gewerbesteuer || 0},
            {label: "Gesamtsteueraufwand", get: d => d.steuerlast || 0},
            {label: "Jahresüberschuss/-fehlbetrag", get: d => d.gewinnNachSteuern || 0}
        ];

        // Header-Zeile erstellen
        const headerRow = buildHeaderRow(config);
        const outputRows = [headerRow];

        // Gruppenhierarchie für BWA
        const bwaGruppen = [
            {titel: "Betriebserlöse (Einnahmen)", count: 10},
            {titel: "Materialaufwand & Wareneinsatz", count: 4},
            {titel: "Betriebsausgaben (Sachkosten)", count: 11},
            {titel: "Abschreibungen & Zinsen", count: 7},
            {titel: "Besondere Posten", count: 3},
            {titel: "Rückstellungen", count: 2},
            {titel: "Betriebsergebnis vor Steuern (EBIT)", count: 1},
            {titel: "Steuern & Vorsteuer", count: 7},
            {titel: "Jahresüberschuss/-fehlbetrag", count: 1}
        ];

        // Ausgabe mit Gruppenhierarchie erstellen
        let posIndex = 0;
        for (let gruppenIndex = 0; gruppenIndex < bwaGruppen.length; gruppenIndex++) {
            const gruppe = bwaGruppen[gruppenIndex];

            // Gruppenüberschrift
            outputRows.push([
                `${gruppenIndex + 1}. ${gruppe.titel}`,
                ...Array(headerRow.length - 1).fill("")
            ]);

            // Gruppenpositionen
            for (let i = 0; i < gruppe.count; i++) {
                outputRows.push(buildOutputRow(positions[posIndex++], bwaData));
            }

            // Leerzeile nach jeder Gruppe außer der letzten
            if (gruppenIndex < bwaGruppen.length - 1) {
                outputRows.push(Array(headerRow.length).fill(""));
            }
        }

        // BWA-Sheet erstellen oder aktualisieren
        let bwaSheet = sheetUtils.getOrCreateSheet(ss, "BWA");
        bwaSheet.clearContents();

        // Daten in das Sheet schreiben (als Batch für Performance)
        const dataRange = bwaSheet.getRange(1, 1, outputRows.length, outputRows[0].length);
        dataRange.setValues(outputRows);

        // Formatierungen anwenden
        applyBwaFormatting(bwaSheet, headerRow.length, bwaGruppen, outputRows.length);

        // BWA-Sheet aktivieren
        ss.setActiveSheet(bwaSheet);

        return true;
    } catch (e) {
        console.error("Fehler bei der BWA-Erstellung:", e);
        return false;
    }
}

/**
 * Wendet Formatierungen auf das BWA-Sheet an
 * @param {Sheet} sheet - Das zu formatierende Sheet
 * @param {number} headerLength - Anzahl der Spalten
 * @param {Array} bwaGruppen - Gruppenhierarchie
 * @param {number} totalRows - Gesamtzahl der Zeilen
 */
function applyBwaFormatting(sheet, headerLength, bwaGruppen, totalRows) {
    // Header formatieren
    sheet.getRange(1, 1, 1, headerLength).setFontWeight("bold").setBackground("#f3f3f3");

    // Gruppenüberschriften formatieren
    let rowIndex = 2;
    for (const gruppe of bwaGruppen) {
        sheet.getRange(rowIndex, 1).setFontWeight("bold");
        rowIndex += gruppe.count + 1; // +1 für die Leerzeile
    }

    // Währungsformat für alle Zahlenwerte
    sheet.getRange(2, 2, totalRows - 1, headerLength - 1).setNumberFormat("#,##0.00 €");

    // Summen-Zeilen hervorheben
    const summenZeilen = [11, 15, 26, 33, 36, 38, 39, 46];
    summenZeilen.forEach(row => {
        if (row <= totalRows) {
            sheet.getRange(row, 1, 1, headerLength).setBackground("#e6f2ff");
        }
    });

    // EBIT und Jahresüberschuss hervorheben
    sheet.getRange(39, 1, 1, headerLength).setFontWeight("bold");
    sheet.getRange(46, 1, 1, headerLength).setFontWeight("bold");

    // Spaltenbreiten anpassen
    sheet.autoResizeColumns(1, headerLength);
}

export default {
    buildHeaderRow,
    buildOutputRow,
    generateBWASheet,
    applyBwaFormatting
};
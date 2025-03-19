// modules/bwaModule/formatter.js
import numberUtils from '../../utils/numberUtils.js';
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Erstellt den Header für die BWA mit Monats- und Quartalsspalten
 * @param {Object} config - Die Konfiguration
 * @returns {Array} Header-Zeile
 */
function buildHeaderRow(config) {
    const headers = ['Kategorie'];

    // Optimierung: Konstruiere Header mit einer For-Schleife statt impliziter Iteration
    for (let q = 0; q < 4; q++) {
        for (let m = q * 3; m < q * 3 + 3; m++) {
            headers.push(`${config.common.months[m]} (€)`);
        }
        headers.push(`Q${q + 1} (€)`);
    }
    headers.push('Jahr (€)');

    return headers;
}

/**
 * Erstellt eine Ausgabezeile für eine Position mit optimierter Berechnungslogik
 * @param {Object} pos - Position mit Label und Wert-Funktion
 * @param {Object} bwaData - BWA-Daten
 * @returns {Array} Formatierte Zeile
 */
function buildOutputRow(pos, bwaData) {
    // Optimierung: Prä-allokiere Arrays für Werte
    const monthly = new Array(12).fill(0);
    const quarters = new Array(4).fill(0);
    let yearly = 0;

    // Optimierung: Hole Monatswerte in einem Durchgang
    for (let m = 1; m <= 12; m++) {
        const val = pos.get(bwaData[m]) || 0;
        monthly[m-1] = val;
        yearly += val;

        // Quartale direkt mitzählen (optimiert)
        quarters[Math.floor((m-1) / 3)] += val;
    }

    // Runde alle Werte für bessere Darstellung
    const roundedMonthly = monthly.map(val => numberUtils.round(val, 2));
    const roundedQuarters = quarters.map(val => numberUtils.round(val, 2));
    const roundedYearly = numberUtils.round(yearly, 2);

    // Zeile zusammenstellen mit optimierter Struktur
    // [Label, Jan, Feb, Mar, Q1, Apr, May, Jun, Q2, ...]
    return [
        pos.label,
        ...roundedMonthly.slice(0, 3),
        roundedQuarters[0],
        ...roundedMonthly.slice(3, 6),
        roundedQuarters[1],
        ...roundedMonthly.slice(6, 9),
        roundedQuarters[2],
        ...roundedMonthly.slice(9, 12),
        roundedQuarters[3],
        roundedYearly,
    ];
}

/**
 * Erstellt das BWA-Sheet basierend auf den BWA-Daten mit optimierter Batch-Verarbeitung
 * @param {Object} bwaData - BWA-Daten nach Monaten
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} true bei Erfolg, false bei Fehler
 */
function generateBWASheet(bwaData, ss, config) {
    try {
        console.log('Generating BWA sheet...');

        // Positionen für die BWA definieren
        const positions = [
            {label: 'Erlöse aus Lieferungen und Leistungen', get: d => d.umsatzerloese || 0},
            {label: 'Provisionserlöse', get: d => d.provisionserloese || 0},
            {label: 'Steuerfreie Inland-Einnahmen', get: d => d.steuerfreieInlandEinnahmen || 0},
            {label: 'Steuerfreie Ausland-Einnahmen', get: d => d.steuerfreieAuslandEinnahmen || 0},
            {label: 'Sonstige betriebliche Erträge', get: d => d.sonstigeErtraege || 0},
            {label: 'Erträge aus Vermietung/Verpachtung', get: d => d.vermietung || 0},
            {label: 'Erträge aus Zuschüssen', get: d => d.zuschuesse || 0},
            {label: 'Erträge aus Währungsgewinnen', get: d => d.waehrungsgewinne || 0},
            {label: 'Erträge aus Anlagenabgängen', get: d => d.anlagenabgaenge || 0},
            {label: 'Betriebserlöse', get: d => d.gesamtErloese || 0},
            {label: 'Wareneinsatz', get: d => d.wareneinsatz || 0},
            {label: 'Bezogene Leistungen', get: d => d.fremdleistungen || 0},
            {label: 'Roh-, Hilfs- & Betriebsstoffe', get: d => d.rohHilfsBetriebsstoffe || 0},
            {label: 'Gesamtkosten Material & Fremdleistungen', get: d => d.gesamtWareneinsatz || 0},
            {label: 'Bruttolöhne & Gehälter', get: d => d.bruttoLoehne || 0},
            {label: 'Soziale Abgaben & Arbeitgeberanteile', get: d => d.sozialeAbgaben || 0},
            {label: 'Sonstige Personalkosten', get: d => d.sonstigePersonalkosten || 0},
            {label: 'Werbung & Marketing', get: d => d.werbungMarketing || 0},
            {label: 'Reisekosten', get: d => d.reisekosten || 0},
            {label: 'Versicherungen', get: d => d.versicherungen || 0},
            {label: 'Telefon & Internet', get: d => d.telefonInternet || 0},
            {label: 'Bürokosten', get: d => d.buerokosten || 0},
            {label: 'Fortbildungskosten', get: d => d.fortbildungskosten || 0},
            {label: 'Kfz-Kosten', get: d => d.kfzKosten || 0},
            {label: 'Sonstige betriebliche Aufwendungen', get: d => d.sonstigeAufwendungen || 0},
            {label: 'Abschreibungen Maschinen', get: d => d.abschreibungenMaschinen || 0},
            {label: 'Abschreibungen Büroausstattung', get: d => d.abschreibungenBueromaterial || 0},
            {label: 'Abschreibungen immaterielle Wirtschaftsgüter', get: d => d.abschreibungenImmateriell || 0},
            {label: 'Zinsen auf Bankdarlehen', get: d => d.zinsenBank || 0},
            {label: 'Zinsen auf Gesellschafterdarlehen', get: d => d.zinsenGesellschafter || 0},
            {label: 'Leasingkosten', get: d => d.leasingkosten || 0},
            {label: 'Gesamt Abschreibungen & Zinsen', get: d => d.gesamtAbschreibungenZinsen || 0},
            {label: 'Eigenkapitalveränderungen', get: d => d.eigenkapitalveraenderungen || 0},
            {label: 'Gesellschafterdarlehen', get: d => d.gesellschafterdarlehen || 0},
            {label: 'Ausschüttungen an Gesellschafter', get: d => d.ausschuettungen || 0},
            {label: 'Steuerrückstellungen', get: d => d.steuerrueckstellungen || 0},
            {label: 'Rückstellungen sonstige', get: d => d.rueckstellungenSonstige || 0},
            {label: 'Betriebsergebnis vor Steuern (EBIT)', get: d => d.ebit || 0},
            {label: 'Umsatzsteuer (abzuführen)', get: d => d.umsatzsteuer || 0},
            {label: 'Vorsteuer', get: d => d.vorsteuer || 0},
            {label: 'Nicht abzugsfähige VSt (Bewirtung)', get: d => d.nichtAbzugsfaehigeVSt || 0},
            {label: 'Körperschaftsteuer', get: d => d.koerperschaftsteuer || 0},
            {label: 'Solidaritätszuschlag', get: d => d.solidaritaetszuschlag || 0},
            {label: 'Gewerbesteuer', get: d => d.gewerbesteuer || 0},
            {label: 'Gesamtsteueraufwand', get: d => d.steuerlast || 0},
            {label: 'Jahresüberschuss/-fehlbetrag', get: d => d.gewinnNachSteuern || 0},
        ];

        // Header-Zeile und Gruppenhierarchie in einem Batch erstellen
        const headerRow = buildHeaderRow(config);

        // Optimierung: Gruppenhierarchie mit besserer Datenstruktur
        const bwaGruppen = [
            {titel: 'Betriebserlöse (Einnahmen)', count: 10},
            {titel: 'Materialaufwand & Wareneinsatz', count: 4},
            {titel: 'Betriebsausgaben (Sachkosten)', count: 11},
            {titel: 'Abschreibungen & Zinsen', count: 7},
            {titel: 'Besondere Posten', count: 3},
            {titel: 'Rückstellungen', count: 2},
            {titel: 'Betriebsergebnis vor Steuern (EBIT)', count: 1},
            {titel: 'Steuern & Vorsteuer', count: 7},
            {titel: 'Jahresüberschuss/-fehlbetrag', count: 1},
        ];

        // Optimierung: Ausgabe-Array vorab mit Platz allokieren
        // (headers + positions + group headers + empty lines)
        const totalRows = 1 + positions.length + bwaGruppen.length + (bwaGruppen.length - 1);
        const outputRows = new Array(totalRows);

        // Header als erste Zeile hinzufügen
        outputRows[0] = headerRow;

        // Ausgabe mit Gruppenhierarchie erstellen
        let rowIndex = 1;
        let posIndex = 0;

        for (let gruppenIndex = 0; gruppenIndex < bwaGruppen.length; gruppenIndex++) {
            const gruppe = bwaGruppen[gruppenIndex];

            // Gruppenüberschrift
            outputRows[rowIndex++] = [
                `${gruppenIndex + 1}. ${gruppe.titel}`,
                ...Array(headerRow.length - 1).fill(''),
            ];

            // Gruppenpositionen
            for (let i = 0; i < gruppe.count; i++) {
                outputRows[rowIndex++] = buildOutputRow(positions[posIndex++], bwaData);
            }

            // Leerzeile nach jeder Gruppe außer der letzten
            if (gruppenIndex < bwaGruppen.length - 1) {
                outputRows[rowIndex++] = Array(headerRow.length).fill('');
            }
        }

        // BWA-Sheet erstellen oder aktualisieren
        const bwaSheet = sheetUtils.getOrCreateSheet(ss, 'BWA');
        bwaSheet.clearContents();

        // Optimierung: Daten in einem Batch-API-Call schreiben
        bwaSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);

        // Formatierungen in optimierten Batches anwenden
        applyBwaFormatting(bwaSheet, headerRow.length, bwaGruppen, outputRows.length);

        // BWA-Sheet aktivieren
        ss.setActiveSheet(bwaSheet);
        console.log('BWA sheet generated successfully');

        return true;
    } catch (e) {
        console.error('Fehler bei der BWA-Erstellung:', e);
        return false;
    }
}

/**
 * Wendet Formatierungen auf das BWA-Sheet an mit optimierten Batch-Formatierungen
 * @param {Sheet} sheet - Das zu formatierende Sheet
 * @param {number} headerLength - Anzahl der Spalten
 * @param {Array} bwaGruppen - Gruppenhierarchie
 * @param {number} totalRows - Gesamtzahl der Zeilen
 */
function applyBwaFormatting(sheet, headerLength, bwaGruppen, totalRows) {
    // Optimierung: Formatierungen in Gruppen für Batch-Anwendung
    const formatGroups = {
        // Header-Formatierung
        header: {
            range: sheet.getRange(1, 1, 1, headerLength),
            formats: {
                fontWeight: 'bold',
                background: '#f3f3f3',
            },
        },

        // Gruppentitel-Formatierungen
        groupTitles: [],

        // Summen-Zeilen Formatierungen
        summaryRows: [],

        // EBIT und Jahresüberschuss-Formatierungen
        highlight: [],
    };

    // Gruppentitel-Zeilen finden
    let rowIndex = 2;
    for (const gruppe of bwaGruppen) {
        formatGroups.groupTitles.push(sheet.getRange(rowIndex, 1));
        rowIndex += gruppe.count + 1; // +1 für die Leerzeile
    }

    // Summen-Zeilen definieren basierend auf den BWA-Zeilen
    const summenZeilen = [11, 15, 26, 33, 36, 38, 39, 46];
    summenZeilen.forEach(row => {
        if (row <= totalRows) {
            formatGroups.summaryRows.push(sheet.getRange(row, 1, 1, headerLength));
        }
    });

    // EBIT und Jahresüberschuss hervorheben
    [39, 46].forEach(row => {
        if (row <= totalRows) {
            formatGroups.highlight.push(sheet.getRange(row, 1, 1, headerLength));
        }
    });

    // Optimierte Batch-Formatierungen anwenden

    // 1. Header formatieren
    formatGroups.header.range
        .setFontWeight(formatGroups.header.formats.fontWeight)
        .setBackground(formatGroups.header.formats.background);

    // 2. Gruppentitel formatieren
    if (formatGroups.groupTitles.length > 0) {
        formatGroups.groupTitles.forEach(range => {
            range.setFontWeight('bold');
        });
    }

    // 3. Summenzeilen formatieren
    if (formatGroups.summaryRows.length > 0) {
        formatGroups.summaryRows.forEach(range => {
            range.setBackground('#e6f2ff');
        });
    }

    // 4. EBIT und Jahresüberschuss hervorheben
    if (formatGroups.highlight.length > 0) {
        formatGroups.highlight.forEach(range => {
            range.setFontWeight('bold');
        });
    }

    // 5. Währungsformat für alle Zahlenwerte in einem Batch
    sheet.getRange(2, 2, totalRows - 1, headerLength - 1).setNumberFormat('#,##0.00 €');

    // 6. Spaltenbreiten anpassen
    sheet.autoResizeColumns(1, headerLength);
}

export default {
    buildHeaderRow,
    buildOutputRow,
    generateBWASheet,
    applyBwaFormatting,
};
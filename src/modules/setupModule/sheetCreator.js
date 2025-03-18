// src/modules/setupModule/sheetCreator.js
import stringUtils from '../../utils/stringUtils.js';
import cellValidator from '../validatorModule/cellValidator.js';


/**
 * Erstellt oder aktualisiert ein Sheet mit Header-Zeile und Formatierung
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {string} sheetName - Name des Sheets
 * @param {Array} headers - Array mit Header-Bezeichnungen
 * @param {Object} config - Die Konfiguration
 * @returns {Sheet} - Das erstellte oder aktualisierte Sheet
 */
function createOrUpdateSheet(ss, sheetName, headers, config) {
    let sheet = ss.getSheetByName(sheetName);
    const isNew = !sheet;

    if (isNew) {
        // Neues Sheet erstellen
        sheet = ss.insertSheet(sheetName);
    }

    // Header-Zeile setzen
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Header formatieren
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');

    // Spaltenbreiten anpassen
    sheet.autoResizeColumns(1, headers.length);

    // Konditionale Formatierung oder andere Sheet-spezifische Einstellungen
    applySheetSpecificSettings(sheet, sheetName, config);

    return sheet;
}

/**
 * Wendet spezifische Einstellungen je nach Sheet-Typ an
 * @param {Sheet} sheet - Das zu formatierende Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {Object} config - Die Konfiguration
 */
function applySheetSpecificSettings(sheet, sheetName, config) {
    // Sheet-spezifische Formatierungen anwenden
    switch (sheetName) {
    case 'Einnahmen':
    case 'Ausgaben':
    case 'Eigenbelege':
        applyDocumentSheetSettings(sheet, sheetName, config);
        break;
    case 'Bankbewegungen':
        applyBankSheetSettings(sheet, config);
        break;
    case 'Gesellschafterkonto':
        applyGesellschafterSheetSettings(sheet, config);
        break;
    case 'Holding Transfers':
        applyHoldingSheetSettings(sheet, config);
        break;
    case 'Änderungshistorie':
        // Einfache Tabelle ohne besondere Formatierung
        break;
    }

    // Validierung für gemeinsame Felder wie Datum, Zahlungsart, etc.
    addCommonValidations(sheet, sheetName, config);
}

/**
 * Wendet Einstellungen für Dokument-Sheets (Einnahmen, Ausgaben, Eigenbelege) an
 * @param {Sheet} sheet - Das zu formatierende Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {Object} config - Die Konfiguration
 */
function applyDocumentSheetSettings(sheet, sheetName, config) {
    const configKey = sheetName.toLowerCase().replace(/\s+/g, '');
    const columns = config[configKey].columns;

    // Währungsformat für Beträge
    const currencyColumns = [
        columns.nettobetrag,
        columns.mwstBetrag,
        columns.bruttoBetrag,
        columns.bezahlt,
        columns.restbetragNetto,
    ];

    currencyColumns.forEach(col => {
        if (col) {
            sheet.getRange(2, col, 1000, 1).setNumberFormat('#,##0.00 €');
        }
    });

    // Prozentformat für MwSt-Satz
    if (columns.mwstSatz) {
        sheet.getRange(2, columns.mwstSatz, 1000, 1).setNumberFormat('0.00%');
    }

    // Datumsformat für Datum und Zahlungsdatum
    const dateColumns = [columns.datum, columns.zahlungsdatum];
    dateColumns.forEach(col => {
        if (col) {
            sheet.getRange(2, col, 1000, 1).setNumberFormat('dd.mm.yyyy');
        }
    });

    // Bedingte Formatierung für Zahlungsstatus
    if (columns.zahlungsstatus) {
        const statusCol = stringUtils.getColumnLetter(columns.zahlungsstatus);
        const conditions = [];

        if (sheetName === 'Eigenbelege') {
            conditions.push(
                {value: 'Offen', background: '#FFC7CE', fontColor: '#9C0006'},
                {value: 'Teilerstattet', background: '#FFEB9C', fontColor: '#9C6500'},
                {value: 'Erstattet', background: '#C6EFCE', fontColor: '#006100'},
            );
        } else {
            conditions.push(
                {value: 'Offen', background: '#FFC7CE', fontColor: '#9C0006'},
                {value: 'Teilbezahlt', background: '#FFEB9C', fontColor: '#9C6500'},
                {value: 'Bezahlt', background: '#C6EFCE', fontColor: '#006100'},
            );
        }

        sheet.getRange(`${statusCol}2:${statusCol}1000`).clearFormat();
        conditions.forEach(condition => {
            const rule = SpreadsheetApp.newConditionalFormatRule()
                .whenTextEqualTo(condition.value)
                .setBackground(condition.background)
                .setFontColor(condition.fontColor)
                .setRanges([sheet.getRange(`${statusCol}2:${statusCol}1000`)])
                .build();

            const rules = sheet.getConditionalFormatRules();
            rules.push(rule);
            sheet.setConditionalFormatRules(rules);
        });
    }
}

/**
 * Wendet Einstellungen für das Bankbewegungen-Sheet an
 * @param {Sheet} sheet - Das zu formatierende Sheet
 * @param {Object} config - Die Konfiguration
 */
function applyBankSheetSettings(sheet, config) {
    const columns = config.bankbewegungen.columns;
    const anfangssaldoDate = new Date(config.tax.year, 0, 1);
    // Anfangssaldo-Zeile
    const initialBalanceRow = [
        anfangssaldoDate, // Datum
        'Anfangssaldo', // Buchungstext
        '', // Betrag
        0, // Saldo
        '', // Transaktionstyp
        '', // Kategorie
        '', // Konto (Soll)
        '', // Gegenkonto (Haben)
        '', // Referenz
        '', // Verwendungszweck
        '', // Match-Info
        '', // Zeitstempel
    ];

    sheet.getRange(2, 1, 1, initialBalanceRow.length).setValues([initialBalanceRow]);

    // Währungsformat für Betrag und Saldo
    if (columns.betrag) {
        sheet.getRange(2, columns.betrag, 1000, 1).setNumberFormat('#,##0.00 €');
    }

    if (columns.saldo) {
        sheet.getRange(2, columns.saldo, 1000, 1).setNumberFormat('#,##0.00 €');
    }

    // Bedingte Formatierung für Transaktionstyp
    if (columns.transaktionstyp) {
        const typeCol = stringUtils.getColumnLetter(columns.transaktionstyp);
        const conditions = [
            {value: 'Einnahme', background: '#C6EFCE', fontColor: '#006100'},
            {value: 'Ausgabe', background: '#FFC7CE', fontColor: '#9C0006'},
        ];

        sheet.getRange(`${typeCol}3:${typeCol}1000`).clearFormat();
        conditions.forEach(condition => {
            const rule = SpreadsheetApp.newConditionalFormatRule()
                .whenTextEqualTo(condition.value)
                .setBackground(condition.background)
                .setFontColor(condition.fontColor)
                .setRanges([sheet.getRange(`${typeCol}3:${typeCol}1000`)])
                .build();

            const rules = sheet.getConditionalFormatRules();
            rules.push(rule);
            sheet.setConditionalFormatRules(rules);
        });
    }

    // Anfangssaldo-Zeile formatieren
    sheet.getRange(2, 1, 1, sheet.getLastColumn())
        .setBackground('#e6f2ff')
        .setFontWeight('bold');
}

/**
 * Wendet Einstellungen für das Gesellschafterkonto-Sheet an
 * @param {Sheet} sheet - Das zu formatierende Sheet
 * @param {Object} config - Die Konfiguration
 */
function applyGesellschafterSheetSettings(sheet, config) {
    const columns = config.gesellschafterkonto.columns;

    // Währungsformat für Betrag
    if (columns.betrag) {
        sheet.getRange(2, columns.betrag, 1000, 1).setNumberFormat('#,##0.00 €');
    }

    // Datumsformat für Datum und Zahlungsdatum
    const dateColumns = [columns.datum, columns.zahlungsdatum];
    dateColumns.forEach(col => {
        if (col) {
            sheet.getRange(2, col, 1000, 1).setNumberFormat('dd.mm.yyyy');
        }
    });
}

/**
 * Wendet Einstellungen für das Holding Transfers-Sheet an
 * @param {Sheet} sheet - Das zu formatierende Sheet
 * @param {Object} config - Die Konfiguration
 */
function applyHoldingSheetSettings(sheet, config) {
    const columns = config.holdingTransfers.columns;

    // Währungsformat für Betrag
    if (columns.betrag) {
        sheet.getRange(2, columns.betrag, 1000, 1).setNumberFormat('#,##0.00 €');
    }

    // Datumsformat für Datum und Zahlungsdatum
    const dateColumns = [columns.datum, columns.zahlungsdatum];
    dateColumns.forEach(col => {
        if (col) {
            sheet.getRange(2, col, 1000, 1).setNumberFormat('dd.mm.yyyy');
        }
    });
}

/**
 * Fügt Validierungen für gemeinsame Felder hinzu
 * @param {Sheet} sheet - Das zu validierende Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {Object} config - Die Konfiguration
 */
function addCommonValidations(sheet, sheetName, config) {
    let configKey;

    switch (sheetName) {
    case 'Einnahmen':
        configKey = 'einnahmen';
        break;
    case 'Ausgaben':
        configKey = 'ausgaben';
        break;
    case 'Eigenbelege':
        configKey = 'eigenbelege';
        break;
    case 'Bankbewegungen':
        configKey = 'bankbewegungen';
        break;
    case 'Gesellschafterkonto':
        configKey = 'gesellschafterkonto';
        break;
    case 'Holding Transfers':
        configKey = 'holdingTransfers';
        break;
    default:
        return;
    }

    const columns = config[configKey].columns;

    // Zahlungsart-Dropdown
    if (columns.zahlungsart) {
        cellValidator.validateDropdown(
            sheet, 2, columns.zahlungsart, 1000, 1,
            config.common.paymentType,
        );
    }

    // Kategorien-Dropdown
    if (columns.kategorie && config[configKey].categories) {
        cellValidator.validateDropdown(
            sheet, 2, columns.kategorie, 1000, 1,
            Object.keys(config[configKey].categories),
        );
    }

    // Weitere spezifische Validierungen je nach Sheet-Typ
    if (sheetName === 'Eigenbelege' && columns.ausgelegtVon) {
        // Dropdown für "Ausgelegt von"
        const ausleger = [
            ...config.common.shareholders,
            ...config.common.employees,
        ];

        cellValidator.validateDropdown(
            sheet, 2, columns.ausgelegtVon, 1000, 1,
            ausleger,
        );
    }

    if (sheetName === 'Gesellschafterkonto' && columns.gesellschafter) {
        // Dropdown für Gesellschafter
        cellValidator.validateDropdown(
            sheet, 2, columns.gesellschafter, 1000, 1,
            config.common.shareholders,
        );
    }

    if (sheetName === 'Holding Transfers' && columns.ziel) {
        // Dropdown für Zielgesellschaft
        cellValidator.validateDropdown(
            sheet, 2, columns.ziel, 1000, 1,
            config.common.companies,
        );
    }
}

/**
 * Erstellt alle notwendigen Sheets für die Buchhaltung
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 */
function createAllSheets(ss, config) {
    // Daten-Sheets erstellen
    createDataSheets(ss, config);

    // Bankbewegungen-Sheet erstellen
    createBankSheet(ss, config);

    // Historien-Sheet erstellen
    createHistorySheet(ss, config);

    // Ergebnisblätter erstellen (UStVA, BWA, Bilanz)
    createResultSheets(ss, config);
}

/**
 * Erstellt die Daten-Sheets (Einnahmen, Ausgaben, Eigenbelege, Gesellschafterkonto, Holding Transfers)
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 */
function createDataSheets(ss, config) {
    // Einnahmen-Sheet erstellen
    const einnahmenHeaders = Object.keys(config.einnahmen.columns).map(key => {
        return getHeaderLabel(key, 'einnahmen');
    });
    createOrUpdateSheet(ss, 'Einnahmen', einnahmenHeaders, config);

    // Ausgaben-Sheet erstellen
    const ausgabenHeaders = Object.keys(config.ausgaben.columns).map(key => {
        return getHeaderLabel(key, 'ausgaben');
    });
    createOrUpdateSheet(ss, 'Ausgaben', ausgabenHeaders, config);

    // Eigenbelege-Sheet erstellen
    const eigenbelegeHeaders = Object.keys(config.eigenbelege.columns).map(key => {
        return getHeaderLabel(key, 'eigenbelege');
    });
    createOrUpdateSheet(ss, 'Eigenbelege', eigenbelegeHeaders, config);

    // Gesellschafterkonto-Sheet erstellen
    const gesellschafterHeaders = Object.keys(config.gesellschafterkonto.columns).map(key => {
        return getHeaderLabel(key, 'gesellschafterkonto');
    });
    createOrUpdateSheet(ss, 'Gesellschafterkonto', gesellschafterHeaders, config);

    // Holding Transfers-Sheet erstellen
    const holdingHeaders = Object.keys(config.holdingTransfers.columns).map(key => {
        return getHeaderLabel(key, 'holdingTransfers');
    });
    createOrUpdateSheet(ss, 'Holding Transfers', holdingHeaders, config);
}

/**
 * Erstellt das Bankbewegungen-Sheet
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 */
function createBankSheet(ss, config) {
    const bankHeaders = Object.keys(config.bankbewegungen.columns).map(key => {
        return getHeaderLabel(key, 'bankbewegungen');
    });
    createOrUpdateSheet(ss, 'Bankbewegungen', bankHeaders, config);
}

/**
 * Erstellt das Änderungshistorie-Sheet
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 */
function createHistorySheet(ss, config) {
    const historyHeaders = Object.keys(config.aenderungshistorie.columns).map(key => {
        return getHeaderLabel(key, 'aenderungshistorie');
    });
    createOrUpdateSheet(ss, 'Änderungshistorie', historyHeaders, config);
}

/**
 * Erstellt die Ergebnisblätter (UStVA, BWA, Bilanz)
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 */
function createResultSheets(ss, config) {
    // UStVA-Sheet - einfache Version, die später durch UStVAModule erweitert wird
    const ustvaHeaders = [
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
    ];
    createOrUpdateSheet(ss, 'UStVA', ustvaHeaders, config);

    // BWA-Sheet - einfache Version, die später durch BWAModule erweitert wird
    const bwaHeaders = [
        'Kategorie',
        'Januar (€)',
        'Februar (€)',
        'März (€)',
        'Q1 (€)',
        'April (€)',
        'Mai (€)',
        'Juni (€)',
        'Q2 (€)',
        'Juli (€)',
        'August (€)',
        'September (€)',
        'Q3 (€)',
        'Oktober (€)',
        'November (€)',
        'Dezember (€)',
        'Q4 (€)',
        'Jahr (€)',
    ];
    createOrUpdateSheet(ss, 'BWA', bwaHeaders, config);

    // Bilanz-Sheet - einfache Version, die später durch BilanzModule erweitert wird
    const bilanzHeaders = ['Position', 'Wert'];
    const bilanzSheet = createOrUpdateSheet(ss, 'Bilanz', bilanzHeaders, config);

    // Anfangsstruktur für Bilanz
    const bilanzRows = [
        ['Aktiva (Vermögenswerte)', ''],
        ['', ''],
        ['1. Anlagevermögen', ''],
        ['1.1 Sachanlagen', ''],
        ['1.2 Immaterielle Vermögensgegenstände', ''],
        ['1.3 Finanzanlagen', ''],
        ['Summe Anlagevermögen', ''],
        ['', ''],
        ['2. Umlaufvermögen', ''],
        ['2.1 Bankguthaben', ''],
        ['2.2 Kasse', ''],
        ['2.3 Forderungen aus Lieferungen und Leistungen', ''],
        ['2.4 Vorräte', ''],
        ['Summe Umlaufvermögen', ''],
        ['', ''],
        ['3. Rechnungsabgrenzungsposten', ''],
        ['', ''],
        ['Summe Aktiva', ''],
    ];

    bilanzSheet.getRange(1, 1, bilanzRows.length, 2).setValues(bilanzRows);
    bilanzSheet.getRange(1, 4, 1, 1).setValue('Passiva (Kapital und Schulden)');
}

/**
 * Gibt die Überschrift für eine Spalte basierend auf dem Schlüssel und Sheettyp zurück
 * @param {string} key - Spaltenschlüssel aus der Konfiguration
 * @param {string} sheetType - Typ des Sheets
 * @returns {string} - Überschrift für die Spalte
 */
function getHeaderLabel(key, sheetType) {
    // Standardlabels für gemeinsame Spalten
    const commonLabels = {
        datum: 'Datum',
        rechnungsnummer: 'Rechnungsnummer',
        notizen: 'Notizen',
        kategorie: 'Kategorie',
        buchungskonto: 'Buchungskonto',
        nettobetrag: 'Nettobetrag',
        mwstSatz: 'MwSt-Satz',
        mwstBetrag: 'MwSt-Betrag',
        bruttoBetrag: 'Bruttobetrag',
        bezahlt: 'Bereits bezahlt',
        restbetragNetto: 'Restbetrag (Netto)',
        quartal: 'Quartal',
        zahlungsstatus: 'Zahlungsstatus',
        zahlungsart: 'Zahlungsart',
        zahlungsdatum: 'Zahlungsdatum',
        bankabgleich: 'Bankabgleich',
        uebertragJahr: 'Übertrag Jahr',
        zeitstempel: 'Letzte Aktualisierung',
        dateiname: 'Dateiname',
        dateilink: 'Link zur Datei',
    };

    // Sheet-spezifische Labels
    const sheetLabels = {
        einnahmen: {
            kunde: 'Kunde',
            leistungsBeschreibung: 'Leistungsbeschreibung',
        },
        ausgaben: {
            kunde: 'Lieferant',
            leistungsBeschreibung: 'Leistungsbeschreibung',
        },
        eigenbelege: {
            rechnungsnummer: 'Belegnummer',
            ausgelegtVon: 'Ausgelegt von',
            beschreibung: 'Beschreibung',
            bezahlt: 'Bereits erstattet',
            zahlungsstatus: 'Erstattungsstatus',
            zahlungsart: 'Erstattungsart',
            zahlungsdatum: 'Erstattungsdatum',
        },
        bankbewegungen: {
            buchungstext: 'Buchungstext',
            betrag: 'Betrag',
            saldo: 'Saldo',
            transaktionstyp: 'Transaktionstyp',
            kontoSoll: 'Konto (Soll)',
            kontoHaben: 'Gegenkonto (Haben)',
            referenz: 'Referenz',
            verwendungszweck: 'Verwendungszweck',
            matchInfo: 'Match-Information',
        },
        gesellschafterkonto: {
            referenz: 'Referenz',
            gesellschafter: 'Gesellschafter',
            betrag: 'Betrag',
        },
        holdingTransfers: {
            referenz: 'Referenz',
            ziel: 'Zielgesellschaft',
            verwendungszweck: 'Verwendungszweck',
            betrag: 'Betrag',
        },
        aenderungshistorie: {
            datum: 'Datum',
            typ: 'Rechnungstyp',
            dateiname: 'Dateiname',
            dateilink: 'Link zur Datei',
        },
    };

    // Sheet-spezifisches Label verwenden, falls vorhanden, sonst auf gemeinsames Label zurückfallen
    if (sheetLabels[sheetType] && sheetLabels[sheetType][key]) {
        return sheetLabels[sheetType][key];
    }

    return commonLabels[key] || key;
}

export default {
    createOrUpdateSheet,
    createAllSheets,
    createDataSheets,
    createBankSheet,
    createHistorySheet,
    createResultSheets,
};
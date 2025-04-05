// src/modules/setupModule/sheetCreator.js
import stringUtils from '../../utils/stringUtils.js';
import cellValidator from '../validatorModule/cellValidator.js';
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Erstellt oder aktualisiert ein Sheet mit Header-Zeile und Formatierung
 * Optimierte Version mit Batch-Operationen
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

    // Header-Zeile und Formatierung in einem Batch-Operation setzen
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');

    // Spaltenbreiten anpassen für bessere Darstellung
    sheet.autoResizeColumns(1, headers.length);

    // Konditionale Formatierung oder andere Sheet-spezifische Einstellungen
    applySheetSpecificSettings(sheet, sheetName, config);

    return sheet;
}

/**
 * Wendet spezifische Einstellungen je nach Sheet-Typ an
 * Optimierte Version mit Batch-Formatierung
 * @param {Sheet} sheet - Das zu formatierende Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {Object} config - Die Konfiguration
 */
function applySheetSpecificSettings(sheet, sheetName, config) {
    // Mapping für Sheet-spezifische Verarbeitung
    const settingsMap = {
        'Einnahmen': () => applyDocumentSheetSettings(sheet, sheetName, config),
        'Ausgaben': () => applyDocumentSheetSettings(sheet, sheetName, config),
        'Eigenbelege': () => applyDocumentSheetSettings(sheet, sheetName, config),
        'Bankbewegungen': () => applyBankSheetSettings(sheet, config),
        'Gesellschafterkonto': () => applyGesellschafterSheetSettings(sheet, config),
        'Holding Transfers': () => applyHoldingSheetSettings(sheet, config),
    };

    // Spezifische Settings anwenden wenn verfügbar
    if (settingsMap[sheetName]) {
        settingsMap[sheetName]();
    }

    // Validierung für gemeinsame Felder wie Datum, Zahlungsart, etc.
    addCommonValidations(sheet, sheetName, config);
}

/**
 * Wendet Einstellungen für Dokument-Sheets an
 * Optimierte Version mit Batch-Formatierung
 * @param {Sheet} sheet - Das zu formatierende Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {Object} config - Die Konfiguration
 */
function applyDocumentSheetSettings(sheet, sheetName, config) {
    const configKey = sheetName.toLowerCase().replace(/\s+/g, '');
    const columns = config[configKey].columns;

    // Alle Formatierungen in Arrays für Batch-Updates sammeln
    const formatBatches = {
        currency: [],
        percent: [],
        date: [],
    };

    // Währungsformat für Beträge
    [
        columns.nettobetrag,
        columns.mwstBetrag,
        columns.bruttoBetrag,
        columns.bezahlt,
        columns.restbetragNetto,
    ].filter(Boolean).forEach(col => {
        formatBatches.currency.push(col);
    });

    // Prozentformat für MwSt-Satz
    if (columns.mwstSatz) {
        formatBatches.percent.push(columns.mwstSatz);
    }

    // Datumsformat für Datum und Zahlungsdatum
    [columns.datum, columns.zahlungsdatum].filter(Boolean).forEach(col => {
        formatBatches.date.push(col);
    });

    // Alle Formatierungen in einem Batch anwenden
    if (formatBatches.currency.length > 0) {
        formatBatches.currency.forEach(col => {
            sheet.getRange(2, col, 1000, 1).setNumberFormat('#,##0.00 €');
        });
    }

    if (formatBatches.percent.length > 0) {
        formatBatches.percent.forEach(col => {
            sheet.getRange(2, col, 1000, 1).setNumberFormat('0.00%');
        });
    }

    if (formatBatches.date.length > 0) {
        formatBatches.date.forEach(col => {
            sheet.getRange(2, col, 1000, 1).setNumberFormat('dd.mm.yyyy');
        });
    }

    // Bedingte Formatierung für Zahlungsstatus
    if (columns.zahlungsstatus) {
        const statusCol = stringUtils.getColumnLetter(columns.zahlungsstatus);
        applyStatusColumnFormatting(sheet, statusCol, sheetName);
    }
}

/**
 * Wendet bedingte Formatierung für Status-Spalten an
 * @param {Sheet} sheet - Das Sheet
 * @param {string} statusCol - Spalte für Status
 * @param {string} sheetName - Name des Sheets
 */
function applyStatusColumnFormatting(sheet, statusCol, sheetName) {
    // Existierende Regeln löschen, um Konflikte zu vermeiden
    try {
        const range = sheet.getRange(`${statusCol}2:${statusCol}1000`);

        // Formatierungsbedingungen je nach Sheet-Typ
        let conditions;

        if (sheetName === 'Eigenbelege') {
            conditions = [
                {value: 'Offen', background: '#FFC7CE', fontColor: '#9C0006'},
                {value: 'Teilerstattet', background: '#FFEB9C', fontColor: '#9C6500'},
                {value: 'Erstattet', background: '#C6EFCE', fontColor: '#006100'},
            ];
        } else {
            conditions = [
                {value: 'Offen', background: '#FFC7CE', fontColor: '#9C0006'},
                {value: 'Teilbezahlt', background: '#FFEB9C', fontColor: '#9C6500'},
                {value: 'Bezahlt', background: '#C6EFCE', fontColor: '#006100'},
            ];
        }

        // Regeln in einem Batch erstellen
        const rules = sheet.getConditionalFormatRules().filter(rule =>
            rule.getRanges().every(r => r.getA1Notation() !== range.getA1Notation()),
        );

        // Neue Regeln hinzufügen
        conditions.forEach(condition => {
            const rule = SpreadsheetApp.newConditionalFormatRule()
                .whenTextEqualTo(condition.value)
                .setBackground(condition.background)
                .setFontColor(condition.fontColor)
                .setRanges([range])
                .build();
            rules.push(rule);
        });

        // Alle Regeln in einem API-Call setzen
        sheet.setConditionalFormatRules(rules);
    } catch (e) {
        console.error(`Fehler bei der Statusformatierung für ${sheetName}:`, e);
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

    // Anfangssaldo-Zeile in einem Batch setzen
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

    // Erste Datenzeile als Anfangssaldo
    sheet.getRange(2, 1, 1, initialBalanceRow.length).setValues([initialBalanceRow]);

    // Formatierungen für Beträge und Status in einem Batch
    const formatBatches = {
        currency: [],
        conditionalFormat: [],
    };

    // Spalten für Währungsformat sammeln
    if (columns.betrag) formatBatches.currency.push(columns.betrag);
    if (columns.saldo) formatBatches.currency.push(columns.saldo);

    // Batch-Währungsformat anwenden
    formatBatches.currency.forEach(col => {
        sheet.getRange(2, col, 1000, 1).setNumberFormat('#,##0.00 €');
    });

    // Formatierung für Anfangssaldo
    sheet.getRange(2, 1, 1, sheet.getLastColumn())
        .setBackground('#e6f2ff')
        .setFontWeight('bold');

    // Bedingte Formatierung für Transaktionstyp
    if (columns.transaktionstyp) {
        const typeCol = stringUtils.getColumnLetter(columns.transaktionstyp);

        // Regeln in einem Batch erstellen und setzen
        const range = sheet.getRange(`${typeCol}3:${typeCol}1000`);
        const rules = sheet.getConditionalFormatRules().filter(rule =>
            rule.getRanges().every(r => r.getA1Notation() !== range.getA1Notation()),
        );

        // Einnahme-Regel
        const incomeRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Einnahme')
            .setBackground('#C6EFCE')
            .setFontColor('#006100')
            .setRanges([range])
            .build();

        // Ausgabe-Regel
        const expenseRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Ausgabe')
            .setBackground('#FFC7CE')
            .setFontColor('#9C0006')
            .setRanges([range])
            .build();

        rules.push(incomeRule, expenseRule);
        sheet.setConditionalFormatRules(rules);
    }
}

/**
 * Wendet Einstellungen für das Gesellschafterkonto-Sheet an
 * @param {Sheet} sheet - Das zu formatierende Sheet
 * @param {Object} config - Die Konfiguration
 */
function applyGesellschafterSheetSettings(sheet, config) {
    const columns = config.gesellschafterkonto.columns;

    // Geld- und Datumsformatierungen in einem Batch
    if (columns.betrag) {
        sheet.getRange(2, columns.betrag, 1000, 1).setNumberFormat('#,##0.00 €');
    }

    [columns.datum, columns.zahlungsdatum].filter(Boolean).forEach(col => {
        sheet.getRange(2, col, 1000, 1).setNumberFormat('dd.mm.yyyy');
    });
}

/**
 * Wendet Einstellungen für das Holding Transfers-Sheet an
 * @param {Sheet} sheet - Das zu formatierende Sheet
 * @param {Object} config - Die Konfiguration
 */
function applyHoldingSheetSettings(sheet, config) {
    const columns = config.holdingTransfers.columns;

    // Geld- und Datumsformatierungen in einem Batch
    if (columns.betrag) {
        sheet.getRange(2, columns.betrag, 1000, 1).setNumberFormat('#,##0.00 €');
    }

    [columns.datum, columns.zahlungsdatum].filter(Boolean).forEach(col => {
        sheet.getRange(2, col, 1000, 1).setNumberFormat('dd.mm.yyyy');
    });
}

/**
 * Fügt Validierungen für gemeinsame Felder hinzu
 * Optimierte Version mit Batch-Validierung
 * @param {Sheet} sheet - Das zu validierende Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {Object} config - Die Konfiguration
 */
function addCommonValidations(sheet, sheetName, config) {
    // Konfigurationsschlüssel für Sheet ermitteln
    let configKey;
    const sheetTypeMap = {
        'Einnahmen': 'einnahmen',
        'Ausgaben': 'ausgaben',
        'Eigenbelege': 'eigenbelege',
        'Bankbewegungen': 'bankbewegungen',
        'Gesellschafterkonto': 'gesellschafterkonto',
        'Holding Transfers': 'holdingTransfers',
    };

    configKey = sheetTypeMap[sheetName];
    if (!configKey) return;

    const columns = config[configKey].columns;

    // Batch-Validierungen vorbereiten
    const validations = [];

    // 1. Zahlungsart-Dropdown
    if (columns.zahlungsart) {
        validations.push({
            column: columns.zahlungsart,
            values: config.common.paymentType,
        });
    }

    // 2. Kategorien-Dropdown
    if (columns.kategorie && config[configKey].categories) {
        validations.push({
            column: columns.kategorie,
            values: Object.keys(config[configKey].categories),
        });
    }

    // 3. Ausland-Dropdown
    if (columns.ausland) {
        validations.push({
            column: columns.ausland,
            values: config.common.auslandType,
        });
    }

    // 4. Sheet-spezifische Validierungen
    if (sheetName === 'Eigenbelege' && columns.ausgelegtVon) {
        // Dropdown für "Ausgelegt von"
        validations.push({
            column: columns.ausgelegtVon,
            values: [...config.common.shareholders, ...config.common.employees],
        });
    }

    if (sheetName === 'Gesellschafterkonto' && columns.gesellschafter) {
        // Dropdown für Gesellschafter
        validations.push({
            column: columns.gesellschafter,
            values: config.common.shareholders,
        });
    }

    if (sheetName === 'Holding Transfers' && columns.ziel) {
        // Dropdown für Zielgesellschaft
        validations.push({
            column: columns.ziel,
            values: config.common.companies,
        });
    }

    // Alle Validierungen in einem Batch anwenden
    validations.forEach(({column, values}) => {
        if (column && values?.length > 0) {
            cellValidator.validateDropdown(
                sheet, 2, column, 1000, 1,
                values,
            );
        }
    });
}

/**
 * Erstellt alle notwendigen Sheets für die Buchhaltung
 * Optimierte Version mit Statustracking
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} - Erfolg der Operation
 */
function createAllSheets(ss, config) {
    try {
        // Status-Tracking
        const status = {
            success: true,
            createdSheets: [],
            errors: [],
        };

        // Versuche alle Sheet-Typen zu erstellen
        try {
            // 1. Daten-Sheets erstellen
            const dataSuccess = createDataSheets(ss, config);
            status.success = status.success && dataSuccess.success;
            status.createdSheets.push(...dataSuccess.sheets);
            status.errors.push(...dataSuccess.errors);

            // 2. Bankbewegungen-Sheet erstellen
            const bankSuccess = createBankSheet(ss, config);
            status.success = status.success && bankSuccess.success;
            if (bankSuccess.sheet) status.createdSheets.push(bankSuccess.sheet);
            status.errors.push(...bankSuccess.errors);

            // 3. Historien-Sheet erstellen
            const historySuccess = createHistorySheet(ss, config);
            status.success = status.success && historySuccess.success;
            if (historySuccess.sheet) status.createdSheets.push(historySuccess.sheet);
            status.errors.push(...historySuccess.errors);

            // 4. Ergebnisblätter erstellen
            const resultSuccess = createResultSheets(ss, config);
            status.success = status.success && resultSuccess.success;
            status.createdSheets.push(...resultSuccess.sheets);
            status.errors.push(...resultSuccess.errors);

        } catch (e) {
            status.success = false;
            status.errors.push(`Unexpected error: ${e.message}`);
        }

        // Bei Fehlern Benutzer informieren
        if (!status.success && status.errors.length > 0) {
            console.error('Fehler beim Erstellen der Sheets:', status.errors);
            SpreadsheetApp.getUi().alert(
                'Fehler beim Einrichten',
                `Einige Sheets konnten nicht erstellt werden: ${status.errors.join(', ')}`,
                SpreadsheetApp.getUi().ButtonSet.OK,
            );
        }

        return status.success;
    } catch (e) {
        console.error('Schwerwiegender Fehler beim Erstellen der Sheets:', e);
        return false;
    }
}

/**
 * Erstellt die Daten-Sheets
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Status-Objekt
 */
function createDataSheets(ss, config) {
    const status = { success: true, sheets: [], errors: [] };

    // Sheets als Array definieren für einfachere Verarbeitung
    const sheetsToCreate = [
        { name: 'Einnahmen', configKey: 'einnahmen' },
        { name: 'Ausgaben', configKey: 'ausgaben' },
        { name: 'Eigenbelege', configKey: 'eigenbelege' },
        { name: 'Gesellschafterkonto', configKey: 'gesellschafterkonto' },
        { name: 'Holding Transfers', configKey: 'holdingTransfers' },
    ];

    // Alle Sheets in einer Schleife erstellen
    sheetsToCreate.forEach(({name, configKey}) => {
        try {
            const headers = Object.keys(config[configKey].columns).map(key => {
                return getHeaderLabel(key, configKey);
            });

            const sheet = createOrUpdateSheet(ss, name, headers, config);
            if (sheet) {
                status.sheets.push(name);
            } else {
                status.success = false;
                status.errors.push(`Failed to create ${name}`);
            }
        } catch (e) {
            status.success = false;
            status.errors.push(`Error creating ${name}: ${e.message}`);
        }
    });

    return status;
}

/**
 * Erstellt das Bankbewegungen-Sheet
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Status-Objekt
 */
function createBankSheet(ss, config) {
    const status = { success: true, sheet: null, errors: [] };

    try {
        const bankHeaders = Object.keys(config.bankbewegungen.columns).map(key => {
            return getHeaderLabel(key, 'bankbewegungen');
        });

        const sheet = createOrUpdateSheet(ss, 'Bankbewegungen', bankHeaders, config);
        if (sheet) {
            status.sheet = 'Bankbewegungen';
        } else {
            status.success = false;
            status.errors.push('Failed to create Bankbewegungen');
        }
    } catch (e) {
        status.success = false;
        status.errors.push(`Error creating Bankbewegungen: ${e.message}`);
    }

    return status;
}

/**
 * Erstellt das Änderungshistorie-Sheet
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Status-Objekt
 */
function createHistorySheet(ss, config) {
    const status = { success: true, sheet: null, errors: [] };

    try {
        const historyHeaders = Object.keys(config.aenderungshistorie.columns).map(key => {
            return getHeaderLabel(key, 'aenderungshistorie');
        });

        const sheet = createOrUpdateSheet(ss, 'Änderungshistorie', historyHeaders, config);
        if (sheet) {
            status.sheet = 'Änderungshistorie';
        } else {
            status.success = false;
            status.errors.push('Failed to create Änderungshistorie');
        }
    } catch (e) {
        status.success = false;
        status.errors.push(`Error creating Änderungshistorie: ${e.message}`);
    }

    return status;
}

/**
 * Erstellt die Ergebnisblätter (UStVA, BWA, Bilanz)
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Status-Objekt
 */
function createResultSheets(ss, config) {
    const status = { success: true, sheets: [], errors: [] };

    try {
        // UStVA-Sheet erstellen
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
        const ustvaSheet = createOrUpdateSheet(ss, 'UStVA', ustvaHeaders, config);
        if (ustvaSheet) {
            status.sheets.push('UStVA');
        } else {
            status.errors.push('Failed to create UStVA');
        }

        // BWA-Sheet erstellen
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
        const bwaSheet = createOrUpdateSheet(ss, 'BWA', bwaHeaders, config);
        if (bwaSheet) {
            status.sheets.push('BWA');
        } else {
            status.errors.push('Failed to create BWA');
        }

        // Bilanz-Sheet erstellen
        const bilanzHeaders = ['Position', 'Wert'];
        const bilanzSheet = createOrUpdateSheet(ss, 'Bilanz', bilanzHeaders, config);
        if (bilanzSheet) {
            status.sheets.push('Bilanz');

            // Anfangsstruktur für Bilanz in einem Batch setzen
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
        } else {
            status.errors.push('Failed to create Bilanz');
        }

        // Status aktualisieren
        status.success = status.errors.length === 0;

    } catch (e) {
        status.success = false;
        status.errors.push(`Error creating result sheets: ${e.message}`);
    }

    return status;
}

/**
 * Gibt die Überschrift für eine Spalte basierend auf dem Schlüssel und Sheettyp zurück
 * Optimierte Version mit Mapping-Tabellen
 * @param {string} key - Spaltenschlüssel aus der Konfiguration
 * @param {string} sheetType - Typ des Sheets
 * @returns {string} - Überschrift für die Spalte
 */
function getHeaderLabel(key, sheetType) {
    // Gemeinsame Labels für alle Sheet-Typen
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

    // Optimierte Label-Suche: Zuerst im Sheet-spezifischen Label, dann Fallback auf gemeinsame Labels
    return (sheetLabels[sheetType] && sheetLabels[sheetType][key]) ||
        commonLabels[key] ||
        key;
}

export default {
    createOrUpdateSheet,
    createAllSheets,
    createDataSheets,
    createBankSheet,
    createHistorySheet,
    createResultSheets,
};
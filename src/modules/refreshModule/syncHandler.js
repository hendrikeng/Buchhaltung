// src/modules/refreshModule/syncHandler.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';

/**
 * Markiert bezahlte Einnahmen, Ausgaben und Eigenbelege farblich
 */
function markPaidInvoices(ss, bankZuordnungen, config) {
    // Alle relevanten Sheets abrufen
    const sheets = {
        einnahmen: ss.getSheetByName('Einnahmen'),
        ausgaben: ss.getSheetByName('Ausgaben'),
        eigenbelege: ss.getSheetByName('Eigenbelege'),
        gesellschafterkonto: ss.getSheetByName('Gesellschafterkonto'),
        holdingTransfers: ss.getSheetByName('Holding Transfers'),
    };

    // Für jedes Sheet Zahlungsdaten aktualisieren
    Object.entries(sheets).forEach(([type, sheet]) => {
        if (sheet && sheet.getLastRow() > 1) {
            markPaidRows(sheet, type, bankZuordnungen, config);
        }
    });
}

/**
 * Markiert bezahlte Zeilen in einem Sheet
 */
function markPaidRows(sheet, sheetType, bankZuordnungen, config) {
    // Konfiguration für das Sheet
    const columns = config[sheetType].columns;

    // Hole Werte aus dem Sheet
    const numRows = sheet.getLastRow() - 1;
    if (numRows <= 0) return;

    // Bestimme die maximale Anzahl an Spalten, die benötigt werden
    const maxCol = Math.max(
        columns.zahlungsdatum || 0,
        columns.bankabgleich || 0,
    );

    if (maxCol === 0) return; // Keine relevanten Spalten verfügbar

    const data = sheet.getRange(2, 1, numRows, maxCol).getValues();

    // Batch-Arrays für die verschiedenen Status
    const rowCategories = {
        fullPaidWithBankRows: [],
        fullPaidRows: [],
        partialPaidWithBankRows: [],
        partialPaidRows: [],
        gutschriftRows: [],
        normalRows: [],
    };

    // Bank-Abgleich-Updates sammeln
    const bankabgleichUpdates = [];

    // Datenzeilen durchgehen und in Kategorien einteilen
    for (let i = 0; i < data.length; i++) {
        const row = i + 2; // Aktuelle Zeile im Sheet
        const rowData = categorizeRow(data[i], row, columns, sheetType, bankZuordnungen, config);

        // Zu entsprechender Kategorie hinzufügen
        if (rowData.category) {
            rowCategories[rowData.category].push(row);
        }

        // Bank-Abgleich-Info hinzufügen wenn vorhanden
        if (rowData.bankInfo) {
            bankabgleichUpdates.push({
                row,
                value: rowData.bankInfo,
            });
        }
    }

    // Färbe Zeilen im Batch-Modus
    applyColorToRowCategories(sheet, rowCategories);

    // Bank-Abgleich-Updates in Batches ausführen
    applyBankInfoUpdates(sheet, bankabgleichUpdates, columns);
}

/**
 * Kategorisiert eine Zeile basierend auf ihren Daten
 */
function categorizeRow(rowData, rowIndex, columns, sheetType, bankZuordnungen, config) {
    const nettobetrag = numberUtils.parseCurrency(rowData[columns.nettobetrag - 1]);
    const bezahltBetrag = numberUtils.parseCurrency(rowData[columns.bezahlt - 1]);
    const zahlungsDatum = rowData[columns.zahlungsdatum - 1];
    const referenz = rowData[columns.rechnungsnummer - 1];

    // Prüfe, ob es eine Gutschrift ist
    const isGutschrift = referenz && referenz.toString().startsWith('G-');

    // Prüfe, ob diese Position im Banking-Sheet zugeordnet wurde
    const prefix = getDocumentTypePrefix(sheetType);
    const zuordnungsKey = `${prefix}#${rowIndex}`;
    const hatBankzuordnung = bankZuordnungen[zuordnungsKey] !== undefined;

    // Zahlungsstatus berechnen
    const mwst = columns.mwstSatz ?
        numberUtils.parseMwstRate(rowData[columns.mwstSatz - 1], config.tax.defaultMwst) / 100 : 0;
    const bruttoBetrag = nettobetrag * (1 + mwst);
    const isPaid = Math.abs(bezahltBetrag) >= Math.abs(bruttoBetrag) * 0.999; // 99.9% bezahlt wegen Rundungsfehlern
    const isPartialPaid = !isPaid && bezahltBetrag > 0;

    // Bankabgleich-Info aus der Zuordnung
    const bankInfo = hatBankzuordnung ?
        getZuordnungsInfo(bankZuordnungen[zuordnungsKey]) :
        undefined;

    // Kategorie bestimmen
    let category = null;

    if (isGutschrift) {
        category = 'gutschriftRows';
    } else if (isPaid) {
        if (zahlungsDatum) {
            category = hatBankzuordnung ? 'fullPaidWithBankRows' : 'fullPaidRows';
        } else {
            // Bezahlt aber kein Datum
            category = 'partialPaidRows';
        }
    } else if (isPartialPaid) {
        category = hatBankzuordnung ? 'partialPaidWithBankRows' : 'partialPaidRows';
    } else {
        // Unbezahlt - normale Zeile
        category = 'normalRows';
    }

    return { category, bankInfo };
}

/**
 * Ermittelt das Präfix für einen Dokumenttyp
 */
function getDocumentTypePrefix(sheetType) {
    return sheetType === 'eigenbelege' ? 'eigenbeleg' :
        sheetType === 'gesellschafterkonto' ? 'gesellschafterkonto' :
            sheetType === 'holdingTransfers' ? 'holdingtransfer' : sheetType;
}

/**
 * Wendet Farben auf alle Zeilenkategorien an
 */
function applyColorToRowCategories(sheet, rowCategories) {
    const colorMap = {
        fullPaidWithBankRows: '#C6EFCE',    // Kräftigeres Grün
        fullPaidRows: '#EAF1DD',            // Helles Grün
        partialPaidWithBankRows: '#FFC7AA', // Kräftigeres Orange
        partialPaidRows: '#FCE4D6',         // Helles Orange
        gutschriftRows: '#E6E0FF',          // Helles Lila
        normalRows: null,                    // Keine Farbe / Zurücksetzen
    };

    Object.entries(rowCategories).forEach(([category, rows]) => {
        if (rows.length > 0) {
            applyColorToRows(sheet, rows, colorMap[category]);
        }
    });
}

/**
 * Wendet eine Farbe auf mehrere Zeilen an
 */
function applyColorToRows(sheet, rows, color) {
    if (!rows || rows.length === 0) return;

    // Verarbeite in Batches von maximal 20 Zeilen
    const batchSize = 20;
    for (let i = 0; i < rows.length; i += batchSize) {
        const batchRows = rows.slice(i, i + batchSize);
        batchRows.forEach(row => {
            try {
                const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
                if (color) {
                    range.setBackground(color);
                } else {
                    range.setBackground(null);
                }
            } catch (e) {
                console.error(`Fehler beim Formatieren von Zeile ${row}:`, e);
            }
        });

        // Kurze Pause um API-Limits zu vermeiden
        if (i + batchSize < rows.length) {
            Utilities.sleep(50);
        }
    }
}

/**
 * Führt Bank-Info-Updates in Batches aus
 */
function applyBankInfoUpdates(sheet, updates, columns) {
    if (updates.length === 0 || !columns.bankabgleich) return;

    // Gruppiere Updates nach Wert für effizientere Batch-Updates
    const groupedUpdates = {};

    updates.forEach(update => {
        const { row, value } = update;
        if (!groupedUpdates[value]) {
            groupedUpdates[value] = [];
        }
        groupedUpdates[value].push(row);
    });

    // Führe Batch-Updates für jede Gruppe aus
    Object.entries(groupedUpdates).forEach(([value, rows]) => {
        // Verarbeite in Batches von maximal 20 Zeilen
        const batchSize = 20;
        for (let i = 0; i < rows.length; i += batchSize) {
            const batchRows = rows.slice(i, i + batchSize);
            batchRows.forEach(row => {
                sheet.getRange(row, columns.bankabgleich).setValue(value);
            });

            // Kurze Pause, um API-Limits zu vermeiden
            if (i + batchSize < rows.length) {
                Utilities.sleep(50);
            }
        }
    });
}

/**
 * Erstellt einen Informationstext für eine Bank-Zuordnung
 */
function getZuordnungsInfo(zuordnung) {
    if (!zuordnung) return '';

    let infoText = '✓ Bank: ' + zuordnung.bankDatum;

    // Wenn es mehrere Zuordnungen gibt (z.B. bei aufgeteilten Zahlungen)
    if (zuordnung.additional && zuordnung.additional.length > 0) {
        infoText += ' + ' + zuordnung.additional.length + ' weitere';
    }

    return infoText;
}

export default {
    markPaidInvoices,
    markPaidRows,
    getZuordnungsInfo,
};
// src/modules/refreshModule/syncHandler.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';

/**
 * Markiert bezahlte Einnahmen, Ausgaben und Eigenbelege farblich
 * und aktualisiert den Zahlungsstatus
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
        sheet.getLastColumn(),  // Use the full sheet to ensure we get all data
        columns.zahlungsdatum || 0,
        columns.bankabgleich || 0,
        columns.zahlungsart || 0,
        columns.bezahlt || 0,
        columns.bruttoBetrag || 0,
        columns.zahlungsstatus || 0,
        columns.nettobetrag || 0,  // Important for gutschrift detection based on negative amount
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

    // Updates für verschiedene Felder sammeln
    const bankabgleichUpdates = [];
    const zahlungsdatumUpdates = [];
    const zahlungsartUpdates = [];
    const bezahltUpdates = [];

    // Datenzeilen durchgehen und in Kategorien einteilen
    for (let i = 0; i < data.length; i++) {
        const row = i + 2; // Aktuelle Zeile im Sheet
        const rowData = categorizeRow(data[i], row, columns, sheetType, bankZuordnungen, config);

        // Zu entsprechender Kategorie hinzufügen
        if (rowData.category) {
            rowCategories[rowData.category].push(row);
        }

        // Updates für verschiedene Felder hinzufügen
        if (rowData.bankInfo && columns.bankabgleich) {
            bankabgleichUpdates.push({
                row,
                value: rowData.bankInfo,
            });
        }

        // Zahlungsdatum setzen, wenn eine Bankbuchung gefunden wurde
        if (rowData.bankDatum && columns.zahlungsdatum) {
            // Always update with bank date when a bank transaction is found
            zahlungsdatumUpdates.push({
                row,
                value: Utilities.formatDate(
                    new Date(rowData.bankDatum),
                    Session.getScriptTimeZone(),
                    'dd.MM.yyyy',
                ),
            });
        }

        // Zahlungsart auf "Überweisung" setzen, wenn eine Bankbuchung gefunden wurde
        if (rowData.bankFound && columns.zahlungsart) {
            zahlungsartUpdates.push({
                row,
                value: 'Überweisung',
            });
        }

        // Bezahlt-Betrag aktualisieren, wenn der Betrag eine Bankbuchung ist
        if (rowData.bezahltBetrag !== null && columns.bezahlt) {
            // Nur aktualisieren, wenn der neue Betrag größer ist
            const currentBezahlt = numberUtils.parseCurrency(data[i][columns.bezahlt - 1]);
            if (rowData.bezahltBetrag > currentBezahlt) {
                bezahltUpdates.push({
                    row,
                    value: rowData.bezahltBetrag,
                });
            }
        }
    }

    // Färbe Zeilen im Batch-Modus
    applyColorToRowCategories(sheet, rowCategories);

    // Bank-Abgleich-Updates in Batches ausführen
    if (bankabgleichUpdates.length > 0) {
        applyUpdates(sheet, bankabgleichUpdates, columns.bankabgleich);
    }

    // Zahlungsdatum-Updates ausführen
    if (zahlungsdatumUpdates.length > 0) {
        applyUpdates(sheet, zahlungsdatumUpdates, columns.zahlungsdatum);
    }

    // Zahlungsart-Updates ausführen
    if (zahlungsartUpdates.length > 0) {
        applyUpdates(sheet, zahlungsartUpdates, columns.zahlungsart);
    }

    // Bezahlt-Updates ausführen
    if (bezahltUpdates.length > 0) {
        applyUpdates(sheet, bezahltUpdates, columns.bezahlt);
    }
}

/**
 * Kategorisiert eine Zeile basierend auf ihren Daten
 */
function categorizeRow(rowData, rowIndex, columns, sheetType, bankZuordnungen, config) {
    const nettobetrag = columns.nettobetrag ? numberUtils.parseCurrency(rowData[columns.nettobetrag - 1]) : 0;
    const bezahltBetrag = columns.bezahlt ? numberUtils.parseCurrency(rowData[columns.bezahlt - 1]) : 0;
    const zahlungsDatum = columns.zahlungsdatum ? rowData[columns.zahlungsdatum - 1] : null;

    // Prüfe, ob es eine Gutschrift ist - jetzt basierend auf negativem Betrag
    const isGutschrift = nettobetrag < 0;

    // Prüfe, ob diese Position im Banking-Sheet zugeordnet wurde
    const prefix = getDocumentTypePrefix(sheetType);
    const zuordnungsKey = `${prefix}#${rowIndex}`;
    const bankzuordnung = bankZuordnungen[zuordnungsKey];
    const hatBankzuordnung = bankzuordnung !== undefined;

    // Zahlungsstatus berechnen - Betrag immer als absoluten Wert verwenden
    const nettoAbs = Math.abs(nettobetrag);
    const mwst = columns.mwstSatz ?
        numberUtils.parseMwstRate(rowData[columns.mwstSatz - 1], config.tax.defaultMwst) / 100 : 0;
    const bruttoBetrag = nettoAbs * (1 + mwst);
    const isPaid = Math.abs(bezahltBetrag) >= bruttoBetrag * 0.999; // 99.9% bezahlt wegen Rundungsfehlern
    const isPartialPaid = !isPaid && bezahltBetrag > 0;

    // Bankabgleich-Info und Datum aus der Zuordnung
    const bankInfo = hatBankzuordnung ? getZuordnungsInfo(bankzuordnung) : undefined;
    const bankDatum = hatBankzuordnung ? bankzuordnung.bankDatum : undefined;

    // Bezahlt-Betrag berechnen
    let newBezahltBetrag = null;
    if (hatBankzuordnung) {
        // Wenn Bankzuordnung gefunden wurde und es eine Vollzahlung ist, setze den Bruttobetrag
        if (bankzuordnung.matchInfo.includes('Vollständige Zahlung')) {
            newBezahltBetrag = bruttoBetrag;
        }
        // Bei Teilzahlung den aktuellen Stand beibehalten oder erhöhen
        else if (bankzuordnung.matchInfo.includes('Teilzahlung')) {
            newBezahltBetrag = Math.max(bezahltBetrag, bruttoBetrag * 0.5); // Mindestens 50%
        }
    }

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

    return {
        category,
        bankInfo,
        bankDatum,
        bankFound: hatBankzuordnung,
        bezahltBetrag: newBezahltBetrag,
        isGutschrift: isGutschrift,
    };
}

/**
 * Ermittelt das Präfix für einen Dokumenttyp
 */
function getDocumentTypePrefix(sheetType) {
    const prefixMap = {
        'einnahmen': 'einnahme',
        'ausgaben': 'ausgabe',
        'eigenbelege': 'eigenbeleg',
        'gesellschafterkonto': 'gesellschafterkonto',
        'holdingTransfers': 'holdingtransfer',
    };
    return prefixMap[sheetType] || sheetType;
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
 * Führt Updates für eine Spalte in Batches aus
 */
function applyUpdates(sheet, updates, columnIndex) {
    if (updates.length === 0 || !columnIndex) return;

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
                sheet.getRange(row, columnIndex).setValue(value);
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

    let infoText = '✓ Bank: ';

    // Format the date to Berlin timezone
    try {
        const date = zuordnung.bankDatum;
        if (date) {
            // Format as DD.MM.YYYY
            const formattedDate = Utilities.formatDate(
                new Date(date),
                Session.getScriptTimeZone(),
                'dd.MM.yyyy',
            );
            infoText += formattedDate;
        } else {
            infoText += 'Datum unbekannt';
        }
    } catch (e) {
        infoText += String(zuordnung.bankDatum);
    }

    // Additional info if available
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
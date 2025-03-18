// src/modules/refreshModule/syncHandler.js
import numberUtils from '../../utils/numberUtils.js';
import dateUtils from '../../utils/dateUtils.js';
import refreshUtils from './refreshUtils.js';

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
        sheet.getLastColumn(),
        columns.zahlungsdatum || 0,
        columns.bankabgleich || 0,
        columns.zahlungsart || 0,
        columns.bezahlt || 0,
        columns.bruttoBetrag || 0,
        columns.zahlungsstatus || 0,
        columns.nettobetrag || 0,
        columns.rechnungsnummer || 0,
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

        // Get the sheet or doctype prefix
        const prefix = refreshUtils.getDocumentTypePrefix(sheetType);

        // For Einnahmen, check both regular and gutschrift matches
        // For other sheets, ONLY use regular matches (no gutschrift)
        let bankzuordnung = null;
        if (sheetType === 'einnahmen') {
            const zuordnungsKey = `${prefix}#${row}`;
            const gutschriftKey = `gutschrift#${row}`;
            bankzuordnung = bankZuordnungen[zuordnungsKey] || bankZuordnungen[gutschriftKey];
        } else {
            const zuordnungsKey = `${prefix}#${row}`;
            bankzuordnung = bankZuordnungen[zuordnungsKey];
        }

        const hatBankzuordnung = bankzuordnung !== undefined;

        // Get the document's reference number for validation
        let docReference = '';
        if (columns.rechnungsnummer) {
            docReference = String(data[i][columns.rechnungsnummer - 1] || '').trim();
        }

        // Skip processing if this is a false match by comparing reference numbers
        if (hatBankzuordnung && docReference &&
            bankzuordnung.originalRef &&
            !refreshUtils.isGoodReferenceMatch(docReference, bankzuordnung.originalRef)) {
            // This is likely a false match - skip this row
            continue;
        }

        // Get gutschrift info from current row - ONLY for Einnahmen
        let isGutschrift = false;
        if (sheetType === 'einnahmen' && columns.nettobetrag) {
            const nettobetrag = numberUtils.parseCurrency(data[i][columns.nettobetrag - 1]);
            isGutschrift = nettobetrag < 0;
        }

        // For gutschrifts with bank matches, we handle them specially
        if (isGutschrift) {
            // Add to gutschrift category for coloring
            rowCategories.gutschriftRows.push(row);

            // Also add bank info for gutschrift
            if (columns.bankabgleich && hatBankzuordnung) {
                bankabgleichUpdates.push({
                    row,
                    value: refreshUtils.getZuordnungsInfo(bankzuordnung),
                });
            }

            // And set payment date for gutschrift if we have bank info
            if (columns.zahlungsdatum && hatBankzuordnung && bankzuordnung.bankDatum) {
                // Only set date if currently empty
                if (!data[i][columns.zahlungsdatum - 1] || data[i][columns.zahlungsdatum - 1] === '') {
                    zahlungsdatumUpdates.push({
                        row,
                        value: refreshUtils.formatDate(bankzuordnung.bankDatum),
                    });
                }
            }

            // Set payment method for gutschrift
            if (columns.zahlungsart && hatBankzuordnung) {
                zahlungsartUpdates.push({
                    row,
                    value: 'Überweisung',
                });
            }

            // Continue to next row - we've handled this gutschrift
            continue;
        }

        // For normal rows, use standard categorization
        const rowData = categorizeRow(data[i], row, columns, sheetType, bankZuordnungen, config);

        // Zu entsprechender Kategorie hinzufügen
        if (rowData.category) {
            rowCategories[rowData.category].push(row);
        }

        // Check if dates differ and we need to add correction note
        let dateCorrection = false;
        if (rowData.bankDatum && columns.zahlungsdatum) {
            const existingDate = data[i][columns.zahlungsdatum - 1];
            if (existingDate && existingDate !== '') {
                try {
                    // Normalize both dates to compare them properly
                    const existingDateObj = dateUtils.parseDate(existingDate);
                    const bankDateObj = dateUtils.parseDate(rowData.bankDatum);

                    if (existingDateObj && bankDateObj) {
                        // Format to YYYY-MM-DD for proper comparison
                        const existingNormalized = Utilities.formatDate(
                            existingDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                        const bankNormalized = Utilities.formatDate(
                            bankDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');

                        // Only mark as corrected if dates are actually different
                        dateCorrection = existingNormalized !== bankNormalized;
                    }
                } catch (e) {
                    console.error('Error comparing dates:', e);
                }
            }
        }

        // Updates für Bank-Abgleich-Info with date correction note if needed
        if (rowData.bankInfo && columns.bankabgleich) {
            let bankInfo = rowData.bankInfo;
            if (dateCorrection) {
                bankInfo += ' (Datum korrigiert)';
            }

            bankabgleichUpdates.push({
                row,
                value: bankInfo,
            });
        }

        // Update payment date if needed
        if (rowData.bankDatum && columns.zahlungsdatum) {
            const existingDate = data[i][columns.zahlungsdatum - 1];

            if (!existingDate || existingDate === '' || dateCorrection) {
                zahlungsdatumUpdates.push({
                    row,
                    value: refreshUtils.formatDate(rowData.bankDatum),
                });
            }
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

    // Enhanced Gutschrift detection - ONLY for Einnahmen
    const isGutschrift = sheetType === 'einnahmen' && nettobetrag < 0;

    // Check if this row has banking matches
    const prefix = refreshUtils.getDocumentTypePrefix(sheetType);
    const zuordnungsKey = `${prefix}#${rowIndex}`;

    // For Einnahmen, check both regular and gutschrift matches
    // For other sheets, ONLY use regular matches
    let bankzuordnung = null;
    if (sheetType === 'einnahmen' && isGutschrift) {
        const gutschriftKey = `gutschrift#${rowIndex}`;
        bankzuordnung = bankZuordnungen[zuordnungsKey] || bankZuordnungen[gutschriftKey];
    } else {
        bankzuordnung = bankZuordnungen[zuordnungsKey];
    }

    const hatBankzuordnung = bankzuordnung !== undefined;

    // Zahlungsstatus berechnen - Betrag immer als absoluten Wert verwenden
    const nettoAbs = Math.abs(nettobetrag);
    const mwst = columns.mwstSatz ?
        numberUtils.parseMwstRate(rowData[columns.mwstSatz - 1], config.tax.defaultMwst) / 100 : 0;
    const bruttoBetrag = nettoAbs * (1 + mwst);
    const isPaid = Math.abs(bezahltBetrag) >= bruttoBetrag * 0.999; // 99.9% bezahlt wegen Rundungsfehlern
    const isPartialPaid = !isPaid && bezahltBetrag > 0;

    // Bankabgleich-Info und Datum aus der Zuordnung
    const bankInfo = hatBankzuordnung ? refreshUtils.getZuordnungsInfo(bankzuordnung) : undefined;
    const bankDatum = hatBankzuordnung ? bankzuordnung.bankDatum : undefined;

    // Bezahlt-Betrag berechnen
    let newBezahltBetrag = null;
    if (hatBankzuordnung) {
        // Wenn Bankzuordnung gefunden wurde und es eine Vollzahlung ist, setze den Bruttobetrag
        if (bankzuordnung.matchInfo && bankzuordnung.matchInfo.includes('Vollständige Zahlung')) {
            newBezahltBetrag = bruttoBetrag;
        }
        // Bei Teilzahlung den aktuellen Stand beibehalten oder erhöhen
        else if (bankzuordnung.matchInfo && bankzuordnung.matchInfo.includes('Teilzahlung')) {
            newBezahltBetrag = Math.max(bezahltBetrag, bruttoBetrag * 0.5); // Mindestens 50%
        }
    }

    // Kategorie bestimmen - handle Gutschriften with priority
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

export default {
    markPaidInvoices,
};
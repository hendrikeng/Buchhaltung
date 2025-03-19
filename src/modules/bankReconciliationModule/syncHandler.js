// src/modules/bankReconciliationModule/syncHandler.js
import numberUtils from '../../utils/numberUtils.js';
import dateUtils from '../../utils/dateUtils.js';

/**
 * Markiert bezahlte Einnahmen, Ausgaben und Eigenbelege farblich
 * und aktualisiert den Zahlungsstatus mit optimierter Batch-Verarbeitung
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} bankZuordnungen - Zuordnungen aus dem Bank-Matching (bereits genehmigt)
 * @param {Object} config - Die Konfiguration
 */
function markPaidInvoices(ss, bankZuordnungen, config) {
    console.log('Marking paid invoices...');

    // Map für sheetTypes und ihre Sheet-Objekte (lazy loading)
    const sheetTypeMap = {
        'einnahmen': 'Einnahmen',
        'ausgaben': 'Ausgaben',
        'eigenbelege': 'Eigenbelege',
        'gesellschafterkonto': 'Gesellschafterkonto',
        'holdingTransfers': 'Holding Transfers',
    };

    // Optimierung: Alle Zuordnungen nach Dokumenttyp gruppieren
    const zuordnungenByType = groupZuordnungenByType(bankZuordnungen);

    // Für jeden Dokumenttyp nur einmal das Sheet laden und verarbeiten
    Object.entries(zuordnungenByType).forEach(([sheetType, relevantZuordnungen]) => {
        if (!relevantZuordnungen || relevantZuordnungen.length === 0) return;

        const sheetName = sheetTypeMap[sheetType];
        if (!sheetName) return;

        const sheet = ss.getSheetByName(sheetName);
        if (!sheet || sheet.getLastRow() <= 1) return;

        // Verarbeite das Sheet mit den relevanten Zuordnungen
        markPaidRowsOptimized(sheet, sheetType, relevantZuordnungen, config);
    });
}

/**
 * Gruppiert Zuordnungen nach Dokumenttyp für effizientere Verarbeitung
 * @param {Object} bankZuordnungen - Die Bank-Zuordnungen
 * @returns {Object} Nach Dokumenttyp gruppierte Zuordnungen
 */
function groupZuordnungenByType(bankZuordnungen) {
    const result = {};

    Object.entries(bankZuordnungen).forEach(([key, zuordnung]) => {
        // Extrahiere den Dokumenttyp aus dem Schlüssel
        const docType = key.split('#')[0];

        // Konvertiere Dokumenttyp in SheetType
        const sheetType = docTypeToSheetType(docType);

        if (!result[sheetType]) {
            result[sheetType] = [];
        }

        result[sheetType].push({
            key: key,
            zuordnung: zuordnung,
        });
    });

    return result;
}

/**
 * Konvertiert einen Dokumenttyp in einen Sheet-Typ
 * @param {string} docType - Der Dokumenttyp
 * @returns {string} Der Sheet-Typ
 */
function docTypeToSheetType(docType) {
    const typeMap = {
        'einnahme': 'einnahmen',
        'gutschrift': 'einnahmen', // Gutschriften sind in Einnahmen
        'ausgabe': 'ausgaben',
        'eigenbeleg': 'eigenbelege',
        'gesellschafterkonto': 'gesellschafterkonto',
        'holdingtransfer': 'holdingTransfers',
    };

    return typeMap[docType] || docType;
}

/**
 * Optimierte Version zum Markieren bezahlter Zeilen
 * @param {Sheet} sheet - Das Sheet
 * @param {string} sheetType - Der Sheet-Typ
 * @param {Array} relevantZuordnungen - Die relevanten Zuordnungen
 * @param {Object} config - Die Konfiguration
 */
function markPaidRowsOptimized(sheet, sheetType, relevantZuordnungen, config) {
    console.log(`Processing sheet ${sheetType} with ${relevantZuordnungen.length} relevant assignments`);

    // Konfiguration für das Sheet
    const columns = config[sheetType].columns;

    // Hole Werte aus dem Sheet in einem einzigen API-Call
    const numRows = sheet.getLastRow() - 1;
    if (numRows <= 0) return;

    // Bestimme die benötigten Spalten für die Verarbeitung
    const requiredCols = [
        columns.rechnungsnummer,
        columns.nettobetrag,
        columns.zahlungsdatum,
        columns.bankabgleich,
        columns.zahlungsart,
        columns.bezahlt,
        columns.bruttoBetrag,
        columns.zahlungsstatus,
        columns.mwstSatz,
    ].filter(col => col !== undefined);

    const maxCol = Math.max(...requiredCols, sheet.getLastColumn());

    // Lade alle Daten in einem API-Call
    const data = sheet.getRange(2, 1, numRows, maxCol).getValues();

    // Zuordnungen nach Zeilen gruppieren für schnellen Lookup
    const zuordnungenByRow = {};
    relevantZuordnungen.forEach(({key, zuordnung}) => {
        const row = zuordnung.row;
        if (row >= 2 && row <= numRows + 1) { // 1-basierter Index prüfen
            zuordnungenByRow[row] = zuordnung;
        }
    });

    // Vorbereitung für Batch-Updates
    const updates = {
        rowCategories: {
            fullPaidWithBankRows: [],
            fullPaidRows: [],
            partialPaidWithBankRows: [],
            partialPaidRows: [],
            gutschriftRows: [],
            normalRows: [],
        },
        bankabgleichUpdates: [],
        zahlungsdatumUpdates: [],
        zahlungsartUpdates: [],
        bezahltUpdates: [],
    };

    // Alle Zeilen in einem Durchgang verarbeiten
    for (let i = 0; i < data.length; i++) {
        const row = i + 2; // Aktuelle Zeile (1-basiert + Header)

        // Prüfe, ob für diese Zeile eine Zuordnung existiert
        const zuordnung = zuordnungenByRow[row];
        const hatBankzuordnung = !!zuordnung;

        // In Einnahmen müssen wir Gutschriften besonders behandeln
        let isGutschrift = false;
        if (sheetType === 'einnahmen' && columns.nettobetrag) {
            const nettobetrag = numberUtils.parseCurrency(data[i][columns.nettobetrag - 1]);
            isGutschrift = nettobetrag < 0;
        }

        // Bei Gutschriften spezielle Logik anwenden
        if (isGutschrift) {
            // Gutschriften für Farbgebung kennzeichnen
            updates.rowCategories.gutschriftRows.push(row);

            // Bankabgleich-Info für Gutschrift
            if (columns.bankabgleich && hatBankzuordnung) {
                updates.bankabgleichUpdates.push({
                    row,
                    value: getZuordnungsInfo(zuordnung),
                });
            }

            // Zahlungsdatum für Gutschrift von Bank übernehmen
            if (columns.zahlungsdatum && hatBankzuordnung && zuordnung.bankDatum) {
                updates.zahlungsdatumUpdates.push({
                    row,
                    value: formatBankDate(zuordnung.bankDatum),
                });
            }

            // Zahlungsart für Gutschrift setzen
            if (columns.zahlungsart && hatBankzuordnung) {
                updates.zahlungsartUpdates.push({
                    row,
                    value: 'Überweisung',
                });
            }

            // Nächste Zeile verarbeiten, wir haben diese Gutschrift erledigt
            continue;
        }

        // Normale Zeilen kategorisieren
        const rowInfo = categorizeRow(data[i], columns, sheetType, zuordnung);

        // Zu entsprechender Kategorie hinzufügen
        if (rowInfo.category) {
            updates.rowCategories[rowInfo.category].push(row);
        }

        // Bank-Abgleich-Info hinzufügen
        if (rowInfo.bankInfo && columns.bankabgleich) {
            updates.bankabgleichUpdates.push({
                row,
                value: rowInfo.bankInfo,
            });
        }

        // Zahlungsdatum aktualisieren wenn nötig
        if (rowInfo.bankDatum && columns.zahlungsdatum) {
            updates.zahlungsdatumUpdates.push({
                row,
                value: formatBankDate(rowInfo.bankDatum),
            });
        }

        // Zahlungsart setzen wenn Bank-Zuordnung vorhanden
        if (rowInfo.bankFound && columns.zahlungsart) {
            updates.zahlungsartUpdates.push({
                row,
                value: 'Überweisung',
            });
        }

        // Bezahlt-Betrag aktualisieren wenn nötig
        if (rowInfo.bezahltBetrag !== null && columns.bezahlt) {
            // Nur aktualisieren, wenn der neue Betrag größer ist
            const currentBezahlt = numberUtils.parseCurrency(data[i][columns.bezahlt - 1]);
            if (rowInfo.bezahltBetrag > currentBezahlt) {
                updates.bezahltUpdates.push({
                    row,
                    value: rowInfo.bezahltBetrag,
                });
            }
        }
    }

    // Alle Updates in optimierten Batches anwenden
    applyUpdatesOptimized(sheet, updates, columns);
}

/**
 * Wendet alle Updates in optimierten Batches an
 * @param {Sheet} sheet - Das Sheet
 * @param {Object} updates - Die Updates
 * @param {Object} columns - Die Spaltenkonfiguration
 */
function applyUpdatesOptimized(sheet, updates, columns) {
    // 1. Zeilenfarben setzen
    applyColorToRowCategories(sheet, updates.rowCategories);

    // 2. Spaltenbasierte Updates in Batches
    if (updates.bankabgleichUpdates.length > 0 && columns.bankabgleich) {
        applyColumnUpdatesInBatches(sheet, updates.bankabgleichUpdates, columns.bankabgleich);
    }

    if (updates.zahlungsdatumUpdates.length > 0 && columns.zahlungsdatum) {
        applyColumnUpdatesInBatches(sheet, updates.zahlungsdatumUpdates, columns.zahlungsdatum);
    }

    if (updates.zahlungsartUpdates.length > 0 && columns.zahlungsart) {
        applyColumnUpdatesInBatches(sheet, updates.zahlungsartUpdates, columns.zahlungsart);
    }

    if (updates.bezahltUpdates.length > 0 && columns.bezahlt) {
        applyColumnUpdatesInBatches(sheet, updates.bezahltUpdates, columns.bezahlt);
    }
}

/**
 * Wendet Spalten-Updates in optimierten Batches an
 * @param {Sheet} sheet - Das Sheet
 * @param {Array} updates - Die Updates
 * @param {number} columnIndex - Der Spaltenindex
 */
function applyColumnUpdatesInBatches(sheet, updates, columnIndex) {
    // Optimierung: Updates nach Wert gruppieren
    const valueGroups = {};

    updates.forEach(update => {
        const { row, value } = update;
        if (!valueGroups[value]) {
            valueGroups[value] = [];
        }
        valueGroups[value].push(row);
    });

    // Für jeden Wert einen Batch-Update durchführen
    Object.entries(valueGroups).forEach(([value, rows]) => {
        // Zusammenhängende Zeilen identifizieren für Range-Updates
        const ranges = getRangesFromRows(rows);

        // Jede Range mit dem Wert aktualisieren
        ranges.forEach(range => {
            try {
                const [startRow, endRow] = range;
                const numRows = endRow - startRow + 1;

                // Wertarray für Range erstellen
                const values = Array(numRows).fill([value]);

                // Range in einem API-Call aktualisieren
                sheet.getRange(startRow, columnIndex, numRows, 1).setValues(values);
            } catch (e) {
                console.error(`Error updating range ${range[0]}-${range[1]} with value ${value}:`, e);

                // Fallback: Einzelne Zeilen aktualisieren
                try {
                    for (let row = range[0]; row <= range[1]; row++) {
                        sheet.getRange(row, columnIndex).setValue(value);
                    }
                } catch (fallbackError) {
                    console.error('Fallback error updating rows:', fallbackError);
                }
            }
        });
    });
}

/**
 * Identifiziert zusammenhängende Bereiche aus einer sortierten Liste von Zeilen
 * @param {Array} rows - Die Zeilen
 * @returns {Array} Array mit [startRow, endRow] Paaren
 */
function getRangesFromRows(rows) {
    if (!rows || rows.length === 0) return [];

    // Sortieren für Range-Erkennung
    rows.sort((a, b) => a - b);

    const ranges = [];
    let currentRange = [rows[0], rows[0]];

    for (let i = 1; i < rows.length; i++) {
        // Falls nächste Zeile direkt anschließt, Range erweitern
        if (rows[i] === currentRange[1] + 1) {
            currentRange[1] = rows[i];
        } else {
            // Sonst aktuelle Range speichern und neue beginnen
            ranges.push([...currentRange]);
            currentRange = [rows[i], rows[i]];
        }
    }

    // Letzte Range hinzufügen
    ranges.push([...currentRange]);

    return ranges;
}

/**
 * Wendet Farben auf alle Zeilenkategorien an mit optimierter Batch-Verarbeitung
 * @param {Sheet} sheet - Das Sheet
 * @param {Object} rowCategories - Die Zeilenkategorien
 */
function applyColorToRowCategories(sheet, rowCategories) {
    const colorMap = {
        fullPaidWithBankRows: '#C6EFCE',    // Kräftigeres Grün
        fullPaidRows: '#EAF1DD',            // Helles Grün
        partialPaidWithBankRows: '#FFC7AA', // Kräftigeres Orange
        partialPaidRows: '#FCE4D6',         // Helles Orange
        gutschriftRows: '#E6E0FF',          // Helles Lila
        normalRows: null,                   // Keine Farbe / Zurücksetzen
    };

    // Optimiert: Wende Farben in optimierten Batches an
    Object.entries(rowCategories).forEach(([category, rows]) => {
        if (rows.length > 0) {
            applyColorToRowsOptimized(sheet, rows, colorMap[category]);
        }
    });
}

/**
 * Wendet Farbe auf Zeilen an mit optimierter Range-basierter Verarbeitung
 * @param {Sheet} sheet - Das Sheet
 * @param {Array} rows - Die Zeilen
 * @param {string} color - Die Farbe
 */
function applyColorToRowsOptimized(sheet, rows, color) {
    if (!rows || rows.length === 0) return;

    // Keine Formatierung wenn keine Farbe
    if (!color) return;

    try {
        // Zusammenhängende Ranges identifizieren
        const ranges = getRangesFromRows(rows);

        // Jede Range effizient formatieren
        ranges.forEach(([startRow, endRow]) => {
            try {
                const numRows = endRow - startRow + 1;
                sheet.getRange(startRow, 1, numRows, sheet.getLastColumn())
                    .setBackground(color);
            } catch (e) {
                console.error(`Error applying color to range ${startRow}-${endRow}:`, e);

                // Fallback: Für jede Zeile einzeln formatieren
                for (let row = startRow; row <= endRow; row++) {
                    try {
                        sheet.getRange(row, 1, 1, sheet.getLastColumn())
                            .setBackground(color);
                    } catch (rowError) {
                        console.error(`Error formatting row ${row}:`, rowError);
                    }
                }
            }
        });
    } catch (e) {
        console.error('Error in batch formatting:', e);
    }
}

/**
 * Kategorisiert eine Zeile basierend auf den Daten
 * @param {Array} rowData - Die Zeilendaten
 * @param {Object} columns - Die Spaltenkonfiguration
 * @param {string} sheetType - Der Sheet-Typ
 * @param {Object} bankzuordnung - Die Bank-Zuordnung
 * @returns {Object} Die kategorisierte Zeile
 */
function categorizeRow(rowData, columns, sheetType, bankzuordnung) {
    // Extrahiere relevante Daten aus der Zeile
    const nettobetrag = columns.nettobetrag ? numberUtils.parseCurrency(rowData[columns.nettobetrag - 1]) : 0;
    const bezahltBetrag = columns.bezahlt ? numberUtils.parseCurrency(rowData[columns.bezahlt - 1]) : 0;
    const zahlungsDatum = columns.zahlungsdatum ? rowData[columns.zahlungsdatum - 1] : null;
    const hatBankzuordnung = !!bankzuordnung;

    // Zahlungsdatum validieren
    let isValidPaymentDate = false;
    if (zahlungsDatum) {
        const paymentDate = dateUtils.parseDate(zahlungsDatum);
        isValidPaymentDate = paymentDate && paymentDate <= new Date();
    }

    // MwSt-Rate und Bruttobetrag berechnen
    const nettoAbs = Math.abs(nettobetrag);
    const mwst = columns.mwstSatz ?
        numberUtils.parseMwstRate(rowData[columns.mwstSatz - 1]) / 100 : 0;
    const bruttoBetrag = nettoAbs * (1 + mwst);

    // Zahlungsstatus ermitteln
    const isPaid = Math.abs(bezahltBetrag) >= bruttoBetrag * 0.999 && isValidPaymentDate;
    const isPartialPaid = !isPaid && bezahltBetrag > 0;

    // Informationen für Spalten-Updates
    const bankInfo = hatBankzuordnung ? getZuordnungsInfo(bankzuordnung) : undefined;
    const bankDatum = hatBankzuordnung ? bankzuordnung.bankDatum : undefined;

    // Bezahlt-Betrag berechnen wenn Bank-Zuordnung vorhanden
    let newBezahltBetrag = null;
    if (hatBankzuordnung) {
        // Bei Vollzahlung den Bruttobetrag setzen
        if (bankzuordnung.matchInfo && bankzuordnung.matchInfo.includes('Vollständige Zahlung')) {
            newBezahltBetrag = bruttoBetrag;
        }
        // Bei Teilzahlung aktuellen Stand beibehalten oder erhöhen
        else if (bankzuordnung.matchInfo && bankzuordnung.matchInfo.includes('Teilzahlung')) {
            newBezahltBetrag = Math.max(bezahltBetrag, bruttoBetrag * 0.5);
        }
    }

    // Kategorie für Farbgebung bestimmen
    let category = null;

    if (isPaid) {
        category = isValidPaymentDate ?
            (hatBankzuordnung ? 'fullPaidWithBankRows' : 'fullPaidRows') :
            'partialPaidRows';
    } else if (isPartialPaid) {
        category = hatBankzuordnung ? 'partialPaidWithBankRows' : 'partialPaidRows';
    } else {
        category = 'normalRows';
    }

    return {
        category,
        bankInfo,
        bankDatum,
        bankFound: hatBankzuordnung,
        bezahltBetrag: newBezahltBetrag,
    };
}

/**
 * Erstellt einen Informationstext für eine Bank-Zuordnung
 * @param {Object} zuordnung - Die Bank-Zuordnung
 * @returns {string} Der Informationstext
 */
function getZuordnungsInfo(zuordnung) {
    if (!zuordnung) return '';

    let infoText = '✓ Bank: ';

    // Bank-Datum formatieren
    if (zuordnung.bankDatum) {
        infoText += formatBankDate(zuordnung.bankDatum);
    } else {
        infoText += 'Datum unbekannt';
    }

    // Zusatzinformationen hinzufügen
    if (zuordnung.additional && zuordnung.additional.length > 0) {
        infoText += ' + ' + zuordnung.additional.length + ' weitere';
    }

    return infoText;
}

/**
 * Formatiert ein Bank-Datum für die Anzeige
 * @param {Date|string} date - Das zu formatierende Datum
 * @returns {string} Das formatierte Datum
 */
function formatBankDate(date) {
    if (date instanceof Date) {
        return Utilities.formatDate(
            date,
            Session.getScriptTimeZone(),
            'dd.MM.yyyy',
        );
    }
    return String(date);
}

/**
 * Prüft, ob zwei Referenznummern gut genug übereinstimmen
 * @param {string} ref1 - Erste Referenz
 * @param {string} ref2 - Zweite Referenz
 * @returns {boolean} true wenn die Referenzen übereinstimmen
 */
function isGoodReferenceMatch(ref1, ref2) {
    // Leere Werte behandeln
    if (!ref1 || !ref2) return false;

    // Frühe Rückkehr bei exakter Übereinstimmung
    if (ref1 === ref2) return true;

    // Schnelle Prüfung, ob eine Referenz die andere enthält
    if (ref1.includes(ref2) || ref2.includes(ref1)) {
        // Bei kurzen Referenzen (weniger als 5 Zeichen) strengere Prüfung
        const shorter = ref1.length <= ref2.length ? ref1 : ref2;
        if (shorter.length < 5) {
            return ref1.startsWith(ref2) || ref2.startsWith(ref1);
        }
        return true;
    }

    // Keine gute Übereinstimmung
    return false;
}

/**
 * Ermittelt das Präfix für einen Dokumenttyp
 * @param {string} sheetType - Sheet type
 * @returns {string} - Document type prefix
 */
function getDocumentTypePrefix(sheetType) {
    const prefixMap = {
        'einnahmen': 'einnahme',
        'ausgaben': 'ausgabe',
        'eigenbelege': 'eigenbeleg',
        'gesellschafterkonto': 'gesellschafterkonto',
        'holdingTransfers': 'holdingtransfer',
        'gutschrift': 'gutschrift',
    };
    return prefixMap[sheetType] || sheetType;
}

export default {
    markPaidInvoices,
    isGoodReferenceMatch,
    getDocumentTypePrefix,
    getZuordnungsInfo,
};
// src/modules/bankReconciliationModule/approvalHandler.js
import stringUtils from '../../utils/stringUtils.js';
import dateUtils from '../../utils/dateUtils.js';
import numberUtils from '../../utils/numberUtils.js';

/**
 * Handler für den interaktiven Genehmigungsprozess beim Bankabgleich
 * Optimierte Version mit verbesserter Performance
 */
const approvalHandler = {
    /**
     * Führt den interaktiven Genehmigungsprozess für Bankabgleich-Zuordnungen durch
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Object} bankZuordnungen - Die Zuordnungen aus dem Matching-Prozess
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Genehmigte Zuordnungen und Statistiken
     */
    processApprovals(ss, bankZuordnungen, config) {
        const ui = SpreadsheetApp.getUi();
        const approvedMatches = {};
        const stats = {
            total: Object.keys(bankZuordnungen).length,
            approved: 0,
            rejected: 0,
            byType: {},
        };

        // Optimierung: Batch-Laden von Sheet-Daten
        console.log('Initializing approval process, loading data...');
        const sheets = {};
        const sheetData = this._preloadSheetData(ss, bankZuordnungen, config);

        // Optimierung: Gruppiere Zuordnungen nach Typ und ähnlichen Änderungen
        const groupedMatches = this._groupMatchesByType(bankZuordnungen);

        // Fortschrittsanzeige - Gesamtzahl der zu verarbeitenden Matches
        ui.alert(
            'Bankabgleich - Interaktive Genehmigung',
            `Es wurden ${stats.total} potentielle Übereinstimmungen gefunden. Diese werden nun zur Bestätigung vorgelegt.`,
            ui.ButtonSet.OK,
        );

        // Verarbeite Gruppen von Zuordnungen
        Object.entries(groupedMatches).forEach(([type, matches]) => {
            // Zeige nur an, welcher Dokumenttyp verarbeitet wird, nicht für jeden einzelnen
            console.log(`Processing ${matches.length} ${this._getTypeName(type)} matches...`);

            // Weitere Optimierung: Gruppiere ähnliche Änderungen zusammen
            const changeGroups = this._groupSimilarChanges(matches, sheetData, config);

            // Verarbeite jede Änderungsgruppe
            Object.entries(changeGroups).forEach(([changeKey, groupMatches]) => {
                // Bei sehr ähnlichen Änderungen können diese als Gruppe genehmigt werden
                if (groupMatches.length > 3 && this._hasSimpleChanges(groupMatches, sheetData, config)) {
                    const approved = this._processBatchMatches(groupMatches, sheetData, ui);
                    if (approved) {
                        groupMatches.forEach(match => {
                            const key = `${match.typ}#${match.row}`;
                            approvedMatches[key] = match;
                            stats.approved++;
                            stats.byType[type] = (stats.byType[type] || 0) + 1;
                        });
                    } else {
                        stats.rejected += groupMatches.length;
                    }
                } else {
                    // Einzelne Verarbeitung für komplexere Änderungen
                    groupMatches.forEach(match => {
                        const approved = this._processMatch(match, sheetData, config, ui);
                        if (approved) {
                            const key = `${match.typ}#${match.row}`;
                            approvedMatches[key] = match;
                            stats.approved++;
                            stats.byType[type] = (stats.byType[type] || 0) + 1;
                        } else {
                            stats.rejected++;
                        }
                    });
                }
            });
        });

        // Zeige Zusammenfassung der Genehmigungen
        this._showSummary(stats, ui);

        return {
            approvedMatches,
            stats,
        };
    },

    /**
     * Lädt Daten aus allen relevanten Sheets vorab für bessere Performance
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Object} bankZuordnungen - Die Zuordnungen
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Vorgeladene Daten
     */
    _preloadSheetData(ss, bankZuordnungen, config) {
        const sheetData = {
            sheets: {},
            rowData: {},
        };

        // Ermittle benötigte Sheets und Zeilen
        const neededData = {};

        Object.values(bankZuordnungen).forEach(match => {
            const sheetKey = this._getSheetKey(match.typ);
            if (!neededData[sheetKey]) {
                neededData[sheetKey] = new Set();
            }
            neededData[sheetKey].add(match.row);
        });

        // Lade alle benötigten Sheets
        Object.entries(neededData).forEach(([sheetKey, rows]) => {
            const sheetName = this._getSheetName(sheetKey);
            const sheet = ss.getSheetByName(sheetName);

            if (sheet) {
                sheetData.sheets[sheetKey] = sheet;

                // Batch-Laden aller Zeilen für dieses Sheet
                if (rows.size > 0) {
                    const rowsArray = Array.from(rows).sort((a, b) => a - b);

                    // Optimierung: Zusammenhängende Zeilen in einem Batch laden
                    const ranges = this._getConsecutiveRanges(rowsArray);

                    ranges.forEach(range => {
                        const [startRow, endRow] = range;
                        const numRows = endRow - startRow + 1;

                        // Lade Daten in einem API-Call
                        const values = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

                        // Speichere Zeilen im Cache
                        for (let i = 0; i < values.length; i++) {
                            const rowIndex = startRow + i;
                            sheetData.rowData[`${sheetKey}_${rowIndex}`] = values[i];
                        }
                    });
                }
            }
        });

        return sheetData;
    },

    /**
     * Findet zusammenhängende Bereiche in einer Liste von Zeilennummern
     * @param {Array} rows - Sortierte Zeilennummern
     * @returns {Array} - Array mit [startRow, endRow] Paaren
     */
    _getConsecutiveRanges(rows) {
        const ranges = [];
        if (rows.length === 0) return ranges;

        let start = rows[0];
        let end = rows[0];

        for (let i = 1; i < rows.length; i++) {
            if (rows[i] === end + 1) {
                end = rows[i];
            } else {
                ranges.push([start, end]);
                start = end = rows[i];
            }
        }

        ranges.push([start, end]);
        return ranges;
    },

    /**
     * Gruppiert Zuordnungen nach Dokumenttyp
     * @param {Object} bankZuordnungen - Die Zuordnungen aus dem Matching-Prozess
     * @returns {Object} - Nach Typ gruppierte Zuordnungen
     */
    _groupMatchesByType(bankZuordnungen) {
        const grouped = {};

        Object.values(bankZuordnungen).forEach(match => {
            const type = match.typ;
            if (!grouped[type]) {
                grouped[type] = [];
            }
            grouped[type].push(match);
        });

        return grouped;
    },

    /**
     * Gruppiert Matches nach ähnlichen Änderungen für effizientere Genehmigung
     * @param {Array} matches - Array von Matches desselben Typs
     * @param {Object} sheetData - Vorgeladene Sheet-Daten
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Nach Änderungstyp gruppierte Matches
     */
    _groupSimilarChanges(matches, sheetData, config) {
        const groups = {};

        matches.forEach(match => {
            const sheetKey = this._getSheetKey(match.typ);
            const configKey = this._getConfigKey(match.typ);
            const sheetConfig = config[configKey]?.columns || {};

            // Lade Zeilendaten aus dem Cache
            const rowDataKey = `${sheetKey}_${match.row}`;
            const rowData = sheetData.rowData[rowDataKey];

            if (!rowData) return;

            // Ermittle erforderliche Änderungen
            const changes = this._detectChanges(match, rowData, sheetConfig);
            const realChanges = changes.filter(change => change.type !== 'bankMarking');

            // Erstelle einen Schlüssel basierend auf den Änderungstypen
            const changeTypes = realChanges.map(c => c.type).sort().join('_');

            // Selbst wenn keine Änderungen erforderlich sind, müssen wir die Zuordnung verarbeiten
            // für den Bankabgleich-Haken
            const actualKey = changeTypes || 'no_changes';

            if (!groups[actualKey]) {
                groups[actualKey] = [];
            }

            groups[actualKey].push(match);
        });

        return groups;
    },

    /**
     * Prüft, ob eine Gruppe von Matches nur einfache Änderungen enthält
     * @param {Array} matches - Die zu prüfenden Matches
     * @param {Object} sheetData - Vorgeladene Sheet-Daten
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - true wenn nur einfache Änderungen
     */
    _hasSimpleChanges(matches, sheetData, config) {
        // Nur einfache Änderungen wie Zahlungsdatum und Zahlungsart
        const simpleChangeTypes = new Set(['date', 'paymentMethod']);

        return matches.every(match => {
            const sheetKey = this._getSheetKey(match.typ);
            const configKey = this._getConfigKey(match.typ);
            const sheetConfig = config[configKey]?.columns || {};

            // Lade Zeilendaten aus dem Cache
            const rowDataKey = `${sheetKey}_${match.row}`;
            const rowData = sheetData.rowData[rowDataKey];

            if (!rowData) return false;

            const changes = this._detectChanges(match, rowData, sheetConfig);
            const realChanges = changes.filter(change => change.type !== 'bankMarking');
            return realChanges.every(change => simpleChangeTypes.has(change.type));
        });
    },

    /**
     * Verarbeitet eine Gruppe ähnlicher Matches in einem Batch
     * @param {Array} matches - Die zu verarbeitenden Matches
     * @param {Object} sheetData - Vorgeladene Sheet-Daten
     * @param {Object} ui - Die UI für Dialoge
     * @returns {boolean} - true wenn genehmigt, false wenn abgelehnt
     */
    _processBatchMatches(matches, sheetData, ui) {
        const sampleMatch = matches[0];
        const sheetKey = this._getSheetKey(sampleMatch.typ);
        const typeName = this._getTypeName(sampleMatch.typ);

        let message = `${matches.length} ähnliche ${typeName}-Übereinstimmungen gefunden.\n\n`;
        message += 'Diese Gruppe beinhaltet Dokumente mit ähnlichen Änderungen, wie Zahlungsdatum oder Zahlungsmethode aktualisieren.\n\n';
        message += `Beispiel:\n${this._createSimpleMatchDescription(sampleMatch)}\n\n`;
        message += `Möchten Sie alle ${matches.length} Änderungen in dieser Gruppe übernehmen?`;

        const result = ui.alert(
            'Bankabgleich - Batch-Genehmigung',
            message,
            ui.ButtonSet.YES_NO,
        );

        return result === ui.Button.YES;
    },

    /**
     * Verarbeitet eine einzelne Zuordnung und fragt um Genehmigung
     * @param {Object} match - Die zu verarbeitende Zuordnung
     * @param {Object} sheetData - Vorgeladene Sheet-Daten
     * @param {Object} config - Die Konfiguration
     * @param {Object} ui - Die UI für Dialoge
     * @returns {boolean} - true wenn genehmigt, false wenn abgelehnt
     */
    _processMatch(match, sheetData, config, ui) {
        try {
            // Hole Sheet-Typ und Konfiguration
            const sheetKey = this._getSheetKey(match.typ);
            const configKey = this._getConfigKey(match.typ);
            const sheetConfig = config[configKey]?.columns || {};

            // Lade Zeilendaten aus dem Cache
            const rowDataKey = `${sheetKey}_${match.row}`;
            const rowData = sheetData.rowData[rowDataKey];

            if (!rowData) return false;

            // Erstelle Beschreibung der Zuordnung
            const matchDescription = this._createMatchDescription(match, rowData, sheetConfig);

            // Erstelle Änderungsvorschläge für die Genehmigung
            const changes = this._detectChanges(match, rowData, sheetConfig);

            // Filter out bankMarking changes for dialog display but keep them for processing
            // This way we'll set the checkmark automatically without asking
            const realChanges = changes.filter(change => change.type !== 'bankMarking');

            // If the only changes are setting the checkmark, automatically approve
            if (realChanges.length === 0) {
                // No substantive changes required, automatically approve
                return true;
            }

            // Baue Dialog-Nachricht auf
            let message = `${matchDescription}\n\n`;

            if (realChanges.length > 0) {
                message += 'Folgende Änderungen werden vorgeschlagen:\n';

                realChanges.forEach((change, index) => {
                    message += `${index + 1}. ${change.description}\n`;
                });
            }

            message += '\nMöchten Sie diese Änderungen übernehmen?';

            // Zeige Genehmigungsdialog
            const result = ui.alert(
                'Bankabgleich - Änderungen genehmigen',
                message,
                ui.ButtonSet.YES_NO,
            );

            return result === ui.Button.YES;
        } catch (e) {
            console.error('Fehler bei der Verarbeitung einer Zuordnung:', e);
            // Bei Fehler sicherheitshalber nicht genehmigen
            return false;
        }
    },

    /**
     * Erstellt eine Beschreibung der Zuordnung für den Benutzer
     * @param {Object} match - Die Zuordnung
     * @param {Array} rowData - Die Zeilendaten aus dem Sheet
     * @param {Object} sheetConfig - Die Spaltenkonfiguration
     * @returns {string} - Beschreibung der Zuordnung
     */
    _createMatchDescription(match, rowData, sheetConfig) {
        // Bestimme relevante Daten aus der Zeile
        const documentNumber = sheetConfig.rechnungsnummer ?
            rowData[sheetConfig.rechnungsnummer - 1] : match.originalRef;

        let documentDate = '';
        if (sheetConfig.datum) {
            const rawDate = rowData[sheetConfig.datum - 1];
            // Format German date nicely using dateUtils
            documentDate = dateUtils.formatDate(rawDate);
        }

        const category = sheetConfig.kategorie ?
            rowData[sheetConfig.kategorie - 1] : match.category;

        const amount = sheetConfig.nettobetrag ?
            numberUtils.formatCurrency(rowData[sheetConfig.nettobetrag - 1]) : '';

        // Erstelle Beschreibung
        return `${this._getTypeName(match.typ)} #${match.row} (${documentNumber})` +
            `${documentDate ? ` vom ${documentDate}` : ''}` +
            `${category ? `, Kategorie: "${category}"` : ''}` +
            `${amount ? `, Betrag: ${amount}` : ''}`;
    },

    /**
     * Erstellt eine einfache Beschreibung eines Matches für Batch-Verarbeitung
     * @param {Object} match - Das Match
     * @returns {string} - Einfache Beschreibung
     */
    _createSimpleMatchDescription(match) {
        return `${this._getTypeName(match.typ)} #${match.row} (${match.originalRef || 'unbekannte Referenz'})`;
    },

    /**
     * Erkennt Unterschiede zwischen Banking und Sheet
     * @param {Object} match - Die Zuordnung
     * @param {Array} rowData - Die Zeilendaten aus dem Sheet
     * @param {Object} sheetConfig - Die Spaltenkonfiguration
     * @returns {Array} - Array mit Änderungsvorschlägen
     */
    _detectChanges(match, rowData, sheetConfig) {
        const changes = [];

        // 1. Prüfe Datum (wenn in beiden vorhanden)
        if (match.bankDatum && sheetConfig.zahlungsdatum) {
            const sheetDate = rowData[sheetConfig.zahlungsdatum - 1];

            if (stringUtils.isEmpty(sheetDate)) {
                // Zahlungsdatum fehlt im Sheet
                const formattedBankDate = typeof match.bankDatum === 'string' ?
                    match.bankDatum : dateUtils.formatDate(match.bankDatum);

                changes.push({
                    type: 'date',
                    field: 'zahlungsdatum',
                    description: `Zahlungsdatum aus Bank (${formattedBankDate}) übernehmen - kein Zahlungsdatum vorhanden`,
                    bankValue: match.bankDatum,
                    sheetValue: null,
                });
            } else {
                // Formatierte Datumsdarstellungen für Vergleich und Anzeige
                const formattedSheetDate = dateUtils.formatDate(sheetDate);
                const formattedBankDate = typeof match.bankDatum === 'string' ?
                    match.bankDatum : dateUtils.formatDate(match.bankDatum);

                if (formattedSheetDate !== formattedBankDate) {
                    changes.push({
                        type: 'date',
                        field: 'zahlungsdatum',
                        description: `Zahlungsdatum aktualisieren: "${formattedSheetDate}" → "${formattedBankDate}"`,
                        bankValue: match.bankDatum,
                        sheetValue: sheetDate,
                    });
                }
            }
        }

        // 2. Prüfe Kategorie
        if (match.category && sheetConfig.kategorie) {
            const sheetCategory = rowData[sheetConfig.kategorie - 1];

            if (stringUtils.isEmpty(sheetCategory)) {
                // Kategorie fehlt im Sheet
                changes.push({
                    type: 'category',
                    field: 'kategorie',
                    description: `Kategorie aus ${this._getTypeName(match.typ)} ("${match.category}") übernehmen`,
                    bankValue: match.category,
                    sheetValue: null,
                });
            } else if (sheetCategory !== match.category) {
                changes.push({
                    type: 'category',
                    field: 'kategorie',
                    description: `Kategorie aktualisieren: "${sheetCategory}" → "${match.category}"`,
                    bankValue: match.category,
                    sheetValue: sheetCategory,
                });
            }
        }

        // 3. Prüfe Zahlungsart
        if (sheetConfig.zahlungsart) {
            const sheetPaymentMethod = rowData[sheetConfig.zahlungsart - 1];

            if (stringUtils.isEmpty(sheetPaymentMethod)) {
                changes.push({
                    type: 'paymentMethod',
                    field: 'zahlungsart',
                    description: 'Zahlungsart "Überweisung" übernehmen',
                    bankValue: 'Überweisung',
                    sheetValue: null,
                });
            }
        }

        // 4. Prüfe Zahlungsstatus
        if (sheetConfig.zahlungsstatus && sheetConfig.bruttoBetrag && sheetConfig.bezahlt) {
            const bruttoBetrag = numberUtils.parseCurrency(rowData[sheetConfig.bruttoBetrag - 1]);
            const bezahlt = numberUtils.parseCurrency(rowData[sheetConfig.bezahlt - 1]);
            const sheetStatus = rowData[sheetConfig.zahlungsstatus - 1];

            // Prüfe, ob es sich um eine Gutschrift handelt
            const nettobetrag = sheetConfig.nettobetrag ?
                numberUtils.parseCurrency(rowData[sheetConfig.nettobetrag - 1]) : 0;
            const isGutschrift = nettobetrag < 0;

            let targetStatus;

            if (isGutschrift) {
                // Logik für Gutschriften (negative Beträge)
                if (Math.abs(bezahlt) <= 0) {
                    targetStatus = 'Offen';
                } else if (numberUtils.isApproximatelyEqual(Math.abs(bezahlt), Math.abs(bruttoBetrag), 0.01)) {
                    targetStatus = 'Bezahlt';
                } else {
                    targetStatus = 'Teilbezahlt';
                }
            } else {
                // Logik für normale Rechnungen (positive Beträge)
                if (bezahlt <= 0) {
                    targetStatus = 'Offen';
                } else if (bezahlt >= bruttoBetrag * 0.999) {
                    // Wenn Betrag bezahlt ist oder sogar mehr (Überzahlung), dann ist es bezahlt
                    targetStatus = match.typ === 'eigenbeleg' ? 'Erstattet' : 'Bezahlt';
                } else {
                    // Nur wenn tatsächlich weniger bezahlt wurde, ist es teilbezahlt
                    targetStatus = match.typ === 'eigenbeleg' ? 'Teilerstattet' : 'Teilbezahlt';
                }
            }

            if (sheetStatus !== targetStatus) {
                changes.push({
                    type: 'status',
                    field: 'zahlungsstatus',
                    description: `Zahlungsstatus aktualisieren: "${sheetStatus || 'Nicht gesetzt'}" → "${targetStatus}"`,
                    bankValue: targetStatus,
                    sheetValue: sheetStatus,
                });
            }
        }

        // 5. Prüfe Bankabgleich-Haken
        if (sheetConfig.bankabgleich) {
            const bankMarkingValue = rowData[sheetConfig.bankabgleich - 1];

            if (stringUtils.isEmpty(bankMarkingValue)) {
                const bankInfo = `✓ Bank: ${formatBankDate(match.bankDatum || new Date())}`;

                changes.push({
                    type: 'bankMarking',
                    field: 'bankabgleich',
                    description: 'Bankabgleich-Haken setzen',
                    bankValue: bankInfo,
                    sheetValue: null,
                });
            }
        }

        return changes;
    },

    /**
     * Konvertiert einen Dokumenttyp in einen nutzerfreundlichen Namen
     * @param {string} type - Der Dokumenttyp
     * @returns {string} - Nutzerfreundlicher Name
     */
    _getTypeName(type) {
        const typeMap = {
            'einnahme': 'Einnahme',
            'ausgabe': 'Ausgabe',
            'eigenbeleg': 'Eigenbeleg',
            'gesellschafterkonto': 'Gesellschafterkonto',
            'holdingtransfer': 'Holding Transfer',
            'gutschrift': 'Gutschrift',
        };

        return typeMap[type] || type;
    },

    /**
     * Konvertiert einen Dokumenttyp in einen Sheet-Schlüssel
     * @param {string} type - Der Dokumenttyp
     * @returns {string} - Sheet-Schlüssel
     */
    _getSheetKey(type) {
        const keyMap = {
            'einnahme': 'einnahmen',
            'gutschrift': 'einnahmen',
            'ausgabe': 'ausgaben',
            'eigenbeleg': 'eigenbelege',
            'gesellschafterkonto': 'gesellschafterkonto',
            'holdingtransfer': 'holdingTransfers',
        };

        return keyMap[type] || type;
    },

    /**
     * Konvertiert einen Sheet-Schlüssel in einen Sheet-Namen
     * @param {string} sheetKey - Der Sheet-Schlüssel
     * @returns {string} - Sheet-Name
     */
    _getSheetName(sheetKey) {
        const nameMap = {
            'einnahmen': 'Einnahmen',
            'ausgaben': 'Ausgaben',
            'eigenbelege': 'Eigenbelege',
            'gesellschafterkonto': 'Gesellschafterkonto',
            'holdingTransfers': 'Holding Transfers',
        };

        return nameMap[sheetKey] || sheetKey;
    },

    /**
     * Konvertiert einen Dokumenttyp in einen Konfigurations-Schlüssel
     * @param {string} type - Der Dokumenttyp
     * @returns {string} - Konfigurations-Schlüssel
     */
    _getConfigKey(type) {
        return this._getSheetKey(type);
    },

    /**
     * Zeigt eine Zusammenfassung der Genehmigungen an
     * @param {Object} stats - Statistiken über die Genehmigungen
     * @param {Object} ui - Die UI für Dialoge
     */
    _showSummary(stats, ui) {
        const parts = [
            '=== ZUSAMMENFASSUNG DER BANKABGLEICH-GENEHMIGUNG ===',
            '',
            `Insgesamt bearbeitet: ${stats.total}`,
            `Genehmigt: ${stats.approved}`,
            `Abgelehnt: ${stats.rejected}`,
            '',
        ];

        if (Object.keys(stats.byType).length > 0) {
            parts.push('Genehmigte Zuordnungen nach Typ:');
            Object.entries(stats.byType).forEach(([type, count]) => {
                parts.push(`- ${this._getTypeName(type)}: ${count}`);
            });
        }

        ui.alert(
            'Bankabgleich - Genehmigungsprozess abgeschlossen',
            parts.join('\n'),
            ui.ButtonSet.OK,
        );
    },
};

// Helper function for consistent date formatting
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

export default approvalHandler;
// modules/bwaModule/calculator.js - FIXED VERSION
import dateUtils from '../../utils/dateUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import stringUtils from '../../utils/stringUtils.js';

// Cache für häufig verwendete Kategoriezuordnungen
const categoryMappingCache = new Map();

/**
 * Verarbeitet Einnahmen und ordnet sie den BWA-Kategorien zu
 * @param {Array} row - Zeile aus dem Einnahmen-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processRevenue(row, bwaData, config) {
    try {
        const columns = config.einnahmen.columns;

        // IST-BESTEUERUNG: Prüfen ob bezahlt (analog zu UStVA)
        // Zahlungsdatum prüfen (nur abgeschlossene Zahlungen)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Monat des Zahlungsdatums für Ist-Besteuerung verwenden
        const paymentMonth = paymentDate.getMonth() + 1; // 1-12

        // Prüfen ob Zahlungsdatum im relevanten Jahr liegt
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (paymentDate.getFullYear() !== targetYear) return;

        // Für Ist-Besteuerung: Verwende den bezahlten Betrag statt des Nettobetrags
        const nettobetrag = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        const bruttobetrag = numberUtils.parseCurrency(row[columns.bruttoBetrag - 1]);
        const bezahlt = numberUtils.parseCurrency(row[columns.bezahlt - 1]);

        // Wenn nichts bezahlt wurde, gibt es nichts zu verarbeiten
        if (bezahlt <= 0) return;

        // Wenn nur teilweise bezahlt, berechne den korrekten Anteil
        const bezahltAnteil = bezahlt / (bruttobetrag || 1);
        const amount = nettobetrag * Math.min(bezahltAnteil, 1); // Begrenze auf 100%

        if (amount === 0) return;

        // Kategorie extrahieren
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // Ausland-Status prüfen
        const auslandWert = columns.ausland && row[columns.ausland - 1] ?
            row[columns.ausland - 1].toString().trim() : 'Inland';
        const isEuAusland = auslandWert === 'EU-Ausland';
        const isNichtEuAusland = auslandWert === 'Nicht-EU-Ausland';
        const isAusland = isEuAusland || isNichtEuAusland;

        // Nicht-betriebliche Buchungen ignorieren
        const nonBusinessCategories = new Set(['Gewinnvortrag', 'Verlustvortrag', 'Gewinnvortrag/Verlustvortrag']);
        if (nonBusinessCategories.has(category)) return;

        // Für EU-Ausland und Nicht-EU-Ausland immer als steuerfreie Auslandseinnahmen buchen
        // und nicht in andere BWA-Kategorien einfließen lassen (verhindert Doppelzählung)
        if (isAusland) {
            bwaData[paymentMonth].steuerfreieAuslandEinnahmen += amount;
            return; // Wichtig: Early return, keine weitere Verarbeitung
        }

        // Cache-Schlüssel für Kategorie-Zuordnung
        const cacheKey = `rev_cat_${category}_${auslandWert}`;

        // Prüfen, ob Zuordnung im Cache
        if (categoryMappingCache.has(cacheKey)) {
            const target = categoryMappingCache.get(cacheKey);
            bwaData[paymentMonth][target] += amount;
            return;
        }

        // Direkte Zuordnung basierend auf der Kategorie
        const directMappings = {
            'Erlöse aus Lieferungen und Leistungen': 'umsatzerloese',
            'Provisionserlöse': 'provisionserloese',
            'Sonstige betriebliche Erträge': 'sonstigeErtraege',
            'Erträge aus Vermietung/Verpachtung': 'vermietung',
            'Erträge aus Zuschüssen': 'zuschuesse',
            'Erträge aus Währungsgewinnen': 'waehrungsgewinne',
            'Erträge aus Anlagenabgängen': 'anlagenabgaenge',
        };

        // BWA-Mapping aus Konfiguration verwenden
        const categoryConfig = config.einnahmen.categories[category];
        let mapping = categoryConfig?.bwaMapping;

        // Direktes Mapping aus Liste wenn verfügbar
        if (!mapping && directMappings[category]) {
            mapping = directMappings[category];
        }

        // Wenn Mapping vorhanden, anwenden
        if (mapping) {
            // Mapping im Cache speichern und anwenden
            categoryMappingCache.set(cacheKey, mapping);
            bwaData[paymentMonth][mapping] += amount;
            return;
        }

        // Fallback-Kategorisierung nach Steuersatz
        if (numberUtils.parseMwstRate(row[columns.mwstSatz - 1], config.tax.defaultMwst) === 0) {
            categoryMappingCache.set(cacheKey, 'steuerfreieInlandEinnahmen');
            bwaData[paymentMonth].steuerfreieInlandEinnahmen += amount;
        } else {
            categoryMappingCache.set(cacheKey, 'umsatzerloese');
            bwaData[paymentMonth].umsatzerloese += amount;
        }
    } catch (e) {
        console.error('Fehler bei der Verarbeitung einer Einnahme:', e);
    }
}

/**
 * Verarbeitet Ausgaben und ordnet sie den BWA-Kategorien zu
 * @param {Array} row - Zeile aus dem Ausgaben-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processExpense(row, bwaData, config) {
    try {
        const columns = config.ausgaben.columns;

        // IST-BESTEUERUNG: Prüfen ob bezahlt (analog zu UStVA)
        // Zahlungsdatum prüfen (nur abgeschlossene Zahlungen)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Monat des Zahlungsdatums für Ist-Besteuerung verwenden
        const paymentMonth = paymentDate.getMonth() + 1; // 1-12

        // Prüfen ob Zahlungsdatum im relevanten Jahr liegt
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (paymentDate.getFullYear() !== targetYear) return;

        // Für Ist-Besteuerung: Verwende den bezahlten Betrag statt des Nettobetrags
        const nettobetrag = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        const bruttobetrag = numberUtils.parseCurrency(row[columns.bruttoBetrag - 1]);
        const bezahlt = numberUtils.parseCurrency(row[columns.bezahlt - 1]);

        // Wenn nichts bezahlt wurde, gibt es nichts zu verarbeiten
        if (bezahlt <= 0) return;

        // Wenn nur teilweise bezahlt, berechne den korrekten Anteil
        const bezahltAnteil = bezahlt / (bruttobetrag || 1);
        const amount = nettobetrag * Math.min(bezahltAnteil, 1); // Begrenze auf 100%

        if (amount === 0) return;

        // Kategorie extrahieren
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // Ausland-Status prüfen
        const auslandWert = columns.ausland && row[columns.ausland - 1] ?
            row[columns.ausland - 1].toString().trim() : 'Inland';
        const isEuAusland = auslandWert === 'EU-Ausland';
        const isNichtEuAusland = auslandWert === 'Nicht-EU-Ausland';
        const isAusland = isEuAusland || isNichtEuAusland;

        // Nicht-betriebliche Positionen ignorieren
        const nonBusinessCategories = new Set([
            'Privatentnahme', 'Privateinlage', 'Holding Transfers',
            'Gewinnvortrag', 'Verlustvortrag', 'Gewinnvortrag/Verlustvortrag',
        ]);

        if (nonBusinessCategories.has(category)) {
            return;
        }

        // Für EU-Ausland und Nicht-EU-Ausland immer als steuerfreie Auslandsausgaben buchen
        // und nicht in andere BWA-Kategorien einfließen lassen (verhindert Doppelzählung)
        if (isAusland) {
            bwaData[paymentMonth].steuerfreieAuslandAusgaben += amount;
            return; // Wichtig: Early return, keine weitere Verarbeitung
        }

        // Cache-Schlüssel für Kategorie-Zuordnung
        const cacheKey = `exp_cat_${category}_${auslandWert}`;

        // Prüfen, ob Zuordnung im Cache
        if (categoryMappingCache.has(cacheKey)) {
            const target = categoryMappingCache.get(cacheKey);
            bwaData[paymentMonth][target] += amount;
            return;
        }

        // Direkte Zuordnung basierend auf der Kategorie
        const directMappings = {
            'Wareneinsatz': 'wareneinsatz',
            'Bezogene Leistungen': 'fremdleistungen',
            'Roh-, Hilfs- & Betriebsstoffe': 'rohHilfsBetriebsstoffe',
            'Bruttolöhne & Gehälter': 'bruttoLoehne',
            'Soziale Abgaben & Arbeitgeberanteile': 'sozialeAbgaben',
            'Sonstige Personalkosten': 'sonstigePersonalkosten',
            'Marketing & Werbung': 'werbungMarketing',
            'Werbung & Marketing': 'werbungMarketing',
            'Reisekosten': 'reisekosten',
            'Versicherungen': 'versicherungen',
            'Telefon & Internet': 'telefonInternet',
            'Bürokosten': 'buerokosten',
            'Fortbildungskosten': 'fortbildungskosten',
            'Kfz-Kosten': 'kfzKosten',
            'Abschreibungen Maschinen': 'abschreibungenMaschinen',
            'Abschreibungen Büroausstattung': 'abschreibungenBueromaterial',
            'Abschreibungen immaterielle Wirtschaftsgüter': 'abschreibungenImmateriell',
            'Zinsen auf Bankdarlehen': 'zinsenBank',
            'Zinsen auf Gesellschafterdarlehen': 'zinsenGesellschafter',
            'Leasingkosten': 'leasingkosten',
            'Gewerbesteuerrückstellungen': 'gewerbesteuerRueckstellungen',
            'Körperschaftsteuer': 'koerperschaftsteuer',
            'Solidaritätszuschlag': 'solidaritaetszuschlag',
            'Sonstige Steuerrückstellungen': 'steuerrueckstellungen',
            'Sonstige betriebliche Aufwendungen': 'sonstigeAufwendungen',
        };

        // Direktes Mapping aus Liste wenn verfügbar
        if (directMappings[category]) {
            categoryMappingCache.set(cacheKey, directMappings[category]);
            bwaData[paymentMonth][directMappings[category]] += amount;
            return;
        }

        // BWA-Mapping aus Konfiguration verwenden
        const categoryConfig = config.ausgaben.categories[category];
        const mapping = categoryConfig?.bwaMapping;

        if (mapping) {
            // Mapping für diese Kategorie cachen und anwenden
            categoryMappingCache.set(cacheKey, mapping);
            bwaData[paymentMonth][mapping] += amount;
            return;
        }

        // Fallback: Als sonstige Aufwendungen erfassen
        categoryMappingCache.set(cacheKey, 'sonstigeAufwendungen');
        bwaData[paymentMonth].sonstigeAufwendungen += amount;
    } catch (e) {
        console.error('Fehler bei der Verarbeitung einer Ausgabe:', e);
    }
}

/**
 * Verarbeitet Eigenbelege und ordnet sie den BWA-Kategorien zu
 * @param {Array} row - Zeile aus dem Eigenbelege-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processEigen(row, bwaData, config) {
    try {
        const columns = config.eigenbelege.columns;

        // IST-BESTEUERUNG: Prüfen ob erstattet/bezahlt (analog zu UStVA)
        // Erstattungsdatum prüfen (nur abgeschlossene Zahlungen)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Monat des Zahlungsdatums für Ist-Besteuerung verwenden
        const paymentMonth = paymentDate.getMonth() + 1; // 1-12

        // Prüfen ob Zahlungsdatum im relevanten Jahr liegt
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (paymentDate.getFullYear() !== targetYear) return;

        // Für Ist-Besteuerung: Verwende den bezahlten Betrag statt des Nettobetrags
        const nettobetrag = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        const bruttobetrag = numberUtils.parseCurrency(row[columns.bruttoBetrag - 1]);
        const bezahlt = numberUtils.parseCurrency(row[columns.bezahlt - 1]);

        // Wenn nichts bezahlt wurde, gibt es nichts zu verarbeiten
        if (bezahlt <= 0) return;

        // Wenn nur teilweise bezahlt, berechne den korrekten Anteil
        const bezahltAnteil = bezahlt / (bruttobetrag || 1);
        const amount = nettobetrag * Math.min(bezahltAnteil, 1); // Begrenze auf 100%

        if (amount === 0) return;

        // Kategorie extrahieren
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // Ausland-Status prüfen
        const auslandWert = columns.ausland && row[columns.ausland - 1] ?
            row[columns.ausland - 1].toString().trim() : 'Inland';
        const isEuAusland = auslandWert === 'EU-Ausland';
        const isNichtEuAusland = auslandWert === 'Nicht-EU-Ausland';
        const isAusland = isEuAusland || isNichtEuAusland;

        // Cache-Schlüssel für Kategorie-Konfiguration
        const cacheKey = `eig_cat_${category}_${auslandWert}`;

        // Eigenbelege nach steuertyp kategorisieren
        let eigenCfg, taxType;

        if (categoryMappingCache.has(cacheKey)) {
            const cached = categoryMappingCache.get(cacheKey);
            taxType = cached.taxType;
        } else {
            eigenCfg = config.eigenbelege.categories[category] ?? {};
            taxType = eigenCfg.taxType ?? 'steuerpflichtig';
            categoryMappingCache.set(cacheKey, { taxType });
        }

        // Korrektes Handling für Ausland - nur als steuerfrei erfassen, keine Doppelzählung
        if (isAusland) {
            bwaData[paymentMonth].eigenbelegeSteuerfrei += amount;
            return; // Wichtig: Early return
        }

        // Basierend auf Steuertyp zuordnen
        if (taxType === 'steuerfrei') {
            bwaData[paymentMonth].eigenbelegeSteuerfrei += amount;
        } else {
            bwaData[paymentMonth].eigenbelegeSteuerpflichtig += amount;
        }

        // BWA-Mapping verwenden wenn vorhanden (mit Cache)
        const mappingCacheKey = `eig_map_${category}_${auslandWert}`;
        let mapping;

        if (categoryMappingCache.has(mappingCacheKey)) {
            mapping = categoryMappingCache.get(mappingCacheKey);
        } else {
            const categoryConfig = config.eigenbelege.categories[category];
            mapping = categoryConfig?.bwaMapping || 'sonstigeAufwendungen';
            categoryMappingCache.set(mappingCacheKey, mapping);
        }

        // Einmal dem BWA-Mapping zuordnen
        bwaData[paymentMonth][mapping] += amount;
    } catch (e) {
        console.error('Fehler bei der Verarbeitung eines Eigenbelegs:', e);
    }
}

/**
 * Verarbeitet Gesellschafterkonto und ordnet Positionen den BWA-Kategorien zu
 * KORRIGIERT für Ist-Besteuerung
 * @param {Array} row - Zeile aus dem Gesellschafterkonto-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processGesellschafter(row, bwaData, config) {
    try {
        const columns = config.gesellschafterkonto.columns;

        // IST-BESTEUERUNG: Prüfen ob bezahlt
        // Zahlungsdatum prüfen (nur abgeschlossene Zahlungen)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Monat des Zahlungsdatums für Ist-Besteuerung verwenden
        const paymentMonth = paymentDate.getMonth() + 1; // 1-12

        // Prüfen ob Zahlungsdatum im relevanten Jahr liegt
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (paymentDate.getFullYear() !== targetYear) return;

        // Betrag extrahieren
        const amount = numberUtils.parseCurrency(row[columns.betrag - 1]);
        if (amount === 0) return;

        // Kategorie extrahieren
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // BWA-Mapping verwenden (mit Cache)
        const mappingCacheKey = `gesell_map_${category}`;
        let mapping;

        if (categoryMappingCache.has(mappingCacheKey)) {
            mapping = categoryMappingCache.get(mappingCacheKey);
        } else {
            const categoryConfig = config.gesellschafterkonto.categories[category];
            mapping = categoryConfig?.bwaMapping;
            categoryMappingCache.set(mappingCacheKey, mapping || null);
        }

        if (mapping) {
            bwaData[paymentMonth][mapping] += amount;
        }
    } catch (e) {
        console.error('Fehler bei der Verarbeitung einer Gesellschafterkonto-Position:', e);
    }
}

/**
 * Verarbeitet Holding Transfers und ordnet Positionen den BWA-Kategorien zu
 * KORRIGIERT für Ist-Besteuerung
 * @param {Array} row - Zeile aus dem Holding Transfers-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processHolding(row, bwaData, config) {
    try {
        const columns = config.holdingTransfers.columns;

        // IST-BESTEUERUNG: Prüfen ob bezahlt
        // Zahlungsdatum prüfen (nur abgeschlossene Zahlungen)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Monat des Zahlungsdatums für Ist-Besteuerung verwenden
        const paymentMonth = paymentDate.getMonth() + 1; // 1-12

        // Prüfen ob Zahlungsdatum im relevanten Jahr liegt
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (paymentDate.getFullYear() !== targetYear) return;

        // Betrag extrahieren
        const amount = numberUtils.parseCurrency(row[columns.betrag - 1]);
        if (amount === 0) return;

        // Kategorie extrahieren
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // BWA-Mapping verwenden (mit Cache)
        const mappingCacheKey = `holding_map_${category}`;
        let mapping;

        if (categoryMappingCache.has(mappingCacheKey)) {
            mapping = categoryMappingCache.get(mappingCacheKey);
        } else {
            const categoryConfig = config.holdingTransfers.categories[category];
            mapping = categoryConfig?.bwaMapping;
            categoryMappingCache.set(mappingCacheKey, mapping || null);
        }

        if (mapping) {
            bwaData[paymentMonth][mapping] += amount;
        }
    } catch (e) {
        console.error('Fehler bei der Verarbeitung eines Holding Transfers:', e);
    }
}

/**
 * Berechnet Gruppensummen und abgeleitete Werte für alle Monate
 * Optimierte Version mit effizienter Berechnungslogik
 * @param {Object} bwaData - BWA-Datenstruktur mit Rohdaten
 * @param {Object} config - Die Konfiguration
 */
function calculateAggregates(bwaData, config) {
    for (let m = 1; m <= 12; m++) {
        const d = bwaData[m];

        // Optimierung: Berechne Summen in einem Durchgang und verwende
        // Zwischenspeicher für mehrfach verwendete Summen

        // Erlöse
        d.gesamtErloese = numberUtils.round(
            d.umsatzerloese + d.provisionserloese + d.steuerfreieInlandEinnahmen +
            d.steuerfreieAuslandEinnahmen + d.sonstigeErtraege + d.vermietung +
            d.zuschuesse + d.waehrungsgewinne + d.anlagenabgaenge,
            2,
        );

        // Materialkosten
        d.gesamtWareneinsatz = numberUtils.round(
            d.wareneinsatz + d.fremdleistungen + d.rohHilfsBetriebsstoffe,
            2,
        );

        // Betriebsausgaben
        d.gesamtBetriebsausgaben = numberUtils.round(
            d.bruttoLoehne + d.sozialeAbgaben + d.sonstigePersonalkosten +
            d.werbungMarketing + d.reisekosten + d.versicherungen + d.telefonInternet +
            d.buerokosten + d.fortbildungskosten + d.kfzKosten + d.sonstigeAufwendungen,
            2,
        );

        // Abschreibungen & Zinsen
        d.gesamtAbschreibungenZinsen = numberUtils.round(
            d.abschreibungenMaschinen + d.abschreibungenBueromaterial +
            d.abschreibungenImmateriell + d.zinsenBank + d.zinsenGesellschafter +
            d.leasingkosten,
            2,
        );

        // Besondere Posten
        d.gesamtBesonderePosten = numberUtils.round(
            d.eigenkapitalveraenderungen + d.gesellschafterdarlehen + d.ausschuettungen,
            2,
        );

        // Rückstellungen
        d.gesamtRueckstellungenTransfers = numberUtils.round(
            d.steuerrueckstellungen + d.rueckstellungenSonstige,
            2,
        );

        // EBIT (Zwischenspeicher für Steuern)
        d.ebit = numberUtils.round(
            d.gesamtErloese - (d.gesamtWareneinsatz + d.gesamtBetriebsausgaben +
                d.gesamtAbschreibungenZinsen + d.gesamtBesonderePosten),
            2,
        );

        // Steuern berechnen (mit optimierten Zugriff auf Config)
        const taxConfig = config.tax.isHolding ? config.tax.holding : config.tax.operative;

        // Für Holdings gelten spezielle Steuersätze wegen Beteiligungsprivileg
        const steuerfaktor = config.tax.isHolding
            ? taxConfig.gewinnUebertragSteuerpflichtig / 100
            : 1;

        // Steuerberechnung (angepasst für Klarheit und Optimierung)
        d.gewerbesteuer = numberUtils.round(d.ebit * (taxConfig.gewerbesteuer / 10000) * steuerfaktor, 2);
        d.koerperschaftsteuer = numberUtils.round(d.ebit * (taxConfig.koerperschaftsteuer / 100) * steuerfaktor, 2);
        d.solidaritaetszuschlag = numberUtils.round(d.koerperschaftsteuer * (taxConfig.solidaritaetszuschlag / 100), 2);

        // Gesamte Steuerlast
        d.steuerlast = numberUtils.round(
            d.koerperschaftsteuer + d.solidaritaetszuschlag + d.gewerbesteuer,
            2,
        );

        // Gewinn nach Steuern
        d.gewinnNachSteuern = numberUtils.round(d.ebit - d.steuerlast, 2);
    }
}

export default {
    processRevenue,
    processExpense,
    processEigen,
    processGesellschafter,
    processHolding,
    calculateAggregates,
};
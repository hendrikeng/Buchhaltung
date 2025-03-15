// modules/bwaModule/calculator.js
import dateUtils from '../../utils/dateUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import stringUtils from '../../utils/stringUtils.js';

/**
 * Verarbeitet Einnahmen und ordnet sie den BWA-Kategorien zu
 * @param {Array} row - Zeile aus dem Einnahmen-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processRevenue(row, bwaData, config) {
    try {
        const columns = config.einnahmen.columns;

        const m = dateUtils.getMonthFromRow(row, "einnahmen", config);
        if (!m) return;

        const amount = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        if (amount === 0) return;

        const category = row[columns.kategorie - 1]?.toString().trim() || "";

        // Nicht-betriebliche Buchungen ignorieren
        if (["Gewinnvortrag", "Verlustvortrag", "Gewinnvortrag/Verlustvortrag"].includes(category)) return;

        // Direkte Zuordnung basierend auf der Kategorie
        switch (category) {
            case "Sonstige betriebliche Erträge":
                bwaData[m].sonstigeErtraege += amount;
                return;
            case "Erträge aus Vermietung/Verpachtung":
                bwaData[m].vermietung += amount;
                return;
            case "Erträge aus Zuschüssen":
                bwaData[m].zuschuesse += amount;
                return;
            case "Erträge aus Währungsgewinnen":
                bwaData[m].waehrungsgewinne += amount;
                return;
            case "Erträge aus Anlagenabgängen":
                bwaData[m].anlagenabgaenge += amount;
                return;
        }

        // BWA-Mapping aus Konfiguration verwenden
        const mapping = config.einnahmen.bwaMapping[category];
        if (mapping === "umsatzerloese" || mapping === "provisionserloese") {
            bwaData[m][mapping] += amount;
            return;
        }

        // Kategorisierung nach Steuersatz
        if (numberUtils.parseMwstRate(row[columns.mwstSatz - 1], config.tax.defaultMwst) === 0) {
            bwaData[m].steuerfreieInlandEinnahmen += amount;
        } else {
            bwaData[m].umsatzerloese += amount;
        }
    } catch (e) {
        console.error("Fehler bei der Verarbeitung einer Einnahme:", e);
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

        const m = dateUtils.getMonthFromRow(row, "ausgaben", config);
        if (!m) return;

        const amount = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        if (amount === 0) return;

        const category = row[columns.kategorie - 1]?.toString().trim() || "";

        // Nicht-betriebliche Positionen ignorieren
        if (["Privatentnahme", "Privateinlage", "Holding Transfers",
            "Gewinnvortrag", "Verlustvortrag", "Gewinnvortrag/Verlustvortrag"].includes(category)) {
            return;
        }

        // Direkte Zuordnung basierend auf der Kategorie
        switch (category) {
            case "Bruttolöhne & Gehälter":
                bwaData[m].bruttoLoehne += amount;
                return;
            case "Soziale Abgaben & Arbeitgeberanteile":
                bwaData[m].sozialeAbgaben += amount;
                return;
            case "Sonstige Personalkosten":
                bwaData[m].sonstigePersonalkosten += amount;
                return;
            case "Gewerbesteuerrückstellungen":
                bwaData[m].gewerbesteuerRueckstellungen += amount;
                return;
            case "Telefon & Internet":
                bwaData[m].telefonInternet += amount;
                return;
            case "Bürokosten":
                bwaData[m].buerokosten += amount;
                return;
            case "Fortbildungskosten":
                bwaData[m].fortbildungskosten += amount;
                return;
        }

        // BWA-Mapping aus Konfiguration verwenden
        const mapping = config.ausgaben.bwaMapping[category];
        if (!mapping) {
            bwaData[m].sonstigeAufwendungen += amount;
            return;
        }

        // Zuordnung basierend auf dem BWA-Mapping
        switch (mapping) {
            case "wareneinsatz":
                bwaData[m].wareneinsatz += amount;
                break;
            case "fremdleistungen":
                bwaData[m].fremdleistungen += amount;
                break;
            case "rohHilfsBetriebsstoffe":
                bwaData[m].rohHilfsBetriebsstoffe += amount;
                break;
            case "werbungMarketing":
                bwaData[m].werbungMarketing += amount;
                break;
            case "reisekosten":
                bwaData[m].reisekosten += amount;
                break;
            case "versicherungen":
                bwaData[m].versicherungen += amount;
                break;
            case "kfzKosten":
                bwaData[m].kfzKosten += amount;
                break;
            case "abschreibungenMaschinen":
                bwaData[m].abschreibungenMaschinen += amount;
                break;
            case "abschreibungenBueromaterial":
                bwaData[m].abschreibungenBueromaterial += amount;
                break;
            case "abschreibungenImmateriell":
                bwaData[m].abschreibungenImmateriell += amount;
                break;
            case "zinsenBank":
                bwaData[m].zinsenBank += amount;
                break;
            case "zinsenGesellschafter":
                bwaData[m].zinsenGesellschafter += amount;
                break;
            case "leasingkosten":
                bwaData[m].leasingkosten += amount;
                break;
            case "koerperschaftsteuer":
                bwaData[m].koerperschaftsteuer += amount;
                break;
            case "solidaritaetszuschlag":
                bwaData[m].solidaritaetszuschlag += amount;
                break;
            case "gewerbesteuer":
                bwaData[m].gewerbesteuer += amount;
                break;
            default:
                bwaData[m].sonstigeAufwendungen += amount;
        }
    } catch (e) {
        console.error("Fehler bei der Verarbeitung einer Ausgabe:", e);
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

        const m = dateUtils.getMonthFromRow(row, "eigenbelege", config);
        if (!m) return;

        const amount = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        if (amount === 0) return;

        const category = row[columns.kategorie - 1]?.toString().trim() || "";
        const eigenCfg = config.eigenbelege.categories[category] ?? {};
        const taxType = eigenCfg.taxType ?? "steuerpflichtig";

        // Basierend auf Steuertyp zuordnen
        if (taxType === "steuerfrei") {
            bwaData[m].eigenbelegeSteuerfrei += amount;
        } else {
            bwaData[m].eigenbelegeSteuerpflichtig += amount;
        }

        // BWA-Mapping verwenden wenn vorhanden
        const mapping = config.eigenbelege.bwaMapping[category];
        if (mapping) {
            bwaData[m][mapping] += amount;
        } else {
            bwaData[m].sonstigeAufwendungen += amount;
        }
    } catch (e) {
        console.error("Fehler bei der Verarbeitung eines Eigenbelegs:", e);
    }
}

/**
 * Verarbeitet Gesellschafterkonto und ordnet Positionen den BWA-Kategorien zu
 * @param {Array} row - Zeile aus dem Gesellschafterkonto-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processGesellschafter(row, bwaData, config) {
    try {
        const columns = config.gesellschafterkonto.columns;

        // Datum aus der Zeile extrahieren und Monat ermitteln
        const zeitstempelDatum = row[columns.zeitstempel - 1];
        if (!zeitstempelDatum) return;

        const d = dateUtils.parseDate(zeitstempelDatum);
        if (!d) return;

        // Prüfen, ob das Jahr mit dem Konfigurationsjahr übereinstimmt
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (d.getFullYear() !== targetYear) return;

        const m = d.getMonth() + 1; // Monat (1-12)

        const amount = numberUtils.parseCurrency(row[columns.betrag - 1]);
        if (amount === 0) return;

        const category = row[columns.kategorie - 1]?.toString().trim() || "";

        // BWA-Mapping verwenden
        const mapping = config.gesellschafterkonto.bwaMapping[category];
        if (mapping) {
            bwaData[m][mapping] += amount;
        }
    } catch (e) {
        console.error("Fehler bei der Verarbeitung einer Gesellschafterkonto-Position:", e);
    }
}

/**
 * Verarbeitet Holding Transfers und ordnet Positionen den BWA-Kategorien zu
 * @param {Array} row - Zeile aus dem Holding Transfers-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processHolding(row, bwaData, config) {
    try {
        const columns = config.holdingTransfers.columns;

        // Datum aus der Zeile extrahieren und Monat ermitteln
        const zeitstempelDatum = row[columns.zeitstempel - 1];
        if (!zeitstempelDatum) return;

        const d = dateUtils.parseDate(zeitstempelDatum);
        if (!d) return;

        // Prüfen, ob das Jahr mit dem Konfigurationsjahr übereinstimmt
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (d.getFullYear() !== targetYear) return;

        const m = d.getMonth() + 1; // Monat (1-12)

        const amount = numberUtils.parseCurrency(row[columns.betrag - 1]);
        if (amount === 0) return;

        const category = row[columns.art - 1]?.toString().trim() || "";

        // BWA-Mapping verwenden
        const mapping = config.holdingTransfers.bwaMapping[category];
        if (mapping) {
            bwaData[m][mapping] += amount;
        }
    } catch (e) {
        console.error("Fehler bei der Verarbeitung eines Holding Transfers:", e);
    }
}

/**
 * Berechnet Gruppensummen und abgeleitete Werte für alle Monate
 * @param {Object} bwaData - BWA-Datenstruktur mit Rohdaten
 * @param {Object} config - Die Konfiguration
 */
function calculateAggregates(bwaData, config) {
    for (let m = 1; m <= 12; m++) {
        const d = bwaData[m];

        // Erlöse
        d.gesamtErloese = numberUtils.round(
            d.umsatzerloese + d.provisionserloese + d.steuerfreieInlandEinnahmen +
            d.steuerfreieAuslandEinnahmen + d.sonstigeErtraege + d.vermietung +
            d.zuschuesse + d.waehrungsgewinne + d.anlagenabgaenge,
            2
        );

        // Materialkosten
        d.gesamtWareneinsatz = numberUtils.round(
            d.wareneinsatz + d.fremdleistungen + d.rohHilfsBetriebsstoffe,
            2
        );

        // Betriebsausgaben
        d.gesamtBetriebsausgaben = numberUtils.round(
            d.bruttoLoehne + d.sozialeAbgaben + d.sonstigePersonalkosten +
            d.werbungMarketing + d.reisekosten + d.versicherungen + d.telefonInternet +
            d.buerokosten + d.fortbildungskosten + d.kfzKosten + d.sonstigeAufwendungen,
            2
        );

        // Abschreibungen & Zinsen
        d.gesamtAbschreibungenZinsen = numberUtils.round(
            d.abschreibungenMaschinen + d.abschreibungenBueromaterial +
            d.abschreibungenImmateriell + d.zinsenBank + d.zinsenGesellschafter +
            d.leasingkosten,
            2
        );

        // Besondere Posten
        d.gesamtBesonderePosten = numberUtils.round(
            d.eigenkapitalveraenderungen + d.gesellschafterdarlehen + d.ausschuettungen,
            2
        );

        // Rückstellungen
        d.gesamtRueckstellungenTransfers = numberUtils.round(
            d.steuerrueckstellungen + d.rueckstellungenSonstige,
            2
        );

        // EBIT
        d.ebit = numberUtils.round(
            d.gesamtErloese - (d.gesamtWareneinsatz + d.gesamtBetriebsausgaben +
                d.gesamtAbschreibungenZinsen + d.gesamtBesonderePosten),
            2
        );

        // Steuern berechnen
        const taxConfig = config.tax.isHolding ? config.tax.holding : config.tax.operative;

        // Für Holdings gelten spezielle Steuersätze wegen Beteiligungsprivileg
        const steuerfaktor = config.tax.isHolding
            ? taxConfig.gewinnUebertragSteuerpflichtig / 100
            : 1;

        d.gewerbesteuer = numberUtils.round(d.ebit * (taxConfig.gewerbesteuer / 10000) * steuerfaktor, 2);
        d.koerperschaftsteuer = numberUtils.round(d.ebit * (taxConfig.koerperschaftsteuer / 100) * steuerfaktor, 2);
        d.solidaritaetszuschlag = numberUtils.round(d.koerperschaftsteuer * (taxConfig.solidaritaetszuschlag / 100), 2);

        // Gesamte Steuerlast
        d.steuerlast = numberUtils.round(
            d.koerperschaftsteuer + d.solidaritaetszuschlag + d.gewerbesteuer,
            2
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
    calculateAggregates
};
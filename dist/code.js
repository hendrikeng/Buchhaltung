/* Bundled Code for Google Apps Script */
var global = this;
function onOpen() {};
function onEdit() {};
function setupTrigger() {};
function refreshSheet() {};
function calculateUStVA() {};
function calculateBWA() {};
function calculateBilanz() {};
function importDriveFiles() {};
(function (factory) {
    typeof define === 'function' && define.amd ? define(factory) :
    factory();
})((function () { 'use strict';

    /**
     * Konfiguration für die Buchhaltungsanwendung
     * Unterstützt die Buchhaltung für Holding und operative GmbH nach SKR04
     */
    const config = {
        // Allgemeine Einstellungen
        common: {
            paymentType: ["Überweisung", "Bar", "Kreditkarte", "Paypal", "Lastschrift"],
            months: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"],
            currentYear: new Date().getFullYear(),
            version: "1.0.0"
        },

        // Steuerliche Einstellungen
        tax: {
            defaultMwst: 19,
            allowedMwst: [0, 7, 19],  // Erlaubte MwSt-Sätze
            stammkapital: 25000,
            year: 2021,  // Geschäftsjahr
            isHolding: false, // true bei Holding

            // Holding-spezifische Steuersätze
            holding: {
                gewerbesteuer: 470,  // Angepasst an lokalen Hebesatz
                koerperschaftsteuer: 15,
                solidaritaetszuschlag: 5.5,
                gewinnUebertragSteuerfrei: 95,  // % der Beteiligungserträge steuerfrei
                gewinnUebertragSteuerpflichtig: 5  // % der Beteiligungserträge steuerpflichtig
            },

            // Operative GmbH Steuersätze
            operative: {
                gewerbesteuer: 470,  // Angepasst an lokalen Hebesatz
                koerperschaftsteuer: 15,
                solidaritaetszuschlag: 5.5,
                gewinnUebertragSteuerfrei: 0,
                gewinnUebertragSteuerpflichtig: 100
            }
        },

        // Einnahmen-Konfiguration
        einnahmen: {
            columns: {
                datum: 1,              // A: Rechnungsdatum
                rechnungsnummer: 2,    // B: Rechnungsnummer
                kategorie: 3,          // C: Kategorie
                kunde: 4,              // D: Kunde
                nettobetrag: 5,        // E: Nettobetrag
                mwstSatz: 6,           // F: MwSt-Satz in %
                mwstBetrag: 7,         // G: MwSt-Betrag (E*F)
                bruttoBetrag: 8,       // H: Bruttobetrag (E+G)
                bezahlt: 9,            // I: Bereits bezahlter Betrag
                restbetragNetto: 10,   // J: Restbetrag Netto
                quartal: 11,           // K: Berechnetes Quartal
                zahlungsstatus: 12,    // L: Zahlungsstatus (Offen/Teilbezahlt/Bezahlt)
                zahlungsart: 13,       // M: Zahlungsart
                zahlungsdatum: 14,     // N: Zahlungsdatum
                bankabgleich: 15,      // O: Bankabgleich-Information
                zeitstempel: 16,       // P: Zeitstempel der letzten Änderung
                dateiname: 17,         // Q: Dateiname (für importierte Dateien)
                dateilink: 18          // R: Link zur Originaldatei
            },
            // Kategorien mit Steuertyp
            categories: {
                "Erlöse aus Lieferungen und Leistungen": {taxType: "steuerpflichtig"},
                "Provisionserlöse": {taxType: "steuerpflichtig"},
                "Sonstige betriebliche Erträge": {taxType: "steuerpflichtig"},
                "Erträge aus Vermietung/Verpachtung": {taxType: "steuerfrei_inland"},
                "Erträge aus Zuschüssen": {taxType: "steuerpflichtig"},
                "Erträge aus Währungsgewinnen": {taxType: "steuerpflichtig"},
                "Erträge aus Anlagenabgängen": {taxType: "steuerpflichtig"},
                "Darlehen": {taxType: "steuerfrei_inland"},
                "Zinsen": {taxType: "steuerfrei_inland"},
                "Gewinnvortrag": {taxType: "steuerfrei_inland"},
                "Verlustvortrag": {taxType: "steuerfrei_inland"}
            },
            // SKR04-konforme Konten-Zuordnung
            kontoMapping: {
                "Erlöse aus Lieferungen und Leistungen": {soll: "1200", gegen: "4400", mwst: "1776"},
                "Provisionserlöse": {soll: "1200", gegen: "4120", mwst: "1776"},
                "Sonstige betriebliche Erträge": {soll: "1200", gegen: "4830", mwst: "1776"},
                "Erträge aus Vermietung/Verpachtung": {soll: "1200", gegen: "4180"},
                "Erträge aus Zuschüssen": {soll: "1200", gegen: "4190", mwst: "1776"},
                "Erträge aus Währungsgewinnen": {soll: "1200", gegen: "4840"},
                "Erträge aus Anlagenabgängen": {soll: "1200", gegen: "4855"},
                "Darlehen": {soll: "1200", gegen: "3100"},
                "Zinsen": {soll: "1200", gegen: "4130"},
                "Gewinnvortrag": {soll: "1200", gegen: "2970"},
                "Verlustvortrag": {soll: "1200", gegen: "2978"}
            },
            // BWA-Mapping für Einnahmen-Kategorien
            bwaMapping: {
                "Erlöse aus Lieferungen und Leistungen": "umsatzerloese",
                "Provisionserlöse": "provisionserloese",
                "Sonstige betriebliche Erträge": "sonstigeErtraege",
                "Erträge aus Vermietung/Verpachtung": "vermietung",
                "Erträge aus Zuschüssen": "zuschuesse",
                "Erträge aus Währungsgewinnen": "waehrungsgewinne",
                "Erträge aus Anlagenabgängen": "anlagenabgaenge",
                "Darlehen": "sonstigeErtraege",
                "Zinsen": "sonstigeErtraege"
            }
        },

        // Ausgaben-Konfiguration
        ausgaben: {
            columns: {
                datum: 1,              // A: Rechnungsdatum
                rechnungsnummer: 2,    // B: Rechnungsnummer
                kategorie: 3,          // C: Kategorie
                kunde: 4,              // D: Lieferant
                nettobetrag: 5,        // E: Nettobetrag
                mwstSatz: 6,           // F: MwSt-Satz in %
                mwstBetrag: 7,         // G: MwSt-Betrag (E*F)
                bruttoBetrag: 8,       // H: Bruttobetrag (E+G)
                bezahlt: 9,            // I: Bereits bezahlter Betrag
                restbetragNetto: 10,   // J: Restbetrag Netto
                quartal: 11,           // K: Berechnetes Quartal
                zahlungsstatus: 12,    // L: Zahlungsstatus (Offen/Teilbezahlt/Bezahlt)
                zahlungsart: 13,       // M: Zahlungsart
                zahlungsdatum: 14,     // N: Zahlungsdatum
                bankabgleich: 15,      // O: Bankabgleich-Information
                zeitstempel: 16,       // P: Zeitstempel der letzten Änderung
                dateiname: 17,         // Q: Dateiname (für importierte Dateien)
                dateilink: 18          // R: Link zur Originaldatei
            },
            // Kategorien mit Steuertyp
            categories: {
                // Materialaufwand & Wareneinsatz
                "Wareneinsatz": {taxType: "steuerpflichtig", group: "material"},
                "Bezogene Leistungen": {taxType: "steuerpflichtig", group: "material"},
                "Roh-, Hilfs- & Betriebsstoffe": {taxType: "steuerpflichtig", group: "material"},

                // Personalkosten
                "Bruttolöhne & Gehälter": {taxType: "steuerfrei_inland", group: "personal"},
                "Soziale Abgaben & Arbeitgeberanteile": {taxType: "steuerfrei_inland", group: "personal"},
                "Sonstige Personalkosten": {taxType: "steuerpflichtig", group: "personal"},

                // Raumkosten
                "Miete": {taxType: "steuerfrei_inland", group: "raum"},
                "Nebenkosten": {taxType: "steuerpflichtig", group: "raum"},

                // Betriebskosten
                "Betriebskosten": {taxType: "steuerpflichtig", group: "betrieb"},
                "Marketing & Werbung": {taxType: "steuerpflichtig", group: "betrieb"},
                "Reisekosten": {taxType: "steuerpflichtig", group: "betrieb"},
                "Versicherungen": {taxType: "steuerfrei_inland", group: "betrieb"},
                "Porto": {taxType: "steuerfrei_inland", group: "betrieb"},
                "Google Ads": {taxType: "steuerfrei_ausland", group: "betrieb"},
                "AWS": {taxType: "steuerfrei_ausland", group: "betrieb"},
                "Facebook Ads": {taxType: "steuerfrei_ausland", group: "betrieb"},
                "Bewirtung": {taxType: "steuerpflichtig", group: "betrieb", besonderheit: "bewirtung"},
                "Telefon & Internet": {taxType: "steuerpflichtig", group: "betrieb"},
                "Bürokosten": {taxType: "steuerpflichtig", group: "betrieb"},
                "Fortbildungskosten": {taxType: "steuerpflichtig", group: "betrieb"},

                // Abschreibungen & Zinsen
                "Abschreibungen Maschinen": {taxType: "steuerpflichtig", group: "abschreibung"},
                "Abschreibungen Büroausstattung": {taxType: "steuerpflichtig", group: "abschreibung"},
                "Abschreibungen immaterielle Wirtschaftsgüter": {taxType: "steuerpflichtig", group: "abschreibung"},
                "Zinsen auf Bankdarlehen": {taxType: "steuerpflichtig", group: "zinsen"},
                "Zinsen auf Gesellschafterdarlehen": {taxType: "steuerpflichtig", group: "zinsen"},
                "Leasingkosten": {taxType: "steuerpflichtig", group: "abschreibung"},

                // Steuern & Rückstellungen
                "Gewerbesteuerrückstellungen": {taxType: "steuerfrei_inland", group: "steuer"},
                "Körperschaftsteuer": {taxType: "steuerfrei_inland", group: "steuer"},
                "Solidaritätszuschlag": {taxType: "steuerfrei_inland", group: "steuer"},
                "Sonstige Steuerrückstellungen": {taxType: "steuerfrei_inland", group: "steuer"},

                // Sonstige Aufwendungen
                "Sonstige betriebliche Aufwendungen": {taxType: "steuerpflichtig", group: "sonstige"}
            },
            // SKR04-konforme Konten-Zuordnung
            kontoMapping: {
                "Wareneinsatz": {soll: "5000", gegen: "1200", vorsteuer: "1576"},
                "Bezogene Leistungen": {soll: "5300", gegen: "1200", vorsteuer: "1576"},
                "Roh-, Hilfs- & Betriebsstoffe": {soll: "5400", gegen: "1200", vorsteuer: "1576"},
                "Bruttolöhne & Gehälter": {soll: "6000", gegen: "1200"},
                "Soziale Abgaben & Arbeitgeberanteile": {soll: "6010", gegen: "1200"},
                "Sonstige Personalkosten": {soll: "6020", gegen: "1200", vorsteuer: "1576"},
                "Miete": {soll: "6310", gegen: "1200"},
                "Nebenkosten": {soll: "6320", gegen: "1200", vorsteuer: "1576"},
                "Betriebskosten": {soll: "6300", gegen: "1200", vorsteuer: "1576"},
                "Marketing & Werbung": {soll: "6600", gegen: "1200", vorsteuer: "1576"},
                "Reisekosten": {soll: "6650", gegen: "1200", vorsteuer: "1576"},
                "Versicherungen": {soll: "6400", gegen: "1200"},
                "Porto": {soll: "6810", gegen: "1200"},
                "Google Ads": {soll: "6600", gegen: "1200"},
                "AWS": {soll: "6500", gegen: "1200"},
                "Facebook Ads": {soll: "6600", gegen: "1200"},
                "Bewirtung": {soll: "6670", gegen: "1200", vorsteuer: "1576"},
                "Telefon & Internet": {soll: "6805", gegen: "1200", vorsteuer: "1576"},
                "Bürokosten": {soll: "6815", gegen: "1200", vorsteuer: "1576"},
                "Fortbildungskosten": {soll: "6830", gegen: "1200", vorsteuer: "1576"},
                "Abschreibungen Maschinen": {soll: "6200", gegen: "1200"},
                "Abschreibungen Büroausstattung": {soll: "6210", gegen: "1200"},
                "Abschreibungen immaterielle Wirtschaftsgüter": {soll: "6220", gegen: "1200"},
                "Zinsen auf Bankdarlehen": {soll: "7300", gegen: "1200"},
                "Zinsen auf Gesellschafterdarlehen": {soll: "7310", gegen: "1200"},
                "Leasingkosten": {soll: "6240", gegen: "1200", vorsteuer: "1576"},
                "Gewerbesteuerrückstellungen": {soll: "7610", gegen: "1200"},
                "Körperschaftsteuer": {soll: "7600", gegen: "1200"},
                "Solidaritätszuschlag": {soll: "7620", gegen: "1200"},
                "Sonstige Steuerrückstellungen": {soll: "7630", gegen: "1200"},
                "Sonstige betriebliche Aufwendungen": {soll: "6800", gegen: "1200", vorsteuer: "1576"}
            },
            // BWA-Mapping für Ausgaben-Kategorien
            bwaMapping: {
                "Wareneinsatz": "wareneinsatz",
                "Bezogene Leistungen": "fremdleistungen",
                "Roh-, Hilfs- & Betriebsstoffe": "rohHilfsBetriebsstoffe",
                "Betriebskosten": "sonstigeAufwendungen",
                "Marketing & Werbung": "werbungMarketing",
                "Reisekosten": "reisekosten",
                "Bruttolöhne & Gehälter": "bruttoLoehne",
                "Soziale Abgaben & Arbeitgeberanteile": "sozialeAbgaben",
                "Sonstige Personalkosten": "sonstigePersonalkosten",
                "Sonstige betriebliche Aufwendungen": "sonstigeAufwendungen",
                "Miete": "mieteNebenkosten",
                "Nebenkosten": "mieteNebenkosten",
                "Versicherungen": "versicherungen",
                "Porto": "buerokosten",
                "Telefon & Internet": "telefonInternet",
                "Bürokosten": "buerokosten",
                "Fortbildungskosten": "fortbildungskosten",
                "Abschreibungen Maschinen": "abschreibungenMaschinen",
                "Abschreibungen Büroausstattung": "abschreibungenBueromaterial",
                "Abschreibungen immaterielle Wirtschaftsgüter": "abschreibungenImmateriell",
                "Zinsen auf Bankdarlehen": "zinsenBank",
                "Zinsen auf Gesellschafterdarlehen": "zinsenGesellschafter",
                "Leasingkosten": "leasingkosten",
                "Google Ads": "werbungMarketing",
                "AWS": "sonstigeAufwendungen",
                "Facebook Ads": "werbungMarketing",
                "Bewirtung": "sonstigeAufwendungen",
                "Gewerbesteuerrückstellungen": "gewerbesteuerRueckstellungen",
                "Körperschaftsteuer": "koerperschaftsteuer",
                "Solidaritätszuschlag": "solidaritaetszuschlag",
                "Sonstige Steuerrückstellungen": "steuerrueckstellungen"
            }
        },

        // Eigenbelege-Konfiguration
        eigenbelege: {
            columns: {
                datum: 1,              // A: Belegdatum
                rechnungsnummer: 2,    // B: Belegnummer
                ausgelegtVon: 3,       // C: Ausgelegt von (Person)
                kategorie: 4,          // D: Kategorie
                beschreibung: 5,       // E: Beschreibung
                nettobetrag: 6,        // F: Nettobetrag
                mwstSatz: 7,           // G: MwSt-Satz in %
                mwstBetrag: 8,         // H: MwSt-Betrag (F*G)
                bruttoBetrag: 9,       // I: Bruttobetrag (F+H)
                bezahlt: 10,           // J: Bereits erstatteter Betrag
                restbetragNetto: 11,   // K: Restbetrag Netto
                quartal: 12,           // L: Berechnetes Quartal
                zahlungsstatus: 13,    // M: Erstattungsstatus (Offen/Erstattet/Gebucht)
                zahlungsart: 14,       // N: Erstattungsart
                zahlungsdatum: 15,     // O: Erstattungsdatum
                bankabgleich: 16,      // P: Bankabgleich-Information
                zeitstempel: 17,       // Q: Zeitstempel der letzten Änderung
                dateiname: 18,         // R: Dateiname (für importierte Dateien)
                dateilink: 19          // S: Link zum Originalbeleg
            },
            // Kategorien mit Steuertyp
            categories: {
                "Kleidung": {taxType: "steuerpflichtig"},
                "Trinkgeld": {taxType: "steuerfrei"},
                "Private Vorauslage": {taxType: "steuerfrei"},
                "Bürokosten": {taxType: "steuerpflichtig"},
                "Reisekosten": {taxType: "steuerpflichtig"},
                "Bewirtung": {taxType: "eigenbeleg", besonderheit: "bewirtung"},
                "Sonstiges": {taxType: "steuerpflichtig"}
            },
            // TODO: Necessary?
            kontoMapping: {},
            // TODO: Necessary?
            bwaMapping: {},
        },

        // Bankbewegungen-Konfiguration
        bankbewegungen: {
            columns: {
                datum: 1,              // A: Buchungsdatum
                buchungstext: 2,       // B: Buchungstext
                betrag: 3,             // C: Betrag
                saldo: 4,              // D: Saldo (berechnet)
                transaktionstyp: 5,    // E: Transaktionstyp (Einnahme/Ausgabe)
                kategorie: 6,          // F: Kategorie
                kontoSoll: 7,          // G: Konto (Soll)
                kontoHaben: 8,         // H: Gegenkonto (Haben)
                referenz: 9,           // I: Referenznummer
                verwendungszweck: 10,  // J: Verwendungszweck
                matchInfo: 11,         // K: Match-Information zu Einnahmen/Ausgaben
                zeitstempel: 12,       // L: Zeitstempel
            },
            types: ["Einnahme", "Ausgabe", "Interne Buchung"],
            defaultAccount: "1200"
        },

        // Gesellschafterkonto-Konfiguration
        gesellschafterkonto: {
            columns: {
                datum: 1,              // A: Datum
                beschreibung: 2,       // B: Beschreibung
                kategorie: 3,          // C: Kategorie (Darlehen/Ausschüttung/Kapitalrückführung)
                betrag: 4,             // D: Betrag
                gesellschafter: 5,     // E: Gesellschafter
                anmerkung: 6,          // F: Anmerkung
                kontoSoll: 7,          // G: Konto (Soll)
                kontoHaben: 8,         // H: Gegenkonto (Haben)
                buchungsdatum: 9,      // I: Buchungsdatum
                beleg: 10,             // J: Beleg/Referenz
                saldo: 11,             // K: Saldo (berechnet)
                zeitstempel: 12        // L: Zeitstempel der letzten Änderung
            },
            categories: ["Gesellschafterdarlehen", "Ausschüttungen", "Kapitalrückführung", "Privatentnahme", "Privateinlage"],
            shareholders: ["Christopher Giebel", "Hendrik Werner"]
        },

        // Holding Transfers-Konfiguration
        holdingTransfers: {
            columns: {
                datum: 1,              // A: Datum
                betrag: 2,             // B: Betrag
                art: 3,                // C: Art (Gewinnübertrag/Kapitalrückführung)
                buchungstext: 4,       // D: Buchungstext
                zahlungsstatus: 5,             // E: Status
                referenz: 6,            // F: Referenznummer zur Bankbewegung
                zeitstempel: 7       // G: Zeitstempel der letzten Änderung
            },
            categories: ["Gewinnübertrag", "Kapitalrückführung"]
        },

        // Änderungshistorie-Konfiguration
        aenderungshistorie: {
            columns: {
                datum: 1,              // A: Datum/Zeitstempel
                typ: 2,                // B: Rechnungstyp (Einnahme/Ausgabe)
                dateiname: 3,          // C: Dateiname
                dateilink: 4           // D: Link zur Datei
            }
        },

        // Kontenplan SKR04 (Auszug der wichtigsten Konten)
        kontenplan: {
            // Bestandskonten (Klasse 0-3)
            "0000": "Eröffnungsbilanz",
            "0100": "Immaterielle Vermögensgegenstände",
            "0400": "Grundstücke ohne Bauten",
            "0500": "Bauten auf eigenen Grundstücken",
            "0650": "Technische Anlagen und Maschinen",
            "0700": "Andere Anlagen, Betriebs- und Geschäftsausstattung",
            "1200": "Bank",
            "1210": "Kasse",
            "1300": "Forderungen aus Lieferungen und Leistungen",
            "1576": "Vorsteuer 19%",
            "1577": "Vorsteuer 7%",
            "1590": "Umsatzsteuer-Vorauszahlungen",
            "1600": "Forderungen gegen Gesellschafter",
            "1776": "Umsatzsteuer 19%",
            "1777": "Umsatzsteuer 7%",
            "1790": "Umsatzsteuer-Vorauszahlungen",
            "2000": "Gezeichnetes Kapital",
            "2100": "Kapitalrücklage",
            "2970": "Gewinnvortrag vor Verwendung",
            "2978": "Verlustvortrag vor Verwendung",
            "3100": "Darlehen",
            "3150": "Verbindlichkeiten gegenüber Kreditinstituten",
            "3200": "Verbindlichkeiten aus Lieferungen und Leistungen",
            "3300": "Verbindlichkeiten gegenüber Gesellschaftern",

            // Erlöskonten (Klasse 4)
            "4000": "Umsatzerlöse",
            "4120": "Provisionserlöse",
            "4130": "Zinsertrag",
            "4180": "Erträge aus Vermietung und Verpachtung",
            "4190": "Erträge aus Zuschüssen und Zulagen",
            "4400": "Erlöse aus Lieferungen und Leistungen",
            "4830": "Sonstige betriebliche Erträge",
            "4840": "Erträge aus Währungsumrechnung",
            "4855": "Erträge aus Anlagenabgang",

            // Aufwandskonten (Klasse 5-7)
            "5000": "Wareneinsatz",
            "5300": "Bezogene Leistungen",
            "5400": "Roh-, Hilfs- und Betriebsstoffe",
            "6000": "Löhne und Gehälter",
            "6010": "Gesetzliche soziale Aufwendungen",
            "6020": "Sonstige Personalkosten",
            "6200": "Abschreibungen auf Sachanlagen",
            "6210": "Abschreibungen auf Büroausstattung",
            "6220": "Abschreibungen auf immaterielle Vermögensgegenstände",
            "6240": "Leasingkosten",
            "6300": "Betriebskosten",
            "6310": "Miete",
            "6320": "Nebenkosten",
            "6400": "Versicherungen",
            "6500": "IT-Kosten",
            "6600": "Werbekosten",
            "6650": "Reisekosten",
            "6670": "Bewirtungskosten",
            "6800": "Sonstige betriebliche Aufwendungen",
            "6805": "Telefon und Internet",
            "6810": "Porto",
            "6815": "Bürobedarf",
            "6830": "Fortbildungskosten",
            "7300": "Zinsaufwendungen für Bankdarlehen",
            "7310": "Zinsaufwendungen für Gesellschafterdarlehen",
            "7600": "Körperschaftsteuer",
            "7610": "Gewerbesteuer",
            "7620": "Solidaritätszuschlag",
            "7630": "Sonstige Steuern",

            // Abschlusskonten (Klasse 8-9)
            "8000": "Eröffnungsbilanzkonto",
            "8400": "Abschlussbilanzkonto",
            "8999": "Gewinn- und Verlustkonto"
        },

        /**
         * Initialisierungsfunktion für abgeleitete Daten
         * Wird automatisch beim Import aufgerufen
         */
        initialize() {
            // Bankkategorien dynamisch aus den Einnahmen- und Ausgaben-Kategorien befüllen
            this.bankbewegungen.categories = [
                ...Object.keys(this.einnahmen.categories),
                ...Object.keys(this.ausgaben.categories),
                ...Object.keys(this.eigenbelege.categories),
                ...this.gesellschafterkonto.categories,
                ...this.holdingTransfers.categories,
            ];

            // Duplikate aus den Kategorien entfernen
            this.bankbewegungen.categories = [...new Set(this.bankbewegungen.categories)];

            return this;
        }
    };

    // Initialisierung ausführen und exportieren
    var config$1 = config.initialize();

    // src/helpers.js

    /**
     * Hilfsmodule für verschiedene häufig benötigte Funktionen
     */
    const Helpers = {
        /**
         * Cache für häufig verwendete Berechnungen
         * Verbessert die Performance bei wiederholten Aufrufen
         */
        _cache: {
            dates: new Map(),
            currency: new Map(),
            mwstRates: new Map(),
            columnLetters: new Map()
        },

        /**
         * Cache leeren
         */
        clearCache() {
            this._cache.dates.clear();
            this._cache.currency.clear();
            this._cache.mwstRates.clear();
            this._cache.columnLetters.clear();
        },

        /**
         * Konvertiert verschiedene Datumsformate in ein gültiges Date-Objekt
         * @param {Date|string} value - Das zu parsende Datum
         * @returns {Date|null} - Das geparste Datum oder null bei ungültigem Format
         */
        parseDate(value) {
            // Cache-Lookup für häufig verwendete Werte
            const cacheKey = value instanceof Date
                ? value.getTime().toString()
                : value ? value.toString() : '';

            if (this._cache.dates.has(cacheKey)) {
                return this._cache.dates.get(cacheKey);
            }

            let result = null;

            // Wenn bereits ein Date-Objekt
            if (value instanceof Date) {
                result = isNaN(value.getTime()) ? null : value;
            }
            // Wenn String
            else if (typeof value === "string") {
                // Deutsche Datumsformate (DD.MM.YYYY) unterstützen
                if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(value)) {
                    const [day, month, year] = value.split('.').map(Number);
                    const date = new Date(year, month - 1, day);
                    result = isNaN(date.getTime()) ? null : date;
                } else {
                    const d = new Date(value);
                    result = isNaN(d.getTime()) ? null : d;
                }
            }

            // Ergebnis cachen
            this._cache.dates.set(cacheKey, result);
            return result;
        },

        /**
         * Konvertiert einen String oder eine Zahl in einen numerischen Währungswert
         * @param {number|string} value - Der zu parsende Wert
         * @returns {number} - Der geparste Währungswert oder 0 bei ungültigem Format
         */
        parseCurrency(value) {
            if (value === null || value === undefined || value === "") return 0;
            if (typeof value === "number") return value;

            // Cache-Lookup für String-Werte
            const stringValue = value.toString();
            if (this._cache.currency.has(stringValue)) {
                return this._cache.currency.get(stringValue);
            }

            // Entferne alle Zeichen außer Ziffern, Komma, Punkt und Minus
            const str = stringValue
                .replace(/[^\d,.-]/g, "")
                .replace(/,/g, "."); // Alle Kommas durch Punkte ersetzen

            // Bei mehreren Punkten nur den letzten als Dezimaltrenner behandeln
            const parts = str.split('.');
            let result;

            if (parts.length > 2) {
                const last = parts.pop();
                result = parseFloat(parts.join('') + '.' + last);
            } else {
                result = parseFloat(str);
            }

            result = isNaN(result) ? 0 : result;

            // Ergebnis cachen
            this._cache.currency.set(stringValue, result);
            return result;
        },

        /**
         * Parst einen MwSt-Satz und normalisiert ihn
         * @param {number|string} value - Der zu parsende MwSt-Satz
         * @returns {number} - Der normalisierte MwSt-Satz (0-100)
         */
        parseMwstRate(value) {
            const defaultMwst = config$1?.tax?.defaultMwst || 19;

            if (value === null || value === undefined || value === "") {
                return defaultMwst;
            }

            // Cache-Lookup für häufig verwendete Werte
            const cacheKey = value.toString();
            if (this._cache.mwstRates.has(cacheKey)) {
                return this._cache.mwstRates.get(cacheKey);
            }

            let result;

            if (typeof value === "number") {
                // Wenn der Wert < 1 ist, nehmen wir an, dass es sich um einen Dezimalwert handelt (z.B. 0.19)
                result = value < 1 ? value * 100 : value;
            } else {
                // String-Wert parsen und bereinigen
                const rateStr = value.toString()
                    .replace(/%/g, "")
                    .replace(/,/g, ".")
                    .trim();

                const rate = parseFloat(rateStr);

                // Wenn der geparste Wert ungültig ist, Standardwert zurückgeben
                if (isNaN(rate)) {
                    result = defaultMwst;
                } else {
                    // Normalisieren: Werte < 1 werden als Dezimalwerte interpretiert (z.B. 0.19 -> 19)
                    result = rate < 1 ? rate * 100 : rate;
                }
            }

            // Ergebnis cachen
            this._cache.mwstRates.set(cacheKey, result);
            return result;
        },

        /**
         * Sucht nach einem Ordner mit bestimmtem Namen innerhalb eines übergeordneten Ordners
         * @param {Folder} parent - Der übergeordnete Ordner
         * @param {string} name - Der gesuchte Ordnername
         * @returns {Folder|null} - Der gefundene Ordner oder null
         */
        getFolderByName(parent, name) {
            if (!parent) return null;

            try {
                const folderIter = parent.getFoldersByName(name);
                return folderIter.hasNext() ? folderIter.next() : null;
            } catch (e) {
                console.error("Fehler beim Suchen des Ordners:", e);
                return null;
            }
        },

        /**
         * Extrahiert ein Datum aus einem Dateinamen in verschiedenen Formaten
         * @param {string} filename - Der Dateiname, aus dem das Datum extrahiert werden soll
         * @returns {string} - Das extrahierte Datum im Format DD.MM.YYYY oder leer
         */
        extractDateFromFilename(filename) {
            if (!filename) return "";

            // Cache-Lookup
            if (this._cache.dates.has(`filename_${filename}`)) {
                return this._cache.dates.get(`filename_${filename}`);
            }

            const nameWithoutExtension = filename.replace(/\.[^/.]+$/, "");
            let result = "";

            // Verschiedene Formate erkennen (vom spezifischsten zum allgemeinsten)

            // 1. Format: DD.MM.YYYY im Dateinamen (deutsches Format)
            let match = nameWithoutExtension.match(/(\d{2}[.]\d{2}[.]\d{4})/);
            if (match?.[1]) {
                result = match[1];
            } else {
                // 2. Format: RE-YYYY-MM-DD oder ähnliches mit Trennzeichen
                match = nameWithoutExtension.match(/[^0-9](\d{4}[-_.\/]\d{2}[-_.\/]\d{2})[^0-9]/);
                if (match?.[1]) {
                    const dateParts = match[1].split(/[-_.\/]/);
                    if (dateParts.length === 3) {
                        const [year, month, day] = dateParts;
                        result = `${day}.${month}.${year}`;
                    }
                } else {
                    // 3. Format: YYYY-MM-DD am Anfang oder Ende
                    match = nameWithoutExtension.match(/(^|[^0-9])(\d{4}[-_.\/]\d{2}[-_.\/]\d{2})($|[^0-9])/);
                    if (match?.[2]) {
                        const dateParts = match[2].split(/[-_.\/]/);
                        if (dateParts.length === 3) {
                            const [year, month, day] = dateParts;
                            result = `${day}.${month}.${year}`;
                        }
                    } else {
                        // 4. Format: DD-MM-YYYY mit verschiedenen Trennzeichen
                        match = nameWithoutExtension.match(/(\d{2})[-_.\/](\d{2})[-_.\/](\d{4})/);
                        if (match) {
                            const [_, day, month, year] = match;
                            result = `${day}.${month}.${year}`;
                        }
                    }
                }
            }

            // Ergebnis cachen
            this._cache.dates.set(`filename_${filename}`, result);
            return result;
        },

        /**
         * Setzt bedingte Formatierung für eine Spalte
         * @param {Sheet} sheet - Das zu formatierende Sheet
         * @param {string} column - Die zu formatierende Spalte (z.B. "A")
         * @param {Array<Object>} conditions - Array mit Bedingungen ({value, background, fontColor, pattern})
         */
        setConditionalFormattingForColumn(sheet, column, conditions) {
            if (!sheet || !column || !conditions || !conditions.length) return;

            try {
                const lastRow = Math.max(sheet.getLastRow(), 2);
                const range = sheet.getRange(`${column}2:${column}${lastRow}`);

                // Bestehende Regeln für die Spalte löschen
                const existingRules = sheet.getConditionalFormatRules();
                const newRules = existingRules.filter(rule => {
                    const ranges = rule.getRanges();
                    return !ranges.some(r =>
                        r.getColumn() === range.getColumn() &&
                        r.getRow() === range.getRow() &&
                        r.getNumColumns() === range.getNumColumns()
                    );
                });

                // Neue Regeln erstellen
                const formatRules = conditions.map(({ value, background, fontColor, pattern }) => {
                    let rule;

                    if (pattern === "beginsWith") {
                        rule = SpreadsheetApp.newConditionalFormatRule()
                            .whenTextStartsWith(value);
                    } else if (pattern === "contains") {
                        rule = SpreadsheetApp.newConditionalFormatRule()
                            .whenTextContains(value);
                    } else {
                        rule = SpreadsheetApp.newConditionalFormatRule()
                            .whenTextEqualTo(value);
                    }

                    return rule
                        .setBackground(background || "#ffffff")
                        .setFontColor(fontColor || "#000000")
                        .setRanges([range])
                        .build();
                });

                // Regeln anwenden
                sheet.setConditionalFormatRules([...newRules, ...formatRules]);
            } catch (e) {
                console.error("Fehler beim Setzen der bedingten Formatierung:", e);
            }
        },

        /**
         * Extrahiert den Monat aus einem Datum in einer Zeile
         * @param {Array} row - Die Zeile mit dem Datum
         * @param {string} sheetName - Der Name des Sheets (für Spaltenkonfiguration)
         * @returns {number} - Die Monatsnummer (1-12) oder 0 bei Fehler
         */
        getMonthFromRow(row, sheetName = null) {
            if (!row) return 0;

            // Spalte für Zeitstempel basierend auf Sheet-Typ bestimmen
            let timestampColumn;

            if (sheetName) {
                // Spaltenkonfiguration aus dem Sheetnamen bestimmen
                const sheetConfig = config$1[sheetName.toLowerCase()]?.columns;
                if (sheetConfig && sheetConfig.zeitstempel) {
                    timestampColumn = sheetConfig.zeitstempel - 1; // 0-basiert
                } else {
                    // Fallback auf Standardposition, falls keine Konfiguration gefunden
                    timestampColumn = 15; // Standard: Spalte P (16. Spalte, 0-basiert: 15)
                }
            } else {
                // Wenn kein Sheetname angegeben, Fallback auf Position 13
                // (entspricht dem ursprünglichen Wert im Code)
                timestampColumn = 13;
            }

            // Sicherstellen, dass die Zeile lang genug ist
            if (row.length <= timestampColumn) return 0;

            const d = this.parseDate(row[timestampColumn]);

            // Auf das Jahr aus der Konfiguration prüfen oder das aktuelle Jahr verwenden
            const targetYear = config$1?.tax?.year || new Date().getFullYear();

            // Wenn kein Datum oder das Jahr nicht übereinstimmt
            if (!d || d.getFullYear() !== targetYear) return 0;

            return d.getMonth() + 1; // JavaScript Monate sind 0-basiert, wir geben 1-12 zurück
        },

        /**
         * Formatiert ein Datum im deutschen Format (DD.MM.YYYY)
         * @param {Date|string} date - Das zu formatierende Datum
         * @returns {string} - Das formatierte Datum oder leer bei ungültigem Datum
         */
        formatDate(date) {
            const d = this.parseDate(date);
            if (!d) return "";

            const day = String(d.getDate()).padStart(2, '0');
            const month = String(d.getMonth() + 1).padStart(2, '0');
            const year = d.getFullYear();

            return `${day}.${month}.${year}`;
        },

        /**
         * Formatiert einen Währungsbetrag im deutschen Format
         * @param {number|string} amount - Der zu formatierende Betrag
         * @param {string} currency - Das Währungssymbol (Standard: "€")
         * @returns {string} - Der formatierte Betrag
         */
        formatCurrency(amount, currency = "€") {
            const value = this.parseCurrency(amount);
            return value.toLocaleString('de-DE', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            }) + " " + currency;
        },

        /**
         * Generiert eine eindeutige ID (für Referenzzwecke)
         * @param {string} prefix - Optional ein Präfix für die ID
         * @returns {string} - Eine eindeutige ID
         */
        generateUniqueId(prefix = "") {
            const timestamp = new Date().getTime();
            const random = Math.floor(Math.random() * 10000);
            return `${prefix}${timestamp}${random}`;
        },

        /**
         * Konvertiert einen Spaltenindex (1-basiert) in einen Spaltenbuchstaben (A, B, C, ...)
         * @param {number} columnIndex - 1-basierter Spaltenindex
         * @returns {string} - Spaltenbuchstabe(n)
         */
        getColumnLetter(columnIndex) {
            // Cache-Lookup für häufig verwendete Indizes
            if (this._cache.columnLetters.has(columnIndex)) {
                return this._cache.columnLetters.get(columnIndex);
            }

            let letter = '';
            let colIndex = columnIndex;

            while (colIndex > 0) {
                const modulo = (colIndex - 1) % 26;
                letter = String.fromCharCode(65 + modulo) + letter;
                colIndex = Math.floor((colIndex - modulo) / 26);
            }

            // Ergebnis cachen
            this._cache.columnLetters.set(columnIndex, letter);
            return letter;
        },

        /**
         * Prüft, ob zwei Zahlenwerte im Rahmen einer bestimmten Toleranz gleich sind
         * @param {number} a - Erster Wert
         * @param {number} b - Zweiter Wert
         * @param {number} tolerance - Toleranzwert (Standard: 0.01)
         * @returns {boolean} - true wenn Werte innerhalb der Toleranz gleich sind
         */
        isApproximatelyEqual(a, b, tolerance = 0.01) {
            return Math.abs(a - b) <= tolerance;
        },

        /**
         * Sicheres Runden eines Werts auf n Dezimalstellen
         * @param {number} value - Der zu rundende Wert
         * @param {number} decimals - Anzahl der Dezimalstellen (Standard: 2)
         * @returns {number} - Gerundeter Wert
         */
        round(value, decimals = 2) {
            const factor = Math.pow(10, decimals);
            return Math.round((value + Number.EPSILON) * factor) / factor;
        },

        /**
         * Prüft, ob ein Wert leer oder undefiniert ist
         * @param {*} value - Der zu prüfende Wert
         * @returns {boolean} - true wenn der Wert leer ist
         */
        isEmpty(value) {
            return value === null || value === undefined || value.toString().trim() === "";
        },

        /**
         * Bereinigt einen Text von Sonderzeichen und macht ihn vergleichbar
         * @param {string} text - Der zu bereinigende Text
         * @returns {string} - Der bereinigte Text
         */
        normalizeText(text) {
            if (!text) return "";
            return text.toString()
                .toLowerCase()
                .replace(/[äöüß]/g, match => {
                    return {
                        'ä': 'ae',
                        'ö': 'oe',
                        'ü': 'ue',
                        'ß': 'ss'
                    }[match];
                })
                .replace(/[^a-z0-9]/g, '');
        },

        /**
         * Optimierte Batch-Verarbeitung für Google Sheets API-Calls
         * Vermeidet häufige API-Calls, die zur Drosselung führen können
         * @param {Sheet} sheet - Das zu aktualisierende Sheet
         * @param {Array} data - Array mit Daten-Zeilen
         * @param {number} startRow - Startzeile (1-basiert)
         * @param {number} startCol - Startspalte (1-basiert)
         */
        batchWriteToSheet(sheet, data, startRow, startCol) {
            if (!sheet || !data || !data.length || !data[0].length) return;

            try {
                // Schreibe alle Daten in einem API-Call
                sheet.getRange(
                    startRow,
                    startCol,
                    data.length,
                    data[0].length
                ).setValues(data);
            } catch (e) {
                console.error("Fehler beim Batch-Schreiben in das Sheet:", e);

                // Fallback: Schreibe in kleineren Blöcken, falls der ursprüngliche Call fehlschlägt
                const BATCH_SIZE = 50; // Kleinere Batch-Größe für Fallback

                for (let i = 0; i < data.length; i += BATCH_SIZE) {
                    const batchData = data.slice(i, i + BATCH_SIZE);
                    try {
                        sheet.getRange(
                            startRow + i,
                            startCol,
                            batchData.length,
                            batchData[0].length
                        ).setValues(batchData);

                        // Kurze Pause, um API-Drosselung zu vermeiden
                        if (i + BATCH_SIZE < data.length) {
                            Utilities.sleep(100);
                        }
                    } catch (innerError) {
                        console.error(`Fehler beim Schreiben von Batch ${i / BATCH_SIZE}:`, innerError);
                    }
                }
            }
        }
    };

    // file: src/importModule.js

    /**
     * Modul für den Import von Dateien aus Google Drive in die Buchhaltungstabelle
     */
    const ImportModule = (() => {
        /**
         * Initialisiert die Änderungshistorie, falls sie nicht existiert
         * @param {Sheet} history - Das Änderungshistorie-Sheet
         * @returns {boolean} - true bei erfolgreicher Initialisierung
         */
        const initializeHistorySheet = (history) => {
            try {
                if (history.getLastRow() === 0) {
                    const historyConfig = config$1.aenderungshistorie.columns;
                    const headerRow = Array(history.getLastColumn()).fill("");

                    headerRow[historyConfig.datum - 1] = "Datum";
                    headerRow[historyConfig.typ - 1] = "Rechnungstyp";
                    headerRow[historyConfig.dateiname - 1] = "Dateiname";
                    headerRow[historyConfig.dateilink - 1] = "Link zur Datei";

                    history.appendRow(headerRow);
                    history.getRange(1, 1, 1, 4).setFontWeight("bold");
                }
                return true;
            } catch (e) {
                console.error("Fehler bei der Initialisierung des History-Sheets:", e);
                return false;
            }
        };

        /**
         * Sammelt bereits importierte Dateien aus der Änderungshistorie
         * @param {Sheet} history - Das Änderungshistorie-Sheet
         * @returns {Set} - Set mit bereits importierten Dateinamen
         */
        const collectExistingFiles = (history) => {
            const existingFiles = new Set();
            try {
                const historyData = history.getDataRange().getValues();
                const historyConfig = config$1.aenderungshistorie.columns;

                // Überschriftenzeile überspringen und alle Dateinamen sammeln
                for (let i = 1; i < historyData.length; i++) {
                    const fileName = historyData[i][historyConfig.dateiname - 1];
                    if (fileName) existingFiles.add(fileName);
                }
            } catch (e) {
                console.error("Fehler beim Sammeln bereits importierter Dateien:", e);
            }
            return existingFiles;
        };

        /**
         * Ruft den übergeordneten Ordner des aktuellen Spreadsheets ab
         * @returns {Folder|null} - Der übergeordnete Ordner oder null bei Fehler
         */
        const getParentFolder = () => {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const file = DriveApp.getFileById(ss.getId());
                const parents = file.getParents();
                return parents.hasNext() ? parents.next() : null;
            } catch (e) {
                console.error("Fehler beim Abrufen des übergeordneten Ordners:", e);
                return null;
            }
        };

        /**
         * Findet oder erstellt einen Ordner mit dem angegebenen Namen
         * @param {Folder} parentFolder - Der übergeordnete Ordner
         * @param {string} folderName - Der zu findende oder erstellende Ordnername
         * @returns {Folder|null} - Der gefundene oder erstellte Ordner oder null bei Fehler
         */
        const findOrCreateFolder = (parentFolder, folderName) => {
            if (!parentFolder) return null;

            try {
                let folder = Helpers.getFolderByName(parentFolder, folderName);

                if (!folder) {
                    const ui = SpreadsheetApp.getUi();
                    const createFolder = ui.alert(
                        `Der Ordner '${folderName}' existiert nicht. Soll er erstellt werden?`,
                        ui.ButtonSet.YES_NO
                    );

                    if (createFolder === ui.Button.YES) {
                        folder = parentFolder.createFolder(folderName);
                    }
                }

                return folder;
            } catch (e) {
                console.error(`Fehler beim Finden/Erstellen des Ordners ${folderName}:`, e);
                return null;
            }
        };

        /**
         * Importiert Dateien aus einem Ordner in die entsprechenden Sheets
         *
         * @param {Folder} folder - Google Drive Ordner mit den zu importierenden Dateien
         * @param {Sheet} mainSheet - Hauptsheet (Einnahmen, Ausgaben oder Eigenbelege)
         * @param {string} type - Typ der Dateien ("Einnahme", "Ausgabe" oder "Eigenbeleg")
         * @param {Sheet} historySheet - Sheet für die Änderungshistorie
         * @param {Set} existingFiles - Set mit bereits importierten Dateinamen
         * @returns {number} - Anzahl der importierten Dateien
         */
        const importFilesFromFolder = (folder, mainSheet, type, historySheet, existingFiles) => {
            if (!folder || !mainSheet || !historySheet) return 0;

            const files = folder.getFiles();
            const newMainRows = [];
            const newHistoryRows = [];
            const timestamp = new Date();
            let importedCount = 0;

            // Konfiguration für das richtige Sheet auswählen
            const sheetTypeMap = {
                "Einnahme": config$1.einnahmen.columns,
                "Ausgabe": config$1.ausgaben.columns,
                "Eigenbeleg": config$1.eigenbelege.columns
            };

            const sheetConfig = sheetTypeMap[type];
            if (!sheetConfig) {
                console.error("Unbekannter Dateityp:", type);
                return 0;
            }

            // Konfiguration für das Änderungshistorie-Sheet
            const historyConfig = config$1.aenderungshistorie.columns;

            // Batch-Verarbeitung der Dateien
            const batchSize = 20;
            let fileCount = 0;
            let currentBatch = [];

            while (files.hasNext()) {
                const file = files.next();
                const fileName = file.getName().replace(/\.[^/.]+$/, ""); // Entfernt Dateiendung

                // Prüfe, ob die Datei bereits importiert wurde
                if (!existingFiles.has(fileName)) {
                    // Extraktion der Rechnungsinformationen
                    const invoiceName = fileName.replace(/^[^ ]* /, ""); // Entfernt Präfix vor erstem Leerzeichen
                    const invoiceDate = Helpers.extractDateFromFilename(fileName);
                    const fileUrl = file.getUrl();

                    // Neue Zeile für das Hauptsheet erstellen
                    const row = Array(mainSheet.getLastColumn()).fill("");

                    // Daten in die richtigen Spalten setzen (0-basiert)
                    row[sheetConfig.datum - 1] = invoiceDate;                  // Datum
                    row[sheetConfig.rechnungsnummer - 1] = invoiceName;        // Rechnungsnummer
                    row[sheetConfig.zeitstempel - 1] = timestamp;              // Zeitstempel
                    row[sheetConfig.dateiname - 1] = fileName;                 // Dateiname
                    row[sheetConfig.dateilink - 1] = fileUrl;                  // Link zur Originaldatei

                    newMainRows.push(row);

                    // Protokolliere den Import in der Änderungshistorie
                    const historyRow = Array(historySheet.getLastColumn()).fill("");

                    // Daten in die richtigen Historie-Spalten setzen (0-basiert)
                    historyRow[historyConfig.datum - 1] = timestamp;           // Zeitstempel
                    historyRow[historyConfig.typ - 1] = type;                  // Typ (Einnahme/Ausgabe/Eigenbeleg)
                    historyRow[historyConfig.dateiname - 1] = fileName;        // Dateiname
                    historyRow[historyConfig.dateilink - 1] = fileUrl;         // Link zur Datei

                    newHistoryRows.push(historyRow);
                    existingFiles.add(fileName); // Zur Liste der importierten Dateien hinzufügen
                    importedCount++;

                    // Batch-Verarbeitung zum Verbessern der Performance
                    fileCount++;
                    currentBatch.push(file);

                    // Verarbeitungslimit erreicht oder letzte Datei
                    if (fileCount % batchSize === 0 || !files.hasNext()) {
                        // Hier könnte zusätzliche Batch-Verarbeitung erfolgen
                        // z.B. Metadaten extrahieren, etc.

                        // Batch zurücksetzen
                        currentBatch = [];

                        // Kurze Pause einfügen um API-Limits zu vermeiden
                        Utilities.sleep(50);
                    }
                }
            }

            // Optimierte Schreibvorgänge mit Helpers
            if (newMainRows.length > 0) {
                Helpers.batchWriteToSheet(
                    mainSheet,
                    newMainRows,
                    mainSheet.getLastRow() + 1,
                    1
                );
            }

            if (newHistoryRows.length > 0) {
                Helpers.batchWriteToSheet(
                    historySheet,
                    newHistoryRows,
                    historySheet.getLastRow() + 1,
                    1
                );
            }

            return importedCount;
        };

        /**
         * Hauptfunktion zum Importieren von Dateien aus den Einnahmen- Ausgaben- und Eigenbelegeordnern
         * @returns {number} Anzahl der importierten Dateien
         */
        const importDriveFiles = () => {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();
            let totalImported = 0;

            try {
                // Hauptsheets für Einnahmen, Ausgaben und Eigenbelege abrufen
                const revenueMain = ss.getSheetByName("Einnahmen");
                const expenseMain = ss.getSheetByName("Ausgaben");
                const receiptsMain = ss.getSheetByName("Eigenbelege");

                if (!revenueMain || !expenseMain || !receiptsMain) {
                    ui.alert("Fehler: Die Sheets 'Einnahmen', 'Ausgaben' oder 'Eigenbelege' existieren nicht!");
                    return 0;
                }

                // Änderungshistorie abrufen oder erstellen
                const history = ss.getSheetByName("Änderungshistorie") || ss.insertSheet("Änderungshistorie");

                // Änderungshistorie initialisieren
                if (!initializeHistorySheet(history)) {
                    ui.alert("Fehler: Die Änderungshistorie konnte nicht initialisiert werden!");
                    return 0;
                }

                // Bereits importierte Dateien sammeln
                const existingFiles = collectExistingFiles(history);

                // Auf übergeordneten Ordner zugreifen
                const parentFolder = getParentFolder();
                if (!parentFolder) {
                    ui.alert("Fehler: Kein übergeordneter Ordner gefunden.");
                    return 0;
                }

                // Unterordner für Einnahmen, Ausgaben und Eigenbelege finden oder erstellen
                const revenueFolder = findOrCreateFolder(parentFolder, "Einnahmen");
                const expenseFolder = findOrCreateFolder(parentFolder, "Ausgaben");
                const receiptsFolder = findOrCreateFolder(parentFolder, "Eigenbelege");

                // Import durchführen wenn Ordner existieren
                let importedRevenue = 0, importedExpense = 0, importedReceipts = 0;

                if (revenueFolder) {
                    try {
                        importedRevenue = importFilesFromFolder(
                            revenueFolder,
                            revenueMain,
                            "Einnahme",
                            history,
                            existingFiles
                        );
                        totalImported += importedRevenue;
                    } catch (e) {
                        console.error("Fehler beim Import der Einnahmen:", e);
                        ui.alert("Fehler beim Import der Einnahmen: " + e.toString());
                    }
                }

                if (expenseFolder) {
                    try {
                        importedExpense = importFilesFromFolder(
                            expenseFolder,
                            expenseMain,
                            "Ausgabe",
                            history,
                            existingFiles  // Das gleiche Set wird für beide Importe verwendet
                        );
                        totalImported += importedExpense;
                    } catch (e) {
                        console.error("Fehler beim Import der Ausgaben:", e);
                        ui.alert("Fehler beim Import der Ausgaben: " + e.toString());
                    }
                }

                if (receiptsFolder) {
                    try {
                        importedReceipts = importFilesFromFolder(
                            receiptsFolder,
                            receiptsMain,
                            "Eigenbeleg",
                            history,
                            existingFiles  // Das gleiche Set wird für beide Importe verwendet
                        );
                        totalImported += importedReceipts;
                    } catch (e) {
                        console.error("Fehler beim Import der Eigenbelege:", e);
                        ui.alert("Fehler beim Import der Eigenbelege: " + e.toString());
                    }
                }

                // Abschluss-Meldung anzeigen
                if (totalImported === 0) {
                    ui.alert("Es wurden keine neuen Dateien gefunden.");
                } else {
                    ui.alert(
                        `Import abgeschlossen.\n\n` +
                        `${importedRevenue} Einnahmen importiert.\n` +
                        `${importedExpense} Ausgaben importiert.\n` +
                        `${importedReceipts} Eigenbelege importiert.`
                    );
                }

                return totalImported;
            } catch (e) {
                console.error("Unerwarteter Fehler beim Import:", e);
                ui.alert("Ein unerwarteter Fehler ist aufgetreten: " + e.toString());
                return 0;
            }
        };

        // Öffentliche API des Moduls
        return {
            importDriveFiles
        };
    })();

    // file: src/validator.js

    /**
     * Modul zur Validierung der Eingaben in den verschiedenen Tabellen
     */
    const Validator = (() => {
        /**
         * Erstellt eine Dropdown-Validierung für einen Bereich
         * @param {Sheet} sheet - Das Sheet, in dem validiert werden soll
         * @param {number} row - Die Start-Zeile
         * @param {number} col - Die Start-Spalte
         * @param {number} numRows - Die Anzahl der Zeilen
         * @param {number} numCols - Die Anzahl der Spalten
         * @param {Array<string>} list - Die Liste der gültigen Werte
         * @returns {Range} - Der validierte Bereich
         */
        const validateDropdown = (sheet, row, col, numRows, numCols, list) => {
            if (!sheet || !list || !list.length) return null;

            try {
                return sheet.getRange(row, col, numRows, numCols).setDataValidation(
                    SpreadsheetApp.newDataValidation()
                        .requireValueInList(list, true)
                        .setAllowInvalid(false)
                        .build()
                );
            } catch (e) {
                console.error("Fehler beim Erstellen der Dropdown-Validierung:", e);
                return null;
            }
        };

        /**
         * Validiert eine Zeile aus einem Dokument (Einnahmen, Ausgaben oder Eigenbelege)
         * @param {Array} row - Die zu validierende Zeile
         * @param {number} rowIndex - Der Index der Zeile (für Fehlermeldungen)
         * @param {string} sheetType - Der Typ des Sheets ("einnahmen", "ausgaben" oder "eigenbelege")
         * @returns {Array<string>} - Array mit Warnungen
         */
        const validateDocumentRow = (row, rowIndex, sheetType = "einnahmen") => {
            const warnings = [];
            const columns = config$1[sheetType].columns;

            /**
             * Validiert eine Zeile anhand einer Liste von Validierungsregeln
             * @param {Array} row - Die zu validierende Zeile
             * @param {number} idx - Der Index der Zeile (für Fehlermeldungen)
             * @param {Array<Object>} rules - Array mit Regeln ({check, message})
             */
            const validateRow = (row, idx, rules) => {
                rules.forEach(({check, message}) => {
                    if (check(row)) warnings.push(`Zeile ${idx}: ${message}`);
                });
            };

            // Grundlegende Validierungsregeln für alle Dokumente
            const baseRules = [
                {check: r => Helpers.isEmpty(r[columns.datum - 1]), message: `${sheetType === "eigenbelege" ? "Beleg" : "Rechnungs"}datum fehlt.`},
                {check: r => Helpers.isEmpty(r[columns.rechnungsnummer - 1]), message: `${sheetType === "eigenbelege" ? "Beleg" : "Rechnungs"}nummer fehlt.`},
                {check: r => Helpers.isEmpty(r[columns.kategorie - 1]), message: "Kategorie fehlt."},
                {check: r => isInvalidNumber(r[columns.nettobetrag - 1]), message: "Nettobetrag fehlt oder ungültig."},
                {
                    check: r => {
                        const mwstStr = r[columns.mwstSatz - 1] == null ? "" : r[columns.mwstSatz - 1].toString().trim();
                        if (Helpers.isEmpty(mwstStr)) return false; // Wird schon durch andere Regel geprüft

                        // MwSt-Satz extrahieren und normalisieren
                        const mwst = Helpers.parseMwstRate(mwstStr);
                        if (isNaN(mwst)) return true;

                        // Prüfe auf erlaubte MwSt-Sätze aus der Konfiguration
                        const allowedRates = config$1?.tax?.allowedMwst || [0, 7, 19];
                        return !allowedRates.includes(Math.round(mwst));
                    },
                    message: `Ungültiger MwSt-Satz. Erlaubt sind: ${config$1?.tax?.allowedMwst?.join('%, ')}% oder leer.`
                }
            ];

            // Dokument-spezifische Regeln
            if (sheetType === "einnahmen" || sheetType === "ausgaben") {
                baseRules.push({
                    check: r => Helpers.isEmpty(r[columns.kunde - 1]),
                    message: `${sheetType === "einnahmen" ? "Kunde" : "Lieferant"} fehlt.`
                });
            } else if (sheetType === "eigenbelege") {
                baseRules.push({
                    check: r => Helpers.isEmpty(r[columns.ausgelegtVon - 1]),
                    message: "Ausgelegt von fehlt."
                });
                baseRules.push({
                    check: r => Helpers.isEmpty(r[columns.beschreibung - 1]),
                    message: "Beschreibung fehlt."
                });
            }

            // Status-abhängige Regeln
            const zahlungsstatus = row[columns.zahlungsstatus - 1] ? row[columns.zahlungsstatus - 1].toString().trim().toLowerCase() : "";
            const isOpen = zahlungsstatus === "offen";

            // Angepasste Bezeichnungen je nach Dokumenttyp
            const paidStatus = sheetType === "eigenbelege" ? "erstattet/teilerstattet" : "bezahlt/teilbezahlt";
            const paymentType = sheetType === "eigenbelege" ? "Erstattungsart" : "Zahlungsart";
            const paymentDate = sheetType === "eigenbelege" ? "Erstattungsdatum" : "Zahlungsdatum";

            // Regeln für offene Zahlungen
            const openPaymentRules = [
                {
                    check: r => !Helpers.isEmpty(r[columns.zahlungsart - 1]),
                    message: `${paymentType} darf bei offener Zahlung nicht gesetzt sein.`
                },
                {
                    check: r => !Helpers.isEmpty(r[columns.zahlungsdatum - 1]),
                    message: `${paymentDate} darf bei offener Zahlung nicht gesetzt sein.`
                }
            ];

            // Regeln für bezahlte/erstattete Zahlungen
            const paidPaymentRules = [
                {
                    check: r => Helpers.isEmpty(r[columns.zahlungsart - 1]),
                    message: `${paymentType} muss bei ${paidStatus} Zahlung gesetzt sein.`
                },
                {
                    check: r => Helpers.isEmpty(r[columns.zahlungsdatum - 1]),
                    message: `${paymentDate} muss bei ${paidStatus} Zahlung gesetzt sein.`
                },
                {
                    check: r => {
                        if (Helpers.isEmpty(r[columns.zahlungsdatum - 1])) return false; // Wird schon durch andere Regel geprüft

                        const paymentDate = Helpers.parseDate(r[columns.zahlungsdatum - 1]);
                        return paymentDate ? paymentDate > new Date() : false;
                    },
                    message: `${paymentDate} darf nicht in der Zukunft liegen.`
                },
                {
                    check: r => {
                        if (Helpers.isEmpty(r[columns.zahlungsdatum - 1]) || Helpers.isEmpty(r[columns.datum - 1])) return false;

                        const paymentDate = Helpers.parseDate(r[columns.zahlungsdatum - 1]);
                        const documentDate = Helpers.parseDate(r[columns.datum - 1]);
                        return paymentDate && documentDate ? paymentDate < documentDate : false;
                    },
                    message: `${paymentDate} darf nicht vor dem ${sheetType === "eigenbelege" ? "Beleg" : "Rechnungs"}datum liegen.`
                }
            ];

            // Regeln basierend auf Zahlungsstatus zusammenstellen
            const paymentRules = isOpen ? openPaymentRules : paidPaymentRules;

            // Alle Regeln kombinieren und anwenden
            const rules = [...baseRules, ...paymentRules];
            validateRow(row, rowIndex, rules);

            return warnings;
        };

        /**
         * Prüft, ob ein Wert keine gültige Zahl ist
         * @param {*} v - Der zu prüfende Wert
         * @returns {boolean} - True, wenn der Wert keine gültige Zahl ist
         */
        const isInvalidNumber = v => Helpers.isEmpty(v) || isNaN(parseFloat(v.toString().trim()));

        /**
         * Validiert das Bankbewegungen-Sheet
         * @param {Sheet} bankSheet - Das zu validierende Sheet
         * @returns {Array<string>} - Array mit Warnungen
         */
        const validateBanking = bankSheet => {
            if (!bankSheet) return ["Bankbewegungen-Sheet nicht gefunden"];

            const data = bankSheet.getDataRange().getValues();
            const warnings = [];
            const columns = config$1.bankbewegungen.columns;

            // Regeln für Header- und Footer-Zeilen
            const headerFooterRules = [
                {check: r => Helpers.isEmpty(r[columns.datum - 1]), message: "Buchungsdatum fehlt."},
                {check: r => Helpers.isEmpty(r[columns.buchungstext - 1]), message: "Buchungstext fehlt."},
                {
                    check: r => !Helpers.isEmpty(r[columns.betrag - 1]) && !isNaN(parseFloat(r[columns.betrag - 1].toString().trim())),
                    message: "Betrag darf nicht gesetzt sein."
                },
                {check: r => Helpers.isEmpty(r[columns.saldo - 1]) || isInvalidNumber(r[columns.saldo - 1]), message: "Saldo fehlt oder ungültig."},
                {check: r => !Helpers.isEmpty(r[columns.transaktionstyp - 1]), message: "Typ darf nicht gesetzt sein."},
                {check: r => !Helpers.isEmpty(r[columns.kategorie - 1]), message: "Kategorie darf nicht gesetzt sein."},
                {check: r => !Helpers.isEmpty(r[columns.kontoSoll - 1]), message: "Konto (Soll) darf nicht gesetzt sein."},
                {check: r => !Helpers.isEmpty(r[columns.kontoHaben - 1]), message: "Gegenkonto (Haben) darf nicht gesetzt sein."}
            ];

            // Regeln für Datenzeilen
            const dataRowRules = [
                {check: r => Helpers.isEmpty(r[columns.datum - 1]), message: "Buchungsdatum fehlt."},
                {check: r => Helpers.isEmpty(r[columns.buchungstext - 1]), message: "Buchungstext fehlt."},
                {check: r => Helpers.isEmpty(r[columns.betrag - 1]) || isInvalidNumber(r[columns.betrag - 1]), message: "Betrag fehlt oder ungültig."},
                {check: r => Helpers.isEmpty(r[columns.saldo - 1]) || isInvalidNumber(r[columns.saldo - 1]), message: "Saldo fehlt oder ungültig."},
                {check: r => Helpers.isEmpty(r[columns.transaktionstyp - 1]), message: "Typ fehlt."},
                {check: r => Helpers.isEmpty(r[columns.kategorie - 1]), message: "Kategorie fehlt."},
                {check: r => Helpers.isEmpty(r[columns.kontoSoll - 1]), message: "Konto (Soll) fehlt."},
                {check: r => Helpers.isEmpty(r[columns.kontoHaben - 1]), message: "Gegenkonto (Haben) fehlt."}
            ];

            /**
             * Validiert eine Zeile anhand von Regeln
             * @param {Array} row - Die zu validierende Zeile
             * @param {number} idx - Der Index der Zeile (für Fehlermeldungen)
             * @param {Array<Object>} rules - Array mit Regeln ({check, message})
             */
            const validateRow = (row, idx, rules) => {
                rules.forEach(({check, message}) => {
                    if (check(row)) warnings.push(`Zeile ${idx}: ${message}`);
                });
            };

            // Zeilen validieren
            data.forEach((row, i) => {
                const idx = i + 1;

                // Header oder Footer
                if (i === 0 || i === data.length - 1) {
                    validateRow(row, idx, headerFooterRules);
                }
                // Datenzeilen
                else if (i > 0 && i < data.length - 1) {
                    validateRow(row, idx, dataRowRules);
                }
            });

            return warnings;
        };

        /**
         * Validiert alle Sheets auf Fehler
         * @param {Sheet} revenueSheet - Das Einnahmen-Sheet
         * @param {Sheet} expenseSheet - Das Ausgaben-Sheet
         * @param {Sheet|null} bankSheet - Das Bankbewegungen-Sheet (optional)
         * @param {Sheet|null} eigenSheet - Das Eigenbelege-Sheet (optional)
         * @returns {boolean} - True, wenn keine Fehler gefunden wurden
         */
        const validateAllSheets = (revenueSheet, expenseSheet, bankSheet = null, eigenSheet = null) => {
            if (!revenueSheet || !expenseSheet) {
                SpreadsheetApp.getUi().alert("Fehler: Benötigte Sheets nicht gefunden!");
                return false;
            }

            try {
                // Warnungen für alle Sheets sammeln
                const allWarnings = [];

                // Einnahmen validieren (wenn Daten vorhanden)
                if (revenueSheet.getLastRow() > 1) {
                    const revenueData = revenueSheet.getDataRange().getValues().slice(1); // Header überspringen
                    const revenueWarnings = validateSheet(revenueData, "einnahmen");
                    if (revenueWarnings.length) {
                        allWarnings.push("Fehler in 'Einnahmen':\n" + revenueWarnings.join("\n"));
                    }
                }

                // Ausgaben validieren (wenn Daten vorhanden)
                if (expenseSheet.getLastRow() > 1) {
                    const expenseData = expenseSheet.getDataRange().getValues().slice(1); // Header überspringen
                    const expenseWarnings = validateSheet(expenseData, "ausgaben");
                    if (expenseWarnings.length) {
                        allWarnings.push("Fehler in 'Ausgaben':\n" + expenseWarnings.join("\n"));
                    }
                }

                // Eigenbelege validieren (wenn vorhanden und Daten vorhanden)
                if (eigenSheet && eigenSheet.getLastRow() > 1) {
                    const eigenData = eigenSheet.getDataRange().getValues().slice(1); // Header überspringen
                    const eigenWarnings = validateSheet(eigenData, "eigenbelege");
                    if (eigenWarnings.length) {
                        allWarnings.push("Fehler in 'Eigenbelege':\n" + eigenWarnings.join("\n"));
                    }
                }

                // Bankbewegungen validieren (wenn vorhanden)
                if (bankSheet) {
                    const bankWarnings = validateBanking(bankSheet);
                    if (bankWarnings.length) {
                        allWarnings.push("Fehler in 'Bankbewegungen':\n" + bankWarnings.join("\n"));
                    }
                }

                // Fehlermeldungen anzeigen, falls vorhanden
                if (allWarnings.length) {
                    const ui = SpreadsheetApp.getUi();
                    // Bei vielen Fehlern ggf. einschränken, um UI-Limits zu vermeiden
                    const maxMsgLength = 1500; // Google Sheets Alert-Dialog hat Beschränkungen
                    let alertMsg = allWarnings.join("\n\n");

                    if (alertMsg.length > maxMsgLength) {
                        alertMsg = alertMsg.substring(0, maxMsgLength) +
                            "\n\n... und weitere Fehler. Bitte beheben Sie die angezeigten Fehler zuerst.";
                    }

                    ui.alert("Validierungsfehler gefunden", alertMsg, ui.ButtonSet.OK);
                    return false;
                }

                return true;
            } catch (e) {
                console.error("Fehler bei der Validierung:", e);
                SpreadsheetApp.getUi().alert("Ein Fehler ist bei der Validierung aufgetreten: " + e.toString());
                return false;
            }
        };

        /**
         * Validiert alle Zeilen in einem Sheet
         * @param {Array} data - Zeilen-Daten (ohne Header)
         * @param {string} sheetType - Typ des Sheets ('einnahmen', 'ausgaben' oder 'eigenbelege')
         * @returns {Array<string>} - Array mit Warnungen
         */
        const validateSheet = (data, sheetType) => {
            return data.reduce((warnings, row, index) => {
                // Nur nicht-leere Zeilen prüfen
                if (row.some(cell => cell !== "")) {
                    const rowWarnings = validateDocumentRow(row, index + 2, sheetType);
                    warnings.push(...rowWarnings);
                }
                return warnings;
            }, []);
        };

        /**
         * Validiert eine einzelne Zelle anhand eines festgelegten Formats
         * @param {*} value - Der zu validierende Wert
         * @param {string} type - Der Validierungstyp (date, number, currency, mwst, text)
         * @returns {Object} - Validierungsergebnis {isValid, message}
         */
        const validateCellValue = (value, type) => {
            switch (type.toLowerCase()) {
                case 'date':
                    const date = Helpers.parseDate(value);
                    return {
                        isValid: !!date,
                        message: date ? "" : "Ungültiges Datumsformat. Bitte verwenden Sie DD.MM.YYYY."
                    };

                case 'number':
                    const num = parseFloat(value);
                    return {
                        isValid: !isNaN(num),
                        message: !isNaN(num) ? "" : "Ungültige Zahl."
                    };

                case 'currency':
                    const amount = Helpers.parseCurrency(value);
                    return {
                        isValid: !isNaN(amount),
                        message: !isNaN(amount) ? "" : "Ungültiger Geldbetrag."
                    };

                case 'mwst':
                    const mwst = Helpers.parseMwstRate(value);
                    const allowedRates = config$1?.tax?.allowedMwst || [0, 7, 19];
                    return {
                        isValid: allowedRates.includes(Math.round(mwst)),
                        message: allowedRates.includes(Math.round(mwst))
                            ? ""
                            : `Ungültiger MwSt-Satz. Erlaubt sind: ${allowedRates.join('%, ')}%.`
                    };

                case 'text':
                    return {
                        isValid: !Helpers.isEmpty(value),
                        message: !Helpers.isEmpty(value) ? "" : "Text darf nicht leer sein."
                    };

                default:
                    return {
                        isValid: true,
                        message: ""
                    };
            }
        };

        /**
         * Validiert ein komplettes Sheet und liefert einen detaillierten Fehlerbericht
         * @param {Sheet} sheet - Das zu validierende Sheet
         * @param {Object} validationRules - Regeln für jede Spalte {spaltenIndex: {type, required}}
         * @returns {Object} - Validierungsergebnis mit Fehlern pro Zeile/Spalte
         */
        const validateSheetWithRules = (sheet, validationRules) => {
            const results = {
                isValid: true,
                errors: [],
                errorsByRow: {},
                errorsByColumn: {}
            };

            if (!sheet) {
                results.isValid = false;
                results.errors.push("Sheet nicht gefunden");
                return results;
            }

            const data = sheet.getDataRange().getValues();
            if (data.length <= 1) return results; // Nur Header oder leer

            // Header-Zeile überspringen
            for (let rowIdx = 1; rowIdx < data.length; rowIdx++) {
                const row = data[rowIdx];

                // Jede Spalte mit Validierungsregeln prüfen
                Object.entries(validationRules).forEach(([colIndex, rules]) => {
                    const colIdx = parseInt(colIndex, 10);
                    if (isNaN(colIdx) || colIdx >= row.length) return;

                    const cellValue = row[colIdx];
                    const { type, required } = rules;

                    // Pflichtfeld-Prüfung
                    if (required && Helpers.isEmpty(cellValue)) {
                        const error = {
                            row: rowIdx + 1,
                            column: colIdx + 1,
                            message: "Pflichtfeld nicht ausgefüllt"
                        };
                        addError(results, error);
                        return;
                    }

                    // Wenn nicht leer, auf Typ prüfen
                    if (!Helpers.isEmpty(cellValue) && type) {
                        const validation = validateCellValue(cellValue, type);
                        if (!validation.isValid) {
                            const error = {
                                row: rowIdx + 1,
                                column: colIdx + 1,
                                message: validation.message
                            };
                            addError(results, error);
                        }
                    }
                });
            }

            results.isValid = results.errors.length === 0;
            return results;
        };

        /**
         * Fügt einen Fehler zum Validierungsergebnis hinzu
         * @param {Object} results - Das Validierungsergebnis
         * @param {Object} error - Der Fehler {row, column, message}
         */
        const addError = (results, error) => {
            results.errors.push(error);

            // Nach Zeile gruppieren
            if (!results.errorsByRow[error.row]) {
                results.errorsByRow[error.row] = [];
            }
            results.errorsByRow[error.row].push(error);

            // Nach Spalte gruppieren
            if (!results.errorsByColumn[error.column]) {
                results.errorsByColumn[error.column] = [];
            }
            results.errorsByColumn[error.column].push(error);
        };

        // Öffentliche API des Moduls
        return {
            validateDropdown,
            validateDocumentRow,
            validateBanking,
            validateAllSheets,
            validateCellValue,
            validateSheetWithRules,
            isEmpty: Helpers.isEmpty,
            isInvalidNumber
        };
    })();

    // file: src/refreshModule.js

    /**
     * Modul zum Aktualisieren der Tabellenblätter und Neuberechnen von Formeln
     */
    const RefreshModule = (() => {
        // Cache für wiederholte Berechnungen
        const _cache = {
            // Sheet-Referenzen Cache, um unnötige getSheetByName-Aufrufe zu vermeiden
            sheets: new Map(),
            // Referenz-Maps für schnellere Suche nach Rechnungsnummern
            references: {
                einnahmen: null,
                ausgaben: null
            }
        };

        /**
         * Cache zurücksetzen
         */
        const clearCache = () => {
            _cache.sheets.clear();
            _cache.references.einnahmen = null;
            _cache.references.ausgaben = null;
        };

        /**
         * Holt ein Sheet aus dem Cache oder vom Spreadsheet
         * @param {string} name - Name des Sheets
         * @returns {Sheet|null} - Das Sheet oder null, wenn nicht gefunden
         */
        const getSheet = (name) => {
            if (_cache.sheets.has(name)) {
                return _cache.sheets.get(name);
            }

            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
            if (sheet) {
                _cache.sheets.set(name, sheet);
            }
            return sheet;
        };

        /**
         * Aktualisiert ein Datenblatt (Einnahmen, Ausgaben, Eigenbelege)
         * @param {Sheet} sheet - Das zu aktualisierende Sheet
         * @returns {boolean} - true bei erfolgreicher Aktualisierung
         */
        const refreshDataSheet = sheet => {
            try {
                const lastRow = sheet.getLastRow();
                if (lastRow < 2) return true; // Keine Daten zum Aktualisieren

                const numRows = lastRow - 1;
                const name = sheet.getName();

                // Passende Spaltenkonfiguration für das entsprechende Sheet auswählen
                let columns;
                if (name === "Einnahmen") {
                    columns = config$1.einnahmen.columns;
                } else if (name === "Ausgaben") {
                    columns = config$1.ausgaben.columns;
                } else if (name === "Eigenbelege") {
                    columns = config$1.eigenbelege.columns;
                } else {
                    return false; // Unbekanntes Sheet
                }

                // Spaltenbuchstaben aus den Indizes generieren (mit Cache für Effizienz)
                const columnLetters = {};
                Object.entries(columns).forEach(([key, index]) => {
                    columnLetters[key] = Helpers.getColumnLetter(index);
                });

                // Batch-Array für Formeln erstellen (effizienter als einzelne Range-Updates)
                const formulasBatch = {};

                // MwSt-Betrag
                formulasBatch[columns.mwstBetrag] = Array.from(
                    {length: numRows},
                    (_, i) => [`=${columnLetters.nettobetrag}${i + 2}*${columnLetters.mwstSatz}${i + 2}`]
                );

                // Brutto-Betrag
                formulasBatch[columns.bruttoBetrag] = Array.from(
                    {length: numRows},
                    (_, i) => [`=${columnLetters.nettobetrag}${i + 2}+${columnLetters.mwstBetrag}${i + 2}`]
                );

                // Bezahlter Betrag - für Teilzahlungen
                formulasBatch[columns.restbetragNetto] = Array.from(
                    {length: numRows},
                    (_, i) => [`=(${columnLetters.bruttoBetrag}${i + 2}-${columnLetters.bezahlt}${i + 2})/(1+${columnLetters.mwstSatz}${i + 2})`]
                );

                // Quartal
                formulasBatch[columns.quartal] = Array.from(
                    {length: numRows},
                    (_, i) => [`=IF(${columnLetters.datum}${i + 2}="";"";ROUNDUP(MONTH(${columnLetters.datum}${i + 2})/3;0))`]
                );

                if (name !== "Eigenbelege") {
                    // Für Einnahmen und Ausgaben: Zahlungsstatus
                    formulasBatch[columns.zahlungsstatus] = Array.from(
                        {length: numRows},
                        (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Bezahlt";"Teilbezahlt"))`]
                    );
                } else {
                    // Für Eigenbelege: Zahlungsstatus
                    formulasBatch[columns.zahlungsstatus] = Array.from(
                        {length: numRows},
                        (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Erstattet";"Teilerstattet"))`]
                    );
                }

                // Formeln in Batches anwenden (weniger API-Calls)
                Object.entries(formulasBatch).forEach(([col, formulas]) => {
                    sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulas);
                });

                // Bezahlter Betrag - Leerzeichen durch 0 ersetzen für Berechnungen
                const bezahltRange = sheet.getRange(2, columns.bezahlt, numRows, 1);
                const bezahltValues = bezahltRange.getValues();
                const updatedBezahltValues = bezahltValues.map(
                    ([val]) => [Helpers.isEmpty(val) ? 0 : val]
                );
                bezahltRange.setValues(updatedBezahltValues);

                // Dropdown-Validierungen je nach Sheet-Typ setzen
                setDropdownValidations(sheet, name, numRows, columns);

                // Bedingte Formatierung für Status-Spalte
                if (name !== "Eigenbelege") {
                    // Für Einnahmen und Ausgaben: Zahlungsstatus
                    Helpers.setConditionalFormattingForColumn(sheet, columnLetters.zahlungsstatus, [
                        {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
                        {value: "Teilbezahlt", background: "#FFEB9C", fontColor: "#9C6500"},
                        {value: "Bezahlt", background: "#C6EFCE", fontColor: "#006100"}
                    ]);
                } else {
                    // Für Eigenbelege: Status
                    Helpers.setConditionalFormattingForColumn(sheet, columnLetters.zahlungsstatus, [
                        {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
                        {value: "Teilerstattet", background: "#FFEB9C", fontColor: "#9C6500"},
                        {value: "Erstattet", background: "#C6EFCE", fontColor: "#006100"}
                    ]);

                }

                // Spaltenbreiten automatisch anpassen
                sheet.autoResizeColumns(1, sheet.getLastColumn());

                return true;
            } catch (e) {
                console.error(`Fehler beim Aktualisieren von ${sheet.getName()}:`, e);
                return false;
            }
        };

        /**
         * Setzt Dropdown-Validierungen für ein Sheet
         * @param {Sheet} sheet - Das Sheet
         * @param {string} sheetName - Name des Sheets
         * @param {number} numRows - Anzahl der Datenzeilen
         * @param {Object} columns - Spaltenkonfiguration
         */
        const setDropdownValidations = (sheet, sheetName, numRows, columns) => {
            if (sheetName === "Einnahmen") {
                Validator.validateDropdown(
                    sheet, 2, columns.kategorie, numRows, 1,
                    Object.keys(config$1.einnahmen.categories)
                );
            } else if (sheetName === "Ausgaben") {
                Validator.validateDropdown(
                    sheet, 2, columns.kategorie, numRows, 1,
                    Object.keys(config$1.ausgaben.categories)
                );
            } else if (sheetName === "Eigenbelege") {
                Validator.validateDropdown(
                    sheet, 2, columns.kategorie, numRows, 1,
                    Object.keys(config$1.eigenbelege.categories)
                );
            }

            // Zahlungsart-Dropdown für alle Blätter
            Validator.validateDropdown(
                sheet, 2, columns.zahlungsart, numRows, 1,
                config$1.common.paymentType
            );
        };

        /**
         * Aktualisiert das Bankbewegungen-Sheet
         * @param {Sheet} sheet - Das Bankbewegungen-Sheet
         * @returns {boolean} - true bei erfolgreicher Aktualisierung
         */
        const refreshBankSheet = sheet => {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const lastRow = sheet.getLastRow();
                if (lastRow < 3) return true; // Nicht genügend Daten zum Aktualisieren

                const firstDataRow = 3; // Erste Datenzeile (nach Header-Zeile)
                const numDataRows = lastRow - firstDataRow + 1;
                const transRows = lastRow - firstDataRow - 1; // Anzahl der Transaktionszeilen ohne die letzte Zeile

                // Bankbewegungen-Konfiguration für Spalten
                const columns = config$1.bankbewegungen.columns;

                // Spaltenbuchstaben aus den Indizes generieren
                const columnLetters = {};
                Object.entries(columns).forEach(([key, index]) => {
                    columnLetters[key] = Helpers.getColumnLetter(index);
                });

                // 1. Saldo-Formeln setzen (jede Zeile verwendet den Saldo der vorherigen Zeile + aktuellen Betrag)
                if (transRows > 0) {
                    sheet.getRange(firstDataRow, columns.saldo, transRows, 1).setFormulas(
                        Array.from({length: transRows}, (_, i) =>
                            [`=${columnLetters.saldo}${firstDataRow + i - 1}+${columnLetters.betrag}${firstDataRow + i}`]
                        )
                    );
                }

                // 2. Transaktionstyp basierend auf dem Betrag setzen (Einnahme/Ausgabe)
                const amounts = sheet.getRange(firstDataRow, columns.betrag, numDataRows, 1).getValues();
                const typeValues = amounts.map(([val]) => {
                    const amt = parseFloat(val) || 0;
                    return [amt > 0 ? "Einnahme" : amt < 0 ? "Ausgabe" : ""];
                });
                sheet.getRange(firstDataRow, columns.transaktionstyp, numDataRows, 1).setValues(typeValues);

                // 3. Dropdown-Validierungen für Typ, Kategorie und Konten
                applyBankSheetValidations(sheet, firstDataRow, numDataRows, columns);

                // 4. Bedingte Formatierung für Transaktionstyp-Spalte
                Helpers.setConditionalFormattingForColumn(sheet, columnLetters.transaktionstyp, [
                    {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100"},
                    {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006"}
                ]);

                // 5. REFERENZEN-MATCHING: Suche nach Referenzen in Einnahmen- und Ausgaben-Sheets
                performBankReferenceMatching(ss, sheet, firstDataRow, numDataRows, lastRow, columns, columnLetters);

                // 6. Endsaldo-Zeile aktualisieren
                updateEndSaldoRow(sheet, lastRow, columns, columnLetters);

                // 7. Spaltenbreiten anpassen
                sheet.autoResizeColumns(1, sheet.getLastColumn());

                return true;
            } catch (e) {
                console.error("Fehler beim Aktualisieren des Bankbewegungen-Sheets:", e);
                return false;
            }
        };

        /**
         * Wendet Validierungen auf das Bankbewegungen-Sheet an
         * @param {Sheet} sheet - Das Sheet
         * @param {number} firstDataRow - Erste Datenzeile
         * @param {number} numDataRows - Anzahl der Datenzeilen
         * @param {Object} columns - Spaltenkonfiguration
         */
        const applyBankSheetValidations = (sheet, firstDataRow, numDataRows, columns) => {
            // Validierung für Transaktionstyp
            Validator.validateDropdown(
                sheet, firstDataRow, columns.transaktionstyp, numDataRows, 1,
                config$1.bankbewegungen.types
            );

            // Validierung für Kategorie
            Validator.validateDropdown(
                sheet, firstDataRow, columns.kategorie, numDataRows, 1,
                config$1.bankbewegungen.categories
            );

            // Konten für Dropdown-Validierung sammeln
            const allowedKontoSoll = new Set();
            const allowedGegenkonto = new Set();

            // Konten aus Einnahmen und Ausgaben sammeln
            Object.values(config$1.einnahmen.kontoMapping).forEach(m => {
                if (m.soll) allowedKontoSoll.add(m.soll);
                if (m.gegen) allowedGegenkonto.add(m.gegen);
            });

            Object.values(config$1.ausgaben.kontoMapping).forEach(m => {
                if (m.soll) allowedKontoSoll.add(m.soll);
                if (m.gegen) allowedGegenkonto.add(m.gegen);
            });

            // Dropdown-Validierungen für Konten setzen
            Validator.validateDropdown(
                sheet, firstDataRow, columns.kontoSoll, numDataRows, 1,
                Array.from(allowedKontoSoll)
            );

            Validator.validateDropdown(
                sheet, firstDataRow, columns.kontoHaben, numDataRows, 1,
                Array.from(allowedGegenkonto)
            );
        };

        /**
         * Aktualisiert die Endsaldo-Zeile im Bankbewegungen-Sheet
         * @param {Sheet} sheet - Das Sheet
         * @param {number} lastRow - Letzte Zeile
         * @param {Object} columns - Spaltenkonfiguration
         * @param {Object} columnLetters - Buchstaben für die Spalten
         */
        const updateEndSaldoRow = (sheet, lastRow, columns, columnLetters) => {
            // Endsaldo-Zeile überprüfen
            const lastRowText = sheet.getRange(lastRow, columns.buchungstext).getDisplayValue().toString().trim().toLowerCase();
            const formattedDate = Utilities.formatDate(
                new Date(),
                Session.getScriptTimeZone(),
                "dd.MM.yyyy"
            );

            if (lastRowText === "endsaldo") {
                // Endsaldo-Zeile aktualisieren
                sheet.getRange(lastRow, columns.datum).setValue(formattedDate);
                sheet.getRange(lastRow, columns.saldo).setFormula(`=${columnLetters.saldo}${lastRow - 1}`);
            } else {
                // Neue Endsaldo-Zeile erstellen
                const newRow = Array(sheet.getLastColumn()).fill("");
                newRow[columns.datum - 1] = formattedDate;
                newRow[columns.buchungstext - 1] = "Endsaldo";
                newRow[columns.saldo - 1] = sheet.getRange(lastRow, columns.saldo).getValue();

                sheet.appendRow(newRow);
            }
        };

        /**
         * Erstellt eine Map aus Referenznummern für schnellere Suche
         * @param {Array} data - Array mit Referenznummern und Beträgen
         * @returns {Object} - Map mit Referenznummern als Keys
         */
        const createReferenceMap = (data) => {
            // Cache prüfen und nutzen
            if (arguments.length === 1 && data === "einnahmen" && _cache.references.einnahmen) {
                return _cache.references.einnahmen;
            }
            if (arguments.length === 1 && data === "ausgaben" && _cache.references.ausgaben) {
                return _cache.references.ausgaben;
            }

            const map = {};

            // Keine Daten vorhanden
            if (!data || !data.length) {
                return map;
            }

            // Da die Indizes bei data[] sich auf die Spalten im Range beziehen (nicht auf das komplette Sheet),
            // müssen wir wissen, welche Spalten in welcher Reihenfolge geladen wurden
            // In der aktuellen Implementation werden diese Spalten geladen:
            // Rechnungsnummer(0), Nettobetrag(3), MwSt-Satz(4), Bezahlt(7)

            for (let i = 0; i < data.length; i++) {
                const ref = data[i][0]; // Referenz (erste Spalte im geladenen Range)
                if (Helpers.isEmpty(ref)) continue;

                // Entferne "G-" Prefix für den Key, falls vorhanden (für Gutschriften)
                let key = ref.trim();
                const isGutschrift = key.startsWith("G-");
                if (isGutschrift) {
                    key = key.substring(2); // Entferne "G-" Prefix für den Schlüssel
                }

                // Alternativschlüssel: Normalisierter Text für unscharfe Suche
                const normalizedKey = Helpers.normalizeText(key);

                // Netto-Betrag aus dritter Spalte im geladenen Range
                let betrag = 0;
                if (!Helpers.isEmpty(data[i][3])) {
                    betrag = Helpers.parseCurrency(data[i][3]);

                    // Bei Gutschriften ist der Betrag im Sheet negativ, wir speichern den Absolutwert
                    betrag = Math.abs(betrag);
                }

                // MwSt-Satz aus vierter Spalte im geladenen Range
                let mwstRate = 0;
                if (!Helpers.isEmpty(data[i][4])) {
                    mwstRate = Helpers.parseMwstRate(data[i][4]);
                }

                // Bezahlter Betrag aus achter Spalte im geladenen Range
                let bezahlt = 0;
                if (!Helpers.isEmpty(data[i][7])) {
                    bezahlt = Helpers.parseCurrency(data[i][7]);
                    // Bei Gutschriften ist der bezahlte Betrag im Sheet negativ, wir speichern den Absolutwert
                    bezahlt = Math.abs(bezahlt);
                }

                // Bruttobetrag berechnen
                const brutto = betrag * (1 + mwstRate/100);

                // Speichere auch den Zeilen-Index und die Beträge
                map[key] = {
                    ref: ref.trim(), // Originale Referenz mit G-Prefix, falls vorhanden
                    normalizedRef: normalizedKey, // Für unscharfe Suche
                    row: i + 2,      // +2 weil wir bei Zeile 2 beginnen (erste Zeile ist Header)
                    betrag: betrag,
                    mwstRate: mwstRate,
                    brutto: brutto,
                    bezahlt: bezahlt,
                    offen: Helpers.round(brutto - bezahlt, 2),
                    isGutschrift: isGutschrift
                };

                // Doppelten Eintrag mit normalisiertem Schlüssel anlegen, falls nötig
                if (normalizedKey !== key && !map[normalizedKey]) {
                    map[normalizedKey] = map[key];
                }
            }

            // Ergebnis cachen, wenn es sich um ein komplettes Sheet handelt
            if (arguments.length === 1) {
                if (data === "einnahmen") {
                    _cache.references.einnahmen = map;
                } else if (data === "ausgaben") {
                    _cache.references.ausgaben = map;
                }
            }

            return map;
        };

        /**
         * Führt das Matching von Bankbewegungen mit Rechnungen durch
         * @param {Spreadsheet} ss - Das Spreadsheet
         * @param {Sheet} sheet - Das Bankbewegungen-Sheet
         * @param {number} firstDataRow - Erste Datenzeile
         * @param {number} numDataRows - Anzahl der Datenzeilen
         * @param {number} lastRow - Letzte Zeile
         * @param {Object} columns - Spaltenkonfiguration
         * @param {Object} columnLetters - Buchstaben für die Spalten
         */
        const performBankReferenceMatching = (ss, sheet, firstDataRow, numDataRows, lastRow, columns, columnLetters) => {
            // Konfigurationen für Spaltenindizes
            const einnahmenCols = config$1.einnahmen.columns;
            const ausgabenCols = config$1.ausgaben.columns;

            // Referenzdaten laden für Einnahmen
            const einnahmenSheet = getSheet("Einnahmen");
            let einnahmenData = [], einnahmenMap = {};

            if (einnahmenSheet && einnahmenSheet.getLastRow() > 1) {
                const numEinnahmenRows = einnahmenSheet.getLastRow() - 1;
                // Die relevanten Spalten laden basierend auf der Konfiguration
                einnahmenData = einnahmenSheet.getRange(2, einnahmenCols.rechnungsnummer, numEinnahmenRows, 8).getDisplayValues();
                einnahmenMap = createReferenceMap(einnahmenData);
            }

            // Referenzdaten laden für Ausgaben
            const ausgabenSheet = getSheet("Ausgaben");
            let ausgabenData = [], ausgabenMap = {};

            if (ausgabenSheet && ausgabenSheet.getLastRow() > 1) {
                const numAusgabenRows = ausgabenSheet.getLastRow() - 1;
                // Die relevanten Spalten laden basierend auf der Konfiguration
                ausgabenData = ausgabenSheet.getRange(2, ausgabenCols.rechnungsnummer, numAusgabenRows, 8).getDisplayValues();
                ausgabenMap = createReferenceMap(ausgabenData);
            }

            // Bankbewegungen Daten für Verarbeitung holen
            const bankData = sheet.getRange(firstDataRow, 1, numDataRows, columns.matchInfo).getDisplayValues();

            // Ergebnis-Arrays für Batch-Update
            const matchResults = [];
            const kontoSollResults = [];
            const kontoHabenResults = [];

            // Banking-Zuordnungen für spätere Synchronisierung mit Einnahmen/Ausgaben
            const bankZuordnungen = {};

            // Sammeln aller gültigen Konten für die Validierung
            const allowedKontoSoll = new Set();
            const allowedGegenkonto = new Set();

            // Konten aus Einnahmen und Ausgaben sammeln
            Object.values(config$1.einnahmen.kontoMapping).forEach(m => {
                if (m.soll) allowedKontoSoll.add(m.soll);
                if (m.gegen) allowedGegenkonto.add(m.gegen);
            });

            Object.values(config$1.ausgaben.kontoMapping).forEach(m => {
                if (m.soll) allowedKontoSoll.add(m.soll);
                if (m.gegen) allowedGegenkonto.add(m.gegen);
            });

            // Fallback-Konto wenn kein Match - das erste Konto aus den erlaubten Konten
            const fallbackKontoSoll = allowedKontoSoll.size > 0 ? Array.from(allowedKontoSoll)[0] : "";
            const fallbackKontoHaben = allowedGegenkonto.size > 0 ? Array.from(allowedGegenkonto)[0] : "";

            // Durchlaufe jede Bankbewegung und suche nach Übereinstimmungen
            for (let i = 0; i < bankData.length; i++) {
                const rowIndex = i + firstDataRow;
                const row = bankData[i];

                // Prüfe, ob es sich um die Endsaldo-Zeile handelt
                const label = row[columns.buchungstext - 1] ? row[columns.buchungstext - 1].toString().trim().toLowerCase() : "";
                if (rowIndex === lastRow && label === "endsaldo") {
                    matchResults.push([""]);
                    kontoSollResults.push([""]);
                    kontoHabenResults.push([""]);
                    continue;
                }

                const tranType = row[columns.transaktionstyp - 1]; // Einnahme/Ausgabe
                const refNumber = row[columns.referenz - 1];       // Referenznummer
                let matchInfo = "";

                // Kontonummern basierend auf Kategorie vorbereiten
                let kontoSoll = "", kontoHaben = "";
                const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";

                // Nur prüfen, wenn Referenz nicht leer ist
                if (!Helpers.isEmpty(refNumber)) {
                    let matchFound = false;
                    const betragValue = Math.abs(Helpers.parseCurrency(row[columns.betrag - 1]));

                    if (tranType === "Einnahme") {
                        // Suche in Einnahmen
                        const matchResult = findMatch(refNumber, einnahmenMap, betragValue);
                        if (matchResult) {
                            matchFound = true;
                            matchInfo = processEinnahmeMatch(matchResult, betragValue, row, columns, einnahmenSheet, einnahmenCols);

                            // Für spätere Markierung merken
                            const key = `einnahme#${matchResult.row}`;
                            bankZuordnungen[key] = {
                                typ: "einnahme",
                                row: matchResult.row,
                                bankDatum: row[columns.datum - 1],
                                matchInfo: matchInfo,
                                transTyp: tranType
                            };
                        }
                    } else if (tranType === "Ausgabe") {
                        // Suche in Ausgaben
                        const matchResult = findMatch(refNumber, ausgabenMap, betragValue);
                        if (matchResult) {
                            matchFound = true;
                            matchInfo = processAusgabeMatch(matchResult, betragValue, row, columns, ausgabenSheet, ausgabenCols);

                            // Für spätere Markierung merken
                            const key = `ausgabe#${matchResult.row}`;
                            bankZuordnungen[key] = {
                                typ: "ausgabe",
                                row: matchResult.row,
                                bankDatum: row[columns.datum - 1],
                                matchInfo: matchInfo,
                                transTyp: tranType
                            };
                        }

                        // FALLS keine Übereinstimmung, auch in Einnahmen suchen (für Gutschriften)
                        if (!matchFound) {
                            const gutschriftMatch = findMatch(refNumber, einnahmenMap);
                            if (gutschriftMatch) {
                                matchFound = true;
                                matchInfo = processGutschriftMatch(gutschriftMatch, betragValue, row, columns, einnahmenSheet, einnahmenCols);

                                // Für spätere Markierung merken
                                const key = `einnahme#${gutschriftMatch.row}`;
                                bankZuordnungen[key] = {
                                    typ: "einnahme",
                                    row: gutschriftMatch.row,
                                    bankDatum: row[columns.datum - 1],
                                    matchInfo: matchInfo,
                                    transTyp: "Gutschrift"
                                };
                            }
                        }
                    }

                    // Spezialfälle prüfen (Gesellschaftskonto, Holding)
                    if (!matchFound) {
                        const lcRef = refNumber.toString().toLowerCase();
                        if (lcRef.includes("gesellschaftskonto")) {
                            matchFound = true;
                            matchInfo = tranType === "Einnahme"
                                ? "Gesellschaftskonto (Einnahme)"
                                : "Gesellschaftskonto (Ausgabe)";
                        } else if (lcRef.includes("holding")) {
                            matchFound = true;
                            matchInfo = tranType === "Einnahme"
                                ? "Holding (Einnahme)"
                                : "Holding (Ausgabe)";
                        }
                    }
                }

                // Kontonummern basierend auf Kategorie setzen
                if (category) {
                    // Den richtigen Mapping-Typ basierend auf der Transaktionsart auswählen
                    const mappingSource = tranType === "Einnahme" ?
                        config$1.einnahmen.kontoMapping :
                        config$1.ausgaben.kontoMapping;

                    // Mapping für die Kategorie finden
                    if (mappingSource && mappingSource[category]) {
                        const mapping = mappingSource[category];
                        // Nutze die Kontonummern aus dem Mapping
                        kontoSoll = mapping.soll || fallbackKontoSoll;
                        kontoHaben = mapping.gegen || fallbackKontoHaben;
                    } else {
                        // Fallback wenn kein Mapping gefunden wurde - erstes zulässiges Konto
                        kontoSoll = fallbackKontoSoll;
                        kontoHaben = fallbackKontoHaben;
                    }
                }

                // Ergebnisse für Batch-Update sammeln
                matchResults.push([matchInfo]);
                kontoSollResults.push([kontoSoll]);
                kontoHabenResults.push([kontoHaben]);
            }

            // Batch-Updates ausführen für bessere Performance
            sheet.getRange(firstDataRow, columns.matchInfo, numDataRows, 1).setValues(matchResults);
            sheet.getRange(firstDataRow, columns.kontoSoll, numDataRows, 1).setValues(kontoSollResults);
            sheet.getRange(firstDataRow, columns.kontoHaben, numDataRows, 1).setValues(kontoHabenResults);

            // Formatiere die gesamten Zeilen basierend auf dem Match-Typ
            formatMatchedRows(sheet, firstDataRow, matchResults, columns);

            // Bedingte Formatierung für Match-Spalte mit verbesserten Farben
            setMatchColumnFormatting(sheet, columnLetters.matchInfo);

            // Setze farbliche Markierung in den Einnahmen/Ausgaben Sheets basierend auf Zahlungsstatus
            markPaidInvoices(einnahmenSheet, ausgabenSheet, bankZuordnungen);
        };
        /**
         * Findet eine Übereinstimmung in der Referenz-Map
         * @param {string} reference - Zu suchende Referenz
         * @param {Object} refMap - Map mit Referenznummern
         * @param {number} betrag - Betrag der Bankbewegung (absoluter Wert)
         * @returns {Object|null} - Gefundene Übereinstimmung oder null
         */
        const findMatch = (reference, refMap, betrag = null) => {
            // Keine Daten
            if (!reference || !refMap) return null;

            // Normalisierte Suche vorbereiten
            const normalizedRef = Helpers.normalizeText(reference);

            // Priorität der Matching-Strategien:
            // 1. Exakte Übereinstimmung
            // 2. Normalisierte Übereinstimmung
            // 3. Enthält-Beziehung (in beide Richtungen)

            // 1. Exakte Übereinstimmung
            if (refMap[reference]) {
                return evaluateMatch(refMap[reference], betrag);
            }

            // 2. Normalisierte Übereinstimmung
            if (normalizedRef && refMap[normalizedRef]) {
                return evaluateMatch(refMap[normalizedRef], betrag);
            }

            // 3. Teilweise Übereinstimmung (mit Cache für Performance)
            const candidateKeys = Object.keys(refMap);

            // Zuerst prüfen wir, ob die Referenz in einem Schlüssel enthalten ist
            for (const key of candidateKeys) {
                if (key.includes(reference) || reference.includes(key)) {
                    return evaluateMatch(refMap[key], betrag);
                }
            }

            // Falls keine exakten Treffer, probieren wir es mit normalisierten Werten
            for (const key of candidateKeys) {
                const normalizedKey = Helpers.normalizeText(key);
                if (normalizedKey.includes(normalizedRef) || normalizedRef.includes(normalizedKey)) {
                    return evaluateMatch(refMap[key], betrag);
                }
            }

            return null;
        };

        /**
         * Bewertet die Qualität einer Übereinstimmung basierend auf Beträgen
         * @param {Object} match - Die gefundene Übereinstimmung
         * @param {number} betrag - Der Betrag zum Vergleich
         * @returns {Object} - Übereinstimmung mit zusätzlichen Infos
         */
        const evaluateMatch = (match, betrag = null) => {
            if (!match) return null;

            // Behalte die ursprüngliche Übereinstimmung
            const result = { ...match };

            // Wenn kein Betrag zum Vergleich angegeben wurde
            if (betrag === null) {
                result.matchType = "Referenz-Match";
                return result;
            }

            // Hole die Beträge aus der Übereinstimmung
            const matchBrutto = Math.abs(match.brutto);
            const matchBezahlt = Math.abs(match.bezahlt);

            // Toleranzwert für Betragsabweichungen (2 Cent)
            const tolerance = 0.02;

            // Fall 1: Betrag entspricht genau dem Bruttobetrag (Vollständige Zahlung)
            if (Helpers.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
                result.matchType = "Vollständige Zahlung";
                return result;
            }

            // Fall 2: Position ist bereits vollständig bezahlt
            if (Helpers.isApproximatelyEqual(matchBezahlt, matchBrutto, tolerance) && matchBezahlt > 0) {
                result.matchType = "Vollständige Zahlung";
                return result;
            }

            // Fall 3: Teilzahlung (Bankbetrag kleiner als Rechnungsbetrag)
            // Nur als Teilzahlung markieren, wenn der Betrag deutlich kleiner ist (> 10% Differenz)
            if (betrag < matchBrutto && (matchBrutto - betrag) > (matchBrutto * 0.1)) {
                result.matchType = "Teilzahlung";
                return result;
            }

            // Fall 4: Betrag ist größer als Bruttobetrag, aber vermutlich trotzdem die richtige Zahlung
            // (z.B. wegen Rundungen oder kleinen Gebühren)
            if (betrag > matchBrutto && Helpers.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
                result.matchType = "Vollständige Zahlung";
                return result;
            }

            // Fall 5: Bei allen anderen Fällen (Beträge weichen stärker ab)
            result.matchType = "Unsichere Zahlung";
            result.betragsDifferenz = Helpers.round(Math.abs(betrag - matchBrutto), 2);
            return result;
        };

        /**
         * Verarbeitet eine Einnahmen-Übereinstimmung
         * @returns {string} Formatierte Match-Information
         */
        const processEinnahmeMatch = (matchResult, betragValue, row, columns, einnahmenSheet, einnahmenCols) => {
            let matchStatus = "";

            // Je nach Match-Typ unterschiedliche Statusinformationen
            if (matchResult.matchType) {
                // Bei "Unsichere Zahlung" auch die Differenz anzeigen
                if (matchResult.matchType === "Unsichere Zahlung" && matchResult.betragsDifferenz) {
                    matchStatus = ` (${matchResult.matchType}, Diff: ${matchResult.betragsDifferenz}€)`;
                } else {
                    matchStatus = ` (${matchResult.matchType})`;
                }
            }

            // Prüfen, ob Zahldatum in Einnahmen/Ausgaben aktualisiert werden soll
            if ((matchResult.matchType === "Vollständige Zahlung" ||
                    matchResult.matchType === "Teilzahlung") &&
                einnahmenSheet) {
                // Bankbewegungsdatum holen (aus Spalte A)
                const zahlungsDatum = row[columns.datum - 1];
                if (zahlungsDatum) {
                    // Zahldatum im Einnahmen-Sheet aktualisieren (nur wenn leer)
                    const einnahmenRow = matchResult.row;
                    const zahldatumRange = einnahmenSheet.getRange(einnahmenRow, einnahmenCols.zahlungsdatum);
                    const aktuellDatum = zahldatumRange.getValue();

                    if (Helpers.isEmpty(aktuellDatum)) {
                        zahldatumRange.setValue(zahlungsDatum);
                        matchStatus += " ✓ Datum aktualisiert";
                    }
                }
            }

            return `Einnahme #${matchResult.row}${matchStatus}`;
        };

        /**
         * Verarbeitet eine Ausgaben-Übereinstimmung
         * @returns {string} Formatierte Match-Information
         */
        const processAusgabeMatch = (matchResult, betragValue, row, columns, ausgabenSheet, ausgabenCols) => {
            let matchStatus = "";

            // Je nach Match-Typ unterschiedliche Statusinformationen für Ausgaben
            if (matchResult.matchType) {
                if (matchResult.matchType === "Unsichere Zahlung" && matchResult.betragsDifferenz) {
                    matchStatus = ` (${matchResult.matchType}, Diff: ${matchResult.betragsDifferenz}€)`;
                } else {
                    matchStatus = ` (${matchResult.matchType})`;
                }
            }

            // Prüfen, ob Zahldatum in Ausgaben aktualisiert werden soll
            if ((matchResult.matchType === "Vollständige Zahlung" ||
                    matchResult.matchType === "Teilzahlung") &&
                ausgabenSheet) {
                // Bankbewegungsdatum holen (aus Spalte A)
                const zahlungsDatum = row[columns.datum - 1];
                if (zahlungsDatum) {
                    // Zahldatum im Ausgaben-Sheet aktualisieren (nur wenn leer)
                    const ausgabenRow = matchResult.row;
                    const zahldatumRange = ausgabenSheet.getRange(ausgabenRow, ausgabenCols.zahlungsdatum);
                    const aktuellDatum = zahldatumRange.getValue();

                    if (Helpers.isEmpty(aktuellDatum)) {
                        zahldatumRange.setValue(zahlungsDatum);
                        matchStatus += " ✓ Datum aktualisiert";
                    }
                }
            }

            return `Ausgabe #${matchResult.row}${matchStatus}`;
        };

        /**
         * Verarbeitet eine Gutschrift-Übereinstimmung
         * @returns {string} Formatierte Match-Information
         */
        const processGutschriftMatch = (gutschriftMatch, betragValue, row, columns, einnahmenSheet, einnahmenCols) => {
            let matchStatus = "";

            // Bei Gutschriften könnte der Betrag abweichen (z.B. Teilgutschrift)
            // Prüfen, ob die Beträge plausibel sind
            const gutschriftBetrag = Math.abs(gutschriftMatch.brutto);

            if (Helpers.isApproximatelyEqual(betragValue, gutschriftBetrag, 0.01)) {
                // Beträge stimmen genau überein
                matchStatus = " (Vollständige Gutschrift)";
            } else if (betragValue < gutschriftBetrag) {
                // Teilgutschrift (Gutschriftbetrag kleiner als ursprünglicher Rechnungsbetrag)
                matchStatus = " (Teilgutschrift)";
            } else {
                // Ungewöhnlicher Fall - Gutschriftbetrag größer als Rechnungsbetrag
                matchStatus = " (Ungewöhnliche Gutschrift)";
            }

            // Bei Gutschriften auch im Einnahmen-Sheet aktualisieren - hier als negative Zahlung
            if (einnahmenSheet) {
                // Bankbewegungsdatum holen (aus Spalte A)
                const gutschriftDatum = row[columns.datum - 1];
                if (gutschriftDatum) {
                    // Gutschriftdatum im Einnahmen-Sheet aktualisieren und "G-" vor die Referenz setzen
                    const einnahmenRow = gutschriftMatch.row;

                    // Zahldatum aktualisieren (nur wenn leer)
                    const zahldatumRange = einnahmenSheet.getRange(einnahmenRow, einnahmenCols.zahlungsdatum);
                    const aktuellDatum = zahldatumRange.getValue();

                    if (Helpers.isEmpty(aktuellDatum)) {
                        zahldatumRange.setValue(gutschriftDatum);
                        matchStatus += " ✓ Datum aktualisiert";
                    }

                    // Optional: Die Referenz mit "G-" kennzeichnen, um Gutschrift zu markieren
                    const refRange = einnahmenSheet.getRange(einnahmenRow, einnahmenCols.rechnungsnummer);
                    const currentRef = refRange.getValue();
                    if (currentRef && !currentRef.toString().startsWith("G-")) {
                        refRange.setValue("G-" + currentRef);
                        matchStatus += " ✓ Ref. markiert";
                    }
                }
            }

            return `Gutschrift zu Einnahme #${gutschriftMatch.row}${matchStatus}`;
        };

        /**
         * Formatiert Zeilen basierend auf dem Match-Typ
         * @param {Sheet} sheet - Das Sheet
         * @param {number} firstDataRow - Erste Datenzeile
         * @param {Array} matchResults - Array mit Match-Infos
         * @param {Object} columns - Spaltenkonfiguration
         */
        const formatMatchedRows = (sheet, firstDataRow, matchResults, columns) => {
            // Performance-optimiertes Batch-Update vorbereiten
            const formatBatches = {
                'Einnahme': { rows: [], color: "#EAF1DD" },  // Helles Grün (Grundfarbe für Einnahmen)
                'Vollständige Zahlung (Einnahme)': { rows: [], color: "#C6EFCE" },  // Kräftiges Grün
                'Teilzahlung (Einnahme)': { rows: [], color: "#FCE4D6" },  // Helles Orange
                'Ausgabe': { rows: [], color: "#FFCCCC" },   // Helles Rosa (Grundfarbe für Ausgaben)
                'Vollständige Zahlung (Ausgabe)': { rows: [], color: "#FFC7CE" },  // Helles Rot
                'Teilzahlung (Ausgabe)': { rows: [], color: "#FCE4D6" },  // Helles Orange
                'Gutschrift': { rows: [], color: "#E6E0FF" },  // Helles Lila
                'Gesellschaftskonto/Holding': { rows: [], color: "#FFEB9C" }  // Helles Gelb
            };

            // Zeilen nach Kategorien gruppieren
            matchResults.forEach((matchInfo, index) => {
                const rowIndex = firstDataRow + index;
                const matchText = (matchInfo && matchInfo[0]) ? matchInfo[0].toString() : "";

                if (!matchText) return; // Überspringe leere Matches

                if (matchText.includes("Einnahme")) {
                    if (matchText.includes("Vollständige Zahlung")) {
                        formatBatches['Vollständige Zahlung (Einnahme)'].rows.push(rowIndex);
                    } else if (matchText.includes("Teilzahlung")) {
                        formatBatches['Teilzahlung (Einnahme)'].rows.push(rowIndex);
                    } else {
                        formatBatches['Einnahme'].rows.push(rowIndex);
                    }
                } else if (matchText.includes("Ausgabe")) {
                    if (matchText.includes("Vollständige Zahlung")) {
                        formatBatches['Vollständige Zahlung (Ausgabe)'].rows.push(rowIndex);
                    } else if (matchText.includes("Teilzahlung")) {
                        formatBatches['Teilzahlung (Ausgabe)'].rows.push(rowIndex);
                    } else {
                        formatBatches['Ausgabe'].rows.push(rowIndex);
                    }
                } else if (matchText.includes("Gutschrift")) {
                    formatBatches['Gutschrift'].rows.push(rowIndex);
                } else if (matchText.includes("Gesellschaftskonto") || matchText.includes("Holding")) {
                    formatBatches['Gesellschaftskonto/Holding'].rows.push(rowIndex);
                }
            });

            // Batch-Formatting anwenden
            Object.values(formatBatches).forEach(batch => {
                if (batch.rows.length > 0) {
                    // Gruppen von maximal 20 Zeilen formatieren (API-Limits vermeiden)
                    const chunkSize = 20;
                    for (let i = 0; i < batch.rows.length; i += chunkSize) {
                        const chunk = batch.rows.slice(i, i + chunkSize);
                        chunk.forEach(rowIndex => {
                            try {
                                sheet.getRange(rowIndex, 1, 1, columns.matchInfo)
                                    .setBackground(batch.color);
                            } catch (e) {
                                console.error(`Fehler beim Formatieren von Zeile ${rowIndex}:`, e);
                            }
                        });

                        // Kurze Pause um API-Limits zu vermeiden
                        if (i + chunkSize < batch.rows.length) {
                            Utilities.sleep(50);
                        }
                    }
                }
            });
        };

        /**
         * Setzt bedingte Formatierung für die Match-Spalte
         * @param {Sheet} sheet - Das Sheet
         * @param {string} columnLetter - Buchstabe für die Spalte
         */
        const setMatchColumnFormatting = (sheet, columnLetter) => {
            Helpers.setConditionalFormattingForColumn(sheet, columnLetter, [
                // Grundlegende Match-Typen mit beginsWith Pattern
                {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100", pattern: "beginsWith"},
                {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006", pattern: "beginsWith"},
                {value: "Gutschrift", background: "#E6E0FF", fontColor: "#5229A3", pattern: "beginsWith"},
                {value: "Gesellschaftskonto", background: "#FFEB9C", fontColor: "#9C6500", pattern: "beginsWith"},
                {value: "Holding", background: "#FFEB9C", fontColor: "#9C6500", pattern: "beginsWith"},

                // Zusätzliche Betragstypen mit contains Pattern
                {value: "Vollständige Zahlung", background: "#C6EFCE", fontColor: "#006100", pattern: "contains"},
                {value: "Teilzahlung", background: "#FCE4D6", fontColor: "#974706", pattern: "contains"},
                {value: "Unsichere Zahlung", background: "#F8CBAD", fontColor: "#843C0C", pattern: "contains"},
                {value: "Vollständige Gutschrift", background: "#E6E0FF", fontColor: "#5229A3", pattern: "contains"},
                {value: "Teilgutschrift", background: "#E6E0FF", fontColor: "#5229A3", pattern: "contains"},
                {value: "Ungewöhnliche Gutschrift", background: "#FFD966", fontColor: "#7F6000", pattern: "contains"}
            ]);
        };

        /**
         * Markiert bezahlte Einnahmen und Ausgaben farblich basierend auf dem Zahlungsstatus
         * @param {Sheet} einnahmenSheet - Das Einnahmen-Sheet
         * @param {Sheet} ausgabenSheet - Das Ausgaben-Sheet
         * @param {Object} bankZuordnungen - Zuordnungen aus dem Bankbewegungen-Sheet
         */
        const markPaidInvoices = (einnahmenSheet, ausgabenSheet, bankZuordnungen) => {
            // Markiere bezahlte Einnahmen
            if (einnahmenSheet && einnahmenSheet.getLastRow() > 1) {
                markPaidRows(einnahmenSheet, "Einnahmen", bankZuordnungen);
            }

            // Markiere bezahlte Ausgaben
            if (ausgabenSheet && ausgabenSheet.getLastRow() > 1) {
                markPaidRows(ausgabenSheet, "Ausgaben", bankZuordnungen);
            }
        };

        /**
         * Markiert bezahlte Zeilen in einem Sheet
         * @param {Sheet} sheet - Das zu aktualisierende Sheet
         * @param {string} sheetType - Typ des Sheets ("Einnahmen" oder "Ausgaben")
         * @param {Object} bankZuordnungen - Zuordnungen aus dem Bankbewegungen-Sheet
         */
        const markPaidRows = (sheet, sheetType, bankZuordnungen) => {
            // Konfiguration für das Sheet
            const columns = sheetType === "Einnahmen"
                ? config$1.einnahmen.columns
                : config$1.ausgaben.columns;

            // Hole Werte aus dem Sheet
            const numRows = sheet.getLastRow() - 1;
            const data = sheet.getRange(2, 1, numRows, columns.zahlungsdatum).getValues();

            // Batch-Arrays für die verschiedenen Status
            const fullPaidWithBankRows = [];
            const fullPaidRows = [];
            const partialPaidWithBankRows = [];
            const partialPaidRows = [];
            const gutschriftRows = [];
            const normalRows = [];

            // Bank-Abgleich-Updates sammeln
            const bankabgleichUpdates = [];

            // Datenzeilen durchgehen und in Kategorien einteilen
            for (let i = 0; i < data.length; i++) {
                const row = i + 2; // Aktuelle Zeile im Sheet
                const nettobetrag = Helpers.parseCurrency(data[i][columns.nettobetrag - 1]);
                const bezahltBetrag = Helpers.parseCurrency(data[i][columns.bezahlt - 1]);
                const zahlungsDatum = data[i][columns.zahlungsdatum - 1];
                const referenz = data[i][columns.rechnungsnummer - 1];

                // Prüfe, ob es eine Gutschrift ist
                const isGutschrift = referenz && referenz.toString().startsWith("G-");

                // Prüfe, ob diese Position im Banking-Sheet zugeordnet wurde
                const zuordnungsKey = `${sheetType.toLowerCase().slice(0, -1)}#${row}`;
                const hatBankzuordnung = bankZuordnungen[zuordnungsKey] !== undefined;

                // Zahlungsstatus berechnen
                const mwst = Helpers.parseMwstRate(data[i][columns.mwstSatz - 1]) / 100;
                const bruttoBetrag = nettobetrag * (1 + mwst);
                const isPaid = Math.abs(bezahltBetrag) >= Math.abs(bruttoBetrag) * 0.999; // 99.9% bezahlt wegen Rundungsfehlern
                const isPartialPaid = !isPaid && bezahltBetrag > 0;

                // Kategorie bestimmen
                if (isGutschrift) {
                    gutschriftRows.push(row);
                    // Bank-Abgleich-Info setzen
                    if (hatBankzuordnung) {
                        bankabgleichUpdates.push({
                            row,
                            value: getZuordnungsInfo(bankZuordnungen[zuordnungsKey])
                        });
                    }
                } else if (isPaid) {
                    if (zahlungsDatum) {
                        if (hatBankzuordnung) {
                            fullPaidWithBankRows.push(row);
                            bankabgleichUpdates.push({
                                row,
                                value: getZuordnungsInfo(bankZuordnungen[zuordnungsKey])
                            });
                        } else {
                            fullPaidRows.push(row);
                        }
                    } else {
                        // Bezahlt aber kein Datum
                        partialPaidRows.push(row);
                    }
                } else if (isPartialPaid) {
                    if (hatBankzuordnung) {
                        partialPaidWithBankRows.push(row);
                        bankabgleichUpdates.push({
                            row,
                            value: getZuordnungsInfo(bankZuordnungen[zuordnungsKey])
                        });
                    } else {
                        partialPaidRows.push(row);
                    }
                } else {
                    // Unbezahlt - normale Zeile
                    normalRows.push(row);
                    // Vorhandene Bank-Abgleich-Info entfernen falls vorhanden
                    if (sheet.getRange(row, columns.bankabgleich).getValue().toString().startsWith("✓ Bank:")) {
                        bankabgleichUpdates.push({
                            row,
                            value: ""
                        });
                    }
                }
            }

            // Färbe Zeilen im Batch-Modus (für bessere Performance)
            applyColorToRows(sheet, fullPaidWithBankRows, "#C6EFCE"); // Kräftigeres Grün
            applyColorToRows(sheet, fullPaidRows, "#EAF1DD"); // Helles Grün
            applyColorToRows(sheet, partialPaidWithBankRows, "#FFC7AA"); // Kräftigeres Orange
            applyColorToRows(sheet, partialPaidRows, "#FCE4D6"); // Helles Orange
            applyColorToRows(sheet, gutschriftRows, "#E6E0FF"); // Helles Lila
            applyColorToRows(sheet, normalRows, null); // Keine Farbe / Zurücksetzen

            // Bank-Abgleich-Updates in Batches ausführen
            if (bankabgleichUpdates.length > 0) {
                // Gruppiere Updates nach Wert für effizientere Batch-Updates
                const groupedUpdates = {};

                bankabgleichUpdates.forEach(update => {
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
        };

        /**
         * Wendet eine Farbe auf mehrere Zeilen an
         * @param {Sheet} sheet - Das Sheet
         * @param {Array} rows - Die zu färbenden Zeilennummern
         * @param {string|null} color - Die Hintergrundfarbe oder null zum Zurücksetzen
         */
        const applyColorToRows = (sheet, rows, color) => {
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
        };

        /**
         * Erstellt einen Informationstext für eine Bank-Zuordnung
         * @param {Object} zuordnung - Die Zuordnungsinformation
         * @returns {string} - Formatierter Informationstext
         */
        const getZuordnungsInfo = (zuordnung) => {
            if (!zuordnung) return "";

            let infoText = "✓ Bank: " + zuordnung.bankDatum;

            // Wenn es mehrere Zuordnungen gibt (z.B. bei aufgeteilten Zahlungen)
            if (zuordnung.additional && zuordnung.additional.length > 0) {
                infoText += " + " + zuordnung.additional.length + " weitere";
            }

            return infoText;
        };

        /**
         * Aktualisiert das aktive Tabellenblatt
         */
        const refreshActiveSheet = () => {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const sheet = ss.getActiveSheet();
                const name = sheet.getName();
                const ui = SpreadsheetApp.getUi();

                // Cache zurücksetzen
                clearCache();

                if (["Einnahmen", "Ausgaben", "Eigenbelege"].includes(name)) {
                    refreshDataSheet(sheet);
                    ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
                } else if (name === "Bankbewegungen") {
                    refreshBankSheet(sheet);
                    ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
                } else {
                    ui.alert("Für dieses Blatt gibt es keine Refresh-Funktion.");
                }
            } catch (e) {
                console.error("Fehler beim Aktualisieren des aktiven Sheets:", e);
                SpreadsheetApp.getUi().alert("Ein Fehler ist beim Aktualisieren aufgetreten: " + e.toString());
            }
        };

        /**
         * Aktualisiert alle relevanten Tabellenblätter
         */
        const refreshAllSheets = () => {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();

                // Cache zurücksetzen
                clearCache();

                // Sheets in der richtigen Reihenfolge aktualisieren, um Abhängigkeiten zu berücksichtigen
                const refreshOrder = ["Einnahmen", "Ausgaben", "Eigenbelege", "Bankbewegungen"];

                for (const name of refreshOrder) {
                    const sheet = getSheet(name);
                    if (sheet) {
                        name === "Bankbewegungen" ? refreshBankSheet(sheet) : refreshDataSheet(sheet);
                        // Kurze Pause einfügen, um API-Limits zu vermeiden
                        Utilities.sleep(100);
                    }
                }
            } catch (e) {
                console.error("Fehler beim Aktualisieren aller Sheets:", e);
            }
        };

        // Öffentliche API des Moduls
        return {
            refreshActiveSheet,
            refreshAllSheets
        };
    })();

    // file: src/uStVACalculator.js

    /**
     * Modul zur Berechnung der Umsatzsteuervoranmeldung (UStVA)
     * Unterstützt die Berechnung nach SKR04 für monatliche und quartalsweise Auswertungen
     */
    const UStVACalculator = (() => {
        // Zwischenspeicher für UStVA-Daten um redundante Berechnung zu vermeiden
        const _cache = {
            ustva: null,
            lastUpdated: null
        };

        /**
         * Cache leeren
         */
        const clearCache = () => {
            _cache.ustva = null;
            _cache.lastUpdated = null;
        };

        /**
         * Erstellt ein leeres UStVA-Datenobjekt mit Nullwerten
         * @returns {Object} Leere UStVA-Datenstruktur
         */
        const createEmptyUStVA = () => ({
            steuerpflichtige_einnahmen: 0,
            steuerfreie_inland_einnahmen: 0,
            steuerfreie_ausland_einnahmen: 0,
            steuerpflichtige_ausgaben: 0,
            steuerfreie_inland_ausgaben: 0,
            steuerfreie_ausland_ausgaben: 0,
            eigenbelege_steuerpflichtig: 0,
            eigenbelege_steuerfrei: 0,
            nicht_abzugsfaehige_vst: 0,
            ust_7: 0,
            ust_19: 0,
            vst_7: 0,
            vst_19: 0
        });

        /**
         * Verarbeitet eine Zeile aus den Einnahmen/Ausgaben/Eigenbelegen für die UStVA
         *
         * @param {Array} row - Datenzeile aus einem Sheet
         * @param {Object} data - UStVA-Datenobjekt nach Monaten
         * @param {boolean} isIncome - Handelt es sich um Einnahmen (true) oder Ausgaben (false)
         * @param {boolean} isEigen - Handelt es sich um Eigenbelege (true oder false)
         */
        const processUStVARow = (row, data, isIncome, isEigen = false) => {
            try {
                // Sheet-Typ bestimmen
                const sheetType = isIncome ? "einnahmen" : isEigen ? "eigenbelege" : "ausgaben";
                const columns = config$1[sheetType].columns;

                // Zahlungsdatum prüfen (nur abgeschlossene Zahlungen)
                const paymentDate = Helpers.parseDate(row[columns.zahlungsdatum - 1]);
                if (!paymentDate || paymentDate > new Date()) return;

                // Monat und Jahr prüfen (nur relevantes Geschäftsjahr)
                const month = Helpers.getMonthFromRow(row, sheetType);
                if (!month || month < 1 || month > 12) return;

                // Beträge aus der Zeile extrahieren
                const netto = Helpers.parseCurrency(row[columns.nettobetrag - 1]);
                const restNetto = Helpers.parseCurrency(row[columns.restbetragNetto - 1]) || 0; // Steuerbemessungsgrundlage für Teilzahlungen
                const gezahlt = netto - restNetto; // Tatsächlich gezahlter/erhaltener Betrag

                // Falls kein Betrag gezahlt wurde, nichts zu verarbeiten
                if (Helpers.isApproximatelyEqual(gezahlt, 0)) return;

                // MwSt-Satz normalisieren
                const mwstRate = Helpers.parseMwstRate(row[columns.mwstSatz - 1]);
                const roundedRate = Math.round(mwstRate);

                // Steuer berechnen
                const tax = Helpers.round(gezahlt * (mwstRate / 100), 2);

                // Kategorie ermitteln
                const category = row[columns.kategorie - 1]?.toString().trim() || "";

                // Verarbeitung basierend auf Zeilentyp
                if (isIncome) {
                    processIncomeRow(data, month, gezahlt, tax, roundedRate, category);
                } else if (isEigen) {
                    processEigenRow(data, month, gezahlt, tax, roundedRate, category);
                } else {
                    processExpenseRow(data, month, gezahlt, tax, roundedRate, category);
                }
            } catch (e) {
                console.error("Fehler bei der Verarbeitung einer UStVA-Zeile:", e);
            }
        };

        /**
         * Verarbeitet eine Einnahmenzeile für die UStVA
         * @param {Object} data - UStVA-Datenobjekt nach Monaten
         * @param {number} month - Monat (1-12)
         * @param {number} gezahlt - Gezahlter Betrag
         * @param {number} tax - Berechnete Steuer
         * @param {number} roundedRate - Gerundeter Steuersatz
         * @param {string} category - Kategorie
         */
        const processIncomeRow = (data, month, gezahlt, tax, roundedRate, category) => {
            // EINNAHMEN
            const catCfg = config$1.einnahmen.categories[category] ?? {};
            const taxType = catCfg.taxType ?? "steuerpflichtig";

            if (taxType === "steuerfrei_inland") {
                // Steuerfreie Einnahmen im Inland (z.B. Vermietung)
                data[month].steuerfreie_inland_einnahmen += gezahlt;
            } else if (taxType === "steuerfrei_ausland" || !roundedRate) {
                // Steuerfreie Einnahmen aus dem Ausland oder ohne MwSt
                data[month].steuerfreie_ausland_einnahmen += gezahlt;
            } else {
                // Steuerpflichtige Einnahmen
                data[month].steuerpflichtige_einnahmen += gezahlt;

                // USt nach Steuersatz addieren
                // Wir prüfen ob es ein gültiger Steuersatz ist
                if (roundedRate === 7 || roundedRate === 19) {
                    data[month][`ust_${roundedRate}`] += tax;
                }
            }
        };

        /**
         * Verarbeitet eine Eigenbeleg-Zeile für die UStVA
         * @param {Object} data - UStVA-Datenobjekt nach Monaten
         * @param {number} month - Monat (1-12)
         * @param {number} gezahlt - Gezahlter Betrag
         * @param {number} tax - Berechnete Steuer
         * @param {number} roundedRate - Gerundeter Steuersatz
         * @param {string} category - Kategorie
         */
        const processEigenRow = (data, month, gezahlt, tax, roundedRate, category) => {
            // EIGENBELEGE
            const eigenCfg = config$1.eigenbelege.categories[category] ?? {};
            const taxType = eigenCfg.taxType ?? "steuerpflichtig";

            if (taxType === "steuerfrei") {
                // Steuerfreie Eigenbelege
                data[month].eigenbelege_steuerfrei += gezahlt;
            } else if (taxType === "eigenbeleg" && eigenCfg.besonderheit === "bewirtung") {
                // Bewirtungsbelege (nur 70% der Vorsteuer absetzbar)
                data[month].eigenbelege_steuerpflichtig += gezahlt;

                if (roundedRate === 7 || roundedRate === 19) {
                    const vst70 = Helpers.round(tax * 0.7, 2); // 70% absetzbare Vorsteuer
                    const vst30 = Helpers.round(tax * 0.3, 2); // 30% nicht absetzbar
                    data[month][`vst_${roundedRate}`] += vst70;
                    data[month].nicht_abzugsfaehige_vst += vst30;
                }
            } else {
                // Normale steuerpflichtige Eigenbelege
                data[month].eigenbelege_steuerpflichtig += gezahlt;

                if (roundedRate === 7 || roundedRate === 19) {
                    data[month][`vst_${roundedRate}`] += tax;
                }
            }
        };

        /**
         * Verarbeitet eine Ausgabenzeile für die UStVA
         * @param {Object} data - UStVA-Datenobjekt nach Monaten
         * @param {number} month - Monat (1-12)
         * @param {number} gezahlt - Gezahlter Betrag
         * @param {number} tax - Berechnete Steuer
         * @param {number} roundedRate - Gerundeter Steuersatz
         * @param {string} category - Kategorie
         */
        const processExpenseRow = (data, month, gezahlt, tax, roundedRate, category) => {
            // AUSGABEN
            const catCfg = config$1.ausgaben.categories[category] ?? {};
            const taxType = catCfg.taxType ?? "steuerpflichtig";

            if (taxType === "steuerfrei_inland") {
                // Steuerfreie Ausgaben im Inland
                data[month].steuerfreie_inland_ausgaben += gezahlt;
            } else if (taxType === "steuerfrei_ausland") {
                // Steuerfreie Ausgaben im Ausland
                data[month].steuerfreie_ausland_ausgaben += gezahlt;
            } else {
                // Steuerpflichtige Ausgaben
                data[month].steuerpflichtige_ausgaben += gezahlt;

                if (roundedRate === 7 || roundedRate === 19) {
                    data[month][`vst_${roundedRate}`] += tax;
                }
            }
        };

        /**
         * Formatiert eine UStVA-Datenzeile für die Ausgabe
         *
         * @param {string} label - Bezeichnung der Zeile (z.B. Monat oder Quartal)
         * @param {Object} d - UStVA-Datenobjekt für den Zeitraum
         * @returns {Array} Formatierte Zeile für die Ausgabe
         */
        const formatUStVARow = (label, d) => {
            // Berechnung der USt-Zahlung: USt minus VSt (abzüglich nicht abzugsfähiger VSt)
            const ustZahlung = Helpers.round(
                (d.ust_7 + d.ust_19) - ((d.vst_7 + d.vst_19) - d.nicht_abzugsfaehige_vst),
                2
            );

            // Berechnung des Gesamtergebnisses: Einnahmen minus Ausgaben (ohne Steueranteil)
            const ergebnis = Helpers.round(
                (d.steuerpflichtige_einnahmen + d.steuerfreie_inland_einnahmen + d.steuerfreie_ausland_einnahmen) -
                (d.steuerpflichtige_ausgaben + d.steuerfreie_inland_ausgaben + d.steuerfreie_ausland_ausgaben +
                    d.eigenbelege_steuerpflichtig + d.eigenbelege_steuerfrei),
                2
            );

            // Formatierte Zeile zurückgeben
            return [
                label,
                d.steuerpflichtige_einnahmen,
                d.steuerfreie_inland_einnahmen,
                d.steuerfreie_ausland_einnahmen,
                d.steuerpflichtige_ausgaben,
                d.steuerfreie_inland_ausgaben,
                d.steuerfreie_ausland_ausgaben,
                d.eigenbelege_steuerpflichtig,
                d.eigenbelege_steuerfrei,
                d.nicht_abzugsfaehige_vst,
                d.ust_7,
                d.ust_19,
                d.vst_7,
                d.vst_19,
                ustZahlung,
                ergebnis
            ];
        };

        /**
         * Aggregiert UStVA-Daten für einen Zeitraum (z.B. Quartal oder Jahr)
         *
         * @param {Object} data - UStVA-Datenobjekt nach Monaten
         * @param {number} start - Startmonat (1-12)
         * @param {number} end - Endmonat (1-12)
         * @returns {Object} Aggregiertes UStVA-Datenobjekt
         */
        const aggregateUStVA = (data, start, end) => {
            const sum = createEmptyUStVA();

            for (let m = start; m <= end; m++) {
                if (!data[m]) continue; // Überspringe, falls keine Daten für den Monat

                for (const key in sum) {
                    sum[key] += data[m][key] || 0;
                }
            }

            return sum;
        };

        /**
         * Erfasst alle UStVA-Daten aus den verschiedenen Sheets
         * @returns {Object|null} UStVA-Daten nach Monaten oder null bei Fehler
         */
        const collectUStVAData = () => {
            // Prüfen, ob der Cache aktuelle Daten enthält
            const now = new Date();
            if (_cache.ustva && _cache.lastUpdated) {
                const cacheAge = now - _cache.lastUpdated;
                if (cacheAge < 5 * 60 * 1000) { // Cache ist jünger als 5 Minuten
                    return _cache.ustva;
                }
            }

            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();

                // Benötigte Sheets abrufen
                const revenueSheet = ss.getSheetByName("Einnahmen");
                const expenseSheet = ss.getSheetByName("Ausgaben");
                const eigenSheet = ss.getSheetByName("Eigenbelege");

                // Prüfen, ob die wichtigsten Sheets vorhanden sind
                if (!revenueSheet || !expenseSheet) {
                    console.error("Fehlende Blätter: 'Einnahmen' oder 'Ausgaben' nicht gefunden");
                    return null;
                }

                // Daten aus den Sheets laden
                const revenueData = revenueSheet.getDataRange().getValues();
                const expenseData = expenseSheet.getDataRange().getValues();
                const eigenData = eigenSheet ? eigenSheet.getDataRange().getValues() : [];

                // Leere UStVA-Datenstruktur für alle Monate erstellen
                const ustvaData = Object.fromEntries(
                    Array.from({length: 12}, (_, i) => [i + 1, createEmptyUStVA()])
                );

                // Daten für jede Zeile verarbeiten
                revenueData.slice(1).forEach(row => processUStVARow(row, ustvaData, true));
                expenseData.slice(1).forEach(row => processUStVARow(row, ustvaData, false));
                if (eigenData.length) {
                    eigenData.slice(1).forEach(row => processUStVARow(row, ustvaData, false, true));
                }

                // Daten cachen
                _cache.ustva = ustvaData;
                _cache.lastUpdated = now;

                return ustvaData;
            } catch (e) {
                console.error("Fehler beim Sammeln der UStVA-Daten:", e);
                return null;
            }
        };

        /**
         * Hauptfunktion zur Berechnung der UStVA
         * Sammelt Daten aus allen relevanten Sheets und erstellt ein UStVA-Sheet
         */
        const calculateUStVA = () => {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const ui = SpreadsheetApp.getUi();

                // Daten sammeln
                const ustvaData = collectUStVAData();
                if (!ustvaData) {
                    ui.alert("Die UStVA konnte nicht berechnet werden. Bitte prüfen Sie die Fehleranzeige.");
                    return;
                }

                // Ausgabe-Header erstellen
                const outputRows = [
                    [
                        "Zeitraum",
                        "Steuerpflichtige Einnahmen",
                        "Steuerfreie Inland-Einnahmen",
                        "Steuerfreie Ausland-Einnahmen",
                        "Steuerpflichtige Ausgaben",
                        "Steuerfreie Inland-Ausgaben",
                        "Steuerfreie Ausland-Ausgaben",
                        "Eigenbelege steuerpflichtig",
                        "Eigenbelege steuerfrei",
                        "Nicht abzugsfähige VSt (Bewirtung)",
                        "USt 7%",
                        "USt 19%",
                        "VSt 7%",
                        "VSt 19%",
                        "USt-Zahlung",
                        "Ergebnis"
                    ]
                ];

                // Monatliche Daten ausgeben
                config$1.common.months.forEach((name, i) => {
                    const month = i + 1;
                    outputRows.push(formatUStVARow(name, ustvaData[month]));

                    // Nach jedem Quartal eine Zusammenfassung einfügen
                    if (month % 3 === 0) {
                        const quartalsNummer = month / 3;
                        const quartalsStart = month - 2;
                        outputRows.push(formatUStVARow(
                            `Quartal ${quartalsNummer}`,
                            aggregateUStVA(ustvaData, quartalsStart, month)
                        ));
                    }
                });

                // Jahresergebnis anfügen
                outputRows.push(formatUStVARow("Gesamtjahr", aggregateUStVA(ustvaData, 1, 12)));

                // UStVA-Sheet erstellen oder aktualisieren
                let ustvaSheet = ss.getSheetByName("UStVA");
                if (!ustvaSheet) {
                    ustvaSheet = ss.insertSheet("UStVA");
                } else {
                    ustvaSheet.clearContents();
                }

                // Daten in das Sheet schreiben mit optimiertem Batch-Update
                const dataRange = ustvaSheet.getRange(1, 1, outputRows.length, outputRows[0].length);
                dataRange.setValues(outputRows);

                // Header formatieren
                ustvaSheet.getRange(1, 1, 1, outputRows[0].length).setFontWeight("bold");

                // Quartalszellen formatieren
                for (let i = 0; i < 4; i++) {
                    const row = 3 * (i + 1) + 1 + i; // Position der Quartalszeile
                    ustvaSheet.getRange(row, 1, 1, outputRows[0].length).setBackground("#e6f2ff");
                }

                // Jahreszeile formatieren
                ustvaSheet.getRange(outputRows.length, 1, 1, outputRows[0].length)
                    .setBackground("#d9e6f2")
                    .setFontWeight("bold");

                // Währungsformat für Beträge (Spalten 2-16)
                ustvaSheet.getRange(2, 2, outputRows.length - 1, 15).setNumberFormat("#,##0.00 €");

                // Spaltenbreiten anpassen
                ustvaSheet.autoResizeColumns(1, outputRows[0].length);

                // UStVA-Sheet aktivieren
                ss.setActiveSheet(ustvaSheet);

                ui.alert("UStVA wurde erfolgreich aktualisiert!");
                return true;
            } catch (e) {
                console.error("Fehler bei der UStVA-Berechnung:", e);
                SpreadsheetApp.getUi().alert("Fehler bei der UStVA-Berechnung: " + e.toString());
                return false;
            }
        };

        // Öffentliche API des Moduls
        return {
            calculateUStVA,
            clearCache,
            // Für Testzwecke könnten hier weitere Funktionen exportiert werden
            _internal: {
                createEmptyUStVA,
                processUStVARow,
                formatUStVARow,
                aggregateUStVA
            }
        };
    })();

    // file: src/bWACalculator.js

    /**
     * Modul zur Berechnung der Betriebswirtschaftlichen Auswertung (BWA)
     */
    const BWACalculator = (() => {
        // Cache für berechnete BWA-Daten
        const _cache = {
            bwaData: null,
            lastCalculated: null
        };

        /**
         * Cache zurücksetzen
         */
        const clearCache = () => {
            _cache.bwaData = null;
            _cache.lastCalculated = null;
        };

        /**
         * Erstellt ein leeres BWA-Datenobjekt mit Nullwerten
         * @returns {Object} Leere BWA-Datenstruktur
         */
        const createEmptyBWA = () => ({
            // Gruppe 1: Betriebserlöse (Einnahmen)
            umsatzerloese: 0,
            provisionserloese: 0,
            steuerfreieInlandEinnahmen: 0,
            steuerfreieAuslandEinnahmen: 0,
            sonstigeErtraege: 0,
            vermietung: 0,
            zuschuesse: 0,
            waehrungsgewinne: 0,
            anlagenabgaenge: 0,
            gesamtErloese: 0,

            // Gruppe 2: Materialaufwand & Wareneinsatz
            wareneinsatz: 0,
            fremdleistungen: 0,
            rohHilfsBetriebsstoffe: 0,
            gesamtWareneinsatz: 0,

            // Gruppe 3: Betriebsausgaben (Sachkosten)
            bruttoLoehne: 0,
            sozialeAbgaben: 0,
            sonstigePersonalkosten: 0,
            werbungMarketing: 0,
            reisekosten: 0,
            versicherungen: 0,
            telefonInternet: 0,
            buerokosten: 0,
            fortbildungskosten: 0,
            kfzKosten: 0,
            sonstigeAufwendungen: 0,
            gesamtBetriebsausgaben: 0,

            // Gruppe 4: Abschreibungen & Zinsen
            abschreibungenMaschinen: 0,
            abschreibungenBueromaterial: 0,
            abschreibungenImmateriell: 0,
            zinsenBank: 0,
            zinsenGesellschafter: 0,
            leasingkosten: 0,
            gesamtAbschreibungenZinsen: 0,

            // Gruppe 5: Besondere Posten (Kapitalbewegungen)
            eigenkapitalveraenderungen: 0,
            gesellschafterdarlehen: 0,
            ausschuettungen: 0,
            gesamtBesonderePosten: 0,

            // Gruppe 6: Rückstellungen
            steuerrueckstellungen: 0,
            rueckstellungenSonstige: 0,
            gesamtRueckstellungenTransfers: 0,

            // Gruppe 7: EBIT
            ebit: 0,

            // Gruppe 8: Steuern & Vorsteuer
            umsatzsteuer: 0,
            vorsteuer: 0,
            nichtAbzugsfaehigeVSt: 0,
            koerperschaftsteuer: 0,
            solidaritaetszuschlag: 0,
            gewerbesteuer: 0,
            gewerbesteuerRueckstellungen: 0,
            sonstigeSteuerrueckstellungen: 0,
            steuerlast: 0,

            // Gruppe 9: Jahresüberschuss/-fehlbetrag
            gewinnNachSteuern: 0,

            // Eigenbelege (zur Aggregation)
            eigenbelegeSteuerfrei: 0,
            eigenbelegeSteuerpflichtig: 0
        });

        /**
         * Verarbeitet Einnahmen und ordnet sie den BWA-Kategorien zu
         * @param {Array} row - Zeile aus dem Einnahmen-Sheet
         * @param {Object} bwaData - BWA-Datenstruktur
         */
        const processRevenue = (row, bwaData) => {
            try {
                const columns = config$1.einnahmen.columns;

                const m = Helpers.getMonthFromRow(row, "einnahmen");
                if (!m) return;

                const amount = Helpers.parseCurrency(row[columns.nettobetrag - 1]);
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
                const mapping = config$1.einnahmen.bwaMapping[category];
                if (mapping === "umsatzerloese" || mapping === "provisionserloese") {
                    bwaData[m][mapping] += amount;
                    return;
                }

                // Kategorisierung nach Steuersatz
                if (Helpers.parseMwstRate(row[columns.mwstSatz - 1]) === 0) {
                    bwaData[m].steuerfreieInlandEinnahmen += amount;
                } else {
                    bwaData[m].umsatzerloese += amount;
                }
            } catch (e) {
                console.error("Fehler bei der Verarbeitung einer Einnahme:", e);
            }
        };

        /**
         * Verarbeitet Ausgaben und ordnet sie den BWA-Kategorien zu
         * @param {Array} row - Zeile aus dem Ausgaben-Sheet
         * @param {Object} bwaData - BWA-Datenstruktur
         */
        const processExpense = (row, bwaData) => {
            try {
                const columns = config$1.ausgaben.columns;

                const m = Helpers.getMonthFromRow(row, "ausgaben");
                if (!m) return;

                const amount = Helpers.parseCurrency(row[columns.nettobetrag - 1]);
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
                const mapping = config$1.ausgaben.bwaMapping[category];
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
        };

        /**
         * Verarbeitet Eigenbelege und ordnet sie den BWA-Kategorien zu
         * @param {Array} row - Zeile aus dem Eigenbelege-Sheet
         * @param {Object} bwaData - BWA-Datenstruktur
         */
        const processEigen = (row, bwaData) => {
            try {
                const columns = config$1.eigenbelege.columns;

                const m = Helpers.getMonthFromRow(row, "eigenbelege");
                if (!m) return;

                const amount = Helpers.parseCurrency(row[columns.nettobetrag - 1]);
                if (amount === 0) return;

                const category = row[columns.kategorie - 1]?.toString().trim() || "";
                const eigenCfg = config$1.eigenbelege.categories[category] ?? {};
                const taxType = eigenCfg.taxType ?? "steuerpflichtig";

                // Basierend auf Steuertyp zuordnen
                if (taxType === "steuerfrei") {
                    bwaData[m].eigenbelegeSteuerfrei += amount;
                } else {
                    bwaData[m].eigenbelegeSteuerpflichtig += amount;
                }
            } catch (e) {
                console.error("Fehler bei der Verarbeitung eines Eigenbelegs:", e);
            }
        };

        /**
         * Berechnet Gruppensummen und abgeleitete Werte für alle Monate
         * @param {Object} bwaData - BWA-Datenstruktur mit Rohdaten
         */
        const calculateAggregates = (bwaData) => {
            for (let m = 1; m <= 12; m++) {
                const d = bwaData[m];

                // Erlöse
                d.gesamtErloese = Helpers.round(
                    d.umsatzerloese + d.provisionserloese + d.steuerfreieInlandEinnahmen +
                    d.steuerfreieAuslandEinnahmen + d.sonstigeErtraege + d.vermietung +
                    d.zuschuesse + d.waehrungsgewinne + d.anlagenabgaenge,
                    2
                );

                // Materialkosten
                d.gesamtWareneinsatz = Helpers.round(
                    d.wareneinsatz + d.fremdleistungen + d.rohHilfsBetriebsstoffe,
                    2
                );

                // Betriebsausgaben
                d.gesamtBetriebsausgaben = Helpers.round(
                    d.bruttoLoehne + d.sozialeAbgaben + d.sonstigePersonalkosten +
                    d.werbungMarketing + d.reisekosten + d.versicherungen + d.telefonInternet +
                    d.buerokosten + d.fortbildungskosten + d.kfzKosten + d.sonstigeAufwendungen,
                    2
                );

                // Abschreibungen & Zinsen
                d.gesamtAbschreibungenZinsen = Helpers.round(
                    d.abschreibungenMaschinen + d.abschreibungenBueromaterial +
                    d.abschreibungenImmateriell + d.zinsenBank + d.zinsenGesellschafter +
                    d.leasingkosten,
                    2
                );

                // Besondere Posten
                d.gesamtBesonderePosten = Helpers.round(
                    d.eigenkapitalveraenderungen + d.gesellschafterdarlehen + d.ausschuettungen,
                    2
                );

                // Rückstellungen
                d.gesamtRueckstellungenTransfers = Helpers.round(
                    d.steuerrueckstellungen + d.rueckstellungenSonstige,
                    2
                );

                // EBIT
                d.ebit = Helpers.round(
                    d.gesamtErloese - (d.gesamtWareneinsatz + d.gesamtBetriebsausgaben +
                        d.gesamtAbschreibungenZinsen + d.gesamtBesonderePosten),
                    2
                );

                // Steuern berechnen
                const taxConfig = config$1.tax.isHolding ? config$1.tax.holding : config$1.tax.operative;

                // Für Holdings gelten spezielle Steuersätze wegen Beteiligungsprivileg
                const steuerfaktor = config$1.tax.isHolding
                    ? taxConfig.gewinnUebertragSteuerpflichtig / 100
                    : 1;

                d.gewerbesteuer = Helpers.round(d.ebit * (taxConfig.gewerbesteuer / 10000) * steuerfaktor, 2);
                d.koerperschaftsteuer = Helpers.round(d.ebit * (taxConfig.koerperschaftsteuer / 100) * steuerfaktor, 2);
                d.solidaritaetszuschlag = Helpers.round(d.koerperschaftsteuer * (taxConfig.solidaritaetszuschlag / 100), 2);

                // Gesamte Steuerlast
                d.steuerlast = Helpers.round(
                    d.koerperschaftsteuer + d.solidaritaetszuschlag + d.gewerbesteuer,
                    2
                );

                // Gewinn nach Steuern
                d.gewinnNachSteuern = Helpers.round(d.ebit - d.steuerlast, 2);
            }
        };

        /**
         * Sammelt alle BWA-Daten aus den verschiedenen Sheets
         * @returns {Object|null} BWA-Datenstruktur oder null bei Fehler
         */
        const aggregateBWAData = () => {
            try {
                // Prüfen ob Cache gültig ist (maximal 5 Minuten alt)
                const now = new Date();
                if (_cache.bwaData && _cache.lastCalculated) {
                    const cacheAge = now - _cache.lastCalculated;
                    if (cacheAge < 5 * 60 * 1000) { // 5 Minuten
                        return _cache.bwaData;
                    }
                }

                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const revenueSheet = ss.getSheetByName("Einnahmen");
                const expenseSheet = ss.getSheetByName("Ausgaben");
                const eigenSheet = ss.getSheetByName("Eigenbelege");

                if (!revenueSheet || !expenseSheet) {
                    console.error("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
                    return null;
                }

                // BWA-Daten für alle Monate initialisieren
                const bwaData = Object.fromEntries(Array.from({length: 12}, (_, i) => [i + 1, createEmptyBWA()]));

                // Daten aus den Sheets verarbeiten
                revenueSheet.getDataRange().getValues().slice(1).forEach(row => processRevenue(row, bwaData));
                expenseSheet.getDataRange().getValues().slice(1).forEach(row => processExpense(row, bwaData));
                if (eigenSheet) eigenSheet.getDataRange().getValues().slice(1).forEach(row => processEigen(row, bwaData));

                // Gruppensummen und weitere Berechnungen
                calculateAggregates(bwaData);

                // Daten im Cache speichern
                _cache.bwaData = bwaData;
                _cache.lastCalculated = now;

                return bwaData;
            } catch (e) {
                console.error("Fehler bei der Aggregation der BWA-Daten:", e);
                return null;
            }
        };

        /**
         * Erstellt den Header für die BWA mit Monats- und Quartalsspalten
         * @returns {Array} Header-Zeile
         */
        const buildHeaderRow = () => {
            const headers = ["Kategorie"];
            for (let q = 0; q < 4; q++) {
                for (let m = q * 3; m < q * 3 + 3; m++) {
                    headers.push(`${config$1.common.months[m]} (€)`);
                }
                headers.push(`Q${q + 1} (€)`);
            }
            headers.push("Jahr (€)");
            return headers;
        };

        /**
         * Erstellt eine Ausgabezeile für eine Position
         * @param {Object} pos - Position mit Label und Wert-Funktion
         * @param {Object} bwaData - BWA-Daten
         * @returns {Array} Formatierte Zeile
         */
        const buildOutputRow = (pos, bwaData) => {
            const monthly = [];
            let yearly = 0;

            // Monatswerte berechnen
            for (let m = 1; m <= 12; m++) {
                const val = pos.get(bwaData[m]) || 0;
                monthly.push(val);
                yearly += val;
            }

            // Quartalswerte berechnen
            const quarters = [0, 0, 0, 0];
            for (let i = 0; i < 12; i++) {
                quarters[Math.floor(i / 3)] += monthly[i];
            }

            // Runde alle Werte für bessere Darstellung
            const roundedMonthly = monthly.map(val => Helpers.round(val, 2));
            const roundedQuarters = quarters.map(val => Helpers.round(val, 2));
            const roundedYearly = Helpers.round(yearly, 2);

            // Zeile zusammenstellen
            return [pos.label,
                ...roundedMonthly.slice(0, 3), roundedQuarters[0],
                ...roundedMonthly.slice(3, 6), roundedQuarters[1],
                ...roundedMonthly.slice(6, 9), roundedQuarters[2],
                ...roundedMonthly.slice(9, 12), roundedQuarters[3],
                roundedYearly];
        };

        /**
         * Hauptfunktion zur Berechnung der BWA
         */
        const calculateBWA = () => {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const bwaData = aggregateBWAData();
                if (!bwaData) {
                    SpreadsheetApp.getUi().alert("Fehler: BWA-Daten konnten nicht generiert werden.");
                    return;
                }

                // Positionen für die BWA definieren
                const positions = [
                    {label: "Erlöse aus Lieferungen und Leistungen", get: d => d.umsatzerloese || 0},
                    {label: "Provisionserlöse", get: d => d.provisionserloese || 0},
                    {label: "Steuerfreie Inland-Einnahmen", get: d => d.steuerfreieInlandEinnahmen || 0},
                    {label: "Steuerfreie Ausland-Einnahmen", get: d => d.steuerfreieAuslandEinnahmen || 0},
                    {label: "Sonstige betriebliche Erträge", get: d => d.sonstigeErtraege || 0},
                    {label: "Erträge aus Vermietung/Verpachtung", get: d => d.vermietung || 0},
                    {label: "Erträge aus Zuschüssen", get: d => d.zuschuesse || 0},
                    {label: "Erträge aus Währungsgewinnen", get: d => d.waehrungsgewinne || 0},
                    {label: "Erträge aus Anlagenabgängen", get: d => d.anlagenabgaenge || 0},
                    {label: "Betriebserlöse", get: d => d.gesamtErloese || 0},
                    {label: "Wareneinsatz", get: d => d.wareneinsatz || 0},
                    {label: "Bezogene Leistungen", get: d => d.fremdleistungen || 0},
                    {label: "Roh-, Hilfs- & Betriebsstoffe", get: d => d.rohHilfsBetriebsstoffe || 0},
                    {label: "Gesamtkosten Material & Fremdleistungen", get: d => d.gesamtWareneinsatz || 0},
                    {label: "Bruttolöhne & Gehälter", get: d => d.bruttoLoehne || 0},
                    {label: "Soziale Abgaben & Arbeitgeberanteile", get: d => d.sozialeAbgaben || 0},
                    {label: "Sonstige Personalkosten", get: d => d.sonstigePersonalkosten || 0},
                    {label: "Werbung & Marketing", get: d => d.werbungMarketing || 0},
                    {label: "Reisekosten", get: d => d.reisekosten || 0},
                    {label: "Versicherungen", get: d => d.versicherungen || 0},
                    {label: "Telefon & Internet", get: d => d.telefonInternet || 0},
                    {label: "Bürokosten", get: d => d.buerokosten || 0},
                    {label: "Fortbildungskosten", get: d => d.fortbildungskosten || 0},
                    {label: "Kfz-Kosten", get: d => d.kfzKosten || 0},
                    {label: "Sonstige betriebliche Aufwendungen", get: d => d.sonstigeAufwendungen || 0},
                    {label: "Abschreibungen Maschinen", get: d => d.abschreibungenMaschinen || 0},
                    {label: "Abschreibungen Büroausstattung", get: d => d.abschreibungenBueromaterial || 0},
                    {label: "Abschreibungen immaterielle Wirtschaftsgüter", get: d => d.abschreibungenImmateriell || 0},
                    {label: "Zinsen auf Bankdarlehen", get: d => d.zinsenBank || 0},
                    {label: "Zinsen auf Gesellschafterdarlehen", get: d => d.zinsenGesellschafter || 0},
                    {label: "Leasingkosten", get: d => d.leasingkosten || 0},
                    {label: "Gesamt Abschreibungen & Zinsen", get: d => d.gesamtAbschreibungenZinsen || 0},
                    {label: "Eigenkapitalveränderungen", get: d => d.eigenkapitalveraenderungen || 0},
                    {label: "Gesellschafterdarlehen", get: d => d.gesellschafterdarlehen || 0},
                    {label: "Ausschüttungen an Gesellschafter", get: d => d.ausschuettungen || 0},
                    {label: "Steuerrückstellungen", get: d => d.steuerrueckstellungen || 0},
                    {label: "Rückstellungen sonstige", get: d => d.rueckstellungenSonstige || 0},
                    {label: "Betriebsergebnis vor Steuern (EBIT)", get: d => d.ebit || 0},
                    {label: "Umsatzsteuer (abzuführen)", get: d => d.umsatzsteuer || 0},
                    {label: "Vorsteuer", get: d => d.vorsteuer || 0},
                    {label: "Nicht abzugsfähige VSt (Bewirtung)", get: d => d.nichtAbzugsfaehigeVSt || 0},
                    {label: "Körperschaftsteuer", get: d => d.koerperschaftsteuer || 0},
                    {label: "Solidaritätszuschlag", get: d => d.solidaritaetszuschlag || 0},
                    {label: "Gewerbesteuer", get: d => d.gewerbesteuer || 0},
                    {label: "Gesamtsteueraufwand", get: d => d.steuerlast || 0},
                    {label: "Jahresüberschuss/-fehlbetrag", get: d => d.gewinnNachSteuern || 0}
                ];

                // Header-Zeile erstellen
                const headerRow = buildHeaderRow();
                const outputRows = [headerRow];

                // Gruppenhierarchie für BWA
                const bwaGruppen = [
                    {titel: "Betriebserlöse (Einnahmen)", count: 10},
                    {titel: "Materialaufwand & Wareneinsatz", count: 4},
                    {titel: "Betriebsausgaben (Sachkosten)", count: 11},
                    {titel: "Abschreibungen & Zinsen", count: 7},
                    {titel: "Besondere Posten", count: 3},
                    {titel: "Rückstellungen", count: 2},
                    {titel: "Betriebsergebnis vor Steuern (EBIT)", count: 1},
                    {titel: "Steuern & Vorsteuer", count: 7},
                    {titel: "Jahresüberschuss/-fehlbetrag", count: 1}
                ];

                // Ausgabe mit Gruppenhierarchie erstellen
                let posIndex = 0;
                for (let gruppenIndex = 0; gruppenIndex < bwaGruppen.length; gruppenIndex++) {
                    const gruppe = bwaGruppen[gruppenIndex];

                    // Gruppenüberschrift
                    outputRows.push([
                        `${gruppenIndex + 1}. ${gruppe.titel}`,
                        ...Array(headerRow.length - 1).fill("")
                    ]);

                    // Gruppenpositionen
                    for (let i = 0; i < gruppe.count; i++) {
                        outputRows.push(buildOutputRow(positions[posIndex++], bwaData));
                    }

                    // Leerzeile nach jeder Gruppe außer der letzten
                    if (gruppenIndex < bwaGruppen.length - 1) {
                        outputRows.push(Array(headerRow.length).fill(""));
                    }
                }

                // BWA-Sheet erstellen oder aktualisieren
                let bwaSheet = ss.getSheetByName("BWA");
                if (!bwaSheet) {
                    bwaSheet = ss.insertSheet("BWA");
                } else {
                    bwaSheet.clearContents();
                }

                // Daten in das Sheet schreiben (als Batch für Performance)
                const dataRange = bwaSheet.getRange(1, 1, outputRows.length, outputRows[0].length);
                dataRange.setValues(outputRows);

                // Formatierungen anwenden
                applyBwaFormatting(bwaSheet, headerRow.length, bwaGruppen, outputRows.length);

                // Erfolgsbenachrichtigung
                SpreadsheetApp.getUi().alert("BWA wurde aktualisiert!");

                // BWA-Sheet aktivieren
                ss.setActiveSheet(bwaSheet);

                return true;

            } catch (e) {
                console.error("Fehler bei der BWA-Berechnung:", e);
                SpreadsheetApp.getUi().alert("Fehler bei der BWA-Berechnung: " + e.toString());
                return false;
            }
        };

        /**
         * Wendet Formatierungen auf das BWA-Sheet an
         * @param {Sheet} sheet - Das zu formatierende Sheet
         * @param {number} headerLength - Anzahl der Spalten
         * @param {Array} bwaGruppen - Gruppenhierarchie
         * @param {number} totalRows - Gesamtzahl der Zeilen
         */
        const applyBwaFormatting = (sheet, headerLength, bwaGruppen, totalRows) => {
            // Header formatieren
            sheet.getRange(1, 1, 1, headerLength).setFontWeight("bold").setBackground("#f3f3f3");

            // Gruppenüberschriften formatieren
            let rowIndex = 2;
            for (const gruppe of bwaGruppen) {
                sheet.getRange(rowIndex, 1).setFontWeight("bold");
                rowIndex += gruppe.count + 1; // +1 für die Leerzeile
            }

            // Währungsformat für alle Zahlenwerte
            sheet.getRange(2, 2, totalRows - 1, headerLength - 1).setNumberFormat("#,##0.00 €");

            // Summen-Zeilen hervorheben
            const summenZeilen = [11, 15, 26, 33, 36, 38, 39, 46];
            summenZeilen.forEach(row => {
                if (row <= totalRows) {
                    sheet.getRange(row, 1, 1, headerLength).setBackground("#e6f2ff");
                }
            });

            // EBIT und Jahresüberschuss hervorheben
            sheet.getRange(39, 1, 1, headerLength).setFontWeight("bold");
            sheet.getRange(46, 1, 1, headerLength).setFontWeight("bold");

            // Spaltenbreiten anpassen
            sheet.autoResizeColumns(1, headerLength);
        };

        // Öffentliche API des Moduls
        return {
            calculateBWA,
            clearCache,
            // Für Testzwecke könnten hier weitere Funktionen exportiert werden
            _internal: {
                createEmptyBWA,
                processRevenue,
                processExpense,
                processEigen,
                aggregateBWAData
            }
        };
    })();

    // file: src/bilanzCalculator.js

    /**
     * Modul zur Erstellung einer Bilanz nach SKR04
     * Erstellt eine standardkonforme Bilanz basierend auf den Daten aus anderen Sheets
     */
    const BilanzCalculator = (() => {
        // Cache für Bilanzdaten
        const _cache = {
            bilanzData: null,
            lastCalculated: null
        };

        /**
         * Cache leeren
         */
        const clearCache = () => {
            _cache.bilanzData = null;
            _cache.lastCalculated = null;
        };

        /**
         * Erstellt eine leere Bilanz-Datenstruktur
         * @returns {Object} Leere Bilanz-Datenstruktur
         */
        const createEmptyBilanz = () => ({
            // Aktiva (Vermögenswerte)
            aktiva: {
                // Anlagevermögen
                sachanlagen: 0,                 // SKR04: 0400-0699, Sachanlagen
                immaterielleVermoegen: 0,       // SKR04: 0100-0199, Immaterielle Vermögensgegenstände
                finanzanlagen: 0,               // SKR04: 0700-0899, Finanzanlagen
                summeAnlagevermoegen: 0,        // Summe Anlagevermögen

                // Umlaufvermögen
                bankguthaben: 0,                // SKR04: 1200, Bank
                kasse: 0,                       // SKR04: 1210, Kasse
                forderungenLuL: 0,              // SKR04: 1300-1370, Forderungen aus Lieferungen und Leistungen
                vorraete: 0,                    // SKR04: 1400-1590, Vorräte
                summeUmlaufvermoegen: 0,        // Summe Umlaufvermögen

                // Rechnungsabgrenzung
                rechnungsabgrenzung: 0,         // SKR04: 1900-1990, Aktiver Rechnungsabgrenzungsposten

                // Gesamtsumme
                summeAktiva: 0                  // Summe aller Aktiva
            },

            // Passiva (Kapital und Schulden)
            passiva: {
                // Eigenkapital
                stammkapital: 0,                // SKR04: 2000, Gezeichnetes Kapital
                kapitalruecklagen: 0,           // SKR04: 2100, Kapitalrücklage
                gewinnvortrag: 0,               // SKR04: 2970, Gewinnvortrag
                verlustvortrag: 0,              // SKR04: 2978, Verlustvortrag
                jahresueberschuss: 0,           // Jahresüberschuss aus BWA
                summeEigenkapital: 0,           // Summe Eigenkapital

                // Verbindlichkeiten
                bankdarlehen: 0,                // SKR04: 3150, Verbindlichkeiten gegenüber Kreditinstituten
                gesellschafterdarlehen: 0,      // SKR04: 3300, Verbindlichkeiten gegenüber Gesellschaftern
                verbindlichkeitenLuL: 0,        // SKR04: 3200, Verbindlichkeiten aus Lieferungen und Leistungen
                steuerrueckstellungen: 0,       // SKR04: 3060, Steuerrückstellungen
                summeVerbindlichkeiten: 0,      // Summe Verbindlichkeiten

                // Rechnungsabgrenzung
                rechnungsabgrenzung: 0,         // SKR04: 3800-3990, Passiver Rechnungsabgrenzungsposten

                // Gesamtsumme
                summePassiva: 0                 // Summe aller Passiva
            }
        });

        /**
         * Sammelt Daten aus verschiedenen Sheets für die Bilanz
         * @returns {Object} Bilanz-Datenstruktur mit befüllten Werten
         */
        const aggregateBilanzData = () => {
            try {
                // Prüfen ob Cache gültig ist (maximal 5 Minuten alt)
                const now = new Date();
                if (_cache.bilanzData && _cache.lastCalculated) {
                    const cacheAge = now - _cache.lastCalculated;
                    if (cacheAge < 5 * 60 * 1000) { // 5 Minuten
                        return _cache.bilanzData;
                    }
                }

                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const bilanzData = createEmptyBilanz();

                // Spalten-Konfigurationen für die verschiedenen Sheets abrufen
                const bankCols = config$1.bankbewegungen.columns;
                const ausgabenCols = config$1.ausgaben.columns;
                const gesellschafterCols = config$1.gesellschafterkonto.columns;

                // 1. Banksaldo aus "Bankbewegungen" (Endsaldo)
                const bankSheet = ss.getSheetByName("Bankbewegungen");
                if (bankSheet) {
                    const lastRow = bankSheet.getLastRow();
                    if (lastRow >= 1) {
                        const label = bankSheet.getRange(lastRow, bankCols.buchungstext).getValue().toString().toLowerCase();
                        if (label === "endsaldo") {
                            bilanzData.aktiva.bankguthaben = Helpers.parseCurrency(
                                bankSheet.getRange(lastRow, bankCols.saldo).getValue()
                            );
                        }
                    }
                }

                // 2. Jahresüberschuss aus "BWA" (Letzte Zeile, sofern dort "Jahresüberschuss" vorkommt)
                const bwaSheet = ss.getSheetByName("BWA");
                if (bwaSheet) {
                    const data = bwaSheet.getDataRange().getValues();
                    for (let i = data.length - 1; i >= 0; i--) {
                        const row = data[i];
                        if (row[0].toString().toLowerCase().includes("jahresüberschuss")) {
                            // Letzte Spalte enthält den Jahreswert
                            bilanzData.passiva.jahresueberschuss = Helpers.parseCurrency(row[row.length - 1]);
                            break;
                        }
                    }
                }

                // 3. Stammkapital aus Konfiguration
                bilanzData.passiva.stammkapital = config$1.tax.stammkapital || 25000;

                // 4. Gesellschafterdarlehen aus dem Gesellschafterkonto-Sheet
                bilanzData.passiva.gesellschafterdarlehen = getDarlehensumme(ss, gesellschafterCols);

                // 5. Steuerrückstellungen aus dem Ausgaben-Sheet
                bilanzData.passiva.steuerrueckstellungen = getSteuerRueckstellungen(ss, ausgabenCols);

                // 6. Berechnung der Summen
                calculateBilanzSummen(bilanzData);

                // Daten im Cache speichern
                _cache.bilanzData = bilanzData;
                _cache.lastCalculated = now;

                return bilanzData;
            } catch (e) {
                console.error("Fehler bei der Sammlung der Bilanzdaten:", e);
                return null;
            }
        };

        /**
         * Ermittelt die Summe der Gesellschafterdarlehen
         * @param {Spreadsheet} ss - Das Spreadsheet
         * @param {Object} gesellschafterCols - Spaltenkonfiguration
         * @returns {number} Summe der Gesellschafterdarlehen
         */
        const getDarlehensumme = (ss, gesellschafterCols) => {
            let darlehenSumme = 0;

            const gesellschafterSheet = ss.getSheetByName("Gesellschafterkonto");
            if (gesellschafterSheet) {
                const data = gesellschafterSheet.getDataRange().getValues();

                // Überschrift überspringen
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    // Prüfen, ob es sich um ein Gesellschafterdarlehen handelt
                    if (row[gesellschafterCols.kategorie - 1] &&
                        row[gesellschafterCols.kategorie - 1].toString().toLowerCase() === "gesellschafterdarlehen") {
                        darlehenSumme += Helpers.parseCurrency(row[gesellschafterCols.betrag - 1] || 0);
                    }
                }
            }

            return darlehenSumme;
        };

        /**
         * Ermittelt die Summe der Steuerrückstellungen
         * @param {Spreadsheet} ss - Das Spreadsheet
         * @param {Object} ausgabenCols - Spaltenkonfiguration
         * @returns {number} Summe der Steuerrückstellungen
         */
        const getSteuerRueckstellungen = (ss, ausgabenCols) => {
            let steuerRueckstellungen = 0;

            const ausSheet = ss.getSheetByName("Ausgaben");
            if (ausSheet) {
                const data = ausSheet.getDataRange().getValues();

                // Array mit Steuerrückstellungs-Kategorien
                const steuerKategorien = [
                    "Gewerbesteuerrückstellungen",
                    "Körperschaftsteuer",
                    "Solidaritätszuschlag",
                    "Sonstige Steuerrückstellungen"
                ];

                // Überschrift überspringen
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const category = row[ausgabenCols.kategorie - 1]?.toString().trim() || "";

                    if (steuerKategorien.includes(category)) {
                        steuerRueckstellungen += Helpers.parseCurrency(row[ausgabenCols.nettobetrag - 1] || 0);
                    }
                }
            }

            return steuerRueckstellungen;
        };

        /**
         * Berechnet die Summen für die Bilanz
         * @param {Object} bilanzData - Die Bilanzdaten
         */
        const calculateBilanzSummen = (bilanzData) => {
            const { aktiva, passiva } = bilanzData;

            // Summen für Aktiva
            aktiva.summeAnlagevermoegen = Helpers.round(
                aktiva.sachanlagen +
                aktiva.immaterielleVermoegen +
                aktiva.finanzanlagen,
                2
            );

            aktiva.summeUmlaufvermoegen = Helpers.round(
                aktiva.bankguthaben +
                aktiva.kasse +
                aktiva.forderungenLuL +
                aktiva.vorraete,
                2
            );

            aktiva.summeAktiva = Helpers.round(
                aktiva.summeAnlagevermoegen +
                aktiva.summeUmlaufvermoegen +
                aktiva.rechnungsabgrenzung,
                2
            );

            // Summen für Passiva
            passiva.summeEigenkapital = Helpers.round(
                passiva.stammkapital +
                passiva.kapitalruecklagen +
                passiva.gewinnvortrag -
                passiva.verlustvortrag +
                passiva.jahresueberschuss,
                2
            );

            passiva.summeVerbindlichkeiten = Helpers.round(
                passiva.bankdarlehen +
                passiva.gesellschafterdarlehen +
                passiva.verbindlichkeitenLuL +
                passiva.steuerrueckstellungen,
                2
            );

            passiva.summePassiva = Helpers.round(
                passiva.summeEigenkapital +
                passiva.summeVerbindlichkeiten +
                passiva.rechnungsabgrenzung,
                2
            );
        };

        /**
         * Konvertiert einen Zahlenwert in eine Zellenformel oder direkte Zahl
         * @param {number} value - Der Wert
         * @param {boolean} useFormula - Ob eine Formel verwendet werden soll
         * @returns {string|number} - Formel als String oder direkter Wert
         */
        const valueOrFormula = (value, useFormula = false) => {
            if (Helpers.isApproximatelyEqual(value, 0) && useFormula) {
                return "";  // Leere Zelle für 0-Werte bei Formeln
            }
            return value;
        };

        /**
         * Erstellt ein Array für die Aktiva-Seite der Bilanz
         * @param {Object} bilanzData - Die Bilanzdaten
         * @returns {Array} Array mit Zeilen für die Aktiva-Seite
         */
        const createAktivaArray = (bilanzData) => {
            const { aktiva } = bilanzData;
            const year = config$1.tax.year || new Date().getFullYear();

            return [
                [`Bilanz ${year} - Aktiva (Vermögenswerte)`, ""],
                ["", ""],
                ["1. Anlagevermögen", ""],
                ["1.1 Sachanlagen", valueOrFormula(aktiva.sachanlagen)],
                ["1.2 Immaterielle Vermögensgegenstände", valueOrFormula(aktiva.immaterielleVermoegen)],
                ["1.3 Finanzanlagen", valueOrFormula(aktiva.finanzanlagen)],
                ["Summe Anlagevermögen", "=SUM(B4:B6)"],
                ["", ""],
                ["2. Umlaufvermögen", ""],
                ["2.1 Bankguthaben", valueOrFormula(aktiva.bankguthaben)],
                ["2.2 Kasse", valueOrFormula(aktiva.kasse)],
                ["2.3 Forderungen aus Lieferungen und Leistungen", valueOrFormula(aktiva.forderungenLuL)],
                ["2.4 Vorräte", valueOrFormula(aktiva.vorraete)],
                ["Summe Umlaufvermögen", "=SUM(B10:B13)"],
                ["", ""],
                ["3. Rechnungsabgrenzungsposten", valueOrFormula(aktiva.rechnungsabgrenzung)],
                ["", ""],
                ["Summe Aktiva", "=B7+B14+B16"]
            ];
        };

        /**
         * Erstellt ein Array für die Passiva-Seite der Bilanz
         * @param {Object} bilanzData - Die Bilanzdaten
         * @returns {Array} Array mit Zeilen für die Passiva-Seite
         */
        const createPassivaArray = (bilanzData) => {
            const { passiva } = bilanzData;
            const year = config$1.tax.year || new Date().getFullYear();

            return [
                [`Bilanz ${year} - Passiva (Kapital und Schulden)`, ""],
                ["", ""],
                ["1. Eigenkapital", ""],
                ["1.1 Gezeichnetes Kapital (Stammkapital)", valueOrFormula(passiva.stammkapital)],
                ["1.2 Kapitalrücklage", valueOrFormula(passiva.kapitalruecklagen)],
                ["1.3 Gewinnvortrag", valueOrFormula(passiva.gewinnvortrag)],
                ["1.4 Verlustvortrag (negativ)", valueOrFormula(passiva.verlustvortrag)],
                ["1.5 Jahresüberschuss/Jahresfehlbetrag", valueOrFormula(passiva.jahresueberschuss)],
                ["Summe Eigenkapital", "=SUM(F4:F8)"],
                ["", ""],
                ["2. Verbindlichkeiten", ""],
                ["2.1 Verbindlichkeiten gegenüber Kreditinstituten", valueOrFormula(passiva.bankdarlehen)],
                ["2.2 Verbindlichkeiten gegenüber Gesellschaftern", valueOrFormula(passiva.gesellschafterdarlehen)],
                ["2.3 Verbindlichkeiten aus Lieferungen und Leistungen", valueOrFormula(passiva.verbindlichkeitenLuL)],
                ["2.4 Steuerrückstellungen", valueOrFormula(passiva.steuerrueckstellungen)],
                ["Summe Verbindlichkeiten", "=SUM(F12:F15)"],
                ["", ""],
                ["3. Rechnungsabgrenzungsposten", valueOrFormula(passiva.rechnungsabgrenzung)],
                ["", ""],
                ["Summe Passiva", "=F9+F16+F18"]
            ];
        };

        /**
         * Formatiert das Bilanz-Sheet
         * @param {Sheet} bilanzSheet - Das zu formatierende Sheet
         * @param {number} aktivaLength - Anzahl der Aktiva-Zeilen
         * @param {number} passivaLength - Anzahl der Passiva-Zeilen
         */
        const formatBilanzSheet = (bilanzSheet, aktivaLength, passivaLength) => {
            // Überschriften formatieren
            bilanzSheet.getRange("A1").setFontWeight("bold").setFontSize(12);
            bilanzSheet.getRange("E1").setFontWeight("bold").setFontSize(12);

            // Zwischensummen und Gesamtsummen formatieren
            const summenZeilenAktiva = [7, 14, 18]; // Zeilen mit Summen in Aktiva
            const summenZeilenPassiva = [9, 16, 20]; // Zeilen mit Summen in Passiva

            summenZeilenAktiva.forEach(row => {
                if (row <= aktivaLength) {
                    bilanzSheet.getRange(row, 1, 1, 2).setFontWeight("bold");
                    if (row === 18) { // Gesamtsumme Aktiva
                        bilanzSheet.getRange(row, 1, 1, 2).setBackground("#e6f2ff");
                    } else {
                        bilanzSheet.getRange(row, 1, 1, 2).setBackground("#f0f0f0");
                    }
                }
            });

            summenZeilenPassiva.forEach(row => {
                if (row <= passivaLength) {
                    bilanzSheet.getRange(row, 5, 1, 2).setFontWeight("bold");
                    if (row === 20) { // Gesamtsumme Passiva
                        bilanzSheet.getRange(row, 5, 1, 2).setBackground("#e6f2ff");
                    } else {
                        bilanzSheet.getRange(row, 5, 1, 2).setBackground("#f0f0f0");
                    }
                }
            });

            // Abschnittsüberschriften formatieren
            [3, 9, 16].forEach(row => {
                if (row <= aktivaLength) {
                    bilanzSheet.getRange(row, 1).setFontWeight("bold");
                }
            });

            [3, 11, 17].forEach(row => {
                if (row <= passivaLength) {
                    bilanzSheet.getRange(row, 5).setFontWeight("bold");
                }
            });

            // Währungsformat für Beträge anwenden
            bilanzSheet.getRange("B4:B18").setNumberFormat("#,##0.00 €");
            bilanzSheet.getRange("F4:F20").setNumberFormat("#,##0.00 €");

            // Spaltenbreiten anpassen
            bilanzSheet.autoResizeColumns(1, 6);
        };

        /**
         * Hauptfunktion zur Erstellung der Bilanz
         * Sammelt Daten und erstellt ein Bilanz-Sheet
         */
        const calculateBilanz = () => {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const ui = SpreadsheetApp.getUi();

                // Bilanzdaten aggregieren
                const bilanzData = aggregateBilanzData();
                if (!bilanzData) {
                    ui.alert("Fehler: Bilanzdaten konnten nicht gesammelt werden.");
                    return false;
                }

                // Bilanz-Arrays erstellen
                const aktivaArray = createAktivaArray(bilanzData);
                const passivaArray = createPassivaArray(bilanzData);

                // Prüfen, ob Aktiva = Passiva
                const aktivaSumme = bilanzData.aktiva.summeAktiva;
                const passivaSumme = bilanzData.passiva.summePassiva;
                const differenz = Math.abs(aktivaSumme - passivaSumme);

                if (differenz > 0.01) {
                    // Bei Differenz die Bilanz trotzdem erstellen, aber warnen
                    ui.alert(
                        "Bilanz ist nicht ausgeglichen",
                        `Die Bilanzsummen von Aktiva (${Helpers.formatCurrency(aktivaSumme)}) und Passiva (${Helpers.formatCurrency(passivaSumme)}) ` +
                        `stimmen nicht überein. Differenz: ${Helpers.formatCurrency(differenz)}. ` +
                        `Bitte überprüfen Sie Ihre Buchhaltungsdaten.`,
                        ui.ButtonSet.OK
                    );
                }

                // Erstelle oder leere das Blatt "Bilanz"
                let bilanzSheet = ss.getSheetByName("Bilanz");
                if (!bilanzSheet) {
                    bilanzSheet = ss.insertSheet("Bilanz");
                } else {
                    bilanzSheet.clearContents();
                }

                // Batch-Write statt einzelner Zellen-Updates
                // Schreibe Aktiva ab Zelle A1 und Passiva ab Zelle E1
                bilanzSheet.getRange(1, 1, aktivaArray.length, 2).setValues(aktivaArray);
                bilanzSheet.getRange(1, 5, passivaArray.length, 2).setValues(passivaArray);

                // Formatierung anwenden
                formatBilanzSheet(bilanzSheet, aktivaArray.length, passivaArray.length);

                // Erfolgsmeldung
                ui.alert("Die Bilanz wurde erfolgreich erstellt!");
                return true;
            } catch (e) {
                console.error("Fehler bei der Bilanzerstellung:", e);
                SpreadsheetApp.getUi().alert("Fehler bei der Bilanzerstellung: " + e.toString());
                return false;
            }
        };

        // Öffentliche API des Moduls
        return {
            calculateBilanz,
            clearCache,
            // Für Testzwecke könnten hier weitere Funktionen exportiert werden
            _internal: {
                createEmptyBilanz,
                aggregateBilanzData
            }
        };
    })();

    // file: src/code.js
    // imports

    // =================== Globale Funktionen ===================
    /**
     * Erstellt das Menü in der Google Sheets UI beim Öffnen der Tabelle
     */
    const onOpen = () => {
        SpreadsheetApp.getUi()
            .createMenu("📂 Buchhaltung")
            .addItem("📥 Dateien importieren", "importDriveFiles")
            .addItem("🔄 Aktuelles Blatt aktualisieren", "refreshSheet")
            .addItem("📊 UStVA berechnen", "calculateUStVA")
            .addItem("📈 BWA berechnen", "calculateBWA")
            .addItem("📝 Bilanz erstellen", "calculateBilanz")
            .addToUi();
    };

    /**
     * Wird bei jeder Bearbeitung des Spreadsheets ausgelöst
     * Fügt Zeitstempel hinzu, wenn bestimmte Blätter bearbeitet werden
     *
     * @param {Object} e - Event-Objekt von Google Sheets
     */
    const onEdit = e => {
        const {range} = e;
        const sheet = range.getSheet();
        const name = sheet.getName();

        // Header-Zeile überspringen (Zeile 1)
        if (range.getRow() <= 1) return;

        // Konvertieren in kleinbuchstaben und leerzeichen entfernen
        let sheetKey = name.toLowerCase().replace(/\s+/g, '');

        // Nach entsprechendem Schlüssel in config suchen (case-insensitive)
        let configKey = null;
        Object.keys(config$1).forEach(key => {
            if (key.toLowerCase() === sheetKey) {
                configKey = key;
            }
        });

        // Wenn kein passender Schlüssel gefunden wurde, abbrechen
        if (!configKey || !config$1[configKey].columns.zeitstempel) return;

        // Spalte für Zeitstempel aus der Konfiguration
        const timestampCol = config$1[configKey].columns.zeitstempel;

        // Prüfen, ob die bearbeitete Spalte bereits die Zeitstempel-Spalte ist
        if (range.getColumn() === timestampCol ||
            (range.getNumColumns() > 1 && range.getColumn() <= timestampCol &&
                range.getColumn() + range.getNumColumns() > timestampCol)) {
            return; // Vermeide Endlosschleife
        }

        // Multi-Zellen-Editierung unterstützen
        const numRows = range.getNumRows();
        const timestampValues = [];
        const now = new Date();

        // Für jede bearbeitete Zeile
        for (let i = 0; i < numRows; i++) {
            const rowIndex = range.getRow() + i;

            // Header-Zeile überspringen
            if (rowIndex <= 1) continue;

            const headerLen = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].length;

            // Prüfen, ob die Zeile leer ist
            const rowValues = sheet.getRange(rowIndex, 1, 1, headerLen).getValues()[0];
            if (rowValues.every(cell => cell === "")) continue;

            // Zeitstempel für diese Zeile hinzufügen
            timestampValues.push([now]);
        }

        // Wenn es Zeitstempel zu setzen gibt
        if (timestampValues.length > 0) {
            sheet.getRange(range.getRow(), timestampCol, timestampValues.length, 1)
                .setValues(timestampValues);
        }
    };

    /**
     * Richtet die notwendigen Trigger für das Spreadsheet ein
     */
    const setupTrigger = () => {
        const triggers = ScriptApp.getProjectTriggers();
        // Prüfe, ob der onOpen Trigger bereits existiert
        if (!triggers.some(t => t.getHandlerFunction() === "onOpen")) {
            ScriptApp.newTrigger("onOpen")
                .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
                .onOpen()
                .create();
        }

        // Prüfe, ob der onEdit Trigger bereits existiert
        if (!triggers.some(t => t.getHandlerFunction() === "onEdit")) {
            ScriptApp.newTrigger("onEdit")
                .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
                .onEdit()
                .create();
        }
    };

    /**
     * Validiert alle relevanten Sheets
     * @returns {boolean} - True wenn alle Sheets valide sind, False sonst
     */
    const validateSheets = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const bankSheet = ss.getSheetByName("Bankbewegungen");
        const eigenSheet = ss.getSheetByName("Eigenbelege");

        return Validator.validateAllSheets(revenueSheet, expenseSheet, bankSheet, eigenSheet);
    };

    /**
     * Gemeinsame Fehlerbehandlungsfunktion für alle Berechnungsfunktionen
     * @param {Function} fn - Die auszuführende Funktion
     * @param {string} errorMessage - Die Fehlermeldung bei einem Fehler
     */
    const executeWithErrorHandling = (fn, errorMessage) => {
        try {
            // Zuerst alle Sheets aktualisieren
            RefreshModule.refreshAllSheets();

            // Dann Validierung durchführen
            if (!validateSheets()) {
                // Validierung fehlgeschlagen - Berechnung abbrechen
                console.error(`${errorMessage}: Validierung fehlgeschlagen`);
                return;
            }

            // Wenn Validierung erfolgreich, Berechnung ausführen
            fn();
        } catch (error) {
            SpreadsheetApp.getUi().alert(`${errorMessage}: ${error.message}`);
            console.error(`${errorMessage}:`, error);
        }
    };

    /**
     * Aktualisiert das aktive Tabellenblatt
     */
    const refreshSheet = () => RefreshModule.refreshActiveSheet();

    /**
     * Berechnet die Umsatzsteuervoranmeldung
     */
    const calculateUStVA = () => {
        executeWithErrorHandling(
            UStVACalculator.calculateUStVA,
            "Fehler bei der UStVA-Berechnung"
        );
    };

    /**
     * Berechnet die BWA (Betriebswirtschaftliche Auswertung)
     */
    const calculateBWA = () => {
        executeWithErrorHandling(
            BWACalculator.calculateBWA,
            "Fehler bei der BWA-Berechnung"
        );
    };

    /**
     * Erstellt die Bilanz
     */
    const calculateBilanz = () => {
        executeWithErrorHandling(
            BilanzCalculator.calculateBilanz,
            "Fehler bei der Bilanzerstellung"
        );
    };

    /**
     * Importiert Dateien aus Google Drive und aktualisiert alle Tabellenblätter
     */
    const importDriveFiles = () => {
        try {
            ImportModule.importDriveFiles();
            RefreshModule.refreshAllSheets();

            // Optional: Nach dem Import auch eine Validierung durchführen
            // validateSheets();
        } catch (error) {
            SpreadsheetApp.getUi().alert("Fehler beim Dateiimport: " + error.message);
            console.error("Import-Fehler:", error);
        }
    };

    // Exportiere alle relevanten Funktionen in den globalen Namensraum,
    // damit sie von Google Apps Script als Trigger und Menüpunkte aufgerufen werden können.
    global.onOpen = onOpen;
    global.onEdit = onEdit;
    global.setupTrigger = setupTrigger;
    global.refreshSheet = refreshSheet;
    global.calculateUStVA = calculateUStVA;
    global.calculateBWA = calculateBWA;
    global.calculateBilanz = calculateBilanz;
    global.importDriveFiles = importDriveFiles;

}));

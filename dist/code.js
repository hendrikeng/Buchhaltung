/* Bundled Code for Google Apps Script */
'use strict';

/**
 * Konfiguration für die Buchhaltungsanwendung
 * Unterstützt die Buchhaltung für Holding und operative GmbH nach SKR04
 */
const config = {
    // Allgemeine Einstellungen
    common: {
        paymentType: ["Überweisung", "Bar", "Kreditkarte", "Paypal", "Lastschrift"],
        months: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"],
        currentYear: new Date().getFullYear()
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

    // Sheet-Struktur Konfiguration
    sheets: {
        // Konfiguration für das Einnahmen-Sheet
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
                steuerbemessung: 10,   // J: Steuerbemessungsgrundlage für Teilzahlungen
                quartal: 11,           // K: Berechnetes Quartal
                zahlungsstatus: 12,    // L: Zahlungsstatus (Offen/Teilbezahlt/Bezahlt)
                zahlungsart: 13,       // M: Zahlungsart
                zahlungsdatum: 14,     // N: Zahlungsdatum
                bankAbgleich: 15,      // O: Bank-Abgleich-Information
                zeitstempel: 16,       // P: Zeitstempel der letzten Änderung
                dateiname: 17,         // Q: Dateiname (für importierte Dateien)
                dateilink: 18          // R: Link zur Originaldatei
            }
        },

        // Konfiguration für das Ausgaben-Sheet (identisch zu Einnahmen)
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
                steuerbemessung: 10,   // J: Steuerbemessungsgrundlage für Teilzahlungen
                quartal: 11,           // K: Berechnetes Quartal
                zahlungsstatus: 12,    // L: Zahlungsstatus (Offen/Teilbezahlt/Bezahlt)
                zahlungsart: 13,       // M: Zahlungsart
                zahlungsdatum: 14,     // N: Zahlungsdatum
                bankAbgleich: 15,      // O: Bank-Abgleich-Information
                zeitstempel: 16,       // P: Zeitstempel der letzten Änderung
                dateiname: 17,         // Q: Dateiname (für importierte Dateien)
                dateilink: 18          // R: Link zur Originaldatei
            }
        },

        // Konfiguration für das Eigenbelege-Sheet
        eigenbelege: {
            columns: {
                datum: 1,              // A: Belegdatum
                belegnummer: 2,        // B: Belegnummer
                kategorie: 3,          // C: Kategorie
                beschreibung: 4,       // D: Beschreibung
                nettobetrag: 5,        // E: Nettobetrag
                mwstSatz: 6,           // F: MwSt-Satz in %
                mwstBetrag: 7,         // G: MwSt-Betrag (E*F)
                bruttoBetrag: 8,       // H: Bruttobetrag (E+G)
                bezahlt: 9,            // I: Bereits bezahlter Betrag
                steuerbemessung: 10,   // J: Steuerbemessungsgrundlage für Teilzahlungen
                quartal: 11,           // K: Berechnetes Quartal
                status: 12,            // L: Status (Offen/Erstattet/Gebucht)
                zahlungsart: 13,       // M: Zahlungsart
                zahlungsdatum: 14,     // N: Erstattungsdatum
                zeitstempel: 16,       // P: Zeitstempel der letzten Änderung
                dateiname: 17,         // Q: Dateiname (für importierte Dateien)
                dateilink: 18          // R: Link zur Originaldatei
            }
        },

        // Konfiguration für das Bankbewegungen-Sheet
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
                anmerkung: 11,         // K: Anmerkung
                matchInfo: 12          // L: Match-Information zu Einnahmen/Ausgaben
            }
        },

        // Konfiguration für das Gesellschafterkonto-Sheet
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
            }
        },

        // Konfiguration für das Holding Transfers-Sheet
        holdingTransfers: {
            columns: {
                datum: 1,              // A: Datum
                beschreibung: 2,       // B: Beschreibung
                kategorie: 3,          // C: Kategorie (Gewinnübertrag/Kapitalrückführung)
                betrag: 4,             // D: Betrag
                anmerkung: 5,          // E: Anmerkung
                zeitstempel: 6         // F: Zeitstempel der letzten Änderung
            }
        },

        // Konfiguration für das Änderungshistorie-Sheet
        aenderungshistorie: {
            columns: {
                datum: 1,              // A: Datum/Zeitstempel
                typ: 2,                // B: Rechnungstyp (Einnahme/Ausgabe)
                dateiname: 3,          // C: Dateiname
                dateilink: 4           // D: Link zur Datei
            }
        }
    },

    // Einnahmen-Konfiguration
    einnahmen: {
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

    // Bankbewegungen-Konfiguration
    bank: {
        // Kombinierte Liste aller Kategorien (automatisch generiert)
        category: [], // Wird dynamisch befüllt

        // Typen von Bankbewegungen
        type: ["Einnahme", "Ausgabe", "Interne Buchung"],

        // Standard-Bankkonto
        defaultAccount: "1200"
    },

    // Gesellschafterkonto-Konfiguration
    gesellschafterkonto: {
        category: ["Gesellschafterdarlehen", "Ausschüttungen", "Kapitalrückführung", "Privatentnahme", "Privateinlage"],
        shareholder: ["Christopher Giebel", "Hendrik Werner"]
    },

    // Eigenbelege-Konfiguration
    eigenbelege: {
        category: ["Kleidung", "Trinkgeld", "Private Vorauslage", "Bürokosten", "Reisekosten", "Bewirtung", "Sonstiges"],
        mapping: {
            "Kleidung": {taxType: "steuerpflichtig"},
            "Trinkgeld": {taxType: "steuerfrei"},
            "Private Vorauslage": {taxType: "steuerfrei"},
            "Bürokosten": {taxType: "steuerpflichtig"},
            "Reisekosten": {taxType: "steuerpflichtig"},
            "Bewirtung": {taxType: "eigenbeleg", besonderheit: "bewirtung"},
            "Sonstiges": {taxType: "steuerpflichtig"}
        },
        status: ["Offen", "Erstattet", "Gebucht"]
    },

    // Holding Transfers-Konfiguration
    holdingTransfers: {
        category: ["Gewinnübertrag", "Kapitalrückführung"]
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

};

// Bankkategorien dynamisch aus den Einnahmen- und Ausgaben-Kategorien befüllen
config.bank.category = [
    ...Object.keys(config.einnahmen.categories),
    ...Object.keys(config.ausgaben.categories),
    ...config.gesellschafterkonto.category,
    ...config.holdingTransfers.category,
    ...config.eigenbelege.category
];

// Duplikate aus den Kategorien entfernen
config.bank.category = [...new Set(config.bank.category)];

// src/helpers.js

/**
 * Hilfsmodule für verschiedene häufig benötigte Funktionen
 */
const Helpers = {
    /**
     * Konvertiert verschiedene Datumsformate in ein gültiges Date-Objekt
     * @param {Date|string} value - Das zu parsende Datum
     * @returns {Date|null} - Das geparste Datum oder null bei ungültigem Format
     */
    parseDate(value) {
        // Wenn bereits ein Date-Objekt
        if (value instanceof Date) return isNaN(value.getTime()) ? null : value;

        // Wenn String
        if (typeof value === "string") {
            // Deutsche Datumsformate (DD.MM.YYYY) unterstützen
            if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(value)) {
                const [day, month, year] = value.split('.').map(Number);
                const date = new Date(year, month - 1, day);
                return isNaN(date.getTime()) ? null : date;
            }

            const d = new Date(value);
            return isNaN(d.getTime()) ? null : d;
        }

        return null;
    },

    /**
     * Konvertiert einen String oder eine Zahl in einen numerischen Währungswert
     * @param {number|string} value - Der zu parsende Wert
     * @returns {number} - Der geparste Währungswert oder 0 bei ungültigem Format
     */
    parseCurrency(value) {
        if (value === null || value === undefined || value === "") return 0;
        if (typeof value === "number") return value;

        // Entferne alle Zeichen außer Ziffern, Komma, Punkt und Minus
        const str = value.toString()
            .replace(/[^\d,.-]/g, "")
            .replace(/,/g, "."); // Alle Kommas durch Punkte ersetzen

        // Bei mehreren Punkten nur den letzten als Dezimaltrenner behandeln
        const parts = str.split('.');
        if (parts.length > 2) {
            const last = parts.pop();
            return parseFloat(parts.join('') + '.' + last);
        }

        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    },

    /**
     * Parst einen MwSt-Satz und normalisiert ihn
     * @param {number|string} value - Der zu parsende MwSt-Satz
     * @returns {number} - Der normalisierte MwSt-Satz (0-100)
     */
    parseMwstRate(value) {
        if (value === null || value === undefined || value === "") {
            // Verwende den Standard-MwSt-Satz aus der Konfiguration oder fallback auf 19%
            return config?.tax?.defaultMwst || 19;
        }

        if (typeof value === "number") {
            // Wenn der Wert < 1 ist, nehmen wir an, dass es sich um einen Dezimalwert handelt (z.B. 0.19)
            return value < 1 ? value * 100 : value;
        }

        // String-Wert parsen und bereinigen
        const rate = parseFloat(
            value.toString()
                .replace(/%/g, "")
                .replace(/,/g, ".")
                .trim()
        );

        // Wenn der geparste Wert ungültig ist, Standardwert zurückgeben
        if (isNaN(rate)) {
            return config?.tax?.defaultMwst || 19;
        }

        // Normalisieren: Werte < 1 werden als Dezimalwerte interpretiert (z.B. 0.19 -> 19)
        return rate < 1 ? rate * 100 : rate;
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

        const nameWithoutExtension = filename.replace(/\.[^/.]+$/, "");

        // Verschiedene Formate erkennen (vom spezifischsten zum allgemeinsten)

        // 1. Format: DD.MM.YYYY im Dateinamen (deutsches Format)
        let match = nameWithoutExtension.match(/(\d{2}[.]\d{2}[.]\d{4})/);
        if (match?.[1]) {
            return match[1];
        }

        // 2. Format: RE-YYYY-MM-DD oder ähnliches mit Trennzeichen
        match = nameWithoutExtension.match(/[^0-9](\d{4}[-_.\/]\d{2}[-_.\/]\d{2})[^0-9]/);
        if (match?.[1]) {
            const dateParts = match[1].split(/[-_.\/]/);
            if (dateParts.length === 3) {
                const [year, month, day] = dateParts;
                return `${day}.${month}.${year}`;
            }
        }

        // 3. Format: YYYY-MM-DD am Anfang oder Ende
        match = nameWithoutExtension.match(/(^|[^0-9])(\d{4}[-_.\/]\d{2}[-_.\/]\d{2})($|[^0-9])/);
        if (match?.[2]) {
            const dateParts = match[2].split(/[-_.\/]/);
            if (dateParts.length === 3) {
                const [year, month, day] = dateParts;
                return `${day}.${month}.${year}`;
            }
        }

        // 4. Format: DD-MM-YYYY mit verschiedenen Trennzeichen
        match = nameWithoutExtension.match(/(\d{2})[-_.\/](\d{2})[-_.\/](\d{4})/);
        if (match) {
            const [_, day, month, year] = match;
            return `${day}.${month}.${year}`;
        }

        // 5. Aktuelles Datum als Fallback (optional)
        const today = new Date();
        const day = String(today.getDate()).padStart(2, '0');
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const year = today.getFullYear();

        // Als Kommentar belassen, da es möglicherweise besser ist, ein leeres Datum zurückzugeben
        // return `${day}.${month}.${year}`;

        return "";
    },

    /**
     * Setzt bedingte Formatierung für eine Spalte
     * @param {Sheet} sheet - Das zu formatierende Sheet
     * @param {string} column - Die zu formatierende Spalte (z.B. "A")
     * @param {Array<Object>} conditions - Array mit Bedingungen ({value, background, fontColor})
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
            const formatRules = conditions.map(({ value, background, fontColor }) =>
                SpreadsheetApp.newConditionalFormatRule()
                    .whenTextEqualTo(value)
                    .setBackground(background || "#ffffff")
                    .setFontColor(fontColor || "#000000")
                    .setRanges([range])
                    .build()
            );

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
            const sheetConfig = config.sheets[sheetName.toLowerCase()]?.columns;
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
        const targetYear = config?.tax?.year || new Date().getFullYear();

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
        let letter = '';
        while (columnIndex > 0) {
            const modulo = (columnIndex - 1) % 26;
            letter = String.fromCharCode(65 + modulo) + letter;
            columnIndex = Math.floor((columnIndex - modulo) / 26);
        }
        return letter;
    }
};

// file: src/importModule.js

/**
 * Modul für den Import von Dateien aus Google Drive in die Buchhaltungstabelle
 */
const ImportModule = (() => {
    /**
     * Importiert Dateien aus einem Ordner in die entsprechenden Sheets
     *
     * @param {Folder} folder - Google Drive Ordner mit den zu importierenden Dateien
     * @param {Sheet} mainSheet - Hauptsheet (Einnahmen oder Ausgaben)
     * @param {string} type - Typ der Dateien ("Einnahme" oder "Ausgabe")
     * @param {Sheet} historySheet - Sheet für die Änderungshistorie
     * @param {Set} existingFiles - Set mit bereits importierten Dateinamen
     * @returns {number} - Anzahl der importierten Dateien
     */
    const importFilesFromFolder = (folder, mainSheet, type, historySheet, existingFiles) => {
        const files = folder.getFiles();
        const newMainRows = [];
        const newHistoryRows = [];
        const timestamp = new Date();
        let importedCount = 0;

        // Konfiguration für das richtige Sheet auswählen
        const sheetConfig = type === "Einnahme"
            ? config.sheets.einnahmen.columns
            : config.sheets.ausgaben.columns;

        // Konfiguration für das Änderungshistorie-Sheet
        const historyConfig = config.sheets.aenderungshistorie.columns;

        while (files.hasNext()) {
            const file = files.next();
            const fileName = file.getName().replace(/\.[^/.]+$/, ""); // Entfernt Dateiendung
            const invoiceName = fileName.replace(/^[^ ]* /, ""); // Entfernt Präfix vor erstem Leerzeichen
            const invoiceDate = Helpers.extractDateFromFilename(fileName);
            const fileUrl = file.getUrl();

            // Prüfe, ob die Datei bereits importiert wurde
            if (!existingFiles.has(fileName)) {
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
                historyRow[historyConfig.typ - 1] = type;                  // Typ (Einnahme/Ausgabe)
                historyRow[historyConfig.dateiname - 1] = fileName;        // Dateiname
                historyRow[historyConfig.dateilink - 1] = fileUrl;         // Link zur Datei

                newHistoryRows.push(historyRow);

                existingFiles.add(fileName); // Zur Liste der importierten Dateien hinzufügen
                importedCount++;
            }
        }

        // Neue Zeilen in die entsprechenden Sheets schreiben
        if (newMainRows.length > 0) {
            mainSheet.getRange(
                mainSheet.getLastRow() + 1,
                1,
                newMainRows.length,
                newMainRows[0].length
            ).setValues(newMainRows);
        }

        if (newHistoryRows.length > 0) {
            historySheet.getRange(
                historySheet.getLastRow() + 1,
                1,
                newHistoryRows.length,
                newHistoryRows[0].length
            ).setValues(newHistoryRows);
        }

        return importedCount;
    };

    /**
     * Hauptfunktion zum Importieren von Dateien aus den Einnahmen- und Ausgabenordnern
     * @returns {number} Anzahl der importierten Dateien
     */
    const importDriveFiles = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const ui = SpreadsheetApp.getUi();
        let totalImported = 0;

        try {
            // Hauptsheets für Einnahmen und Ausgaben abrufen
            const revenueMain = ss.getSheetByName("Einnahmen");
            const expenseMain = ss.getSheetByName("Ausgaben");

            if (!revenueMain || !expenseMain) {
                ui.alert("Fehler: Die Sheets 'Einnahmen' oder 'Ausgaben' existieren nicht!");
                return 0;
            }

            // Änderungshistorie abrufen oder erstellen
            const history = ss.getSheetByName("Änderungshistorie") || ss.insertSheet("Änderungshistorie");

            // Header-Zeile für Änderungshistorie initialisieren, falls nötig
            if (history.getLastRow() === 0) {
                const historyConfig = config.sheets.aenderungshistorie.columns;
                const headerRow = ["", "", "", ""];
                headerRow[historyConfig.datum - 1] = "Datum";
                headerRow[historyConfig.typ - 1] = "Rechnungstyp";
                headerRow[historyConfig.dateiname - 1] = "Dateiname";
                headerRow[historyConfig.dateilink - 1] = "Link zur Datei";

                history.appendRow(headerRow);
                history.getRange(1, 1, 1, 4).setFontWeight("bold");
            }

            // Bereits importierte Dateien aus der Änderungshistorie erfassen
            const historyData = history.getDataRange().getValues();
            const existingFiles = new Set();
            const historyConfig = config.sheets.aenderungshistorie.columns;

            // Überschriftenzeile überspringen und alle Dateinamen sammeln
            for (let i = 1; i < historyData.length; i++) {
                existingFiles.add(historyData[i][historyConfig.dateiname - 1]); // Dateiname aus der entsprechenden Spalte
            }

            // Auf übergeordneten Ordner zugreifen
            let parentFolder;
            try {
                const file = DriveApp.getFileById(ss.getId());
                const parents = file.getParents();
                parentFolder = parents.hasNext() ? parents.next() : null;

                if (!parentFolder) {
                    ui.alert("Fehler: Kein übergeordneter Ordner gefunden.");
                    return 0;
                }
            } catch (e) {
                ui.alert("Fehler beim Zugriff auf Google Drive: " + e.toString());
                return 0;
            }

            // Unterordner für Einnahmen und Ausgaben finden oder erstellen
            let revenueFolder, expenseFolder;
            let importedRevenue = 0, importedExpense = 0;

            try {
                revenueFolder = Helpers.getFolderByName(parentFolder, "Einnahmen");
                if (!revenueFolder) {
                    const createFolder = ui.alert(
                        "Der Ordner 'Einnahmen' existiert nicht. Soll er erstellt werden?",
                        ui.ButtonSet.YES_NO
                    );
                    if (createFolder === ui.Button.YES) {
                        revenueFolder = parentFolder.createFolder("Einnahmen");
                    }
                }
            } catch (e) {
                ui.alert("Fehler beim Zugriff auf den Einnahmen-Ordner: " + e.toString());
            }

            try {
                expenseFolder = Helpers.getFolderByName(parentFolder, "Ausgaben");
                if (!expenseFolder) {
                    const createFolder = ui.alert(
                        "Der Ordner 'Ausgaben' existiert nicht. Soll er erstellt werden?",
                        ui.ButtonSet.YES_NO
                    );
                    if (createFolder === ui.Button.YES) {
                        expenseFolder = parentFolder.createFolder("Ausgaben");
                    }
                }
            } catch (e) {
                ui.alert("Fehler beim Zugriff auf den Ausgaben-Ordner: " + e.toString());
            }

            // Import durchführen wenn Ordner existieren
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

            // Abschluss-Meldung anzeigen
            if (totalImported === 0) {
                ui.alert("Es wurden keine neuen Dateien gefunden.");
            } else {
                ui.alert(
                    `Import abgeschlossen.\n\n` +
                    `${importedRevenue} Einnahmen importiert.\n` +
                    `${importedExpense} Ausgaben importiert.`
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
     * Prüft, ob ein Wert leer ist
     * @param {*} v - Der zu prüfende Wert
     * @returns {boolean} - True, wenn der Wert leer ist
     */
    const isEmpty = v => v == null || v.toString().trim() === "";

    /**
     * Prüft, ob ein Wert keine gültige Zahl ist
     * @param {*} v - Der zu prüfende Wert
     * @returns {boolean} - True, wenn der Wert keine gültige Zahl ist
     */
    const isInvalidNumber = v => isEmpty(v) || isNaN(parseFloat(v.toString().trim()));

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
     * Validiert eine Einnahmen- oder Ausgaben-Zeile
     * @param {Array} row - Die zu validierende Zeile
     * @param {number} rowIndex - Der Index der Zeile (für Fehlermeldungen)
     * @returns {Array<string>} - Array mit Warnungen
     */
    const validateRevenueAndExpenses = (row, rowIndex) => {
        const warnings = [];

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

        // Grundlegende Validierungsregeln
        const baseRules = [
            {check: r => isEmpty(r[0]), message: "Rechnungsdatum fehlt."},
            {check: r => isEmpty(r[1]), message: "Rechnungsnummer fehlt."},
            {check: r => isEmpty(r[2]), message: "Kategorie fehlt."},
            {check: r => isEmpty(r[3]), message: "Kunde fehlt."},
            {check: r => isInvalidNumber(r[4]), message: "Nettobetrag fehlt oder ungültig."},
            {
                check: r => {
                    const mwstStr = r[5] == null ? "" : r[5].toString().trim();
                    if (isEmpty(mwstStr)) return false; // Wird schon durch andere Regel geprüft

                    // MwSt-Satz extrahieren und normalisieren
                    const mwst = Helpers.parseMwstRate(mwstStr);
                    if (isNaN(mwst)) return true;

                    // Prüfe auf erlaubte MwSt-Sätze aus der Konfiguration
                    const allowedRates = config?.tax?.allowedMwst || [0, 7, 19];
                    return !allowedRates.includes(Math.round(mwst));
                },
                message: `Ungültiger MwSt-Satz. Erlaubt sind: ${config?.tax?.allowedMwst?.join('%, ')}% oder leer.`
            }
        ];

        // Status-abhängige Regeln
        const status = row[11] ? row[11].toString().trim().toLowerCase() : "";

        // Regeln für offene Zahlungen
        const openPaymentRules = [
            {check: r => !isEmpty(r[12]), message: 'Zahlungsart darf bei offener Zahlung nicht gesetzt sein.'},
            {check: r => !isEmpty(r[13]), message: 'Zahlungsdatum darf bei offener Zahlung nicht gesetzt sein.'}
        ];

        // Regeln für bezahlte Zahlungen
        const paidPaymentRules = [
            {
                check: r => isEmpty(r[12]),
                message: 'Zahlungsart muss bei bezahlter/teilbezahlter Zahlung gesetzt sein.'
            },
            {
                check: r => isEmpty(r[13]),
                message: 'Zahlungsdatum muss bei bezahlter/teilbezahlter Zahlung gesetzt sein.'
            },
            {
                check: r => {
                    if (isEmpty(r[13])) return false; // Wird schon durch andere Regel geprüft

                    const paymentDate = Helpers.parseDate(r[13]);
                    return paymentDate ? paymentDate > new Date() : false;
                },
                message: "Zahlungsdatum darf nicht in der Zukunft liegen."
            },
            {
                check: r => {
                    if (isEmpty(r[13]) || isEmpty(r[0])) return false;

                    const paymentDate = Helpers.parseDate(r[13]);
                    const invoiceDate = Helpers.parseDate(r[0]);
                    return paymentDate && invoiceDate ? paymentDate < invoiceDate : false;
                },
                message: "Zahlungsdatum darf nicht vor dem Rechnungsdatum liegen."
            }
        ];

        // Regeln basierend auf Zahlungsstatus zusammenstellen
        const paymentRules = status === "offen" ? openPaymentRules : paidPaymentRules;

        // Alle Regeln kombinieren und anwenden
        const rules = [...baseRules, ...paymentRules];
        validateRow(row, rowIndex, rules);

        return warnings;
    };

    /**
     * Validiert das Bankbewegungen-Sheet
     * @param {Sheet} bankSheet - Das zu validierende Sheet
     * @returns {Array<string>} - Array mit Warnungen
     */
    const validateBanking = bankSheet => {
        if (!bankSheet) return ["Bankbewegungen-Sheet nicht gefunden"];

        const data = bankSheet.getDataRange().getValues();
        const warnings = [];

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

        // Regeln für Header- und Footer-Zeilen
        const headerFooterRules = [
            {check: r => isEmpty(r[0]), message: "Buchungsdatum fehlt."},
            {check: r => isEmpty(r[1]), message: "Buchungstext fehlt."},
            {
                check: r => !isEmpty(r[2]) && !isNaN(parseFloat(r[2].toString().trim())),
                message: "Betrag darf nicht gesetzt sein."
            },
            {check: r => isEmpty(r[3]) || isInvalidNumber(r[3]), message: "Saldo fehlt oder ungültig."},
            {check: r => !isEmpty(r[4]), message: "Typ darf nicht gesetzt sein."},
            {check: r => !isEmpty(r[5]), message: "Kategorie darf nicht gesetzt sein."},
            {check: r => !isEmpty(r[6]), message: "Konto (Soll) darf nicht gesetzt sein."},
            {check: r => !isEmpty(r[7]), message: "Gegenkonto (Haben) darf nicht gesetzt sein."}
        ];

        // Regeln für Datenzeilen
        const dataRowRules = [
            {check: r => isEmpty(r[0]), message: "Buchungsdatum fehlt."},
            {check: r => isEmpty(r[1]), message: "Buchungstext fehlt."},
            {check: r => isEmpty(r[2]) || isInvalidNumber(r[2]), message: "Betrag fehlt oder ungültig."},
            {check: r => isEmpty(r[3]) || isInvalidNumber(r[3]), message: "Saldo fehlt oder ungültig."},
            {check: r => isEmpty(r[4]), message: "Typ fehlt."},
            {check: r => isEmpty(r[5]), message: "Kategorie fehlt."},
            {check: r => isEmpty(r[6]), message: "Konto (Soll) fehlt."},
            {check: r => isEmpty(r[7]), message: "Gegenkonto (Haben) fehlt."}
        ];

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
     * @returns {boolean} - True, wenn keine Fehler gefunden wurden
     */
    const validateAllSheets = (revenueSheet, expenseSheet, bankSheet = null) => {
        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Fehler: Benötigte Sheets nicht gefunden!");
            return false;
        }

        try {
            // Daten aus den Sheets lesen
            const revData = revenueSheet.getDataRange().getValues();
            const expData = expenseSheet.getDataRange().getValues();

            // Einnahmen validieren (Header-Zeile überspringen)
            const revenueWarnings = revData.length > 1
                ? revData.slice(1).reduce((acc, row, i) => {
                    if (row.some(cell => cell !== "")) { // Nur nicht-leere Zeilen prüfen
                        return acc.concat(validateRevenueAndExpenses(row, i + 2));
                    }
                    return acc;
                }, [])
                : [];

            // Ausgaben validieren (Header-Zeile überspringen)
            const expenseWarnings = expData.length > 1
                ? expData.slice(1).reduce((acc, row, i) => {
                    if (row.some(cell => cell !== "")) { // Nur nicht-leere Zeilen prüfen
                        return acc.concat(validateRevenueAndExpenses(row, i + 2));
                    }
                    return acc;
                }, [])
                : [];

            // Bank validieren (falls verfügbar)
            const bankWarnings = bankSheet ? validateBanking(bankSheet) : [];

            // Fehlermeldungen zusammenstellen
            const msgArr = [];
            if (revenueWarnings.length) {
                msgArr.push("Fehler in 'Einnahmen':\n" + revenueWarnings.join("\n"));
            }

            if (expenseWarnings.length) {
                msgArr.push("Fehler in 'Ausgaben':\n" + expenseWarnings.join("\n"));
            }

            if (bankWarnings.length) {
                msgArr.push("Fehler in 'Bankbewegungen':\n" + bankWarnings.join("\n"));
            }

            // Fehlermeldungen anzeigen, falls vorhanden
            if (msgArr.length) {
                const ui = SpreadsheetApp.getUi();
                // Bei vielen Fehlern ggf. einschränken, um UI-Limits zu vermeiden
                const maxMsgLength = 1500; // Google Sheets Alert-Dialog hat Beschränkungen
                let alertMsg = msgArr.join("\n\n");

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
                const allowedRates = config?.tax?.allowedMwst || [0, 7, 19];
                return {
                    isValid: allowedRates.includes(Math.round(mwst)),
                    message: allowedRates.includes(Math.round(mwst))
                        ? ""
                        : `Ungültiger MwSt-Satz. Erlaubt sind: ${allowedRates.join('%, ')}%.`
                };

            case 'text':
                return {
                    isValid: !isEmpty(value),
                    message: !isEmpty(value) ? "" : "Text darf nicht leer sein."
                };

            default:
                return {
                    isValid: true,
                    message: ""
                };
        }
    };

    // Öffentliche API des Moduls
    return {
        validateDropdown,
        validateRevenueAndExpenses,
        validateBanking,
        validateAllSheets,
        validateCellValue,
        isEmpty,
        isInvalidNumber
    };
})();

// file: src/refreshModule.js

/**
 * Modul zum Aktualisieren der Tabellenblätter und Neuberechnen von Formeln
 */
const RefreshModule = (() => {
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
                columns = config.sheets.einnahmen.columns;
            } else if (name === "Ausgaben") {
                columns = config.sheets.ausgaben.columns;
            } else if (name === "Eigenbelege") {
                columns = config.sheets.eigenbelege.columns;
            } else {
                return false; // Unbekanntes Sheet
            }

            // Spaltenbuchstaben aus den Indizes generieren
            const columnLetters = {};
            Object.entries(columns).forEach(([key, index]) => {
                columnLetters[key] = Helpers.getColumnLetter(index);
            });

            // Formeln für verschiedene Spalten setzen (mit konfigurierten Spaltenbuchstaben)
            const formulas = {};

            // MwSt-Betrag
            formulas[columns.mwstBetrag] = row =>
                `=${columnLetters.nettobetrag}${row}*${columnLetters.mwstSatz}${row}`;

            // Brutto-Betrag
            formulas[columns.bruttoBetrag] = row =>
                `=${columnLetters.nettobetrag}${row}+${columnLetters.mwstBetrag}${row}`;

            // Steuerbemessungsgrundlage - für Teilzahlungen
            formulas[columns.steuerbemessung] = row =>
                `=(${columnLetters.bruttoBetrag}${row}-${columnLetters.bezahlt}${row})/(1+${columnLetters.mwstSatz}${row})`;

            // Quartal
            formulas[columns.quartal] = row =>
                `=IF(${columnLetters.datum}${row}="";"";ROUNDUP(MONTH(${columnLetters.datum}${row})/3;0))`;

            // Zahlungsstatus
            formulas[columns.zahlungsstatus] = row =>
                `=IF(VALUE(${columnLetters.bezahlt}${row})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${row})>=VALUE(${columnLetters.bruttoBetrag}${row});"Bezahlt";"Teilbezahlt"))`;

            // Formeln für jede Spalte anwenden
            Object.entries(formulas).forEach(([col, formulaFn]) => {
                const formulasArray = Array.from({length: numRows}, (_, i) => [formulaFn(i + 2)]);
                sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulasArray);
            });

            // Bezahlter Betrag - Leerzeichen durch 0 ersetzen für Berechnungen
            const bezahltRange = sheet.getRange(2, columns.bezahlt, numRows, 1);
            const bezahltValues = bezahltRange.getValues().map(([val]) => (val === "" || val === null ? 0 : val));
            bezahltRange.setValues(bezahltValues.map(val => [val]));

            // Dropdown-Validierungen je nach Sheet-Typ setzen
            if (name === "Einnahmen") {
                Validator.validateDropdown(
                    sheet, 2, columns.kategorie, numRows, 1,
                    Object.keys(config.einnahmen.categories)
                );
            } else if (name === "Ausgaben") {
                Validator.validateDropdown(
                    sheet, 2, columns.kategorie, numRows, 1,
                    Object.keys(config.ausgaben.categories)
                );
            } else if (name === "Eigenbelege") {
                Validator.validateDropdown(
                    sheet, 2, columns.kategorie, numRows, 1,
                    config.eigenbelege.category
                );

                // Für Eigenbelege: Status-Dropdown hinzufügen
                Validator.validateDropdown(
                    sheet, 2, columns.status, numRows, 1,
                    config.eigenbelege.status
                );

                // Bedingte Formatierung für Status-Spalte (nur für Eigenbelege)
                Helpers.setConditionalFormattingForColumn(sheet, columnLetters.status, [
                    {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
                    {value: "Erstattet", background: "#FFEB9C", fontColor: "#9C6500"},
                    {value: "Gebucht", background: "#C6EFCE", fontColor: "#006100"}
                ]);
            }

            // Zahlungsart-Dropdown für alle Blätter
            Validator.validateDropdown(
                sheet, 2, columns.zahlungsart, numRows, 1,
                config.common.paymentType
            );

            // Bedingte Formatierung für Zahlungsstatus-Spalte (für alle außer Eigenbelege)
            if (name !== "Eigenbelege") {
                Helpers.setConditionalFormattingForColumn(sheet, columnLetters.zahlungsstatus, [
                    {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
                    {value: "Teilbezahlt", background: "#FFEB9C", fontColor: "#9C6500"},
                    {value: "Bezahlt", background: "#C6EFCE", fontColor: "#006100"}
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

            // Saldo-Formeln setzen (jede Zeile verwendet den Saldo der vorherigen Zeile + aktuellen Betrag)
            if (transRows > 0) {
                sheet.getRange(firstDataRow, 4, transRows, 1).setFormulas(
                    Array.from({length: transRows}, (_, i) =>
                        [`=D${firstDataRow + i - 1}+C${firstDataRow + i}`]
                    )
                );
            }

            // Transaktionstyp basierend auf dem Betrag setzen (Einnahme/Ausgabe)
            const amounts = sheet.getRange(firstDataRow, 3, numDataRows, 1).getValues();
            const typeValues = amounts.map(([val]) => {
                const amt = parseFloat(val) || 0;
                return [amt > 0 ? "Einnahme" : amt < 0 ? "Ausgabe" : ""];
            });
            sheet.getRange(firstDataRow, 5, numDataRows, 1).setValues(typeValues);

            // Dropdown-Validierungen für Typ, Kategorie und Konten
            Validator.validateDropdown(
                sheet, firstDataRow, 5, numDataRows, 1,
                config.bank.type
            );

            Validator.validateDropdown(
                sheet, firstDataRow, 6, numDataRows, 1,
                config.bank.category
            );

            // Konten für Dropdown-Validierung sammeln
            const allowedKontoSoll = Object.values(config.einnahmen.kontoMapping)
                .concat(Object.values(config.ausgaben.kontoMapping))
                .map(m => m.soll);

            const allowedGegenkonto = Object.values(config.einnahmen.kontoMapping)
                .concat(Object.values(config.ausgaben.kontoMapping))
                .map(m => m.gegen);

            // Dropdown-Validierungen für Konten setzen
            Validator.validateDropdown(
                sheet, firstDataRow, 7, numDataRows, 1,
                allowedKontoSoll
            );

            Validator.validateDropdown(
                sheet, firstDataRow, 8, numDataRows, 1,
                allowedGegenkonto
            );

            // Bedingte Formatierung für Transaktionstyp-Spalte
            Helpers.setConditionalFormattingForColumn(sheet, "E", [
                {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100"},
                {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006"}
            ]);

            // REFERENZEN-MATCHING: Suche nach Referenzen in Einnahmen- und Ausgaben-Sheets

            // Daten aus Einnahmen-Sheet
            const einnahmenSheet = ss.getSheetByName("Einnahmen");
            let einnahmenData = [];
            if (einnahmenSheet && einnahmenSheet.getLastRow() > 1) {
                const numEinnahmenRows = einnahmenSheet.getLastRow() - 1;
                // Jetzt auch Beträge (E) und bezahlte Beträge (I) laden
                einnahmenData = einnahmenSheet.getRange(2, 2, numEinnahmenRows, 8).getDisplayValues();
                // B (0), E (3), I (7) Spaltenindizes nach 0-basierter Nummerierung
            }

            // Daten aus Ausgaben-Sheet
            const ausgabenSheet = ss.getSheetByName("Ausgaben");
            let ausgabenData = [];
            if (ausgabenSheet && ausgabenSheet.getLastRow() > 1) {
                const numAusgabenRows = ausgabenSheet.getLastRow() - 1;
                // Jetzt auch Beträge (E) und bezahlte Beträge (I) laden
                ausgabenData = ausgabenSheet.getRange(2, 2, numAusgabenRows, 8).getDisplayValues();
                // B (0), E (3), I (7) Spaltenindizes nach 0-basierter Nummerierung
            }

            // Bankbewegungen Daten für Verarbeitung holen
            const bankData = sheet.getRange(firstDataRow, 1, numDataRows, 12).getDisplayValues();

            // Cache für schnellere Suche
            const einnahmenMap = createReferenceMap(einnahmenData);
            const ausgabenMap = createReferenceMap(ausgabenData);

            // Durchlaufe jede Bankbewegung und suche nach Übereinstimmungen
            for (let i = 0; i < bankData.length; i++) {
                const rowIndex = i + firstDataRow;
                const row = bankData[i];

                // Prüfe, ob es sich um die Endsaldo-Zeile handelt
                const label = row[1] ? row[1].toString().trim().toLowerCase() : "";
                if (rowIndex === lastRow && label === "endsaldo") continue;

                const tranType = row[4]; // Spalte E: Einnahme/Ausgabe
                const refNumber = row[8]; // Spalte I: Referenznummer

                // Nur prüfen, wenn Referenz nicht leer ist
                if (refNumber && refNumber.trim() !== "") {
                    let matchFound = false;
                    let matchInfo = "";

                    const refTrimmed = refNumber.toString().trim();

                    // Betrag für den Vergleich (als absoluter Wert)
                    const betragValue = Math.abs(parseFloat(row[2]) || 0);

                    // In Abhängigkeit vom Typ im entsprechenden Sheet suchen
                    if (tranType === "Einnahme") {
                        // Optimierte Suche in Einnahmen mittels Map
                        const matchResult = findMatch(refTrimmed, einnahmenMap, betragValue);
                        if (matchResult) {
                            matchFound = true;
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
                                tranType === "Einnahme") {
                                // Bankbewegungsdatum holen (aus Spalte A)
                                const zahlungsDatum = row[0];
                                if (zahlungsDatum) {
                                    // Zahldatum im Einnahmen-Sheet aktualisieren (nur wenn leer)
                                    const einnahmenRow = matchResult.row;
                                    const zahldatumRange = einnahmenSheet.getRange(einnahmenRow, 14); // Spalte N für Zahldatum
                                    const aktuellDatum = zahldatumRange.getValue();

                                    if (!aktuellDatum || aktuellDatum === "") {
                                        zahldatumRange.setValue(zahlungsDatum);
                                        matchStatus += " ✓ Datum aktualisiert";
                                    }
                                }
                            }

                            matchInfo = `Einnahme #${matchResult.row}${matchStatus}`;
                        }
                    } else if (tranType === "Ausgabe") {
                        // Optimierte Suche in Ausgaben mittels Map
                        const matchResult = findMatch(refTrimmed, ausgabenMap, betragValue);
                        if (matchResult) {
                            matchFound = true;
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
                                const zahlungsDatum = row[0];
                                if (zahlungsDatum) {
                                    // Zahldatum im Ausgaben-Sheet aktualisieren (nur wenn leer)
                                    const ausgabenRow = matchResult.row;
                                    const zahldatumRange = ausgabenSheet.getRange(ausgabenRow, 14); // Spalte N für Zahldatum
                                    const aktuellDatum = zahldatumRange.getValue();

                                    if (!aktuellDatum || aktuellDatum === "") {
                                        zahldatumRange.setValue(zahlungsDatum);
                                        matchStatus += " ✓ Datum aktualisiert";
                                    }
                                }
                            }

                            matchInfo = `Ausgabe #${matchResult.row}${matchStatus}`;
                        }

                        // FALLS keine Übereinstimmung, auch in Einnahmen suchen (für Gutschriften)
                        if (!matchFound) {
                            // Bei einer Ausgabe, die möglicherweise eine Gutschrift ist,
                            // ignorieren wir den Betrag für den ersten Vergleich, um die entsprechende Einnahme zu finden
                            const gutschriftMatch = findMatch(refTrimmed, einnahmenMap);

                            if (gutschriftMatch) {
                                matchFound = true;
                                let matchStatus = "";

                                // Bei Gutschriften könnte der Betrag abweichen (z.B. Teilgutschrift)
                                // Prüfen, ob die Beträge plausibel sind
                                const gutschriftBetrag = Math.abs(gutschriftMatch.betrag);

                                if (Math.abs(betragValue - gutschriftBetrag) <= 0.01) {
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
                                    const gutschriftDatum = row[0];
                                    if (gutschriftDatum) {
                                        // Gutschriftdatum im Einnahmen-Sheet aktualisieren und "G-" vor die Referenz setzen
                                        const einnahmenRow = gutschriftMatch.row;

                                        // Zahldatum aktualisieren (nur wenn leer)
                                        const zahldatumRange = einnahmenSheet.getRange(einnahmenRow, 14); // Spalte N für Zahldatum
                                        const aktuellDatum = zahldatumRange.getValue();

                                        if (!aktuellDatum || aktuellDatum === "") {
                                            zahldatumRange.setValue(gutschriftDatum);
                                            matchStatus += " ✓ Datum aktualisiert";
                                        }

                                        // Optional: Die Referenz mit "G-" kennzeichnen, um Gutschrift zu markieren
                                        const refRange = einnahmenSheet.getRange(einnahmenRow, 2); // Spalte B für Referenz
                                        const currentRef = refRange.getValue();
                                        if (currentRef && !currentRef.toString().startsWith("G-")) {
                                            refRange.setValue("G-" + currentRef);
                                            matchStatus += " ✓ Ref. markiert";
                                        }
                                    }
                                }

                                matchInfo = `Gutschrift zu Einnahme #${gutschriftMatch.row}${matchStatus}`;
                            }
                        }
                    }

                    // Spezialfälle prüfen
                    if (!matchFound) {
                        const lcRef = refTrimmed.toLowerCase();
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

                    // Ergebnis in Spalte L speichern
                    row[11] = matchFound ? matchInfo : "";
                } else {
                    row[11] = ""; // Leere Spalte L
                }

                // Kontonummern basierend auf Kategorie setzen
                const category = row[5] || "";
                let mapping = null;

                if (tranType === "Einnahme") {
                    mapping = config.einnahmen.kontoMapping[category];
                } else if (tranType === "Ausgabe") {
                    mapping = config.ausgaben.kontoMapping[category];
                }

                if (!mapping) {
                    mapping = {soll: "Manuell prüfen", gegen: "Manuell prüfen"};
                }

                row[6] = mapping.soll;
                row[7] = mapping.gegen;
            }

            // Zuerst nur Spalte L aktualisieren (für bessere Performance und Fehlerbehandlung)
            const matchColumn = bankData.map(row => [row[11]]);
            sheet.getRange(firstDataRow, 12, numDataRows, 1).setValues(matchColumn);

            // Dann die restlichen Daten zurückschreiben
            sheet.getRange(firstDataRow, 1, numDataRows, 11).setValues(
                bankData.map(row => row.slice(0, 11))
            );

            // Verzögerung hinzufügen, um sicherzustellen, dass die Daten verarbeitet wurden
            console.log("Daten wurden zurückgeschrieben, warte 1 Sekunde vor der Formatierung...");
            Utilities.sleep(1000);

            // Log zur Überprüfung
            console.log("Beginne mit Zeilenformatierung für " + bankData.length + " Zeilen");

            // Formatiere die gesamten Zeilen basierend auf dem Match-Typ
            // Wir verarbeiten jede Zeile einzeln, um Probleme zu isolieren
            for (let i = 0; i < bankData.length; i++) {
                try {
                    const rowIndex = firstDataRow + i;
                    const matchInfo = bankData[i][11]; // Spalte L mit Match-Info

                    // Nur formatieren, wenn die Zeile existiert und eine Match-Info hat
                    if (rowIndex <= sheet.getLastRow() && matchInfo && matchInfo.trim() !== "") {
                        console.log(`Versuche Formatierung für Zeile ${rowIndex}, Match: "${matchInfo}"`);

                        // Zuerst den Hintergrund zurücksetzen
                        const rowRange = sheet.getRange(rowIndex, 1, 1, 12);
                        rowRange.setBackground(null);

                        // Kurze Pause, um die Sheets-API nicht zu überlasten
                        if (i % 5 === 0) Utilities.sleep(100);

                        // Dann die neue Farbe anwenden, je nach Match-Typ
                        if (matchInfo.includes("Einnahme")) {
                            if (matchInfo.includes("Vollständige Zahlung")) {
                                console.log(`Setze Grün für Zeile ${rowIndex}`);
                                rowRange.setBackground("#C6EFCE"); // Kräftiges Grün
                            } else if (matchInfo.includes("Teilzahlung")) {
                                console.log(`Setze Orange für Zeile ${rowIndex}`);
                                rowRange.setBackground("#FCE4D6"); // Helles Orange
                            } else {
                                console.log(`Setze Hellgrün für Zeile ${rowIndex}`);
                                rowRange.setBackground("#EAF1DD"); // Helles Grün
                            }
                        } else if (matchInfo.includes("Ausgabe")) {
                            if (matchInfo.includes("Vollständige Zahlung")) {
                                console.log(`Setze Rot für Zeile ${rowIndex}`);
                                rowRange.setBackground("#FFC7CE"); // Helles Rot
                            } else if (matchInfo.includes("Teilzahlung")) {
                                console.log(`Setze Orange für Zeile ${rowIndex}`);
                                rowRange.setBackground("#FCE4D6"); // Helles Orange
                            } else {
                                console.log(`Setze Rosa für Zeile ${rowIndex}`);
                                rowRange.setBackground("#FFCCCC"); // Helles Rosa
                            }
                        } else if (matchInfo.includes("Gutschrift")) {
                            console.log(`Setze Lila für Zeile ${rowIndex}`);
                            rowRange.setBackground("#E6E0FF"); // Helles Lila
                        } else if (matchInfo.includes("Gesellschaftskonto") || matchInfo.includes("Holding")) {
                            console.log(`Setze Gelb für Zeile ${rowIndex}`);
                            rowRange.setBackground("#FFEB9C"); // Helles Gelb
                        }
                    }
                } catch (err) {
                    console.error(`Fehler beim Formatieren von Zeile ${firstDataRow + i}:`, err);
                }
            }

            console.log("Zeilenformatierung abgeschlossen");

            // Bedingte Formatierung für Match-Spalte mit verbesserten Farben
            Helpers.setConditionalFormattingForColumn(sheet, "L", [
                // Grundlegende Match-Typen
                {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100", pattern: "beginsWith"},
                {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006", pattern: "beginsWith"},
                {value: "Gutschrift", background: "#E6E0FF", fontColor: "#5229A3", pattern: "beginsWith"},
                {value: "Gesellschaftskonto", background: "#FFEB9C", fontColor: "#9C6500", pattern: "beginsWith"},
                {value: "Holding", background: "#FFEB9C", fontColor: "#9C6500", pattern: "beginsWith"},

                // Zusätzliche Betragstypen
                {value: "Vollständige Zahlung", background: "#C6EFCE", fontColor: "#006100", pattern: "contains"},
                {value: "Teilzahlung", background: "#FCE4D6", fontColor: "#974706", pattern: "contains"},
                {value: "Unsichere Zahlung", background: "#F8CBAD", fontColor: "#843C0C", pattern: "contains"},
                {value: "Vollständige Gutschrift", background: "#E6E0FF", fontColor: "#5229A3", pattern: "contains"},
                {value: "Teilgutschrift", background: "#E6E0FF", fontColor: "#5229A3", pattern: "contains"},
                {value: "Ungewöhnliche Gutschrift", background: "#FFD966", fontColor: "#7F6000", pattern: "contains"}
            ]);

            // Endsaldo-Zeile aktualisieren
            const lastRowText = sheet.getRange(lastRow, 2).getDisplayValue().toString().trim().toLowerCase();
            const formattedDate = Utilities.formatDate(
                new Date(),
                Session.getScriptTimeZone(),
                "dd.MM.yyyy"
            );

            if (lastRowText === "endsaldo") {
                sheet.getRange(lastRow, 1).setValue(formattedDate);
                sheet.getRange(lastRow, 4).setFormula(`=D${lastRow - 1}`);
            } else {
                sheet.appendRow([
                    formattedDate,
                    "Endsaldo",
                    "",
                    sheet.getRange(lastRow, 4).getValue(),
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    ""
                ]);
            }

            // Spaltenbreiten anpassen
            sheet.autoResizeColumns(1, sheet.getLastColumn());

            // Verzögerung vor dem Aufruf von markPaidInvoices
            console.log("Warte 1 Sekunde vor dem Markieren der bezahlten Rechnungen...");
            Utilities.sleep(1000);

            // Setze farbliche Markierung in den Einnahmen/Ausgaben Sheets basierend auf Zahlungsstatus
            markPaidInvoices(einnahmenSheet, ausgabenSheet);

            return true;
        } catch (e) {
            console.error("Fehler beim Aktualisieren des Bankbewegungen-Sheets:", e);
            return false;
        }
    };

    /**
     * Erstellt eine Map aus Referenznummern für schnellere Suche
     * @param {Array} data - Array mit Referenznummern und Beträgen
     * @returns {Object} - Map mit Referenznummern als Keys
     */
    function createReferenceMap(data) {
        const map = {};
        for (let i = 0; i < data.length; i++) {
            const ref = data[i][0]; // Referenz in Spalte B (Index 0)
            if (ref && ref.trim() !== "") {
                // Entferne "G-" Prefix für den Key, falls vorhanden (für Gutschriften)
                let key = ref.trim();
                const isGutschrift = key.startsWith("G-");
                if (isGutschrift) {
                    key = key.substring(2); // Entferne "G-" Prefix für den Schlüssel
                }

                // Netto-Betrag aus Spalte E (Index 3)
                let betrag = 0;
                if (data[i][3] !== undefined && data[i][3] !== null && data[i][3] !== "") {
                    // Betragsstring säubern
                    const betragStr = data[i][3].toString().replace(/[^0-9.,-]/g, "").replace(",", ".");
                    betrag = parseFloat(betragStr) || 0;

                    // Bei Gutschriften ist der Betrag im Sheet negativ, wir speichern den Absolutwert
                    betrag = Math.abs(betrag);
                }

                // MwSt-Satz aus Spalte F (Index 4)
                let mwstRate = 0;
                if (data[i][4] !== undefined && data[i][4] !== null && data[i][4] !== "") {
                    // MwSt-Satz säubern und parsen
                    const mwstStr = data[i][4].toString().replace(/[^0-9.,-]/g, "").replace(",", ".");
                    mwstRate = parseFloat(mwstStr) || 0;

                    // Wenn der Wert > 1 ist, nehmen wir an, dass es sich um einen Prozentsatz handelt
                    if (mwstRate > 1) {
                        mwstRate = mwstRate;
                    } else {
                        // Ansonsten als Dezimalwert interpretieren und in Prozent umrechnen
                        mwstRate = mwstRate * 100;
                    }
                }

                // Bezahlter Betrag aus Spalte I (Index 7)
                let bezahlt = 0;
                if (data[i][7] !== undefined && data[i][7] !== null && data[i][7] !== "") {
                    // Betragsstring säubern
                    const bezahltStr = data[i][7].toString().replace(/[^0-9.,-]/g, "").replace(",", ".");
                    bezahlt = parseFloat(bezahltStr) || 0;

                    // Bei Gutschriften ist der bezahlte Betrag im Sheet negativ, wir speichern den Absolutwert
                    bezahlt = Math.abs(bezahlt);
                }

                // Speichere auch den Zeilen-Index und die Beträge
                map[key] = {
                    ref: ref.trim(), // Originale Referenz mit G-Prefix, falls vorhanden
                    row: i + 2,
                    betrag: betrag,
                    mwstRate: mwstRate,
                    bezahlt: bezahlt,
                    offen: betrag * (1 + mwstRate/100) - bezahlt,
                    isGutschrift: isGutschrift
                };
            }
        }
        return map;
    }

    /**
     * Markiert bezahlte Einnahmen und Ausgaben farblich basierend auf dem Zahlungsstatus
     * @param {Sheet} einnahmenSheet - Das Einnahmen-Sheet
     * @param {Sheet} ausgabenSheet - Das Ausgaben-Sheet
     */
    function markPaidInvoices(einnahmenSheet, ausgabenSheet) {
        // Sammle zugeordnete Referenzen aus dem Bankbewegungen-Sheet
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bankSheet = ss.getSheetByName("Bankbewegungen");

        // Map zum Speichern der zugeordneten Referenzen und ihrer Bankbewegungsinformationen
        const bankZuordnungen = {};

        if (bankSheet && bankSheet.getLastRow() > 2) {
            const bankData = bankSheet.getRange(3, 1, bankSheet.getLastRow() - 2, 12).getDisplayValues();

            for (const row of bankData) {
                const matchInfo = row[11]; // Spalte L: Match-Info
                const transTyp = row[4];   // Spalte E: Einnahme/Ausgabe
                const bankDatum = row[0];  // Spalte A: Datum

                if (matchInfo && matchInfo.trim() !== "") {
                    // Spezialfall: Gutschriften zu Einnahmen
                    // Das Format der Match-Info ist "Gutschrift zu Einnahme #12"
                    const gutschriftPattern = matchInfo.match(/Gutschrift zu Einnahme #(\d+)/i);

                    if (gutschriftPattern && gutschriftPattern.length >= 2) {
                        const rowNum = parseInt(gutschriftPattern[1]);

                        // Schlüssel für Einnahme erstellen, nicht für Gutschrift
                        const key = `einnahme#${rowNum}`;

                        // Informationen zur Bankbewegung speichern oder aktualisieren
                        if (!bankZuordnungen[key]) {
                            bankZuordnungen[key] = {
                                typ: "einnahme",
                                row: rowNum,
                                bankDatum: bankDatum,
                                matchInfo: matchInfo,
                                transTyp: "Gutschrift"
                            };
                        } else {
                            // Wenn bereits ein Eintrag existiert, füge diesen als zusätzliche Info hinzu
                            bankZuordnungen[key].additional = bankZuordnungen[key].additional || [];
                            bankZuordnungen[key].additional.push({
                                bankDatum: bankDatum,
                                matchInfo: matchInfo
                            });
                        }
                    }
                    // Standardfall: Normale Einnahmen/Ausgaben
                    else {
                        // Z.B. "Einnahme #5 (Vollständige Zahlung)" oder "Ausgabe #7 (Teilzahlung)"
                        const matchParts = matchInfo.match(/(Einnahme|Ausgabe|Gutschrift)[^#]*#(\d+)/i);

                        if (matchParts && matchParts.length >= 3) {
                            const typ = matchParts[1].toLowerCase(); // "einnahme", "ausgabe", "gutschrift"
                            const rowNum = parseInt(matchParts[2]);  // Zeilennummer im Sheet

                            // Schlüssel für die Map erstellen
                            const key = `${typ}#${rowNum}`;

                            // Informationen zur Bankbewegung speichern oder aktualisieren
                            if (!bankZuordnungen[key]) {
                                bankZuordnungen[key] = {
                                    typ: typ,
                                    row: rowNum,
                                    bankDatum: bankDatum,
                                    matchInfo: matchInfo,
                                    transTyp: transTyp
                                };
                            } else {
                                // Wenn bereits ein Eintrag existiert, füge diesen als zusätzliche Info hinzu
                                bankZuordnungen[key].additional = bankZuordnungen[key].additional || [];
                                bankZuordnungen[key].additional.push({
                                    bankDatum: bankDatum,
                                    matchInfo: matchInfo
                                });
                            }
                        }
                    }
                }
            }
        }

        // Markiere bezahlte Einnahmen
        if (einnahmenSheet && einnahmenSheet.getLastRow() > 1) {
            const numEinnahmenRows = einnahmenSheet.getLastRow() - 1;

            // Hole Werte aus dem Einnahmen-Sheet
            const einnahmenData = einnahmenSheet.getRange(2, 1, numEinnahmenRows, 14).getValues();

            // Für jede Zeile prüfen
            for (let i = 0; i < einnahmenData.length; i++) {
                const row = i + 2; // Aktuelle Zeile im Sheet
                const nettoBetrag = parseFloat(einnahmenData[i][4]) || 0; // Spalte E
                const bezahltBetrag = parseFloat(einnahmenData[i][8]) || 0; // Spalte I
                const zahlungsDatum = einnahmenData[i][13]; // Spalte N
                const referenz = einnahmenData[i][1]; // Spalte B

                // Prüfe, ob es eine Gutschrift ist
                const isGutschrift = referenz && referenz.toString().startsWith("G-");

                // Prüfe, ob diese Einnahme im Banking-Sheet zugeordnet wurde
                const zuordnungsKey = `einnahme#${row}`;
                const hatBankzuordnung = bankZuordnungen[zuordnungsKey] !== undefined;

                // Zeilenbereich für die Formatierung
                const rowRange = einnahmenSheet.getRange(row, 1, 1, 14);

                // Status basierend auf Zahlung setzen
                if (Math.abs(bezahltBetrag) >= Math.abs(nettoBetrag) * 0.999) { // 99.9% bezahlt wegen Rundungsfehlern
                    // Vollständig bezahlt
                    if (zahlungsDatum) {
                        if (isGutschrift) {
                            // Gutschriften in Lila markieren
                            rowRange.setBackground("#E6E0FF"); // Helles Lila
                        } else if (hatBankzuordnung) {
                            // Bezahlte Rechnungen mit Bank-Zuordnung in kräftigerem Grün markieren
                            rowRange.setBackground("#C6EFCE"); // Kräftigeres Grün
                        } else {
                            // Bezahlte Rechnungen ohne Bank-Zuordnung in hellerem Grün markieren
                            rowRange.setBackground("#EAF1DD"); // Helles Grün
                        }
                    } else {
                        // Bezahlt aber kein Datum - in Orange markieren
                        rowRange.setBackground("#FCE4D6"); // Helles Orange
                    }
                } else if (bezahltBetrag > 0) {
                    // Teilzahlung
                    if (hatBankzuordnung) {
                        rowRange.setBackground("#FFC7AA"); // Kräftigeres Orange für Teilzahlungen mit Bank-Zuordnung
                    } else {
                        rowRange.setBackground("#FCE4D6"); // Helles Orange für normale Teilzahlungen
                    }
                } else {
                    // Unbezahlt - keine spezielle Farbe
                    rowRange.setBackground(null);
                }

                // Optional: Wenn eine Bankzuordnung existiert, in Spalte O einen Hinweis setzen
                if (hatBankzuordnung) {
                    const zuordnungsInfo = bankZuordnungen[zuordnungsKey];
                    let infoText = "✓ Bank: " + zuordnungsInfo.bankDatum;

                    // Wenn es mehrere Zuordnungen gibt (z.B. bei aufgeteilten Zahlungen)
                    if (zuordnungsInfo.additional && zuordnungsInfo.additional.length > 0) {
                        infoText += " + " + zuordnungsInfo.additional.length + " weitere";
                    }

                    // Titelzeile für Spalte O setzen, falls noch nicht vorhanden
                    if (einnahmenSheet.getRange(1, 15).getValue() === "") {
                        einnahmenSheet.getRange(1, 15).setValue("Bank-Abgleich");
                    }

                    einnahmenSheet.getRange(row, 15).setValue(infoText);
                } else {
                    // Titelzeile für Spalte O setzen, falls noch nicht vorhanden
                    if (einnahmenSheet.getRange(1, 15).getValue() === "") {
                        einnahmenSheet.getRange(1, 15).setValue("Bank-Abgleich");
                    }

                    // Setze die Zelle leer, falls keine Zuordnung existiert
                    const currentValue = einnahmenSheet.getRange(row, 15).getValue();
                    if (currentValue && currentValue.toString().startsWith("✓ Bank:")) {
                        einnahmenSheet.getRange(row, 15).setValue("");
                    }
                }
            }
        }

        // Markiere bezahlte Ausgaben (ähnliche Logik wie bei Einnahmen)
        if (ausgabenSheet && ausgabenSheet.getLastRow() > 1) {
            const numAusgabenRows = ausgabenSheet.getLastRow() - 1;

            // Hole Werte aus dem Ausgaben-Sheet
            const ausgabenData = ausgabenSheet.getRange(2, 1, numAusgabenRows, 14).getValues();

            // Für jede Zeile prüfen
            for (let i = 0; i < ausgabenData.length; i++) {
                const row = i + 2; // Aktuelle Zeile im Sheet
                const nettoBetrag = parseFloat(ausgabenData[i][4]) || 0; // Spalte E
                const bezahltBetrag = parseFloat(ausgabenData[i][8]) || 0; // Spalte I
                const zahlungsDatum = ausgabenData[i][13]; // Spalte N

                // Prüfe, ob diese Ausgabe im Banking-Sheet zugeordnet wurde
                const zuordnungsKey = `ausgabe#${row}`;
                const hatBankzuordnung = bankZuordnungen[zuordnungsKey] !== undefined;

                // Zeilenbereich für die Formatierung
                const rowRange = ausgabenSheet.getRange(row, 1, 1, 14);

                // Status basierend auf Zahlung setzen
                if (Math.abs(bezahltBetrag) >= Math.abs(nettoBetrag) * 0.999) { // 99.9% bezahlt wegen Rundungsfehlern
                    // Vollständig bezahlt
                    if (zahlungsDatum) {
                        if (hatBankzuordnung) {
                            // Bezahlte Ausgaben mit Bank-Zuordnung in kräftigerem Grün markieren
                            rowRange.setBackground("#C6EFCE"); // Kräftigeres Grün
                        } else {
                            // Bezahlte Ausgaben ohne Bank-Zuordnung in hellerem Grün markieren
                            rowRange.setBackground("#EAF1DD"); // Helles Grün
                        }
                    } else {
                        // Bezahlt aber kein Datum - in Orange markieren
                        rowRange.setBackground("#FCE4D6"); // Helles Orange
                    }
                } else if (bezahltBetrag > 0) {
                    // Teilzahlung
                    if (hatBankzuordnung) {
                        rowRange.setBackground("#FFC7AA"); // Kräftigeres Orange für Teilzahlungen mit Bank-Zuordnung
                    } else {
                        rowRange.setBackground("#FCE4D6"); // Helles Orange für normale Teilzahlungen
                    }
                } else {
                    // Unbezahlt - keine spezielle Farbe
                    rowRange.setBackground(null);
                }

                // Optional: Wenn eine Bankzuordnung existiert, in Spalte O einen Hinweis setzen
                if (hatBankzuordnung) {
                    const zuordnungsInfo = bankZuordnungen[zuordnungsKey];
                    let infoText = "✓ Bank: " + zuordnungsInfo.bankDatum;

                    // Wenn es mehrere Zuordnungen gibt (z.B. bei aufgeteilten Zahlungen)
                    if (zuordnungsInfo.additional && zuordnungsInfo.additional.length > 0) {
                        infoText += " + " + zuordnungsInfo.additional.length + " weitere";
                    }

                    // Titelzeile für Spalte O setzen, falls noch nicht vorhanden
                    if (ausgabenSheet.getRange(1, 15).getValue() === "") {
                        ausgabenSheet.getRange(1, 15).setValue("Bank-Abgleich");
                    }

                    ausgabenSheet.getRange(row, 15).setValue(infoText);
                } else {
                    // Titelzeile für Spalte O setzen, falls noch nicht vorhanden
                    if (ausgabenSheet.getRange(1, 15).getValue() === "") {
                        ausgabenSheet.getRange(1, 15).setValue("Bank-Abgleich");
                    }

                    // Setze die Zelle leer, falls keine Zuordnung existiert
                    const currentValue = ausgabenSheet.getRange(row, 15).getValue();
                    if (currentValue && currentValue.toString().startsWith("✓ Bank:")) {
                        ausgabenSheet.getRange(row, 15).setValue("");
                    }
                }
            }
        }
    }

    /**
     * Findet eine Übereinstimmung in der Referenz-Map
     * @param {string} reference - Zu suchende Referenz
     * @param {Object} refMap - Map mit Referenznummern
     * @param {number} betrag - Betrag der Bankbewegung (absoluter Wert)
     * @returns {Object|null} - Gefundene Übereinstimmung oder null
     */
    function findMatch(reference, refMap, betrag = null) {
        // 1. Exakte Übereinstimmung
        if (refMap[reference]) {
            const match = refMap[reference];

            // Wenn ein Betrag angegeben ist
            if (betrag !== null) {
                const matchNetto = Math.abs(match.betrag);
                const matchMwstRate = parseFloat(match.mwstRate || 0) / 100;

                // Bruttobetrag berechnen (Netto + MwSt)
                const matchBrutto = matchNetto * (1 + matchMwstRate);
                const matchBezahlt = Math.abs(match.bezahlt);

                // Beträge mit größerer Toleranz vergleichen (2 Cent Unterschied erlauben)
                const tolerance = 0.02;

                // Fall 1: Betrag entspricht genau dem Bruttobetrag (Vollständige Zahlung)
                if (Math.abs(betrag - matchBrutto) <= tolerance) {
                    match.matchType = "Vollständige Zahlung";
                    return match;
                }

                // Fall 2: Position ist bereits vollständig bezahlt
                if (Math.abs(matchBezahlt - matchBrutto) <= tolerance && matchBezahlt > 0) {
                    match.matchType = "Vollständige Zahlung";
                    return match;
                }

                // Fall 3: Teilzahlung (Bankbetrag kleiner als Rechnungsbetrag)
                // Nur als Teilzahlung markieren, wenn der Betrag deutlich kleiner ist (> 10% Differenz)
                if (betrag < matchBrutto && (matchBrutto - betrag) > (matchBrutto * 0.1)) {
                    match.matchType = "Teilzahlung";
                    return match;
                }

                // Fall 4: Betrag ist größer als Bruttobetrag, aber vermutlich trotzdem die richtige Zahlung
                // (z.B. wegen Rundungen oder kleinen Gebühren)
                if (betrag > matchBrutto && (betrag - matchBrutto) <= tolerance) {
                    match.matchType = "Vollständige Zahlung";
                    return match;
                }

                // Fall 5: Bei allen anderen Fällen (Beträge weichen stärker ab)
                match.matchType = "Unsichere Zahlung";
                match.betragsDifferenz = Math.abs(betrag - matchBrutto).toFixed(2);
                return match;
            } else {
                // Ohne Betragsvergleich
                match.matchType = "Referenz-Match";
                return match;
            }
        }

        // 2. Teilweise Übereinstimmung (in beide Richtungen)
        for (const key in refMap) {
            if (reference.includes(key) || key.includes(reference)) {
                const match = refMap[key];

                // Nur Betragsvergleich wenn nötig
                if (betrag !== null) {
                    const matchNetto = Math.abs(match.betrag);
                    const matchMwstRate = parseFloat(match.mwstRate || 0) / 100;

                    // Bruttobetrag berechnen (Netto + MwSt)
                    const matchBrutto = matchNetto * (1 + matchMwstRate);
                    const matchBezahlt = Math.abs(match.bezahlt);

                    // Größere Toleranz für Teilweise Übereinstimmungen
                    const tolerance = 0.02;

                    // Beträge stimmen mit Toleranz überein
                    if (Math.abs(betrag - matchBrutto) <= tolerance) {
                        match.matchType = "Vollständige Zahlung";
                        return match;
                    }

                    // Wenn Position bereits vollständig bezahlt ist
                    if (Math.abs(matchBezahlt - matchBrutto) <= tolerance && matchBezahlt > 0) {
                        match.matchType = "Vollständige Zahlung";
                        return match;
                    }

                    // Teilzahlung mit größerer Toleranz für nicht-exakte Referenzen
                    // Nur als Teilzahlung markieren, wenn der Betrag deutlich kleiner ist (> 10% Differenz)
                    if (betrag < matchBrutto && (matchBrutto - betrag) > (matchBrutto * 0.1)) {
                        match.matchType = "Teilzahlung";
                        return match;
                    }

                    // Bei allen anderen Fällen: Unsichere Zahlung
                    match.matchType = "Unsichere Zahlung";
                    match.betragsDifferenz = Math.abs(betrag - matchBrutto).toFixed(2);
                    return match;
                } else {
                    // Ohne Betragsvergleich
                    match.matchType = "Teilw. Match";
                    return match;
                }
            }
        }

        return null;
    }

    /**
     * Aktualisiert das aktive Tabellenblatt
     */
    const refreshActiveSheet = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sheet = ss.getActiveSheet();
            const name = sheet.getName();
            const ui = SpreadsheetApp.getUi();

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

            ["Einnahmen", "Ausgaben", "Eigenbelege", "Bankbewegungen"].forEach(name => {
                const sheet = ss.getSheetByName(name);
                if (sheet) {
                    name === "Bankbewegungen" ? refreshBankSheet(sheet) : refreshDataSheet(sheet);
                }
            });
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
            // Zahlungsdatum prüfen (nur abgeschlossene Zahlungen)
            const paymentDate = Helpers.parseDate(row[13]);
            if (!paymentDate || paymentDate > new Date()) return;

            // Monat und Jahr prüfen (nur relevantes Geschäftsjahr)
            const month = Helpers.getMonthFromRow(row);
            if (!month || month < 1 || month > 12) return;

            // Beträge aus der Zeile extrahieren
            const netto = Helpers.parseCurrency(row[4]);
            const restNetto = Helpers.parseCurrency(row[10]) || 0; // Steuerbemessungsgrundlage für Teilzahlungen
            const gezahlt = netto - restNetto; // Tatsächlich gezahlter/erhaltener Betrag

            // Falls kein Betrag gezahlt wurde, nichts zu verarbeiten
            if (gezahlt === 0) return;

            // MwSt-Satz normalisieren
            const mwstRate = Helpers.parseMwstRate(row[5]);
            const roundedRate = Math.round(mwstRate);

            // Steuer berechnen
            const tax = gezahlt * (mwstRate / 100);

            // Kategorie ermitteln
            const category = row[2]?.toString().trim() || "";

            // Je nach Typ (Einnahme/Ausgabe/Eigenbeleg) unterschiedlich verarbeiten
            if (isIncome) {
                // EINNAHMEN
                const catCfg = config.einnahmen.categories[category] ?? {};
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
                    } else {
                        console.warn(`Unbekannter MwSt-Satz für Einnahme: ${roundedRate}%`);
                    }
                }
            } else if (isEigen) {
                // EIGENBELEGE
                const eigenCfg = config.eigenbelege.mapping[category] ?? {};
                const taxType = eigenCfg.taxType ?? "steuerpflichtig";

                if (taxType === "steuerfrei") {
                    // Steuerfreie Eigenbelege
                    data[month].eigenbelege_steuerfrei += gezahlt;
                } else if (taxType === "eigenbeleg" && eigenCfg.besonderheit === "bewirtung") {
                    // Bewirtungsbelege (nur 70% der Vorsteuer absetzbar)
                    data[month].eigenbelege_steuerpflichtig += gezahlt;

                    if (roundedRate === 7 || roundedRate === 19) {
                        data[month][`vst_${roundedRate}`] += tax * 0.7; // 70% absetzbare Vorsteuer
                        data[month].nicht_abzugsfaehige_vst += tax * 0.3; // 30% nicht absetzbar
                    } else {
                        console.warn(`Unbekannter MwSt-Satz für Bewirtung: ${roundedRate}%`);
                    }
                } else {
                    // Normale steuerpflichtige Eigenbelege
                    data[month].eigenbelege_steuerpflichtig += gezahlt;

                    if (roundedRate === 7 || roundedRate === 19) {
                        data[month][`vst_${roundedRate}`] += tax;
                    } else {
                        console.warn(`Unbekannter MwSt-Satz für Eigenbeleg: ${roundedRate}%`);
                    }
                }
            } else {
                // AUSGABEN
                const catCfg = config.ausgaben.categories[category] ?? {};
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
                    } else if (roundedRate !== 0) {
                        // 0% ist kein Fehler, daher nur für andere Sätze warnen
                        console.warn(`Unbekannter MwSt-Satz für Ausgabe: ${roundedRate}%`);
                    }
                }
            }
        } catch (e) {
            console.error("Fehler bei der Verarbeitung einer UStVA-Zeile:", e);
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
        const ustZahlung = (d.ust_7 + d.ust_19) - ((d.vst_7 + d.vst_19) - d.nicht_abzugsfaehige_vst);

        // Berechnung des Gesamtergebnisses: Einnahmen minus Ausgaben (ohne Steueranteil)
        const ergebnis = (d.steuerpflichtige_einnahmen + d.steuerfreie_inland_einnahmen + d.steuerfreie_ausland_einnahmen) -
            (d.steuerpflichtige_ausgaben + d.steuerfreie_inland_ausgaben + d.steuerfreie_ausland_ausgaben +
                d.eigenbelege_steuerpflichtig + d.eigenbelege_steuerfrei);

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
     * Hauptfunktion zur Berechnung der UStVA
     * Sammelt Daten aus allen relevanten Sheets und erstellt ein UStVA-Sheet
     */
    const calculateUStVA = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();

            // Benötigte Sheets abrufen
            const revenueSheet = ss.getSheetByName("Einnahmen");
            const expenseSheet = ss.getSheetByName("Ausgaben");
            const eigenSheet = ss.getSheetByName("Eigenbelege");

            // Prüfen, ob alle benötigten Sheets vorhanden sind
            if (!revenueSheet || !expenseSheet) {
                ui.alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben' wurde nicht gefunden.");
                return;
            }

            // Sheets validieren
            if (!Validator.validateAllSheets(revenueSheet, expenseSheet)) {
                ui.alert("Die UStVA-Berechnung wurde abgebrochen, da Fehler in den Daten gefunden wurden.");
                return;
            }

            // Daten aus den Sheets laden
            const revenueData = revenueSheet.getDataRange().getValues();
            const expenseData = expenseSheet.getDataRange().getValues();
            const eigenData = eigenSheet ? eigenSheet.getDataRange().getValues() : [];

            // Leere UStVA-Datenstruktur für alle Monate erstellen
            const ustvaData = Object.fromEntries(
                Array.from({length: 12}, (_, i) => [i + 1, createEmptyUStVA()])
            );

            // Helfer-Funktion zum Verarbeiten von Datenzeilen
            const processRows = (data, isIncome, isEigen = false) => {
                data.slice(1).forEach(row => { // Ab Zeile 2 (nach Header)
                    processUStVARow(row, ustvaData, isIncome, isEigen);
                });
            };

            // Daten verarbeiten
            processRows(revenueData, true);         // Einnahmen
            processRows(expenseData, false);        // Ausgaben
            if (eigenData.length) {
                processRows(eigenData, false, true); // Eigenbelege
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
            config.common.months.forEach((name, i) => {
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
            const ustvaSheet = ss.getSheetByName("UStVA") || ss.insertSheet("UStVA");
            ustvaSheet.clearContents();

            // Daten in das Sheet schreiben
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

            // Zahlenformate anwenden
            // Währungsformat für Beträge (Spalten 2-16)
            ustvaSheet.getRange(2, 2, outputRows.length - 1, 15).setNumberFormat("#,##0.00 €");

            // Spaltenbreiten anpassen
            ustvaSheet.autoResizeColumns(1, outputRows[0].length);

            // UStVA-Sheet aktivieren
            ss.setActiveSheet(ustvaSheet);

            ui.alert("UStVA wurde erfolgreich aktualisiert!");
        } catch (e) {
            console.error("Fehler bei der UStVA-Berechnung:", e);
            SpreadsheetApp.getUi().alert("Fehler bei der UStVA-Berechnung: " + e.toString());
        }
    };

    // Öffentliche API des Moduls
    return {
        calculateUStVA,
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
     *
     * @param {Array} row - Zeile aus dem Einnahmen-Sheet
     * @param {Object} bwaData - BWA-Datenstruktur
     */
    const processRevenue = (row, bwaData) => {
        try {
            const m = Helpers.getMonthFromRow(row);
            if (!m) return;

            const amount = Helpers.parseCurrency(row[4]);
            const category = row[2]?.toString().trim() || "";

            if (["Gewinnvortrag", "Verlustvortrag", "Gewinnvortrag/Verlustvortrag"].includes(category)) return;

            // Spezielle Kategorien direkt zuordnen
            if (category === "Sonstige betriebliche Erträge") return void (bwaData[m].sonstigeErtraege += amount);
            if (category === "Erträge aus Vermietung/Verpachtung") return void (bwaData[m].vermietung += amount);
            if (category === "Erträge aus Zuschüssen") return void (bwaData[m].zuschuesse += amount);
            if (category === "Erträge aus Währungsgewinnen") return void (bwaData[m].waehrungsgewinne += amount);
            if (category === "Erträge aus Anlagenabgängen") return void (bwaData[m].anlagenabgaenge += amount);

            // BWA-Mapping aus Konfiguration verwenden
            const mapping = config.einnahmen.bwaMapping[category];
            if (["umsatzerloese", "provisionserloese"].includes(mapping)) {
                bwaData[m][mapping] += amount;
            } else if (Helpers.parseMwstRate(row[5]) === 0) {
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
     *
     * @param {Array} row - Zeile aus dem Ausgaben-Sheet
     * @param {Object} bwaData - BWA-Datenstruktur
     */
    const processExpense = (row, bwaData) => {
        try {
            const m = Helpers.getMonthFromRow(row);
            if (!m) return;

            const amount = Helpers.parseCurrency(row[4]);
            const category = row[2]?.toString().trim() || "";

            // Nicht-betriebliche Positionen ignorieren
            if (["Privatentnahme", "Privateinlage", "Holding Transfers",
                "Gewinnvortrag", "Verlustvortrag", "Gewinnvortrag/Verlustvortrag"].includes(category)) return;

            // Spezielle Kategorien direkt zuordnen
            if (category === "Bruttolöhne & Gehälter") return void (bwaData[m].bruttoLoehne += amount);
            if (category === "Soziale Abgaben & Arbeitgeberanteile") return void (bwaData[m].sozialeAbgaben += amount);
            if (category === "Sonstige Personalkosten") return void (bwaData[m].sonstigePersonalkosten += amount);
            if (category === "Gewerbesteuerrückstellungen") return void (bwaData[m].gewerbesteuerRueckstellungen += amount);
            if (category === "Telefon & Internet") return void (bwaData[m].telefonInternet += amount);
            if (category === "Bürokosten") return void (bwaData[m].buerokosten += amount);
            if (category === "Fortbildungskosten") return void (bwaData[m].fortbildungskosten += amount);

            // BWA-Mapping aus Konfiguration verwenden
            const mapping = config.ausgaben.bwaMapping[category];
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
     *
     * @param {Array} row - Zeile aus dem Eigenbelege-Sheet
     * @param {Object} bwaData - BWA-Datenstruktur
     */
    const processEigen = (row, bwaData) => {
        try {
            const m = Helpers.getMonthFromRow(row);
            if (!m) return;

            const amount = Helpers.parseCurrency(row[4]);
            const category = row[2]?.toString().trim() || "";
            const eigenCfg = config.eigenbelege.mapping[category] ?? {};
            const taxType = eigenCfg.taxType ?? "steuerpflichtig";

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
     * Sammelt alle BWA-Daten aus den verschiedenen Sheets
     *
     * @returns {Object|null} BWA-Datenstruktur oder null bei Fehler
     */
    const aggregateBWAData = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const revenueSheet = ss.getSheetByName("Einnahmen");
            const expenseSheet = ss.getSheetByName("Ausgaben");
            const eigenSheet = ss.getSheetByName("Eigenbelege");

            if (!revenueSheet || !expenseSheet) {
                SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
                return null;
            }

            // BWA-Daten für alle Monate initialisieren
            const bwaData = Object.fromEntries(Array.from({length: 12}, (_, i) => [i + 1, createEmptyBWA()]));

            // Daten aus den Sheets verarbeiten
            revenueSheet.getDataRange().getValues().slice(1).forEach(processRevenue);
            expenseSheet.getDataRange().getValues().slice(1).forEach(processExpense);
            if (eigenSheet) eigenSheet.getDataRange().getValues().slice(1).forEach(processEigen);

            // Gruppensummen und weitere Berechnungen
            for (let m = 1; m <= 12; m++) {
                const d = bwaData[m];

                // Erlöse
                d.gesamtErloese = d.umsatzerloese + d.provisionserloese + d.steuerfreieInlandEinnahmen +
                    d.steuerfreieAuslandEinnahmen + d.sonstigeErtraege + d.vermietung +
                    d.zuschuesse + d.waehrungsgewinne + d.anlagenabgaenge;

                // Materialkosten
                d.gesamtWareneinsatz = d.wareneinsatz + d.fremdleistungen + d.rohHilfsBetriebsstoffe;

                // Betriebsausgaben
                d.gesamtBetriebsausgaben = d.bruttoLoehne + d.sozialeAbgaben + d.sonstigePersonalkosten +
                    d.werbungMarketing + d.reisekosten + d.versicherungen + d.telefonInternet +
                    d.buerokosten + d.fortbildungskosten + d.kfzKosten + d.sonstigeAufwendungen;

                // Abschreibungen & Zinsen
                d.gesamtAbschreibungenZinsen = d.abschreibungenMaschinen + d.abschreibungenBueromaterial +
                    d.abschreibungenImmateriell + d.zinsenBank + d.zinsenGesellschafter +
                    d.leasingkosten;

                // Besondere Posten
                d.gesamtBesonderePosten = d.eigenkapitalveraenderungen + d.gesellschafterdarlehen + d.ausschuettungen;

                // Rückstellungen
                d.gesamtRueckstellungenTransfers = d.steuerrueckstellungen + d.rueckstellungenSonstige;

                // EBIT
                d.ebit = d.gesamtErloese - (d.gesamtWareneinsatz + d.gesamtBetriebsausgaben +
                    d.gesamtAbschreibungenZinsen + d.gesamtBesonderePosten);

                // Steuern berechnen
                const taxConfig = config.tax.isHolding ? config.tax.holding : config.tax.operative;

                // Für Holdings gelten spezielle Steuersätze wegen Beteiligungsprivileg
                const steuerfaktor = config.tax.isHolding
                    ? taxConfig.gewinnUebertragSteuerpflichtig / 100
                    : 1;

                d.gewerbesteuer = d.ebit * (taxConfig.gewerbesteuer / 100) * steuerfaktor;
                d.koerperschaftsteuer = d.ebit * (taxConfig.koerperschaftsteuer / 100) * steuerfaktor;
                d.solidaritaetszuschlag = d.koerperschaftsteuer * (taxConfig.solidaritaetszuschlag / 100);

                // Gesamte Steuerlast
                d.steuerlast = d.koerperschaftsteuer + d.solidaritaetszuschlag + d.gewerbesteuer;

                // Gewinn nach Steuern
                d.gewinnNachSteuern = d.ebit - d.steuerlast;
            }

            return bwaData;
        } catch (e) {
            console.error("Fehler bei der Aggregation der BWA-Daten:", e);
            SpreadsheetApp.getUi().alert("Fehler bei der BWA-Berechnung: " + e.toString());
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
                headers.push(`${config.common.months[m]} (€)`);
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

        // Zeile zusammenstellen
        return [pos.label,
            ...monthly.slice(0, 3), quarters[0],
            ...monthly.slice(3, 6), quarters[1],
            ...monthly.slice(6, 9), quarters[2],
            ...monthly.slice(9, 12), quarters[3],
            yearly];
    };

    /**
     * Hauptfunktion zur Berechnung der BWA
     */
    const calculateBWA = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const bwaData = aggregateBWAData();
            if (!bwaData) return;

            // Positionen definieren
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
            const bwaSheet = ss.getSheetByName("BWA") || ss.insertSheet("BWA");
            bwaSheet.clearContents();

            // Daten in das Sheet schreiben
            const dataRange = bwaSheet.getRange(1, 1, outputRows.length, outputRows[0].length);
            dataRange.setValues(outputRows);

            // Formatierungen anwenden
            // Header formatieren
            bwaSheet.getRange(1, 1, 1, headerRow.length).setFontWeight("bold").setBackground("#f3f3f3");

            // Gruppenüberschriften formatieren
            for (let i = 0, rowIndex = 2; i < bwaGruppen.length; i++) {
                bwaSheet.getRange(rowIndex, 1).setFontWeight("bold");
                rowIndex += bwaGruppen[i].count + 1; // +1 für die Leerzeile
            }

            // Währungsformat für alle Zahlenwerte
            bwaSheet.getRange(2, 2, outputRows.length - 1, headerRow.length - 1).setNumberFormat("#,##0.00 €");

            // Summen-Zeilen hervorheben
            const summenZeilen = [11, 15, 26, 33, 36, 38, 39, 46];
            summenZeilen.forEach(row => {
                bwaSheet.getRange(row, 1, 1, headerRow.length).setBackground("#e6f2ff");
            });

            // EBIT und Jahresüberschuss hervorheben
            bwaSheet.getRange(39, 1, 1, headerRow.length).setFontWeight("bold");
            bwaSheet.getRange(46, 1, 1, headerRow.length).setFontWeight("bold");

            // Spaltenbreiten anpassen
            bwaSheet.autoResizeColumns(1, headerRow.length);

            // Erfolgsbenachrichtigung
            SpreadsheetApp.getUi().alert("BWA wurde aktualisiert!");

            // BWA-Sheet aktivieren
            ss.setActiveSheet(bwaSheet);

        } catch (e) {
            console.error("Fehler bei der BWA-Berechnung:", e);
            SpreadsheetApp.getUi().alert("Fehler bei der BWA-Berechnung: " + e.toString());
        }
    };

    // Öffentliche API des Moduls
    return {calculateBWA};
})();

// file: src/bilanzCalculator.js

/**
 * Modul zur Erstellung einer Bilanz nach SKR04
 * Erstellt eine standardkonforme Bilanz basierend auf den Daten aus anderen Sheets
 */
const BilanzCalculator = (() => {
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
     *
     * @returns {Object} Bilanz-Datenstruktur mit befüllten Werten
     */
    const aggregateBilanzData = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const bilanzData = createEmptyBilanz();

            // 1. Banksaldo aus "Bankbewegungen" (Endsaldo)
            const bankSheet = ss.getSheetByName("Bankbewegungen");
            if (bankSheet) {
                const lastRow = bankSheet.getLastRow();
                if (lastRow >= 1) {
                    const label = bankSheet.getRange(lastRow, 2).getValue().toString().toLowerCase();
                    if (label === "endsaldo") {
                        bilanzData.aktiva.bankguthaben = Helpers.parseCurrency(bankSheet.getRange(lastRow, 4).getValue());
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
            bilanzData.passiva.stammkapital = config.tax.stammkapital || 25000;

            // 4. Suche nach Gesellschafterdarlehen im Gesellschafterkonto-Sheet
            const gesellschafterSheet = ss.getSheetByName("Gesellschafterkonto");
            if (gesellschafterSheet) {
                let darlehenSumme = 0;
                const data = gesellschafterSheet.getDataRange().getValues();

                // Überschrift überspringen
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    // Prüfen, ob es sich um ein Gesellschafterdarlehen handelt
                    if (row[2] && row[2].toString().toLowerCase() === "gesellschafterdarlehen") {
                        darlehenSumme += Helpers.parseCurrency(row[3] || 0);
                    }
                }

                bilanzData.passiva.gesellschafterdarlehen = darlehenSumme;
            }

            // 5. Steuerrückstellungen aus BWA oder Ausgaben-Sheet
            const ausSheet = ss.getSheetByName("Ausgaben");
            if (ausSheet) {
                let steuerRueckstellungen = 0;
                const data = ausSheet.getDataRange().getValues();

                // Überschrift überspringen
                for (let i = 1; i < data.length; i++) {
                    const row = data[i];
                    const category = row[2]?.toString().trim() || "";

                    if (["Gewerbesteuerrückstellungen", "Körperschaftsteuer", "Solidaritätszuschlag", "Sonstige Steuerrückstellungen"].includes(category)) {
                        steuerRueckstellungen += Helpers.parseCurrency(row[4] || 0);
                    }
                }

                bilanzData.passiva.steuerrueckstellungen = steuerRueckstellungen;
            }

            // 6. Berechnung der Summen
            bilanzData.aktiva.summeAnlagevermoegen =
                bilanzData.aktiva.sachanlagen +
                bilanzData.aktiva.immaterielleVermoegen +
                bilanzData.aktiva.finanzanlagen;

            bilanzData.aktiva.summeUmlaufvermoegen =
                bilanzData.aktiva.bankguthaben +
                bilanzData.aktiva.kasse +
                bilanzData.aktiva.forderungenLuL +
                bilanzData.aktiva.vorraete;

            bilanzData.aktiva.summeAktiva =
                bilanzData.aktiva.summeAnlagevermoegen +
                bilanzData.aktiva.summeUmlaufvermoegen +
                bilanzData.aktiva.rechnungsabgrenzung;

            bilanzData.passiva.summeEigenkapital =
                bilanzData.passiva.stammkapital +
                bilanzData.passiva.kapitalruecklagen +
                bilanzData.passiva.gewinnvortrag -
                bilanzData.passiva.verlustvortrag +
                bilanzData.passiva.jahresueberschuss;

            bilanzData.passiva.summeVerbindlichkeiten =
                bilanzData.passiva.bankdarlehen +
                bilanzData.passiva.gesellschafterdarlehen +
                bilanzData.passiva.verbindlichkeitenLuL +
                bilanzData.passiva.steuerrueckstellungen;

            bilanzData.passiva.summePassiva =
                bilanzData.passiva.summeEigenkapital +
                bilanzData.passiva.summeVerbindlichkeiten +
                bilanzData.passiva.rechnungsabgrenzung;

            return bilanzData;
        } catch (e) {
            console.error("Fehler bei der Sammlung der Bilanzdaten:", e);
            SpreadsheetApp.getUi().alert("Fehler bei der Bilanzerstellung: " + e.toString());
            return null;
        }
    };

    /**
     * Konvertiert einen Zahlenwert in eine Zellenformel oder direkte Zahl
     * @param {number} value - Der Wert
     * @param {boolean} useFormula - Ob eine Formel verwendet werden soll
     * @returns {string|number} - Formel als String oder direkter Wert
     */
    const valueOrFormula = (value, useFormula = false) => {
        if (value === 0 && useFormula) {
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
        const year = config.tax.year || new Date().getFullYear();

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
        const year = config.tax.year || new Date().getFullYear();

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
     * Hauptfunktion zur Erstellung der Bilanz
     * Sammelt Daten und erstellt ein Bilanz-Sheet
     */
    const calculateBilanz = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();

            // Bilanzdaten aggregieren
            const bilanzData = aggregateBilanzData();
            if (!bilanzData) return;

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
                    `Die Bilanzsummen von Aktiva (${aktivaSumme} €) und Passiva (${passivaSumme} €) ` +
                    `stimmen nicht überein. Differenz: ${differenz.toFixed(2)} €. ` +
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

            // Schreibe Aktiva ab Zelle A1 und Passiva ab Zelle E1
            bilanzSheet.getRange(1, 1, aktivaArray.length, 2).setValues(aktivaArray);
            bilanzSheet.getRange(1, 5, passivaArray.length, 2).setValues(passivaArray);

            // Formatierung anwenden
            // Überschriften formatieren
            bilanzSheet.getRange("A1").setFontWeight("bold").setFontSize(12);
            bilanzSheet.getRange("E1").setFontWeight("bold").setFontSize(12);

            // Zwischensummen und Gesamtsummen formatieren
            const summenZeilenAktiva = [7, 14, 18]; // Zeilen mit Summen in Aktiva
            const summenZeilenPassiva = [9, 16, 20]; // Zeilen mit Summen in Passiva

            summenZeilenAktiva.forEach(row => {
                bilanzSheet.getRange(row, 1, 1, 2).setFontWeight("bold");
                if (row === 18) { // Gesamtsumme Aktiva
                    bilanzSheet.getRange(row, 1, 1, 2).setBackground("#e6f2ff");
                } else {
                    bilanzSheet.getRange(row, 1, 1, 2).setBackground("#f0f0f0");
                }
            });

            summenZeilenPassiva.forEach(row => {
                bilanzSheet.getRange(row, 5, 1, 2).setFontWeight("bold");
                if (row === 20) { // Gesamtsumme Passiva
                    bilanzSheet.getRange(row, 5, 1, 2).setBackground("#e6f2ff");
                } else {
                    bilanzSheet.getRange(row, 5, 1, 2).setBackground("#f0f0f0");
                }
            });

            // Abschnittsüberschriften formatieren
            [3, 9, 11, 16].forEach(row => {
                bilanzSheet.getRange(row, 1).setFontWeight("bold");
            });

            [3, 11, 17].forEach(row => {
                bilanzSheet.getRange(row, 5).setFontWeight("bold");
            });

            // Währungsformat für Beträge anwenden
            bilanzSheet.getRange("B4:B18").setNumberFormat("#,##0.00 €");
            bilanzSheet.getRange("F4:F20").setNumberFormat("#,##0.00 €");

            // Spaltenbreiten anpassen
            bilanzSheet.autoResizeColumns(1, 6);

            // Erfolgsmeldung
            ui.alert("Die Bilanz wurde erfolgreich erstellt!");
        } catch (e) {
            console.error("Fehler bei der Bilanzerstellung:", e);
            SpreadsheetApp.getUi().alert("Fehler bei der Bilanzerstellung: " + e.toString());
        }
    };

    // Öffentliche API des Moduls
    return {
        calculateBilanz,
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

    // Mapping von Blattname zu Zeitstempel-Spalte
    const mapping = {
        "Einnahmen": 16,
        "Ausgaben": 16,
        "Eigenbelege": 16,
        "Bankbewegungen": 11,
        "Gesellschafterkonto": 12,
        "Holding Transfers": 6
    };

    // Prüfen, ob wir dieses Sheet bearbeiten sollen
    if (!(name in mapping)) return;

    // Header-Zeile ignorieren
    if (range.getRow() === 1) return;

    // Spalte für Zeitstempel
    const timestampCol = mapping[name];

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
        SpreadsheetApp.newTrigger("onOpen")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onOpen()
            .create();
    }

    // Prüfe, ob der onEdit Trigger bereits existiert
    if (!triggers.some(t => t.getHandlerFunction() === "onEdit")) {
        SpreadsheetApp.newTrigger("onEdit")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onEdit()
            .create();
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
    try {
        RefreshModule.refreshAllSheets();
        UStVACalculator.calculateUStVA();
    } catch (error) {
        SpreadsheetApp.getUi().alert("Fehler bei der UStVA-Berechnung: " + error.message);
        console.error("UStVA-Fehler:", error);
    }
};

/**
 * Berechnet die BWA (Betriebswirtschaftliche Auswertung)
 */
const calculateBWA = () => {
    try {
        RefreshModule.refreshAllSheets();
        BWACalculator.calculateBWA();
    } catch (error) {
        SpreadsheetApp.getUi().alert("Fehler bei der BWA-Berechnung: " + error.message);
        console.error("BWA-Fehler:", error);
    }
};

/**
 * Erstellt die Bilanz
 */
const calculateBilanz = () => {
    try {
        RefreshModule.refreshAllSheets();
        BilanzCalculator.calculateBilanz(); // Aktualisierter Aufruf
    } catch (error) {
        SpreadsheetApp.getUi().alert("Fehler bei der Bilanzerstellung: " + error.message);
        console.error("Bilanz-Fehler:", error);
    }
};

/**
 * Importiert Dateien aus Google Drive und aktualisiert alle Tabellenblätter
 */
const importDriveFiles = () => {
    try {
        ImportModule.importDriveFiles();
        RefreshModule.refreshAllSheets();
    } catch (error) {
        SpreadsheetApp.getUi().alert("Fehler beim Dateiimport: " + error.message);
        console.error("Import-Fehler:", error);
    }
};

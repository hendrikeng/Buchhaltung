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
     * Allgemeine Einstellungen für die Buchhaltungsanwendung
     */
    var common = {
        paymentType: ["Überweisung", "Bar", "Kreditkarte", "Paypal", "Lastschrift"],
        months: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"],
        shareholders: ["Christopher Giebel", "Hendrik Werner"],
        employees: [],
        currentYear: new Date().getFullYear(),
        version: "1.0.0"
    };

    /**
     * Steuerliche Einstellungen für die Buchhaltungsanwendung
     */
    var tax = {
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
    };

    /**
     * Sheet-spezifische Konfigurationen
     */
    var sheets = {
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
            // Kategorien mit einheitlicher Struktur
            categories: {
                "Erlöse aus Lieferungen und Leistungen": {
                    taxType: "steuerpflichtig",
                    group: "umsatz",
                    besonderheit: null
                },
                "Provisionserlöse": {
                    taxType: "steuerpflichtig",
                    group: "umsatz",
                    besonderheit: null
                },
                "Sonstige betriebliche Erträge": {
                    taxType: "steuerpflichtig",
                    group: "sonstige_ertraege",
                    besonderheit: null
                },
                "Erträge aus Vermietung/Verpachtung": {
                    taxType: "steuerfrei_inland",
                    group: "vermietung",
                    besonderheit: null
                },
                "Erträge aus Zuschüssen": {
                    taxType: "steuerpflichtig",
                    group: "zuschuesse",
                    besonderheit: null
                },
                "Erträge aus Währungsgewinnen": {
                    taxType: "steuerpflichtig",
                    group: "sonstige_ertraege",
                    besonderheit: null
                },
                "Erträge aus Anlagenabgängen": {
                    taxType: "steuerpflichtig",
                    group: "anlagenabgaenge",
                    besonderheit: null
                },
                "Darlehen": {
                    taxType: "steuerfrei_inland",
                    group: "finanzen",
                    besonderheit: null
                },
                "Zinsen": {
                    taxType: "steuerfrei_inland",
                    group: "finanzen",
                    besonderheit: null
                },
                "Gewinnvortrag": {
                    taxType: "steuerfrei_inland",
                    group: "eigenkapital",
                    besonderheit: null
                },
                "Verlustvortrag": {
                    taxType: "steuerfrei_inland",
                    group: "eigenkapital",
                    besonderheit: null
                }
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
            // Kategorien mit einheitlicher Struktur
            categories: {
                // Materialaufwand & Wareneinsatz
                "Wareneinsatz": {
                    taxType: "steuerpflichtig",
                    group: "material",
                    besonderheit: null
                },
                "Bezogene Leistungen": {
                    taxType: "steuerpflichtig",
                    group: "material",
                    besonderheit: null
                },
                "Roh-, Hilfs- & Betriebsstoffe": {
                    taxType: "steuerpflichtig",
                    group: "material",
                    besonderheit: null
                },

                // Personalkosten
                "Bruttolöhne & Gehälter": {
                    taxType: "steuerfrei_inland",
                    group: "personal",
                    besonderheit: null
                },
                "Soziale Abgaben & Arbeitgeberanteile": {
                    taxType: "steuerfrei_inland",
                    group: "personal",
                    besonderheit: null
                },
                "Sonstige Personalkosten": {
                    taxType: "steuerpflichtig",
                    group: "personal",
                    besonderheit: null
                },

                // Raumkosten
                "Miete": {
                    taxType: "steuerfrei_inland",
                    group: "raum",
                    besonderheit: null
                },
                "Nebenkosten": {
                    taxType: "steuerpflichtig",
                    group: "raum",
                    besonderheit: null
                },

                // Betriebskosten
                "Betriebskosten": {
                    taxType: "steuerpflichtig",
                    group: "betrieb",
                    besonderheit: null
                },
                "Marketing & Werbung": {
                    taxType: "steuerpflichtig",
                    group: "betrieb",
                    besonderheit: null
                },
                "Reisekosten": {
                    taxType: "steuerpflichtig",
                    group: "betrieb",
                    besonderheit: null
                },
                "Versicherungen": {
                    taxType: "steuerfrei_inland",
                    group: "betrieb",
                    besonderheit: null
                },
                "Porto": {
                    taxType: "steuerfrei_inland",
                    group: "betrieb",
                    besonderheit: null
                },
                "Google Ads": {
                    taxType: "steuerfrei_ausland",
                    group: "betrieb",
                    besonderheit: null
                },
                "AWS": {
                    taxType: "steuerfrei_ausland",
                    group: "betrieb",
                    besonderheit: null
                },
                "Facebook Ads": {
                    taxType: "steuerfrei_ausland",
                    group: "betrieb",
                    besonderheit: null
                },
                "Bewirtung": {
                    taxType: "steuerpflichtig",
                    group: "betrieb",
                    besonderheit: "bewirtung"
                },
                "Telefon & Internet": {
                    taxType: "steuerpflichtig",
                    group: "betrieb",
                    besonderheit: null
                },
                "Bürokosten": {
                    taxType: "steuerpflichtig",
                    group: "betrieb",
                    besonderheit: null
                },
                "Fortbildungskosten": {
                    taxType: "steuerpflichtig",
                    group: "betrieb",
                    besonderheit: null
                },

                // Abschreibungen & Zinsen
                "Abschreibungen Maschinen": {
                    taxType: "steuerpflichtig",
                    group: "abschreibung",
                    besonderheit: null
                },
                "Abschreibungen Büroausstattung": {
                    taxType: "steuerpflichtig",
                    group: "abschreibung",
                    besonderheit: null
                },
                "Abschreibungen immaterielle Wirtschaftsgüter": {
                    taxType: "steuerpflichtig",
                    group: "abschreibung",
                    besonderheit: null
                },
                "Zinsen auf Bankdarlehen": {
                    taxType: "steuerpflichtig",
                    group: "zinsen",
                    besonderheit: null
                },
                "Zinsen auf Gesellschafterdarlehen": {
                    taxType: "steuerpflichtig",
                    group: "zinsen",
                    besonderheit: null
                },
                "Leasingkosten": {
                    taxType: "steuerpflichtig",
                    group: "abschreibung",
                    besonderheit: null
                },

                // Steuern & Rückstellungen
                "Gewerbesteuerrückstellungen": {
                    taxType: "steuerfrei_inland",
                    group: "steuer",
                    besonderheit: null
                },
                "Körperschaftsteuer": {
                    taxType: "steuerfrei_inland",
                    group: "steuer",
                    besonderheit: null
                },
                "Solidaritätszuschlag": {
                    taxType: "steuerfrei_inland",
                    group: "steuer",
                    besonderheit: null
                },
                "Sonstige Steuerrückstellungen": {
                    taxType: "steuerfrei_inland",
                    group: "steuer",
                    besonderheit: null
                },

                // Sonstige Aufwendungen
                "Sonstige betriebliche Aufwendungen": {
                    taxType: "steuerpflichtig",
                    group: "sonstige",
                    besonderheit: null
                }
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
            // Kategorien mit einheitlicher Struktur
            categories: {
                "Kleidung": {
                    taxType: "steuerpflichtig",
                    group: "betrieb",
                    besonderheit: null
                },
                "Trinkgeld": {
                    taxType: "steuerfrei",
                    group: "betrieb",
                    besonderheit: null
                },
                "Private Vorauslage": {
                    taxType: "steuerfrei",
                    group: "sonstige",
                    besonderheit: null
                },
                "Bürokosten": {
                    taxType: "steuerpflichtig",
                    group: "betrieb",
                    besonderheit: null
                },
                "Reisekosten": {
                    taxType: "steuerpflichtig",
                    group: "betrieb",
                    besonderheit: null
                },
                "Bewirtung": {
                    taxType: "eigenbeleg",
                    group: "betrieb",
                    besonderheit: "bewirtung"
                },
                "Sonstiges": {
                    taxType: "steuerpflichtig",
                    group: "sonstige",
                    besonderheit: null
                }
            },
            // Konten-Mapping für Eigenbelege
            kontoMapping: {
                "Kleidung": {soll: "6800", gegen: "1200", vorsteuer: "1576"},
                "Trinkgeld": {soll: "6800", gegen: "1200"},
                "Private Vorauslage": {soll: "6800", gegen: "1200"},
                "Bürokosten": {soll: "6815", gegen: "1200", vorsteuer: "1576"},
                "Reisekosten": {soll: "6650", gegen: "1200", vorsteuer: "1576"},
                "Bewirtung": {soll: "6670", gegen: "1200", vorsteuer: "1576"},
                "Sonstiges": {soll: "6800", gegen: "1200", vorsteuer: "1576"}
            },
            // BWA-Mapping für Eigenbelege
            bwaMapping: {
                "Kleidung": "sonstigeAufwendungen",
                "Trinkgeld": "sonstigeAufwendungen",
                "Private Vorauslage": "sonstigeAufwendungen",
                "Bürokosten": "buerokosten",
                "Reisekosten": "reisekosten",
                "Bewirtung": "sonstigeAufwendungen",
                "Sonstiges": "sonstigeAufwendungen"
            }
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
            // Kategorien als Objekte mit einheitlicher Struktur
            categories: {
                "Gesellschafterdarlehen": {
                    taxType: "steuerfrei_inland",
                    group: "gesellschafter",
                    besonderheit: null
                },
                "Ausschüttungen": {
                    taxType: "steuerfrei_inland",
                    group: "gesellschafter",
                    besonderheit: null
                },
                "Kapitalrückführung": {
                    taxType: "steuerfrei_inland",
                    group: "gesellschafter",
                    besonderheit: null
                },
                "Privatentnahme": {
                    taxType: "steuerfrei_inland",
                    group: "gesellschafter",
                    besonderheit: null
                },
                "Privateinlage": {
                    taxType: "steuerfrei_inland",
                    group: "gesellschafter",
                    besonderheit: null
                }
            },
            // Konten-Mapping für Gesellschafterkonto
            kontoMapping: {
                "Gesellschafterdarlehen": {soll: "1200", gegen: "3300"},
                "Ausschüttungen": {soll: "2000", gegen: "1200"},
                "Kapitalrückführung": {soll: "2000", gegen: "1200"},
                "Privatentnahme": {soll: "1600", gegen: "1200"},
                "Privateinlage": {soll: "1200", gegen: "1600"}
            },
            // BWA-Mapping für Gesellschafterkonto
            bwaMapping: {
                "Gesellschafterdarlehen": "gesellschafterdarlehen",
                "Ausschüttungen": "ausschuettungen",
                "Kapitalrückführung": "eigenkapitalveraenderungen",
                "Privatentnahme": "eigenkapitalveraenderungen",
                "Privateinlage": "eigenkapitalveraenderungen"
            },
        },

        // Holding Transfers-Konfiguration
        holdingTransfers: {
            columns: {
                datum: 1,              // A: Datum
                betrag: 2,             // B: Betrag
                art: 3,                // C: Art (Gewinnübertrag/Kapitalrückführung)
                buchungstext: 4,       // D: Buchungstext
                zahlungsstatus: 5,     // E: Status
                referenz: 6,           // F: Referenznummer zur Bankbewegung
                zeitstempel: 7         // G: Zeitstempel der letzten Änderung
            },
            // Kategorien als Objekte mit einheitlicher Struktur
            categories: {
                "Gewinnübertrag": {
                    taxType: "steuerfrei_inland",
                    group: "holding",
                    besonderheit: null
                },
                "Kapitalrückführung": {
                    taxType: "steuerfrei_inland",
                    group: "holding",
                    besonderheit: null
                }
            },
            // Konten-Mapping für Holding Transfers
            kontoMapping: {
                "Gewinnübertrag": {soll: "1200", gegen: "8999"},
                "Kapitalrückführung": {soll: "1200", gegen: "2000"}
            },
            // BWA-Mapping für Holding Transfers
            bwaMapping: {
                "Gewinnübertrag": "gesamtRueckstellungenTransfers",
                "Kapitalrückführung": "eigenkapitalveraenderungen"
            }
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

        // Änderungshistorie-Konfiguration
        aenderungshistorie: {
            columns: {
                datum: 1,              // A: Datum/Zeitstempel
                typ: 2,                // B: Rechnungstyp (Einnahme/Ausgabe)
                dateiname: 3,          // C: Dateiname
                dateilink: 4           // D: Link zur Datei
            }
        }
    };

    /**
     * Kontenplan SKR04 (Auszug der wichtigsten Konten)
     */
    var accounts = {
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
    };

    // config/index.js

    /**
     * Konfiguration für die Buchhaltungsanwendung
     * Unterstützt die Buchhaltung für Holding und operative GmbH nach SKR04
     */
    const config = {
        // Allgemeine Einstellungen
        common,

        // Steuerliche Einstellungen
        tax,

        // Sheet-Konfigurationen
        ...sheets,

        // Kontenplan
        kontenplan: accounts,

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
                ...Object.keys(this.gesellschafterkonto.categories),
                ...Object.keys(this.holdingTransfers.categories),
            ];

            // Duplikate aus den Kategorien entfernen
            this.bankbewegungen.categories = [...new Set(this.bankbewegungen.categories)];

            return this;
        }
    };

    // Initialisierung ausführen und exportieren
    var config$1 = config.initialize();

    // utils/sheetUtils.js
    /**
     * Funktionen für die Arbeit mit Google Sheets
     */

    /**
     * Optimierte Batch-Verarbeitung für Google Sheets API-Calls
     * Vermeidet häufige API-Calls, die zur Drosselung führen können
     * @param {Sheet} sheet - Das zu aktualisierende Sheet
     * @param {Array} data - Array mit Daten-Zeilen
     * @param {number} startRow - Startzeile (1-basiert)
     * @param {number} startCol - Startspalte (1-basiert)
     */
    function batchWriteToSheet(sheet, data, startRow, startCol) {
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

    /**
     * Setzt bedingte Formatierung für eine Spalte
     * @param {Sheet} sheet - Das zu formatierende Sheet
     * @param {string} column - Die zu formatierende Spalte (z.B. "A")
     * @param {Array<Object>} conditions - Array mit Bedingungen ({value, background, fontColor, pattern})
     */
    function setConditionalFormattingForColumn(sheet, column, conditions) {
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
    }

    /**
     * Sucht nach einem Ordner mit bestimmtem Namen innerhalb eines übergeordneten Ordners
     * @param {Folder} parent - Der übergeordnete Ordner
     * @param {string} name - Der gesuchte Ordnername
     * @returns {Folder|null} - Der gefundene Ordner oder null
     */
    function getFolderByName(parent, name) {
        if (!parent) return null;

        try {
            const folderIter = parent.getFoldersByName(name);
            return folderIter.hasNext() ? folderIter.next() : null;
        } catch (e) {
            console.error("Fehler beim Suchen des Ordners:", e);
            return null;
        }
    }

    /**
     * Holt ein Sheet aus dem Spreadsheet oder erstellt es, wenn es nicht existiert
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {string} sheetName - Name des Sheets
     * @returns {Sheet} - Das vorhandene oder neu erstellte Sheet
     */
    function getOrCreateSheet(ss, sheetName) {
        let sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
        }

        return sheet;
    }

    var sheetUtils$1 = {
        batchWriteToSheet,
        setConditionalFormattingForColumn,
        getFolderByName,
        getOrCreateSheet
    };

    // modules/importModule/fileManager.js

    /**
     * Ruft den übergeordneten Ordner des aktuellen Spreadsheets ab
     * @returns {Folder|null} - Der übergeordnete Ordner oder null bei Fehler
     */
    function getParentFolder() {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const file = DriveApp.getFileById(ss.getId());
            const parents = file.getParents();
            return parents.hasNext() ? parents.next() : null;
        } catch (e) {
            console.error("Fehler beim Abrufen des übergeordneten Ordners:", e);
            return null;
        }
    }

    /**
     * Findet oder erstellt einen Ordner mit dem angegebenen Namen
     * @param {Folder} parentFolder - Der übergeordnete Ordner
     * @param {string} folderName - Der zu findende oder erstellende Ordnername
     * @param {UI} ui - UI-Objekt für Dialoge
     * @returns {Folder|null} - Der gefundene oder erstellte Ordner oder null bei Fehler
     */
    function findOrCreateFolder(parentFolder, folderName, ui) {
        if (!parentFolder) return null;

        try {
            let folder = sheetUtils$1.getFolderByName(parentFolder, folderName);

            if (!folder) {
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
    }

    var fileManager = {
        getParentFolder,
        findOrCreateFolder
    };

    // utils/numberUtils.js
    /**
     * Funktionen für die Verarbeitung und Formatierung von Zahlen und Währungen
     */

    /**
     * Konvertiert einen String oder eine Zahl in einen numerischen Währungswert
     * @param {number|string} value - Der zu parsende Wert
     * @returns {number} - Der geparste Währungswert oder 0 bei ungültigem Format
     */
    function parseCurrency$1(value) {
        if (value === null || value === undefined || value === "") return 0;
        if (typeof value === "number") return value;

        // Entferne alle Zeichen außer Ziffern, Komma, Punkt und Minus
        const str = value.toString()
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

        return isNaN(result) ? 0 : result;
    }

    /**
     * Parst einen MwSt-Satz und normalisiert ihn
     * @param {number|string} value - Der zu parsende MwSt-Satz
     * @param {number} defaultMwst - Standard-MwSt-Satz, falls value ungültig ist
     * @returns {number} - Der normalisierte MwSt-Satz (0-100)
     */
    function parseMwstRate$1(value, defaultMwst = 19) {
        if (value === null || value === undefined || value === "") {
            return defaultMwst;
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

        return result;
    }

    /**
     * Formatiert einen Währungsbetrag im deutschen Format
     * @param {number|string} amount - Der zu formatierende Betrag
     * @param {string} currency - Das Währungssymbol (Standard: "€")
     * @returns {string} - Der formatierte Betrag
     */
    function formatCurrency$1(amount, currency = "€") {
        const value = parseCurrency$1(amount);
        return value.toLocaleString('de-DE', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }) + " " + currency;
    }

    /**
     * Prüft, ob zwei Zahlenwerte im Rahmen einer bestimmten Toleranz gleich sind
     * @param {number} a - Erster Wert
     * @param {number} b - Zweiter Wert
     * @param {number} tolerance - Toleranzwert (Standard: 0.01)
     * @returns {boolean} - true wenn Werte innerhalb der Toleranz gleich sind
     */
    function isApproximatelyEqual$1(a, b, tolerance = 0.01) {
        return Math.abs(a - b) <= tolerance;
    }

    /**
     * Sicheres Runden eines Werts auf n Dezimalstellen
     * @param {number} value - Der zu rundende Wert
     * @param {number} decimals - Anzahl der Dezimalstellen (Standard: 2)
     * @returns {number} - Gerundeter Wert
     */
    function round$1(value, decimals = 2) {
        const factor = Math.pow(10, decimals);
        return Math.round((value + Number.EPSILON) * factor) / factor;
    }

    /**
     * Prüft, ob ein Wert keine gültige Zahl ist
     * @param {*} v - Der zu prüfende Wert
     * @returns {boolean} - True, wenn der Wert keine gültige Zahl ist
     */
    function isInvalidNumber$1(v) {
        if (v === null || v === undefined || v === "") return true;
        return isNaN(parseFloat(v.toString().trim()));
    }

    var dateUtils = {
        parseCurrency: parseCurrency$1,
        parseMwstRate: parseMwstRate$1,
        formatCurrency: formatCurrency$1,
        isApproximatelyEqual: isApproximatelyEqual$1,
        round: round$1,
        isInvalidNumber: isInvalidNumber$1
    };

    // modules/importModule/dataProcessor.js

    /**
     * Importiert Dateien aus einem Ordner in die entsprechenden Sheets
     *
     * @param {Folder} folder - Google Drive Ordner mit den zu importierenden Dateien
     * @param {Sheet} mainSheet - Hauptsheet (Einnahmen, Ausgaben oder Eigenbelege)
     * @param {string} type - Typ der Dateien ("Einnahme", "Ausgabe" oder "Eigenbeleg")
     * @param {Sheet} historySheet - Sheet für die Änderungshistorie
     * @param {Set} existingFiles - Set mit bereits importierten Dateinamen
     * @param {Object} config - Die Konfiguration
     * @returns {number} - Anzahl der importierten Dateien
     */
    function importFilesFromFolder(folder, mainSheet, type, historySheet, existingFiles, config) {
        if (!folder || !mainSheet || !historySheet) return 0;

        const files = folder.getFiles();
        const newMainRows = [];
        const newHistoryRows = [];
        const timestamp = new Date();
        let importedCount = 0;

        // Konfiguration für das richtige Sheet auswählen
        const sheetTypeMap = {
            "Einnahme": config.einnahmen.columns,
            "Ausgabe": config.ausgaben.columns,
            "Eigenbeleg": config.eigenbelege.columns
        };

        const sheetConfig = sheetTypeMap[type];
        if (!sheetConfig) {
            console.error("Unbekannter Dateityp:", type);
            return 0;
        }

        // Konfiguration für das Änderungshistorie-Sheet
        const historyConfig = config.aenderungshistorie.columns;

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
                const invoiceDate = dateUtils.extractDateFromFilename(fileName);
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
            sheetUtils$1.batchWriteToSheet(
                mainSheet,
                newMainRows,
                mainSheet.getLastRow() + 1,
                1
            );
        }

        if (newHistoryRows.length > 0) {
            sheetUtils$1.batchWriteToSheet(
                historySheet,
                newHistoryRows,
                historySheet.getLastRow() + 1,
                1
            );
        }

        return importedCount;
    }

    var dataProcessor = {
        importFilesFromFolder
    };

    // modules/importModule/historyTracker.js

    /**
     * Initialisiert die Änderungshistorie, falls sie nicht existiert
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Object} config - Die Konfiguration
     * @returns {Sheet} - Das initialisierte Historie-Sheet
     */
    function getOrCreateHistorySheet(ss, config) {
        const history = sheetUtils$1.getOrCreateSheet(ss, "Änderungshistorie");

        try {
            if (history.getLastRow() === 0) {
                const historyConfig = config.aenderungshistorie.columns;
                const headerRow = Array(Math.max(...Object.values(historyConfig))).fill("");

                headerRow[historyConfig.datum - 1] = "Datum";
                headerRow[historyConfig.typ - 1] = "Rechnungstyp";
                headerRow[historyConfig.dateiname - 1] = "Dateiname";
                headerRow[historyConfig.dateilink - 1] = "Link zur Datei";

                history.appendRow(headerRow);
                history.getRange(1, 1, 1, 4).setFontWeight("bold");
            }

            return history;
        } catch (e) {
            console.error("Fehler bei der Initialisierung des History-Sheets:", e);
            throw e;
        }
    }

    /**
     * Sammelt bereits importierte Dateien aus der Änderungshistorie
     * @param {Sheet} history - Das Änderungshistorie-Sheet
     * @returns {Set} - Set mit bereits importierten Dateinamen
     */
    function collectExistingFiles(history) {
        const existingFiles = new Set();

        try {
            const historyData = history.getDataRange().getValues();

            // Spaltenindex für Dateinamen ermitteln (sollte Spalte C sein)
            const fileNameIndex = 2;

            // Überschriftenzeile überspringen und alle Dateinamen sammeln
            for (let i = 1; i < historyData.length; i++) {
                const fileName = historyData[i][fileNameIndex];
                if (fileName) existingFiles.add(fileName);
            }
        } catch (e) {
            console.error("Fehler beim Sammeln bereits importierter Dateien:", e);
        }

        return existingFiles;
    }

    var historyTracker = {
        getOrCreateHistorySheet,
        collectExistingFiles
    };

    // modules/importModule/index.js

    /**
     * Modul für den Import von Dateien aus Google Drive in die Buchhaltungstabelle
     */
    const ImportModule = {
        /**
         * Importiert Dateien aus Google Drive und aktualisiert alle Tabellenblätter
         * @param {Object} config - Die Konfiguration
         * @returns {number} - Anzahl der importierten Dateien
         */
        importDriveFiles(config) {
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
                const history = historyTracker.getOrCreateHistorySheet(ss, config);

                // Auf übergeordneten Ordner zugreifen
                const parentFolder = fileManager.getParentFolder();
                if (!parentFolder) {
                    ui.alert("Fehler: Kein übergeordneter Ordner gefunden.");
                    return 0;
                }

                // Unterordner für Einnahmen, Ausgaben und Eigenbelege finden oder erstellen
                const revenueFolder = fileManager.findOrCreateFolder(parentFolder, "Einnahmen", ui);
                const expenseFolder = fileManager.findOrCreateFolder(parentFolder, "Ausgaben", ui);
                const receiptsFolder = fileManager.findOrCreateFolder(parentFolder, "Eigenbelege", ui);

                // Bereits importierte Dateien sammeln
                const existingFiles = historyTracker.collectExistingFiles(history);

                // Import durchführen wenn Ordner existieren
                let importedRevenue = 0, importedExpense = 0, importedReceipts = 0;

                if (revenueFolder) {
                    try {
                        importedRevenue = dataProcessor.importFilesFromFolder(
                            revenueFolder,
                            revenueMain,
                            "Einnahme",
                            history,
                            existingFiles,
                            config
                        );
                        totalImported += importedRevenue;
                    } catch (e) {
                        console.error("Fehler beim Import der Einnahmen:", e);
                        ui.alert("Fehler beim Import der Einnahmen: " + e.toString());
                    }
                }

                if (expenseFolder) {
                    try {
                        importedExpense = dataProcessor.importFilesFromFolder(
                            expenseFolder,
                            expenseMain,
                            "Ausgabe",
                            history,
                            existingFiles,
                            config
                        );
                        totalImported += importedExpense;
                    } catch (e) {
                        console.error("Fehler beim Import der Ausgaben:", e);
                        ui.alert("Fehler beim Import der Ausgaben: " + e.toString());
                    }
                }

                if (receiptsFolder) {
                    try {
                        importedReceipts = dataProcessor.importFilesFromFolder(
                            receiptsFolder,
                            receiptsMain,
                            "Eigenbeleg",
                            history,
                            existingFiles,
                            config
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
        }
    };

    // utils/stringUtils.js
    /**
     * Funktionen für die Verarbeitung von Strings
     */

    /**
     * Prüft, ob ein Wert leer oder undefiniert ist
     * @param {*} value - Der zu prüfende Wert
     * @returns {boolean} - true wenn der Wert leer ist
     */
    function isEmpty(value) {
        return value === null || value === undefined || value.toString().trim() === "";
    }

    /**
     * Bereinigt einen Text von Sonderzeichen und macht ihn vergleichbar
     * @param {string} text - Der zu bereinigende Text
     * @returns {string} - Der bereinigte Text
     */
    function normalizeText(text) {
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
    }

    /**
     * Konvertiert einen Spaltenindex (1-basiert) in einen Spaltenbuchstaben (A, B, C, ...)
     * @param {number} columnIndex - 1-basierter Spaltenindex
     * @returns {string} - Spaltenbuchstabe(n)
     */
    function getColumnLetter(columnIndex) {
        let letter = '';
        let colIndex = columnIndex;

        while (colIndex > 0) {
            const modulo = (colIndex - 1) % 26;
            letter = String.fromCharCode(65 + modulo) + letter;
            colIndex = Math.floor((colIndex - modulo) / 26);
        }

        return letter;
    }

    /**
     * Generiert eine eindeutige ID (für Referenzzwecke)
     * @param {string} prefix - Optional ein Präfix für die ID
     * @returns {string} - Eine eindeutige ID
     */
    function generateUniqueId(prefix = "") {
        const timestamp = new Date().getTime();
        const random = Math.floor(Math.random() * 10000);
        return `${prefix}${timestamp}${random}`;
    }

    var stringUtils = {
        isEmpty,
        normalizeText,
        getColumnLetter,
        generateUniqueId
    };

    // modules/refreshModule/formattingHandler.js

    /**
     * Setzt bedingte Formatierung für eine Statusspalte
     * @param {Sheet} sheet - Das Sheet
     * @param {string} column - Spaltenbezeichnung (z.B. "L")
     * @param {Array<Object>} conditions - Bedingungen (value, background, fontColor)
     */
    function setConditionalFormattingForStatusColumn(sheet, column, conditions) {
        sheetUtils$1.setConditionalFormattingForColumn(sheet, column, conditions);
    }

    /**
     * Setzt Dropdown-Validierungen für ein Sheet
     * @param {Sheet} sheet - Das Sheet
     * @param {string} sheetName - Name des Sheets
     * @param {number} numRows - Anzahl der Datenzeilen
     * @param {Object} columns - Spaltenkonfiguration
     * @param {Object} config - Die Konfiguration
     */
    function setDropdownValidations(sheet, sheetName, numRows, columns, config) {
        // Validator-Funktion lokal implementieren, um Abhängigkeit zu vermeiden
        function validateDropdown(sheet, row, col, numRows, numCols, list) {
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
        }

        if (sheetName === "Einnahmen") {
            validateDropdown(
                sheet, 2, columns.kategorie, numRows, 1,
                Object.keys(config.einnahmen.categories)
            );
        } else if (sheetName === "Ausgaben") {
            validateDropdown(
                sheet, 2, columns.kategorie, numRows, 1,
                Object.keys(config.ausgaben.categories)
            );
        } else if (sheetName === "Eigenbelege") {
            // Kategorie-Dropdown für Eigenbelege
            validateDropdown(
                sheet, 2, columns.kategorie, numRows, 1,
                Object.keys(config.eigenbelege.categories)
            );

            // Dropdown für "Ausgelegt von" hinzufügen (merged aus shareholders und employees)
            const ausleger = [
                ...config.common.shareholders,
                ...config.common.employees
            ];

            validateDropdown(
                sheet, 2, columns.ausgelegtVon, numRows, 1,
                ausleger
            );
        } else if (sheetName === "Gesellschafterkonto") {
            // Kategorie-Dropdown für Gesellschafterkonto
            validateDropdown(
                sheet, 2, columns.kategorie, numRows, 1,
                Object.keys(config.gesellschafterkonto.categories)
            );

            // Dropdown für Gesellschafter
            validateDropdown(
                sheet, 2, columns.gesellschafter, numRows, 1,
                config.common.shareholders
            );
        } else if (sheetName === "Holding Transfers") {
            // Art-Dropdown für Holding Transfers
            validateDropdown(
                sheet, 2, columns.art, numRows, 1,
                Object.keys(config.holdingTransfers.categories)
            );
        }

        // Zahlungsart-Dropdown für alle Blätter mit Zahlungsart-Spalte
        if (columns.zahlungsart) {
            validateDropdown(
                sheet, 2, columns.zahlungsart, numRows, 1,
                config.common.paymentType
            );
        }
    }

    /**
     * Wendet Validierungen auf das Bankbewegungen-Sheet an
     * @param {Sheet} sheet - Das Sheet
     * @param {number} firstDataRow - Erste Datenzeile
     * @param {number} numDataRows - Anzahl der Datenzeilen
     * @param {Object} columns - Spaltenkonfiguration
     * @param {Object} config - Die Konfiguration
     */
    function applyBankSheetValidations(sheet, firstDataRow, numDataRows, columns, config) {
        // Validator-Funktion lokal implementieren, um Abhängigkeit zu vermeiden
        function validateDropdown(sheet, row, col, numRows, numCols, list) {
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
        }

        // Validierung für Transaktionstyp
        validateDropdown(
            sheet, firstDataRow, columns.transaktionstyp, numDataRows, 1,
            config.bankbewegungen.types
        );

        // Validierung für Kategorie
        validateDropdown(
            sheet, firstDataRow, columns.kategorie, numDataRows, 1,
            config.bankbewegungen.categories
        );

        // Konten für Dropdown-Validierung sammeln
        const allowedKontoSoll = new Set();
        const allowedGegenkonto = new Set();

        // Konten aus allen relevanten Mappings sammeln
        [
            config.einnahmen.kontoMapping,
            config.ausgaben.kontoMapping,
            config.eigenbelege.kontoMapping,
            config.gesellschafterkonto.kontoMapping,
            config.holdingTransfers.kontoMapping
        ].forEach(mapping => {
            Object.values(mapping).forEach(m => {
                if (m.soll) allowedKontoSoll.add(m.soll);
                if (m.gegen) allowedGegenkonto.add(m.gegen);
            });
        });

        // Dropdown-Validierungen für Konten setzen
        validateDropdown(
            sheet, firstDataRow, columns.kontoSoll, numDataRows, 1,
            Array.from(allowedKontoSoll)
        );

        validateDropdown(
            sheet, firstDataRow, columns.kontoHaben, numDataRows, 1,
            Array.from(allowedGegenkonto)
        );
    }

    var formattingHandler = {
        setConditionalFormattingForStatusColumn,
        setDropdownValidations,
        applyBankSheetValidations
    };

    // modules/refreshModule/dataSheetHandler.js

    /**
     * Aktualisiert ein Datenblatt (Einnahmen, Ausgaben, Eigenbelege)
     * @param {Sheet} sheet - Das zu aktualisierende Sheet
     * @param {string} sheetName - Name des Sheets
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - true bei erfolgreicher Aktualisierung
     */
    function refreshDataSheet(sheet, sheetName, config) {
        try {
            const lastRow = sheet.getLastRow();
            if (lastRow < 2) return true; // Keine Daten zum Aktualisieren

            const numRows = lastRow - 1;

            // Passende Spaltenkonfiguration für das entsprechende Sheet auswählen
            let columns;
            if (sheetName === "Einnahmen") {
                columns = config.einnahmen.columns;
            } else if (sheetName === "Ausgaben") {
                columns = config.ausgaben.columns;
            } else if (sheetName === "Eigenbelege") {
                columns = config.eigenbelege.columns;
            } else if (sheetName === "Gesellschafterkonto") {
                columns = config.gesellschafterkonto.columns;
            } else if (sheetName === "Holding Transfers") {
                columns = config.holdingTransfers.columns;
            } else {
                return false; // Unbekanntes Sheet
            }

            // Spaltenbuchstaben aus den Indizes generieren
            const columnLetters = {};
            Object.entries(columns).forEach(([key, index]) => {
                columnLetters[key] = stringUtils.getColumnLetter(index);
            });

            // Batch-Array für Formeln erstellen (effizienter als einzelne Range-Updates)
            const formulasBatch = {};

            // Standardformeln setzen
            setStandardFormulas(formulasBatch, columnLetters, numRows, sheetName);

            // Spezifische Formeln für bestimmte Sheet-Typen
            if (sheetName === "Gesellschafterkonto") {
                // Hier später spezifische Formeln für Gesellschafterkonto ergänzen
            } else if (sheetName === "Holding Transfers") {
                // Hier später spezifische Formeln für Holding Transfers ergänzen
            }

            // Formeln in Batches anwenden (weniger API-Calls)
            Object.entries(formulasBatch).forEach(([col, formulas]) => {
                sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulas);
            });

            // Bezahlter Betrag - Leerzeichen durch 0 ersetzen für Berechnungen
            if (columns.bezahlt) {
                const bezahltRange = sheet.getRange(2, columns.bezahlt, numRows, 1);
                const bezahltValues = bezahltRange.getValues();
                const updatedBezahltValues = bezahltValues.map(
                    ([val]) => [stringUtils.isEmpty(val) ? 0 : val]
                );
                bezahltRange.setValues(updatedBezahltValues);
            }

            // Dropdown-Validierungen je nach Sheet-Typ setzen
            formattingHandler.setDropdownValidations(sheet, sheetName, numRows, columns, config);

            // Bedingte Formatierung für Status-Spalte
            if (columns.zahlungsstatus) {
                setStatusFormatting(sheet, sheetName, columnLetters.zahlungsstatus);
            }

            // Spaltenbreiten automatisch anpassen
            sheet.autoResizeColumns(1, sheet.getLastColumn());

            return true;
        } catch (e) {
            console.error(`Fehler beim Aktualisieren von ${sheetName}:`, e);
            return false;
        }
    }

    /**
     * Setzt Standardformeln für ein Datenblatt
     * @param {Object} formulasBatch - Objekt für Batch-Formeln
     * @param {Object} columnLetters - Spaltenbuchstaben
     * @param {number} numRows - Anzahl der Datenzeilen
     * @param {string} sheetName - Name des Sheets
     */
    function setStandardFormulas(formulasBatch, columnLetters, numRows, sheetName) {
        // Prüfen, ob die benötigten Spalten vorhanden sind
        const hasRequiredColumns = (
            columnLetters.mwstBetrag &&
            columnLetters.nettobetrag &&
            columnLetters.mwstSatz &&
            columnLetters.bruttoBetrag &&
            columnLetters.restbetragNetto &&
            columnLetters.bezahlt &&
            columnLetters.datum &&
            columnLetters.quartal
        );

        if (!hasRequiredColumns) return;

        // MwSt-Betrag
        formulasBatch[columnLetters.mwstBetrag] = Array.from(
            {length: numRows},
            (_, i) => [`=${columnLetters.nettobetrag}${i + 2}*${columnLetters.mwstSatz}${i + 2}/100`]
        );

        // Brutto-Betrag
        formulasBatch[columnLetters.bruttoBetrag] = Array.from(
            {length: numRows},
            (_, i) => [`=${columnLetters.nettobetrag}${i + 2}+${columnLetters.mwstBetrag}${i + 2}`]
        );

        // Bezahlter Betrag - für Teilzahlungen
        formulasBatch[columnLetters.restbetragNetto] = Array.from(
            {length: numRows},
            (_, i) => [`=(${columnLetters.bruttoBetrag}${i + 2}-${columnLetters.bezahlt}${i + 2})/(1+${columnLetters.mwstSatz}${i + 2}/100)`]
        );

        // Quartal
        formulasBatch[columnLetters.quartal] = Array.from(
            {length: numRows},
            (_, i) => [`=IF(${columnLetters.datum}${i + 2}="";"";ROUNDUP(MONTH(${columnLetters.datum}${i + 2})/3;0))`]
        );

        // Zahlungsstatus
        if (columnLetters.zahlungsstatus) {
            if (sheetName !== "Eigenbelege") {
                // Für Einnahmen und Ausgaben: Zahlungsstatus
                formulasBatch[columnLetters.zahlungsstatus] = Array.from(
                    {length: numRows},
                    (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Bezahlt";"Teilbezahlt"))`]
                );
            } else {
                // Für Eigenbelege: Zahlungsstatus
                formulasBatch[columnLetters.zahlungsstatus] = Array.from(
                    {length: numRows},
                    (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Erstattet";"Teilerstattet"))`]
                );
            }
        }
    }

    /**
     * Setzt bedingte Formatierung für die Status-Spalte
     * @param {Sheet} sheet - Das Sheet
     * @param {string} sheetName - Name des Sheets
     * @param {string} statusColumn - Statusspaltenbuchstabe
     */
    function setStatusFormatting(sheet, sheetName, statusColumn) {
        if (sheetName !== "Eigenbelege") {
            // Für Einnahmen und Ausgaben: Zahlungsstatus
            formattingHandler.setConditionalFormattingForStatusColumn(sheet, statusColumn, [
                {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
                {value: "Teilbezahlt", background: "#FFEB9C", fontColor: "#9C6500"},
                {value: "Bezahlt", background: "#C6EFCE", fontColor: "#006100"}
            ]);
        } else {
            // Für Eigenbelege: Status
            formattingHandler.setConditionalFormattingForStatusColumn(sheet, statusColumn, [
                {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
                {value: "Teilerstattet", background: "#FFEB9C", fontColor: "#9C6500"},
                {value: "Erstattet", background: "#C6EFCE", fontColor: "#006100"}
            ]);
        }
    }

    var dataSheetHandler = {
        refreshDataSheet
    };

    // utils/numberUtils.js
    /**
     * Funktionen für die Verarbeitung und Formatierung von Zahlen und Währungen
     */

    /**
     * Konvertiert einen String oder eine Zahl in einen numerischen Währungswert
     * @param {number|string} value - Der zu parsende Wert
     * @returns {number} - Der geparste Währungswert oder 0 bei ungültigem Format
     */
    function parseCurrency(value) {
        if (value === null || value === undefined || value === "") return 0;
        if (typeof value === "number") return value;

        // Entferne alle Zeichen außer Ziffern, Komma, Punkt und Minus
        const str = value.toString()
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

        return isNaN(result) ? 0 : result;
    }

    /**
     * Parst einen MwSt-Satz und normalisiert ihn
     * @param {number|string} value - Der zu parsende MwSt-Satz
     * @param {number} defaultMwst - Standard-MwSt-Satz, falls value ungültig ist
     * @returns {number} - Der normalisierte MwSt-Satz (0-100)
     */
    function parseMwstRate(value, defaultMwst = 19) {
        if (value === null || value === undefined || value === "") {
            return defaultMwst;
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

        return result;
    }

    /**
     * Formatiert einen Währungsbetrag im deutschen Format
     * @param {number|string} amount - Der zu formatierende Betrag
     * @param {string} currency - Das Währungssymbol (Standard: "€")
     * @returns {string} - Der formatierte Betrag
     */
    function formatCurrency(amount, currency = "€") {
        const value = parseCurrency(amount);
        return value.toLocaleString('de-DE', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }) + " " + currency;
    }

    /**
     * Prüft, ob zwei Zahlenwerte im Rahmen einer bestimmten Toleranz gleich sind
     * @param {number} a - Erster Wert
     * @param {number} b - Zweiter Wert
     * @param {number} tolerance - Toleranzwert (Standard: 0.01)
     * @returns {boolean} - true wenn Werte innerhalb der Toleranz gleich sind
     */
    function isApproximatelyEqual(a, b, tolerance = 0.01) {
        return Math.abs(a - b) <= tolerance;
    }

    /**
     * Sicheres Runden eines Werts auf n Dezimalstellen
     * @param {number} value - Der zu rundende Wert
     * @param {number} decimals - Anzahl der Dezimalstellen (Standard: 2)
     * @returns {number} - Gerundeter Wert
     */
    function round(value, decimals = 2) {
        const factor = Math.pow(10, decimals);
        return Math.round((value + Number.EPSILON) * factor) / factor;
    }

    /**
     * Prüft, ob ein Wert keine gültige Zahl ist
     * @param {*} v - Der zu prüfende Wert
     * @returns {boolean} - True, wenn der Wert keine gültige Zahl ist
     */
    function isInvalidNumber(v) {
        if (v === null || v === undefined || v === "") return true;
        return isNaN(parseFloat(v.toString().trim()));
    }

    var numberUtils$1 = {
        parseCurrency,
        parseMwstRate,
        formatCurrency,
        isApproximatelyEqual,
        round,
        isInvalidNumber
    };

    // utils/cacheUtils.js
    /**
     * Funktionen für Caching zur Verbesserung der Performance
     */

    /**
     * Cache-Objekt für verschiedene Datentypen
     */
    const globalCache = {
        // Verschiedene Cache-Maps für unterschiedliche Datentypen
        _store: {
            dates: new Map(),
            currency: new Map(),
            mwstRates: new Map(),
            columnLetters: new Map(),
            sheets: new Map(),
            references: {
                einnahmen: null,
                ausgaben: null,
                eigenbelege: null,
                gesellschafterkonto: null,
                holdingTransfers: null
            },
            computed: new Map()
        },

        /**
         * Löscht den gesamten Cache oder einen bestimmten Teilbereich
         * @param {string} type - Optional: Der zu löschende Cache-Bereich
         */
        clear(type = null) {
            if (type === null) {
                // Alles löschen
                this._store.dates.clear();
                this._store.currency.clear();
                this._store.mwstRates.clear();
                this._store.columnLetters.clear();
                this._store.sheets.clear();
                this._store.computed.clear();
                this._store.references.einnahmen = null;
                this._store.references.ausgaben = null;
                this._store.references.eigenbelege = null;
                this._store.references.gesellschafterkonto = null;
                this._store.references.holdingTransfers = null;
            } else if (type === 'references') {
                this._store.references.einnahmen = null;
                this._store.references.ausgaben = null;
                this._store.references.eigenbelege = null;
                this._store.references.gesellschafterkonto = null;
                this._store.references.holdingTransfers = null;
            } else if (this._store[type]) {
                this._store[type].clear();
            }
        },

        /**
         * Speichert einen Wert im Cache
         * @param {string} type - Cache-Typ (dates, currency, mwstRates, columnLetters, computed)
         * @param {string} key - Schlüssel für den gecacheten Wert
         * @param {*} value - Zu cachender Wert
         */
        set(type, key, value) {
            if (this._store[type] instanceof Map) {
                this._store[type].set(key, value);
            } else if (type === 'references' && key in this._store.references) {
                this._store.references[key] = value;
            }
        },

        /**
         * Prüft ob ein Wert im Cache vorhanden ist
         * @param {string} type - Cache-Typ
         * @param {string} key - Schlüssel des zu prüfenden Werts
         * @returns {boolean} true wenn der Wert im Cache vorhanden ist
         */
        has(type, key) {
            if (this._store[type] instanceof Map) {
                return this._store[type].has(key);
            } else if (type === 'references') {
                return this._store.references[key] !== null;
            }
            return false;
        },

        /**
         * Ruft einen Wert aus dem Cache ab
         * @param {string} type - Cache-Typ
         * @param {string} key - Schlüssel des abzurufenden Werts
         * @returns {*} Der gecachete Wert oder undefined wenn nicht vorhanden
         */
        get(type, key) {
            if (this._store[type] instanceof Map) {
                return this._store[type].get(key);
            } else if (type === 'references') {
                return this._store.references[key];
            }
            return undefined;
        },

        /**
         * Ruft einen Wert aus dem Cache ab oder berechnet ihn, wenn nicht vorhanden
         * @param {string} type - Cache-Typ
         * @param {string} key - Schlüssel des abzurufenden Werts
         * @param {Function} computeFn - Funktion zur Berechnung des Werts, wenn nicht im Cache
         * @returns {*} Der gecachete oder berechnete Wert
         */
        getOrCompute(type, key, computeFn) {
            if (this.has(type, key)) {
                return this.get(type, key);
            }

            const value = computeFn();
            this.set(type, key, value);
            return value;
        }
    };

    // modules/refreshModule/matchingHandler.js

    /**
     * Führt das Matching von Bankbewegungen mit Rechnungen durch
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Sheet} sheet - Das Bankbewegungen-Sheet
     * @param {number} firstDataRow - Erste Datenzeile
     * @param {number} numDataRows - Anzahl der Datenzeilen
     * @param {number} lastRow - Letzte Zeile
     * @param {Object} columns - Spaltenkonfiguration
     * @param {Object} columnLetters - Buchstaben für die Spalten
     * @param {Object} config - Die Konfiguration
     */
    function performBankReferenceMatching(ss, sheet, firstDataRow, numDataRows, lastRow, columns, columnLetters, config) {
        // Konfigurationen für Spaltenindizes
        const einnahmenCols = config.einnahmen.columns;
        const ausgabenCols = config.ausgaben.columns;
        const eigenbelegeCols = config.eigenbelege.columns;
        const gesellschafterCols = config.gesellschafterkonto.columns;
        const holdingCols = config.holdingTransfers.columns;

        // Referenzdaten laden für alle relevanten Sheets
        const sheetData = loadReferenceData(ss, {
            einnahmen: einnahmenCols,
            ausgaben: ausgabenCols,
            eigenbelege: eigenbelegeCols,
            gesellschafterkonto: gesellschafterCols,
            holdingTransfers: holdingCols
        });

        // Bankbewegungen Daten für Verarbeitung holen
        const bankData = sheet.getRange(firstDataRow, 1, numDataRows, columns.matchInfo).getDisplayValues();

        // Ergebnis-Arrays für Batch-Update
        const matchResults = [];
        const kontoSollResults = [];
        const kontoHabenResults = [];

        // Banking-Zuordnungen für spätere Synchronisierung mit anderen Sheets
        const bankZuordnungen = {};

        // Sammeln aller gültigen Konten für die Validierung
        const allowedKontoSoll = new Set();
        const allowedGegenkonto = new Set();

        // Konten aus allen relevanten Mappings sammeln
        [
            config.einnahmen.kontoMapping,
            config.ausgaben.kontoMapping,
            config.eigenbelege.kontoMapping,
            config.gesellschafterkonto.kontoMapping,
            config.holdingTransfers.kontoMapping
        ].forEach(mapping => {
            Object.values(mapping).forEach(m => {
                if (m.soll) allowedKontoSoll.add(m.soll);
                if (m.gegen) allowedGegenkonto.add(m.gegen);
            });
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

            // Matching-Logik für Bankbewegungen
            if (!stringUtils.isEmpty(refNumber)) {
                let matchFound = false;
                const betragValue = Math.abs(numberUtils$1.parseCurrency(row[columns.betrag - 1]));

                // Je nach Transaktionstyp in unterschiedlichen Sheets suchen
                if (tranType === "Einnahme") {
                    // Einnahmen prüfen
                    processEinnahmeMatching(
                        refNumber, betragValue, row, columns, sheetData.einnahmen,
                        einnahmenCols, matchResults, kontoSollResults, kontoHabenResults,
                        bankZuordnungen, matchFound, config
                    );
                } else if (tranType === "Ausgabe") {
                    // Ausgaben, Eigenbelege, Gesellschafterkonto und Holding Transfers prüfen
                    processAusgabeMatching(
                        refNumber, betragValue, row, columns, {
                            ausgaben: sheetData.ausgaben,
                            eigenbelege: sheetData.eigenbelege,
                            gesellschafterkonto: sheetData.gesellschafterkonto,
                            holdingTransfers: sheetData.holdingTransfers
                        },
                        {
                            ausgabenCols, eigenbelegeCols, gesellschafterCols, holdingCols
                        },
                        matchResults, kontoSollResults, kontoHabenResults,
                        bankZuordnungen, matchFound, config
                    );
                }
            }

            // Kontonummern basierend auf Kategorie setzen
            if (category) {
                // Den richtigen Mapping-Typ basierend auf der Transaktionsart auswählen
                const mappingSource = tranType === "Einnahme" ?
                    config.einnahmen.kontoMapping :
                    config.ausgaben.kontoMapping;

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

            // Sicherstellen, dass die Arrays die richtigen Werte haben
            if (matchResults.length <= i) {
                matchResults.push([matchInfo]);
            }

            if (kontoSollResults.length <= i) {
                kontoSollResults.push([kontoSoll]);
            }

            if (kontoHabenResults.length <= i) {
                kontoHabenResults.push([kontoHaben]);
            }
        }

        // Batch-Updates ausführen für bessere Performance
        sheet.getRange(firstDataRow, columns.matchInfo, numDataRows, 1).setValues(matchResults);
        sheet.getRange(firstDataRow, columns.kontoSoll, numDataRows, 1).setValues(kontoSollResults);
        sheet.getRange(firstDataRow, columns.kontoHaben, numDataRows, 1).setValues(kontoHabenResults);

        // Formatiere die gesamten Zeilen basierend auf dem Match-Typ
        formatMatchedRows(sheet, firstDataRow, matchResults, columns);

        // Bedingte Formatierung für Match-Spalte mit verbesserten Farben
        setMatchColumnFormatting(sheet, columnLetters.matchInfo);

        // Setze farbliche Markierung in den Einnahmen/Ausgaben/Eigenbelege Sheets basierend auf Zahlungsstatus
        markPaidInvoices(ss, bankZuordnungen, config);
    }

    /**
     * Lädt Referenzdaten aus allen relevanten Sheets
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Object} columns - Spaltenkonfigurationen für alle Sheets
     * @returns {Object} - Referenzdaten für alle Sheets
     */
    function loadReferenceData(ss, columns) {
        const result = {};

        // Sheets und ihre Konfigurationen
        const sheets = {
            einnahmen: { name: "Einnahmen", cols: columns.einnahmen },
            ausgaben: { name: "Ausgaben", cols: columns.ausgaben },
            eigenbelege: { name: "Eigenbelege", cols: columns.eigenbelege },
            gesellschafterkonto: { name: "Gesellschafterkonto", cols: columns.gesellschafterkonto },
            holdingTransfers: { name: "Holding Transfers", cols: columns.holdingTransfers }
        };

        // Für jedes Sheet Referenzdaten laden
        Object.entries(sheets).forEach(([key, config]) => {
            const sheet = ss.getSheetByName(config.name);

            if (sheet && sheet.getLastRow() > 1) {
                result[key] = loadSheetReferenceData(sheet, config.cols, key);
            } else {
                result[key] = {};
            }
        });

        return result;
    }

    /**
     * Lädt Referenzdaten aus einem einzelnen Sheet
     * @param {Sheet} sheet - Das Sheet
     * @param {Object} columns - Spaltenkonfiguration
     * @param {string} sheetType - Typ des Sheets
     * @returns {Object} - Referenzdaten für das Sheet
     */
    function loadSheetReferenceData(sheet, columns, sheetType) {
        // Cache prüfen
        if (globalCache.has('references', sheetType)) {
            return globalCache.get('references', sheetType);
        }

        const map = {};
        const numRows = sheet.getLastRow() - 1;

        if (numRows <= 0) return map;

        // Die benötigten Spalten bestimmen
        const columnsToLoad = Math.max(
            columns.rechnungsnummer,
            columns.nettobetrag,
            columns.mwstSatz || 0,
            columns.bezahlt || 0
        );

        // Die relevanten Spalten laden
        const data = sheet.getRange(2, 1, numRows, columnsToLoad).getValues();

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const ref = row[columns.rechnungsnummer - 1];

            if (stringUtils.isEmpty(ref)) continue;

            // Entferne "G-" Prefix für den Key, falls vorhanden (für Gutschriften)
            let key = ref.toString().trim();
            const isGutschrift = key.startsWith("G-");

            if (isGutschrift) {
                key = key.substring(2);
            }

            const normalizedKey = stringUtils.normalizeText(key);

            // Werte extrahieren
            let betrag = 0;
            if (!stringUtils.isEmpty(row[columns.nettobetrag - 1])) {
                betrag = numberUtils$1.parseCurrency(row[columns.nettobetrag - 1]);
                betrag = Math.abs(betrag);
            }

            let mwstRate = 0;
            if (columns.mwstSatz && !stringUtils.isEmpty(row[columns.mwstSatz - 1])) {
                mwstRate = numberUtils$1.parseMwstRate(row[columns.mwstSatz - 1]);
            }

            let bezahlt = 0;
            if (columns.bezahlt && !stringUtils.isEmpty(row[columns.bezahlt - 1])) {
                bezahlt = numberUtils$1.parseCurrency(row[columns.bezahlt - 1]);
                bezahlt = Math.abs(bezahlt);
            }

            // Bruttobetrag berechnen
            const brutto = betrag * (1 + mwstRate/100);

            map[key] = {
                ref: ref.toString().trim(),
                normalizedRef: normalizedKey,
                row: i + 2, // 1-basiert + Überschrift
                betrag: betrag,
                mwstRate: mwstRate,
                brutto: brutto,
                bezahlt: bezahlt,
                offen: numberUtils$1.round(brutto - bezahlt, 2),
                isGutschrift: isGutschrift
            };

            // Normalisierter Key als Alternative
            if (normalizedKey !== key && !map[normalizedKey]) {
                map[normalizedKey] = map[key];
            }
        }

        // Referenzdaten cachen
        globalCache.set('references', sheetType, map);

        return map;
    }

    /**
     * Verarbeitet Matching für Einnahmen
     * @param {string} refNumber - Referenznummer
     * @param {number} betragValue - Betrag
     * @param {Array} row - Zeile aus dem Bankbewegungen-Sheet
     * @param {Object} columns - Spaltenkonfiguration
     * @param {Object} einnahmenData - Referenzdaten für Einnahmen
     * @param {Object} einnahmenCols - Spaltenkonfiguration für Einnahmen
     * @param {Array} matchResults - Array für Match-Ergebnisse
     * @param {Array} kontoSollResults - Array für Konto-Soll-Ergebnisse
     * @param {Array} kontoHabenResults - Array für Konto-Haben-Ergebnisse
     * @param {Object} bankZuordnungen - Speicher für Bankzuordnungen
     * @param {boolean} matchFound - Flag für gefundene Übereinstimmung
     * @param {Object} config - Die Konfiguration
     */
    function processEinnahmeMatching(refNumber, betragValue, row, columns, einnahmenData, einnahmenCols, matchResults, kontoSollResults, kontoHabenResults, bankZuordnungen, matchFound, config) {
        // Einnahmen-Match prüfen
        const matchResult = findMatch(refNumber, einnahmenData, betragValue);

        if (matchResult) {
            const matchInfo = processEinnahmeMatch(matchResult);
            matchResults.push([matchInfo]);

            // Konten aus dem Mapping setzen
            const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";
            let kontoSoll = "", kontoHaben = "";

            if (category && config.einnahmen.kontoMapping[category]) {
                const mapping = config.einnahmen.kontoMapping[category];
                kontoSoll = mapping.soll || "";
                kontoHaben = mapping.gegen || "";
            }

            kontoSollResults.push([kontoSoll]);
            kontoHabenResults.push([kontoHaben]);

            // Für spätere Markierung merken
            const key = `einnahme#${matchResult.row}`;
            bankZuordnungen[key] = {
                typ: "einnahme",
                row: matchResult.row,
                bankDatum: row[columns.datum - 1],
                matchInfo: matchInfo,
                transTyp: "Einnahme"
            };
        }
    }

    /**
     * Verarbeitet Matching für Ausgaben und andere Belege
     * @param {string} refNumber - Referenznummer
     * @param {number} betragValue - Betrag
     * @param {Array} row - Zeile aus dem Bankbewegungen-Sheet
     * @param {Object} columns - Spaltenkonfiguration
     * @param {Object} sheetData - Referenzdaten für verschiedene Sheets
     * @param {Object} sheetCols - Spaltenkonfigurationen für verschiedene Sheets
     * @param {Array} matchResults - Array für Match-Ergebnisse
     * @param {Array} kontoSollResults - Array für Konto-Soll-Ergebnisse
     * @param {Array} kontoHabenResults - Array für Konto-Haben-Ergebnisse
     * @param {Object} bankZuordnungen - Speicher für Bankzuordnungen
     * @param {boolean} matchFound - Flag für gefundene Übereinstimmung
     * @param {Object} config - Die Konfiguration
     */
    function processAusgabeMatching(refNumber, betragValue, row, columns, sheetData, sheetCols, matchResults, kontoSollResults, kontoHabenResults, bankZuordnungen, matchFound, config) {
        // Ausgaben-Match prüfen
        const ausgabenMatch = findMatch(refNumber, sheetData.ausgaben, betragValue);
        if (ausgabenMatch) {
            matchFound = true;
            const matchInfo = processAusgabeMatch(ausgabenMatch, betragValue, row, columns, sheetCols.ausgabenCols);
            matchResults.push([matchInfo]);

            // Konten aus dem Mapping setzen
            const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";
            let kontoSoll = "", kontoHaben = "";

            if (category && config.ausgaben.kontoMapping[category]) {
                const mapping = config.ausgaben.kontoMapping[category];
                kontoSoll = mapping.soll || "";
                kontoHaben = mapping.gegen || "";
            }

            kontoSollResults.push([kontoSoll]);
            kontoHabenResults.push([kontoHaben]);

            // Für spätere Markierung merken
            const key = `ausgabe#${ausgabenMatch.row}`;
            bankZuordnungen[key] = {
                typ: "ausgabe",
                row: ausgabenMatch.row,
                bankDatum: row[columns.datum - 1],
                matchInfo: matchInfo,
                transTyp: "Ausgabe"
            };
            return;
        }

        // Wenn keine Übereinstimmung in Ausgaben, dann in Eigenbelegen suchen
        if (!matchFound) {
            const eigenbelegMatch = findMatch(refNumber, sheetData.eigenbelege, betragValue);
            if (eigenbelegMatch) {
                matchFound = true;
                const matchInfo = processEigenbelegMatch(eigenbelegMatch, betragValue, row, columns, sheetCols.eigenbelegeCols);
                matchResults.push([matchInfo]);

                // Konten aus dem Mapping setzen
                const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";
                let kontoSoll = "", kontoHaben = "";

                if (category && config.eigenbelege.kontoMapping[category]) {
                    const mapping = config.eigenbelege.kontoMapping[category];
                    kontoSoll = mapping.soll || "";
                    kontoHaben = mapping.gegen || "";
                }

                kontoSollResults.push([kontoSoll]);
                kontoHabenResults.push([kontoHaben]);

                // Für spätere Markierung merken
                const key = `eigenbeleg#${eigenbelegMatch.row}`;
                bankZuordnungen[key] = {
                    typ: "eigenbeleg",
                    row: eigenbelegMatch.row,
                    bankDatum: row[columns.datum - 1],
                    matchInfo: matchInfo,
                    transTyp: "Ausgabe"
                };
                return;
            }
        }

        // Wenn keine Übereinstimmung, auch in Gesellschafterkonto suchen
        if (!matchFound) {
            const gesellschafterMatch = findMatch(refNumber, sheetData.gesellschafterkonto, betragValue);
            if (gesellschafterMatch) {
                matchFound = true;
                const matchInfo = processGesellschafterMatch(gesellschafterMatch, betragValue, row, columns, sheetCols.gesellschafterCols);
                matchResults.push([matchInfo]);

                // Konten aus dem Mapping setzen
                const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";
                let kontoSoll = "", kontoHaben = "";

                if (category && config.gesellschafterkonto.kontoMapping[category]) {
                    const mapping = config.gesellschafterkonto.kontoMapping[category];
                    kontoSoll = mapping.soll || "";
                    kontoHaben = mapping.gegen || "";
                }

                kontoSollResults.push([kontoSoll]);
                kontoHabenResults.push([kontoHaben]);

                // Für spätere Markierung merken
                const key = `gesellschafterkonto#${gesellschafterMatch.row}`;
                bankZuordnungen[key] = {
                    typ: "gesellschafterkonto",
                    row: gesellschafterMatch.row,
                    bankDatum: row[columns.datum - 1],
                    matchInfo: matchInfo,
                    transTyp: "Ausgabe"
                };
                return;
            }
        }

        // Wenn keine Übereinstimmung, auch in Holding Transfers suchen
        if (!matchFound) {
            const holdingMatch = findMatch(refNumber, sheetData.holdingTransfers, betragValue);
            if (holdingMatch) {
                matchFound = true;
                const matchInfo = processHoldingMatch(holdingMatch, betragValue, row, columns, sheetCols.holdingCols);
                matchResults.push([matchInfo]);

                // Konten aus dem Mapping setzen
                const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";
                let kontoSoll = "", kontoHaben = "";

                if (category && config.holdingTransfers.kontoMapping[category]) {
                    const mapping = config.holdingTransfers.kontoMapping[category];
                    kontoSoll = mapping.soll || "";
                    kontoHaben = mapping.gegen || "";
                }

                kontoSollResults.push([kontoSoll]);
                kontoHabenResults.push([kontoHaben]);

                // Für spätere Markierung merken
                const key = `holdingtransfer#${holdingMatch.row}`;
                bankZuordnungen[key] = {
                    typ: "holdingtransfer",
                    row: holdingMatch.row,
                    bankDatum: row[columns.datum - 1],
                    matchInfo: matchInfo,
                    transTyp: "Ausgabe"
                };
                return;
            }
        }

        // FALLS keine Übereinstimmung, auch in Einnahmen suchen (für Gutschriften)
        if (!matchFound) {
            const gutschriftMatch = findMatch(refNumber, sheetData.einnahmen);
            if (gutschriftMatch) {
                matchFound = true;
                const matchInfo = processGutschriftMatch(gutschriftMatch, betragValue, row, columns, sheetCols.einnahmenCols);
                matchResults.push([matchInfo]);

                // Konten aus dem Mapping setzen
                const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";
                let kontoSoll = "", kontoHaben = "";

                if (category && config.einnahmen.kontoMapping[category]) {
                    const mapping = config.einnahmen.kontoMapping[category];
                    kontoSoll = mapping.soll || "";
                    kontoHaben = mapping.gegen || "";
                }

                kontoSollResults.push([kontoSoll]);
                kontoHabenResults.push([kontoHaben]);

                // Für spätere Markierung merken
                const key = `einnahme#${gutschriftMatch.row}`;
                bankZuordnungen[key] = {
                    typ: "einnahme",
                    row: gutschriftMatch.row,
                    bankDatum: row[columns.datum - 1],
                    matchInfo: matchInfo,
                    transTyp: "Gutschrift"
                };
                return;
            }
        }

        // Wenn keine Übereinstimmung gefunden wurde, leere Ergebnisse hinzufügen
        if (!matchFound) {
            matchResults.push([""]);
            kontoSollResults.push([""]);
            kontoHabenResults.push([""]);
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
        // Keine Daten
        if (!reference || !refMap) return null;

        // Normalisierte Suche vorbereiten
        const normalizedRef = stringUtils.normalizeText(reference);

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

        // 3. Teilweise Übereinstimmung
        const candidateKeys = Object.keys(refMap);

        // Zuerst prüfen wir, ob die Referenz in einem Schlüssel enthalten ist
        for (const key of candidateKeys) {
            if (key.includes(reference) || reference.includes(key)) {
                return evaluateMatch(refMap[key], betrag);
            }
        }

        // Falls keine exakten Treffer, probieren wir es mit normalisierten Werten
        for (const key of candidateKeys) {
            const normalizedKey = stringUtils.normalizeText(key);
            if (normalizedKey.includes(normalizedRef) || normalizedRef.includes(normalizedKey)) {
                return evaluateMatch(refMap[key], betrag);
            }
        }

        return null;
    }

    /**
     * Bewertet die Qualität einer Übereinstimmung basierend auf Beträgen
     * @param {Object} match - Die gefundene Übereinstimmung
     * @param {number} betrag - Der Betrag zum Vergleich
     * @returns {Object} - Übereinstimmung mit zusätzlichen Infos
     */
    function evaluateMatch(match, betrag = null) {
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
        if (numberUtils$1.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
            result.matchType = "Vollständige Zahlung";
            return result;
        }

        // Fall 2: Position ist bereits vollständig bezahlt
        if (numberUtils$1.isApproximatelyEqual(matchBezahlt, matchBrutto, tolerance) && matchBezahlt > 0) {
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
        if (betrag > matchBrutto && numberUtils$1.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
            result.matchType = "Vollständige Zahlung";
            return result;
        }

        // Fall 5: Bei allen anderen Fällen (Beträge weichen stärker ab)
        result.matchType = "Unsichere Zahlung";
        result.betragsDifferenz = numberUtils$1.round(Math.abs(betrag - matchBrutto), 2);
        return result;
    }

    /**
     * Verarbeitet eine Einnahmen-Übereinstimmung
     * @param {Object} matchResult - Das Match-Ergebnis
     * @param {number} betragValue - Der Betrag
     * @param {Array} row - Die Bankbewegungszeile
     * @param {Object} columns - Die Spaltenkonfiguration
     * @param {Object} einnahmenCols - Die Spaltenkonfiguration für Einnahmen
     * @returns {string} Formatierte Match-Information
     */
    function processEinnahmeMatch(matchResult, betragValue, row, columns, einnahmenCols) {
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

        return `Einnahme #${matchResult.row}${matchStatus}`;
    }

    /**
     * Verarbeitet eine Ausgaben-Übereinstimmung
     * @param {Object} matchResult - Das Match-Ergebnis
     * @param {number} betragValue - Der Betrag
     * @param {Array} row - Die Bankbewegungszeile
     * @param {Object} columns - Die Spaltenkonfiguration
     * @param {Object} ausgabenCols - Die Spaltenkonfiguration für Ausgaben
     * @returns {string} Formatierte Match-Information
     */
    function processAusgabeMatch(matchResult, betragValue, row, columns, ausgabenCols) {
        let matchStatus = "";

        // Je nach Match-Typ unterschiedliche Statusinformationen für Ausgaben
        if (matchResult.matchType) {
            if (matchResult.matchType === "Unsichere Zahlung" && matchResult.betragsDifferenz) {
                matchStatus = ` (${matchResult.matchType}, Diff: ${matchResult.betragsDifferenz}€)`;
            } else {
                matchStatus = ` (${matchResult.matchType})`;
            }
        }

        return `Ausgabe #${matchResult.row}${matchStatus}`;
    }

    /**
     * Verarbeitet eine Eigenbeleg-Übereinstimmung
     * @param {Object} eigenbelegMatch - Das Match-Ergebnis
     * @param {number} betragValue - Der Betrag
     * @param {Array} row - Die Bankbewegungszeile
     * @param {Object} columns - Die Spaltenkonfiguration
     * @param {Object} eigenbelegeCols - Die Spaltenkonfiguration für Eigenbelege
     * @returns {string} Formatierte Match-Information
     */
    function processEigenbelegMatch(eigenbelegMatch, betragValue, row, columns, eigenbelegeCols) {
        let matchStatus = "";

        // Je nach Match-Typ unterschiedliche Statusinformationen
        if (eigenbelegMatch.matchType) {
            // Bei "Unsichere Zahlung" auch die Differenz anzeigen
            if (eigenbelegMatch.matchType === "Unsichere Zahlung" && eigenbelegMatch.betragsDifferenz) {
                matchStatus = ` (${eigenbelegMatch.matchType}, Diff: ${eigenbelegMatch.betragsDifferenz}€)`;
            } else {
                matchStatus = ` (${eigenbelegMatch.matchType})`;
            }
        }

        return `Eigenbeleg #${eigenbelegMatch.row}${matchStatus}`;
    }

    /**
     * Verarbeitet eine Gesellschafterkonto-Übereinstimmung
     * @param {Object} gesellschafterMatch - Das Match-Ergebnis
     * @param {number} betragValue - Der Betrag
     * @param {Array} row - Die Bankbewegungszeile
     * @param {Object} columns - Die Spaltenkonfiguration
     * @param {Object} gesellschafterCols - Die Spaltenkonfiguration für Gesellschafterkonto
     * @returns {string} Formatierte Match-Information
     */
    function processGesellschafterMatch(gesellschafterMatch, betragValue, row, columns, gesellschafterCols) {
        let matchStatus = "";

        // Je nach Match-Typ unterschiedliche Statusinformationen
        if (gesellschafterMatch.matchType) {
            // Bei "Unsichere Zahlung" auch die Differenz anzeigen
            if (gesellschafterMatch.matchType === "Unsichere Zahlung" && gesellschafterMatch.betragsDifferenz) {
                matchStatus = ` (${gesellschafterMatch.matchType}, Diff: ${gesellschafterMatch.betragsDifferenz}€)`;
            } else {
                matchStatus = ` (${gesellschafterMatch.matchType})`;
            }
        }

        return `Gesellschafterkonto #${gesellschafterMatch.row}${matchStatus}`;
    }

    /**
     * Verarbeitet eine Holding-Transfers-Übereinstimmung
     * @param {Object} holdingMatch - Das Match-Ergebnis
     * @param {number} betragValue - Der Betrag
     * @param {Array} row - Die Bankbewegungszeile
     * @param {Object} columns - Die Spaltenkonfiguration
     * @param {Object} holdingCols - Die Spaltenkonfiguration für Holding Transfers
     * @returns {string} Formatierte Match-Information
     */
    function processHoldingMatch(holdingMatch, betragValue, row, columns, holdingCols) {
        let matchStatus = "";

        // Je nach Match-Typ unterschiedliche Statusinformationen
        if (holdingMatch.matchType) {
            // Bei "Unsichere Zahlung" auch die Differenz anzeigen
            if (holdingMatch.matchType === "Unsichere Zahlung" && holdingMatch.betragsDifferenz) {
                matchStatus = ` (${holdingMatch.matchType}, Diff: ${holdingMatch.betragsDifferenz}€)`;
            } else {
                matchStatus = ` (${holdingMatch.matchType})`;
            }
        }

        return `Holding Transfer #${holdingMatch.row}${matchStatus}`;
    }

    /**
     * Verarbeitet eine Gutschrift-Übereinstimmung
     * @param {Object} gutschriftMatch - Das Match-Ergebnis
     * @param {number} betragValue - Der Betrag
     * @param {Array} row - Die Bankbewegungszeile
     * @param {Object} columns - Die Spaltenkonfiguration
     * @param {Object} einnahmenCols - Die Spaltenkonfiguration für Einnahmen
     * @returns {string} Formatierte Match-Information
     */
    function processGutschriftMatch(gutschriftMatch, betragValue, row, columns, einnahmenCols) {
        let matchStatus = "";

        // Bei Gutschriften könnte der Betrag abweichen (z.B. Teilgutschrift)
        // Prüfen, ob die Beträge plausibel sind
        const gutschriftBetrag = Math.abs(gutschriftMatch.brutto);

        if (numberUtils$1.isApproximatelyEqual(betragValue, gutschriftBetrag, 0.01)) {
            // Beträge stimmen genau überein
            matchStatus = " (Vollständige Gutschrift)";
        } else if (betragValue < gutschriftBetrag) {
            // Teilgutschrift (Gutschriftbetrag kleiner als ursprünglicher Rechnungsbetrag)
            matchStatus = " (Teilgutschrift)";
        } else {
            // Ungewöhnlicher Fall - Gutschriftbetrag größer als Rechnungsbetrag
            matchStatus = " (Ungewöhnliche Gutschrift)";
        }

        return `Gutschrift zu Einnahme #${gutschriftMatch.row}${matchStatus}`;
    }

    /**
     * Formatiert Zeilen basierend auf dem Match-Typ
     * @param {Sheet} sheet - Das Sheet
     * @param {number} firstDataRow - Erste Datenzeile
     * @param {Array} matchResults - Array mit Match-Infos
     * @param {Object} columns - Spaltenkonfiguration
     */
    function formatMatchedRows(sheet, firstDataRow, matchResults, columns) {
        // Performance-optimiertes Batch-Update vorbereiten
        const formatBatches = {
            'Einnahme': { rows: [], color: "#EAF1DD" },  // Helles Grün (Grundfarbe für Einnahmen)
            'Vollständige Zahlung (Einnahme)': { rows: [], color: "#C6EFCE" },  // Kräftiges Grün
            'Teilzahlung (Einnahme)': { rows: [], color: "#FCE4D6" },  // Helles Orange
            'Ausgabe': { rows: [], color: "#FFCCCC" },   // Helles Rosa (Grundfarbe für Ausgaben)
            'Vollständige Zahlung (Ausgabe)': { rows: [], color: "#FFC7CE" },  // Helles Rot
            'Teilzahlung (Ausgabe)': { rows: [], color: "#FCE4D6" },  // Helles Orange
            'Eigenbeleg': { rows: [], color: "#DDEBF7" },  // Helles Blau (Grundfarbe für Eigenbelege)
            'Vollständige Zahlung (Eigenbeleg)': { rows: [], color: "#9BC2E6" },  // Kräftigeres Blau
            'Teilzahlung (Eigenbeleg)': { rows: [], color: "#FCE4D6" },  // Helles Orange
            'Gesellschafterkonto': { rows: [], color: "#E2EFDA" },  // Helles Grün für Gesellschafterkonto
            'Holding Transfer': { rows: [], color: "#FFF2CC" },  // Helles Gelb für Holding Transfers
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
            } else if (matchText.includes("Eigenbeleg")) {
                if (matchText.includes("Vollständige Zahlung")) {
                    formatBatches['Vollständige Zahlung (Eigenbeleg)'].rows.push(rowIndex);
                } else if (matchText.includes("Teilzahlung")) {
                    formatBatches['Teilzahlung (Eigenbeleg)'].rows.push(rowIndex);
                } else {
                    formatBatches['Eigenbeleg'].rows.push(rowIndex);
                }
            } else if (matchText.includes("Gesellschafterkonto")) {
                formatBatches['Gesellschafterkonto'].rows.push(rowIndex);
            } else if (matchText.includes("Holding Transfer")) {
                formatBatches['Holding Transfer'].rows.push(rowIndex);
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
    }

    /**
     * Setzt bedingte Formatierung für die Match-Spalte
     * @param {Sheet} sheet - Das Sheet
     * @param {string} columnLetter - Buchstabe für die Spalte
     */
    function setMatchColumnFormatting(sheet, columnLetter) {
        const conditions = [
            // Grundlegende Match-Typen mit beginsWith Pattern
            {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100", pattern: "beginsWith"},
            {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006", pattern: "beginsWith"},
            {value: "Eigenbeleg", background: "#DDEBF7", fontColor: "#2F5597", pattern: "beginsWith"},
            {value: "Gesellschafterkonto", background: "#E2EFDA", fontColor: "#375623", pattern: "beginsWith"},
            {value: "Holding Transfer", background: "#FFF2CC", fontColor: "#7F6000", pattern: "beginsWith"},
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
        ];

        // Bedingte Formatierung für die Match-Spalte setzen
        sheetUtils.setConditionalFormattingForColumn(sheet, columnLetter, conditions);
    }

    /**
     * Markiert bezahlte Einnahmen, Ausgaben und Eigenbelege farblich basierend auf dem Zahlungsstatus
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Object} bankZuordnungen - Zuordnungen aus dem Bankbewegungen-Sheet
     * @param {Object} config - Die Konfiguration
     */
    function markPaidInvoices(ss, bankZuordnungen, config) {
        // Alle relevanten Sheets abrufen
        const sheets = {
            einnahmen: ss.getSheetByName("Einnahmen"),
            ausgaben: ss.getSheetByName("Ausgaben"),
            eigenbelege: ss.getSheetByName("Eigenbelege"),
            gesellschafterkonto: ss.getSheetByName("Gesellschafterkonto"),
            holdingTransfers: ss.getSheetByName("Holding Transfers")
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
     * @param {Sheet} sheet - Das zu aktualisierende Sheet
     * @param {string} sheetType - Typ des Sheets ("einnahmen", "ausgaben", "eigenbelege", etc.)
     * @param {Object} bankZuordnungen - Zuordnungen aus dem Bankbewegungen-Sheet
     * @param {Object} config - Die Konfiguration
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
            columns.bankabgleich || 0
        );

        if (maxCol === 0) return; // Keine relevanten Spalten verfügbar

        const data = sheet.getRange(2, 1, numRows, maxCol).getValues();

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
            const nettobetrag = numberUtils$1.parseCurrency(data[i][columns.nettobetrag - 1]);
            const bezahltBetrag = numberUtils$1.parseCurrency(data[i][columns.bezahlt - 1]);
            const zahlungsDatum = data[i][columns.zahlungsdatum - 1];
            const referenz = data[i][columns.rechnungsnummer - 1];

            // Prüfe, ob es eine Gutschrift ist
            const isGutschrift = referenz && referenz.toString().startsWith("G-");

            // Prüfe, ob diese Position im Banking-Sheet zugeordnet wurde
            // Bei Eigenbelegen verwenden wir den key "eigenbeleg#row"
            const prefix = sheetType === "eigenbelege" ? "eigenbeleg" :
                sheetType === "gesellschafterkonto" ? "gesellschafterkonto" :
                    sheetType === "holdingTransfers" ? "holdingtransfer" :
                        sheetType;

            const zuordnungsKey = `${prefix}#${row}`;
            const hatBankzuordnung = bankZuordnungen[zuordnungsKey] !== undefined;

            // Zahlungsstatus berechnen
            const mwst = columns.mwstSatz ? numberUtils$1.parseMwstRate(data[i][columns.mwstSatz - 1], config.tax.defaultMwst) / 100 : 0;
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
                if (columns.bankabgleich && data[i][columns.bankabgleich - 1]?.toString().startsWith("✓ Bank:")) {
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
        if (bankabgleichUpdates.length > 0 && columns.bankabgleich) {
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
    }

    /**
     * Wendet eine Farbe auf mehrere Zeilen an
     * @param {Sheet} sheet - Das Sheet
     * @param {Array} rows - Die zu färbenden Zeilennummern
     * @param {string|null} color - Die Hintergrundfarbe oder null zum Zurücksetzen
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
     * Erstellt einen Informationstext für eine Bank-Zuordnung
     * @param {Object} zuordnung - Die Zuordnungsinformation
     * @returns {string} - Formatierter Informationstext
     */
    function getZuordnungsInfo(zuordnung) {
        if (!zuordnung) return "";

        let infoText = "✓ Bank: " + zuordnung.bankDatum;

        // Wenn es mehrere Zuordnungen gibt (z.B. bei aufgeteilten Zahlungen)
        if (zuordnung.additional && zuordnung.additional.length > 0) {
            infoText += " + " + zuordnung.additional.length + " weitere";
        }

        return infoText;
    }

    var matchingHandler = {
        performBankReferenceMatching,
        processEinnahmeMatching,
        processAusgabeMatching,
        findMatch,
        evaluateMatch,
        formatMatchedRows,
        setMatchColumnFormatting,
        markPaidInvoices
    };

    // modules/refreshModule/bankSheetHandler.js

    /**
     * Aktualisiert das Bankbewegungen-Sheet
     * @param {Sheet} sheet - Das Bankbewegungen-Sheet
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - true bei erfolgreicher Aktualisierung
     */
    function refreshBankSheet(sheet, config) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const lastRow = sheet.getLastRow();
            if (lastRow < 3) return true; // Nicht genügend Daten zum Aktualisieren

            const firstDataRow = 3; // Erste Datenzeile (nach Header-Zeile)
            const numDataRows = lastRow - firstDataRow + 1;
            const transRows = lastRow - firstDataRow - 1; // Anzahl der Transaktionszeilen ohne die letzte Zeile

            // Bankbewegungen-Konfiguration für Spalten
            const columns = config.bankbewegungen.columns;

            // Spaltenbuchstaben aus den Indizes generieren
            const columnLetters = {};
            Object.entries(columns).forEach(([key, index]) => {
                columnLetters[key] = stringUtils.getColumnLetter(index);
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
                const amt = numberUtils$1.parseCurrency(val);
                return [amt > 0 ? "Einnahme" : amt < 0 ? "Ausgabe" : ""];
            });
            sheet.getRange(firstDataRow, columns.transaktionstyp, numDataRows, 1).setValues(typeValues);

            // 3. Dropdown-Validierungen für Typ, Kategorie und Konten
            formattingHandler.applyBankSheetValidations(sheet, firstDataRow, numDataRows, columns, config);

            // 4. Bedingte Formatierung für Transaktionstyp-Spalte
            formattingHandler.setConditionalFormattingForStatusColumn(sheet, columnLetters.transaktionstyp, [
                {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100"},
                {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006"}
            ]);

            // 5. REFERENZEN-MATCHING: Suche nach Referenzen in Einnahmen- und Ausgaben-Sheets
            matchingHandler.performBankReferenceMatching(ss, sheet, firstDataRow, numDataRows, lastRow, columns, columnLetters, config);

            // 6. Endsaldo-Zeile aktualisieren
            updateEndSaldoRow(sheet, lastRow, columns, columnLetters);

            // 7. Spaltenbreiten anpassen
            sheet.autoResizeColumns(1, sheet.getLastColumn());

            return true;
        } catch (e) {
            console.error("Fehler beim Aktualisieren des Bankbewegungen-Sheets:", e);
            return false;
        }
    }

    /**
     * Aktualisiert die Endsaldo-Zeile im Bankbewegungen-Sheet
     * @param {Sheet} sheet - Das Sheet
     * @param {number} lastRow - Letzte Zeile
     * @param {Object} columns - Spaltenkonfiguration
     * @param {Object} columnLetters - Buchstaben für die Spalten
     */
    function updateEndSaldoRow(sheet, lastRow, columns, columnLetters) {
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
    }

    var bankSheetHandler = {
        refreshBankSheet
    };

    // modules/refreshModule/index.js

    /**
     * Modul zum Aktualisieren der Tabellenblätter und Neuberechnen von Formeln
     */
    const RefreshModule = {
        /**
         * Cache zurücksetzen
         */
        clearCache() {
            globalCache.clear();
        },

        /**
         * Aktualisiert das aktive Tabellenblatt
         * @param {Object} config - Die Konfiguration
         */
        refreshActiveSheet(config) {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const sheet = ss.getActiveSheet();
                const name = sheet.getName();
                const ui = SpreadsheetApp.getUi();

                // Cache zurücksetzen
                this.clearCache();

                if (["Einnahmen", "Ausgaben", "Eigenbelege"].includes(name)) {
                    dataSheetHandler.refreshDataSheet(sheet, name, config);
                    ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
                } else if (name === "Bankbewegungen") {
                    bankSheetHandler.refreshBankSheet(sheet, config);
                    ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
                } else if (name === "Gesellschafterkonto") {
                    // Hier später Logik für Gesellschafterkonto ergänzen
                    ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
                } else if (name === "Holding Transfers") {
                    // Hier später Logik für Holding Transfers ergänzen
                    ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
                } else {
                    ui.alert("Für dieses Blatt gibt es keine Refresh-Funktion.");
                }
            } catch (e) {
                console.error("Fehler beim Aktualisieren des aktiven Sheets:", e);
                SpreadsheetApp.getUi().alert("Ein Fehler ist beim Aktualisieren aufgetreten: " + e.toString());
            }
        },

        /**
         * Aktualisiert alle relevanten Tabellenblätter
         * @param {Object} config - Die Konfiguration
         */
        refreshAllSheets(config) {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();

                // Cache zurücksetzen
                this.clearCache();

                // Sheets in der richtigen Reihenfolge aktualisieren, um Abhängigkeiten zu berücksichtigen
                const refreshOrder = ["Einnahmen", "Ausgaben", "Eigenbelege", "Gesellschafterkonto", "Holding Transfers", "Bankbewegungen"];

                for (const name of refreshOrder) {
                    const sheet = ss.getSheetByName(name);
                    if (!sheet) continue;

                    if (name === "Bankbewegungen") {
                        bankSheetHandler.refreshBankSheet(sheet, config);
                    } else if (name === "Gesellschafterkonto") {
                        // Hier später Logik für Gesellschafterkonto ergänzen
                    } else if (name === "Holding Transfers") {
                        // Hier später Logik für Holding Transfers ergänzen
                    } else {
                        dataSheetHandler.refreshDataSheet(sheet, name, config);
                    }

                    // Kurze Pause einfügen, um API-Limits zu vermeiden
                    Utilities.sleep(100);
                }
            } catch (e) {
                console.error("Fehler beim Aktualisieren aller Sheets:", e);
                throw e; // Fehlermeldung weiterleiten, damit sie in der Hauptfunktion angezeigt wird
            }
        }
    };

    // modules/ustvaModule/dataModel.js
    /**
     * Erstellt ein leeres UStVA-Datenobjekt mit Nullwerten
     * @returns {Object} Leere UStVA-Datenstruktur
     */
    function createEmptyUStVA() {
        return {
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
        };
    }

    var dataModel$2 = {
        createEmptyUStVA
    };

    // modules/ustvaModule/calculator.js

    /**
     * Verarbeitet eine Zeile aus den Einnahmen/Ausgaben/Eigenbelegen für die UStVA
     *
     * @param {Array} row - Datenzeile aus einem Sheet
     * @param {Object} data - UStVA-Datenobjekt nach Monaten
     * @param {boolean} isIncome - Handelt es sich um Einnahmen (true) oder Ausgaben (false)
     * @param {boolean} isEigen - Handelt es sich um Eigenbelege (true oder false)
     * @param {Object} config - Die Konfiguration
     */
    function processUStVARow(row, data, isIncome, isEigen = false, config) {
        try {
            // Sheet-Typ bestimmen
            const sheetType = isIncome ? "einnahmen" : isEigen ? "eigenbelege" : "ausgaben";
            const columns = config[sheetType].columns;

            // Zahlungsdatum prüfen (nur abgeschlossene Zahlungen)
            const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
            if (!paymentDate || paymentDate > new Date()) return;

            // Monat und Jahr prüfen (nur relevantes Geschäftsjahr)
            const month = dateUtils.getMonthFromRow(row, sheetType, config);
            if (!month || month < 1 || month > 12) return;

            // Beträge aus der Zeile extrahieren
            const netto = numberUtils$1.parseCurrency(row[columns.nettobetrag - 1]);
            const restNetto = numberUtils$1.parseCurrency(row[columns.restbetragNetto - 1]) || 0; // Steuerbemessungsgrundlage für Teilzahlungen
            const gezahlt = netto - restNetto; // Tatsächlich gezahlter/erhaltener Betrag

            // Falls kein Betrag gezahlt wurde, nichts zu verarbeiten
            if (numberUtils$1.isApproximatelyEqual(gezahlt, 0)) return;

            // MwSt-Satz normalisieren
            const mwstRate = numberUtils$1.parseMwstRate(row[columns.mwstSatz - 1], config.tax.defaultMwst);
            const roundedRate = Math.round(mwstRate);

            // Steuer berechnen
            const tax = numberUtils$1.round(gezahlt * (mwstRate / 100), 2);

            // Kategorie ermitteln
            const category = row[columns.kategorie - 1]?.toString().trim() || "";

            // Verarbeitung basierend auf Zeilentyp
            if (isIncome) {
                processIncomeRow(data, month, gezahlt, tax, roundedRate, category, config);
            } else if (isEigen) {
                processEigenRow(data, month, gezahlt, tax, roundedRate, category, config);
            } else {
                processExpenseRow(data, month, gezahlt, tax, roundedRate, category, config);
            }
        } catch (e) {
            console.error("Fehler bei der Verarbeitung einer UStVA-Zeile:", e);
        }
    }

    /**
     * Verarbeitet eine Einnahmenzeile für die UStVA
     * @param {Object} data - UStVA-Datenobjekt nach Monaten
     * @param {number} month - Monat (1-12)
     * @param {number} gezahlt - Gezahlter Betrag
     * @param {number} tax - Berechnete Steuer
     * @param {number} roundedRate - Gerundeter Steuersatz
     * @param {string} category - Kategorie
     * @param {Object} config - Die Konfiguration
     */
    function processIncomeRow(data, month, gezahlt, tax, roundedRate, category, config) {
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
            }
        }
    }

    /**
     * Verarbeitet eine Eigenbeleg-Zeile für die UStVA
     * @param {Object} data - UStVA-Datenobjekt nach Monaten
     * @param {number} month - Monat (1-12)
     * @param {number} gezahlt - Gezahlter Betrag
     * @param {number} tax - Berechnete Steuer
     * @param {number} roundedRate - Gerundeter Steuersatz
     * @param {string} category - Kategorie
     * @param {Object} config - Die Konfiguration
     */
    function processEigenRow(data, month, gezahlt, tax, roundedRate, category, config) {
        // EIGENBELEGE
        const eigenCfg = config.eigenbelege.categories[category] ?? {};
        const taxType = eigenCfg.taxType ?? "steuerpflichtig";

        if (taxType === "steuerfrei") {
            // Steuerfreie Eigenbelege
            data[month].eigenbelege_steuerfrei += gezahlt;
        } else if (taxType === "eigenbeleg" && eigenCfg.besonderheit === "bewirtung") {
            // Bewirtungsbelege (nur 70% der Vorsteuer absetzbar)
            data[month].eigenbelege_steuerpflichtig += gezahlt;

            if (roundedRate === 7 || roundedRate === 19) {
                const vst70 = numberUtils$1.round(tax * 0.7, 2); // 70% absetzbare Vorsteuer
                const vst30 = numberUtils$1.round(tax * 0.3, 2); // 30% nicht absetzbar
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
    }

    /**
     * Verarbeitet eine Ausgabenzeile für die UStVA
     * @param {Object} data - UStVA-Datenobjekt nach Monaten
     * @param {number} month - Monat (1-12)
     * @param {number} gezahlt - Gezahlter Betrag
     * @param {number} tax - Berechnete Steuer
     * @param {number} roundedRate - Gerundeter Steuersatz
     * @param {string} category - Kategorie
     * @param {Object} config - Die Konfiguration
     */
    function processExpenseRow(data, month, gezahlt, tax, roundedRate, category, config) {
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
            }
        }
    }

    /**
     * Aggregiert UStVA-Daten für einen Zeitraum (z.B. Quartal oder Jahr)
     *
     * @param {Object} data - UStVA-Datenobjekt nach Monaten
     * @param {number} start - Startmonat (1-12)
     * @param {number} end - Endmonat (1-12)
     * @returns {Object} Aggregiertes UStVA-Datenobjekt
     */
    function aggregateUStVA(data, start, end) {
        const sum = dataModel$2.createEmptyUStVA();

        for (let m = start; m <= end; m++) {
            if (!data[m]) continue; // Überspringe, falls keine Daten für den Monat

            for (const key in sum) {
                sum[key] += data[m][key] || 0;
            }
        }

        return sum;
    }

    var calculator$2 = {
        processUStVARow,
        aggregateUStVA
    };

    // modules/ustvaModule/collector.js

    /**
     * Erfasst alle UStVA-Daten aus den verschiedenen Sheets
     * @param {Object} config - Die Konfiguration
     * @returns {Object|null} UStVA-Daten nach Monaten oder null bei Fehler
     */
    function collectUStVAData(config) {
        // Prüfen, ob der Cache aktuelle Daten enthält
        if (globalCache.has('computed', 'ustva')) {
            return globalCache.get('computed', 'ustva');
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
                Array.from({length: 12}, (_, i) => [i + 1, dataModel$2.createEmptyUStVA()])
            );

            // Daten für jede Zeile verarbeiten
            revenueData.slice(1).forEach(row => calculator$2.processUStVARow(row, ustvaData, true, false, config));
            expenseData.slice(1).forEach(row => calculator$2.processUStVARow(row, ustvaData, false, false, config));
            if (eigenData.length) {
                eigenData.slice(1).forEach(row => calculator$2.processUStVARow(row, ustvaData, false, true, config));
            }

            // Daten cachen
            globalCache.set('computed', 'ustva', ustvaData);

            return ustvaData;
        } catch (e) {
            console.error("Fehler beim Sammeln der UStVA-Daten:", e);
            return null;
        }
    }

    var collector$2 = {
        collectUStVAData
    };

    // modules/ustvaModule/formatter.js

    /**
     * Formatiert eine UStVA-Datenzeile für die Ausgabe
     *
     * @param {string} label - Bezeichnung der Zeile (z.B. Monat oder Quartal)
     * @param {Object} d - UStVA-Datenobjekt für den Zeitraum
     * @returns {Array} Formatierte Zeile für die Ausgabe
     */
    function formatUStVARow(label, d) {
        // Berechnung der USt-Zahlung: USt minus VSt (abzüglich nicht abzugsfähiger VSt)
        const ustZahlung = numberUtils$1.round(
            (d.ust_7 + d.ust_19) - ((d.vst_7 + d.vst_19) - d.nicht_abzugsfaehige_vst),
            2
        );

        // Berechnung des Gesamtergebnisses: Einnahmen minus Ausgaben (ohne Steueranteil)
        const ergebnis = numberUtils$1.round(
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
    }

    /**
     * Erstellt das UStVA-Sheet basierend auf den UStVA-Daten
     *
     * @param {Object} ustvaData - UStVA-Daten nach Monaten
     * @param {Spreadsheet} ss - Das aktive Spreadsheet
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} true bei Erfolg, false bei Fehler
     */
    function generateUStVASheet(ustvaData, ss, config) {
        try {
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
                        calculator$2.aggregateUStVA(ustvaData, quartalsStart, month)
                    ));
                }
            });

            // Jahresergebnis anfügen
            outputRows.push(formatUStVARow("Gesamtjahr", calculator$2.aggregateUStVA(ustvaData, 1, 12)));

            // UStVA-Sheet erstellen oder aktualisieren
            let ustvaSheet = sheetUtils$1.getOrCreateSheet(ss, "UStVA");
            ustvaSheet.clearContents();

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

            return true;
        } catch (e) {
            console.error("Fehler bei der Generierung des UStVA-Sheets:", e);
            return false;
        }
    }

    var formatter$2 = {
        formatUStVARow,
        generateUStVASheet
    };

    // modules/ustvaModule/index.js

    /**
     * Modul zur Berechnung der Umsatzsteuervoranmeldung (UStVA)
     * Unterstützt die Berechnung nach SKR04 für monatliche und quartalsweise Auswertungen
     */
    const UStVAModule = {
        /**
         * Cache leeren
         */
        clearCache() {
            globalCache.clear('ustva');
        },

        /**
         * Hauptfunktion zur Berechnung der UStVA
         * Sammelt Daten aus allen relevanten Sheets und erstellt ein UStVA-Sheet
         * @param {Object} config - Die Konfiguration
         * @returns {boolean} - true bei Erfolg, false bei Fehler
         */
        calculateUStVA(config) {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const ui = SpreadsheetApp.getUi();

                // Daten sammeln
                const ustvaData = collector$2.collectUStVAData(config);
                if (!ustvaData) {
                    ui.alert("Die UStVA konnte nicht berechnet werden. Bitte prüfen Sie die Fehleranzeige.");
                    return false;
                }

                // UStVA-Sheet erstellen oder aktualisieren
                const success = formatter$2.generateUStVASheet(ustvaData, ss, config);

                if (success) {
                    ui.alert("UStVA wurde erfolgreich aktualisiert!");
                    return true;
                } else {
                    ui.alert("Bei der Erstellung der UStVA ist ein Fehler aufgetreten.");
                    return false;
                }
            } catch (e) {
                console.error("Fehler bei der UStVA-Berechnung:", e);
                SpreadsheetApp.getUi().alert("Fehler bei der UStVA-Berechnung: " + e.toString());
                return false;
            }
        },

        // Methoden für Testzwecke und erweiterte Funktionalität
        _internal: {
            createEmptyUStVA: dataModel$2.createEmptyUStVA,
            processUStVARow: calculator$2.processUStVARow,
            formatUStVARow: formatter$2.formatUStVARow,
            aggregateUStVA: calculator$2.aggregateUStVA,
            collectUStVAData: collector$2.collectUStVAData
        }
    };

    // modules/bwaModule/dataModel.js
    /**
     * Erstellt ein leeres BWA-Datenobjekt mit Nullwerten
     * @returns {Object} Leere BWA-Datenstruktur
     */
    function createEmptyBWA() {
        return {
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
        };
    }

    var dataModel$1 = {
        createEmptyBWA
    };

    // modules/bwaModule/calculator.js

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

            const amount = numberUtils$1.parseCurrency(row[columns.nettobetrag - 1]);
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
            if (numberUtils$1.parseMwstRate(row[columns.mwstSatz - 1], config.tax.defaultMwst) === 0) {
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

            const amount = numberUtils$1.parseCurrency(row[columns.nettobetrag - 1]);
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

            const amount = numberUtils$1.parseCurrency(row[columns.nettobetrag - 1]);
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

            const amount = numberUtils$1.parseCurrency(row[columns.betrag - 1]);
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

            const amount = numberUtils$1.parseCurrency(row[columns.betrag - 1]);
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
            d.gesamtErloese = numberUtils$1.round(
                d.umsatzerloese + d.provisionserloese + d.steuerfreieInlandEinnahmen +
                d.steuerfreieAuslandEinnahmen + d.sonstigeErtraege + d.vermietung +
                d.zuschuesse + d.waehrungsgewinne + d.anlagenabgaenge,
                2
            );

            // Materialkosten
            d.gesamtWareneinsatz = numberUtils$1.round(
                d.wareneinsatz + d.fremdleistungen + d.rohHilfsBetriebsstoffe,
                2
            );

            // Betriebsausgaben
            d.gesamtBetriebsausgaben = numberUtils$1.round(
                d.bruttoLoehne + d.sozialeAbgaben + d.sonstigePersonalkosten +
                d.werbungMarketing + d.reisekosten + d.versicherungen + d.telefonInternet +
                d.buerokosten + d.fortbildungskosten + d.kfzKosten + d.sonstigeAufwendungen,
                2
            );

            // Abschreibungen & Zinsen
            d.gesamtAbschreibungenZinsen = numberUtils$1.round(
                d.abschreibungenMaschinen + d.abschreibungenBueromaterial +
                d.abschreibungenImmateriell + d.zinsenBank + d.zinsenGesellschafter +
                d.leasingkosten,
                2
            );

            // Besondere Posten
            d.gesamtBesonderePosten = numberUtils$1.round(
                d.eigenkapitalveraenderungen + d.gesellschafterdarlehen + d.ausschuettungen,
                2
            );

            // Rückstellungen
            d.gesamtRueckstellungenTransfers = numberUtils$1.round(
                d.steuerrueckstellungen + d.rueckstellungenSonstige,
                2
            );

            // EBIT
            d.ebit = numberUtils$1.round(
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

            d.gewerbesteuer = numberUtils$1.round(d.ebit * (taxConfig.gewerbesteuer / 10000) * steuerfaktor, 2);
            d.koerperschaftsteuer = numberUtils$1.round(d.ebit * (taxConfig.koerperschaftsteuer / 100) * steuerfaktor, 2);
            d.solidaritaetszuschlag = numberUtils$1.round(d.koerperschaftsteuer * (taxConfig.solidaritaetszuschlag / 100), 2);

            // Gesamte Steuerlast
            d.steuerlast = numberUtils$1.round(
                d.koerperschaftsteuer + d.solidaritaetszuschlag + d.gewerbesteuer,
                2
            );

            // Gewinn nach Steuern
            d.gewinnNachSteuern = numberUtils$1.round(d.ebit - d.steuerlast, 2);
        }
    }

    var calculator$1 = {
        processRevenue,
        processExpense,
        processEigen,
        processGesellschafter,
        processHolding,
        calculateAggregates
    };

    // modules/bwaModule/collector.js

    /**
     * Sammelt alle BWA-Daten aus den verschiedenen Sheets
     * @param {Object} config - Die Konfiguration
     * @returns {Object|null} BWA-Daten nach Monaten oder null bei Fehler
     */
    function aggregateBWAData(config) {
        try {
            // Prüfen ob Cache gültig ist
            if (globalCache.has('computed', 'bwa')) {
                return globalCache.get('computed', 'bwa');
            }

            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const revenueSheet = ss.getSheetByName("Einnahmen");
            const expenseSheet = ss.getSheetByName("Ausgaben");
            const eigenSheet = ss.getSheetByName("Eigenbelege");
            const gesellschafterSheet = ss.getSheetByName("Gesellschafterkonto");
            const holdingSheet = ss.getSheetByName("Holding Transfers");

            if (!revenueSheet || !expenseSheet) {
                console.error("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
                return null;
            }

            // BWA-Daten für alle Monate initialisieren
            const bwaData = Object.fromEntries(Array.from({length: 12}, (_, i) => [i + 1, dataModel$1.createEmptyBWA()]));

            // Daten aus den Sheets verarbeiten
            revenueSheet.getDataRange().getValues().slice(1).forEach(row => calculator$1.processRevenue(row, bwaData, config));
            expenseSheet.getDataRange().getValues().slice(1).forEach(row => calculator$1.processExpense(row, bwaData, config));

            if (eigenSheet) {
                eigenSheet.getDataRange().getValues().slice(1).forEach(row => calculator$1.processEigen(row, bwaData, config));
            }

            if (gesellschafterSheet) {
                gesellschafterSheet.getDataRange().getValues().slice(1).forEach(row => calculator$1.processGesellschafter(row, bwaData, config));
            }

            if (holdingSheet) {
                holdingSheet.getDataRange().getValues().slice(1).forEach(row => calculator$1.processHolding(row, bwaData, config));
            }

            // Gruppensummen und weitere Berechnungen
            calculator$1.calculateAggregates(bwaData, config);

            // Daten cachen
            globalCache.set('computed', 'bwa', bwaData);

            return bwaData;
        } catch (e) {
            console.error("Fehler bei der Aggregation der BWA-Daten:", e);
            return null;
        }
    }

    var collector$1 = {
        aggregateBWAData
    };

    // modules/bwaModule/formatter.js

    /**
     * Erstellt den Header für die BWA mit Monats- und Quartalsspalten
     * @param {Object} config - Die Konfiguration
     * @returns {Array} Header-Zeile
     */
    function buildHeaderRow(config) {
        const headers = ["Kategorie"];
        for (let q = 0; q < 4; q++) {
            for (let m = q * 3; m < q * 3 + 3; m++) {
                headers.push(`${config.common.months[m]} (€)`);
            }
            headers.push(`Q${q + 1} (€)`);
        }
        headers.push("Jahr (€)");
        return headers;
    }

    /**
     * Erstellt eine Ausgabezeile für eine Position
     * @param {Object} pos - Position mit Label und Wert-Funktion
     * @param {Object} bwaData - BWA-Daten
     * @returns {Array} Formatierte Zeile
     */
    function buildOutputRow(pos, bwaData) {
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
        const roundedMonthly = monthly.map(val => numberUtils$1.round(val, 2));
        const roundedQuarters = quarters.map(val => numberUtils$1.round(val, 2));
        const roundedYearly = numberUtils$1.round(yearly, 2);

        // Zeile zusammenstellen
        return [pos.label,
            ...roundedMonthly.slice(0, 3), roundedQuarters[0],
            ...roundedMonthly.slice(3, 6), roundedQuarters[1],
            ...roundedMonthly.slice(6, 9), roundedQuarters[2],
            ...roundedMonthly.slice(9, 12), roundedQuarters[3],
            roundedYearly];
    }

    /**
     * Erstellt das BWA-Sheet basierend auf den BWA-Daten
     * @param {Object} bwaData - BWA-Daten nach Monaten
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} true bei Erfolg, false bei Fehler
     */
    function generateBWASheet(bwaData, ss, config) {
        try {
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
            const headerRow = buildHeaderRow(config);
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
            let bwaSheet = sheetUtils$1.getOrCreateSheet(ss, "BWA");
            bwaSheet.clearContents();

            // Daten in das Sheet schreiben (als Batch für Performance)
            const dataRange = bwaSheet.getRange(1, 1, outputRows.length, outputRows[0].length);
            dataRange.setValues(outputRows);

            // Formatierungen anwenden
            applyBwaFormatting(bwaSheet, headerRow.length, bwaGruppen, outputRows.length);

            // BWA-Sheet aktivieren
            ss.setActiveSheet(bwaSheet);

            return true;
        } catch (e) {
            console.error("Fehler bei der BWA-Erstellung:", e);
            return false;
        }
    }

    /**
     * Wendet Formatierungen auf das BWA-Sheet an
     * @param {Sheet} sheet - Das zu formatierende Sheet
     * @param {number} headerLength - Anzahl der Spalten
     * @param {Array} bwaGruppen - Gruppenhierarchie
     * @param {number} totalRows - Gesamtzahl der Zeilen
     */
    function applyBwaFormatting(sheet, headerLength, bwaGruppen, totalRows) {
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
    }

    var formatter$1 = {
        buildHeaderRow,
        buildOutputRow,
        generateBWASheet,
        applyBwaFormatting
    };

    // modules/bwaModule/index.js

    /**
     * Modul zur Berechnung der Betriebswirtschaftlichen Auswertung (BWA)
     */
    const BWAModule = {
        /**
         * Cache leeren
         */
        clearCache() {
            globalCache.clear('bwa');
        },

        /**
         * Hauptfunktion zur Berechnung der BWA
         * @param {Object} config - Die Konfiguration
         * @returns {boolean} - true bei Erfolg, false bei Fehler
         */
        calculateBWA(config) {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const ui = SpreadsheetApp.getUi();

                // Daten sammeln
                const bwaData = collector$1.aggregateBWAData(config);
                if (!bwaData) {
                    ui.alert("BWA-Daten konnten nicht generiert werden.");
                    return false;
                }

                // BWA-Sheet generieren
                const success = formatter$1.generateBWASheet(bwaData, ss, config);

                if (success) {
                    ui.alert("BWA wurde aktualisiert!");
                    return true;
                } else {
                    ui.alert("Bei der Erstellung der BWA ist ein Fehler aufgetreten.");
                    return false;
                }
            } catch (e) {
                console.error("Fehler bei der BWA-Berechnung:", e);
                SpreadsheetApp.getUi().alert("Fehler bei der BWA-Berechnung: " + e.toString());
                return false;
            }
        },

        // Methoden für Testzwecke und erweiterte Funktionalität
        _internal: {
            createEmptyBWA: dataModel$1.createEmptyBWA,
            processRevenue: calculator$1.processRevenue,
            processExpense: calculator$1.processExpense,
            processEigen: calculator$1.processEigen,
            aggregateBWAData: collector$1.aggregateBWAData
        }
    };

    // modules/bilanzModule/dataModel.js
    /**
     * Erstellt eine leere Bilanz-Datenstruktur
     * @returns {Object} Leere Bilanz-Datenstruktur
     */
    function createEmptyBilanz() {
        return {
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
        };
    }

    var dataModel = {
        createEmptyBilanz
    };

    // modules/bilanzModule/calculator.js

    /**
     * Ermittelt die Summe der Gesellschafterdarlehen
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Object} gesellschafterCols - Spaltenkonfiguration
     * @param {Object} config - Die Konfiguration
     * @returns {number} Summe der Gesellschafterdarlehen
     */
    function getDarlehensumme(ss, gesellschafterCols, config) {
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
                    darlehenSumme += numberUtils$1.parseCurrency(row[gesellschafterCols.betrag - 1] || 0);
                }
            }
        }

        return darlehenSumme;
    }

    /**
     * Ermittelt die Summe der Steuerrückstellungen
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Object} ausgabenCols - Spaltenkonfiguration
     * @param {Object} config - Die Konfiguration
     * @returns {number} Summe der Steuerrückstellungen
     */
    function getSteuerRueckstellungen(ss, ausgabenCols, config) {
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
                    steuerRueckstellungen += numberUtils$1.parseCurrency(row[ausgabenCols.nettobetrag - 1] || 0);
                }
            }
        }

        return steuerRueckstellungen;
    }

    /**
     * Berechnet die Summen für die Bilanz
     * @param {Object} bilanzData - Die Bilanzdaten
     */
    function calculateBilanzSummen(bilanzData) {
        const { aktiva, passiva } = bilanzData;

        // Summen für Aktiva
        aktiva.summeAnlagevermoegen = numberUtils$1.round(
            aktiva.sachanlagen +
            aktiva.immaterielleVermoegen +
            aktiva.finanzanlagen,
            2
        );

        aktiva.summeUmlaufvermoegen = numberUtils$1.round(
            aktiva.bankguthaben +
            aktiva.kasse +
            aktiva.forderungenLuL +
            aktiva.vorraete,
            2
        );

        aktiva.summeAktiva = numberUtils$1.round(
            aktiva.summeAnlagevermoegen +
            aktiva.summeUmlaufvermoegen +
            aktiva.rechnungsabgrenzung,
            2
        );

        // Summen für Passiva
        passiva.summeEigenkapital = numberUtils$1.round(
            passiva.stammkapital +
            passiva.kapitalruecklagen +
            passiva.gewinnvortrag -
            passiva.verlustvortrag +
            passiva.jahresueberschuss,
            2
        );

        passiva.summeVerbindlichkeiten = numberUtils$1.round(
            passiva.bankdarlehen +
            passiva.gesellschafterdarlehen +
            passiva.verbindlichkeitenLuL +
            passiva.steuerrueckstellungen,
            2
        );

        passiva.summePassiva = numberUtils$1.round(
            passiva.summeEigenkapital +
            passiva.summeVerbindlichkeiten +
            passiva.rechnungsabgrenzung,
            2
        );
    }

    var calculator = {
        getDarlehensumme,
        getSteuerRueckstellungen,
        calculateBilanzSummen
    };

    // modules/bilanzModule/collector.js

    /**
     * Sammelt Daten aus verschiedenen Sheets für die Bilanz
     * @param {Object} config - Die Konfiguration
     * @returns {Object} Bilanz-Datenstruktur mit befüllten Werten
     */
    function aggregateBilanzData(config) {
        try {
            // Prüfen ob Cache gültig ist
            if (globalCache.has('computed', 'bilanz')) {
                return globalCache.get('computed', 'bilanz');
            }

            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const bilanzData = dataModel.createEmptyBilanz();

            // Spalten-Konfigurationen für die verschiedenen Sheets abrufen
            const bankCols = config.bankbewegungen.columns;
            const ausgabenCols = config.ausgaben.columns;
            const gesellschafterCols = config.gesellschafterkonto.columns;

            // 1. Banksaldo aus "Bankbewegungen" (Endsaldo)
            const bankSheet = ss.getSheetByName("Bankbewegungen");
            if (bankSheet) {
                const lastRow = bankSheet.getLastRow();
                if (lastRow >= 1) {
                    const label = bankSheet.getRange(lastRow, bankCols.buchungstext).getValue().toString().toLowerCase();
                    if (label === "endsaldo") {
                        bilanzData.aktiva.bankguthaben = numberUtils$1.parseCurrency(
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
                        bilanzData.passiva.jahresueberschuss = numberUtils$1.parseCurrency(row[row.length - 1]);
                        break;
                    }
                }
            }

            // 3. Stammkapital aus Konfiguration
            bilanzData.passiva.stammkapital = config.tax.stammkapital || 25000;

            // 4. Gesellschafterdarlehen aus dem Gesellschafterkonto-Sheet
            bilanzData.passiva.gesellschafterdarlehen = calculator.getDarlehensumme(ss, gesellschafterCols, config);

            // 5. Steuerrückstellungen aus dem Ausgaben-Sheet
            bilanzData.passiva.steuerrueckstellungen = calculator.getSteuerRueckstellungen(ss, ausgabenCols, config);

            // 6. Berechnung der Summen
            calculator.calculateBilanzSummen(bilanzData);

            // Daten im Cache speichern
            globalCache.set('computed', 'bilanz', bilanzData);

            return bilanzData;
        } catch (e) {
            console.error("Fehler bei der Sammlung der Bilanzdaten:", e);
            return null;
        }
    }

    var collector = {
        aggregateBilanzData
    };

    // modules/bilanzModule/formatter.js

    /**
     * Konvertiert einen Zahlenwert in eine Zellenformel oder direkte Zahl
     * @param {number} value - Der Wert
     * @param {boolean} useFormula - Ob eine Formel verwendet werden soll
     * @returns {string|number} - Formel als String oder direkter Wert
     */
    function valueOrFormula(value, useFormula = false) {
        if (numberUtils$1.isApproximatelyEqual(value, 0) && useFormula) {
            return "";  // Leere Zelle für 0-Werte bei Formeln
        }
        return value;
    }

    /**
     * Erstellt ein Array für die Aktiva-Seite der Bilanz
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Array} Array mit Zeilen für die Aktiva-Seite
     */
    function createAktivaArray(bilanzData, config) {
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
    }

    /**
     * Erstellt ein Array für die Passiva-Seite der Bilanz
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Array} Array mit Zeilen für die Passiva-Seite
     */
    function createPassivaArray(bilanzData, config) {
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
    }

    /**
     * Erstellt das Bilanz-Sheet
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} true bei Erfolg, false bei Fehler
     */
    function generateBilanzSheet(bilanzData, ss, config) {
        try {
            // Bilanz-Arrays erstellen
            const aktivaArray = createAktivaArray(bilanzData, config);
            const passivaArray = createPassivaArray(bilanzData, config);

            // Erstelle oder leere das Blatt "Bilanz"
            let bilanzSheet = sheetUtils$1.getOrCreateSheet(ss, "Bilanz");
            bilanzSheet.clearContents();

            // Batch-Write statt einzelner Zellen-Updates
            // Schreibe Aktiva ab Zelle A1 und Passiva ab Zelle E1
            bilanzSheet.getRange(1, 1, aktivaArray.length, 2).setValues(aktivaArray);
            bilanzSheet.getRange(1, 5, passivaArray.length, 2).setValues(passivaArray);

            // Formatierung anwenden
            formatBilanzSheet(bilanzSheet, aktivaArray.length, passivaArray.length);

            // Bilanz-Sheet aktivieren
            ss.setActiveSheet(bilanzSheet);

            return true;
        } catch (e) {
            console.error("Fehler bei der Generierung des Bilanz-Sheets:", e);
            return false;
        }
    }

    /**
     * Formatiert das Bilanz-Sheet
     * @param {Sheet} bilanzSheet - Das zu formatierende Sheet
     * @param {number} aktivaLength - Anzahl der Aktiva-Zeilen
     * @param {number} passivaLength - Anzahl der Passiva-Zeilen
     */
    function formatBilanzSheet(bilanzSheet, aktivaLength, passivaLength) {
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
    }

    var formatter = {
        valueOrFormula,
        createAktivaArray,
        createPassivaArray,
        generateBilanzSheet,
        formatBilanzSheet
    };

    // modules/bilanzModule/index.js

    /**
     * Modul zur Erstellung einer Bilanz nach SKR04
     * Erstellt eine standardkonforme Bilanz basierend auf den Daten aus anderen Sheets
     */
    const BilanzModule = {
        /**
         * Cache leeren
         */
        clearCache() {
            globalCache.clear('bilanz');
        },

        /**
         * Hauptfunktion zur Erstellung der Bilanz
         * Sammelt Daten und erstellt ein Bilanz-Sheet
         * @param {Object} config - Die Konfiguration
         * @returns {boolean} true bei Erfolg, false bei Fehler
         */
        calculateBilanz(config) {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const ui = SpreadsheetApp.getUi();

                // Bilanzdaten aggregieren
                const bilanzData = collector.aggregateBilanzData(config);
                if (!bilanzData) {
                    ui.alert("Fehler: Bilanzdaten konnten nicht gesammelt werden.");
                    return false;
                }

                // Bilanz-Sheet erstellen
                const success = formatter.generateBilanzSheet(bilanzData, ss, config);

                // Prüfen, ob Aktiva = Passiva
                const aktivaSumme = bilanzData.aktiva.summeAktiva;
                const passivaSumme = bilanzData.passiva.summePassiva;
                const differenz = Math.abs(aktivaSumme - passivaSumme);

                if (differenz > 0.01) {
                    // Bei Differenz die Bilanz trotzdem erstellen, aber warnen
                    ui.alert(
                        "Bilanz ist nicht ausgeglichen",
                        `Die Bilanzsummen von Aktiva (${numberUtils.formatCurrency(aktivaSumme)}) und Passiva (${numberUtils.formatCurrency(passivaSumme)}) ` +
                        `stimmen nicht überein. Differenz: ${numberUtils.formatCurrency(differenz)}. ` +
                        `Bitte überprüfen Sie Ihre Buchhaltungsdaten.`,
                        ui.ButtonSet.OK
                    );
                }

                if (success) {
                    ui.alert("Die Bilanz wurde erfolgreich erstellt!");
                    return true;
                } else {
                    ui.alert("Bei der Erstellung der Bilanz ist ein Fehler aufgetreten.");
                    return false;
                }
            } catch (e) {
                console.error("Fehler bei der Bilanzerstellung:", e);
                SpreadsheetApp.getUi().alert("Fehler bei der Bilanzerstellung: " + e.toString());
                return false;
            }
        },

        // Methoden für Testzwecke und erweiterte Funktionalität
        _internal: {
            createEmptyBilanz: dataModel.createEmptyBilanz,
            aggregateBilanzData: collector.aggregateBilanzData
        }
    };

    // modules/validatorModule/documentValidator.js

    /**
     * Validiert eine Zeile aus einem Dokument (Einnahmen, Ausgaben oder Eigenbelege)
     * @param {Array} row - Die zu validierende Zeile
     * @param {number} rowIndex - Der Index der Zeile (für Fehlermeldungen)
     * @param {string} sheetType - Der Typ des Sheets ("einnahmen", "ausgaben" oder "eigenbelege")
     * @param {Object} config - Die Konfiguration
     * @returns {Array<string>} - Array mit Warnungen
     */
    function validateDocumentRow(row, rowIndex, sheetType = "einnahmen", config) {
        const warnings = [];
        const columns = config[sheetType].columns;

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
            {check: r => stringUtils.isEmpty(r[columns.datum - 1]), message: `${sheetType === "eigenbelege" ? "Beleg" : "Rechnungs"}datum fehlt.`},
            {check: r => stringUtils.isEmpty(r[columns.rechnungsnummer - 1]), message: `${sheetType === "eigenbelege" ? "Beleg" : "Rechnungs"}nummer fehlt.`},
            {check: r => stringUtils.isEmpty(r[columns.kategorie - 1]), message: "Kategorie fehlt."},
            {check: r => numberUtils$1.isInvalidNumber(r[columns.nettobetrag - 1]), message: "Nettobetrag fehlt oder ungültig."},
            {
                check: r => {
                    const mwstStr = r[columns.mwstSatz - 1] == null ? "" : r[columns.mwstSatz - 1].toString().trim();
                    if (stringUtils.isEmpty(mwstStr)) return false; // Wird schon durch andere Regel geprüft

                    // MwSt-Satz extrahieren und normalisieren
                    const mwst = numberUtils$1.parseMwstRate(mwstStr, config.tax.defaultMwst);
                    if (isNaN(mwst)) return true;

                    // Prüfe auf erlaubte MwSt-Sätze aus der Konfiguration
                    const allowedRates = config?.tax?.allowedMwst || [0, 7, 19];
                    return !allowedRates.includes(Math.round(mwst));
                },
                message: `Ungültiger MwSt-Satz. Erlaubt sind: ${config?.tax?.allowedMwst?.join('%, ')}% oder leer.`
            }
        ];

        // Dokument-spezifische Regeln
        if (sheetType === "einnahmen" || sheetType === "ausgaben") {
            baseRules.push({
                check: r => stringUtils.isEmpty(r[columns.kunde - 1]),
                message: `${sheetType === "einnahmen" ? "Kunde" : "Lieferant"} fehlt.`
            });
        } else if (sheetType === "eigenbelege") {
            baseRules.push({
                check: r => stringUtils.isEmpty(r[columns.ausgelegtVon - 1]),
                message: "Ausgelegt von fehlt."
            });
            baseRules.push({
                check: r => stringUtils.isEmpty(r[columns.beschreibung - 1]),
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
                check: r => !stringUtils.isEmpty(r[columns.zahlungsart - 1]),
                message: `${paymentType} darf bei offener Zahlung nicht gesetzt sein.`
            },
            {
                check: r => !stringUtils.isEmpty(r[columns.zahlungsdatum - 1]),
                message: `${paymentDate} darf bei offener Zahlung nicht gesetzt sein.`
            }
        ];

        // Regeln für bezahlte/erstattete Zahlungen
        const paidPaymentRules = [
            {
                check: r => stringUtils.isEmpty(r[columns.zahlungsart - 1]),
                message: `${paymentType} muss bei ${paidStatus} Zahlung gesetzt sein.`
            },
            {
                check: r => stringUtils.isEmpty(r[columns.zahlungsdatum - 1]),
                message: `${paymentDate} muss bei ${paidStatus} Zahlung gesetzt sein.`
            },
            {
                check: r => {
                    if (stringUtils.isEmpty(r[columns.zahlungsdatum - 1])) return false; // Wird schon durch andere Regel geprüft

                    const paymentDate = dateUtils.parseDate(r[columns.zahlungsdatum - 1]);
                    return paymentDate ? paymentDate > new Date() : false;
                },
                message: `${paymentDate} darf nicht in der Zukunft liegen.`
            },
            {
                check: r => {
                    if (stringUtils.isEmpty(r[columns.zahlungsdatum - 1]) || stringUtils.isEmpty(r[columns.datum - 1])) return false;

                    const paymentDate = dateUtils.parseDate(r[columns.zahlungsdatum - 1]);
                    const documentDate = dateUtils.parseDate(r[columns.datum - 1]);
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
    }

    /**
     * Validiert alle Zeilen in einem Sheet
     * @param {Array} data - Zeilen-Daten (ohne Header)
     * @param {string} sheetType - Typ des Sheets ('einnahmen', 'ausgaben' oder 'eigenbelege')
     * @param {Object} config - Die Konfiguration
     * @returns {Array<string>} - Array mit Warnungen
     */
    function validateSheet(data, sheetType, config) {
        return data.reduce((warnings, row, index) => {
            // Nur nicht-leere Zeilen prüfen
            if (row.some(cell => cell !== "")) {
                const rowWarnings = validateDocumentRow(row, index + 2, sheetType, config);
                warnings.push(...rowWarnings);
            }
            return warnings;
        }, []);
    }

    /**
     * Validiert alle Sheets auf Fehler
     * @param {Sheet} revenueSheet - Das Einnahmen-Sheet
     * @param {Sheet} expenseSheet - Das Ausgaben-Sheet
     * @param {Sheet|null} bankSheet - Das Bankbewegungen-Sheet (optional)
     * @param {Sheet|null} eigenSheet - Das Eigenbelege-Sheet (optional)
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - True, wenn keine Fehler gefunden wurden
     */
    function validateAllSheets(revenueSheet, expenseSheet, bankSheet = null, eigenSheet = null, config) {
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
                const revenueWarnings = validateSheet(revenueData, "einnahmen", config);
                if (revenueWarnings.length) {
                    allWarnings.push("Fehler in 'Einnahmen':\n" + revenueWarnings.join("\n"));
                }
            }

            // Ausgaben validieren (wenn Daten vorhanden)
            if (expenseSheet.getLastRow() > 1) {
                const expenseData = expenseSheet.getDataRange().getValues().slice(1); // Header überspringen
                const expenseWarnings = validateSheet(expenseData, "ausgaben", config);
                if (expenseWarnings.length) {
                    allWarnings.push("Fehler in 'Ausgaben':\n" + expenseWarnings.join("\n"));
                }
            }

            // Eigenbelege validieren (wenn vorhanden und Daten vorhanden)
            if (eigenSheet && eigenSheet.getLastRow() > 1) {
                const eigenData = eigenSheet.getDataRange().getValues().slice(1); // Header überspringen
                const eigenWarnings = validateSheet(eigenData, "eigenbelege", config);
                if (eigenWarnings.length) {
                    allWarnings.push("Fehler in 'Eigenbelege':\n" + eigenWarnings.join("\n"));
                }
            }

            // Bankbewegungen validieren (wenn vorhanden)
            if (bankSheet) {
                const bankWarnings = validateBanking(bankSheet, config);
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
    }

    var documentValidator = {
        validateDocumentRow,
        validateSheet,
        validateAllSheets
    };

    // modules/validatorModule/bankValidator.js

    /**
     * Validiert das Bankbewegungen-Sheet
     * @param {Sheet} bankSheet - Das zu validierende Sheet
     * @param {Object} config - Die Konfiguration
     * @returns {Array<string>} - Array mit Warnungen
     */
    function validateBanking$1(bankSheet, config) {
        if (!bankSheet) return ["Bankbewegungen-Sheet nicht gefunden"];

        const data = bankSheet.getDataRange().getValues();
        const warnings = [];
        const columns = config.bankbewegungen.columns;

        // Regeln für Header- und Footer-Zeilen
        const headerFooterRules = [
            {check: r => stringUtils.isEmpty(r[columns.datum - 1]), message: "Buchungsdatum fehlt."},
            {check: r => stringUtils.isEmpty(r[columns.buchungstext - 1]), message: "Buchungstext fehlt."},
            {
                check: r => !stringUtils.isEmpty(r[columns.betrag - 1]) && !isNaN(parseFloat(r[columns.betrag - 1].toString().trim())),
                message: "Betrag darf nicht gesetzt sein."
            },
            {check: r => stringUtils.isEmpty(r[columns.saldo - 1]) || numberUtils$1.isInvalidNumber(r[columns.saldo - 1]), message: "Saldo fehlt oder ungültig."},
            {check: r => !stringUtils.isEmpty(r[columns.transaktionstyp - 1]), message: "Typ darf nicht gesetzt sein."},
            {check: r => !stringUtils.isEmpty(r[columns.kategorie - 1]), message: "Kategorie darf nicht gesetzt sein."},
            {check: r => !stringUtils.isEmpty(r[columns.kontoSoll - 1]), message: "Konto (Soll) darf nicht gesetzt sein."},
            {check: r => !stringUtils.isEmpty(r[columns.kontoHaben - 1]), message: "Gegenkonto (Haben) darf nicht gesetzt sein."}
        ];

        // Regeln für Datenzeilen
        const dataRowRules = [
            {check: r => stringUtils.isEmpty(r[columns.datum - 1]), message: "Buchungsdatum fehlt."},
            {check: r => stringUtils.isEmpty(r[columns.buchungstext - 1]), message: "Buchungstext fehlt."},
            {check: r => stringUtils.isEmpty(r[columns.betrag - 1]) || numberUtils$1.isInvalidNumber(r[columns.betrag - 1]), message: "Betrag fehlt oder ungültig."},
            {check: r => stringUtils.isEmpty(r[columns.saldo - 1]) || numberUtils$1.isInvalidNumber(r[columns.saldo - 1]), message: "Saldo fehlt oder ungültig."},
            {check: r => stringUtils.isEmpty(r[columns.transaktionstyp - 1]), message: "Typ fehlt."},
            {check: r => stringUtils.isEmpty(r[columns.kategorie - 1]), message: "Kategorie fehlt."},
            {check: r => stringUtils.isEmpty(r[columns.kontoSoll - 1]), message: "Konto (Soll) fehlt."},
            {check: r => stringUtils.isEmpty(r[columns.kontoHaben - 1]), message: "Gegenkonto (Haben) fehlt."}
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
    }

    var bankValidator = {
        validateBanking: validateBanking$1
    };

    // modules/validatorModule/cellValidator.js

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
    function validateDropdown(sheet, row, col, numRows, numCols, list) {
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
    }

    /**
     * Validiert eine einzelne Zelle anhand eines festgelegten Formats
     * @param {*} value - Der zu validierende Wert
     * @param {string} type - Der Validierungstyp (date, number, currency, mwst, text)
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Validierungsergebnis {isValid, message}
     */
    function validateCellValue(value, type, config) {
        switch (type.toLowerCase()) {
            case 'date':
                const date = dateUtils.parseDate(value);
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
                const amount = numberUtils$1.parseCurrency(value);
                return {
                    isValid: !isNaN(amount),
                    message: !isNaN(amount) ? "" : "Ungültiger Geldbetrag."
                };

            case 'mwst':
                const mwst = numberUtils$1.parseMwstRate(value, config.tax.defaultMwst);
                const allowedRates = config?.tax?.allowedMwst || [0, 7, 19];
                return {
                    isValid: allowedRates.includes(Math.round(mwst)),
                    message: allowedRates.includes(Math.round(mwst))
                        ? ""
                        : `Ungültiger MwSt-Satz. Erlaubt sind: ${allowedRates.join('%, ')}%.`
                };

            case 'text':
                return {
                    isValid: !stringUtils.isEmpty(value),
                    message: !stringUtils.isEmpty(value) ? "" : "Text darf nicht leer sein."
                };

            default:
                return {
                    isValid: true,
                    message: ""
                };
        }
    }

    /**
     * Validiert ein komplettes Sheet und liefert einen detaillierten Fehlerbericht
     * @param {Sheet} sheet - Das zu validierende Sheet
     * @param {Object} validationRules - Regeln für jede Spalte {spaltenIndex: {type, required}}
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Validierungsergebnis mit Fehlern pro Zeile/Spalte
     */
    function validateSheetWithRules(sheet, validationRules, config) {
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
                if (required && stringUtils.isEmpty(cellValue)) {
                    const error = {
                        row: rowIdx + 1,
                        column: colIdx + 1,
                        message: "Pflichtfeld nicht ausgefüllt"
                    };
                    addError(results, error);
                    return;
                }

                // Wenn nicht leer, auf Typ prüfen
                if (!stringUtils.isEmpty(cellValue) && type) {
                    const validation = validateCellValue(cellValue, type, config);
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
    }

    /**
     * Fügt einen Fehler zum Validierungsergebnis hinzu
     * @param {Object} results - Das Validierungsergebnis
     * @param {Object} error - Der Fehler {row, column, message}
     */
    function addError(results, error) {
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
    }

    var cellValidator = {
        validateDropdown,
        validateCellValue,
        validateSheetWithRules
    };

    // modules/validatorModule/index.js

    /**
     * Hauptexport der Validator-Funktionalität
     */
    var ValidatorModule = {
        ...documentValidator,
        ...bankValidator,
        ...cellValidator
    };

    // file: src/code.js
    // imports

    // =================== Globale Funktionen ===================
    /**
     * Erstellt das Menü in der Google Sheets UI beim Öffnen der Tabelle
     */
    function onOpen() {
        SpreadsheetApp.getUi()
            .createMenu("📂 Buchhaltung")
            .addItem("📥 Dateien importieren", "importDriveFiles")
            .addItem("🔄 Aktuelles Blatt aktualisieren", "refreshSheet")
            .addItem("📊 UStVA berechnen", "calculateUStVA")
            .addItem("📈 BWA berechnen", "calculateBWA")
            .addItem("📝 Bilanz erstellen", "calculateBilanz")
            .addToUi();
    }

    /**
     * Wird bei jeder Bearbeitung des Spreadsheets ausgelöst
     * Fügt Zeitstempel hinzu, wenn bestimmte Blätter bearbeitet werden
     *
     * @param {Object} e - Event-Objekt von Google Sheets
     */
    function onEdit(e) {
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
    }

    /**
     * Richtet die notwendigen Trigger für das Spreadsheet ein
     */
    function setupTrigger() {
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
    }

    /**
     * Validiert alle relevanten Sheets
     * @returns {boolean} - True wenn alle Sheets valide sind, False sonst
     */
    function validateSheets() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const bankSheet = ss.getSheetByName("Bankbewegungen");
        const eigenSheet = ss.getSheetByName("Eigenbelege");

        return ValidatorModule.validateAllSheets(revenueSheet, expenseSheet, bankSheet, eigenSheet, config$1);
    }

    /**
     * Gemeinsame Fehlerbehandlungsfunktion für alle Berechnungsfunktionen
     * @param {Function} fn - Die auszuführende Funktion
     * @param {string} errorMessage - Die Fehlermeldung bei einem Fehler
     */
    function executeWithErrorHandling(fn, errorMessage) {
        try {
            // Zuerst alle Sheets aktualisieren
            RefreshModule.refreshAllSheets(config$1);

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
    }

    /**
     * Aktualisiert das aktive Tabellenblatt
     */
    function refreshSheet() {
        RefreshModule.refreshActiveSheet(config$1);
    }

    /**
     * Berechnet die Umsatzsteuervoranmeldung
     */
    function calculateUStVA() {
        executeWithErrorHandling(
            () => UStVAModule.calculateUStVA(config$1),
            "Fehler bei der UStVA-Berechnung"
        );
    }

    /**
     * Berechnet die BWA (Betriebswirtschaftliche Auswertung)
     */
    function calculateBWA() {
        executeWithErrorHandling(
            () => BWAModule.calculateBWA(config$1),
            "Fehler bei der BWA-Berechnung"
        );
    }

    /**
     * Erstellt die Bilanz
     */
    function calculateBilanz() {
        executeWithErrorHandling(
            () => BilanzModule.calculateBilanz(config$1),
            "Fehler bei der Bilanzerstellung"
        );
    }

    /**
     * Importiert Dateien aus Google Drive und aktualisiert alle Tabellenblätter
     */
    function importDriveFiles() {
        try {
            ImportModule.importDriveFiles(config$1);
            RefreshModule.refreshAllSheets(config$1);

            // Optional: Nach dem Import auch eine Validierung durchführen
            // validateSheets();
        } catch (error) {
            SpreadsheetApp.getUi().alert("Fehler beim Dateiimport: " + error.message);
            console.error("Import-Fehler:", error);
        }
    }

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

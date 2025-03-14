/**
 * Konfiguration für die Buchhaltungsanwendung
 * Unterstützt die Buchhaltung für Holding und operative GmbH nach SKR04
 */
const config = {
    // Allgemeine Einstellungen
    common: {
        paymentType: ["Überweisung", "Bar", "Kreditkarte", "Paypal", "Lastschrift"],
        months: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"],
        shareholders: ["Christopher Giebel", "Hendrik Werner"],
        employees: [],
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
        // Konten-Mapping für Eigenbelege hinzugefügt
        kontoMapping: {
            "Kleidung": {soll: "6800", gegen: "1200", vorsteuer: "1576"},
            "Trinkgeld": {soll: "6800", gegen: "1200"},
            "Private Vorauslage": {soll: "6800", gegen: "1200"},
            "Bürokosten": {soll: "6815", gegen: "1200", vorsteuer: "1576"},
            "Reisekosten": {soll: "6650", gegen: "1200", vorsteuer: "1576"},
            "Bewirtung": {soll: "6670", gegen: "1200", vorsteuer: "1576"},
            "Sonstiges": {soll: "6800", gegen: "1200", vorsteuer: "1576"}
        },
        // BWA-Mapping für Eigenbelege hinzugefügt
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
        // Konten-Mapping für Gesellschafterkonto hinzugefügt
        kontoMapping: {
            "Gesellschafterdarlehen": {soll: "1200", gegen: "3300"},
            "Ausschüttungen": {soll: "2000", gegen: "1200"},
            "Kapitalrückführung": {soll: "2000", gegen: "1200"},
            "Privatentnahme": {soll: "1600", gegen: "1200"},
            "Privateinlage": {soll: "1200", gegen: "1600"}
        },
        // BWA-Mapping für Gesellschafterkonto hinzugefügt
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
            betrag: 2,              // B: Betrag
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
        // Konten-Mapping für Holding Transfers hinzugefügt
        kontoMapping: {
            "Gewinnübertrag": {soll: "1200", gegen: "8999"},
            "Kapitalrückführung": {soll: "1200", gegen: "2000"}
        },
        // BWA-Mapping für Holding Transfers hinzugefügt
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
            ...Object.keys(this.gesellschafterkonto.categories),
            ...Object.keys(this.holdingTransfers.categories),
        ];

        // Duplikate aus den Kategorien entfernen
        this.bankbewegungen.categories = [...new Set(this.bankbewegungen.categories)];

        return this;
    }
};

// Initialisierung ausführen und exportieren
export default config.initialize();
/**
 * Sheet-spezifische Konfigurationen
 */
export default {
    // Einnahmen-Konfiguration
    einnahmen: {
        columns: {
            datum: 1,                  // A: Rechnungsdatum
            rechnungsnummer: 2,        // B: Rechnungsnummer
            kunde: 3,                  // C: Kunde
            leistungsBeschreibung: 4,  // D: Leistungsbeschreibung
            notizen: 5,                // E: Notizen
            kategorie: 6,              // F: Kategorie
            buchungskonto: 7,          // G: Buchungskonto
            nettobetrag: 8,            // H: Nettobetrag
            mwstSatz: 9,               // I: MwSt-Satz in %
            mwstBetrag: 10,            // J: MwSt-Betrag (H*I)
            bruttoBetrag: 11,          // K: Bruttobetrag (H+J)
            bezahlt: 12,               // L: Bereits bezahlter Betrag
            restbetragNetto: 13,       // M: Restbetrag Netto
            quartal: 14,               // N: Berechnetes Quartal
            zahlungsstatus: 15,        // O: Zahlungsstatus (Offen/Teilbezahlt/Bezahlt)
            zahlungsart: 16,           // P: Zahlungsart
            zahlungsdatum: 17,         // Q: Zahlungsdatum
            bankabgleich: 18,          // R: Bankabgleich-Information
            uebertragJahr: 19,         // S: Übertrag Jahr
            zeitstempel: 20,           // T: Zeitstempel der letzten Änderung
            dateiname: 21,             // U: Dateiname (für importierte Dateien)
            dateilink: 22,             // V: Link zur Originaldatei
        },
        // Kategorien mit einheitlicher Struktur
        categories: {
            'Erlöse aus Lieferungen und Leistungen': {
                taxType: 'steuerpflichtig',
                group: 'umsatz',
                besonderheit: null,
            },
            'Provisionserlöse': {
                taxType: 'steuerpflichtig',
                group: 'umsatz',
                besonderheit: null,
            },
            'Sonstige betriebliche Erträge': {
                taxType: 'steuerpflichtig',
                group: 'sonstige_ertraege',
                besonderheit: null,
            },
            'Erträge aus Vermietung/Verpachtung': {
                taxType: 'steuerfrei_inland',
                group: 'vermietung',
                besonderheit: null,
            },
            'Erträge aus Zuschüssen': {
                taxType: 'steuerpflichtig',
                group: 'zuschuesse',
                besonderheit: null,
            },
            'Erträge aus Währungsgewinnen': {
                taxType: 'steuerpflichtig',
                group: 'sonstige_ertraege',
                besonderheit: null,
            },
            'Erträge aus Anlagenabgängen': {
                taxType: 'steuerpflichtig',
                group: 'anlagenabgaenge',
                besonderheit: null,
            },
            'Darlehen': {
                taxType: 'steuerfrei_inland',
                group: 'finanzen',
                besonderheit: null,
            },
            'Zinsen': {
                taxType: 'steuerfrei_inland',
                group: 'finanzen',
                besonderheit: null,
            },
            'Gewinnvortrag': {
                taxType: 'steuerfrei_inland',
                group: 'eigenkapital',
                besonderheit: null,
            },
            'Verlustvortrag': {
                taxType: 'steuerfrei_inland',
                group: 'eigenkapital',
                besonderheit: null,
            },
            'Umsatzsteuererstattungen': {
                taxType: 'steuerfrei_inland',
                group: 'steuerkonto',
                besonderheit: 'erstattung',
            },
            'Gutschriften (Warenrückgabe)': {
                taxType: 'steuerpflichtig',
                group: 'umsatz',
                besonderheit: 'erloesschmaelerung',
            },
        },
        // SKR04-konforme Konten-Zuordnung
        kontoMapping: {
            'Erlöse aus Lieferungen und Leistungen': {soll: '1200', gegen: '4400', mwst: '1776'},
            'Provisionserlöse': {soll: '1200', gegen: '4120', mwst: '1776'},
            'Sonstige betriebliche Erträge': {soll: '1200', gegen: '4830', mwst: '1776'},
            'Erträge aus Vermietung/Verpachtung': {soll: '1200', gegen: '4180'},
            'Erträge aus Zuschüssen': {soll: '1200', gegen: '4190', mwst: '1776'},
            'Erträge aus Währungsgewinnen': {soll: '1200', gegen: '4840'},
            'Erträge aus Anlagenabgängen': {soll: '1200', gegen: '4855'},
            'Darlehen': {soll: '1200', gegen: '3100'},
            'Zinsen': {soll: '1200', gegen: '4130'},
            'Gewinnvortrag': {soll: '1200', gegen: '2970'},
            'Verlustvortrag': {soll: '1200', gegen: '2978'},
            'Umsatzsteuererstattungen': {
                soll: '1800',
                gegen: '3820',
            },
            'Gutschriften (Warenrückgabe)': {
                soll: '4200',  // Erlösschmälerung 19%
                gegen: '1200', // Bank oder ggf. Debitorenkonto
                mwst: '1776',
            },
        },
        // BWA-Mapping für Einnahmen-Kategorien
        bwaMapping: {
            'Erlöse aus Lieferungen und Leistungen': 'umsatzerloese',
            'Provisionserlöse': 'provisionserloese',
            'Sonstige betriebliche Erträge': 'sonstigeErtraege',
            'Erträge aus Vermietung/Verpachtung': 'vermietung',
            'Erträge aus Zuschüssen': 'zuschuesse',
            'Erträge aus Währungsgewinnen': 'waehrungsgewinne',
            'Erträge aus Anlagenabgängen': 'anlagenabgaenge',
            'Darlehen': 'sonstigeErtraege',
            'Zinsen': 'sonstigeErtraege',
            // Umsatzsteuererstattungen nicht relevant für BWA
            'Umsatzsteuererstattungen': 'steuerlicheKorrekturen',
            'Gutschriften (Warenrückgabe)': 'umsatzerloese',
        },
    },

    // Ausgaben-Konfiguration
    ausgaben: {
        columns: {
            datum: 1,                  // A: Rechnungsdatum
            rechnungsnummer: 2,        // B: Rechnungsnummer
            kunde: 3,                  // C: Lieferant
            leistungsBeschreibung: 4,  // D: Leistungsbeschreibung
            notizen: 5,                // E: Notizen
            kategorie: 6,              // F: Kategorie
            buchungskonto: 7,          // G: Buchungskonto
            nettobetrag: 8,            // H: Nettobetrag
            mwstSatz: 9,               // I: MwSt-Satz in %
            mwstBetrag: 10,            // J: MwSt-Betrag (H*I)
            bruttoBetrag: 11,          // K: Bruttobetrag (H+J)
            bezahlt: 12,               // L: Bereits bezahlter Betrag
            restbetragNetto: 13,       // M: Restbetrag Netto
            quartal: 14,               // N: Berechnetes Quartal
            zahlungsstatus: 15,        // O: Zahlungsstatus (Offen/Teilbezahlt/Bezahlt)
            zahlungsart: 16,           // P: Zahlungsart
            zahlungsdatum: 17,         // Q: Zahlungsdatum
            bankabgleich: 18,          // R: Bankabgleich-Information
            uebertragJahr: 19,         // S: Übertrag Jahr
            zeitstempel: 20,           // T: Zeitstempel der letzten Änderung
            dateiname: 21,             // U: Dateiname (für importierte Dateien)
            dateilink: 22,             // V: Link zur Originaldatei
        },
        // Kategorien mit einheitlicher Struktur
        categories: {
            // Materialaufwand & Wareneinsatz
            'Wareneinsatz': {
                taxType: 'steuerpflichtig',
                group: 'material',
                besonderheit: null,
            },
            'Bezogene Leistungen': {
                taxType: 'steuerpflichtig',
                group: 'material',
                besonderheit: null,
            },
            'Roh-, Hilfs- & Betriebsstoffe': {
                taxType: 'steuerpflichtig',
                group: 'material',
                besonderheit: null,
            },

            // Provisionen
            'Provisionszahlungen': {
                taxType: 'steuerpflichtig',
                group: 'leistungen',
                besonderheit: null,
            },

            // Personalkosten
            'Bruttolöhne & Gehälter': {
                taxType: 'steuerfrei_inland',
                group: 'personal',
                besonderheit: null,
            },
            'Soziale Abgaben & Arbeitgeberanteile': {
                taxType: 'steuerfrei_inland',
                group: 'personal',
                besonderheit: null,
            },
            'Sonstige Personalkosten': {
                taxType: 'steuerpflichtig',
                group: 'personal',
                besonderheit: null,
            },

            // Raumkosten
            'Miete': {
                taxType: 'steuerfrei_inland',
                group: 'raum',
                besonderheit: null,
            },
            'Nebenkosten': {
                taxType: 'steuerpflichtig',
                group: 'raum',
                besonderheit: null,
            },

            // Betriebskosten
            'Betriebskosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
            },
            'Marketing & Werbung': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
            },
            'Reisekosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
            },
            'Versicherungen': {
                taxType: 'steuerfrei_inland',
                group: 'betrieb',
                besonderheit: null,
            },
            'Porto': {
                taxType: 'steuerfrei_inland',
                group: 'betrieb',
                besonderheit: null,
            },
            'Bewirtung': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: 'bewirtung',
            },
            'Telefon & Internet': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
            },
            'Bürokosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
            },
            'Fortbildungskosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
            },
            'IT-Kosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
            },

            // Abschreibungen & Zinsen
            'Abschreibungen Maschinen': {
                taxType: 'steuerpflichtig',
                group: 'abschreibung',
                besonderheit: null,
            },
            'Abschreibungen Büroausstattung': {
                taxType: 'steuerpflichtig',
                group: 'abschreibung',
                besonderheit: null,
            },
            'Abschreibungen immaterielle Wirtschaftsgüter': {
                taxType: 'steuerpflichtig',
                group: 'abschreibung',
                besonderheit: null,
            },
            'Zinsen auf Bankdarlehen': {
                taxType: 'steuerpflichtig',
                group: 'zinsen',
                besonderheit: null,
            },
            'Zinsen auf Gesellschafterdarlehen': {
                taxType: 'steuerpflichtig',
                group: 'zinsen',
                besonderheit: null,
            },
            'Leasingkosten': {
                taxType: 'steuerpflichtig',
                group: 'abschreibung',
                besonderheit: null,
            },

            // Steuern & Rückstellungen
            'Gewerbesteuerrückstellungen': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: null,
            },
            'Körperschaftsteuer': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: null,
            },
            'Solidaritätszuschlag': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: null,
            },
            'Sonstige Steuerrückstellungen': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: null,
            },
            'Steuerzahlungen': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: 'abgabe',
            },

            // Sonstige Aufwendungen
            'Sonstige betriebliche Aufwendungen': {
                taxType: 'steuerpflichtig',
                group: 'sonstige',
                besonderheit: null,
            },

            // Mobilität
            'Kfz-Kosten': {
                taxType: 'steuerpflichtig',
                group: 'mobilitaet',
                besonderheit: null,
            },

            // Finanzen
            'Bankgebühren': {
                taxType: 'steuerfrei_inland',
                group: 'finanzen',
                besonderheit: null,
            },
        },
        // SKR04-konforme Konten-Zuordnung
        kontoMapping: {
            'Wareneinsatz': {soll: '5000', gegen: '1200', vorsteuer: '1576'},
            'Bezogene Leistungen': {soll: '5300', gegen: '1200', vorsteuer: '1576'},
            'Roh-, Hilfs- & Betriebsstoffe': {soll: '5400', gegen: '1200', vorsteuer: '1576'},
            'Bruttolöhne & Gehälter': {soll: '6000', gegen: '1200'},
            'Soziale Abgaben & Arbeitgeberanteile': {soll: '6010', gegen: '1200'},
            'Sonstige Personalkosten': {soll: '6020', gegen: '1200', vorsteuer: '1576'},
            'Miete': {soll: '6310', gegen: '1200'},
            'Nebenkosten': {soll: '6320', gegen: '1200', vorsteuer: '1576'},
            'Betriebskosten': {soll: '6300', gegen: '1200', vorsteuer: '1576'},
            'Marketing & Werbung': {soll: '6600', gegen: '1200', vorsteuer: '1576'},
            'Reisekosten': {soll: '6650', gegen: '1200', vorsteuer: '1576'},
            'Versicherungen': {soll: '6400', gegen: '1200'},
            'Porto': {soll: '6810', gegen: '1200'},
            'Bewirtung': {soll: '6670', gegen: '1200', vorsteuer: '1576'},
            'Telefon & Internet': {soll: '6805', gegen: '1200', vorsteuer: '1576'},
            'Bürokosten': {soll: '6815', gegen: '1200', vorsteuer: '1576'},
            'Fortbildungskosten': {soll: '6830', gegen: '1200', vorsteuer: '1576'},
            'Abschreibungen Maschinen': {soll: '6200', gegen: '1200'},
            'Abschreibungen Büroausstattung': {soll: '6210', gegen: '1200'},
            'Abschreibungen immaterielle Wirtschaftsgüter': {soll: '6220', gegen: '1200'},
            'Zinsen auf Bankdarlehen': {soll: '7300', gegen: '1200'},
            'Zinsen auf Gesellschafterdarlehen': {soll: '7310', gegen: '1200'},
            'Leasingkosten': {soll: '6240', gegen: '1200', vorsteuer: '1576'},
            'Gewerbesteuerrückstellungen': {soll: '7610', gegen: '1200'},
            'Körperschaftsteuer': {soll: '7600', gegen: '1200'},
            'Solidaritätszuschlag': {soll: '7620', gegen: '1200'},
            'Sonstige Steuerrückstellungen': {soll: '7630', gegen: '1200'},
            'Sonstige betriebliche Aufwendungen': {soll: '6800', gegen: '1200', vorsteuer: '1576'},
            'Provisionszahlungen': { soll: '4920', gegen: '1200', vorsteuer: '1576' },
            'IT-Kosten': { soll: '6570', gegen: '1200', vorsteuer: '1576' },
            'Bankgebühren': { soll: '6855', gegen: '1200' },
            'Kfz-Kosten': { soll: '4550', gegen: '1200', vorsteuer: '1576' },
            'Steuerzahlungen': { soll: '7610', gegen: '1200' },

        },
        // BWA-Mapping für Ausgaben-Kategorien
        bwaMapping: {
            'Wareneinsatz': 'wareneinsatz',
            'Bezogene Leistungen': 'fremdleistungen',
            'Roh-, Hilfs- & Betriebsstoffe': 'rohHilfsBetriebsstoffe',
            'Betriebskosten': 'sonstigeAufwendungen',
            'Marketing & Werbung': 'werbungMarketing',
            'Reisekosten': 'reisekosten',
            'Bruttolöhne & Gehälter': 'bruttoLoehne',
            'Soziale Abgaben & Arbeitgeberanteile': 'sozialeAbgaben',
            'Sonstige Personalkosten': 'sonstigePersonalkosten',
            'Sonstige betriebliche Aufwendungen': 'sonstigeAufwendungen',
            'Miete': 'mieteNebenkosten',
            'Nebenkosten': 'mieteNebenkosten',
            'Versicherungen': 'versicherungen',
            'Porto': 'buerokosten',
            'Telefon & Internet': 'telefonInternet',
            'Bürokosten': 'buerokosten',
            'Fortbildungskosten': 'fortbildungskosten',
            'Abschreibungen Maschinen': 'abschreibungenMaschinen',
            'Abschreibungen Büroausstattung': 'abschreibungenBueromaterial',
            'Abschreibungen immaterielle Wirtschaftsgüter': 'abschreibungenImmateriell',
            'Zinsen auf Bankdarlehen': 'zinsenBank',
            'Zinsen auf Gesellschafterdarlehen': 'zinsenGesellschafter',
            'Leasingkosten': 'leasingkosten',
            'Google Ads': 'werbungMarketing',
            'AWS': 'sonstigeAufwendungen',
            'Facebook Ads': 'werbungMarketing',
            'Bewirtung': 'sonstigeAufwendungen',
            'Gewerbesteuerrückstellungen': 'gewerbesteuerRueckstellungen',
            'Körperschaftsteuer': 'koerperschaftsteuer',
            'Solidaritätszuschlag': 'solidaritaetszuschlag',
            'Sonstige Steuerrückstellungen': 'steuerrueckstellungen',
        },
    },

    // Eigenbelege-Konfiguration
    eigenbelege: {
        columns: {
            datum: 1,              // A: Belegdatum
            rechnungsnummer: 2,    // B: Belegnummer
            ausgelegtVon: 3,       // C: Ausgelegt von (Person)
            beschreibung: 4,       // D: Beschreibung
            notizen: 5,            // E: Notizen
            kategorie: 6,          // F: Kategorie
            buchungskonto: 7,      // G: Buchungskonto
            nettobetrag: 8,        // H: Nettobetrag
            mwstSatz: 9,           // I: MwSt-Satz in %
            mwstBetrag: 10,        // J: MwSt-Betrag (H*I)
            bruttoBetrag: 11,      // K: Bruttobetrag (H+J)
            bezahlt: 12,           // L: Bereits erstatteter Betrag
            restbetragNetto: 13,   // M: Restbetrag Netto
            quartal: 14,           // N: Berechnetes Quartal
            zahlungsstatus: 15,    // O: Erstattungsstatus (Offen/Erstattet/Gebucht)
            zahlungsart: 16,       // P: Erstattungsart
            zahlungsdatum: 17,     // Q: Erstattungsdatum
            bankabgleich: 18,      // R: Bankabgleich-Information
            uebertragJahr: 19,     // S: Übertrag Jahr
            zeitstempel: 20,       // T: Zeitstempel der letzten Änderung
            dateiname: 21,         // U: Dateiname (für importierte Dateien)
            dateilink: 22,         // V: Link zum Originalbeleg
        },
        // Kategorien mit einheitlicher Struktur
        categories: {
            'Kleidung': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
            },
            'Trinkgeld': {
                taxType: 'steuerfrei',
                group: 'betrieb',
                besonderheit: null,
            },
            'Private Vorauslage': {
                taxType: 'steuerfrei',
                group: 'sonstige',
                besonderheit: null,
            },
            'Bürokosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
            },
            'Reisekosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
            },
            'Bewirtung': {
                taxType: 'eigenbeleg',
                group: 'betrieb',
                besonderheit: 'bewirtung',
            },
            'Sonstiges': {
                taxType: 'steuerpflichtig',
                group: 'sonstige',
                besonderheit: null,
            },
        },
        // Konten-Mapping für Eigenbelege
        kontoMapping: {
            'Kleidung': {soll: '6800', gegen: '1200', vorsteuer: '1576'},
            'Trinkgeld': {soll: '6800', gegen: '1200'},
            'Private Vorauslage': {soll: '6800', gegen: '1890'},
            'Bürokosten': {soll: '6815', gegen: '1200', vorsteuer: '1576'},
            'Reisekosten': {soll: '6650', gegen: '1200', vorsteuer: '1576'},
            'Bewirtung': {soll: '6670', gegen: '1200', vorsteuer: '1576'},
            'Sonstiges': {soll: '6800', gegen: '1200', vorsteuer: '1576'},
        },
        // BWA-Mapping für Eigenbelege
        bwaMapping: {
            'Kleidung': 'sonstigeAufwendungen',
            'Trinkgeld': 'sonstigeAufwendungen',
            'Private Vorauslage': 'sonstigeAufwendungen',
            'Bürokosten': 'buerokosten',
            'Reisekosten': 'reisekosten',
            'Bewirtung': 'sonstigeAufwendungen',
            'Sonstiges': 'sonstigeAufwendungen',
        },
    },

    // Gesellschafterkonto-Konfiguration
    gesellschafterkonto: {
        columns: {
            datum: 1,              // A: Datum
            referenz: 2,           // B: Referenz/Belegnummer
            gesellschafter: 3,     // C: Gesellschafter
            notizen: 4,            // D: Notizen/Kommentare
            kategorie: 5,          // E: Kategorie
            buchungskonto: 6,      // F: Buchungskonto
            betrag: 7,             // G: Betrag
            quartal: 8,            // H: Berechnetes Quartal
            zahlungsart: 9,        // I: Zahlungsart
            zahlungsdatum: 10,     // J: Zahlungsdatum
            bankabgleich: 11,      // K: Bankabgleich-Information
            uebertragJahr: 12,     // L: Übertrag Jahr
            zeitstempel: 13,       // M: Zeitstempel der letzten Änderung
        },
        // Kategorien als Objekte mit einheitlicher Struktur
        categories: {
            'Gesellschafterdarlehen': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
            },
            'Ausschüttungen': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
            },
            'Kapitalrückführung': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
            },
            'Privatentnahme': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
            },
            'Privateinlage': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
            },
        },
        // Konten-Mapping für Gesellschafterkonto
        kontoMapping: {
            'Gesellschafterdarlehen': {soll: '1200', gegen: '0950'},
            'Ausschüttungen': {soll: '2000', gegen: '1200'},
            'Kapitalrückführung': {soll: '1890', gegen: '1200'},
            'Privatentnahme': {soll: '1800', gegen: '1200'},
            'Privateinlage': {soll: '1200', gegen: '1890'},
        },
        // BWA-Mapping für Gesellschafterkonto
        bwaMapping: {
            'Gesellschafterdarlehen': 'gesellschafterdarlehen',
            'Ausschüttungen': 'ausschuettungen',
            'Kapitalrückführung': 'eigenkapitalveraenderungen',
            'Privatentnahme': 'eigenkapitalveraenderungen',
            'Privateinlage': 'eigenkapitalveraenderungen',
        },
    },

    // Holding Transfers-Konfiguration
    holdingTransfers: {
        columns: {
            datum: 1,              // A: Buchungsdatum
            referenz: 2,           // B: Referenz/Belegnummer
            ziel: 3,               // C: Zielgesellschaft
            verwendungszweck: 4,   // D: Verwendungszweck
            notizen: 5,            // E: Notizen/Kommentare
            kategorie: 6,          // F: Kategorie
            buchungskonto: 7,      // G: Buchungskonto
            betrag: 8,             // H: Betrag
            quartal: 9,            // I: Quartal
            zahlungsart: 10,       // J: Zahlungsart
            zahlungsdatum: 11,     // K: Zahlungsdatum
            bankabgleich: 12,      // L: Bankabgleich-Information
            uebertragJahr: 13,     // M: Übertrag Jahr
            zeitstempel: 14,       // N: Zeitstempel der letzten Änderung
        },
        // Kategorien als Objekte mit einheitlicher Struktur
        categories: {
            'Gewinnübertrag': {
                taxType: 'steuerfrei_inland',
                group: 'holding',
                besonderheit: null,
            },
            'Kapitalrückführung': {
                taxType: 'steuerfrei_inland',
                group: 'holding',
                besonderheit: null,
            },
        },
        // Konten-Mapping für Holding Transfers
        kontoMapping: {
            'Gewinnübertrag': {soll: '1200', gegen: '8999'},
            'Kapitalrückführung': {soll: '1200', gegen: '2000'},
        },
        // BWA-Mapping für Holding Transfers
        bwaMapping: {
            'Gewinnübertrag': 'gesamtRueckstellungenTransfers',
            'Kapitalrückführung': 'eigenkapitalveraenderungen',
        },
    },

    // Bankbewegungen-Konfiguration
    bankbewegungen: {
        columns: {
            datum: 1,              // A: Buchungsdatum
            buchungstext: 2,       // B: Buchungstext
            notizen: 3,            // C: Notizen/Kommentare
            referenz: 4,           // D: Referenznummer
            matchInfo: 5,          // E: Match-Information zu Einnahmen/Ausgaben
            betrag: 6,             // F: Betrag
            saldo: 7,              // G: Saldo (berechnet)
            transaktionstyp: 8,    // H: Transaktionstyp (Einnahme/Ausgabe)
            kategorie: 9,          // I: Kategorie
            kontoSoll: 10,         // J: Konto (Soll)
            kontoHaben: 11,        // K: Gegenkonto (Haben)
            zeitstempel: 12,       // L: Zeitstempel
        },
        types: ['Einnahme', 'Ausgabe', 'Interne Buchung'],
        defaultAccount: '1200',
    },

    // Änderungshistorie-Konfiguration
    aenderungshistorie: {
        columns: {
            datum: 1,              // A: Datum/Zeitstempel
            typ: 2,                // B: Rechnungstyp (Einnahme/Ausgabe)
            dateiname: 3,          // C: Dateiname
            dateilink: 4,           // D: Link zur Datei
        },
    },
};
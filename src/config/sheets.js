/**
 * Sheet-spezifische Konfigurationen
 * Optimierte Struktur mit reduzierter Redundanz
 */
export default {
    // Einnahmen-Konfiguration
    einnahmen: {
        columns: {
            datum: 1,                  // A: Rechnungsdatum
            rechnungsnummer: 2,        // B: Rechnungsnummer
            kunde: 3,                  // C: Kunde
            ausland: 4,                // D: Ausland/EU-Ausland
            leistungsBeschreibung: 5,  // E: Leistungsbeschreibung
            notizen: 6,                // F: Notizen
            kategorie: 7,              // G: Kategorie
            buchungskonto: 8,          // H: Buchungskonto
            nettobetrag: 9,            // I: Nettobetrag
            mwstSatz: 10,              // J: MwSt-Satz in %
            mwstBetrag: 11,            // K: MwSt-Betrag (H*I)
            bruttoBetrag: 12,          // L: Bruttobetrag (H+K)
            bezahlt: 13,               // M: Bereits bezahlter Betrag
            restbetragNetto: 14,       // N: Restbetrag Netto
            quartal: 15,               // O: Berechnetes Quartal
            zahlungsstatus: 16,        // P: Zahlungsstatus (Offen/Teilbezahlt/Bezahlt)
            zahlungsart: 17,           // Q: Zahlungsart
            zahlungsdatum: 18,         // R: Zahlungsdatum
            bankabgleich: 19,          // S: Bankabgleich-Information
            uebertragJahr: 20,         // T: Übertrag Jahr
            zeitstempel: 21,           // U: Zeitstempel der letzten Änderung
            dateiname: 22,             // V: Dateiname (für importierte Dateien)
            dateilink: 23,             // W: Link zur Originaldatei
        },
        // Kategorien mit integrierter Konto- und BWA-Zuordnung
        categories: {
            'Erlöse aus Lieferungen und Leistungen': {
                taxType: 'steuerpflichtig',
                group: 'umsatz',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4400', mwst: '1776'},
                bwaMapping: 'erloeseLieferungenLeistungen',
            },
            'Provisionserlöse': {
                taxType: 'steuerpflichtig',
                group: 'umsatz',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4120', mwst: '1776'},
                bwaMapping: 'provisionserloese',
            },
            'Sonstige betriebliche Erträge': {
                taxType: 'steuerpflichtig',
                group: 'sonstige_ertraege',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4830', mwst: '1776'},
                bwaMapping: 'sonstigeBetrieblicheErtraege',
            },
            'Erträge aus Vermietung/Verpachtung': {
                taxType: 'steuerfrei_inland',
                group: 'vermietung',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4180'},
                bwaMapping: 'ertraegeVermietungVerpachtung',
            },
            'Erträge aus Zuschüssen': {
                taxType: 'steuerpflichtig',
                group: 'zuschuesse',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4190', mwst: '1776'},
                bwaMapping: 'ertraegeZuschuesse',
            },
            'Erträge aus Währungsgewinnen': {
                taxType: 'steuerpflichtig',
                group: 'sonstige_ertraege',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4840'},
                bwaMapping: 'ertraegeKursgewinne',
            },
            'Erträge aus Anlagenabgängen': {
                taxType: 'steuerpflichtig',
                group: 'anlagenabgaenge',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4855'},
                bwaMapping: 'ertraegeAnlagenabgaenge',
            },
            'Darlehen': {
                taxType: 'steuerfrei_inland',
                group: 'finanzen',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '3100'},
                bwaMapping: null,  // Kein Bestandteil der BWA! (Finanzierung, neutral)
            },
            'Zinsen': {
                taxType: 'steuerfrei_inland',
                group: 'finanzen',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4130'},
                bwaMapping: 'sonstigeBetrieblicheErtraege',
            },
            'Gewinnvortrag': {
                taxType: 'steuerfrei_inland',
                group: 'eigenkapital',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '2970'},
                bwaMapping: null, // Eigenkapitalbewegung – gehört nicht in die BWA!
            },
            'Verlustvortrag': {
                taxType: 'steuerfrei_inland',
                group: 'eigenkapital',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '2978'},
                bwaMapping: null, // Ebenfalls EK, nicht BWA-relevant
            },
            'Umsatzsteuererstattungen': {
                taxType: 'steuerfrei_inland',
                group: 'steuerkonto',
                besonderheit: 'erstattung',
                kontoMapping: {soll: '1800', gegen: '3820'},
                bwaMapping: null, // Steuerkonto – KEIN Bestandteil der BWA!
            },
            'Gutschriften (Warenrückgabe)': {
                taxType: 'steuerpflichtig',
                group: 'umsatz',
                besonderheit: 'erloesschmaelerung',
                kontoMapping: {soll: '4200', gegen: '1200', mwst: '1776'},
                bwaMapping: 'erloeseLieferungenLeistungen',
            },
        },
    },

    // Ausgaben-Konfiguration
    ausgaben: {
        columns: {
            datum: 1,                  // A: Rechnungsdatum
            rechnungsnummer: 2,        // B: Rechnungsnummer
            kunde: 3,                  // C: Lieferant
            ausland: 4,                // D: Ausland/EU-Ausland
            leistungsBeschreibung: 5,  // E: Leistungsbeschreibung
            notizen: 6,                // F: Notizen
            kategorie: 7,              // G: Kategorie
            buchungskonto: 8,          // H: Buchungskonto
            nettobetrag: 9,            // I: Nettobetrag
            mwstSatz: 10,              // J: MwSt-Satz in %
            mwstBetrag: 11,            // K: MwSt-Betrag (H*I)
            bruttoBetrag: 12,          // L: Bruttobetrag (H+K)
            bezahlt: 13,               // M: Bereits bezahlter Betrag
            restbetragNetto: 14,       // N: Restbetrag Netto
            quartal: 15,               // O: Berechnetes Quartal
            zahlungsstatus: 16,        // P: Zahlungsstatus (Offen/Teilbezahlt/Bezahlt)
            zahlungsart: 17,           // Q: Zahlungsart
            zahlungsdatum: 18,         // R: Zahlungsdatum
            bankabgleich: 19,          // S: Bankabgleich-Information
            uebertragJahr: 20,         // T: Übertrag Jahr
            zeitstempel: 21,           // U: Zeitstempel der letzten Änderung
            dateiname: 22,             // V: Dateiname (für importierte Dateien)
            dateilink: 23,             // W: Link zur Originaldatei
        },
        // Kategorien mit integrierter Konto- und BWA-Zuordnung
        categories: {
            // Materialaufwand & Wareneinsatz
            'Wareneinsatz': {
                taxType: 'steuerpflichtig',
                group: 'material',
                besonderheit: null,
                kontoMapping: {soll: '5000', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'wareneinsatz',
            },
            'Bezogene Leistungen': {
                taxType: 'steuerpflichtig',
                group: 'material',
                besonderheit: null,
                kontoMapping: {soll: '5300', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'bezogeneLeistungen',
            },
            'Roh-, Hilfs- & Betriebsstoffe': {
                taxType: 'steuerpflichtig',
                group: 'material',
                besonderheit: null,
                kontoMapping: {soll: '5400', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'rohHilfsBetriebsstoffe',
            },

            // Provisionen
            'Provisionszahlungen': {
                taxType: 'steuerpflichtig',
                group: 'leistungen',
                besonderheit: null,
                kontoMapping: {soll: '4920', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'provisionszahlungenDritte',
            },

            // Personalkosten
            'Bruttolöhne & Gehälter': {
                taxType: 'steuerfrei_inland',
                group: 'personal',
                besonderheit: null,
                kontoMapping: {soll: '6000', gegen: '1200'},
                bwaMapping: 'loehneGehaelter',
            },
            'Soziale Abgaben & Arbeitgeberanteile': {
                taxType: 'steuerfrei_inland',
                group: 'personal',
                besonderheit: null,
                kontoMapping: {soll: '6010', gegen: '1200'},
                bwaMapping: 'sozialeAbgaben',
            },
            'Sonstige Personalkosten': {
                taxType: 'steuerpflichtig',
                group: 'personal',
                besonderheit: null,
                kontoMapping: {soll: '6020', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'sonstigePersonalkosten',
            },

            // Raumkosten
            'Miete': {
                taxType: 'steuerfrei_inland',
                group: 'raum',
                besonderheit: null,
                kontoMapping: {soll: '6310', gegen: '1200'},
                bwaMapping: 'mieteLeasing',
            },
            'Nebenkosten': {
                taxType: 'steuerpflichtig',
                group: 'raum',
                besonderheit: null,
                kontoMapping: {soll: '6320', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'mieteLeasing',
            },

            // IT-Kosten
            'IT-Kosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6570', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'itKosten',
            },

            // Betriebskosten
            'Betriebskosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6300', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'sonstigeBetrieblicheAufwendungen',
            },
            'Marketing & Werbung': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6600', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'werbungMarketing',
            },
            'Reisekosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6650', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'reisekosten',
            },
            'Versicherungen': {
                taxType: 'steuerfrei_inland',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6400', gegen: '1200'},
                bwaMapping: 'versicherungen',
            },
            'Porto': {
                taxType: 'steuerfrei_inland',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6810', gegen: '1200'},
                bwaMapping: 'buerokosten',
            },
            'Bewirtung': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: 'bewirtung',
                kontoMapping: {soll: '6670', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'bewirtungskosten',
            },
            'Telefon & Internet': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6805', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'telefonInternet',
            },
            'Bürokosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6815', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'buerokosten',
            },
            'Fortbildungskosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6830', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'fortbildungskosten',
            },

            // Abschreibungen & Zinsen
            'Abschreibungen Maschinen': {
                taxType: 'steuerpflichtig',
                group: 'abschreibung',
                besonderheit: null,
                kontoMapping: {soll: '6200', gegen: '1200'},
                bwaMapping: 'abschreibungenSachanlagen',
            },
            'Abschreibungen Büroausstattung': {
                taxType: 'steuerpflichtig',
                group: 'abschreibung',
                besonderheit: null,
                kontoMapping: {soll: '6210', gegen: '1200'},
                bwaMapping: 'abschreibungenSachanlagen',
            },
            'Abschreibungen immaterielle Wirtschaftsgüter': {
                taxType: 'steuerpflichtig',
                group: 'abschreibung',
                besonderheit: null,
                kontoMapping: {soll: '6220', gegen: '1200'},
                bwaMapping: 'abschreibungenImmaterielleVG',
            },
            'Zinsen auf Bankdarlehen': {
                taxType: 'steuerpflichtig',
                group: 'zinsen',
                besonderheit: null,
                kontoMapping: {soll: '7300', gegen: '1200'},
                bwaMapping: 'zinsenBankdarlehen',
            },
            'Zinsen auf Gesellschafterdarlehen': {
                taxType: 'steuerpflichtig',
                group: 'zinsen',
                besonderheit: null,
                kontoMapping: {soll: '7310', gegen: '1200'},
                bwaMapping: 'zinsenGesellschafterdarlehen',
            },
            'Leasingkosten': {
                taxType: 'steuerpflichtig',
                group: 'abschreibung',
                besonderheit: null,
                kontoMapping: {soll: '6240', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'leasingzinsen',
            },

            // Beiträge und Abgaben
            'Beiträge & Abgaben': {
                taxType: 'steuerfrei_inland',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6845', gegen: '1200'},
                bwaMapping: 'beitraegeAbgaben',
            },

            // Steuern & Rückstellungen
            'Gewerbesteuerrückstellungen': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: null,
                kontoMapping: {soll: '7610', gegen: '1200'},
                bwaMapping: null, // Nicht BWA-relevant, gehört ins Jahresergebnis
            },
            'Körperschaftsteuer': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: null,
                kontoMapping: {soll: '7600', gegen: '1200'},
                bwaMapping: null, // Nicht BWA-relevant
            },
            'Solidaritätszuschlag': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: null,
                kontoMapping: {soll: '7620', gegen: '1200'},
                bwaMapping: null, // Nicht BWA-relevant,
            },
            'Sonstige Steuerrückstellungen': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: null,
                kontoMapping: {soll: '7630', gegen: '1200'},
                bwaMapping: null, // Nicht BWA-relevant,
            },
            'Steuerzahlungen': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: 'abgabe',
                kontoMapping: {soll: '7610', gegen: '1200'},
                bwaMapping: null, // Nicht in BWA (Ergebnisverwendung, kein Aufwand)
            },

            // Sonstige Aufwendungen
            'Sonstige betriebliche Aufwendungen': {
                taxType: 'steuerpflichtig',
                group: 'sonstige',
                besonderheit: null,
                kontoMapping: {soll: '6800', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'sonstigeBetrieblicheAufwendungen',
            },

            // Mobilität
            'Kfz-Kosten': {
                taxType: 'steuerpflichtig',
                group: 'mobilitaet',
                besonderheit: null,
                kontoMapping: {soll: '4550', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'kfzKosten',
            },

            // Finanzen
            'Bankgebühren': {
                taxType: 'steuerfrei_inland',
                group: 'finanzen',
                besonderheit: null,
                kontoMapping: {soll: '6855', gegen: '1200'},
                bwaMapping: 'sonstigeBetrieblicheAufwendungen',
            },
        },
    },

    // Eigenbelege-Konfiguration
    eigenbelege: {
        columns: {
            datum: 1,              // A: Belegdatum
            rechnungsnummer: 2,    // B: Belegnummer
            ausgelegtVon: 3,       // C: Ausgelegt von (Person)
            beschreibung: 4,       // D: Beschreibung
            ausland: 5,            // E: Ausland/EU-Ausland
            notizen: 6,            // F: Notizen
            kategorie: 7,          // G: Kategorie
            buchungskonto: 8,      // H: Buchungskonto
            nettobetrag: 9,        // I: Nettobetrag
            mwstSatz: 10,           // J: MwSt-Satz in %
            mwstBetrag: 11,        // K: MwSt-Betrag (H*I)
            bruttoBetrag: 12,      // L: Bruttobetrag (H+J)
            bezahlt: 13,           // M: Bereits erstatteter Betrag
            restbetragNetto: 14,   // N: Restbetrag Netto
            quartal: 15,           // O: Berechnetes Quartal
            zahlungsstatus: 16,    // P: Erstattungsstatus (Offen/Erstattet/Gebucht)
            zahlungsart: 17,       // Q: Erstattungsart
            zahlungsdatum: 18,     // R: Erstattungsdatum
            bankabgleich: 19,      // S: Bankabgleich-Information
            uebertragJahr: 20,     // T: Übertrag Jahr
            zeitstempel: 21,       // U: Zeitstempel der letzten Änderung
            dateiname: 22,         // V: Dateiname (für importierte Dateien)
            dateilink: 23,         // W: Link zum Originalbeleg
        },
        // Kategorien mit integrierter Konto- und BWA-Zuordnung
        categories: {
            'Kleidung': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6800', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'sonstigeBetrieblicheAufwendungen',
            },
            'Trinkgeld': {
                taxType: 'steuerfrei',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6800', gegen: '1200'},
                bwaMapping: 'sonstigeBetrieblicheAufwendungen',
            },
            'Bürokosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6815', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'buerokosten',
            },
            'Reisekosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6650', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'reisekosten',
            },
            'Bewirtung': {
                taxType: 'eigenbeleg',
                group: 'betrieb',
                besonderheit: 'bewirtung',
                kontoMapping: {soll: '6670', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'bewirtungskosten',
            },
            'Sonstiges': {
                taxType: 'steuerpflichtig',
                group: 'sonstige',
                besonderheit: null,
                kontoMapping: {soll: '6800', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'sonstigeBetrieblicheAufwendungen',
            },
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
        // Kategorien mit integrierter Konto- und BWA-Zuordnung
        categories: {
            'Gesellschafterdarlehen': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '0950'},
                bwaMapping: null, // nicht in der BWA aufführen, da Darlehen keine Betriebsausgabe/-einnahme ist
            },
            'Ausschüttungen': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
                kontoMapping: {soll: '2000', gegen: '1200'},
                bwaMapping: null, // Ausschüttungen gehören nicht in die BWA, sie sind nicht betriebsbezogen, sondern Ergebnisverwendung
            },
            'Kapitalrückführung': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
                kontoMapping: {soll: '1890', gegen: '1200'},
                bwaMapping: null, // nicht in die BWA, ist rein eigenkapitalbezogen
            },
            'Privatentnahme': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
                kontoMapping: {soll: '1800', gegen: '1200'},
                bwaMapping: null, // Nicht relevant für BWA
            },
            'Privateinlage': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '1890'},
                bwaMapping: null, // Nicht relevant für BWA
            },
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
        // Kategorien mit integrierter Konto- und BWA-Zuordnung
        categories: {
            'Gewinnübertrag': {
                taxType: 'steuerfrei_inland',
                group: 'holding',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '8999'},
                bwaMapping: null, // Nicht relevant für BWA
            },
            'Kapitalrückführung': {
                taxType: 'steuerfrei_inland',
                group: 'holding',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '2000'},
                bwaMapping: null, // Nicht relevant für BWA
            },
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
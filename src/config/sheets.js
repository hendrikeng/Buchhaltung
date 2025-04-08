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
                bilanzMapping: {
                    // Offene Forderungen erscheinen in der Bilanz
                    positiv: 'aktiva.umlaufvermoegen.forderungen.forderungenLuL',
                    negativ: null, // Gutschriften reduzieren Forderungen
                },
            },
            'Provisionserlöse': {
                taxType: 'steuerpflichtig',
                group: 'umsatz',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4120', mwst: '1776'},
                bwaMapping: 'provisionserloese',
                bilanzMapping: {
                    positiv: 'aktiva.umlaufvermoegen.forderungen.forderungenLuL',
                    negativ: null,
                },
            },
            'Sonstige betriebliche Erträge': {
                taxType: 'steuerpflichtig',
                group: 'sonstige_ertraege',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4830', mwst: '1776'},
                bwaMapping: 'sonstigeBetrieblicheErtraege',
                bilanzMapping: {
                    positiv: 'aktiva.umlaufvermoegen.forderungen.forderungenLuL',
                    negativ: null,
                },
            },
            'Erträge aus Vermietung/Verpachtung': {
                taxType: 'steuerfrei_inland',
                group: 'vermietung',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4180'},
                bwaMapping: 'ertraegeVermietungVerpachtung',
                bilanzMapping: {
                    positiv: 'aktiva.umlaufvermoegen.forderungen.forderungenLuL',
                    negativ: null,
                },
            },
            'Erträge aus Zuschüssen': {
                taxType: 'steuerpflichtig',
                group: 'zuschuesse',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4190', mwst: '1776'},
                bwaMapping: 'ertraegeZuschuesse',
                bilanzMapping: {
                    positiv: 'aktiva.umlaufvermoegen.forderungen.forderungenLuL',
                    negativ: null,
                },
            },
            'Erträge aus Währungsgewinnen': {
                taxType: 'steuerpflichtig',
                group: 'sonstige_ertraege',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4840'},
                bwaMapping: 'ertraegeKursgewinne',
                bilanzMapping: {
                    positiv: 'aktiva.umlaufvermoegen.forderungen.forderungenLuL',
                    negativ: null,
                },
            },
            'Erträge aus Anlagenabgängen': {
                taxType: 'steuerpflichtig',
                group: 'anlagenabgaenge',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4855'},
                bwaMapping: 'ertraegeAnlagenabgaenge',
                bilanzMapping: {
                    positiv: 'aktiva.umlaufvermoegen.forderungen.forderungenLuL',
                    negativ: null,
                },
            },
            'Darlehen': {
                taxType: 'steuerfrei_inland',
                group: 'finanzen',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '3100'},
                bwaMapping: null,  // Kein Bestandteil der BWA! (Finanzierung, neutral)
                bilanzMapping: {
                    positiv: 'aktiva.anlagevermoegen.finanzanlagen.ausleihungenVerbundene',
                    negativ: 'passiva.verbindlichkeiten.verbindlichkeitenKreditinstitute',
                },
            },
            'Zinsen': {
                taxType: 'steuerfrei_inland',
                group: 'finanzen',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '4130'},
                bwaMapping: 'sonstigeBetrieblicheErtraege',
                bilanzMapping: {
                    positiv: 'aktiva.umlaufvermoegen.forderungen.sonstigeVermoegensgegenstaende',
                    negativ: null,
                },
            },
            'Gewinnvortrag': {
                taxType: 'steuerfrei_inland',
                group: 'eigenkapital',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '2970'},
                bwaMapping: null, // Eigenkapitalbewegung – gehört nicht in die BWA!
                bilanzMapping: {
                    positiv: 'passiva.eigenkapital.gewinnvortrag',
                    negativ: 'passiva.eigenkapital.gewinnvortrag',
                },
            },
            'Verlustvortrag': {
                taxType: 'steuerfrei_inland',
                group: 'eigenkapital',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '2978'},
                bwaMapping: null, // Ebenfalls EK, nicht BWA-relevant
                bilanzMapping: {
                    positiv: 'passiva.eigenkapital.gewinnvortrag',
                    negativ: 'passiva.eigenkapital.gewinnvortrag',
                },
            },
            'Umsatzsteuererstattungen': {
                taxType: 'steuerfrei_inland',
                group: 'steuerkonto',
                besonderheit: 'erstattung',
                kontoMapping: {soll: '1800', gegen: '3820'},
                bwaMapping: null, // Steuerkonto – KEIN Bestandteil der BWA!
                bilanzMapping: {
                    positiv: 'aktiva.umlaufvermoegen.forderungen.sonstigeVermoegensgegenstaende',
                    negativ: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                },
            },
            'Gutschriften (Warenrückgabe)': {
                taxType: 'steuerpflichtig',
                group: 'umsatz',
                besonderheit: 'erloesschmaelerung',
                kontoMapping: {soll: '4200', gegen: '1200', mwst: '1776'},
                bwaMapping: 'erloeseLieferungenLeistungen',
                bilanzMapping: {
                    positiv: null,
                    negativ: 'aktiva.umlaufvermoegen.forderungen.forderungenLuL',
                },
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
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },
            'Bezogene Leistungen': {
                taxType: 'steuerpflichtig',
                group: 'material',
                besonderheit: null,
                kontoMapping: {soll: '5300', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'bezogeneLeistungen',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },
            'Roh-, Hilfs- & Betriebsstoffe': {
                taxType: 'steuerpflichtig',
                group: 'material',
                besonderheit: null,
                kontoMapping: {soll: '5400', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'rohHilfsBetriebsstoffe',
                bilanzMapping: {
                    positiv: 'aktiva.umlaufvermoegen.vorraete.rohHilfsBetriebsstoffe',
                    negativ: null,
                },
            },

            // Provisionen
            'Provisionszahlungen': {
                taxType: 'steuerpflichtig',
                group: 'leistungen',
                besonderheit: null,
                kontoMapping: {soll: '4920', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'provisionszahlungenDritte',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },

            // Personalkosten
            'Bruttolöhne & Gehälter': {
                taxType: 'steuerfrei_inland',
                group: 'personal',
                besonderheit: null,
                kontoMapping: {soll: '6000', gegen: '1200'},
                bwaMapping: 'loehneGehaelter',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                    negativ: null,
                },
            },
            'Soziale Abgaben & Arbeitgeberanteile': {
                taxType: 'steuerfrei_inland',
                group: 'personal',
                besonderheit: null,
                kontoMapping: {soll: '6010', gegen: '1200'},
                bwaMapping: 'sozialeAbgaben',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                    negativ: null,
                },
            },
            'Sonstige Personalkosten': {
                taxType: 'steuerpflichtig',
                group: 'personal',
                besonderheit: null,
                kontoMapping: {soll: '6020', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'sonstigePersonalkosten',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                    negativ: null,
                },
            },

            // Raumkosten
            'Miete': {
                taxType: 'steuerfrei_inland',
                group: 'raum',
                besonderheit: null,
                kontoMapping: {soll: '6310', gegen: '1200'},
                bwaMapping: 'mieteLeasing',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },
            'Nebenkosten': {
                taxType: 'steuerpflichtig',
                group: 'raum',
                besonderheit: null,
                kontoMapping: {soll: '6320', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'mieteLeasing',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },

            // IT-Kosten
            'IT-Kosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6570', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'itKosten',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },

            // Betriebskosten
            'Betriebskosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6300', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'sonstigeBetrieblicheAufwendungen',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },
            'Marketing & Werbung': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6600', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'werbungMarketing',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },
            'Reisekosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6650', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'reisekosten',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },
            'Versicherungen': {
                taxType: 'steuerfrei_inland',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6400', gegen: '1200'},
                bwaMapping: 'versicherungen',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },
            'Porto': {
                taxType: 'steuerfrei_inland',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6810', gegen: '1200'},
                bwaMapping: 'buerokosten',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },
            'Bewirtung': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: 'bewirtung',
                kontoMapping: {soll: '6670', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'bewirtungskosten',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },
            'Telefon & Internet': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6805', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'telefonInternet',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },
            'Bürokosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6815', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'buerokosten',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },
            'Fortbildungskosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6830', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'fortbildungskosten',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },

            // Abschreibungen & Zinsen
            'Abschreibungen Maschinen': {
                taxType: 'steuerpflichtig',
                group: 'abschreibung',
                besonderheit: null,
                kontoMapping: {soll: '6200', gegen: '1200'},
                bwaMapping: 'abschreibungenSachanlagen',
                bilanzMapping: {
                    positiv: 'aktiva.anlagevermoegen.sachanlagen.technischeAnlagenMaschinen',
                    negativ: null,
                },
            },
            'Abschreibungen Büroausstattung': {
                taxType: 'steuerpflichtig',
                group: 'abschreibung',
                besonderheit: null,
                kontoMapping: {soll: '6210', gegen: '1200'},
                bwaMapping: 'abschreibungenSachanlagen',
                bilanzMapping: {
                    positiv: 'aktiva.anlagevermoegen.sachanlagen.andereAnlagen',
                    negativ: null,
                },
            },
            'Abschreibungen immaterielle Wirtschaftsgüter': {
                taxType: 'steuerpflichtig',
                group: 'abschreibung',
                besonderheit: null,
                kontoMapping: {soll: '6220', gegen: '1200'},
                bwaMapping: 'abschreibungenImmaterielleVG',
                bilanzMapping: {
                    positiv: 'aktiva.anlagevermoegen.immaterielleVermoegensgegenstaende.konzessionen',
                    negativ: null,
                },
            },
            'Zinsen auf Bankdarlehen': {
                taxType: 'steuerpflichtig',
                group: 'zinsen',
                besonderheit: null,
                kontoMapping: {soll: '7300', gegen: '1200'},
                bwaMapping: 'zinsenBankdarlehen',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenKreditinstitute',
                    negativ: null,
                },
            },
            'Zinsen auf Gesellschafterdarlehen': {
                taxType: 'steuerpflichtig',
                group: 'zinsen',
                besonderheit: null,
                kontoMapping: {soll: '7310', gegen: '1200'},
                bwaMapping: 'zinsenGesellschafterdarlehen',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                    negativ: null,
                },
            },
            'Leasingkosten': {
                taxType: 'steuerpflichtig',
                group: 'abschreibung',
                besonderheit: null,
                kontoMapping: {soll: '6240', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'leasingzinsen',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },

            // Beiträge und Abgaben
            'Beiträge & Abgaben': {
                taxType: 'steuerfrei_inland',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6845', gegen: '1200'},
                bwaMapping: 'beitraegeAbgaben',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },

            // Steuern & Rückstellungen
            'Gewerbesteuerrückstellungen': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: null,
                kontoMapping: {soll: '7610', gegen: '1200'},
                bwaMapping: null, // Nicht BWA-relevant, gehört ins Jahresergebnis
                bilanzMapping: {
                    positiv: 'passiva.rueckstellungen.steuerrueckstellungen',
                    negativ: null,
                },
            },
            'Körperschaftsteuer': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: null,
                kontoMapping: {soll: '7600', gegen: '1200'},
                bwaMapping: null, // Nicht BWA-relevant
                bilanzMapping: {
                    positiv: 'passiva.rueckstellungen.steuerrueckstellungen',
                    negativ: null,
                },
            },
            'Solidaritätszuschlag': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: null,
                kontoMapping: {soll: '7620', gegen: '1200'},
                bwaMapping: null, // Nicht BWA-relevant,
                bilanzMapping: {
                    positiv: 'passiva.rueckstellungen.steuerrueckstellungen',
                    negativ: null,
                },
            },
            'Sonstige Steuerrückstellungen': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: null,
                kontoMapping: {soll: '7630', gegen: '1200'},
                bwaMapping: null, // Nicht BWA-relevant,
                bilanzMapping: {
                    positiv: 'passiva.rueckstellungen.steuerrueckstellungen',
                    negativ: null,
                },
            },
            'Steuerzahlungen': {
                taxType: 'steuerfrei_inland',
                group: 'steuer',
                besonderheit: 'abgabe',
                kontoMapping: {soll: '7610', gegen: '1200'},
                bwaMapping: null, // Nicht in BWA (Ergebnisverwendung, kein Aufwand)
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                    negativ: null,
                },
            },

            // Sonstige Aufwendungen
            'Sonstige betriebliche Aufwendungen': {
                taxType: 'steuerpflichtig',
                group: 'sonstige',
                besonderheit: null,
                kontoMapping: {soll: '6800', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'sonstigeBetrieblicheAufwendungen',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },

            // Mobilität
            'Kfz-Kosten': {
                taxType: 'steuerpflichtig',
                group: 'mobilitaet',
                besonderheit: null,
                kontoMapping: {soll: '4550', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'kfzKosten',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenLuL',
                    negativ: null,
                },
            },

            // Finanzen
            'Bankgebühren': {
                taxType: 'steuerfrei_inland',
                group: 'finanzen',
                besonderheit: null,
                kontoMapping: {soll: '6855', gegen: '1200'},
                bwaMapping: 'sonstigeBetrieblicheAufwendungen',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.verbindlichkeitenKreditinstitute',
                    negativ: null,
                },
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
            mwstSatz: 10,          // J: MwSt-Satz in %
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
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                    negativ: null,
                },
            },
            'Trinkgeld': {
                taxType: 'steuerfrei',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6800', gegen: '1200'},
                bwaMapping: 'sonstigeBetrieblicheAufwendungen',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                    negativ: null,
                },
            },
            'Bürokosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6815', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'buerokosten',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                    negativ: null,
                },
            },
            'Reisekosten': {
                taxType: 'steuerpflichtig',
                group: 'betrieb',
                besonderheit: null,
                kontoMapping: {soll: '6650', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'reisekosten',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                    negativ: null,
                },
            },
            'Bewirtung': {
                taxType: 'eigenbeleg',
                group: 'betrieb',
                besonderheit: 'bewirtung',
                kontoMapping: {soll: '6670', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'bewirtungskosten',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                    negativ: null,
                },
            },
            'Sonstiges': {
                taxType: 'steuerpflichtig',
                group: 'sonstige',
                besonderheit: null,
                kontoMapping: {soll: '6800', gegen: '1200', vorsteuer: '1576'},
                bwaMapping: 'sonstigeBetrieblicheAufwendungen',
                bilanzMapping: {
                    positiv: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                    negativ: null,
                },
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
                bilanzMapping: {
                    // Positiver Betrag: Gesellschaft gewährt Darlehen an Gesellschafter (Vermögenswert)
                    positiv: 'aktiva.anlagevermoegen.finanzanlagen.ausleihungenVerbundene',
                    // Negativer Betrag: Gesellschafter gewährt Darlehen an Gesellschaft (Verbindlichkeit)
                    negativ: 'passiva.verbindlichkeiten.sonstigeVerbindlichkeiten',
                },
            },
            'Ausschüttungen': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
                kontoMapping: {soll: '2000', gegen: '1200'},
                bwaMapping: null, // Ausschüttungen gehören nicht in die BWA, sie sind nicht betriebsbezogen, sondern Ergebnisverwendung
                bilanzMapping: {
                    positiv: 'passiva.eigenkapital.jahresueberschuss',
                    negativ: 'passiva.eigenkapital.gewinnvortrag',
                },
            },
            'Kapitalrückführung': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
                kontoMapping: {soll: '1890', gegen: '1200'},
                bwaMapping: null, // nicht in die BWA, ist rein eigenkapitalbezogen
                bilanzMapping: {
                    positiv: 'passiva.eigenkapital.gezeichnetesKapital',
                    negativ: 'passiva.eigenkapital.gezeichnetesKapital',
                },
            },
            'Privatentnahme': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
                kontoMapping: {soll: '1800', gegen: '1200'},
                bwaMapping: null, // Nicht relevant für BWA
                bilanzMapping: {
                    positiv: 'passiva.eigenkapital.gewinnvortrag',
                    negativ: 'passiva.eigenkapital.jahresueberschuss',
                },
            },
            'Privateinlage': {
                taxType: 'steuerfrei_inland',
                group: 'gesellschafter',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '1890'},
                bwaMapping: null, // Nicht relevant für BWA
                bilanzMapping: {
                    positiv: 'passiva.eigenkapital.kapitalruecklage',
                    negativ: null,
                },
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
                bilanzMapping: {
                    // Für Holding: Positiv = Einnahme = Beteiligung
                    // Für operative GmbH: Negativ = Ausgabe = Eigenkapitalminderung
                    positiv: 'aktiva.anlagevermoegen.finanzanlagen.anteileVerbundeneUnternehmen',
                    negativ: 'passiva.eigenkapital.jahresueberschuss',
                },
            },
            'Kapitalrückführung': {
                taxType: 'steuerfrei_inland',
                group: 'holding',
                besonderheit: null,
                kontoMapping: {soll: '1200', gegen: '2000'},
                bwaMapping: null, // Nicht relevant für BWA
                bilanzMapping: {
                    // Für Holding: Negativ = Ausgabe = Minderung der Beteiligung
                    // Für operative GmbH: Positiv = Einnahme = keine Änderung in Bilanz (ggf. Bankguthaben)
                    positiv: 'aktiva.umlaufvermoegen.liquideMittel.bankguthaben',
                    negativ: 'aktiva.anlagevermoegen.finanzanlagen.anteileVerbundeneUnternehmen',
                },
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
            dateilink: 4,          // D: Link zur Datei
        },
    },
};
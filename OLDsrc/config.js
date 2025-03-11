const config = {
    common: {
        paymentType: ["Überweisung", "Bar", "Kreditkarte", "Paypal"],
        months: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"],
    },
    tax: {
        defaultMwst: 19,
        stammkapital: 25000,
        year: 2021,
        isHolding: false, // true bei Holding
        holding: {
            gewerbesteuer: 16.45,
            koerperschaftsteuer: 5.5,
            solidaritaetszuschlag: 5.5,
            gewinnUebertragSteuerfrei: 95,
            gewinnUebertragSteuerpflichtig: 5
        },
        operative: {
            gewerbesteuer: 16.45,
            koerperschaftsteuer: 15,
            solidaritaetszuschlag: 5.5,
            gewinnUebertragSteuerfrei: 0,
            gewinnUebertragSteuerpflichtig: 100
        }
    },
    einnahmen: {
        categories: {
            "Erlöse aus Lieferungen und Leistungen": {taxType: "steuerpflichtig"},
            "Provisionserlöse": {taxType: "steuerpflichtig"},
            "Sonstige betriebliche Erträge": {taxType: "steuerpflichtig"},
            "Erträge aus Vermietung/Verpachtung": {taxType: "steuerfrei_inland"},
            "Erträge aus Zuschüssen": {taxType: "steuerpflichtig"},
            "Erträge aus Währungsgewinnen": {taxType: "steuerpflichtig"},
            "Erträge aus Anlagenabgängen": {taxType: "steuerpflichtig"},
            "Darlehen": {taxType: "steuerfrei_inland"},
            "Zinsen": {taxType: "steuerfrei_inland"}
        },
        kontoMapping: {
            "Erlöse aus Lieferungen und Leistungen": {soll: "1200 Bank", gegen: "4400 Erlöse aus L&L"},
            "Provisionserlöse": {soll: "1200 Bank", gegen: "4420 Provisionserlöse"},
            "Sonstige betriebliche Erträge": {soll: "1200 Bank", gegen: "4490 Sonstige betriebliche Erträge"},
            "Erträge aus Vermietung/Verpachtung": {soll: "1200 Bank", gegen: "4410 Vermietung/Verpachtung"},
            "Erträge aus Zuschüssen": {soll: "1200 Bank", gegen: "4430 Zuschüsse"},
            "Erträge aus Währungsgewinnen": {soll: "1200 Bank", gegen: "4440 Währungsgewinne"},
            "Erträge aus Anlagenabgängen": {soll: "1200 Bank", gegen: "4450 Anlagenabgänge"},
            "Darlehen": {soll: "1200 Bank", gegen: "3000 Darlehen"},
            "Zinsen": {soll: "1200 Bank", gegen: "2650 Zinserträge"}
        },
        bwaMapping: {
            "Erlöse aus Lieferungen und Leistungen": "umsatzerloese",
            "Provisionserlöse": "provisionserloese",
            "Sonstige betriebliche Erträge": "sonstigeErtraege",
            "Erträge aus Vermietung/Verpachtung": "vermietung",
            "Erträge aus Zuschüssen": "zuschuesse",
            "Erträge aus Währungsgewinnen": "waehrungsgewinne",
            "Erträge aus Anlagenabgängen": "anlagenabgaenge"
        }
    },
    ausgaben: {
        categories: {
            "Wareneinsatz": {taxType: "steuerpflichtig"},
            "Bezogene Leistungen": {taxType: "steuerpflichtig"},
            "Roh-, Hilfs- & Betriebsstoffe": {taxType: "steuerpflichtig"},
            "Betriebskosten": {taxType: "steuerpflichtig"},
            "Marketing & Werbung": {taxType: "steuerpflichtig"},
            "Reisekosten": {taxType: "steuerpflichtig"},
            "Personalkosten": {taxType: "steuerpflichtig"},
            "Bruttolöhne & Gehälter": {taxType: "steuerpflichtig"},
            "Soziale Abgaben & Arbeitgeberanteile": {taxType: "steuerpflichtig"},
            "Sonstige Personalkosten": {taxType: "steuerpflichtig"},
            "Sonstige betriebliche Aufwendungen": {taxType: "steuerpflichtig"},
            "Miete": {taxType: "steuerfrei_inland"},
            "Versicherungen": {taxType: "steuerfrei_inland"},
            "Porto": {taxType: "steuerfrei_inland"},
            "Google Ads": {taxType: "steuerfrei_ausland"},
            "AWS": {taxType: "steuerfrei_ausland"},
            "Facebook Ads": {taxType: "steuerfrei_ausland"},
            "Bewirtung": {taxType: "steuerpflichtig"},
            "Telefon & Internet": {taxType: "steuerpflichtig"},
            "Bürokosten": {taxType: "steuerpflichtig"},
            "Fortbildungskosten": {taxType: "steuerpflichtig"},
            "Abschreibungen Maschinen": {taxType: "steuerpflichtig"},
            "Abschreibungen Büroausstattung": {taxType: "steuerpflichtig"},
            "Abschreibungen immaterielle Wirtschaftsgüter": {taxType: "steuerpflichtig"},
            "Zinsen auf Bankdarlehen": {taxType: "steuerpflichtig"},
            "Zinsen auf Gesellschafterdarlehen": {taxType: "steuerpflichtig"},
            "Leasingkosten": {taxType: "steuerpflichtig"},
            "Gewerbesteuerrückstellungen": {taxType: "steuerpflichtig"},
            "Körperschaftsteuer": {taxType: "steuerpflichtig"},
            "Solidaritätszuschlag": {taxType: "steuerpflichtig"},
            "Sonstige Steuerrückstellungen": {taxType: "steuerpflichtig"}
        },
        kontoMapping: {
            "Wareneinsatz": {soll: "4900 Wareneinsatz", gegen: "1200 Bank"},
            "Bezogene Leistungen": {soll: "4900 Bezogene Leistungen", gegen: "1200 Bank"},
            "Roh-, Hilfs- & Betriebsstoffe": {soll: "4900 Roh-, Hilfs- & Betriebsstoffe", gegen: "1200 Bank"},
            "Betriebskosten": {soll: "4900 Betriebskosten", gegen: "1200 Bank"},
            "Marketing & Werbung": {soll: "4900 Marketing & Werbung", gegen: "1200 Bank"},
            "Reisekosten": {soll: "4900 Reisekosten", gegen: "1200 Bank"},
            "Bruttolöhne & Gehälter": {soll: "4900 Personalkosten", gegen: "1200 Bank"},
            "Soziale Abgaben & Arbeitgeberanteile": {soll: "4900 Personalkosten", gegen: "1200 Bank"},
            "Sonstige Personalkosten": {soll: "4900 Personalkosten", gegen: "1200 Bank"},
            "Sonstige betriebliche Aufwendungen": {soll: "4900 Sonstige betriebliche Aufwendungen", gegen: "1200 Bank"},
            "Miete": {soll: "4900 Miete & Nebenkosten", gegen: "1200 Bank"},
            "Versicherungen": {soll: "4900 Versicherungen", gegen: "1200 Bank"},
            "Porto": {soll: "4900 Betriebskosten", gegen: "1200 Bank"},
            "Google Ads": {soll: "4900 Marketing & Werbung", gegen: "1200 Bank"},
            "AWS": {soll: "4900 Marketing & Werbung", gegen: "1200 Bank"},
            "Facebook Ads": {soll: "4900 Marketing & Werbung", gegen: "1200 Bank"},
            "Bewirtung": {soll: "4900 Bewirtung", gegen: "1200 Bank"},
            "Telefon & Internet": {soll: "4900 Telefon & Internet", gegen: "1200 Bank"},
            "Bürokosten": {soll: "4900 Bürokosten", gegen: "1200 Bank"},
            "Fortbildungskosten": {soll: "4900 Fortbildungskosten", gegen: "1200 Bank"},
            "Abschreibungen Maschinen": {soll: "4900 Abschreibungen Maschinen", gegen: "1200 Bank"},
            "Abschreibungen Büroausstattung": {soll: "4900 Abschreibungen Büroausstattung", gegen: "1200 Bank"},
            "Abschreibungen immaterielle Wirtschaftsgüter": {
                soll: "4900 Abschreibungen Immateriell",
                gegen: "1200 Bank"
            },
            "Zinsen auf Bankdarlehen": {soll: "4900 Zinsen Bankdarlehen", gegen: "1200 Bank"},
            "Zinsen auf Gesellschafterdarlehen": {soll: "4900 Zinsen Gesellschafterdarlehen", gegen: "1200 Bank"},
            "Leasingkosten": {soll: "4900 Leasingkosten", gegen: "1200 Bank"},
            "Gewerbesteuerrückstellungen": {soll: "4900 Rückstellungen", gegen: "1200 Bank"},
            "Körperschaftsteuer": {soll: "4900 Körperschaftsteuer", gegen: "1200 Bank"},
            "Solidaritätszuschlag": {soll: "4900 Solidaritätszuschlag", gegen: "1200 Bank"},
            "Sonstige Steuerrückstellungen": {soll: "4900 Sonstige Steuerrückstellungen", gegen: "1200 Bank"}
        },
        bwaMapping: {
            "Wareneinsatz": "wareneinsatz",
            "Bezogene Leistungen": "bezogeneLeistungen",
            "Roh-, Hilfs- & Betriebsstoffe": "rohHilfsBetriebsstoffe",
            "Betriebskosten": "betriebskosten",
            "Marketing & Werbung": "werbungMarketing",
            "Reisekosten": "reisekosten",
            "Bruttolöhne & Gehälter": "bruttoLoehne",
            "Soziale Abgaben & Arbeitgeberanteile": "sozialeAbgaben",
            "Sonstige Personalkosten": "sonstigePersonalkosten",
            "Sonstige betriebliche Aufwendungen": "sonstigeAufwendungen",
            "Miete": "mieteNebenkosten",
            "Versicherungen": "versicherungen",
            "Porto": "betriebskosten",
            "Telefon & Internet": "telefonInternet",
            "Bürokosten": "buerokosten",
            "Fortbildungskosten": "fortbildungskosten"
        }
    },
    // Hier wird nun die Bank-Konfiguration angepasst:
    bank: {
        category: [
            // Kategorien aus Einnahmen
            "Erlöse aus Lieferungen und Leistungen",
            "Provisionserlöse",
            "Sonstige betriebliche Erträge",
            "Erträge aus Vermietung/Verpachtung",
            "Erträge aus Zuschüssen",
            "Erträge aus Währungsgewinnen",
            "Erträge aus Anlagenabgängen",
            "Darlehen",
            "Zinsen",
            // Kategorien aus Ausgaben
            "Wareneinsatz",
            "Bezogene Leistungen",
            "Roh-, Hilfs- & Betriebsstoffe",
            "Betriebskosten",
            "Marketing & Werbung",
            "Reisekosten",
            "Personalkosten",
            "Bruttolöhne & Gehälter",
            "Soziale Abgaben & Arbeitgeberanteile",
            "Sonstige Personalkosten",
            "Sonstige betriebliche Aufwendungen",
            "Miete",
            "Versicherungen",
            "Porto",
            "Google Ads",
            "AWS",
            "Facebook Ads",
            "Bewirtung",
            "Telefon & Internet",
            "Bürokosten",
            "Fortbildungskosten",
            // Kategorien aus Gesellschafterkonto
            "Gesellschafterdarlehen",
            "Ausschüttungen",
            "Kapitalrückführung",
            // Kategorien aus Holding Transfers
            "Gewinnübertrag",
            // Kategorien aus Eigenbelege
            "Kleidung",
            "Trinkgeld",
            "Private Vorauslage",
            "Sonstiges"
        ],
        type: ["Einnahme", "Ausgabe"]
    },
    gesellschafterkonto: {
        category: ["Gesellschafterdarlehen", "Ausschüttungen", "Kapitalrückführung"],
        shareholder: ["Christopher Giebel", "Hendrik Werner"]
    },
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
    holdingTransfers: {
        category: ["Gewinnübertrag", "Kapitalrückführung"]
    }
};

export default config
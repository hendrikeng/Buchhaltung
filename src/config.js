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
            "Erlöse aus Lieferungen und Leistungen": {soll: "1200 Bank", gegen: "8400 Erlöse aus L&L"},
            "Provisionserlöse": {soll: "1200 Bank", gegen: "8120 Provisionserlöse"},
            "Sonstige betriebliche Erträge": {soll: "1200 Bank", gegen: "8200 Sonstige betriebliche Erträge"},
            "Erträge aus Vermietung/Verpachtung": {soll: "1200 Bank", gegen: "8180 Vermietung/Verpachtung"},
            "Erträge aus Zuschüssen": {soll: "1200 Bank", gegen: "8190 Zuschüsse"},
            "Erträge aus Währungsgewinnen": {soll: "1200 Bank", gegen: "8130 Währungsgewinne"},
            "Erträge aus Anlagenabgängen": {soll: "1200 Bank", gegen: "8210 Anlagenabgänge"},
            "Darlehen": {soll: "1200 Bank", gegen: "3000 Darlehen"},
            "Zinsen": {soll: "1200 Bank", gegen: "8150 Zinserträge"}
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
            "Wareneinsatz": {soll: "5000 Wareneinsatz", gegen: "1200 Bank"},
            "Bezogene Leistungen": {soll: "5300 Bezogene Leistungen", gegen: "1200 Bank"},
            "Roh-, Hilfs- & Betriebsstoffe": {soll: "5400 Roh-, Hilfs- & Betriebsstoffe", gegen: "1200 Bank"},
            "Betriebskosten": {soll: "6300 Betriebskosten", gegen: "1200 Bank"},
            "Marketing & Werbung": {soll: "6600 Marketing & Werbung", gegen: "1200 Bank"},
            "Reisekosten": {soll: "6650 Reisekosten", gegen: "1200 Bank"},
            "Bruttolöhne & Gehälter": {soll: "6000 Personalkosten", gegen: "1200 Bank"},
            "Soziale Abgaben & Arbeitgeberanteile": {soll: "6010 Sozialabgaben", gegen: "1200 Bank"},
            "Sonstige Personalkosten": {soll: "6020 Sonstige Personalkosten", gegen: "1200 Bank"},
            "Sonstige betriebliche Aufwendungen": {soll: "6800 Sonstige betriebliche Aufwendungen", gegen: "1200 Bank"},
            "Miete": {soll: "6310 Miete & Nebenkosten", gegen: "1200 Bank"},
            "Versicherungen": {soll: "6400 Versicherungen", gegen: "1200 Bank"},
            "Porto": {soll: "6810 Betriebskosten", gegen: "1200 Bank"},
            "Telefon & Internet": {soll: "6805 Telefon & Internet", gegen: "1200 Bank"},
            "Bürokosten": {soll: "6815 Bürokosten", gegen: "1200 Bank"},
            "Fortbildungskosten": {soll: "6830 Fortbildungskosten", gegen: "1200 Bank"}
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
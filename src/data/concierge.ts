// Business-Concierge: neue Empfehlungen einfach als weiteren Eintrag ergänzen.
// partnerlink leer = Anfrage läuft über mein Formular (siehe empfehlungAnfrage in site.ts).

export type Empfehlung = {
  name: string;
  art: 'Tool' | 'Partner';
  text: string; // ein Satz: was es ist
  einsatz: string; // wofür ich es einsetze bzw. empfehle
  partnerlink?: string;
};

export const gruppen: { titel: string; eintraege: Empfehlung[] }[] = [
  {
    titel: 'KI und Automatisierung',
    eintraege: [
      {
        name: 'Langdock',
        art: 'Tool',
        text: 'DSGVO-konforme KI-Plattform: KI-Assistenten, Unternehmenswissen und Automatisierungen an einem Ort.',
        einsatz: 'Basis vieler meiner KI-Einführungen. Ich richte die Plattform ein, binde sie an vorhandene Systeme an und bringe sie mit Ihrem Team in den Alltag.',
      },
      {
        name: 'DevCTRL',
        art: 'Tool',
        text: 'KI-Assistent für Teams mit eingebauter Kontrolle: Freigaben für kritische Aktionen, Schutz sensibler Daten, lückenlose Protokolle.',
        einsatz: 'Wenn KI echte Aufgaben übernehmen soll und Datenschutz und Compliance mitreden müssen.',
      },
      {
        name: 'Jotform',
        art: 'Tool',
        text: 'Online-Formulare und Workflows: Anfragen, Aufträge und Freigaben digital erfassen und weiterleiten.',
        einsatz: 'Ersetzt Papier, E-Mail-Pingpong und Abtippen. Meine eigenen Anfragestrecken laufen ebenfalls darüber.',
      },
      {
        name: 'Perspective',
        art: 'Tool',
        text: 'Funnel-Builder für mobile Landingpages mit eingebauter Kontaktverwaltung.',
        einsatz: 'Für Kampagnen, die aus Anzeigen oder Social Media direkt qualifizierte Anfragen machen sollen.',
      },
    ],
  },
  {
    titel: 'Büro und Finanzen',
    eintraege: [
      {
        name: 'Lexware Office',
        art: 'Tool',
        text: 'Buchhaltung, Rechnungen und Belege online, mit Anbindung an Bank und Steuerkanzlei.',
        einsatz: 'Für einen digitalen Rechnungsworkflow ohne Medienbrüche.',
      },
      {
        name: 'Google Workspace',
        art: 'Tool',
        text: 'E-Mail, Kalender, Dokumente und Videokonferenzen für das ganze Team.',
        einsatz: 'Wenn Zusammenarbeit heute noch über Dateianhänge und lokale Ordner läuft.',
      },
    ],
  },
  {
    titel: 'Datenschutz und Sicherheit',
    eintraege: [
      {
        name: 'Insecco',
        art: 'Partner',
        text: 'Datenschutz und Informationssicherheit für den Mittelstand.',
        einsatz: 'Damit KI-Projekte und neue Prozesse von Anfang an rechtssicher aufgestellt sind.',
      },
    ],
  },
  {
    titel: 'Förderung und Weiterbildung',
    eintraege: [
      {
        name: 'Propero',
        art: 'Partner',
        text: 'Kooperationspartner für BAFA-geförderte Unternehmensberatungen.',
        einsatz: 'Ein Teil der Beratungskosten kann über das BAFA bezuschusst werden. Den Antrag stellen wir vor Beratungsbeginn.',
      },
      {
        name: 'Fördermittel-Finder',
        art: 'Partner',
        text: 'Passende Förderprogramme für Ihr Vorhaben finden.',
        einsatz: 'Für Digitalisierung, KI und Investitionen gibt es oft Zuschüsse, die ungenutzt bleiben.',
      },
      {
        name: 'Wegweiser digital',
        art: 'Partner',
        text: 'Geförderte KI-Weiterbildungen für Mitarbeitende finden.',
        einsatz: 'Damit Ihr Team KI sicher nutzt, zum Teil ohne eigene Kosten.',
      },
    ],
  },
  {
    titel: 'Dienstleister',
    eintraege: [
      {
        name: 'Marktplatz Software und Design GmbH',
        art: 'Partner',
        text: 'Marktplatz für Agenturdienstleistungen: Software, Design und Marketing.',
        einsatz: 'Wenn für ein Projekt zusätzliche Umsetzungskapazität gebraucht wird.',
      },
      {
        name: 'PendlerPortal',
        art: 'Partner',
        text: 'Mobilitätslösungen für Mittelständler: Arbeitswege und Fahrgemeinschaften einfacher organisieren.',
        einsatz: 'Ein Hebel für Mitarbeiterbindung und Gewinnung, gerade im ländlichen Raum.',
      },
    ],
  },
];

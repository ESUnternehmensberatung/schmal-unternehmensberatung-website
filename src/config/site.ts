// Zentrale Einstellungen. Sobald Google-Terminbuchung und Jotform-Formulare
// stehen, hier die URLs eintragen – alle Buttons der Seite ziehen automatisch mit.

import { url } from '../utils/url';

const email = 'info@schmal-unternehmensberatung.de';

function mailto(subject: string, body: string) {
  return `mailto:${email}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}

// Firmendaten: einzige Quelle für Website, Impressum, strukturierte Daten und
// Verzeichniseinträge (siehe briefing/eintraege-checkliste.md). Überall exakt so verwenden.
// Concierge: Empfehlungen ohne eigenen Partnerlink laufen über das eingebettete
// Jotform-Formular auf /empfehlung/, damit die Anfrage nachweislich über mich kommt.
// Der Name wird im Feld „wofurInteressieren“ vorausgewählt.
export function empfehlungAnfrage(name: string) {
  return `${url('/empfehlung/')}?wofurInteressieren=${encodeURIComponent(name)}`;
}

export const site = {
  name: 'Schmal Unternehmensberatung',
  owner: 'Eduard Schmal',
  legal: 'Schmal Unternehmensberatung, Inhaber Eduard Schmal',
  category: 'Unternehmensberatung',
  claim: 'Digitale Transformation für Mittelstand und freie Berufe',
  url: 'https://schmal-unternehmensberatung.de',
  email,
  phone: '+49 1522 685 95 24',
  phoneHref: 'tel:+4915226859524',
  address: ['Im Kamp 7', '21376 Eyendorf'],

  links: {
    // Formularseiten mit eingebettetem Jotform (src/layouts/Formular.astro, src/pages/*).
    quickWin: url('/quick-win/'),
    // Google-Terminbuchung: verlinkt auf der Danke-Seite und in der Antwort-Mail des Quick-Win-Formulars.
    termin: 'https://calendar.app.google/Eg49DKaRsUyQXEFH6',
    angebot: url('/angebot/'),
    whitepaper: url('/whitepaper/'),
    kiStarttag: url('/ki-starttag/'),
    // Bewertungsprofile (verlinkt unter den Kundenstimmen bei den Ergebnissen).
    provenExpert: 'https://www.provenexpert.com/de-de/es-unternehmensberatung/',
    googleProfil: 'https://maps.google.com/?cid=3350734573139222131',
    linkedin: 'https://www.linkedin.com/in/eduard-schmal-998611253',
  },
};

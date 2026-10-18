// Zentrale Einstellungen. Sobald Google-Terminbuchung und Jotform-Formulare
// stehen, hier die URLs eintragen – alle Buttons der Seite ziehen automatisch mit.

const email = 'info@schmal-unternehmensberatung.de';

function mailto(subject: string, body: string) {
  return `mailto:${email}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}

// Firmendaten: einzige Quelle für Website, Impressum, strukturierte Daten und
// Verzeichniseinträge (siehe briefing/eintraege-checkliste.md). Überall exakt so verwenden.
// Concierge: Empfehlungen ohne eigenen Partnerlink laufen über ein Formular,
// damit die Anfrage nachweislich über mich kommt. Sobald das Jotform-Formular
// steht, hier die URL eintragen; der Name der Empfehlung wird per URL-Parameter
// „wofurInteressieren“ (Mehrfachauswahl in Jotform) vorausgewählt.
const jotformEmpfehlung = 'https://form.jotform.com/262843738420056';

export function empfehlungAnfrage(name: string) {
  if (jotformEmpfehlung) return `${jotformEmpfehlung}?wofurInteressieren=${encodeURIComponent(name)}`;
  return mailto(
    `Empfehlung: ${name}`,
    `Guten Tag Herr Schmal,\n\nich interessiere mich für ${name} und bitte um Kontakt.\n\nName:\nFirma:\nTelefon:\nAnliegen:\n\nViele Grüße\n`,
  );
}

export const site = {
  name: 'Schmal Unternehmensberatung',
  owner: 'Eduard Schmal',
  legal: 'Schmal Unternehmensberatung, Inhaber Eduard Schmal',
  category: 'Unternehmensberatung',
  claim: 'Digitale Transformation und KI für den Mittelstand',
  url: 'https://schmal-unternehmensberatung.de',
  email,
  phone: '+49 1522 685 95 24',
  phoneHref: 'tel:+4915226859524',
  address: ['Im Kamp 7', '21376 Eyendorf'],

  links: {
    // Jotform „Quick-Win-Analyse anfragen“ (Feld „quelle“ zeigt, woher die Anfrage kommt).
    // Später: nach dem Absenden auf die Google-Terminbuchung weiterleiten.
    quickWin: 'https://form.jotform.com/262902467726060?quelle=website',
    // Später: Jotform „Angebot anfordern“ (Name, Firma, Mitarbeiterzahl, Anliegen).
    angebot: mailto(
      'Angebot anfordern',
      'Guten Tag Herr Schmal,\n\nbitte senden Sie mir ein Angebot.\n\n' +
        'Name:\nFirma:\nMitarbeiterzahl:\nAnliegen:\n\nViele Grüße\n',
    ),
    // Jotform: Whitepaper anfordern (Feld „quelle“ per URL vorbelegt).
    whitepaper: 'https://form.jotform.com/262843975491067?quelle=website',
    // Link zum Google-Unternehmensprofil (für die Bewertungen). Leer = Sektion ausgeblendet.
    googleProfil: '',
  },
};

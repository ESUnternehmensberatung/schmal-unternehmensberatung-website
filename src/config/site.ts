// Zentrale Einstellungen. Sobald Google-Terminbuchung und Jotform-Formulare
// stehen, hier die URLs eintragen – alle Buttons der Seite ziehen automatisch mit.

const email = 'es@schmal-unternehmensberatung.de';

function mailto(subject: string, body: string) {
  return `mailto:${email}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}

export const site = {
  name: 'ES | Schmal Unternehmensberatung',
  owner: 'Eduard Schmal',
  url: 'https://schmal-unternehmensberatung.de',
  email,
  phone: '+49 4172 / 431 90 17',
  phoneHref: 'tel:+4941724319017',
  address: ['Im Kamp 7', '21376 Eyendorf'],

  links: {
    // Später: Jotform-Vorabfragen, die auf die Google-Terminbuchung weiterleiten.
    quickWin: mailto(
      'Quick-Win-Analyse anfragen',
      'Guten Tag Herr Schmal,\n\nich interessiere mich für die kostenfreie Quick-Win-Analyse (45 Minuten).\n\n' +
        'Unternehmensgröße (Mitarbeitende):\n' +
        'Größter Zeitfresser im Moment:\n' +
        'Meine Rolle im Unternehmen:\n\n' +
        'Mögliche Termine:\n\nViele Grüße\n',
    ),
    // Später: Jotform „Angebot anfordern“ (Name, Firma, Mitarbeiterzahl, Anliegen).
    angebot: mailto(
      'Angebot anfordern',
      'Guten Tag Herr Schmal,\n\nbitte senden Sie mir ein Angebot.\n\n' +
        'Name:\nFirma:\nMitarbeiterzahl:\nAnliegen:\n\nViele Grüße\n',
    ),
    // Später: Jotform Whitepaper-Download mit Double-Opt-in.
    whitepaper: mailto(
      'Whitepaper vormerken',
      'Guten Tag Herr Schmal,\n\nbitte senden Sie mir das Whitepaper „Die 5 größten Fehler bei der digitalen Transformation im Mittelstand“, sobald es erscheint.\n\n' +
        'Name:\nFirma:\n\nMir ist bekannt, dass ich diese Einwilligung jederzeit widerrufen kann.\n\nViele Grüße\n',
    ),
    // Link zum Google-Unternehmensprofil (für die Bewertungen). Leer = Sektion ausgeblendet.
    googleProfil: '',
  },
};

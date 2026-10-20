// Einwilligungen für Cookies und Tracking (Cookie-Banner).
//
// Solange die Liste leer ist, erscheint KEIN Banner: Die Website setzt dann keine
// einwilligungspflichtigen Cookies, und ein Banner wäre überflüssig.
// Sobald hier ein Dienst steht, zeigt die Seite das Banner, und der Dienst wird erst
// nach Zustimmung geladen. Zusätzlich erscheint im Footer „Cookie-Einstellungen“.
//
// Beispiel:
// {
//   id: 'matomo',
//   name: 'Matomo',
//   kategorie: 'statistik',
//   anbieter: 'Eigenes Hosting',
//   zweck: 'Anonyme Besucherstatistik, um die Website zu verbessern.',
//   skript: 'https://statistik.example.de/matomo.js',
// }
//
// Beim Hinzufügen eines Dienstes: Version erhöhen (alle Besucher werden neu gefragt)
// und den Dienst in der Datenschutzerklärung beschreiben.

export type Kategorie = 'statistik' | 'marketing';

export type Dienst = {
  id: string;
  name: string;
  kategorie: Kategorie;
  anbieter: string;
  zweck: string;
  skript?: string; // externe Skript-URL, wird erst nach Einwilligung geladen
  inline?: string; // alternativ: Inline-Code, wird erst nach Einwilligung ausgeführt
};

export const consentVersion = 1;

export const kategorien: Record<Kategorie, { titel: string; text: string }> = {
  statistik: {
    titel: 'Statistik',
    text: 'Hilft mir zu verstehen, welche Inhalte gelesen werden, damit ich die Website verbessern kann.',
  },
  marketing: {
    titel: 'Marketing',
    text: 'Misst den Erfolg von Anzeigen und ermöglicht passende Werbung auf anderen Plattformen.',
  },
};

export const dienste: Dienst[] = [];

// Interne Pfade relativ zum Veröffentlichungsort: auf der eigenen Domain "/",
// auf GitHub Pages ohne Domain "/schmal-unternehmensberatung-website/".
const base = import.meta.env.BASE_URL.replace(/\/$/, '');

export function url(path = '/') {
  return `${base}${path.startsWith('/') ? path : `/${path}`}`;
}

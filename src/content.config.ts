import { defineCollection } from 'astro:content';
import { glob } from 'astro/loaders';
import { z } from 'astro/zod';

const rechtliches = defineCollection({
  loader: glob({ pattern: '*.md', base: './src/content/rechtliches' }),
  schema: z.object({
    title: z.string(),
    description: z.string(),
    quelle: z.string().optional(),
  }),
});

const faq = defineCollection({
  loader: glob({ pattern: '*.md', base: './src/content/faq' }),
  schema: z.object({
    frage: z.string(),
    reihenfolge: z.number(),
  }),
});

// Blog: Fachartikel (eigene Themen, ohne Provision) und Erfahrungsberichte zu
// Business-Concierge-Produkten (mit Provisionshinweis). `produkt` ist nur bei
// Concierge-Artikeln Pflicht.
const blog = defineCollection({
  loader: glob({ pattern: '**/*.md', base: './src/content/blog' }),
  schema: z
    .object({
      title: z.string(),
      description: z.string(),
      kategorie: z.enum(['concierge', 'fachartikel']).default('concierge'),
      produkt: z.string().optional(),
      datum: z.coerce.date(),
      entwurf: z.boolean().default(false),
    })
    .refine((d) => d.kategorie !== 'concierge' || !!d.produkt, {
      message: 'Concierge-Artikel brauchen das Feld "produkt".',
      path: ['produkt'],
    }),
});

export const collections = { rechtliches, faq, blog };

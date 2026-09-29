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

// Blog: je Business-Concierge-Produkt ein Artikel.
const blog = defineCollection({
  loader: glob({ pattern: '**/*.md', base: './src/content/blog' }),
  schema: z.object({
    title: z.string(),
    description: z.string(),
    produkt: z.string(),
    datum: z.coerce.date(),
    entwurf: z.boolean().default(false),
  }),
});

export const collections = { rechtliches, faq, blog };

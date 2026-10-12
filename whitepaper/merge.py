"""Setzt Cover und Innenseiten zum fertigen Whitepaper zusammen (pip install pymupdf)."""
import os
import pymupdf

hier = os.path.dirname(os.path.abspath(__file__))
ziel = os.path.join(hier, 'Whitepaper-KI-Projekte-Schmal-Unternehmensberatung.pdf')

doc = pymupdf.open(os.path.join(hier, '.cover.pdf'))
doc.insert_pdf(pymupdf.open(os.path.join(hier, '.rest.pdf')))
doc.set_metadata({
    'title': '5 Gründe, warum die meisten KI-Projekte bei Mittelständlern und Freiberuflern scheitern',
    'author': 'Eduard Schmal, Schmal Unternehmensberatung',
    'subject': 'Whitepaper KI-Projekte im Mittelstand',
})
doc.save(ziel, garbage=3, deflate=True)
for teil in ('.cover.pdf', '.rest.pdf'):
    os.remove(os.path.join(hier, teil))
print(ziel, len(doc), 'Seiten')

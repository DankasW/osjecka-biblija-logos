# Osječka Biblija za Logos

Osječka Biblija pripremljena kao osobna knjiga (Personal Book) za [Logos Bible Software](https://www.logos.com/). Svaki stih ima Logos milestone, pa se knjiga u Logosu ponaša kao prava Biblija: prati reference, radi u paralelnom prikazu i povezuje se s komentarima.

Ovo je izvedeno djelo projekta **Osječka Biblija**:

- čitanje na webu: [biblija.hr](https://biblija.hr)
- izvorni podaci: [github.com/dario-nop/osjecka_biblija](https://github.com/dario-nop/osjecka_biblija)

Sav tekst Biblije dolazi iz njihovog repozitorija. Ovdje je samo obrada za Logos.

## Što je u repozitoriju

| Datoteka | Sadržaj |
| --- | --- |
| `osjecka_biblija_logos.docx` | gotova knjiga za Logos (66 knjiga, 31.084 stiha), iz verzije 2.0 od 24. 11. 2026. |
| `obh_json_u_md.py` | pretvara službeni JSON u mape s jednom Markdown datotekom po poglavlju |
| `logos_bible.py` | iz tih mapa radi `.docx` s Logos oznakama |
| `LICENSE` | CC BY-SA 4.0, ista licenca kao izvornik |

## Kako dodati u Logos

1. Preuzmi `osjecka_biblija_logos.docx`.
2. U Logosu otvori **Tools → Personal Books** i dodaj novu knjigu.
3. Dodaj `.docx`, upiši naslov (npr. "Osječka Biblija") i pokreni **Build Book**.

Knjiga koristi Logos datatype `bible` (standardna numeracija). Build prolazi bez upozorenja.

## Kako ponovo napraviti knjigu (nova verzija izvornika)

Potreban je Python 3 i paket `python-docx` (`pip install python-docx`).

```
git clone https://github.com/dario-nop/osjecka_biblija
python obh_json_u_md.py osjecka_biblija/osjecka_biblija_v2.0.json "Osječka Biblija"
python logos_bible.py "Osječka Biblija" osjecka_biblija_logos.docx --type bible
```

`obh_json_u_md.py` zapisuje i `id_stiha.tsv`, popis stalnih ID-jeva stihova iz izvornika, da se izmjene u novim verzijama lako pronađu.

## Kako je tekst obrađen

- Svaki stih dobiva milestone `[[@bible:Ge 1:1]]`, a sam tekst stiha je između `{{field-on:bible}}` i `{{field-off:bible}}`. Broj stiha ostaje vidljiv, ali je izvan tog polja.
- Zapisi s tekstom `#` se preskaču. Prema napomeni u izvornom repozitoriju oni služe samo za poravnanje s drugim numeracijama i nemaju svoj sadržaj.
- Tekst stiha se ne mijenja. Navodnici i interpunkcija ostaju kao u izvorniku.
- Hebrejski i grčki izvornik i transliteracija iz izvornog JSON-a **nisu** uključeni u ovu knjigu.

## Poznata ograničenja

Na nekoliko mjesta izvornik broji stihove po hebrejskom poglavlju, a ne po standardnoj numeraciji (npr. Jona 1,17 je u izvorniku Jona 2,1; isto vrijedi za Br 29,40, 1 Sam 23,29, Job 41,27–34, Pj 6,13, Ez 20,45–49, Hoš 11,12, Dj 19,41 i 2 Kor 13,14). Logos te stihove prihvaća, ali je u paralelnom prikazu s drugim Biblijama tekst ondje pomaknut za jedan stih. Ispravak treba doći u izvornim podacima, pa je javljen uredništvu.

## Licenca i navođenje izvora

Hrvatski tekst: **Osječka Biblija, biblija.hr, CC BY-SA 4.0.** Izmijenjena inačica ekumenskog prijevoda Tomislava Dretara (1998.), izvorno objavljenog na hrvatskom Wikizvoru pod licencom CC BY-SA. Modernizacija i uređivanje: biblija.hr. Izvor podataka: [github.com/dario-nop/osjecka_biblija](https://github.com/dario-nop/osjecka_biblija). Cjeloviti uvjeti: [biblija.hr/licenca-i-izvori](https://biblija.hr/licenca-i-izvori).

Ova obrada (Logos oznake i skripte) objavljena je pod istom licencom, [CC BY-SA 4.0](LICENSE). Tko je dijeli ili mijenja, navodi izvor kao gore i dijeli pod istom licencom.

---

## English summary

The Osječka Biblija (a Croatian Bible translation) prepared as a Personal Book for Logos Bible Software. Every verse carries a Logos milestone (datatype `bible`), so the book works as a real Bible in Logos. All text comes from the [Osječka Biblija project](https://github.com/dario-nop/osjecka_biblija) ([biblija.hr](https://biblija.hr)), licensed CC BY-SA 4.0; this derived work uses the same license. Download `osjecka_biblija_logos.docx` and add it under Tools → Personal Books.

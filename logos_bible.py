"""Tagira cijelu Bibliju iz Markdown mape (jedna .md datoteka po poglavlju) u Logos PB .docx.

Ulaz:  mapa s podmapama "01 Postanak", "02 Izlazak", ... i datotekama "Postanak 1.md", ...
       Stih = redak koji pocinje s **N**. Sve ostalo (naslovi #, brojevi u naslovima,
       oznake <5a>) nije stih i ne dobiva milestone.
Izlaz: <izlaz>.docx i <izlaz>_report.md. Ulazna mapa se nikad ne mijenja.

Pokretanje: python logos_bible.py "KS Biblija" "Biblija KS (tagirano).docx" [--type bible+ehsg]
Test:       python logos_bible.py --selftest
"""
import datetime
import os
import re
import sys

import docx
from docx.enum.text import WD_BREAK

MARKER = "[Logos PB by D@nkas]"
VERSION = "1.0"  # povecaj kad mijenjas pravila, da znas kojom verzijom je fajl napravljen

# Ime mape (bez broja) -> Logos kratica. Isti oblici kao u rucno tagiranim KS fajlovima, bez njihovih gresaka.
BOOKS = {
    "Postanak": "Ge", "Izlazak": "Ex", "Levitski zakonik": "Lev", "Brojevi": "Nu", "Ponovljeni zakon": "Deut",
    "Jošua": "Jos", "Suci": "Jdg", "Ruta": "Ru", "1 Samuelova": "1Sam", "2 Samuelova": "2Sam",
    "1 Kraljevima": "1Kgs", "2 Kraljevima": "2Kgs", "1 Ljetopisa": "1Chr", "2 Ljetopisa": "2Chr",
    "Ezra": "Ezr", "Nehemija": "Neh", "Tobija": "Tob", "Judita": "Jdt", "Estera": "Est",
    "1 Makabejska": "1Mac", "2 Makabejska": "2Mac", "Job": "Job", "Psalmi": "Ps", "Mudre izreke": "Pro",
    "Propovjednik": "Ecc", "Pjesma nad pjesmama": "Song", "Mudrost": "Wis", "Sirah": "Sir",
    "Izaija": "Isa", "Jeremija": "Jer", "Tužaljke": "Lam", "Baruh": "Bar", "Ezekiel": "Ezk",
    "Daniel": "Dn", "Hošea": "Hos", "Joel": "Jl", "Amos": "Am", "Obadija": "Ob", "Jona": "Jon",
    "Mihej": "Mic", "Nahum": "Nah", "Habakuk": "Hab", "Sefanija": "Zep", "Hagaj": "Hag",
    "Zaharija": "Zec", "Malahija": "Mal",
    "Matej": "Mt", "Marko": "Mk", "Luka": "Lk", "Ivan": "Jn", "Djela apostolska": "Act",
    "Rimljanima": "Rom", "1 Korinćanima": "1Cor", "2 Korinćanima": "2Cor", "Galaćanima": "Gal",
    "Efežanima": "Eph", "Filipljanima": "Php", "Kološanima": "Col", "1 Solunjanima": "1Thess",
    "2 Solunjanima": "2Thess", "1 Timoteju": "1Tim", "2 Timoteju": "2Tim", "Titu": "Tit",
    "Filemonu": "Phm", "Hebrejima": "Heb", "Jakovljeva": "Jas", "1 Petrova": "1Pe", "2 Petrova": "2Pe",
    "1 Ivanova": "1Jn", "2 Ivanova": "2Jn", "3 Ivanova": "3Jn", "Judina": "Jd", "Otkrivenje": "Rev",
}

VERSE = re.compile(r"^\*\*(\d{1,3})\*\*\s*(.*)$")  # jedino ovo je broj stiha
HEAD = re.compile(r"^(#{1,6})\s+(.*)$")
ITAL = re.compile(r"\*([^*\n]+)\*")
PARALLEL = re.compile(r"^\(.*\)$")
MARK = re.compile(r"<\d{1,3}[a-z]{1,2}>")  # premjesteni dijelovi stihova (KS): urednicka oznaka, ne ide u field
CAPS = re.compile(r"^(?=.*[A-ZČĆŽŠĐ]{3})[^a-zčćžšđ]+$")  # podnaslov bez #, npr. "A. MERIBAAL"
FON, FOFF ="{{field-on:bible}}", "{{field-off:bible}}"


def num_key(name, last=False):
    """Prvi broj (mapa "09 1 Samuelova") ili zadnji (poglavlje "1 Samuelova 2.md")."""
    n = re.findall(r"\d+", name)
    return int(n[-1 if last else 0]) if n else 0


def blocks(text):
    """Odlomci odvojeni praznim redom; svaki je lista redaka."""
    return [b.splitlines() for b in re.split(r"\n\s*\n", text.replace("\r", "").strip()) if b.strip()]


class Book:
    def __init__(self, d, dtype, report):
        self.d, self.dtype, self.rep = d, dtype, report

    def ms(self, ref):
        return f"[[@{self.dtype}:{ref}]]"

    def chapter(self, abbr, ch, text, where, single=False):
        d, rep = self.d, self.rep
        last = 0  # zadnji stih u poglavlju
        after_head = False
        for lines in blocks(text):
            h = HEAD.match(lines[0])
            if after_head and PARALLEL.match(lines[0]):
                d.add_paragraph(lines.pop(0))  # paralelna mjesta ispod podnaslova, npr. (Mk 1, 1–8)
                if not lines:
                    continue
            after_head = bool(h) or all(CAPS.match(x) for x in lines)
            if after_head:  # naslov (i bez #, npr. "A. MERIBAAL"); moze imati vise redaka; nikad se ne tagira
                level = len(h.group(1)) if h else 4
                title = " ".join([h.group(2).strip() if h else lines[0].strip()] + [x.strip() for x in lines[1:]])
                if level == 1:
                    if not single:  # u knjizi s jednim poglavljem "Phm 1" je stih 1 (Logos: duplicate milestone)
                        d.add_paragraph(self.ms(f"{abbr} {ch}"))
                    d.add_heading(f"[[{title}>>{abbr} {ch}]]", 2)
                else:
                    d.add_heading(title, min(level + 1, 9))
                continue
            p = d.add_paragraph()
            st = {"open": False}

            def put(s):
                """Biblijski tekst ide u field; oznake <5a> ostaju izvan fielda."""
                for piece in re.split(f"({MARK.pattern})", s):
                    if not piece:
                        continue
                    is_mark = bool(MARK.fullmatch(piece))
                    if is_mark == st["open"]:  # mijenja se vrsta teksta: otvori ili zatvori field
                        p.add_run(FOFF if st["open"] else FON)
                        st["open"] = not is_mark
                    for j, bit in enumerate(ITAL.split(piece)):  # *rijec* = kurziv (VAB: dodane rijeci)
                        if bit:
                            p.add_run(bit).italic = j % 2 == 1 or None

            for i, line in enumerate(lines):
                m = VERSE.match(line)
                if i and not m:  # novi redak pjesme; novi stih u istom odlomku ide u isti redak (nize)
                    p.add_run().add_break(WD_BREAK.LINE)
                if m:
                    v = int(m.group(1))
                    if v != last + 1:
                        rep.append(f"- {where}: stih {v} dolazi poslije {last}")
                    last = v
                    if not m.group(2).strip():
                        rep.append(f"- {where}: stih {v} je prazan u izvoru (milestone bez fielda)")
                    if st["open"]:
                        p.add_run(FOFF)
                        st["open"] = False
                    if i:
                        p.add_run(" ")
                    p.add_run(self.ms(f"{abbr} {ch}:{v}"))
                    p.add_run(str(v)).font.superscript = True
                    p.add_run(" ")
                    put(m.group(2))
                elif last:  # nastavak stiha (npr. iza podnaslova ili novi red pjesme)
                    put(line)
                else:
                    rep.append(f"- {where}: tekst prije prvog stiha ostao bez taga: {line[:60]!r}")
                    p.add_run(line)
            if st["open"]:
                p.add_run(FOFF)
        return last


def run(src, out, dtype="bible+ehsg"):
    d = docx.Document()
    d.add_paragraph(MARKER)
    d.add_paragraph(f"Logos bible tagger ver. {VERSION} ({datetime.date.today():%Y-%m-%d})")
    izvor = os.path.join(src, "IZVOR.txt")  # navod izvora/licence (npr. CC BY-SA), ako ga preuzimanje zapise
    if os.path.exists(izvor):
        with open(izvor, encoding="utf-8") as fh:
            d.add_paragraph(fh.read().strip())
    rep, verses = [], 0
    bk = Book(d, dtype, rep)
    for folder in sorted(os.listdir(src), key=num_key):
        path = os.path.join(src, folder)
        if not os.path.isdir(path):
            continue
        name = re.sub(r"^\d+\s*", "", folder)
        abbr = BOOKS.get(name)
        if not abbr:
            rep.append(f"- PRESKOCENO: nepoznata knjiga {folder!r} (dodaj je u BOOKS)")
            continue
        d.add_heading(name, 1)
        files = sorted((f for f in os.listdir(path) if f.endswith(".md")), key=lambda f: num_key(f, True))
        for f in files:
            with open(os.path.join(path, f), encoding="utf-8") as fh:
                verses += bk.chapter(abbr, num_key(f, True), fh.read(), f[:-3], single=len(files) == 1)
    d.save(out)
    rp = os.path.splitext(out)[0] + "_report.md"
    with open(rp, "w", encoding="utf-8") as fh:
        fh.write(f"# Izvjestaj: {os.path.basename(out)}\n\nIzvor: {src}\nTip: @{dtype}\n"
                 f"Stihova (zbroj zadnjih stihova po poglavlju): {verses}\n\n## Za rucnu provjeru ({len(rep)})\n\n")
        fh.write("\n".join(rep) + "\n")
    print(f"Gotovo: {out}\nIzvjestaj: {rp} ({len(rep)} stavki)")


def selftest():
    d = docx.Document()
    rep = []
    md = ("# Postanak 2\n\n## Naslov 3. dio\n\n**1** Prvi.\n\n**2** Drugi\nred dva.\n\n"
          "#### Pod 4\ndrugi red (1. dio)\n\n(Mk 1, 2–3)\n\nNastavak.\n\n**4** Cetiri <5a> pet.\n\n"
          "A. VELIKI\n\n**5** \n\n#### X\n\n(Mt 1, 1)\n<5b> Dalje.\n")
    last = Book(d, "bible+njb", rep).chapter("Ge", 2, md, "Postanak 2")
    t = [p.text for p in d.paragraphs]
    assert t[0] == "[[@bible+njb:Ge 2]]" and t[1] == "[[Postanak 2>>Ge 2]]", t
    assert t[2] == "Naslov 3. dio", t  # broj u naslovu se ne tagira
    assert t[3] == "[[@bible+njb:Ge 2:1]]1 {{field-on:bible}}Prvi.{{field-off:bible}}", t[3]
    assert t[4] == "[[@bible+njb:Ge 2:2]]2 {{field-on:bible}}Drugi\nred dva.{{field-off:bible}}", t[4]
    assert t[5] == "Pod 4 drugi red (1. dio)", t[5]
    assert t[6] == "(Mk 1, 2–3)", t[6]
    assert t[7] == "{{field-on:bible}}Nastavak.{{field-off:bible}}", t[7]
    assert t[8] == "[[@bible+njb:Ge 2:4]]4 {{field-on:bible}}Cetiri {{field-off:bible}}<5a>{{field-on:bible}} pet.{{field-off:bible}}", t[8]
    assert t[9] == "A. VELIKI" and t[10] == "[[@bible+njb:Ge 2:5]]5 ", t[9:11]
    assert t[12] == "(Mt 1, 1)" and t[13] == "<5b>{{field-on:bible}} Dalje.{{field-off:bible}}", t[12:]
    d2 = docx.Document()
    Book(d2, "Bible", []).chapter("Ge", 3, "# Postanak 3\n\n**1** Ti si *mi* dao.\n", "x")
    p = d2.paragraphs[-1]
    assert p.text == "[[@Bible:Ge 3:1]]1 {{field-on:bible}}Ti si mi dao.{{field-off:bible}}", p.text
    assert [r.text for r in p.runs if r.italic] == ["mi"], [(r.text, r.italic) for r in p.runs]
    Book(d2, "Bible", []).chapter("Ge", 4, "**1** A.\n**2** B.\n", "x")  # dva stiha u jednom odlomku
    assert d2.paragraphs[-1].text == ("[[@Bible:Ge 4:1]]1 {{field-on:bible}}A.{{field-off:bible}} "
                                      "[[@Bible:Ge 4:2]]2 {{field-on:bible}}B.{{field-off:bible}}"), d2.paragraphs[-1].text
    assert last == 5 and any("stih 4 dolazi poslije 2" in r for r in rep), rep
    assert num_key("1 Ljetopisa 2.md", True) == 2 and num_key("13 1 Ljetopisa") == 13
    print("selftest OK")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    a = sys.argv[1:]
    if a[:1] == ["--selftest"]:
        selftest()
    elif len(a) >= 2:
        run(a[0], a[1], a[a.index("--type") + 1] if "--type" in a else "bible+ehsg")
    else:
        print(__doc__)

"""Osjecka Biblija iz sluzbenog JSON-a (github.com/dario-nop/osjecka_biblija) u Markdown:
<izlaz>/NN Knjiga/Knjiga N.md, isti oblik kao KS/VAB, pa ga logos_bible.py tagira bez promjena.

Zapisi s tekstom "#" su samo poravnanje s drugim numeracijama (README repoa) i preskacu se.
id_stiha se ne gubi: zapisuje se u <izlaz>/id_stiha.tsv (knjiga, poglavlje, stih, id) za buduca azuriranja.

Pokretanje: python obh_json_u_md.py osjecka_biblija_v2.0.json "Osječka Biblija"
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from logos_bible import BOOKS  # noqa: E402

DEUTERO = {"Tobija", "Judita", "1 Makabejska", "2 Makabejska", "Mudrost", "Sirah", "Baruh"}
FOLDERS = [b for b in BOOKS if b not in DEUTERO]  # 66 knjiga redom kao id_knjige 1..66

ATTRIBUTION = ("Osječka Biblija, biblija.hr, CC BY-SA 4.0. Hrvatski tekst: izmijenjena inačica ekumenskog "
               "prijevoda Tomislava Dretara (1998.), izvorno objavljenog na hrvatskom Wikizvoru pod licencom "
               "CC BY-SA; modernizacija i uređivanje: biblija.hr. Izvor podataka: github.com/dario-nop/osjecka_biblija")


def convert(records, dst):
    chapters = {}  # (id_knjige, poglavlje) -> [(stih, tekst)]
    for r in records:
        if r["tekst"].strip() == "#":
            continue
        chapters.setdefault((r["id_knjige"], r["poglavlje"]), []).append((r["stih"], r["tekst"].strip(), r["id_stiha"]))
    ids = []
    for (b, ch), verses in sorted(chapters.items()):
        name = FOLDERS[b - 1]
        d = os.path.join(dst, f"{b:02d} {name}")
        os.makedirs(d, exist_ok=True)
        with open(os.path.join(d, f"{name} {ch}.md"), "w", encoding="utf-8") as fh:
            fh.write("\n".join([f"# {name} {ch}", ""] + [f"**{v}** {t}\n" for v, t, _ in verses]))
        ids += [f"{name}\t{ch}\t{v}\t{i}" for v, _, i in verses]
    with open(os.path.join(dst, "IZVOR.txt"), "w", encoding="utf-8") as fh:
        fh.write(ATTRIBUTION + f" (verzija {records[0]['verzija']}, {records[0]['datum']})\n")
    with open(os.path.join(dst, "id_stiha.tsv"), "w", encoding="utf-8") as fh:
        fh.write("knjiga\tpoglavlje\tstih\tid_stiha\n" + "\n".join(ids) + "\n")
    return len(chapters), len(ids)


def selftest(tmp):
    recs = [dict(id_stiha=1, id_knjige=2, poglavlje=7, stih=25, tekst="A.", verzija="2.0", datum="d"),
            dict(id_stiha=2, id_knjige=2, poglavlje=7, stih=26, tekst="#", verzija="2.0", datum="d")]
    assert len(FOLDERS) == 66 and FOLDERS[18] == "Psalmi" and FOLDERS[65] == "Otkrivenje"
    assert convert(recs, tmp) == (1, 1)
    with open(os.path.join(tmp, "02 Izlazak", "Izlazak 7.md"), encoding="utf-8") as fh:
        assert fh.read() == "# Izlazak 7\n\n**25** A.\n", "md"


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        selftest(tmp)
    if sys.argv[1:] != ["--selftest"]:
        with open(sys.argv[1], encoding="utf-8") as fh:
            data = json.load(fh)
        assert len({r["id_knjige"] for r in data}) <= 66
        print("Poglavlja: %d, stihova: %d" % convert(data, sys.argv[2]))

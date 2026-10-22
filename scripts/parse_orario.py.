#!/usr/bin/env python3
"""Converte l'export Excel dell'orario IULM in lezioni.json.

Uso: python scripts/parse_orario.py export.xlsx lezioni.json

- Se l'orario non e' cambiato rispetto a lezioni.json, non scrive nulla.
- Se e' cambiato, aggiorna lezioni.json e scrive il riepilogo in commit_msg.txt.
- Se il file sembra sbagliato (poche lezioni), si ferma con errore senza toccare nulla.
"""
import datetime
import json
import os
import re
import sys

import openpyxl

MIN_LESSONS = 20  # sotto questa soglia l'export e' considerato anomalo


def parse(xlsx_path):
    ws = openpyxl.load_workbook(xlsx_path, data_only=True).active
    lessons, seen = [], set()
    for row in ws.iter_rows(values_only=True):
        c = list(row) + [None] * 8
        giorno, ora, materia, aula = c[1], c[2], c[3], c[6]
        if not (isinstance(giorno, str) and isinstance(ora, str) and isinstance(materia, str)):
            continue
        try:
            d = datetime.datetime.strptime(giorno.strip(), "%d-%m-%Y").date()
        except ValueError:
            continue
        m = re.match(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})", ora.strip())
        if not m:
            continue
        h1, m1, h2, m2 = map(int, m.groups())
        subject = re.sub(r"\s+", " ", materia.split("\n")[0].strip())
        lesson = {
            "date": d.isoformat(),
            "start": f"{h1:02d}:{m1:02d}",
            "end": f"{h2:02d}:{m2:02d}",
            "subject": subject,
            "room": (aula or "").strip(),
            "cancelled": "Annullato" in materia,
        }
        key = (lesson["date"], lesson["start"], lesson["end"], subject)
        if key in seen:
            continue
        seen.add(key)
        lessons.append(lesson)
    lessons.sort(key=lambda l: (l["date"], l["start"], l["subject"]))
    return lessons


def short(l):
    name = l["subject"].replace(" (Gruppo 1)", "").replace(" GRUPPO 1", "")
    return f'{l["date"]} {l["start"]}-{l["end"]} {name} ({l["room"] or "aula n.d."})'


def diff(old, new):
    key = lambda l: (l["date"], l["subject"], l["start"])
    om, nm = {key(l): l for l in old}, {key(l): l for l in new}
    lines = []
    for k in sorted(set(nm) - set(om)):
        lines.append("+ aggiunta: " + short(nm[k]))
    for k in sorted(set(om) - set(nm)):
        lines.append("- rimossa: " + short(om[k]))
    for k in sorted(set(om) & set(nm)):
        a, b = om[k], nm[k]
        ch = []
        if a["end"] != b["end"]:
            ch.append(f'fine {a["end"]} -> {b["end"]}')
        if a["room"] != b["room"]:
            ch.append(f'aula {a["room"] or "n.d."} -> {b["room"] or "n.d."}')
        if a["cancelled"] != b["cancelled"]:
            ch.append("ANNULLATA" if b["cancelled"] else "non piu' annullata")
        if ch:
            lines.append("~ modificata: " + short(b) + " [" + ", ".join(ch) + "]")
    return lines


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    xlsx, out = sys.argv[1], sys.argv[2]
    new = parse(xlsx)
    if len(new) < MIN_LESSONS:
        sys.exit(f"ERRORE: trovate solo {len(new)} lezioni (minimo {MIN_LESSONS}). "
                 "Il file scaricato non sembra un orario valido: non aggiorno nulla.")

    old = []
    if os.path.exists(out):
        with open(out, encoding="utf-8") as f:
            old = json.load(f).get("lessons", [])

    if old == new:
        print(f"Nessuna modifica ({len(new)} lezioni).")
        return

    changes = diff(old, new) if old else ["Primo caricamento dell'orario"]
    data = {
        "updated": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "lessons": new,
    }
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
        f.write("\n")

    summary = "Aggiornamento orario\n\n" + "\n".join(changes) + "\n"
    with open("commit_msg.txt", "w", encoding="utf-8") as f:
        f.write(summary)
    print(summary)


if __name__ == "__main__":
    main()

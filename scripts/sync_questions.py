"""Convert the quiz spreadsheet (xlsx) into questions.json for the quiz page.

Usage: python scripts/sync_questions.py quiz.xlsx questions.json

Columns (header row is found automatically, order doesn't matter):
  Muzyk (poprawna odpowiedź) | Tytuł utworu | Sygnatura czasowa | Link YouTube | Link Spotify |
  Błędna odpowiedź 1 | Błędna odpowiedź 2 | Błędna odpowiedź 3 | Instrument | Album (rok) | Ciekawostka
The old layout with answer columns A–D is still understood.
Rows without a YouTube link or a timestamp are skipped until they're filled in.
"""
import datetime as dt
import json
import re
import sys

import openpyxl


def norm(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def youtube_id(link):
    link = norm(link).replace("\\", "")
    for pat in (r"youtu\.be/([\w-]{11})", r"[?&]v=([\w-]{11})", r"/(?:embed|shorts|live)/([\w-]{11})"):
        m = re.search(pat, link)
        if m:
            return m.group(1)
    return None


def seconds(v):
    """Timestamp → seconds. Handles '1:29', '1:02:03', time/timedelta cells and day fractions."""
    if isinstance(v, dt.datetime):
        v = v.time()
    if isinstance(v, dt.time):
        h, m, s = v.hour, v.minute, v.second
        # Sheets/Excel turn a typed "1:29" or "0:41" into h:mm:00; a solo timestamp is meant as m:ss
        return h * 60 + m if s == 0 else h * 3600 + m * 60 + s
    if isinstance(v, dt.timedelta):
        total = int(round(v.total_seconds()))
        h, rest = divmod(total, 3600)
        m, s = divmod(rest, 60)
        return h * 60 + m if s == 0 else total
    if isinstance(v, (int, float)):
        if 0 < v < 1:  # fraction of a day
            return seconds(dt.timedelta(days=v))
        return int(v)
    parts = [int(p) for p in re.findall(r"\d+", str(v))]
    if not parts:
        return None
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return parts[0] * 60 + parts[1]
    h, m, s = parts[:3]
    return h * 60 + m if s == 0 else h * 3600 + m * 60 + s


def main(src, dst):
    ws = openpyxl.load_workbook(src, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    header_idx = next(i for i, r in enumerate(rows) if any("youtube" in norm(c).lower() for c in r))
    header = [norm(c).lower() for c in rows[header_idx]]

    def col(*names):
        for i, h in enumerate(header):
            if any(n in h for n in names):
                return i
        return None

    c_name = col("nazwisko", "muzyk", "poprawna")
    c_title = col("tytuł", "tytul")
    c_time = col("sygnatura", "czas")
    c_yt = col("youtube")
    c_instr = col("instrument")
    c_album = col("album")
    c_fact = col("ciekawostk", "opis", "kontekst")
    c_wrong = [i for i, h in enumerate(header)
               if h.startswith(("błędna", "bledna", "zła", "zla")) or h in ("a", "b", "c", "d")]

    questions = []
    for r in rows[header_idx + 1:]:
        get = lambda i: r[i] if i is not None and i < len(r) else None
        answer, vid = norm(get(c_name)), youtube_id(get(c_yt))
        raw_time = get(c_time)
        if not answer or not vid or raw_time is None or norm(raw_time) == "":
            continue
        start = seconds(raw_time)
        if start is None:
            continue
        opts = [answer]
        for i in c_wrong:
            o = norm(get(i))
            if o and o not in opts:
                opts.append(o)
        q = {"answer": answer, "title": norm(get(c_title)), "youtube": vid, "start": start, "options": opts}
        for key, c in (("instrument", c_instr), ("album", c_album), ("fact", c_fact)):
            val = norm(get(c))
            if val:
                q[key] = val
        questions.append(q)

    if not questions:
        sys.exit("No questions found — not overwriting questions.json")
    with open(dst, "w", encoding="utf-8") as f:
        json.dump(questions, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print(f"Wrote {len(questions)} questions to {dst}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])

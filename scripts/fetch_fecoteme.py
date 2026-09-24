#!/usr/bin/env python3
"""Download new ranking workbooks published on the FECOTEME website.

Reads each page listed in sources.json ("pages"), finds every .xlsx link, and
downloads the ones not seen before into inbox/<season>/. A small "<file>.json"
sidecar is written next to each download with what the link label says
(stage, circuit, category, gender), which convert.py uses as its first guess.

Already-downloaded URLs are remembered in data/fetched.json, so a file is
only downloaded once. A corrected file that FECOTEME re-uploads gets a new URL
and is picked up again (and replaces the earlier stage).
"""
import html
import json
import re
import sys
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SOURCES = ROOT / "sources.json"
SEEN = ROOT / "data" / "fetched.json"
INBOX = ROOT / "inbox"
UA = {"User-Agent": "Mozilla/5.0 (ranking-tm importer; +https://github.com)"}

ORD = {"1er": 1, "1ro": 1, "2do": 2, "3er": 3, "3ro": 3, "3cer": 3, "4to": 4, "5to": 5, "6to": 6, "7mo": 7, "8vo": 8}
CIRCUITS = {"menor": "Liga Menor", "mayor": "Liga Mayor", "open_femenino": "Open Femenino", "open": "Open Femenino",
            "master": "Master", "ptt": "PTT"}


def get(url):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def parse_label(label):
    """'4to_ranking_nacional_menor_u13_masc.xlsx' -> stage/circuit/division/gender."""
    t = label.lower().replace(".xlsx", "").replace("-", "_").replace(" ", "_")
    meta = {}
    m = re.match(r"(\d+(?:er|ro|do|to|cer|mo|vo))_ranking_nacional_(.+)$", t)
    if not m:
        return meta
    meta["stage"] = ORD.get(m.group(1))
    rest = m.group(2)
    for key in sorted(CIRCUITS, key=len, reverse=True):
        if rest.startswith(key):
            meta["circuit"] = CIRCUITS[key]
            rest = rest[len(key):].strip("_")
            break
    if re.search(r"(^|_)fem", rest):
        meta["gender"] = "F"
    elif re.search(r"(^|_)masc", rest):
        meta["gender"] = "M"
    mu = re.search(r"u_?(\d{1,2})", rest)
    if mu:
        meta["division"] = f"U{mu.group(1)}"
    elif re.search(r"open_?a", rest):
        meta["division"] = "Open A"
    elif re.search(r"open_?b", rest):
        meta["division"] = "Open B"
    return {k: v for k, v in meta.items() if v}


def links(page_html, base):
    for m in re.finditer(r'<a\b[^>]*href=["\']([^"\']+\.xlsx)["\'][^>]*>(.*?)</a>', page_html, re.I | re.S):
        href = urllib.parse.urljoin(base, html.unescape(m.group(1)))
        text = re.sub(r"<[^>]+>", " ", m.group(2))
        text = " ".join(html.unescape(text).split())
        yield href, text


def main():
    cfg = json.loads(SOURCES.read_text(encoding="utf-8"))
    seen = json.loads(SEEN.read_text(encoding="utf-8")) if SEEN.exists() else {}
    new = 0
    for page in cfg.get("pages", []):
        url, season = page["url"], int(page["season"])
        try:
            body = get(url).decode("utf-8", "replace")
        except Exception as e:  # noqa: BLE001
            print(f"WARN could not read {url}: {e}")
            continue
        found = list(links(body, url))
        print(f"{url}: {len(found)} .xlsx links")
        for href, text in found:
            if href in seen:
                continue
            fname = urllib.parse.unquote(href.rsplit("/", 1)[-1])
            fname = re.sub(r"[^\w.\-]+", "_", fname)
            dest_dir = INBOX / str(season)
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest = dest_dir / fname
            try:
                dest.write_bytes(get(href))
            except Exception as e:  # noqa: BLE001
                print(f"WARN could not download {href}: {e}")
                continue
            hint = parse_label(text) or parse_label(fname)
            hint.update({"season": season, "sourceUrl": href, "sourceLabel": text})
            (dest_dir / (fname + ".hint.json")).write_text(json.dumps(hint, ensure_ascii=False, indent=1), encoding="utf-8")
            seen[href] = {"label": text, "file": fname, "season": season}
            new += 1
            print(f"NEW {text or fname} -> inbox/{season}/{fname}")
    SEEN.parent.mkdir(parents=True, exist_ok=True)
    SEEN.write_text(json.dumps(seen, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{new} new file(s) downloaded")


if __name__ == "__main__":
    sys.exit(main())

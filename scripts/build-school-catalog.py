#!/usr/bin/env python3
"""Build the national folio.exam school-name catalog from MEXT's latest published code list."""
from __future__ import annotations
import json
import re
import urllib.request
from pathlib import Path

import pandas as pd

BASE = "https://www.mext.go.jp/content/20260529-mxt_chousa01-000011635"
FILES = [BASE + "_1.xlsx", BASE + "_3.xlsx"]
OUT = Path("school-catalog.js")
META = Path("school-catalog-meta.json")
SHEET_HEADER = 1
SOURCE_URL = "https://www.mext.go.jp/b_menu/toukei/mext_01087.html"
SOURCE_LABEL = "文部科学省 学校コード一覧（令和8年5月1日時点・暫定版）"

TYPE_MAP = {"1(国)": "国立", "2(公)": "公立", "3(私)": "私立"}
INCLUDED = {"D1(高校)", "D2(中等)", "E1(養護)", "E1(特支盲)", "E1(特支聾)"}

def clean(v: object) -> str:
    return str(v).strip() if v is not None else ""

def get_place(address: str) -> str:
    if not address:
        return ""
    m = re.match(r"^(北海道|東京都|京都府|大阪府|.{2,3}県)(.*)$", address)
    if not m:
        return ""
    pref, rest = m.groups()
    patterns = (
        r"^(.+?市.+?区)",
        r"^(.+?市)",
        r"^(.+?郡.+?[町村])",
        r"^(.+?区)",
        r"^(.+?[町村])",
    )
    for pat in patterns:
        hit = re.match(pat, rest)
        if hit:
            return pref + hit.group(1)
    return pref

def download(url: str, path: Path) -> None:
    request = urllib.request.Request(url, headers={"User-Agent": "folio-exam-school-catalog/1.0"})
    with urllib.request.urlopen(request, timeout=90) as response:
        path.write_bytes(response.read())

def main() -> None:
    records: dict[str, list[str]] = {}
    raw_counts: dict[str, int] = {}
    for n, url in enumerate(FILES, start=1):
        path = Path(f"/tmp/mext-school-codes-{n}.xlsx")
        download(url, path)
        frame = pd.read_excel(path, sheet_name=0, header=SHEET_HEADER, dtype=str).fillna("")
        required = ["学校コード", "学校種", "設置\n区分", "本分校", "学校名", "学校所在地", "郵便番号", "属性情報\n廃止年月日"]
        missing = [c for c in required if c not in frame.columns]
        if missing:
            raise RuntimeError(f"Unexpected MEXT workbook headers in {url}: {missing}")
        for _, row in frame.iterrows():
            code = clean(row["学校コード"])
            kind = clean(row["学校種"])
            branch = clean(row["本分校"])
            if kind not in INCLUDED or branch not in {"1(本)", "2(分)"}:
                continue
            if clean(row["属性情報\n廃止年月日"]):
                continue
            if not code or not clean(row["学校名"]):
                continue
            address = clean(row["学校所在地"])
            code_kind = kind.split("(", 1)[0]
            subtype = kind[kind.find("(") + 1:kind.rfind(")")] if "(" in kind else kind
            record = [
                clean(row["学校名"]),
                get_place(address),
                TYPE_MAP.get(clean(row["設置\n区分"]), "区分未確認"),
                code,
                address,
                code_kind,
                clean(row["郵便番号"]),
                "本校" if branch == "1(本)" else "分校",
                subtype,
            ]
            if code in records and records[code] != record:
                raise RuntimeError(f"Conflicting duplicate school code: {code}")
            records[code] = record
            raw_counts[code_kind] = raw_counts.get(code_kind, 0) + 1

    rows = list(records.values())
    rows.sort(key=lambda x: (x[4][:3], x[5], x[0], x[4], x[3]))
    codes = [r[3] for r in rows]
    if len(codes) != len(set(codes)):
        raise RuntimeError("Duplicate school codes remain after merge.")

    # Check for exact duplicate identity tuples. Different school codes remain
    # separate only when MEXT assigns distinct identifiers; the audit records them.
    identity_counts: dict[tuple[str, str, str], int] = {}
    for r in rows:
        key = (r[0], r[4], r[5])
        identity_counts[key] = identity_counts.get(key, 0) + 1
    exact_identity_duplicates = sum(v - 1 for v in identity_counts.values() if v > 1)

    js = (
        "/* Generated from the official MEXT school-code list. Do not edit by hand. */\n"
        f"window.FOLIO_EXTRA_SCHOOL_CATALOG = {json.dumps(rows, ensure_ascii=False, separators=(',', ':'))};\n"
        f"window.FOLIO_EXTRA_SCHOOL_CATALOG_SOURCE = {json.dumps({'label': SOURCE_LABEL, 'url': SOURCE_URL, 'published': '2026-05-29', 'asOf': '2026-05-01'}, ensure_ascii=False)};\n"
    )
    OUT.write_text(js, encoding="utf-8")
    meta = {
        "source": SOURCE_LABEL,
        "sourceUrl": SOURCE_URL,
        "published": "2026-05-29",
        "asOf": "2026-05-01",
        "totalUniqueSchoolCodes": len(rows),
        "bySchoolKindCode": {k: sum(1 for r in rows if r[5] == k) for k in sorted({r[5] for r in rows})},
        "byCampusType": {k: sum(1 for r in rows if r[7] == k) for k in ("本校", "分校")},
        "exactNameAddressKindDuplicateRows": exact_identity_duplicates,
        "note": "D1には全日制・定時制・通信制の各課程が含まれます。学校コード一覧だけでは課程別の学校名を重複なく切り分けられないため、1学校コードにつき1件で登録し、課程情報は別資料で確認します。E1は特別支援学校（学部構成はこの一覧では未確認）、D2は中等教育学校です。廃止校は除外しています。",
    }
    META.write_text(json.dumps(meta, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False))

if __name__ == "__main__":
    main()

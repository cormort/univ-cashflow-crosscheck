#!/usr/bin/env python3
"""從別的勾稽表抽出「平衡表科目調整 ↔ 現流項目」配對，評估能不能併進主表。

只讀不改：不動來源檔、不動 勾稽配對表.json。要併是另一個決定。

配對藏在兩個地方，有一個就抽得出來：
  1. 借貸金額是公式、直接指到現流列（`=J50*-1`）——我們這邊約佔八成
  2. C/E 的借貸代號與 I 欄的代號相同者互相配對——約佔兩成
第 2 種只要求「該檔案內部代號自洽」，不需要跟我們的編碼一致。

用法：python3 extract_pairs.py 別人的勾稽表.xlsx [輸出.csv]
"""
import collections
import csv
import json
import pathlib
import re
import sys

import openpyxl

MASTER_FILE = "現流代號主檔.json"
PAIRING_FILE = "勾稽配對表.json"
JREF = re.compile(r"(?<![A-Z$])J(\d+)")
COL = {"科目": 1, "借代號": 3, "借金額": 4, "貸代號": 5, "貸金額": 6,
       "現流項目": 8, "現流代號": 9}


def indent(s):
    n = 0
    for ch in str(s):
        if ch != "　":
            break
        n += 1
    return n


def skeleton(ws):
    """骨架範圍：從 H 欄開始有項目，到 A 欄出現資料區塊表頭（或 H 欄用完）為止。"""
    lo = next((r for r in range(1, 40) if ws.cell(r, COL["現流項目"]).value not in (None, "")), None)
    if lo is None:
        raise SystemExit("✗ 這份檔案的第 8 欄沒有現流項目，版面與預期不符")
    hi = next((r - 2 for r in range(lo + 5, ws.max_row + 1)
               if str(ws.cell(r, 1).value).strip() in ("科目", "項      目")), None)
    if hi is None:                       # 沒有資料區塊就用 H 欄最後一列
        hi = max(r for r in range(lo, ws.max_row + 1)
                 if ws.cell(r, COL["現流項目"]).value not in (None, ""))
    return lo, hi


def extract(path):
    wb = openpyxl.load_workbook(path)
    sheets = [n for n in wb.sheetnames if wb[n].cell(1, COL["現流項目"]).value is not None
              or wb[n].max_column >= COL["現流代號"]]
    ws = wb[sheets[0]] if len(wb.sheetnames) < 3 else wb[wb.sheetnames[2]]
    lo, hi = skeleton(ws)

    stack, item, parent = [], {}, {}
    for r in range(lo, hi + 1):
        h = ws.cell(r, COL["現流項目"]).value
        if h in (None, ""):
            continue
        d = indent(h)
        del stack[d:]
        while len(stack) < d:
            stack.append("")
        stack.append(str(h).strip())
        item[r] = stack[-1]
        parent[r] = next((x for x in reversed(stack[:-1]) if x), "")
    by_code = {str(ws.cell(r, COL["現流代號"]).value).strip(): r for r in range(lo, hi + 1)
               if ws.cell(r, COL["現流代號"]).value not in (None, "")}

    def owner(r):
        for x in range(r, lo - 1, -1):
            v = ws.cell(x, COL["科目"]).value
            if isinstance(v, str) and v.strip() not in ("", "0"):
                return v.strip()
        return "?"

    pairs, seen, how = [], collections.Counter(), collections.Counter()
    for r in range(lo, hi + 1):
        for cc, ac, side in ((COL["借代號"], COL["借金額"], "借"),
                             (COL["貸代號"], COL["貸金額"], "貸")):
            v = ws.cell(r, cc).value
            if v in (None, "") or (isinstance(v, str) and v.startswith("=")):
                continue
            amt = str(ws.cell(r, ac).value or "")
            j = {int(m.group(1)) for m in JREF.finditer(amt)}
            by_f = next(iter(j)) if len(j) == 1 else None
            tgt = by_f if by_f in item else by_code.get(str(v).strip())
            how["公式" if by_f in item else ("代號" if tgt else "無法解析")] += 1
            o = owner(r)
            seen[(o, side)] += 1
            pairs.append({"平衡表科目": o, "借貸": side, "第N筆": seen[(o, side)],
                          "原代號": str(v).strip(),
                          "現流項目": item.get(tgt, ""), "現流上層": parent.get(tgt, ""),
                          "金額公式": amt, "來源": pathlib.Path(path).name})
    return ws.title, (lo, hi), pairs, how


def main():
    path = sys.argv[1]
    title, (lo, hi), pairs, how = extract(path)
    print(f"工作表「{title}」骨架 {lo}–{hi} 列；抽出 {len(pairs)} 筆配對　來源 {dict(how)}\n")

    if not pathlib.Path(MASTER_FILE).exists():
        print(f"（找不到 {MASTER_FILE}，跳過與公版的比對）")
    else:
        master = json.loads(pathlib.Path(MASTER_FILE).read_text())["代號"]
        names = {k.split("\t")[1] for k in master}
        got = {p["現流項目"] for p in pairs if p["現流項目"]}
        hit = got & names
        print(f"① 名稱與公版的重疊：{len(hit)}/{len(got)}"
              f"（{100 * len(hit) / max(len(got), 1):.0f}%）"
              + ("　→ 可直接併" if len(hit) == len(got)
                 else "　→ 需要名稱對照，不重疊的："))
        for n in sorted(got - names)[:8]:
            print(f"      {n[:44]}")

        have = set(json.loads(pathlib.Path(PAIRING_FILE).read_text())["配對"].values()) \
            if pathlib.Path(PAIRING_FILE).exists() else set()
        have_names = {k.split("\t")[1] for k in have}
        new = hit - have_names
        print(f"\n② 能補的缺口：{len(new)} 個現流項目是我們沒有配對過的"
              f"（公版可編碼 {len(master)} 個、目前覆蓋 {len(have)} 個）")
        for n in sorted(new)[:10]:
            print(f"      {n[:44]}")

    out = sys.argv[2] if len(sys.argv) > 2 else None
    if out:
        with open(out, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(pairs[0]))
            w.writeheader()
            w.writerows(pairs)
        print(f"\n→ {out}")
    print("\n（只讀不改；要併進 勾稽配對表.json 是另一個決定，"
          "建議先跑歸零檢定確認配對正確）")


if __name__ == "__main__":
    main()

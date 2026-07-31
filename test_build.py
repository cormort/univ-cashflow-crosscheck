#!/usr/bin/env python3
"""自我檢查：不靠 Excel，直接在 Python 裡解析產出檔的 B/J/L 三欄公式，
算出它們會取到什麼值，再跟參考檔（大學116…）的快取值逐格比對。

用法：python3 test_build.py 大學115現金流量表勾稽檔-47所大學.xlsx 大學116現金流量表勾稽檔-47所大學.xlsx
"""
import re
import sys

import openpyxl

TOTAL_SHEET = "大學-自己調整"
NOTE_SHEET = "筆記"
SKEL_ROWS = 135
DETAIL_SHIFT = 1     # 明細表整體下移，與總表對齊
FROZEN_ROWS = 3

# 已查明原因、與參考檔本來就該不同的格（列號為產出檔座標）
KNOWN = {
    ("東華", "L128"): "參考檔裡只有這張表的查表範圍較短（漏掉合計列），眾數樣板已修正",
    ("臺科", "B118"): "參考檔的上年區塊把累積餘絀手動拆成賸餘與短絀，來源檔的上年調整後預算只有淨額",
    ("臺科", "B122"): "同上",
}

VLOOKUP_RE = re.compile(
    r"^=VLOOKUP\(A(\d+),'?" + re.escape(TOTAL_SHEET) + r"'?!\$A\$(\d+):\$[A-Z]+\$(\d+),\$K\$1,FALSE\)$")
OFFSET_RE = re.compile(
    r"^=OFFSET\('?" + re.escape(TOTAL_SHEET) + r"'?!\$A\$(\d+),0,\$K\$1-1\)$")


def lbl(v):
    """0 / None / 空字串 / 公式（不是科目名）都視為空白。"""
    if v in (None, 0, "") or (isinstance(v, str) and v.startswith("=")):
        return None
    return v


def main():
    out_path, ref_path = sys.argv[1], sys.argv[2]
    wb = openpyxl.load_workbook(out_path)                 # 公式
    wbd = openpyxl.load_workbook(out_path, data_only=True)  # 我們自己寫進去的字面值
    ref = openpyxl.load_workbook(ref_path, data_only=True)

    total = wbd[TOTAL_SHEET]
    names = [n for n in wb.sheetnames if n not in (NOTE_SHEET, TOTAL_SHEET)]
    ref_names = [n for n in ref.sheetnames if n not in (NOTE_SHEET, TOTAL_SHEET)]
    assert len(names) == len(ref_names) == 47, (len(names), len(ref_names))

    checked = mismatch = unparsed = 0
    bad, known = [], []
    for idx, (name, rname) in enumerate(zip(names, ref_names), start=1):
        ws, rws = wb[name], ref[rname]
        k1 = idx + 2                                      # K1 = L1 + 2
        for r in range(1, SKEL_ROWS + DETAIL_SHIFT + 1):
            r_ref = r - DETAIL_SHIFT if r > FROZEN_ROWS else r
            for c in (2, 10, 12):                         # B, J, L
                f = ws.cell(r, c).value
                if not isinstance(f, str) or not f.startswith("="):
                    continue
                m = VLOOKUP_RE.match(f)
                if m:
                    label = ws.cell(int(m.group(1)), 1).value
                    got = None
                    for br in range(int(m.group(2)), int(m.group(3)) + 1):
                        if total.cell(br, 1).value == label:
                            got = total.cell(br, k1).value
                            break
                else:
                    m = OFFSET_RE.match(f)
                    if not m:
                        if "VLOOKUP" in f or "OFFSET" in f:
                            unparsed += 1
                        continue
                    got = total.cell(int(m.group(1)), k1).value

                want = rws.cell(r_ref, c).value
                checked += 1
                if (got or 0) != (want or 0):
                    key = (name, {2: "B", 10: "J", 12: "L"}[c] + str(r))
                    if key in KNOWN:
                        known.append((key, got, want))
                        continue
                    mismatch += 1
                    if len(bad) < 25:
                        bad.append((name, {2:'B',10:'J',12:'L'}[c] + str(r), got, want, f[:60]))

    print(f"比對 {checked} 格（B/J/L）：不符 {mismatch} 格，"
          f"已知差異 {len(known)} 格，無法解析的查表公式 {unparsed} 個")
    for key, got, want in known:
        print(f"  [已知] {key[0]} {key[1]}: 產出={got!r} 參考={want!r} — {KNOWN[key]}")
    for name, cell, got, want, f in bad:
        print(f"  {name} {cell}: 產出={got!r} 參考={want!r}  {f}")
    assert unparsed == 0, "有查表公式沒被改寫成指向總表"
    assert mismatch == 0, "B/J/L 取值與參考檔不符"
    print("✓ B/J/L 三欄與參考檔完全一致")

    # 各校表與總表逐列對齊：科目/現流項目要同列，總表的跨表加總要指同一列
    wt = wb[TOTAL_SHEET]
    off = bad_align = 0
    wording = set()
    for name in names:
        ws = wb[name]
        for r in range(2, SKEL_ROWS + DETAIL_SHIFT + 1):   # 第 1 列是校名/表名，本就不同
            for c in (1, 8):                               # A 科目、H 現流項目
                a, b = lbl(ws.cell(r, c).value), lbl(wt.cell(r, c).value)
                if (a is None) != (b is None):             # 有沒有項目必須一致 = 列有對齊
                    bad_align += 1
                    if bad_align <= 10:
                        print(f"  未對齊 {name} {'AH'[c // 8]}{r}: {a!r} vs 總表 {b!r}")
                elif a != b and not str(b).startswith("="):
                    wording.add(f"{'AH'[c // 8]}{r}")
    if wording:
        print(f"  （列有對齊、只是用字略有出入：{'、'.join(sorted(wording))}）")
    sum3d = re.compile(r"^=SUM\([^!]+!\$?[A-Z]{1,2}\$?(\d+)\)$")
    for r in range(1, SKEL_ROWS + DETAIL_SHIFT + 1):
        for c in range(1, 20):
            m = sum3d.match(str(wt.cell(r, c).value))
            if m and int(m.group(1)) != r:
                off += 1
                if off <= 5:
                    print(f"  總表 {r},{c} 加總指到第 {m.group(1)} 列")
    assert bad_align == 0, "各校表與總表的科目列未對齊"
    assert off == 0, "總表的跨表加總沒有指到同一列"
    print("✓ 47 張明細表與總表逐列對齊，總表跨表加總全部指同一列")


if __name__ == "__main__":
    main()

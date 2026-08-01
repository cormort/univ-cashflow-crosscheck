#!/usr/bin/env python3
"""由「12.平衡」與「3.現流」兩份來源檔，產生大學現金流量表勾稽底稿。

作法：以既有勾稽檔為範本（保留全部公式與格式），
      1) 47 張明細表逐格取眾數，消掉各表之間的漂移
      2) 五個原始資料區塊只寫一份在總表底下
      3) 明細表的 VLOOKUP 範圍 / OFFSET 錨點改指總表，並依新資料列數重算
"""
import argparse
import collections
import difflib
import json
import pathlib
import re
import sys

import openpyxl
from openpyxl.utils import get_column_letter


class BuildError(Exception):
    """來源檔或範本不符預期，訊息直接給使用者看。"""

TOTAL_SHEET = "大學-自己調整"
NOTE_SHEET = "筆記"
SKEL_ROWS = 135          # 範本明細表的骨架列數
DETAIL_SHIFT = 1         # 明細表整體下移，讓列號與總表對齊
FROZEN_ROWS = 3          # 前三列（標題）不下移
SKEL_COLS = 18           # A~R
N_UNI = 47
TOTAL_COL = 2            # 資料區塊的「合計」欄
BLOCK_LAST_COL = TOTAL_COL + N_UNI   # 49 = AW

# 五個資料區塊的順序（範本與產出一致）
BLOCK_NAMES = ["本年資產", "本年負債淨值", "上年資產", "上年負債淨值", "現流"]


# --------------------------------------------------------------------------
# 來源檔讀取
# --------------------------------------------------------------------------
def pick_sheet(wb, tag=None, name_prefix=None):
    """依 A1 標記或表名前綴挑工作表。

    排除「差異」表；優先取名稱帶最新四位日期後綴者，同日期取名稱最短者
    （避開 `114調後2`、`細修2` 這種暫存版）。
    """
    cands = []
    for name in wb.sheetnames:
        if "差異" in name:
            continue
        if tag is not None and str(wb[name]["A1"].value).strip() != tag:
            continue
        if name_prefix is not None and not name.startswith(name_prefix):
            continue
        m = re.search(r"(\d{4})$", name)
        cands.append(((int(m.group(1)) if m else -1, -len(name)), name))
    if not cands:
        raise BuildError(f"找不到工作表（tag={tag!r} prefix={name_prefix!r}）")
    return wb[max(cands)[1]]


def read_matrix(ws, header_row=4):
    """回傳 (校名清單, [(科目, [合計, *47校])])，保留空白列以便切段。"""
    unis = [ws.cell(header_row, c).value for c in range(TOTAL_COL + 1, BLOCK_LAST_COL + 1)]
    if any(u is None for u in unis):
        raise BuildError(f"{ws.title}: 第 {header_row} 列的校名不足 {N_UNI} 校")
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        label = ws.cell(r, 1).value
        vals = [ws.cell(r, c).value for c in range(TOTAL_COL, BLOCK_LAST_COL + 1)]
        rows.append((label, vals))
    while rows and rows[-1][0] is None:
        rows.pop()
    return unis, rows


def split_sections(rows, sheet_title):
    """把平衡表切成 (資產段, 負債淨值段)，各自含最後的「合計」列。"""
    idx_total = [i for i, (lb, _) in enumerate(rows)
                 if isinstance(lb, str) and lb.split() == ["合", "計"]]
    idx_liab = [i for i, (lb, _) in enumerate(rows)
                if isinstance(lb, str) and lb.strip() == "負債"]
    if len(idx_total) < 2 or not idx_liab:
        raise BuildError(f"{sheet_title}: 找不到資產/負債的分段點（合計列 {len(idx_total)} 個）")
    return rows[: idx_total[0] + 1], rows[idx_liab[0]: idx_total[-1] + 1]


def parse_year(wb_bal, wb_cf):
    """從來源檔標題解析本年度（民國年）。"""
    found = []
    for wb, tag, prefix in ((wb_bal, None, "彙總修"), (wb_cf, None, "總修")):
        try:
            ws = pick_sheet(wb, name_prefix=prefix)
        except BuildError:
            continue
        for r in range(1, 8):
            for c in range(1, 8):
                v = ws.cell(r, c).value
                if isinstance(v, str):
                    m = re.search(r"中華民國(\d{2,3})年", v)
                    if m:
                        found.append((ws.title, v.strip(), int(m.group(1))))
                        break
            if found and found[-1][0] == ws.title:
                break
    if not found:
        raise BuildError("無法從來源檔標題解析年度，請用 --year 指定")
    years = {y for _, _, y in found}
    if len(years) > 1:
        print(f"! 兩份來源檔的年度不一致：{found}", file=sys.stderr)
    return found, min(years)


# --------------------------------------------------------------------------
# 範本：區塊掃描與眾數樣板
# --------------------------------------------------------------------------
Block = collections.namedtuple("Block", "header start end")


def scan_blocks(ws):
    """掃出總表在骨架之後的五個資料區塊（ws 需為 data_only 讀取）。"""
    headers = [r for r in range(SKEL_ROWS + 1, ws.max_row + 1)
               if isinstance(ws.cell(r, 1).value, str)
               and ws.cell(r, 1).value.strip() in ("科目", "項      目")]
    if len(headers) != len(BLOCK_NAMES):
        raise BuildError(f"{ws.title}: 掃到 {len(headers)} 個資料區塊，預期 {len(BLOCK_NAMES)} 個")
    blocks = {}
    for i, h in enumerate(headers):
        # 下一個區塊的表頭列前面是它的標示列（如「115(本年度)」），不屬於本區塊
        limit = headers[i + 1] - 1 if i + 1 < len(headers) else ws.max_row + 1
        end = h
        for r in range(h + 1, limit):
            if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip():
                end = r
        blocks[BLOCK_NAMES[i]] = Block(h, h + 1, end)
    return blocks


MIRROR_RE = re.compile(r"^='?" + re.escape(TOTAL_SHEET) + r"'?!\$?A\$?(\d+)$")


def scan_mirror(ws):
    """明細表骨架之後是總表資料區塊的鏡射公式；回傳 {明細列: 總表列}。"""
    mirror = {}
    for r in range(SKEL_ROWS + 1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str):
            m = MIRROR_RE.match(v.strip())
            if m:
                mirror[r] = int(m.group(1))
    if len(mirror) < 100:
        raise BuildError(f"{ws.title}: 只認出 {len(mirror)} 列鏡射公式，範本結構與預期不符")
    return mirror


def load_mirror(template_path, ws):
    """先讀鏡射公式；瘦身過的範本沒有那些列，改讀旁邊的 .mirror.json。"""
    try:
        return scan_mirror(ws)
    except BuildError:
        side = pathlib.Path(str(template_path) + ".mirror.json")
        if not side.exists():
            raise BuildError(f"範本沒有鏡射列，也找不到 {side.name}")
        return {int(k): v for k, v in json.loads(side.read_text()).items()}


def strip_template(src, dst):
    """把 14MB 範本瘦身成骨架範本：

    - 刪掉 47 張明細表的鏡射資料列（鏡射對照另存 JSON）
    - 清掉總表資料區塊裡的金額，只留科目名稱
      （產生時本來就只讀科目名稱來對列，金額一律重寫；
       範本因此不帶任何一年的預算數字，可以安心公開）
    """
    wb = openpyxl.load_workbook(src)
    names = [n for n in wb.sheetnames if n not in (NOTE_SHEET, TOTAL_SHEET)]
    mirror = scan_mirror(wb[names[0]])
    for n in names:
        ws = wb[n]
        if ws.max_row > SKEL_ROWS:
            ws.delete_rows(SKEL_ROWS + 1, ws.max_row - SKEL_ROWS + 1)
    wt = wb[TOTAL_SHEET]
    wiped = 0
    for r in range(SKEL_ROWS + 1, wt.max_row + 1):
        for c in range(TOTAL_COL, BLOCK_LAST_COL + 1):
            if wt.cell(r, c).value is not None:
                put(wt, r, c, None)
                wiped += 1

    # 47 張明細表一律換成眾數樣板：抹掉各校去年的人工填數與表名註記，
    # 範本因此不含任何一年的金額。（產生時本來就會取眾數，結果不變）
    tpl, _ = modal_template(wb, names)
    plugs = 0
    for i, n in enumerate(names, start=1):
        ws = wb[n]
        for (r, c), v in tpl.items():
            if ws.cell(r, c).value != v:
                if isinstance(ws.cell(r, c).value, (int, float)):
                    plugs += 1
                put(ws, r, c, v)
        put(ws, 1, 12, i)
        ws.title = n.split("-")[0]
    print(f"  已清掉範本裡 {wiped} 格舊年度金額、{plugs} 筆各校人工填數，表名註記也一併移除")
    wb.save(dst)
    pathlib.Path(str(dst) + ".mirror.json").write_text(
        json.dumps({str(k): v for k, v in mirror.items()}))
    print(f"✓ {dst}（{pathlib.Path(dst).stat().st_size / 1e6:.1f}MB）"
          f"＋ {pathlib.Path(dst).name}.mirror.json")


def modal_template(wb, detail_names):
    """47 張明細表逐格取眾數，回傳 (樣板, 分歧清單)。"""
    sheets = [wb[n] for n in detail_names]
    tpl, drift = {}, []
    for r in range(1, SKEL_ROWS + 1):
        for c in range(1, SKEL_COLS + 1):
            vals = [s.cell(r, c).value for s in sheets]
            cnt = collections.Counter(repr(v) for v in vals)
            best, n = cnt.most_common(1)[0]
            tpl[(r, c)] = vals[[repr(v) for v in vals].index(best)]
            if len(cnt) > 1:
                drift.append((r, c, cnt.most_common(3)))
    return tpl, drift


# --------------------------------------------------------------------------
# 公式改寫
# --------------------------------------------------------------------------
RANGE_RE = re.compile(r"\$?([A-Z]{1,2})\$?(\d+):\$?([A-Z]{1,2})\$?(\d+)")
CELL_RE = re.compile(r"(?<![\w$!:])(\$?)([A-Z]{1,2})(\$?)(\d+)(?![\w(:])")
ANYREF_RE = re.compile(r"(\$?)([A-Z]{1,2})(\$?)(\d+)")


def shift_local_rows(formula, delta):
    """把公式裡指向骨架（列 4~SKEL_ROWS）的列號整體位移；資料區塊列號不動。"""
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula

    def bump(m):
        row = int(m.group(4))
        if not (FROZEN_ROWS < row <= SKEL_ROWS):
            return m.group(0)
        return f"{m.group(1)}{m.group(2)}{m.group(3)}{row + delta}"

    return ANYREF_RE.sub(bump, formula)


def build_rowmap(old_blocks, new_blocks, old_ws, new_rows_by_block):
    """總表老列號 → 新列號，與 老列號 → 區塊名。

    用序列對齊（不是查表）比對新舊科目清單，因為現流表有重複的項目名稱
    （「收取利息」「機械及設備」都出現多次），純查表會對到第一筆。
    """
    rowmap, blockmap = {}, {}
    for name in BLOCK_NAMES:
        ob, nb = old_blocks[name], new_blocks[name]
        rowmap[ob.header] = nb.header
        for r in range(ob.header, ob.end + 1):
            blockmap[r] = name
        old_labels = [old_ws.cell(r, 1).value for r in range(ob.start, ob.end + 1)]
        new_labels = [lb for lb, _ in new_rows_by_block[name]]
        for tag, i1, i2, j1, _ in difflib.SequenceMatcher(
                None, old_labels, new_labels, autojunk=False).get_opcodes():
            if tag == "equal":
                for k in range(i2 - i1):
                    rowmap[ob.start + i1 + k] = nb.start + j1 + k
    return rowmap, blockmap


def rewrite(formula, rowmap, blockmap, new_blocks, prefix, skel_rows=SKEL_ROWS):
    """把公式裡指向資料區塊的參照改成新位置（必要時加上總表前綴）。"""
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula, True, set()

    ok = True
    unresolved = set()

    def fix_range(m):
        nonlocal ok
        r1 = int(m.group(2))
        if r1 <= skel_rows:
            return m.group(0)
        name = blockmap.get(r1)
        if name is None:
            ok = False
            unresolved.add(r1)
            return m.group(0)
        nb = new_blocks[name]
        return (f"{prefix}$A${nb.start}:${get_column_letter(BLOCK_LAST_COL)}${nb.end}")

    out = RANGE_RE.sub(fix_range, formula)

    def fix_cell(m):
        nonlocal ok
        row = int(m.group(4))
        if row <= skel_rows:
            return m.group(0)
        new = rowmap.get(row)
        if new is None:
            ok = False
            unresolved.add(row)
            return m.group(0)
        return f"{prefix}${m.group(2)}${new}"

    out = CELL_RE.sub(fix_cell, out)
    return out, ok, unresolved


# --------------------------------------------------------------------------
# 寫入
# --------------------------------------------------------------------------
def put(ws, r, c, value):
    """寫入儲存格；合併儲存格的非左上角略過（openpyxl 不允許寫入）。"""
    cell = ws.cell(r, c)
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        return None
    cell.value = value
    return cell


def write_block(ws, marker_row, header_row, marker, head_label, unis, rows, num_fmt):
    put(ws, marker_row, 1, marker)
    put(ws, header_row, 1, head_label)
    put(ws, header_row, TOTAL_COL, "合    計")
    for i, u in enumerate(unis):
        put(ws, header_row, TOTAL_COL + 1 + i, u)
    for i, (label, vals) in enumerate(rows):
        r = header_row + 1 + i
        put(ws, r, 1, label)
        for j, v in enumerate(vals):
            cell = put(ws, r, TOTAL_COL + j, v)
            if cell is not None and isinstance(v, (int, float)):
                cell.number_format = num_fmt
    return header_row + len(rows)


DEFAULT_TEMPLATE = "template.xlsx"


def build(balance, cashflow, template=DEFAULT_TEMPLATE, output=None, year=None,
          bal_sheet=None, prev_sheet=None, cf_sheet=None, outdir="."):
    """產生勾稽底稿，回傳 (輸出路徑, 報告文字)。"""
    lines = []
    def report(*a):
        lines.append(" ".join(str(x) for x in a))

    args = argparse.Namespace(balance=balance, cashflow=cashflow, template=template,
                              output=output, year=year, bal_sheet=bal_sheet,
                              prev_sheet=prev_sheet, cf_sheet=cf_sheet)

    # ---- 來源 ----
    wb_bal = openpyxl.load_workbook(args.balance, data_only=True)
    wb_cf = openpyxl.load_workbook(args.cashflow, data_only=True)

    evidence, parsed = parse_year(wb_bal, wb_cf)
    year = args.year or parsed
    report("年度解析：")
    for title, text, y in evidence:
        report(f"  {title}: {text} → {y}")
    report(f"  ⇒ 本年度 {year}、上年度 {year - 1}\n")

    ws_cur = wb_bal[args.bal_sheet] if args.bal_sheet else pick_sheet(wb_bal, tag="1.本年預算數")
    ws_prev = wb_bal[args.prev_sheet] if args.prev_sheet else pick_sheet(wb_bal, tag="5.上年調整後預算數")
    ws_cfd = wb_cf[args.cf_sheet] if args.cf_sheet else pick_sheet(wb_cf, name_prefix="細修")
    report(f"來源工作表：本年={ws_cur.title}　上年={ws_prev.title}　現流={ws_cfd.title}\n")

    unis_cur, rows_cur = read_matrix(ws_cur)
    unis_prev, rows_prev = read_matrix(ws_prev)
    unis_cf, rows_cf = read_matrix(ws_cfd)
    for tag, u in (("上年平衡表", unis_prev), ("現流表", unis_cf)):
        if u != unis_cur:
            raise BuildError(f"{tag}的校名順序與本年平衡表不一致")

    # 合計欄勾稽
    bad = []
    for tag, rows in (("本年", rows_cur), ("上年", rows_prev), ("現流", rows_cf)):
        for label, vals in rows:
            if not isinstance(vals[0], (int, float)):
                continue
            s = sum(v for v in vals[1:] if isinstance(v, (int, float)))
            if abs(s - vals[0]) > 0.5:
                bad.append((tag, label, vals[0], s))
    if bad:
        report("! 合計 ≠ 47 校橫加：")
        for t, lb, a, b in bad[:20]:
            report(f"   {t} {str(lb).strip()}: 合計 {a} vs 橫加 {b}")
        raise BuildError("來源資料未通過勾稽，中止")

    cur_a, cur_l = split_sections(rows_cur, ws_cur.title)
    prev_a, prev_l = split_sections(rows_prev, ws_prev.title)
    rows_cf = [r for r in rows_cf if r[0] is not None]
    new_data = {
        "本年資產": cur_a, "本年負債淨值": cur_l,
        "上年資產": prev_a, "上年負債淨值": prev_l,
        "現流": rows_cf,
    }
    markers = {
        "本年資產": f"{year}(本年度)", "本年負債淨值": f"{year}(本年度)",
        "上年資產": f"{year - 1}(上年度)", "上年負債淨值": f"{year - 1}(上年度)",
        "現流": "現金流量表",
    }

    # ---- 範本 ----
    wb = openpyxl.load_workbook(args.template)
    wbv = openpyxl.load_workbook(args.template, data_only=True)
    detail_names = [n for n in wb.sheetnames if n not in (NOTE_SHEET, TOTAL_SHEET)]
    if len(detail_names) != N_UNI:
        raise BuildError(f"範本有 {len(detail_names)} 張明細表，預期 {N_UNI} 張")

    old_total_blocks = scan_blocks(wbv[TOTAL_SHEET])
    mirror = load_mirror(args.template, wb[detail_names[0]])
    num_fmt = {name: wb[TOTAL_SHEET].cell(b.start, TOTAL_COL).number_format
               for name, b in old_total_blocks.items()}

    tpl, drift = modal_template(wb, detail_names)

    # ---- 新區塊配置（全部寫在總表）----
    new_blocks, cursor = {}, SKEL_ROWS + 2
    for name in BLOCK_NAMES:
        header = cursor + 1                       # cursor = 標示列
        new_blocks[name] = Block(header, header + 1, header + len(new_data[name]))
        cursor = new_blocks[name].end + 3
    rowmap_total, blockmap_total = build_rowmap(
        old_total_blocks, new_blocks, wbv[TOTAL_SHEET], new_data)
    # 明細表的區塊是總表的鏡射，經 mirror 轉一手即可
    rowmap_detail = {r: rowmap_total[t] for r, t in mirror.items() if t in rowmap_total}
    blockmap_detail = {r: blockmap_total[t] for r, t in mirror.items() if t in blockmap_total}

    # ---- 明細表：套樣板 + 改寫公式 + 清掉原有資料區塊 ----
    failed = []
    renamed = []
    for idx, name in enumerate(detail_names, start=1):
        ws = wb[name]
        # openpyxl 的 insert_rows 不會跟著搬合併儲存格，得自己搬
        merges = [str(rng) for rng in ws.merged_cells.ranges]
        for rng in merges:
            ws.unmerge_cells(rng)
        ws.insert_rows(FROZEN_ROWS + 1, DETAIL_SHIFT)     # 空出與總表對齊用的列
        for rng in merges:
            mc = openpyxl.worksheet.cell_range.CellRange(rng)
            if mc.min_row > FROZEN_ROWS:
                mc.shift(row_shift=DETAIL_SHIFT)
            ws.merge_cells(str(mc))
        for (r, c), v in sorted(tpl.items(), reverse=True):
            v = shift_local_rows(v, DETAIL_SHIFT)
            new, ok, unresolved = rewrite(v, rowmap_detail, blockmap_detail, new_blocks,
                                          f"'{TOTAL_SHEET}'!", SKEL_ROWS + DETAIL_SHIFT)
            if not ok:                       # 對不到新區塊位置，只能清空
                skel = SKEL_ROWS + DETAIL_SHIFT
                stale = (not (unresolved & set(blockmap_detail))
                         and min(unresolved) > skel)
                failed.append((name, r + DETAIL_SHIFT if r > FROZEN_ROWS else r, c, v, stale))
                new = None
            put(ws, r + (DETAIL_SHIFT if r > FROZEN_ROWS else 0), c, new)
        for c in range(1, SKEL_COLS + 4):                 # 借總表的欄位說明填新空列
            put(ws, FROZEN_ROWS + 1, c, wb[TOTAL_SHEET].cell(FROZEN_ROWS + 1, c).value)
        put(ws, 1, 12, idx)                               # L1 = 校序號
        put(ws, 3, 1, f"中華民國{year}年度")
        put(ws, 5, 2, f"{year - 1}.12.31")
        put(ws, 5, 7, f"{year}.12.31")
        put(ws, 7, 12, f"{year}年12月")
        last = SKEL_ROWS + DETAIL_SHIFT
        ws.delete_rows(last + 1, ws.max_row - last + 1)
        base = name.split("-")[0]
        if base != name:
            renamed.append((name, base))
            ws.title = base

    # ---- 總表：改寫公式 + 寫入資料區塊 ----
    wt = wb[TOTAL_SHEET]
    for r in range(1, SKEL_ROWS + 1):
        for c in range(1, SKEL_COLS + 4):
            v = wt.cell(r, c).value
            new, ok, unresolved = rewrite(v, rowmap_total, blockmap_total, new_blocks, "")
            if not ok:
                stale = (not (unresolved & set(blockmap_total))
                         and min(unresolved) > SKEL_ROWS)
                failed.append((TOTAL_SHEET, r, c, v, stale))
                new = None
            if new != v:
                put(wt, r, c, new)
    put(wt, 3, 1, f"中華民國{year}年度")
    put(wt, 5, 2, f"{year - 1}.12.31")
    put(wt, 5, 7, f"{year}.12.31")
    first, last = wb.sheetnames[2], wb.sheetnames[-1]
    realigned = set()
    for r in range(1, SKEL_ROWS + 1):
        for c in range(1, SKEL_COLS + 1):
            v = wt.cell(r, c).value
            if isinstance(v, str) and ":" in v and "!" in v and v.startswith("=SUM("):
                v = re.sub(r"[^(]+!", f"{first}:{last}!", v)
                # 一律指向同一列，才能整欄往下拉；範本原有幾格是跳列的
                fixed = re.sub(r"!(\$?[A-Z]{1,2}\$?)\d+", lambda m: f"!{m.group(1)}{r}", v)
                if fixed != re.sub(r"!(\$?[A-Z]{1,2}\$?)(\d+)",
                                   lambda m: f"!{m.group(1)}{int(m.group(2)) + DETAIL_SHIFT}", v):
                    realigned.add(r)
                put(wt, r, c, fixed)

    wt.delete_rows(SKEL_ROWS + 1, wt.max_row - SKEL_ROWS + 1)
    for name in BLOCK_NAMES:
        b = new_blocks[name]
        head = "項      目" if name == "現流" else "科目"
        write_block(wt, b.header - 1, b.header, markers[name], head,
                    unis_cur, new_data[name], num_fmt[name])

    # ---- 報告 ----
    report(f"資料區塊（總表）：")
    for name in BLOCK_NAMES:
        b = new_blocks[name]
        report(f"  {name:<6} 表頭 {b.header:>4}　資料 {b.start}–{b.end}（{len(new_data[name])} 列）")

    skel_labels = {tpl[(r, 1)] for r in range(7, SKEL_ROWS + 1)} | \
                  {tpl[(r, 8)] for r in range(7, SKEL_ROWS + 1)}
    src_labels = {lb for name in BLOCK_NAMES for lb, _ in new_data[name]}
    only_src = sorted(l for l in src_labels - skel_labels if isinstance(l, str) and l.strip())
    only_skel = sorted(l for l in skel_labels - src_labels if isinstance(l, str) and l.strip())
    report(f"\n科目對照：來源有而骨架無 {len(only_src)} 項；骨架有而來源無 {len(only_skel)} 項")
    for l in only_skel:
        report(f"  骨架缺料（會 #N/A）：{l.strip()}")

    if realigned:
        report(f"\n總表跨表加總已改為一律指同一列（範本原本跳列的第 "
              f"{'、'.join(str(r) for r in sorted(realigned))} 列已修正）")

    if renamed:
        report(f"\n表名已去除去年的註記後綴：")
        for old, new in renamed:
            report(f"  {old}  →  {new}")
    if failed:
        broken = [f for f in failed if "#REF!" in str(f[3])]
        stale = [f for f in failed if "#REF!" not in str(f[3]) and f[4]]
        lost = [f for f in failed if "#REF!" not in str(f[3]) and not f[4]]

        def cells(items):
            return "、".join(sorted({f"{get_column_letter(c)}{r}" for _, r, c, *_ in items},
                                   key=lambda s: (s[0], int(s[1:]))))

        report(f"\n已清空 {len(failed)} 個公式（座標為產出檔）：")
        if broken:
            report(f"  範本裡本來就是 #REF! 斷鏈：{cells(broken)}（{len(broken)} 格）")
        if stale:
            report(f"  指向明細表已移除的舊鏡射區、新版沒有對應位置：{cells(stale)}（{len(stale)} 格）")
        if lost:
            report(f"  ! 對不到位置且不在舊鏡射區，可能是來源資料變動或邊界判斷有誤："
                   f"{cells(lost)}（{len(lost)} 格）")
            for name, r, c, v, _ in lost[:5]:
                report(f"      {name} {get_column_letter(c)}{r}: {v}")

    CHECK_COLS = range(14, SKEL_COLS + 1)          # N~R：檢查欄，各表擺放位置本就不一
    ties = [(r, c, mc) for r, c, mc in drift
            if c not in CHECK_COLS and c != 12 and mc[0][1] < N_UNI * 0.75]
    n_check = sum(1 for r, c, _ in drift if c in CHECK_COLS)
    if ties:
        report(f"\n樣板取眾數時的近平手（{len(ties)} 處，建議人工確認）：")
        for r, c, mc in ties:
            opts = "　/　".join(f"{str(v)[:32]}×{n}" for v, n in mc)
            report(f"  {get_column_letter(c)}{r}: {opts}")
    report(f"（另有 N~R 檢查欄 {n_check} 格各表擺法不一，已統一為眾數）")

    out = args.output or str(
        pathlib.Path(outdir) / f"大學{year}現金流量表勾稽檔-{N_UNI}所大學.xlsx")
    wb.save(out)
    report(f"\n✓ 已寫出 {pathlib.Path(out).name}")
    return out, "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="產生大學現金流量表勾稽底稿")
    ap.add_argument("balance", help="12.平衡xxxx.xlsx")
    ap.add_argument("cashflow", help="3.現流xxxx.xlsx")
    ap.add_argument("-t", "--template", default=DEFAULT_TEMPLATE)
    ap.add_argument("-o", "--output")
    ap.add_argument("--year", type=int, help="覆寫解析到的年度")
    ap.add_argument("--bal-sheet"), ap.add_argument("--prev-sheet"), ap.add_argument("--cf-sheet")
    ap.add_argument("--strip-template", metavar="OUT",
                    help="把範本瘦身成只留骨架的 OUT（另存 OUT.mirror.json），不做其他事")
    args = ap.parse_args()
    if args.strip_template:
        strip_template(args.template, args.strip_template)
        return
    try:
        _, text = build(args.balance, args.cashflow, args.template, args.output,
                        args.year, args.bal_sheet, args.prev_sheet, args.cf_sheet)
    except BuildError as e:
        raise SystemExit(f"✗ {e}")
    print(text)


if __name__ == "__main__":
    main()

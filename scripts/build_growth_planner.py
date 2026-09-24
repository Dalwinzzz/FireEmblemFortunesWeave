"""生成《固定成长养成规划器》Excel。

数据源：data/characters.csv（角色成长率）、data/classes.csv（职业补正）
输出：tools/固定成长养成规划器.xlsx

用法：python scripts/build_growth_planner.py
"""
import csv
import os

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "tools", "固定成长养成规划器.xlsx")

STATS = ["HP", "力量", "魔力", "速度", "技巧", "防守", "魔防", "幸运", "魅力"]
NS = len(STATS)
DATA_ROWS = 150          # 数据表预留行数（第3行起），方便以后追加角色/职业
DATA_LAST = 2 + DATA_ROWS
COMBO_ROWS = 6000        # 推荐路线搜索的组合上限
TL_ROWS = 100            # 逐级成长表行数

SH_MAIN, SH_CHAR, SH_CLASS = "养成规划", "角色成长率", "职业"
CHAR_Q, CLASS_Q = f"'{SH_CHAR}'", f"'{SH_CLASS}'"

# ---------- 样式 ----------
FONT = "Microsoft YaHei"
f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_title = Font(name=FONT, size=16, bold=True, color="1F3864")
f_sec = Font(name=FONT, size=12, bold=True, color="FFFFFF")
f_hdr = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_input = Font(name=FONT, size=10, color="0000FF", bold=True)
f_note = Font(name=FONT, size=9, color="595959")
f_warn = Font(name=FONT, size=10, bold=True, color="C00000")
fill_sec = PatternFill("solid", fgColor="1F3864")
fill_hdr = PatternFill("solid", fgColor="2F5597")
fill_input = PatternFill("solid", fgColor="FFF2CC")
fill_info = PatternFill("solid", fgColor="DEEBF7")
fill_grey = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
left_nowrap = Alignment(horizontal="left", vertical="center")


def num(v):
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except (TypeError, ValueError):
        return v


def load_csv(name):
    with open(os.path.join(ROOT, "data", name), encoding="utf-8-sig") as fp:
        rows = list(csv.reader(fp))
    return rows[0], rows[1:]


def style_range(ws, ref, font=None, fill=None, align=None, brd=True):
    for row in ws[ref]:
        for c in row:
            if font:
                c.font = font
            if fill:
                c.fill = fill
            if align:
                c.alignment = align
            if brd:
                c.border = border


def section(ws, row, text, last_col="M"):
    ws.merge_cells(f"A{row}:{last_col}{row}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = f_sec
    c.fill = fill_sec
    c.alignment = left_nowrap
    ws.row_dimensions[row].height = 22


def header(ws, row, values, start_col=1):
    for i, v in enumerate(values):
        c = ws.cell(row, start_col + i, v)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = center
        c.border = border


# ======================================================================
# 数据表
# ======================================================================
def build_data_sheet(ws, title, hdr, rows, front, growth_cols, extra_formula=None, widths=None):
    """front：放在成长率之前的列名；growth_cols：9 个成长率列名；其余列原样追加在后面。"""
    rest = [h for h in hdr if h not in front and h not in growth_cols]
    cols = front + growth_cols + (["成长合计"] if extra_formula else []) + [h for h in rest if h != "成长合计"]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="1F3864")
    header(ws, 2, cols)
    ws.row_dimensions[2].height = 32
    idx = {h: i for i, h in enumerate(hdr)}
    for r, row in enumerate(rows, start=3):
        for c, h in enumerate(cols, start=1):
            if h == "成长合计" and extra_formula:
                v = extra_formula(r)
            else:
                v = row[idx[h]] if h in idx else ""
                v = num(v) if h in growth_cols or h in ("序号", "最低转职等级", "移动力") else v
                if v == "":
                    v = None
            cell = ws.cell(r, c, v)
            cell.font = f_base
            cell.border = border
            cell.alignment = center if (h in growth_cols or h in front or h in ("成长合计", "移动力")) else left
    g0 = len(front) + 1
    for r in range(3, 3 + len(rows)):
        for c in range(g0, g0 + NS):
            ws.cell(r, c).fill = fill_info
    ws.freeze_panes = ws.cell(3, len(front) + 1)
    ws.auto_filter.ref = f"A2:{CL(len(cols))}{2 + len(rows)}"
    for i, h in enumerate(cols, start=1):
        w = (widths or {}).get(h)
        if w is None:
            w = 7 if h in growth_cols else 14
        ws.column_dimensions[CL(i)].width = w
    return cols


def build_char_sheet(wb):
    hdr, rows = load_csv("characters.csv")
    ws = wb.create_sheet(SH_CHAR)
    front = ["序号", "阵营", "角色", "日文名", "初始兵种"]
    widths = {"序号": 6, "阵营": 14, "角色": 12, "日文名": 14, "初始兵种": 10, "成长合计": 9,
              "职业偏向（图鉴·战斗特征）": 36, "个人技能": 36, "说明": 60, "血印1": 30, "血印2": 30}
    cols = build_data_sheet(ws, "角色成长率（个人成长率 %，可直接在此订正；角色名需与下拉框一致）", hdr, rows, front, STATS,
                            extra_formula=lambda r: f'=IF(COUNT(F{r}:N{r})=0,"—",SUM(F{r}:N{r}))', widths=widths)
    assert cols[2] == "角色" and cols[5:14] == STATS
    return len(rows)


def build_class_sheet(wb):
    hdr, rows = load_csv("classes.csv")
    ws = wb.create_sheet(SH_CLASS)
    front = ["序号", "阶级", "兵种名", "最低转职等级", "纳入推荐"]
    widths = {"序号": 6, "阶级": 9, "兵种名": 13, "最低转职等级": 9, "纳入推荐": 8, "转职条件": 22,
              "使用武器": 26, "中文说明": 50, "解放条件": 30, "代表角色": 24}
    cols = build_data_sheet(ws, "职业（职业成长率补正 %，可直接在此订正/追加）", hdr, rows, front, STATS, widths=widths)
    assert cols[2] == "兵种名" and cols[5:14] == STATS
    for r in range(3, 3 + len(rows)):
        for c in (4, 5):
            ws.cell(r, c).fill = fill_input
            ws.cell(r, c).font = f_input
    dv = DataValidation(type="list", formula1='"是,否"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"E3:E{DATA_LAST}")
    note_r = 3 + len(rows) + 1
    ws.cell(note_r, 1, "说明：").font = f_bold
    notes = [
        "「最低转职等级」取自转职条件里的推荐等级；没写推荐等级的按阶级默认（基本1/初级5/中级20/上级35/最上级45/神将45，神将职为推测值）。自动推荐只会在达到该等级后才转入此职业。",
        "「纳入推荐」=否 的职业不会出现在自动推荐里（手动路线仍可选）。舞者、飞天女神将默认设为否（疑似蕾达专属）；没解锁或不想用的职业也可以改成否。",
        "要追加职业，在表格末尾接着填一行即可（第152行以内），兵种名会自动出现在养成规划的下拉框里。",
    ]
    for i, t in enumerate(notes, start=1):
        ws.merge_cells(start_row=note_r + i, start_column=1, end_row=note_r + i, end_column=20)
        c = ws.cell(note_r + i, 1, t)
        c.font = f_note
        c.alignment = left_nowrap
    return len(rows)


# ======================================================================
# 主页面
# ======================================================================
S0 = 4  # 属性列起点 D
SC = [CL(S0 + j) for j in range(NS)]          # 主页面属性列 D..L
HC0 = 19                                        # 辅助区属性列起点 S
HC = [CL(HC0 + j) for j in range(NS)]         # S..AA
DC = [CL(6 + j) for j in range(NS)]           # 数据表属性列 F..N

CHAR_NAMES = f"{CHAR_Q}!$C$3:$C${DATA_LAST}"
CLASS_NAMES = f"{CLASS_Q}!$C$3:$C${DATA_LAST}"


def class_col(letter):
    return f"{CLASS_Q}!{letter}$3:{letter}${DATA_LAST}"


def build_main(wb):
    ws = wb.active
    ws.title = SH_MAIN
    ws.sheet_view.showGridLines = False
    for row in ws.iter_rows(min_row=1, max_row=160, max_col=13):
        for c in row:
            c.font = f_base

    widths = {"A": 12, "B": 11, "C": 13, "M": 62}
    for col in "ABCDEFGHIJKLM":
        ws.column_dimensions[col].width = widths.get(col, 8.5)

    # ---- 标题 ----
    ws.merge_cells("A1:M1")
    ws["A1"] = "火焰纹章 万缕千丝 ｜ 固定成长 · 养成路线规划器"
    ws["A1"].font = f_title
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:M2")
    ws["A2"] = ("用法：黄底蓝字是输入格。① 选角色 → ② 填起点/目标的等级、职业、属性（目标属性不填=不作要求）→ "
                "③ 看推荐路线，在「显示路线」里切换方案或改用手动路线 → ④ 下方逐级表显示每一级的职业和属性。"
                "规则：每升1级，各属性经验槽 +（角色成长率+职业补正，最低0），满100则该属性+1。")
    ws["A2"].font = f_note
    ws["A2"].alignment = left
    ws.row_dimensions[2].height = 42

    # ---- ① 角色 ----
    ws["A4"] = "① 角色"
    ws["A4"].font = f_bold
    ws.merge_cells("B4:C4")
    ws["B4"] = "蕾达"
    style_range(ws, "B4:C4", font=f_input, fill=fill_input, align=center)
    char_match = f"MATCH($B$4,{CHAR_NAMES},0)"
    ws["E4"] = "阵营"
    ws.merge_cells("F4:G4")
    ws["F4"] = f'=IFERROR(INDEX({CHAR_Q}!$B$3:$B${DATA_LAST},{char_match}),"")'
    ws["H4"] = "初始兵种"
    ws.merge_cells("I4:J4")
    ws["I4"] = f'=IFERROR(INDEX({CHAR_Q}!$E$3:$E${DATA_LAST},{char_match}),"")'
    ws.merge_cells("K4:M4")
    ws["K4"] = (f'=IF($B$4="","请选择角色",IF(ISNA({char_match}),"⚠ 角色表里找不到该角色",'
                f'IF(COUNT(INDEX({CHAR_Q}!$F$3:$N${DATA_LAST},{char_match},0))=0,"⚠ 该角色成长率数据缺失（按0计算）","✓ 成长率已读取")))')
    for ref in ("E4", "H4"):
        ws[ref].font = f_bold
        ws[ref].alignment = center
    style_range(ws, "F4:G4", fill=fill_grey, align=center)
    style_range(ws, "I4:J4", fill=fill_grey, align=center)
    ws["K4"].font = f_bold
    ws["K4"].alignment = left_nowrap

    # ---- ② 起点 / 目标 ----
    header(ws, 6, ["项目", "等级", "职业"] + STATS + ["说明"])
    labels = {7: "角色成长率%", 8: "起点职业补正%", 9: "起点", 10: "目标", 11: "需提升", 12: "每级需成长%"}
    notes = {7: "来自「角色成长率」表", 8: "来自「职业」表，仅供参考",
             9: "填当前等级、职业和面板属性", 10: "目标属性留空=不作要求；目标职业留空=45级后沿用上一职业",
             11: "目标 − 起点", 12: "达成目标平均每级需要的（角色+职业）成长，可对照职业补正"}
    for r, t in labels.items():
        ws.cell(r, 1, t).font = f_bold
        ws.cell(r, 1).alignment = center
        ws.cell(r, 1).border = border
        ws.cell(r, 13, notes[r]).font = f_note
        ws.cell(r, 13).alignment = left_nowrap
        ws.cell(r, 13).border = border
    for j in range(NS):
        s, d = SC[j], DC[j]
        ws[f"{s}7"] = f"=IFERROR(INDEX({CHAR_Q}!{d}$3:{d}${DATA_LAST},{char_match})+0,0)"
        ws[f"{s}8"] = f"=IFERROR(INDEX({class_col(d)},MATCH($C$9,{CLASS_NAMES},0))+0,0)"
        ws[f"{s}11"] = f'=IF(OR({s}10="",{s}9=""),"",{s}10-{s}9)'
        ws[f"{s}12"] = (f'=IF(OR({s}10="",$B$10="",$B$9="",$B$10<=$B$9),"",'
                        f'ROUND(MAX(0,({s}10-{s}9)*100-$B$14)/($B$10-$B$9),1))')
    for r in (7, 8, 11, 12):
        style_range(ws, f"B{r}:L{r}", fill=fill_grey, align=center)
    # 示例输入（蕾达，数值为示例）
    ws["B9"], ws["C9"] = 10, "猎兵"
    ws["B10"], ws["C10"] = 40, "狙击手"
    start = [26, 10, 6, 14, 11, 7, 6, 5, 9]
    target = [40, 22, None, 30, 27, None, None, None, None]
    for j in range(NS):
        ws[f"{SC[j]}9"] = start[j]
        ws[f"{SC[j]}10"] = target[j]
    style_range(ws, "B9:L10", font=f_input, fill=fill_input, align=center)
    ws["M9"] = "填当前等级、职业和面板属性（示例数据，请替换）"

    # ---- 参数 ----
    ws["A14"] = "初始经验槽%"
    ws["A14"].font = f_bold
    ws["B14"] = 0
    ws.merge_cells("C14:D14")
    ws["C14"] = "推荐路线转职节点Lv"
    ws["C14"].font = f_bold
    ws["C14"].alignment = center
    for col, v in zip("EFGH", (5, 20, 35, 45)):
        ws[f"{col}14"] = v
    ws.merge_cells("I14:J14")
    ws["I14"] = "显示路线"
    ws["I14"].font = f_bold
    ws["I14"].alignment = center
    ws.merge_cells("K14:L14")
    ws["K14"] = "推荐1"
    style_range(ws, "B14:B14", font=f_input, fill=fill_input, align=center)
    style_range(ws, "E14:H14", font=f_input, fill=fill_input, align=center)
    style_range(ws, "K14:L14", font=f_input, fill=fill_input, align=center)
    ws["M14"] = ("经验槽默认0=最保守（实际通常≥0，结果只会更好）。"
                 "节点：推荐路线只在这些等级换职业（对应初级/中级/上级/最上级的推荐等级）。")
    ws["M14"].font = f_note
    ws["M14"].alignment = left
    ws.row_dimensions[14].height = 30

    # ---- 辅助区（P 列起，隐藏） ----
    build_helpers(ws)

    # ---- 警告 ----
    ws.merge_cells("A15:M15")
    target_min = f"INDEX({class_col('D')},MATCH($C$10,{CLASS_NAMES},0))"
    entry_lv = "IF($T$9>0,$S$9,IF($U$5>0,INDEX($S$5:$S$9,$U$5+1),$B$9))"
    manual_bad = "OR(" + ",".join(f"$BV${r+1}<$BV${r}" for r in range(43, 48)) + ")"
    ws["A15"] = (
        '=IF(OR($B$9="",$B$10=""),"⚠ 请填写起点等级和目标等级；",IF($B$10<=$B$9,"⚠ 目标等级必须大于起点等级；",""))'
        f'&IF($B$10-$B$9>{TL_ROWS - 1},"⚠ 等级跨度超过{TL_ROWS - 1}，逐级表只显示前{TL_ROWS}级；","")'
        f'&IF(ISNA(MATCH($C$9,{CLASS_NAMES},0)),"⚠ 起点职业不在职业表中（补正按0计）；","")'
        f'&IF(AND($C$10<>"",ISNA(MATCH($C$10,{CLASS_NAMES},0))),"⚠ 目标职业不在职业表中；","")'
        f'&IFERROR(IF(AND($C$10<>"",{entry_lv}<{target_min}),"⚠ 目标职业最低转职等级为Lv"&{target_min}&"，但按目标等级只能在Lv"&{entry_lv}&"转入；",""),"")'
        f'&IF($U$9>{COMBO_ROWS},"⚠ 候选组合"&$U$9&"个，超过{COMBO_ROWS}，只搜索了前{COMBO_ROWS}个（可在职业表把用不到的职业设为「否」）；","")'
        '&IF(NOT($W$5),"⚠ 某个阶段没有可选职业，请检查职业表「纳入推荐」；","")'
        '&IF(AND($BV$26>0,NOT($BZ$19)),"⚠ 没有可行的推荐路线；","")'
        f'&IF(AND($BV$26=0,{manual_bad}),"⚠ 手动路线的转职等级需从上到下递增、中间不要空行；","")'
    )
    ws["A15"].font = f_warn
    ws["A15"].alignment = left_nowrap

    # ---- ③ 推荐路线 ----
    section(ws, 17, "③ 推荐路线 Top5（按：达成项数 ＞ 缺口合计最小 ＞ 最终属性总和最大 排序）")
    header(ws, 18, ["方案", "达成项", "缺口合计"] + STATS + ["路线（LvX 职业 = 从该等级起在此职业升级）"])
    for i in range(5):
        r = 19 + i
        ws.cell(r, 1, f"推荐{i + 1}")
        ws[f"B{r}"] = f'=IF($BZ{r},INDEX($BQ$13:$BQ${12 + COMBO_ROWS},$BV{r})&"/"&$W$6,"—")'
        ws[f"C{r}"] = f'=IF($BZ{r},INDEX($BR$13:$BR${12 + COMBO_ROWS},$BV{r}),"")'
        for j in range(NS):
            ws[f"{SC[j]}{r}"] = (f'=IF($BZ{r},{SC[j]}$9+INT(({HC[j]}$12+INDEX($AJ$13:$AR${DATA_LAST + 10},$BW{r},{j + 1})'
                                 f'+INDEX($AT$13:$BB${DATA_LAST + 10},$BX{r},{j + 1})'
                                 f'+INDEX($BD$13:$BL${DATA_LAST + 10},$BY{r},{j + 1}))/100),"")')
        parts = ['IF($T$5>0,"Lv"&$S$5&" "&$CA{r}&" → ","")'.format(r=r)]
        prev = ["CA", "CB", "CC", "CD"]
        for k, cc in enumerate(["CB", "CC", "CD", "CE"], start=1):
            parts.append(f'IF(AND($T${5 + k}>0,OR($S${5 + k}=$B$9,${cc}{r}<>${prev[k - 1]}{r})),"Lv"&$S${5 + k}&" "&${cc}{r}&" → ","")')
        none_txt = '"无可行组合"' if i == 0 else '"—"'
        ws[f"M{r}"] = f'=IF($BZ{r},{"&".join(parts)}&"Lv"&$B$10&" 完成",{none_txt})'
        style_range(ws, f"A{r}:M{r}", align=center)
        ws[f"M{r}"].alignment = left_nowrap
        ws[f"A{r}"].font = f_bold

    # ---- ④ 当前显示路线 ----
    section(ws, 25, "④ 当前显示路线：几级转什么职业、在哪个职业升几级（由上方「显示路线」切换）")
    header(ws, 26, ["阶段", "起始等级", "职业", "结束等级", "升级次数"])
    ws.merge_cells("F26:M26")
    header(ws, 26, ["说明"], start_col=6)
    for j in range(6):
        r = 27 + j
        n = f"$BX{r}"
        ws.cell(r, 1, f"阶段{j + 1}")
        ws[f"B{r}"] = f'=IF({n}>0,MAX($B$9,$BV{r}),"")'
        ws[f"C{r}"] = f'=IF({n}>0,$BW{r},"")'
        ws[f"D{r}"] = f'=IF({n}>0,B{r}+{n},"")'
        ws[f"E{r}"] = f'=IF({n}>0,{n},"")'
        prev_cls = "$BW$26" if j == 0 else f"$BW{r - 1}"
        ws.merge_cells(f"F{r}:M{r}")
        ws[f"F{r}"] = (f'=IF({n}>0,"Lv"&B{r}&IF(B{r}=$B$9,IF(C{r}=$C$9," 以【"&C{r}&"】起步"," 立即转职为【"&C{r}&"】"),IF(C{r}={prev_cls}," 继续【"&C{r}&"】"," 转职为【"&C{r}&"】"))'
                       f'&"，在该职业升 "&{n}&" 级 → Lv"&D{r},"")')
        style_range(ws, f"A{r}:M{r}", align=center)
        ws[f"F{r}"].alignment = left_nowrap
        ws[f"A{r}"].font = f_bold

    # ---- 最终属性对比 ----
    section(ws, 34, "最终属性对比（当前显示路线，目标等级时）")
    header(ws, 35, ["项目", "", ""] + STATS + [""])
    for r, t in zip(range(36, 40), ["路线结果", "目标", "差值", "判定"]):
        ws.cell(r, 1, t).font = f_bold
    tl_first, tl_last = 52, 52 + TL_ROWS - 1
    for j in range(NS):
        s = SC[j]
        ws[f"{s}36"] = f'=IFERROR(INDEX({s}${tl_first}:{s}${tl_last},$B$10-$B$9+1),"")'
        ws[f"{s}37"] = f'=IF({s}$10="","",{s}$10)'
        ws[f"{s}38"] = f'=IF(OR({s}37="",{s}36=""),"",{s}36-{s}37)'
        ws[f"{s}39"] = f'=IF({s}38="","—",IF({s}38>=0,"✓","✗ 差"&-{s}38))'
    ws.merge_cells("B39:C39")
    ws["B39"] = f'=COUNTIF(D39:L39,"✓")&" / "&$W$6&" 项达成"'
    ws["B39"].font = f_bold
    style_range(ws, "A36:L39", align=center)

    # ---- ⑤ 手动路线 ----
    section(ws, 41, "⑤ 手动路线（「显示路线」选「手动」时生效：自己填几级转什么职业，下方逐级表会按此计算）")
    header(ws, 42, ["段", "转职等级", "职业"])
    ws.merge_cells("D42:M42")
    header(ws, 42, ["说明"], start_col=4)
    for j in range(6):
        r = 43 + j
        ws.cell(r, 1, f"段{j + 1}").font = f_bold
        ws.merge_cells(f"D{r}:M{r}")
        if j == 0:
            ws[f"B{r}"] = "=$B$9"
            ws[f"C{r}"] = "=$C$9"
            ws[f"D{r}"] = "起点（自动取上方起点等级/职业）"
            style_range(ws, f"B{r}:C{r}", fill=fill_grey, align=center)
        else:
            style_range(ws, f"B{r}:C{r}", font=f_input, fill=fill_input, align=center)
            ws[f"D{r}"] = "在此等级转职为该职业；不用的段留空" if j == 1 else None
        style_range(ws, f"A{r}:A{r}", align=center)
        style_range(ws, f"D{r}:M{r}", font=f_note, align=left_nowrap)
    ws["B44"], ws["C44"] = 20, "弓箭手"
    ws["B45"], ws["C45"] = 35, "狙击手"

    # ---- ⑥ 逐级成长表 ----
    section(ws, 50, "⑥ 逐级成长表（当前显示路线；绿色=该级属性+1，黄色行=转职）")
    header(ws, 51, ["等级", "阶段", "职业"] + STATS + ["备注"])
    for i in range(TL_ROWS):
        r = tl_first + i
        ws[f"A{r}"] = f'=IF(OR($B$9="",$B$10=""),"",IF($B$9+{i}<=$B$10,$B$9+{i},""))'
        ws[f"B{r}"] = f'=IF($A{r}="","",MATCH($A{r},$BV$27:$BV$32,1))'
        ws[f"C{r}"] = f'=IF($A{r}="","",INDEX($BW$27:$BW$32,$B{r}))'
        for j in range(NS):
            ws[f"{SC[j]}{r}"] = f'=IF($A{r}="","",{SC[j]}$9+INT(($B$14+{CL(85 + j)}{r})/100))'
        if i == 0:
            ws[f"M{r}"] = (f'=IF($A{r}="","",IF($C{r}=$C$9,"起点：【"&$C{r}&"】",'
                           f'"起点：【"&$C$9&"】 ★ 立即转职 → 【"&$C{r}&"】"))')
        else:
            ws[f"M{r}"] = (f'=IF($A{r}="","",IF($C{r}<>$C{r - 1},"★ 转职 → 【"&$C{r}&"】","")'
                           f'&IF($A{r}=$B$10,"（目标等级）",""))')
        style_range(ws, f"A{r}:M{r}", align=center)
        ws[f"M{r}"].alignment = left_nowrap

    # ---- 条件格式 ----
    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    ws.conditional_formatting.add(
        f"D{tl_first + 1}:L{tl_last}",
        FormulaRule(formula=[f"AND(ISNUMBER(D{tl_first + 1}),ISNUMBER(D{tl_first}),D{tl_first + 1}>D{tl_first})"],
                    fill=green, font=Font(color="006100", bold=True)))
    ws.conditional_formatting.add(
        f"A{tl_first}:C{tl_last}", FormulaRule(formula=[f'LEFT($M{tl_first},1)="★"'], fill=yellow))
    ws.conditional_formatting.add(
        f"M{tl_first}:M{tl_last}", FormulaRule(formula=[f'LEFT($M{tl_first},1)="★"'], fill=yellow))
    ws.conditional_formatting.add("D39:L39", FormulaRule(formula=['LEFT(D39,1)="✓"'], fill=green, font=Font(color="006100")))
    ws.conditional_formatting.add("D39:L39", FormulaRule(formula=['LEFT(D39,1)="✗"'], fill=red, font=Font(color="9C0006")))
    ws.conditional_formatting.add("D38:L38", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="9C0006", bold=True)))
    for r in range(19, 24):
        ws.conditional_formatting.add(
            f"D{r}:L{r}", FormulaRule(formula=[f'AND(D$10<>"",ISNUMBER(D{r}),D{r}<D$10)'], fill=red, font=Font(color="9C0006")))
    ws.conditional_formatting.add(
        "B19:B23", FormulaRule(formula=['AND($BZ19,INDEX($BQ$13:$BQ$6012,$BV19)=$W$6)'], fill=green, font=Font(color="006100", bold=True)))

    # ---- 数据验证 ----
    dv_char = DataValidation(type="list", formula1=CHAR_NAMES, allow_blank=True)
    dv_class = DataValidation(type="list", formula1=CLASS_NAMES, allow_blank=True)
    dv_route = DataValidation(type="list", formula1='"推荐1,推荐2,推荐3,推荐4,推荐5,手动"', allow_blank=False)
    dv_lv = DataValidation(type="whole", operator="between", formula1="1", formula2="200", allow_blank=True)
    dv_acc = DataValidation(type="whole", operator="between", formula1="0", formula2="99", allow_blank=False)
    for dv in (dv_char, dv_class, dv_route, dv_lv, dv_acc):
        ws.add_data_validation(dv)
    dv_char.add("B4")
    dv_class.add("C9:C10")
    dv_class.add("C44:C48")
    dv_route.add("K14")
    dv_lv.add("B9:B10")
    dv_lv.add("B44:B48")
    dv_lv.add("E14:H14")
    dv_acc.add("B14")

    ws.freeze_panes = "A7"
    # 隐藏辅助列
    ws.column_dimensions.group("N", "CZ", hidden=True)
    return ws


def build_helpers(ws):
    """辅助计算区：P 列起。"""
    H = lambda ref, v: ws.__setitem__(ref, v)
    H("P3", "【辅助计算区，请勿修改】")
    # 窗口表：行5~9 对应 k=0..4；Q=下限 R=上限 S=本窗口起始等级 T=本窗口升级次数
    header(ws, 4, ["k", "lo", "hi", "e", "n", "参数"], start_col=16)
    los = ["0", "$E$14", "$F$14", "$G$14", "$H$14"]
    his = ["$E$14", "$F$14", "$G$14", "$H$14", "999"]
    for k in range(5):
        r = 5 + k
        H(f"P{r}", k)
        H(f"Q{r}", f"={los[k]}")
        H(f"R{r}", f"={his[k]}")
        H(f"S{r}", f"=IFERROR(MIN(MAX($B$9,Q{r}),$B$10),0)")
        H(f"T{r}", f"=IFERROR(MAX(0,MIN($B$10,R{r})-MAX($B$9,Q{r})),0)")
    # U5 强制为目标职业的窗口；U6~U8 各窗口组合数；U9 总组合；U10 候选是否齐全；U11 目标项数
    H("U5", '=IF(OR($C$10="",$T$9>0),0,IF($T$8>0,3,IF($T$7>0,2,IF($T$6>0,1,0))))')
    for k, col in zip((1, 2, 3), ("AB", "AC", "AD")):
        H(f"V{5 + k}", f"=COUNTIF({col}$13:{col}${DATA_LAST + 10},TRUE)")
        H(f"U{5 + k}", f"=IF($T${5 + k}=0,1,MAX(1,V{5 + k}))")
    H("U9", "=U6*U7*U8")
    H("W5", "=AND(OR($T$6=0,$V$6>0),OR($T$7=0,$V$7>0),OR($T$8=0,$V$8>0))")
    H("W6", "=COUNT($D$10:$L$10)")
    # 行10 起点职业窗口贡献 C0；行11 目标职业窗口贡献 C4；行12 常量 K = 经验槽+C0+C4
    for j in range(NS):
        h, s, d = HC[j], SC[j], DC[j]
        g = lambda cls: f"MAX(0,{s}$7+IFERROR(INDEX({class_col(d)},MATCH({cls},{CLASS_NAMES},0))+0,0))"
        H(f"{h}10", f"=$T$5*{g('$C$9')}")
        H(f"{h}11", f'=IF($C$10="",0,$T$9*{g("$C$10")})')
        H(f"{h}12", f"=$B$14+{h}10+{h}11")
    # 职业辅助表：行13~162 对应职业表第3~152行
    header(ws, 12, ["职业", "最低Lv", "纳入"], start_col=16)
    elig_lo = {1: "$Q$6", 2: "$Q$7", 3: "$Q$8"}
    for i in range(DATA_ROWS):
        r, cr = 13 + i, 3 + i
        H(f"P{r}", f'=IF({CLASS_Q}!$C{cr}="","",{CLASS_Q}!$C{cr})')
        H(f"Q{r}", f"=N({CLASS_Q}!$D{cr})")
        H(f"R{r}", f"={CLASS_Q}!$E{cr}&\"\"")
        for j in range(NS):
            H(f"{HC[j]}{r}", f'=IF($P{r}="",0,MAX(0,{SC[j]}$7+N({CLASS_Q}!{DC[j]}{cr})))')
        for k, (ec, cc) in enumerate((("AB", "AE"), ("AC", "AF"), ("AD", "AG")), start=1):
            H(f"{ec}{r}", (f'=AND($P{r}<>"",IF($U$5={k},$P{r}=$C$10,'
                           f'OR($P{r}=$C$9,AND($R{r}="是",$Q{r}<=MAX($B$9,{elig_lo[k]})))))'))
            H(f"{cc}{r}", f'=IF({ec}{r},COUNTIF({ec}$13:{ec}{r},TRUE),"")')
    # 各窗口候选列表：AI(名)+AJ..AR / AS+AT..BB / BC+BD..BL
    n_eff = {1: "$T$6", 2: "$T$7", 3: '($T$8+IF($C$10="",$T$9,0))'}
    for k, (name_col, cnt_col) in enumerate((("AI", "AE"), ("AS", "AF"), ("BC", "AG")), start=1):
        base = ws[name_col + "1"].column
        for i in range(DATA_ROWS):
            r, p = 13 + i, i + 1
            m = f"MATCH({p},${cnt_col}$13:${cnt_col}${DATA_LAST + 10},0)"
            H(f"{name_col}{r}", f'=IFERROR(INDEX($P$13:$P${DATA_LAST + 10},{m}),"")')
            for j in range(NS):
                H(f"{CL(base + 1 + j)}{r}", f"=IFERROR(INDEX({HC[j]}$13:{HC[j]}${DATA_LAST + 10},{m}),0)*{n_eff[k]}")
    # 组合枚举：BN~BP 三个窗口的候选序号，BQ 达成项数，BR 缺口，BS 属性总和，BT 评分
    header(ws, 12, ["i1", "i2", "i3", "达成", "缺口", "总和", "评分"], start_col=66)
    tgt = "$D$10:$L$10"
    for i in range(COMBO_ROWS):
        r = 13 + i
        ok = f"{i}<$U$9"
        H(f"BN{r}", f"=IF({ok},INT({i}/($U$7*$U$8))+1,1)")
        H(f"BO{r}", f"=IF({ok},MOD(INT({i}/$U$8),$U$7)+1,1)")
        H(f"BP{r}", f"=IF({ok},MOD({i},$U$8)+1,1)")
        fin = (f"($D$9:$L$9+INT(($S$12:$AA$12+INDEX($AJ$13:$AR${DATA_LAST + 10},BN{r},0)"
               f"+INDEX($AT$13:$BB${DATA_LAST + 10},BO{r},0)+INDEX($BD$13:$BL${DATA_LAST + 10},BP{r},0))/100))")
        H(f"BQ{r}", f'=IF({ok},SUMPRODUCT(({tgt}<>"")*({fin}>={tgt})),0)')
        H(f"BR{r}", f'=IF({ok},SUMPRODUCT(({tgt}<>"")*(({tgt}-{fin})>0)*({tgt}-{fin})),0)')
        H(f"BS{r}", f"=IF({ok},SUMPRODUCT({fin}),0)")
        H(f"BT{r}", f"=IF(AND({ok},$W$5),BQ{r}*1000000-MIN(BR{r},999)*1000+BS{r}-{i}/100000,-1000000000-{i})")
    # Top5：BV 组合位置，BW~BY 候选序号，BZ 是否有效，CA~CE 各窗口职业
    header(ws, 18, ["pos", "i1", "i2", "i3", "有效", "k0", "k1", "k2", "k3", "k4"], start_col=74)
    sc = f"$BT$13:$BT${12 + COMBO_ROWS}"
    for i in range(5):
        r = 19 + i
        H(f"BV{r}", f"=MATCH(LARGE({sc},{i + 1}),{sc},0)")
        for col, src in zip(("BW", "BX", "BY"), ("BN", "BO", "BP")):
            H(f"{col}{r}", f"=INDEX(${src}$13:${src}${12 + COMBO_ROWS},BV{r})")
        H(f"BZ{r}", f"=INDEX({sc},BV{r})>-100000000")
        H(f"CA{r}", "=$C$9")
        H(f"CB{r}", f"=IF($T$6>0,INDEX($AI$13:$AI${DATA_LAST + 10},BW{r}),CA{r})")
        H(f"CC{r}", f"=IF($T$7>0,INDEX($AS$13:$AS${DATA_LAST + 10},BX{r}),CB{r})")
        H(f"CD{r}", f"=IF($T$8>0,INDEX($BC$13:$BC${DATA_LAST + 10},BY{r}),CC{r})")
        H(f"CE{r}", f'=IF($C$10<>"",$C$10,CD{r})')
    # 当前显示路线：BV26=方案号(0=手动)，BV27~32 起始等级，BW27~32 职业，BX27~32 升级次数
    H("BU26", "sel")
    H("BV26", '=IF($K$14="手动",0,IFERROR(VALUE(RIGHT($K$14,1)),1))')
    H("BW26", "")
    for k, cc in enumerate(("CA", "CB", "CC", "CD", "CE")):
        r = 27 + k
        H(f"BV{r}", f"=IF($BV$26=0,BV{43 + k},$S${5 + k})")
        H(f"BW{r}", f"=IF($BV$26=0,BW{43 + k},INDEX({cc}$19:{cc}$23,$BV$26))")
    H("BV32", "=IF($BV$26=0,BV48,999)")
    H("BW32", '=IF($BV$26=0,BW48,"")')
    for r in range(27, 33):
        nxt = f"BV{r + 1}" if r < 32 else "999"
        H(f"BX{r}", f"=IFERROR(MAX(0,MIN($B$10,{nxt})-MAX($B$9,BV{r})),0)")
    # 手动路线辅助：BV43~48 起始等级（空=999），BW 职业
    H("BV43", "=$B$9")
    H("BW43", "=$C$9")
    for r in range(44, 49):
        H(f"BV{r}", f'=IF(B{r}="",999,B{r})')
        H(f"BW{r}", f'=IF(B{r}="","",C{r})')
    # 逐级表辅助：CF=职业行号，CG..CO=累计成长
    for i in range(TL_ROWS):
        r = 52 + i
        H(f"CF{r}", f'=IF($A{r}="","",IFERROR(MATCH($C{r},{CLASS_NAMES},0),0))')
        for j in range(NS):
            col = CL(85 + j)
            if i == 0:
                H(f"{col}{r}", f'=IF($A{r}="","",0)')
            else:
                H(f"{col}{r}", (f'=IF($A{r}="","",{col}{r - 1}+MAX(0,{SC[j]}$7+IF($CF{r - 1}=0,0,'
                                f'IFERROR(INDEX({class_col(DC[j])},$CF{r - 1})+0,0))))'))


def main():
    assert CL(85) == "CG" and CL(66) == "BN" and CL(74) == "BV"
    wb = Workbook()
    build_main(wb)
    n_char = build_char_sheet(wb)
    n_class = build_class_sheet(wb)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    print(f"saved {OUT}: {n_char} 角色, {n_class} 职业")


if __name__ == "__main__":
    main()

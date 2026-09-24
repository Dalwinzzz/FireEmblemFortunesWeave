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
from openpyxl.workbook.defined_name import DefinedName

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "tools", "固定成长养成规划器.xlsx")

STATS = ["HP", "力量", "魔力", "速度", "技巧", "防守", "魔防", "幸运", "魅力"]
NS = len(STATS)
DATA_ROWS = 150          # 数据表预留行数（第3行起），方便以后追加角色/职业
DATA_LAST = 2 + DATA_ROWS
COMBO_ROWS = 6000        # 推荐路线搜索的组合上限
REF_ROWS = 4000          # 中途换职优化的候选上限
TL_ROWS = 100            # 逐级成长表行数
# 武器熟练度：角色表列名 / 在职业表文本里搜索用的关键字
SKILLS = ["剑术", "枪术", "斧术", "弓术", "格斗术", "黑魔法", "白魔法", "指挥", "步兵", "马术", "重装", "飞行"]
SKILL_KEYS = ["剑术", "枪术", "斧术", "弓术", "格斗术", "黑魔法", "白魔法", "指挥", "步兵术", "马术", "重装术", "飞行术"]
NK = len(SKILLS)
RANKS = ["F", "E", "E+", "D", "D+", "C", "C+", "B", "B+", "A", "A+", "S", "S+"]

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
NUMERIC = ({"序号", "最低转职等级", "移动力", "精通加成值", "初始等级", "精通所需EXP"}
           | {"初始" + s for s in STATS} | {"基础" + s for s in STATS})


def build_data_sheet(ws, title, hdr, rows, front, growth_cols, extra_formula=None, widths=None, exclude=()):
    """front：放在成长率之前的列名；growth_cols：9 个成长率列名；其余列原样追加在后面（exclude 中的列不显示）。"""
    rest = [h for h in hdr if h not in front and h not in growth_cols and h not in exclude]
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
                v = num(v) if h in growth_cols or h in NUMERIC else v
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
    front = ["编号", "所属势力", "角色", "性别", "初始兵种"]
    widths = {"编号": 6, "所属势力": 14, "角色": 12, "性别": 6, "初始兵种": 10, "成长合计": 9, "初始等级": 8,
              "个人特技说明": 36, "血印说明": 36, "最早可加入章节": 30, "加入条件": 40, "挖角条件": 50,
              "装备特技": 30, "可习得战技": 40, "擅长": 22, "弱项": 18}
    widths.update({"初始" + s: 7 for s in STATS})
    widths.update({s: 6 for s in SKILLS})
    cols = build_data_sheet(ws, "角色（个人成长率%、初始能力值〔不含兵种基础值〕、初始技能等级；可直接订正，角色名需与下拉框一致）",
                            hdr, rows, front, STATS,
                            extra_formula=lambda r: f'=IF(COUNT(F{r}:N{r})=0,"—",SUM(F{r}:N{r}))', widths=widths)
    assert cols[2] == "角色" and cols[5:14] == STATS
    return len(rows), cols


def build_class_sheet(wb):
    hdr, rows = load_csv("classes.csv")
    ws = wb.create_sheet(SH_CLASS)
    front = ["序号", "阶级", "兵种名", "最低转职等级", "限定"]
    widths = {"序号": 6, "阶级": 9, "兵种名": 13, "最低转职等级": 9, "限定": 7, "转职条件": 60,
              "使用武器": 34, "技能EXP加成": 30, "兵种特性": 16, "兵种固有特技": 30, "精通特技": 18,
              "转职·主要技能": 18, "转职·选择技能": 22}
    widths.update({"基础" + s: 7 for s in STATS})
    cols = build_data_sheet(ws, "职业（成长率补正%、兵种基础能力值、转职条件；可直接订正/追加）", hdr, rows, front, STATS,
                            widths=widths, exclude=("纳入推荐",))
    assert cols[2] == "兵种名" and cols[5:14] == STATS
    ws.cell(2, 5).value = "限定\n(女/角色名)"
    for r in range(3, 3 + len(rows)):
        for c in (4, 5, cols.index("精通加成属性") + 1, cols.index("精通加成值") + 1):
            ws.cell(r, c).fill = fill_input
            ws.cell(r, c).font = f_input
    mcol = CL(cols.index("精通加成属性") + 1)
    dv_m = DataValidation(type="list", formula1='"' + ",".join(STATS) + '"', allow_blank=True)
    ws.add_data_validation(dv_m)
    dv_m.add(f"{mcol}3:{mcol}{DATA_LAST}")
    note_r = 3 + len(rows) + 1
    ws.cell(note_r, 1, "说明：").font = f_bold
    notes = [
        "「最低转职等级」：初级5、中级20、上级35、最上级45（暂定）、神将45（暂定）。自动推荐只会在达到该等级后才转入此职业。",
        "「限定」：女=女性专用；填角色名=该角色专用（如舞者=蕾达）。自动推荐会跳过不符合的职业。参与推荐与否在「养成规划」页⑧勾选。",
        "「基础」开头的列是兵种基础能力值：面板属性 = 个人能力 + 当前兵种基础值，换职业时面板会随之变化（空白按0计）。",
        "「精通加成属性/值」：在该职业累计练满「精通需练级数」（规划页，默认5级）即视为精通，最终属性加上该加成（如天翼兵 魔防+3）。可按实际数据补充其他职业。",
        "「转职·主要技能」为转职必需的武器熟练度（全部满足），「转职·选择技能」为满足其一即可，由「转职条件」原文解析而来；「使用武器」决定在该职业期间能练哪些熟练度。",
        "要追加职业，在表格末尾接着填一行即可（第152行以内），兵种名会自动出现在养成规划的下拉框里。",
    ]
    for i, t in enumerate(notes, start=1):
        ws.merge_cells(start_row=note_r + i, start_column=1, end_row=note_r + i, end_column=20)
        c = ws.cell(note_r + i, 1, t)
        c.font = f_note
        c.alignment = left_nowrap
    defaults = [(row[hdr.index("兵种名")], row[hdr.index("纳入推荐")]) for row in rows]
    return len(rows), cols, defaults


# ======================================================================
# 主页面
# ======================================================================
SC = [CL(4 + j) for j in range(NS)]           # 主页面属性列 D..L
WC = [CL(2 + s) for s in range(NK)]           # 主页面熟练度列 B..M
DC = [CL(6 + j) for j in range(NS)]           # 数据表属性列 F..N
LAST_VIS = "P"                                # 可见区最右列
I_STR, I_MAG, I_SPD, I_DEF, I_RES = 1, 2, 3, 5, 6   # 力量/魔力/速度/防守/魔防 在 STATS 中的下标

CHAR_NAMES = f"{CHAR_Q}!$C$3:$C${DATA_LAST}"
CLASS_NAMES = f"{CLASS_Q}!$C$3:$C${DATA_LAST}"
H0 = 5                                        # 辅助表首行
HZ = DATA_ROWS + 1                            # 职业辅助表的「空行」序号：序号0统一映射到这一行（全为0）
CLS_LAST = H0 + DATA_ROWS                     # 职业辅助表末行（含空行）

# 可见区行号
R_WSEC, R_WHDR, R_WINIT, R_WCUR, R_WUSE = 17, 18, 19, 20, 21
R_BSEC, R_BHDR, R_B0 = 23, 24, 25
R_OSEC, R_OHDR, R_O0 = 31, 32, 33
NSEG = 10                                     # 当前显示路线 / 手动路线的段数
R_SSEC, R_SHDR, R_S0 = 39, 40, 41             # 41~50
R_FSEC, R_FHDR, R_F0 = 52, 53, 54             # 54 成长结果 55 精通加成 56 最终 57 目标 58 差值 59 判定
R_ESEC, R_E0 = 61, 62                         # 62 我方武器 63 速度 64 攻击 65 物防 66 魔防 67 综合
R_MSEC, R_MHDR, R_M0 = 69, 70, 71             # 71~80
R_CSEC, R_CHDR, R_C0 = 82, 83, 84             # 候选职业勾选 84~99
C_ROWS, C_GROUPS = 16, 4
R_TSEC, R_THDR, R_T0 = 101, 102, 103
R_TLAST = R_T0 + TL_ROWS - 1
SERIES = ["斗士系", "猎兵系", "士兵系", "魔道系", "飞行系"]

# 实战评估参数（可见区右上 N3:P12；我方武器在实战评估区）
EN = {"spd": "$O$4", "patk": "$O$5", "matk": "$O$6", "pdef": "$O$7", "mdef": "$O$8",
      "type": "$O$9", "speed": "$O$10", "front": "$O$11", "base": "$O$12",
      "might": f"$C${R_E0}", "wt": f"$E${R_E0}"}
# 候选职业勾选格：第 i 个职业（0 起）→ 名称格 / 勾选格
GRID_NAME_COLS = ["A", "C", "E", "H"]
GRID_MARK_COLS = ["B", "D", "G", "J"]
SERIES_NAME_COL, SERIES_MARK_COL = "L", "M"


def grid_cell(i, kind):
    g, r = divmod(i, C_ROWS)
    col = (GRID_NAME_COLS if kind == "name" else GRID_MARK_COLS)[g]
    return f"{col}{R_C0 + r}"


def class_col(letter, last=DATA_LAST + 1):
    return f"{CLASS_Q}!${letter}$3:${letter}${last}"


def Z(x):
    """职业序号 0（找不到）→ 指向全 0 的空行，避免 INDEX(...,0) 返回整列。"""
    return f"IF({x}=0,{HZ},{x})"


class Alloc:
    def __init__(self, start):
        self.c = start

    def take(self, n=1):
        cols = [CL(self.c + i) for i in range(n)]
        self.c += n
        return cols if n > 1 else cols[0]


def rng(col, r1, r2):
    return f"${col}${r1}:${col}${r2}"


def mrng(c1, c2, r1, r2):
    return f"${c1}${r1}:${c2}${r2}"


def build_main(wb, char_cols, class_cols, class_defaults):
    ws = wb.active
    ws.title = SH_MAIN
    ws.sheet_view.showGridLines = False
    for row in ws.iter_rows(min_row=1, max_row=R_TLAST, max_col=16):
        for c in row:
            c.font = f_base
    widths = {"A": 12, "B": 10, "C": 13, "M": 50, "N": 24, "O": 40, "P": 40}
    for col in "ABCDEFGHIJKLMNOP":
        ws.column_dimensions[col].width = widths.get(col, 8)

    # 数据表中的列字母
    ccol = {h: CL(i + 1) for i, h in enumerate(char_cols)}
    kcol = {h: CL(i + 1) for i, h in enumerate(class_cols)}
    SK_CHAR = [ccol[s] for s in SKILLS]
    C_USE, C_MAIN, C_SEL = kcol["使用武器"], kcol["转职·主要技能"], kcol["转职·选择技能"]
    C_COND, C_FEAT, C_RESTR = kcol["转职条件"], kcol["兵种特性"], kcol["限定"]
    C_BASE = [kcol["基础" + s] for s in STATS]
    C_MBA, C_MBV = kcol["精通加成属性"], kcol["精通加成值"]

    # ------------------------------------------------------------------
    # 辅助区列分配（Q 列起全部隐藏）
    # ------------------------------------------------------------------
    A = Alloc(18)
    SCL_L, SCL_V = A.take(), A.take()
    W_LO, W_HI, W_E, W_N = A.take(), A.take(), A.take(), A.take()
    V9 = A.take(NS)            # 行5 C0，行6 C4，行7 K，行8 KT，行9 当前路线精通加成，行10 起点兵种基础值
    V13 = A.take(13)           # 行5 等级表，行6 当前熟练度，行7 可练标记
    CLS = {"name": A.take(), "min": A.take(), "avail": A.take(), "low": A.take(),
           "G": A.take(NS), "T": A.take(NK), "RM": A.take(NK), "RS": A.take(NK), "NM": A.take(NK), "SU": A.take(NK),
           "SA": A.take(), "SS": A.take(), "EL": A.take(3), "CN": A.take(3), "ELU": A.take(4), "CNU": A.take(4),
           "MB": A.take(NS), "MBANY": A.take(), "MBTXT": A.take(), "BS": A.take(NS), "ok": A.take()}
    FL = [{"idx": A.take(), "name": A.take(), "C": A.take(NS)} for _ in range(3)]
    UL = A.take(4)
    CMB = {k: A.take() for k in ("i1", "i2", "i3", "id1", "id2", "id3", "id4", "w0", "w1", "w2", "w3", "w4",
                                 "low", "met", "short", "total", "combat", "score")}
    CMB["F"] = A.take(NS)
    REF = {k: A.take() for k in ("k", "j", "x", "B", "P", "fin", "valid", "w0", "w1", "w2", "w3", "w4", "wB",
                                 "low", "met", "short", "total", "combat", "score", "best")}
    REF["F"] = A.take(NS)
    RK = {k: A.take() for k in ("pos", "valid", "met", "short", "total", "bad", "low", "combat", "mtxt",
                                "i1", "i2", "i3", "k", "x", "B", "P", "head", "split", "first", "second")}
    RK["FROM"], RK["IDX"], RK["NAME"], RK["N"], RK["F"] = A.take(6), A.take(6), A.take(6), A.take(6), A.take(NS)
    SEG = {k: A.take() for k in ("from", "idx", "name", "n", "prev", "piece", "scan", "lv", "first", "w")}
    SEG["U"] = A.take(NK)
    MAN = {k: A.take() for k in ("from", "idx", "name", "n", "prev", "piece", "scan")}
    MAN["U"] = A.take(NK)
    TLH = {"idx": A.take(), "cum": A.take(NS)}
    SEGI = SEG["idx"]
    LAST_HELPER = A.c - 1

    H = ws.__setitem__
    # 窗口表：k=0..4 → 行 5..9
    E = [f"${W_E}${H0 + k}" for k in range(5)]
    N = [f"${W_N}${H0 + k}" for k in range(5)]
    LO = [f"${W_LO}${H0 + k}" for k in range(5)]
    EW = rng(W_E, H0 + 1, H0 + 4)        # k=1..4
    NW = rng(W_N, H0 + 1, H0 + 4)

    # 标量：先分配位置（公式互相引用），再写入
    names_order = (["startCls", "startIdx", "tgZ", "tgtIdx", "tgtMin", "force", "cnt1", "cnt2", "cnt3", "m1", "m2", "m3",
                    "M", "listOK", "tgtCnt", "mm1", "mm2", "mm3", "mm4", "sz1", "sz2", "sz3", "sz4",
                    "O1", "O2", "O3", "O4", "R", "NN1", "NN2", "NN3", "NN4", "head1", "head2", "head3", "head4",
                    "P1", "P2", "P3", "P4", "validT", "lowT", "entryLv", "sel", "rankMissing", "manualBad",
                    "Nm", "clsRef", "typeMag", "front", "fast", "eD", "combatN", "useBase", "lastWin",
                    "charName", "gender", "manualHas"]
                   + [f"bLv{k}" for k in range(5)] + [f"bFirst{k}" for k in range(5)])
    scal = {nm: f"${SCL_V}${H0 + i}" for i, nm in enumerate(names_order)}

    def S(name, formula):
        r = int(scal[name].split("$")[-1])
        H(f"{SCL_L}{r}", name)
        H(f"{SCL_V}{r}", formula)

    def ref(name):
        return scal[name]

    CS = ref("startCls")      # 起点职业（C9 留空 = 角色初始兵种）
    si = ref("startIdx")
    # 常用区域定义为工作簿名称，大幅缩短公式（文件体积和解析时间）
    defnames = {}

    def dn(name, ref_):
        defnames[name] = ref_
        return name

    NAMES = dn("z_NAMES", rng(CLS["name"], H0, CLS_LAST))
    MINC = dn("z_MIN", rng(CLS["min"], H0, CLS_LAST))
    LOWC = dn("z_LOW", rng(CLS["low"], H0, CLS_LAST))
    MBANY = dn("z_MBANY", rng(CLS["MBANY"], H0, CLS_LAST))
    MBTXT = dn("z_MBTXT", rng(CLS["MBTXT"], H0, CLS_LAST))
    SAC, SSC = dn("z_SA", rng(CLS["SA"], H0, CLS_LAST)), dn("z_SS", rng(CLS["SS"], H0, CLS_LAST))
    GMAT = dn("z_GMAT", mrng(CLS["G"][0], CLS["G"][-1], H0, CLS_LAST))
    TMAT = dn("z_TMAT", mrng(CLS["T"][0], CLS["T"][-1], H0, CLS_LAST))
    NMMAT = dn("z_NMMAT", mrng(CLS["NM"][0], CLS["NM"][-1], H0, CLS_LAST))
    SUMAT = dn("z_SUMAT", mrng(CLS["SU"][0], CLS["SU"][-1], H0, CLS_LAST))
    _g = [dn(f"z_G{j}", rng(CLS["G"][j], H0, CLS_LAST)) for j in range(NS)]
    _mb = [dn(f"z_MB{j}", rng(CLS["MB"][j], H0, CLS_LAST)) for j in range(NS)]
    _bs = [dn(f"z_BS{j}", rng(CLS["BS"][j], H0, CLS_LAST)) for j in range(NS)]
    gcol = lambda j: _g[j]
    mbcol = lambda j: _mb[j]
    bscol = lambda j: _bs[j]
    BSTART = [f"${V9[j]}${H0 + 5}" for j in range(NS)]   # 起点兵种基础值

    def base_adj(j, idx):
        """换到职业 idx 后面板的兵种基础值变化（开关关闭时为 0）。"""
        return f"{ref('useBase')}*(INDEX({bscol(j)},{idx})-{BSTART[j]})"
    RANKROW = dn("z_RANKS", mrng(V13[0], V13[12], H0, H0))
    CURROW = dn("z_CUR", mrng(V13[0], V13[NK - 1], H0 + 1, H0 + 1))
    CTROW = dn("z_CT", mrng(V13[0], V13[NK - 1], H0 + 2, H0 + 2))
    TGT = "$D$10:$L$10"
    FLI = [dn(f"z_L{k + 1}IDX", rng(FL[k]["idx"], H0, H0 + DATA_ROWS - 1)) for k in range(3)]
    _flc = [[dn(f"z_L{k + 1}C{j}", rng(FL[k]["C"][j], H0, H0 + DATA_ROWS - 1)) for j in range(NS)] for k in range(3)]
    flc = lambda k, j: _flc[k][j]
    ULMAT = dn("z_ULMAT", mrng(UL[0], UL[3], H0, H0 + DATA_ROWS - 1))
    OROW = dn("z_OFS", f"{ref('O1')}:{ref('O4')}")
    NNROW = dn("z_NN", f"{ref('NN1')}:{ref('NN4')}")
    HEADROW = dn("z_HEAD", f"{ref('head1')}:{ref('head4')}")
    PROW = dn("z_P", f"{ref('P1')}:{ref('P4')}")
    Nm = ref("Nm")
    trow = lambda x: f"INDEX({TMAT},{Z(x)},0)"

    def badc(x, u):
        """转入职业 x 时，熟练度需注意（之前职业练不了）的项数；u = 之前职业累计可练向量表达式。"""
        zx = Z(x)
        return (f"(SUMPRODUCT(INDEX({NMMAT},{zx},0)*((({u})*{CTROW})=0))"
                f"+IF(AND(INDEX({SAC},{zx})=1,INDEX({SSC},{zx})=0,"
                f"SUMPRODUCT(INDEX({SUMAT},{zx},0)*((({u})*{CTROW})>0))=0),1,0))")

    def aspd(F):
        """攻速 = 速度 − MAX(0, 武器重量 − 体格)，体格 = INT(力量/5)。"""
        return f"({F[I_SPD]}-MAX(0,N({EN['wt']})-INT({F[I_STR]}/5)))"

    def combat(F):
        """实战评估通过项数。F = 9 个最终属性单元格。"""
        t, fa, fr, eD = ref("typeMag"), ref("fast"), ref("front"), ref("eD")
        a = aspd(F)
        return (f'IF({EN["spd"]}="",0,IF({fa}=1,{a}>={EN["spd"]}+4,{a}>={EN["spd"]}-3)*1)'
                f'+IF({eD}="",0,(IF({t}=1,{F[I_MAG]},{F[I_STR]})+N({EN["might"]})>{eD})*1)'
                f'+IF(AND({fr}=1,{EN["patk"]}<>""),({F[I_DEF]}>={EN["patk"]}-5)*1,0)'
                f'+IF(AND({fr}=1,{EN["matk"]}<>""),({F[I_RES]}>={EN["matk"]}-5)*1,0)')

    def score(ok, low, met, short, total, cmb, F, i):
        # 排序：达成项 > 实战评估 > 不含靠后职业 > 缺口 > 力或魔 > 属性总和（熟练度只作提示，不参与排序）；
        # 量级控制在 14 位有效数字内（LibreOffice 只保留 15 位），否则区分并列用的小数会丢失
        atk = f"IF({ref('typeMag')}=1,{F[I_MAG]},{F[I_STR]})"
        return (f"=IF({ok},{met}*1E+9+{cmb}*1E+8+({low}=0)*1E+7"
                f"-MIN({short},99)*1E+5+MIN({atk},99)*1E+3+MIN({total},999)-{i}/1E+4,-1E+11-{i})")

    def eval_cells(c, F, ok):
        frow = f"{F[0]}:{F[-1]}"
        H(c["met"], f'=IF({ok},SUMPRODUCT(({TGT}<>"")*({frow}>={TGT})),0)')
        H(c["short"], f'=IF({ok},SUMPRODUCT(({TGT}<>"")*(({TGT}-{frow})>0)*({TGT}-{frow})),0)')
        H(c["total"], f"=IF({ok},SUM({frow}),0)")
        H(c["combat"], f"=IF({ok},{combat(F)},0)")

    # ------------------------------------------------------------------
    # 标题 / 输入区
    # ------------------------------------------------------------------
    ws.merge_cells(f"A1:{LAST_VIS}1")
    ws["A1"] = "火焰纹章 万缕千丝 ｜ 固定成长 · 养成路线规划器"
    ws["A1"].font = f_title
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A2:{LAST_VIS}2")
    ws["A2"] = ("用法：黄底蓝字是输入格。① 选角色 → 填起点/目标的等级、职业、属性（目标属性不填=不作要求）→ 右侧⑦填敌方参考值 → ② 填当前武器熟练度 → "
                "③ 看推荐路线与中途换职优化，在「显示路线」切换 → ④ 每段的转职等级、所需熟练度与解锁条件 → 最终属性与实战评估 → ⑥ 逐级成长表。"
                "规则：每升1级，各属性经验槽 +（角色成长率+职业补正，最低0），满100则该属性+1。")
    ws["A2"].font = f_note
    ws["A2"].alignment = left
    ws.row_dimensions[2].height = 42

    ws["A4"] = "① 角色"
    ws["A4"].font = f_bold
    ws.merge_cells("B4:C4")
    ws["B4"] = "蕾达"
    style_range(ws, "B4:C4", font=f_input, fill=fill_input, align=center)
    char_match = f"MATCH($B$4,{CHAR_NAMES},0)"
    ws["E4"] = "势力"
    ws.merge_cells("F4:G4")
    ws["F4"] = f'=IFERROR(INDEX({CHAR_Q}!$B$3:$B${DATA_LAST},{char_match})&"","")'
    ws["H4"] = "初始兵种\n(自动)"
    ws.merge_cells("I4:J4")
    ws["I4"] = f'=IFERROR(INDEX({CHAR_Q}!$E$3:$E${DATA_LAST},{char_match})&"","")'
    ws.merge_cells("K4:M4")
    ws["K4"] = (f'=IF($B$4="","请选择角色",IF(ISNA({char_match}),"⚠ 角色表里找不到该角色",'
                f'IF(COUNT(INDEX({CHAR_Q}!$F$3:$N${DATA_LAST},{char_match},0))=0,"⚠ 该角色成长率数据缺失（按0计算）",'
                f'"✓ 成长率已读取（初始兵种随角色自动显示；起点职业在C9选，留空=用初始兵种）")))')
    for r in ("E4", "H4"):
        ws[r].font = f_bold
        ws[r].alignment = center
    style_range(ws, "F4:G4", fill=fill_grey, align=center)
    style_range(ws, "I4:J4", fill=fill_grey, align=center)
    ws["K4"].font = f_bold
    ws["K4"].alignment = left
    ws.row_dimensions[4].height = 30

    header(ws, 6, ["项目", "等级", "职业"] + STATS + ["说明"])
    labels = {7: "角色成长率%", 8: "资料初始值", 9: "起点", 10: "目标", 11: "需提升", 12: "每级需成长%"}
    notes = {7: "来自「角色」表", 8: "角色表的初始等级/兵种/能力值（面板=个人能力+兵种基础值），可照抄到起点",
             9: "填当前等级、职业（留空=初始兵种）和面板属性（示例=蕾达初始值）",
             10: "目标属性留空=不作要求；目标职业留空=45级后沿用上一职业",
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
        ic = ccol["初始" + STATS[j]]
        ws[f"{s}8"] = (f'=IFERROR(INDEX({CHAR_Q}!{ic}$3:{ic}${DATA_LAST},{char_match})'
                       f'+{ref("useBase")}*INDEX({bscol(j)},IFERROR(MATCH($I$4,{NAMES},0),{HZ})),"")')
        ws[f"{s}11"] = f'=IF(OR({s}10="",{s}9=""),"",{s}10-{s}9)'
        ws[f"{s}12"] = (f'=IF(OR({s}10="",$B$10="",$B$9="",$B$10<=$B$9),"",'
                        f'ROUND(MAX(0,({s}10-{s}9)*100-$B$14)/($B$10-$B$9),1))')
    for r in (7, 8, 11, 12):
        style_range(ws, f"B{r}:L{r}", fill=fill_grey, align=center)
    ws["B8"] = f'=IFERROR(INDEX({CHAR_Q}!{ccol["初始等级"]}$3:{ccol["初始等级"]}${DATA_LAST},{char_match})+0,"")'
    ws["C8"] = "=$I$4"
    ws["B9"], ws["C9"] = 1, "平民"
    ws["B10"], ws["C10"] = 40, "狙击手"
    start = [25, 8, 6, 12, 8, 6, 6, 5, 10]
    target = [40, 22, None, 30, 27, None, None, None, None]
    for j in range(NS):
        ws[f"{SC[j]}9"] = start[j]
        ws[f"{SC[j]}10"] = target[j]
    style_range(ws, "B9:L10", font=f_input, fill=fill_input, align=center)

    # 参数
    ws["A13"] = "换职最少练级"
    ws["A13"].font = f_bold
    ws["B13"] = 3
    ws.merge_cells("C13:F13")
    ws["C13"] = "「优化」中途换职前后两段各至少练几级"
    ws.merge_cells("G13:H13")
    ws["G13"] = "精通需练级数"
    ws["G13"].font = f_bold
    ws["G13"].alignment = center
    ws["I13"] = 5
    ws.merge_cells("J13:L13")
    ws["J13"] = "同一职业累计练满即算精通（如天翼兵 魔防+3）"
    for r in ("C13", "J13"):
        ws[r].font = f_note
        ws[r].alignment = left_nowrap
    style_range(ws, "B13:B13", font=f_input, fill=fill_input, align=center)
    style_range(ws, "I13:I13", font=f_input, fill=fill_input, align=center)
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
    ws["K14"] = "自动"
    style_range(ws, "B14:B14", font=f_input, fill=fill_input, align=center)
    style_range(ws, "E14:H14", font=f_input, fill=fill_input, align=center)
    style_range(ws, "K14:L14", font=f_input, fill=fill_input, align=center)
    ws.merge_cells(f"M14:{LAST_VIS}14")
    ws["M14"] = (f'="当前显示："&IF({ref("sel")}=0,"手动路线",IF({ref("sel")}>5,"优化"&({ref("sel")}-5),"推荐"&{ref("sel")}))'
                 f'&"　｜显示路线：自动=⑤手动路线填了内容就显示手动，否则显示推荐1；也可指定 推荐1~5 / 优化1~5 / 手动。'
                 f'经验槽默认0=最保守；节点=推荐路线换职业的等级（初级5/中级20/上级35/最上级45）"')
    ws["M14"].font = f_note
    ws["M14"].alignment = left
    ws.row_dimensions[14].height = 30

    # ⑦ 实战评估参数（右上）
    ws.merge_cells("N3:P3")
    ws["N3"] = "⑦ 实战评估参数（留空=不评估该项）"
    ws["N3"].font = f_hdr
    ws["N3"].fill = fill_hdr
    ws["N3"].alignment = left_nowrap
    en_rows = [
        ("spd", "敌方速度", 30, "我方攻速 ≥ 敌速+4 可追击；敌速 ≥ 我方攻速+4 会被追击（示例值）"),
        ("patk", "敌方物理攻击", 30, "前排：防守 ≥ 敌物攻−5 视为达标"),
        ("matk", "敌方魔法攻击", 25, "前排：魔防 ≥ 敌魔攻−5 视为达标"),
        ("pdef", "敌方防守", 15, "物理角色：力量+武器威力 > 敌防守 = 破甲"),
        ("mdef", "敌方魔防", 10, "魔法角色：魔力+武器威力 > 敌魔防 = 破甲"),
        ("type", "我方攻击类型", "自动", None),
        ("speed", "速度定位", "自动", None),
        ("front", "前排定位", "自动", None),
        ("base", "计入兵种基础值", "是", "面板=个人能力+当前兵种基础值，换职业面板随之变化"),
    ]
    for i, (key, label, default, note) in enumerate(en_rows):
        r = 4 + i
        ws[f"N{r}"] = label
        ws[f"N{r}"].font = f_bold
        ws[f"N{r}"].alignment = center
        ws[f"N{r}"].border = border
        ws[f"O{r}"] = default
        style_range(ws, f"O{r}:O{r}", font=f_input, fill=fill_input, align=center)
        ws[f"P{r}"] = note
        ws[f"P{r}"].font = f_note
        ws[f"P{r}"].alignment = left_nowrap
        ws[f"P{r}"].border = border
    ws["P9"] = f'="→ "&IF({ref("typeMag")}=1,"魔法","物理")&"（自动=按目标职业的力/魔成长判断）"'
    ws["P10"] = f'="→ "&IF({ref("fast")}=1,"高速：以追击为目标","低速：以不被追击为目标")&"（自动=前排按低速，否则速度成长≥50为高速）"'
    ws["P11"] = f'="→ "&IF({ref("front")}=1,"前排：评估物防/魔防","非前排")&"（自动=目标职业为重装）"'

    # ② 武器熟练度
    section(ws, R_WSEC, "② 武器熟练度（用于提示转职条件；每个职业只能练它「使用武器」里的熟练度）", LAST_VIS)
    header(ws, R_WHDR, ["项目"] + SKILLS + ["说明"])
    ws.merge_cells(f"N{R_WHDR}:{LAST_VIS}{R_WHDR}")
    for r, t, note in ((R_WINIT, "角色初始", "来自「角色成长率」表（—=数据缺失）"),
                       (R_WCUR, "当前(可填)", "填游戏里当前的熟练度，留空=用角色初始；×=无法习得"),
                       (R_WUSE, "计算用", "缺失按E计算")):
        ws.cell(r, 1, t).font = f_bold
        ws.cell(r, 1).alignment = center
        ws.cell(r, 1).border = border
        ws.merge_cells(f"N{r}:{LAST_VIS}{r}")
        ws[f"N{r}"] = note
        style_range(ws, f"N{r}:{LAST_VIS}{r}", font=f_note, align=left_nowrap)
    for s in range(NK):
        w, cc = WC[s], SK_CHAR[s]
        ws[f"{w}{R_WINIT}"] = f'=IFERROR(INDEX({CHAR_Q}!{cc}$3:{cc}${DATA_LAST},{char_match})&"","")'
        v = f'IF({w}{R_WCUR}<>"",{w}{R_WCUR},{w}{R_WINIT})'
        ws[f"{w}{R_WUSE}"] = f'=IF({v}="×","×",IFERROR(INDEX({RANKROW},1,MATCH({v},{RANKROW},0)),"E"))'
    style_range(ws, f"B{R_WINIT}:M{R_WINIT}", fill=fill_grey, align=center)
    style_range(ws, f"B{R_WCUR}:M{R_WCUR}", font=f_input, fill=fill_input, align=center)
    style_range(ws, f"B{R_WUSE}:M{R_WUSE}", fill=fill_grey, align=center)

    # ------------------------------------------------------------------
    # 辅助区：窗口表 / 标量 / 向量
    # ------------------------------------------------------------------
    H(f"{SCL_L}4", "【辅助计算区，请勿修改】")
    header(ws, 4, ["lo", "hi", "e", "n"], start_col=ws[f"{W_LO}1"].column)
    los = ["0", "$E$14", "$F$14", "$G$14", "$H$14"]
    his = ["$E$14", "$F$14", "$G$14", "$H$14", "999"]
    for k in range(5):
        r = H0 + k
        H(f"{W_LO}{r}", f"={los[k]}")
        H(f"{W_HI}{r}", f"={his[k]}")
        H(f"{W_E}{r}", f"=IFERROR(MIN(MAX($B$9,{W_LO}{r}),$B$10),0)")
        H(f"{W_N}{r}", f"=IFERROR(MAX(0,MIN($B$10,{W_HI}{r})-MAX($B$9,{W_LO}{r})),0)")

    S("startCls", '=IF($C$9="",$I$4,$C$9)')
    S("startIdx", f"=IFERROR(MATCH({CS},{NAMES},0),{HZ})")  # 找不到时指向全0空行
    S("tgtIdx", f'=IF($C$10="",0,IFERROR(MATCH($C$10,{NAMES},0),0))')
    S("tgZ", f"={Z(ref('tgtIdx'))}")
    S("tgtMin", f"=INDEX({MINC},{Z(ref('tgtIdx'))})")
    S("force", f'=IF(OR($C$10="",{N[4]}>0),0,IF({N[3]}>0,3,IF({N[2]}>0,2,IF({N[1]}>0,1,0))))')
    for k in range(3):
        S(f"cnt{k + 1}", f"=COUNTIF({rng(CLS['EL'][k], H0, CLS_LAST)},TRUE)")
        S(f"m{k + 1}", f"=IF({N[k + 1]}=0,1,MAX(1,{ref(f'cnt{k + 1}')}))")
    S("M", f"={ref('m1')}*{ref('m2')}*{ref('m3')}")
    S("listOK", "=AND(" + ",".join(f"OR({N[k + 1]}=0,{ref(f'cnt{k + 1}')}>0)" for k in range(3)) + ")")
    S("tgtCnt", "=COUNT($D$10:$L$10)")
    for k in range(4):
        S(f"mm{k + 1}", f"=COUNTIF({rng(CLS['ELU'][k], H0, CLS_LAST)},TRUE)")
        S(f"sz{k + 1}", f"=IF({N[k + 1]}>=2,{ref(f'mm{k + 1}')}*({N[k + 1]}-1),0)")
        S(f"NN{k + 1}", f"=MAX(1,{N[k + 1]}-1)")
        S(f"P{k + 1}", f"=${RK['IDX'][k + 1]}${H0}")
    S("O1", "=0")
    for k in range(1, 4):
        S(f"O{k + 1}", f"={ref(f'O{k}')}+{ref(f'sz{k}')}")
    S("R", f"={ref('O4')}+{ref('sz4')}")
    for k in range(3):
        S(f"head{k + 1}", f"=IF({ref('force')}={k + 1},1,0)")
    S("head4", f'=IF(AND($C$10<>"",{N[4]}>0),1,0)')
    S("validT", f"=${RK['valid']}${H0}")
    S("lowT", f"=${RK['low']}${H0}")
    S("entryLv", f"=IF({N[4]}>0,{E[4]},IF({ref('force')}>0,INDEX({rng(W_E, H0, H0 + 4)},{ref('force')}+1),$B$9))")
    S("manualHas", f"=COUNTA($B${R_M0 + 1}:$B${R_M0 + NSEG - 1})>0")
    S("sel", (f'=IF($K$14="手动",0,IF(OR($K$14="自动",$K$14=""),IF({ref("manualHas")},0,1),'
              f'IF(LEFT($K$14,2)="优化",5,0)+IFERROR(VALUE(RIGHT($K$14,1)),1)))'))
    S("rankMissing", f'=AND(COUNTIF($B${R_WINIT}:$M${R_WINIT},"—")+COUNTBLANK($B${R_WINIT}:$M${R_WINIT})>=12,'
                     f'COUNTA($B${R_WCUR}:$M${R_WCUR})=0)')
    S("manualBad", "=OR(" + ",".join(f"${MAN['from']}${H0 + j + 1}<${MAN['from']}${H0 + j}" for j in range(NSEG - 1)) + ")")
    S("Nm", "=MAX(1,N($I$13))")
    S("useBase", f'=IF({EN["base"]}="否",0,1)')
    S("lastWin", f"=IF({N[4]}>0,4,IF({N[3]}>0,3,IF({N[2]}>0,2,IF({N[1]}>0,1,0))))")
    S("charName", "=$B$4")
    S("gender", f'=IFERROR(INDEX({CHAR_Q}!${ccol["性别"]}$3:${ccol["性别"]}${DATA_LAST},{char_match})&"","")')
    S("clsRef", f"=IF({ref('tgtIdx')}>0,{ref('tgtIdx')},{si})")
    zc = Z(ref("clsRef"))
    S("typeMag", (f'=IF({EN["type"]}="魔法",1,IF({EN["type"]}="物理",0,'
                  f'IF(INDEX({gcol(I_MAG)},{zc})>INDEX({gcol(I_STR)},{zc}),1,0)))'))
    S("front", (f'=IF({EN["front"]}="是",1,IF({EN["front"]}="否",0,'
                f'IF(ISNUMBER(SEARCH("重装",INDEX({class_col(C_FEAT)},{zc}))),1,0)))'))
    S("fast", (f'=IF({EN["speed"]}="高速追击",1,IF({EN["speed"]}="低速防追",0,'
               f'IF({ref("front")}=1,0,IF(INDEX({gcol(I_SPD)},{zc})>=50,1,0))))'))
    S("eD", (f'=IF({ref("typeMag")}=1,IF({EN["mdef"]}="","",{EN["mdef"]}),'
             f'IF({EN["pdef"]}="","",{EN["pdef"]}))'))
    S("combatN", (f'=({EN["spd"]}<>"")+({ref("eD")}<>"")+{ref("front")}*({EN["patk"]}<>"")'
                  f'+{ref("front")}*({EN["matk"]}<>"")'))
    # 推荐1 各区间职业的累计级数 / 是否首次出现（用于精通判断）
    bIDX = [f"${RK['IDX'][k]}${H0}" for k in range(5)]
    for k in range(5):
        S(f"bLv{k}", "=" + "+".join(f"{N[j]}*({bIDX[j]}={bIDX[k]})" for j in range(5)))
        S(f"bFirst{k}", "=AND(TRUE" + "".join(f",NOT(AND({N[j]}>0,{bIDX[j]}={bIDX[k]}))" for j in range(k)) + ")")

    # 向量：C0 / C4 / K / KT / 当前路线精通加成
    for j in range(NS):
        v, g = V9[j], gcol(j)
        H(f"{v}{H0}", f"={N[0]}*IF({si}={HZ},MAX(0,{SC[j]}$7),INDEX({g},{si}))")
        H(f"{v}{H0 + 1}", f"=IF({ref('tgtIdx')}=0,0,{N[4]}*INDEX({g},{ref('tgtIdx')}))")
        H(f"{v}{H0 + 2}", f"=$B$14+{v}{H0}+{v}{H0 + 1}")
        H(f"{v}{H0 + 3}", (f"={v}{H0 + 2}+INDEX({flc(0, j)},${RK['i1']}${H0})"
                           f"+INDEX({flc(1, j)},${RK['i2']}${H0})+INDEX({flc(2, j)},${RK['i3']}${H0})"))
        H(f"{v}{H0 + 4}", "=" + "+".join(f"${SEG['w']}${H0 + s}*INDEX({mbcol(j)},{Z(f'${SEGI}${H0 + s}')})"
                                         for s in range(NSEG)))
        H(f"{v}{H0 + 5}", f"=INDEX({bscol(j)},{si})")
    # 等级表 / 当前熟练度 / 可练 / 推荐1累计可练
    for i, rk in enumerate(RANKS):
        H(f"{V13[i]}{H0}", rk)
    for s in range(NK):
        v = f'IF(${WC[s]}${R_WCUR}<>"",${WC[s]}${R_WCUR},${WC[s]}${R_WINIT})'
        H(f"{V13[s]}{H0 + 1}", f'=IF({v}="×",0,IFERROR(MATCH({v},{RANKROW},0),2))')
        H(f"{V13[s]}{H0 + 2}", f'=IF({v}="×",0,1)')

    # ------------------------------------------------------------------
    # 职业辅助表（行5~154 对应职业表第3~152行，行155为全0空行）
    # ------------------------------------------------------------------
    elig_lo = [LO[1], LO[2], LO[3], LO[4]]
    for i in range(DATA_ROWS):
        r, cr = H0 + i, 3 + i
        nm, mn, av = f"${CLS['name']}{r}", f"${CLS['min']}{r}", f"${CLS['avail']}{r}"
        H(f"{CLS['name']}{r}", f'=IF({CLASS_Q}!$C{cr}="","",{CLASS_Q}!$C{cr})')
        H(f"{CLS['min']}{r}", f"=N({CLASS_Q}!$D{cr})")
        # 参与推荐：⑧勾选格（✓/靠后）+ 所属系列已勾选 + 性别/角色限定
        mark = f"${grid_cell(i, 'mark')[0]}${grid_cell(i, 'mark')[1:]}" if i < C_ROWS * C_GROUPS else '"✗"'
        H(f"{CLS['avail']}{r}", f'={mark}&""')
        H(f"{CLS['low']}{r}", f'=IF({av}="靠后",1,0)')
        feat = f"{CLASS_Q}!${C_FEAT}{cr}"
        rs_ = f"{CLASS_Q}!${C_RESTR}{cr}"
        series_ok = ",".join(f'NOT(AND(ISNUMBER(SEARCH("{sr}",{feat})),${SERIES_MARK_COL}${R_C0 + k}<>"✓"))'
                             for k, sr in enumerate(SERIES))
        g_ = ref("gender")
        H(f"{CLS['ok']}{r}", (f'=AND(OR({av}="✓",{av}="靠后"),{series_ok},'
                              f'OR({rs_}="",{rs_}={ref("charName")},AND({rs_}="女",OR({g_}="女",{g_}="自选")),'
                              f'AND({rs_}="男",OR({g_}="男",{g_}="自选"))))'))
        for j in range(NS):
            H(f"{CLS['BS'][j]}{r}", f'=IF({nm}="",0,N({CLASS_Q}!${C_BASE[j]}{cr}))')
        for j in range(NS):
            H(f"{CLS['G'][j]}{r}", f'=IF({nm}="",0,MAX(0,{SC[j]}$7+N({CLASS_Q}!{DC[j]}{cr})))')
            H(f"{CLS['MB'][j]}{r}", f'=IF({nm}="",0,IF({CLASS_Q}!${C_MBA}{cr}="{STATS[j]}",IFERROR(1*{CLASS_Q}!${C_MBV}{cr},0),0))')
        H(f"{CLS['MBANY']}{r}", f"=IF(SUMPRODUCT(ABS({CLS['MB'][0]}{r}:{CLS['MB'][-1]}{r}))>0,1,0)")
        H(f"{CLS['MBTXT']}{r}", f'=IF(${CLS["MBANY"]}{r}=1,{CLASS_Q}!${C_MBA}{cr}&"+"&{CLASS_Q}!${C_MBV}{cr},"")')
        for s in range(NK):
            key, L = SKILL_KEYS[s], len(SKILL_KEYS[s])
            H(f"{CLS['T'][s]}{r}", f'=IF({nm}="",0,IF(ISNUMBER(SEARCH("{key}",{CLASS_Q}!${C_USE}{cr})),1,0))')
            for tag, col in (("RM", C_MAIN), ("RS", C_SEL)):
                t = f"{CLASS_Q}!${col}{cr}"
                p = f'SEARCH("{key}",{t})'
                H(f"{CLS[tag][s]}{r}", (f'=IF({nm}="",0,IFERROR(MATCH(MID({t},{p}+{L},1)'
                                         f'&IF(MID({t},{p}+{L + 1},1)="+","+",""),{RANKROW},0),0))'))
            cur = f"{V13[s]}${H0 + 1}"
            H(f"{CLS['NM'][s]}{r}", f"=IF({CLS['RM'][s]}{r}>{cur},1,0)")
            H(f"{CLS['SU'][s]}{r}", f"=IF({CLS['RS'][s]}{r}>{cur},1,0)")
        rs_row = f"{CLS['RS'][0]}{r}:{CLS['RS'][-1]}{r}"
        H(f"{CLS['SA']}{r}", f'=IF(COUNTIF({rs_row},">0")>0,1,0)')
        H(f"{CLS['SS']}{r}", f"=IF(SUMPRODUCT(({rs_row}>0)*({rs_row}<={CURROW}))>0,1,0)")
        ok_avail = f"${CLS['ok']}{r}"
        for k in range(3):
            el = CLS["EL"][k]
            H(f"{el}{r}", (f'=AND({nm}<>"",IF({ref("force")}={k + 1},{nm}=$C$10,'
                           f'OR({nm}={CS},AND({ok_avail},{mn}<=MAX($B$9,{elig_lo[k]})))))'))
            H(f"{CLS['CN'][k]}{r}", f'=IF({el}{r},COUNTIF({el}${H0}:{el}{r},TRUE),"")')
        for k in range(4):
            el = CLS["ELU"][k]
            H(f"{el}{r}", f'=AND({nm}<>"",OR({nm}={CS},AND({ok_avail},{mn}<=MAX($B$9,{elig_lo[k]}))))')
            H(f"{CLS['CNU'][k]}{r}", f'=IF({el}{r},COUNTIF({el}${H0}:{el}{r},TRUE),"")')

    # 各窗口候选列表
    n_eff = [N[1], N[2], f'({N[3]}+IF($C$10="",{N[4]},0))']
    for k in range(3):
        cn = rng(CLS["CN"][k], H0, CLS_LAST)
        for i in range(DATA_ROWS):
            r, p = H0 + i, i + 1
            ix = f"{FL[k]['idx']}{r}"
            H(ix, f"=IFERROR(MATCH({p},{cn},0),{HZ})")
            H(f"{FL[k]['name']}{r}", f"=INDEX({NAMES},{ix})")
            for j in range(NS):
                H(f"{FL[k]['C'][j]}{r}", f"=INDEX({gcol(j)},{ix})*{n_eff[k]}")
    for k in range(4):
        cn = rng(CLS["CNU"][k], H0, CLS_LAST)
        for i in range(DATA_ROWS):
            H(f"{UL[k]}{H0 + i}", f"=IFERROR(MATCH({i + 1},{cn},0),{HZ})")

    # ------------------------------------------------------------------
    # 组合枚举（每个区间一个职业）
    # ------------------------------------------------------------------
    header(ws, 4, [k for k in CMB if k != "F"] + STATS, start_col=ws[f"{CMB['i1']}1"].column)
    m1, m2, m3, M = ref("m1"), ref("m2"), ref("m3"), ref("M")
    tg = ref("tgtIdx")
    for i in range(COMBO_ROWS):
        r = H0 + i
        ok = f"{i}<{M}"
        c = {k: f"{v}{r}" for k, v in CMB.items() if k != "F"}
        F = [f"{col}{r}" for col in CMB["F"]]
        H(c["i1"], f"=IF({ok},INT({i}/({m2}*{m3}))+1,1)")
        H(c["i2"], f"=IF({ok},MOD(INT({i}/{m3}),{m2})+1,1)")
        H(c["i3"], f"=IF({ok},MOD({i},{m3})+1,1)")
        H(c["id1"], f"=IF({N[1]}>0,INDEX({FLI[0]},{c['i1']}),{si})")
        H(c["id2"], f"=IF({N[2]}>0,INDEX({FLI[1]},{c['i2']}),{c['id1']})")
        H(c["id3"], f"=IF({N[3]}>0,INDEX({FLI[2]},{c['i3']}),{c['id2']})")
        H(c["id4"], f'=IF($C$10<>"",{ref("tgZ")},{c["id3"]})')
        ids = [si, c["id1"], c["id2"], c["id3"], c["id4"]]
        # 精通：该职业在路线中累计级数 ≥ Nm，且只在首次出现的区间计一次
        for k in range(5):
            lv = "+".join(f"{N[j]}*({ids[j]}={ids[k]})" for j in range(5))
            first = "".join(f",NOT(AND({N[j]}>0,{ids[j]}={ids[k]}))" for j in range(k))
            H(c[f"w{k}"], f"=IF(AND({ok},{N[k]}>0{first},{lv}>={Nm},INDEX({MBANY},{ids[k]})=1),1,0)")
        H(c["low"], "=IF(" + ok + "," + "+".join(f"({N[k]}>0)*INDEX({LOWC},{ids[k]})" for k in range(1, 5)) + ",0)")
        for j in range(NS):
            bonus = "+".join(f"{c[f'w{k}']}*INDEX({mbcol(j)},{ids[k]})" for k in range(5))
            H(F[j], (f"=IF({ok},{SC[j]}$9+INT(({V9[j]}${H0 + 2}+INDEX({flc(0, j)},{c['i1']})"
                     f"+INDEX({flc(1, j)},{c['i2']})+INDEX({flc(2, j)},{c['i3']}))/100)+{bonus}"
                     f"+{base_adj(j, c['id4'])},0)"))
        eval_cells(c, F, ok)
        H(c["score"], score(f"AND({ok},{ref('listOK')})", c["low"], c["met"], c["short"], c["total"],
                            c["combat"], F, i))

    # ------------------------------------------------------------------
    # 中途换职优化（基于推荐1，在某个区间内再换一次职业）
    # ------------------------------------------------------------------
    header(ws, 4, [k for k in REF if k != "F"] + STATS, start_col=ws[f"{REF['k']}1"].column)
    Rn = ref("R")
    sc_full = rng(REF["score"], H0, H0 + REF_ROWS - 1)
    for i in range(REF_ROWS):
        r = H0 + i
        ok = f"{i}<{Rn}"
        c = {k: f"{v}{r}" for k, v in REF.items() if k != "F"}
        F = [f"{col}{r}" for col in REF["F"]]
        k_ = c["k"]
        H(k_, f"=IF({ok},MATCH({i},{OROW},1),1)")
        H(c["j"], f"=IF({ok},INT(({i}-INDEX({OROW},{k_}))/INDEX({NNROW},{k_}))+1,1)")
        H(c["x"], f"=IF({ok},MOD({i}-INDEX({OROW},{k_}),INDEX({NNROW},{k_}))+1,1)")
        H(c["B"], f"=IF({ok},INDEX({ULMAT},{c['j']},{k_}),{HZ})")
        H(c["P"], f"=INDEX({PROW},{k_})")
        B, P, x = c["B"], c["P"], c["x"]
        hd = f"INDEX({HEADROW},{k_})"
        valid = c["valid"]
        H(valid, (f"=AND({ok},{ref('validT')},{B}<>{HZ},{B}<>{P},{x}>=$B$13,INDEX({NW},{k_})-{x}>=$B$13,"
                  f"OR({hd}=0,{ref('tgtMin')}<=INDEX({EW},{k_})+{x}))"))
        for k in range(5):
            H(c[f"w{k}"], (f"=IF(AND({valid},{N[k]}>0,{ref(f'bFirst{k}')},INDEX({MBANY},{bIDX[k]})=1,"
                           f"{ref(f'bLv{k}')}-{x}*({bIDX[k]}={P})+{x}*({bIDX[k]}={B})>={Nm}),1,0)"))
        inbase = "+".join(f"({N[k]}>0)*({bIDX[k]}={B})" for k in range(5))
        H(c["wB"], f"=IF(AND({valid},INDEX({MBANY},{B})=1,{x}>={Nm},({inbase})=0),1,0)")
        H(c["low"], f"=IF({valid},{ref('lowT')}+INDEX({LOWC},{B}),0)")
        # 最终职业：在最后一个区间「尾部」换职时变成 B，否则同推荐1
        H(c["fin"], f"=IF(AND({hd}=0,{k_}={ref('lastWin')}),{B},{bIDX[4]})")
        for j in range(NS):
            bonus = "+".join(f"{c[f'w{k}']}*INDEX({mbcol(j)},{bIDX[k]})" for k in range(5))
            bonus += f"+{c['wB']}*INDEX({mbcol(j)},{B})"
            H(F[j], (f"=IF({valid},{SC[j]}$9+INT(({V9[j]}${H0 + 3}+{x}*(INDEX({gcol(j)},{B})"
                     f"-INDEX({gcol(j)},{P})))/100)+{bonus}+{base_adj(j, c['fin'])},0)"))
        eval_cells(c, F, valid)
        H(c["score"], score(valid, c["low"], c["met"], c["short"], c["total"], c["combat"], F, i))
        # 同一区间同一职业（连续的 x 行）只保留换职等级最优的一行，避免 Top5 被同职业不同等级占满
        b0 = f"({i}-{x}+2)"
        b1 = f"MIN({REF_ROWS},{b0}+INDEX({NNROW},{k_})-1)"
        H(c["best"], (f"=IF({c['score']}>=MAX(INDEX({sc_full},{b0}):INDEX({sc_full},{b1})),"
                      f"{c['score']},-1E+12-{i})"))

    # ------------------------------------------------------------------
    # 排名：行5~9 推荐1~5，行10~14 优化1~5；每条路线展开为 6 段（起始等级 FROM / 职业序号 IDX）
    # ------------------------------------------------------------------
    rk_hdr = [k for k in RK if isinstance(RK[k], str)]
    header(ws, 4, rk_hdr, start_col=ws[f"{RK['pos']}1"].column)
    csc = rng(CMB["score"], H0, H0 + COMBO_ROWS - 1)
    rsc = rng(REF["best"], H0, H0 + REF_ROWS - 1)
    for rr in range(10):
        r = H0 + rr
        is_opt = rr >= 5
        rank = rr % 5 + 1
        c = {k: f"{v}{r}" for k, v in RK.items() if isinstance(v, str)}
        a = {k: f"${v}${r}" for k, v in RK.items() if isinstance(v, str)}
        src, sc_ = (REF, rsc) if is_opt else (CMB, csc)
        top = H0 + (REF_ROWS if is_opt else COMBO_ROWS) - 1
        H(c["pos"], f"=MATCH(LARGE({sc_},{rank}),{sc_},0)")
        H(c["valid"], f"=INDEX({sc_},{a['pos']})>-1E+10")
        for key in ("met", "short", "total", "low", "combat"):
            H(c[key], f"=INDEX({rng(src[key], H0, top)},{a['pos']})")
        for j in range(NS):
            H(f"{RK['F'][j]}{r}", f"=INDEX({rng(src['F'][j], H0, top)},{a['pos']})")
        FROM = [f"${v}${r}" for v in RK["FROM"]]
        IDX = [f"${v}${r}" for v in RK["IDX"]]
        wget = lambda key: f"INDEX({rng(src[key], H0, top)},{a['pos']})"
        if not is_opt:
            for key in ("i1", "i2", "i3"):
                H(c[key], f"=INDEX({rng(CMB[key], H0, top)},{a['pos']})")
            H(IDX[0][1:].replace("$", ""), f"={si}")
            for k in (1, 2, 3, 4):
                H(IDX[k][1:].replace("$", ""), f"=INDEX({rng(CMB[f'id{k}'], H0, top)},{a['pos']})")
            H(IDX[5][1:].replace("$", ""), f"={IDX[4]}")
            for k in range(5):
                H(FROM[k][1:].replace("$", ""), f"={E[k]}")
            H(FROM[5][1:].replace("$", ""), "=999")
            H(c["mtxt"], "=" + "&".join(f'IF({wget(f"w{k}")}=1,INDEX({MBTXT},{Z(IDX[k])})&" ","")' for k in range(5)))
        else:
            for key in ("k", "x", "B"):
                H(c[key], f"=INDEX({rng(REF[key], H0, top)},{a['pos']})")
            k_ = a["k"]
            H(c["P"], f"=INDEX({PROW},{k_})")
            H(c["head"], f"=INDEX({HEADROW},{k_})")
            H(c["split"], f"=INDEX({EW},{k_})+IF({a['head']}=1,{a['x']},INDEX({NW},{k_})-{a['x']})")
            H(c["first"], f"=IF({a['head']}=1,{a['B']},{a['P']})")
            H(c["second"], f"=IF({a['head']}=1,{a['P']},{a['B']})")
            bFROM = [f"${v}${H0}" for v in RK["FROM"]]
            bIDX6 = [f"${v}${H0}" for v in RK["IDX"]]
            for i in range(6):
                prev = max(i - 1, 0)
                H(IDX[i][1:].replace("$", ""),
                  f"=IF({i}<{k_},{bIDX6[i]},IF({i}={k_},{a['first']},IF({i}={k_}+1,{a['second']},{bIDX6[prev]})))")
                H(FROM[i][1:].replace("$", ""),
                  f"=IF({i}<{k_},{bFROM[i]},IF({i}={k_},{bFROM[min(i, 4)]},IF({i}={k_}+1,{a['split']},{bFROM[prev]})))")
            H(c["mtxt"], "=" + "&".join(f'IF({wget(f"w{k}")}=1,INDEX({MBTXT},{Z(bIDX[k])})&" ","")' for k in range(5))
              + f'&IF({wget("wB")}=1,INDEX({MBTXT},{Z(a["B"])})&" ","")')
        NAME = [f"${v}${r}" for v in RK["NAME"]]
        NN_ = [f"${v}${r}" for v in RK["N"]]
        for i in range(6):
            H(NAME[i][1:].replace("$", ""), f'=IF(OR({IDX[i]}=0,{IDX[i]}={HZ}),{CS if i == 0 else chr(34) * 2},INDEX({NAMES},{IDX[i]}))')
            nxt = FROM[i + 1] if i < 5 else "999"
            H(NN_[i][1:].replace("$", ""), f"=MAX(0,MIN($B$10,{nxt})-MAX($B$9,{FROM[i]}))")
        # 熟练度提示（只对展示的路线计算）：转入之前没待过的职业时，所需熟练度是否能在之前的职业练
        terms, u = [], trow(IDX[0])
        for j in range(1, 6):
            seen = "".join(f",NOT(AND({NN_[i]}>0,{IDX[i]}={IDX[j]}))" for i in range(j))
            terms.append(f"IF(AND({NN_[j]}>0,{IDX[j]}<>{si}{seen}),{badc(IDX[j], u)},0)")
            u = f"{u}+{trow(IDX[j])}"
        H(c["bad"], "=" + "+".join(terms))

    # ------------------------------------------------------------------
    # 可见区：③ 推荐 / ③+ 优化
    # ------------------------------------------------------------------
    def route_text(r):
        FROM = [f"${v}${r}" for v in RK["FROM"]]
        NAME = [f"${v}${r}" for v in RK["NAME"]]
        NN_ = [f"${v}${r}" for v in RK["N"]]
        parts = [f'IF({NN_[0]}>0,"Lv"&MAX($B$9,{FROM[0]})&" "&{NAME[0]}&" → ","")']
        for i in range(1, 6):
            parts.append(f'IF(AND({NN_[i]}>0,OR(MAX($B$9,{FROM[i]})=$B$9,{NAME[i]}<>{NAME[i - 1]})),'
                         f'"Lv"&MAX($B$9,{FROM[i]})&" "&{NAME[i]}&" → ","")')
        return "&".join(parts) + '&"Lv"&$B$10&" 完成"'

    cn_ = ref("combatN")
    for block, (rsec, rhdr, r0, title) in enumerate(
            ((R_BSEC, R_BHDR, R_B0, "③ 推荐路线 Top5（每个区间一个职业；排序：达成项 ＞ 实战评估 ＞ 不含靠后职业 ＞ 缺口小 ＞ 力/魔高 ＞ 属性总和；属性含精通加成；熟练度仅提示）"),
             (R_OSEC, R_OHDR, R_O0, "③+ 中途换职优化 Top5（在推荐1的某个区间中途再换一次职业，枚举换哪个职业、第几级换）"))):
        section(ws, rsec, title, LAST_VIS)
        header(ws, rhdr, ["方案", "达成项", "缺口合计"] + STATS + ["路线（LvX 职业 = 从该等级起在此职业升级）", "熟练度/精通", "实战评估 / 备注"])
        ws.merge_cells(f"O{rhdr}:{LAST_VIS}{rhdr}")
        for i in range(5):
            r, hr = r0 + i, H0 + block * 5 + i
            c = {k: f"${v}${hr}" for k, v in RK.items() if isinstance(v, str)}
            ws.cell(r, 1, ("优化" if block else "推荐") + str(i + 1))
            ws[f"B{r}"] = f'=IF({c["valid"]},{c["met"]}&"/"&{ref("tgtCnt")},"—")'
            ws[f"C{r}"] = f'=IF({c["valid"]},{c["short"]},"")'
            for j in range(NS):
                ws[f"{SC[j]}{r}"] = f'=IF({c["valid"]},${RK["F"][j]}${hr},"")'
            none_txt = ('"无可行组合"' if i == 0 else '"—"')
            ws[f"M{r}"] = f"=IF({c['valid']},{route_text(hr)},{none_txt})"
            ws[f"N{r}"] = (f'=IF({c["valid"]},IF({c["bad"]}=0,"✓ 熟练度可行","⚠ "&{c["bad"]}&"项熟练度需注意")'
                           f'&IF({c["low"]}>0,"·含靠后职业","")&IF({c["mtxt"]}<>""," · 精通:"&{c["mtxt"]},""),"")')
            ws.merge_cells(f"O{r}:{LAST_VIS}{r}")
            cmb_txt = f'IF({cn_}>0,"实战 "&{c["combat"]}&"/"&{cn_}&" 项通过","")'
            if block == 0:
                tail = '&"（基准路线，下方「优化」在此基础上中途换职）"' if i == 0 else ""
                ws[f"O{r}"] = f'=IF({c["valid"]},{cmb_txt}{tail},"")'
            else:
                b = {k: f"${RK[k]}${H0}" for k in ("met", "short", "total")}
                ws[f"O{r}"] = (f'=IF({c["valid"]},{cmb_txt}&"｜对比推荐1：达成"&TEXT({c["met"]}-{b["met"]},"+0;-0;0")'
                               f'&"，缺口"&TEXT({c["short"]}-{b["short"]},"+0;-0;0")'
                               f'&"，属性总和"&TEXT({c["total"]}-{b["total"]},"+0;-0;0"),"")')
            style_range(ws, f"A{r}:{LAST_VIS}{r}", align=center)
            for col in "MNO":
                ws[f"{col}{r}"].alignment = left_nowrap
            ws[f"A{r}"].font = f_bold

    # ------------------------------------------------------------------
    # 当前显示路线 / 手动路线：辅助 6 段
    # ------------------------------------------------------------------
    sel = ref("sel")
    rk_rng = lambda col: rng(col, H0, H0 + 9)
    for j in range(NSEG):
        r = H0 + j
        if j < 6:
            H(f"{SEG['from']}{r}", f"=IF({sel}=0,${MAN['from']}${r},INDEX({rk_rng(RK['FROM'][j])},{sel}))")
            H(f"{SEG['idx']}{r}", f"=IF({sel}=0,${MAN['idx']}${r},INDEX({rk_rng(RK['IDX'][j])},{sel}))")
            H(f"{SEG['name']}{r}", f"=IF({sel}=0,${MAN['name']}${r},INDEX({rk_rng(RK['NAME'][j])},{sel}))")
        else:
            H(f"{SEG['from']}{r}", f"=IF({sel}=0,${MAN['from']}${r},999)")
            H(f"{SEG['idx']}{r}", f"=IF({sel}=0,${MAN['idx']}${r},{HZ})")
            H(f"{SEG['name']}{r}", f'=IF({sel}=0,${MAN["name"]}${r},"")')
    for j in range(NSEG):
        r, mr = H0 + j, R_M0 + j
        if j == 0:
            H(f"{MAN['from']}{r}", "=$B$9")
            H(f"{MAN['name']}{r}", f"={CS}")
            H(f"{MAN['idx']}{r}", f"={si}")
        else:
            H(f"{MAN['from']}{r}", f'=IF($B{mr}="",999,$B{mr})')
            H(f"{MAN['name']}{r}", f'=IF($B{mr}="","",$C{mr})')
            H(f"{MAN['idx']}{r}", f"=IFERROR(MATCH({MAN['name']}{r},{NAMES},0),{HZ})")
    SL = H0 + NSEG - 1
    seg_idx = rng(SEG["idx"], H0, SL)
    seg_n = rng(SEG["n"], H0, SL)
    for G_ in (SEG, MAN):
        for j in range(NSEG):
            r = H0 + j
            nxt = f"{G_['from']}{r + 1}" if j < NSEG - 1 else "999"
            H(f"{G_['n']}{r}", f"=MAX(0,MIN($B$10,{nxt})-MAX($B$9,{G_['from']}{r}))")
            H(f"{G_['prev']}{r}", f"={si}" if j == 0 else f"={G_['idx']}{r - 1}")
            for s in range(NK):
                tcol = rng(CLS["T"][s], H0, CLS_LAST)
                if j == 0:
                    H(f"{G_['U'][s]}{r}", f"=INDEX({tcol},{Z(si)})")
                else:
                    H(f"{G_['U'][s]}{r}", f"={G_['U'][s]}{r - 1}+INDEX({tcol},{Z(G_['idx'] + str(r - 1))})")
            ix = Z(f"{G_['idx']}{r}")
            pieces, sel_pieces, scan_terms = [], [], []
            scan = f"{G_['scan']}{r}"
            for s in range(NK):
                can = f"({G_['U'][s]}{r}*{V13[s]}${H0 + 2}>0)"
                nm_ = rng(CLS["NM"][s], H0, CLS_LAST)
                su_ = rng(CLS["SU"][s], H0, CLS_LAST)
                rm_ = rng(CLS["RM"][s], H0, CLS_LAST)
                rs_ = rng(CLS["RS"][s], H0, CLS_LAST)
                pieces.append(f'IF(INDEX({nm_},{ix})=1,"{SKILLS[s]}"&INDEX({RANKROW},1,INDEX({rm_},{ix}))'
                              f'&IF({can},"(可练) ","(之前职业练不了) "),"")')
                # 选一：有能练的就只列能练的；一个都练不了才全部列出
                sel_pieces.append(f'IF(AND(INDEX({su_},{ix})=1,OR({scan}=FALSE,{can})),'
                                  f'"{SKILLS[s]}"&INDEX({RANKROW},1,INDEX({rs_},{ix}))'
                                  f'&IF({can},"(可练) ","(之前职业练不了) "),"")')
                scan_terms.append(f"AND(INDEX({su_},{ix})=1,{can})")
            H(scan, "=OR(" + ",".join(scan_terms) + ")")
            H(f"{G_['piece']}{r}", ("=" + "&".join(pieces)
                                    + f'&IF(AND(INDEX({SAC},{ix})=1,INDEX({SSC},{ix})=0),"｜选一："&'
                                    + "&".join(sel_pieces) + ',"")'))
    for j in range(NSEG):
        r = H0 + j
        idx_j, n_j = f"{SEG['idx']}{r}", f"{SEG['n']}{r}"
        H(f"{SEG['lv']}{r}", f"=SUMPRODUCT(({seg_idx}={idx_j})*{seg_n})")
        if j == 0:
            H(f"{SEG['first']}{r}", "=TRUE")
        else:
            H(f"{SEG['first']}{r}", (f"=SUMPRODUCT(({SEG['idx']}${H0}:{SEG['idx']}{r - 1}={idx_j})"
                                     f"*({SEG['n']}${H0}:{SEG['n']}{r - 1}>0))=0"))
        H(f"{SEG['w']}{r}", (f"=IF(AND({n_j}>0,{SEG['first']}{r},{SEG['lv']}{r}>={Nm},"
                             f"INDEX({MBANY},{Z(idx_j)})=1),1,0)"))

    def req_cells(r, G_, hr, n_ref):
        """N=所需熟练度，O=熟练度提示，P=转职条件/解锁条件"""
        ix = Z(f"${G_['idx']}${hr}")
        m = f"INDEX({class_col(C_MAIN)},{ix})"
        s_ = f"INDEX({class_col(C_SEL)},{ix})"
        m_none, s_none = f'OR({m}="",{m}="—")', f'OR({s_}="",{s_}="—")'
        idx = f"${G_['idx']}${hr}"
        ws[f"N{r}"] = (f'=IF({n_ref}=0,"",IF(OR({idx}=0,{idx}={HZ}),"",IF(AND({m_none},{s_none}),"无",'
                       f'IF({m_none},"",{m})&IF({s_none},"",IF({m_none},"","；")&"选一："&{s_}))))')
        piece = f"${G_['piece']}${hr}"
        ws[f"O{r}"] = (f'=IF({n_ref}=0,"",IF(OR({idx}=0,{idx}={HZ}),"⚠ 职业不在职业表",'
                       f'IF(OR({idx}=${G_["prev"]}${hr},{idx}={si}),"—（已在该职业）",'
                       f'IF({piece}="","✓ 满足","需练："&IF(LEFT({piece},1)="｜",MID({piece},2,999),{piece})))))')
        cond = f"INDEX({class_col(C_COND)},{ix})"
        ws[f"P{r}"] = f'=IF({n_ref}=0,"",IF(OR({idx}=0,{idx}={HZ}),"",IF(OR({cond}="--",{cond}=""),"无",{cond}&"")))'

    # ④ 当前显示路线
    section(ws, R_SSEC, "④ 当前显示路线：几级转什么职业、在哪个职业升几级、需要什么熟练度（由上方「显示路线」切换）", LAST_VIS)
    header(ws, R_SHDR, ["阶段", "起始等级", "职业", "结束等级", "升级次数"])
    ws.merge_cells(f"F{R_SHDR}:M{R_SHDR}")
    header(ws, R_SHDR, ["说明"], start_col=6)
    header(ws, R_SHDR, ["所需熟练度", "熟练度提示（可练=之前的职业能练）", "转职条件 / 解锁条件"], start_col=14)
    for j in range(NSEG):
        r, hr = R_S0 + j, H0 + j
        n = f"${SEG['n']}${hr}"
        ws.cell(r, 1, f"阶段{j + 1}")
        ws[f"B{r}"] = f'=IF({n}>0,MAX($B$9,${SEG["from"]}${hr}),"")'
        ws[f"C{r}"] = f'=IF({n}>0,${SEG["name"]}${hr},"")'
        ws[f"D{r}"] = f'=IF({n}>0,B{r}+{n},"")'
        ws[f"E{r}"] = f'=IF({n}>0,{n},"")'
        prev_name = CS if j == 0 else f"${SEG['name']}${hr - 1}"
        ws.merge_cells(f"F{r}:M{r}")
        ws[f"F{r}"] = (f'=IF({n}>0,"Lv"&B{r}&IF(B{r}=$B$9,IF(C{r}={CS}," 以【"&C{r}&"】起步"," 立即转职为【"&C{r}&"】"),'
                       f'IF(C{r}={prev_name}," 继续【"&C{r}&"】"," 转职为【"&C{r}&"】"))'
                       f'&"，在该职业升 "&{n}&" 级 → Lv"&D{r}'
                       f'&IF(${SEG["w"]}${hr}=1,"；累计练满"&{Nm}&"级可精通："&INDEX({MBTXT},{Z(f"${SEGI}${hr}")}),""),"")')
        req_cells(r, SEG, hr, n)
        style_range(ws, f"A{r}:{LAST_VIS}{r}", align=center)
        for col in "FNOP":
            ws[f"{col}{r}"].alignment = left
        ws[f"A{r}"].font = f_bold
        ws.row_dimensions[r].height = 30

    # 最终属性对比
    section(ws, R_FSEC, "最终属性对比（当前显示路线，目标等级时；「最终」= 面板成长结果〔含兵种基础值〕 + 精通加成）", LAST_VIS)
    header(ws, R_FHDR, ["项目", "", ""] + STATS)
    rows_f = ["面板结果", "精通加成", "最终", "目标", "差值", "判定"]
    for k, t in enumerate(rows_f):
        ws.cell(R_F0 + k, 1, t).font = f_bold
    rg, rb, rf, rt, rd, rj = (R_F0 + k for k in range(6))
    for j in range(NS):
        s = SC[j]
        ws[f"{s}{rg}"] = f'=IFERROR(INDEX({s}${R_T0}:{s}${R_TLAST},$B$10-$B$9+1),"")'
        ws[f"{s}{rb}"] = f'=IF({V9[j]}${H0 + 4}=0,"",{V9[j]}${H0 + 4})'
        ws[f"{s}{rf}"] = f'=IF({s}{rg}="","",{s}{rg}+N({s}{rb}))'
        ws[f"{s}{rt}"] = f'=IF({s}$10="","",{s}$10)'
        ws[f"{s}{rd}"] = f'=IF(OR({s}{rt}="",{s}{rf}=""),"",{s}{rf}-{s}{rt})'
        ws[f"{s}{rj}"] = (f'=IF({s}{rd}="","—",IF({s}{rd}>=0,IF({s}{rg}>={s}{rt},"✓","✓(含精通)"),'
                          f'"✗ 差"&-{s}{rd}))')
    ws.merge_cells(f"B{rj}:C{rj}")
    ws[f"B{rj}"] = f'=COUNTIF(D{rj}:L{rj},"✓*")&" / "&{ref("tgtCnt")}&" 项达成"'
    ws[f"B{rj}"].font = f_bold
    style_range(ws, f"A{R_F0}:L{rj}", align=center)
    for col in SC:
        ws[f"{col}{rf}"].font = f_bold

    # 实战评估（第一行：我方武器输入）
    section(ws, R_ESEC, "实战评估（当前显示路线的最终属性 vs 右上⑦敌方参考值；仅作参考，最终选择由玩家决定）", LAST_VIS)
    ws.cell(R_E0, 1, "我方武器").font = f_bold
    ws.cell(R_E0, 1).alignment = center
    ws[f"B{R_E0}"], ws[f"C{R_E0}"], ws[f"D{R_E0}"], ws[f"E{R_E0}"] = "威力", 5, "重量", 5
    for c_ in ("B", "D"):
        ws[f"{c_}{R_E0}"].font = f_bold
        ws[f"{c_}{R_E0}"].alignment = center
    for c_ in ("C", "E"):
        style_range(ws, f"{c_}{R_E0}:{c_}{R_E0}", font=f_input, fill=fill_input, align=center)
    fv = {j: f"${SC[j]}${rf}" for j in range(NS)}
    ws.merge_cells(f"F{R_E0}:{LAST_VIS}{R_E0}")
    ws[f"F{R_E0}"] = (f'="体格 = INT(力量/5) = "&IF({fv[I_STR]}="","—",INT({fv[I_STR]}/5))'
                      f'&"　｜攻速 = 速度 − MAX(0, 重量 − 体格) = "&IF({fv[I_SPD]}="","—",{aspd(fv)})'
                      f'&"　｜改威力/重量后评估结果立即重算（推荐排序也会随之更新）"')
    ws[f"F{R_E0}"].font = f_note
    ws[f"F{R_E0}"].alignment = left_nowrap
    t, fa, fr, eD = ref("typeMag"), ref("fast"), ref("front"), ref("eD")
    spd, es = aspd(fv), EN["spd"]
    atk = f"IF({t}=1,{fv[I_MAG]},{fv[I_STR]})"
    dmg = f"({atk}+N({EN['might']})-{eD})"
    can_double = f"AND({es}<>\"\",{spd}-{es}>=4)"
    eval_rows = [
        ("速度", (f'=IF(OR({es}="",{fv[I_SPD]}=""),"—（未填敌方速度）",IF({fa}=1,"【高速·追击】"&IF({spd}-{es}>=4,'
                f'"✓ 可追击：攻速"&{spd}&" vs 敌"&{es}&"（差+"&({spd}-{es})&"，≥4）",'
                f'"✗ 追击不足：攻速"&{spd}&" vs 敌"&{es}&"，还差"&(4-({spd}-{es}))&"点"),'
                f'"【低速·防追】"&IF({es}-{spd}<4,"✓ 不会被追击：攻速"&{spd}&" vs 敌"&{es},'
                f'"✗ 会被追击：攻速"&{spd}&" vs 敌"&{es}&"，还差"&({es}-{spd}-3)&"点")))')),
        ("攻击", (f'=IF(OR({eD}="",{fv[I_STR]}=""),"—（未填敌方防御）",IF({t}=1,"【魔法】魔力","【物理】力量")&{atk}'
                f'&"+威力"&N({EN["might"]})&" − 敌"&IF({t}=1,"魔防","防守")&{eD}&" = "&{dmg}&" → "'
                f'&IF({dmg}>0,"✓ 破甲，每击"&{dmg}&IF({can_double},"，可追击共"&2*{dmg},""),'
                f'"✗ 未破甲，还差"&(1-{dmg})&"点"))')),
        ("物防", (f'=IF({fr}=0,"—（非前排，不评估）",IF(OR({EN["patk"]}="",{fv[I_DEF]}=""),"—（未填敌方物理攻击）",'
                f'IF({fv[I_DEF]}>={EN["patk"]}-5,"✓ 防守"&{fv[I_DEF]}&" ≥ 敌物攻"&{EN["patk"]}&"−5，每击受伤"&MAX(0,{EN["patk"]}-{fv[I_DEF]}),'
                f'"✗ 防守"&{fv[I_DEF]}&"，距 敌物攻−5 还差"&({EN["patk"]}-5-{fv[I_DEF]})&"点")))')),
        ("魔防", (f'=IF({fr}=0,"—（非前排，不评估）",IF(OR({EN["matk"]}="",{fv[I_RES]}=""),"—（未填敌方魔法攻击）",'
                f'IF({fv[I_RES]}>={EN["matk"]}-5,"✓ 魔防"&{fv[I_RES]}&" ≥ 敌魔攻"&{EN["matk"]}&"−5，每击受伤"&MAX(0,{EN["matk"]}-{fv[I_RES]}),'
                f'"✗ 魔防"&{fv[I_RES]}&"，距 敌魔攻−5 还差"&({EN["matk"]}-5-{fv[I_RES]})&"点")))')),
    ]
    for k, (label, f) in enumerate(eval_rows):
        r = R_E0 + 1 + k
        ws.cell(r, 1, label).font = f_bold
        ws.merge_cells(f"B{r}:{LAST_VIS}{r}")
        ws[f"B{r}"] = f
        style_range(ws, f"A{r}:{LAST_VIS}{r}", align=left_nowrap)
        ws[f"A{r}"].alignment = center
    r = R_E0 + 5
    ws.cell(r, 1, "综合").font = f_bold
    ws.merge_cells(f"B{r}:{LAST_VIS}{r}")
    ws[f"B{r}"] = (f'=IF({cn_}=0,"未填敌方参考值",IF({fv[0]}="","","实战评估 "&({combat([fv[j] for j in range(NS)])})'
                   f'&" / "&{cn_}&" 项通过"))&"　｜定位："&IF({t}=1,"魔法","物理")&" · "&IF({fa}=1,"高速追击","低速防追")'
                   f'&" · "&IF({fr}=1,"前排","非前排")&"　｜力/魔越高越好，推荐排序已考虑以上各项"')
    style_range(ws, f"A{r}:{LAST_VIS}{r}", align=left_nowrap)
    ws[f"A{r}"].alignment = center
    ws[f"B{r}"].font = f_bold

    # ⑤ 手动路线
    section(ws, R_MSEC, "⑤ 手动路线（填了转职等级/职业后，「显示路线=自动」会直接显示这条路线，逐级表和评估同步联动）", LAST_VIS)
    header(ws, R_MHDR, ["段", "转职等级", "职业"])
    ws.merge_cells(f"D{R_MHDR}:M{R_MHDR}")
    header(ws, R_MHDR, ["说明"], start_col=4)
    header(ws, R_MHDR, ["所需熟练度", "熟练度提示（可练=之前的职业能练）", "转职条件 / 解锁条件"], start_col=14)
    for j in range(NSEG):
        r, hr = R_M0 + j, H0 + j
        ws.cell(r, 1, f"段{j + 1}").font = f_bold
        ws.merge_cells(f"D{r}:M{r}")
        if j == 0:
            ws[f"B{r}"] = "=$B$9"
            ws[f"C{r}"] = f"={CS}"
            ws[f"D{r}"] = (f'="起点（自动取上方起点等级/职业）　｜当前显示："&IF({sel}=0,"手动路线 ✓",'
                           f'"推荐/优化路线（手动路线未生效：清空下方段落或把「显示路线」改为自动/手动）")')
            style_range(ws, f"B{r}:C{r}", fill=fill_grey, align=center)
        else:
            style_range(ws, f"B{r}:C{r}", font=f_input, fill=fill_input, align=center)
            ws[f"D{r}"] = "在此等级转职为该职业；从上往下按等级递增填写，不用的段留空" if j == 1 else None
        req_cells(r, MAN, hr, f"${MAN['n']}${hr}")
        style_range(ws, f"A{r}:A{r}", align=center)
        style_range(ws, f"D{r}:M{r}", font=f_note, align=left_nowrap)
        style_range(ws, f"N{r}:{LAST_VIS}{r}", align=left)
        ws.row_dimensions[r].height = 30

    # ⑧ 自动推荐候选职业（勾选）
    section(ws, R_CSEC, "⑧ 自动推荐候选职业：✓=参与，靠后=参与但排后，✗/留空=不参与（起点/目标职业不受限制）；右侧可按系一键开关", LAST_VIS)
    for g in range(C_GROUPS):
        header(ws, R_CHDR, ["职业", "参与"], start_col=ws[f"{GRID_NAME_COLS[g]}1"].column)
        if GRID_NAME_COLS[g] in ("E", "H"):
            nxt = CL(ws[f"{GRID_NAME_COLS[g]}1"].column + 1)
            ws.merge_cells(f"{GRID_NAME_COLS[g]}{R_CHDR}:{nxt}{R_CHDR}")
            header(ws, R_CHDR, ["参与"], start_col=ws[f"{GRID_MARK_COLS[g]}1"].column)
    default_mark = {"是": "✓", "靠后": "靠后", "否": "✗"}
    for i in range(C_ROWS * C_GROUPS):
        nc, mc = grid_cell(i, "name"), grid_cell(i, "mark")
        col, row_ = nc[0], int(nc[1:])
        if col in ("E", "H"):
            ws.merge_cells(f"{nc}:{CL(ws[col + '1'].column + 1)}{row_}")
        ws[nc] = f'=IF({CLASS_Q}!$C{3 + i}="","",{CLASS_Q}!$B{3 + i}&"·"&{CLASS_Q}!$C{3 + i})'
        ws[nc].font = f_base
        ws[nc].alignment = left_nowrap
        ws[nc].border = border
        ws[mc] = default_mark.get(class_defaults[i][1], "✓") if i < len(class_defaults) else None
        style_range(ws, f"{mc}:{mc}", font=f_input, fill=fill_input, align=center)
    header(ws, R_CHDR, ["系列", "开关"], start_col=ws[f"{SERIES_NAME_COL}1"].column)
    ws.merge_cells(f"N{R_CHDR}:{LAST_VIS}{R_CHDR}")
    header(ws, R_CHDR, ["说明"], start_col=14)
    for k, sr in enumerate(SERIES):
        r = R_C0 + k
        ws[f"{SERIES_NAME_COL}{r}"] = sr
        ws[f"{SERIES_NAME_COL}{r}"].font = f_bold
        ws[f"{SERIES_NAME_COL}{r}"].alignment = center
        ws[f"{SERIES_NAME_COL}{r}"].border = border
        ws[f"{SERIES_MARK_COL}{r}"] = "✓"
        style_range(ws, f"{SERIES_MARK_COL}{r}:{SERIES_MARK_COL}{r}", font=f_input, fill=fill_input, align=center)
    notes_c = ["系列开关=✗ 时，该系所有职业都不参与（按职业表「兵种特性」判断）",
               "例：只想走猎兵系，就把其他系设为 ✗；想排除斗拳手→弓箭手这类跨系路线同理",
               "平民/贵族不属于任何系，只受单个勾选控制",
               "女性专用/角色限定职业会按所选角色自动排除",
               "表内职业名 = 阶级·兵种名，顺序同「职业」表"]
    for k, t in enumerate(notes_c):
        r = R_C0 + k
        ws.merge_cells(f"N{r}:{LAST_VIS}{r}")
        ws[f"N{r}"] = t
        ws[f"N{r}"].font = f_note
        ws[f"N{r}"].alignment = left_nowrap

    # ⑥ 逐级成长表
    section(ws, R_TSEC, "⑥ 逐级成长表（当前显示路线的面板值〔含当前兵种基础值，不含精通加成〕；绿色=该级属性+1，黄色行=转职）", LAST_VIS)
    header(ws, R_THDR, ["等级", "阶段", "职业"] + STATS + ["备注"])
    seg_from = rng(SEG["from"], H0, SL)
    for i in range(TL_ROWS):
        r = R_T0 + i
        ws[f"A{r}"] = f'=IF(OR($B$9="",$B$10=""),"",IF($B$9+{i}<=$B$10,$B$9+{i},""))'
        ws[f"B{r}"] = f'=IF($A{r}="","",MATCH($A{r},{seg_from},1))'
        ws[f"C{r}"] = f'=IF($A{r}="","",INDEX({rng(SEG["name"], H0, SL)},$B{r}))'
        tix = f"${TLH['idx']}{r}"
        H(tix[1:], f'=IF($A{r}="",{HZ},{Z(f"INDEX({seg_idx},$B{r})")})')
        for j in range(NS):
            cum = TLH["cum"][j]
            if i == 0:
                H(f"{cum}{r}", f'=IF($A{r}="","",0)')
            else:
                H(f"{cum}{r}", (f'=IF($A{r}="","",{cum}{r - 1}+IF(${TLH["idx"]}{r - 1}={HZ},MAX(0,{SC[j]}$7),'
                                f'INDEX({gcol(j)},${TLH["idx"]}{r - 1})))'))
            ws[f"{SC[j]}{r}"] = f'=IF($A{r}="","",{SC[j]}$9+INT(($B$14+{cum}{r})/100)+{base_adj(j, tix)})'
        if i == 0:
            ws[f"M{r}"] = (f'=IF($A{r}="","",IF($C{r}={CS},"起点：【"&$C{r}&"】",'
                           f'"起点：【"&{CS}&"】 ★ 立即转职 → 【"&$C{r}&"】"))')
        else:
            ws[f"M{r}"] = (f'=IF($A{r}="","",IF($C{r}<>$C{r - 1},"★ 转职 → 【"&$C{r}&"】","")'
                           f'&IF($A{r}=$B$10,"（目标等级）",""))')
        style_range(ws, f"A{r}:M{r}", align=center)
        ws[f"M{r}"].alignment = left_nowrap

    # ------------------------------------------------------------------
    # 警告
    # ------------------------------------------------------------------
    ws.merge_cells(f"A15:{LAST_VIS}15")
    tmin, elv = ref("tgtMin"), ref("entryLv")
    ws["A15"] = (
        '=IF(OR($B$9="",$B$10=""),"⚠ 请填写起点等级和目标等级；",IF($B$10<=$B$9,"⚠ 目标等级必须大于起点等级；",""))'
        f'&IF($B$10-$B$9>{TL_ROWS - 1},"⚠ 等级跨度超过{TL_ROWS - 1}，逐级表只显示前{TL_ROWS}级；","")'
        f'&IF({si}={HZ},"⚠ 起点职业「"&{CS}&"」不在职业表中（补正按0计），请在C9选择起点职业；","")'
        f'&IF(AND($C$10<>"",{ref("tgtIdx")}=0),"⚠ 目标职业不在职业表中；","")'
        f'&IF(AND({ref("tgtIdx")}>0,{elv}<{tmin}),"⚠ 目标职业最低转职等级为Lv"&{tmin}&"，但按目标等级只能在Lv"&{elv}&"转入；","")'
        f'&IF({M}>{COMBO_ROWS},"⚠ 候选组合"&{M}&"个，超过{COMBO_ROWS}，只搜索了前{COMBO_ROWS}个（可在⑧取消勾选用不到的职业/系列）；","")'
        f'&IF({Rn}>{REF_ROWS},"⚠ 中途换职候选"&{Rn}&"个，超过{REF_ROWS}，只搜索了前{REF_ROWS}个；","")'
        f'&IF(NOT({ref("listOK")}),"⚠ 某个阶段没有可选职业，请检查⑧候选职业勾选；","")'
        f'&IF(NOT({ref("validT")}),"⚠ 没有可行的推荐路线；","")'
        f'&IF(AND({sel}=0,{ref("manualBad")}),"⚠ 手动路线的转职等级需从上到下递增、中间不要空行；","")'
        f'&IF({ref("rankMissing")},"提示：该角色没有熟练度数据，按E计算，可在②填写当前熟练度；","")'
    )
    ws["A15"].font = f_warn
    ws["A15"].alignment = left_nowrap

    # ------------------------------------------------------------------
    # 条件格式 / 数据验证
    # ------------------------------------------------------------------
    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")
    yellow = PatternFill("solid", fgColor="FFEB9C")
    gfont, rfont, yfont = Font(color="006100", bold=True), Font(color="9C0006", bold=True), Font(color="9C5700", bold=True)
    cf = ws.conditional_formatting
    cf.add(f"D{R_T0 + 1}:L{R_TLAST}", FormulaRule(
        formula=[f"AND(ISNUMBER(D{R_T0 + 1}),ISNUMBER(D{R_T0}),D{R_T0 + 1}>D{R_T0})"], fill=green, font=gfont))
    cf.add(f"A{R_T0}:C{R_TLAST}", FormulaRule(formula=[f'LEFT($M{R_T0},1)="★"'], fill=fill_input))
    cf.add(f"M{R_T0}:M{R_TLAST}", FormulaRule(formula=[f'ISNUMBER(SEARCH("★",$M{R_T0}))'], fill=fill_input))
    cf.add(f"D{rj}:L{rj}", FormulaRule(formula=[f'D{rj}="✓(含精通)"'], fill=yellow, font=yfont))
    cf.add(f"D{rj}:L{rj}", FormulaRule(formula=[f'LEFT(D{rj},1)="✓"'], fill=green, font=gfont))
    cf.add(f"D{rj}:L{rj}", FormulaRule(formula=[f'LEFT(D{rj},1)="✗"'], fill=red, font=rfont))
    cf.add(f"D{rd}:L{rd}", CellIsRule(operator="lessThan", formula=["0"], font=rfont))
    cf.add(f"D{rb}:L{rb}", FormulaRule(formula=[f'ISNUMBER(D{rb})'], fill=yellow, font=yfont))
    er = f"B{R_E0 + 1}:B{R_E0 + 4}"
    cf.add(er, FormulaRule(formula=[f'ISNUMBER(SEARCH("✓",B{R_E0 + 1}))'], fill=green, font=gfont))
    cf.add(er, FormulaRule(formula=[f'ISNUMBER(SEARCH("✗",B{R_E0 + 1}))'], fill=yellow, font=yfont))
    for i in range(C_ROWS * C_GROUPS):
        mc = grid_cell(i, "mark")
        cf.add(mc, FormulaRule(formula=[f'{mc}="✓"'], fill=green, font=gfont))
        cf.add(mc, FormulaRule(formula=[f'OR({mc}="✗",{mc}="")'], fill=PatternFill("solid", fgColor="EDEDED"),
                               font=Font(color="7F7F7F")))
    for r0 in (R_B0, R_O0):
        cf.add(f"D{r0}:L{r0 + 4}", FormulaRule(
            formula=[f'AND(D$10<>"",ISNUMBER(D{r0}),D{r0}<D$10)'], fill=red, font=Font(color="9C0006")))
        cf.add(f"B{r0}:B{r0 + 4}", FormulaRule(formula=[f'B{r0}=({ref("tgtCnt")}&"/"&{ref("tgtCnt")})'], fill=green, font=gfont))
        cf.add(f"N{r0}:N{r0 + 4}", FormulaRule(formula=[f'LEFT(N{r0},1)="✓"'], fill=green, font=gfont))
        cf.add(f"N{r0}:N{r0 + 4}", FormulaRule(formula=[f'LEFT(N{r0},1)="⚠"'], fill=yellow, font=yfont))
    for r0 in (R_S0, R_M0):
        rr = f"O{r0}:O{r0 + NSEG - 1}"
        cf.add(rr, FormulaRule(formula=[f'LEFT(O{r0},2)="需练"'], fill=yellow, font=yfont))
        cf.add(rr, FormulaRule(formula=[f'LEFT(O{r0},1)="✓"'], fill=green, font=gfont))

    dv_char = DataValidation(type="list", formula1=CHAR_NAMES, allow_blank=True)
    dv_class = DataValidation(type="list", formula1=CLASS_NAMES, allow_blank=True)
    dv_route = DataValidation(type="list", formula1='"自动,推荐1,推荐2,推荐3,推荐4,推荐5,优化1,优化2,优化3,优化4,优化5,手动"')
    dv_lv = DataValidation(type="whole", operator="between", formula1="1", formula2="200", allow_blank=True)
    dv_acc = DataValidation(type="whole", operator="between", formula1="0", formula2="99")
    dv_rank = DataValidation(type="list", formula1='"' + ",".join(RANKS + ["×"]) + '"', allow_blank=True)
    dv_type = DataValidation(type="list", formula1='"自动,物理,魔法"')
    dv_speed = DataValidation(type="list", formula1='"自动,高速追击,低速防追"')
    dv_front = DataValidation(type="list", formula1='"自动,是,否"')
    dv_num = DataValidation(type="whole", operator="between", formula1="0", formula2="999", allow_blank=True)
    dv_yes = DataValidation(type="list", formula1='"是,否"')
    dv_mark = DataValidation(type="list", formula1='"✓,靠后,✗"', allow_blank=True)
    dv_series = DataValidation(type="list", formula1='"✓,✗"')
    for dv in (dv_char, dv_class, dv_route, dv_lv, dv_acc, dv_rank, dv_type, dv_speed, dv_front, dv_num,
               dv_yes, dv_mark, dv_series):
        ws.add_data_validation(dv)
    dv_char.add("B4")
    dv_class.add("C9:C10")
    dv_class.add(f"C{R_M0 + 1}:C{R_M0 + NSEG - 1}")
    dv_route.add("K14")
    dv_lv.add("B9:B10")
    dv_lv.add(f"B{R_M0 + 1}:B{R_M0 + NSEG - 1}")
    dv_lv.add("E14:H14")
    dv_lv.add("B13")
    dv_lv.add("I13")
    dv_acc.add("B14")
    dv_rank.add(f"B{R_WCUR}:M{R_WCUR}")
    dv_type.add("O9")
    dv_speed.add("O10")
    dv_front.add("O11")
    dv_num.add("O4:O8")
    dv_num.add(f"C{R_E0}")
    dv_num.add(f"E{R_E0}")
    dv_yes.add("O12")
    for i in range(C_ROWS * C_GROUPS):
        dv_mark.add(grid_cell(i, "mark"))
    dv_series.add(f"{SERIES_MARK_COL}{R_C0}:{SERIES_MARK_COL}{R_C0 + len(SERIES) - 1}")

    ws.freeze_panes = "A7"
    ws.column_dimensions.group("Q", CL(LAST_HELPER), hidden=True)
    for name, r_ in defnames.items():
        wb.defined_names[name] = DefinedName(name, attr_text=f"'{SH_MAIN}'!{r_}")
    return ws


def main():
    wb = Workbook()  # 默认第一个 sheet 作为主页面
    n_char, char_cols = build_char_sheet(wb)
    n_class, class_cols, class_defaults = build_class_sheet(wb)
    build_main(wb, char_cols, class_cols, class_defaults)
    wb.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print(f"saved {OUT}: {n_char} 角色, {n_class} 职业")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
医学统计图：KM 生存曲线、森林图、瀑布图、泳道图。

只需要 Python 3.8+、pandas、numpy、matplotlib、openpyxl（WorkBuddy 自带）。

用法（在命令行里运行，路径有空格或中文时加英文双引号）：
  python statfig.py doctor                         检查电脑环境和中文字体
  python statfig.py template KM 输出.xlsx           生成空白模板（KM / 森林图 / 瀑布图 / 泳道图）
  python statfig.py import KM 原表.xlsx 新文件.xlsx  把使用者的表格（xlsx/csv）按列名搬进模板（不改原表）
  python statfig.py set 新文件.xlsx 时间单位=月 终点名称=DFS   填写设置项
  python statfig.py report 新文件.xlsx 组别=安慰剂 例数=60 中位时间=13.5   KM：写入原文报告值
  python statfig.py check 数据.xlsx                 只检查，不画图
  python statfig.py draw 数据.xlsx                  检查通过后画图
  python statfig.py draw 数据.xlsx --confirm "使用者的原话"
                                                   有"需要确认"的问题，使用者回复后才能画

退出码：0 成功；2 有"必须修改"的问题；3 有"需要确认"的问题但没有 --confirm；1 其他错误。
"""

import sys
import os
import math
import re
import datetime
import warnings

warnings.filterwarnings("ignore")

try:
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except Exception:
    pass

VERSION = "1.0（2026-10-30）"
Z95 = 1.959963984540054


def _need(modname):
    try:
        return __import__(modname)
    except Exception:
        print(f"【环境问题】缺少 Python 库 {modname}。请告诉使用者：电脑环境缺少 {modname}，"
              f"需要先安装（在命令行运行 pip install {modname}），装好后再试。不要改用别的画图方法。")
        sys.exit(1)


np = _need("numpy")
pd = _need("pandas")
_need("openpyxl")
mpl = _need("matplotlib")
mpl.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager, transforms  # noqa: E402
from matplotlib.patches import Polygon, Patch  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402


# ---------------------------------------------------------------------------
# 通用：问题清单
# ---------------------------------------------------------------------------

class Issues:
    def __init__(self):
        self.errors = []    # 必须修改：数据或设置有硬伤，改好 Excel 才能画
        self.confirms = []  # 需要确认：可疑或不完整，使用者确认后才能画
        self.infos = []     # 提示：不影响画图

    def error(self, msg):
        self.errors.append(msg)

    def confirm(self, msg):
        self.confirms.append(msg)

    def info(self, msg):
        self.infos.append(msg)


def rows_str(idx_list, limit=8):
    """把 DataFrame 行号转成 Excel 行号（表头是第 1 行）。"""
    rows = [int(i) + 2 for i in idx_list]
    s = "、".join(str(r) for r in rows[:limit])
    if len(rows) > limit:
        return f"第 {s} 行等，共 {len(rows)} 行"
    return f"第 {s} 行"


# ---------------------------------------------------------------------------
# 读 Excel
# ---------------------------------------------------------------------------

TYPE_ALIASES = {
    "km": "km", "km曲线": "km", "km 曲线": "km", "生存曲线": "km", "kaplan-meier": "km", "km图": "km",
    "森林图": "forest", "forest": "forest", "亚组森林图": "forest", "亚组分析": "forest",
    "瀑布图": "waterfall", "waterfall": "waterfall",
    "泳道图": "swimmer", "swimmer": "swimmer", "游泳图": "swimmer",
}
TYPE_NAMES = {"km": "KM 曲线", "forest": "森林图", "waterfall": "瀑布图", "swimmer": "泳道图"}

PRIVACY_WORDS = ["姓名", "名字", "病历", "住院号", "门诊号", "身份证", "电话", "手机", "联系方式", "住址", "地址",
                 "name", "mrn", "phone", "address", "id card", "医保号", "床号"]


def norm_type(s):
    if s is None:
        return None
    k = str(s).strip().lower().replace("（", "(").replace("）", ")")
    return TYPE_ALIASES.get(k)


def cell_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return pd.Timestamp(v).strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def read_workbook(path):
    if not os.path.exists(path):
        print(f"【错误】找不到文件：{path}。请确认路径（路径里有空格或中文时，前后要加英文双引号）。")
        sys.exit(1)
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".xls"):
        print("【错误】只接受 Excel 文件（.xlsx）。请先用 template 命令生成模板，让使用者把数据填进去。")
        sys.exit(1)
    try:
        xl = pd.ExcelFile(path)
    except Exception as e:
        print(f"【错误】打不开这个 Excel：{e}。如果文件正在 Excel 里打开，请先关闭再试。")
        sys.exit(1)
    sheets = xl.sheet_names
    if "设置" not in sheets or "数据" not in sheets:
        print("【错误】这个 Excel 不是本工具的模板：必须有名为“设置”和“数据”的两个工作表。"
              "请用 template 命令生成模板，把数据复制进去。")
        sys.exit(1)
    st = pd.read_excel(xl, "设置", header=0, dtype=object)
    settings = {}
    for _, r in st.iterrows():
        if len(r) < 2:
            continue
        k = cell_str(r.iloc[0])
        if not k:
            continue
        settings[k] = cell_str(r.iloc[1])
    data = pd.read_excel(xl, "数据", header=0, dtype=object)
    data.columns = [cell_str(c) for c in data.columns]
    data = data.loc[:, [c for c in data.columns if c and not c.startswith("Unnamed")]]
    # 去掉全空行
    data = data.dropna(how="all")
    data = data[~data.apply(lambda r: all(cell_str(v) == "" for v in r), axis=1)]
    reported = None
    if "原文报告值" in sheets:
        reported = pd.read_excel(xl, "原文报告值", header=0, dtype=object)
        reported.columns = [cell_str(c) for c in reported.columns]
        reported = reported.dropna(how="all")
        reported = reported[~reported.apply(lambda r: all(cell_str(v) == "" for v in r), axis=1)]
    return settings, data, reported


def find_col(df, aliases):
    low = {c.lower().replace(" ", ""): c for c in df.columns}
    for a in aliases:
        k = a.lower().replace(" ", "")
        if k in low:
            return low[k]
    return None


def is_blank(v):
    return v is None or (isinstance(v, float) and math.isnan(v))


def to_num(v):
    s = cell_str(v)
    if s == "":
        return None
    s = s.replace("，", ".").replace(",", "").replace("％", "%").replace("%", "").replace("−", "-").replace("–", "-")
    try:
        return float(s)
    except ValueError:
        return "BAD"


def num_setting(settings, key, issues=None, default=None):
    s = settings.get(key, "")
    if s == "":
        return default
    v = to_num(s)
    if v == "BAD" or v is None:
        if issues is not None:
            issues.error(f"设置表“{key}”填的是“{s}”，这里需要一个数字。")
        return default
    return v


def yes(settings, key, default=False):
    s = settings.get(key, "").strip().lower()
    if s == "":
        return default
    return s in ("是", "要", "显示", "yes", "y", "true", "1", "需要")


def check_common(settings, data, issues, required):
    # 隐私
    bad_cols = [c for c in data.columns if any(w in c.lower() for w in PRIVACY_WORDS)]
    if bad_cols:
        issues.error("数据表里有可能识别患者身份的列：" + "、".join(bad_cols) +
                     "。请使用者删掉这些列（患者编号用 P01、P02 这样的代号即可）后再给我。我不会读取这些列的内容。")
    # 模板没替换
    for k, v in settings.items():
        if "示例" in v and "请替换" in v:
            issues.error(f"设置表“{k}”还是模板里的示例内容（{v}），请换成真实信息。")
    for key, why in required:
        if settings.get(key, "") == "":
            issues.error(f"设置表缺少“{key}”：{why}请问使用者后填进设置表。")
    if settings.get("数据截止日期", "") in ("不详", "不知道", "未知"):
        issues.confirm("数据截止日期写的是“不详”。核对单里会写“数据截止日期不详”，对外使用前最好补上。是否接受？")
    if settings.get("数据来源", "") and len(settings.get("数据来源", "")) < 4:
        issues.confirm(f"数据来源只写了“{settings.get('数据来源')}”，太简单，以后没法追溯。"
                       "请写清楚：哪份文件 / 哪个版本 / 第几页或哪张表（例如：XX 研究全文稿 2026-09-03，Table 2）。")
    if len(data) == 0:
        issues.error("数据表是空的，没有任何数据行。")


# ---------------------------------------------------------------------------
# 统计：KM、log-rank、Cox（Efron）
# 与 R survival 包 survfit / survdiff / coxph 的默认算法一致（置信区间方法可选）
# ---------------------------------------------------------------------------

def chi2_sf(x, df):
    """卡方分布上尾概率（不依赖 scipy）。"""
    if x <= 0:
        return 1.0
    try:
        from scipy.stats import chi2
        return float(chi2.sf(x, df))
    except Exception:
        pass
    a = df / 2.0
    xx = x / 2.0
    # 正则化不完全伽马函数 Q(a, x)
    if xx < a + 1:
        ap, s, d = a, 1.0 / a, 1.0 / a
        for _ in range(1000):
            ap += 1
            d *= xx / ap
            s += d
            if abs(d) < abs(s) * 1e-15:
                break
        p = s * math.exp(-xx + a * math.log(xx) - math.lgamma(a))
        return max(0.0, 1.0 - p)
    b = xx + 1 - a
    c = 1e300
    d = 1.0 / b
    h = d
    for i in range(1, 1000):
        an = -i * (i - a)
        b += 2
        d = an * d + b
        if abs(d) < 1e-300:
            d = 1e-300
        c = b + an / c
        if abs(c) < 1e-300:
            c = 1e-300
        d = 1.0 / d
        de = d * c
        h *= de
        if abs(de - 1) < 1e-15:
            break
    return h * math.exp(-xx + a * math.log(xx) - math.lgamma(a))


def norm_sf(z):
    return 0.5 * math.erfc(z / math.sqrt(2))


def km_fit(time, event, conf_type="log-log"):
    """返回逐时间点表：time, n_risk, n_event, n_censor, surv, se(累计风险的标准误), lower, upper。"""
    time = np.asarray(time, float)
    event = np.asarray(event, int)
    uniq = np.unique(time)
    out = []
    S, vs = 1.0, 0.0
    for u in uniq:
        m = time == u
        d = int(event[m].sum())
        c = int(m.sum()) - d
        nr = int((time >= u).sum())
        if d > 0:
            S *= (1.0 - d / nr)
            vs = vs + d / (nr * (nr - d)) if nr > d else float("inf")
        se = math.sqrt(vs)
        lo, up = km_ci(S, se, conf_type)
        out.append((u, nr, d, c, S, se, lo, up))
    return pd.DataFrame(out, columns=["time", "n_risk", "n_event", "n_censor", "surv", "se", "lower", "upper"])


def km_ci(S, se, conf_type):
    nan = float("nan")
    if conf_type == "log-log":
        if S <= 0:
            return 0.0, nan
        if S >= 1:
            return nan, 1.0
        if not np.isfinite(se):
            return 0.0, nan
        ll = math.log(-math.log(S))
        # R survfit conf.type="log-log"；log(S)<0
        lo =math.exp(-math.exp(ll - Z95 * se / math.log(S)))
        up = math.exp(-math.exp(ll + Z95 * se / math.log(S)))
        return lo, up
    # "log"（R survfit 默认）
    if S <= 0 or not np.isfinite(se):
        return nan, nan
    lo = math.exp(math.log(S) - Z95 * se)
    up = min(1.0, math.exp(math.log(S) + Z95 * se))
    return lo, up


def _quantile_from_curve(times, vals, p=0.5, tol=1e-9):
    """仿 R survival:::quantile.survfit：曲线首次 ≤ p 的时间；恰好等于 p 且有平台时取平台两端中点。"""
    pts = [(float(t), float(v)) for t, v in zip(times, vals) if v is not None and not np.isnan(v)]
    ded = []
    for t, v in pts:
        if ded and abs(ded[-1][1] - v) < 1e-12:
            continue
        ded.append((t, v))
    for i, (t, v) in enumerate(ded):
        if v <= p + tol:
            if abs(v - p) <= tol:
                if i + 1 < len(ded):
                    return (t + ded[i + 1][0]) / 2.0
                return float("nan")
            return t
    return float("nan")


def km_median(tab, p=0.5):
    med = _quantile_from_curve(tab["time"], tab["surv"], p)
    # 中位数的 CI（Brookmeyer-Crowley，同 R）：下置信曲线先跌破 0.5 → CI 下限；上置信曲线 → CI 上限
    lo = _quantile_from_curve(tab["time"], tab["lower"], p)
    up = _quantile_from_curve(tab["time"], tab["upper"], p)
    return med, lo, up


def km_at(tab, t):
    """t 时刻的生存率及 CI（阶梯函数，取 ≤t 的最后一个时间点）。超出最大随访时间返回 None。"""
    if t > tab["time"].max():
        return None
    sub = tab[tab["time"] <= t]
    if len(sub) == 0:
        return (1.0, float("nan"), float("nan"))
    r = sub.iloc[-1]
    return (float(r["surv"]), float(r["lower"]), float(r["upper"]))


def logrank(time, event, group):
    time = np.asarray(time, float)
    event = np.asarray(event, int)
    group = np.asarray(group)
    levels = list(dict.fromkeys(group.tolist()))
    k = len(levels)
    if k < 2:
        return None
    O = np.zeros(k)
    E = np.zeros(k)
    V = np.zeros((k, k))
    for u in np.unique(time[event == 1]):
        at = time >= u
        n = at.sum()
        dm = (time == u) & (event == 1)
        d = dm.sum()
        ng = np.array([(at & (group == g)).sum() for g in levels], float)
        dg = np.array([(dm & (group == g)).sum() for g in levels], float)
        O += dg
        E += d * ng / n
        if n > 1:
            f = d * (n - d) / (n - 1) / n
            V += f * (np.diag(ng) - np.outer(ng, ng) / n)
    U = (O - E)[:-1]
    Vr = V[:-1, :-1]
    try:
        chi = float(U @ np.linalg.solve(Vr, U))
    except np.linalg.LinAlgError:
        return None
    return {"levels": levels, "O": O, "E": E, "chisq": chi, "df": k - 1, "p": chi2_sf(chi, k - 1)}


def cox_efron(time, event, X):
    """Cox 比例风险模型，Efron 处理同时发生的事件（R coxph 默认）。返回 beta, se, loglik。"""
    time = np.asarray(time, float)
    event = np.asarray(event, int)
    X = np.asarray(X, float)
    if X.ndim == 1:
        X = X[:, None]
    n, p = X.shape
    etimes = np.unique(time[event == 1])
    risk_idx = [np.where(time >= u)[0] for u in etimes]
    death_idx = [np.where((time == u) & (event == 1))[0] for u in etimes]

    def ll_grad_hess(beta):
        eta = X @ beta
        w = np.exp(eta)
        ll = 0.0
        g = np.zeros(p)
        H = np.zeros((p, p))
        for R, D in zip(risk_idx, death_idx):
            m = len(D)
            s0r = w[R].sum()
            s1r = (w[R, None] * X[R]).sum(0)
            s2r = (w[R, None, None] * X[R][:, :, None] * X[R][:, None, :]).sum(0)
            s0d = w[D].sum()
            s1d = (w[D, None] * X[D]).sum(0)
            s2d = (w[D, None, None] * X[D][:, :, None] * X[D][:, None, :]).sum(0)
            ll += eta[D].sum()
            g += X[D].sum(0)
            for l in range(m):
                f = l / m
                s0 = s0r - f * s0d
                s1 = s1r - f * s1d
                s2 = s2r - f * s2d
                ll -= math.log(s0)
                g -= s1 / s0
                H -= s2 / s0 - np.outer(s1, s1) / s0 ** 2
        return ll, g, H

    beta = np.zeros(p)
    ll, g, H = ll_grad_hess(beta)
    for _ in range(100):
        try:
            step = np.linalg.solve(-H, g)
        except np.linalg.LinAlgError:
            return None
        nb = beta + step
        nll, ng, nH = ll_grad_hess(nb)
        halv = 0
        while nll < ll - 1e-10 and halv < 30:
            step /= 2
            nb = beta + step
            nll, ng, nH = ll_grad_hess(nb)
            halv += 1
        conv = abs(nll - ll) < 1e-10 * max(1.0, abs(ll))
        beta, ll, g, H = nb, nll, ng, nH
        if conv:
            break
    try:
        cov = np.linalg.inv(-H)
    except np.linalg.LinAlgError:
        return None
    se = np.sqrt(np.diag(cov))
    if not np.all(np.isfinite(se)) or np.any(np.abs(beta) > 20):
        return None
    return {"beta": beta, "se": se, "loglik": ll}


def reverse_km_median(time, event):
    tab = km_fit(time, 1 - np.asarray(event, int), "log-log")
    return km_median(tab)


# ---------------------------------------------------------------------------
# 格式化
# ---------------------------------------------------------------------------

def f_num(x, nd=1, nr="NR"):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return nr
    return f"{x:.{nd}f}"


def f_p(p):
    if p is None or np.isnan(p):
        return "NA"
    if p < 0.001:
        return "<0.001"
    if p < 0.01:
        return f"{p:.3f}"
    return f"{p:.3f}" if p < 0.1 else f"{p:.2f}"


def f_p_text(p, label="P"):
    s = f_p(p)
    return f"{label} < {s[1:]}" if s.startswith("<") else f"{label} = {s}"


def dash():
    return "–"


def decs(v):
    """原文数字的小数位数（Excel 里 18.0 读出来是 18，按同一组数字里最多的小数位补齐）。"""
    if not isinstance(v, float) or np.isnan(v):
        return 0
    t = repr(round(v, 6))
    return 0 if t.endswith(".0") or "." not in t else len(t.split(".")[1])


def fmt_rep(v, nd):
    if v is None or v == "NR":
        return "NR"
    return f"{v:.{nd}f}"


def pct(x):
    return "NA" if x is None or np.isnan(x) else f"{x * 100:.1f}"


def f_p_text_rep(p, text=None):
    """原文 P 值照原样写，不重新取舍。"""
    if text:
        return text
    if p < 0.001:
        return "P < 0.001"
    t = repr(round(p, 6))
    return f"P = {t}"


def rep_for(rep, lv, levels):
    if lv in rep:
        return rep[lv]
    if len(levels) == 1:
        return rep.get("全部患者") or rep.get("总体") or {}
    return {}


# ---------------------------------------------------------------------------
# 字体
# ---------------------------------------------------------------------------

def setup_fonts(verbose=False):
    """优先：Arial（英文）+ 微软雅黑（中文）。返回 (是否有中文字体, 字体说明)。"""
    candidates_cjk = [
        r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyh.ttf", r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\simsun.ttc", r"C:\Windows\Fonts\Deng.ttf",
        "/System/Library/Fonts/PingFang.ttc", "/System/Library/Fonts/STHeiti Medium.ttc",
        "/Library/Fonts/Arial Unicode.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    ]
    candidates_latin = [r"C:\Windows\Fonts\arial.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf",
                        "/Library/Fonts/Arial.ttf", "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf"]
    fams = []
    latin_name = None
    for f in candidates_latin:
        if os.path.exists(f):
            try:
                font_manager.fontManager.addfont(f)
                latin_name = font_manager.FontProperties(fname=f).get_name()
                fams.append(latin_name)
                break
            except Exception:
                pass
    cjk_name = None
    for f in candidates_cjk:
        if os.path.exists(f):
            try:
                font_manager.fontManager.addfont(f)
                cjk_name = font_manager.FontProperties(fname=f).get_name()
                fams.append(cjk_name)
                break
            except Exception:
                pass
    if latin_name is None:
        fams.insert(0, "DejaVu Sans")
    fams.append("DejaVu Sans")
    # 必须直接把字体列表写进 font.family，matplotlib 才会逐字回退（英文用 Arial，中文用雅黑）
    plt.rcParams["font.family"] = list(dict.fromkeys(fams))
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["pdf.fonttype"] = 42
    plt.rcParams["svg.fonttype"] = "none"
    desc = f"英文字体 {latin_name or 'DejaVu Sans（没有 Arial）'}；中文字体 {cjk_name or '没有找到'}"
    return cjk_name is not None, desc


def has_cjk(s):
    return bool(re.search(r"[\u4e00-\u9fff]", s or ""))


# ---------------------------------------------------------------------------
# KM 曲线
# ---------------------------------------------------------------------------

EVENT_YES = {"1", "1.0", "是", "发生", "事件", "有", "死亡", "复发", "进展", "yes", "y", "true", "event", "dead", "death"}
EVENT_NO = {"0", "0.0", "否", "删失", "无", "未发生", "存活", "无事件", "no", "n", "false", "censored", "censor", "alive"}

ENDPOINT_EN = {"OS": "Overall survival", "PFS": "Progression-free survival", "DFS": "Disease-free survival",
               "EFS": "Event-free survival", "RFS": "Recurrence-free survival", "MFS": "Metastasis-free survival",
               "CNS-PFS": "CNS progression-free survival", "DOR": "Duration of response",
               "IDFS": "Invasive disease-free survival", "TTF": "Time to treatment failure",
               "TTP": "Time to progression", "LM-PFS": "LM progression-free survival"}
ENDPOINT_CN = {"OS": "总生存率", "PFS": "无进展生存率", "DFS": "无病生存率", "EFS": "无事件生存率",
               "RFS": "无复发生存率", "CNS-PFS": "CNS 无进展生存率", "DOR": "持续缓解率"}
UNIT_EN = {"月": "months", "天": "days", "日": "days", "周": "weeks", "年": "years"}
UNIT_CN = {"月": "月", "天": "天", "日": "天", "周": "周", "年": "年"}


def parse_km(settings, data, reported, issues):
    check_common(settings, data, issues, [
        ("终点名称", "图的纵轴和标题要用（例如 OS、PFS、DFS）。"),
        ("时间单位", "必须知道“时间”这一列是天、周、月还是年，否则横轴会标错。"),
        ("数据来源", "核对单要写数据出自哪份文件、哪个版本。"),
        ("数据截止日期", "生存数据必须注明数据截止日期（不知道就填“不详”）。"),
    ])
    unit = settings.get("时间单位", "").strip()
    if unit and unit not in UNIT_EN:
        issues.error(f"时间单位填的是“{unit}”，只能填：月、天、周、年。")
    if settings.get("终点名称", "") and settings.get("终点定义", "") == "":
        issues.confirm(f"终点“{settings.get('终点名称')}”没有写定义（例如 DFS 从哪天算起、哪些算事件、死亡算不算事件）。"
                       "同一个名称在不同研究里定义可能不同。请确认定义，或确认不需要写。")
    ct = settings.get("置信区间方法", "").strip().lower()
    if ct in ("", "log-log", "loglog", "log log"):
        ct = "log-log"
    elif ct in ("log", "plain"):
        ct = ct if ct == "log" else "log"
    else:
        issues.error(f"置信区间方法填的是“{settings.get('置信区间方法')}”，只能填 log-log（默认，与 SAS 一致）或 log（R 默认）。")
        ct = "log-log"

    tcol = find_col(data, ["时间", "生存时间", "随访时间", "time", "survival time", "os_time", "pfs_time"])
    ecol = find_col(data, ["事件", "状态", "结局", "event", "status"])
    gcol = find_col(data, ["组别", "分组", "组", "治疗组", "arm", "group"])
    idcol = find_col(data, ["患者编号", "编号", "受试者编号", "id", "subject", "patient"])
    ccol = find_col(data, ["删失", "censor", "censored", "是否删失"])
    from_dates = False
    # 只有日期、没有时间：用日期算
    sdc = find_col(data, ["起始日期", "起点日期", "入组日期", "随机日期", "随机化日期", "手术日期", "开始日期", "治疗开始日期"])
    edc = find_col(data, ["结束日期", "事件或末次随访日期", "事件/末次随访日期", "终点日期", "末次随访日期"])
    if not tcol and sdc and edc:
        unit_ = settings.get("时间单位", "").strip()
        sd = pd.to_datetime(data[sdc], errors="coerce")
        ed = pd.to_datetime(data[edc], errors="coerce")
        bad_d = data.index[sd.isna() | ed.isna()]
        if len(bad_d):
            issues.error(f"“{sdc}”或“{edc}”有空白或不是日期：{rows_str(bad_d)}。失访患者要填最后一次有效随访日期，"
                         "不能空着，也不能用收到文件的日期代替。请使用者补上。")
        rev = data.index[(ed < sd).fillna(False)]
        if len(rev):
            issues.error(f"“{edc}”早于“{sdc}”：{rows_str(rev)}。")
        cut = pd.to_datetime(settings.get("数据截止日期", ""), errors="coerce")
        if pd.notna(cut):
            late = data.index[(ed > cut).fillna(False)]
            if len(late):
                issues.confirm(f"有 {len(late)} 个患者的“{edc}”晚于数据截止日期 {cut.date()}（{rows_str(late)}）。"
                               "通常应截到数据截止日。请确认这些患者怎么处理。")
        days = (ed - sd).dt.days
        div = {"月": 30.4375, "天": 1, "日": 1, "周": 7, "年": 365.25}.get(unit_, None)
        if div is None:
            issues.error("用日期计算时间时，必须在设置表“时间单位”填 月、天、周 或 年。")
        else:
            data = data.copy()
            data["时间"] = (days / div).round(6)
            tcol = "时间"
            from_dates = True
            issues.confirm(f"数据里没有“时间”列，我用“{edc}”减“{sdc}”算出时间"
                           + (f"（天数 ÷ 30.4375 换算成月）" if unit_ == "月" else "") +
                           f"。请确认：起点是“{sdc}”；“{edc}”对发生事件的患者是事件日期，对没发生事件的患者是末次随访日期。")
    # 分组可能有不死时间偏倚
    gcol0 = find_col(data, ["组别", "分组", "组", "治疗组", "arm", "group"])
    if gcol0:
        gtxt = "、".join(sorted(set(data[gcol0].map(cell_str))))
        if re.search(r"曾经|是否|接受|未接受|缓解|应答|responder|分流|后续|进展后|减量|停药|依从", gtxt + gcol0, re.I):
            issues.confirm(f"分组看起来是按治疗过程中才知道的情况划分的（组别：{gtxt[:40]}）。"
                           "例如“是否曾经分流”“是否缓解”：患者要活到那件事发生才会进入该组，这会让该组生存显得更好（不死时间偏倚）。"
                           "这类比较只能作探索性，不能写成因果结论。请确认分组方式，以及图的用途。")
    if ccol and not ecol:
        issues.error(f"数据里只有“{ccol}”列（删失标记），没有“事件”列。为避免把 0/1 弄反，请使用者改成“事件”列：1 = 发生事件，0 = 删失。")
    if not tcol:
        issues.error("数据表里找不到“时间”列（列名请写“时间”）。")
    if not ecol and not ccol:
        issues.error("数据表里找不到“事件”列（列名请写“事件”，1 = 发生事件，0 = 删失）。")
    if not tcol or not ecol:
        return None
    if not idcol:
        issues.info("数据里没有“患者编号”列，查不了重复患者。建议加上（用 P01、P02 这样的代号）。")
    else:
        ids = data[idcol].map(cell_str)
        dup = ids[(ids != "") & ids.duplicated(keep=False)]
        if len(dup):
            issues.error(f"患者编号有重复：{'、'.join(sorted(set(dup))[:8])}（{rows_str(dup.index)}）。KM 曲线每个患者只能有一行。")

    t = data[tcol].map(to_num)
    bad_t = t[t.map(lambda v: is_blank(v) or v == "BAD")]
    if len(bad_t) and not from_dates:
        issues.error(f"“时间”列有空白或不是数字：{rows_str(bad_t.index)}。请使用者补上或说明这些患者如何处理（不能由我猜）。")
    tnum = t.map(lambda v: v if isinstance(v, float) else np.nan).astype(float)
    neg = tnum[tnum < 0]
    if len(neg):
        issues.error(f"“时间”列有负数：{rows_str(neg.index)}。生存时间不可能是负数。")

    e_raw = data[ecol].map(cell_str).str.lower()
    ev = e_raw.map(lambda s: 1 if s in EVENT_YES else (0 if s in EVENT_NO else None))
    bad_e = ev[ev.isna()]
    if len(bad_e):
        vals = sorted(set(e_raw[bad_e.index]))
        issues.error(f"“事件”列有无法识别的值（{'、'.join(v if v else '空白' for v in vals[:6])}）：{rows_str(bad_e.index)}。"
                     "只能填 1（发生事件）或 0（删失）。")
    mapped_words = sorted(set(e_raw[(ev.notna()) & (~e_raw.isin(["0", "1", "0.0", "1.0"]))]))
    if mapped_words:
        issues.info(f"“事件”列里的文字已按下面理解：{'、'.join(mapped_words)} → "
                    + "、".join(f"{w}={'1' if w in EVENT_YES else '0'}" for w in mapped_words))
    mp = dict(x.split("←", 1) for x in settings.get("原始列对应", "").split("；") if "←" in x)
    if mp.get("事件") and mp["事件"] != "事件":
        issues.confirm(f"“事件”列来自原表的“{mp['事件']}”列。请确认这一列里 1 代表发生了{settings.get('终点名称') or '终点'}事件"
                       "（如死亡、复发、进展），0 代表删失（截至最后随访仍未发生）。如果原表是反过来的（1 = 删失），图会完全画错。")
    if mp.get("时间") and mp["时间"] != "时间" and not any(u in mp["时间"] for u in ["月", "天", "日", "周", "年", "month", "day", "week", "year"]):
        issues.confirm(f"“时间”列来自原表的“{mp['时间']}”列，列名里看不出单位。请确认单位是“{settings.get('时间单位') or '？'}”。")
    if ecol and any(w in ecol for w in ["死亡"]) and settings.get("终点名称", "").upper() in ("PFS", "DFS", "EFS", "RFS"):
        issues.confirm(f"终点是 {settings.get('终点名称')}，但事件列叫“{ecol}”。请确认这一列记录的是 {settings.get('终点名称')} 的事件（复发/进展或死亡），而不是只记录了死亡。")

    if gcol:
        g = data[gcol].map(cell_str)
        bad_g = g[g == ""]
        if len(bad_g):
            issues.error(f"“组别”列有空白：{rows_str(bad_g.index)}。")
    else:
        g = pd.Series(["全部患者"] * len(data), index=data.index)

    if issues.errors:
        return None

    df = pd.DataFrame({"time": tnum.values, "event": ev.astype(int).values, "group": g.values}, index=data.index)
    # 组别顺序
    order_s = settings.get("组别顺序", "")
    levels = list(dict.fromkeys(df["group"].tolist()))
    if order_s:
        want = [x.strip() for x in re.split(r"[,，、;；]", order_s) if x.strip()]
        missing = [x for x in want if x not in levels]
        extra = [x for x in levels if x not in want]
        if missing:
            issues.error(f"设置表“组别顺序”里的 {'、'.join(missing)} 在数据里找不到。数据里的组别是：{'、'.join(levels)}。")
        if extra:
            issues.error(f"数据里的组别 {'、'.join(extra)} 没有写进“组别顺序”。")
        if not missing and not extra:
            levels = want
    ref = settings.get("对照组", "").strip()
    if len(levels) >= 2:
        if not ref:
            issues.confirm(f"有 {len(levels)} 个组（{'、'.join(levels)}），但没有填“对照组”。HR 默认按“{levels[-1]}”作对照组计算"
                           f"（即 HR = 其他组 vs {levels[-1]}）。请确认哪个是对照组。")
            ref = levels[-1]
        elif ref not in levels:
            issues.error(f"设置表“对照组”填的是“{ref}”，数据里没有这个组。数据里的组别是：{'、'.join(levels)}。")
    if len(levels) > 4:
        issues.confirm(f"一共有 {len(levels)} 个组，曲线会很挤。请确认是否真的要画在一张图里。")

    # 合理性
    if unit:
        mx = df["time"].max()
        if unit == "月" and mx > 240:
            issues.confirm(f"时间单位是“月”，但最长时间是 {mx:g}（超过 20 年）。这一列会不会其实是“天”？")
        if unit in ("天", "日") and mx < 100:
            issues.confirm(f"时间单位是“天”，但最长时间只有 {mx:g} 天。这一列会不会其实是“月”？")
        if unit == "年" and mx > 30:
            issues.confirm(f"时间单位是“年”，但最长时间是 {mx:g} 年，不太合理，请确认单位。")
        if unit == "周" and mx > 1000:
            issues.confirm(f"时间单位是“周”，但最长时间是 {mx:g} 周，请确认单位。")
    z = df[(df["time"] == 0)]
    if len(z):
        issues.confirm(f"有 {len(z)} 个患者时间为 0（{rows_str(z.index)}），其中 {int(z['event'].sum())} 个记为事件。"
                       "时间为 0 通常是录入错误或起点定义问题，请确认。")
    nonint = df["time"][(df["time"] % 1 != 0)]
    if unit in ("天", "日") and len(nonint):
        issues.confirm(f"时间单位是“天”，但有 {len(nonint)} 个时间带小数（{rows_str(nonint.index)}）。请确认单位。")
    for lv in levels:
        sub = df[df["group"] == lv]
        if len(sub) < 10:
            issues.confirm(f"组“{lv}”只有 {len(sub)} 例，曲线和中位数会很不稳定。确认要画吗？")
        if len(sub) and sub["event"].sum() == 0:
            issues.info(f"组“{lv}”没有任何事件，曲线会一直是 100%，中位数为“未达到”。")
        if len(sub):
            mx_t = sub["time"].max()
            last = sub[sub["time"] == mx_t]
            if last["event"].sum() == len(last) and (sub["time"] >= mx_t).sum() <= 3:
                issues.info(f"组“{lv}”时间最长的患者（{mx_t:g}）发生了事件，曲线最后会直接掉到 0%。统计上没错，但尾部只剩极少数患者，"
                            "这一段很不稳定；需要的话可以用“横轴最大值”把横轴截短（不改变计算）。")
        if len(sub) and sub["event"].sum() == len(sub):
            issues.confirm(f"组“{lv}”所有患者都是事件、没有删失。临床研究里少见，请确认事件列没有填反（1 = 发生事件，0 = 删失）。")
        # 尾部：最后一个时间点只有一个人
    if df["event"].mean() > 0.95 and len(df) >= 20:
        issues.confirm("几乎所有患者都记为事件。请确认“事件”列 1/0 没有填反。")

    # 关注时间点
    tps = []
    for x in re.split(r"[,，、;；\s]+", settings.get("关注时间点", "")):
        if x.strip():
            v = to_num(x)
            if v == "BAD" or v is None:
                issues.error(f"“关注时间点”里的“{x}”不是数字。")
            else:
                tps.append(v)

    # 原文报告值
    rep = {}
    if reported is not None and len(reported):
        gc = find_col(reported, ["组别", "分组", "组"])
        for _, r in reported.iterrows():
            name = cell_str(r[gc]) if gc else "全部患者"
            if not name:
                continue
            rec = {}
            for key, als in [("n", ["例数", "n"]), ("events", ["事件数", "events"]),
                             ("median", ["中位时间", "中位数", "median"]),
                             ("med_lo", ["中位下限", "中位95%CI下限", "下限"]),
                             ("med_hi", ["中位上限", "中位95%CI上限", "上限"])]:
                c = find_col(reported, als)
                if c is not None:
                    v = to_num(r[c])
                    if v == "BAD":
                        s = cell_str(r[c]).upper()
                        if s in ("NR", "NE", "未达到", "NA", "未达"):
                            v = "NR"
                        else:
                            issues.error(f"原文报告值表里“{name}”的“{c}”填的是“{s}”，需要数字或 NR。")
                            v = None
                    rec[key] = v
            rep[name] = rec
        for name in rep:
            if name not in levels and not (name in ("全部患者", "总体") and len(levels) == 1):
                issues.error(f"原文报告值表里的组别“{name}”在数据里找不到（数据里的组别：{'、'.join(levels)}）。")
    praw = settings.get("原文P值", "").replace("＜", "<").replace(" ", "")
    if praw.startswith("<"):
        settings = dict(settings)
        settings["原文P值"] = ""
    rep_hr = {k: num_setting(settings, k, issues) for k in ["原文HR", "原文HR下限", "原文HR上限", "原文P值"]}
    if praw.startswith("<"):
        v = to_num(praw[1:])
        if isinstance(v, float):
            rep_hr["原文P值"] = v * 0.999
            rep_hr["原文P值文本"] = "P < " + praw[1:]
        else:
            issues.error(f"原文P值填的是“{praw}”，看不懂。请填数字，例如 0.012 或 <0.001。")
    if rep_hr["原文HR"] is not None:
        lo, hi = rep_hr["原文HR下限"], rep_hr["原文HR上限"]
        if lo is not None and hi is not None and not (lo <= rep_hr["原文HR"] <= hi):
            issues.error(f"原文 HR {rep_hr['原文HR']:g} 不在它的 95% CI（{lo:g}–{hi:g}）之内，肯定有一个数字抄错了。")
        if settings.get("原文HR是否分层", "") == "":
            issues.confirm("填了原文 HR，但没有说明原文 HR 是否为分层 Cox 模型（论文常用分层 HR）。"
                           "本工具算的是不分层 HR，两者可能略有不同。请在设置表“原文HR是否分层”填 是 / 否 / 不知道。")
    if rep_hr["原文P值"] is not None and not (0 <= rep_hr["原文P值"] <= 1):
        issues.error(f"原文 P 值填的是 {rep_hr['原文P值']:g}，P 值必须在 0 到 1 之间。")

    src = settings.get("图中统计量", "").strip() or "自动"
    if src not in ("自动", "原文", "计算", "不显示"):
        issues.error(f"“图中统计量”填的是“{src}”，只能填：自动、原文、计算、不显示。")
    return {"stat_src": src, "df": df, "levels": levels, "ref": ref, "unit": unit, "conf_type": ct, "timepoints": tps,
            "reported": rep, "rep_hr": rep_hr, "idcol": idcol}


def km_compute(P, issues):
    df, levels, ref = P["df"], P["levels"], P["ref"]
    res = {"groups": {}}
    for lv in levels:
        sub = df[df["group"] == lv]
        tab = km_fit(sub["time"].values, sub["event"].values, P["conf_type"])
        med = km_median(tab)
        fu = reverse_km_median(sub["time"].values, sub["event"].values)
        rates = {tp: km_at(tab, tp) for tp in P["timepoints"]}
        res["groups"][lv] = {"tab": tab, "n": len(sub), "events": int(sub["event"].sum()),
                             "median": med, "followup": fu, "rates": rates,
                             "tmax": float(sub["time"].max())}
    for lv in levels:
        for tp, r in res["groups"][lv]["rates"].items():
            if r is None:
                issues.confirm(f"关注时间点 {tp:g} {P['unit']} 超过了组“{lv}”的最长随访时间，算不出这个时间点的生存率，会写“无法估计”。")
            else:
                at = int((df[(df['group'] == lv)]['time'] >= tp).sum())
                if at < 10:
                    issues.confirm(f"组“{lv}”在 {tp:g} {P['unit']} 时只剩 {at} 人仍在随访（风险人数），这个时间点的生存率很不稳定。")
    if len(levels) >= 2:
        lr = logrank(df["time"].values, df["event"].values, df["group"].values)
        res["logrank"] = lr
        others = [lv for lv in levels if lv != ref]
        X = np.column_stack([(df["group"].values == lv).astype(float) for lv in others])
        cx = cox_efron(df["time"].values, df["event"].values, X)
        res["cox"] = {}
        if cx is None:
            issues.confirm("HR 算不出来（可能某组没有事件）。图上不显示 HR。")
        else:
            for i, lv in enumerate(others):
                b, s = cx["beta"][i], cx["se"][i]
                res["cox"][lv] = {"hr": math.exp(b), "lo": math.exp(b - Z95 * s), "hi": math.exp(b + Z95 * s),
                                  "p": 2 * norm_sf(abs(b / s))}
    # 与原文比较
    rep = P["reported"]
    for name, rec in rep.items():
        key = name if name in res["groups"] else (levels[0] if len(levels) == 1 else None)
        if key is None:
            continue
        G = res["groups"][key]
        if isinstance(rec.get("n"), float) and int(rec["n"]) != G["n"]:
            issues.confirm(f"组“{name}”：原文例数 {int(rec['n'])}，数据里是 {G['n']} 例，对不上。"
                           "可能是数据版本不同、分析人群不同（ITT / 安全性分析集）或有遗漏。请确认用哪个。")
        if isinstance(rec.get("events"), float) and int(rec["events"]) != G["events"]:
            issues.confirm(f"组“{name}”：原文事件数 {int(rec['events'])}，数据里是 {G['events']}，对不上。请确认数据版本和事件定义。")
        med = G["median"][0]
        rm = rec.get("median")
        if rm == "NR" and not np.isnan(med):
            issues.confirm(f"组“{name}”：原文中位数“未达到”，数据算出来是 {med:.2f} {P['unit']}。请确认。")
        elif isinstance(rm, float):
            if np.isnan(med):
                issues.confirm(f"组“{name}”：原文中位数 {rm:g}，数据算出来“未达到”。请确认数据版本。")
            elif abs(rm - med) > max(0.06, 0.01 * abs(rm)):
                issues.confirm(f"组“{name}”：原文中位数 {rm:g}，数据算出来 {med:.2f} {P['unit']}，对不上。"
                               "可能是数据版本、时间单位换算（天→月）或事件定义不同。请确认。")
        for k, idx, lab in [("med_lo", 1, "中位数 95% CI 下限"), ("med_hi", 2, "中位数 95% CI 上限")]:
            rv = rec.get(k)
            cv = G["median"][idx]
            if isinstance(rv, float) and not np.isnan(cv) and abs(rv - cv) > max(0.06, 0.01 * abs(rv)):
                issues.confirm(f"组“{name}”：原文{lab} {rv:g}，算出来 {cv:.2f}。置信区间算法可能不同"
                               f"（本次用 {P['conf_type']}；SAS 默认 log-log，R 默认 log）。请确认。")
    rh = P["rep_hr"]
    if rh["原文HR"] is not None and res.get("cox"):
        others = list(res["cox"].keys())
        if len(others) == 1:
            c = res["cox"][others[0]]
            if abs(c["hr"] - rh["原文HR"]) > 0.015:
                issues.confirm(f"原文 HR {rh['原文HR']:g}，数据算出来 {c['hr']:.2f}（{others[0]} vs {ref}）。"
                               "差别可能来自：分层因素、HR 方向（谁比谁）、数据版本。请确认图上用哪个。")
            if abs(math.log(rh["原文HR"])) > 0.05 and math.log(rh["原文HR"]) * math.log(c["hr"]) < 0:
                issues.confirm(f"原文 HR 和算出来的 HR 方向相反（一个 <1，一个 >1）。很可能对照组设反了：现在对照组是“{ref}”。")
    if rh["原文P值"] is not None and res.get("logrank"):
        p = res["logrank"]["p"]
        if (p < 0.05) != (rh["原文P值"] < 0.05):
            issues.confirm(f"原文 P 值 {rh['原文P值']:g}，算出来的 log-rank P = {f_p(p)}，一个显著一个不显著。请确认。")
    res["stat_mode"] = decide_stat_mode(P, issues)
    return res


def decide_stat_mode(P, issues):
    """一张图里不混用原文值和计算值。"""
    levels, rep, rh, src = P["levels"], P["reported"], P["rep_hr"], P["stat_src"]
    have, miss = [], []
    for lv in levels:
        (have if rep_for(rep, lv, levels).get("median") is not None else miss).append(f"“{lv}”中位数")
    if len(levels) == 2:
        (have if rh["原文HR"] is not None else miss).append("HR")
    if len(levels) >= 2:
        (have if rh["原文P值"] is not None else miss).append("P 值")
    if src == "不显示":
        return "不显示"
    if src == "计算":
        if have:
            issues.info("设置为“计算”：图上的中位数、HR、P 值都用本工具算的值，原文值只用来核对。")
        return "计算"
    if src == "自动" and not have:
        issues.info("没有填原文报告值：图上的中位数、HR、P 值都是本工具算的（核对单里会标“AI 计算，待统计师复核”）。")
        return "计算"
    if miss:
        issues.confirm("原文报告值不全（缺：" + "、".join(miss) + "）。为了不在一张图里混用原文数字和计算数字，"
                       "图上只显示原文有的数字，缺的不显示。如果想全部改用计算值，请在设置表“图中统计量”填“计算”。请确认。")
    if P["timepoints"]:
        issues.info("图上用原文值时，“关注时间点”的生存率（计算值）不画在图上，只写在核对单里。")
    return "原文"


def nice_step(xmax, unit):
    if unit == "月":
        cands = [1, 2, 3, 6, 12, 24]
    elif unit in ("天", "日"):
        cands = [7, 14, 30, 60, 90, 180, 365]
    elif unit == "周":
        cands = [1, 2, 4, 8, 12, 26, 52]
    else:
        cands = [0.5, 1, 2, 5]
    for c in cands:
        if xmax / c <= 10:
            return c
    return cands[-1]


PALETTE = ["#1F4E79", "#C00000", "#548235", "#ED7D31", "#7030A0", "#808080"]


def parse_colors(settings, n):
    s = settings.get("颜色", "")
    cols = [c.strip() for c in re.split(r"[,，、;；\s]+", s) if c.strip()]
    out = []
    for c in cols:
        if re.fullmatch(r"#?[0-9a-fA-F]{6}", c):
            out.append(c if c.startswith("#") else "#" + c)
    return (out + PALETTE)[:max(n, 1)] if out else PALETTE[:max(n, 1)]


def draw_km(settings, P, R, outbase):
    lang_en = settings.get("图中文字", "英文").strip() not in ("中文", "中")
    unit = P["unit"]
    ep = settings.get("终点名称", "").strip()
    levels = P["levels"]
    colors = parse_colors(settings, len(levels))
    disp = {lv: ("All patients" if (lv == "全部患者" and lang_en) else lv) for lv in levels}
    lab_s = settings.get("图例名称", "")
    if lab_s:
        labs = [x.strip() for x in re.split(r"[,，;；]", lab_s) if x.strip()]
        if len(labs) == len(levels):
            disp = dict(zip(levels, labs))
    xmax = num_setting(settings, "横轴最大值") or max(R["groups"][lv]["tmax"] for lv in levels)
    step = num_setting(settings, "横轴间隔") or nice_step(xmax, unit)
    ticks = np.arange(0, xmax + step * 0.001, step)
    if ticks[-1] < xmax - 1e-9:
        ticks = np.append(ticks, ticks[-1] + step)
    show_table = yes(settings, "显示风险表", True)
    show_ci = yes(settings, "显示置信区间阴影", False)
    show_med = yes(settings, "显示中位线", False)
    show_censor = yes(settings, "显示删失标记", True)

    nrow = len(levels)
    fig_h = 5.2 + (0.32 * nrow + 0.55 if show_table else 0)
    fig = plt.figure(figsize=(7.6, fig_h), dpi=100)
    if show_table:
        gs = fig.add_gridspec(2, 1, height_ratios=[5.2, 0.32 * nrow + 0.4], hspace=0.28)
        ax = fig.add_subplot(gs[0])
        axt = fig.add_subplot(gs[1], sharex=ax)
    else:
        ax = fig.add_subplot(111)
        axt = None

    for lv, col in zip(levels, colors):
        G = R["groups"][lv]
        tab = G["tab"]
        xs = [0.0] + tab["time"].tolist()
        ys = [1.0] + tab["surv"].tolist()
        ax.step(xs, [y * 100 for y in ys], where="post", color=col, lw=1.8, label=disp[lv])
        if show_ci:
            lo = [1.0] + tab["lower"].ffill().fillna(1.0).tolist()
            up = [1.0] + tab["upper"].ffill().fillna(1.0).tolist()
            ax.fill_between(xs, [v * 100 for v in lo], [v * 100 for v in up], step="post", color=col, alpha=0.12, lw=0)
        if show_censor:
            cz = tab[tab["n_censor"] > 0]
            ax.plot(cz["time"], cz["surv"] * 100, ls="none", marker="+", ms=7, mew=1.2, color=col)
        if show_med and not np.isnan(G["median"][0]):
            m = G["median"][0]
            ax.plot([0, m], [50, 50], ls="--", lw=0.8, color="#7F7F7F")
            ax.plot([m, m], [0, 50], ls="--", lw=0.8, color=col)

    ax.set_xlim(0, ticks[-1] * 1.02)
    ax.set_ylim(0, 103)
    ax.set_yticks([0, 20, 40, 60, 80, 100])
    ax.set_xticks(ticks)
    ax.set_xticklabels([f"{t:g}" for t in ticks])
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ytitle = settings.get("纵轴标题", "")
    if not ytitle:
        key = ep.upper()
        ytitle = (ENDPOINT_EN.get(key, f"{ep} probability") + " (%)") if lang_en else (ENDPOINT_CN.get(key, f"{ep} 率") + "（%）")
    xtitle = settings.get("横轴标题", "") or (f"Time ({UNIT_EN.get(unit, unit)})" if lang_en else f"时间（{UNIT_CN.get(unit, unit)}）")
    ax.set_ylabel(ytitle, fontsize=11)
    if axt is None:
        ax.set_xlabel(xtitle, fontsize=11)
    ax.tick_params(labelsize=10)
    title = settings.get("图标题", "")
    if title:
        ax.set_title(title, fontsize=12, loc="left", fontweight="bold")

    # 统计量文字：一张图里要么全用原文值，要么全用计算值，不混用
    mode = R.get("stat_mode", "计算")
    lines = []
    used = {}  # 图中出现的数字及来源
    rep = P["reported"]
    rh = P["rep_hr"]
    for lv in levels:
        G = R["groups"][lv]
        used[f"{lv} 事件数/例数（仅核对单）"] = (f"{G['events']}/{G['n']}", "AI 由数据计数")
    if mode != "不显示":
        unit_w = UNIT_EN.get(unit, unit) if lang_en else UNIT_CN.get(unit, unit)
        med_lines = []
        for lv, col in zip(levels, colors):
            G = R["groups"][lv]
            nm = disp[lv]
            if mode == "原文":
                rr = rep_for(rep, lv, levels)
                if rr.get("median") is None:
                    continue
                vals = [rr.get("median"), rr.get("med_lo"), rr.get("med_hi")]
                nd = max([decs(v) for v in vals if isinstance(v, float)] + [0])
                med_s = fmt_rep(vals[0], nd)
                if vals[1] is not None or vals[2] is not None:
                    med_s += f" ({fmt_rep(vals[1], nd)}{dash()}{fmt_rep(vals[2], nd)})"
                used[f"{lv} 中位数"] = (med_s, "原文报告值")
            else:
                m, lo, hi = G["median"]
                med_s = "NR" if (np.isnan(m) and np.isnan(lo)) else f"{f_num(m)} ({f_num(lo)}{dash()}{f_num(hi)})"
                used[f"{lv} 中位数"] = (med_s, f"AI 计算（{P['conf_type']} CI），待统计师复核")
            med_lines.append((f"{nm}: {med_s}" if lang_en else f"{nm}：{med_s}", col))
        if med_lines:
            lines.append((f"Median {ep}, {unit_w} (95% CI)" if lang_en else f"中位 {ep}，{unit_w}（95% CI）", None))
            lines.extend(med_lines)
        if mode == "计算":
            for tp in P["timepoints"]:
                parts = []
                for lv in levels:
                    r = R["groups"][lv]["rates"].get(tp)
                    v = "NE" if r is None else f"{r[0] * 100:.1f}%"
                    parts.append(f"{disp[lv]} {v}" if len(levels) > 1 else v)
                    used[f"{lv} {tp:g} {unit} 生存率"] = (v if r is None else f"{r[0] * 100:.1f}%（95% CI {pct(r[1])}–{pct(r[2])}）",
                                                         "AI 计算，待统计师复核")
                uw = {"月": "month", "天": "day", "日": "day", "周": "week", "年": "year"}.get(unit, unit)
                lab = f"{tp:g}-{uw} {ep} rate: " if lang_en else f"{tp:g} {UNIT_CN.get(unit, unit)} {ep} 率："
                lines.append((lab + ("; ".join(parts) if lang_en else "；".join(parts)), None))
        if len(levels) == 2:
            other = [lv for lv in levels if lv != P["ref"]][0]
            hr_s = p_s = None
            if mode == "原文":
                if rh["原文HR"] is not None:
                    vals = [rh["原文HR"], rh["原文HR下限"], rh["原文HR上限"]]
                    nd = max([decs(v) for v in vals if v is not None] + [2])
                    hr_s = f"HR {vals[0]:.{nd}f}"
                    if vals[1] is not None and vals[2] is not None:
                        hr_s += f" (95% CI {vals[1]:.{nd}f}{dash()}{vals[2]:.{nd}f})"
                    used[f"HR（{other} vs {P['ref']}）"] = (hr_s, "原文报告值")
                if rh["原文P值"] is not None:
                    p_s = f_p_text_rep(rh["原文P值"], rh.get("原文P值文本"))
                    used["P 值"] = (p_s, "原文报告值")
            elif R.get("cox") and other in R["cox"]:
                c = R["cox"][other]
                hr_s = f"HR {c['hr']:.2f} (95% CI {c['lo']:.2f}{dash()}{c['hi']:.2f})"
                used[f"HR（{other} vs {P['ref']}）"] = (hr_s, "AI 计算（Cox 模型，Efron，不分层），待统计师复核")
                p_s = f_p_text(R["logrank"]["p"], "log-rank P")
                used["P 值"] = (p_s, "AI 计算（log-rank 检验），待统计师复核")
            hp = [x for x in (hr_s, p_s) if x]
            if hp:
                lines.append(("; ".join(hp) if lang_en else "；".join(hp), None))
        elif len(levels) > 2:
            if mode == "原文" and rh["原文P值"] is not None:
                p_s = f_p_text_rep(rh["原文P值"], rh.get("原文P值文本"))
                used["P 值（多组）"] = (p_s, "原文报告值")
                lines.append((p_s, None))
            elif mode == "计算" and R.get("logrank"):
                p_s = f_p_text(R["logrank"]["p"], "log-rank P")
                used["P 值（多组）"] = (p_s, "AI 计算（log-rank 检验），待统计师复核")
                lines.append((p_s, None))

    min_s = min(float(R["groups"][lv]["tab"]["surv"].min()) for lv in levels)
    low_place = min_s > 0.55  # 曲线都在上半部分：统计量放左下；否则放右上
    if len(levels) > 1:
        ax.legend(loc="upper right" if (low_place and not lines) else "lower left" if not low_place else "center right",
                  frameon=False, fontsize=10, handlelength=2.2)
    if lines:
        txt = "\n".join(l for l, _ in lines)
        if low_place:
            ax.text(0.03, 0.04, txt, transform=ax.transAxes, fontsize=9, va="bottom", ha="left", linespacing=1.45,
                    bbox=dict(boxstyle="round,pad=0.4", fc="white", ec="#BFBFBF", lw=0.6))
        else:
            ax.text(0.97, 0.96, txt, transform=ax.transAxes, fontsize=9, va="top", ha="right", linespacing=1.45,
                    bbox=dict(boxstyle="round,pad=0.4", fc="white", ec="#BFBFBF", lw=0.6))

    if axt is not None:
        ax.set_xlabel(xtitle, fontsize=11)
        axt.set_ylim(-0.5, nrow - 0.5)
        axt.invert_yaxis()
        axt.axis("off")
        lab_tr = transforms.offset_copy(transforms.blended_transform_factory(axt.transAxes, axt.transData),
                                        fig=fig, x=-16, y=0, units="points")
        for i, (lv, col) in enumerate(zip(levels, colors)):
            sub = P["df"][P["df"]["group"] == lv]["time"].values
            for tk in ticks:
                n_at = int((sub >= tk).sum())
                axt.text(tk, i, str(n_at), ha="center", va="center", fontsize=9, color=col)
            axt.text(0, i, disp[lv], transform=lab_tr, ha="right", va="center", fontsize=9, color=col)
        axt.text(0, -0.75, "No. at risk" if lang_en else "风险人数", transform=lab_tr, ha="right", va="bottom",
                 fontsize=9.5, fontweight="bold")
    foot = settings.get("图注", "")
    if foot:
        fig.text(0.01, 0.005, foot, fontsize=8, color="#404040", ha="left", va="bottom", wrap=True)
    return fig, used


def km_summary_lines(P, R):
    L = []
    for lv in P["levels"]:
        G = R["groups"][lv]
        m, lo, hi = G["median"]
        fm = G["followup"][0]
        L.append(f"- 组“{lv}”：{G['n']} 例，事件 {G['events']} 例；中位 {f_num(m, 2)}（95% CI {f_num(lo, 2)}–{f_num(hi, 2)}）{P['unit']}；"
                 f"中位随访（反向 KM）{f_num(fm, 2)} {P['unit']}；最长时间 {G['tmax']:g}")
        for tp, r in G["rates"].items():
            if r is None:
                L.append(f"  - {tp:g} {P['unit']} 生存率：无法估计（超过最长随访）")
            else:
                L.append(f"  - {tp:g} {P['unit']} 生存率：{r[0] * 100:.1f}%（95% CI {f_num(r[1] * 100 if not np.isnan(r[1]) else float('nan'), 1, 'NA')}–{f_num(r[2] * 100 if not np.isnan(r[2]) else float('nan'), 1, 'NA')}）")
    if R.get("logrank"):
        L.append(f"- log-rank 检验：卡方 {R['logrank']['chisq']:.3f}，自由度 {R['logrank']['df']}，P = {R['logrank']['p']:.4g}")
    for lv, c in (R.get("cox") or {}).items():
        L.append(f"- HR（{lv} vs {P['ref']}，Cox 模型，Efron，不分层）：{c['hr']:.3f}（95% CI {c['lo']:.3f}–{c['hi']:.3f}），Wald P = {c['p']:.4g}")
    return L


# ---------------------------------------------------------------------------
# 森林图
# ---------------------------------------------------------------------------

def parse_forest(settings, data, issues):
    check_common(settings, data, issues, [
        ("终点名称", "森林图标题和核对单要写清是哪个终点（例如 PFS）。"),
        ("试验组名称", "横轴下方要写“有利于哪一组”。"),
        ("对照组名称", "横轴下方要写“有利于哪一组”。"),
        ("数据来源", "核对单要写数据出自哪份文件、哪个版本、哪张图表。"),
    ])
    measure = settings.get("效应量", "").strip().upper() or "HR"
    if measure not in ("HR", "OR", "RR"):
        issues.error(f"效应量填的是“{settings.get('效应量')}”，只能填 HR、OR 或 RR。")
    left_good = settings.get("小于1有利于", "").strip()
    if left_good == "":
        if measure == "HR":
            issues.info("没填“小于1有利于”，按常规：HR < 1 有利于试验组（事件是复发、进展、死亡这类坏事时成立）。")
            left_good = "试验组"
        else:
            issues.confirm(f"效应量是 {measure}，但没有说明 {measure} < 1 有利于哪一组。"
                           f"如果终点是缓解率这类“好事”，{measure} > 1 才有利于试验组。请确认后在设置表“小于1有利于”填 试验组 或 对照组。")
            left_good = "试验组"
    elif left_good not in ("试验组", "对照组"):
        issues.error("“小于1有利于”只能填 试验组 或 对照组。")
    if settings.get("HR方向", "") == "" and measure == "HR":
        issues.confirm(f"请确认 HR 的方向：HR = {settings.get('试验组名称') or '试验组'} vs {settings.get('对照组名称') or '对照组'}"
                       "（试验组比对照组）。如果原文是反过来的，图会完全读反。确认后在设置表“HR方向”填“试验组 vs 对照组”。")

    sc = find_col(data, ["亚组", "亚组变量", "subgroup", "变量"])
    lc = find_col(data, ["分层", "分类", "类别", "水平", "level", "category"])
    hc = find_col(data, ["HR", "OR", "RR", "效应量", "estimate"])
    loc_ = find_col(data, ["下限", "95%CI下限", "CI下限", "lower", "lcl"])
    upc = find_col(data, ["上限", "95%CI上限", "CI上限", "upper", "ucl"])
    pc = find_col(data, ["交互P", "交互作用P", "交互P值", "p interaction", "p for interaction", "pinteraction"])
    ne1 = find_col(data, ["试验组事件数"])
    nn1 = find_col(data, ["试验组例数"])
    ne0 = find_col(data, ["对照组事件数"])
    nn0 = find_col(data, ["对照组例数"])
    ntot = find_col(data, ["例数", "总例数", "n"])
    for c, nm in [(sc, "亚组"), (hc, "HR"), (loc_, "下限"), (upc, "上限")]:
        if c is None:
            issues.error(f"数据表里找不到“{nm}”列。")
    if issues.errors:
        return None
    rows = []
    for i, r in data.iterrows():
        sub = cell_str(r[sc])
        lev = cell_str(r[lc]) if lc else ""
        if not sub:
            issues.error(f"第 {i + 2} 行“亚组”是空的。")
            continue
        rec = {"row": i + 2, "sub": sub, "lev": lev}
        for key, col in [("hr", hc), ("lo", loc_), ("hi", upc), ("p", pc), ("e1", ne1), ("n1", nn1),
                         ("e0", ne0), ("n0", nn0), ("n", ntot)]:
            if col is None:
                rec[key] = None
                continue
            v = to_num(r[col])
            if v == "BAD":
                s = cell_str(r[col])
                if key == "p" and re.match(r"^[<＜]\s*0?\.\d+$", s):
                    rec[key] = float(re.sub(r"[<＜\s]", "", s)) / 2
                    rec["p_text"] = s.replace("＜", "<")
                    continue
                if key in ("hr", "lo", "hi") and s.upper() in ("NE", "NA", "NR", "不可估计"):
                    rec[key] = None
                    rec["ne"] = True
                    continue
                issues.error(f"第 {i + 2} 行“{col}”填的是“{s}”，需要数字。")
                v = None
            rec[key] = v
        rows.append(rec)
    if issues.errors:
        return None
    label = f"{sub} {lev}".strip()
    seen = set()
    for r in rows:
        label = f"{r['sub']} / {r['lev']}" if r["lev"] else r["sub"]
        if label in seen:
            issues.error(f"第 {r['row']} 行“{label}”重复出现了。")
        seen.add(label)
        if r.get("ne"):
            issues.info(f"第 {r['row']} 行“{label}”的效应量为 NE（不可估计），图上只写 NE，不画点。")
            continue
        hr, lo, hi = r["hr"], r["lo"], r["hi"]
        if hr is None or lo is None or hi is None:
            issues.error(f"第 {r['row']} 行“{label}”的 {measure}、下限、上限有空白。")
            continue
        if hr <= 0 or lo <= 0 or hi <= 0:
            issues.error(f"第 {r['row']} 行“{label}”：{measure} 和置信区间必须大于 0（现在是 {hr:g}、{lo:g}、{hi:g}）。"
                         "如果这是对数值（log HR），请使用者换成原始 HR。")
            continue
        if lo > hi:
            issues.error(f"第 {r['row']} 行“{label}”：下限 {lo:g} 比上限 {hi:g} 还大，可能两列填反了。")
            continue
        if not (lo <= hr <= hi):
            issues.error(f"第 {r['row']} 行“{label}”：{measure} {hr:g} 不在 95% CI（{lo:g}–{hi:g}）里，一定有数字抄错了，请对照原文改正。")
            continue
        if lo < hr < hi:
            mid = (math.log(lo) + math.log(hi)) / 2
            width = math.log(hi) - math.log(lo)
            if width > 0 and abs(math.log(hr) - mid) > 0.12 * width + 0.02:
                issues.confirm(f"第 {r['row']} 行“{label}”：{measure} {hr:g}（{lo:g}–{hi:g}）在对数尺度上明显不在置信区间中间，"
                               "通常意味着某个数字抄错了（或原文用的不是 Wald 区间）。请对照原文确认。")
        if r["p"] is not None and not (0 <= r["p"] <= 1):
            issues.error(f"第 {r['row']} 行“{label}”：交互 P 值 {r['p']:g} 不在 0–1 之间。")
        for e, n, who in [("e1", "n1", "试验组"), ("e0", "n0", "对照组")]:
            if r[e] is not None and r[n] is not None and r[e] > r[n]:
                issues.error(f"第 {r['row']} 行“{label}”：{who}事件数 {r[e]:g} 大于例数 {r[n]:g}。")
            for k in (e, n):
                if r[k] is not None and r[k] % 1 != 0:
                    issues.error(f"第 {r['row']} 行“{label}”：{who}的人数 {r[k]:g} 不是整数。")
        if r["hi"] / r["lo"] > 50:
            issues.confirm(f"第 {r['row']} 行“{label}”：置信区间非常宽（{lo:g}–{hi:g}），通常是例数或事件很少。确认照画吗？")
    # 总体行
    overall_names = ("总体", "全部", "全部患者", "overall", "all", "all patients", "itt")
    ov = [r for r in rows if r["sub"].strip().lower() in overall_names]
    if not ov:
        issues.confirm("没有“总体”这一行（亚组填“总体”）。森林图一般要把总体结果放在最上面作参照。确认不要总体行吗？")
    else:
        o = ov[0]
        # 各亚组例数之和与总体比较
        tot_key = None
        for k in ("n", "n1"):
            if o.get(k) is not None:
                tot_key = k
                break
        if tot_key:
            groups = {}
            for r in rows:
                if r is o or r[tot_key] is None:
                    continue
                groups.setdefault(r["sub"], []).append(r)
            for sname, rs in groups.items():
                if len(rs) < 2:
                    continue
                s = sum(r[tot_key] for r in rs)
                if tot_key == "n1" and o.get("n0") is not None:
                    s0 = sum((r["n0"] or 0) for r in rs)
                    if s > o["n1"] or s0 > o["n0"]:
                        issues.error(f"亚组“{sname}”各层例数加起来（试验组 {s:g}，对照组 {s0:g}）超过了总体（{o['n1']:g}，{o['n0']:g}），一定有错。")
                    elif s < o["n1"] or s0 < o["n0"]:
                        issues.info(f"亚组“{sname}”各层例数加起来（试验组 {s:g}，对照组 {s0:g}）少于总体（{o['n1']:g}，{o['n0']:g}），"
                                    "可能有缺失或未知类别没列出。核对单里会写明。")
                else:
                    if s > o[tot_key]:
                        issues.error(f"亚组“{sname}”各层例数加起来 {s:g} 超过了总体 {o[tot_key]:g}，一定有错。")
                    elif s < o[tot_key]:
                        issues.info(f"亚组“{sname}”各层例数加起来 {s:g} 少于总体 {o[tot_key]:g}，可能有缺失或未知类别没列出。")
    # 交互 P 只应写在亚组标题上
    if pc:
        for sname in dict.fromkeys(r["sub"] for r in rows):
            ps = {r["p"] for r in rows if r["sub"] == sname and r["p"] is not None}
            if len(ps) > 1:
                issues.confirm(f"亚组“{sname}”的各层填了不同的 P 值（{'、'.join(f'{p:g}' for p in ps)}）。"
                               "森林图通常每个亚组变量只有一个交互 P 值。请确认这一列是交互 P 还是各层自己的 P。")
    sig_inter = [r for r in rows if r["p"] is not None and r["p"] < 0.05 and r["sub"].lower() not in overall_names]
    if sig_inter:
        issues.info("有交互 P < 0.05 的亚组：" + "、".join(dict.fromkeys(r["sub"] for r in sig_inter)) +
                    "。亚组分析多为探索性，表述时不要写成“某亚组获益更大”的确证结论。")
    return {"rows": rows, "measure": measure, "left_good": left_good,
            "has_counts": all(k is not None for k in (nn1, nn0)), "has_events": ne1 is not None and ne0 is not None,
            "has_n": ntot is not None, "has_p": pc is not None, "overall_names": overall_names}


def nice_log_ticks(lo, hi):
    cands = [0.01, 0.02, 0.05, 0.1, 0.2, 0.25, 0.5, 1, 2, 4, 5, 10, 20, 50, 100]
    t = [c for c in cands if lo <= c <= hi]
    if 1 not in t:
        t.append(1)
    t = sorted(set(t))
    if 0.2 in t and 0.25 in t:
        t.remove(0.2)
    if 4 in t and 5 in t:
        t.remove(5)
    # 刻度在对数尺度上至少间隔 0.25，从 1 向两边挑
    keep = [1]
    for side in (sorted([x for x in t if x > 1]), sorted([x for x in t if x < 1], reverse=True)):
        last = 1
        for x in side:
            if abs(math.log10(x) - math.log10(last)) >= 0.25:
                keep.append(x)
                last = x
    return sorted(keep)


def draw_forest(settings, F):
    lang_en = settings.get("图中文字", "英文").strip() not in ("中文", "中")
    rows = F["rows"]
    meas = F["measure"]
    trt = settings.get("试验组名称", "试验组")
    ctl = settings.get("对照组名称", "对照组")
    # 组织显示行
    disp = []
    order = list(dict.fromkeys(r["sub"] for r in rows))
    ovn = F["overall_names"]
    ov_first = sorted(order, key=lambda s: 0 if s.lower() in ovn else 1)
    for s in ov_first:
        rs = [r for r in rows if r["sub"] == s]
        if len(rs) == 1 and not rs[0]["lev"]:
            disp.append(("data", s, rs[0], s.lower() in ovn))
        else:
            pval = next((r for r in rs if r["p"] is not None), None)
            disp.append(("head", s, pval, False))
            for r in rs:
                disp.append(("data", "    " + r["lev"], r, False))
    n = len(disp)
    title = settings.get("图标题", "")
    row_h = 0.30
    top_rows = 2.2 + (1.3 if title else 0)
    bottom_in = 0.95
    fig_h = (n + top_rows) * row_h + bottom_in
    fig = plt.figure(figsize=(10.5, fig_h), dpi=100)
    ax = fig.add_axes([0.485, bottom_in / fig_h, 0.245, n * row_h / fig_h])
    ax.set_ylim(n - 0.5, -0.5)
    vals = [r[k] for _, _, r, _ in [d for d in disp if d[0] == "data"] for k in ("lo", "hi") if r.get(k)]
    xlo = num_setting(settings, "横轴最小值") or min(min(vals) / 1.15, 0.5)
    xhi = num_setting(settings, "横轴最大值") or max(max(vals) * 1.15, 2)
    ax.set_xscale("log")
    ax.set_xlim(xlo, xhi)
    ticks = nice_log_ticks(xlo, xhi)
    ax.set_xticks(ticks)
    ax.set_xticklabels([f"{t:g}" for t in ticks], fontsize=9)
    ax.minorticks_off()
    ax.axvline(1, color="#404040", lw=0.8)
    for s_ in ("top", "right", "left"):
        ax.spines[s_].set_visible(False)
    ax.set_yticks([])
    tr = transforms.blended_transform_factory(fig.transFigure, ax.transData)
    x_label, x_c1, x_c2, x_est, x_p = 0.015, 0.305, 0.415, 0.75, 0.94
    head_y = -1.25
    fs = 9.5
    colc = settings.get("颜色", "").strip().split(",")[0].strip()
    colc = colc if re.fullmatch(r"#?[0-9a-fA-F]{6}", colc or "") else "#1F4E79"
    colc = colc if colc.startswith("#") else "#" + colc
    show_counts = F["has_counts"]
    fig.text(x_label, head_y, "Subgroup" if lang_en else "亚组", transform=tr, fontsize=fs, fontweight="bold", va="center")
    if show_counts:
        lab = ("No. of events/patients" if F["has_events"] else "No. of patients") if lang_en else \
              ("事件数/例数" if F["has_events"] else "例数")
        fig.text(x_c1, head_y - 0.3, trt, transform=tr, fontsize=fs, fontweight="bold", ha="center", va="center")
        fig.text(x_c2, head_y - 0.3, ctl, transform=tr, fontsize=fs, fontweight="bold", ha="center", va="center")
        fig.text((x_c1 + x_c2) / 2, head_y + 0.45, lab, transform=tr, fontsize=fs - 1, ha="center", va="center",
                 color="#404040")
    elif F["has_n"]:
        fig.text(x_c1, head_y, "N" if lang_en else "例数", transform=tr, fontsize=fs, fontweight="bold", ha="center",
                 va="center")
    fig.text(x_est, head_y, f"{meas} (95% CI)", transform=tr, fontsize=fs, fontweight="bold", va="center")
    if F["has_p"]:
        fig.text(x_p, head_y, "P for\ninteraction" if lang_en else "交互 P", transform=tr, fontsize=fs,
                 fontweight="bold", ha="center", va="center", linespacing=1.1)
    fig.add_artist(Line2D([x_label, 0.985], [0, 0], transform=transforms.blended_transform_factory(fig.transFigure, ax.transData),
                          color="#808080", lw=0.6))
    fig.artists[-1].set_ydata([-0.65, -0.65])
    if title:
        fig.text(x_label, head_y - 1.35, title, transform=tr, fontsize=12, fontweight="bold", va="center")
    # 权重
    ws = []
    for d in disp:
        if d[0] == "data" and d[2].get("lo") and not d[2].get("ne"):
            se = (math.log(d[2]["hi"]) - math.log(d[2]["lo"])) / (2 * Z95)
            ws.append(1 / se ** 2 if se > 0 else 1)
    wmax = max(ws) if ws else 1
    used = {}
    for i, d in enumerate(disp):
        kind, lab, r, is_ov = d
        if kind == "head":
            fig.text(x_label, i, lab, transform=tr, fontsize=fs, fontweight="bold", va="center")
            if r is not None and F["has_p"]:
                ptxt = r.get("p_text") or f_p(r["p"])
                fig.text(x_p, i, ptxt, transform=tr, fontsize=fs, ha="center", va="center")
                used[f"{lab} 交互 P"] = ptxt
            continue
        fig.text(x_label, i, lab if not is_ov else lab, transform=tr, fontsize=fs, va="center",
                 fontweight="bold" if is_ov else "normal")
        name = r["sub"] + (" / " + r["lev"] if r["lev"] else "")
        if show_counts:
            for xc, e, nn in [(x_c1, "e1", "n1"), (x_c2, "e0", "n0")]:
                if r.get(nn) is not None:
                    t = f"{int(r[e])}/{int(r[nn])}" if (F["has_events"] and r.get(e) is not None) else f"{int(r[nn])}"
                    fig.text(xc, i, t, transform=tr, fontsize=fs, ha="center", va="center")
        elif F["has_n"] and r.get("n") is not None:
            fig.text(x_c1, i, f"{int(r['n'])}", transform=tr, fontsize=fs, ha="center", va="center")
        if r.get("ne"):
            fig.text(x_est, i, "NE", transform=tr, fontsize=fs, va="center")
            used[name] = "NE"
            continue
        hr, lo, hi = r["hr"], r["lo"], r["hi"]
        est = f"{hr:.2f} ({lo:.2f}{dash()}{hi:.2f})"
        fig.text(x_est, i, est, transform=tr, fontsize=fs, va="center", fontweight="bold" if is_ov else "normal")
        used[name] = est
        if not lab.startswith("    ") and r.get("p") is not None and F["has_p"] and not is_ov:
            ptxt = r.get("p_text") or f_p(r["p"])
            fig.text(x_p, i, ptxt, transform=tr, fontsize=fs, ha="center", va="center")
        cl, ch = max(lo, xlo), min(hi, xhi)
        if is_ov and yes(settings, "总体用菱形", True):
            ax.add_patch(Polygon([[lo, i], [hr, i - 0.28], [hi, i], [hr, i + 0.28]], closed=True, fc=colc, ec=colc))
        else:
            ax.plot([cl, ch], [i, i], color="#404040", lw=1.1, solid_capstyle="butt")
            se = (math.log(hi) - math.log(lo)) / (2 * Z95)
            w = (1 / se ** 2) if se > 0 else wmax
            ms = 4 + 7 * math.sqrt(w / wmax)
            ax.plot([hr], [i], marker="s", ms=ms, color=colc, ls="none")
            if lo < xlo:
                ax.annotate("", xy=(xlo, i), xytext=(xlo * 1.25, i), arrowprops=dict(arrowstyle="->", lw=1, color="#404040"))
            if hi > xhi:
                ax.annotate("", xy=(xhi, i), xytext=(xhi / 1.25, i), arrowprops=dict(arrowstyle="->", lw=1, color="#404040"))
    # 有利于
    lg = trt if F["left_good"] == "试验组" else ctl
    rg = ctl if F["left_good"] == "试验组" else trt
    lt = f"← Favors {lg}" if lang_en else f"← 有利于{lg}"
    rt = f"Favors {rg} →" if lang_en else f"有利于{rg} →"
    one = (math.log(1) - math.log(xlo)) / (math.log(xhi) - math.log(xlo))
    ax.annotate(lt, xy=(one, 0), xycoords="axes fraction", xytext=(-4, -30), textcoords="offset points",
                ha="right", va="top", fontsize=8.5)
    ax.annotate(rt, xy=(one, 0), xycoords="axes fraction", xytext=(4, -30), textcoords="offset points",
                ha="left", va="top", fontsize=8.5)
    foot = settings.get("图注", "")
    if foot:
        fig.text(x_label, 0.08 / fig_h, foot, fontsize=8, color="#404040", va="bottom")
    return fig, used


# ---------------------------------------------------------------------------
# 瀑布图
# ---------------------------------------------------------------------------

RESP_MAP = {"CR": "CR", "完全缓解": "CR", "PR": "PR", "部分缓解": "PR", "SD": "SD", "疾病稳定": "SD", "稳定": "SD",
            "PD": "PD", "疾病进展": "PD", "进展": "PD", "NE": "NE", "不可评估": "NE", "NA": "NE",
            "UPR": "uPR", "uPR": "uPR", "未确认PR": "uPR", "UCR": "uCR", "未确认CR": "uCR", "NON-CR/NON-PD": "SD"}
RESP_COL = {"CR": "#1F4E79", "PR": "#2E75B6", "uCR": "#5B9BD5", "uPR": "#9DC3E6", "SD": "#A5A5A5", "PD": "#C00000", "NE": "#D9D9D9"}


def parse_waterfall(settings, data, issues):
    check_common(settings, data, issues, [
        ("数据来源", "核对单要写数据出自哪份文件、哪个版本。"),
        ("疗效评价标准", "例如 RECIST 1.1、RANO、iRECIST，决定阈值线和缓解类别。"),
    ])
    idc = find_col(data, ["患者编号", "编号", "id"])
    vc = find_col(data, ["最佳变化百分比", "最大变化百分比", "变化百分比", "best change", "change"])
    rc = find_col(data, ["最佳疗效", "最佳总体疗效", "疗效", "bor", "response"])
    gc = find_col(data, ["组别", "分组", "group"])
    if vc is None:
        issues.error("数据表里找不到“最佳变化百分比”列（靶病灶直径总和相对基线的最佳变化，%）。")
        return None
    v = data[vc].map(to_num)
    bad = v[v == "BAD"]
    if len(bad):
        issues.error(f"“最佳变化百分比”有不是数字的值：{rows_str(bad.index)}。")
    miss = v[v.map(is_blank)]
    vv = v.map(lambda x: x if isinstance(x, float) else np.nan).astype(float)
    if len(miss):
        issues.confirm(f"有 {len(miss)} 个患者没有“最佳变化百分比”（{rows_str(miss.index)}），这些人画不进瀑布图。"
                       "图注会写“n = 可测量患者数”。请确认。")
    lt = vv[vv < -100]
    if len(lt):
        issues.error(f"最佳变化百分比小于 -100%：{rows_str(lt.index)}。直径不可能缩小超过 100%。")
    small = vv.dropna()
    if len(small) and small.abs().max() <= 1.0 and len(small) >= 5:
        issues.confirm("所有变化值都在 -1 到 1 之间，看起来像是小数（-0.35）而不是百分比（-35）。请确认单位。")
    big = vv[vv > 300]
    if len(big):
        issues.confirm(f"有变化超过 +300% 的值（{rows_str(big.index)}），请确认没有录入错误。")
    resp = None
    if rc:
        rr = data[rc].map(cell_str)
        resp = rr.map(lambda s: RESP_MAP.get(s.upper(), RESP_MAP.get(s)) if s else None)
        badr = rr[(rr != "") & resp.isna()]
        if len(badr):
            issues.error(f"“最佳疗效”有无法识别的值（{'、'.join(sorted(set(badr))[:6])}）。只能填 CR、PR、SD、PD、NE（未确认的写 uPR / uCR）。")
        crit = settings.get("疗效评价标准", "").upper()
        if "RECIST" in crit and not issues.errors:
            for i in data.index:
                x = vv[i]
                r = resp[i]
                if np.isnan(x) or r is None:
                    continue
                if r == "PR" and x > -30:
                    issues.confirm(f"第 {i + 2} 行：最佳疗效 PR，但靶病灶只变化 {x:g}%（RECIST 1.1 的 PR 需要缩小 ≥30%）。请确认。")
                if r == "CR" and x > -30:
                    issues.confirm(f"第 {i + 2} 行：最佳疗效 CR，但靶病灶变化 {x:g}%。请确认（淋巴结靶病灶缩到 <10 mm 时可以不到 -100%，但不会这么小）。")
                if r == "SD" and x <= -30:
                    issues.info(f"第 {i + 2} 行：靶病灶缩小 {x:g}%，最佳疗效却是 SD（可能是 PR 未确认或非靶病灶进展）。")
                if r == "PD" and x < 20:
                    issues.info(f"第 {i + 2} 行：最佳疗效 PD，靶病灶变化 {x:g}% 未达 +20%（可能是新病灶或非靶病灶进展）。")
    else:
        issues.confirm("没有“最佳疗效”列，柱子没法按 CR/PR/SD/PD 上色，只能用同一种颜色。确认这样画吗？")
    if idc:
        ids = data[idc].map(cell_str)
        dup = ids[(ids != "") & ids.duplicated(keep=False)]
        if len(dup):
            issues.error(f"患者编号有重复：{'、'.join(sorted(set(dup))[:8])}。瀑布图每个患者一根柱子。")
    rep_orr = num_setting(settings, "原文ORR", issues)
    df = pd.DataFrame({"v": vv, "resp": resp if resp is not None else None,
                       "g": data[gc].map(cell_str) if gc else ""}, index=data.index).dropna(subset=["v"])
    if rep_orr is not None and resp is not None and len(df):
        n_all = len(data)
        orr = (resp.isin(["CR", "PR"])).sum() / n_all * 100
        if abs(orr - rep_orr) > 0.15:
            issues.confirm(f"原文 ORR {rep_orr:g}%，数据算出来 {orr:.1f}%（{int(resp.isin(['CR', 'PR']).sum())}/{n_all}，只算确认的 CR/PR）。"
                           "可能分母不同（可评估人群 vs 全部入组）或是否计入未确认缓解。请确认。")
    return {"df": df, "n_all": len(data)}


def draw_waterfall(settings, W):
    lang_en = settings.get("图中文字", "英文").strip() not in ("中文", "中")
    df = W["df"].sort_values("v", ascending=False).reset_index(drop=True)
    fig, ax = plt.subplots(figsize=(max(6.5, min(14, 0.12 * len(df) + 4)), 4.6), dpi=100)
    cols = [RESP_COL.get(r, "#1F4E79") if isinstance(r, str) else "#1F4E79" for r in df["resp"]]
    ax.bar(np.arange(len(df)), df["v"], color=cols, width=0.8, edgecolor="none")
    ax.axhline(0, color="#404040", lw=0.8)
    if yes(settings, "显示阈值线", True):
        ax.axhline(20, color="#7F7F7F", lw=0.8, ls="--")
        ax.axhline(-30, color="#7F7F7F", lw=0.8, ls="--")
    ax.set_xticks([])
    ax.set_xlim(-0.8, len(df) - 0.2)
    ymin = min(-100, df["v"].min() - 5)
    ymax = max(40, df["v"].max() + 10)
    ax.set_ylim(ymin, ymax)
    for s in ("top", "right", "bottom"):
        ax.spines[s].set_visible(False)
    yl = settings.get("纵轴标题", "") or ("Best change from baseline in\nsum of target lesion diameters (%)" if lang_en
                                       else "靶病灶直径总和相对基线的\n最佳变化（%）")
    ax.set_ylabel(yl, fontsize=10)
    present = [r for r in ["CR", "uCR", "PR", "uPR", "SD", "PD", "NE"] if r in set(df["resp"].dropna())]
    if present:
        ax.legend(handles=[Patch(color=RESP_COL[r], label=r) for r in present], frameon=False, fontsize=9,
                  loc="upper right", ncol=len(present))
    ax.text(0.0, 1.0, f"n = {len(df)}", transform=ax.transAxes, fontsize=9, va="bottom")
    title = settings.get("图标题", "")
    if title:
        ax.set_title(title, fontsize=12, loc="left", fontweight="bold", pad=16)
    foot = settings.get("图注", "")
    if foot:
        fig.text(0.01, 0.01, foot, fontsize=8, color="#404040")
    used = {"图中患者数": (f"{len(df)}（数据共 {W['n_all']} 行）", "AI 由数据计数")}
    return fig, used


# ---------------------------------------------------------------------------
# 泳道图
# ---------------------------------------------------------------------------

def parse_swimmer(settings, data, issues):
    check_common(settings, data, issues, [
        ("时间单位", "横轴需要单位（月、周、天）。"),
        ("数据来源", "核对单要写数据出自哪份文件、哪个版本。"),
        ("数据截止日期", "泳道图里“仍在治疗”的箭头只在数据截止日成立（不知道就填“不详”）。"),
    ])
    idc = find_col(data, ["患者编号", "编号", "id"])
    dc = find_col(data, ["治疗时长", "时长", "治疗持续时间", "duration"])
    oc = find_col(data, ["仍在治疗", "是否仍在治疗", "ongoing"])
    cc = find_col(data, ["分类", "最佳疗效", "颜色分组", "组别", "category"])
    evcols = {k: find_col(data, a) for k, a in [("resp", ["缓解时间", "缓解开始", "首次缓解时间"]),
                                                   ("prog", ["进展时间", "疾病进展时间"]),
                                                   ("death", ["死亡时间"])]}
    if dc is None:
        issues.error("数据表里找不到“治疗时长”列。")
        return None
    if idc is None:
        issues.error("数据表里找不到“患者编号”列（用 P01、P02 这样的代号）。")
        return None
    d = data[dc].map(to_num)
    bad = d[d.map(lambda x: is_blank(x) or x == "BAD")]
    if len(bad):
        issues.error(f"“治疗时长”有空白或不是数字：{rows_str(bad.index)}。")
    dv = d.map(lambda x: x if isinstance(x, float) else np.nan).astype(float)
    neg = dv[dv <= 0]
    if len(neg):
        issues.error(f"“治疗时长”有 0 或负数：{rows_str(neg.index)}。")
    ids = data[idc].map(cell_str)
    dup = ids[(ids != "") & ids.duplicated(keep=False)]
    if len(dup):
        issues.error(f"患者编号有重复：{'、'.join(sorted(set(dup))[:8])}。")
    ong = data[oc].map(cell_str).map(lambda s: s in ("是", "1", "yes", "y", "仍在", "在治")) if oc else pd.Series(False, index=data.index)
    if not oc:
        issues.confirm("没有“仍在治疗”列，图上不会画表示“截至数据截止仍在治疗”的箭头。确认吗？")
    ev = {}
    for k, c in evcols.items():
        if c:
            x = data[c].map(to_num)
            badx = x[x == "BAD"]
            if len(badx):
                issues.error(f"“{c}”有不是数字的值：{rows_str(badx.index)}。")
            ev[k] = x.map(lambda y: y if isinstance(y, float) else np.nan).astype(float)
    for k, s in ev.items():
        over = s[(s > dv + 1e-9)]
        if len(over):
            nm = {"resp": "缓解时间", "prog": "进展时间", "death": "死亡时间"}[k]
            issues.confirm(f"有 {len(over)} 个患者的{nm}晚于治疗结束（{rows_str(over.index)}）。如果柱子表示“治疗时长”，这些标记会画在柱子外面。请确认柱子表示治疗时长还是随访时长。")
    if "death" in ev:
        both = ev["death"][ev["death"].notna() & ong]
        if len(both):
            issues.error(f"有患者同时标了“仍在治疗”和“死亡时间”：{rows_str(both.index)}。")
    if "resp" in ev and "prog" in ev:
        bad_o = ev["resp"][(ev["resp"].notna()) & (ev["prog"].notna()) & (ev["resp"] > ev["prog"])]
        if len(bad_o):
            issues.confirm(f"有患者的缓解时间晚于进展时间（{rows_str(bad_o.index)}），请确认。")
    if len(data) > 60:
        issues.confirm(f"一共 {len(data)} 个患者，泳道图会很长、字很小。确认要全部画吗？")
    cat = data[cc].map(cell_str) if cc else pd.Series("", index=data.index)
    df = pd.DataFrame({"id": ids, "dur": dv, "ong": ong, "cat": cat}, index=data.index)
    for k, s in ev.items():
        df[k] = s
    return {"df": df.dropna(subset=["dur"]), "catname": cc}


def draw_swimmer(settings, S):
    lang_en = settings.get("图中文字", "英文").strip() not in ("中文", "中")
    unit = settings.get("时间单位", "月")
    df = S["df"].sort_values("dur", ascending=True).reset_index(drop=True)
    n = len(df)
    fig, ax = plt.subplots(figsize=(8, max(3, 0.24 * n + 1.4)), dpi=100)
    cats = list(dict.fromkeys(df["cat"].tolist()))
    pal = parse_colors(settings, max(len(cats), 1))
    cmap = {c: (RESP_COL.get(RESP_MAP.get(c.upper(), ""), None) or pal[i % len(pal)]) for i, c in enumerate(cats)}
    for i, r in df.iterrows():
        ax.barh(i, r["dur"], color=cmap.get(r["cat"], "#1F4E79"), height=0.65, edgecolor="none")
        if r["ong"]:
            ax.annotate("", xy=(r["dur"] + max(df["dur"]) * 0.04, i), xytext=(r["dur"], i),
                        arrowprops=dict(arrowstyle="-|>", lw=1, color="#404040"))
    mk = {"resp": ("^", "#548235", "Response onset" if lang_en else "缓解开始"),
          "prog": ("D", "#C00000", "Progression" if lang_en else "疾病进展"),
          "death": ("X", "#000000", "Death" if lang_en else "死亡")}
    handles = []
    for k, (m, c, lab) in mk.items():
        if k in df.columns:
            sub = df[df[k].notna()]
            ax.plot(sub[k], sub.index, ls="none", marker=m, ms=5.5, color=c)
            if len(sub):
                handles.append(Line2D([], [], ls="none", marker=m, color=c, label=lab))
    if df["ong"].any():
        handles.append(Line2D([], [], color="#404040", marker=">", ls="-", label="Ongoing" if lang_en else "仍在治疗"))
    for c in cats:
        if c:
            handles.insert(0, Patch(color=cmap[c], label=c))
    ax.set_yticks(range(n))
    ax.set_yticklabels(df["id"], fontsize=7 if n > 30 else 8.5)
    ax.set_ylim(-0.7, n - 0.3)
    ax.set_xlim(0, max(df["dur"]) * 1.1)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_xlabel(settings.get("横轴标题", "") or (f"Time on treatment ({UNIT_EN.get(unit, unit)})" if lang_en
                                                    else f"治疗时间（{UNIT_CN.get(unit, unit)}）"), fontsize=10)
    ax.set_ylabel("Patient" if lang_en else "患者", fontsize=10)
    if handles:
        ax.legend(handles=handles, frameon=False, fontsize=8.5, loc="lower right")
    title = settings.get("图标题", "")
    if title:
        ax.set_title(title, fontsize=12, loc="left", fontweight="bold")
    return fig, {"图中患者数": (str(n), "AI 由数据计数")}


# ---------------------------------------------------------------------------
# 输出与流程
# ---------------------------------------------------------------------------

def print_report(kind, issues, path):
    tn = TYPE_NAMES[kind]
    print(f"========== {tn}：数据检查结果 ==========")
    print(f"文件：{path}")
    print(f"必须修改：{len(issues.errors)} 条；需要使用者确认：{len(issues.confirms)} 条；提示：{len(issues.infos)} 条")
    if issues.infos:
        print("\n【提示】（不影响画图，可顺带告诉使用者）")
        for i, m in enumerate(issues.infos, 1):
            print(f"  {i}. {m}")
    if issues.errors or issues.confirms:
        print("\n========== 请把下面这段原样发给使用者 ==========")
        print(f"我检查了您给的{tn}数据，发现下面的问题，需要您处理后我才能画图：")
        k = 1
        if issues.errors:
            print("\n一、必须修改（数据或信息有错误、缺失，请改正或补充）：")
            for m in issues.errors:
                print(f"  {k}. {m}")
                k += 1
        if issues.confirms:
            print("\n二、需要您确认（可能有问题，请逐条回复）：")
            for m in issues.confirms:
                print(f"  {k}. {m}")
                k += 1
        print("\n请按编号逐条回复。我不会自己猜或修改数字。")
        print("========== 以上发给使用者 ==========")
    print("\n========== 给 AI 助手的下一步 ==========")
    if issues.errors:
        print("有【必须修改】的问题：不能画图。把上面那段发给使用者，等使用者给出改正后的数据或补充信息，"
              "按使用者的回复更新 Excel（数据数字只按使用者给的改，不能自己编），再运行 check。")
        return 2
    if issues.confirms:
        print("没有【必须修改】，但有【需要确认】：把上面那段发给使用者，等回复。"
              "使用者逐条回复后：需要改的先改 Excel；全部确认后运行\n"
              f"  python statfig.py draw \"{path}\" --confirm \"使用者回复的原话\"")
        return 3
    print(f"检查通过，可以画图：python statfig.py draw \"{path}\"")
    return 0


def write_checklist(path, kind, settings, issues, used, extra_lines, confirm_text, outputs, font_desc):
    L = []
    tn = TYPE_NAMES[kind]
    L.append(f"# {tn} 核对单")
    L.append("")
    L.append(f"- 生成时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}（statfig {VERSION}）")
    for k in ["研究名称", "终点名称", "数据来源", "数据截止日期", "图标题"]:
        if settings.get(k):
            L.append(f"- {k}：{settings[k]}")
    L.append(f"- 输出文件：{'、'.join(os.path.basename(o) for o in outputs)}")
    L.append(f"- 字体：{font_desc}")
    L.append("")
    L.append("## 图中出现的数字及来源")
    L.append("")
    L.append("| 位置 | 图中写法 | 来源 |")
    L.append("|---|---|---|")
    for k, v in used.items():
        if isinstance(v, tuple):
            L.append(f"| {k} | {v[0]} | {v[1]} |")
        else:
            L.append(f"| {k} | {v} | 使用者提供的数据表（照抄） |")
    L.append("")
    L.append("标“AI 计算”的数字是本工具从逐例数据算出来的，不是原文报告值：对外使用前请统计师或对照原文复核。")
    if extra_lines:
        L.append("")
        L.append("## 计算明细（供复核）")
        L.append("")
        L.extend(extra_lines)
    L.append("")
    L.append("## 使用者确认过的问题")
    L.append("")
    if issues.confirms:
        for i, m in enumerate(issues.confirms, 1):
            L.append(f"{i}. {m}")
        L.append("")
        L.append(f"使用者的回复（原话）：{confirm_text}")
    else:
        L.append("无")
    if issues.infos:
        L.append("")
        L.append("## 提示")
        L.append("")
        for i, m in enumerate(issues.infos, 1):
            L.append(f"{i}. {m}")
    L.append("")
    L.append("## 使用提醒")
    L.append("")
    L.append("- 不同研究的曲线或数字放在一起时，必须注明：非头对头比较，研究设计与人群不同，仅供参考。")
    L.append("- 原文已有的图（论文、PPT 里的原图）应直接使用原图并注明出处，不要用本工具重画。")
    L.append("- 亚组分析、探索性分析不能写成确证性结论。")
    with open(path, "w", encoding="utf-8-sig") as f:
        f.write("\n".join(L) + "\n")


def get_kind(settings, path):
    k = norm_type(settings.get("图类型", ""))
    if k is None:
        print(f"【错误】设置表“图类型”填的是“{settings.get('图类型', '')}”，只能填：KM、森林图、瀑布图、泳道图。")
        sys.exit(1)
    return k


def run_check(path, quiet=False):
    settings, data, reported = read_workbook(path)
    kind = get_kind(settings, path)
    issues = Issues()
    parsed, extra = None, None
    if kind == "km":
        parsed = parse_km(settings, data, reported, issues)
        if parsed is not None and not issues.errors:
            extra = km_compute(parsed, issues)
    elif kind == "forest":
        parsed = parse_forest(settings, data, issues)
    elif kind == "waterfall":
        parsed = parse_waterfall(settings, data, issues)
    elif kind == "swimmer":
        parsed = parse_swimmer(settings, data, issues)
    return kind, settings, issues, parsed, extra


def cmd_check(path):
    kind, settings, issues, parsed, extra = run_check(path)
    code = print_report(kind, issues, path)
    if kind == "km" and code != 2 and extra is not None:
        print("\n（参考）计算结果预览：")
        for l in km_summary_lines(parsed, extra):
            print("  " + l)
    return code


def cmd_draw(path, confirm_text):
    kind, settings, issues, parsed, extra = run_check(path)
    if issues.errors:
        print("【不能画图】数据还有必须修改的问题。")
        return print_report(kind, issues, path)
    if issues.confirms and not confirm_text:
        print("【不能画图】还有需要使用者确认的问题，没有收到使用者的回复（--confirm）。")
        return print_report(kind, issues, path)
    if issues.confirms and len(confirm_text.strip()) < 2:
        print("【不能画图】--confirm 后面要写使用者回复的原话，不能为空。")
        return 3
    has_cjk_font, font_desc = setup_fonts()
    lang_cn = settings.get("图中文字", "英文").strip() in ("中文", "中")
    all_text = " ".join(settings.values())
    if (lang_cn or has_cjk(all_text) or (kind == "km" and any(has_cjk(lv) for lv in parsed["levels"]))) and not has_cjk_font:
        print("【注意】这台电脑没有找到中文字体，图里的中文会显示成方框。请把图例、组别改成英文，或把“图中文字”设为英文。")
    if kind == "km":
        fig, used = draw_km(settings, parsed, extra, None)
        extra_lines = km_summary_lines(parsed, extra)
    elif kind == "forest":
        fig, used = draw_forest(settings, parsed)
        extra_lines = []
    elif kind == "waterfall":
        fig, used = draw_waterfall(settings, parsed)
        extra_lines = []
    else:
        fig, used = draw_swimmer(settings, parsed)
        extra_lines = []
    base_dir = os.path.dirname(os.path.abspath(path))
    stem = settings.get("输出文件名", "") or os.path.splitext(os.path.basename(path))[0]
    stem = re.sub(r'[\\/:*?"<>|]', "_", stem)
    outdir = os.path.join(base_dir, stem + "_出图")
    os.makedirs(outdir, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    png = os.path.join(outdir, f"{stem}_{stamp}.png")
    pdf = os.path.join(outdir, f"{stem}_{stamp}.pdf")
    fig.savefig(png, dpi=300, bbox_inches="tight", facecolor="white")
    fig.savefig(pdf, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    chk = os.path.join(outdir, f"{stem}_{stamp}_核对单.md")
    write_checklist(chk, kind, settings, issues, used, extra_lines, confirm_text, [png, pdf], font_desc)
    print("========== 画图完成 ==========")
    print(f"图片（放 PPT 用）：{png}")
    print(f"PDF（矢量，可无限放大）：{pdf}")
    print(f"核对单：{chk}")
    print("\n========== 给 AI 助手的下一步 ==========")
    print("1. 打开 PNG 看一眼：文字有没有重叠、方框乱码、超出边界。有问题按 SKILL.md 第 6 步调整设置后重画。")
    print("2. 把图和核对单交给使用者，并用一两句话说明：图中哪些数字是 AI 计算的、需要复核（见核对单）。")
    return 0


# ---------------------------------------------------------------------------
# 模板
# ---------------------------------------------------------------------------

def template_spec(kind):
    common_tail = [
        ("图中文字", "英文", "英文 或 中文（默认英文，投稿和壁报常用英文）"),
        ("图标题", "", "可空。图上方的标题"),
        ("图注", "", "可空。图下方的小字，例如数据截止日期、缩写说明"),
        ("颜色", "", "可空。用十六进制色号，多个用逗号隔开，例如 #1F4E79,#C00000"),
        ("输出文件名", "", "可空。默认用 Excel 文件名"),
    ]
    if kind == "km":
        settings = [
            ("图类型", "KM", "不要改"),
            ("研究名称", "", "必填。例如 XX 研究"),
            ("终点名称", "", "必填。例如 OS、PFS、DFS"),
            ("终点定义", "", "建议填。例如：从随机化到疾病复发或任何原因死亡"),
            ("时间单位", "", "必填。“时间”列的单位：月、天、周、年"),
            ("数据来源", "", "必填。哪份文件、哪个版本，例如：XX 研究全文稿 2026-09-03 逐例数据"),
            ("数据截止日期", "", "必填。不知道就填“不详”"),
            ("组别顺序", "", "可空。多组时按此顺序画，逗号隔开，必须和数据里的组名完全一样"),
            ("对照组", "", "多组时必填。HR = 其他组 vs 对照组"),
            ("图例名称", "", "可空。图例里显示的组名，按组别顺序，逗号隔开"),
            ("关注时间点", "", "可空。例如 12,24,36：在图上和核对单里给出这些时间点的生存率"),
            ("置信区间方法", "log-log", "log-log（默认，与 SAS 一致）或 log（R 默认）"),
            ("图中统计量", "自动", "自动（有原文值用原文，没有就用计算值）/ 原文 / 计算 / 不显示"),
            ("原文HR", "", "可空。论文或 PPT 报告的 HR，用来核对并显示在图上"),
            ("原文HR下限", "", "可空"),
            ("原文HR上限", "", "可空"),
            ("原文HR是否分层", "", "填了原文 HR 时必填：是 / 否 / 不知道"),
            ("原文P值", "", "可空。原文报告的 P 值"),
            ("显示风险表", "是", "是 / 否（图下方的 No. at risk）"),
            ("显示删失标记", "是", "是 / 否（曲线上的 + 号）"),
            ("显示中位线", "否", "是 / 否"),
            ("显示置信区间阴影", "否", "是 / 否"),
            ("横轴间隔", "", "可空。例如 6 表示每 6 个月一个刻度"),
            ("横轴最大值", "", "可空"),
            ("纵轴标题", "", "可空。默认按终点名称自动生成"),
            ("横轴标题", "", "可空。默认 Time (months)"),
        ] + common_tail
        cols = ["患者编号", "组别", "时间", "事件"]
        example = [["P001", "试验药", 13.2, 1], ["P002", "试验药", 24.0, 0], ["P003", "安慰剂", 6.5, 1],
                   ["P004", "安慰剂", 18.1, 0]]
        notes = ["每个患者一行。", "时间：从起点（如随机化、入组、手术）到事件或最后一次随访的时间，单位填在设置表“时间单位”。",
                 "事件：1 = 发生了终点事件（如复发、进展、死亡），0 = 删失（截至最后随访仍未发生）。",
                 "只有日期没有时间时：把“时间”列换成“起始日期”和“事件或末次随访日期”两列，工具会自己算（失访者填最后一次有效随访日期）。",
                 "组别：单臂研究可以不要这一列。", "不要放姓名、病历号、身份证号、电话。",
                 "如果原文报告了每组的例数、事件数、中位数，请填到“原文报告值”表，工具会自动核对。"]
        reported_cols = ["组别", "例数", "事件数", "中位时间", "中位下限", "中位上限"]
    elif kind == "forest":
        settings = [
            ("图类型", "森林图", "不要改"),
            ("研究名称", "", "必填"),
            ("终点名称", "", "必填。例如 PFS"),
            ("数据来源", "", "必填。哪份文件、哪张图表，例如：NEJM 2023 Figure 3"),
            ("数据截止日期", "", "建议填"),
            ("效应量", "HR", "HR、OR 或 RR"),
            ("试验组名称", "", "必填。例如 Drug A"),
            ("对照组名称", "", "必填。例如 Placebo"),
            ("HR方向", "", "必填。确认原文 HR 是“试验组 vs 对照组”后填：试验组 vs 对照组"),
            ("小于1有利于", "", "可空。HR 默认 <1 有利于试验组；OR/RR 用于缓解率时要填"),
            ("总体用菱形", "是", "是 / 否"),
            ("横轴最小值", "", "可空"),
            ("横轴最大值", "", "可空"),
        ] + common_tail
        cols = ["亚组", "分层", "试验组事件数", "试验组例数", "对照组事件数", "对照组例数", "HR", "下限", "上限", "交互P"]
        example = [["总体", "", 50, 200, 90, 200, 0.52, 0.37, 0.73, ""],
                   ["年龄", "<65 岁", 30, 120, 55, 118, 0.50, 0.32, 0.78, 0.68],
                   ["年龄", "≥65 岁", 20, 80, 35, 82, 0.56, 0.32, 0.97, 0.68]]
        notes = ["每个亚组的每一层一行；“总体”一行的“亚组”填“总体”，“分层”留空。",
                 "HR、下限、上限照抄原文，不要自己换算；原文写 NE 的照写 NE。",
                 "交互P：同一个亚组的几层填同一个值（或只在第一层填）；原文写 <0.001 的照写。",
                 "事件数、例数可以不填（删掉这几列即可），图上就不显示人数。"]
        reported_cols = None
    elif kind == "waterfall":
        settings = [
            ("图类型", "瀑布图", "不要改"),
            ("研究名称", "", "必填"),
            ("数据来源", "", "必填"),
            ("数据截止日期", "", "建议填"),
            ("疗效评价标准", "", "必填。例如 RECIST 1.1"),
            ("原文ORR", "", "可空。原文报告的 ORR（%），用来核对"),
            ("显示阈值线", "是", "是 / 否（+20% 和 -30% 虚线）"),
            ("纵轴标题", "", "可空"),
        ] + common_tail
        cols = ["患者编号", "最佳变化百分比", "最佳疗效"]
        example = [["P001", -45, "PR"], ["P002", 12, "SD"], ["P003", 25, "PD"]]
        notes = ["每个患者一行。", "最佳变化百分比：靶病灶直径总和相对基线的最佳（最小）变化，单位 %，缩小写负数，例如 -45。",
                 "最佳疗效：CR、PR、SD、PD、NE；未确认的写 uPR、uCR。"]
        reported_cols = None
    else:
        settings = [
            ("图类型", "泳道图", "不要改"),
            ("研究名称", "", "必填"),
            ("数据来源", "", "必填"),
            ("数据截止日期", "", "必填。不知道就填“不详”"),
            ("时间单位", "月", "月、周、天"),
            ("横轴标题", "", "可空"),
        ] + common_tail
        cols = ["患者编号", "治疗时长", "仍在治疗", "分类", "缓解时间", "进展时间", "死亡时间"]
        example = [["P001", 14.5, "是", "PR", 1.9, "", ""], ["P002", 6.2, "否", "SD", "", 6.0, ""]]
        notes = ["每个患者一行。", "治疗时长：从开始治疗到停药（或数据截止）的时间。", "仍在治疗：是 / 否（截至数据截止日）。",
                 "分类：柱子颜色依据，例如最佳疗效。", "缓解时间、进展时间、死亡时间：从开始治疗算起的时间，没有就空着。"]
        reported_cols = None
    return settings, cols, example, notes, reported_cols


def cmd_template(kind_s, out, quiet=False):
    kind = norm_type(kind_s)
    if kind is None:
        print("【错误】图类型只能是：KM、森林图、瀑布图、泳道图。")
        return 1
    if not out.lower().endswith(".xlsx"):
        out += ".xlsx"
    if os.path.exists(out):
        print(f"【错误】{out} 已经存在，为避免覆盖，请换一个文件名。")
        return 1
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    settings, cols, example, notes, reported_cols = template_spec(kind)
    wb = Workbook()
    ws = wb.active
    ws.title = "设置"
    ws.append(["项目", "填写", "说明"])
    for k, v, d in settings:
        ws.append([k, v, d])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    for r in ws.iter_rows(min_row=2):
        if "必填" in str(r[2].value):
            r[0].font = Font(bold=True, color="C00000")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 70
    wd = wb.create_sheet("数据")
    wd.append(cols)
    for c in wd[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    for i in range(len(cols)):
        wd.column_dimensions[chr(65 + i)].width = 14
    if reported_cols:
        wr = wb.create_sheet("原文报告值")
        wr.append(reported_cols)
        for c in wr[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="DDEBF7")
    we = wb.create_sheet("示例（不会被读取）")
    we.append(["下面是填写格式示例（虚构数据），真实数据请填在“数据”表"])
    we.append(cols)
    for row in example:
        we.append(row)
    wn = wb.create_sheet("填写说明")
    for line in notes:
        wn.append([line])
    wn.column_dimensions["A"].width = 110
    for w in (we, wn):
        for row in w.iter_rows():
            for c in row:
                c.alignment = Alignment(wrap_text=False)
    wb.save(out)
    if quiet:
        return 0
    print(f"已生成模板：{out}")
    print("把这个文件发给使用者，请他：1）在“设置”表填好红色的必填项；2）把数据粘贴到“数据”表（第一行列名不要改）。")
    return 0


# ---------------------------------------------------------------------------
# import：把使用者自己的表格搬进模板；set：填写设置项
# ---------------------------------------------------------------------------

IMPORT_ALIASES = {
    "km": {"患者编号": ["患者编号", "编号", "受试者编号", "id", "subject", "patient", "病例编号", "研究编号"],
           "组别": ["组别", "分组", "组", "治疗组", "arm", "group", "treatment"],
           "时间": ["时间", "生存时间", "随访时间", "time", "os_time", "pfs_time", "dfs_time", "os月", "pfs月", "dfs月",
                  "os(月)", "pfs(月)", "dfs(月)", "os_月", "pfs_月", "dfs_月"],
           "事件": ["事件", "状态", "结局", "event", "status", "os事件", "pfs事件", "dfs事件", "os_event", "pfs_event"],
           "起始日期": ["起始日期", "起点日期", "入组日期", "随机日期", "随机化日期", "手术日期", "开始日期", "治疗开始日期"],
           "事件或末次随访日期": ["事件或末次随访日期", "事件/末次随访日期", "结束日期", "终点日期", "末次随访日期"]},
    "forest": {"亚组": ["亚组", "亚组变量", "subgroup", "变量"], "分层": ["分层", "分类", "类别", "水平", "level", "category"],
               "试验组事件数": ["试验组事件数"], "试验组例数": ["试验组例数"], "对照组事件数": ["对照组事件数"],
               "对照组例数": ["对照组例数"], "例数": ["例数", "总例数", "n"],
               "HR": ["hr", "or", "rr", "效应量", "estimate"], "下限": ["下限", "95%ci下限", "ci下限", "lower", "lcl"],
               "上限": ["上限", "95%ci上限", "ci上限", "upper", "ucl"],
               "交互P": ["交互p", "交互作用p", "交互p值", "p interaction", "p for interaction", "pinteraction"]},
    "waterfall": {"患者编号": ["患者编号", "编号", "id", "subject", "patient"],
                  "最佳变化百分比": ["最佳变化百分比", "最大变化百分比", "变化百分比", "best change", "change", "最佳变化"],
                  "最佳疗效": ["最佳疗效", "最佳总体疗效", "疗效", "bor", "response"], "组别": ["组别", "分组", "group"]},
    "swimmer": {"患者编号": ["患者编号", "编号", "id", "subject", "patient"],
                "治疗时长": ["治疗时长", "时长", "治疗持续时间", "duration"], "仍在治疗": ["仍在治疗", "是否仍在治疗", "ongoing"],
                "分类": ["分类", "最佳疗效", "颜色分组", "category"], "缓解时间": ["缓解时间", "缓解开始", "首次缓解时间"],
                "进展时间": ["进展时间", "疾病进展时间"], "死亡时间": ["死亡时间"]},
}


def read_any_table(path, sheet=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".txt"):
        for enc in ("utf-8-sig", "gbk", "utf-16"):
            try:
                return pd.read_csv(path, dtype=object, encoding=enc)
            except Exception:
                continue
        print("【错误】读不了这个 CSV 文件（编码不对）。请另存为 UTF-8 编码的 CSV 或 Excel。")
        sys.exit(1)
    xl = pd.ExcelFile(path)
    if sheet is None:
        if "数据" in xl.sheet_names:
            sheet = "数据"
        elif len(xl.sheet_names) > 1:
            print(f"这个 Excel 有多个工作表：{'、'.join(xl.sheet_names)}。默认读第一个“{xl.sheet_names[0]}”。"
                  "如果数据不在这个表，请加 --sheet 表名 重新运行。")
            sheet = xl.sheet_names[0]
        else:
            sheet = xl.sheet_names[0]
    return pd.read_excel(xl, sheet, dtype=object)


def cmd_import(kind_s, src, out, mapping_s=None, sheet=None):
    kind = norm_type(kind_s)
    if kind is None:
        print("【错误】图类型只能是：KM、森林图、瀑布图、泳道图。")
        return 1
    if not os.path.exists(src):
        print(f"【错误】找不到文件：{src}")
        return 1
    df = read_any_table(src, sheet)
    df.columns = [cell_str(c) for c in df.columns]
    df = df.loc[:, [c for c in df.columns if c and not c.startswith("Unnamed")]].dropna(how="all")
    aliases = IMPORT_ALIASES[kind]
    manual = {}
    if mapping_s:
        for part in re.split(r"[,，;；]", mapping_s):
            if "=" in part:
                k, v = part.split("=", 1)
                manual[k.strip()] = v.strip()
    used_src = {}
    for std, als in aliases.items():
        if std in manual:
            if manual[std] not in df.columns:
                print(f"【错误】--map 里写的列“{manual[std]}”在文件里找不到。文件里的列：{'、'.join(df.columns)}")
                return 1
            used_src[std] = manual[std]
            continue
        low = {c.lower().replace(" ", ""): c for c in df.columns}
        for a in als:
            if a.lower().replace(" ", "") in low:
                used_src[std] = low[a.lower().replace(" ", "")]
                break
    privacy = [c for c in df.columns if any(w in c.lower() for w in PRIVACY_WORDS)]
    for std, c in list(used_src.items()):
        if c in privacy:
            del used_src[std]
    ignored = [c for c in df.columns if c not in used_src.values()]
    need = {"km": [["时间", "起始日期"], ["事件"]], "forest": [["亚组"], ["HR"], ["下限"], ["上限"]],
            "waterfall": [["最佳变化百分比"]], "swimmer": [["患者编号"], ["治疗时长"]]}[kind]
    missing = [opts[0] for opts in need if not any(o in used_src for o in opts)]
    if kind == "km" and "起始日期" in used_src and "事件或末次随访日期" not in used_src and "时间" not in used_src:
        missing.append("事件或末次随访日期")
    print(f"文件里的列：{'、'.join(df.columns)}")
    print("对应关系：" + ("；".join(f"{k} ← {v}" for k, v in used_src.items()) or "无"))
    if privacy:
        print(f"【没有复制】可能识别患者身份的列：{'、'.join(privacy)}（隐私保护，不会复制到新文件）")
    if ignored:
        others = [c for c in ignored if c not in privacy]
        if others:
            print(f"【没有复制】用不到或认不出的列：{'、'.join(others)}")
    if missing:
        print(f"【错误】找不到这些必需的列：{'、'.join(missing)}。")
        print("请问使用者：上面哪一列是“" + "”、“".join(missing) + "”？然后加 --map 重新运行，例如：")
        print(f'  python statfig.py import {kind_s} "{src}" "{out}" --map "{missing[0]}=原文件里的列名"')
        return 2
    if kind == "km" and "时间" in used_src and "事件" in used_src:
        if any(w in used_src["事件"] for w in ["删失", "censor"]):
            print("【错误】对应到“事件”的那一列看起来是删失标记（1=删失），方向相反。请使用者确认后改成事件列。")
            return 2
    if os.path.exists(out):
        print(f"【错误】{out} 已经存在，为避免覆盖，请换一个文件名。")
        return 1
    rc = cmd_template(kind_s, out, quiet=True)
    if rc != 0:
        return rc
    from openpyxl import load_workbook
    wb = load_workbook(out)
    ws = wb["数据"]
    cols = list(used_src.keys())
    ws.delete_rows(1, ws.max_row)
    ws.append(cols)
    for _, r in df.iterrows():
        row = []
        for std in cols:
            v = r[used_src[std]]
            if isinstance(v, float) and math.isnan(v):
                v = None
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            row.append(v)
        ws.append(row)
    wsS = wb["设置"]
    wsS.append(["原始列对应", "；".join(f"{k}←{v}" for k, v in used_src.items()),
                f"import 自动记录（来自 {os.path.basename(src)}），不要改"])
    wb.save(out)
    print(f"已生成：{out}（数据 {len(df)} 行，原文件没有改动）")
    print("下一步：用 set 命令把使用者告诉你的信息填进设置表，然后运行 check。")
    return 0


def cmd_set(path, pairs):
    from openpyxl import load_workbook
    if not os.path.exists(path):
        print(f"【错误】找不到文件：{path}")
        return 1
    try:
        wb = load_workbook(path)
    except PermissionError:
        print("【错误】文件正被 Excel 打开，请先关闭再试。")
        return 1
    if "设置" not in wb.sheetnames:
        print("【错误】这个文件没有“设置”表，不是本工具的模板。")
        return 1
    ws = wb["设置"]
    keys = {cell_str(r[0].value): r for r in ws.iter_rows(min_row=2)}
    changed = []
    for p_ in pairs:
        if "=" not in p_:
            print(f"【错误】“{p_}”格式不对，要写成 项目=内容，例如 时间单位=月")
            return 1
        k, v = p_.split("=", 1)
        k, v = k.strip(), v.strip()
        if k not in keys:
            print(f"【错误】设置表里没有“{k}”这一项。可以填的项目：{'、'.join(x for x in keys if x)}")
            return 1
        if k == "图类型":
            print("【错误】“图类型”不能改。换图类型请重新生成模板。")
            return 1
        old = cell_str(keys[k][1].value)
        keys[k][1].value = v
        changed.append(f"{k}：{old or '（空）'} → {v or '（空）'}")
    wb.save(path)
    print("已更新设置：")
    for c in changed:
        print("  " + c)
    return 0


def cmd_report(path, pairs):
    """KM：写入一组的原文报告值（例数、事件数、中位时间及 95% CI）。同名组别会被替换。"""
    from openpyxl import load_workbook
    if not os.path.exists(path):
        print(f"【错误】找不到文件：{path}")
        return 1
    wb = load_workbook(path)
    if "原文报告值" not in wb.sheetnames:
        print("【错误】这个文件没有“原文报告值”表（只有 KM 模板有）。")
        return 1
    ws = wb["原文报告值"]
    header = [cell_str(c.value) for c in ws[1]]
    rec = {}
    for p_ in pairs:
        if "=" not in p_:
            print(f"【错误】“{p_}”格式不对，要写成 项目=内容，例如 例数=60")
            return 1
        k, v = [x.strip() for x in p_.split("=", 1)]
        if k not in header:
            print(f"【错误】没有“{k}”这一项。可以填：{'、'.join(header)}")
            return 1
        if k != "组别":
            n = to_num(v)
            if n == "BAD" and v.upper() not in ("NR", "NE", "未达到"):
                print(f"【错误】{k} 填的是“{v}”，需要数字或 NR。")
                return 1
            v = n if isinstance(n, float) else ("NR" if v else None)
        rec[k] = v
    if "组别" not in rec:
        print("【错误】必须写 组别=…（和数据里的组名完全一样；单臂研究写 组别=全部患者）。")
        return 1
    target = None
    for row in ws.iter_rows(min_row=2):
        if cell_str(row[0].value) == rec["组别"]:
            target = row
    if target is None:
        ws.append([rec.get(h) for h in header])
    else:
        for i, h in enumerate(header):
            if h in rec:
                target[i].value = rec[h]
    wb.save(path)
    print(("已更新" if target is not None else "已写入") + "原文报告值：" +
          "；".join(f"{k}={cell_str(v) if v is not None else '（空）'}" for k, v in rec.items()))
    return 0


def cmd_doctor():
    print(f"statfig {VERSION}")
    print(f"Python {sys.version.split()[0]}；numpy {np.__version__}；pandas {pd.__version__}；matplotlib {mpl.__version__}")
    ok, desc = setup_fonts()
    print(f"字体：{desc}")
    if not ok:
        print("【注意】没有中文字体：图里只能用英文。")
    print("环境正常，可以使用。")
    return 0


def main(argv):
    if len(argv) < 2 or argv[1] in ("-h", "--help", "help"):
        print(__doc__)
        return 1
    cmd = argv[1].lower()
    if cmd == "doctor":
        return cmd_doctor()
    if cmd == "template":
        if len(argv) < 4:
            print("用法：python statfig.py template KM|森林图|瀑布图|泳道图 输出.xlsx")
            return 1
        return cmd_template(argv[2], argv[3])
    if cmd == "import":
        if len(argv) < 5:
            print('用法：python statfig.py import KM|森林图|瀑布图|泳道图 使用者的表格.xlsx 新文件.xlsx [--map "时间=OS月数"] [--sheet 表名]')
            return 1
        mp = argv[argv.index("--map") + 1] if "--map" in argv and argv.index("--map") + 1 < len(argv) else None
        sh = argv[argv.index("--sheet") + 1] if "--sheet" in argv and argv.index("--sheet") + 1 < len(argv) else None
        return cmd_import(argv[2], argv[3], argv[4], mp, sh)
    if cmd == "set":
        if len(argv) < 4:
            print("用法：python statfig.py set 文件.xlsx 项目=内容 [项目=内容 ...]")
            return 1
        return cmd_set(argv[2], argv[3:])
    if cmd == "report":
        if len(argv) < 4:
            print("用法：python statfig.py report 文件.xlsx 组别=安慰剂 例数=60 事件数=40 中位时间=13.5 中位下限=10.2 中位上限=18.0")
            return 1
        return cmd_report(argv[2], argv[3:])
    if cmd == "check":
        if len(argv) < 3:
            print("用法：python statfig.py check 数据.xlsx")
            return 1
        return cmd_check(argv[2])
    if cmd == "draw":
        if len(argv) < 3:
            print("用法：python statfig.py draw 数据.xlsx [--confirm \"使用者原话\"]")
            return 1
        conf = ""
        if "--confirm" in argv:
            i = argv.index("--confirm")
            conf = " ".join(argv[i + 1:]) if i + 1 < len(argv) else ""
        return cmd_draw(argv[2], conf)
    print(f"不认识的命令：{argv[1]}。可用命令：doctor、template、import、set、report、check、draw")
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))

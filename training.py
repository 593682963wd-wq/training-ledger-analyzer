"""作者 Wangyuan，2026-10-29。显式参数的离线培训台账分析及无窗口绘图。"""
from __future__ import annotations
import json,math,re
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import pandas as pd
from PIL import Image,ImageDraw,ImageFont
from openpyxl import load_workbook
_FONT=ContextVar('training_font',default=None)
@contextmanager
def font_context(path: Path | None = None):
    """只使用调用方明确提供的字体，不扫描系统目录。"""
    if path is not None: ImageFont.truetype(str(path),size=16)
    token=_FONT.set(path)
    try: yield
    finally: _FONT.reset(token)
def font(size:int,bold:bool=False):
    """字体由当前显式上下文提供，默认Pillow内置英文字体。"""
    path=_FONT.get()
    return ImageFont.truetype(str(path),size=size) if path else ImageFont.load_default(size=size)

PALETTE = {
    "bg": "#F4F7FB",
    "panel": "#FFFFFF",
    "ink": "#1F2937",
    "muted": "#667085",
    "line": "#D7DEE8",
    "teal": "#11A8A0",
    "blue": "#3A7BD5",
    "amber": "#F5A524",
    "coral": "#F26D5B",
    "green": "#50B47B",
    "purple": "#7C6CE4",
    "navy": "#18263A",
}

@dataclass
class AnalysisResult:
    """保存清洗后的台账、各统计表、候选结果及处理说明。"""
    df: pd.DataFrame
    kpis: dict
    dept_metrics: pd.DataFrame
    trend: pd.DataFrame
    top_hours: pd.DataFrame
    top_fees: pd.DataFrame
    form_share: pd.DataFrame
    quality_bins: pd.DataFrame
    excellent: pd.DataFrame
    latest_month_label: str
    notes: list[str]

def normalize_name(name: object) -> str:
    """归一表头空白，用于检测和显式映射。"""
    return re.sub(r"\s+", "", str(name).replace("\n", "").strip())

def rename_columns(df:pd.DataFrame, aliases:dict[str,str]|None=None)->pd.DataFrame:
    """按显式原列名→规范列名映射重命名，拒绝冲突列。"""
    names=[normalize_name(x) for x in df.columns]
    if len(names)!=len(set(names)):raise ValueError("重复或归一后冲突列")
    mapping={normalize_name(k):str(v) for k,v in (aliases or {}).items()}
    result=df.copy();result.columns=[mapping.get(n,n) for n in names]
    if result.columns.duplicated().any():raise ValueError("映射后列名冲突")
    return result

def to_number(series: pd.Series) -> pd.Series:
    """清除常见计量后缀后转数字；无效值由调用方校验。"""
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("元", "", regex=False)
        .str.replace("小时", "", regex=False)
        .str.strip()
    )
    cleaned = cleaned.mask(cleaned.isin(["", "nan", "None", "_", "-", "—"]))
    return pd.to_numeric(cleaned, errors="coerce")

def parse_month_value(value:object,default_year:int|None=None):
    """校验完整日期或明确月份；缺失年份只使用调用方提供的年份。"""
    if pd.isna(value):return pd.NaT
    if isinstance(value,(datetime,pd.Timestamp)):return pd.Timestamp(value.year,value.month,1)
    text=str(value).strip()
    match=re.fullmatch(r"(\d{4})[-/.年](\d{1,2})(?:[-/.月](\d{1,2})日?)?月?",text)
    if match:
        year,month,day=match.groups()
        try:return pd.Timestamp(datetime(int(year),int(month),int(day or 1)).replace(day=1))
        except ValueError:return pd.NaT
    match=re.fullmatch(r"(\d{1,2})月",text)
    if match and default_year is not None:
        try:return pd.Timestamp(datetime(default_year,int(match[1]),1))
        except ValueError:return pd.NaT
    return pd.NaT

def format_month(value: pd.Timestamp | pd.NaT) -> str:
    """格式化年月，未知日期保持明确标识。"""
    if pd.isna(value):
        return "undated"
    return f"{value.year:04d}-{value.month:02d}"

def clean_ledger(input_path:Path,sheet_name:str|None=None,*,aliases:dict[str,str]|None=None,default_year:int|None=None,summary_labels:tuple[str,...]=(),missing_numbers:str="error") -> tuple[pd.DataFrame,list[str]]:
    """读取明确工作簿；只剔除明确合计标记/全空行，坏数字拒绝。"""
    if missing_numbers not in {"error","zero"}:raise ValueError("missing_numbers必须为error或zero")
    if default_year is not None and (isinstance(default_year,bool) or not isinstance(default_year,int) or not 1<=default_year<=9999):raise ValueError("default_year无效")
    workbook=load_workbook(input_path,read_only=True,data_only=True)
    try:
        sheet=workbook[sheet_name] if sheet_name is not None else workbook.worksheets[0]
        headers=[normalize_name(x) for x in next(sheet.iter_rows(values_only=True),()) if x is not None]
        if len(headers)!=len(set(headers)):raise ValueError("工作簿存在重复表头")
    finally:workbook.close()
    raw=pd.read_excel(input_path,sheet_name=sheet_name if sheet_name is not None else 0,dtype=object)
    df=rename_columns(raw,aliases)
    required=["report_dept","course","date","instructor","participants","hours","satisfaction","fee"]
    missing=[c for c in required if c not in df]
    if missing:raise ValueError("台账缺少必要字段: "+", ".join(missing))
    notes=[];before=len(df)
    summary=df[["course","instructor"]].apply(lambda row:any(str(v).strip() in summary_labels for v in row),axis=1)
    df=df.loc[~summary & ~df.isna().all(axis=1)].copy()
    if before!=len(df):notes.append(f"Removed {before-len(df)} explicitly marked summary/empty rows")
    for col in ["report_dept","course","instructor"]:
        if df[col].isna().any() or df[col].astype(str).str.strip().eq("").any():raise ValueError("缺失标识字段: "+col)
    for col in ["report_dept","instructor_dept","course","form","instructor"]:
        if col in df:df[col]=df[col].fillna("unspecified").astype(str).str.strip()
    if "form" not in df:df["form"]="unspecified"
    for col in ["hours","participants","satisfaction","fee"]:
        numeric=to_number(df[col]);blank=df[col].isna()|df[col].astype(str).str.strip().isin(["","_","-","—"])
        if (numeric.isna() & ~blank).any():raise ValueError("无法解析数字: "+col)
        if numeric.dropna().map(lambda v:not math.isfinite(float(v)) or v<0).any():raise ValueError("数字必须为有限非负数: "+col)
        if col!="satisfaction" and numeric.isna().any():
            if missing_numbers=="error":raise ValueError("数字缺失: "+col)
            notes.append(f"{col}: filled {int(numeric.isna().sum())} missing values with zero")
            numeric=numeric.fillna(0)
        df[col+"_num"]=numeric
    if df["participants_num"].mod(1).ne(0).any():raise ValueError("participants必须为整数")
    df["month"]=[parse_month_value(v,default_year) for v in df["date"]]
    df["month"]=pd.to_datetime(df["month"])
    df["month_label"]=df["month"].map(format_month)
    count=int(df["month"].isna().sum())
    if count:notes.append(f"{count} rows excluded from monthly trend: invalid/undated month")
    return df,notes

def weighted_average(values:pd.Series,weights:pd.Series)->float|None:
    """仅正权重有效评分参与加权；没有样本时返回None而非零分。"""
    valid=values.notna()&weights.fillna(0).gt(0)
    return float((values[valid]*weights[valid]).sum()/weights[valid].sum()) if valid.any() else None

def analyze(df: pd.DataFrame, config: dict, notes: list[str]) -> AnalysisResult:
    """按明确量表及筛选参数生成聚合结果，不推断组织规则。"""
    if not isinstance(config,dict):raise ValueError("配置必须为对象")
    rule=config["selection_rule"];min_score=float(rule["min_satisfaction"]);top_n=rule["top_n"]
    edges=[float(x) for x in config["quality_edges"]]
    if not math.isfinite(min_score) or min_score<0:raise ValueError("筛选阈值必须为有限非负数")
    if isinstance(top_n,bool) or not isinstance(top_n,int) or not 1<=top_n<=5:raise ValueError("top_n限1至5")
    if not 2<=len(edges)<=5 or not all(math.isfinite(x) for x in edges) or any(a>=b for a,b in zip(edges,edges[1:])):raise ValueError("分箱边界必须有限且严格递增，最多4箱")
    if df["satisfaction_num"].dropna().map(lambda x:not edges[0]<=x<=edges[-1]).any():raise ValueError("评分超出明确量表范围")
    latest_month_cfg = config.get("latest_month", "auto")
    valid_months = df.loc[df["month"].notna(), "month"]
    latest_month = valid_months.max() if len(valid_months) else pd.NaT
    if latest_month_cfg and latest_month_cfg != "auto":
        latest_month = pd.Timestamp(latest_month_cfg).replace(day=1)
    latest_month_label = format_month(latest_month)

    kpis = {
        "total_hours": float(df["hours_num"].sum()),
        "total_cost": float(df["fee_num"].sum()),
        "weighted_score": weighted_average(df["satisfaction_num"], df["participants_num"]),
        "participants": int(df["participants_num"].sum()),
        "courses": int(df["course"].nunique()),
        "instructors": int(df["instructor"].nunique()),
    }

    dept_metrics = (
        df.groupby("report_dept", dropna=False)
        .agg(
            total_hours=("hours_num", "sum"),
            total_fee=("fee_num", "sum"),
            participants=("participants_num", "sum"),
            satisfaction=("satisfaction_num", "mean"),
            courses=("course", "nunique"),
        )
        .reset_index()
        .sort_values(["total_hours", "participants"], ascending=False)
    )

    trend = (
        df.loc[df["month"].notna()]
        .groupby("month")
        .agg(hours=("hours_num", "sum"), participants=("participants_num", "sum"))
        .reset_index()
        .sort_values("month")
    )
    trend["label"] = trend["month"].map(lambda x: f"{x.year}-{x.month:02d}")

    top_hours = (
        df.groupby("instructor")
        .agg(total_hours=("hours_num", "sum"), participants=("participants_num", "sum"), satisfaction=("satisfaction_num", "mean"))
        .reset_index()
        .sort_values(["total_hours", "satisfaction"], ascending=False)
    )
    top_fees = (
        df.groupby("instructor")
        .agg(total_fee=("fee_num", "sum"), total_hours=("hours_num", "sum"), satisfaction=("satisfaction_num", "mean"))
        .reset_index()
        .sort_values(["total_fee", "total_hours"], ascending=False)
    )
    form_share = (
        df.groupby("form")
        .agg(hours=("hours_num", "sum"), count=("course", "count"))
        .reset_index()
        .sort_values(["hours", "count"], ascending=False)
    )
    if form_share["hours"].sum() <= 0:
        form_share["share_value"] = form_share["count"]
    else:
        form_share["share_value"] = form_share["hours"]

    labels=[f"{'[' if i==0 else '('}{edges[i]!r}, {edges[i+1]!r}]" for i in range(len(edges)-1)]
    grouped=pd.cut(df["satisfaction_num"],edges,labels=labels,include_lowest=True,right=True)
    quality_bins=pd.DataFrame([{"label":label,"count":int(grouped.eq(label).sum())} for label in labels])
    latest_df = df.loc[df["month"].eq(latest_month)].copy() if not pd.isna(latest_month) else df.iloc[:0].copy()
    eligible = latest_df.loc[latest_df["satisfaction_num"].ge(min_score)].copy()
    excellent = (
        eligible.groupby("instructor")
        .agg(
            total_hours=("hours_num", "sum"),
            participants=("participants_num", "sum"),
            satisfaction=("satisfaction_num", "mean"),
            total_fee=("fee_num", "sum"),
            course=("course", lambda s: s.mode().iat[0] if not s.mode().empty else s.iloc[0]),
            report_dept=("report_dept", lambda s: s.mode().iat[0] if not s.mode().empty else s.iloc[0]),
        )
        .reset_index()
        .sort_values(["total_hours", "satisfaction", "participants"], ascending=False)
        .head(top_n)
    )

    return AnalysisResult(
        df=df,
        kpis=kpis,
        dept_metrics=dept_metrics,
        trend=trend,
        top_hours=top_hours,
        top_fees=top_fees,
        form_share=form_share,
        quality_bins=quality_bins,
        excellent=excellent,
        latest_month_label=latest_month_label,
        notes=notes,
    )

def text_size(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.ImageFont) -> tuple[int, int]:
    """测量文字边界用于字体和布局适配。"""
    bbox = draw.textbbox((0, 0), text, font=fnt)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]

def draw_text(draw: ImageDraw.ImageDraw, xy: tuple[int, int], text: str, fnt: ImageFont.ImageFont, fill: str = PALETTE["ink"]) -> None:
    """向内存画布写入文本，不打开图形窗口。"""
    draw.text(xy, text, font=fnt, fill=fill)

def money_wan(value: float) -> str:
    """将数值缩写为万量级，不进行货币换算。"""
    if abs(value) >= 10000:
        return f"{value / 10000:.1f} x10k"
    return f"{value:,.0f}"

def number_short(value: float) -> str:
    """格式化整数或一位小数的显示值。"""
    if float(value).is_integer():
        return f"{int(value):,}"
    return f"{value:,.1f}"

def panel(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], title: str) -> None:
    """在内存画布绘制一个统计面板。"""
    draw.rounded_rectangle(box, radius=22, fill=PALETTE["panel"], outline="#E5EBF2", width=2)
    draw_text(draw, (box[0] + 28, box[1] + 22), title, font(30, True), PALETTE["ink"])

def card(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], label: str, value: str, accent: str, foot: str = "") -> None:
    """绘制单项指标及说明。"""
    draw.rounded_rectangle(box, radius=24, fill=PALETTE["panel"], outline="#E5EBF2", width=2)
    draw.rounded_rectangle((box[0], box[1], box[0] + 12, box[3]), radius=8, fill=accent)
    draw_text(draw, (box[0] + 34, box[1] + 28), label, font(27), PALETTE["muted"])
    draw_text(draw, (box[0] + 34, box[1] + 70), value, font(44, True), PALETTE["ink"])
    if foot:
        draw_text(draw, (box[0] + 34, box[1] + 122), foot, font(20), PALETTE["muted"])

def chart_bounds(box: tuple[int, int, int, int], top: int = 78, right: int = 42, bottom: int = 54, left: int = 68) -> tuple[int, int, int, int]:
    """取得图表绘图区边界。"""
    return box[0] + left, box[1] + top, box[2] - right, box[3] - bottom

def draw_line_chart(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], trend: pd.DataFrame) -> None:
    """绘制分别归一高度的两项趋势，真实数值另导出表格。"""
    panel(draw, box, "Monthly hours / participants (independent scales)")
    x1, y1, x2, y2 = chart_bounds(box, left=84)
    draw.line((x1, y2, x2, y2), fill=PALETTE["line"], width=2)
    draw.line((x1, y1, x1, y2), fill=PALETTE["line"], width=2)
    if trend.empty:
        draw_text(draw, (x1 + 100, y1 + 120), "No dated rows", font(28), PALETTE["muted"])
        return

    for i in range(5):
        y = y2 - i * (y2 - y1) / 4
        draw.line((x1, y, x2, y), fill="#EEF2F7", width=1)

    labels = trend["label"].tolist()
    hours = trend["hours"].tolist()
    people = trend["participants"].tolist()
    n = len(labels)
    xs = [x1 + (x2 - x1) * (i / max(n - 1, 1)) for i in range(n)]

    def points(values: list[float]) -> list[tuple[float, float]]:
        max_v = max(max(values), 1)
        return [(xs[i], y2 - (float(values[i]) / max_v) * (y2 - y1) * 0.86) for i in range(n)]

    hour_pts = points(hours)
    people_pts = points(people)
    if len(hour_pts) > 1:
        draw.line(hour_pts, fill=PALETTE["blue"], width=6, joint="curve")
        draw.line(people_pts, fill=PALETTE["teal"], width=6, joint="curve")
    for i, (x, label) in enumerate(zip(xs, labels)):
        draw_text(draw, (int(x) - 22, y2 + 18), label, font(20), PALETTE["muted"])
        for pt, color in [(hour_pts[i], PALETTE["blue"]), (people_pts[i], PALETTE["teal"])]:
            draw.ellipse((pt[0] - 8, pt[1] - 8, pt[0] + 8, pt[1] + 8), fill=color)
    draw.rounded_rectangle((x2 - 250, y1 - 42, x2 - 10, y1 - 10), radius=14, fill="#F7FAFC")
    draw.ellipse((x2 - 232, y1 - 33, x2 - 216, y1 - 17), fill=PALETTE["blue"])
    draw_text(draw, (x2 - 206, y1 - 38), "Hours", font(18), PALETTE["muted"])
    draw.ellipse((x2 - 132, y1 - 33, x2 - 116, y1 - 17), fill=PALETTE["teal"])
    draw_text(draw, (x2 - 106, y1 - 38), "People", font(18), PALETTE["muted"])

def draw_bar_chart(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], title: str, df: pd.DataFrame, name_col: str, value_col: str, color: str, suffix: str) -> None:
    """绘制排名条形图，保留姓名和数值。"""
    panel(draw, box, title)
    x1, y1, x2, y2 = chart_bounds(box, top=82, left=42, right=34, bottom=32)
    if df.empty:
        draw_text(draw, (x1 + 40, y1 + 70), "No data", font(26), PALETTE["muted"])
        return
    max_value = max(float(df[value_col].max()), 1.0)
    row_h = (y2 - y1) / max(len(df), 1)
    for i, row in df.reset_index(drop=True).iterrows():
        y = y1 + i * row_h + 6
        name = str(row[name_col])
        value = float(row[value_col])
        bar_w = (x2 - x1 - 280) * value / max_value
        draw_text(draw, (x1, int(y) + 10), name, fit_text(draw,name,170,22,min_size=12), PALETTE["ink"])
        draw.rounded_rectangle((x1 + 180, int(y) + 10, x2 - 90, int(y) + 38), radius=12, fill="#EEF2F7")
        draw.rounded_rectangle((x1 + 180, int(y) + 10, x1 + 180 + bar_w, int(y) + 38), radius=12, fill=color)
        draw_text(draw, (x2 - 78, int(y) + 6), f"{number_short(value)}{suffix}", font(20), PALETTE["muted"])

def draw_pie(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], form_share: pd.DataFrame) -> None:
    """绘制类型占比，超过五类时合并尾项保持正确分母。"""
    panel(draw, box, "Training forms")
    x1, y1, x2, y2 = chart_bounds(box, top=76, left=42, right=34, bottom=28)
    values=form_share["share_value"].tolist();labels=form_share["form"].astype(str).tolist()
    if len(values)>5:values=values[:4]+[sum(values[4:])];labels=labels[:4]+["other"]
    total = sum(values)
    if total <= 0:
        draw_text(draw, (x1 + 40, y1 + 70), "No data", font(26), PALETTE["muted"])
        return
    colors = [PALETTE["teal"], PALETTE["blue"], PALETTE["amber"], PALETTE["coral"], PALETTE["purple"]]
    size = min(y2 - y1, 220)
    cx, cy = x1 + 140, y1 + 120
    pie_box = (cx - size // 2, cy - size // 2, cx + size // 2, cy + size // 2)
    start = -90
    for value, color in zip(values, colors):
        angle = 360 * value / total
        draw.pieslice(pie_box, start=start, end=start + angle, fill=color)
        start += angle
    draw.ellipse((cx - 54, cy - 54, cx + 54, cy + 54), fill=PALETTE["panel"])
    draw_text(draw, (cx - 38, cy - 18), "Forms", font(25, True), PALETTE["ink"])
    lx = x1 + 300
    for i, (label, value) in enumerate(zip(labels, values)):
        y = y1 + 24 + i * 42
        pct = value / total * 100
        draw.rounded_rectangle((lx, y + 7, lx + 22, y + 29), radius=5, fill=colors[i])
        draw_text(draw, (lx + 34, y), f"{label[:12]} {pct:.1f}%", font(21), PALETTE["ink"])

def draw_dept_table(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], dept_metrics: pd.DataFrame) -> None:
    """展示前七个部门的统计，完整结果保留在工作簿。"""
    panel(draw, box, "Department metrics")
    x1, y1, x2, y2 = chart_bounds(box, top=82, left=36, right=36, bottom=26)
    rows = dept_metrics.head(7).reset_index(drop=True)
    headers = ["Department", "Hours", "Cost", "People", "Score"]
    col_x = [x1, x1 + 330, x1 + 480, x1 + 650, x1 + 790]
    for header, x in zip(headers, col_x):
        draw_text(draw, (x, y1), header, font(21, True), PALETTE["muted"])
    draw.line((x1, y1 + 34, x2, y1 + 34), fill=PALETTE["line"], width=2)
    for i, row in rows.iterrows():
        y = y1 + 50 + i * 38
        fill = "#F8FBFD" if i % 2 == 0 else PALETTE["panel"]
        draw.rounded_rectangle((x1 - 8, y - 5, x2, y + 30), radius=10, fill=fill)
        draw_text(draw, (col_x[0], y), str(row["report_dept"])[:15], font(20), PALETTE["ink"])
        draw_text(draw, (col_x[1], y), number_short(float(row["total_hours"])), font(20), PALETTE["ink"])
        draw_text(draw, (col_x[2], y), money_wan(float(row["total_fee"])), font(20), PALETTE["ink"])
        draw_text(draw, (col_x[3], y), number_short(float(row["participants"])), font(20), PALETTE["ink"])
        sat = row["satisfaction"]
        draw_text(draw, (col_x[4], y), "-" if pd.isna(sat) else f"{sat:.2f}", font(20), PALETTE["ink"])

def draw_quality(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], result: AnalysisResult) -> None:
    """展示明确分箱及所选月份候选。"""
    panel(draw, box, "Score bins and selected candidates")
    x1, y1, x2, y2 = chart_bounds(box, top=82, left=44, right=44, bottom=30)
    bins = result.quality_bins
    total = max(int(bins["count"].sum()), 1)
    colors = [PALETTE["green"], PALETTE["teal"], PALETTE["amber"], PALETTE["coral"]]
    bar_x, bar_w = x1, 720
    for i, row in bins.iterrows():
        y = y1 + i * 48
        pct = row["count"] / total
        draw_text(draw, (bar_x, y), row["label"], font(22), PALETTE["ink"])
        draw.rounded_rectangle((bar_x + 130, y + 5, bar_x + 130 + bar_w, y + 31), radius=12, fill="#EEF2F7")
        draw.rounded_rectangle((bar_x + 130, y + 5, bar_x + 130 + bar_w * pct, y + 31), radius=12, fill=colors[i])
        draw_text(draw, (bar_x + 880, y), f"{int(row['count'])} rows", font(22), PALETTE["muted"])

    ex = result.excellent.reset_index(drop=True)
    title_x = x1 + 1080
    draw_text(draw, (title_x, y1), f"{result.latest_month_label} selected: {len(ex)}", font(24, True), PALETTE["ink"])
    if ex.empty:
        draw_text(draw, (title_x, y1 + 56), "No eligible candidates", font(22), PALETTE["muted"])
    else:
        for i, row in ex.iterrows():
            y = y1 + 42 + i * 36
            badge = [PALETTE["amber"], PALETTE["teal"], PALETTE["blue"], PALETTE["purple"], PALETTE["green"]][i]
            draw.rounded_rectangle((title_x, y + 2, title_x + 32, y + 32), radius=8, fill=badge)
            draw_text(draw, (title_x + 9, y + 2), str(i + 1), font(19, True), "#FFFFFF")
            label = f"{row['instructor']}  {number_short(float(row['total_hours']))}h score: {float(row['satisfaction']):.2f}"
            draw_text(draw, (title_x + 46, y), label, font(21), PALETTE["ink"])

def render_dashboard(result: AnalysisResult, output: Path) -> None:
    """生成固定尺寸概览图，不读取外部素材。"""
    width, height = 2400, 1600
    img = Image.new("RGB", (width, height), PALETTE["bg"])
    draw = ImageDraw.Draw(img)
    draw_text(draw, (72, 54), "Offline Training Analytics", font(54, True), PALETTE["ink"])
    subtitle = f"Explicit workbook | latest month: {result.latest_month_label} | rows: {len(result.df)}"
    draw_text(draw, (76, 122), subtitle, font(25), PALETTE["muted"])

    k = result.kpis
    card_y, card_w, card_h, gap = 186, 430, 160, 28
    cards = [
        ("Total hours", f"{number_short(k['total_hours'])}h", PALETTE["blue"], "Recorded hours"),
        ("Total cost", money_wan(k["total_cost"]), PALETTE["teal"], "Recorded cost"),
        ("Weighted score", "N/A" if k["weighted_score"] is None else f"{k['weighted_score']:.2f}", PALETTE["green"], "Positive participant weights only"),
        ("Participants", number_short(k["participants"]), PALETTE["amber"], "Recorded attendance"),
        ("Courses / instructors", f"{k['courses']} / {k['instructors']}", PALETTE["coral"], "Distinct counts"),
    ]
    for i, (label, value, accent, foot) in enumerate(cards):
        x = 72 + i * (card_w + gap)
        card(draw, (x, card_y, x + card_w, card_y + card_h), label, value, accent, foot)

    draw_line_chart(draw, (72, 390, 1180, 830), result.trend)
    draw_dept_table(draw, (1210, 390, 2328, 830), result.dept_metrics)
    draw_bar_chart(draw, (72, 870, 800, 1194), "Teaching hours: top 5", result.top_hours.head(5), "instructor", "total_hours", PALETTE["blue"], "h")
    draw_bar_chart(draw, (832, 870, 1560, 1194), "Recorded cost: top 5", result.top_fees.head(5), "instructor", "total_fee", PALETTE["teal"], "")
    draw_pie(draw, (1592, 870, 2328, 1194), result.form_share)
    draw_quality(draw, (72, 1230, 2328, 1538), result)

    if result.notes:
        note = "；".join(result.notes[:2])
        draw_text(draw, (78, 1550), f"Notes: {note}", font(20), PALETTE["muted"])

    output.parent.mkdir(parents=True, exist_ok=True)
    img.save(output, quality=95)

def gradient_background(width: int, height: int) -> Image.Image:
    """用像素渐变生成背景，不加载现有图片。"""
    img = Image.new("RGB", (width, height))
    pix = img.load()
    top = (18, 38, 58)
    bottom = (17, 154, 148)
    for y in range(height):
        t = y / (height - 1)
        for x in range(width):
            side = 0.08 * math.sin((x / width) * math.pi)
            r = int(top[0] * (1 - t) + bottom[0] * t + 20 * side)
            g = int(top[1] * (1 - t) + bottom[1] * t + 28 * side)
            b = int(top[2] * (1 - t) + bottom[2] * t + 16 * side)
            pix[x, y] = (r, g, b)
    return img

def fit_text(draw: ImageDraw.ImageDraw, text: str, max_width: int, size: int, bold: bool = False, min_size: int = 28) -> ImageFont.ImageFont:
    """在限定字号内缩小文本以适应给定宽度。"""
    current = size
    while current >= min_size:
        fnt = font(current, bold)
        if text_size(draw, text, fnt)[0] <= max_width:
            return fnt
        current -= 2
    return font(min_size, bold)

def wrap_text(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.ImageFont, max_width: int) -> list[str]:
    """按实际字体宽度换行。"""
    lines: list[str] = []
    current = ""
    for char in text:
        if text_size(draw, current + char, fnt)[0] <= max_width:
            current += char
        else:
            if current:
                lines.append(current)
            current = char
    if current:
        lines.append(current)
    return lines

def short_copy(row: pd.Series) -> str:
    """根据候选实际指标生成短说明。"""
    hours = number_short(float(row["total_hours"]))
    sat = float(row["satisfaction"])
    people = int(row["participants"])
    copy = f"Score{sat:.2f}, {people} people, {hours}h"
    return copy[:30]

def render_poster(result: AnalysisResult, config: dict, output: Path) -> None:
    """仅在存在真实计算候选时生成海报。"""
    width, height = 1080, 1620
    img = gradient_background(width, height)
    draw = ImageDraw.Draw(img, "RGBA")

    for i in range(18):
        x = 80 + i * 58
        draw.line((x, 190, x + 260, 640), fill=(255, 255, 255, 18), width=2)
    for radius in [380, 520, 660]:
        draw.arc((width - radius, 120, width + radius // 2, 120 + radius), 195, 332, fill=(245, 165, 36, 70), width=3)
    for i, color in enumerate([(58, 123, 213, 80), (17, 168, 160, 95), (245, 165, 36, 85)]):
        draw.ellipse((760 - i * 80, 108 + i * 90, 1130 - i * 40, 478 + i * 90), outline=color, width=3)

    draw.rounded_rectangle((70, 66, 314, 116), radius=24, fill=(255, 255, 255, 34), outline=(255, 255, 255, 65))
    draw_text(draw, (94, 78), "Selected training", font(25, True), "#FFFFFF")

    if result.excellent.empty:raise ValueError("没有满足条件的候选，不生成虚构获选海报")
    top_rows=result.excellent.reset_index(drop=True)
    champion=top_rows.iloc[0]

    course = str(champion["course"])
    title_template = config.get("poster", {}).get("title_template", "Selected: {course}")
    title = title_template.format(course=course[:18])
    title_font = fit_text(draw, title, 920, 64, True, 42)
    title_lines = wrap_text(draw, title, title_font, 920)
    y = 178
    for line in title_lines[:2]:
        draw_text(draw, (72, y), line, title_font, "#FFFFFF")
        y += text_size(draw, line, title_font)[1] + 14

    slogan = config.get("poster", {}).get("slogan", "Explicit selection, reviewed inputs")
    draw_text(draw, (76, y + 4), slogan, font(30), "#DDEFF2")

    draw.rounded_rectangle((72, 472, 1008, 744), radius=34, fill=(255, 255, 255, 242))
    draw_text(draw, (112, 514), str(champion["instructor"]), font(72, True), PALETTE["navy"])
    dept_line = f"{champion.get('report_dept', '')} | {result.latest_month_label}"
    draw_text(draw, (116, 612), dept_line, font(30), PALETTE["muted"])
    copy = short_copy(champion)
    draw.rounded_rectangle((112, 658, 770, 710), radius=22, fill="#EAF8F7")
    draw_text(draw, (138, 668), copy, font(29, True), PALETTE["teal"])

    metrics = [
        ("Score", f"{float(champion['satisfaction']):.2f}", PALETTE["green"]),
        ("participants", f"{int(champion['participants'])}", PALETTE["amber"]),
        ("Recorded hours", f"{number_short(float(champion['total_hours']))}h", PALETTE["blue"]),
    ]
    metric_y = 790
    metric_w = 292
    for i, (label, value, accent) in enumerate(metrics):
        x = 72 + i * (metric_w + 30)
        draw.rounded_rectangle((x, metric_y, x + metric_w, metric_y + 180), radius=30, fill=(255, 255, 255, 235))
        draw.rounded_rectangle((x + 28, metric_y + 26, x + 72, metric_y + 70), radius=14, fill=accent)
        draw_text(draw, (x + 28, metric_y + 86), label, font(28), PALETTE["muted"])
        val_font = fit_text(draw, value, metric_w - 54, 50, True, 32)
        draw_text(draw, (x + 28, metric_y + 122), value, val_font, PALETTE["navy"])

    draw.rounded_rectangle((72, 1028, 1008, 1396), radius=34, fill=(255, 255, 255, 232))
    draw_text(draw, (112, 1068), f"Selected instructors: {len(top_rows)}", font(36, True), PALETTE["navy"])
    for i, row in top_rows.head(5).iterrows():
        y = 1130 + i * 52
        colors = [PALETTE["amber"], PALETTE["teal"], PALETTE["blue"], PALETTE["purple"], PALETTE["green"]]
        draw.rounded_rectangle((112, y, 156, y + 42), radius=12, fill=colors[i % len(colors)])
        draw_text(draw, (126, y + 2), str(i + 1), font(25, True), "#FFFFFF")
        name = str(row["instructor"])
        item = f"{name}  {number_short(float(row['total_hours']))}h score: {float(row['satisfaction']):.2f}"
        draw_text(draw, (176, y + 2), item, fit_text(draw,item,510,28,True,18), PALETTE["ink"])
        draw_text(draw, (720, y + 5), str(row["course"]), fit_text(draw,str(row["course"]),240,22,min_size=14), PALETTE["muted"])

    draw.rounded_rectangle((72, 1442, 1008, 1532), radius=30, fill=(255, 255, 255, 30), outline=(255, 255, 255, 80))
    footer = config.get("poster", {}).get("footer", "Offline training analytics")
    draw_text(draw, (108, 1472), f"{footer} | score >= {config['selection_rule']['min_satisfaction']}; ranked by hours", font(25), "#E8FBFA")

    output.parent.mkdir(parents=True, exist_ok=True)
    img.convert("RGB").save(output, quality=96)

def save_summary(result: AnalysisResult, output: Path) -> None:
    """导出严格JSON摘要，未知评分使用null。"""
    payload = {
        "latest_month": result.latest_month_label,
        "kpis": result.kpis,
        "notes": result.notes,
        "selected_instructors": result.excellent.to_dict(orient="records"),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, allow_nan=False)

def save_excellent_csv(result: AnalysisResult, output: Path) -> None:
    """导出候选列表并保护公式型文本。"""
    output.parent.mkdir(parents=True, exist_ok=True)
    _safe_frame(result.excellent).to_csv(output, index=False, encoding="utf-8-sig")

def _safe_frame(frame:pd.DataFrame)->pd.DataFrame:
    """导出表格文本保留为文本，避免被表格软件当公式解释。"""
    return frame.map(lambda v:"'"+v if isinstance(v,str) and v.lstrip().startswith(("=","+","-","@")) else v)

def save_analysis_workbook(result:AnalysisResult,output:Path)->None:
    """导出实际聚合工作表，不附原模板或外部链接。"""
    tables={"kpis":pd.DataFrame([result.kpis]),"departments":result.dept_metrics,"monthly":result.trend,"hours_rank":result.top_hours,"cost_rank":result.top_fees,"forms":result.form_share,"score_bins":result.quality_bins,"selected":result.excellent}
    with pd.ExcelWriter(output,engine="openpyxl") as writer:
        for name,frame in tables.items():_safe_frame(frame).to_excel(writer,sheet_name=name,index=False)

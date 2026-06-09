# -*- coding: utf-8 -*-
"""
NailVesta · 深深达名单生成器  (zero-dependency 版)
=====================================================
输入两份文件：
  1) 深度达人 List（既有名单，键 = handle）
  2) TikTok Creator 数据（某时段，键 = Creator username，含 GMV / orders / items 等）

逻辑：按清洗后的 handle 比对，过滤掉这段时间「已经没在出单」的深度达人，
      保留仍在持续出单的，组成一份更高级的「深深达」名单。

本版自带纯标准库的 .xlsx 读写（不依赖 openpyxl），
所以只要环境有 streamlit（已内含 pandas）即可运行，无需 requirements.txt。
"""

import io
import re
import zipfile
import xml.etree.ElementTree as ET
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="NailVesta · 深深达名单生成器", page_icon="💅", layout="wide")

# =========================================================================== #
#  纯标准库 .xlsx 读写
# =========================================================================== #
def _col_to_idx(ref):
    m = re.match(r"([A-Z]+)", ref or "")
    if not m:
        return None
    s, n = m.group(1), 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def read_xlsx(file_bytes):
    """读取第一个工作表，回传 (headers, rows)。"""
    z = zipfile.ZipFile(io.BytesIO(file_bytes))
    names = z.namelist()
    shared = []
    if "xl/sharedStrings.xml" in names:
        for _, el in ET.iterparse(io.BytesIO(z.read("xl/sharedStrings.xml"))):
            if el.tag.endswith("}si") or el.tag == "si":
                shared.append("".join(t.text or "" for t in el.iter()
                                      if (t.tag.endswith("}t") or t.tag == "t")))
                el.clear()
    sheets = sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
    sheet = sheets[0] if sheets else next(n for n in names if n.startswith("xl/worksheets/"))
    rows, max_cols = [], 0
    for _, el in ET.iterparse(io.BytesIO(z.read(sheet))):
        if not (el.tag.endswith("}row") or el.tag == "row"):
            continue
        cells, auto = {}, 0
        for c in el:
            if not (c.tag.endswith("}c") or c.tag == "c"):
                continue
            idx = _col_to_idx(c.attrib.get("r", ""))
            if idx is None:
                idx = auto
            auto = idx + 1
            t = c.attrib.get("t")
            v_text = is_text = None
            for ch in c:
                if ch.tag.endswith("}v") or ch.tag == "v":
                    v_text = ch.text
                elif ch.tag.endswith("}is") or ch.tag == "is":
                    is_text = "".join(x.text or "" for x in ch.iter()
                                      if (x.tag.endswith("}t") or x.tag == "t"))
            if t == "s" and v_text is not None:
                val = shared[int(v_text)]
            elif t == "inlineStr" and is_text is not None:
                val = is_text
            elif v_text is not None:
                val = v_text
                if t is None or t == "n":
                    try:
                        f = float(v_text)
                        val = int(f) if f.is_integer() else f
                    except (TypeError, ValueError):
                        pass
            else:
                val = None
            cells[idx] = val
            if idx + 1 > max_cols:
                max_cols = idx + 1
        rows.append(cells)
        el.clear()
    table = [[r.get(i) for i in range(max_cols)] for r in rows]
    headers = [("" if v is None else str(v)) for v in table[0]] if table else []
    return headers, table[1:]


def _esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _idx_to_col(i):
    s, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


_STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
           '<fonts count="2"><font><sz val="11"/><name val="Arial"/></font>'
           '<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Arial"/></font></fonts>'
           '<fills count="3"><fill><patternFill patternType="none"/></fill>'
           '<fill><patternFill patternType="gray125"/></fill>'
           '<fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill></fills>'
           '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
           '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
           '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
           '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1">'
           '<alignment horizontal="center"/></xf></cellXfs>'
           '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
           '</styleSheet>')


def write_xlsx(sheets):
    """sheets: list of (name, headers, rows_of_lists) -> bytes"""
    buf = io.BytesIO()
    z = zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED)
    sheet_ct = wb_sheets = wb_rels = ""
    for i, (name, headers, rows) in enumerate(sheets, 1):
        sheet_ct += (f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                     f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        wb_sheets += f'<sheet name="{_esc(name)}" sheetId="{i}" r:id="rId{i}"/>'
        wb_rels += (f'<Relationship Id="rId{i}" '
                    f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                    f'Target="worksheets/sheet{i}.xml"/>')
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
        for r, row in enumerate([headers] + rows, 1):
            out.append(f'<row r="{r}">')
            for c, val in enumerate(row):
                ref = f"{_idx_to_col(c)}{r}"
                style = ' s="1"' if r == 1 else ''
                if val is None or val == "":
                    out.append(f'<c r="{ref}"{style}/>')
                elif isinstance(val, bool):
                    out.append(f'<c r="{ref}"{style}><v>{1 if val else 0}</v></c>')
                elif isinstance(val, (int, float)):
                    out.append(f'<c r="{ref}"{style}><v>{val}</v></c>')
                else:
                    out.append(f'<c r="{ref}" t="inlineStr"{style}>'
                               f'<is><t xml:space="preserve">{_esc(val)}</t></is></c>')
            out.append('</row>')
        out.append('</sheetData></worksheet>')
        z.writestr(f"xl/worksheets/sheet{i}.xml", "".join(out))
    z.writestr("[Content_Types].xml",
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
               '<Default Extension="xml" ContentType="application/xml"/>'
               '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
               '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
               + sheet_ct + '</Types>')
    z.writestr("_rels/.rels",
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    z.writestr("xl/workbook.xml",
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
               '<sheets>' + wb_sheets + '</sheets></workbook>')
    z.writestr("xl/_rels/workbook.xml.rels",
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               + wb_rels +
               '<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    z.writestr("xl/styles.xml", _STYLES)
    z.close()
    return buf.getvalue()


@st.cache_data(show_spinner="读取文件中…")
def load_df(file_bytes):
    headers, rows = read_xlsx(file_bytes)
    return pd.DataFrame(rows, columns=headers)


def df_to_rows(df):
    def conv(v):
        if v is None:
            return None
        if isinstance(v, float) and v != v:        # NaN
            return None
        if isinstance(v, np.integer):
            return int(v)
        if isinstance(v, np.floating):
            f = float(v)
            return None if f != f else f
        if isinstance(v, (np.bool_, bool)):
            return bool(v)
        return v
    return [[conv(v) for v in rec] for rec in df.itertuples(index=False, name=None)]


# =========================================================================== #
#  清洗 / 工具
# =========================================================================== #
def clean_handle(h, strip_paren=True, strip_suffix=True):
    if h is None or (isinstance(h, float) and h != h):
        return None
    s = str(h).strip().lower()
    if strip_paren:
        s = re.sub(r"（.*?）", "", s)
        s = re.sub(r"\(.*?\)", "", s)
    s = s.lstrip("@")
    if strip_suffix:
        s = re.sub(r"-\d+$", "", s)
    return s.strip() or None


def to_num(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False)
        .replace({"--": None, "": None, "nan": None, "None": None}),
        errors="coerce",
    ).fillna(0)


def first_non_null(s):
    s = s.dropna()
    return s.iloc[0] if len(s) else None


def pick(cols, candidates):
    for c in candidates:
        if c in cols:
            return c
    return None


DEEP_KEY_CANDIDATES = ["handle", "Handle", "Creator username", "username", "用户名"]
TT_KEY_CANDIDATES = ["Creator username", "creator username", "username", "handle", "用户名"]
GMV_CANDIDATES = ["Affiliate GMV", "GMV", "affiliate gmv"]
ORDER_CANDIDATES = ["Affiliate orders", "Orders", "affiliate orders"]
ITEM_CANDIDATES = ["Items sold", "items sold", "Affiliate items sold"]
REFUND_GMV_CANDIDATES = ["Affiliate refunded GMV", "Refunded GMV"]
FOLLOWERS_CANDIDATES = ["Affiliate followers", "Followers"]
DEEP_CARRY = ["Level", "评级", "深度合作Status", "深度前出单数", "终止合作",
              "名字", "Email", "Phone", "Tiktok Link", "aff link"]

# =========================================================================== #
#  侧边栏
# =========================================================================== #
st.sidebar.header("① 上传文件")
deep_file = st.sidebar.file_uploader("深度达人 List (.xlsx)", type=["xlsx"])
tt_file = st.sidebar.file_uploader("TikTok Creator 数据 (.xlsx)", type=["xlsx"])

st.sidebar.header("② Handle 清洗")
strip_paren = st.sidebar.checkbox("去除括号内容 （改名…）", value=True)
strip_suffix = st.sidebar.checkbox("去除结尾 -数字 后缀（gicellex-2 → gicellex）", value=True)

st.sidebar.header("③ 出单门槛")
metric_label = st.sidebar.selectbox(
    "判定「仍在出单」的指标",
    ["出单数 (Affiliate orders)", "GMV (Affiliate GMV)", "件数 (Items sold)"], index=0)
metric_key = {"出单数": "orders", "GMV": "gmv", "件数": "items"}[metric_label.split(" ")[0]]
if metric_key == "gmv":
    min_val = st.sidebar.number_input("最小 GMV（含）", min_value=0.0, value=1.0, step=10.0)
elif metric_key == "items":
    min_val = st.sidebar.number_input("最小件数（含）", min_value=0, value=1, step=1)
else:
    min_val = st.sidebar.number_input("最小出单数（含）", min_value=0, value=1, step=1)
exclude_terminated = st.sidebar.checkbox("排除已标记『终止合作』的达人", value=False)

st.title("💅 NailVesta · 深深达名单生成器")
st.caption("过滤掉这段时间已经没在出单的深度达人，产出一份更精炼的深深达名单。")

if not deep_file or not tt_file:
    st.info("请在左侧分别上传 **深度达人 List** 与 **TikTok Creator 数据** 两份 .xlsx 文件。")
    st.stop()

# =========================================================================== #
#  读取 + 键名/指标对应
# =========================================================================== #
deep_raw = load_df(deep_file.getvalue())
tt_raw = load_df(tt_file.getvalue())

deep_key = pick(deep_raw.columns, DEEP_KEY_CANDIDATES)
tt_key = pick(tt_raw.columns, TT_KEY_CANDIDATES)
with st.sidebar.expander("⚙️ 键名对应（自动侦测，可手动改）", expanded=False):
    deep_key = st.selectbox("深度名单 handle 栏", list(deep_raw.columns),
                            index=list(deep_raw.columns).index(deep_key) if deep_key else 0)
    tt_key = st.selectbox("TikTok username 栏", list(tt_raw.columns),
                          index=list(tt_raw.columns).index(tt_key) if tt_key else 0)

gmv_col = pick(tt_raw.columns, GMV_CANDIDATES)
order_col = pick(tt_raw.columns, ORDER_CANDIDATES)
item_col = pick(tt_raw.columns, ITEM_CANDIDATES)
refund_col = pick(tt_raw.columns, REFUND_GMV_CANDIDATES)
fol_col = pick(tt_raw.columns, FOLLOWERS_CANDIDATES)

# =========================================================================== #
#  清洗 + 去重 + 聚合
# =========================================================================== #
deep = deep_raw.copy()
deep["__key"] = deep[deep_key].map(lambda x: clean_handle(x, strip_paren, strip_suffix))
deep = deep[deep["__key"].notna()]
agg = {c: ("max" if c == "深度前出单数" else first_non_null)
       for c in DEEP_CARRY if c in deep.columns}
deep_u = deep.groupby("__key").agg(agg).reset_index().rename(columns={"__key": "handle"})

tt = tt_raw.copy()
tt["__key"] = tt[tt_key].map(lambda x: clean_handle(x, strip_paren, strip_suffix))
tt = tt[tt["__key"].notna()]
tt["gmv"] = to_num(tt[gmv_col]) if gmv_col else 0
tt["orders"] = to_num(tt[order_col]) if order_col else 0
tt["items"] = to_num(tt[item_col]) if item_col else 0
tt["refund_gmv"] = to_num(tt[refund_col]) if refund_col else 0
tt["followers"] = to_num(tt[fol_col]) if fol_col else 0
tt_g = tt.groupby("__key").agg(
    gmv=("gmv", "sum"), orders=("orders", "sum"), items=("items", "sum"),
    refund_gmv=("refund_gmv", "sum"), followers=("followers", "max"),
).reset_index().rename(columns={"__key": "handle"})

merged = deep_u.merge(tt_g, on="handle", how="left")
for c in ["gmv", "orders", "items", "refund_gmv", "followers"]:
    merged[c] = merged[c].fillna(0)
merged["in_period"] = merged["handle"].isin(set(tt_g["handle"]))
merged["退款率"] = (merged["refund_gmv"] / merged["gmv"]).where(merged["gmv"] > 0, 0).round(3)

qualifies = merged[metric_key] >= min_val
if exclude_terminated and "终止合作" in merged.columns:
    terminated = pd.to_numeric(merged["终止合作"], errors="coerce").fillna(0) > 0
    qualifies = qualifies & ~terminated
merged["深深达"] = qualifies

kept = merged[merged["深深达"]].sort_values("gmv", ascending=False).reset_index(drop=True)
dropped = merged[~merged["深深达"]].sort_values("gmv", ascending=False).reset_index(drop=True)

# =========================================================================== #
#  概览
# =========================================================================== #
c1, c2, c3, c4 = st.columns(4)
c1.metric("深度达人总数", f"{len(merged):,}")
c2.metric("时段内有数据", f"{int(merged['in_period'].sum()):,}")
c3.metric("✅ 深深达（保留）", f"{len(kept):,}")
c4.metric("❌ 已剔除", f"{len(dropped):,}")
c5, c6, c7 = st.columns(3)
c5.metric("深深达 总 GMV", f"${kept['gmv'].sum():,.0f}")
c6.metric("深深达 总出单数", f"{int(kept['orders'].sum()):,}")
c7.metric("深深达 平均 GMV", f"${kept['gmv'].mean():,.0f}" if len(kept) else "$0")
st.caption(
    f"门槛：**{metric_label} ≥ {min_val}**"
    + ("，且排除『终止合作』" if exclude_terminated else "")
    + f"。未出现在 TikTok 数据中的 {int((~merged['in_period']).sum())} 位达人视为未出单、一律剔除。")

# =========================================================================== #
#  展示
# =========================================================================== #
display_cols = ["handle", "Level", "深度前出单数", "gmv", "orders", "items",
                "退款率", "followers", "名字", "Email", "Phone", "Tiktok Link", "aff link"]
display_cols = [c for c in display_cols if c in merged.columns]
rename_map = {"gmv": "时段GMV", "orders": "时段出单数", "items": "时段件数", "followers": "粉丝数"}

tab1, tab2, tab3 = st.tabs(["✅ 深深达名单", "❌ 已剔除", "📊 GMV 分布"])
with tab1:
    st.dataframe(kept[display_cols].rename(columns=rename_map),
                 use_container_width=True, hide_index=True)
with tab2:
    reason = pd.Series("低于门槛", index=dropped.index)
    reason[~dropped["in_period"]] = "时段内无数据/未出单"
    if exclude_terminated and "终止合作" in dropped.columns:
        reason[pd.to_numeric(dropped["终止合作"], errors="coerce").fillna(0) > 0] = "已终止合作"
    d = dropped[display_cols].rename(columns=rename_map).copy()
    d.insert(1, "剔除原因", reason.values)
    st.dataframe(d, use_container_width=True, hide_index=True)
with tab3:
    pos = kept[kept["gmv"] > 0]
    if len(pos):
        st.bar_chart(pos.set_index("handle")["gmv"].head(30))
        st.caption("Top 30 深深达（按时段 GMV）")
    else:
        st.write("没有 GMV > 0 的达人。")

# =========================================================================== #
#  导出
# =========================================================================== #
def build_excel():
    kept_d = kept[display_cols].rename(columns=rename_map)
    drop_d = dropped[display_cols].rename(columns=rename_map)
    reason = pd.Series("低于门槛", index=dropped.index)
    reason[~dropped["in_period"]] = "时段内无数据/未出单"
    if exclude_terminated and "终止合作" in dropped.columns:
        reason[pd.to_numeric(dropped["终止合作"], errors="coerce").fillna(0) > 0] = "已终止合作"
    drop_d = drop_d.copy(); drop_d.insert(1, "剔除原因", reason.values)
    full = merged.drop(columns=["深深达"], errors="ignore")
    return write_xlsx([
        ("深深达名单", list(kept_d.columns), df_to_rows(kept_d)),
        ("已剔除", list(drop_d.columns), df_to_rows(drop_d)),
        ("全部比对", list(full.columns), df_to_rows(full)),
    ])


st.sidebar.header("④ 导出")
st.sidebar.download_button(
    "📥 下载结果 Excel",
    data=build_excel(),
    file_name="NailVesta_深深达名单.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

"""플랫폼사업부 통합 현황 대시보드

데일리 업데이트 엑셀(dailyupdate) 한 파일만 올리면
  1) 매출 · 쿠팡 PO : 상품군별 목표 달성 카드, 상품군/채널별 월 매출표, 납품 예정
  2) 재고 · 결품 : 쿠팡 VF 재고, 소진 예상일, 납품 데드라인, SKU 상세
  3) 데이터 설정 · 점검 : 상품군 목표, 재고 계산 설정, 데이터 이상 확인
을 보여 줍니다.

읽는 시트: RAW_매입, RAW_출고(벤플), SKU별 재고 현황(쿠팡), 요약(목표매출 행), 상품군목표(선택)
"""

import io
import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="플랫폼사업부 통합 현황", page_icon="📊", layout="wide")

# ─────────────────────────────────────────────
# 설정
# ─────────────────────────────────────────────
BUY_SHEET = "RAW_매입"
OUT_SHEET = "RAW_출고(벤플)"
STOCK_SHEET = "SKU별 재고 현황(쿠팡)"
SUMMARY_SHEET = "요약"
GROUP_TARGET_SHEET = "상품군목표"  # 선택: 월 | 상품군 | 목표(원)

TYPE_ORDER = ["볼", "파쇄", "두부", "마이티펫", "장작", "필름", "리빙용품", "기저귀", "필터", "트리"]
# 쿠팡 납품 리드타임(일) 기본값: 쿠팡 PO를 받고 VF에 입고되기까지. 사이드바에서 바꿀 수 있음
DEFAULT_LEAD = {t: 5 for t in TYPE_ORDER}
TODAY = datetime.now(ZoneInfo("Asia/Seoul")).date()


def norm(s) -> str:
    return re.sub(r"\s+", "", str(s)).lower()


def type_key(t) -> int:
    return TYPE_ORDER.index(t) if t in TYPE_ORDER else len(TYPE_ORDER)


def sort_types(types):
    return sorted([t for t in types if pd.notna(t)], key=type_key)


def won_m(x) -> str:
    """원 → 백만원 문자열"""
    if x is None or pd.isna(x):
        return "-"
    return f"{x / 1e6:,.1f}백만"


def pct(x) -> str:
    if x is None or pd.isna(x) or np.isinf(x):
        return "-"
    return f"{x * 100:+.1f}%"


# ─────────────────────────────────────────────
# 엑셀 읽기
# ─────────────────────────────────────────────
def parse_stock(xls: pd.ExcelFile):
    """SKU별 재고 현황(쿠팡) 시트 → (DataFrame, 현재고 기준일)"""
    if STOCK_SHEET not in xls.sheet_names:
        return None, None
    raw = pd.read_excel(xls, STOCK_SHEET, header=None)

    h = None
    for i in range(min(20, len(raw))):
        vals = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        if "sku id" in vals:
            h = i
            break
    if h is None:
        return None, None

    hdr = [str(v).strip() if pd.notna(v) else "" for v in raw.iloc[h].tolist()]
    low = [x.lower() for x in hdr]
    start = low.index("sku id")
    end = low.index("sku id", start + 1) if "sku id" in low[start + 1:] else len(hdr)

    def find(pred):
        for j in range(start, end):
            if pred(hdr[j]):
                return j
        return None

    c_name = find(lambda x: x == "품명")
    c_type = find(lambda x: x == "유형")
    c_cur = find(lambda x: x.startswith("현재고"))
    c_in = find(lambda x: x == "입고예정")
    if None in (c_name, c_type, c_cur):
        return None, None

    # 기준일: '현재고(11/17)' 의 날짜, 연도는 시트 안의 날짜 셀에서
    year = TODAY.year
    for i in range(h):
        for v in raw.iloc[i].tolist():
            if isinstance(v, (datetime, pd.Timestamp)):
                year = v.year
    m = re.search(r"(\d{1,2})\s*/\s*(\d{1,2})", hdr[c_cur])
    stock_date = pd.Timestamp(year, int(m.group(1)), int(m.group(2))) if m else pd.Timestamp(TODAY)

    body = raw.iloc[h + 1:]
    df = pd.DataFrame({
        "sku id": pd.to_numeric(body.iloc[:, start], errors="coerce"),
        "품명": body.iloc[:, c_name].astype("string").str.strip(),
        "유형": body.iloc[:, c_type].astype("string").str.strip(),
        "현재고": pd.to_numeric(body.iloc[:, c_cur], errors="coerce"),
        "입고예정": pd.to_numeric(body.iloc[:, c_in], errors="coerce") if c_in is not None else 0,
    })
    df = df[df["sku id"].notna()].copy()
    df["sku id"] = df["sku id"].astype("int64")
    df[["현재고", "입고예정"]] = df[["현재고", "입고예정"]].fillna(0)
    return df.reset_index(drop=True), stock_date


def parse_summary(xls: pd.ExcelFile, fallback_year: int):
    """요약 시트의 '목표매출', '전년(YYYY)' 행 → {(연, 월): 금액}"""
    target, prior = {}, {}
    if SUMMARY_SHEET not in xls.sheet_names:
        return target, prior
    raw = pd.read_excel(xls, SUMMARY_SHEET, header=None)

    def month_cols(i):
        cols = {}
        for j, v in enumerate(raw.iloc[i].tolist()):
            mm = re.match(r"^\s*(\d{1,2})\s*월", str(v))
            if mm:
                cols[j] = int(mm.group(1))
        return cols

    for i in range(len(raw)):
        row = [str(v).strip() if pd.notna(v) else "" for v in raw.iloc[i].tolist()]
        if "목표매출" not in row:
            continue
        # 바로 위쪽에서 월 머리글 행 찾기
        cols = {}
        for k in range(i - 1, max(-1, i - 4), -1):
            cols = month_cols(k)
            if len(cols) >= 6:
                break
        if not cols:
            continue
        # 전년 행 / 연도
        year = fallback_year
        prior_row = None
        for k in range(i + 1, min(len(raw), i + 6)):
            rr = [str(v).strip() if pd.notna(v) else "" for v in raw.iloc[k].tolist()]
            for cell in rr:
                py = re.match(r"^전년\s*\((\d{4})\)", cell)
                if py:
                    year = int(py.group(1)) + 1
                    prior_row = k
        for j, mth in cols.items():
            v = pd.to_numeric(raw.iat[i, j], errors="coerce")
            if pd.notna(v):
                target[(year, mth)] = float(v)
            if prior_row is not None:
                pv = pd.to_numeric(raw.iat[prior_row, j], errors="coerce")
                if pd.notna(pv):
                    prior[(year - 1, mth)] = float(pv)
        break
    return target, prior


@st.cache_data(show_spinner="엑셀 읽는 중...")
def load_workbook(data: bytes):
    xls = pd.ExcelFile(io.BytesIO(data))
    missing = [s for s in (BUY_SHEET, OUT_SHEET) if s not in xls.sheet_names]
    if missing:
        raise ValueError(f"필수 시트가 없습니다: {', '.join(missing)}")

    # 매입(=쿠팡·네이버·이마트 수주) ── 매출 기준
    buy = pd.read_excel(xls, BUY_SHEET)
    buy.columns = [str(c).strip() for c in buy.columns]
    for c in ["품목", "공급채널", "구분", "유형", "입고처", "주차", "비고"]:
        if c in buy.columns:
            buy[c] = buy[c].astype("string").str.strip()
    buy["날짜"] = pd.to_datetime(buy["날짜"], errors="coerce")
    buy["발주일"] = pd.to_datetime(buy["발주일"], errors="coerce") if "발주일" in buy.columns else pd.NaT
    for c in ["수량(BOX)", "수량(톤)", "금액(원)"]:
        buy[c] = pd.to_numeric(buy[c], errors="coerce") if c in buy.columns else np.nan
    buy = buy[buy["날짜"].notna() & buy["품목"].notna()].copy()
    buy["금액(원)"] = buy["금액(원)"].fillna(0)

    # 출고(벤플)
    out = pd.read_excel(xls, OUT_SHEET)
    out.columns = [str(c).strip() for c in out.columns]
    out["날짜"] = pd.to_datetime(out["날짜"], errors="coerce")
    out["sku id"] = pd.to_numeric(out["sku id"], errors="coerce")
    out["출고량"] = pd.to_numeric(out["출고량"], errors="coerce").fillna(0)
    out["품명"] = out["품명"].astype("string").str.strip()
    out["유형"] = out["유형"].astype("string").str.strip()
    out = out[out["날짜"].notna() & out["sku id"].notna()].copy()
    out["sku id"] = out["sku id"].astype("int64")

    stock, stock_date = parse_stock(xls)
    fallback_year = int(buy["날짜"].max().year)
    target, prior = parse_summary(xls, fallback_year)

    # 품명 → SKU 매핑 (RAW_매입에는 SKU가 없어 이름+유형으로 연결)
    sku_map = {}
    sources = [out[["sku id", "품명", "유형"]]]
    if stock is not None:
        sources.insert(0, stock[["sku id", "품명", "유형"]])
    for src in sources:
        for sku, name, typ in src.dropna().itertuples(index=False):
            sku_map.setdefault((norm(name), str(typ)), int(sku))
    buy["sku id"] = [sku_map.get((norm(n), str(t))) for n, t in zip(buy["품목"], buy["유형"])]

    # SKU 대표 이름: 재고현황 시트 이름 > 출고 최신 이름
    names = out.sort_values("날짜").groupby("sku id").agg(품명=("품명", "last"), 유형=("유형", "last"))
    if stock is not None:
        names = stock.set_index("sku id")[["품명", "유형"]].combine_first(names)

    # 선택 시트: 상품군목표 (월, 상품군, 목표)
    grp_target = {}
    if GROUP_TARGET_SHEET in xls.sheet_names:
        gt = pd.read_excel(xls, GROUP_TARGET_SHEET)
        gt.columns = [str(c).strip() for c in gt.columns]
        if {"월", "상품군", "목표"} <= set(gt.columns):
            for mo, grp, v in gt[["월", "상품군", "목표"]].itertuples(index=False):
                try:
                    per = pd.Period(pd.to_datetime(str(mo).strip()[:7] if not isinstance(mo, (datetime, pd.Timestamp)) else mo), "M")
                except Exception:  # noqa: BLE001
                    continue
                v = pd.to_numeric(v, errors="coerce")
                if pd.notna(v) and pd.notna(grp):
                    grp_target[(str(per), str(grp).strip())] = float(v)

    return buy, out, stock, stock_date, target, prior, names, grp_target


# ─────────────────────────────────────────────
# 상품군 / 표시 형식
# ─────────────────────────────────────────────
GROUPS = [
    ("벤토 모래(볼)", ["볼"]),
    ("벤토 모래(파쇄)", ["파쇄"]),
    ("두부모래", ["두부"]),
    ("마이티펫", ["마이티펫"]),
    ("장작", ["장작"]),
    ("모니터보호필름", ["필름"]),
    ("강아지 기저귀", ["기저귀"]),
    ("주방용품", ["리빙용품"]),
    ("차량용 공기필터", ["필터"]),
    ("기타", ["트리"]),
]
TYPE2GRP = {t: g for g, ts in GROUPS for t in ts}
GRP_ORDER = [g for g, _ in GROUPS]
BRANDS = ("코멧", "탐사", "홈플래닛")
BRAND_GROUPS = ("강아지 기저귀", "주방용품")  # 브랜드별로 펼칠 상품군


def fmt_krw(x) -> str:
    if x is None or pd.isna(x) or abs(x) < 0.5:
        return "-"
    if abs(x) >= 1e8:
        return f"{x / 1e8:.2f}억"
    return f"{x / 1e4:,.0f}만"


CSS = """
<style>
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(215px,1fr));gap:12px;margin:6px 0 4px}
.card{border:1px solid rgba(127,127,127,.25);border-left:5px solid var(--c);border-radius:10px;
      padding:12px 14px 10px;background:rgba(127,127,127,.03)}
.card .hd{display:flex;justify-content:space-between;align-items:center;font-weight:700;font-size:15px}
.badge{font-size:11.5px;font-weight:600;padding:2px 8px;border-radius:10px;color:var(--c);background:var(--bg)}
.card .big{font-size:23px;font-weight:800;margin-top:4px}
.card .big small{font-size:12.5px;font-weight:500;opacity:.65;margin-left:4px}
.bar{position:relative;display:flex;height:8px;background:rgba(127,127,127,.15);border-radius:4px;margin:9px 0 8px}
.bar .s1{background:#2f5d8a}.bar .s2{background:#3aa59a}.bar .s3{background:#9cc9e3}
.bar .s1:first-child{border-radius:4px 0 0 4px}
.bar .mk{position:absolute;top:-4px;width:2px;height:16px;background:currentColor}
.card .sub{font-size:12px;opacity:.78;line-height:1.65}
.legend{font-size:12px;opacity:.75;margin:4px 0 14px}
.legend i{display:inline-block;width:10px;height:10px;border-radius:2px;margin:0 4px 0 10px;vertical-align:-1px}
.htw{overflow-x:auto;border:1px solid rgba(127,127,127,.25);border-radius:10px}
table.ht{border-collapse:collapse;width:100%;font-size:13px;white-space:nowrap}
table.ht th{font-weight:600;padding:9px 8px;text-align:right;border-bottom:1px solid rgba(127,127,127,.25)}
table.ht th small{display:block;font-weight:400;opacity:.6;font-size:11px}
table.ht th:first-child,table.ht td:first-child{text-align:left;padding-left:14px;min-width:140px}
table.ht td{padding:9px 8px;text-align:right;border-bottom:1px solid rgba(127,127,127,.12)}
table.ht td:first-child{font-weight:600}
table.ht .cur{border-left:2px solid #2f5d8a;border-right:2px solid #2f5d8a}
table.ht .ex{border-left:1px solid rgba(127,127,127,.25)}
table.ht tr.subr td:first-child{font-weight:400;padding-left:30px;opacity:.85}
table.ht tr.subr td{font-size:12px}
table.ht tr.tot td{font-weight:800;border-top:2px solid rgba(127,127,127,.45)}
.sec{font-size:19px;font-weight:700;margin:26px 0 2px}
.cap{font-size:12.5px;opacity:.7;margin-bottom:8px}
</style>
"""

# ─────────────────────────────────────────────
# 사이드바 · 데이터 읽기
# ─────────────────────────────────────────────
st.markdown(CSS, unsafe_allow_html=True)
st.sidebar.title("📊 플랫폼사업부")
up = st.sidebar.file_uploader("데일리 업데이트 엑셀 (.xlsx)", type=["xlsx"])

if up is None:
    st.title("플랫폼사업부 통합 현황")
    st.info("왼쪽에서 데일리 업데이트 엑셀(dailyupdate_*.xlsx)을 올려 주세요.")
    st.caption(f"읽는 시트: {BUY_SHEET}, {OUT_SHEET}, {STOCK_SHEET}, {SUMMARY_SHEET}(목표매출 행), "
               f"{GROUP_TARGET_SHEET}(선택)")
    st.stop()

try:
    buy, out, stock, stock_date, target, prior, sku_names, grp_target_xl = load_workbook(up.getvalue())
except Exception as e:  # noqa: BLE001
    st.error(f"엑셀을 읽지 못했습니다: {e}")
    st.stop()

buy = buy.copy()
buy["상품군"] = buy["유형"].map(TYPE2GRP).fillna("기타")
first_word = buy["품목"].astype(str).str.split().str[0]
buy["브랜드"] = np.where(first_word.isin(BRANDS), first_word, "기타")
buy["ym"] = buy["날짜"].dt.to_period("M")

base_date = st.sidebar.date_input("기준일", value=TODAY)
base_ts = pd.Timestamp(base_date)
cur_p = pd.Period(base_ts, "M")
month_end = cur_p.end_time.normalize()

st.sidebar.markdown("---")
st.sidebar.caption(
    f"매출 데이터: {buy['날짜'].min():%Y-%m-%d} ~ {buy['날짜'].max():%Y-%m-%d}\n\n"
    f"출고 데이터: ~ {out['날짜'].max():%Y-%m-%d}\n\n"
    + (f"쿠팡 현재고 기준: {stock_date:%Y-%m-%d}" if stock_date is not None else "재고 현황 시트 없음")
)

tab_sales, tab_inv, tab_set = st.tabs(["💰 매출 · 쿠팡 PO", "📦 재고 · 결품", "⚙️ 데이터 설정 · 점검"])

# ─────────────────────────────────────────────
# 3. 데이터 설정 (다른 탭이 이 값을 쓰므로 먼저 그림)
# ─────────────────────────────────────────────
active_groups = [g for g in GRP_ORDER
                 if buy[(buy["상품군"] == g) & (buy["ym"] > cur_p - 12) & (buy["ym"] <= cur_p)]["금액(원)"].sum() > 0]

with tab_set:
    st.subheader(f"{cur_p.month}월 상품군별 목표")
    st.caption(
        f"엑셀에 '{GROUP_TARGET_SHEET}' 시트(열: 월 | 상품군 | 목표)를 만들어 두면 자동으로 읽습니다. "
        "여기서 고친 값은 이 화면을 새로고침하면 사라지니, 계속 쓸 목표는 엑셀에 넣어 주세요."
    )
    tgt_df = pd.DataFrame({
        "상품군": active_groups,
        "목표(만원)": [grp_target_xl.get((str(cur_p), g), np.nan) / 1e4 for g in active_groups],
    })
    edited = st.data_editor(
        tgt_df, hide_index=True, use_container_width=False, key=f"grp_tgt_{cur_p}",
        disabled=["상품군"],
        column_config={"목표(만원)": st.column_config.NumberColumn(min_value=0, step=100, format="%.0f")},
    )
    grp_target = {g: float(v) * 1e4 for g, v in zip(edited["상품군"], edited["목표(만원)"]) if pd.notna(v) and v > 0}
    total_tgt = target.get((cur_p.year, cur_p.month))
    st.caption(f"전체 목표는 요약 시트의 목표매출 행에서 읽습니다: {cur_p.month}월 {fmt_krw(total_tgt)}")

    st.markdown("---")
    st.subheader("재고 계산 설정")
    sc1, sc2 = st.columns([1, 2])
    window = sc1.selectbox("일평균 출고 기준 기간", [7, 14, 30, 90], index=2, format_func=lambda d: f"최근 {d}일")
    types_in_stock = sort_types(stock["유형"].unique()) if stock is not None else []
    lead = {}
    with sc2.expander("유형별 쿠팡 납품 리드타임(일)", expanded=False):
        st.caption("쿠팡 PO를 받고 VF에 입고되기까지 걸리는 기간입니다. 실제에 맞게 고쳐 주세요.")
        lc = st.columns(4)
        for i, t in enumerate(types_in_stock):
            lead[t] = lc[i % 4].number_input(t, min_value=0, max_value=180, value=DEFAULT_LEAD.get(t, 5),
                                             step=1, key=f"lead_{t}")

def compute_inv():
    if stock is None:
        return None, None
    last_out = out["날짜"].max()
    win = out[out["날짜"] > last_out - pd.Timedelta(days=window)]
    avg = win.groupby("sku id")["출고량"].sum() / window

    inv = stock.copy()
    inv["일평균출고"] = inv["sku id"].map(avg).fillna(0)
    inv["가용재고"] = inv["현재고"] + inv["입고예정"]
    with np.errstate(divide="ignore", invalid="ignore"):
        inv["현재고_일수"] = np.where(inv["일평균출고"] > 0, inv["현재고"] / inv["일평균출고"], np.inf)
        inv["가용_일수"] = np.where(inv["일평균출고"] > 0, inv["가용재고"] / inv["일평균출고"], np.inf)
    inv["리드타임"] = inv["유형"].map(lambda t: lead.get(t, DEFAULT_LEAD.get(t, 5)))

    def calc(r):
        if r["일평균출고"] <= 0:
            return pd.Series({"소진예상일": pd.NaT, "데드라인": pd.NaT, "남은일": np.nan, "상태": "⚪ 출고없음"})
        run_out = stock_date + pd.Timedelta(days=float(r["가용_일수"]))
        dl = run_out - pd.Timedelta(days=int(r["리드타임"]))
        left = (dl.normalize() - base_ts).days
        if r["가용_일수"] < r["리드타임"] or left < 0:
            s = "🔴 결품 위험"
        elif left <= 7:
            s = "🟠 1주 내 PO 예상"
        elif left <= 14:
            s = "🟡 2주 내 PO 예상"
        elif r["가용_일수"] > 90:
            s = "🔵 과재고"
        else:
            s = "🟢 정상"
        return pd.Series({"소진예상일": run_out.normalize(), "데드라인": dl.normalize(), "남은일": left, "상태": s})

    inv = pd.concat([inv, inv.apply(calc, axis=1)], axis=1)
    inv["_t"] = inv["유형"].map(type_key)
    inv = inv.sort_values(["_t", "남은일"], na_position="last").drop(columns="_t").reset_index(drop=True)
    return inv, last_out

INV, LAST_OUT = compute_inv()


# ─────────────────────────────────────────────
# 1. 매출 · 쿠팡 PO
# ─────────────────────────────────────────────
def month_stats(df: pd.DataFrame):
    """당월 실적(기준일까지 납품) / 확정(기준일 이후 납품 PO) / 예측(최근 4주 일평균 × PO 없는 남은 날)"""
    cur = df[df["ym"] == cur_p]
    act = cur.loc[cur["날짜"] <= base_ts, "금액(원)"].sum()
    conf = cur.loc[cur["날짜"] > base_ts, "금액(원)"].sum()
    recent = df[(df["날짜"] > base_ts - pd.Timedelta(days=28)) & (df["날짜"] <= base_ts)]["금액(원)"].sum() / 28
    last_conf = max(base_ts, cur["날짜"].max()) if len(cur) else base_ts
    remaining = max(0, (month_end - last_conf.normalize()).days)
    fc = recent * remaining if month_end >= base_ts else 0
    return float(act), float(conf), float(fc)


def month_value(df: pd.DataFrame, p):
    return float(df.loc[df["ym"] == p, "금액(원)"].sum())


def card_html(name, df, tgt):
    act, conf, fc = month_stats(df)
    done, exp = act + conf, act + conf + fc
    prev, ly = month_value(df, cur_p - 1), month_value(df, cur_p - 12)
    if tgt:
        if done >= tgt:
            c, bg, badge = "#1f9d55", "rgba(31,157,85,.13)", "달성"
        elif exp >= tgt:
            c, bg, badge = "#2a9d8f", "rgba(42,157,143,.13)", "달성 예상"
        elif exp >= tgt * 0.9:
            c, bg, badge = "#d4a017", "rgba(212,160,23,.15)", "근접"
        else:
            c, bg, badge = "#e05a47", "rgba(224,90,71,.13)", "미달 예상"
        rate = f"{done / tgt * 100:.0f}%"
        gap = exp - tgt
        line1 = f"목표 <b>{fmt_krw(tgt)}</b> · 예상 {fmt_krw(exp)}"
        line2 = f"예상 {exp / tgt * 100:.0f}% · {fmt_krw(abs(gap))} {'초과' if gap >= 0 else '부족'}"
        scale = max(tgt, exp) * 1.06
    else:
        c, bg, badge = "#8a94a6", "rgba(138,148,166,.15)", "목표 없음"
        rate = ""
        line1 = f"예상 <b>{fmt_krw(exp)}</b>"
        yoy = f" ({(exp / ly - 1) * 100:+.0f}%)" if ly else ""
        line2 = f"전월 {fmt_krw(prev)} · 전년 {fmt_krw(ly)}{yoy}"
        scale = max(exp, prev, 1) * 1.06
    w = lambda v: f"{max(v, 0) / scale * 100:.2f}%"  # noqa: E731
    mk = f'<div class="mk" style="left:{tgt / scale * 100:.2f}%"></div>' if tgt else ""
    return (
        f'<div class="card" style="--c:{c};--bg:{bg}">'
        f'<div class="hd"><span>{name}</span><span class="badge">{badge}</span></div>'
        f'<div class="big">{fmt_krw(done) if done else "0원"}<small>{rate}</small></div>'
        f'<div class="bar"><div class="s1" style="width:{w(act)}"></div><div class="s2" style="width:{w(conf)}"></div>'
        f'<div class="s3" style="width:{w(fc)}"></div>{mk}</div>'
        f'<div class="sub">{line1}<br>{line2}</div></div>'
    )


def heat_table(rows, months, extra_share=False):
    """rows: [(라벨, df, 'normal'|'sub'|'tot')] → 행마다 색 농도를 따로 매기는 월 매출표 HTML"""
    head = "<tr><th>구분</th>"
    for p in months:
        yr = f"{p.year % 100}년" if (p.month == 1 or p == months[0]) else ""
        if p == cur_p:
            head += f'<th class="cur">{p.month}월<small>(예상)</small></th>'
        else:
            head += f"<th>{p.month}월<small>{yr}</small></th>"
    head += '<th class="ex">전년비</th>'
    if extra_share:
        head += "<th>당월 비중</th>"
    head += f"<th>{cur_p.year % 100}년 누적<small>1~{cur_p.month}월</small></th></tr>"

    total_cur = None
    if extra_share:
        total_cur = sum(sum(month_stats(df)) for _, df, kind in rows if kind == "normal")

    body = ""
    for label, df, kind in rows:
        act, conf, fc = month_stats(df)
        vals = [month_value(df, p) if p != cur_p else act + conf + fc for p in months]
        vmax = max(vals) if vals and max(vals) > 0 else 1
        ly = month_value(df, cur_p - 12)
        cur_v = act + conf + fc
        yoy = f"{(cur_v / ly - 1) * 100:+.0f}%" if ly > 0 else "-"
        ytd = sum(month_value(df, pd.Period(year=cur_p.year, month=mm, freq="M")) for mm in range(1, cur_p.month)) + act + conf
        cls = {"sub": ' class="subr"', "tot": ' class="tot"'}.get(kind, "")
        body += f"<tr{cls}><td>{label}</td>"
        for p, v in zip(months, vals):
            a = 0.06 + 0.6 * (v / vmax) if v > 0 else 0
            style = f"background:rgba(47,93,138,{a:.2f});" + ("color:#fff;" if a > 0.42 else "")
            ccls = ' class="cur"' if p == cur_p else ""
            body += f'<td{ccls} style="{style}">{fmt_krw(v)}</td>'
        body += f'<td class="ex">{yoy}</td>'
        if extra_share:
            share = f"{cur_v / total_cur * 100:.1f}%" if total_cur else "-"
            body += f"<td>{share if kind != 'tot' else '100%'}</td>"
        body += f"<td><b>{fmt_krw(ytd)}</b></td></tr>"
    return f'<div class="htw"><table class="ht">{head}{body}</table></div>'


with tab_sales:
    months = [cur_p - i for i in range(11, -1, -1)]

    # ── 상품군별 목표 달성 카드
    st.markdown(f'<div class="sec">🎯 {cur_p.month}월 상품군별 목표 달성</div>'
                f'<div class="cap">큰 숫자 = 달성액(실적+확정)과 목표 대비 % · 검은 선 = 목표 · 기준일 {base_date:%m/%d}</div>',
                unsafe_allow_html=True)
    cards = card_html("전체", buy, total_tgt)
    for g in active_groups:
        cards += card_html(g, buy[buy["상품군"] == g], grp_target.get(g))
    st.markdown(f'<div class="cards">{cards}</div>'
                '<div class="legend"><i style="background:#2f5d8a"></i>실적(기준일까지 납품)'
                '<i style="background:#3aa59a"></i>확정(기준일 이후 납품 PO)'
                '<i style="background:#9cc9e3"></i>예측(최근 4주 일평균 × PO 없는 남은 날)</div>',
                unsafe_allow_html=True)

    # ── 요약 줄: 최근 수주 / 납품 예정 / 결품 위험
    e1, e2, e3 = st.columns(3)
    recent = buy[buy["발주일"].notna() & (buy["발주일"] <= base_ts)]
    if not recent.empty:
        last_po = recent["발주일"].max()
        today_po = recent[recent["발주일"] == last_po]
        with e1.expander(f"🧾 최근 수주 {last_po:%m/%d} · {fmt_krw(today_po['금액(원)'].sum())}"):
            g = (today_po.groupby("상품군").agg(금액=("금액(원)", "sum"), 수량=("수량(BOX)", "sum"),
                                              납기=("날짜", "max")).reset_index())
            g = g.sort_values("상품군", key=lambda s: s.map(GRP_ORDER.index))
            g["금액"] = g["금액"].map(fmt_krw)
            g["수량"] = g["수량"].fillna(0).map(lambda v: f"{v:,.0f}")
            g["납기"] = g["납기"].dt.strftime("~%m/%d")
            st.dataframe(g, hide_index=True, use_container_width=True)

    fut = buy[buy["날짜"] > base_ts]
    with e2.expander(f"🚚 납품 예정 · {len(fut)}건 · {fmt_krw(fut['금액(원)'].sum())}"):
        if fut.empty:
            st.caption("기준일 이후 납품 예정 PO가 없습니다.")
        else:
            pv = fut.pivot_table(index=fut["날짜"].dt.strftime("%m/%d"), columns="상품군",
                                 values="금액(원)", aggfunc="sum", fill_value=0)
            pv = pv[[c for c in GRP_ORDER if c in pv.columns]]
            pv["합계"] = pv.sum(axis=1)
            st.dataframe(pv.map(fmt_krw), use_container_width=True)

    if INV is not None:
        risk = INV[INV["상태"].isin(["🔴 결품 위험", "🟠 1주 내 PO 예상"])]
        n_red = int((INV["상태"] == "🔴 결품 위험").sum())
        n_or = int((INV["상태"] == "🟠 1주 내 PO 예상").sum())
        with e3.expander(f"📦 쿠팡 재고 · 🔴 {n_red} 🟠 {n_or}"):
            st.dataframe(risk[["상태", "품명", "현재고", "입고예정", "소진예상일"]].assign(
                현재고=lambda d: d["현재고"].round(0), 입고예정=lambda d: d["입고예정"].round(0)),
                hide_index=True, use_container_width=True,
                column_config={"소진예상일": st.column_config.DateColumn(format="MM/DD")})
            st.caption("자세한 내용은 '재고 · 결품' 탭에서 확인하세요.")

    # ── 상품군별 월 매출표
    st.markdown('<div class="sec">📅 상품군별 월 매출 (최근 12개월)</div>'
                '<div class="cap">색이 진할수록 그 행 안에서 매출이 큰 달 · 이번 달은 실적+확정+예측 · '
                '누적은 올해 1월부터 (이번 달은 실적+확정만)</div>', unsafe_allow_html=True)
    split = st.toggle("강아지 기저귀 · 주방용품 브랜드별로 펼치기", value=False)
    rows = []
    for g in active_groups:
        gdf = buy[buy["상품군"] == g]
        rows.append((g, gdf, "normal"))
        if split and g in BRAND_GROUPS:
            for b in [b for b in BRANDS if b in gdf["브랜드"].unique()]:
                rows.append((f"└ {b}", gdf[gdf["브랜드"] == b], "sub"))
    rows.append(("합계", buy, "tot"))
    st.markdown(heat_table(rows, months), unsafe_allow_html=True)

    # ── 채널별 월 매출표 (행마다 색을 따로 매겨 규모 차이가 커도 추이가 보이게)
    st.markdown('<div class="sec">🛒 채널별 월 매출</div>'
                '<div class="cap">채널마다 규모 차이가 커서, 색은 채널(행)별로 따로 매깁니다 · 당월 비중 = 이번 달 예상 기준</div>',
                unsafe_allow_html=True)
    ch_tot = buy[buy["ym"] > cur_p - 12].groupby("공급채널")["금액(원)"].sum().sort_values(ascending=False)
    ch_rows = [(ch, buy[buy["공급채널"] == ch], "normal") for ch in ch_tot.index if ch_tot[ch] > 0]
    ch_rows.append(("합계", buy, "tot"))
    st.markdown(heat_table(ch_rows, months, extra_share=True), unsafe_allow_html=True)


# ─────────────────────────────────────────────
# 2. 재고 · 결품
# ─────────────────────────────────────────────
with tab_inv:
    if stock is None:
        st.warning(f"'{STOCK_SHEET}' 시트를 찾지 못해 재고 계산을 할 수 없습니다.")
    else:
        inv, last_out = INV, LAST_OUT


        st.subheader("쿠팡 VF 재고 · 결품 위험")
        st.caption(
            f"현재고 기준 {stock_date:%m/%d} (벤플 수불)  ·  일평균 출고 = {last_out:%m/%d}까지 최근 {window}일  ·  "
            "소진예상일 = 현재고 기준일 + (현재고+입고예정) ÷ 일평균  ·  납품 데드라인 = 소진예상일 − 납품 리드타임"
        )

        cnt = inv["상태"].value_counts()
        kk = st.columns(5)
        for i, s in enumerate(["🔴 결품 위험", "🟠 1주 내 PO 예상", "🟡 2주 내 PO 예상", "🟢 정상", "🔵 과재고"]):
            kk[i].metric(s, f"{int(cnt.get(s, 0))}개 SKU")

        f1, f2 = st.columns([3, 2])
        sel_types = f1.multiselect("유형", types_in_stock, default=types_in_stock)
        sel_status = f2.multiselect("상태", sorted(inv["상태"].unique()), default=sorted(inv["상태"].unique()))
        view = inv[inv["유형"].isin(sel_types) & inv["상태"].isin(sel_status)].reset_index(drop=True)

        cols = ["상태", "유형", "품명", "sku id", "현재고", "입고예정", "일평균출고", "현재고_일수", "가용_일수",
                "소진예상일", "리드타임", "데드라인", "남은일"]
        disp = view[cols].copy()
        disp["현재고_일수"] = disp["현재고_일수"].replace(np.inf, np.nan)
        disp["가용_일수"] = disp["가용_일수"].replace(np.inf, np.nan)
        disp[["현재고", "입고예정"]] = disp[["현재고", "입고예정"]].round(0)
        st.caption("👇 행을 클릭하면 아래에 SKU 상세가 나옵니다.")
        ev = st.dataframe(
            disp, hide_index=True, use_container_width=True, height=min(38 * (len(disp) + 1), 640),
            on_select="rerun", selection_mode="single-row", key="inv_table",
            column_config={
                "sku id": st.column_config.NumberColumn("SKU", format="%d"),
                "현재고": st.column_config.NumberColumn(format="localized"),
                "입고예정": st.column_config.NumberColumn(format="localized"),
                "일평균출고": st.column_config.NumberColumn(format="%.1f"),
                "현재고_일수": st.column_config.NumberColumn("현재고만(일)", format="%.0f"),
                "가용_일수": st.column_config.NumberColumn("입고포함(일)", format="%.0f"),
                "소진예상일": st.column_config.DateColumn(format="MM/DD"),
                "리드타임": st.column_config.NumberColumn("납품 리드타임", format="%d"),
                "데드라인": st.column_config.DateColumn("납품 데드라인", format="MM/DD"),
                "남은일": st.column_config.NumberColumn("남은 일수", format="%d"),
            })

        # ── SKU 상세
        st.markdown("---")
        sel_rows = getattr(getattr(ev, "selection", None), "rows", []) or []
        options = view["sku id"].tolist()
        if not options:
            st.info("조건에 맞는 SKU가 없습니다.")
        else:
            default_idx = sel_rows[0] if sel_rows and sel_rows[0] < len(options) else 0
            sku = st.selectbox(
                "SKU 상세", options, index=default_idx,
                format_func=lambda s: f"[{sku_names.loc[s, '유형'] if s in sku_names.index else ''}] "
                                      f"{sku_names.loc[s, '품명'] if s in sku_names.index else s} ({s})")
            r = view[view["sku id"] == sku].iloc[0]

            d1, d2, d3, d4 = st.columns(4)
            d1.metric("현재고", f"{r['현재고']:,.0f}")
            d2.metric("입고예정", f"{r['입고예정']:,.0f}")
            d3.metric(f"일평균 출고 ({window}일)", f"{r['일평균출고']:,.1f}")
            d4.metric("상태", r["상태"], f"데드라인 {r['데드라인']:%m/%d}" if pd.notna(r["데드라인"]) else None,
                      delta_color="off")

            g1, g2 = st.columns(2)
            with g1:
                so = out[out["sku id"] == sku].groupby("날짜")["출고량"].sum()
                if not so.empty:
                    idx = pd.date_range(last_out - pd.Timedelta(days=119), last_out)
                    so = so.reindex(idx, fill_value=0)
                    fig = go.Figure()
                    fig.add_scatter(x=so.index, y=so.values, name="일 출고", mode="lines", line=dict(width=1))
                    fig.add_scatter(x=so.index, y=so.rolling(7, min_periods=1).mean(), name="7일 평균",
                                    mode="lines", line=dict(width=3))
                    fig.update_layout(title="일별 출고 (최근 120일)", height=340,
                                      legend=dict(orientation="h", y=-0.2), margin=dict(t=50, b=10))
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("출고 기록이 없습니다.")

            with g2:
                # 재고 예상 곡선: 현재고에서 일평균만큼 줄고, RAW_매입의 기준일 이후 납품 예정분이 더해짐
                po = buy[(buy["sku id"] == sku)]
                fut = po[po["날짜"] > stock_date].groupby("날짜")["수량(BOX)"].sum()
                days = pd.date_range(stock_date, stock_date + pd.Timedelta(days=60))
                level, lv = [], float(r["현재고"])
                for d in days:
                    if d != stock_date:
                        lv -= r["일평균출고"]
                    lv += float(fut.get(d, 0))
                    level.append(lv)
                fig = go.Figure()
                fig.add_scatter(x=days, y=level, name="예상 재고", mode="lines", fill="tozeroy")
                if len(fut):
                    fig.add_bar(x=fut.index, y=fut.values, name="납품 예정(PO)", opacity=0.6)
                fig.add_hline(y=0, line_dash="dot")
                if pd.notna(r["데드라인"]):
                    fig.add_vline(x=r["데드라인"], line_dash="dash", line_color="red")
                fig.update_layout(title="쿠팡 재고 예상 (60일)", height=340,
                                  legend=dict(orientation="h", y=-0.2), margin=dict(t=50, b=10))
                st.plotly_chart(fig, use_container_width=True)
                if len(fut) and abs(fut.sum() - r["입고예정"]) > 0.5:
                    st.caption(f"※ RAW_매입의 납품 예정 합계({fut.sum():,.0f})와 재고 시트의 입고예정"
                               f"({r['입고예정']:,.0f})이 다릅니다. 곡선은 RAW_매입 기준입니다.")

            hist = po.assign(**{"금액(원)": po["금액(원)"].round(0)}).sort_values("날짜", ascending=False)[["날짜", "발주일", "품목", "수량(BOX)", "금액(원)", "입고처"]].head(30)
            with st.expander(f"쿠팡 PO 이력 (최근 {len(hist)}건)"):
                st.dataframe(hist, hide_index=True, use_container_width=True,
                             column_config={"날짜": st.column_config.DateColumn("납품일", format="YYYY-MM-DD"),
                                            "발주일": st.column_config.DateColumn(format="YYYY-MM-DD"),
                                            "수량(BOX)": st.column_config.NumberColumn(format="localized"),
                                            "금액(원)": st.column_config.NumberColumn(format="localized")})


# ─────────────────────────────────────────────
# 3. 데이터 점검 (설정 탭 아래)
# ─────────────────────────────────────────────
with tab_set:
    st.markdown("---")
    st.subheader("데이터 점검")
    st.caption("숫자가 이상할 때 먼저 확인할 것들입니다. 대시보드는 엑셀 내용을 그대로 계산합니다.")

    a, b, c = st.columns(3)
    stale = (base_ts - out["날짜"].max()).days
    a.metric("출고 데이터 마지막 날짜", f"{out['날짜'].max():%m/%d}", f"기준일보다 {stale}일 전" if stale > 0 else "최신",
             delta_color="inverse" if stale > 3 else "off")
    b.metric("목표매출 행", f"{len(target)}개월 읽음" if target else "못 읽음")
    c.metric("재고 현황 시트", f"{len(stock)}개 SKU" if stock is not None else "없음")

    dup = buy[buy.drop(columns=["ym"], errors="ignore").duplicated(keep=False)]
    with st.expander(f"RAW_매입 완전 중복 행: {len(dup)}행 (매출에 그대로 포함됨)"):
        st.caption("같은 날 같은 수량의 PO가 실제로 여러 건이면 정상입니다. 입력 실수면 엑셀에서 지워 주세요.")
        st.dataframe(dup[["날짜", "품목", "유형", "수량(BOX)", "금액(원)", "입고처", "발주일"]], hide_index=True,
                     use_container_width=True)

    no_amt = buy[(buy["금액(원)"] == 0)]
    with st.expander(f"금액이 비었거나 0인 행: {len(no_amt)}행"):
        st.dataframe(no_amt[["날짜", "품목", "유형", "수량(BOX)", "입고처", "비고"]], hide_index=True,
                     use_container_width=True)

    unm = (buy[buy["sku id"].isna() & (buy["공급채널"] == "쿠팡")]
           .groupby(["유형", "품목"]).agg(건수=("날짜", "size"), 마지막=("날짜", "max")).reset_index())
    with st.expander(f"SKU와 연결 안 된 쿠팡 품목: {len(unm)}개"):
        st.caption("RAW_매입 품목명이 출고/재고 시트 품명과 달라 SKU를 못 찾은 경우입니다. "
                   "출고 기록이 없는 신규 품목이면 정상입니다.")
        st.dataframe(unm, hide_index=True, use_container_width=True)

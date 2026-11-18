"""플랫폼사업부 통합 현황 대시보드

데일리 업데이트 엑셀(dailyupdate) 한 파일만 올리면
  1) 사업부 매출 현황 (목표 대비, 전월·전년 비교, 품목·채널·주차별)
  2) 쿠팡 VF 재고 · 결품 위험 · 공급 준비 데드라인
  3) 데이터 점검
을 보여 줍니다.

읽는 시트: RAW_매입, RAW_출고(벤플), SKU별 재고 현황(쿠팡), 요약(목표매출 행)
"""

import io
import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import plotly.express as px
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

    return buy, out, stock, stock_date, target, prior, names


# ─────────────────────────────────────────────
# 사이드바
# ─────────────────────────────────────────────
st.sidebar.title("📊 플랫폼사업부")
up = st.sidebar.file_uploader("데일리 업데이트 엑셀 (.xlsx)", type=["xlsx"])

if up is None:
    st.title("플랫폼사업부 통합 현황")
    st.info("왼쪽에서 데일리 업데이트 엑셀(dailyupdate_*.xlsx)을 올려 주세요.")
    st.caption(f"읽는 시트: {BUY_SHEET}, {OUT_SHEET}, {STOCK_SHEET}, {SUMMARY_SHEET}(목표매출 행)")
    st.stop()

try:
    buy, out, stock, stock_date, target, prior, sku_names = load_workbook(up.getvalue())
except Exception as e:  # noqa: BLE001
    st.error(f"엑셀을 읽지 못했습니다: {e}")
    st.stop()

base_date = st.sidebar.date_input("기준일", value=TODAY)
base_ts = pd.Timestamp(base_date)

st.sidebar.markdown("---")
st.sidebar.subheader("재고 계산 설정")
window = st.sidebar.selectbox("일평균 출고 기준 기간", [7, 14, 30, 90], index=2, format_func=lambda d: f"최근 {d}일")
types_in_stock = sort_types(stock["유형"].unique()) if stock is not None else []
lead = {}
with st.sidebar.expander("유형별 쿠팡 납품 리드타임(일)"):
    st.caption("쿠팡 PO를 받고 VF에 입고되기까지 걸리는 기간입니다. 실제에 맞게 고쳐 주세요.")
    for t in types_in_stock:
        lead[t] = st.number_input(t, min_value=0, max_value=180, value=DEFAULT_LEAD.get(t, 5), step=1, key=f"lead_{t}")

st.sidebar.markdown("---")
st.sidebar.caption(
    f"매출 데이터: {buy['날짜'].min():%Y-%m-%d} ~ {buy['날짜'].max():%Y-%m-%d}\n\n"
    f"출고 데이터: ~ {out['날짜'].max():%Y-%m-%d}\n\n"
    + (f"쿠팡 현재고 기준: {stock_date:%Y-%m-%d}" if stock_date is not None else "재고 현황 시트 없음")
)

tab1, tab2, tab3 = st.tabs(["📈 사업부 현황", "📦 쿠팡 재고 · 데드라인", "🔍 데이터 점검"])

# ─────────────────────────────────────────────
# 1. 사업부 현황
# ─────────────────────────────────────────────
with tab1:
    y, m = base_ts.year, base_ts.month
    buy["ym"] = buy["날짜"].dt.to_period("M")
    cur_p = pd.Period(year=y, month=m, freq="M")
    prev_p = cur_p - 1
    ly_p = cur_p - 12

    month_tot = buy.groupby("ym")["금액(원)"].sum()

    def month_sales(p):
        if p in month_tot.index and month_tot[p] != 0:
            return float(month_tot[p])
        return prior.get((p.year, p.month))

    cur_rows = buy[buy["ym"] == cur_p]
    cur_total = cur_rows["금액(원)"].sum()
    delivered = cur_rows.loc[cur_rows["날짜"] <= base_ts, "금액(원)"].sum()
    scheduled = cur_total - delivered
    tgt = target.get((y, m))

    # 전월 동기간 (1일 ~ 같은 날짜)
    prev_start = prev_p.start_time
    prev_same_end = min(prev_start + pd.Timedelta(days=base_ts.day - 1), prev_p.end_time)
    prev_same = buy[(buy["날짜"] >= prev_start) & (buy["날짜"] <= prev_same_end)]["금액(원)"].sum()
    prev_full = month_sales(prev_p)
    ly = month_sales(ly_p)

    ytd = buy[(buy["날짜"].dt.year == y) & (buy["ym"] <= cur_p)]["금액(원)"].sum()
    ytd_tgt = sum(v for (yy, mm), v in target.items() if yy == y and mm <= m)

    st.subheader(f"{y}년 {m}월 현황  ·  기준일 {base_date:%m/%d}")
    k = st.columns(5)
    k[0].metric("당월 매출 (확정+예정)", won_m(cur_total),
                help="RAW_매입의 날짜(납품일)가 이번 달인 수주 합계. 기준일 이후 납품 예정분 포함.")
    k[1].metric("목표 달성률", f"{cur_total / tgt * 100:.1f}%" if tgt else "-",
                f"목표 {won_m(tgt)}" if tgt else None, delta_color="off")
    k[2].metric("전월 동기간 대비", pct(delivered / prev_same - 1) if prev_same else "-",
                f"1~{base_ts.day}일 납품분 비교", delta_color="off",
                help="이번 달 기준일까지 납품분 vs 지난달 같은 날짜까지 납품분")
    k[3].metric("전년 동월 대비", pct(cur_total / ly - 1) if ly else "-",
                f"전년 {won_m(ly)}" if ly else None, delta_color="off")
    k[4].metric("연 누계 달성률", f"{ytd / ytd_tgt * 100:.1f}%" if ytd_tgt else "-",
                f"{won_m(ytd)} / {won_m(ytd_tgt)}" if ytd_tgt else won_m(ytd), delta_color="off")
    st.caption(f"당월 매출 중 기준일까지 납품 {won_m(delivered)}, 이후 납품 예정 {won_m(scheduled)}  ·  "
               f"전월 전체 {won_m(prev_full)}")

    # 금일(최근) 수주
    recent = buy[buy["발주일"].notna() & (buy["발주일"] <= base_ts)]
    if not recent.empty:
        last_po = recent["발주일"].max()
        today_po = recent[recent["발주일"] == last_po]
        with st.expander(f"🧾 최근 수주 ({last_po:%m/%d} 발주분) 합계 {today_po['금액(원)'].sum():,.0f}원", expanded=True):
            g = (today_po.groupby(["구분", "유형"], dropna=False)
                 .agg(금액=("금액(원)", "sum"), 수량_BOX=("수량(BOX)", "sum"), 납기_마지막=("날짜", "max"))
                 .reset_index())
            g = g.sort_values("유형", key=lambda s: s.map(type_key))
            g["납기_마지막"] = g["납기_마지막"].dt.strftime("~%m/%d")
            g[["금액", "수량_BOX"]] = g[["금액", "수량_BOX"]].round(0)
            st.dataframe(g, hide_index=True, use_container_width=True,
                         column_config={"금액": st.column_config.NumberColumn("금액(원)", format="localized"),
                                        "수량_BOX": st.column_config.NumberColumn("수량(BOX)", format="localized")})

    c1, c2 = st.columns(2)

    # 월별 매출 vs 목표 (당해)
    with c1:
        yr = buy[buy["날짜"].dt.year == y]
        mg = yr.groupby([yr["날짜"].dt.month, "구분"])["금액(원)"].sum().unstack(fill_value=0)
        fig = go.Figure()
        for col in [c for c in ["제품", "상품"] if c in mg.columns] + [c for c in mg.columns if c not in ("제품", "상품")]:
            fig.add_bar(x=[f"{i}월" for i in mg.index], y=mg[col] / 1e6, name=str(col))
        tm = sorted((mm, v) for (yy, mm), v in target.items() if yy == y)
        if tm:
            fig.add_scatter(x=[f"{mm}월" for mm, _ in tm], y=[v / 1e6 for _, v in tm],
                            name="목표", mode="lines+markers", line=dict(dash="dash"))
        fig.update_layout(barmode="stack", title=f"{y}년 월별 매출 vs 목표 (백만원)", height=380,
                          legend=dict(orientation="h", y=-0.15), margin=dict(t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    # 유형별 월 매출 (최근 13개월)
    with c2:
        span = buy[(buy["ym"] > cur_p - 13) & (buy["ym"] <= cur_p)]
        tg = span.groupby(["ym", "유형"])["금액(원)"].sum().reset_index()
        tg["월"] = tg["ym"].dt.strftime("%y.%m")
        tg["백만원"] = tg["금액(원)"] / 1e6
        fig = px.bar(tg, x="월", y="백만원", color="유형",
                     category_orders={"유형": sort_types(tg["유형"].unique())},
                     title="품목(유형)별 월 매출 — 최근 13개월")
        fig.update_layout(height=380, legend=dict(orientation="h", y=-0.15), margin=dict(t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)

    # 당월 주차별
    with c3:
        wk = cur_rows.groupby(["주차", "유형"])["금액(원)"].sum().reset_index()
        if not wk.empty:
            wk["백만원"] = wk["금액(원)"] / 1e6
            fig = px.bar(wk, x="주차", y="백만원", color="유형",
                         category_orders={"유형": sort_types(wk["유형"].unique()),
                                          "주차": sorted(wk["주차"].dropna().unique())},
                         title=f"{m}월 주차별 매출")
            fig.update_layout(height=360, legend=dict(orientation="h", y=-0.2), margin=dict(t=50, b=10))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("이번 달 매출 데이터가 없습니다.")

    # 채널별 추이
    with c4:
        ch = span.groupby(["ym", "공급채널"])["금액(원)"].sum().reset_index()
        ch["월"] = ch["ym"].dt.strftime("%y.%m")
        ch["백만원"] = ch["금액(원)"] / 1e6
        fig = px.line(ch, x="월", y="백만원", color="공급채널", markers=True, title="채널별 월 매출")
        fig.update_layout(height=360, legend=dict(orientation="h", y=-0.2), margin=dict(t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    # 유형별 당월 요약 표
    st.markdown("##### 품목(유형)별 당월 실적")
    prev_rows = buy[buy["ym"] == prev_p]
    ly_rows = buy[buy["ym"] == ly_p]
    tbl = pd.DataFrame({
        "당월": cur_rows.groupby("유형")["금액(원)"].sum(),
        "전월": prev_rows.groupby("유형")["금액(원)"].sum(),
        "전년동월": ly_rows.groupby("유형")["금액(원)"].sum(),
        "당월수량(BOX)": cur_rows.groupby("유형")["수량(BOX)"].sum(),
    }).fillna(0)
    tbl = tbl.loc[sort_types(tbl.index)]
    tbl["비중"] = tbl["당월"] / tbl["당월"].sum() if tbl["당월"].sum() else 0
    tbl["전년比"] = np.where(tbl["전년동월"] > 0, tbl["당월"] / tbl["전년동월"] - 1, np.nan)
    show = tbl.copy()
    show["당월수량(BOX)"] = show["당월수량(BOX)"].round(0)
    for c in ["당월", "전월", "전년동월"]:
        show[c] = show[c] / 1e6
    st.dataframe(
        show.reset_index().rename(columns={"index": "유형"}), hide_index=True, use_container_width=True,
        column_config={
            "당월": st.column_config.NumberColumn("당월(백만원)", format="%.1f"),
            "전월": st.column_config.NumberColumn("전월 전체(백만원)", format="%.1f"),
            "전년동월": st.column_config.NumberColumn("전년동월(백만원)", format="%.1f"),
            "당월수량(BOX)": st.column_config.NumberColumn(format="localized"),
            "비중": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1),
            "전년比": st.column_config.NumberColumn("전년 대비", format="percent"),
        })

# ─────────────────────────────────────────────
# 2. 쿠팡 재고 · 데드라인
# ─────────────────────────────────────────────
with tab2:
    if stock is None:
        st.warning(f"'{STOCK_SHEET}' 시트를 찾지 못해 재고 계산을 할 수 없습니다.")
    else:
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
# 3. 데이터 점검
# ─────────────────────────────────────────────
with tab3:
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

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
STOCK_IN_SHEET = "재고입고관리"  # 통합: 분류 | 품목 | 구분(현재고/입고예정) | 날짜 | 수량 | 단위 | 비고
RAW_SHEET = STOCK_IN_SHEET  # 화면 문구용 (예전 '모래원재료' 시트도 계속 읽음)
RAW_SHEET_NAMES = ["모래원재료", "원재료"]  # 이 이름 중 먼저 찾은 시트를 읽음
RAW_MATS = ["볼", "파쇄"]
TOFU_SHEET_NAMES = ["두부모래재고", "두부재고"]  # 선택: 품목 | 구분(현재고/입고예정) | 날짜 | 수량(박스) | 비고
TOFU_FLAVORS = ["오리지널", "라벤더", "녹차", "복숭아", "베이비", "숯"]


def flavor_of(name):
    for f in TOFU_FLAVORS:
        if f in str(name):
            return f
    return None

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


def to_month(v):
    """셀 값 → 'YYYY-MM' (날짜 셀, '2026-11', '2026.11', '2026년 11월' 모두 허용)"""
    if isinstance(v, (datetime, pd.Timestamp)):
        return str(pd.Period(v, "M"))
    m = re.match(r"^\s*(\d{4})\s*[-./년]\s*(\d{1,2})", str(v))
    return f"{int(m.group(1))}-{int(m.group(2)):02d}" if m else None


def parse_group_target(raw: pd.DataFrame) -> dict:
    """상품군목표 시트 → {('YYYY-MM', 행 이름): 목표금액}"""
    res = {}
    top = min(10, len(raw))
    # 세로형: 머리글에 월 / 상품군 / 목표
    for i in range(top):
        row = [str(v).strip() if pd.notna(v) else "" for v in raw.iloc[i].tolist()]
        if {"월", "상품군", "목표"} <= set(row):
            cm, cg, cv = row.index("월"), row.index("상품군"), row.index("목표")
            for k in range(i + 1, len(raw)):
                per, lab = to_month(raw.iat[k, cm]), raw.iat[k, cg]
                v = pd.to_numeric(raw.iat[k, cv], errors="coerce")
                if per and pd.notna(lab) and pd.notna(v):
                    key = (per, str(lab).strip())
                    res[key] = res.get(key, 0) + float(v)
            return res
    # 가로형: 머리글 행에 월이 가로로, 왼쪽 칸들이 상품군 이름
    for i in range(top):
        cols = {j: to_month(v) for j, v in enumerate(raw.iloc[i].tolist()) if pd.notna(v) and to_month(v)}
        if not cols:
            continue
        first = min(cols)
        for k in range(i + 1, len(raw)):
            parts = [str(raw.iat[k, j]).strip() for j in range(first) if pd.notna(raw.iat[k, j])]
            lab = " ".join(p for p in parts if p)
            if not lab:
                continue
            for j, per in cols.items():
                v = pd.to_numeric(raw.iat[k, j], errors="coerce")
                if pd.notna(v):
                    res[(per, lab)] = res.get((per, lab), 0) + float(v)
        break
    return res


def parse_raw_sheet(xls: pd.ExcelFile):
    """원재료 시트 → DataFrame(원재료, 구분, 날짜, 수량(톤), 비고). 없으면 None"""
    name = next((n for n in RAW_SHEET_NAMES if n in xls.sheet_names), None)
    if name is None:
        return None
    raw = pd.read_excel(xls, name, header=None)
    for i in range(min(10, len(raw))):
        row = [str(v).strip() if pd.notna(v) else "" for v in raw.iloc[i].tolist()]
        if "원재료" in row and "구분" in row:
            def col(*names):
                for n in names:
                    for j, h in enumerate(row):
                        if h.replace(" ", "").startswith(n):
                            return j
                return None
            cm, cg, cd, cq = col("원재료"), col("구분"), col("날짜", "입고예정일", "일자"), col("수량", "톤")
            cb, co = col("비고"), col("발주일")
            if None in (cm, cg, cq):
                return None
            body = raw.iloc[i + 1:]
            df = pd.DataFrame({
                "원재료": body.iloc[:, cm].astype("string").str.strip(),
                "구분": body.iloc[:, cg].astype("string").str.strip(),
                "날짜": pd.to_datetime(body.iloc[:, cd], errors="coerce") if cd is not None else pd.NaT,
                "수량(톤)": pd.to_numeric(body.iloc[:, cq], errors="coerce"),
                "비고": body.iloc[:, cb].astype("string") if cb is not None else "",
            })
            if co is not None:  # 입고일이 비어 있고 발주일만 있으면 비고에 남겨 둠
                od = pd.to_datetime(body.iloc[:, co], errors="coerce")
                df["발주일"] = od.values
            df["원재료"] = df["원재료"].str.replace("벤토나이트", "").str.replace("_", "").str.replace("모래", "").str.strip()
            df.loc[df["원재료"].str.contains("볼", na=False), "원재료"] = "볼"
            df.loc[df["원재료"].str.contains("파쇄", na=False), "원재료"] = "파쇄"
            df.loc[df["구분"].str.contains("현재", na=False), "구분"] = "현재고"
            df.loc[df["구분"].str.contains("입고|예정", na=False), "구분"] = "입고예정"
            return df[df["원재료"].isin(RAW_MATS) & df["수량(톤)"].notna()].reset_index(drop=True)
    return None


def parse_stock_in_sheet(xls: pd.ExcelFile):
    """재고입고관리 시트(원재료·두부모래·마이티펫 통합) → (원재료 df, 두부 df, 마이티펫 df). 없으면 (None, None, None)"""
    if STOCK_IN_SHEET not in xls.sheet_names:
        return None, None, None
    raw = pd.read_excel(xls, STOCK_IN_SHEET, header=None)
    for i in range(min(10, len(raw))):
        row = [str(v).strip().replace(" ", "") if pd.notna(v) else "" for v in raw.iloc[i].tolist()]
        if "분류" in row and "구분" in row:
            def col(*names):
                for n in names:
                    for j, h in enumerate(row):
                        if h.startswith(n):
                            return j
                return None
            ck, cn, cg, cd, cq, cb = col("분류"), col("품목", "품명"), col("구분"), col("날짜"), col("수량"), col("비고")
            if None in (ck, cn, cg, cq):
                break
            body = raw.iloc[i + 1:]
            df = pd.DataFrame({
                "분류": body.iloc[:, ck].astype("string").str.strip(),
                "품목": body.iloc[:, cn].astype("string").str.strip(),
                "구분": body.iloc[:, cg].astype("string").str.strip(),
                "날짜": pd.to_datetime(body.iloc[:, cd], errors="coerce") if cd is not None else pd.NaT,
                "수량": pd.to_numeric(body.iloc[:, cq], errors="coerce"),
                "비고": body.iloc[:, cb].astype("string").fillna("") if cb is not None else "",
            })
            df = df[df["분류"].notna() & df["수량"].notna()]
            df.loc[df["구분"].str.contains("현재", na=False), "구분"] = "현재고"
            df.loc[df["구분"].str.contains("입고|예정", na=False), "구분"] = "입고예정"

            rm = df[df["분류"].str.contains("원재료", na=False)].copy()
            rm["원재료"] = np.where(rm["품목"].str.contains("볼", na=False), "볼",
                                 np.where(rm["품목"].str.contains("파쇄", na=False), "파쇄", None))
            rm = rm[rm["원재료"].notna()].rename(columns={"수량": "수량(톤)"})
            tf = df[df["분류"].str.contains("두부", na=False) & ~df["분류"].str.contains("마이티펫", na=False)].copy()
            tf = tf[tf["품목"].map(flavor_of).notna()].rename(columns={"수량": "수량(박스)"})
            tf["맛"] = tf["품목"].map(flavor_of)
            mp = df[df["분류"].str.contains("마이티펫", na=False)].copy()
            return (rm[["원재료", "구분", "날짜", "수량(톤)", "비고"]].reset_index(drop=True),
                    tf[["맛", "품목", "구분", "날짜", "수량(박스)", "비고"]].reset_index(drop=True),
                    mp.reset_index(drop=True))
    return None, None, None


def parse_tofu_sheet(xls: pd.ExcelFile):
    """두부모래재고 시트 → DataFrame(맛, 구분, 날짜, 수량(박스), 비고). 없으면 None"""
    name = next((n for n in TOFU_SHEET_NAMES if n in xls.sheet_names), None)
    if name is None:
        return None
    raw = pd.read_excel(xls, name, header=None)
    for i in range(min(10, len(raw))):
        row = [str(v).strip().replace(" ", "") if pd.notna(v) else "" for v in raw.iloc[i].tolist()]
        if "구분" in row and any(h.startswith(("품목", "품명")) for h in row):
            def col(*names):
                for n in names:
                    for j, h in enumerate(row):
                        if h.startswith(n):
                            return j
                return None
            cn, cg, cd, cq, cb = col("품목", "품명"), col("구분"), col("날짜", "입고"), col("수량"), col("비고")
            if None in (cn, cg, cq):
                return None
            body = raw.iloc[i + 1:]
            df = pd.DataFrame({
                "맛": body.iloc[:, cn].map(flavor_of),
                "품목": body.iloc[:, cn].astype("string").str.strip(),
                "구분": body.iloc[:, cg].astype("string").str.strip(),
                "날짜": pd.to_datetime(body.iloc[:, cd], errors="coerce") if cd is not None else pd.NaT,
                "수량(박스)": pd.to_numeric(body.iloc[:, cq], errors="coerce"),
                "비고": body.iloc[:, cb].astype("string") if cb is not None else "",
            })
            df.loc[df["구분"].str.contains("현재", na=False), "구분"] = "현재고"
            df.loc[df["구분"].str.contains("입고|예정", na=False), "구분"] = "입고예정"
            return df[df["맛"].notna() & df["수량(박스)"].notna()].reset_index(drop=True)
    return None


def parse_summary_tofu_stock(xls: pd.ExcelFile) -> dict:
    """요약 시트 입고예정표의 두부모래 현재고 (두부모래재고 시트가 없을 때 임시로 씀)"""
    res = {}
    if SUMMARY_SHEET not in xls.sheet_names:
        return res
    raw = pd.read_excel(xls, SUMMARY_SHEET, header=None)
    for i in range(len(raw)):
        vals = raw.iloc[i].tolist()
        for j, v in enumerate(vals):
            sv = str(v)
            if "두부" in sv and "마이티펫" not in sv and flavor_of(sv) and flavor_of(sv) not in res:
                n = pd.to_numeric(vals[j + 1] if j + 1 < len(vals) else None, errors="coerce")
                if pd.notna(n):
                    res[flavor_of(sv)] = float(n)
    return res


def parse_summary_raw_stock(xls: pd.ExcelFile) -> dict:
    """요약 시트 입고예정표의 '벤토나이트_볼/파쇄' 현재고 (원재료 시트가 없을 때 임시로 씀)"""
    res = {}
    if SUMMARY_SHEET not in xls.sheet_names:
        return res
    raw = pd.read_excel(xls, SUMMARY_SHEET, header=None)
    for i in range(len(raw)):
        vals = raw.iloc[i].tolist()
        for j, v in enumerate(vals):
            name = str(v).replace(" ", "")
            for mat in RAW_MATS:
                if name in (f"벤토나이트_{mat}", f"벤토나이트{mat}") and mat not in res:
                    for w in vals[j + 1:]:
                        n = pd.to_numeric(w, errors="coerce")
                        if pd.notna(n):
                            res[mat] = float(n)
                            break
    return res


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

    # 선택 시트: 상품군목표 — 가로형(상품군 × 월) 또는 세로형(월 | 상품군 | 목표)
    grp_target = {}
    if GROUP_TARGET_SHEET in xls.sheet_names:
        grp_target = parse_group_target(pd.read_excel(xls, GROUP_TARGET_SHEET, header=None))

    raw_mat, tofu, mightypet = parse_stock_in_sheet(xls)
    if raw_mat is None:
        raw_mat = parse_raw_sheet(xls)
    raw_fallback = parse_summary_raw_stock(xls) if raw_mat is None else {}

    if tofu is None:
        tofu = parse_tofu_sheet(xls)
    tofu_fallback = parse_summary_tofu_stock(xls) if tofu is None else {}

    return (buy, out, stock, stock_date, target, prior, names, grp_target, raw_mat, raw_fallback,
            tofu, tofu_fallback, mightypet)


# ─────────────────────────────────────────────
# 상품군 / 표시 형식
# ─────────────────────────────────────────────
GROUPS = [
    ("벤토 모래(볼·파쇄)", ["볼", "파쇄"]),
    ("두부모래(쿠팡)", ["두부"]),
    ("마이티펫(네이버)", ["마이티펫"]),
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
TYPE_SPLIT_GROUPS = ("벤토 모래(볼·파쇄)",)  # 유형별로 펼칠 상품군


def label_to_group(label):
    """상품군목표 시트의 행 이름 → 대시보드 상품군 (합계 행·모르는 이름은 None)"""
    s = norm(label)
    if "합계" in s or "total" in s:
        return None
    for g in GRP_ORDER:
        if s == norm(g):
            return g
    if "마이티펫" in s:
        return "마이티펫(네이버)"
    if "두부" in s:
        return "두부모래(쿠팡)"
    if any(k in s for k in ("벤토", "볼", "파쇄", "제조")):
        return "벤토 모래(볼·파쇄)"
    if "장작" in s:
        return "장작"
    if "필름" in s:
        return "모니터보호필름"
    if "기저귀" in s:
        return "강아지 기저귀"
    if any(k in s for k in ("주방", "리빙", "냄비")):
        return "주방용품"
    if "필터" in s:
        return "차량용 공기필터"
    return None


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
    (buy, out, stock, stock_date, target, prior, sku_names, grp_target_xl,
     raw_mat_xl, raw_fallback, tofu_xl, tofu_fallback, mightypet_xl) = load_workbook(up.getvalue())
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

st.sidebar.caption(f"📄 {up.name}")
st.sidebar.markdown("---")
st.sidebar.caption(
    f"매출 데이터: {buy['날짜'].min():%Y-%m-%d} ~ {buy['날짜'].max():%Y-%m-%d}\n\n"
    f"출고 데이터: ~ {out['날짜'].max():%Y-%m-%d}\n\n"
    + (f"쿠팡 현재고 기준: {stock_date:%Y-%m-%d}" if stock_date is not None else "재고 현황 시트 없음")
)

tab_sales, tab_inv, tab_order, tab_set = st.tabs(
    ["💰 매출 · 쿠팡 PO", "📦 재고 · 결품", "🏭 발주 시뮬레이터", "⚙️ 데이터 설정 · 점검"])
tab_raw, tab_tofu = tab_order.tabs(["원재료 (벤토나이트 볼·파쇄)", "두부모래"])

# ─────────────────────────────────────────────
# 3. 데이터 설정 (다른 탭이 이 값을 쓰므로 먼저 그림)
# ─────────────────────────────────────────────
active_groups = [g for g in GRP_ORDER
                 if buy[(buy["상품군"] == g) & (buy["ym"] > cur_p - 12) & (buy["ym"] <= cur_p)]["금액(원)"].sum() > 0]

with tab_set:
    st.subheader(f"{cur_p.month}월 상품군별 목표")
    st.caption(
        f"엑셀 '{GROUP_TARGET_SHEET}' 시트에서 자동으로 읽습니다 (상품군 × 월 표). "
        "여기서 고친 값은 새로고침하면 사라지니, 계속 쓸 목표는 엑셀에 넣어 주세요."
    )
    xl_month, unmatched = {}, []
    for (per, lab), v in grp_target_xl.items():
        if per != str(cur_p):
            continue
        g = label_to_group(lab)
        if g:
            xl_month[g] = xl_month.get(g, 0) + v
        elif not any(k in norm(lab) for k in ("합계", "total")):
            unmatched.append(lab)
    if unmatched:
        st.warning("상품군을 알아보지 못한 목표 행: " + ", ".join(unmatched))
    tgt_groups = active_groups + [g for g in xl_month if g not in active_groups]
    tgt_df = pd.DataFrame({
        "상품군": tgt_groups,
        "목표(만원)": [xl_month.get(g, np.nan) / 1e4 for g in tgt_groups],
    })
    edited = st.data_editor(
        tgt_df, hide_index=True, use_container_width=False, key=f"grp_tgt_{cur_p}",
        disabled=["상품군"],
        column_config={"목표(만원)": st.column_config.NumberColumn(min_value=0, step=100, format="%.0f")},
    )
    grp_target = {g: float(v) * 1e4 for g, v in zip(edited["상품군"], edited["목표(만원)"]) if pd.notna(v) and v > 0}
    summary_tgt = target.get((cur_p.year, cur_p.month))
    grp_sum = sum(grp_target.values())
    # 공식 목표 = 상품군 목표 합계. 상품군 목표가 하나도 없을 때만 요약 시트 목표매출을 씀
    total_tgt = grp_sum if grp_sum else summary_tgt
    st.caption(f"'전체' 목표 = 상품군 목표 합계 {fmt_krw(grp_sum)}" if grp_sum else
               f"상품군 목표가 없어 '전체'는 요약 시트 목표매출({fmt_krw(summary_tgt)})을 씁니다.")

    st.markdown("---")
    st.subheader("재고 계산 설정")
    sc1, sc2 = st.columns([1, 2])
    window = sc1.selectbox("일평균 출고 기준 기간 (현재고 기준일까지)", [7, 14, 30, 90], index=0,
                           format_func=lambda d: f"최근 {d}일")
    types_in_stock = sort_types(stock["유형"].unique()) if stock is not None else []
    lead = {}
    with sc2.expander("유형별 쿠팡 납품 리드타임(일)", expanded=False):
        st.caption("쿠팡 PO를 받고 VF에 입고되기까지 걸리는 기간입니다. 실제에 맞게 고쳐 주세요.")
        lc = st.columns(4)
        for i, t in enumerate(types_in_stock):
            lead[t] = lc[i % 4].number_input(t, min_value=0, max_value=180, value=DEFAULT_LEAD.get(t, 5),
                                             step=1, key=f"lead_{t}")

    st.markdown("---")
    st.subheader("원재료(벤토나이트) 발주 설정")
    r1, r2, r3, r4 = st.columns(4)
    raw_lead = r1.number_input("발주 → 공장 입고 (일)", min_value=0, max_value=180, value=40, step=1, key="raw_lead")
    raw_safety = r2.number_input("안전재고 (며칠분)", min_value=0, max_value=90, value=14, step=1, key="raw_safety")
    raw_yield = r3.number_input("수율 (제품 톤 ÷ 원재료 톤)", min_value=0.5, max_value=1.2, value=1.0, step=0.01,
                                key="raw_yield")
    raw_window = r4.selectbox("하루 사용량 기준 기간", [30, 60, 90], index=2, format_func=lambda d: f"최근 {d}일",
                              key="raw_window")
    u1, u2, _, _ = st.columns(4)
    raw_unit = {
        "볼": u1.selectbox("볼 발주 단위(톤)", [220, 330], index=1, key="unit_볼"),
        "파쇄": u2.selectbox("파쇄 발주 단위(톤)", [220, 330], index=0, key="unit_파쇄"),
    }

    if raw_mat_xl is not None:
        st.caption(f"원재료 현재고·입고예정은 엑셀 '{RAW_SHEET}' 시트에서 읽었습니다. 아래에서 고친 값은 새로고침하면 사라집니다.")
        raw_init = raw_mat_xl[["원재료", "구분", "날짜", "수량(톤)", "비고"]].copy()
    else:
        st.warning(
            f"엑셀에 '{RAW_SHEET}' 시트가 없어 요약 시트 입고예정표의 현재고만 임시로 넣었습니다. "
            "입고 예정분은 아래 표에 직접 추가하거나, 엑셀에 시트를 만들어 주세요 "
            "(열: 분류 | 품목 | 구분 | 날짜 | 수량 | 단위 | 비고)."
        )
        raw_init = pd.DataFrame([
            {"원재료": m, "구분": "현재고", "날짜": base_ts, "수량(톤)": v, "비고": "요약 시트에서 읽음 · 확인 필요"}
            for m, v in raw_fallback.items()
        ], columns=["원재료", "구분", "날짜", "수량(톤)", "비고"])
    raw_init["날짜"] = pd.to_datetime(raw_init["날짜"], errors="coerce").dt.date
    raw_init["비고"] = raw_init["비고"].fillna("").astype(str)
    raw_rows = st.data_editor(
        raw_init, num_rows="dynamic", hide_index=True, use_container_width=True, key="raw_editor",
        column_config={
            "원재료": st.column_config.SelectboxColumn(options=RAW_MATS, required=True),
            "구분": st.column_config.SelectboxColumn(options=["현재고", "입고예정"], required=True),
            "날짜": st.column_config.DateColumn("날짜 (현재고 기준일 / 입고예정일)", format="YYYY-MM-DD"),
            "수량(톤)": st.column_config.NumberColumn(min_value=0, format="%.1f"),
            "비고": st.column_config.TextColumn("비고"),
        },
    )


with tab_set:
    st.markdown("---")
    st.subheader("두부모래 발주 설정")
    t1, t2, t3, t4 = st.columns(4)
    tofu_lead = t1.number_input("발주 → 물류 입고 (일)", min_value=0, max_value=180, value=40, step=1, key="tofu_lead")
    tofu_safety = t2.number_input("안전재고 (며칠분)", min_value=0, max_value=90, value=14, step=1, key="tofu_safety")
    tofu_unit = t3.number_input("발주 단위 (박스)", min_value=1, max_value=20000, value=1320, step=60, key="tofu_unit")
    tofu_window = t4.selectbox("일평균 출고 기준 기간", [7, 14, 30, 60, 90], index=2, format_func=lambda d: f"최근 {d}일",
                               key="tofu_window")
    tofu_incl_vf = st.checkbox("쿠팡 VF 현재고도 우리 재고에 더해서 보기", value=True, key="tofu_vf",
                               help="켜면 '평택 재고 + 쿠팡 VF 재고'가 소비자 출고로 줄어드는 것으로 계산합니다.")
    if tofu_xl is not None:
        st.caption("두부모래 평택 재고·입고예정은 엑셀 '재고입고관리' 시트에서 읽었습니다. 아래에서 고친 값은 새로고침하면 사라집니다.")
        tofu_init = tofu_xl[["품목", "구분", "날짜", "수량(박스)", "비고"]].copy()
    else:
        st.warning("엑셀에 '재고입고관리' 시트가 없어 요약 시트의 현재고만 임시로 넣었습니다 (입고예정 없음). "
                   "시트를 만들어 주세요 (열: 분류 | 품목 | 구분 | 날짜 | 수량 | 단위 | 비고).")
        tofu_init = pd.DataFrame([
            {"품목": f"7L 두부_{k}", "구분": "현재고", "날짜": base_ts, "수량(박스)": v, "비고": "요약 시트에서 읽음 · 확인 필요"}
            for k, v in tofu_fallback.items()
        ], columns=["품목", "구분", "날짜", "수량(박스)", "비고"])
    tofu_init["날짜"] = pd.to_datetime(tofu_init["날짜"], errors="coerce").dt.date
    tofu_init["비고"] = tofu_init["비고"].fillna("").astype(str)
    tofu_rows = st.data_editor(
        tofu_init, num_rows="dynamic", hide_index=True, use_container_width=True, key="tofu_editor",
        column_config={
            "구분": st.column_config.SelectboxColumn(options=["현재고", "입고예정"], required=True),
            "날짜": st.column_config.DateColumn("날짜 (현재고 기준일 / 입고예정일)", format="YYYY-MM-DD"),
            "수량(박스)": st.column_config.NumberColumn(min_value=0, format="%d"),
            "비고": st.column_config.TextColumn("비고"),
        },
    )

AVG_INFO = {}


def compute_inv():
    if stock is None:
        return None, None
    last_out = out["날짜"].max()
    # 일평균 출고: 현재고 기준일까지 최근 window일. 출고 데이터가 기준일보다 짧으면 있는 날짜만으로 나눔
    anchor = stock_date if stock_date is not None else last_out
    w_start = anchor - pd.Timedelta(days=window - 1)
    w_end = min(anchor, last_out)
    covered = max((w_end - w_start).days + 1, 1)
    win = out[(out["날짜"] >= w_start) & (out["날짜"] <= w_end)]
    avg = win.groupby("sku id")["출고량"].sum() / covered
    AVG_INFO.update(start=w_start, end=w_end, days=covered, anchor=anchor)

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
    card_groups = [g for g in GRP_ORDER if g in active_groups or g in grp_target]
    for g in card_groups:
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
    split = st.toggle("세부 펼치기 (벤토 모래 → 볼·파쇄 / 기저귀·주방용품 → 브랜드)", value=False)
    rows = []
    for g in active_groups:
        gdf = buy[buy["상품군"] == g]
        rows.append((g, gdf, "normal"))
        if split and g in TYPE_SPLIT_GROUPS:
            for t in sort_types(gdf["유형"].unique()):
                rows.append((f"└ {t}", gdf[gdf["유형"] == t], "sub"))
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
            f"현재고 기준 {stock_date:%m/%d} (벤플 수불)  ·  일평균 출고 = {AVG_INFO['start']:%m/%d}~{AVG_INFO['end']:%m/%d} "
            f"{AVG_INFO['days']}일 평균"
            + (f" (출고 데이터가 {last_out:%m/%d}까지라 {window}일 중 {AVG_INFO['days']}일만 반영)"
               if AVG_INFO['days'] < window else "") + "  ·  "
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
            d3.metric(f"일평균 출고 ({AVG_INFO['start']:%m/%d}~{AVG_INFO['end']:%m/%d})", f"{r['일평균출고']:,.1f}")
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
# 원재료 발주
# ─────────────────────────────────────────────
RAW_HORIZON = 150


def raw_plan(mat: str, rows: pd.DataFrame):
    rows = rows.copy()
    rows["날짜"] = pd.to_datetime(rows["날짜"], errors="coerce")
    rows = rows[(rows["원재료"] == mat) & rows["수량(톤)"].notna()]
    st_rows = rows[rows["구분"] == "현재고"]
    if st_rows.empty:
        return None
    srow = st_rows.sort_values("날짜", na_position="first").iloc[-1]
    start = srow["날짜"].normalize() if pd.notna(srow["날짜"]) else base_ts
    stock0 = float(srow["수량(톤)"])

    inb = rows[rows["구분"] == "입고예정"]
    inb = inb[~inb["비고"].astype(str).str.contains("완료", na=False)]
    undated = inb[inb["날짜"].isna()]
    inb = inb[inb["날짜"].notna()]
    # 현재고 날짜 이전(같은 날 포함) 입고분은 이미 현재고에 들어 있다고 보고 뺌
    included = inb[inb["날짜"].dt.normalize() <= start]
    inb = inb[inb["날짜"].dt.normalize() > start]
    arrivals = {}
    for d, q in zip(inb["날짜"], inb["수량(톤)"]):
        d = d.normalize()
        arrivals[d] = arrivals.get(d, 0) + float(q)

    prod = buy[(buy["구분"] == "제품") & (buy["유형"] == mat) & (buy["날짜"] <= base_ts)
               & (buy["날짜"] > base_ts - pd.Timedelta(days=raw_window))]
    daily = float(prod["수량(톤)"].sum()) / raw_window / raw_yield
    safety = daily * raw_safety
    days = pd.date_range(start, periods=RAW_HORIZON + 1)

    def sim(arr):
        # 재고는 0 밑으로 내려가지 않음 (모자란 만큼은 생산을 못 한 것으로 봄)
        lv, res = stock0, []
        for i, d in enumerate(days):
            if i > 0:
                lv = max(lv - daily, 0.0)
            lv += arr.get(d, 0)
            res.append(lv)
        return np.array(res)

    base_traj = sim(arrivals)
    neg = np.where(base_traj <= 0)[0] if daily > 0 else []
    runout = days[neg[0]] if len(neg) else None

    # 오늘 발주해도 리드타임 안에는 못 받으므로, 그 기간의 부족은 '발주로 못 막는 부족'으로 따로 셈
    earliest = base_ts + pd.Timedelta(days=int(raw_lead))
    in_lead = days < earliest
    gap_in_lead = int(((base_traj <= 0) & in_lead).sum()) if daily > 0 else 0
    low_in_lead = bool(((base_traj < safety) & in_lead).any())

    arr2, planned, last_arrive = dict(arrivals), [], earliest - pd.Timedelta(days=1)
    if daily > 0:
        for _ in range(8):
            traj = sim(arr2)
            cand = [i for i in np.where(traj < safety)[0] if days[i] > last_arrive]
            if not cand:
                break
            need = days[cand[0]]
            order = max(need - pd.Timedelta(days=int(raw_lead)), base_ts)
            arrive = order + pd.Timedelta(days=int(raw_lead))
            arr2[arrive] = arr2.get(arrive, 0) + raw_unit[mat]
            planned.append({"발주일": order, "입고예정일": arrive, "수량(톤)": raw_unit[mat],
                            "재고가 안전재고 아래로": need})
            last_arrive = arrive
            if arrive > days[-1]:
                break
    return {"start": start, "stock0": stock0, "arrivals": arrivals, "undated": undated, "daily": daily,
            "included": included,
            "safety": safety, "days": days, "base": base_traj, "plan_traj": sim(arr2), "runout": runout,
            "planned": planned, "gap_in_lead": gap_in_lead, "low_in_lead": low_in_lead, "earliest": earliest}


with tab_raw:
    st.subheader("원재료(벤토나이트) 발주 필요 시기")
    st.caption("입고가 끝나면 현재고 수량과 날짜만 바꾸면 됩니다. 현재고 날짜 이전 입고예정은 자동으로 빠집니다.")
    st.caption(
        f"하루 사용량 = 최근 {raw_window}일 볼·파쇄 제품 납품 톤 ÷ {raw_window} ÷ 수율 {raw_yield:g}  ·  "
        f"발주 → 공장 입고 {raw_lead}일  ·  안전재고 {raw_safety}일분  ·  "
        "발주 필요일 = 재고가 안전재고 아래로 내려가는 날 − 리드타임.  "
        "평택 완제품 재고 변화는 반영하지 않습니다 (설정은 '데이터 설정' 탭)."
    )
    if raw_mat_xl is None:
        st.error(
            f"⚠️ 올린 엑셀에 '{RAW_SHEET}' 시트가 없어, 숨겨진 요약 시트의 현재고만 임시로 쓰고 있습니다 "
            "(입고 예정분 없음 → 결과가 틀립니다). '재고입고관리' 시트가 있는 최신 엑셀을 저장한 뒤 다시 올려 주세요."
        )
    else:
        st.success(f"'{RAW_SHEET}' 시트에서 원재료 현재고·입고예정을 읽었습니다.")
    for mat in RAW_MATS:
        P = raw_plan(mat, raw_rows)
        st.markdown(f"#### 벤토나이트 {mat}")
        if P is None:
            st.info(f"{mat} 원재료 현재고가 없습니다. '데이터 설정' 탭 표나 엑셀 '{RAW_SHEET}' 시트에 넣어 주세요.")
            continue
        planned = P["planned"]
        nxt = planned[0] if planned else None
        if P["daily"] <= 0:
            status = "⚪ 최근 사용량 없음"
        elif nxt is None:
            status = f"🟢 {RAW_HORIZON}일 안에 추가 발주 불필요"
        elif (nxt["발주일"] - base_ts).days <= 0:
            status = f"🟠 오늘 발주 필요 ({nxt['수량(톤)']}톤 → {nxt['입고예정일']:%m/%d} 입고)"
        elif (nxt["발주일"] - base_ts).days <= 7:
            status = f"🟠 {nxt['발주일']:%m/%d}까지 발주 ({(nxt['발주일'] - base_ts).days}일 남음)"
        else:
            status = f"🟢 {nxt['발주일']:%m/%d}까지 발주 ({(nxt['발주일'] - base_ts).days}일 남음)"
        st.markdown(f"**{status}**")
        if P["gap_in_lead"]:
            st.error(f"🔴 앞으로 {raw_lead}일({P['earliest']:%m/%d}) 안에 약 {P['gap_in_lead']}일 재고 공백이 예상됩니다. "
                     "지금 발주해도 그 전에는 못 받으니, 이미 발주한 물량의 입고를 앞당기거나 생산 계획을 조정해야 합니다.")
        elif P["low_in_lead"]:
            st.caption(f"⚠️ {P['earliest']:%m/%d} 전에 재고가 안전재고 아래로 잠깐 내려갑니다 (재고 공백은 없음).")

        m1, m2, m3, m4, m5 = st.columns(5)
        cover = P["stock0"] / P["daily"] if P["daily"] > 0 else np.inf
        m1.metric(f"현재고 ({P['start']:%m/%d})", f"{P['stock0']:,.0f}톤",
                  f"{cover:,.1f}일분" if np.isfinite(cover) else None, delta_color="off")
        m2.metric("입고 예정 (확정)", f"{sum(P['arrivals'].values()):,.0f}톤",
                  f"{len(P['arrivals'])}건", delta_color="off")
        m3.metric("하루 사용량", f"{P['daily']:,.1f}톤", f"월 {P['daily'] * 30:,.0f}톤", delta_color="off")
        m4.metric("소진 예상일 (추가 발주 없을 때)", f"{P['runout']:%m/%d}" if P["runout"] is not None else "-",
                  f"{(P['runout'] - base_ts).days}일 뒤" if P["runout"] is not None else f"{RAW_HORIZON}일 이상",
                  delta_color="off")
        m5.metric("다음 발주 필요일", f"{nxt['발주일']:%m/%d}" if nxt else "-",
                  f"{nxt['수량(톤)']}톤" if nxt else None, delta_color="off")

        if len(P["undated"]):
            st.warning(f"입고일이 비어 있는 입고예정 {len(P['undated'])}건({P['undated']['수량(톤)'].sum():,.0f}톤)은 "
                       "계산에서 뺐습니다. 날짜를 넣어 주세요.")
        if len(P["included"]):
            st.caption("✔️ 현재고 날짜 이전 입고분은 현재고에 이미 들어 있다고 보고 뺐습니다: "
                       + ", ".join(f"{d:%m/%d} {q:,.0f}톤" for d, q in zip(P["included"]["날짜"], P["included"]["수량(톤)"])))
        if P["start"] < base_ts - pd.Timedelta(days=3):
            st.caption(f"⚠️ 현재고 기준일({P['start']:%m/%d})이 기준일보다 오래됐습니다. 최신 재고로 바꿔 주세요.")

        g1, g2 = st.columns([3, 2])
        with g1:
            fig = go.Figure()
            fig.add_scatter(x=P["days"], y=P["base"], name="확정 입고만", mode="lines", line=dict(width=2))
            if planned:
                fig.add_scatter(x=P["days"], y=P["plan_traj"], name="권장 발주 포함", mode="lines",
                                line=dict(width=2, dash="dash"))
            if P["arrivals"]:
                fig.add_bar(x=list(P["arrivals"].keys()), y=list(P["arrivals"].values()), name="입고 예정",
                            opacity=0.55)
            if planned:
                fig.add_bar(x=[p["입고예정일"] for p in planned], y=[p["수량(톤)"] for p in planned],
                            name="권장 발주 입고", opacity=0.35)
            fig.add_hline(y=P["safety"], line_dash="dot", annotation_text=f"안전재고 {P['safety']:,.0f}톤")
            fig.add_hline(y=0, line_color="red", line_width=1)
            fig.update_layout(title=f"벤토나이트 {mat} 원재료 재고 예상 ({RAW_HORIZON}일, 톤)", height=360,
                              legend=dict(orientation="h", y=-0.18), margin=dict(t=50, b=10))
            st.plotly_chart(fig, use_container_width=True)
        with g2:
            st.markdown("**권장 발주 일정**")
            if planned:
                pdf = pd.DataFrame(planned)
                st.dataframe(
                    pdf[["발주일", "입고예정일", "수량(톤)", "재고가 안전재고 아래로"]], hide_index=True,
                    use_container_width=True,
                    column_config={"발주일": st.column_config.DateColumn(format="MM/DD"),
                                   "입고예정일": st.column_config.DateColumn(format="MM/DD"),
                                   "재고가 안전재고 아래로": st.column_config.DateColumn("안전재고 미달일", format="MM/DD")})
            else:
                st.caption("앞으로 추가 발주가 필요하지 않습니다.")
            if P["daily"] > 0:
                st.caption(f"발주 단위별 버티는 기간: 220톤 ≈ {220 / P['daily']:,.0f}일, "
                           f"330톤 ≈ {330 / P['daily']:,.0f}일")
        st.markdown("---")


# ─────────────────────────────────────────────
# 두부모래 발주 시뮬레이터
# ─────────────────────────────────────────────
def plan_stock(stock0, start, arrivals, daily, lead, safety_days, unit, horizon=150):
    """재고 예상 + 권장 발주 일정. 재고는 0 밑으로 안 내려감(모자란 날 = 공백)"""
    days = pd.date_range(start, periods=horizon + 1)
    safety = daily * safety_days

    def sim(arr):
        lv, res = stock0, []
        for i, d in enumerate(days):
            if i > 0:
                lv = max(lv - daily, 0.0)
            lv += arr.get(d, 0)
            res.append(lv)
        return np.array(res)

    base_traj = sim(arrivals)
    neg = np.where(base_traj <= 0)[0] if daily > 0 else []
    runout = days[neg[0]] if len(neg) else None
    earliest = base_ts + pd.Timedelta(days=int(lead))
    in_lead = days < earliest
    gap_in_lead = int(((base_traj <= 0) & in_lead).sum()) if daily > 0 else 0

    arr2, planned, last_arrive = dict(arrivals), [], earliest - pd.Timedelta(days=1)
    if daily > 0:
        for _ in range(8):
            traj = sim(arr2)
            cand = [i for i in np.where(traj < safety)[0] if days[i] > last_arrive]
            if not cand:
                break
            need = days[cand[0]]
            order = max(need - pd.Timedelta(days=int(lead)), base_ts)
            arrive = order + pd.Timedelta(days=int(lead))
            arr2[arrive] = arr2.get(arrive, 0) + unit
            planned.append({"발주일": order, "입고예정일": arrive, "수량": unit, "안전재고 미달일": need})
            last_arrive = arrive
            if arrive > days[-1]:
                break
    return {"days": days, "base": base_traj, "plan_traj": sim(arr2), "runout": runout, "planned": planned,
            "safety": safety, "gap_in_lead": gap_in_lead, "earliest": earliest, "daily": daily}


def plan_status(P):
    nxt = P["planned"][0] if P["planned"] else None
    if P["daily"] <= 0:
        return "⚪ 출고 없음"
    if P["gap_in_lead"]:
        return f"🔴 {P['gap_in_lead']}일 결품 예상"
    if nxt is None:
        return "🟢 발주 불필요"
    left = (nxt["발주일"] - base_ts).days
    if left <= 0:
        return "🟠 오늘 발주"
    if left <= 7:
        return f"🟠 {left}일 내 발주"
    return f"🟢 {left}일 남음"


with tab_tofu:
    st.subheader("두부모래 발주 시뮬레이터")
    st.caption("입고가 끝나면 현재고 수량과 날짜만 바꾸면 됩니다. 현재고 날짜 이전 입고예정은 자동으로 빠집니다.")
    if tofu_xl is None:
        st.error("⚠️ 엑셀에 '재고입고관리' 시트가 없어 요약 시트 현재고만 쓰고 있습니다 (입고예정 없음 → 결과가 틀립니다).")
    tofu_skus = []
    if stock is not None:
        tofu_skus = [(int(r["sku id"]), r["품명"]) for _, r in stock[stock["유형"] == "두부"].iterrows()]
    if not tofu_skus:
        tofu_skus = [(int(k), v) for k, v in out[out["유형"] == "두부"].groupby("sku id")["품명"].last().items()]

    end = min(stock_date if stock_date is not None else out["날짜"].max(), out["날짜"].max())
    t_start = end - pd.Timedelta(days=tofu_window - 1)
    st.caption(
        f"재고 = 평택 재고{' + 쿠팡 VF 현재고' if tofu_incl_vf else ''} + 입고예정  ·  "
        f"하루 출고 = 쿠팡 소비자 출고 {t_start:%m/%d}~{end:%m/%d} 평균  ·  발주 → 물류 입고 {tofu_lead}일  ·  "
        f"안전재고 {tofu_safety}일분  ·  발주 단위 {tofu_unit:,}박스 (설정은 '데이터 설정' 탭)"
    )

    trows = tofu_rows.copy()
    trows["맛"] = trows["품목"].map(flavor_of)
    trows["날짜"] = pd.to_datetime(trows["날짜"], errors="coerce")
    plans, summary = {}, []
    for sku, name in tofu_skus:
        fl = flavor_of(name)
        rows = trows[(trows["맛"] == fl) & trows["수량(박스)"].notna()]
        cur = rows[rows["구분"] == "현재고"].sort_values("날짜", na_position="first")
        wh = float(cur["수량(박스)"].iloc[-1]) if len(cur) else 0.0
        start = cur["날짜"].iloc[-1].normalize() if len(cur) and pd.notna(cur["날짜"].iloc[-1]) else base_ts
        vf = 0.0
        if tofu_incl_vf and stock is not None:
            hit = stock.loc[stock["sku id"] == sku, "현재고"]
            vf = float(hit.iloc[0]) if len(hit) else 0.0
        inb = rows[(rows["구분"] == "입고예정") & ~rows["비고"].astype(str).str.contains("완료", na=False)]
        undated = int(inb["날짜"].isna().sum())
        dated = inb[inb["날짜"].notna()]
        included = dated[dated["날짜"].dt.normalize() <= start]  # 현재고 날짜 이전 입고분은 이미 현재고에 포함
        dated = dated[dated["날짜"].dt.normalize() > start]
        arrivals = {}
        for d, q in zip(dated["날짜"], dated["수량(박스)"]):
            d = d.normalize()
            arrivals[d] = arrivals.get(d, 0) + float(q)
        so = out[(out["sku id"] == sku) & (out["날짜"] >= t_start) & (out["날짜"] <= end)]["출고량"].sum()
        daily = float(so) / tofu_window
        P = plan_stock(wh + vf, start, arrivals, daily, tofu_lead, tofu_safety, int(tofu_unit))
        P.update(sku=sku, name=name, flavor=fl, wh=wh, vf=vf, arrivals=arrivals, start=start, undated=undated,
                 included=included)
        plans[sku] = P
        nxt = P["planned"][0] if P["planned"] else None
        total = wh + vf
        summary.append({
            "상태": plan_status(P), "품목": name, "평택 재고": wh, "쿠팡 VF": vf if tofu_incl_vf else np.nan,
            "입고예정": sum(arrivals.values()), "하루 출고": daily,
            "재고 일수": (total / daily) if daily > 0 else np.nan,
            "입고 포함 일수": ((total + sum(arrivals.values())) / daily) if daily > 0 else np.nan,
            "소진 예상일": P["runout"], "다음 발주일": nxt["발주일"] if nxt else pd.NaT,
            "권장 수량": nxt["수량"] if nxt else np.nan, "sku": sku,
        })

    sdf = pd.DataFrame(summary).sort_values("다음 발주일", na_position="last").reset_index(drop=True)
    n_red = int(sdf["상태"].str.startswith("🔴").sum())
    n_or = int(sdf["상태"].str.startswith("🟠").sum())
    k1, k2, k3 = st.columns(3)
    k1.metric("🔴 결품 예상", f"{n_red}개 품목")
    k2.metric("🟠 1주 내 발주", f"{n_or}개 품목")
    k3.metric("합계 권장 발주 (다음 1회)", f"{sdf['권장 수량'].sum():,.0f}박스")

    st.caption("👇 행을 클릭하면 아래에 품목별 재고 예상이 나옵니다.")
    ev = st.dataframe(
        sdf.drop(columns="sku"), hide_index=True, use_container_width=True, on_select="rerun",
        selection_mode="single-row", key="tofu_table",
        column_config={
            "평택 재고": st.column_config.NumberColumn(format="localized"),
            "쿠팡 VF": st.column_config.NumberColumn(format="localized"),
            "입고예정": st.column_config.NumberColumn(format="localized"),
            "하루 출고": st.column_config.NumberColumn(format="%.1f"),
            "재고 일수": st.column_config.NumberColumn(format="%.0f"),
            "입고 포함 일수": st.column_config.NumberColumn(format="%.0f"),
            "소진 예상일": st.column_config.DateColumn(format="MM/DD"),
            "다음 발주일": st.column_config.DateColumn(format="MM/DD"),
            "권장 수량": st.column_config.NumberColumn(format="localized"),
        })

    sel = getattr(getattr(ev, "selection", None), "rows", []) or []
    if len(sdf):
        idx0 = sel[0] if sel and sel[0] < len(sdf) else 0
        sku = st.selectbox("품목 상세", sdf["sku"].tolist(), index=idx0,
                           format_func=lambda s: plans[s]["name"], key="tofu_pick")
        P = plans[sku]
        st.markdown(f"#### {P['name']}  ·  {plan_status(P)}")
        if P["gap_in_lead"]:
            st.error(f"앞으로 {tofu_lead}일({P['earliest']:%m/%d}) 안에 약 {P['gap_in_lead']}일 결품이 예상됩니다. "
                     "지금 발주해도 그 전에는 못 받으니, 입고를 앞당기거나 쿠팡 PO 대응을 조정해야 합니다.")
        if P["undated"]:
            st.warning(f"입고일이 비어 있는 입고예정 {P['undated']}건은 계산에서 뺐습니다.")
        if len(P["included"]):
            st.caption("✔️ 현재고 날짜 이전 입고분은 현재고에 이미 들어 있다고 보고 뺐습니다: "
                       + ", ".join(f"{d:%m/%d} {q:,.0f}박스" for d, q in zip(P["included"]["날짜"], P["included"]["수량(박스)"])))
        nxt = P["planned"][0] if P["planned"] else None
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric(f"평택 재고 ({P['start']:%m/%d})", f"{P['wh']:,.0f}")
        m2.metric("쿠팡 VF 현재고", f"{P['vf']:,.0f}" if tofu_incl_vf else "제외")
        m3.metric("입고 예정", f"{sum(P['arrivals'].values()):,.0f}", f"{len(P['arrivals'])}건", delta_color="off")
        m4.metric("하루 출고", f"{P['daily']:,.1f}", f"월 {P['daily'] * 30:,.0f}", delta_color="off")
        m5.metric("다음 발주일", f"{nxt['발주일']:%m/%d}" if nxt else "-",
                  f"{nxt['수량']:,}박스 → {nxt['입고예정일']:%m/%d} 입고" if nxt else None, delta_color="off")
        g1, g2 = st.columns([3, 2])
        with g1:
            fig = go.Figure()
            fig.add_scatter(x=P["days"], y=P["base"], name="확정 입고만", mode="lines", line=dict(width=2))
            if P["planned"]:
                fig.add_scatter(x=P["days"], y=P["plan_traj"], name="권장 발주 포함", mode="lines",
                                line=dict(width=2, dash="dash"))
            if P["arrivals"]:
                fig.add_bar(x=list(P["arrivals"].keys()), y=list(P["arrivals"].values()), name="입고 예정", opacity=0.55)
            if P["planned"]:
                fig.add_bar(x=[p["입고예정일"] for p in P["planned"]], y=[p["수량"] for p in P["planned"]],
                            name="권장 발주 입고", opacity=0.35)
            fig.add_hline(y=P["safety"], line_dash="dot", annotation_text=f"안전재고 {P['safety']:,.0f}")
            fig.add_hline(y=0, line_color="red", line_width=1)
            fig.update_layout(title="재고 예상 (150일, 박스)", height=360,
                              legend=dict(orientation="h", y=-0.18), margin=dict(t=50, b=10))
            st.plotly_chart(fig, use_container_width=True)
        with g2:
            st.markdown("**권장 발주 일정**")
            if P["planned"]:
                st.dataframe(pd.DataFrame(P["planned"]), hide_index=True, use_container_width=True,
                             column_config={"발주일": st.column_config.DateColumn(format="MM/DD"),
                                            "입고예정일": st.column_config.DateColumn(format="MM/DD"),
                                            "안전재고 미달일": st.column_config.DateColumn(format="MM/DD"),
                                            "수량": st.column_config.NumberColumn("수량(박스)", format="localized")})
            else:
                st.caption("150일 안에 추가 발주가 필요하지 않습니다.")
            if P["daily"] > 0:
                st.caption(f"발주 1회({tofu_unit:,}박스)로 버티는 기간 ≈ {tofu_unit / P['daily']:,.0f}일")

    if mightypet_xl is not None and len(mightypet_xl):
        with st.expander(f"마이티펫 두부모래 (네이버 · 글로브 직송) 재고·입고 {len(mightypet_xl)}건"):
            st.caption("마이티펫은 발주 계산은 하지 않고, 재고입고관리 시트 내용만 보여 줍니다.")
            st.dataframe(mightypet_xl[["품목", "구분", "날짜", "수량", "비고"]], hide_index=True,
                         use_container_width=True,
                         column_config={"날짜": st.column_config.DateColumn(format="YYYY-MM-DD"),
                                        "수량": st.column_config.NumberColumn(format="localized")})


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

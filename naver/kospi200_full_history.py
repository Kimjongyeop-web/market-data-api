# -*- coding: utf-8 -*-
"""
KOSPI200 407종목 FULL HISTORY 재구축
====================================

목적
----
- 2008~2026 역사적 KOSPI200 종목 407개의 가격을 "편입 여부와 무관하게" 수집한다.
- 어떤 종목이 KOSPI200에 들어오기 전/나간 후에도, 해당 날짜에 네이버 가격이 있으면 그대로 보존한다.
- membership은 가격 수집/마스킹에 절대 사용하지 않고, QC와 연구 유니버스 판정에만 사용한다.
- 기존 엑셀의 가격/거래량은 fallback으로 보존하고, 네이버에서 추가로 받은 관측치로 빈 구간을 채운다.
- 최종 결과 엑셀은 정확히 2개만 만든다.

필수 파일(이 .py와 같은 폴더)
-----------------------------
데이터 파일 후보 중 1개:
    00_KOSPI200_전체데이터_문자형_2008_2026.xlsx
    KOSPI200_분석데이터_2008_2026.xlsx

QC 파일 후보 중 1개:
    00_데이터_안내_QC검토표_2008_2026.xlsx
    00_KOSPI200_2008_2026_QC검토표.xlsx
    KOSPI200_QC안내검토표_2008_2026.xlsx

필수 패키지
-----------
pip install pandas numpy requests openpyxl

최종 출력(딱 2개)
----------------
FINAL_FULL_HISTORY/
    KOSPI200_분석데이터_FULL_HISTORY_2008_2026.xlsx
    KOSPI200_QC안내검토표_FULL_HISTORY_2008_2026.xlsx
"""

from __future__ import annotations

import ast
import html
import re
import time
from pathlib import Path

import numpy as np
import pandas as pd
import requests


# =============================================================================
# 0. SETTINGS
# =============================================================================

try:
    BASE_DIR = Path(__file__).resolve().parent
except NameError:
    BASE_DIR = Path.cwd()

DATA_INPUT_CANDIDATES = [
    "00_KOSPI200_전체데이터_문자형_2008_2026.xlsx",
    "KOSPI200_분석데이터_2008_2026.xlsx",
    "00_KOSPI200_전체데이터_2008_2026.xlsx",
]

QC_INPUT_CANDIDATES = [
    "00_데이터_안내_QC검토표_2008_2026.xlsx",
    "00_KOSPI200_2008_2026_QC검토표.xlsx",
    "KOSPI200_QC안내검토표_2008_2026.xlsx",
]

OUT_DIR = BASE_DIR / "FINAL_FULL_HISTORY"
OUT_DATA = OUT_DIR / "KOSPI200_분석데이터_FULL_HISTORY_2008_2026.xlsx"
OUT_QC = OUT_DIR / "KOSPI200_QC안내검토표_FULL_HISTORY_2008_2026.xlsx"

# 기존 kp.py raw가 있으면 fallback/cache로 읽기만 한다.
LEGACY_RAW_DIR = BASE_DIR / "data_kospi200_reversal" / "03_prices" / "naver_raw"

NAVER_URL = "https://fchart.stock.naver.com/siseJson.nhn"
REQUEST_TIMEOUT = 25
MAX_RETRIES = 3
REQUEST_SLEEP_SECONDS = 0.45

# 첫 연구일 수익률 계산에 필요한 직전 가격을 조금 더 받는다.
PRE_SAMPLE_BUFFER_DAYS = 45

# 네이버 전체범위 단일 요청이 실패할 때만 연도 chunk로 재시도.
FALLBACK_CHUNK_YEARS = 5

# 원본에 407개가 있는지 확인. 다르면 중단해서 잘못된 파일을 만들지 않는다.
EXPECTED_TICKERS = 407


# =============================================================================
# 1. HELPERS
# =============================================================================

def normalize_ticker(value) -> str:
    s = str(value).strip()

    if s.lower() in {"", "nan", "none"}:
        return ""

    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    # 영숫자 임시 종목코드(예: 0126Z0)는 원문 유지.
    if s.isdigit():
        s = s.zfill(6)

    return s.upper()


def locate_file(candidates: list[str], label: str) -> Path:
    for name in candidates:
        p = BASE_DIR / name
        if p.exists():
            return p

    # 정확한 이름이 조금 달라도 찾기.
    xlsx = sorted(BASE_DIR.glob("*.xlsx"))
    for p in xlsx:
        text = p.name.lower()
        if label == "data" and ("전체데이터" in p.name or "분석데이터" in p.name):
            return p
        if label == "qc" and ("qc" in text or "검토표" in p.name):
            return p

    raise FileNotFoundError(
        f"{label} 입력 엑셀을 찾지 못했습니다.\n"
        f"현재 폴더: {BASE_DIR}\n"
        f"후보: {candidates}"
    )


def find_sheet_name(xls: pd.ExcelFile, candidates: list[str]) -> str | None:
    normalized = {
        re.sub(r"[\s_]+", "", str(s)).lower(): s
        for s in xls.sheet_names
    }

    for c in candidates:
        key = re.sub(r"[\s_]+", "", c).lower()
        if key in normalized:
            return normalized[key]

    # 부분 포함 fallback
    for s in xls.sheet_names:
        key = re.sub(r"[\s_]+", "", str(s)).lower()
        for c in candidates:
            ckey = re.sub(r"[\s_]+", "", c).lower()
            if ckey in key or key in ckey:
                return s

    return None


def parse_dates(series: pd.Series) -> pd.Series:
    # Excel serial / 문자열 / datetime 모두 처리.
    s = series.copy()

    numeric = pd.to_numeric(s, errors="coerce")
    serial = numeric.between(20000, 80000)

    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")

    if serial.any():
        out.loc[serial] = (
            pd.Timestamp("1899-12-30")
            + pd.to_timedelta(numeric.loc[serial], unit="D")
        )

    other = ~serial
    if other.any():
        out.loc[other] = pd.to_datetime(
            s.loc[other],
            errors="coerce",
        )

    return out


def clean_matrix(
    df: pd.DataFrame,
    ticker_order: list[str] | None = None,
    numeric: bool = True,
) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    df = df.copy()

    date_col = next(
        (
            c for c in df.columns
            if str(c).strip().lower() in {"날짜", "date"}
        ),
        df.columns[0],
    )

    rename = {}
    for c in df.columns:
        if c == date_col:
            continue
        rename[c] = normalize_ticker(c)

    df = df.rename(columns=rename)
    df[date_col] = parse_dates(df[date_col])

    df = (
        df.dropna(subset=[date_col])
        .drop_duplicates(subset=[date_col], keep="last")
        .sort_values(date_col)
        .set_index(date_col)
    )
    df.index.name = "날짜"

    # 빈/중복 column 제거
    df = df.loc[:, [c for c in df.columns if c]]
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated(keep="last")]

    if numeric:
        df = df.apply(pd.to_numeric, errors="coerce")

    if ticker_order is not None:
        df = df.reindex(columns=ticker_order)

    return df


def to_excel_ready(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out = out.reset_index()

    if out.columns[0] != "날짜":
        out = out.rename(columns={out.columns[0]: "날짜"})

    return out


# =============================================================================
# 2. LOAD BASE WORKBOOK
# =============================================================================

def load_base():
    data_path = locate_file(DATA_INPUT_CANDIDATES, "data")
    qc_path = locate_file(QC_INPUT_CANDIDATES, "qc")

    print("=" * 96)
    print("INPUT")
    print("=" * 96)
    print("데이터:", data_path.name)
    print("QC    :", qc_path.name)

    data_xls = pd.ExcelFile(data_path, engine="openpyxl")
    qc_xls = pd.ExcelFile(qc_path, engine="openpyxl")

    price_sheet = find_sheet_name(
        data_xls,
        ["수정주가", "수정주가full", "adjustedclose"],
    )
    volume_sheet = find_sheet_name(
        data_xls,
        ["거래량", "volume"],
    )
    member_sheet = find_sheet_name(
        data_xls,
        ["KOSPI200편입여부", "편입여부", "membership"],
    )

    if price_sheet is None:
        raise ValueError(
            f"{data_path.name}에서 수정주가 시트를 찾지 못했습니다.\n"
            f"실제 시트={data_xls.sheet_names}"
        )

    # 407개 universe는 기존 수정주가 header 자체에서 가져온다.
    header = pd.read_excel(
        data_path,
        sheet_name=price_sheet,
        nrows=0,
        engine="openpyxl",
    )

    date_col = next(
        (
            c for c in header.columns
            if str(c).strip().lower() in {"날짜", "date"}
        ),
        header.columns[0],
    )

    tickers = [
        normalize_ticker(c)
        for c in header.columns
        if c != date_col and normalize_ticker(c)
    ]

    # 순서 보존 + 중복 제거
    tickers = list(dict.fromkeys(tickers))

    if len(tickers) != EXPECTED_TICKERS:
        raise RuntimeError(
            f"역사적 ticker 수가 {len(tickers)}개입니다. "
            f"예상 {EXPECTED_TICKERS}개와 다릅니다.\n"
            "잘못된 입력 파일을 선택했을 가능성이 있으므로 중단합니다."
        )

    print(f"역사적 ticker: {len(tickers):,}개")

    price_raw = pd.read_excel(
        data_path,
        sheet_name=price_sheet,
        engine="openpyxl",
        dtype=object,
    )
    old_close = clean_matrix(price_raw, tickers, numeric=True)

    if volume_sheet is not None:
        volume_raw = pd.read_excel(
            data_path,
            sheet_name=volume_sheet,
            engine="openpyxl",
            dtype=object,
        )
        old_volume = clean_matrix(volume_raw, tickers, numeric=True)
    else:
        old_volume = pd.DataFrame(
            index=old_close.index,
            columns=tickers,
            dtype=float,
        )
        old_volume.index.name = "날짜"

    if member_sheet is not None:
        member_raw = pd.read_excel(
            data_path,
            sheet_name=member_sheet,
            engine="openpyxl",
            dtype=object,
        )
        membership = clean_matrix(member_raw, tickers, numeric=True)
        membership = membership.fillna(0)
        membership = membership.eq(1).astype("int8")
    else:
        membership = pd.DataFrame(
            0,
            index=old_close.index,
            columns=tickers,
            dtype="int8",
        )
        membership.index.name = "날짜"

    research_dates = old_close.index.sort_values()

    print(
        f"연구 calendar: {research_dates.min().date()} ~ "
        f"{research_dates.max().date()} | {len(research_dates):,}일"
    )

    return {
        "data_path": data_path,
        "qc_path": qc_path,
        "data_xls": data_xls,
        "qc_xls": qc_xls,
        "price_sheet": price_sheet,
        "volume_sheet": volume_sheet,
        "member_sheet": member_sheet,
        "tickers": tickers,
        "research_dates": research_dates,
        "old_close": old_close.reindex(research_dates),
        "old_volume": old_volume.reindex(research_dates),
        "membership": membership.reindex(research_dates).fillna(0).astype("int8"),
    }


# =============================================================================
# 3. NAVER DOWNLOAD
# =============================================================================

def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/151.0 Safari/537.36"
            ),
            "Referer": "https://finance.naver.com/",
            "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
        }
    )
    return s


def parse_naver_response(text: str) -> pd.DataFrame:
    if text is None:
        return pd.DataFrame()

    cleaned = html.unescape(text)
    cleaned = re.sub(
        r"<[^>]+>",
        "",
        cleaned,
        flags=re.MULTILINE,
    ).strip()

    if not cleaned:
        return pd.DataFrame()

    parsed_rows = []

    # 1) literal_eval
    try:
        parsed = ast.literal_eval(cleaned)

        if isinstance(parsed, list):
            rows = parsed[1:] if len(parsed) >= 2 else []

            for row in rows:
                if not isinstance(row, (list, tuple)) or len(row) < 6:
                    continue

                date_raw = (
                    str(row[0])
                    .strip()
                    .replace('"', "")
                    .replace("'", "")
                )

                if not re.fullmatch(r"\d{8}", date_raw):
                    continue

                parsed_rows.append(
                    {
                        "date": date_raw,
                        "open": row[1],
                        "high": row[2],
                        "low": row[3],
                        "close": row[4],
                        "volume": row[5],
                    }
                )
    except Exception:
        pass

    # 2) regex fallback
    if not parsed_rows:
        pattern = re.compile(
            r'\[\s*["\']?(\d{8})["\']?\s*,'
            r'\s*([-+]?\d+(?:\.\d+)?)\s*,'
            r'\s*([-+]?\d+(?:\.\d+)?)\s*,'
            r'\s*([-+]?\d+(?:\.\d+)?)\s*,'
            r'\s*([-+]?\d+(?:\.\d+)?)\s*,'
            r'\s*([-+]?\d+(?:\.\d+)?)'
        )

        for m in pattern.finditer(cleaned):
            parsed_rows.append(
                {
                    "date": m.group(1),
                    "open": m.group(2),
                    "high": m.group(3),
                    "low": m.group(4),
                    "close": m.group(5),
                    "volume": m.group(6),
                }
            )

    if not parsed_rows:
        return pd.DataFrame()

    df = pd.DataFrame(parsed_rows)

    df["date"] = pd.to_datetime(
        df["date"],
        format="%Y%m%d",
        errors="coerce",
    )

    for c in ["open", "high", "low", "close", "volume"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.dropna(subset=["date", "close"])
    df = df.loc[df["close"] > 0]

    return (
        df.sort_values("date")
        .drop_duplicates(subset=["date"], keep="last")
        .reset_index(drop=True)
    )


def request_range(
    session: requests.Session,
    ticker: str,
    start: pd.Timestamp,
    end: pd.Timestamp,
) -> tuple[pd.DataFrame, str]:
    params = {
        "symbol": ticker,
        "requestType": 1,
        "startTime": pd.Timestamp(start).strftime("%Y%m%d"),
        "endTime": pd.Timestamp(end).strftime("%Y%m%d"),
        "timeframe": "day",
    }

    last_error = ""

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = session.get(
                NAVER_URL,
                params=params,
                timeout=REQUEST_TIMEOUT,
            )
            r.raise_for_status()

            df = parse_naver_response(r.text)

            if not df.empty:
                return df, "OK"

            last_error = "EMPTY_OR_PARSE_FAIL"

        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"

        if attempt < MAX_RETRIES:
            time.sleep(REQUEST_SLEEP_SECONDS * attempt)

    return pd.DataFrame(), last_error


def chunk_ranges(
    start: pd.Timestamp,
    end: pd.Timestamp,
    years: int,
):
    cur = pd.Timestamp(start)

    while cur <= end:
        nxt = min(
            cur + pd.DateOffset(years=years) - pd.Timedelta(days=1),
            end,
        )
        yield cur, nxt
        cur = nxt + pd.Timedelta(days=1)


def download_full_ticker(
    session: requests.Session,
    ticker: str,
    request_start: pd.Timestamp,
    request_end: pd.Timestamp,
) -> tuple[pd.DataFrame, str]:
    # 먼저 kp.py와 같은 방식으로 전체 범위를 한 번에 요청.
    df, status = request_range(
        session,
        ticker,
        request_start,
        request_end,
    )

    if not df.empty:
        return df, "OK_FULL_RANGE"

    # 전체 요청 실패 시에만 5년 단위 fallback.
    parts = []
    errors = [status]

    for s, e in chunk_ranges(
        request_start,
        request_end,
        FALLBACK_CHUNK_YEARS,
    ):
        part, part_status = request_range(
            session,
            ticker,
            s,
            e,
        )

        if not part.empty:
            parts.append(part)
        else:
            errors.append(
                f"{s.date()}~{e.date()}:{part_status}"
            )

        time.sleep(REQUEST_SLEEP_SECONDS)

    if parts:
        out = (
            pd.concat(parts, ignore_index=True)
            .sort_values("date")
            .drop_duplicates("date", keep="last")
            .reset_index(drop=True)
        )
        return out, "OK_CHUNK_FALLBACK"

    return pd.DataFrame(), " | ".join(errors[-5:])


def read_legacy_raw(ticker: str) -> pd.DataFrame:
    p = LEGACY_RAW_DIR / f"{ticker}.csv"

    if not p.exists():
        return pd.DataFrame()

    try:
        df = pd.read_csv(p, low_memory=False)

        date_col = next(
            (c for c in df.columns if str(c).lower() == "date"),
            None,
        )
        close_col = next(
            (c for c in df.columns if str(c).lower() == "close"),
            None,
        )

        if date_col is None or close_col is None:
            return pd.DataFrame()

        out = pd.DataFrame(
            {
                "date": pd.to_datetime(df[date_col], errors="coerce"),
                "close": pd.to_numeric(df[close_col], errors="coerce"),
                "volume": pd.to_numeric(
                    df["volume"] if "volume" in df.columns else np.nan,
                    errors="coerce",
                ),
            }
        )

        out = out.dropna(subset=["date", "close"])
        out = out.loc[out["close"] > 0]

        return (
            out.sort_values("date")
            .drop_duplicates("date", keep="last")
            .reset_index(drop=True)
        )

    except Exception:
        return pd.DataFrame()


# =============================================================================
# 4. BUILD FULL-HISTORY PANELS
# =============================================================================

def build_panels(base: dict):
    tickers = base["tickers"]
    dates = base["research_dates"]
    old_close = base["old_close"]
    old_volume = base["old_volume"]

    # 기존 값은 fallback으로 시작.
    final_close = old_close.copy()
    final_volume = old_volume.copy()

    # 첫 연구일 로그수익률용 buffer.
    request_start = dates.min() - pd.Timedelta(
        days=PRE_SAMPLE_BUFFER_DAYS
    )
    request_end = dates.max()

    session = make_session()

    status_rows = []
    full_return_series = {}

    print()
    print("=" * 96)
    print("407종목 FULL HISTORY 수집")
    print("가격은 membership으로 마스킹하지 않습니다.")
    print(
        f"요청 범위: {request_start.date()} ~ "
        f"{request_end.date()}"
    )
    print("=" * 96)

    for i, ticker in enumerate(tickers, start=1):
        print(
            f"[{i:>3}/{len(tickers)}] {ticker}",
            end=" -> ",
            flush=True,
        )

        legacy = read_legacy_raw(ticker)

        downloaded, dl_status = download_full_ticker(
            session=session,
            ticker=ticker,
            request_start=request_start,
            request_end=request_end,
        )

        # legacy + 새 download 병합
        pieces = []

        if not legacy.empty:
            pieces.append(legacy[["date", "close", "volume"]])

        if not downloaded.empty:
            pieces.append(downloaded[["date", "close", "volume"]])

        if pieces:
            hist = (
                pd.concat(pieces, ignore_index=True)
                .sort_values("date")
                .drop_duplicates("date", keep="last")
                .reset_index(drop=True)
            )
        else:
            hist = pd.DataFrame(
                columns=["date", "close", "volume"]
            )

        old_nonnull = int(
            old_close[ticker].notna().sum()
        )

        if not hist.empty:
            h = hist.set_index("date").sort_index()

            # 연구 calendar에 맞추되 membership은 전혀 보지 않는다.
            new_close = h["close"].reindex(dates)
            new_volume = h["volume"].reindex(dates)

            # 네이버/legacy에 값이 있는 셀은 우선 적용.
            final_close.loc[
                new_close.notna(), ticker
            ] = new_close.loc[new_close.notna()]

            final_volume.loc[
                new_volume.notna(), ticker
            ] = new_volume.loc[new_volume.notna()]

            # log return은 buffer까지 포함된 원시 history에서 계산.
            r = np.log(
                h["close"] / h["close"].shift(1)
            )
            full_return_series[ticker] = r.reindex(dates)

            first_raw = h["close"].dropna().index.min()
            last_raw = h["close"].dropna().index.max()

            status = (
                dl_status
                if not downloaded.empty
                else "LEGACY_ONLY"
            )
        else:
            first_raw = pd.NaT
            last_raw = pd.NaT
            status = "ORIGINAL_ONLY_DOWNLOAD_FAIL"

        final_nonnull = int(
            final_close[ticker].notna().sum()
        )

        added = int(
            (
                final_close[ticker].notna()
                & old_close[ticker].isna()
            ).sum()
        )

        print(
            f"{status} | 기존={old_nonnull:,} "
            f"최종={final_nonnull:,} 추가={added:,}"
        )

        status_rows.append(
            {
                "ticker": ticker,
                "status": status,
                "download_detail": dl_status,
                "raw_first_date": first_raw,
                "raw_last_date": last_raw,
                "old_price_obs": old_nonnull,
                "final_price_obs": final_nonnull,
                "added_price_cells": added,
            }
        )

        time.sleep(REQUEST_SLEEP_SECONDS)

    # 다운로드 history가 있으면 buffer 기반 return,
    # 없으면 최종 panel에서 계산한 return으로 fallback.
    panel_log_return = np.log(
        final_close / final_close.shift(1)
    )

    final_log_return = panel_log_return.copy()

    for ticker, s in full_return_series.items():
        mask = s.notna()
        final_log_return.loc[mask, ticker] = s.loc[mask]

    final_close.index.name = "날짜"
    final_volume.index.name = "날짜"
    final_log_return.index.name = "날짜"

    status_df = pd.DataFrame(status_rows)

    return (
        final_close,
        final_volume,
        final_log_return,
        status_df,
    )


# =============================================================================
# 5. QC
# =============================================================================

def build_full_history_qc(
    base: dict,
    final_close: pd.DataFrame,
    final_volume: pd.DataFrame,
    status_df: pd.DataFrame,
):
    dates = base["research_dates"]
    tickers = base["tickers"]
    old_close = base["old_close"]
    membership = base["membership"]

    coverage_rows = []

    for ticker in tickers:
        price = final_close[ticker]
        old = old_close[ticker]
        member = membership[ticker].eq(1)

        valid = price.notna()

        member_dates = dates[member.values]

        first_member = (
            member_dates.min()
            if len(member_dates) else pd.NaT
        )
        last_member = (
            member_dates.max()
            if len(member_dates) else pd.NaT
        )

        before_first = (
            valid
            & (dates < first_member)
            if pd.notna(first_member)
            else pd.Series(False, index=dates)
        )

        after_last = (
            valid
            & (dates > last_member)
            if pd.notna(last_member)
            else pd.Series(False, index=dates)
        )

        nonmember_price = valid & (~member)

        coverage_rows.append(
            {
                "ticker": ticker,
                "first_price_date": (
                    price.index[valid].min()
                    if valid.any()
                    else pd.NaT
                ),
                "last_price_date": (
                    price.index[valid].max()
                    if valid.any()
                    else pd.NaT
                ),
                "old_price_obs": int(old.notna().sum()),
                "final_price_obs": int(valid.sum()),
                "added_price_cells": int(
                    (valid & old.isna()).sum()
                ),
                "first_member_date": first_member,
                "last_member_date": last_member,
                "membership_days": int(member.sum()),
                "price_while_member_obs": int(
                    (valid & member).sum()
                ),
                "missing_price_while_member": int(
                    ((~valid) & member).sum()
                ),
                "price_while_nonmember_obs": int(
                    nonmember_price.sum()
                ),
                "price_before_first_entry_obs": int(
                    before_first.sum()
                ),
                "price_after_last_exit_obs": int(
                    after_last.sum()
                ),
                "has_pre_entry_history": int(
                    before_first.any()
                ),
                "has_post_exit_history": int(
                    after_last.any()
                ),
            }
        )

    coverage = pd.DataFrame(coverage_rows)

    daily = pd.DataFrame(
        {
            "날짜": dates,
            "407종목_가격보유수": final_close.notna().sum(axis=1).to_numpy(),
            "KOSPI200_구성종목수": membership.sum(axis=1).to_numpy(),
            "편입종목_가격보유수": (
                final_close.notna()
                & membership.eq(1)
            ).sum(axis=1).to_numpy(),
            "비편입종목_가격보유수": (
                final_close.notna()
                & membership.eq(0)
            ).sum(axis=1).to_numpy(),
            "407종목_거래량보유수": final_volume.notna().sum(axis=1).to_numpy(),
        }
    )

    summary = pd.DataFrame(
        [
            ["연구시작일", dates.min()],
            ["연구종료일", dates.max()],
            ["거래일수", len(dates)],
            ["역사적종목수", len(tickers)],
            [
                "기존수정주가_비결측셀",
                int(old_close.notna().sum().sum()),
            ],
            [
                "FULL_HISTORY_수정주가_비결측셀",
                int(final_close.notna().sum().sum()),
            ],
            [
                "추가된_수정주가셀",
                int(
                    (
                        final_close.notna()
                        & old_close.isna()
                    ).sum().sum()
                ),
            ],
            [
                "비편입기간_가격관측셀",
                int(
                    (
                        final_close.notna()
                        & membership.eq(0)
                    ).sum().sum()
                ),
            ],
            [
                "편입중_가격누락셀",
                int(
                    (
                        final_close.isna()
                        & membership.eq(1)
                    ).sum().sum()
                ),
            ],
            [
                "편입전가격history_보유종목수",
                int(
                    coverage["has_pre_entry_history"].sum()
                ),
            ],
            [
                "최종편출후가격history_보유종목수",
                int(
                    coverage["has_post_exit_history"].sum()
                ),
            ],
            [
                "다운로드성공/부분성공종목수",
                int(
                    status_df["status"]
                    .astype(str)
                    .str.contains("OK|LEGACY", regex=True)
                    .sum()
                ),
            ],
            [
                "원본만사용_다운로드실패종목수",
                int(
                    status_df["status"]
                    .eq("ORIGINAL_ONLY_DOWNLOAD_FAIL")
                    .sum()
                ),
            ],
        ],
        columns=["항목", "값"],
    )

    guide = pd.DataFrame(
        [
            [
                "핵심",
                "수정주가/거래량은 KOSPI200 편입 여부로 마스킹하지 않았습니다.",
            ],
            [
                "Full History",
                "407개 역사적 종목 각각에 대해 2008~2026 연구 calendar에서 "
                "네이버에 가격이 존재하는 날짜를 모두 보존했습니다.",
            ],
            [
                "편입 전/후",
                "편입 전 또는 편출 후에도 상장되어 가격이 있으면 수정주가 시트에 값이 남습니다.",
            ],
            [
                "Membership",
                "KOSPI200편입여부는 투자/PCA 유니버스 판정용이며 가격 수집 기준이 아닙니다.",
            ],
            [
                "결측",
                "상장 전, 상장폐지 후, 원자료 미제공 등 실제 가격이 없는 구간은 빈 셀입니다.",
            ],
            [
                "수익률",
                "일간로그수익률은 같은 종목의 연속 관측 수정주가로 계산했습니다.",
            ],
            [
                "QC",
                "FullHistory_Coverage에서 종목별 편입 전/후 가격 관측 수와 추가된 셀 수를 확인합니다.",
            ],
        ],
        columns=["구분", "내용"],
    )

    return summary, guide, coverage, daily


# =============================================================================
# 6. EXCEL OUTPUT
# =============================================================================

def write_data_workbook(
    base: dict,
    final_close: pd.DataFrame,
    final_volume: pd.DataFrame,
    final_log_return: pd.DataFrame,
):
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    data_path = base["data_path"]
    data_xls = base["data_xls"]
    membership = base["membership"]

    core_aliases = {
        base["price_sheet"],
        base["volume_sheet"],
        base["member_sheet"],
        "수정주가",
        "거래량",
        "일간로그수익률",
        "KOSPI200편입여부",
        "편입여부",
    }
    core_aliases.discard(None)

    guide = pd.DataFrame(
        [
            ["자료범위", f"{final_close.index.min():%Y-%m-%d} ~ {final_close.index.max():%Y-%m-%d}"],
            ["거래일수", len(final_close)],
            ["역사적종목수", final_close.shape[1]],
            ["가격기준", "407개 종목의 편입 여부와 무관한 Full History"],
            ["수정주가", "네이버 fchart 기준. 기존 값은 fallback으로 보존"],
            ["편입여부", "가격 마스킹에 사용하지 않음. PCA/투자 유니버스 판정에만 사용"],
            ["결측", "실제 가격 미존재/원자료 미제공 구간은 빈 셀"],
        ],
        columns=["구분", "내용"],
    )

    print()
    print("[1/2] 분석데이터 Excel 저장:", OUT_DATA.name)

    with pd.ExcelWriter(
        OUT_DATA,
        engine="openpyxl",
    ) as writer:
        guide.to_excel(
            writer,
            sheet_name="파일안내",
            index=False,
        )

        to_excel_ready(final_close).to_excel(
            writer,
            sheet_name="수정주가",
            index=False,
        )

        to_excel_ready(final_volume).to_excel(
            writer,
            sheet_name="거래량",
            index=False,
        )

        to_excel_ready(final_log_return).to_excel(
            writer,
            sheet_name="일간로그수익률",
            index=False,
        )

        to_excel_ready(membership).to_excel(
            writer,
            sheet_name="KOSPI200편입여부",
            index=False,
        )

        # 기존 데이터 workbook의 다른 참고 시트는 그대로 복사.
        for sheet in data_xls.sheet_names:
            if sheet in core_aliases or sheet == "파일안내":
                continue

            try:
                df = pd.read_excel(
                    data_path,
                    sheet_name=sheet,
                    engine="openpyxl",
                    dtype=object,
                )

                safe = str(sheet)[:31]

                if safe in writer.book.sheetnames:
                    continue

                df.to_excel(
                    writer,
                    sheet_name=safe,
                    index=False,
                )
            except Exception as exc:
                print(
                    f"  [참고시트 SKIP] {sheet}: "
                    f"{type(exc).__name__}"
                )

        # 최소한의 사용성
        for ws in writer.book.worksheets:
            ws.freeze_panes = "B2" if ws.max_column > 3 else "A2"
            if ws.max_row >= 1 and ws.max_column >= 1:
                ws.auto_filter.ref = ws.dimensions


def write_qc_workbook(
    base: dict,
    summary: pd.DataFrame,
    guide: pd.DataFrame,
    status_df: pd.DataFrame,
    coverage: pd.DataFrame,
    daily: pd.DataFrame,
):
    qc_path = base["qc_path"]
    qc_xls = base["qc_xls"]

    print("[2/2] QC Excel 저장:", OUT_QC.name)

    with pd.ExcelWriter(
        OUT_QC,
        engine="openpyxl",
    ) as writer:
        guide.to_excel(
            writer,
            sheet_name="파일안내",
            index=False,
        )

        summary.to_excel(
            writer,
            sheet_name="FullHistory_요약",
            index=False,
        )

        status_df.to_excel(
            writer,
            sheet_name="다운로드상태",
            index=False,
        )

        coverage.to_excel(
            writer,
            sheet_name="FullHistory_Coverage",
            index=False,
        )

        daily.to_excel(
            writer,
            sheet_name="FullHistory_일별",
            index=False,
        )

        # 기존 QC 시트도 보존.
        for sheet in qc_xls.sheet_names:
            if sheet == "파일안내":
                continue

            try:
                df = pd.read_excel(
                    qc_path,
                    sheet_name=sheet,
                    engine="openpyxl",
                    dtype=object,
                )

                safe = str(sheet)[:31]

                if safe in writer.book.sheetnames:
                    safe = ("기존_" + safe)[:31]

                df.to_excel(
                    writer,
                    sheet_name=safe,
                    index=False,
                )
            except Exception as exc:
                print(
                    f"  [기존 QC시트 SKIP] {sheet}: "
                    f"{type(exc).__name__}"
                )

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            if ws.max_row >= 1 and ws.max_column >= 1:
                ws.auto_filter.ref = ws.dimensions


# =============================================================================
# 7. MAIN
# =============================================================================

def main():
    print()
    print("=" * 96)
    print("KOSPI200 407종목 FULL HISTORY BUILDER")
    print("=" * 96)
    print(
        "원칙: 가격은 membership과 독립적으로 수집. "
        "membership=0이어도 가격이 있으면 보존."
    )
    print()

    base = load_base()

    (
        final_close,
        final_volume,
        final_log_return,
        status_df,
    ) = build_panels(base)

    (
        summary,
        guide,
        coverage,
        daily,
    ) = build_full_history_qc(
        base=base,
        final_close=final_close,
        final_volume=final_volume,
        status_df=status_df,
    )

    write_data_workbook(
        base=base,
        final_close=final_close,
        final_volume=final_volume,
        final_log_return=final_log_return,
    )

    write_qc_workbook(
        base=base,
        summary=summary,
        guide=guide,
        status_df=status_df,
        coverage=coverage,
        daily=daily,
    )

    print()
    print("=" * 96)
    print("완료")
    print("=" * 96)
    print("최종 결과는 아래 2개뿐입니다.")
    print("1.", OUT_DATA)
    print("2.", OUT_QC)
    print()
    print(
        "추가된 수정주가 셀:",
        int(
            (
                final_close.notna()
                & base["old_close"].isna()
            ).sum().sum()
        ),
    )
    print(
        "비편입기간인데 가격이 존재하는 셀:",
        int(
            (
                final_close.notna()
                & base["membership"].eq(0)
            ).sum().sum()
        ),
    )
    print(
        "=> 이 값이 0이 아니면 membership 밖 Full History가 "
        "실제로 포함된 것입니다."
    )


if __name__ == "__main__":
    main()

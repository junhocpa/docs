#!/usr/bin/env python3
"""YMK 매입 전표 검토용 로더 헬퍼.

더존/구매시스템 export 공통 이슈(openpyxl 스타일 오류, 선행 0 소실, 합계행)를
처리한 로더와 미발행 추출·전표키 유틸. 분석 로직은 상황마다 다르므로 여기서는
'읽기'까지만 표준화한다.

사용 예:
    from load_helpers import load_tax_list, load_journal, extract_unissued, norm_pj
    tax = load_tax_list("acb2050_taxData.xlsx")
    mi  = extract_unissued(tax)                # 매입 & 미발행
    j   = load_journal("ACA0050Grid.xlsx")     # 전표키 컬럼 포함
"""
import pandas as pd

STR_COLS = ["거래처코드", "사업자(주민)번호", "사업자번호", "계정과목코드"]


PAD = {"거래처코드": 5, "계정과목코드": 7}  # 선행 0 복원 자릿수


def _read(path, header=0):
    """calamine 엔진 우선(더존 export는 openpyxl 실패 잦음), 문자열 컬럼 보존."""
    try:
        df = pd.read_excel(path, engine="calamine", header=header)
    except Exception:
        df = pd.read_excel(path, header=header)
    for c in STR_COLS:
        if c in df.columns:
            s = df[c].astype(str).str.replace(r"\.0$", "", regex=True)
            if c in PAD:  # 숫자로 읽혀 선행 0이 소실된 코드 복원 (예: 4192→04192)
                s = s.where(~s.str.fullmatch(r"\d+"), s.str.zfill(PAD[c]))
            df[c] = s
    return df


def load_tax_list(path):
    """ACB2050 세금계산서 리스트."""
    return _read(path)


def load_journal(path):
    """ACA0050 분개장. 합계행 제거 + 전표키(작성일-작성번호) 부여."""
    j = _read(path)
    j = j[j["승인일"].astype(str) != "합계"].copy()
    j["전표키"] = j["작성일"].astype(str) + "-" + j["작성번호"].astype(str)
    return j


def load_purchase(path):
    """자재진행리스트/입고등록마스타/입고현황 공통: 제목행 스킵(header=1),
    순번 없는 행 제거, PJ 정규화, TAX발행일 date 문자열화."""
    df = _read(path, header=1)
    if "순번" in df.columns:
        df = df[df["순번"].notna()].copy()
    if "PJT코드" in df.columns:
        df["PJ"] = df["PJT코드"].map(norm_pj)
    if "TAX발행일" in df.columns:
        df["TAX"] = pd.to_datetime(df["TAX발행일"], errors="coerce").dt.date.astype(str)
    return df


def norm_pj(x):
    """구매시스템 export의 선행 0 소실 복원: 5자리→앞에 0, '71'·6자리는 그대로."""
    s = str(x).strip().split(".")[0]
    if s in ("nan", "None", ""):
        return ""
    return "0" + s if len(s) == 5 else s


def extract_unissued(tax, gubun="매입"):
    """전표 미발행 건 추출."""
    return tax[(tax["구분"] == gubun) & (tax["전표여부"] == "미발행")].copy()


def offset_pairs(mi):
    """같은 거래처+작성일 그룹의 부호 포함 합이 0이면 상쇄쌍 후보로 태깅해 반환."""
    g = mi.groupby(["거래처코드", "작성일"])["공급가액"].transform("sum")
    grp_n = mi.groupby(["거래처코드", "작성일"])["공급가액"].transform("size")
    mi = mi.copy()
    mi["상쇄쌍후보"] = (g == 0) & (grp_n > 1)
    return mi


def vendor_vouchers(j, code, buy_only=True, last=3):
    """분개장에서 거래처의 최근 전표(전표키 그룹) 반환. buy_only면 매입유형만."""
    sub = j[j["거래처코드"] == str(code)]
    if buy_only:
        sub = sub[sub["전표유형명"].astype(str).str.contains("매")]
    keys = sorted(sub["전표키"].unique())[-last:]
    return {k: j[j["전표키"] == k].sort_values("전표라인번호") for k in keys}


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)
    tax = load_tax_list(sys.argv[1])
    mi = offset_pairs(extract_unissued(tax))
    cols = ["No", "분류", "작성일", "거래처코드", "거래처명", "공급가액", "세액", "상쇄쌍후보"]
    print(mi[[c for c in cols if c in mi.columns]].to_string(index=False))
    print(f"\n미발행 {len(mi)}건 | 공급가액 합 {mi['공급가액'].sum():,.0f} | 상쇄쌍후보 {int(mi['상쇄쌍후보'].sum())}건")

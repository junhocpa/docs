#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
YMK 재고수불 자동전표 워크북 생성 템플릿 (지침 v1.3)

사용법:
  1) 재고현황(출고현황)·분개장을 로드해 규칙 1~11을 적용하고, 아래 ROWS/COPY_REF/
     VENDORS/PJTS/ACCTS 를 이번 달 데이터로 채운다. (규칙 적용은 references/rules.md 참조)
  2) python build_voucher.py 실행 → outputs 폴더에 워크북 생성.
  3) recalc.py 로 재계산 후, 선행 0·차대검증을 재확인한다.

핵심: 코드성 값(작성일자/회계단위/거래처/PJT)은 문자열로 넣고 텍스트 서식('@')을 적용해
선행 0을 보존한다. T~V(전표복사 참조)는 더존 업로드 대상이 아니다(A~S만 업로드).
타계정구분(규칙 11)은 업로드 양식에 없다 — 입력 후 더존에서 지정하고 원장으로 검증한다.
"""
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

# ============== 이번 달 파라미터 (매월 수정) ==============
YEAR, MONTH = 2026, 5
VERSION = "v1.3"
# 출력 경로: 환경변수 > Claude.ai outputs 폴더 > 현재 작업폴더 순으로 자동 선택
OUTPUT_DIR = os.environ.get("YMK_OUTPUT_DIR") or (
    "/mnt/user-data/outputs" if os.path.isdir("/mnt/user-data/outputs") else "."
)
OUT = os.path.join(OUTPUT_DIR, f"YMK_재고수불_자동전표_{YEAR}년{MONTH:02d}월_{VERSION}.xlsx")

# 자동전표 라인: (작성일 YYYYMMDD, 전표번호, 라인순번, 계정코드, 차대(3/4), 적요, 금액, 거래처5자리, 부서 or None, PJT6자리 or None, 품의)
# 타프로젝트대체는 양쪽 차변(3)에 +/- 금액, 적요 선례형식(규칙 8):
#   (+) "타프로젝트에서 대체액" 새PJT / (-) "타프로젝트로 대체액" 원PJT, 품의 "타프로젝트대체"
# 연구개발비대체(규칙 10): 차변 경상연구개발비(제조 523)/PJT "000093", 대변 원재료.
ROWS = [
    ("20260515", 1, 1, 1460000, 3, "타계정에서 대체액", 77873, "03340", None,   None,     "타계정대체"),
    ("20260515", 1, 2, 1490000, 4, "타계정으로 대체액", 77873, "03340", "2000", "000001", "타계정대체"),
    ("20260526", 1, 1, 1490000, 3, "타계정에서 대체액", 98167, "00321", "2000", "021101", "타계정대체"),
    ("20260526", 1, 2, 1490000, 4, "타계정으로 대체액", 98167, "00321", "2000", "000001", "타계정대체"),
]
# 전표복사 참조(규칙 9): 작성일 -> (참조일, 참조번호, 유형). 전표 첫 라인에만 표기.
COPY_REF = {
    "20260515": ("2026-04-22", 12, "상품대체"),
    "20260526": ("2026-04-17", 5,  "원재료 타계정대체"),
}
# ③거래처매핑: (인벤코드, 거래처명, 회계코드, 사업번호, 상태, 건수, 금액합, 특이사항)
VENDORS = [
    ("51165", "에이엔티",        "03340", "1221977121", "매칭됨", 2, 77873, "정상 (분개장 일치)"),
    ("31311", "요도가와메덱(일본)", "00321", "1111111111", "매칭됨", 1, 98167, "더미 사업번호/해외 — 통화 확인 권장"),
]
# ⑥월별계정변동: (계정명, 코드, 건수, 금액)  — 출고후/차변 기준
ACCTS = [("상품", "1460000", 2, 77873), ("원재료", "1490000", 1, 98167)]

# ============== 서식 상수 ==============
FONT = "맑은 고딕"
HEAD_FILL = PatternFill("solid", fgColor="305496"); HEAD_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
SUB_FILL = PatternFill("solid", fgColor="D9E1F2");  SUB_FONT = Font(name=FONT, color="808080", size=9)
REF_FILL = PatternFill("solid", fgColor="EDEDED")
TITLE_FONT = Font(name=FONT, bold=True, size=14, color="1F4E79")
TOT_FILL = PatternFill("solid", fgColor="FFE699")
DR_FILL = PatternFill("solid", fgColor="DDEBF7"); CR_FILL = PatternFill("solid", fgColor="FCE4D6")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
base = Font(name=FONT, size=10)
thin = Side(style="thin", color="BFBFBF"); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
TEXT_COLS = ["A", "B", "C", "D", "L", "N", "Q"]  # 선행 0 보존용 텍스트 서식 대상

def header2(ws, kor, eng, r=1):
    for c, (k, e) in enumerate(zip(kor, eng), 1):
        a = ws.cell(r, c, k); a.fill = HEAD_FILL; a.font = HEAD_FONT
        a.alignment = Alignment("center", "center", wrap_text=True); a.border = BORDER
        b = ws.cell(r + 1, c, e); b.fill = SUB_FILL; b.font = SUB_FONT
        b.alignment = Alignment("center", "center"); b.border = BORDER

def build():
    wb = Workbook()

    # ① 요약
    ws = wb.active; ws.title = "①요약"
    ws["A1"] = f"YMK 재고수불 자동전표 처리 결과 — {YEAR}년 {MONTH}월"; ws["A1"].font = TITLE_FONT
    dr = sum(r[6] for r in ROWS if r[4] == 3); cr = sum(r[6] for r in ROWS if r[4] == 4)
    summ = [("항목", "결과"),
            ("의미있는 변동건수", f"통합 {len(set((r[0],r[1]) for r in ROWS))}전표 {len(ROWS)}라인"),
            ("변동 금액 합계", f"{dr:,}원"),
            ("차/대변 검증", f"차변 {dr:,} / 대변 {cr:,} / 차이 {dr-cr}"),
            ("회계단위(B컬럼)", "1000 (★ YMK 회계단위 확인 후 수정)"),
            ("전표복사 참조", " / ".join(f"{v[2]}→{v[0]} #{v[1]}" for v in COPY_REF.values()))]
    for i, (a, b) in enumerate(summ):
        ra = ws.cell(3 + i, 1, a); rb = ws.cell(3 + i, 2, b); ra.border = BORDER; rb.border = BORDER
        ra.font = HEAD_FONT if i == 0 else base; rb.font = HEAD_FONT if i == 0 else base
        if i == 0: ra.fill = HEAD_FILL; rb.fill = HEAD_FILL
    ws.column_dimensions["A"].width = 24; ws.column_dimensions["B"].width = 52

    # ② 자동전표업로드
    ws = wb.create_sheet("②자동전표업로드")
    kor = ["작성일자","회계단위","사업장","프로젝트/부서코드","발행일","작성번호","라인순번","계정과목","차대구분","적요","계정금액","거래처","사용부서","프로젝트","품의내역","전표유형","수출신고번호 등","환종","외화금액"]
    eng = ["MENU_DT","IN_DIV_CD","M_DIV_CD","MGT_CD","ISSUE_DT","MENU_SQ","MENU_LN_SQ","ACCT_CD","DRCR_FG","RMK_DC","ACCT_AM","TR_CD","CT_DEPT","PJT_CD","ISU_DOC","DOCU_TY","CT_NB","DUMMY1","CASH_AM"]
    header2(ws, kor, eng)
    for col, txt in (("T", "전표복사 참조일"), ("U", "참조 전표번호"), ("V", "참조유형")):
        h = ws[col + "1"]; h.value = txt; h.fill = REF_FILL; h.font = Font(name=FONT, bold=True, size=9, color="595959")
        h.alignment = Alignment("center", "center", wrap_text=True); h.border = BORDER
        s = ws[col + "2"]; s.value = "(업로드 제외)"; s.fill = REF_FILL; s.font = SUB_FONT
        s.alignment = Alignment("center", "center"); s.border = BORDER
    rr = 3
    for (dt, sq, ln, acct, drcr, rmk, amt, tr, dept, pjt, isu) in ROWS:
        vals = [dt, "1000", None, pjt, None, sq, ln, acct, drcr, rmk, amt, tr, dept, pjt, isu, "1", None, None, None]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(rr, c, v); cell.border = BORDER; cell.font = base
            cell.alignment = Alignment("center", "center") if c not in (10, 15) else Alignment("left", "center")
            if c == 11: cell.number_format = "#,##0"
            cell.fill = DR_FILL if drcr == 3 else CR_FILL
        if ln == 1 and dt in COPY_REF:
            rd, rno, rtype = COPY_REF[dt]
            for col, val in (("T", rd), ("U", rno), ("V", rtype)):
                rc = ws[col + str(rr)]; rc.value = val; rc.fill = REF_FILL; rc.border = BORDER
                rc.font = Font(name=FONT, size=9, color="595959"); rc.alignment = Alignment("center", "center")
                if col == "U": rc.number_format = "0"
        else:
            for col in ("T", "U", "V"): ws[col + str(rr)].fill = REF_FILL; ws[col + str(rr)].border = BORDER
        rr += 1
    ws.cell(rr, 9, "합계").font = Font(name=FONT, bold=True)
    ws.cell(rr, 10, "차변합계").font = Font(name=FONT, bold=True)
    ws.cell(rr, 11, f"=SUMIF(I3:I{rr-1},3,K3:K{rr-1})").number_format = "#,##0"
    ws.cell(rr+1, 10, "대변합계").font = Font(name=FONT, bold=True)
    ws.cell(rr+1, 11, f"=SUMIF(I3:I{rr-1},4,K3:K{rr-1})").number_format = "#,##0"
    ws.cell(rr+2, 10, "차-대 (0이어야 함)").font = Font(name=FONT, bold=True)
    ws.cell(rr+2, 11, f"=K{rr}-K{rr+1}").number_format = "#,##0"
    for x in (rr, rr+1, rr+2):
        ws.cell(x, 10).fill = TOT_FILL; ws.cell(x, 11).fill = TOT_FILL; ws.cell(x, 11).font = Font(name=FONT, bold=True)
    ws.cell(rr+4, 1, "※ T~V열(전표복사 참조)은 더존 업로드 대상이 아닙니다. 업로드는 A~S열만 사용하세요.").font = Font(name=FONT, size=9, color="808080")
    for i, w in enumerate([11,9,8,15,9,9,9,10,8,16,12,9,9,11,12,9,14,8,9], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for col, w in zip(["T","U","V"], [14,12,16]): ws.column_dimensions[col].width = w

    # ③ 거래처매핑
    ws = wb.create_sheet("③거래처매핑")
    ws["A1"] = "거래처 자동매칭 + 특이사항"; ws["A1"].font = TITLE_FONT
    for c, h in enumerate(["인벤코드","거래처명","회계코드","사업번호","상태","건수","금액합","특이사항"], 1):
        cell = ws.cell(3, c, h); cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER; cell.alignment = Alignment("center", "center")
    for i, row in enumerate(VENDORS):
        for c, v in enumerate(row, 1):
            cell = ws.cell(4 + i, c, v); cell.border = BORDER; cell.font = base
            if c == 7: cell.number_format = "#,##0"
            if c == 8 and ("해외" in str(row[7]) or "더미" in str(row[7])): cell.fill = WARN_FILL
    for col, w in zip("ABCDEFGH", [10,18,10,14,9,7,12,52]): ws.column_dimensions[col].width = w

    # ④ PJT매핑
    ws = wb.create_sheet("④PJT매핑")
    ws["A1"] = "프로젝트 매칭 / 92변환 / 전기재고 강제변환(규칙2)"; ws["A1"].font = TITLE_FONT
    ws["A3"] = "사용/변환 PJT는 references/master-data.md 및 이번 달 데이터 기준으로 기재."; ws["A3"].font = Font(name=FONT, size=10, color="808080")
    ws["A4"] = "26기 → 000092 변환 건수: (해당 시 기재)"; ws["A4"].font = Font(name=FONT, bold=True, color="C00000")

    # ⑤ 분개검토표 (②와 동일 내용 + 참조/변환 메모)
    ws = wb.create_sheet("⑤분개검토표")
    ws["A1"] = f"{YEAR}년 {MONTH}월 분개 검토표"; ws["A1"].font = TITLE_FONT
    hdr = ["작성일","전표번호","전표복사 참조","계정코드","차대","금액","거래처","PJT","적요","품의","변환/특이"]
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(3, c, h); cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER; cell.alignment = Alignment("center", "center", wrap_text=True)
    for i, (dt, sq, ln, acct, drcr, rmk, amt, tr, dept, pjt, isu) in enumerate(ROWS):
        ref = ""
        if ln == 1 and dt in COPY_REF:
            rd, rno, rtype = COPY_REF[dt]; ref = f"{rd} #{rno} ({rtype})"
        vals = [dt, sq, ref, acct, "차변" if drcr == 3 else "대변", amt, tr, pjt or "(공란)", rmk, isu, ""]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(4 + i, c, v); cell.border = BORDER; cell.font = base
            if c == 6: cell.number_format = "#,##0"
            cell.fill = DR_FILL if drcr == 3 else CR_FILL
    for col, w in zip("ABCDEFGHIJK", [11,10,22,10,6,11,10,10,16,12,20]): ws.column_dimensions[col].width = w

    # ⑥ 월별계정변동
    ws = wb.create_sheet("⑥월별계정변동")
    ws["A1"] = f"{YEAR}년 {MONTH}월 계정 변동액 (출고후/차변 기준)"; ws["A1"].font = TITLE_FONT
    for c, h in enumerate(["계정과목","계정코드","건수","금액","비중%"], 1):
        cell = ws.cell(3, c, h); cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER; cell.alignment = Alignment("center", "center")
    n = len(ACCTS); tot_row = 4 + n
    for i, (nm, cd, cnt, amt) in enumerate(ACCTS):
        r = 4 + i
        for c, v in enumerate([nm, cd, cnt, amt], 1):
            cell = ws.cell(r, c, v); cell.border = BORDER; cell.font = base
            if c == 4: cell.number_format = "#,##0"
        pc = ws.cell(r, 5, f"=D{r}/$D${tot_row}"); pc.number_format = "0.0%"; pc.border = BORDER; pc.font = base
    ws.cell(tot_row, 1, "합계").font = Font(name=FONT, bold=True)
    ws.cell(tot_row, 3, f"=SUM(C4:C{tot_row-1})"); ws.cell(tot_row, 4, f"=SUM(D4:D{tot_row-1})").number_format = "#,##0"
    ws.cell(tot_row, 5, f"=D{tot_row}/$D${tot_row}").number_format = "0.0%"
    for c in range(1, 6): ws.cell(tot_row, c).fill = TOT_FILL; ws.cell(tot_row, c).border = BORDER; ws.cell(tot_row, c).font = Font(name=FONT, bold=True)
    for col, w in zip("ABCDE", [12,12,8,14,10]): ws.column_dimensions[col].width = w

    wb.save(OUT)
    # 텍스트 서식(선행 0 보존) 재적용
    wb2 = load_workbook(OUT); ws2 = wb2["②자동전표업로드"]
    for r in range(3, 3 + len(ROWS)):
        for col in TEXT_COLS: ws2[col + str(r)].number_format = "@"
    wb2.save(OUT)
    print("saved:", OUT)
    print("차변합계 =", dr, "/ 대변합계 =", cr, "/ 차이 =", dr - cr)
    if dr != cr:
        print("!!! 차대 불일치 — 데이터 재확인 필요")

if __name__ == "__main__":
    build()
    # 저장 후 수식(SUMIF 등) 캐시값을 채우려면 재계산이 필요하다.
    #  - Claude.ai(파일 생성 기능): python /mnt/skills/public/xlsx/scripts/recalc.py <OUT>
    #  - Claude Code(로컬): recalc.py가 없으면 LibreOffice로 재계산
    #      soffice --headless --calc --convert-to xlsx --outdir <dir> <OUT>
    #      LibreOffice가 없으면 재계산은 생략해도 됨(엑셀에서 파일을 열면 수식이 계산된다).
    # 재계산 후 반드시 03340/000001/20260515 선행 0 유지를 재확인한다.

#!/usr/bin/env python3
"""위멤버스 수임처정보 엑셀 → Airtable upsert 배치 JSON 변환.

사용법:
  python3 transform.py <위멤버스.xlsx> <existing_masters.json> <출력폴더>

existing_masters.json 형식 (SKILL.md 절차 1에서 생성):
  [{"id": "rec..", "bizno": "123-45-67890", "code": "ST-0001", "sources": ["위멤버스"]}, ...]

출력:
  master_upsert_N.json / source_upsert_N.json / access_upsert_N.json  (배치당 50건)
  report.json  (신규/갱신/누락 요약)
"""
import sys, json, os, re
import openpyxl

# ── 필드 ID (references/ids.md와 동일) ──────────────────────────────
M = dict(상호='fld5iJfds7xzFIXzO', 사업자번호='fldABqqWpY2cmXpcY', 대표자='fldkotucZZGaA6is8',
         업종='fldONn3f8YE38P8P5', 구분='fldEN0gRtZfEuJ9CI', 과세유형='fldEDuKrw7HPlOH7v',
         담당자='fldy2zO3xIUCuTuBX', 수임일='fldlS7HLJru1gdYA7', 상태='fldPRiqWe6nyRMlDy',
         연락처='fldDwYGDGHORpQOMc', 이메일='flda5xOQLSg2DsZnP', 기장료='fldWTj2mGswzJibz8',
         메모='fld3daCrlWVlZqnXh', 코드='fldxkvlkkj2CNVrIS', 소스='fldwAUHqApbu1PI0C')
S = dict(원본명='fldlxW8UqBeZ6uaov', 사업자번호='flddr4FHa5hLQ17zj', 대표자='fldx4UJyMpWAE4I1z',
         연락처='fldTBpPZtyMbYPh5f', 이메일='fldTxcshnK6ewf1LL', 원본데이터='fldjdnED4nObRJikF',
         매칭상태='fldYlJANj6nQnES2S', 링크='fldRbT8AsNNSh002e')
A = dict(항목명='fldMBUkbAxfEnQ9MF', 사이트='fldNLxDYNwpy6AWce', 인증방식='fldIIPgXvBKXmpR1D',
         동의='fldXAXkoMw5Bq7Q3l', 계정='fldEuokUtd1A6sC2V', 비번='fldcJRMIdvhFScyBz',
         링크='flduKFOC2ymJUTIiN')

# ── 소스_위멤버스 컬럼별 필드 (원본 컬럼 → (필드ID, 타입 t/d/n/i)) ────────
SRC_COLS = {
 '담당자': ('fldH8rBSlof59ncH6','t'), '구분': ('fldFkCWsw1WtG8j9V','t'), '과세유형': ('fldiwL2f6KwnP5i4T','t'),
 '과세유형 상세': ('fldJNxtSnbFwRU2xS','t'), '신고프로그램': ('fldlFau0lPOHCkkno','t'),
 '개업일': ('fldPptCYZW93jl36I','d'), '폐업일': ('fldp2nGMdiQryiblu','d'), '계약상태': ('fldgJoQdU8QAViq6z','t'),
 '업태': ('fld9r07GWBJRStyg0','t'), '종목': ('fldQ4jGXgvSaLhro0','t'),
 '주소 (지번)': ('fldvbUkG6ihagQ1v1','t'), '도로명 주소': ('fldXureGV0hbTz6h5','t'),
 '성실신고대상': ('fldnlfe1TI8LiZ32c','t'), '특이사항': ('fldexsRgrSrPSDrh4','t'), '공동대표': ('fldI6yPCLNu4s7gJr','t'),
 '홈택스 계정': ('fldoILw0dnmktih1y','t'), '홈택스 비번': ('fldqqJH0gCHS8pYxi','t'),
 '여신금융협회 계정': ('fldbFjVIhrdSnKpD8','t'), '여신금융협회 비번': ('fldxfjWUVuyx7BpFt','t'),
 'SNS 매출 계정': ('fldmspF20NARzN5xZ','t'), '현금영수증 가맹여부': ('fldI97uPx5CM1o9vd','t'),
 '신용카드 가맹여부': ('fldkV3nFRMmcyC3Ld','t'), '총괄납부 주사업장': ('flduwJcf4fNeNwtmZ','t'),
 '전자세금계산서 발급의무 대상자여부': ('fldW2cbhHxiDipXsj','t'), '원천세신고유형': ('fldFn1iPo3hqH8Ic7','t'),
 '관할세무서': ('fldjkzMBqTE2L4GDf','t'), '세무서 담당': ('fldSjuhb6WWvLmTmd','t'), '주업종코드': ('fldjhfAXERiO5Ye88','t'),
 '홈택스 수임동의': ('fldTN9sNUvQqVrqL6','t'), '수임기준일': ('fldfPtWYYo5DjE5Av','d'),
 '정보제공범위': ('fldRydwv79IdithPt','t'), '가입경로': ('fldUfTLtOwzj9LPwm','t'), '가입경로상세': ('fldMkIFLM7nkQV2JQ','t'),
 '기장료 출금시작일': ('fldk2oHqVWN26I0iW','d'), '월 기장료': ('fldAUWJOUKN7omMRE','n'),
 'CMS 회원번호': ('fldC0lNXOySvFaMmM','t'), '기장료 출금계좌': ('fldNnHFecMejOvW0E','t'),
 '사업용계좌개수': ('fldJqLnegIIFzPhyD','i'), '사업용계좌 조회일시': ('fldeE2phQRtPM4Szj','t'),
 '사업용카드개수': ('fldtK732lunsH5bhS','i'), '사업용카드 조회일시': ('fldxQkLEJ9GPPGBKP','t'),
 '업무량': ('fldm7xQfXXI0YrRZp','t'), '난이도': ('fldbemuY5o6eZUoIA','t'),
 '사업용도로명주소': ('fldZGLPXhhSXfYMEF','t'), '사업용법정동주소': ('fldSIFYSfIBvU2jeD','t'),
 '대표자도로명주소': ('flda19pcD3nPha9qu','t'), '대표자법정동주소': ('fldShWfNJfEgLyhYv','t'),
 '법인번호': ('fldQcMDWO65Y6r8aM','t'), '거래처 등록일': ('fldTAIGnJzNDX51fN','d'),
 '홈택스정보 조회일시': ('fldy1eyh7cAIZofgi','t'),
}
REP = dict(항목명='fld7YVXHxlPMAdb9T', 거래처명='fldjpZSuacRNIo6SU', 사업자번호='fld7jDic6pkvRDPqk',
           이름='fldW0nLbhvavSA8Gp', 주민='fldu3bJ7OzluwwdUl', 지분='fldAWknIqqxN8GNIQ',
           주대표='fldC3L7Y3xz7aOY1v', 취임='fldhfsbhiXaYAa0OJ', ID='fldt6INbyX1erqe83',
           PW='fldMvmkdOuZoVV5ju', 링크='fldJhetH57pQoSj2e')
YMK_REC = 'recjhgoOjfon8BRCV'
YMK_BIZNO = '135-81-42023'


def typed(v, ty):
    if ty == 'd':
        return v[:10]
    if ty == 'n':
        try:
            x = float(v)
            return x if x > 0 else None
        except ValueError:
            return None
    if ty == 'i':
        try:
            return int(float(v))
        except ValueError:
            return None
    return v



def load_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['회사정보']
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(h).strip() if h else '' for h in rows[0]]
    idx = {h: i for i, h in enumerate(hdr)}
    return [{h: (str(r[i]).strip() if r[i] is not None else '') for h, i in idx.items()}
            for r in rows[1:] if r[0]]


def vat_type(d):
    det = d.get('과세유형 상세', '')
    if '간이' in det or '간이' in d.get('과세유형', ''):
        return '간이과세'
    if '면세' in det:
        return '면세'
    return '일반과세'


def wonbon(d):
    parts = []
    def add(label, *keys):
        vals = [d[k] for k in keys if d.get(k)]
        if vals:
            parts.append(f"{label}: {' / '.join(vals)}")
    add('계약상태', '계약상태'); add('개업일', '개업일')
    add('업태/종목', '업태', '종목'); add('주소', '도로명 주소')
    add('관할세무서', '관할세무서'); add('세무서담당', '세무서 담당')
    add('주업종코드', '주업종코드'); add('홈택스ID', '홈택스 계정')
    add('홈택스 수임동의', '홈택스 수임동의'); add('원천세신고', '원천세신고유형')
    add('급여일', '급여일'); add('급여대상', '급여대상'); add('급여지급기준', '급여지급기준')
    add('가입경로', '가입경로', '가입경로상세'); add('기장료출금시작일', '기장료 출금시작일')
    add('CMS회원번호', 'CMS 회원번호'); add('CMS수납구분', 'CMS 수납구분'); add('출금계좌', '기장료 출금계좌')
    add('업무량/난이도', '업무량', '난이도'); add('미수잔액', '미수잔액')
    add('성실신고대상', '성실신고대상'); add('공동대표', '공동대표')
    add('폐업일', '폐업일'); add('폐업사유', '폐업사유'); add('특이사항', '특이사항')
    return '\n'.join(parts)


def master_fields(d, is_new, code=None):
    f = {M['사업자번호']: d['사업자번호'], M['상호']: d['거래처명'],
         M['상태']: '수임중', M['과세유형']: vat_type(d)}
    if d['대표자명'] and not d['대표자명'].startswith('O*'):
        f[M['대표자']] = d['대표자명']
    if d['업태'] or d['종목']:
        f[M['업종']] = ' / '.join(x for x in (d['업태'], d['종목']) if x)
    if d['구분'] in ('법인', '개인'):
        f[M['구분']] = d['구분']
    if d['담당자']:
        f[M['담당자']] = d['담당자']
    if d['수임기준일']:
        f[M['수임일']] = d['수임기준일'][:10]
    if d['대표연락처']:
        f[M['연락처']] = d['대표연락처']
    if d['대표메일']:
        f[M['이메일']] = d['대표메일']
    try:
        fee = float(d['월 기장료'])
        if fee > 0:
            f[M['기장료']] = fee
    except (ValueError, KeyError):
        pass
    if is_new:
        f[M['코드']] = code
        f[M['소스']] = ['위멤버스']
        if d['폐업일']:
            f[M['메모']] = f"폐업 ({d['폐업일']}{', ' + d['폐업사유'] if d['폐업사유'] else ''})"
    return f


def main():
    xlsx, existing_path, outdir = sys.argv[1], sys.argv[2], sys.argv[3]
    os.makedirs(outdir, exist_ok=True)
    data = load_rows(xlsx)
    existing = json.load(open(existing_path))
    by_bizno = {e['bizno']: e for e in existing if e.get('bizno')}
    max_code = max((int(m.group(1)) for e in existing
                    if (m := re.match(r'ST-(\d+)', e.get('code') or ''))), default=1)

    masters, sources, access = [], [], []
    new_list, upd_list = [], []
    file_biznos = set()
    for d in data:
        bizno, name = d['사업자번호'], d['거래처명']
        file_biznos.add(bizno)
        is_new = bizno not in by_bizno
        if is_new:
            max_code += 1
            masters.append({'fields': master_fields(d, True, f'ST-{max_code:04d}')})
            new_list.append(f'ST-{max_code:04d} {name}')
        else:
            masters.append({'fields': master_fields(d, False)})
            upd_list.append(name)

        link = [YMK_REC] if bizno == YMK_BIZNO else [name]
        sf = {S['원본명']: name, S['사업자번호']: bizno,
              S['매칭상태']: '신규생성' if is_new else '자동매칭', S['링크']: link}
        for k, col in (('대표자', '대표자명'), ('연락처', '대표연락처'), ('이메일', '대표메일')):
            if d[col]:
                sf[S[k]] = d[col]
        for col, (fid, ty) in SRC_COLS.items():
            v = d.get(col, '')
            if v:
                tv = typed(v, ty)
                if tv is not None:
                    sf[fid] = tv
        sources.append({'fields': sf})

        alink = [YMK_REC] if bizno == YMK_BIZNO else [name]
        dong = d['홈택스 수임동의']
        if d['홈택스 계정'] or d['홈택스 비번'] or dong:
            af = {A['항목명']: f'홈택스 — {name}', A['사이트']: '홈택스',
                  A['인증방식']: 'ID/비밀번호', A['링크']: alink}
            if d['홈택스 계정']: af[A['계정']] = d['홈택스 계정']
            if d['홈택스 비번']: af[A['비번']] = d['홈택스 비번']
            if dong: af[A['동의']] = '완료' if dong == '동의' else '대기'
            access.append({'fields': af})
        if d['여신금융협회 계정'] or d['여신금융협회 비번']:
            af = {A['항목명']: f'여신금융협회 — {name}', A['사이트']: '여신금융협회',
                  A['인증방식']: 'ID/비밀번호', A['링크']: alink}
            if d['여신금융협회 계정']: af[A['계정']] = d['여신금융협회 계정']
            if d['여신금융협회 비번']: af[A['비번']] = d['여신금융협회 비번']
            access.append({'fields': af})

    reps = []
    wb2 = openpyxl.load_workbook(xlsx, data_only=True)
    if '대표자정보' in wb2.sheetnames:
        ws2 = wb2['대표자정보']
        rows2 = list(ws2.iter_rows(values_only=True))
        hdr2 = [str(h).strip() if h else '' for h in rows2[0]]
        idx2 = {h: i for i, h in enumerate(hdr2)}
        for r in rows2[1:]:
            if not r[0]:
                continue
            d2 = {h: (str(r[i]).strip() if r[i] is not None else '') for h, i in idx2.items()}
            nm, bz = d2['거래처명'], d2['사업자번호']
            f2 = {REP['항목명']: (f"{d2['이름']} — {nm}" if d2['이름'] else nm),
                  REP['거래처명']: nm, REP['사업자번호']: bz,
                  REP['링크']: [YMK_REC] if bz == YMK_BIZNO else [nm]}
            if d2['이름']: f2[REP['이름']] = d2['이름']
            if d2['주민등록번호']: f2[REP['주민']] = d2['주민등록번호']
            if d2['지분율']:
                try: f2[REP['지분']] = int(float(d2['지분율']))
                except ValueError: pass
            if d2['주대표'] in ('Y', 'N'): f2[REP['주대표']] = d2['주대표']
            if d2['취임일']: f2[REP['취임']] = d2['취임일'][:10]
            if d2['홈택스ID']: f2[REP['ID']] = d2['홈택스ID']
            if d2['홈택스PW']: f2[REP['PW']] = d2['홈택스PW']
            reps.append({'fields': f2})

    for prefix, recs in (('master_upsert', masters), ('source_upsert', sources),
                         ('access_upsert', access), ('rep_upsert', reps)):
        for i in range(0, len(recs), 50):
            with open(os.path.join(outdir, f'{prefix}_{i//50}.json'), 'w') as fp:
                json.dump(recs[i:i+50], fp, ensure_ascii=False)

    missing = [e for e in existing
               if e.get('bizno') and e['bizno'] not in file_biznos
               and '위멤버스' in (e.get('sources') or [])]
    report = {'file_rows': len(data), 'new': len(new_list), 'updated': len(upd_list),
              'new_clients': new_list,
              'missing_from_file': [f"{e.get('code', '?')} (bizno {e['bizno']})" for e in missing],
              'rep_rows': len(reps),
              'batches': {'master': (len(masters)+49)//50, 'source': (len(sources)+49)//50,
                          'access': (len(access)+49)//50, 'rep': (len(reps)+49)//50}}
    with open(os.path.join(outdir, 'report.json'), 'w') as fp:
        json.dump(report, fp, ensure_ascii=False, indent=1)
    print(json.dumps(report, ensure_ascii=False, indent=1))


if __name__ == '__main__':
    main()

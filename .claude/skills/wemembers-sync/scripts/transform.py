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

        sf = {S['원본명']: name, S['사업자번호']: bizno, S['원본데이터']: wonbon(d),
              S['매칭상태']: '신규생성' if is_new else '자동매칭', S['링크']: [name]}
        for k, col in (('대표자', '대표자명'), ('연락처', '대표연락처'), ('이메일', '대표메일')):
            if d[col]:
                sf[S[k]] = d[col]
        sources.append({'fields': sf})

        dong = d['홈택스 수임동의']
        if d['홈택스 계정'] or d['홈택스 비번'] or dong:
            af = {A['항목명']: f'홈택스 — {name}', A['사이트']: '홈택스',
                  A['인증방식']: 'ID/비밀번호', A['링크']: [name]}
            if d['홈택스 계정']: af[A['계정']] = d['홈택스 계정']
            if d['홈택스 비번']: af[A['비번']] = d['홈택스 비번']
            if dong: af[A['동의']] = '완료' if dong == '동의' else '대기'
            access.append({'fields': af})
        if d['여신금융협회 계정'] or d['여신금융협회 비번']:
            af = {A['항목명']: f'여신금융협회 — {name}', A['사이트']: '여신금융협회',
                  A['인증방식']: 'ID/비밀번호', A['링크']: [name]}
            if d['여신금융협회 계정']: af[A['계정']] = d['여신금융협회 계정']
            if d['여신금융협회 비번']: af[A['비번']] = d['여신금융협회 비번']
            access.append({'fields': af})

    for prefix, recs in (('master_upsert', masters), ('source_upsert', sources), ('access_upsert', access)):
        for i in range(0, len(recs), 50):
            with open(os.path.join(outdir, f'{prefix}_{i//50}.json'), 'w') as fp:
                json.dump(recs[i:i+50], fp, ensure_ascii=False)

    missing = [e for e in existing
               if e.get('bizno') and e['bizno'] not in file_biznos
               and '위멤버스' in (e.get('sources') or [])]
    report = {'file_rows': len(data), 'new': len(new_list), 'updated': len(upd_list),
              'new_clients': new_list,
              'missing_from_file': [f"{e.get('code', '?')} (bizno {e['bizno']})" for e in missing],
              'batches': {'master': (len(masters)+49)//50, 'source': (len(sources)+49)//50,
                          'access': (len(access)+49)//50}}
    with open(os.path.join(outdir, 'report.json'), 'w') as fp:
        json.dump(report, fp, ensure_ascii=False, indent=1)
    print(json.dumps(report, ensure_ascii=False, indent=1))


if __name__ == '__main__':
    main()

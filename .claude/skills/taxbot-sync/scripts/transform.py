# -*- coding: utf-8 -*-
"""택스봇 수임고객정보 → Airtable payload 변환.

phase1: 마스터 신규/갱신 payload
    python3 transform.py <수임처등록정보.xlsx> <연락처정보.xlsx> <existing_masters.json> <출력폴더>
phase2: 소스_택스봇 / 연락처 / 접속정보 payload (마스터 반영·재조회 후)
    python3 transform.py --phase2 <출력폴더> <master_map.json> <existing_access.json>

existing_masters.json: [{"id","bizno","code","sources","name","ceo","mgr","phone","email","gubun","date","status"}, ...]
master_map.json: {"by_biz": {bizno: {"id","code","name"}}, "by_code": {code: {...}}}
existing_access.json: [{"id","item","site","login","pw","consent"}, ...]
"""
import json, sys, re, collections, os

REAL_MGRS = {'이동희', '이태우', '류재상', '이준호'}

MF = {'상호': 'fld5iJfds7xzFIXzO', '사업자번호': 'fldABqqWpY2cmXpcY', '대표자': 'fldkotucZZGaA6is8',
      '업종': 'fldONn3f8YE38P8P5', '구분': 'fldEN0gRtZfEuJ9CI', '과세유형': 'fldEDuKrw7HPlOH7v',
      '담당자': 'fldy2zO3xIUCuTuBX', '수임일': 'fldlS7HLJru1gdYA7', '상태': 'fldPRiqWe6nyRMlDy',
      '연락처': 'fldDwYGDGHORpQOMc', '이메일': 'flda5xOQLSg2DsZnP', '메모': 'fld3daCrlWVlZqnXh',
      '거래처코드': 'fldxkvlkkj2CNVrIS', '소스': 'fldwAUHqApbu1PI0C'}

SF = {'원본명': 'fldY5cWQWelPkL7mU', '사업자번호': 'fldalu48zlznSBqM4', '대표자': 'fldRGjuG8BFR85Q7E',
      '연락처': 'fldzQkRktlYsAyQtQ', '이메일': 'fld0pA3FE3Av9i8bX', '매칭상태': 'fldZjPMSWMomUSE2n',
      '거래처': 'fldWYodIsGMJ7WSqO',
      '담당자명(세무대리인)': 'fldfbiTZO4gvw1WSk', '홈택스부서사용자ID': 'fldlSfHmM4NIMx34C',
      '주민등록번호': 'fldfgKHoCrdUffpyC', '대리구분': 'fldIzptnskQAjA7oR',
      '세무프로그램 회사코드': 'fldXsSaixEUOLVsPq', '사업자구분': 'fldZGIemZYNFXu8a2',
      '정보제공범위': 'fldcTCQtY8hHAR7d6', '사업장전화번호': 'fldFiChmFGIAwLXnx',
      '휴대전화번호': 'fld6LeIHwKzYTGIZF', '사업장도로명주소': 'fldhvKJuTBHpUIHS5',
      '사업장법정동주소': 'fldbtztePUkt0mVZ2', '개업일': 'fldtxB7kRGxRrmVHe', '폐업일': 'fldKKMDSok4awtPz0',
      '주업종코드': 'fldAPdZpBi1PnQciG', '업태': 'fldmEqizPAZzSwu86', '종목': 'fld3KmGqhJC0hkbrZ',
      '현금영수증가맹여부': 'fldCsAJKzH9sYQsWc', '신용카드가맹여부': 'fld9rO1YXRJo0M19e',
      '원천징수의무구분': 'fldMhNMZNuBcIzuf7', '전자세금계산서발급의무대상자여부': 'fldzGEg5mA9ElKO5g',
      '관할세무서/담당자': 'fldcBcywFGhSGLwpe', '총괄납부주업장여부': 'fldAjQ8NUBue9KIFR',
      '법인등록번호': 'fldGu0FE4ISXMowal', '회계시작일': 'fldbbw5lWJkHYiIWQ', '회계종료일': 'fldEkoXKwSAyj7Jdo',
      '수임일자': 'fldWL8zlYRfVqQvcO', '동의일자': 'fldT9yLFkdK1KEm2e', '해임일자': 'fldSQSolfSlKXpxP1',
      '홈택스ID': 'fldzJjjUQhoqpttM8', '홈택스PW': 'fldLT0mdkxcNxA9QM',
      '홈택스2차인증번호': 'fldFdW2USgVM81PqR', '환급계좌 은행': 'fld0SiACDwD7Wvfj3',
      '계좌번호': 'fldfNi1n5JZEw4hMi', '예금주명': 'fldjmqnLDAmmnkYQj'}
COMMON = ('원본명', '사업자번호', '대표자', '연락처', '이메일', '매칭상태', '거래처')
PASSTHRU = [k for k in SF if k not in COMMON]

CF = {'항목명': 'fldO64eddwm02L4Kc', '거래처명': 'fldZcwEYF0gyviLTn', '사업자번호': 'fldYWdC7u5o63mzzM',
      '대표자명': 'fldNMShX4da6i49Zm', '구분': 'fldxcpL7uqmwJkc7l', '성명': 'fldXsLEH9t4uSLlkO',
      '핸드폰': 'fldtryVR5JwW7Bm8z', '이메일': 'fldDiwPZN9H1ryydC', '알림톡': 'fld5JczQZRKU5X1oM',
      'LMS': 'fldWHTySsAzySppQW', '이메일채널': 'fldygiZgFw5foIZGa', '자동반영': 'fldoPRp2R2ExA51TN',
      '세무담당자': 'fldK46EO8PLXwWaU9', '부서ID': 'fldiC5wfrwOBKtKGI', '거래처': 'fldbLSUR3IQHrWr18'}

AF = {'항목명': 'fldMBUkbAxfEnQ9MF', '사이트': 'fldNLxDYNwpy6AWce', '인증방식': 'fldIIPgXvBKXmpR1D',
      '수임동의': 'fldXAXkoMw5Bq7Q3l', '계정ID': 'fldEuokUtd1A6sC2V', '비밀번호': 'fldcJRMIdvhFScyBz',
      '거래처': 'flduKFOC2ymJUTIiN', '메모': 'fldTtdRuB1dErIUUD'}


def cell(v):
    if v is None:
        return ''
    s = str(v).strip()
    if s.endswith(' 00:00:00'):
        s = s[:10]
    return s


def is_test(row):
    return (row['상호'] or row['대표자명']).lower() == 'test' or row['사업자등록번호'] == '111-22-33333'


def read_sheet(path, header_row):
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True)['Sheet1']
    rows = list(ws.iter_rows(values_only=True))
    hdr = [cell(c) for c in rows[header_row - 1]]
    # 연락처정보 3번째 컬럼은 헤더 공란 → 홈택스부서사용자ID
    hdr = [h if h else '홈택스부서사용자ID' for h in hdr]
    out = []
    for r in rows[header_row:]:
        vals = [cell(c) for c in r]
        if any(vals):
            out.append(dict(zip(hdr, vals)))
    return out


def merge_clients(rows):
    """사업자번호 중복(공동대표) 행 병합 + 개인(번호 없음) 분리."""
    merged, order, nobiz = {}, [], []
    for d in rows:
        b = d['사업자등록번호']
        if not b:
            nobiz.append(dict(d))
            continue
        if b not in merged:
            merged[b] = dict(d)
            order.append(b)
        else:
            m = merged[b]
            names = [x.strip() for x in re.split('[,，]', m['대표자명']) if x.strip()]
            for x in re.split('[,，]', d['대표자명']):
                x = x.strip()
                if x and x not in names:
                    names.append(x)
            m['대표자명'] = ', '.join(names)
            for k, v in d.items():
                if k not in ('대표자명', 'no') and v and not m.get(k):
                    m[k] = v
    return [merged[b] for b in order] + nobiz


def best_contact(row, contacts_by_biz, contacts):
    """마스터용 실번호 전화/이메일. 빈 사업자번호 키는 절대 쓰지 않는다."""
    b = row['사업자등록번호']
    cs = contacts_by_biz.get(b, []) if b else []
    if not cs and not b:
        nm = row['상호'] or row['대표자명']
        cs = [c for c in contacts if c['상호'] == nm or (not c['상호'] and c['대표자명'] == row['대표자명'])]
    ceo_first = row['대표자명'].split(',')[0].strip()
    phone = email = ''
    for c in sorted(cs, key=lambda c: (c['성명'] != ceo_first, c['구분'] != '대표자')):
        if not phone and c['핸드폰']:
            phone = c['핸드폰']
        if not email and c['이메일']:
            email = c['이메일']
    return phone, email


def phase1(client_xlsx, contact_xlsx, masters_json, outdir):
    rows = read_sheet(client_xlsx, 5)
    contacts = read_sheet(contact_xlsx, 4)
    masters = json.load(open(masters_json))
    clients = merge_clients(rows)
    contacts_by_biz = collections.defaultdict(list)
    for c in contacts:
        contacts_by_biz[c['사업자등록번호']].append(c)

    bmap = {m['bizno']: m for m in masters if m['bizno']}
    max_code = max(int(m['code'][3:]) for m in masters if m['code'])
    next_code = max_code + 1

    new_payload, upd_payload, new_meta = [], [], []
    dismissed, closed = [], []
    for row in clients:
        b = row['사업자등록번호']
        name = row['상호'] or row['대표자명']
        if is_test(row):
            row['_match'] = '신규생성'
            continue
        phone, email = best_contact(row, contacts_by_biz, contacts)
        email = row['이메일주소'] or email
        mgr = row['담당자명(세무대리인)'] if row['담당자명(세무대리인)'] in REAL_MGRS else ''
        if b and b in bmap:
            m = bmap[b]
            f = {}
            if '택스봇' not in (m['sources'] or []):
                f[MF['소스']] = (m['sources'] or []) + ['택스봇']
            for key, val in (('ceo', row['대표자명']), ('mgr', mgr), ('phone', phone),
                             ('email', email), ('date', row['수임일자'])):
                fld = {'ceo': '대표자', 'mgr': '담당자', 'phone': '연락처', 'email': '이메일', 'date': '수임일'}[key]
                if not m[key] and val:
                    f[MF[fld]] = val
            if not m['gubun'] and row['사업자구분']:
                f[MF['구분']] = '법인' if row['사업자구분'] == '법인사업자' else '개인'
            if f:
                upd_payload.append({'id': m['id'], 'fields': f})
            if row['해임일자']:
                dismissed.append(f"{m['name']} ({b}) 해임일자 {row['해임일자']} — 기존 마스터, 상태 미변경")
            row['_match'] = '자동매칭'
        else:
            code = f'ST-{next_code:04d}'
            next_code += 1
            f = {MF['상호']: name, MF['대표자']: row['대표자명'], MF['거래처코드']: code,
                 MF['소스']: ['택스봇'], MF['상태']: '수임중'}
            if b:
                f[MF['사업자번호']] = b
            if row['사업자구분']:
                gubun = '법인' if row['사업자구분'] == '법인사업자' else '개인'
                f[MF['구분']] = gubun
                if gubun == '법인':
                    f[MF['과세유형']] = '법인'
            if row['업태'] or row['종목']:
                f[MF['업종']] = ' / '.join(x for x in (row['업태'], row['종목']) if x)
            if mgr:
                f[MF['담당자']] = mgr
            if row['수임일자']:
                f[MF['수임일']] = row['수임일자']
            if phone:
                f[MF['연락처']] = phone
            if email:
                f[MF['이메일']] = email
            memo = []
            if not b:
                memo.append('사업자번호 없음 (개인 신고대리 고객)')
            if row['대리구분'] == '신고대리':
                memo.append('신고대리')
            if row['폐업일']:
                memo.append(f"폐업일 {row['폐업일']}")
            if row['해임일자']:
                memo.append(f"해임일자 {row['해임일자']}")
                f[MF['상태']] = '해지'
                dismissed.append(f"{name} ({b or '개인'}) 해임일자 {row['해임일자']} — 신규, 상태 해지로 생성")
            if row['담당자명(세무대리인)'] in ('대상제외', '담당자지정필요'):
                memo.append(f"택스봇 담당자: {row['담당자명(세무대리인)']}")
            if memo:
                f[MF['메모']] = '[택스봇] ' + ' · '.join(memo)
            new_payload.append({'fields': f})
            new_meta.append({'code': code, 'name': name, 'bizno': b})
            row['_match'] = '신규생성'
        if row['폐업일']:
            closed.append(f"{name} ({b or '개인'}) 폐업일 {row['폐업일']}")

    taxbot_biz = {r['사업자등록번호'] for r in clients if r['사업자등록번호']}
    missing = [f"{m['code']} {m['name']} ({m['bizno']})" for m in masters if m['bizno'] not in taxbot_biz]
    report = {'신규': len(new_payload), '기존갱신': len(upd_payload),
              '신규코드': f"ST-{max_code + 1:04d} ~ ST-{next_code - 1:04d}" if new_payload else '없음',
              '해임': dismissed, '폐업': closed, '택스봇에_없는_기존마스터': missing}
    os.makedirs(outdir, exist_ok=True)
    json.dump(clients, open(f'{outdir}/clients_merged.json', 'w'), ensure_ascii=False, indent=1)
    json.dump(contacts, open(f'{outdir}/taxbot_contacts.json', 'w'), ensure_ascii=False, indent=1)
    json.dump(new_payload, open(f'{outdir}/master_new.json', 'w'), ensure_ascii=False)
    json.dump(upd_payload, open(f'{outdir}/master_update.json', 'w'), ensure_ascii=False)
    json.dump(new_meta, open(f'{outdir}/master_new_meta.json', 'w'), ensure_ascii=False, indent=1)
    json.dump(report, open(f'{outdir}/report_phase1.json', 'w'), ensure_ascii=False, indent=1)
    print(json.dumps({k: (v if isinstance(v, (int, str)) else len(v)) for k, v in report.items()}, ensure_ascii=False))


def phase2(outdir, map_json, access_json):
    clients = json.load(open(f'{outdir}/clients_merged.json'))
    contacts = json.load(open(f'{outdir}/taxbot_contacts.json'))
    new_meta = json.load(open(f'{outdir}/master_new_meta.json'))
    mm = json.load(open(map_json))
    mm_by_biz, mm_by_code = mm['by_biz'], mm['by_code']
    access_existing = json.load(open(access_json))
    name_to_code = {m['name']: m['code'] for m in new_meta if not m['bizno']}

    def master_for(row):
        b = row['사업자등록번호']
        if b and b in mm_by_biz:
            return mm_by_biz[b]
        code = name_to_code.get(row['상호'] or row['대표자명'])
        return mm_by_code.get(code) if code else None

    src_up, src_cr, unlinked = [], [], []
    for row in clients:
        m = master_for(row) if not is_test(row) else None
        nm = row['상호'] or row['대표자명']
        f = {SF['원본명']: nm}
        for src, dst in (('사업자등록번호', '사업자번호'), ('대표자명', '대표자'),
                         ('휴대전화번호', '연락처'), ('이메일주소', '이메일')):
            if row[src]:
                f[SF[dst]] = row[src]
        f[SF['매칭상태']] = '신규생성' if row.get('_match') == '신규생성' else '자동매칭'
        if m:
            f[SF['거래처']] = [m['id']]
        else:
            unlinked.append(nm)
        for col in PASSTHRU:
            if row.get(col):
                f[SF[col]] = row[col]
        (src_up if row['사업자등록번호'] else src_cr).append({'fields': f})

    seen, keycount, con_cr = set(), collections.Counter(), []
    for c in contacts:
        dk = (c['성명'], c['상호'], c['사업자등록번호'], c['핸드폰'], c['구분'])
        if dk in seen:
            continue
        seen.add(dk)
        b = c['사업자등록번호']
        m = mm_by_biz.get(b)
        if not m and not b:
            code = name_to_code.get(c['대표자명'])
            m = mm_by_code.get(code) if code else None
        key = f"{c['성명']} — {c['상호'] or c['대표자명']}"
        keycount[key] += 1
        if keycount[key] > 1:
            key = f"{key} ({keycount[key]})"
        f = {CF['항목명']: key, CF['성명']: c['성명']}
        for src, dst in (('상호', '거래처명'), ('사업자등록번호', '사업자번호'), ('대표자명', '대표자명'),
                         ('구분', '구분'), ('핸드폰', '핸드폰'), ('이메일', '이메일'),
                         ('발송채널_알림톡', '알림톡'), ('발송채널_LMS', 'LMS'),
                         ('발송채널_이메일', '이메일채널'), ('발송하기자동반영', '자동반영'),
                         ('세무담당자', '세무담당자'), ('홈택스부서사용자ID', '부서ID')):
            if c.get(src):
                f[CF[dst]] = c[src]
        if m:
            f[CF['거래처']] = [m['id']]
        con_cr.append({'fields': f})

    acc_by_item = {a['item']: a for a in access_existing}
    acc_cr, acc_up, acc_conf = [], [], []
    for row in clients:
        if not row['홈택스ID'] or is_test(row):
            continue
        m = master_for(row)
        item = f"홈택스 — {m['name'] if m else (row['상호'] or row['대표자명'])}"
        consent = '완료' if row['동의일자'] else '대기'
        ex = acc_by_item.get(item)
        if ex:
            if not ex['login'] or ex['login'] == row['홈택스ID']:
                f = {AF['계정ID']: row['홈택스ID'], AF['비밀번호']: row['홈택스PW']}
                if consent == '완료' and ex.get('consent') != '완료':
                    f[AF['수임동의']] = '완료'
                if ex['login'] == row['홈택스ID'] and ex['pw'] == row['홈택스PW'] and AF['수임동의'] not in f:
                    continue
                acc_up.append({'id': ex['id'], 'fields': f})
            else:
                acc_conf.append(f"{item}: 기존 계정ID {ex['login']} ↔ 택스봇 {row['홈택스ID']} (미변경)")
        else:
            f = {AF['항목명']: item, AF['사이트']: '홈택스', AF['인증방식']: 'ID/비밀번호',
                 AF['수임동의']: consent, AF['계정ID']: row['홈택스ID'], AF['비밀번호']: row['홈택스PW']}
            if row['홈택스2차인증번호']:
                f[AF['메모']] = f"홈택스 2차인증번호: {row['홈택스2차인증번호']}"
            if m:
                f[AF['거래처']] = [m['id']]
            acc_cr.append({'fields': f})

    json.dump(src_up, open(f'{outdir}/p2_source_upserts.json', 'w'), ensure_ascii=False)
    json.dump(src_cr, open(f'{outdir}/p2_source_creates.json', 'w'), ensure_ascii=False)
    json.dump(con_cr, open(f'{outdir}/p2_contacts.json', 'w'), ensure_ascii=False)
    json.dump(acc_cr, open(f'{outdir}/p2_access_creates.json', 'w'), ensure_ascii=False)
    json.dump(acc_up, open(f'{outdir}/p2_access_updates.json', 'w'), ensure_ascii=False)
    json.dump(acc_conf, open(f'{outdir}/p2_access_conflicts.json', 'w'), ensure_ascii=False, indent=1)
    print(json.dumps({'source_upsert': len(src_up), 'source_create': len(src_cr), 'contacts': len(con_cr),
                      'access_create': len(acc_cr), 'access_update': len(acc_up),
                      'access_conflict': len(acc_conf), 'unlinked': unlinked}, ensure_ascii=False))


if __name__ == '__main__':
    if sys.argv[1] == '--phase2':
        phase2(*sys.argv[2:5])
    else:
        phase1(*sys.argv[1:5])

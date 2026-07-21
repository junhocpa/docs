# Airtable ID 참조 — 택스봇 동기화 (apprWMoKzGUuIuNdW)

## 거래처 (마스터) — tblM3ud4wVMVQlc2c

위멤버스 스킬(`../wemembers-sync/references/ids.md`)과 동일. 택스봇에서 쓰는 필드:

| 필드 | ID | 택스봇 원천 |
|---|---|---|
| 상호 (primary) | fld5iJfds7xzFIXzO | 상호 (없으면 대표자명) |
| 사업자번호 | fldABqqWpY2cmXpcY | 사업자등록번호 |
| 대표자 | fldkotucZZGaA6is8 | 대표자명 (공동대표는 `, ` 병합) |
| 업종 | fldONn3f8YE38P8P5 | 업태 + " / " + 종목 |
| 구분 | fldEN0gRtZfEuJ9CI | 사업자구분 (개인사업자→개인, 법인사업자→법인) |
| 과세유형 | fldEDuKrw7HPlOH7v | 법인만 "법인" (개인은 판정 불가 — 미전송) |
| 담당자 | fldy2zO3xIUCuTuBX | 담당자명(세무대리인) — 실명만 |
| 수임일 | fldlS7HLJru1gdYA7 | 수임일자 |
| 상태 | fldPRiqWe6nyRMlDy | 해임일자 있으면 해지, 없으면 수임중 |
| 연락처 | fldDwYGDGHORpQOMc | 연락처정보 파일의 실번호 |
| 이메일 | flda5xOQLSg2DsZnP | 이메일주소 (없으면 연락처정보의 이메일) |
| 메모 | fld3daCrlWVlZqnXh | `[택스봇] ...` (신규만) |
| 거래처코드 | fldxkvlkkj2CNVrIS | ST-#### (신규만) |
| 소스 | fldwAUHqApbu1PI0C | "택스봇" (기존은 조회값과 병합해 전송) |

## 소스_택스봇 — tbl7tkjO8cKsHm0Sz

공통 필드: 원본명(primary) fldY5cWQWelPkL7mU · 사업자번호 fldalu48zlznSBqM4 (upsert 병합 키) · 대표자 fldRGjuG8BFR85Q7E · 연락처 fldzQkRktlYsAyQtQ · 이메일 fld0pA3FE3Av9i8bX · 원본데이터 fldFzwP6HD3fPMu6w (미사용 — 컬럼별 필드 사용) · 매칭상태 fldZjPMSWMomUSE2n (자동매칭/신규생성) · 거래처 fldWYodIsGMJ7WSqO (record id 배열)

수임처등록정보 컬럼별 필드 (2026-07-21 생성):

| 원천 컬럼 | 필드 ID | 타입 |
|---|---|---|
| 담당자명(세무대리인) | fldfbiTZO4gvw1WSk | singleLineText (필드명: 담당자) |
| 홈택스부서사용자ID | fldlSfHmM4NIMx34C | singleLineText |
| 주민등록번호 | fldfgKHoCrdUffpyC | singleLineText |
| 대리구분 | fldIzptnskQAjA7oR | singleSelect(기장대리/신고대리) |
| 세무프로그램 회사코드 | fldXsSaixEUOLVsPq | singleLineText |
| 사업자구분 | fldZGIemZYNFXu8a2 | singleSelect(개인사업자/법인사업자) |
| 정보제공범위 | fldcTCQtY8hHAR7d6 | singleLineText |
| 사업장전화번호 | fldFiChmFGIAwLXnx | singleLineText |
| 휴대전화번호 | fld6LeIHwKzYTGIZF | singleLineText (마스킹 원본) |
| 사업장도로명주소 | fldhvKJuTBHpUIHS5 | singleLineText |
| 사업장법정동주소 | fldbtztePUkt0mVZ2 | singleLineText |
| 개업일 | fldtxB7kRGxRrmVHe | date(iso) |
| 폐업일 | fldKKMDSok4awtPz0 | date(iso) |
| 주업종코드 | fldAPdZpBi1PnQciG | singleLineText |
| 업태 | fldmEqizPAZzSwu86 | singleLineText |
| 종목 | fld3KmGqhJC0hkbrZ | singleLineText |
| 현금영수증가맹여부 | fldCsAJKzH9sYQsWc | singleLineText |
| 신용카드가맹여부 | fld9rO1YXRJo0M19e | singleLineText |
| 원천징수의무구분 | fldMhNMZNuBcIzuf7 | singleLineText |
| 전자세금계산서발급의무대상자여부 | fldzGEg5mA9ElKO5g | singleLineText (필드명: 전자세금계산서 발급의무) |
| 관할세무서/담당자 | fldcBcywFGhSGLwpe | singleLineText |
| 총괄납부주업장여부 | fldAjQ8NUBue9KIFR | singleLineText (필드명: 총괄납부 주사업장) |
| 법인등록번호 | fldGu0FE4ISXMowal | singleLineText |
| 회계시작일 | fldbbw5lWJkHYiIWQ | singleLineText (MMDD) |
| 회계종료일 | fldEkoXKwSAyj7Jdo | singleLineText (MMDD) |
| 수임일자 | fldWL8zlYRfVqQvcO | date(iso) |
| 동의일자 | fldT9yLFkdK1KEm2e | date(iso) |
| 해임일자 | fldSQSolfSlKXpxP1 | date(iso) |
| 홈택스ID | fldzJjjUQhoqpttM8 | singleLineText |
| 홈택스PW | fldLT0mdkxcNxA9QM | singleLineText |
| 홈택스2차인증번호 | fldFdW2USgVM81PqR | singleLineText |
| 환급계좌 은행 | fld0SiACDwD7Wvfj3 | singleLineText |
| 계좌번호 | fldfNi1n5JZEw4hMi | singleLineText (필드명: 환급계좌번호) |
| 예금주명 | fldjmqnLDAmmnkYQj | singleLineText |

## 소스_택스봇_연락처 — tbliolNHdLfhGpoD7 (2026-07-21 생성)

| 필드 | ID | 원천 |
|---|---|---|
| 항목명 (primary) | fldO64eddwm02L4Kc | `성명 — 상호` (중복 시 ` (2)` 부여) |
| 거래처명 | fldZcwEYF0gyviLTn | 상호 |
| 사업자번호 | fldYWdC7u5o63mzzM | 사업자등록번호 |
| 대표자명 | fldNMShX4da6i49Zm | 대표자명 |
| 구분 | fldxcpL7uqmwJkc7l | singleSelect(대표자/담당자) |
| 성명 | fldXsLEH9t4uSLlkO | 성명 |
| 핸드폰 | fldtryVR5JwW7Bm8z | 핸드폰 (실번호) |
| 이메일 | fldDiwPZN9H1ryydC | 이메일 |
| 발송채널_알림톡 | fld5JczQZRKU5X1oM | 여/부 |
| 발송채널_LMS | fldWHTySsAzySppQW | 여/부 |
| 발송채널_이메일 | fldygiZgFw5foIZGa | 여/부 |
| 발송하기자동반영 | fldoPRp2R2ExA51TN | 여/부 |
| 세무담당자 | fldK46EO8PLXwWaU9 | 세무담당자 |
| 홈택스부서사용자ID | fldiC5wfrwOBKtKGI | (헤더 공란 3번째 컬럼) |
| 거래처 | fldbLSUR3IQHrWr18 | record id 배열 |

## 접속정보 — tblSVIJn4Dixg1nEn

위멤버스 스킬과 동일: 항목명 fldMBUkbAxfEnQ9MF · 사이트 fldNLxDYNwpy6AWce · 인증방식 fldIIPgXvBKXmpR1D · 수임동의 상태 fldXAXkoMw5Bq7Q3l · 계정ID fldEuokUtd1A6sC2V · 비밀번호 fldcJRMIdvhFScyBz · 메모 fldTtdRuB1dErIUUD · 거래처 flduKFOC2ymJUTIiN

- 항목명: `홈택스 — {마스터 상호}` — **위멤버스가 만든 기존 항목과 같은 표기를 써야 중복이 안 생긴다**
- 계정ID가 기존과 다르면 덮어쓰지 않고 충돌 보고

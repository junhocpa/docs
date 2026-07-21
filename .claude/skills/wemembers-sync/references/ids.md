# Airtable ID 참조 — 솔루션택스 업무관리 (apprWMoKzGUuIuNdW)

## 거래처 (마스터) — tblM3ud4wVMVQlc2c

| 필드 | ID | 타입 | 파일 원천 컬럼 |
|---|---|---|---|
| 상호 (primary) | fld5iJfds7xzFIXzO | singleLineText | 거래처명 |
| 사업자번호 | fldABqqWpY2cmXpcY | singleLineText | 사업자번호 (upsert 병합 키) |
| 대표자 | fldkotucZZGaA6is8 | singleLineText | 대표자명 |
| 업종 | fldONn3f8YE38P8P5 | singleLineText | 업태 + " / " + 종목 |
| 구분 | fldEN0gRtZfEuJ9CI | singleSelect(법인/개인) | 구분 |
| 과세유형 | fldEDuKrw7HPlOH7v | singleSelect(일반과세/간이과세/면세/법인) | 과세유형 상세에서 판정 |
| 결산월 | fldNzqti8aRIaXNKp | number | (위멤버스 파일에 없음) |
| 담당자 | fldy2zO3xIUCuTuBX | singleLineText | 담당자 |
| 수임일 | fldlS7HLJru1gdYA7 | date(iso) | 수임기준일 |
| 상태 | fldPRiqWe6nyRMlDy | singleSelect(수임중/상담중/해지) | 항상 수임중 (해지는 수동) |
| 연락처 | fldDwYGDGHORpQOMc | phoneNumber | 대표연락처 |
| 이메일 | flda5xOQLSg2DsZnP | email | 대표메일 |
| 월 기장료 | fldWTj2mGswzJibz8 | currency | 월 기장료 (>0일 때만) |
| 메모 | fld3daCrlWVlZqnXh | multilineText | 신규 시 폐업 정보만 (기존 upsert 시 미전송) |
| 거래처코드 | fldxkvlkkj2CNVrIS | singleLineText | ST-#### (기존 upsert 시 미전송) |
| 소스 | fldwAUHqApbu1PI0C | multipleSelects | 신규 시 ["위멤버스"] (기존 upsert 시 미전송) |

## 소스_위멤버스 — tblZubQ3Qg0qcesyH

| 필드 | ID | 파일 원천 |
|---|---|---|
| 원본명 (primary) | fldlxW8UqBeZ6uaov | 거래처명 |
| 사업자번호 | flddr4FHa5hLQ17zj | 사업자번호 (upsert 병합 키) |
| 대표자 | fldx4UJyMpWAE4I1z | 대표자명 |
| 연락처 | fldTBpPZtyMbYPh5f | 대표연락처 |
| 이메일 | fldTxcshnK6ewf1LL | 대표메일 |
| 원본데이터 | fldjdnED4nObRJikF | 주요 컬럼 key: value 병합 (transform.py 참조) |
| 매칭상태 | fldYlJANj6nQnES2S | 자동매칭(기존)/신규생성(신규) |
| 거래처 링크 | fldRbT8AsNNSh002e | 상호 문자열 (typecast) |

## 접속정보 — tblSVIJn4Dixg1nEn

| 필드 | ID | 파일 원천 |
|---|---|---|
| 항목명 (primary) | fldMBUkbAxfEnQ9MF | "홈택스 — {거래처명}" / "여신금융협회 — {거래처명}" (upsert 병합 키) |
| 사이트 | fldNLxDYNwpy6AWce | 홈택스/여신금융협회 |
| 인증방식 | fldIIPgXvBKXmpR1D | ID/비밀번호 |
| 인증서 만료일 | fldmEowDmTMaR0ScK | (수동 관리) |
| 수임동의 상태 | fldXAXkoMw5Bq7Q3l | 홈택스 수임동의: 동의→완료, 그 외 값→대기 |
| 보관위치 | fldsvBVJDVvhABFRx | (수동 관리) |
| 계정ID | fldEuokUtd1A6sC2V | 홈택스 계정 / 여신금융협회 계정 |
| 비밀번호 | fldcJRMIdvhFScyBz | 홈택스 비번 / 여신금융협회 비번 |
| 거래처 링크 | flduKFOC2ymJUTIiN | 상호 문자열 (typecast) |

## 기타 테이블 (동기화 대상 아님)

- 신고일정 tblDKpFPsf1yvONXG · 수수료수금 tblQG4aTGfLSRheOd · 업무태스크 tblPfNQnSNxglZHY4
- 소스_CMS tblXYKKE3PdAMxtXN · 소스_채널톡 tbloNwXaKg2E7o2vT · 소스_카카오톡 tblh7tnGJD9bjTIrZ · 소스_내부관리 tbl6dD2t8wbtvWtNO · 소스_택스봇 tbl7tkjO8cKsHm0Sz
- YMK 마스터 레코드: recjhgoOjfon8BRCV (거래처코드 ST-0001)

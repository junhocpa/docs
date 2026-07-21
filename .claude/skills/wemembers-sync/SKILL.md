---
name: wemembers-sync
description: 위멤버스 수임처정보 엑셀을 Airtable "솔루션택스 업무관리" 베이스에 동기화하는 스킬. 사용자가 위멤버스 수임처정보 엑셀(.xlsx, 회사정보/대표자정보 시트 포함)을 업로드하며 "위멤버스 반영해줘", "수임처 동기화", "위멤버스 업데이트", "거래처 목록 반영" 등을 요청하면 반드시 이 스킬을 사용한다. 신규 거래처는 마스터에 코드 부여해 생성, 기존 거래처는 갱신(upsert), 소스_위멤버스 원본과 접속정보(홈택스 계정·비밀번호 포함)를 함께 동기화하고, 이번 파일에서 빠진 거래처를 해지 후보로 보고한다.
---

# 위멤버스 수임처정보 → Airtable 동기화

## 대상

- 입력: 위멤버스 수임처정보 엑셀 (`회사정보` 시트 기준, 1행 헤더 69컬럼)
- 출력: Airtable 베이스 `apprWMoKzGUuIuNdW` (솔루션택스 업무관리)의 3개 테이블 upsert
- 비밀번호(홈택스·여신금융협회)는 접속정보 테이블에 **그대로 저장**한다 (대표 결정 사항 — 제외하거나 별도 보관 권유 금지)

## 테이블·필드 ID

`references/ids.md` 참조. 핵심:

| 테이블 | ID | upsert 병합 키 |
|---|---|---|
| 거래처 (마스터) | `tblM3ud4wVMVQlc2c` | 사업자번호 `fldABqqWpY2cmXpcY` |
| 소스_위멤버스 | `tblZubQ3Qg0qcesyH` | 사업자번호 `flddr4FHa5hLQ17zj` |
| 접속정보 | `tblSVIJn4Dixg1nEn` | 항목명 `fldMBUkbAxfEnQ9MF` (예: `홈택스 — 상호`) |

## 절차

1. **기존 마스터 조회**: ToolSearch로 Airtable 도구 로드 후 `list_records_for_table`(마스터, 필드: 사업자번호·거래처코드·소스)를 페이지네이션으로 전부 조회해 `existing_masters.json`으로 저장한다. 형식: `[{"id": "rec..", "bizno": "...", "code": "ST-0001", "sources": ["위멤버스"]}, ...]`
2. **변환 스크립트 실행**: `python3 scripts/transform.py <업로드파일.xlsx> <existing_masters.json> <출력폴더>` → 출력폴더에 `master_upsert_*.json`, `source_upsert_*.json`, `access_upsert_*.json`, `report.json` 생성
3. **upsert 실행** (반드시 이 순서로, 각 파일을 읽어 `update_records_for_table` 호출):
   - 마스터: `performUpsert: {fieldIdsToMergeOn: ["fldABqqWpY2cmXpcY"]}`, `typecast: true`
   - 소스_위멤버스: `performUpsert: {fieldIdsToMergeOn: ["flddr4FHa5hLQ17zj"]}`, `typecast: true`
   - 접속정보: `performUpsert: {fieldIdsToMergeOn: ["fldMBUkbAxfEnQ9MF"]}`, `typecast: true`
   - 배치당 최대 50건. 레코드 수가 많으면 대량 입력을 general-purpose 서브에이전트에 위임한다.
4. **검증·보고**: `report.json`의 내용(신규 N건·갱신 N건·부여 코드 범위·**이번 파일에서 빠진 기존 거래처 목록**)을 사용자에게 보고한다. 빠진 거래처는 해지 후보일 뿐이므로 **삭제·상태변경은 하지 말고** 목록만 보고한다.

## 변환 규칙 (transform.py가 구현)

- **신규/기존 판정**: 사업자번호 기준. 신규는 기존 최대 코드 다음 번호부터 `ST-####` 부여
- **기존 거래처의 마스터 upsert**: `거래처코드`·`소스`·`메모` 필드는 **보내지 않는다** (수기 편집·타 소스 태그 보존). 나머지(상호·대표자·연락처·기장료·담당자 등)는 파일 값으로 갱신
- **신규 거래처**: 전체 필드 + 새 코드 + 소스 `["위멤버스"]`
- **소스_위멤버스**: 원본 그대로 (원본명·사업자번호·대표자·연락처·이메일·원본데이터·매칭상태), 거래처 링크는 상호 문자열(typecast 매칭). 마스터 upsert가 먼저 완료되어야 이름 매칭이 성립한다
- **접속정보**: 홈택스(계정·비밀번호·수임동의), 여신금융협회(계정·비밀번호) — 값이 있는 것만. 비밀번호 변경은 upsert로 자동 반영
- 빈 문자열 필드는 payload에서 제거 (email/phone 타입 오류 방지)

## 주의

- 마스터 `상호`는 파일 값으로 갱신된다 (예: "㈜와이엠케이" → "주식회사 와이엠케이"). 링크 매칭 정합성을 위해 의도된 동작
- upsert에서 "multiple matches" 오류가 나면 해당 병합 키가 중복된 것 — 중복 레코드를 사용자에게 보고하고 해당 건은 건너뛴다
- 동기화 후 마스터 총 레코드 수가 (기존 + 신규)와 다르면 typecast 링크가 새 마스터를 만든 것이므로 즉시 확인한다

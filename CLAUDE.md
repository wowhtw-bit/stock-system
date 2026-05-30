# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

소규모 로제트 키링 재료 재고관리 시스템. **빌드/번들러/패키지 매니저 없음** — 세 개의 파일만 다룸:

| 파일 | 역할 | 배포 방식 |
|---|---|---|
| `dashboard.html` | 단일 HTML 파일 (vanilla JS, 인라인 CSS). 사용자가 로컬에서 더블클릭으로 실행. | 사용자에게 파일 전달 |
| `apps_script.gs` | Google Apps Script 백엔드. | 구글시트의 Apps Script 편집기에 **수동 복붙** 후 "새 배포"로 웹 앱 URL 발급 |
| `재고관리_템플릿.xlsx` | 초기 시트 양식 (6개 시트). | 구글드라이브 업로드 → Google Sheets로 변환 |

테스트 스위트, lint, CI 없음. 변경 검증은 (1) 코드 리뷰, (2) 필요시 사용자가 직접 Apps Script에 붙여넣고 브라우저에서 dashboard.html 열어 확인.

## 아키텍처 핵심

```
브라우저(dashboard.html)  →  fetch POST JSON  →  Apps Script 웹 앱  →  Google Sheets + Drive
```

- **인증 없음.** 웹 앱 URL이 곧 인증. 사용자는 첫 실행 시 URL을 localStorage에 저장 (`apiUrl`).
- **단일 엔드포인트.** `doPost`/`doGet`이 `handle_(e)`로 라우팅, `body.action` 문자열로 디스패치 (apps_script.gs:38–53). 새 기능 추가 시 이 switch에 case를 추가하고 dashboard.html의 `api('actionName', payload)` 호출과 일치시킬 것.
- **응답 형식 고정.** 항상 `{ ok: true, data }` 또는 `{ ok: false, error }`. `api()` 헬퍼(dashboard.html:585)가 `ok:false`면 throw.
- **상태 모델.** 클라이언트는 단일 `state` 객체에 전체 시트를 캐싱. 거의 모든 mutation 후 `refresh()` 호출하여 `getAll`로 전부 다시 로드 (낙관적 업데이트 없음). 시트가 커도 수십 행 수준이라 OK.

## 시트 스키마 — 반드시 지켜야 할 규칙

**1행 = 안내문(무시), 2행 = 헤더, 3행부터 = 데이터.** `readSheet_`가 이 구조를 가정하며 깨지면 모든 작업이 실패한다. 마이그레이션 함수 외에는 헤더 이름·위치를 변경하지 말 것.

| 시트 | 컬럼 |
|---|---|
| Materials (M) | `id`, `재료명`, `단위`, `현재재고`, `임계수량`, `메모`, `이미지URL` |
| Products (P) | `id`, `제품명`, `메모` |
| BOM (B) | `제품id`, `재료id`, `1개당_소요량` |
| ProductionLog (PL) | `logId`, `일시`, `제품id`, `생산수량`, `메모`, `취소여부` |
| StockLog (SL) | `logId`, `일시`, `재료id`, `변동수량`, `종류`, `메모`, `취소여부` |
| Settings (S) | `key`, `value`, `desc` |

- ID 발급: 재료=`M001`, 제품=`P001`, 생산로그=`L0001`, 재고로그=`S0001` (`nextEntityId_` / `nextLogId_`).
- 로그는 **append-only + 소프트 취소** (`취소여부` TRUE). `undoLog_`는 재고를 역산해 복원하고 플래그만 토글 — 행을 절대 삭제하지 않는다.

## 동시성·트랜잭션

쓰기 작업은 모두 `LockService.getScriptLock().waitLock(10000)`으로 직렬화 (이미지 업로드는 15초). 새 mutation 핸들러를 추가할 때 동일 패턴(`try { ... } finally { lock.releaseLock(); }`)을 따를 것 — 생산 입력은 BOM 조회 → 재고 검증 → 다중 재료 차감 → 로그 append를 한 락 안에서 수행해야 정합성이 깨지지 않는다.

## 이미지 처리

- 클라이언트가 캔버스로 리사이즈/JPEG 압축(`resizeImage`, dashboard.html:950) 후 base64로 전송.
- 서버는 `재고관리_이미지` Drive 폴더(없으면 자동 생성, ID는 ScriptProperties에 캐시)에 저장 → `ANYONE_WITH_LINK / VIEW` 공유 → `drive.google.com/thumbnail?id=...&sz=w800` URL을 시트에 저장.
- 재료 이미지 교체/삭제 시 이전 파일은 URL에서 ID를 파싱해 `setTrashed(true)`.

## 마이그레이션

스키마 변경이 필요하면 `apps_script.gs`에 `migrate_v2_threshold` 같은 일회성 함수를 추가하고, 사용자가 Apps Script 편집기에서 함수 선택 후 ▶️로 실행하게 안내한다. **자동 마이그레이션 없음** — 멱등하게 짜고 (이미 적용됐는지 헤더로 감지) `SpreadsheetApp.getUi().alert()`로 결과를 보고할 것.

## 보안·시크릿

- `my-project-*.json`, `*service-account*.json` 등은 `.gitignore`로 차단됨. 절대 커밋 금지.
- 웹 앱 URL은 코드에 하드코딩하지 않는다 — 사용자별로 localStorage에 저장.
- 민감 데이터를 다루는 배포라면 웹 앱 액세스를 "조직 내"로 제한하도록 README가 안내함.

## 코드 스타일 메모

- Apps Script는 V8이지만 코드 스타일은 ES5(var, function 선언, 콜백) — 일관성을 위해 유지.
- dashboard.html은 ES2017+ (async/await, template literals, optional chaining 일부) 사용 가능. 빌드 단계가 없으므로 최신 브라우저 가정.
- 한글 컬럼명을 JS 객체 키로 그대로 사용 (`row['재료명']`). 헤더 문자열과 정확히 일치해야 함.

# PO Sheet Sorter

Outlook 데스크톱으로 들어오는 **PO SHEET (PDF 첨부)** 메일을 읽어서,
PDF 안의 스타일#를 기준으로 지정 폴더에 규칙에 맞는 이름으로 저장하는 Outlook VBA 매크로입니다.

- Python이나 외부 프로그램 설치가 **필요 없습니다.** (Word의 PDF 텍스트 변환 기능 이용)
- 추출에 실패한 파일은 완료 처리하지 않고 **확인필요/오류로 분리**합니다.
- 처리 원장을 남겨 **같은 내용의 첨부를 두 번 처리하지 않습니다.** (파일 내용 SHA-256 기준)
- Outlook이 꺼져 있던 동안 온 메일도 **재시작 시 다시 찾아 처리**합니다.
- 순수 파싱/검증 로직은 Outlook 없이 **자동 회귀 테스트**로 검증할 수 있습니다.
- 특정 바이어 PO SHEET 양식 기준으로 작성되었습니다.

---

## 처리 흐름

```
새 메일 도착 (또는 시작 시 최근 N일 재검색)
        ↓
대상 메일 검증      제목 키워드 · 허용 발신자 · 첨부 개수
        ↓
첨부 저장           충돌하지 않는 고유 임시파일명
        ↓
첨부 검증           크기 상한 · %PDF- 시그니처
        ↓
중복 검사           첨부 내용 SHA-256이 원장에 이미 있으면 SKIP
        ↓
Word로 텍스트 변환   실패 시 → PO_Error
        ↓
필드 추출 · 검증     필수값 존재 · 값의 유일성(스타일#/오더번호/납기일/국가) · 날짜 형식
        ├─ 정상  → 스타일#별 폴더  [COMPLETED]
        ├─ 모호  → PO_Review      [REVIEW]
        └─ 오류  → PO_Error       [ERROR]
        ↓
로그 + 원장 기록 + 메일에 상태 속성 기록
```

파일 이동은 **모든 검증을 통과한 마지막 단계에서만** 수행합니다.

---

## 파일명 규칙

```
ORDER.{D코드} ({국가} PO#{오더번호} {납기일}).pdf
```

예시: `ORDER.D010125T000000 (AUS PO#10000000 250115).pdf`

| 항목 | 출처 | 필수 |
|---|---|:---:|
| `D010125T000000` | 바이어가 보낸 **원본 첨부파일명** 안의 `D######T######` | ✅ |
| `AUS`/`NZ` | PDF에 SYDNEY / MELBOURNE / BRISBANE / PERTH(AUS) 또는 AUCKLAND(NZ) 언급 | ✅ * |
| `10000000` | PDF의 `ORDER NUMBER` | ✅ |
| `250115` | PDF의 `DLV CONS DATE` (dd/mm/yy → yymmdd) | ✅ |

\* `COUNTRY_REQUIRED = False`로 두면 판별 실패 시 `XXX`로 저장됩니다 (기본값은 확인필요 처리).
서로 다른 국가의 도시명이 문서 안에 동시에 등장하면(청구지 주소 등과 섞이는 경우) 자동 판별하지
않고 확인필요로 분리합니다.

**필수값이 하나라도 없으면 완료 폴더로 가지 않습니다.** `PO_Review`로 분리되고 사유가 로그에 남습니다.

### 스타일# 인식 방식

```
STYLE PACK: 90000001 12AB34CDE567 SAMPLE PRODUCT DESCRIPTION
                     ^^^^^^^^^^^^ 이 값을 스타일#로 사용
```

스타일#·오더번호·납기일·국가 중 하나라도 PDF 안에서 **서로 다른 값이 2개 이상** 발견되면
자동 판별하지 않고 확인필요로 분리합니다. (서로 다른 PO 구간의 값이 섞이는 것을 방지)

---

## 상태 분류

| 상태 | 조건 | 파일 위치 |
|---|---|---|
| `COMPLETED` | 모든 필수값 검증 및 최종 이동 성공 | 스타일#별 폴더 |
| `REVIEW` | 필수값 누락, 값 다중, 국가 불명/모호, 스타일 폴더 없음, 텍스트 없음 | `PO_Review` |
| `ERROR` | Word 변환 실패, 손상/암호화 PDF, 시그니처 불일치, 크기 초과 | `PO_Error` |
| `SKIP` | 내용이 동일한 첨부 재수신, 처리 중 도착, 허용되지 않은 발신자 | 이동 없음 |

결과는 세 곳에 기록됩니다.

- `PO_Log.txt` — 사람이 읽는 처리 로그 (UTF-8)
- `PO_Ledger.csv` — 중복 방지용 원장 (UTF-8, `Key, ProcessedAt, EntryID, AttachmentIndex, AttachmentName, Status, Destination, Error`).
  `Key`는 첨부 내용의 SHA-256 해시(`sha256:...`)이며, 해시 계산에 실패한 환경에서는 `EntryID|첨부순번|파일명` 형식으로 폴백합니다.
- 원본 메일의 사용자 속성 — `PO_SORT_STATUS`, `PO_SORT_UPDATED_AT`, `PO_SORT_RESULT`

로그나 원장 파일 쓰기 자체가 실패해도(다른 프로그램이 파일을 열어둔 경우 등) 처리 파이프라인은
멈추지 않고 `PO_Log.txt.fallback.YYYYMMDD.txt`에 대신 기록합니다. `RunReconcileNow` 종료 시
실패 건수가 있으면 한 번만 알려줍니다.

---

## 요구 사항

- Windows용 Outlook 데스크톱 (설치형)
- Microsoft Word (같은 PC에 설치)
- 스타일#별 하위 폴더가 이미 만들어져 있는 상위 폴더

---

## 빠른 시작

1. `Alt + F11` → VBA 편집기 실행
2. `파일 > 파일 가져오기`로 [`modConfig.bas`](modConfig.bas), [`modParse.bas`](modParse.bas), [`modIO.bas`](modIO.bas), [`modTest.bas`](modTest.bas) 4개를 가져오기
3. 왼쪽 트리에서 `ThisOutlookSession` 더블클릭 → [`ThisOutlookSession.bas`](ThisOutlookSession.bas) 전체 내용 붙여넣기
4. `modConfig.bas`의 설정 상수 수정 — 특히 **`ALLOWED_SENDERS`는 반드시 실제 바이어 도메인으로 변경**
   (초기값 `@example.com` 그대로면 모든 메일이 건너뛰어집니다)
5. 매크로 보안 허용 후 Outlook 재시작

자세한 절차와 문제 해결은 **[`docs/SETUP.md`](docs/SETUP.md)** 를 참고하세요.

---

## 운영 시 주의

- **지정 PC 한 대에서만 실행하세요.** 여러 팀원이 동시에 실행하면 원장이 충돌해 중복 방지가 깨집니다.
- 공유 사서함을 쓸 경우 `SHARED_MAILBOX`에 주소를 넣되, 실행 책임자를 한 명으로 지정하세요.
- Outlook이 실행 중이어야 실시간 감지가 됩니다. 꺼져 있던 동안의 메일은 재시작 시 재검색으로 처리됩니다.
- `PO_Review` / `PO_Error`로 분류된 메일은 자동 재검색 대상에서 제외됩니다. 사유를 해소한 뒤
  다시 처리하려면 Outlook에서 해당 메일을 선택하고 `ForceReprocessSelected`를 실행하세요.
- **스캔 이미지 PDF는 인식되지 않습니다** (텍스트 기반 PDF만 지원). OCR이 필요하면 별도 도구가 필요합니다.
- PDF마다 Word 문서 변환이 필요하므로 대량 처리 시 Outlook 이벤트 처리가 지연될 수 있습니다.
  (Word 인스턴스는 재사용되지만, 변환 자체에 하드 타임아웃은 없습니다 — [알려진 한계](#알려진-한계) 참고)
- `PO_Review` / `PO_Error` 폴더는 자동 재처리되지 않으니 주기적으로 확인하세요.

---

## 커스터마이징

### 다른 국가 추가

`modParse.bas`의 `ExtractCountries` 함수에 조건을 추가합니다. (AUCKLAND → `NZ`는 이미 반영되어 있습니다)

```vb
If InStr(u, "AUCKLAND") > 0 Or InStr(u, "WELLINGTON") > 0 Then
    result.Add "NZ"
End If
```

국가별로 도시 조건을 하나의 `If`로 묶어 `result.Add`를 한 번만 호출하면, 같은 국가 코드가
`result`에 중복으로 들어가 유일성 검증(#10.2)이 오작동하는 것을 피할 수 있습니다.

### 다른 PO 양식 추가

`modParse.bas`의 `ParseAndValidate` 안의 정규식(`MatchDistinct` 호출부)을 수정하거나 양식별 분기를
추가합니다. 공급처가 여러 곳이 되면 추출 규칙을 공급처별 함수로 분리하는 편이 유지보수에 낫습니다.
`modTest.bas`에 회귀 테스트를 추가해두면 정규식을 고칠 때마다 기존 케이스가 안 깨지는지 바로 확인할 수 있습니다.

---

## 테스트

`modParse.bas`의 로직은 Outlook·Word·파일시스템에 의존하지 않는 순수 함수라, VBA 편집기에서
`modTest.bas`의 `RunAllTests`에 커서를 두고 F5를 누르면 Outlook 재시작이나 실제 PO PDF 없이
즉시 결과를 확인할 수 있습니다(직접 실행 창 `Ctrl+G`에 출력). Outlook/Word 연동이 필요한
나머지 시나리오는 `docs/SETUP.md`의 수동 테스트 표를 따르세요.

---

## 알려진 한계

코드 리뷰에서 지적된 항목 중 아직 반영하지 않았거나, 구조적으로 완전히 해결하기 어려운 것들입니다.

- Word `Documents.Open` 자체에는 하드 타임아웃이 없습니다. 인스턴스는 재사용하지만(20건마다 재시작),
  손상된 PDF가 Word의 복구 모드에 오래 머무르면 그 동안 Outlook 이벤트 처리가 블로킹됩니다.
  VBA 단일 스레드 구조상 완전한 해결은 어렵고, 근본적으로는 `pdftotext.exe` 같은 외부 프로세스로
  변환을 옮기는 것이 대안입니다(README 상단의 "외부 프로그램 불필요" 설계와는 상충).
- 원장(`PO_Ledger.csv`)이 CSV 파일 하나라 동시 접근에 취약합니다. 여러 프로세스가 동시에 쓰면
  깨질 수 있으므로 "지정 PC 한 대에서만 실행" 규칙이 여전히 중요합니다. 팀 규모가 커지면
  SQLite/Access 전환을 권장합니다.
- 원장은 로그와 달리 자동 로테이션하지 않습니다(로테이션 시 중복 방지 이력이 끊기기 때문). 파일이
  매우 커지면 `LedgerHasEntry`가 시작 시 1회 전체 스캔하는 비용이 늘어납니다.
- `FileSha256`이 쓰는 `System.Security.Cryptography.SHA256Managed`는 .NET Framework의 COM 등록에
  의존합니다. 등록되지 않은 환경에서는 자동으로 `EntryID|첨부순번|파일명` 키로 폴백하며, 이 경우
  README 상단에 적힌 "메일 이동에 안전"이 적용되지 않습니다.

---

## 라이선스

MIT License - [`LICENSE`](LICENSE) 참고

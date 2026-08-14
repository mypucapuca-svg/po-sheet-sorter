# PO Sheet Sorter

Outlook 데스크톱으로 들어오는 **PO SHEET (PDF 첨부)** 메일을 읽어서,
PDF 안의 스타일#를 기준으로 지정 폴더에 규칙에 맞는 이름으로 저장하는 Outlook VBA 매크로입니다.

- Python이나 외부 프로그램 설치가 **필요 없습니다.** (Word의 PDF 텍스트 변환 기능 이용)
- 추출에 실패한 파일은 완료 처리하지 않고 **확인필요/오류로 분리**합니다.
- 처리 원장을 남겨 **같은 메일을 두 번 처리하지 않습니다.**
- Outlook이 꺼져 있던 동안 온 메일도 **재시작 시 다시 찾아 처리**합니다.
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
Word로 텍스트 변환   실패 시 → PO_Error
        ↓
필드 추출 · 검증     필수값 존재 · 값의 유일성 · 날짜 형식
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

**필수값이 하나라도 없으면 완료 폴더로 가지 않습니다.** `PO_Review`로 분리되고 사유가 로그에 남습니다.

### 스타일# 인식 방식

```
STYLE PACK: 90000001 12AB34CDE567 SAMPLE PRODUCT DESCRIPTION
                     ^^^^^^^^^^^^ 이 값을 스타일#로 사용
```

스타일#·오더번호·납기일 중 하나라도 PDF 안에서 **서로 다른 값이 2개 이상** 발견되면
자동 판별하지 않고 확인필요로 분리합니다. (서로 다른 PO 구간의 값이 섞이는 것을 방지)

---

## 상태 분류

| 상태 | 조건 | 파일 위치 |
|---|---|---|
| `COMPLETED` | 모든 필수값 검증 및 최종 이동 성공 | 스타일#별 폴더 |
| `REVIEW` | 필수값 누락, 값 다중, 국가 불명, 스타일 폴더 없음, 텍스트 없음 | `PO_Review` |
| `ERROR` | Word 변환 실패, 손상/암호화 PDF, 시그니처 불일치, 크기 초과 | `PO_Error` |
| `SKIP` | 이미 처리 완료된 첨부, 허용되지 않은 발신자 | 이동 없음 |

결과는 세 곳에 기록됩니다.

- `PO_Log.txt` — 사람이 읽는 처리 로그
- `PO_Ledger.csv` — 중복 방지용 원장 (`Key, ProcessedAt, EntryID, AttachmentIndex, AttachmentName, Status, Destination, Error`)
- 원본 메일의 사용자 속성 — `PO_SORT_STATUS`, `PO_SORT_UPDATED_AT`, `PO_SORT_RESULT`

---

## 요구 사항

- Windows용 Outlook 데스크톱 (설치형)
- Microsoft Word (같은 PC에 설치)
- 스타일#별 하위 폴더가 이미 만들어져 있는 상위 폴더

---

## 빠른 시작

1. `Alt + F11` → `ThisOutlookSession` 더블클릭
2. [`ThisOutlookSession.bas`](ThisOutlookSession.bas) 전체 내용 붙여넣기
3. 코드 상단 설정 수정 — 특히 **`ALLOWED_SENDERS`는 반드시 실제 바이어 도메인으로 변경**
   (초기값 `@example.com` 그대로면 모든 메일이 건너뛰어집니다)
4. 매크로 보안 허용 후 Outlook 재시작

자세한 절차와 문제 해결은 **[`docs/SETUP.md`](docs/SETUP.md)** 를 참고하세요.

---

## 운영 시 주의

- **지정 PC 한 대에서만 실행하세요.** 여러 팀원이 동시에 실행하면 원장이 충돌해 중복 방지가 깨집니다.
- 공유 사서함을 쓸 경우 `SHARED_MAILBOX`에 주소를 넣되, 실행 책임자를 한 명으로 지정하세요.
- Outlook이 실행 중이어야 실시간 감지가 됩니다. 꺼져 있던 동안의 메일은 재시작 시 재검색으로 처리됩니다.
- `EntryID`는 메일을 다른 저장소로 옮기면 바뀌므로, 처리된 메일을 다른 사서함으로 이동하면 중복 처리될 수 있습니다.
- **스캔 이미지 PDF는 인식되지 않습니다** (텍스트 기반 PDF만 지원). OCR이 필요하면 별도 도구가 필요합니다.
- PDF마다 Word를 실행하므로 대량 처리 시 Outlook 이벤트 처리가 지연될 수 있습니다.
- `PO_Review` / `PO_Error` 폴더는 자동 재처리되지 않으니 주기적으로 확인하세요.

---

## 커스터마이징

### 다른 국가 추가

`ExtractCountry` 함수에 조건을 추가합니다. (AUCKLAND → `NZ`는 이미 반영되어 있습니다)

```vb
ElseIf InStr(u, "WELLINGTON") > 0 Then
    ExtractCountry = "NZ"
```

### 다른 PO 양식 추가

`ParseAndValidate` 안의 정규식(`MatchDistinct` 호출부)을 수정하거나 양식별 분기를 추가합니다.
공급처가 여러 곳이 되면 추출 규칙을 공급처별 함수로 분리하는 편이 유지보수에 낫습니다.

---

## 알려진 한계

코드 리뷰에서 지적된 항목 중 아직 반영하지 않은 것들입니다.

- Word 인스턴스를 첨부마다 새로 생성합니다. 대량 처리 시 재사용이나 처리 대기열이 더 안정적입니다.
- 원장이 CSV라 동시 접근에 취약합니다. 팀 규모가 커지면 SQLite/Access 전환을 권장합니다.
- 실제 PO 샘플을 이용한 자동 회귀 테스트가 없습니다. `docs/SETUP.md`의 테스트 표를 수동으로 수행해야 합니다.

---

## 라이선스

MIT License - [`LICENSE`](LICENSE) 참고

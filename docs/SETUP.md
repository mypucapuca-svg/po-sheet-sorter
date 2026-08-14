# 설치 및 설정 가이드

> **코드 원본은 이 저장소의 `.bas` 파일들뿐입니다.**
> [`../ThisOutlookSession.bas`](../ThisOutlookSession.bas)(이벤트/오케스트레이션),
> [`../modConfig.bas`](../modConfig.bas)(설정), [`../modParse.bas`](../modParse.bas)(파싱/검증),
> [`../modIO.bas`](../modIO.bas)(파일·로그·원장·Word), [`../modTest.bas`](../modTest.bas)(테스트).
> 이 문서에는 코드를 중복해서 싣지 않습니다. 항상 저장소의 파일을 가져와 사용하세요.

---

## 1. 사전 준비

| 항목 | 확인 방법 |
|---|---|
| Windows용 Outlook 데스크톱 | 설치형 Outlook 실행 |
| Microsoft Word | 시작 메뉴에서 "Word" 검색해 실행되는지 확인 |
| 스타일#별 폴더 | `BASE_FOLDER` 아래 스타일#와 **정확히 같은 이름**의 폴더가 있어야 함 |

> Word는 PDF를 텍스트로 변환하는 데만 쓰입니다. Python이나 외부 프로그램은 필요 없습니다.

---

## 2. 작업 폴더 구조

매크로가 시작될 때 아래 폴더를 자동 생성합니다 (상위 폴더부터 순차 생성).
`BASE_FOLDER`만은 **자동 생성하지 않으며**, 없으면 매크로가 시작되지 않습니다.

```
C:\PO_SHEETS\                 ← BASE_FOLDER (직접 준비, 스타일#별 하위 폴더 포함)
C:\PO_WORK\
├── PO_Incoming\              ← 임시 저장 (7일 지난 파일은 시작 시 자동 정리)
├── PO_Review\                ← 확인필요
├── PO_Error\                 ← 오류
└── logs\
    ├── PO_Log.txt                        ← 처리 로그 (UTF-8, 5MB 넘으면 자동 로테이션)
    ├── PO_Log.txt.fallback.YYYYMMDD.txt   ← 로그 쓰기 자체가 실패했을 때만 생성되는 폴백
    └── PO_Ledger.csv                     ← 처리 원장(중복 방지용, UTF-8)
```

---

## 3. 매크로 등록

이 프로젝트는 5개의 `.bas` 파일로 나뉘어 있습니다. `ThisOutlookSession`은 Outlook 프로젝트에
이미 존재하는 특수 모듈이라 **붙여넣기**로, 나머지 4개는 **파일 가져오기**로 등록합니다.

1. Outlook에서 `Alt + F11` → VBA 편집기 실행
2. 왼쪽 트리에서 프로젝트(보통 `VbaProject.OTM`)를 우클릭 → `파일 가져오기`
3. [`modConfig.bas`](../modConfig.bas), [`modParse.bas`](../modParse.bas), [`modIO.bas`](../modIO.bas),
   [`modTest.bas`](../modTest.bas) 4개 파일을 **각각** 가져오기 (모듈 이름은 파일에 이미 지정되어 있어
   자동으로 잡힙니다)
4. 왼쪽 트리에서 **`ThisOutlookSession`** 더블클릭
5. [`ThisOutlookSession.bas`](../ThisOutlookSession.bas) 파일을 열어 **전체 내용을 복사**해 붙여넣기
6. 아래 [4. 설정값 수정](#4-설정값-수정)을 진행
7. 저장 (디스켓 아이콘)

---

## 4. 설정값 수정

설정 상수는 모두 **`modConfig.bas`**에 있습니다. 그 상수만 본인 환경에 맞게 수정하면 되고,
`ThisOutlookSession.bas`나 다른 모듈은 건드릴 필요 없습니다.

### 반드시 확인해야 하는 값

| 상수 | 설명 |
|---|---|
| `BASE_FOLDER` | 스타일#별 폴더들이 들어있는 상위 경로. 끝에 `\` 필수 |
| `SUBJECT_KEYWORD` | 처리 대상 메일을 고르는 제목 키워드 |
| `ALLOWED_SENDERS` | 허용 발신자. 세미콜론(`;`)으로 여러 개. `@도메인` 또는 전체 주소 |

### 발신자 검증 (`ALLOWED_SENDERS`)

**초기값 `@example.com`은 반드시 실제 바이어 도메인으로 바꿔야 합니다.** 그렇지 않으면
모든 메일이 "허용되지 않은 발신자"로 건너뛰어집니다.

```vb
' 도메인 전체 허용
Public Const ALLOWED_SENDERS As String = "@buyer.example.com"

' 여러 개 지정
Public Const ALLOWED_SENDERS As String = "@buyer.example.com;@agent.example.com;contact@example.com"
```

검증을 끄고 싶다면 `REQUIRE_SENDER_CHECK = False`로 두되, 이 경우 제목과 확장자만 맞으면
외부 메일의 PDF도 Word로 열리므로 권장하지 않습니다.

> 더 안전한 대안: Outlook **규칙**으로 신뢰하는 발신자의 메일만 전용 폴더로 옮기고,
> 그 폴더만 감시하도록 `BindInboxEvents`(`ThisOutlookSession.bas`)를 수정하는 방법입니다.
> VBA가 지는 검증 책임이 줄어듭니다.

### 기타 설정

| 상수 | 기본값 | 설명 |
|---|---|---|
| `MAX_PDF_SIZE_MB` | 20 | 이보다 큰 첨부는 오류 처리 |
| `MAX_ATTACHMENTS` | 10 | 첨부가 이보다 많으면 메일 전체를 확인필요 |
| `SHARED_MAILBOX` | (비움) | 공유 사서함 주소. 비우면 기본 받은편지함 |
| `RECONCILE_DAYS` | 14 | 시작 시 재검색할 기간 |
| `MAX_RECONCILE_ITEMS` | 200 | 재검색 1회 최대 처리 건수 |
| `COUNTRY_REQUIRED` | True | 국가 판별 실패를 확인필요로 볼지 여부 |
| `LOG_MAX_BYTES` | 5MB | 이 크기를 넘으면 로그 파일을 자동 로테이션 |
| `TEMP_FILE_MAX_AGE_DAYS` | 7 | 이보다 오래된 임시 파일을 시작 시 정리 |
| `WORD_MAX_USE_COUNT` | 20 | 이 건수마다 Word 인스턴스를 재시작(누수 방지) |

---

## 5. 매크로 보안 설정

1. Outlook → 파일 → 옵션 → 보안 센터 → 보안 센터 설정 → 매크로 설정
2. **"모든 매크로에 대해 알림 표시"** 선택
3. Outlook 완전히 종료 후 재시작
4. 매크로 경고가 뜨면 **"매크로 사용"** 클릭

매번 뜨는 창이 번거로우면 `SelfCert.exe`(Office 동봉)로 자체 서명 인증서를 만들어
매크로에 서명하고, 해당 인증서를 "신뢰할 수 있는 게시자"에 등록하면 됩니다.

---

## 6. 첫 실행 확인

Outlook 재시작 시:

- 시작에 성공하면 `PO_Log.txt`에 `[INFO] ... 매크로 시작됨` 이 기록됩니다.
- 폴더 생성이나 쓰기 권한에 문제가 있으면 **경고창이 뜨고 매크로가 시작되지 않습니다.**
  경고창이 떴다면 설정의 경로와 권한을 확인하세요.
- 오래된 `WINWORD.EXE` 프로세스가 남아 있으면 안내성 로그가 한 줄 남습니다(자동 종료하지 않음).
- 이어서 최근 `RECONCILE_DAYS`일치 **아직 한 번도 분류되지 않은** 메일을 자동으로 재검색합니다.
  이미 `PO_Review`/`PO_Error`로 분류된 메일은 자동 재검색 대상이 아닙니다 → [8. 수동 재검색](#8-수동-재검색-및-강제-재처리) 참고.

---

## 7. 테스트

### 자동 테스트 (Outlook/Word 불필요)

`modParse.bas`의 순수 파싱 로직은 회귀 테스트로 즉시 검증할 수 있습니다.

1. VBA 편집기에서 `modTest` 모듈을 열고, `RunAllTests` 프로시저 안에 커서를 둠
2. `F5` 실행
3. `Ctrl+G`(직접 실행 창)에서 `ALL PASS` 확인. 실패가 있으면 `FAIL ...` 줄에 원인이 출력됩니다.

아래 표의 6개 케이스 중 4개(정상 케이스, 스캔 이미지 PDF, 없는 스타일#, 값 다중/납기일 오류)는
이 자동 테스트가 이미 커버합니다. 나머지 2개는 Outlook/Word 연동이 필요해 수동으로 확인하세요.

### 정상 케이스 (수동, Outlook 실사용)

제목에 `PO SHEET`가 들어가고 정상 PO SHEET PDF가 첨부된 메일을 **허용된 발신자 주소로**
받아보세요. (본인에게 보내 테스트하려면 본인 주소를 `ALLOWED_SENDERS`에 임시로 추가)

결과: 스타일# 폴더에 `ORDER.D010125T000000 (AUS PO#10000000 250115).pdf` 생성

### 이상 케이스 (권장 테스트)

| 테스트 | 준비물 | 기대 결과 |
|---|---|---|
| 스캔 이미지 PDF | 이미지만 있는 PDF | `PO_Error` 또는 `PO_Review`, 로그에 사유 기록 |
| 확장자만 PDF | txt 파일 이름을 `.pdf`로 변경 | `PO_Error`, "PDF 시그니처 없음" |
| 없는 스타일# | 폴더 없는 스타일의 PO | `PO_Review`, "스타일 폴더 없음" |
| 동일 내용 첨부 재수신 | 같은 PDF를 다른 메일로 재전달 | 로그에 `[SKIP] 이미 처리된 첨부(내용 동일)` |
| 첨부 2개 중 1개 손상 | 정상 PDF + 손상 PDF | 정상 1건은 완료, 손상 1건만 오류 |
| Outlook 껐다 켜기 | 꺼진 동안 PO 메일 수신 | 재시작 시 자동으로 찾아 처리 |

---

## 8. 수동 재검색 및 강제 재처리

- **`RunReconcileNow`** — Outlook을 오래 꺼두었거나 처리 누락이 의심될 때, VBA 편집기에서
  실행(F5)하면 즉시 재검색합니다. 단, `PO_Review`/`PO_Error`로 이미 분류된 메일은 대상이 아닙니다
  (재시작마다 무한 재처리되는 것을 막기 위한 설계입니다).
- **`ForceReprocessSelected`** — 스타일 폴더를 새로 만드는 등 `PO_Review`/`PO_Error` 사유를
  해소한 뒤, 특정 메일을 즉시 재처리하고 싶을 때 사용합니다. Outlook 메일 목록에서 재처리할
  메일을 선택한 상태로 VBA 편집기에서 실행(F5)하세요.

---

## 9. 문제 해결

| 증상 | 확인할 것 |
|---|---|
| 아무 반응이 없음 | 로그에 `매크로 시작됨`이 있는지 / 매크로 보안 허용했는지 |
| 전부 `[SKIP] 허용되지 않은 발신자` | `modConfig.bas`의 `ALLOWED_SENDERS`가 아직 `@example.com`인지 확인 |
| 시작 시 경고창 | `BASE_FOLDER`가 실제로 존재하는지, 작업 폴더 쓰기 권한이 있는지 |
| 계속 `PO_Review`로 감 | 로그의 사유 확인. 필수값 누락인지, 값이 여러 개인지, 국가가 모호한지 |
| 같은 메일인데 재처리됨 | 첨부 내용이 실제로 달라졌는지 확인(SHA-256 키는 내용이 1바이트라도 다르면 새 항목) |
| 처리가 느림 | PDF마다 Word를 여는 구조. 대량 처리 시 지연될 수 있음 (`WORD_MAX_USE_COUNT`로 인스턴스는 재사용됨) |
| 로그/원장 폴백 파일 생성됨 | `PO_Log.txt.fallback.*.txt` 존재 = 원본 파일이 다른 프로그램에 잠겨있었을 가능성. 파일을 닫고 필요하면 폴백 내용을 원본에 옮기세요 |
| `RunAllTests`에서 `FileSha256` 관련 실패는 없지만 실제 운영 원장 키가 `EntryID\|...` 형식만 보임 | 해당 PC에 `SHA256Managed` COM이 등록되지 않은 것 — 기능은 정상 동작하지만 [알려진 한계](../README.md#알려진-한계) 참고 |
| WINWORD.EXE가 계속 쌓임 | 시작 시 안내 로그가 남는지 확인. 작업 관리자에서 숨겨진 WINWORD.EXE를 정리하고, 반복되면 이슈로 남겨주세요 |

---

## 10. 운영 시 주의

- **여러 팀원이 동시에 실행하지 마세요.** 지정 PC 한 대에서만 실행하는 것을 권장합니다.
  원장(`PO_Ledger.csv`)이 여러 곳에서 동시에 갱신되면 중복 방지가 깨질 수 있습니다.
- 원장 키가 첨부 내용의 SHA-256 해시이므로, 처리된 메일을 다른 폴더/사서함으로 옮기거나
  첨부 순서가 바뀌어도 중복 방지는 유지됩니다(해시 계산이 가능한 환경 기준).
- `PO_Review`와 `PO_Error` 폴더는 주기적으로 비워주세요. 폴더의 파일 자체는 자동 재처리되지
  않지만, 원인을 해소한 뒤에는 `ForceReprocessSelected`로 명시적으로 재처리해야 합니다.

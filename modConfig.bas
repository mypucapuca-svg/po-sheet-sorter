Attribute VB_Name = "modConfig"
Option Explicit

'==============================================================================
'  PO Sheet Sorter - 설정 값
'  본인 환경에 맞게 이 모듈의 상수만 수정하세요. 코드는 건드릴 필요 없습니다.
'==============================================================================

Public Const BASE_FOLDER    As String = "C:\PO_SHEETS\"              ' 스타일#별 폴더 상위 경로
Public Const TEMP_FOLDER    As String = "C:\PO_WORK\PO_Incoming\"    ' 임시 저장
Public Const REVIEW_FOLDER  As String = "C:\PO_WORK\PO_Review\"      ' 확인필요
Public Const ERROR_FOLDER   As String = "C:\PO_WORK\PO_Error\"       ' 오류
Public Const LOG_FILE       As String = "C:\PO_WORK\logs\PO_Log.txt"
Public Const LEDGER_FILE    As String = "C:\PO_WORK\logs\PO_Ledger.csv"

Public Const SUBJECT_KEYWORD As String = "PO SHEET"   ' 메일 제목 키워드

' 발신자 검증: 세미콜론(;)으로 구분. "@도메인" 또는 전체 이메일 주소 모두 가능.
Public Const REQUIRE_SENDER_CHECK As Boolean = True
Public Const ALLOWED_SENDERS As String = "@buyer.example.com"

Public Const MAX_PDF_SIZE_MB  As Long = 20    ' 첨부 크기 상한
Public Const MAX_ATTACHMENTS  As Long = 10    ' 메일당 첨부 개수 상한

' 공유 사서함을 감시하려면 주소 입력. 비워두면 기본 받은편지함.
' 주의: 여러 팀원이 동시에 실행하지 말고 지정 PC 한 대에서만 실행할 것.
Public Const SHARED_MAILBOX As String = ""

Public Const RECONCILE_DAYS      As Long = 14   ' 시작 시 재검색할 기간(일)
Public Const MAX_RECONCILE_ITEMS As Long = 200  ' 재검색 1회 최대 처리 건수

Public Const COUNTRY_REQUIRED As Boolean = True ' 국가 판별 실패를 확인필요로 볼지

Public Const LOG_MAX_BYTES          As Long = 5 * 1024 * 1024  ' 로그 파일 로테이션 임계값(바이트)
Public Const TEMP_FILE_MAX_AGE_DAYS As Long = 7                ' 임시 폴더 정리 기준(일)
Public Const WORD_MAX_USE_COUNT     As Long = 20                ' Word 인스턴스 재시작 주기(건)

'==============================================================================

' Outlook 메일에 기록할 사용자 속성 이름
Public Const PROP_STATUS  As String = "PO_SORT_STATUS"
Public Const PROP_UPDATED As String = "PO_SORT_UPDATED_AT"
Public Const PROP_RESULT  As String = "PO_SORT_RESULT"

' 처리 상태값
Public Const ST_COMPLETED As String = "COMPLETED"
Public Const ST_REVIEW    As String = "REVIEW"
Public Const ST_ERROR     As String = "ERROR"

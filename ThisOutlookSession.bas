Option Explicit

'==============================================================================
'  PO Sheet Sorter  v3
'  Outlook VBA 편집기(Alt+F11) > ThisOutlookSession 에 전체 붙여넣기
'
'  이 파일은 Outlook 이벤트 바인딩과 메일/첨부 단위 오케스트레이션만 담당합니다.
'  설정값은 modConfig.bas, 순수 파싱/검증 로직은 modParse.bas, 파일·로그·원장·Word
'  변환은 modIO.bas에 있습니다. 네 파일 모두 같은 VBA 프로젝트에 있어야 합니다.
'  (docs/SETUP.md 참고)
'
'  v3 대비 변경점 (GitHub 이슈 #1~#11 반영)
'   #1  재검색 날짜 필터를 로케일 무관 DASL 쿼리로 교체 (한국어 등에서 무동작 방지)
'   #2  최종 이동을 Name 문 대신 FileCopy+Kill로 교체 (다른 볼륨/네트워크 공유 대응)
'   #3  mBusy로 건너뛴 메일을 로그에 남기고, 다음 재검색이 자동으로 찾아 처리하도록 함
'   #4  REVIEW/ERROR 메일이 재시작마다 무한 재처리되던 문제 수정 (미분류 메일만 재검색)
'   #5  로그/원장 기록 실패 시 모달 대화상자로 파이프라인을 막지 않고 폴백 파일 + 카운터 사용
'   #6  재검색 컬렉션 순회 중 mail.Save로 항목이 건너뛰어지던 문제 수정 (역순 인덱스)
'   #7  Word 인스턴스를 재사용하여 프로세스 생성 비용과 고아 프로세스 누적을 줄임
'   #8  원장 키를 첨부 내용의 SHA-256 해시로 교체 (메일 이동/첨부순번 변경에 안전)
'   #9  존재하지 않는 날짜(2/30 등)를 유효로 판정하던 버그 수정
'   #10 스타일# 경로 sanitize, 국가 유일성 검증, 로그/원장 UTF-8 기록, 로그 로테이션,
'       임시 폴더 정리, 폴더 쓰기 테스트 파일명 충돌 등 견고성 개선 묶음
'   #11 순수 파싱 로직을 modParse.bas로 분리해 Outlook 없이 자동 회귀 테스트 가능(modTest.bas)
'
'  v1->v2 변경점은 git 이력을 참고하세요.
'==============================================================================


Private WithEvents InboxItems As Outlook.Items
Private mInitOk As Boolean
Private mBusy As Boolean


'==============================================================================
'  시작 / 이벤트 바인딩
'==============================================================================

Private Sub Application_Startup()
    mInitOk = False

    If Not InitializeFolders() Then
        MsgBox "PO Sheet Sorter 시작 실패: 작업 폴더를 만들 수 없거나 쓰기 권한이 없습니다." & vbCrLf & _
               "설정의 폴더 경로를 확인하세요.", vbExclamation, "PO Sheet Sorter"
        Exit Sub
    End If

    If Not BindInboxEvents() Then
        LogMessage ST_ERROR, "", "", "", "", "", "받은편지함 연결 실패"
        MsgBox "PO Sheet Sorter 시작 실패: 감시할 받은편지함을 열 수 없습니다.", _
               vbExclamation, "PO Sheet Sorter"
        Exit Sub
    End If

    mInitOk = True
    ResetLogFailureCount
    ResetLedgerFailureCount
    LogMessage "INFO", "", "", "", "", "", "매크로 시작됨"

    CleanupOldTempFiles
    LogOrphanWordProcesses

    ReconcilePendingMail RECONCILE_DAYS
End Sub


Private Function BindInboxEvents() As Boolean
    Dim olNS As Outlook.NameSpace
    Dim rcp As Outlook.Recipient

    On Error GoTo Fail
    Set olNS = Application.GetNamespace("MAPI")

    If Len(Trim$(SHARED_MAILBOX)) > 0 Then
        Set rcp = olNS.CreateRecipient(SHARED_MAILBOX)
        rcp.Resolve
        If Not rcp.Resolved Then
            LogMessage ST_ERROR, "", "", "", "", "", "공유 사서함을 찾을 수 없음: " & SHARED_MAILBOX
            BindInboxEvents = False
            Exit Function
        End If
        Set InboxItems = olNS.GetSharedDefaultFolder(rcp, olFolderInbox).Items
    Else
        Set InboxItems = olNS.GetDefaultFolder(olFolderInbox).Items
    End If

    BindInboxEvents = True
    Exit Function
Fail:
    BindInboxEvents = False
End Function


' #3: 처리 중 도착한 메일을 조용히 버리지 않고 로그로 남깁니다. 이 메일은 아직
' 상태가 분류되지 않았으므로(GetMailStatus = "") 다음 ReconcilePendingMail(재시작
' 또는 RunReconcileNow)이 자동으로 다시 찾아 처리합니다.
Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrHandler
    If Not mInitOk Then Exit Sub

    If TypeOf Item Is Outlook.MailItem Then
        If mBusy Then
            LogMessage "SKIP", SafeEntryId(Item), SafeSender(Item), SafeSubject(Item), "", "", _
                "다른 메일 처리 중이라 건너뜀 - 다음 재검색 시 자동 처리됨"
            Exit Sub
        End If

        mBusy = True
        ProcessMail Item
        mBusy = False
    End If
    Exit Sub

ErrHandler:
    mBusy = False
    LogMessage ST_ERROR, "", "", "", "", "", "ItemAdd 처리 중 오류: " & Err.Number & " - " & Err.Description
End Sub


' 수동 실행용 (VBA 편집기에서 F5, 또는 리본 버튼에 연결)
Public Sub RunReconcileNow()
    If Not mInitOk Then
        If Not InitializeFolders() Then Exit Sub
        If Not BindInboxEvents() Then Exit Sub
        mInitOk = True
    End If

    ResetLogFailureCount
    ResetLedgerFailureCount
    ReconcilePendingMail RECONCILE_DAYS

    Dim msg As String
    msg = "미처리 메일 재검색을 완료했습니다. 로그를 확인하세요." & vbCrLf & LOG_FILE

    If GetLogFailureCount() > 0 Then
        msg = msg & vbCrLf & vbCrLf & "경고: 로그 기록 실패 " & GetLogFailureCount() & "건 - " & _
              LOG_FILE & ".fallback.*.txt 파일을 확인하세요."
    End If
    If GetLedgerFailureCount() > 0 Then
        msg = msg & vbCrLf & "경고: 원장 기록 실패 " & GetLedgerFailureCount() & "건 - 위 폴백 파일을 확인하세요."
    End If

    MsgBox msg, vbInformation, "PO Sheet Sorter"
End Sub


' #4: 사람이 REVIEW/ERROR 사유(스타일 폴더 생성 등)를 해소한 뒤, 자동 재검색을 기다리지
' 않고 즉시 다시 돌리고 싶을 때 사용합니다. Outlook 목록에서 재처리할 메일을 선택하고
' 실행하세요. (자동 재검색은 #4 수정으로 더 이상 REVIEW/ERROR 메일을 자동으로 건드리지
' 않으므로, 의도적인 재처리는 이렇게 명시적으로 트리거해야 합니다)
Public Sub ForceReprocessSelected()
    Dim sel As Object
    Dim itm As Object
    Dim n As Long

    If Not mInitOk Then
        If Not InitializeFolders() Then Exit Sub
        If Not BindInboxEvents() Then Exit Sub
        mInitOk = True
    End If

    On Error Resume Next
    Set sel = Application.ActiveExplorer.Selection
    On Error GoTo 0

    If sel Is Nothing Then
        MsgBox "재처리할 메일을 Outlook 목록에서 먼저 선택하세요.", vbExclamation, "PO Sheet Sorter"
        Exit Sub
    End If

    For Each itm In sel
        If TypeOf itm Is Outlook.MailItem Then
            mBusy = True
            ProcessMail itm
            mBusy = False
            n = n + 1
        End If
    Next itm

    MsgBox n & "건을 재처리했습니다. 로그를 확인하세요.", vbInformation, "PO Sheet Sorter"
End Sub


'==============================================================================
'  H3. 누락 메일 재검색
'==============================================================================

Private Sub ReconcilePendingMail(ByVal daysBack As Long)
    Dim filtered As Outlook.Items
    Dim itm As Object
    Dim filterStr As String
    Dim processed As Long
    Dim scanned As Long
    Dim i As Long

    On Error GoTo ErrHandler
    If InboxItems Is Nothing Then Exit Sub

    ' #1: 기존 Format(Now, "ddddd h:nn AMPM")은 시스템 로케일을 타서 한국어 Windows 등에서
    ' Outlook의 Restrict 파서가 이해하지 못하는 문자열("오후" 등)을 만들어 재검색이 항상
    ' 실패했습니다. urn:schemas 기반 DASL 쿼리는 로케일과 무관합니다.
    filterStr = "@SQL=" & Chr(34) & "urn:schemas:httpmail:datereceived" & Chr(34) & " >= '" & _
                Format$(Now - daysBack, "yyyy\/mm\/dd hh:nn") & "'"

    Set filtered = InboxItems.Restrict(filterStr)
    filtered.Sort "[ReceivedTime]", True

    ' #6: 순회 중 ProcessMail -> MarkMail이 mail.Save를 호출해 Restrict 컬렉션(라이브 뷰)이
    ' 재평가되면서 For Each 커서가 어긋나 일부 메일이 조용히 건너뛰어졌습니다.
    ' 역순 인덱스 루프는 아직 방문하지 않은 낮은 인덱스가 영향받지 않아 안전합니다.
    For i = filtered.Count To 1 Step -1
        Set itm = filtered.Item(i)
        scanned = scanned + 1
        If processed >= MAX_RECONCILE_ITEMS Then Exit For

        If TypeOf itm Is Outlook.MailItem Then
            ' #4: COMPLETED뿐 아니라 REVIEW/ERROR로 이미 분류된 메일도 재검색 대상에서
            ' 제외해야 무한 재처리(파일 _01, _02... 무한 증식)가 발생하지 않습니다.
            ' 아직 한 번도 분류되지 않은 메일만 후보로 삼습니다.
            If Len(GetMailStatus(itm)) = 0 Then
                If IsTargetMail(itm) Then
                    mBusy = True
                    ProcessMail itm
                    mBusy = False
                    processed = processed + 1
                End If
            End If
        End If
    Next i

    LogMessage "INFO", "", "", "", "", "", _
        "재검색 완료: 검사 " & scanned & "건, 처리 " & processed & "건 (최근 " & daysBack & "일)"
    Exit Sub

ErrHandler:
    mBusy = False
    LogMessage ST_ERROR, "", "", "", "", "", "재검색 중 오류: " & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'  메일 단위 처리
'==============================================================================

Private Sub ProcessMail(ByVal mail As Outlook.MailItem)
    Dim att As Outlook.Attachment
    Dim entryId As String
    Dim pdfCount As Long

    On Error GoTo ErrHandler

    If Not IsTargetMail(mail) Then Exit Sub

    entryId = SafeEntryId(mail)

    If mail.Attachments.Count > MAX_ATTACHMENTS Then
        LogMessage ST_REVIEW, entryId, SafeSender(mail), mail.Subject, "", "", _
            "첨부 개수 초과(" & mail.Attachments.Count & " > " & MAX_ATTACHMENTS & ") - 수동 확인 필요"
        MarkMail mail, ST_REVIEW, "첨부 개수 초과"
        Exit Sub
    End If

    ' M4. 첨부 하나의 실패가 다른 첨부를 막지 않도록 첨부 단위로 분리 호출
    For Each att In mail.Attachments
        If IsPdfAttachment(att) Then
            pdfCount = pdfCount + 1
            ProcessAttachment mail, att, entryId
        End If
    Next att

    If pdfCount = 0 Then
        LogMessage ST_REVIEW, entryId, SafeSender(mail), mail.Subject, "", "", _
            "제목은 일치하나 PDF 첨부가 없음"
        MarkMail mail, ST_REVIEW, "PDF 첨부 없음"
    End If
    Exit Sub

ErrHandler:
    LogMessage ST_ERROR, entryId, SafeSender(mail), SafeSubject(mail), "", "", _
        "메일 처리 오류: " & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'  첨부 단위 처리 (M4: 자체 오류 처리기 보유)
'==============================================================================

Private Sub ProcessAttachment(ByVal mail As Outlook.MailItem, _
                              ByVal att As Outlook.Attachment, _
                              ByVal entryId As String)
    Dim tempPath As String
    Dim ledgerKey As String
    Dim pdfText As String
    Dim convertError As String
    Dim validateError As String
    Dim po As PoData
    Dim destFolder As String
    Dim destPath As String
    Dim sender As String
    Dim subj As String
    Dim contentHash As String

    On Error GoTo ErrHandler

    sender = SafeSender(mail)
    subj = SafeSubject(mail)
    ' 내용 해시 계산 전까지 쓰는 폴백 키. 해시 계산에 성공하면 아래에서 교체됩니다.
    ledgerKey = entryId & "|" & CStr(att.Index) & "|" & att.FileName

    ' ---- M3. 충돌하지 않는 임시 파일명 --------------------------------------
    tempPath = BuildUniquePath(TEMP_FOLDER, _
                   Format(Now, "yyyymmdd_hhnnss") & "_" & ShortHash(entryId) & _
                   "_" & Format(att.Index, "00") & "_" & SanitizeFileName(BaseName(att.FileName)), _
                   ".pdf")

    att.SaveAsFile tempPath

    ' ---- H2. 첨부 검증 (Word로 열기 전에) -----------------------------------
    If Not ValidateAttachmentFile(tempPath, validateError) Then
        MoveTo ERROR_FOLDER, tempPath, destPath
        LogMessage ST_ERROR, entryId, sender, subj, att.FileName, destPath, validateError
        LedgerWrite ledgerKey, entryId, att.Index, att.FileName, ST_ERROR, destPath, validateError
        MarkMail mail, ST_ERROR, validateError
        Exit Sub
    End If

    ' ---- #8. 내용 해시 기반 중복 처리 방지 -----------------------------------
    ' 메일 EntryID(다른 폴더로 이동 시 변경됨)나 첨부 순번/파일명 대신 파일 내용의
    ' SHA-256을 키로 쓰면 같은 PDF가 어느 메일로 오든, 몇 번째 첨부든 동일하게 잡힙니다.
    contentHash = FileSha256(tempPath)
    If Len(contentHash) > 0 Then ledgerKey = "sha256:" & contentHash

    If LedgerHasEntry(ledgerKey) Then
        LogMessage "SKIP", entryId, sender, subj, att.FileName, "", "이미 처리된 첨부(내용 동일) - 건너뜀"
        On Error Resume Next
        Kill tempPath
        On Error GoTo ErrHandler
        Exit Sub
    End If

    ' ---- M1. Word 변환 실패와 미매칭 구분 -----------------------------------
    If Not TryGetPdfTextViaWord(tempPath, pdfText, convertError) Then
        MoveTo ERROR_FOLDER, tempPath, destPath
        LogMessage ST_ERROR, entryId, sender, subj, att.FileName, destPath, convertError
        LedgerWrite ledgerKey, entryId, att.Index, att.FileName, ST_ERROR, destPath, convertError
        MarkMail mail, ST_ERROR, convertError
        Exit Sub
    End If

    If Len(Trim$(pdfText)) = 0 Then
        MoveTo REVIEW_FOLDER, tempPath, destPath
        LogMessage ST_REVIEW, entryId, sender, subj, att.FileName, destPath, _
            "Word 변환은 성공했으나 추출된 텍스트가 없음(스캔 이미지 PDF 가능성)"
        LedgerWrite ledgerKey, entryId, att.Index, att.FileName, ST_REVIEW, destPath, "텍스트 없음"
        MarkMail mail, ST_REVIEW, "텍스트 없음"
        Exit Sub
    End If

    ' ---- 필드 추출 및 검증 (H1, M2, #10.2 국가 유일성) -----------------------
    If Not ParseAndValidate(pdfText, po, validateError) Then
        MoveTo REVIEW_FOLDER, tempPath, destPath
        LogMessage ST_REVIEW, entryId, sender, subj, att.FileName, destPath, validateError
        LedgerWrite ledgerKey, entryId, att.Index, att.FileName, ST_REVIEW, destPath, validateError
        MarkMail mail, ST_REVIEW, validateError
        Exit Sub
    End If

    ' D코드는 PDF가 아니라 원본 첨부파일명에서 추출 (필수값)
    Dim dCode As String
    dCode = ExtractDCode(att.FileName)
    If Len(dCode) = 0 Then
        MoveTo REVIEW_FOLDER, tempPath, destPath
        LogMessage ST_REVIEW, entryId, sender, subj, att.FileName, destPath, _
            "필수값 누락: D코드(첨부파일명에서 D######T###### 형식을 찾지 못함)"
        LedgerWrite ledgerKey, entryId, att.Index, att.FileName, ST_REVIEW, destPath, "D코드 누락"
        MarkMail mail, ST_REVIEW, "D코드 누락"
        Exit Sub
    End If

    ' ---- 목적지 확인 (#10.1: 스타일#도 경로 조합 전 sanitize) -----------------
    destFolder = BASE_FOLDER & SanitizeFileName(po.StyleNo) & "\"
    If Dir(destFolder, vbDirectory) = "" Then
        MoveTo REVIEW_FOLDER, tempPath, destPath
        LogMessage ST_REVIEW, entryId, sender, subj, att.FileName, destPath, _
            "스타일 폴더가 존재하지 않음: " & destFolder
        LedgerWrite ledgerKey, entryId, att.Index, att.FileName, ST_REVIEW, destPath, "스타일 폴더 없음"
        MarkMail mail, ST_REVIEW, "스타일 폴더 없음"
        Exit Sub
    End If

    ' ---- 커밋: 모든 검증 통과 후에만 최종 이동 (H1) --------------------------
    Dim finalBase As String
    finalBase = "ORDER." & dCode & " (" & po.Country & " PO#" & po.OrderNo & " " & po.DlvDate & ")"
    destPath = BuildUniquePath(destFolder, SanitizeFileName(finalBase), ".pdf")

    ' #2: Name 문은 같은 볼륨 내 이동만 지원해 BASE_FOLDER가 TEMP_FOLDER와 다른 드라이브/
    ' 네트워크 공유면 오류 75로 실패했습니다. FileCopy는 볼륨 경계를 넘습니다.
    FileCopy tempPath, destPath
    If Dir(destPath) = "" Then Err.Raise vbObjectError + 513, , "대상 파일 생성 실패: " & destPath
    Kill tempPath

    LogMessage ST_COMPLETED, entryId, sender, subj, att.FileName, destPath, _
        "STYLE=" & po.StyleNo & " ORDER=" & po.OrderNo & " DLV=" & po.DlvDate & " CTRY=" & po.Country
    LedgerWrite ledgerKey, entryId, att.Index, att.FileName, ST_COMPLETED, destPath, ""
    MarkMail mail, ST_COMPLETED, destPath
    Exit Sub

ErrHandler:
    ' 첨부 단위 격리: 이 첨부만 실패로 남기고 호출부는 다음 첨부를 계속 처리
    Dim em As String
    em = "첨부 처리 오류: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Len(tempPath) > 0 Then
        If Dir(tempPath) <> "" Then MoveTo ERROR_FOLDER, tempPath, destPath
    End If
    LogMessage ST_ERROR, entryId, sender, subj, att.FileName, destPath, em
    LedgerWrite ledgerKey, entryId, att.Index, att.FileName, ST_ERROR, destPath, em
    MarkMail mail, ST_ERROR, em
End Sub


'==============================================================================
'  H2. 대상 메일 / 첨부 검증
'==============================================================================

Private Function IsTargetMail(ByVal mail As Outlook.MailItem) As Boolean
    Dim subj As String

    On Error GoTo Fail
    IsTargetMail = False

    subj = SafeSubject(mail)
    If InStr(1, subj, SUBJECT_KEYWORD, vbTextCompare) = 0 Then Exit Function

    If REQUIRE_SENDER_CHECK Then
        If Not IsAllowedSender(SafeSender(mail)) Then
            LogMessage "SKIP", SafeEntryId(mail), SafeSender(mail), subj, "", "", _
                "허용되지 않은 발신자 - 처리하지 않음"
            Exit Function
        End If
    End If

    IsTargetMail = True
    Exit Function
Fail:
    IsTargetMail = False
End Function


Private Function IsPdfAttachment(ByVal att As Outlook.Attachment) As Boolean
    On Error GoTo Fail
    IsPdfAttachment = (LCase$(Right$(att.FileName, 4)) = ".pdf")
    Exit Function
Fail:
    IsPdfAttachment = False
End Function


'==============================================================================
'  메일 속성 / 안전 접근자
'==============================================================================

Private Sub MarkMail(ByVal mail As Outlook.MailItem, _
                     ByVal status As String, _
                     ByVal result As String)
    Dim up As Outlook.UserProperty

    On Error GoTo Done

    ' 이미 COMPLETED인 메일을 REVIEW/ERROR로 내리지 않음 (첨부 여러 개인 경우)
    If GetMailStatus(mail) = ST_COMPLETED And status <> ST_COMPLETED Then GoTo Done

    Set up = mail.UserProperties.Find(PROP_STATUS)
    If up Is Nothing Then Set up = mail.UserProperties.Add(PROP_STATUS, olText)
    up.Value = status

    Set up = mail.UserProperties.Find(PROP_UPDATED)
    If up Is Nothing Then Set up = mail.UserProperties.Add(PROP_UPDATED, olText)
    up.Value = Format(Now, "yyyy-mm-dd hh:nn:ss")

    Set up = mail.UserProperties.Find(PROP_RESULT)
    If up Is Nothing Then Set up = mail.UserProperties.Add(PROP_RESULT, olText)
    up.Value = Left$(result, 250)

    mail.Save
Done:
End Sub


Private Function GetMailStatus(ByVal mail As Outlook.MailItem) As String
    Dim up As Outlook.UserProperty

    GetMailStatus = ""
    On Error GoTo Done

    Set up = mail.UserProperties.Find(PROP_STATUS)
    If Not up Is Nothing Then GetMailStatus = CStr(up.Value)
Done:
End Function


Private Function SafeEntryId(ByVal mail As Outlook.MailItem) As String
    On Error GoTo Fail
    SafeEntryId = mail.entryId
    Exit Function
Fail:
    SafeEntryId = ""
End Function


Private Function SafeSubject(ByVal mail As Outlook.MailItem) As String
    On Error GoTo Fail
    SafeSubject = CStr(mail.Subject)
    Exit Function
Fail:
    SafeSubject = ""
End Function


' Exchange 내부 발신자는 SenderEmailAddress가 SMTP가 아니므로 PropertyAccessor 우선
Private Function SafeSender(ByVal mail As Outlook.MailItem) As String
    Const PR_SENDER_SMTP As String = _
        "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
    Dim addr As String

    On Error Resume Next

    addr = mail.PropertyAccessor.GetProperty(PR_SENDER_SMTP)

    If Len(addr) = 0 Then
        If LCase$(mail.SenderEmailType) = "ex" Then
            addr = mail.Sender.GetExchangeUser.PrimarySmtpAddress
        End If
    End If

    If Len(addr) = 0 Then addr = mail.SenderEmailAddress

    SafeSender = LCase$(Trim$(addr))
End Function

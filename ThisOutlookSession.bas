Option Explicit

'==============================================================================
'  PO Sheet Sorter  v2
'  Outlook VBA 편집기(Alt+F11) > ThisOutlookSession 에 전체 붙여넣기
'
'  v1 대비 변경점 (코드리뷰 반영)
'   H1 필수값 누락 시 완료 처리하지 않고 확인필요(Review)로 분리
'   H2 발신자 / 첨부 크기 / PDF 시그니처 검증 후에만 Word 변환
'   H3 시작 시 최근 N일 미처리 메일 재검색(Reconcile)
'   H4 원장(CSV)에 메일EntryID+첨부순번 기록하여 중복 처리 방지
'   M1 Word 변환 실패와 스타일 미매칭을 별도 상태로 구분
'   M2 오더번호/납기일이 PDF 내에서 유일하지 않으면 확인필요로 분리
'   M3 임시파일명에 메일해시+첨부순번+충돌카운터 사용
'   M4 첨부 단위 오류 격리
'   M5 로그 경로 설정화 + 로그/폴더 접근 실패 감지
'==============================================================================


'============================== 설정 ==========================================
' 아래 값만 본인 환경에 맞게 수정하세요.

Private Const BASE_FOLDER    As String = "C:\PO_SHEETS\"              ' 스타일#별 폴더 상위 경로
Private Const TEMP_FOLDER    As String = "C:\PO_WORK\PO_Incoming\"    ' 임시 저장
Private Const REVIEW_FOLDER  As String = "C:\PO_WORK\PO_Review\"      ' 확인필요
Private Const ERROR_FOLDER   As String = "C:\PO_WORK\PO_Error\"       ' 오류
Private Const LOG_FILE       As String = "C:\PO_WORK\logs\PO_Log.txt"
Private Const LEDGER_FILE    As String = "C:\PO_WORK\logs\PO_Ledger.csv"

Private Const SUBJECT_KEYWORD As String = "PO SHEET"   ' 메일 제목 키워드

' 발신자 검증: 세미콜론(;)으로 구분. "@도메인" 또는 전체 이메일 주소 모두 가능.
Private Const REQUIRE_SENDER_CHECK As Boolean = True
Private Const ALLOWED_SENDERS As String = "@buyer.example.com"

Private Const MAX_PDF_SIZE_MB  As Long = 20    ' 첨부 크기 상한
Private Const MAX_ATTACHMENTS  As Long = 10    ' 메일당 첨부 개수 상한

' 공유 사서함을 감시하려면 주소 입력. 비워두면 기본 받은편지함.
' 주의: 여러 팀원이 동시에 실행하지 말고 지정 PC 한 대에서만 실행할 것.
Private Const SHARED_MAILBOX As String = ""

Private Const RECONCILE_DAYS      As Long = 14   ' 시작 시 재검색할 기간(일)
Private Const MAX_RECONCILE_ITEMS As Long = 200  ' 재검색 1회 최대 처리 건수

Private Const COUNTRY_REQUIRED As Boolean = True ' 국가 판별 실패를 확인필요로 볼지
'==============================================================================


' Outlook 메일에 기록할 사용자 속성 이름
Private Const PROP_STATUS  As String = "PO_SORT_STATUS"
Private Const PROP_UPDATED As String = "PO_SORT_UPDATED_AT"
Private Const PROP_RESULT  As String = "PO_SORT_RESULT"

' 처리 상태값
Private Const ST_COMPLETED As String = "COMPLETED"
Private Const ST_REVIEW    As String = "REVIEW"
Private Const ST_ERROR     As String = "ERROR"

Private Type PoData
    StyleNo As String
    OrderNo As String
    DlvDate As String
    Country As String
End Type

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
    LogMessage "INFO", "", "", "", "", "", "매크로 시작됨"

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


Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrHandler
    If Not mInitOk Then Exit Sub
    If mBusy Then Exit Sub

    If TypeOf Item Is Outlook.MailItem Then
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
    ReconcilePendingMail RECONCILE_DAYS
    MsgBox "미처리 메일 재검색을 완료했습니다. 로그를 확인하세요." & vbCrLf & LOG_FILE, _
           vbInformation, "PO Sheet Sorter"
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

    On Error GoTo ErrHandler
    If InboxItems Is Nothing Then Exit Sub

    filterStr = "[ReceivedTime] >= '" & _
                Format(Now - daysBack, "ddddd h:nn AMPM") & "'"

    Set filtered = InboxItems.Restrict(filterStr)
    filtered.Sort "[ReceivedTime]", True

    For Each itm In filtered
        scanned = scanned + 1
        If processed >= MAX_RECONCILE_ITEMS Then Exit For

        If TypeOf itm Is Outlook.MailItem Then
            If GetMailStatus(itm) <> ST_COMPLETED Then
                If IsTargetMail(itm) Then
                    mBusy = True
                    ProcessMail itm
                    mBusy = False
                    processed = processed + 1
                End If
            End If
        End If
    Next itm

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

    On Error GoTo ErrHandler

    sender = SafeSender(mail)
    subj = SafeSubject(mail)

    ' ---- H4. 중복 처리 방지 -------------------------------------------------
    ledgerKey = entryId & "|" & CStr(att.Index) & "|" & att.FileName
    If LedgerHasCompleted(ledgerKey) Then
        LogMessage "SKIP", entryId, sender, subj, att.FileName, "", "이미 처리 완료된 첨부 - 건너뜀"
        Exit Sub
    End If

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

    ' ---- 필드 추출 및 검증 (H1, M2) -----------------------------------------
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

    ' ---- 목적지 확인 --------------------------------------------------------
    destFolder = BASE_FOLDER & po.StyleNo & "\"
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

    Name tempPath As destPath

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


Private Function IsAllowedSender(ByVal smtpAddress As String) As Boolean
    Dim parts() As String
    Dim i As Long
    Dim rule As String
    Dim addr As String

    IsAllowedSender = False
    addr = LCase$(Trim$(smtpAddress))
    If Len(addr) = 0 Then Exit Function

    parts = Split(ALLOWED_SENDERS, ";")
    For i = LBound(parts) To UBound(parts)
        rule = LCase$(Trim$(parts(i)))
        If Len(rule) > 0 Then
            If Left$(rule, 1) = "@" Then
                ' 도메인 규칙: 주소가 해당 도메인으로 끝나야 함
                If Len(addr) > Len(rule) Then
                    If Right$(addr, Len(rule)) = rule Then
                        IsAllowedSender = True
                        Exit Function
                    End If
                End If
            Else
                If addr = rule Then
                    IsAllowedSender = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function


Private Function IsPdfAttachment(ByVal att As Outlook.Attachment) As Boolean
    On Error GoTo Fail
    IsPdfAttachment = (LCase$(Right$(att.FileName, 4)) = ".pdf")
    Exit Function
Fail:
    IsPdfAttachment = False
End Function


Private Function ValidateAttachmentFile(ByVal filePath As String, _
                                        ByRef errorMessage As String) As Boolean
    Dim sizeBytes As Double

    On Error GoTo Fail
    ValidateAttachmentFile = False
    errorMessage = ""

    If Dir(filePath) = "" Then
        errorMessage = "첨부 저장 실패: 파일이 존재하지 않음"
        Exit Function
    End If

    sizeBytes = FileLen(filePath)

    If sizeBytes = 0 Then
        errorMessage = "첨부 검증 실패: 빈 파일(0바이트)"
        Exit Function
    End If

    If sizeBytes > CDbl(MAX_PDF_SIZE_MB) * 1024# * 1024# Then
        errorMessage = "첨부 검증 실패: 크기 초과(" & Format(sizeBytes / 1048576#, "0.0") & _
                       "MB > " & MAX_PDF_SIZE_MB & "MB)"
        Exit Function
    End If

    If Not HasPdfSignature(filePath) Then
        errorMessage = "첨부 검증 실패: PDF 시그니처(%PDF-) 없음 - 확장자만 PDF인 파일일 수 있음"
        Exit Function
    End If

    ValidateAttachmentFile = True
    Exit Function
Fail:
    errorMessage = "첨부 검증 중 오류: " & Err.Number & " - " & Err.Description
    ValidateAttachmentFile = False
End Function


Private Function HasPdfSignature(ByVal filePath As String) As Boolean
    Dim f As Integer
    Dim buf As String

    On Error GoTo Fail
    f = FreeFile
    Open filePath For Binary Access Read As #f
    buf = Space$(5)
    Get #f, 1, buf
    Close #f

    HasPdfSignature = (buf = "%PDF-")
    Exit Function
Fail:
    On Error Resume Next
    Close #f
    HasPdfSignature = False
End Function


'==============================================================================
'  M1. Word 변환 (성공여부 / 텍스트 / 오류를 분리 반환)
'==============================================================================

Private Function TryGetPdfTextViaWord(ByVal pdfPath As String, _
                                      ByRef pdfText As String, _
                                      ByRef errorMessage As String) As Boolean
    Dim wdApp As Object
    Dim wdDoc As Object

    pdfText = ""
    errorMessage = ""

    On Error GoTo Fail

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    wdApp.DisplayAlerts = False

    Set wdDoc = wdApp.Documents.Open( _
        FileName:=pdfPath, _
        ConfirmConversions:=False, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        PasswordDocument:="__no_password__")

    pdfText = wdDoc.Content.Text
    TryGetPdfTextViaWord = True

Cleanup:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Exit Function

Fail:
    errorMessage = "Word PDF 변환 실패: " & Err.Number & " - " & Err.Description
    TryGetPdfTextViaWord = False
    Resume Cleanup
End Function


'==============================================================================
'  필드 추출 및 검증 (H1, M2)
'==============================================================================

Private Function ParseAndValidate(ByVal pdfText As String, _
                                  ByRef po As PoData, _
                                  ByRef errorMessage As String) As Boolean
    Dim styles As Collection
    Dim orders As Collection
    Dim dates_ As Collection

    ParseAndValidate = False
    errorMessage = ""

    po.StyleNo = "": po.OrderNo = "": po.DlvDate = "": po.Country = ""

    ' 스타일#  : "STYLE PACK: 90000001 12AB34CDE567 ..." 의 두 번째 코드
    Set styles = MatchDistinct(pdfText, "STYLE\s*PACK\s*:\s*\d+\s+([A-Za-z0-9]+)")
    ' 오더번호  : "ORDER NUMBER : 10000000"  (페이지마다 반복되지만 값은 같아야 함)
    Set orders = MatchDistinct(pdfText, "ORDER NUMBER\s*:\s*(\d+)")
    ' 납기일    : "DLV CONS DATE : 15/01/25" -> 250115
    Set dates_ = MatchDistinctDlvDate(pdfText)

    ' --- 필수값 존재 검증 (H1) ---
    Dim missing As String
    missing = ""
    If styles.Count = 0 Then missing = missing & "스타일#,"
    If orders.Count = 0 Then missing = missing & "오더번호,"
    If dates_.Count = 0 Then missing = missing & "납기일,"

    If Len(missing) > 0 Then
        errorMessage = "필수값 누락: " & Left$(missing, Len(missing) - 1)
        Exit Function
    End If

    ' --- 유일성 검증 (M2: 서로 다른 PO 구간의 값이 섞이는 것 방지) ---
    Dim ambiguous As String
    ambiguous = ""
    If styles.Count > 1 Then ambiguous = ambiguous & "스타일#(" & JoinCollection(styles) & ") "
    If orders.Count > 1 Then ambiguous = ambiguous & "오더번호(" & JoinCollection(orders) & ") "
    If dates_.Count > 1 Then ambiguous = ambiguous & "납기일(" & JoinCollection(dates_) & ") "

    If Len(ambiguous) > 0 Then
        errorMessage = "값이 여러 개라 자동 판별 불가 - 수동 확인 필요: " & Trim$(ambiguous)
        Exit Function
    End If

    po.StyleNo = styles(1)
    po.OrderNo = orders(1)
    po.DlvDate = dates_(1)
    po.Country = ExtractCountry(pdfText)

    If Len(po.Country) = 0 Then
        If COUNTRY_REQUIRED Then
            errorMessage = "국가코드 판별 실패(SYDNEY/MELBOURNE/BRISBANE/PERTH 미발견)"
            Exit Function
        Else
            po.Country = "XXX"
        End If
    End If

    ' 날짜 형식 자체 검증
    If Not IsValidYymmdd(po.DlvDate) Then
        errorMessage = "납기일 형식이 올바르지 않음: " & po.DlvDate
        Exit Function
    End If

    ParseAndValidate = True
End Function


Private Function MatchDistinct(ByVal text As String, ByVal pattern As String) As Collection
    Dim re As Object, mc As Object, m As Object
    Dim result As Collection
    Dim v As String

    Set result = New Collection

    On Error GoTo Done
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True
    re.pattern = pattern

    Set mc = re.Execute(text)
    For Each m In mc
        v = Trim$(m.SubMatches(0))
        If Len(v) > 0 Then
            If Not InCollection(result, v) Then result.Add v
        End If
    Next m

Done:
    Set MatchDistinct = result
End Function


Private Function MatchDistinctDlvDate(ByVal text As String) As Collection
    Dim re As Object, mc As Object, m As Object
    Dim result As Collection
    Dim v As String

    Set result = New Collection

    On Error GoTo Done
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.pattern = "DLV CONS DATE\s*:\s*(\d{2})/(\d{2})/(\d{2})"

    Set mc = re.Execute(text)
    For Each m In mc
        ' dd/mm/yy -> yymmdd
        v = m.SubMatches(2) & m.SubMatches(1) & m.SubMatches(0)
        If Not InCollection(result, v) Then result.Add v
    Next m

Done:
    Set MatchDistinctDlvDate = result
End Function


Private Function ExtractCountry(ByVal text As String) As String
    ' 호주 4개 도시 중 하나라도 있으면 AUS, 오클랜드가 있으면 NZ
    ' 다른 국가가 추가되면 이 함수에 규칙을 추가하세요.
    Dim u As String
    u = UCase$(text)

    If InStr(u, "SYDNEY") > 0 Or InStr(u, "MELBOURNE") > 0 Or _
       InStr(u, "BRISBANE") > 0 Or InStr(u, "PERTH") > 0 Then
        ExtractCountry = "AUS"
    ElseIf InStr(u, "AUCKLAND") > 0 Then
        ExtractCountry = "NZ"
    Else
        ExtractCountry = ""
    End If
End Function


Private Function ExtractDCode(ByVal fileName As String) As String
    Dim re As Object, mc As Object

    ExtractDCode = ""
    On Error GoTo Done

    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "D\d{6}T\d{6}"
    re.IgnoreCase = True
    re.Global = True

    Set mc = re.Execute(fileName)
    If mc.Count = 1 Then
        ExtractDCode = mc(0).Value
    ElseIf mc.Count > 1 Then
        ' 여러 개면 판별 불가 -> 빈 값 반환하여 확인필요로 보냄
        ExtractDCode = ""
    End If
Done:
End Function


Private Function IsValidYymmdd(ByVal s As String) As Boolean
    Dim yy As Integer, mm As Integer, dd As Integer

    IsValidYymmdd = False
    If Len(s) <> 6 Then Exit Function
    If Not IsNumeric(s) Then Exit Function

    yy = CInt(Left$(s, 2))
    mm = CInt(Mid$(s, 3, 2))
    dd = CInt(Right$(s, 2))

    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function

    On Error GoTo Fail
    Dim d As Date
    d = DateSerial(2000 + yy, mm, dd)
    IsValidYymmdd = True
    Exit Function
Fail:
    IsValidYymmdd = False
End Function


'==============================================================================
'  파일/폴더 유틸
'==============================================================================

Private Function InitializeFolders() As Boolean
    On Error GoTo Fail
    InitializeFolders = False

    If Not EnsureFolder(TEMP_FOLDER) Then Exit Function
    If Not EnsureFolder(REVIEW_FOLDER) Then Exit Function
    If Not EnsureFolder(ERROR_FOLDER) Then Exit Function
    If Not EnsureFolder(GetParentPath(LOG_FILE)) Then Exit Function
    If Not EnsureFolder(GetParentPath(LEDGER_FILE)) Then Exit Function

    If Dir(BASE_FOLDER, vbDirectory) = "" Then Exit Function
    If Not CanWriteTo(TEMP_FOLDER) Then Exit Function
    If Not CanWriteTo(GetParentPath(LOG_FILE)) Then Exit Function

    InitializeFolders = True
    Exit Function
Fail:
    InitializeFolders = False
End Function


' 상위 폴더부터 순차 생성 (한 단계 MkDir에 의존하지 않음)
Private Function EnsureFolder(ByVal folderPath As String) As Boolean
    Dim parts() As String
    Dim i As Long
    Dim cur As String

    On Error GoTo Fail
    EnsureFolder = False

    If Len(Trim$(folderPath)) = 0 Then Exit Function
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)

    parts = Split(folderPath, "\")
    cur = parts(LBound(parts))          ' "C:" 또는 UNC 앞부분

    If Left$(folderPath, 2) = "\\" Then
        ' UNC 경로: \\server\share 까지는 그대로 사용
        cur = "\\" & parts(2) & "\" & parts(3)
        For i = 4 To UBound(parts)
            cur = cur & "\" & parts(i)
            If Dir(cur, vbDirectory) = "" Then MkDir cur
        Next i
    Else
        For i = LBound(parts) + 1 To UBound(parts)
            cur = cur & "\" & parts(i)
            If Dir(cur, vbDirectory) = "" Then MkDir cur
        Next i
    End If

    EnsureFolder = (Dir(cur, vbDirectory) <> "")
    Exit Function
Fail:
    EnsureFolder = False
End Function


Private Function CanWriteTo(ByVal folderPath As String) As Boolean
    Dim testPath As String
    Dim f As Integer

    On Error GoTo Fail
    CanWriteTo = False

    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    testPath = folderPath & "__write_test_" & Format(Now, "hhnnss") & ".tmp"

    f = FreeFile
    Open testPath For Output As #f
    Print #f, "test"
    Close #f

    Kill testPath
    CanWriteTo = True
    Exit Function
Fail:
    On Error Resume Next
    Close #f
    CanWriteTo = False
End Function


' M3. 존재하지 않는 이름을 찾을 때까지 카운터 증가
Private Function BuildUniquePath(ByVal folderPath As String, _
                                 ByVal baseName As String, _
                                 ByVal extension As String) As String
    Dim candidate As String
    Dim counter As Long

    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    candidate = folderPath & baseName & extension
    counter = 1

    Do While Dir(candidate) <> ""
        candidate = folderPath & baseName & "_" & Format(counter, "00") & extension
        counter = counter + 1
        If counter > 999 Then
            candidate = folderPath & baseName & "_" & Format(Now, "hhnnss") & extension
            Exit Do
        End If
    Loop

    BuildUniquePath = candidate
End Function


Private Sub MoveTo(ByVal targetFolder As String, _
                   ByVal sourcePath As String, _
                   ByRef resultPath As String)
    On Error GoTo Fail
    resultPath = ""
    If Dir(sourcePath) = "" Then Exit Sub

    resultPath = BuildUniquePath(targetFolder, BaseName(FileNameOnly(sourcePath)), ".pdf")
    Name sourcePath As resultPath
    Exit Sub
Fail:
    resultPath = sourcePath   ' 이동 실패 시 원래 위치를 기록
End Sub


' Windows 금지문자 제거 + 예약어 회피 + 길이 제한
Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant
    Dim i As Long
    Dim out As String
    Dim upperOut As String

    out = s
    bad = Array("<", ">", ":", """", "/", "\", "|", "?", "*")
    For i = LBound(bad) To UBound(bad)
        out = Replace(out, bad(i), "_")
    Next i

    ' 제어문자 제거
    For i = 0 To 31
        out = Replace(out, Chr$(i), "")
    Next i

    out = Trim$(out)
    Do While Right$(out, 1) = "." Or Right$(out, 1) = " "
        out = Left$(out, Len(out) - 1)
        If Len(out) = 0 Then Exit Do
    Loop

    If Len(out) = 0 Then out = "UNNAMED"

    ' Windows 예약 이름 회피
    upperOut = UCase$(out)
    Select Case upperOut
        Case "CON", "PRN", "AUX", "NUL", _
             "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
             "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            out = "_" & out
    End Select

    If Len(out) > 120 Then out = Left$(out, 120)

    SanitizeFileName = out
End Function


Private Function BaseName(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 1 Then
        BaseName = Left$(fileName, p - 1)
    Else
        BaseName = fileName
    End If
End Function


Private Function FileNameOnly(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, "\")
    If p > 0 Then
        FileNameOnly = Mid$(fullPath, p + 1)
    Else
        FileNameOnly = fullPath
    End If
End Function


Private Function GetParentPath(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, "\")
    If p > 0 Then
        GetParentPath = Left$(fullPath, p)
    Else
        GetParentPath = ""
    End If
End Function


Private Function ShortHash(ByVal s As String) As String
    Dim i As Long
    Dim h As Long

    h = 5381
    For i = 1 To Len(s)
        h = (h * 33 + Asc(Mid$(s, i, 1))) Mod 8388593&
    Next i
    ShortHash = Right$("000000" & Hex$(h), 6)
End Function


'==============================================================================
'  H4. 처리 원장 (CSV)
'==============================================================================

Private Function LedgerHasCompleted(ByVal ledgerKey As String) As Boolean
    Dim f As Integer
    Dim lineText As String
    Dim needle As String

    LedgerHasCompleted = False
    On Error GoTo Done

    If Dir(LEDGER_FILE) = "" Then Exit Function

    needle = """" & Replace(ledgerKey, """", """""") & ""","
    f = FreeFile
    Open LEDGER_FILE For Input As #f
    Do While Not EOF(f)
        Line Input #f, lineText
        If InStr(1, lineText, needle, vbTextCompare) > 0 Then
            If InStr(1, lineText, """" & ST_COMPLETED & """", vbTextCompare) > 0 Then
                LedgerHasCompleted = True
                Exit Do
            End If
        End If
    Loop
    Close #f
    Exit Function

Done:
    On Error Resume Next
    Close #f
    LedgerHasCompleted = False
End Function


Private Sub LedgerWrite(ByVal ledgerKey As String, _
                        ByVal entryId As String, _
                        ByVal attIndex As Long, _
                        ByVal attName As String, _
                        ByVal status As String, _
                        ByVal destination As String, _
                        ByVal errorText As String)
    Dim f As Integer
    Dim isNew As Boolean

    On Error GoTo Done

    isNew = (Dir(LEDGER_FILE) = "")

    f = FreeFile
    Open LEDGER_FILE For Append As #f
    If isNew Then
        Print #f, "Key,ProcessedAt,EntryID,AttachmentIndex,AttachmentName,Status,Destination,Error"
    End If
    Print #f, CsvField(ledgerKey) & "," & _
              CsvField(Format(Now, "yyyy-mm-dd hh:nn:ss")) & "," & _
              CsvField(entryId) & "," & _
              CsvField(CStr(attIndex)) & "," & _
              CsvField(attName) & "," & _
              CsvField(status) & "," & _
              CsvField(destination) & "," & _
              CsvField(errorText)
    Close #f
    Exit Sub

Done:
    On Error Resume Next
    Close #f
End Sub


Private Function CsvField(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    CsvField = """" & Replace(s, """", """""") & """"
End Function


'==============================================================================
'  M5. 로그 (경로 설정화 + 실패 감지)
'==============================================================================

Private Sub LogMessage(ByVal status As String, _
                       ByVal entryId As String, _
                       ByVal sender As String, _
                       ByVal subj As String, _
                       ByVal attName As String, _
                       ByVal destination As String, _
                       ByVal message As String)
    Dim f As Integer
    Dim lineText As String

    On Error GoTo Fail

    lineText = Format(Now, "yyyy-mm-dd hh:nn:ss") & _
           " [" & status & "]" & _
           " | mail=" & Left$(entryId, 12) & _
           " | from=" & sender & _
           " | subj=" & Left$(subj, 60) & _
           " | att=" & attName & _
           " | dest=" & destination & _
           " | " & message

    f = FreeFile
    Open LOG_FILE For Append As #f
    Print #f, lineText
    Close #f
    Exit Sub

Fail:
    ' 로그 자체가 실패하면 조용히 넘기지 않고 사용자에게 알림 (M5)
    On Error Resume Next
    Close #f
    Debug.Print "LOG WRITE FAILED: " & lineText
    MsgBox "PO Sheet Sorter: 로그 파일에 기록하지 못했습니다." & vbCrLf & _
           "경로: " & LOG_FILE & vbCrLf & _
           "내용: " & lineText, vbExclamation, "PO Sheet Sorter"
End Sub


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


Private Function InCollection(ByVal col As Collection, ByVal val As String) As Boolean
    Dim itm As Variant
    InCollection = False
    For Each itm In col
        If CStr(itm) = val Then
            InCollection = True
            Exit Function
        End If
    Next itm
End Function


Private Function JoinCollection(ByVal col As Collection) As String
    Dim i As Long
    Dim s As String
    For i = 1 To col.Count
        If Len(s) > 0 Then s = s & ","
        s = s & CStr(col(i))
    Next i
    JoinCollection = s
End Function

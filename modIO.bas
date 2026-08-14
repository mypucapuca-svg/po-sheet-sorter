Attribute VB_Name = "modIO"
Option Explicit

'==============================================================================
'  PO Sheet Sorter - 파일/폴더, 로그, 원장, Word 변환
'==============================================================================

Private mLogFailureCount As Long
Private mLedgerFailureCount As Long

Private mLedgerKeys As Object       ' Scripting.Dictionary: ledgerKey -> True
Private mLedgerLoaded As Boolean

Private mWdApp As Object            ' 재사용되는 숨은 Word 인스턴스 (#7)
Private mWdUseCount As Long


' ---- 폴더 초기화 -----------------------------------------------------------

Public Function InitializeFolders() As Boolean
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
Public Function EnsureFolder(ByVal folderPath As String) As Boolean
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


' #10.6: 초 단위 타임스탬프만 쓰면 같은 초에 연속 호출될 때 파일명이 충돌하므로
' Timer의 밀리초 성분을 섞어 유일성을 높임
Public Function CanWriteTo(ByVal folderPath As String) As Boolean
    Dim testPath As String
    Dim f As Integer

    On Error GoTo Fail
    CanWriteTo = False

    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    testPath = folderPath & "__write_test_" & Format(Now, "hhnnss") & "_" & _
               Right$("000" & CStr(Int(Timer * 1000) Mod 1000), 3) & ".tmp"

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
Public Function BuildUniquePath(ByVal folderPath As String, _
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


' #2: VBA의 Name 문은 같은 볼륨 내 이동만 지원합니다(오류 75). BASE_FOLDER/REVIEW_FOLDER 등이
' TEMP_FOLDER와 다른 드라이브·네트워크 공유일 수 있으므로 볼륨 경계에 안전한 FileCopy+Kill 사용.
Public Sub MoveTo(ByVal targetFolder As String, _
                   ByVal sourcePath As String, _
                   ByRef resultPath As String)
    On Error GoTo Fail
    resultPath = ""
    If Dir(sourcePath) = "" Then Exit Sub

    resultPath = BuildUniquePath(targetFolder, BaseName(FileNameOnly(sourcePath)), ".pdf")
    FileCopy sourcePath, resultPath
    If Dir(resultPath) = "" Then
        resultPath = sourcePath   ' 복사 실패 시 원래 위치를 기록
        Exit Sub
    End If
    Kill sourcePath
    Exit Sub
Fail:
    resultPath = sourcePath   ' 이동 실패 시 원래 위치를 기록
End Sub


' #10.5: TEMP_FOLDER에 남는 오래된 파일 정리 (이동 실패 잔재 등)
Public Sub CleanupOldTempFiles()
    Dim fileName As String
    Dim fullPath As String
    Dim cutoff As Date

    On Error Resume Next
    cutoff = Now - TEMP_FILE_MAX_AGE_DAYS

    fileName = Dir(TEMP_FOLDER & "*.pdf")
    Do While Len(fileName) > 0
        fullPath = TEMP_FOLDER & fileName
        If FileDateTime(fullPath) < cutoff Then Kill fullPath
        fileName = Dir()
    Loop
End Sub


' ---- 첨부 검증 --------------------------------------------------------------

Public Function ValidateAttachmentFile(ByVal filePath As String, _
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


Public Function HasPdfSignature(ByVal filePath As String) As Boolean
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


' #8: 첨부 파일 내용의 SHA-256. 원장 키를 EntryID/첨부순번/파일명 대신 이 값으로 쓰면
' 메일을 다른 폴더로 옮기거나 첨부 순서가 바뀌어도 중복 판정이 깨지지 않습니다.
' SHA256Managed는 .NET Framework의 COM 등록에 의존하므로 환경에 따라 실패할 수 있고,
' 그 경우 빈 문자열을 반환하여 호출부가 기존 키 형식으로 폴백하도록 합니다.
Public Function FileSha256(ByVal filePath As String) As String
    Dim stm As Object, sha As Object, bytes() As Byte
    Dim i As Long, s As String

    On Error GoTo Fail
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                       ' adTypeBinary
    stm.Open
    stm.LoadFromFile filePath
    bytes = stm.Read
    stm.Close

    Set sha = CreateObject("System.Security.Cryptography.SHA256Managed")
    bytes = sha.ComputeHash_2(bytes)

    For i = LBound(bytes) To UBound(bytes)
        s = s & Right$("0" & Hex$(bytes(i)), 2)
    Next i
    FileSha256 = s
    Exit Function
Fail:
    FileSha256 = ""                    ' 실패 시 호출부에서 기존 키로 폴백
End Function


'==============================================================================
'  #7. Word 변환 - 인스턴스 재사용 + 실패 시 강제 정리
'==============================================================================

Private Function GetWordApp() As Object
    If mWdApp Is Nothing Then
        Set mWdApp = CreateObject("Word.Application")
        mWdApp.Visible = False
        mWdApp.DisplayAlerts = False
        mWdUseCount = 0
    End If
    Set GetWordApp = mWdApp
End Function


Private Sub ReleaseWordApp(Optional ByVal forceQuit As Boolean = False)
    mWdUseCount = mWdUseCount + 1
    If forceQuit Or mWdUseCount >= WORD_MAX_USE_COUNT Then
        On Error Resume Next
        If Not mWdApp Is Nothing Then mWdApp.Quit
        Set mWdApp = Nothing
        mWdUseCount = 0
    End If
End Sub


Public Function TryGetPdfTextViaWord(ByVal pdfPath As String, _
                                      ByRef pdfText As String, _
                                      ByRef errorMessage As String) As Boolean
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim failed As Boolean

    pdfText = ""
    errorMessage = ""
    failed = False

    On Error GoTo Fail

    Set wdApp = GetWordApp()

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
    Set wdDoc = Nothing
    ReleaseWordApp forceQuit:=failed
    Exit Function

Fail:
    failed = True
    errorMessage = "Word PDF 변환 실패: " & Err.Number & " - " & Err.Description
    TryGetPdfTextViaWord = False
    Resume Cleanup
End Function


' 시작 시 숨겨진 WINWORD.EXE가 남아 있는지 확인만 하고 로그로 남깁니다.
' 사용자가 Word로 실제 작업 중일 수 있으므로 여기서 강제 종료는 하지 않습니다.
Public Sub LogOrphanWordProcesses()
    Dim svc As Object, procs As Object

    On Error Resume Next
    Set svc = GetObject("winmgmts:\\.\root\cimv2")
    If svc Is Nothing Then Exit Sub
    Set procs = svc.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='WINWORD.EXE'")
    If procs Is Nothing Then Exit Sub

    If procs.Count > 0 Then
        LogMessage "INFO", "", "", "", "", "", _
            "WINWORD.EXE " & procs.Count & "개 실행 중 - 이전 세션의 고아 프로세스일 수 있음(수동 확인 권장)"
    End If
End Sub


'==============================================================================
'  #5. 로그 - 실패해도 파이프라인을 막지 않도록 모달 MsgBox 대신 폴백 파일 + 카운터 사용
'==============================================================================

Public Function LogMessage(ByVal status As String, _
                            ByVal entryId As String, _
                            ByVal sender As String, _
                            ByVal subj As String, _
                            ByVal attName As String, _
                            ByVal destination As String, _
                            ByVal message As String) As Boolean
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

    RotateLogIfNeeded
    AppendUtf8Line LOG_FILE, lineText
    LogMessage = True
    Exit Function

Fail:
    ' 로그 자체가 실패해도 모달 대화상자로 이벤트 스레드를 막지 않음(#5).
    ' 대신 실패 횟수를 세고 폴백 파일에 남겨, RunReconcileNow 종료 시 한 번만 사용자에게 알림.
    mLogFailureCount = mLogFailureCount + 1
    Debug.Print "LOG WRITE FAILED: " & lineText
    WriteFallbackLog lineText
    LogMessage = False
End Function


Private Sub WriteFallbackLog(ByVal lineText As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open LOG_FILE & ".fallback." & Format$(Now, "yyyymmdd") & ".txt" For Append As #f
    Print #f, lineText
    Close #f
End Sub


' #10.4: 로그 파일이 일정 크기를 넘으면 타임스탬프를 붙여 옆으로 물러나게 함
Public Sub RotateLogIfNeeded()
    Dim dotPos As Long
    Dim rotatedName As String

    On Error Resume Next
    If Dir(LOG_FILE) = "" Then Exit Sub
    If FileLen(LOG_FILE) < LOG_MAX_BYTES Then Exit Sub

    dotPos = InStrRev(LOG_FILE, ".")
    If dotPos > 0 Then
        rotatedName = Left$(LOG_FILE, dotPos - 1) & "_" & Format$(Now, "yyyymmdd_hhnnss") & Mid$(LOG_FILE, dotPos)
    Else
        rotatedName = LOG_FILE & "_" & Format$(Now, "yyyymmdd_hhnnss")
    End If

    Name LOG_FILE As rotatedName
End Sub


Public Function GetLogFailureCount() As Long
    GetLogFailureCount = mLogFailureCount
End Function


Public Sub ResetLogFailureCount()
    mLogFailureCount = 0
End Sub


'==============================================================================
'  #8. 처리 원장 (CSV) - 내용 해시 키 + 존재 여부 캐시(Dictionary)
'==============================================================================

Private Sub EnsureLedgerCacheLoaded()
    Dim f As Integer
    Dim lineText As String
    Dim p As Long
    Dim keyPart As String

    If mLedgerLoaded Then Exit Sub
    Set mLedgerKeys = CreateObject("Scripting.Dictionary")
    mLedgerLoaded = True

    On Error GoTo Done
    If Dir(LEDGER_FILE) = "" Then Exit Sub

    f = FreeFile
    Open LEDGER_FILE For Input As #f
    Do While Not EOF(f)
        Line Input #f, lineText
        ' Key는 항상 CsvField로 감싸인 첫 번째 컬럼이고 그 값 자체에는 따옴표가 없으므로
        ' 전체 CSV를 파싱하지 않고 첫 필드만 잘라내도 안전합니다.
        If Left$(lineText, 1) = """" Then
            p = InStr(2, lineText, """,")
            If p > 2 Then
                keyPart = Mid$(lineText, 2, p - 2)
                If keyPart <> "Key" Then
                    If Not mLedgerKeys.Exists(keyPart) Then mLedgerKeys.Add keyPart, True
                End If
            End If
        End If
    Loop
    Close #f
    Exit Sub

Done:
    On Error Resume Next
    Close #f
End Sub


' #4: 이전에는 상태가 COMPLETED인 항목만 "처리됨"으로 봐서 REVIEW/ERROR로 남은 첨부가
' 재시작마다 무한 재처리됐습니다. 이제는 원장에 키가 존재하기만 하면(상태 무관) 처리 이력으로 봅니다.
Public Function LedgerHasEntry(ByVal ledgerKey As String) As Boolean
    EnsureLedgerCacheLoaded
    LedgerHasEntry = mLedgerKeys.Exists(ledgerKey)
End Function


Public Sub LedgerWrite(ByVal ledgerKey As String, _
                        ByVal entryId As String, _
                        ByVal attIndex As Long, _
                        ByVal attName As String, _
                        ByVal status As String, _
                        ByVal destination As String, _
                        ByVal errorText As String)
    Dim lineText As String

    On Error GoTo Fail

    lineText = CsvField(ledgerKey) & "," & _
              CsvField(Format(Now, "yyyy-mm-dd hh:nn:ss")) & "," & _
              CsvField(entryId) & "," & _
              CsvField(CStr(attIndex)) & "," & _
              CsvField(attName) & "," & _
              CsvField(status) & "," & _
              CsvField(destination) & "," & _
              CsvField(errorText)

    EnsureLedgerCacheLoaded

    If Dir(LEDGER_FILE) = "" Then
        AppendUtf8Line LEDGER_FILE, "Key,ProcessedAt,EntryID,AttachmentIndex,AttachmentName,Status,Destination,Error"
    End If
    AppendUtf8Line LEDGER_FILE, lineText

    If Not mLedgerKeys.Exists(ledgerKey) Then mLedgerKeys.Add ledgerKey, True
    Exit Sub

Fail:
    ' 원장 실패는 중복 방지가 깨지는 심각한 상황이므로, 로그와 동일하게 침묵하지 않고 카운트+폴백 기록
    mLedgerFailureCount = mLedgerFailureCount + 1
    On Error Resume Next
    WriteFallbackLog "[LEDGER WRITE FAILED] " & lineText
End Sub


Public Function GetLedgerFailureCount() As Long
    GetLedgerFailureCount = mLedgerFailureCount
End Function


Public Sub ResetLedgerFailureCount()
    mLedgerFailureCount = 0
End Sub


'==============================================================================
'  #10.3 UTF-8(BOM) 기록 - Print #f의 시스템 ANSI 코드페이지 대신 사용
'  Excel/다른 로케일 PC에서 열어도 한글이 깨지지 않도록, 그리고 매 호출마다
'  파일 전체를 다시 읽지 않도록 바이트를 이어쓰는 방식으로 구현합니다.
'==============================================================================

Private Sub AppendUtf8Line(ByVal filePath As String, ByVal lineText As String)
    Dim f As Integer
    Dim isNew As Boolean
    Dim bom(2) As Byte

    isNew = (Dir(filePath) = "")

    f = FreeFile
    Open filePath For Binary Access Write As #f
    If isNew Then
        bom(0) = &HEF: bom(1) = &HBB: bom(2) = &HBF
        Put #f, 1, bom
        Put #f, 4, Utf8Bytes(lineText & vbCrLf)
    Else
        Put #f, LOF(f) + 1, Utf8Bytes(lineText & vbCrLf)
    End If
    Close #f
End Sub


Private Function Utf8Bytes(ByVal s As String) As Byte()
    Dim stm As Object
    Dim b() As Byte

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                  ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText s
    stm.Position = 0
    stm.Type = 1                  ' adTypeBinary로 전환 - UTF-8 BOM(3바이트)이 앞에 붙어서 읽힘
    stm.Position = 3              ' BOM 건너뜀
    b = stm.Read
    stm.Close
    Utf8Bytes = b
End Function

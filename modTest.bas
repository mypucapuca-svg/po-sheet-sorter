Attribute VB_Name = "modTest"
Option Explicit

'==============================================================================
'  PO Sheet Sorter - modParse.bas 회귀 테스트
'
'  Outlook 재시작도, 실제 메일 발송도, PDF 준비도 필요 없습니다.
'  VBA 편집기에서 이 모듈을 열고 RunAllTests 안에 커서를 둔 뒤 F5로 실행하세요.
'  결과는 즉시 창(Ctrl+G)에 출력됩니다.
'==============================================================================

Public Sub RunAllTests()
    Dim failed As Long

    failed = 0
    failed = failed + Test_ParseAndValidate()
    failed = failed + Test_ExtractCountries()
    failed = failed + Test_ExtractDCode()
    failed = failed + Test_IsValidYymmdd()
    failed = failed + Test_SanitizeFileName()
    failed = failed + Test_IsAllowedSender()

    Debug.Print "----------------------------------------"
    If failed = 0 Then
        Debug.Print "ALL PASS"
    Else
        Debug.Print failed & " FAILED"
    End If
End Sub


Private Function Test_ParseAndValidate() As Long
    Dim po As PoData
    Dim e As String
    Dim txt As String
    Dim n As Long

    ' 정상 케이스 (AUS)
    txt = "ORDER NUMBER : 10000000" & vbCrLf & _
          "STYLE PACK: 90000001 12AB34CDE567 SAMPLE PRODUCT" & vbCrLf & _
          "DLV CONS DATE : 15/01/25" & vbCrLf & _
          "DELIVER TO: SYDNEY NSW"
    If Not ParseAndValidate(txt, po, e) Then
        n = n + 1: Debug.Print "FAIL Test_ParseAndValidate/normal: " & e
    End If
    If po.StyleNo <> "12AB34CDE567" Then n = n + 1: Debug.Print "FAIL Test_ParseAndValidate/style"
    If po.OrderNo <> "10000000" Then n = n + 1: Debug.Print "FAIL Test_ParseAndValidate/order"
    If po.DlvDate <> "250115" Then n = n + 1: Debug.Print "FAIL Test_ParseAndValidate/date"
    If po.Country <> "AUS" Then n = n + 1: Debug.Print "FAIL Test_ParseAndValidate/country"

    ' 정상 케이스 (NZ)
    txt = "ORDER NUMBER : 20000000" & vbCrLf & _
          "STYLE PACK: 90000002 99ZZ88YYY777 SAMPLE PRODUCT" & vbCrLf & _
          "DLV CONS DATE : 01/03/25" & vbCrLf & _
          "DELIVER TO: AUCKLAND"
    If Not ParseAndValidate(txt, po, e) Then
        n = n + 1: Debug.Print "FAIL Test_ParseAndValidate/nz: " & e
    End If
    If po.Country <> "NZ" Then n = n + 1: Debug.Print "FAIL Test_ParseAndValidate/nz country"

    ' 오더번호 값 다중 -> 거부되어야 함
    txt = "ORDER NUMBER : 10000000" & vbCrLf & _
          "STYLE PACK: 90000001 12AB34CDE567 SAMPLE PRODUCT" & vbCrLf & _
          "DLV CONS DATE : 15/01/25" & vbCrLf & _
          "DELIVER TO: SYDNEY NSW" & vbCrLf & _
          "ORDER NUMBER : 20000000"
    If ParseAndValidate(txt, po, e) Then
        n = n + 1: Debug.Print "FAIL Test_ParseAndValidate: 오더번호 다중을 통과시킴"
    End If

    ' 국가 다중(SYDNEY + AUCKLAND 동시 등장) -> 거부되어야 함
    txt = "ORDER NUMBER : 10000000" & vbCrLf & _
          "STYLE PACK: 90000001 12AB34CDE567 SAMPLE PRODUCT" & vbCrLf & _
          "DLV CONS DATE : 15/01/25" & vbCrLf & _
          "DELIVER TO: SYDNEY / AUCKLAND"
    If ParseAndValidate(txt, po, e) Then
        n = n + 1: Debug.Print "FAIL Test_ParseAndValidate: 국가 다중을 통과시킴"
    End If

    ' 필수값 누락 -> 거부되어야 함
    If ParseAndValidate("STYLE PACK: 90000001 12AB34CDE567", po, e) Then
        n = n + 1: Debug.Print "FAIL Test_ParseAndValidate: 필수값 누락을 통과시킴"
    End If

    Test_ParseAndValidate = n
End Function


Private Function Test_ExtractCountries() As Long
    Dim n As Long

    If ExtractCountries("SYDNEY").Count <> 1 Then n = n + 1: Debug.Print "FAIL Test_ExtractCountries/AUS"
    If ExtractCountries("AUCKLAND").Count <> 1 Then n = n + 1: Debug.Print "FAIL Test_ExtractCountries/NZ"
    If ExtractCountries("SYDNEY AUCKLAND").Count <> 2 Then n = n + 1: Debug.Print "FAIL Test_ExtractCountries/both"
    If ExtractCountries("NOTHING HERE").Count <> 0 Then n = n + 1: Debug.Print "FAIL Test_ExtractCountries/none"

    Test_ExtractCountries = n
End Function


Private Function Test_ExtractDCode() As Long
    Dim n As Long

    If ExtractDCode("D010125T000000.pdf") <> "D010125T000000" Then
        n = n + 1: Debug.Print "FAIL Test_ExtractDCode/normal"
    End If
    If ExtractDCode("noD코드있음.pdf") <> "" Then
        n = n + 1: Debug.Print "FAIL Test_ExtractDCode/missing"
    End If
    If ExtractDCode("D010125T000000_D020126T000001.pdf") <> "" Then
        n = n + 1: Debug.Print "FAIL Test_ExtractDCode/duplicate"
    End If

    Test_ExtractDCode = n
End Function


Private Function Test_IsValidYymmdd() As Long
    Dim n As Long

    If IsValidYymmdd("250115") <> True Then n = n + 1: Debug.Print "FAIL Test_IsValidYymmdd/normal"
    If IsValidYymmdd("240229") <> True Then n = n + 1: Debug.Print "FAIL Test_IsValidYymmdd/leap"     ' 2024 윤년
    If IsValidYymmdd("250229") <> False Then n = n + 1: Debug.Print "FAIL Test_IsValidYymmdd/non-leap" ' 2025 평년
    If IsValidYymmdd("250230") <> False Then n = n + 1: Debug.Print "FAIL Test_IsValidYymmdd/feb30"
    If IsValidYymmdd("250431") <> False Then n = n + 1: Debug.Print "FAIL Test_IsValidYymmdd/apr31"
    If IsValidYymmdd("251301") <> False Then n = n + 1: Debug.Print "FAIL Test_IsValidYymmdd/month13"
    If IsValidYymmdd("250100") <> False Then n = n + 1: Debug.Print "FAIL Test_IsValidYymmdd/day00"
    If IsValidYymmdd("25011") <> False Then n = n + 1: Debug.Print "FAIL Test_IsValidYymmdd/short"
    If IsValidYymmdd("2501155") <> False Then n = n + 1: Debug.Print "FAIL Test_IsValidYymmdd/long"
    If IsValidYymmdd("1e0115") <> False Then n = n + 1: Debug.Print "FAIL Test_IsValidYymmdd/notdigits"

    Test_IsValidYymmdd = n
End Function


Private Function Test_SanitizeFileName() As Long
    Dim n As Long

    If SanitizeFileName("A/B\C:D") <> "A_B_C_D" Then n = n + 1: Debug.Print "FAIL Test_SanitizeFileName/bad chars"
    If SanitizeFileName("CON") <> "_CON" Then n = n + 1: Debug.Print "FAIL Test_SanitizeFileName/reserved"
    If SanitizeFileName("  trailing.  ") <> "trailing" Then n = n + 1: Debug.Print "FAIL Test_SanitizeFileName/trim"
    If SanitizeFileName("") <> "UNNAMED" Then n = n + 1: Debug.Print "FAIL Test_SanitizeFileName/empty"

    Test_SanitizeFileName = n
End Function


Private Function Test_IsAllowedSender() As Long
    Dim n As Long

    ' ALLOWED_SENDERS 기본값은 modConfig.bas의 "@buyer.example.com" 기준
    If IsAllowedSender("someone@buyer.example.com") <> True Then
        n = n + 1: Debug.Print "FAIL Test_IsAllowedSender/domain match"
    End If
    If IsAllowedSender("someone@evil.com") <> False Then
        n = n + 1: Debug.Print "FAIL Test_IsAllowedSender/no match"
    End If
    If IsAllowedSender("") <> False Then
        n = n + 1: Debug.Print "FAIL Test_IsAllowedSender/empty"
    End If

    Test_IsAllowedSender = n
End Function

Attribute VB_Name = "modParse"
Option Explicit

'==============================================================================
'  PO Sheet Sorter - 순수 파싱/검증 로직
'
'  이 모듈의 모든 프로시저는 Outlook·Word·파일시스템에 의존하지 않는 순수 함수입니다.
'  Outlook 없이도 VBA 편집기에서 바로 호출/테스트할 수 있습니다. (modTest.bas 참고)
'==============================================================================

Public Type PoData
    StyleNo As String
    OrderNo As String
    DlvDate As String
    Country As String
End Type


' ---- 필드 추출 및 검증 (H1, M2) --------------------------------------------

Public Function ParseAndValidate(ByVal pdfText As String, _
                                  ByRef po As PoData, _
                                  ByRef errorMessage As String) As Boolean
    Dim styles As Collection
    Dim orders As Collection
    Dim dates_ As Collection
    Dim countries As Collection

    ParseAndValidate = False
    errorMessage = ""

    po.StyleNo = "": po.OrderNo = "": po.DlvDate = "": po.Country = ""

    ' 스타일#  : "STYLE PACK: 90000001 12AB34CDE567 ..." 의 두 번째 코드
    Set styles = MatchDistinct(pdfText, "STYLE\s*PACK\s*:\s*\d+\s+([A-Za-z0-9]+)")
    ' 오더번호  : "ORDER NUMBER : 10000000"  (페이지마다 반복되지만 값은 같아야 함)
    Set orders = MatchDistinct(pdfText, "ORDER NUMBER\s*:\s*(\d+)")
    ' 납기일    : "DLV CONS DATE : 15/01/25" -> 250115
    Set dates_ = MatchDistinctDlvDate(pdfText)
    ' 국가      : 배송지 도시명 (여러 국가 도시가 동시에 매칭되면 유일성 검증에서 걸림)
    Set countries = ExtractCountries(pdfText)

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
    If countries.Count > 1 Then ambiguous = ambiguous & "국가(" & JoinCollection(countries) & ") "

    If Len(ambiguous) > 0 Then
        errorMessage = "값이 여러 개라 자동 판별 불가 - 수동 확인 필요: " & Trim$(ambiguous)
        Exit Function
    End If

    po.StyleNo = styles(1)
    po.OrderNo = orders(1)
    po.DlvDate = dates_(1)

    If countries.Count = 1 Then
        po.Country = countries(1)
    Else
        If COUNTRY_REQUIRED Then
            errorMessage = "국가코드 판별 실패(SYDNEY/MELBOURNE/BRISBANE/PERTH/AUCKLAND 미발견)"
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


Public Function MatchDistinct(ByVal text As String, ByVal pattern As String) As Collection
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


Public Function MatchDistinctDlvDate(ByVal text As String) As Collection
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


' 배송지 도시명으로 국가 코드를 판별합니다. 문서 안에 서로 다른 국가의 도시명이
' 동시에 등장하면(청구지/본사 주소 등과 섞이는 경우) 매칭된 국가를 전부 반환하므로,
' 호출부(ParseAndValidate)에서 다른 필드와 동일하게 유일성 검증에 걸립니다.
' 다른 국가를 추가하려면 이 함수에 조건을 추가하세요. (AUCKLAND -> NZ는 이미 반영됨)
Public Function ExtractCountries(ByVal text As String) As Collection
    Dim result As Collection
    Dim u As String

    Set result = New Collection
    u = UCase$(text)

    If InStr(u, "SYDNEY") > 0 Or InStr(u, "MELBOURNE") > 0 Or _
       InStr(u, "BRISBANE") > 0 Or InStr(u, "PERTH") > 0 Then
        result.Add "AUS"
    End If

    If InStr(u, "AUCKLAND") > 0 Then
        result.Add "NZ"
    End If

    Set ExtractCountries = result
End Function


Public Function ExtractDCode(ByVal fileName As String) As String
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


' DateSerial은 범위를 벗어난 값을 오류 없이 다음 달로 정규화하므로(예: 2/30 -> 3/2),
' 정규화 결과가 입력과 일치하는지 역검증해야 존재하지 않는 날짜를 걸러낼 수 있습니다.
Public Function IsValidYymmdd(ByVal s As String) As Boolean
    Dim yy As Integer, mm As Integer, dd As Integer
    Dim d As Date
    Dim i As Long

    IsValidYymmdd = False
    If Len(s) <> 6 Then Exit Function

    For i = 1 To 6
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i

    yy = CInt(Left$(s, 2))
    mm = CInt(Mid$(s, 3, 2))
    dd = CInt(Right$(s, 2))

    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function

    On Error GoTo Fail
    d = DateSerial(2000 + yy, mm, dd)
    If Day(d) <> dd Or Month(d) <> mm Or Year(d) <> 2000 + yy Then Exit Function

    IsValidYymmdd = True
    Exit Function
Fail:
    IsValidYymmdd = False
End Function


' Windows 금지문자 제거 + 예약어 회피 + 길이 제한
Public Function SanitizeFileName(ByVal s As String) As String
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


Public Function BaseName(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 1 Then
        BaseName = Left$(fileName, p - 1)
    Else
        BaseName = fileName
    End If
End Function


Public Function FileNameOnly(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, "\")
    If p > 0 Then
        FileNameOnly = Mid$(fullPath, p + 1)
    Else
        FileNameOnly = fullPath
    End If
End Function


Public Function GetParentPath(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, "\")
    If p > 0 Then
        GetParentPath = Left$(fullPath, p)
    Else
        GetParentPath = ""
    End If
End Function


Public Function ShortHash(ByVal s As String) As String
    Dim i As Long
    Dim h As Long

    h = 5381
    For i = 1 To Len(s)
        h = (h * 33 + Asc(Mid$(s, i, 1))) Mod 8388593&
    Next i
    ShortHash = Right$("000000" & Hex$(h), 6)
End Function


Public Function CsvField(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    CsvField = """" & Replace(s, """", """""") & """"
End Function


Public Function IsAllowedSender(ByVal smtpAddress As String) As Boolean
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


Public Function JoinCollection(ByVal col As Collection) As String
    Dim i As Long
    Dim s As String
    For i = 1 To col.Count
        If Len(s) > 0 Then s = s & ","
        s = s & CStr(col(i))
    Next i
    JoinCollection = s
End Function

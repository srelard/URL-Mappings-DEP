Attribute VB_Name = "Module1"
Option Explicit

'=========================================================
' CONFIG (MUST match your sheet/table/column names exactly)
'=========================================================
Private Const SHEET_OW As String = "OW URL (Paste Here)"
Private Const SHEET_CAT As String = "Product Categories"
Private Const TABLE_CAT As String = "tblCat"
Private Const CAT_COL_NORM As String = "NormKey"
Private Const CAT_COL_CODE As String = "henkel code"
Private Const SHEET_CHECK As String = "error check DEP urls"

'=========================================================
' Helpers: progress (StatusBar)
'=========================================================
Private Sub SetProgress(ByVal msg As String)
    Application.StatusBar = msg
End Sub

Private Sub ClearProgress()
    Application.StatusBar = False
End Sub

'=========================================================
' Helpers: restore Excel state
'=========================================================
Private Sub ResetExcelState()
    On Error Resume Next
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error GoTo 0
End Sub

'=========================================================
' Robust getters (show friendly error if name mismatch)
'=========================================================
Private Function GetWs(ByVal wb As Workbook, ByVal wsName As String) As Worksheet
    On Error Resume Next
    Set GetWs = wb.Worksheets(wsName)
    On Error GoTo 0

    If GetWs Is Nothing Then
        Dim ws As Worksheet, msg As String
        msg = "Sheet not found: '" & wsName & "'" & vbCrLf & vbCrLf & _
              "Available sheets:" & vbCrLf
        For Each ws In wb.Worksheets
            msg = msg & " - " & ws.Name & vbCrLf
        Next ws
        MsgBox msg, vbCritical, "VBA - Missing sheet"
    End If
End Function

Private Function GetTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set GetTable = ws.ListObjects(tableName)
    On Error GoTo 0

    If GetTable Is Nothing Then
        Dim lo As ListObject, msg As String
        msg = "Table not found: '" & tableName & "' on sheet '" & ws.Name & "'" & vbCrLf & vbCrLf & _
              "Available tables on that sheet:" & vbCrLf
        For Each lo In ws.ListObjects
            msg = msg & " - " & lo.Name & vbCrLf
        Next lo
        MsgBox msg, vbCritical, "VBA - Missing table"
    End If
End Function

Private Function GetTableColumn(ByVal lo As ListObject, ByVal colName As String) As ListColumn
    On Error Resume Next
    Set GetTableColumn = lo.ListColumns(colName)
    On Error GoTo 0

    If GetTableColumn Is Nothing Then
        Dim lc As ListColumn, msg As String
        msg = "Column not found: '" & colName & "' in table '" & lo.Name & "'" & vbCrLf & vbCrLf & _
              "Available columns:" & vbCrLf
        For Each lc In lo.ListColumns
            msg = msg & " - " & lc.Name & vbCrLf
        Next lc
        MsgBox msg, vbCritical, "VBA - Missing column"
    End If
End Function

'=========================================================
' Helpers: string normalization + matching
'=========================================================
Private Function NormText(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    If Len(t) = 0 Then
        NormText = ""
        Exit Function
    End If

    t = Replace(t, "-", "")
    t = Replace(t, " ", "")
    t = Replace(t, ChrW(160), "")
    NormText = t
End Function

Private Function ContainsText(ByVal hay As String, ByVal needle As String) As Boolean
    ContainsText = (InStr(1, hay, needle, vbTextCompare) > 0)
End Function

Private Function StripQueryAndHash(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, "?", vbTextCompare)
    If p > 0 Then s = Left$(s, p - 1)

    p = InStr(1, s, "#", vbTextCompare)
    If p > 0 Then s = Left$(s, p - 1)

    StripQueryAndHash = s
End Function

Private Function RemoveTrailingSlash(ByVal s As String) As String
    Do While Len(s) > 0 And Right$(s, 1) = "/"
        s = Left$(s, Len(s) - 1)
    Loop
    RemoveTrailingSlash = s
End Function

Private Function GetAfter(ByVal s As String, ByVal token As String) As String
    Dim p As Long
    p = InStr(1, s, token, vbTextCompare)
    If p = 0 Then
        GetAfter = ""
    Else
        GetAfter = Mid$(s, p + Len(token))
    End If
End Function

Private Function RemoveExtensionHtml(ByVal s As String) As String
    If LCase$(Right$(s, 5)) = ".html" Then
        RemoveExtensionHtml = Left$(s, Len(s) - 5)
    Else
        RemoveExtensionHtml = s
    End If
End Function

'=========================================================
' Build lookup dictionary from tblCat: NormKey -> henkel code
'=========================================================
Private Function BuildCatDict() As Object
    On Error GoTo Fail

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = GetWs(wb, SHEET_CAT)
    If ws Is Nothing Then Exit Function

    Dim lo As ListObject: Set lo = GetTable(ws, TABLE_CAT)
    If lo Is Nothing Then Exit Function

    If lo.DataBodyRange Is Nothing Then
        MsgBox "Table '" & TABLE_CAT & "' on sheet '" & SHEET_CAT & "' has no data rows.", vbExclamation, "VBA - Empty table"
        Exit Function
    End If

    Dim colNorm As ListColumn, colCode As ListColumn
    Set colNorm = GetTableColumn(lo, CAT_COL_NORM)
    If colNorm Is Nothing Then Exit Function

    Set colCode = GetTableColumn(lo, CAT_COL_CODE)
    If colCode Is Nothing Then Exit Function

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 'TextCompare

    Dim normVals As Variant, codeVals As Variant
    normVals = colNorm.DataBodyRange.Value2
    codeVals = colCode.DataBodyRange.Value2

    Dim i As Long
    Dim k As String, v As String

    For i = 1 To UBound(normVals, 1)
        k = Trim$(CStr(normVals(i, 1)))
        If Len(k) > 0 Then
            v = Trim$(CStr(codeVals(i, 1)))
            If Not dict.Exists(k) Then
                dict.Add k, v
            End If
        End If
    Next i

    Set BuildCatDict = dict
    Exit Function

Fail:
    MsgBox "Error while building category dictionary: " & Err.Description, vbCritical, "VBA Error"
End Function

'=========================================================
' PART A: Generate New URLs into OW URL (Paste Here) col B
'=========================================================
Public Sub GenerateNewUrls_FromOldUrls()
    On Error GoTo CleanFail

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = GetWs(wb, SHEET_OW)
    If ws Is Nothing Then Exit Sub

    Dim dict As Object
    Set dict = BuildCatDict()
    If dict Is Nothing Then Exit Sub

    Dim row As Long: row = 2
    Dim processed As Long: processed = 0
    Dim oldUrl As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    SetProgress "Generating new URLs... (starting)"

    Do
        oldUrl = Trim$(CStr(ws.Cells(row, 1).Value2)) ' column A
        If Len(oldUrl) = 0 Then Exit Do

        oldUrl = StripQueryAndHash(oldUrl)
        oldUrl = RemoveTrailingSlash(oldUrl)

        ws.Cells(row, 2).Value2 = BuildNewUrlFromOld(oldUrl, dict) ' column B

        processed = processed + 1
        If (processed Mod 200) = 0 Then
            SetProgress "Generating new URLs... " & processed & " done (row " & row & ")"
            DoEvents
        End If

        row = row + 1
    Loop

    ResetExcelState
    MsgBox "Done. Generated " & processed & " new URLs into '" & SHEET_OW & "' column B.", vbInformation
    Exit Sub

CleanFail:
    ResetExcelState
    MsgBox "GenerateNewUrls_FromOldUrls failed: " & Err.Description & " (row " & row & ")", vbCritical, "VBA Error"
End Sub

Private Function BuildNewUrlFromOld(ByVal oldUrl As String, ByVal dict As Object) As String
    Dim afterCom As String
    afterCom = GetAfter(oldUrl, ".com/")

    If Len(afterCom) = 0 Then
        BuildNewUrlFromOld = ""
        Exit Function
    End If

    Dim parts As Variant
    parts = Split(afterCom, "/")

    Dim country As String, lang As String, section As String
    country = ""
    lang = ""
    section = ""

    If UBound(parts) >= 0 Then country = Trim$(CStr(parts(0)))
    If UBound(parts) >= 1 Then lang = Trim$(CStr(parts(1)))
    If UBound(parts) >= 2 Then section = Trim$(CStr(parts(2)))

    If Len(country) = 0 Or Len(lang) = 0 Then
        BuildNewUrlFromOld = ""
        Exit Function
    End If

    Dim base As String
    base = "https://next.henkel-adhesives.com/" & country & "/" & lang

    Dim prodGeneric As String
    prodGeneric = base & "/products.html/producttype_industrial-root-producttype.html"

    If Len(section) = 0 Then
        BuildNewUrlFromOld = base & ".html"
        Exit Function
    End If

    If ContainsText(section, "applications") Then
        BuildNewUrlFromOld = base & "/applications.html"
    ElseIf ContainsText(section, "industries") Then
        BuildNewUrlFromOld = base & "/industries.html"
    ElseIf ContainsText(section, "insights") Then
        BuildNewUrlFromOld = base & "/knowledge.html"
    ElseIf ContainsText(section, "search") Then
        BuildNewUrlFromOld = prodGeneric
    ElseIf ContainsText(section, "services") Then
        BuildNewUrlFromOld = base & "/support.html"
    ElseIf ContainsText(section, "spotlights") Then
        BuildNewUrlFromOld = base & "/knowledge.html"
    ElseIf ContainsText(section, "about") Then
        BuildNewUrlFromOld = base & ".html"
    ElseIf ContainsText(section, "product") Then
        BuildNewUrlFromOld = BuildProductUrl(oldUrl, base, prodGeneric, dict)
    Else
        BuildNewUrlFromOld = base & ".html"
    End If
End Function

Private Function BuildProductUrl(ByVal oldUrl As String, ByVal base As String, ByVal prodGeneric As String, ByVal dict As Object) As String
    Dim afterCom As String
    afterCom = GetAfter(oldUrl, ".com/")
    If Len(afterCom) = 0 Then
        BuildProductUrl = prodGeneric
        Exit Function
    End If

    Dim parts As Variant
    parts = Split(afterCom, "/")
    If UBound(parts) < 0 Then
        BuildProductUrl = prodGeneric
        Exit Function
    End If

    Dim lastIdx As Long
    lastIdx = UBound(parts)

    Dim code As String
    Dim candidate As String
    Dim normCandidate As String
    code = ""

    '1) Try last slug (file name without .html)
    candidate = RemoveExtensionHtml(CStr(parts(lastIdx)))
    normCandidate = NormText(candidate)

    If Len(normCandidate) > 0 Then
        If dict.Exists(normCandidate) Then
            code = Trim$(CStr(dict(normCandidate)))
            If Len(code) > 0 Then
                BuildProductUrl = base & "/products.html/producttype_" & code & ".html"
                Exit Function
            End If
        End If
    End If

    '2) Try folders before the file segment (skip indices 0/1/2 = country/lang/products)
    Dim i As Long
    If lastIdx >= 3 Then
        For i = lastIdx - 1 To 3 Step -1
            candidate = CStr(parts(i))
            normCandidate = NormText(candidate)

            If Len(normCandidate) > 0 Then
                If dict.Exists(normCandidate) Then
                    code = Trim$(CStr(dict(normCandidate)))
                    If Len(code) > 0 Then
                        BuildProductUrl = base & "/products.html/producttype_" & code & ".html"
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If

    BuildProductUrl = prodGeneric
End Function

'=========================================================
' PART B: Error Check button behavior
'=========================================================
Public Sub ErrorCheck_Run()
    On Error GoTo CleanFail

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsSrc As Worksheet: Set wsSrc = GetWs(wb, SHEET_OW)
    Dim wsChk As Worksheet: Set wsChk = GetWs(wb, SHEET_CHECK)

    If wsSrc Is Nothing Or wsChk Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    SetProgress "Error check: preparing list..."

    '1) Read new URLs from OW col B until last used row
    Dim lastCell As Range
    Set lastCell = wsSrc.Columns(2).Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    If lastCell Is Nothing Or lastCell.row < 2 Then GoTo CleanUpNoUrls

    Dim srcLastRow As Long
    srcLastRow = lastCell.row

    Dim srcVals As Variant
    srcVals = wsSrc.Range(wsSrc.Cells(2, 2), wsSrc.Cells(srcLastRow, 2)).Value2

    Dim i As Long, usedN As Long
    usedN = 0

    For i = 1 To UBound(srcVals, 1)
        If Len(Trim$(CStr(srcVals(i, 1)))) = 0 Then Exit For
        usedN = usedN + 1
    Next i

    If usedN = 0 Then GoTo CleanUpNoUrls

    Dim urls() As Variant
    ReDim urls(1 To usedN, 1 To 1)

    For i = 1 To usedN
        urls(i, 1) = Trim$(CStr(srcVals(i, 1)))
    Next i

    '2) Clear old data on error check sheet
    Dim lastA As Range, lastB As Range, lastRow As Long
    Set lastA = wsChk.Columns(1).Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set lastB = wsChk.Columns(2).Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    lastRow = 1
    If Not lastA Is Nothing Then lastRow = Application.WorksheetFunction.Max(lastRow, lastA.row)
    If Not lastB Is Nothing Then lastRow = Application.WorksheetFunction.Max(lastRow, lastB.row)

    If lastRow >= 2 Then
        wsChk.Range(wsChk.Cells(2, 1), wsChk.Cells(lastRow, 2)).ClearContents
    End If

    '3) Copy URLs into error check col A
    wsChk.Range("A2").Resize(usedN, 1).Value2 = urls

    '4) HTTP status check
    SetProgress "Error check: checking HTTP status... 0 / " & usedN

    Dim outStatus() As Variant
    ReDim outStatus(1 To usedN, 1 To 1)

    Dim cache As Object
    Set cache = CreateObject("Scripting.Dictionary")
    cache.CompareMode = 1

    Dim u As String, code As Long
    For i = 1 To usedN
        u = CStr(urls(i, 1))

        If cache.Exists(u) Then
            code = CLng(cache(u))
        Else
            code = GetHttpStatus(u, False)
            cache.Add u, code
        End If

        outStatus(i, 1) = HttpStatusLabel(code)

        If (i Mod 200) = 0 Then
            SetProgress "Error check: checking HTTP status... " & i & " / " & usedN
            DoEvents
        End If
    Next i

    wsChk.Range("B2").Resize(usedN, 1).Value2 = outStatus

    ResetExcelState
    MsgBox "Done. Copied " & usedN & " URLs and checked status on '" & SHEET_CHECK & "'.", vbInformation
    Exit Sub

CleanUpNoUrls:
    ResetExcelState
    MsgBox "No generated URLs found in '" & SHEET_OW & "' column B (starting at B2).", vbInformation
    Exit Sub

CleanFail:
    ResetExcelState
    MsgBox "ErrorCheck_Run failed: " & Err.Description, vbCritical, "VBA Error"
End Sub

Private Function HttpStatusLabel(ByVal statusCode As Long) As String
    Select Case statusCode
        Case 200 To 299: HttpStatusLabel = "OK"
        Case 300 To 399: HttpStatusLabel = "Redirect"
        Case 401: HttpStatusLabel = "Unauthorized (401)"
        Case 403: HttpStatusLabel = "Forbidden (403)"
        Case 404: HttpStatusLabel = "Not Found (404)"
        Case 408: HttpStatusLabel = "Request Timeout (408)"
        Case 429: HttpStatusLabel = "Too Many Requests (429)"
        Case 500 To 599: HttpStatusLabel = "Server Error (" & statusCode & ")"
        Case -1: HttpStatusLabel = "Request failed"
        Case Else: HttpStatusLabel = "HTTP " & statusCode
    End Select
End Function

Private Function GetHttpStatus(ByVal url As String, Optional ByVal followRedirects As Boolean = False) As Long
    On Error GoTo TryGet

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Option(6) = followRedirects
    http.SetTimeouts 2000, 2000, 4000, 4000

    http.Open "HEAD", url, False
    http.Send

    GetHttpStatus = CLng(http.Status)
    Exit Function

TryGet:
    On Error GoTo Fail

    Dim http2 As Object
    Set http2 = CreateObject("WinHttp.WinHttpRequest.5.1")

    http2.Option(6) = followRedirects
    http2.SetTimeouts 2000, 2000, 4000, 4000

    http2.Open "GET", url, False
    http2.Send

    GetHttpStatus = CLng(http2.Status)
    Exit Function

Fail:
    GetHttpStatus = -1
End Function


Attribute VB_Name = "Module2"
Option Explicit

Public Sub Run_OW_URL_Status_Check()

    Const SRC_SHEET As String = "OW URL (Paste Here)"
    Const DST_SHEET As String = "Error Check OW URL"

    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim lastSrcRow As Long, n As Long
    Dim arrUrls As Variant, arrRes() As Variant
    Dim i As Long, url As String
    Dim statusCode As Long, errMsg As String

    On Error GoTo CleanFail

    Set wb = ThisWorkbook
    Set wsSrc = wb.Worksheets(SRC_SHEET)
    Set wsDst = wb.Worksheets(DST_SHEET)

    'Speed up Excel while running
    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationManual

    '1) Find last row in source (Column A)
    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).row
    If lastSrcRow < 2 Then GoTo CleanExit 'nothing to copy

    '2) Read source URLs into memory
    arrUrls = wsSrc.Range("A2:A" & lastSrcRow).Value2
    n = UBound(arrUrls, 1)

    '3) Clear destination columns A:B (row 2 down) so it REPLACES anything
    wsDst.Range("A2:B" & wsDst.Rows.Count).ClearContents

    '4) Paste URLs into destination column A
    wsDst.Range("A2").Resize(n, 1).Value2 = arrUrls

    '5) Prepare results array for column B
    ReDim arrRes(1 To n, 1 To 1)

    'Optional: cache duplicates in the same run (faster)
    Dim cache As Object
    Set cache = CreateObject("Scripting.Dictionary")
    cache.CompareMode = 1 'TextCompare

    Dim key As String, resultText As String

    '6) Check each URL; write status to column B
    For i = 1 To n
        url = Trim$(CStr(arrUrls(i, 1)))

        If Len(url) = 0 Then
            arrRes(i, 1) = vbNullString
        Else
            key = NormalizeUrlKey(url)

            If cache.Exists(key) Then
                arrRes(i, 1) = cache(key)
            Else
                errMsg = vbNullString
                statusCode = GetHttpStatusCode(url, errMsg)

                resultText = StatusLabel_WithRedirectOK(statusCode, errMsg)

                cache.Add key, resultText
                arrRes(i, 1) = resultText
            End If
        End If

        If i Mod 50 = 0 Then
            Application.StatusBar = "Checking URLs: " & i & " / " & n
            DoEvents
        End If
    Next i

    '7) Write results back to destination column B
    wsDst.Range("B2").Resize(n, 1).Value2 = arrRes

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = oldCalc
    Exit Sub

CleanFail:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = oldCalc
    MsgBox "URL check failed: " & Err.Description, vbExclamation, "OW URL Status Check"
End Sub


Private Function StatusLabel_WithRedirectOK(ByVal statusCode As Long, ByVal errMsg As String) As String
    If statusCode = 0 Then
        StatusLabel_WithRedirectOK = "ERROR (" & errMsg & ")"
        Exit Function
    End If

    Select Case statusCode
        Case 200 To 299
            StatusLabel_WithRedirectOK = "OK (" & statusCode & ")"

        Case 301, 302
            StatusLabel_WithRedirectOK = "Redirect OK (" & statusCode & ")"

        Case 300 To 399
            StatusLabel_WithRedirectOK = "Redirect (" & statusCode & ")"

        Case 404
            StatusLabel_WithRedirectOK = "404 Not Found"

        Case 400 To 499
            StatusLabel_WithRedirectOK = "ERROR (" & statusCode & ")"

        Case 500 To 599
            StatusLabel_WithRedirectOK = "Server Error (" & statusCode & ")"

        Case Else
            StatusLabel_WithRedirectOK = "HTTP " & statusCode
    End Select
End Function


Private Function NormalizeUrlKey(ByVal url As String) As String
    'Normalize so "same" URLs match:
    ' - trim
    ' - remove fragment (#something)
    ' - remove trailing slash
    ' - lowercase
    Dim s As String
    s = Trim$(url)

    Dim p As Long
    p = InStr(1, s, "#", vbTextCompare)
    If p > 0 Then s = Left$(s, p - 1)

    Do While Len(s) > 0 And Right$(s, 1) = "/"
        s = Left$(s, Len(s) - 1)
    Loop

    NormalizeUrlKey = LCase$(s)
End Function


Private Function GetHttpStatusCode(ByVal url As String, ByRef errMsg As String) As Long
    'Returns HTTP status (200/301/404/500 etc). Returns 0 if request fails.

    On Error GoTo TryGet

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Option(6) = True 'follow redirects
    http.SetTimeouts 5000, 5000, 5000, 12000

    http.Open "HEAD", url, False
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 (Excel VBA)"
    http.Send

    GetHttpStatusCode = CLng(http.Status)
    Exit Function

TryGet:
    Err.Clear
    On Error GoTo Fail

    Dim http2 As Object
    Set http2 = CreateObject("WinHttp.WinHttpRequest.5.1")

    http2.Option(6) = True
    http2.SetTimeouts 5000, 5000, 5000, 15000

    http2.Open "GET", url, False
    http2.SetRequestHeader "User-Agent", "Mozilla/5.0 (Excel VBA)"
    http2.Send

    GetHttpStatusCode = CLng(http2.Status)
    Exit Function

Fail:
    errMsg = Err.Description
    GetHttpStatusCode = 0
End Function


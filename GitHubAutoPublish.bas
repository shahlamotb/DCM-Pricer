Attribute VB_Name = "GitHubAutoPublish"
Option Explicit

' =====================================================================
' Sends data.js to GitHub automatically, every few minutes.
' Needs the module ExportPricerData (already in this workbook).
'
' ONE-TIME SETUP
'   1. GH_OWNER and GH_REPO below are already set to your repository.
'   2. Run SetGitHubToken once (Alt+F8) and paste your token.
'   3. Run PublishNow once to test.
' DAILY USE
'   Run StartAutoPublish. Keep Excel (and Bloomberg) open.
'   Run StopAutoPublish when you want it to stop.
' =====================================================================

Private Const GH_OWNER As String = "shahlamotb"
Private Const GH_REPO As String = "DCM-Pricer"
Private Const GH_BRANCH As String = "main"
Private Const GH_FILE As String = "data.js"
Private Const EVERY_MINUTES As Long = 5
' Every Nth upload also re-reads the method values and comps for all tenors (this briefly steps the
' Tenor cell on the dashboard and puts it back). 1 = every upload. Raise it (e.g. 3) if Excel feels slow.
Private Const FULL_SCAN_EVERY As Long = 1

Private nextRun As Date
Private lastSent As String
Private tick As Long

' Saves the token on THIS PC (Windows user settings), not inside the workbook.
Public Sub SetGitHubToken()
    Dim t As String
    t = InputBox("Paste your GitHub token here (it is stored on this PC only):", "GitHub token")
    If Trim$(t) = "" Then Exit Sub
    SaveSetting "DCMPricer", "GitHub", "Token", Trim$(t)
    MsgBox "Token saved on this PC.", vbInformation
End Sub

' Builds data.js and uploads it once, and tells you what happened.
Public Sub PublishNow()
    MsgBox DoPublish(True), vbInformation, "DCM Pricer"
End Sub

Public Sub StartAutoPublish()
    Dim msg As String
    tick = 0
    msg = DoPublish(True)
    ScheduleNext
    MsgBox "Auto-publish is ON (every " & EVERY_MINUTES & " minutes)." & vbCrLf & vbCrLf & "First run: " & msg & vbCrLf & vbCrLf & _
           "Keep this workbook open. Do not leave a cell in edit mode, or the timer waits.", vbInformation, "DCM Pricer"
End Sub

Public Sub StopAutoPublish()
    On Error Resume Next
    Application.OnTime EarliestTime:=nextRun, Procedure:="'" & ThisWorkbook.Name & "'!AutoTick", Schedule:=False
    On Error GoTo 0
    Application.StatusBar = False
    MsgBox "Auto-publish is OFF.", vbInformation, "DCM Pricer"
End Sub

' Called by the timer. Never shows pop-ups; the result appears in Excel's status bar.
Public Sub AutoTick()
    Dim msg As String
    tick = tick + 1
    msg = DoPublish((tick Mod FULL_SCAN_EVERY) = 0)
    Application.StatusBar = "DCM Pricer " & Format$(Now, "hh:nn") & ": " & msg
    ScheduleNext
End Sub

Private Sub ScheduleNext()
    nextRun = Now + TimeSerial(0, EVERY_MINUTES, 0)
    Application.OnTime EarliestTime:=nextRun, Procedure:="'" & ThisWorkbook.Name & "'!AutoTick"
End Sub

Private Function DoPublish(ByVal scan As Boolean) As String
    Dim token As String, path As String
    token = GetSetting("DCMPricer", "GitHub", "Token", "")
    If token = "" Then DoPublish = "No token saved yet. Run SetGitHubToken first.": Exit Function

    On Error GoTo Fail
    Application.Calculate
    ExportPricerData.gSilent = True
    ExportPricerData.ExportForTimer scan
    ExportPricerData.gSilent = False
    If Not ExportPricerData.gOk Then DoPublish = "Export failed: " & ExportPricerData.gMsg: Exit Function

    path = ExportPricerData.OutDir() & "\data.js"
    Dim b64 As String
    b64 = FileToBase64(path)
    If b64 = lastSent Then DoPublish = "No change since the last upload.": Exit Function

    Dim url As String, status As Long, resp As String, sha As String
    url = "https://api.github.com/repos/" & GH_OWNER & "/" & GH_REPO & "/contents/" & GH_FILE
    resp = Http("GET", url & "?ref=" & GH_BRANCH, token, "", status)
    If status = 200 Then
        sha = ExtractValue(resp, "sha")
    ElseIf status <> 404 Then
        DoPublish = "GitHub said " & status & " when reading the file. Check the username, repository name and token."
        Exit Function
    End If

    Dim body As String
    body = "{""message"":""Pricer data " & Format$(Now, "yyyy-mm-dd hh:nn") & """,""branch"":""" & GH_BRANCH & """,""content"":""" & b64 & """"
    If sha <> "" Then body = body & ",""sha"":""" & sha & """"
    body = body & "}"
    resp = Http("PUT", url, token, body, status)
    If status = 200 Or status = 201 Then
        lastSent = b64
        DoPublish = "Uploaded at " & Format$(Now, "hh:nn") & ". The website updates in about a minute."
    Else
        DoPublish = "Upload failed (GitHub " & status & "). Check that the token can write to the repository."
    End If
    Exit Function

Fail:
    ExportPricerData.gSilent = False
    DoPublish = "Error: " & Err.Description
End Function

Private Function Http(ByVal verb As String, ByVal url As String, ByVal token As String, ByVal body As String, ByRef status As Long) As String
    Dim h As Object
    Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    h.setTimeouts 10000, 10000, 20000, 30000
    h.Open verb, url, False
    h.setRequestHeader "Authorization", "Bearer " & token
    h.setRequestHeader "Accept", "application/vnd.github+json"
    h.setRequestHeader "User-Agent", "DCM-Pricer-Excel"
    h.setRequestHeader "X-GitHub-Api-Version", "2022-11-28"
    If body <> "" Then
        h.setRequestHeader "Content-Type", "application/json"
        h.send body
    Else
        h.send
    End If
    status = h.Status
    Http = h.responseText
End Function

Private Function FileToBase64(ByVal path As String) As String
    Dim stm As Object, xml As Object, el As Object, bytes As Variant
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.LoadFromFile path
    bytes = stm.Read
    stm.Close
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set el = xml.createElement("b64")
    el.DataType = "bin.base64"
    el.nodeTypedValue = bytes
    FileToBase64 = Replace(Replace(el.Text, vbLf, ""), vbCr, "")
End Function

' Pulls the first "key":"value" out of a JSON reply.
Private Function ExtractValue(ByVal json As String, ByVal key As String) As String
    Dim p As Long, q As Long
    p = InStr(1, json, """" & key & """:""")
    If p = 0 Then Exit Function
    p = p + Len(key) + 4
    q = InStr(p, json, """")
    If q > p Then ExtractValue = Mid$(json, p, q - p)
End Function

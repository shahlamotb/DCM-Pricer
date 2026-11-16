Attribute VB_Name = "ExportPricerData"
Option Explicit

' =====================================================================
' Writes data.js for the mobile DCM Pricer page (final model layout).
' Run: Alt+F8 -> ExportPricerData -> Run
' Refresh Bloomberg and let the workbook finish calculating first.
' =====================================================================

' ---- settings (edit if you move things in the workbook) ----
Private Const OUT_DIR_OVERRIDE As String = ""          ' e.g. "C:\Users\me\pricer"; blank = folder of this workbook
Private Const BLOCK_SHEET As String = "AllCurvesPricing"
Private Const INPUT_SHEET As String = "INPUTS"
Private Const NIP_ADDR As String = "D20"               ' New Issue Premium
Private Const DISC_ADDR As String = "D18"              ' Sukuk Discount
Private Const PIFDELTA_ADDR As String = "D22"          ' PIF manual delta
Private Const PIFMODE_ADDR As String = "D24"           ' PIF delta mode (Secondaries / Manual)
Private Const SOURCE_ADDR As String = "D8"             ' Bloomberg data source (e.g. CBBT)
Private Const UST_SHEET As String = "Helper.Sheet"
Private Const UST_FIRST_ROW As Long = 38               ' CT3, CT5, CT7, CT10, CT30 down one column
Private Const UST_COL As Long = 11                     ' column K ("UST % (Current)")
Private Const SUMMARY_TITLE As String = "Fixed Curves Summary"   ' title cell in column C of AllCurvesPricing
Private Const DEFAULT_ISSUER As String = "KSA"         ' for headers like "3Y Bond Pricer" with no issuer name
Private Const DASH_SHEET As String = "Pricer DB"       ' the dashboard sheet
Private Const TENOR_ADDR As String = "F4"              ' the Tenor drop-down cell on the dashboard
' -------------------------------------------------------------

' Used by GitHubAutoPublish: when True, no pop-ups are shown and the result is stored here.
Public gSilent As Boolean
Public gOk As Boolean
Public gMsg As String
Public gScan As Boolean            ' True = re-read method values and comps for every tenor now
Private cacheMethods As String     ' last scan, reused between scans
Private cacheComps As String

Private Sub Say(ByVal text As String, ByVal style As Long)
    gMsg = text
    If Not gSilent Then MsgBox text, style
End Sub

Public Sub ExportPricerData()
    gScan = True                   ' run by hand: always do the full scan
    ExportCore
End Sub

' Called by the timer module. scan = False reuses the last scan.
Public Sub ExportForTimer(ByVal scan As Boolean)
    gScan = scan
    ExportCore
End Sub

Private Sub ExportCore()
    gOk = False
    gMsg = ""
    Dim wsB As Worksheet, wsI As Worksheet, wsU As Worksheet
    On Error Resume Next
    Set wsB = ThisWorkbook.Worksheets(BLOCK_SHEET)
    Set wsI = ThisWorkbook.Worksheets(INPUT_SHEET)
    Set wsU = ThisWorkbook.Worksheets(UST_SHEET)
    On Error GoTo 0
    If wsB Is Nothing Or wsI Is Nothing Or wsU Is Nothing Then
        Say "Could not find one of the sheets: " & BLOCK_SHEET & ", " & INPUT_SHEET & ", " & UST_SHEET, vbExclamation
        Exit Sub
    End If

    ' Treasury yields for 3,5,7,10,30Y
    Dim ust(0 To 4) As Double, i As Long, v As Variant
    For i = 0 To 4
        v = wsU.Cells(UST_FIRST_ROW + i, UST_COL).Value
        If Not IsNum(v) Then
            Say "Treasury yield is not a number at " & wsU.Cells(UST_FIRST_ROW + i, UST_COL).Address(False, False) & " on " & UST_SHEET & ".", vbExclamation
            Exit Sub
        End If
        ust(i) = v
    Next i

    Dim nipV As Variant, discV As Variant, pdV As Variant, modeS As String, srcS As String
    nipV = wsI.Range(NIP_ADDR).Value
    discV = wsI.Range(DISC_ADDR).Value
    pdV = wsI.Range(PIFDELTA_ADDR).Value
    modeS = Trim$(wsI.Range(PIFMODE_ADDR).Value & "")
    srcS = Trim$(wsI.Range(SOURCE_ADDR).Value & "")
    If Not IsNum(nipV) Then nipV = 0
    If Not IsNum(discV) Then discV = 0
    If Not IsNum(pdV) Then pdV = 0
    If LCase$(modeS) <> "manual" Then modeS = "Secondaries" Else modeS = "Manual"
    If srcS = "" Then srcS = "CBBT"

    ' ---- final FV grid: "Fixed Curves Summary" on AllCurvesPricing ----
    Dim sumRow As Long, r As Long, k As Long
    For r = 1 To 60
        If Trim$(wsB.Cells(r, 3).Value & "") = SUMMARY_TITLE Then sumRow = r: Exit For
    Next r
    If sumRow = 0 Then
        Say "Could not find '" & SUMMARY_TITLE & "' in column C of " & BLOCK_SHEET & ".", vbExclamation
        Exit Sub
    End If
    Dim fixedJs As String, nm As String, bondS As String, sukS As String, issLine As String, nIss As Long
    For r = sumRow + 3 To sumRow + 17
        nm = Trim$(wsB.Cells(r, 3).Value & "")
        If nm <> "" And Left$(LCase$(nm), 5) <> "delta" Then
            bondS = GridPart(wsB, r, sumRow + 2, 4)    ' Bonds in columns D:H
            sukS = GridPart(wsB, r, sumRow + 2, 9)     ' Sukuk in columns I:M
            issLine = ""
            If bondS <> "" Then issLine = """Bond"":{" & bondS & "}"
            If sukS <> "" Then issLine = issLine & IIf(issLine = "", "", ",") & """Sukuk"":{" & sukS & "}"
            If issLine <> "" Then
                fixedJs = fixedJs & IIf(fixedJs = "", "", "," & vbLf) & "   " & Q(nm) & ": {" & issLine & "}"
                nIss = nIss + 1
            End If
        End If
    Next r
    If nIss = 0 Then
        Say "No issuer rows found under '" & SUMMARY_TITLE & "'. Nothing was written.", vbExclamation
        Exit Sub
    End If

    ' ---- own-curve pricing blocks: "<n>Y <Issuer> <Bond|Sukuk> Pricer" ----
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\s*(\d+)Y\s+(?:(\w+)\s+)?(?:(Bond|Sukuk)\s*)?Pricer"
    re.IgnoreCase = True

    Dim d As Object, di As Object, dt As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim arr As Variant, r0 As Long, c0 As Long, rr As Long, cc As Long
    Dim m As Object, iss As String, typ As String, ten As String, js As String
    Dim nBlocks As Long
    arr = wsB.UsedRange.Value
    r0 = wsB.UsedRange.Row
    c0 = wsB.UsedRange.Column

    For rr = 1 To UBound(arr, 1)
        For cc = 1 To UBound(arr, 2)
            If VarType(arr(rr, cc)) = vbString Then
                If re.Test(arr(rr, cc)) Then
                    Set m = re.Execute(arr(rr, cc))(0)
                    ten = m.SubMatches(0)
                    iss = m.SubMatches(1) & ""
                    typ = m.SubMatches(2) & ""
                    If typ = "" And (LCase$(iss) = "bond" Or LCase$(iss) = "sukuk") Then
                        typ = iss
                        iss = DEFAULT_ISSUER
                    End If
                    If iss <> "" And typ <> "" Then
                        typ = UCase$(Left$(typ, 1)) & LCase$(Mid$(typ, 2))
                        js = ParseBlock(wsB, r0 + rr - 1, c0 + cc - 1)
                        If js <> "" Then
                            If Not d.Exists(iss) Then Set d(iss) = CreateObject("Scripting.Dictionary")
                            Set di = d(iss)
                            If Not di.Exists(typ) Then Set di(typ) = CreateObject("Scripting.Dictionary")
                            Set dt = di(typ)
                            dt(CStr(CLng(ten))) = js
                            nBlocks = nBlocks + 1
                        End If
                    End If
                End If
            End If
        Next cc
    Next rr

    If nBlocks = 0 Then
        Say "No pricing blocks found on " & BLOCK_SHEET & ". Nothing was written.", vbExclamation
        Exit Sub
    End If

    ' ---- method values and comps for every tenor (the Excel charts only hold the selected tenor) ----
    If gScan Or cacheMethods = "" Then
        Dim wsD As Worksheet
        On Error Resume Next
        Set wsD = ThisWorkbook.Worksheets(DASH_SHEET)
        On Error GoTo 0
        If Not wsD Is Nothing Then ScanTenors wsD
    End If

    ' ---- write data.js ----
    Dim path As String, f As Integer, k1 As Variant, k2 As Variant, k3 As Variant
    path = OutDir() & "\data.js"
    f = FreeFile
    On Error GoTo WriteErr
    Open path For Output As #f
    Print #f, "// Written by the Excel macro ExportPricerData. Do not edit by hand; re-run the macro instead."
    Print #f, "const DATA = {"
    Print #f, "  ""asOf"": " & Q(Format$(Now, "d mmm yyyy hh:nn")) & ","
    If UtcStamp() <> "" Then Print #f, "  ""updatedUtc"": " & Q(UtcStamp()) & ","
    Print #f, "  ""source"": " & Q("Bloomberg (" & srcS & ")") & ","
    Print #f, "  ""ust"": {""3"": " & N(ust(0), "0.000000") & ", ""5"": " & N(ust(1), "0.000000") & ", ""7"": " & N(ust(2), "0.000000") & ", ""10"": " & N(ust(3), "0.000000") & ", ""30"": " & N(ust(4), "0.000000") & "},"
    Print #f, "  ""nip"": " & N(nipV, "0.##") & ","
    Print #f, "  ""sukukDiscount"": " & N(discV, "0.##") & ","
    Print #f, "  ""pifManualDelta"": " & N(pdV, "0.##") & ","
    Print #f, "  ""pifMode"": " & Q(modeS) & ","
    Print #f, "  ""fixed"": {"
    Print #f, fixedJs
    Print #f, "  },"
    If cacheMethods <> "" Then Print #f, "  ""methods"": {" & vbLf & cacheMethods & vbLf & "  },"
    If cacheComps <> "" Then Print #f, "  ""comps"": {" & vbLf & cacheComps & vbLf & "  },"
    Print #f, "  ""blocks"": {"
    Dim iIdx As Long, tIdx As Long, nIdx As Long, sep As String
    iIdx = 0
    For Each k1 In d.Keys
        iIdx = iIdx + 1
        Print #f, "   " & Q(CStr(k1)) & ": {"
        Set di = d(k1)
        tIdx = 0
        For Each k2 In di.Keys
            tIdx = tIdx + 1
            Print #f, "    " & Q(CStr(k2)) & ": {"
            Set dt = di(k2)
            nIdx = 0
            For Each k3 In dt.Keys
                nIdx = nIdx + 1
                sep = IIf(nIdx < dt.Count, ",", "")
                Print #f, "      " & Q(CStr(k3)) & ": " & dt(k3) & sep
            Next k3
            Print #f, "    }" & IIf(tIdx < di.Count, ",", "")
        Next k2
        Print #f, "   }" & IIf(iIdx < d.Count, ",", "")
    Next k1
    Print #f, "  }"
    Print #f, "};"
    Close #f
    On Error GoTo 0

    gOk = True
    Say "Wrote " & nIss & " issuers and " & nBlocks & " pricing blocks to:" & vbCrLf & path, vbInformation
    Exit Sub

WriteErr:
    On Error Resume Next
    Close #f
    Say "Could not write " & path & vbCrLf & Err.Description & vbCrLf & "Set OUT_DIR_OVERRIDE at the top of the macro to a folder you can write to.", vbExclamation
End Sub

' Reads five tenor columns of the summary table (starting at column c1) for one issuer row.
' tenorRow holds the tenor numbers (3, 5, 7, 10, 30). Returns e.g. "3":95,"5":100 or "" if the row has no numbers.
Private Function GridPart(ws As Worksheet, ByVal r As Long, ByVal tenorRow As Long, ByVal c1 As Long) As String
    Dim j As Long, v As Variant, tv As Variant, s As String
    For j = 0 To 4
        v = ws.Cells(r, c1 + j).Value
        tv = ws.Cells(tenorRow, c1 + j).Value
        If IsNum(v) And IsNum(tv) Then s = s & IIf(s = "", "", ",") & """" & CStr(CLng(tv)) & """:" & N(v, "0.##")
    Next j
    GridPart = s
End Function

' Reads one pricer block whose title is at (r0, col). Returns a JSON object string, or "" if the block can't be read.
Private Function ParseBlock(ws As Worksheet, ByVal r0 As Long, ByVal col As Long) As String
    Dim r As Long, fvRow As Long, v As Variant
    For r = r0 + 1 To r0 + 40
        v = ws.Cells(r, col).Value
        If VarType(v) = vbString Then
            If InStr(1, v, "fv from", vbTextCompare) > 0 Then
                fvRow = r
                Exit For
            End If
        End If
    Next r
    If fvRow = 0 Then Exit Function

    Dim refs As String, k As Long, nm As String, nCount As Long, n1 As Double, n2 As Double
    For r = r0 + 2 To fvRow - 1
        nm = "": nCount = 0
        For k = col To col + 3
            v = ws.Cells(r, k).Value
            If VarType(v) = vbString Then
                If nm = "" And Trim$(v) <> "" And Trim$(v) <> "NA" Then nm = Trim$(v)
            ElseIf IsNum(v) Then
                nCount = nCount + 1
                If nCount = 1 Then
                    n1 = v
                ElseIf nCount = 2 Then
                    n2 = v
                End If
            End If
        Next k
        If nm <> "" And nCount >= 2 Then refs = refs & IIf(refs = "", "", ",") & Trip(nm, n1, n2)
    Next r

    Dim methods As String, avg As String, rr As Long, a As Variant, labV As Variant, lab As String
    Dim used As String, j As Long, mv As Variant, valS As String
    avg = "null"
    For rr = fvRow + 2 To fvRow + 24
        a = ws.Cells(rr, col).Value
        labV = ws.Cells(rr, col + 3).Value
        If VarType(a) = vbString Then
            If Trim$(a) = "FV Average" Then
                If IsNum(labV) Then avg = N(labV)
                Exit For
            End If
        End If
        If VarType(labV) = vbString Then
            lab = MethodName(Trim$(labV))
            If lab <> "" Then
                used = ""
                For j = rr To rr + 1
                    nm = Trim$(ws.Cells(j, col).Value & "")
                    If nm <> "" And nm <> "NA" And IsNum(ws.Cells(j, col + 1).Value) And IsNum(ws.Cells(j, col + 2).Value) Then
                        used = used & IIf(used = "", "", ",") & Trip(nm, ws.Cells(j, col + 1).Value, ws.Cells(j, col + 2).Value)
                    End If
                Next j
                mv = ws.Cells(rr + 1, col + 3).Value
                If IsNum(mv) Then valS = N(mv) Else valS = "null"
                methods = methods & IIf(methods = "", "", ",") & Q(lab) & ":{""value"":" & valS & ",""used"":[" & used & "]}"
            End If
        End If
    Next rr

    ParseBlock = "{""refs"":[" & refs & "],""methods"":{" & methods & "},""avg"":" & avg & "}"
End Function

Private Function MethodName(ByVal s As String) As String
    Select Case LCase$(s)
        Case "interpolation": MethodName = "Interpolation"
        Case "extrapolation (lower)": MethodName = "Extrapolation (Lower)"
        Case "curve extension": MethodName = "Curve Extension"
        Case "curve reduction": MethodName = "Curve Reduction"
    End Select
End Function

Public Function OutDir() As String
    Dim p As String
    p = OUT_DIR_OVERRIDE
    If p = "" Then p = ThisWorkbook.Path
    If p = "" Or LCase$(Left$(p, 4)) = "http" Then p = Environ$("USERPROFILE") & "\Desktop"
    OutDir = p
End Function

Private Function IsNum(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency: IsNum = True
    End Select
End Function

' number with a "." decimal separator whatever the Windows locale is
Private Function N(ByVal v As Variant, Optional ByVal fmt As String = "0.00") As String
    N = Replace(Format$(v, fmt), ",", ".")
    If Right$(N, 1) = "." Then N = Left$(N, Len(N) - 1)
    If N = "" Or N = "-" Then N = "0"
End Function

Private Function Q(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    Q = """" & s & """"
End Function

Private Function Trip(ByVal nm As String, ByVal t As Double, ByVal s As Double) As String
    Trip = "[" & Q(nm) & "," & N(t) & "," & N(s) & "]"
End Function


' ---------------------------------------------------------------------
' Full scan: the Excel charts only hold the tenor selected on the dashboard,
' so this steps the Tenor cell through 3, 5, 7, 10 and 30, reads the cells the
' dashboard reads, then puts the original tenor back.
' ---------------------------------------------------------------------
Private Sub ScanTenors(wsD As Worksheet)
    Dim tens As Variant, combos As Variant, i As Long, j As Long
    Dim origT As Variant, calcOld As Long
    Dim dM As Object, dC As Object
    Set dM = CreateObject("Scripting.Dictionary")
    Set dC = CreateObject("Scripting.Dictionary")
    tens = Array(3, 5, 7, 10, 30)
    combos = Array("KSA|Bond", "KSA|Sukuk", "PIF|Bond", "PIF|Sukuk", "Aramco|Bond", "Aramco|Sukuk", "SECO|Sukuk", "STC|Sukuk", "Avilease|Bond", "Maaden|Sukuk", "SRC|Sukuk")

    origT = wsD.Range(TENOR_ADDR).Value
    On Error GoTo Restore
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    For i = 0 To UBound(tens)
        wsD.Range(TENOR_ADDR).Value = tens(i)
        Application.Calculate
        For j = 0 To UBound(combos)
            Dim parts As Variant, js As String
            parts = Split(combos(j), "|")
            js = MethodsJson(CStr(parts(0)), CStr(parts(1)))
            If js <> "" Then NestSet dM, CStr(parts(0)), CStr(parts(1)), CStr(tens(i)), js
        Next j
        CompsToDict dC, CStr(tens(i))
    Next i
    cacheMethods = NestJson(dM)
    cacheComps = NestJson(dC)
Restore:
    On Error Resume Next
    wsD.Range(TENOR_ADDR).Value = origT
    Application.Calculate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' {"Interpolation":..,"Extrapolation (Lower)":..,"Curve Extension":..,"Curve Reduction":..} from the named cells, or "" if the names are missing.
Private Function MethodsJson(ByVal iss As String, ByVal ty As String) As String
    Dim labs As Variant, pre As Variant, i As Long, v As Variant, s As String, found As Boolean, r As Range
    labs = Array("Interpolation", "Extrapolation (Lower)", "Curve Extension", "Curve Reduction")
    pre = Array("Interp", "Extrap", "Exten", "Reduc")
    For i = 0 To 3
        Set r = Nothing
        On Error Resume Next
        Set r = ThisWorkbook.Names(pre(i) & iss & ty).RefersToRange
        On Error GoTo 0
        If r Is Nothing Then
            v = Empty
        Else
            found = True
            v = r.Cells(1, 1).Value
        End If
        s = s & IIf(s = "", "", ",") & Q(CStr(labs(i))) & ":" & IIf(IsNum(v), N(v), "null")
    Next i
    If found Then MethodsJson = "{" & s & "}"
End Function

' Reads every table called <Issuer>_<Bond|Sukuk>_FV (the comps tables on the issuer sheets).
Private Sub CompsToDict(d As Object, ByVal ten As String)
    Dim ws As Worksheet, lo As ListObject, parts As Variant, c As Long, nm As String, v As Variant, s As String
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            parts = Split(lo.Name, "_")
            If UBound(parts) = 2 Then
                If parts(2) = "FV" And (parts(1) = "Bond" Or parts(1) = "Sukuk") Then
                    s = ""
                    For c = 1 To lo.ListColumns.Count
                        nm = Trim$(lo.HeaderRowRange.Cells(1, c).Value & "")
                        v = lo.DataBodyRange.Cells(1, c).Value
                        If nm <> "" And IsNum(v) Then s = s & IIf(s = "", "", ",") & "[" & Q(nm) & "," & N(v) & "]"
                    Next c
                    If s <> "" Then NestSet d, CStr(parts(0)), CStr(parts(1)), ten, "[" & s & "]"
                End If
            End If
        Next lo
    Next ws
End Sub

' d(issuer)(type)(tenor) = json text
Private Sub NestSet(d As Object, ByVal a As String, ByVal b As String, ByVal c As String, ByVal js As String)
    If Not d.Exists(a) Then Set d(a) = CreateObject("Scripting.Dictionary")
    If Not d(a).Exists(b) Then Set d(a)(b) = CreateObject("Scripting.Dictionary")
    d(a)(b)(c) = js
End Sub

Private Function NestJson(d As Object) As String
    Dim k1 As Variant, k2 As Variant, k3 As Variant, s1 As String, s2 As String, s3 As String
    For Each k1 In d.Keys
        s2 = ""
        For Each k2 In d(k1).Keys
            s3 = ""
            For Each k3 In d(k1)(k2).Keys
                s3 = s3 & IIf(s3 = "", "", ",") & Q(CStr(k3)) & ":" & d(k1)(k2)(k3)
            Next k3
            s2 = s2 & IIf(s2 = "", "", "," & vbLf) & "    " & Q(CStr(k2)) & ": {" & s3 & "}"
        Next k2
        s1 = s1 & IIf(s1 = "", "", "," & vbLf) & "   " & Q(CStr(k1)) & ": {" & vbLf & s2 & vbLf & "   }"
    Next k1
    NestJson = s1
End Function

' UTC time of this PC clock as ISO text (uses the Windows time-zone offset). "" if it cannot be read.
Private Function UtcStamp() As String
    On Error GoTo Fail
    Dim sh As Object, bias As Variant, d As Date
    Set sh = CreateObject("WScript.Shell")
    bias = sh.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
    d = DateAdd("n", CLng(bias), Now)
    UtcStamp = Format$(d, "yyyy-mm-dd") & "T" & Format$(d, "hh:nn:ss") & "Z"
    Exit Function
Fail:
    UtcStamp = ""
End Function

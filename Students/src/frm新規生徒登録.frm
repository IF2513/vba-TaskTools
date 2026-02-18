VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmV‹K¶“k“o˜^ 
   Caption         =   "V‹K¶“k“o˜^"
   ClientHeight    =   10845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6945
   OleObjectBlob   =   "frmV‹K¶“k“o˜^.frx":0000
   StartUpPosition =   1  'ƒI[ƒi[ ƒtƒH[ƒ€‚Ì’†‰›
End
Attribute VB_Name = "frmV‹K¶“k“o˜^"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' === ƒtƒB[ƒ‹ƒh ===
Private schCode As String
Private schTerm As String

' === ƒuƒbƒN/ƒV[ƒg’è” ===
Private Const WB_STUDENTS As String = "Students.xlsm"
Private Const SH_STUINFO  As String = "¶“kî•ñˆê——"
Private Const SH_SCHOOL   As String = "ŠwZî•ñ"
Private Const SH_TUTORS   As String = "utˆê——(from Tutors.xlsm)"   ' A=ut”Ô†, B=ut–¼i“¯ƒuƒbƒNj
Private Const SH_ASSIGN   As String = "óuE’S“–utî•ñ"            ' A:‰ïˆõ”Ô† B:–¼ C:‹³‰È D:‰È–Ú E:—j“ú F:ƒRƒ} G:ut”Ô† H:ut–¼

' === ‰Šú‰» ===
Private Sub UserForm_Initialize()
    ' === ¶“kî•ñ ===
    With cmbGrade
        .Clear
        .AddItem ""
        .AddItem "¬ŠwZ1”N": .AddItem "¬ŠwZ2”N": .AddItem "¬ŠwZ3”N": .AddItem "¬ŠwZ4”N": .AddItem "¬ŠwZ5”N": .AddItem "¬ŠwZ6”N"
        .AddItem "’†ŠwZ1”N": .AddItem "’†ŠwZ2”N": .AddItem "’†ŠwZ3”N"
        .AddItem "‚“™ŠwZ1”N": .AddItem "‚“™ŠwZ2”N": .AddItem "‚“™ŠwZ3”N"
        .AddItem "Šù‘²"
        .ListIndex = 0
    End With

    ' ŠwZ‘I‘ğiŠwZî•ñƒV[ƒg‚ÌB—ñ‚©‚çj
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Workbooks(WB_STUDENTS).Worksheets(SH_SCHOOL)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Dim lastRow As Long: lastRow = ws.Cells(ws.rowS.Count, 2).End(xlUp).Row ' B—ñ
        Dim i As Long, nm As String
        With cmbSchool
            .Clear
            .AddItem "" ' –¢‘I‘ğ
            For i = 2 To lastRow
                nm = CStr(ws.Cells(i, 2).Value)
                If Len(nm) > 0 Then .AddItem nm
            Next i
            .ListIndex = 0
        End With
    Else
        MsgBox "ŠwZî•ñƒV[ƒg‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñBŠwZˆê——‚Í‹ó‚Å‹N“®‚µ‚Ü‚·B", vbExclamation
        cmbSchool.Clear: cmbSchool.AddItem "": cmbSchool.ListIndex = 0
    End If

    ' === óuE’S“–utî•ñ ===
    ' —j“ú
    With cmbDay
        .Clear
        .AddItem "": .AddItem "Œ": .AddItem "‰Î": .AddItem "…": .AddItem "–Ø": .AddItem "‹à": .AddItem "“y"
        .ListIndex = 0
    End With
    ' ƒRƒ}
    With cmbPeriod
        .Clear
        .AddItem "": .AddItem "6": .AddItem "7": .AddItem "8"
        .ListIndex = 0
    End With
    ' ‹³‰È
    With cmbCourse
        .Clear
        .AddItem "": .AddItem "‰pŒê": .AddItem "”Šw": .AddItem "‘Œê": .AddItem "—‰È": .AddItem "Ğ‰ï": .AddItem "‘¼"
        .ListIndex = 0
    End With
    ' ‰È–Úi‹³‰È‘I‘ğ‚Å’†g‚ª“ü‚éj
    cmbSubject.Clear

    ' uti“¯ƒuƒbƒN‚Ìutˆê——‚©‚çj
    LoadTutorsToCombo

    ' ˆê——ListBox
    With lstAssignments
        .Clear
        .ColumnCount = 6  ' ‹³‰È, ‰È–Ú, —j“ú, ƒRƒ}, ut”Ô†, ut–¼
        .ColumnHeads = False
        .ColumnWidths = "45 pt;85 pt;30 pt;30 pt;0 pt;85 pt"
    End With

    ' Šù‘¶ƒf[ƒ^‚Ì“Ç‚İ‚İi‰ïˆõ”Ô†‚ª–‘O“ü—Í‚³‚ê‚Ä‚¢‚ê‚Îj
    Dim sid As String: sid = Trim$(txtID.Value)
    If Len(sid) > 0 Then LoadAssignmentsForStudent sid
End Sub

Private Sub cmbGrade_Change()
    If cmbGrade.Value = "Šù‘²" Then cmbSchool.Value = "Šù‘²"
End Sub

' ŠwZ‘I‘ğ‚ÉƒR[ƒh‚ÆŠwŠú§‚ğE‚¤
Private Sub cmbSchool_Change()
    Dim ws As Worksheet, m As Variant, schoolName As String
    schCode = "": schTerm = ""
    schoolName = Trim$(cmbSchool.Value)
    If schoolName = "" Then Exit Sub

    On Error Resume Next
    Set ws = Workbooks(WB_STUDENTS).Worksheets(SH_SCHOOL)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    m = Application.Match(schoolName, ws.Columns(2), 0) ' B—ñFŠwZ–¼
    If Not IsError(m) Then
        schCode = ws.Cells(m, 1).Value ' A—ñ = ŠwZƒR[ƒh
        schTerm = ws.Cells(m, 6).Value ' F—ñ = ŠwŠú§
    End If
End Sub

' ===== “o˜^ƒ{ƒ^ƒ“FŠî–{î•ñ ¨ óuE’S“–utî•ñ ‚ğˆêŠ‡•Û‘¶ =====
Private Sub cmdRegister_Click()
    ' ===== •K{ƒ`ƒFƒbƒN =====
    If Trim$(txtID.Value) = "" Then MsgBox "‰ïˆõ”Ô†‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B", vbExclamation: txtID.SetFocus: Exit Sub
    If Trim$(txtFamName.Value) = "" Or Trim$(txtFirstName.Value) = "" Then MsgBox "–¼i©E–¼j‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B", vbExclamation: txtFamName.SetFocus: Exit Sub
    If Trim$(txtFamKana.Value) = "" Or Trim$(txtFirstKana.Value) = "" Then MsgBox "‚Ó‚è‚ª‚Èi‚¹‚¢E‚ß‚¢j‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B", vbExclamation: txtFamKana.SetFocus: Exit Sub
    If Trim$(cmbGrade.Value) = "" Then MsgBox "Šw”N‚ğ‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢B", vbExclamation: cmbGrade.SetFocus: Exit Sub
    If Trim$(cmbSchool.Value) = "" Then MsgBox "ŠwZ‚ğ‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢B", vbExclamation: cmbSchool.SetFocus: Exit Sub

    ' ===== –¼®Œ`i”¼ŠpƒXƒy[ƒX1ŒÂj =====
    Dim nameJP As String, nameKana As String, s As String
    s = Trim$(Replace$(txtFamName.Value, "@", " ")) & " " & Trim$(Replace$(txtFirstName.Value, "@", " "))
    nameJP = Trim$(Replace$(s, "@", " ")): Do While InStr(nameJP, "  ") > 0: nameJP = Replace$(nameJP, "  ", " "): Loop
    s = Trim$(Replace$(txtFamKana.Value, "@", " ")) & " " & Trim$(Replace$(txtFirstKana.Value, "@", " "))
    nameKana = Trim$(Replace$(s, "@", " ")): Do While InStr(nameKana, "  ") > 0: nameKana = Replace$(nameKana, "  ", " "): Loop

    ' ===== ¶“kî•ñˆê——‚Ö•Û‘¶iƒV[ƒg‘¶İ‚Í‘O’ñj =====
    Dim ws As Worksheet
    Set ws = Workbooks(WB_STUDENTS).Worksheets(SH_STUINFO)

    Dim keyId As String: keyId = Trim$(txtID.Value)
    Dim found As Range
    Set found = ws.Columns(1).Find(What:=keyId, LookAt:=xlWhole, MatchCase:=False)

    If Not found Is Nothing Then
        If MsgBox("‚±‚Ì‰ïˆõ”Ô†‚ÍŠù‚É“o˜^‚³‚ê‚Ä‚¢‚Ü‚·Bã‘‚«isíœ¨––”ö‚ÉÄ“o˜^j‚µ‚Ü‚·‚©H", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        ws.rowS(found.Row).Delete         ' š‚±‚±‚ğC³irowS¨Rowsj
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row + 1   ' š‚±‚±‚àC³

    ws.Cells(lastRow, 1).Value = keyId
    ws.Cells(lastRow, 2).Value = nameJP
    ws.Cells(lastRow, 3).Value = nameKana
    ws.Cells(lastRow, 4).Value = schCode
    ws.Cells(lastRow, 5).Value = Trim$(cmbSchool.Value)
    ws.Cells(lastRow, 6).Value = Trim$(cmbGrade.Value)
    ws.Cells(lastRow, 7).Value = schTerm

    ' ===== óuE’S“–utî•ñ‚ğ‘‚«o‚µ =====
    SaveTutorAssignments keyId, nameJP

    ' ===== ¶“kî•ñˆê——(I`N)‚Ö’S“–ut‚Ì’Zkƒ‰ƒxƒ‹‚ğ”½‰f =====
    UpdateTutorSummaryForStudent keyId

    MsgBox "“o˜^‚µ‚Ü‚µ‚½BiŠî–{î•ñ{óuE’S“–utî•ñj", vbInformation
End Sub


' === ”Ä—p ===
Private Function GetSheetIfExists(wb As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetSheetIfExists = wb.Worksheets(name)
    On Error GoTo 0
End Function

Private Function EnsureAssignSheet() As Worksheet
    Dim wb As Workbook: Set wb = Workbooks(WB_STUDENTS)
    Dim ws As Worksheet: Set ws = GetSheetIfExists(wb, SH_ASSIGN)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.name = SH_ASSIGN
        ws.Range("A1:H1").Value = Array("‰ïˆõ”Ô†", "‰ïˆõ–¼", "‹³‰È", "‰È–Ú", "—j“ú", "ƒRƒ}", "ut”Ô†", "ut–¼")
    End If
    Set EnsureAssignSheet = ws
End Function

' === ‹³‰È¨‰È–Ú‚Ìƒ}ƒbƒv ===
Private Function SubjectMap() As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    m("‰pŒê") = Array("¬Šw‰pŒê", "’†Šw‰pŒê", "OC1", "ƒ‰ƒCƒeƒBƒ“ƒO", "ƒŠ[ƒfƒBƒ“ƒO", "‰pŒê1", "‰pŒê2", "‰pŒê’·•¶", "‰p•¶–@", "óŒ±‰pŒê", "‚Z‰pŒê")
    m("”Šw") = Array("¬ŠwZ”", "’†Šw”Šw", "}Œ`", "”—Ê", "”Šw1A", "”Šw2B", "”Šw3C", "”Šw1", "”Šw2", "”Šw3", "”ŠwA", "”ŠwB", "”ŠwC", "”ŠwŠî‘b", "óŒ±”Šw", "‚Z”Šw")
    m("‘Œê") = Array("¬Šw‘Œê", "¬Šwì•¶", "’†Šw‘Œê", "’†Šwì•¶", "Œ»‘ã•¶", "ŒÃ“T", "ŒÃ“T•¶–@", "Š¿•¶", "‘Œê‘‡", "‘Œê•\Œ»1", "‘Œê•\Œ»2", "¬˜_•¶", "óŒ±‘Œê", "‚Z‘Œê")
    m("—‰È") = Array("¬Šw—‰È", "’†Šw—‰È", "•¨—Šî‘b", "•¨—", "¶•¨Šî‘b", "¶•¨", "‰»ŠwŠî‘b", "‰»Šw", "’nŠwŠî‘b", "’nŠw", "—‰È‘‡A", "—‰È‘‡B", "—‰ÈŠî‘b", "óŒ±—‰È", "‚Z—‰È")
    m("Ğ‰ï") = Array("¬ŠwĞ‰ï", "’†ŠwĞ‰ï", "’†Šw—ğj", "’†Šw’n—", "“ú–{jA", "“ú–{jB", "¢ŠEjA", "¢ŠEjB", "’n—A", "’n—B", "­¡EŒoÏ", "Œ»‘ãĞ‰ï", "óŒ±Ğ‰ï", "‚ZĞ‰ï")
    m("‘¼") = Array("¬—¬Ğ", "¬Z¬‘", "¬‰p¬Z", "’†‰p’†‘", "’†‰p’†”", "’†—’†Ğ", "’†”’†‘", "‚‰p‚”", "‚‰p‚‘", "‚‰p‚—", "‚‰p‚Ğ", "‚”‚‘", "‚”‚—", "‚‘‚Ğ", "‚R‚W", "—””Šw1", "—”‰»Šw", "—”•¨—", "‰Æ’ëŠî‘b", "î•ñA", "‘S”Ê")
    Set SubjectMap = m
End Function

' === ‹³‰È‘I‘ğ ¨ ‰È–Úƒvƒ‹ƒ_ƒEƒ“‚ğÄ\¬ ===
Private Sub cmbCourse_Change()
    Dim m As Object: Set m = SubjectMap()
    Dim key As String: key = Trim$(cmbCourse.Value)
    Dim i As Long
    With cmbSubject
        .Clear
        If Len(key) = 0 Then Exit Sub
        If m.Exists(key) Then
            For i = LBound(m(key)) To UBound(m(key))
                .AddItem m(key)(i)
            Next
        End If
        If .ListCount > 0 Then .ListIndex = 0
    End With
End Sub

' === utˆê——‚Ì“Çi“¯ƒuƒbƒNj ===
Private Sub LoadTutorsToCombo()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Workbooks(WB_STUDENTS).Worksheets(SH_TUTORS)
    On Error GoTo 0

    With cmbTeacher
        .Clear
        .ColumnCount = 2            ' 0=ut”Ô†, 1=ut–¼
        .BoundColumn = 2            ' Value‚Íut–¼
        .ColumnWidths = "0 pt;120 pt" ' ”Ô†‚Í”ñ•\¦E–¼‚Ì‚İŒ©‚¹‚é

        If ws Is Nothing Then
            .AddItem: .List(.ListCount - 1, 0) = "": .List(.ListCount - 1, 1) = "iutˆê——ƒV[ƒg‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñj"
            .ListIndex = -1
            MsgBox "“¯ƒuƒbƒN“à‚ÌƒV[ƒg '" & SH_TUTORS & "' ‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñB", vbExclamation
            Exit Sub
        End If

        Dim lastR As Long, r As Long, tid As String, tnm As String, added As Long
        lastR = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
        For r = 2 To lastR
            tid = Trim$(CStr(ws.Cells(r, 1).Value))   ' A=ut”Ô†
            tnm = Trim$(CStr(ws.Cells(r, 2).Value))   ' B=ut–¼
            If Len(tnm) > 0 Then
                .AddItem
                .List(.ListCount - 1, 0) = tid
                .List(.ListCount - 1, 1) = tnm
                added = added + 1
            End If
        Next

        If added = 0 Then
            .AddItem: .List(.ListCount - 1, 0) = "": .List(.ListCount - 1, 1) = "iutƒf[ƒ^‚ª‹ó‚Å‚·j"
        End If

        .ListIndex = -1  ' ƒfƒtƒHƒ‹ƒg–¢‘I‘ğ
        .Value = ""
    End With
End Sub

' === ˆê——“à‚ÌŠ®‘Sˆê’vd•¡i‹³‰È/‰È–Ú/—j/ƒRƒ}/ut”Ô†jƒ`ƒFƒbƒN ===
Private Function ExistsExactRow(ByVal course As String, ByVal subj As String, ByVal dayW As String, ByVal period As String, ByVal tid As String) As Boolean
    Dim i As Long
    For i = 0 To lstAssignments.ListCount - 1
        If lstAssignments.List(i, 0) = course _
        And lstAssignments.List(i, 1) = subj _
        And lstAssignments.List(i, 2) = dayW _
        And lstAssignments.List(i, 3) = period _
        And lstAssignments.List(i, 4) = tid Then
            ExistsExactRow = True
            Exit Function
        End If
    Next
End Function

' === óus ’Ç‰Á ===
Private Sub cmdAddRow_Click()
    Dim course As String: course = Trim$(cmbCourse.Value)
    Dim subj   As String: subj = Trim$(cmbSubject.Value)
    Dim dayW   As String: dayW = Trim$(cmbDay.Value)
    Dim period As String: period = Trim$(cmbPeriod.Value)
    If course = "" Then MsgBox "‹³‰È‚ğ‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢B", vbExclamation: cmbCourse.SetFocus: Exit Sub
    If subj = "" Then MsgBox "‰È–Ú‚ğ‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢B", vbExclamation: cmbSubject.SetFocus: Exit Sub
    If dayW = "" Then MsgBox "—j“ú‚ğ‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢B", vbExclamation: cmbDay.SetFocus: Exit Sub
    If period = "" Then MsgBox "ƒRƒ}‚ğ‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢B", vbExclamation: cmbPeriod.SetFocus: Exit Sub

    Dim tid As String, tname As String
    If cmbTeacher.ListIndex >= 0 Then
        tid = CStr(cmbTeacher.List(cmbTeacher.ListIndex, 0))
        tname = CStr(cmbTeacher.List(cmbTeacher.ListIndex, 1))
    Else
        tid = "": tname = "" ' ut–¢‘I‘ğOK
    End If

    If ExistsExactRow(course, subj, dayW, period, tid) Then
        MsgBox "“¯ˆê‚Ìsi‹³‰È/‰È–Ú/—j“ú/ƒRƒ}/utj‚ªŠù‚É‚ ‚è‚Ü‚·B", vbInformation
        Exit Sub
    End If

    lstAssignments.AddItem
    lstAssignments.List(lstAssignments.ListCount - 1, 0) = course
    lstAssignments.List(lstAssignments.ListCount - 1, 1) = subj
    lstAssignments.List(lstAssignments.ListCount - 1, 2) = dayW
    lstAssignments.List(lstAssignments.ListCount - 1, 3) = period
    lstAssignments.List(lstAssignments.ListCount - 1, 4) = tid
    lstAssignments.List(lstAssignments.ListCount - 1, 5) = tname
End Sub

' === ˆê——ƒNƒŠƒbƒN ¨ •ÒW—“‚Ö”½‰f ===
Private Sub lstAssignments_Click()
    Dim i As Long: i = lstAssignments.ListIndex
    If i < 0 Then Exit Sub

    cmbCourse.Value = lstAssignments.List(i, 0)
    Call cmbCourse_Change
    cmbSubject.Value = lstAssignments.List(i, 1)
    cmbDay.Value = lstAssignments.List(i, 2)
    cmbPeriod.Value = lstAssignments.List(i, 3)

    Dim tid As String: tid = lstAssignments.List(i, 4)
    Dim idx As Long, hit As Boolean
    For idx = 0 To cmbTeacher.ListCount - 1
        If CStr(cmbTeacher.List(idx, 0)) = tid Then
            cmbTeacher.ListIndex = idx: hit = True: Exit For
        End If
    Next
    If Not hit Then cmbTeacher.ListIndex = -1: cmbTeacher.Value = ""
End Sub

' === óus XV ===
Private Sub cmdUpdateRow_Click()
    Dim i As Long: i = lstAssignments.ListIndex
    If i < 0 Then MsgBox "XV‚·‚és‚ğ‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢B", vbExclamation: Exit Sub

    Dim course As String: course = Trim$(cmbCourse.Value)
    Dim subj   As String: subj = Trim$(cmbSubject.Value)
    Dim dayW   As String: dayW = Trim$(cmbDay.Value)
    Dim period As String: period = Trim$(cmbPeriod.Value)
    If course = "" Or subj = "" Or dayW = "" Or period = "" Then
        MsgBox "‹³‰ÈE‰È–ÚE—j“úEƒRƒ}‚Í•K{‚Å‚·B", vbExclamation: Exit Sub
    End If

    Dim tid As String, tname As String
    If cmbTeacher.ListIndex >= 0 Then
        tid = CStr(cmbTeacher.List(cmbTeacher.ListIndex, 0))
        tname = CStr(cmbTeacher.List(cmbTeacher.ListIndex, 1))
    Else
        tid = "": tname = ""
    End If

    Dim j As Long
    For j = 0 To lstAssignments.ListCount - 1
        If j <> i Then
            If lstAssignments.List(j, 0) = course _
            And lstAssignments.List(j, 1) = subj _
            And lstAssignments.List(j, 2) = dayW _
            And lstAssignments.List(j, 3) = period _
            And lstAssignments.List(j, 4) = tid Then
                MsgBox "“¯ˆê‚Ìsi‹³‰È/‰È–Ú/—j“ú/ƒRƒ}/utj‚ªŠù‚É‚ ‚è‚Ü‚·B", vbInformation
                Exit Sub
            End If
        End If
    Next

    lstAssignments.List(i, 0) = course
    lstAssignments.List(i, 1) = subj
    lstAssignments.List(i, 2) = dayW
    lstAssignments.List(i, 3) = period
    lstAssignments.List(i, 4) = tid
    lstAssignments.List(i, 5) = tname
End Sub

' === óus íœ ===
Private Sub cmdRemoveRow_Click()
    Dim i As Long: i = lstAssignments.ListIndex
    If i < 0 Then MsgBox "íœ‚·‚és‚ğ‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢B", vbExclamation: Exit Sub
    lstAssignments.RemoveItem i
End Sub

' === “ü—Í—“ƒNƒŠƒA ===
Private Sub cmdClearRow_Click()
    cmbCourse.ListIndex = 0
    cmbSubject.Clear
    cmbDay.ListIndex = 0
    cmbPeriod.ListIndex = 0
    cmbTeacher.ListIndex = -1
    cmbTeacher.Value = ""
End Sub

' === óuE’S“–utî•ñ‚Ì•Û‘¶iA`H—ñj ===
Private Sub SaveTutorAssignments(ByVal studentId As String, ByVal studentName As String)
    Dim ws As Worksheet: Set ws = EnsureAssignSheet()

    ' “–ŠY‰ïˆõ‚ÌŠù‘¶s‚ğíœi‰º‚©‚çj
    Dim lastR As Long, r As Long
    lastR = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
    For r = lastR To 2 Step -1
        If CStr(ws.Cells(r, 1).Value) = studentId Then ws.rowS(r).Delete
    Next

    ' ListBox‚Ì‘Ss‚ğ‘‚«o‚µiut–¢‘I‘ğ‚Í‹ó‚Ì‚Ü‚Üj
    Dim i As Long, nxt As Long
    For i = 0 To lstAssignments.ListCount - 1
        nxt = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row + 1
        ws.Cells(nxt, 1).Value = studentId
        ws.Cells(nxt, 2).Value = studentName
        ws.Cells(nxt, 3).Value = lstAssignments.List(i, 0) ' ‹³‰È
        ws.Cells(nxt, 4).Value = lstAssignments.List(i, 1) ' ‰È–Ú
        ws.Cells(nxt, 5).Value = lstAssignments.List(i, 2) ' —j“ú
        ws.Cells(nxt, 6).Value = lstAssignments.List(i, 3) ' ƒRƒ}
        ws.Cells(nxt, 7).Value = lstAssignments.List(i, 4) ' ut”Ô†
        ws.Cells(nxt, 8).Value = lstAssignments.List(i, 5) ' ut–¼
    Next
End Sub

' === ’S“–utî•ñ‚Ì“Çi•ÒW—pj ===
Private Sub LoadAssignmentsForStudent(ByVal studentId As String)
    Dim ws As Worksheet: Set ws = GetSheetIfExists(Workbooks(WB_STUDENTS), SH_ASSIGN)
    If ws Is Nothing Then Exit Sub

    Dim lastR As Long, r As Long
    lastR = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
    lstAssignments.Clear
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value) = studentId Then
            lstAssignments.AddItem
            lstAssignments.List(lstAssignments.ListCount - 1, 0) = CStr(ws.Cells(r, 3).Value) ' ‹³‰È
            lstAssignments.List(lstAssignments.ListCount - 1, 1) = CStr(ws.Cells(r, 4).Value) ' ‰È–Ú
            lstAssignments.List(lstAssignments.ListCount - 1, 2) = CStr(ws.Cells(r, 5).Value) ' —j“ú
            lstAssignments.List(lstAssignments.ListCount - 1, 3) = CStr(ws.Cells(r, 6).Value) ' ƒRƒ}
            lstAssignments.List(lstAssignments.ListCount - 1, 4) = CStr(ws.Cells(r, 7).Value) ' ut”Ô†
            lstAssignments.List(lstAssignments.ListCount - 1, 5) = CStr(ws.Cells(r, 8).Value) ' ut–¼
        End If
    Next
End Sub

' •Â‚¶‚é
Private Sub cmdClose_Click()
    Unload Me
End Sub



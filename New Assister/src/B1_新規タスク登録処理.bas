Attribute VB_Name = "B1_新規タスク登録処理"
Option Explicit

'==== シート名の定数 ===='
Private Const SH_TASKLIST As String = "TaskList"
Private Const SH_STATUS   As String = "TaskStatus"
Private Const SH_STUDENT  As String = "Students"

'==== フォームを開く（リボン/ボタン/ショートカットから呼ぶ用） ===='
Public Sub 新規タスク登録フォームを開く()
    frm新規タスク登録.Show
End Sub

'==== 次のTaskIDを採番（T001, T002, ...） ===='
Public Function GetNextTaskID() As String
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim maxNum As Long, v As Variant, s As String, n As Long

    Set ws = ThisWorkbook.Sheets(SH_TASKLIST)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    maxNum = 0

    For i = 2 To lastRow
        v = ws.Cells(i, 1).Value
        s = CStr(v)
        If Len(s) >= 2 And UCase$(Left$(s, 1)) = "T" Then
            On Error Resume Next
            n = CLng(Mid$(s, 2))
            On Error GoTo 0
            If n > maxNum Then maxNum = n
        End If
    Next i

    GetNextTaskID = "T" & Format$(maxNum + 1, "000")
End Function

'==== TaskList に1行追記し、行番号を返す ===='
Public Function AppendTaskRow(taskId As String, taskName As String, _
                              startDate As Variant, dueDate As Variant, endDate As Variant, _
                              commentText As String, _
                              condGrade As String, condSchool As String, condTerm As String, condDiv As String) As Long
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets(SH_TASKLIST)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, "A").Value = taskId
    ws.Cells(r, "B").Value = taskName
    ws.Cells(r, "C").Value = IIf(IsDate(startDate), CDate(startDate), "")
    ws.Cells(r, "D").Value = IIf(IsDate(dueDate), CDate(dueDate), "")
    ws.Cells(r, "E").Value = IIf(IsDate(endDate), CDate(endDate), "")
    ws.Cells(r, "F").Value = commentText
    ws.Cells(r, "G").Value = condGrade
    ws.Cells(r, "H").Value = condSchool
    ws.Cells(r, "I").Value = condTerm
    ws.Cells(r, "J").Value = condDiv

    AppendTaskRow = r
End Function

'==== 条件のMultiSelect結果（ListBox）を「,」連結にする ===='
Public Function JoinSelected(listBox As MSForms.listBox) As String
    Dim i As Long, s As String
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            s = s & listBox.List(i) & ","
        End If
    Next
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    JoinSelected = s
End Function

'==== 文字列→日付の安全変換（空欄OK、無効はEmptyを返す） ===='
Public Function ParseDateOrEmpty(text As String) As Variant
    Dim t As String: t = Trim$(text)
    t = Replace(t, "年", "/")
    t = Replace(t, "月", "/")
    t = Replace(t, "日", "")
    t = Replace(t, "-", "/")
    If Len(t) = 0 Then
        ParseDateOrEmpty = Empty
    ElseIf IsDate(t) Then
        ParseDateOrEmpty = CDate(t)
    Else
        ParseDateOrEmpty = Empty
    End If
End Function

'==== 新規タスク展開（TaskStatusへ対象者のレコードを付与） ===='
Public Sub ExpandTaskToStatus(newTaskID As String)
    ' 既存の「初期化処理.AssignTargetsToStatus」を流用するのが最短
    ' ここでは新タスクだけ展開したいケースに合わせ軽量版を実装
    Dim wsTask As Worksheet, wsStu As Worksheet, wsStatus As Worksheet
    Dim rowTask As Variant, lastStu As Long, j As Long
    Dim today As Date: today = Date

    Set wsTask = ThisWorkbook.Sheets(SH_TASKLIST)
    Set wsStu = ThisWorkbook.Sheets(SH_STUDENT)
    Set wsStatus = ThisWorkbook.Sheets(SH_STATUS)

    rowTask = Application.Match(newTaskID, wsTask.Columns(1), 0)
    If IsError(rowTask) Then Exit Sub

    ' 期間外はスキップ
    If Not IsWithinPublicationRange(wsTask.Cells(rowTask, "C").Value, wsTask.Cells(rowTask, "E").Value, today) Then Exit Sub

    lastStu = wsStu.Cells(wsStu.Rows.Count, 1).End(xlUp).Row
    For j = 2 To lastStu
        If IsTargetStudent(wsStu.Rows(j), wsTask.Rows(rowTask)) Then
            EnsureStatusRow wsStatus, CStr(wsStu.Cells(j, "A").Value2), newTaskID
        End If
    Next j
End Sub






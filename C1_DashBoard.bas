Attribute VB_Name = "C1_DashBoard"
Option Explicit

' ===== シート・レイアウト前提 =====
Private Const SH_STATUS As String = "TaskStatus"
Private Const SH_DASH   As String = "トップページ"
Private Const SH_TASKLIST As String = "TaskList"


Private Const FIRST_TASK_COL As Long = 6    ' F列からタスク
Private Const ROW_TASK_ID    As Long = 1    ' 1行目 = Task ID（空なら走査終了）
Private Const ROW_TASK_NAME  As Long = 2    ' 2行目 = タスク名
Private Const ROW_COMMENT    As Long = 3    ' 3行目 = コメント
Private Const ROW_DEADLINE   As Long = 4    ' 4行目 = 締切
Private Const ROW_FIRST_STU  As Long = 6    ' 6行目 = 生徒開始

' TaskList のF列（コメント）を TaskID で引く
Private Function GetTaskComment(ByVal taskId As String) As String
    Dim ws As Worksheet, f As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_TASKLIST)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    ' IDの位置は列固定とせず、シート全体から完全一致検索
    Set f = ws.UsedRange.Find(What:=taskId, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then
        GetTaskComment = CStr(ws.Cells(f.Row, "F").Value) ' ★ F列=コメント
    Else
        GetTaskComment = ""
    End If
End Function


' ===== 更新：ActiveX ListBox に5列で表示 =====
Public Sub NA_RefreshDashboard_ActiveX()
    Dim wsS As Worksheet, wsD As Worksheet
    Set wsS = ThisWorkbook.Worksheets(SH_STATUS)
    Set wsD = ThisWorkbook.Worksheets(SH_DASH)

    Dim lastRow As Long
    lastRow = wsS.Cells(wsS.Rows.Count, "A").End(xlUp).Row
    If lastRow < ROW_FIRST_STU Then
        MsgBox "TaskStatus に生徒データがありません。", vbExclamation
        Exit Sub
    End If

    ' 一旦全タスク数をざっくり見積もって配列確保（後で ReDim Preserve）
    Dim maxCols As Long
    maxCols = wsS.Cells(ROW_TASK_ID, wsS.Columns.Count).End(xlToLeft).Column - FIRST_TASK_COL + 1
    If maxCols < 1 Then maxCols = 1

    Dim buf() As Variant, outN As Long
    ReDim buf(1 To maxCols, 1 To 5) ' 5列固定

    Dim c As Long: c = FIRST_TASK_COL
    Do
        Dim taskId As String
        taskId = Trim$(CStr(wsS.Cells(ROW_TASK_ID, c).Value))
        If Len(taskId) = 0 Then Exit Do  ' ★ 1行目が空→終了

        ' 完了率計算：対象=「-」以外、完了=日付
        Dim rng As Range
        Set rng = wsS.Range(wsS.Cells(ROW_FIRST_STU, c), wsS.Cells(lastRow, c))

        Dim targetCnt As Long, doneCnt As Long
        targetCnt = rng.Rows.Count - Application.WorksheetFunction.CountIf(rng, "-") ' 「-」以外（空白含む）
        doneCnt = Application.WorksheetFunction.Count(rng)                            ' 数値=日付の個数

        Dim rate As Double
        If targetCnt > 0 Then rate = doneCnt / targetCnt Else rate = 0


        ' 実施中のみ（対象>0 かつ 完了<対象）
        If targetCnt > 0 And doneCnt < targetCnt Then
            Dim taskName As String, comment As String, deadline As Variant
            taskName = wsS.Cells(ROW_TASK_NAME, c).Value
            comment = GetTaskComment(taskId)
            deadline = wsS.Cells(ROW_DEADLINE, c).Value
            rate = CDbl(doneCnt) / CDbl(targetCnt)

            outN = outN + 1
            If outN > UBound(buf, 1) Then ReDim Preserve buf(1 To outN, 1 To 5)

            buf(outN, 1) = taskId
            buf(outN, 2) = taskName
            buf(outN, 3) = IIf(IsDate(deadline), Format$(deadline, "m/d"), "")
            buf(outN, 4) = Format$(rate, "0.0%")
            buf(outN, 5) = comment
        End If

        c = c + 1
    Loop

    Dim lb As MSForms.listBox
    On Error Resume Next
    Set lb = wsD.OLEObjects("lstTasksAx").Object
    On Error GoTo 0
    If lb Is Nothing Then
        MsgBox "ActiveX リストボックス 'lstTasksAx' が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' ListBox 設定
    With lb
        .Clear
        .ColumnCount = 5
        ' 文字幅に合わせて調整（お好みで）
        .ColumnWidths = "40 pt;170 pt;70 pt;70 pt;200 pt"
        .IntegralHeight = False
        .MultiSelect = fmMultiSelectSingle
    End With

    If outN = 0 Then Exit Sub

    ' 0-based 2次元配列に詰め替えて一括代入（List プロパティ）
    Dim dataArr() As Variant, i As Long, j As Long
    ReDim dataArr(0 To outN - 1, 0 To 4)
    For i = 0 To outN - 1
        For j = 0 To 4
            dataArr(i, j) = buf(i + 1, j + 1)
        Next
    Next
    lb.List = dataArr
    lb.ListIndex = -1 ' 未選択
End Sub

' ===== 出力：選択行の Task ID を使って新規ブックに成形 =====
Public Sub NA_ExportSelectedTask_ActiveX()
    Dim wsS As Worksheet, wsD As Worksheet
    Set wsS = ThisWorkbook.Worksheets(SH_STATUS)
    Set wsD = ThisWorkbook.Worksheets(SH_DASH)

    Dim lb As MSForms.listBox
    On Error Resume Next
    Set lb = wsD.OLEObjects("lstTasksAx").Object
    On Error GoTo 0
    If lb Is Nothing Then
        MsgBox "ActiveX リストボックス 'lstTasksAx' が見つかりません。", vbExclamation
        Exit Sub
    End If
    If lb.ListIndex < 0 Then
        MsgBox "リストからタスクを選択してください。", vbInformation
        Exit Sub
    End If

    Dim taskId As String, taskName As String
    taskId = CStr(lb.List(lb.ListIndex, 0))
    taskName = CStr(lb.List(lb.ListIndex, 1))

    ' TaskStatus 内で TaskID の列を特定（1行目一致・空で打ち切り）
    Dim lastCol As Long: lastCol = wsS.Cells(ROW_TASK_ID, wsS.Columns.Count).End(xlToLeft).Column
    Dim c As Long, hitCol As Long: hitCol = 0
    For c = FIRST_TASK_COL To lastCol
        Dim idHere As String: idHere = CStr(wsS.Cells(ROW_TASK_ID, c).Value)
        If Len(idHere) = 0 Then Exit For
        If StrComp(idHere, taskId, vbTextCompare) = 0 Then
            hitCol = c: Exit For
        End If
    Next
    If hitCol = 0 Then
        MsgBox "TaskStatus 上で Task ID が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' === 新規ブックへ整形 ===
    Dim lastRow As Long: lastRow = wsS.Cells(wsS.Rows.Count, "A").End(xlUp).Row
    Dim rowsCnt As Long: rowsCnt = lastRow - ROW_FIRST_STU + 1

    Dim wb As Workbook, ws As Worksheet
    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "完了表"

    ' ヘッダ
    ws.Range("A1:F1").Value = Array("Task ID", "タスク名", "会員番号", "氏名", "学年", "完了")
    ' 本文
    ws.Range("A2").Resize(rowsCnt, 1).Value = taskId
    ws.Range("B2").Resize(rowsCnt, 1).Value = taskName
    ws.Range("C2").Resize(rowsCnt, 1).Value = wsS.Range("A" & ROW_FIRST_STU).Resize(rowsCnt, 1).Value ' 会員番号
    ws.Range("D2").Resize(rowsCnt, 1).Value = wsS.Range("C" & ROW_FIRST_STU).Resize(rowsCnt, 1).Value ' 氏名
    ws.Range("E2").Resize(rowsCnt, 1).Value = wsS.Range("B" & ROW_FIRST_STU).Resize(rowsCnt, 1).Value ' 学年
    ws.Range("F2").Resize(rowsCnt, 1).Value = wsS.Cells(ROW_FIRST_STU, hitCol).Resize(rowsCnt, 1).Value ' 完了セル

    ' 体裁
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With
    ws.Columns("A:F").AutoFit
End Sub

' ====== 補助 ======
Private Function IsCompletedDate(ByVal v As Variant) As Boolean
    If IsDate(v) Then
        IsCompletedDate = True
    ElseIf (VarType(v) = vbDouble Or VarType(v) = vbDate) And CDbl(v) > 0 Then
        IsCompletedDate = True
    Else
        IsCompletedDate = False
    End If
End Function


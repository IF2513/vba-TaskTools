Attribute VB_Name = "Module1"
Option Explicit

' ===== シート名・コントロール名 =====
Private Const SH_TOP As String = "トップページ"           ' ListBoxがあるシート
Private Const LB_NAME As String = "lstTasksAx"            ' ActiveX ListBox 名
Private Const SH_STATUS As String = "TaskStatus"          ' 進捗の元データ
Private Const SH_TASKLIST As String = "TaskList"          ' コメント(F列)の参照元
Private Const SH_FORMAT As String = "完了状況出力Format"   ' ひな型シート名

' ===== TaskStatus レイアウト =====
Private Const COL_ID As Long = 1      ' A: 会員番号
Private Const COL_NAME As Long = 3    ' C: 氏名
Private Const COL_TUTOR As Long = 4   ' D: 担当講師（連結済み）
Private Const FIRST_TASK_COL As Long = 6 ' F: タスク開始列

Private Const ROW_TASKID As Long = 1
Private Const ROW_TASKNAME As Long = 2
Private Const ROW_DEADLINE As Long = 4
Private Const ROW_STU_FIRST As Long = 6

' ===== メイン入口 =====
Public Sub ExportSelectedTask_CompletionSheet()
    Dim wsTop As Worksheet: Set wsTop = ThisWorkbook.Worksheets(SH_TOP)
    Dim lb As MSForms.listBox
    Set lb = wsTop.OLEObjects(LB_NAME).Object

    If lb.ListCount = 0 Or lb.ListIndex < 0 Then
        MsgBox "タスクを1つ選択してください（" & SH_TOP & " の " & LB_NAME & "）。", vbExclamation
        Exit Sub
    End If

    Dim taskId As String
    taskId = CStr(lb.List(lb.ListIndex, 0))  ' 先頭列=TaskID を想定
    If Len(Trim$(taskId)) = 0 Then
        MsgBox "選択中の行に TaskID がありません。", vbExclamation
        Exit Sub
    End If

    Call ExportTask_ToNewBook(taskId)
End Sub

' ====== 実体：TaskID を指定して新規ブックに出力 ======
Private Sub ExportTask_ToNewBook(ByVal taskId As String)
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Worksheets(SH_STATUS)

    ' 1) Task列の特定
    Dim cTask As Long: cTask = FindTaskCol(wsS, taskId)
    If cTask = 0 Then
        MsgBox "TaskStatusで TaskID [" & taskId & "] が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 2) ヘッダ情報
    Dim taskName As String: taskName = CStr(wsS.Cells(ROW_TASKNAME, cTask).Value)
    Dim deadline As Variant: deadline = wsS.Cells(ROW_DEADLINE, cTask).Value
    Dim comment As String: comment = GetTaskComment(taskId)

' 3) 対象データを配列化（"-" は対象外）
Dim lastR As Long: lastR = wsS.Cells(wsS.Rows.Count, "A").End(xlUp).Row

' 3-1. 件数カウント
Dim cnt As Long, r As Long, v As Variant
cnt = 0
For r = ROW_STU_FIRST To lastR
    v = wsS.Cells(r, cTask).Value
    If Not IsDashOnly(v) Then
        cnt = cnt + 1
    End If
Next

If cnt = 0 Then
    MsgBox "対象者がいません（全員「-」でした）。", vbInformation
    Exit Sub
End If

' 3-2. ちょうどのサイズで一括確保（列=4, 行=cnt）
Dim buf() As Variant
ReDim buf(1 To 4, 1 To cnt)

' 3-3. 実データを詰める
Dim i As Long: i = 1
For r = ROW_STU_FIRST To lastR
    v = wsS.Cells(r, cTask).Value
    If Not IsDashOnly(v) Then
        buf(1, i) = wsS.Cells(r, COL_ID).Value
        buf(2, i) = wsS.Cells(r, COL_NAME).Value
        buf(3, i) = wsS.Cells(r, COL_TUTOR).Value
        buf(4, i) = IIf(IsCompletedDate(v), "済", "")
        i = i + 1
    End If
Next

Dim n As Long: n = cnt   ' ← 以降のロジックが使う総件数

    ' 4) 新規ブックにテンプレをコピー
    Dim wsFmt As Worksheet: Set wsFmt = ThisWorkbook.Worksheets(SH_FORMAT)
    wsFmt.Copy
    Dim wbOut As Workbook: Set wbOut = ActiveWorkbook
    Dim wsOut As Worksheet: Set wsOut = wbOut.Worksheets(1)
    wsOut.Name = Left(taskId, 31)

    ' 5) ヘッダ埋め込み
    wsOut.Range("A1").Value = taskId
    wsOut.Range("B1").Value = taskName
    wsOut.Range("B2").Value = comment
    If IsDate(deadline) Then
        wsOut.Range("M4").Value = CDate(deadline)
        wsOut.Range("M4").NumberFormatLocal = "yyyy/m/d"
    Else
        wsOut.Range("M4").ClearContents
    End If
    wsOut.Range("M58").Value = Now
    wsOut.Range("M58").NumberFormatLocal = "yyyy/m/d"

    ' 6) パネル定義（1シート=2ページ分）
    Dim panels1(1 To 3) As Range  ' 1ページ目：各50
    Set panels1(1) = wsOut.Range("A7:D56")
    Set panels1(2) = wsOut.Range("F7:I56")
    Set panels1(3) = wsOut.Range("K7:N56")

    Dim panels2(1 To 3) As Range  ' 2ページ目：各52
    Set panels2(1) = wsOut.Range("A62:D113")
    Set panels2(2) = wsOut.Range("F62:I113")
    Set panels2(3) = wsOut.Range("K62:N113")

    ' 7) パネルに順次流し込み（足りなければ次のテンプレを追加）
    Dim idx As Long: idx = 1
    Call FillPanels(buf, idx, panels1)
    If idx <= n Then
        ' 2ページ目も使用
        Call FillPanels(buf, idx, panels2)
        wsOut.Range("M115").Value = Now
        wsOut.Range("M115").NumberFormatLocal = "yyyy/m/d"
    End If

    ' さらにデータが残るなら、テンプレを増殖して続きに出力
    Do While idx <= n
        wsFmt.Copy After:=wbOut.Worksheets(wbOut.Worksheets.Count)
        Dim wsNext As Worksheet
        Set wsNext = wbOut.Worksheets(wbOut.Worksheets.Count)
        wsNext.Name = Left(taskId & "_" & (wbOut.Worksheets.Count), 31)

        ' ヘッダ
        wsNext.Range("A1").Value = taskId
        wsNext.Range("B1").Value = taskName
        wsNext.Range("B2").Value = comment
        If IsDate(deadline) Then
            wsNext.Range("M4").Value = CDate(deadline)
            wsNext.Range("M4").NumberFormatLocal = "yyyy/m/d"
        Else
            wsNext.Range("M4").ClearContents
        End If
        wsNext.Range("M58").Value = Now
        wsNext.Range("M58").NumberFormatLocal = "yyyy/m/d h:mm"

        ' パネル参照を差し替えて再利用
        Set panels1(1) = wsNext.Range("A7:D56")
        Set panels1(2) = wsNext.Range("F7:I56")
        Set panels1(3) = wsNext.Range("K7:N56")
        Set panels2(1) = wsNext.Range("A62:D113")
        Set panels2(2) = wsNext.Range("F62:I113")
        Set panels2(3) = wsNext.Range("K62:N113")

        Call FillPanels(buf, idx, panels1)
        If idx <= n Then
            Call FillPanels(buf, idx, panels2)
            wsNext.Range("M115").Value = Now
            wsNext.Range("M115").NumberFormatLocal = "yyyy/m/d h:mm"
        End If
    Loop

    MsgBox "出力が完了しました。", vbInformation
End Sub

' ====== 指定されたパネル群へ buf(idx～) を順に流しこむ ======
Private Sub FillPanels(ByRef buf() As Variant, ByRef idx As Long, ByRef panels() As Range)
    Dim i As Long
    For i = LBound(panels) To UBound(panels)
        Dim cap As Long: cap = panels(i).Rows.Count
        Dim remain As Long: remain = UBound(buf, 2) - idx + 1   ' ★ 2次元=行数
        Dim take As Long: take = IIf(remain <= 0, 0, Application.Min(cap, remain))
        panels(i).ClearContents
        If take > 0 Then
            panels(i).Resize(take, 4).Value = SliceRows(buf, idx, take)
            idx = idx + take
        End If
    Next i
End Sub


' ====== 2D配列から idx 以降 take 行×4列の配列を返す ======
Private Function SliceRows(ByRef src() As Variant, ByVal idx As Long, ByVal take As Long) As Variant
    Dim i As Long, j As Long
    Dim outArr() As Variant
    ReDim outArr(1 To take, 1 To 4)
    ' src(列=1..4, 行=idx..idx+take-1) → outArr(行,列)
    For i = 1 To take
        For j = 1 To 4
            outArr(i, j) = src(j, idx + i - 1)
        Next j
    Next i
    SliceRows = outArr
End Function


' ====== TaskStatus: TaskIDの列を特定 ======
Private Function FindTaskCol(ws As Worksheet, ByVal taskId As String) As Long
    Dim f As Range
    Set f = ws.Rows(ROW_TASKID).Find(What:=taskId, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then FindTaskCol = f.Column
End Function

' ====== TaskList F列 コメント取得 ======
Private Function GetTaskComment(ByVal taskId As String) As String
    Dim ws As Worksheet, f As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_TASKLIST)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Set f = ws.UsedRange.Find(What:=taskId, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then GetTaskComment = CStr(ws.Cells(f.Row, "F").Value)
End Function

' ====== 完了セルが日付なら True ======
Private Function IsCompletedDate(ByVal v As Variant) As Boolean
    If IsDate(v) Then
        IsCompletedDate = True
    ElseIf (VarType(v) = vbDouble Or VarType(v) = vbDate) And CDbl(v) > 0 Then
        IsCompletedDate = True
    Else
        IsCompletedDate = False
    End If
End Function

' 値が「ダッシュだけ（前後空白含む）」なら True
Private Function IsDashOnly(ByVal v As Variant) As Boolean
    Dim s As String
    If IsError(v) Then Exit Function

    s = CStr(v)
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbCrLf, "")
    s = Replace(s, ChrW(160), "")     ' NBSP
    s = Replace$(s, "　", "")         ' 全角スペース
    s = Trim$(s)

    ' 全角→半角（英数）
    s = StrConv(s, vbNarrow)

    ' 各種ダッシュを半角ハイフンに寄せる
    s = Replace$(s, ChrW(&H2010), "-") ' ハイフン
    s = Replace$(s, ChrW(&H2011), "-") ' ノンブレークハイフン
    s = Replace$(s, ChrW(&H2013), "-") ' ENダッシュ
    s = Replace$(s, ChrW(&H2014), "-") ' EMダッシュ
    s = Replace$(s, ChrW(&H2212), "-") ' マイナス記号

    IsDashOnly = (s = "-")
End Function


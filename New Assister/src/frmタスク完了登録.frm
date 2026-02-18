VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmタスク完了登録 
   Caption         =   "タスク完了登録"
   ClientHeight    =   8850.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8730.001
   OleObjectBlob   =   "frmタスク完了登録.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmタスク完了登録"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TOP_OFFSET As Single = 30
Private Const ROW_HEIGHT As Single = 22
Private Const LEFT_NAME As Single = 10
Private Const LEFT_CHECK_START As Single = 120
Private Const DAY_WIDTH As Single = 45

Private chkEvents As Collection   ' クラスイベント保持用
Private DaysOfWeek As Variant
Private StudentIDs As Collection ' 生徒の会員番号

'------------------------------------------------------------
' 初期化：タスク一覧をセット
'------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim wsList As Worksheet, wsStatus As Worksheet
    Set wsList = ThisWorkbook.Sheets("TaskList")
    Set wsStatus = ThisWorkbook.Sheets("TaskStatus")

    Dim lastRow As Long, i As Long
    lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row

    cmbTask.Clear

    ' === 未完了・期限内のタスクのみ追加 ===
    For i = 2 To lastRow
        If Len(wsList.Cells(i, "A").Value) > 0 Then
            Dim taskCol As Variant
            taskCol = Application.Match(wsList.Cells(i, "A").Value, wsStatus.Rows(1), 0)
            
            Dim isCompleted As Boolean
            isCompleted = False
            
            ' TaskStatus上で全員完了しているか？
            If Not IsError(taskCol) Then
                Dim blanks As Long
                blanks = WorksheetFunction.CountBlank(wsStatus.Range(wsStatus.Cells(6, taskCol), wsStatus.Cells(wsStatus.Rows.Count, taskCol).End(xlUp)))
                If blanks = 0 Then isCompleted = True
            End If
            
            ' 掲載終了日チェック
            Dim isExpired As Boolean
            If IsDate(wsList.Cells(i, "E").Value) Then
                isExpired = (wsList.Cells(i, "E").Value < Date)
            End If
            
            ' どちらにも該当しなければリストに追加
            If Not (isCompleted Or isExpired) Then
                cmbTask.AddItem wsList.Cells(i, "A").Value & "：" & wsList.Cells(i, "B").Value
            End If
        End If
    Next i

    ' === 曜日配列 ===
    DaysOfWeek = Array("月", "火", "水", "木", "金", "土")

    ' === FrameScroll の設定 ===
    With Me.FrameScroll
        .ScrollBars = fmScrollBarsVertical
        .KeepScrollBarsVisible = fmScrollBarsVertical
        .ScrollHeight = .Height
    End With

    ' === 週選択コンボボックスの生成 ===
    Dim startDate As Date, weekStart As Date, weekEnd As Date
    startDate = Date - Weekday(Date, vbMonday) + 1 ' 今週の月曜日

    cmbWeek.Clear
    For i = 0 To 4
        weekStart = startDate - (7 * i)
        weekEnd = weekStart + 6
        cmbWeek.AddItem Format$(weekStart, "mm/dd") & "～" & Format$(weekEnd, "mm/dd")
    Next i

    cmbWeek.ListIndex = 0 ' デフォルト：今週
End Sub

Private Sub cmbTask_Change()
    Dim wsList As Worksheet, wsStatus As Worksheet
    Set wsList = ThisWorkbook.Sheets("TaskList")
    Set wsStatus = ThisWorkbook.Sheets("TaskStatus")

    ' ===================== Frame内のコントロール削除 =====================
    With Me.FrameScroll
        Dim tmpHeight As Single: tmpHeight = .ScrollHeight
        .ScrollBars = fmScrollBarsNone
        .ScrollHeight = .Height
        DoEvents

        Dim t As Single: t = Timer
        Do While Timer - t < 0.05: DoEvents: Loop

        Dim ctl As Control, names() As String, i As Long
        If .Controls.Count > 0 Then
            ReDim names(0 To .Controls.Count - 1)
            For i = 0 To .Controls.Count - 1
                names(i) = .Controls(i).Name
            Next i
            For i = LBound(names) To UBound(names)
                On Error Resume Next
                .Controls.Remove names(i)
                On Error GoTo 0
            Next i
        End If
        .ScrollHeight = tmpHeight
        .ScrollBars = fmScrollBarsVertical
    End With

    ' ===================== 新しい一覧を生成 =====================
    Set chkEvents = New Collection ' ← ★ここで初期化
    Set StudentIDs = New Collection

    Dim taskId As String
    If cmbTask.ListIndex < 0 Then Exit Sub
    taskId = Split(cmbTask.Value, "：")(0)

    Dim m As Variant
    m = Application.Match(taskId, wsList.Columns(1), 0)
    If IsError(m) Then Exit Sub

    Dim condGrade As String
    condGrade = Trim$(wsList.Cells(m, 7).Value)

    Dim lastRow As Long, r As Long, Y As Single
    lastRow = wsStatus.Cells(wsStatus.Rows.Count, "A").End(xlUp).Row
    Y = TOP_OFFSET

    Dim colTask As Variant
    colTask = Application.Match(taskId, wsStatus.Rows(1), 0)
    If IsError(colTask) Then Exit Sub

    ' --- 対象生徒を抽出 ---
    For r = 6 To lastRow
        Dim stuName As String, stuGrade As String, stuID As String, cellVal As Variant
        stuGrade = wsStatus.Cells(r, "B").Value
        stuName = wsStatus.Cells(r, "C").Value
        stuID = wsStatus.Cells(r, "A").Value

        cellVal = wsStatus.Cells(r, colTask).Value
        If Trim$(cellVal) <> "" Then GoTo ContinueStudent

        If GradeMatches(condGrade, stuGrade) Then
            ' 生徒名ラベル
            Me.FrameScroll.Controls.Add "Forms.Label.1", "lbl_" & r
            With Me.FrameScroll.Controls("lbl_" & r)
                .Caption = stuGrade & " " & stuName
                .Left = LEFT_NAME
                .Top = Y
                .Width = 100
            End With

            ' 曜日チェックボックス
            Dim j As Long
            For j = LBound(DaysOfWeek) To UBound(DaysOfWeek)
                Dim chk As MSForms.CheckBox
                Dim chkEvt As clsCheckEvent

                Set chk = Me.FrameScroll.Controls.Add("Forms.CheckBox.1", "chk_" & r & "_" & j)
                With chk
                    .Caption = DaysOfWeek(j)
                    .Left = LEFT_CHECK_START + DAY_WIDTH * j
                    .Top = Y
                    .Width = DAY_WIDTH - 5
                    .Tag = r
                End With

                ' ★ クラスイベント紐づけ
                Set chkEvt = New clsCheckEvent
                chkEvt.Init chk, Me
                chkEvents.Add chkEvt
            Next j

            StudentIDs.Add r
            Y = Y + ROW_HEIGHT
        End If

ContinueStudent:
    Next r

    ' スクロール設定
    With Me.FrameScroll
        If Y > .Height Then
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = Y + 10
        Else
            .ScrollBars = fmScrollBarsNone
        End If
    End With
End Sub

Private Sub cmdRegister_Click()
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets("TaskLog")

    ' --- タスク未選択チェック ---
    If cmbTask.ListIndex = -1 Or Len(cmbTask.Value) = 0 Then
        MsgBox "タスクが選択されていません。", vbExclamation
        Exit Sub
    End If

    ' --- 週選択チェック ---
    If cmbWeek.ListIndex = -1 Or Len(cmbWeek.Value) = 0 Then
        MsgBox "登録する週を選択してください。", vbExclamation
        Exit Sub
    End If

    Dim taskId As String
    taskId = Split(cmbTask.Value, "：")(0)
    
    Dim baseDate As Date, compDate As Date
    On Error Resume Next
    baseDate = DateValue(Left$(cmbWeek.Value, 5))
    On Error GoTo 0
    If baseDate = 0 Then
        MsgBox "週の選択が不正です。", vbCritical
        Exit Sub
    End If

    Dim key As Variant, dayIdx As Long
    Dim hasChecked As Boolean: hasChecked = False
    
    '=== TaskLog最終行取得 ===
    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1

    '=== チェックボックス走査 ===
    For Each key In StudentIDs
        For dayIdx = 0 To 5
            Dim chkName As String
            chkName = "chk_" & key & "_" & dayIdx
            
            Dim chk As MSForms.CheckBox
            On Error Resume Next
            Set chk = Me.FrameScroll.Controls(chkName)
            On Error GoTo 0
            
            If Not chk Is Nothing Then
                If chk.Value = True Then
                    hasChecked = True
                    compDate = baseDate + dayIdx
                    
                    Dim stuID As String
                    stuID = ThisWorkbook.Sheets("TaskStatus").Cells(key, "A").Value
                    
                    ' --- TaskLog内でTaskID×StudentIDの行を探す ---
                    Dim found As Range
                    Dim firstAddr As String
                    
                    Set found = Nothing
                    On Error Resume Next
                    Set found = wsLog.Range("A:A").Find(What:=taskId, LookAt:=xlWhole)
                    On Error GoTo 0
                    
                    Dim foundRow As Long
                    foundRow = 0
                    
                    Do While Not found Is Nothing
                        If wsLog.Cells(found.Row, "B").Value = stuID Then
                            foundRow = found.Row
                            Exit Do
                        End If
                        Set found = wsLog.Range("A:A").FindNext(found)
                        If found.Address = firstAddr Then Exit Do
                    Loop
                    
                    ' --- 見つかった場合：完了日・登録日を更新 ---
                    If foundRow > 0 Then
                        wsLog.Cells(foundRow, "E").Value = Format$(compDate, "yyyy/mm/dd")       ' CompletedDate
                        wsLog.Cells(foundRow, "F").Value = Format$(Now, "yyyy/mm/dd")           ' RecordedDate
                    Else
                    ' --- 見つからない場合：新規追加 ---
                    Dim stuName As String, stuGrade As String
                    With ThisWorkbook.Sheets("TaskStatus")
                        stuName = .Cells(key, "C").Value    ' TaskStatusのC列＝名前
                        stuGrade = .Cells(key, "B").Value   ' TaskStatusのB列＝学年
                    End With

                    wsLog.Cells(lastRow, "A").Value = taskId           ' TaskID
                    wsLog.Cells(lastRow, "B").Value = stuID            ' StudentID
                    wsLog.Cells(lastRow, "C").Value = stuName          ' Name
                    wsLog.Cells(lastRow, "D").Value = stuGrade         ' Grade
                    wsLog.Cells(lastRow, "E").Value = Format$(compDate, "yyyy/mm/dd")  ' CompletedDate
                    wsLog.Cells(lastRow, "F").Value = Format$(Now, "yyyy/mm/dd")       ' RecordedDate
                    
                    lastRow = lastRow + 1

                    End If
                    Exit For
                End If
            End If
        Next dayIdx
    Next key

    If Not hasChecked Then
        MsgBox "完了したタスクのチェックをしてから登録を押してください。", vbExclamation
        Exit Sub
    End If

    TaskLog反映
    
    MsgBox "完了日を登録しました。", vbInformation
    Unload Me
End Sub



'------------------------------------------------------------
' 閉じる
'------------------------------------------------------------
Private Sub cmdClose_Click()
    Unload Me
End Sub

'------------------------------------------------------------
' 学年判定
'------------------------------------------------------------
Private Function GradeMatches(condGrades As String, stuGrade As String) As Boolean
    Dim arr() As String, g As Variant
    GradeMatches = False
    If condGrades = "" Then
        GradeMatches = True
        Exit Function
    End If
    arr = Split(condGrades, ",")
    For Each g In arr
        If StrComp(Trim(g), Trim(stuGrade), vbTextCompare) = 0 Then
            GradeMatches = True
            Exit Function
        End If
    Next
End Function


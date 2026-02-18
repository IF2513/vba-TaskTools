VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm新規タスク登録 
   Caption         =   "タスク登録"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11805
   OleObjectBlob   =   "frm新規タスク登録.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frm新規タスク登録"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TaskList")
    
    ' タスクID採番（T001, T002, ...）
    Dim taskId As String
    taskId = GetNextTaskID(ws)
    
    lblTaskID.Caption = taskId
End Sub

'============================
' フォーム：基本操作
'============================
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo EH

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("TaskList")

    ' 入力値チェック
    Dim taskName As String: taskName = Trim$(txtTaskName.Value)
    If taskName = "" Then
        MsgBox "タスク名を入力してください。", vbExclamation
        txtTaskName.SetFocus
        Exit Sub
    End If

    ' ★ ParseDateOrEmpty を使う（DateOrEmptyの代わり）
    Dim startDate As Variant: startDate = ParseDateOrEmpty(txtStart.Value)
    Dim dueDate   As Variant: dueDate = ParseDateOrEmpty(txtDue.Value)
    Dim endDate   As Variant: endDate = ParseDateOrEmpty(txtEnd.Value)

    If Not IsEmpty(startDate) And Not IsEmpty(endDate) Then
        If startDate > endDate Then
            MsgBox "掲載開始日が掲載終了日を超えています。", vbExclamation
            txtStart.SetFocus
            Exit Sub
        End If
    End If

    Dim taskId As String: taskId = lblTaskID.Caption

    ' TaskListシートに書き込み
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = taskId
    ws.Cells(r, 2).Value = taskName
    ws.Cells(r, 3).Value = IIf(IsDate(startDate), CDate(startDate), "")
    ws.Cells(r, 4).Value = IIf(IsDate(dueDate), CDate(dueDate), "")
    ws.Cells(r, 5).Value = IIf(IsDate(endDate), CDate(endDate), "")
    ws.Cells(r, 6).Value = txtComment.Value

    ' ★ G/H/I に条件CSV書き込み
    Dim csvGrade As String, csvDiv As String, csvTerm As String
    csvGrade = BuildCondGrade()   ' G: 対象学年
    csvDiv = BuildCondDiv()       ' H: 対象設置区分
    csvTerm = BuildCondTerm()     ' I: 対象学期制

    ws.Cells(r, 7).Value = csvGrade
    ws.Cells(r, 8).Value = csvDiv
    ws.Cells(r, 9).Value = csvTerm
    ws.Cells(r, 10).Value = ""    ' J: 予備（必要なら）

    ' TaskStatus展開（任意）
    On Error Resume Next
    Application.Run "タスク登録処理.ExpandTaskToStatus", taskId
    On Error GoTo 0
    
    実行タスク反映toTaskStatus       '実施中のタスクを表示
    
    Task条件を生徒に適用                'タスク条件を各セルに反映&TaskLogに行追加

    MsgBox "登録しました（ID: " & taskId & "）。", vbInformation
    Unload Me
    Exit Sub
EH:
    MsgBox "保存中にエラー: " & Err.Number & " " & Err.Description, vbCritical


End Sub


Private Function GetNextTaskID(ws As Worksheet) As String
    Dim i As Long, maxNum As Long, s As String
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        s = CStr(ws.Cells(i, 1).Value)
        If Left$(s, 1) = "T" Then maxNum = Application.Max(maxNum, Val(Mid$(s, 2)))
    Next
    GetNextTaskID = "T" & Format$(maxNum + 1, "000")
End Function

Private Sub ClearFormFields()
    txtTaskName.Value = ""
    txtStart.Value = ""
    txtDue.Value = ""
    txtEnd.Value = ""
    txtComment.Value = ""
End Sub

Private Sub lblTaskID_Click()

End Sub

'============================
' 日付入力補助（8桁自動整形＋半角変換）
'============================

Private Sub txtStart_Change(): FormatDateTyping txtStart: End Sub
Private Sub txtDue_Change():   FormatDateTyping txtDue:   End Sub
Private Sub txtEnd_Change():   FormatDateTyping txtEnd:   End Sub

Private Sub txtStart_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    AllowDigitsOnly KeyAscii
End Sub
Private Sub txtDue_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    AllowDigitsOnly KeyAscii
End Sub
Private Sub txtEnd_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    AllowDigitsOnly KeyAscii
End Sub

Private Sub txtStart_AfterUpdate()
    ValidateDateField txtStart, "開始日"
End Sub
Private Sub txtDue_AfterUpdate()
    ValidateDateField txtDue, "終了日"
End Sub
Private Sub txtEnd_AfterUpdate()
    ValidateDateField txtEnd, "掲載終了日"
End Sub

Private Sub AllowDigitsOnly(ByRef KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8, 9, 48 To 57 ' Backspace, Tab, 0-9
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub FormatDateTyping(tb As Object)
    Static inFormat As Boolean
    If inFormat Then Exit Sub
    inFormat = True

    Dim raw As String: raw = Hankaku(tb.Value)
    Dim digits As String, i As Long, ch As String

    ' 数字のみ抽出
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next
    If Len(digits) > 8 Then digits = Left$(digits, 8)

    ' スラッシュ挿入
    Dim s As String
    Select Case Len(digits)
        Case 0 To 4: s = digits
        Case 5 To 6: s = Left$(digits, 4) & "/" & Mid$(digits, 5)
        Case Else:   s = Left$(digits, 4) & "/" & Mid$(digits, 5, 2) & "/" & Mid$(digits, 7)
    End Select

    ' キャレット位置補正
    Dim caret As Long, newCaret As Long
    For i = 1 To tb.SelStart
        ch = Mid$(tb.Value, i, 1)
        If ch >= "0" And ch <= "9" Then caret = caret + 1
    Next
    Select Case caret
        Case 0 To 4: newCaret = caret
        Case 5 To 6: newCaret = caret + 1
        Case Else:   newCaret = caret + 2
    End Select
    If newCaret > Len(s) Then newCaret = Len(s)

    tb.Value = s
    tb.SelStart = newCaret
    tb.SelLength = 0

    inFormat = False
End Sub

Private Sub ValidateDateField(tb As Object, ByVal labelText As String)
    Dim t As String: t = tb.Value
    If Len(t) = 0 Then Exit Sub

    If Len(t) <> 10 Or Mid$(t, 5, 1) <> "/" Or Mid$(t, 8, 1) <> "/" Or Not IsDate(t) Then
        MsgBox labelText & "は 8桁の数字で正しい日付を入力してください（例：20250817）。", vbExclamation
        tb.SetFocus
        tb.SelStart = 0: tb.SelLength = Len(tb.Value)
        Exit Sub
    End If

    tb.Value = Format$(CDate(t), "yyyy/mm/dd")
End Sub

'============================
' 学年チェック：一括ON/OFF
'============================
Private Sub chkG_All_Click()
    Dim c As Variant
    For Each c In Array( _
        chkG_Grad, _
        chkG_H3, chkG_H2, chkG_H1, _
        chkG_J3, chkG_J2, chkG_J1, _
        chkG_E6, chkG_E5, chkG_E4, chkG_E3, chkG_E2, chkG_E1)
        c.Value = chkG_All.Value
    Next
End Sub

'=== 条件 → CSV（frmタスク登録 内に置く） ===
Private Function BuildCondGrade() As String
    Dim a As New Collection
    If chkG_Grad.Value Then a.Add "既卒"
    If chkG_H3.Value Then a.Add "高3"
    If chkG_H2.Value Then a.Add "高2"
    If chkG_H1.Value Then a.Add "高1"
    If chkG_J3.Value Then a.Add "中3"
    If chkG_J2.Value Then a.Add "中2"
    If chkG_J1.Value Then a.Add "中1"
    If chkG_E6.Value Then a.Add "小6"
    If chkG_E5.Value Then a.Add "小5"
    If chkG_E4.Value Then a.Add "小4"
    If chkG_E3.Value Then a.Add "小3"
    If chkG_E2.Value Then a.Add "小2"
    If chkG_E1.Value Then a.Add "小1"
    BuildCondGrade = JoinCsv(a)
End Function

Private Function BuildCondDiv() As String
    Dim a As New Collection
    If chkS_Public.Value Then a.Add "公立"
    If chkS_Kokuritsu.Value Then a.Add "国立"
    If chkS_Private.Value Then a.Add "私立"
    If chkS_Toritsu.Value Then a.Add "都立"
    If chkS_Kenritsu.Value Then a.Add "県立"            ' 県名入りも後段でヒットさせる前提
    If chkS_Machida.Value Then a.Add "町田市立"
    If chkS_Sagamihara.Value Then a.Add "相模原市立"
    If chkS_Hachioji.Value Then a.Add "八王子市立"
    BuildCondDiv = JoinCsv(a)
End Function

Private Function BuildCondTerm() As String
    Dim a As New Collection
    If chkT_3.Value Then a.Add "3学期制"
    If chkT_2.Value Then a.Add "2学期制"
    BuildCondTerm = JoinCsv(a)
End Function

Private Function JoinCsv(c As Collection) As String
    Dim i As Long, s As String
    For i = 1 To c.Count
        s = s & c(i) & ","
    Next
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    JoinCsv = s
End Function


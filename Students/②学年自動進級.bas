Attribute VB_Name = "②学年自動進級"
Option Explicit

Private Const SETTINGS_SHEET As String = "config"
Private Const CELL_LAST_YEAR As String = "B1"
Private Const CELL_LAST_DATE As String = "B2"
Private Const STUDENT_SHEET  As String = "生徒情報"
Private Const COL_SCHOOL     As Long = 4   ' D列
Private Const COL_GRADE      As Long = 5   ' E列

' 開いた時にチェック＆必要なら自動進級
Public Sub 学年自動進級()
    Dim wsCfg As Worksheet
    Set wsCfg = ThisWorkbook.Worksheets(SETTINGS_SHEET)

    Dim lastY As Long, curY As Long
    If IsNumeric(wsCfg.Range(CELL_LAST_YEAR).Value) Then
        lastY = CLng(wsCfg.Range(CELL_LAST_YEAR).Value)
    Else
        lastY = 0
    End If

    curY = 学年度(Date)

    If curY > lastY Then
        ' 新年度なら進級処理
        Call 学年一括進級_実行
        ' configシートの年度と日時を更新
        wsCfg.Range(CELL_LAST_YEAR).Value = curY
        wsCfg.Range(CELL_LAST_DATE).Value = Now
    End If
End Sub

' 学年を1つ進める処理
Private Sub 学年一括進級_実行()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(STUDENT_SHEET)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim lastRow As Long: lastRow = ws.Cells(ws.rowS.Count, COL_GRADE).End(xlUp).Row
    Dim r As Long, g As String, newG As String

    For r = 2 To lastRow
        g = Trim$(CStr(ws.Cells(r, COL_GRADE).Value))
        If g <> "" Then
            newG = 次学年(g)
            If newG <> g Then
                ws.Cells(r, COL_GRADE).Value = newG
                If newG = "高1" Then ws.Cells(r, COL_SCHOOL).ClearContents   ' 中3→高1で学校空欄
                If newG = "既卒" Then ws.Cells(r, COL_SCHOOL).Value = "既卒" ' 高3→既卒で既卒
            End If
        End If
    Next r

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' 学年マップ
Private Function 次学年(ByVal g As String) As String
    Select Case g
        Case "小1": 次学年 = "小2"
        Case "小2": 次学年 = "小3"
        Case "小3": 次学年 = "小4"
        Case "小4": 次学年 = "小5"
        Case "小5": 次学年 = "小6"
        Case "小6": 次学年 = "中1"
        Case "中1": 次学年 = "中2"
        Case "中2": 次学年 = "中3"
        Case "中3": 次学年 = "高1"
        Case "高1": 次学年 = "高2"
        Case "高2": 次学年 = "高3"
        Case "高3": 次学年 = "既卒"
        Case Else:  次学年 = g   ' 浪人・既卒・その他は据え置き
    End Select
End Function

' 今日が属する学年度（4/1起点）
Private Function 学年度(ByVal d As Date) As Long
    If Month(d) >= 4 Then
        学年度 = Year(d)
    Else
        学年度 = Year(d) - 1
    End If
End Function


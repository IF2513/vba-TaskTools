Attribute VB_Name = "A2_学校情報同期"
Option Explicit

Private Const SRC_FILE  As String = "Students.xlsm"
Private Const SRC_SHEET As String = "学校情報"
Private Const DST_SHEET As String = "学校情報 from Students.xlsm"
Private Const COLS As Long = 4      ' A:D
Private Const KEY_COL As Long = 1   ' 学校コード列(A)

Public Sub 学校情報同期()
    'Dim scr As Boolean, evt As Boolean, calc As XlCalculation, alerts As Boolean
    'scr = Application.ScreenUpdating: Application.ScreenUpdating = False
    'evt = Application.EnableEvents:   Application.EnableEvents = False
    'calc = Application.Calculation:   Application.Calculation = xlCalculationManual
    'alerts = Application.DisplayAlerts: Application.DisplayAlerts = False

    On Error GoTo Fin

    Dim wbThis As Workbook: Set wbThis = ThisWorkbook
    Dim wsDst As Worksheet: Set wsDst = wbThis.Worksheets(DST_SHEET)

    ' --- ソースブックを開く（必要時のみ） ---
    Dim fullPath As String: fullPath = wbThis.Path & "\" & SRC_FILE
    Dim wbSrc As Workbook, openedHere As Boolean
    On Error Resume Next
    Set wbSrc = Workbooks(SRC_FILE)
    On Error GoTo 0
    If wbSrc Is Nothing Then
        If Dir(fullPath) = "" Then Err.Raise vbObjectError + 100, , "外部ファイルが見つかりません: " & fullPath
        Set wbSrc = Workbooks.Open(Filename:=fullPath, ReadOnly:=True)
        If wbSrc.Windows.Count > 0 Then wbSrc.Windows(1).Visible = False
        openedHere = True
    End If
    Dim wsSrc As Worksheet: Set wsSrc = wbSrc.Worksheets(SRC_SHEET)

    ' --- 実データの最終行 ---
    Dim lastSrc As Long: lastSrc = LastDataRow(wsSrc, "A:F") ' ソースはA,Fまで見ておく（D,E,F列使用）
    Dim lastDst As Long: lastDst = LastDataRow(wsDst, "A:D")
    If lastDst < 1 Then lastDst = 1

    ' --- 出力側ID辞書（キー＝学校コード） ---
    Dim dstIdx As Object: Set dstIdx = CreateObject("Scripting.Dictionary")
    dstIdx.CompareMode = 1
    Dim r As Long, code As String
    For r = 2 To lastDst
        code = Trim$(CStr(wsDst.Cells(r, KEY_COL).Value))
        If Len(code) > 0 Then If Not dstIdx.Exists(code) Then dstIdx.Add code, r
    Next

    ' --- ソース側コード集合（削除判定用） ---
    Dim srcCodes As Object: Set srcCodes = CreateObject("Scripting.Dictionary")
    srcCodes.CompareMode = 1

    ' --- 追加・更新（A,D,E,F列のみ反映） ---
    For r = 2 To lastSrc
        code = Trim$(CStr(wsSrc.Cells(r, "A").Value))
        If Len(code) = 0 Then GoTo NextR

        srcCodes(code) = True

        If dstIdx.Exists(code) Then
            Dim dstRow As Long: dstRow = CLng(dstIdx(code))
            If Not RowsEqual_School(wsSrc, wsDst, r, dstRow) Then
                wsDst.Cells(dstRow, "A").Value = wsSrc.Cells(r, "A").Value ' コード
                wsDst.Cells(dstRow, "B").Value = wsSrc.Cells(r, "D").Value ' 学校名（整形）
                wsDst.Cells(dstRow, "C").Value = wsSrc.Cells(r, "E").Value ' 設置区分
                wsDst.Cells(dstRow, "D").Value = wsSrc.Cells(r, "F").Value ' 学期制
            End If
        Else
            Dim newRow As Long: newRow = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).Row + 1
            wsDst.Cells(newRow, "A").Value = wsSrc.Cells(r, "A").Value
            wsDst.Cells(newRow, "B").Value = wsSrc.Cells(r, "D").Value
            wsDst.Cells(newRow, "C").Value = wsSrc.Cells(r, "E").Value
            wsDst.Cells(newRow, "D").Value = wsSrc.Cells(r, "F").Value
            dstIdx.Add code, newRow
        End If
NextR:
    Next

    ' --- 削除処理（ソースに存在しない学校を削除） ---
    Dim i As Long
    For i = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        code = Trim$(CStr(wsDst.Cells(i, KEY_COL).Value))
        If Len(code) = 0 Then
            wsDst.Rows(i).Delete
        ElseIf Not srcCodes.Exists(code) Then
            wsDst.Rows(i).Delete
        End If
    Next

Fin:
    If openedHere Then
        On Error Resume Next
        If wbSrc.Windows.Count > 0 Then wbSrc.Windows(1).Visible = True
        wbSrc.Close SaveChanges:=False
        On Error GoTo 0
    End If

    'Application.DisplayAlerts = alerts
    'Application.Calculation = calc
    'Application.EnableEvents = evt
    'Application.ScreenUpdating = scr

    'MsgBox "学校情報を同期しました。", vbInformation
End Sub


' --- 比較：A,D,E,F列の差分検出 ---
Private Function RowsEqual_School(wsSrc As Worksheet, wsDst As Worksheet, rSrc As Long, rDst As Long) As Boolean
    RowsEqual_School = _
        Normalize(wsSrc.Cells(rSrc, "A").Value) = Normalize(wsDst.Cells(rDst, "A").Value) And _
        Normalize(wsSrc.Cells(rSrc, "D").Value) = Normalize(wsDst.Cells(rDst, "B").Value) And _
        Normalize(wsSrc.Cells(rSrc, "E").Value) = Normalize(wsDst.Cells(rDst, "C").Value) And _
        Normalize(wsSrc.Cells(rSrc, "F").Value) = Normalize(wsDst.Cells(rDst, "D").Value)
End Function

' --- 値正規化（Null, 空, 日付対応） ---
Private Function Normalize(v As Variant) As String
    If IsError(v) Then
        Normalize = "#ERR!"
    ElseIf IsEmpty(v) Or v = "" Then
        Normalize = ""
    ElseIf IsDate(v) Then
        Normalize = CStr(CDbl(CDate(v)))
    Else
        Normalize = CStr(v)
    End If
End Function

' --- 値ベースで最終行を取得 ---
Private Function LastDataRow(ws As Worksheet, ByVal addr As String) As Long
    Dim f As Range
    Set f = ws.Range(addr).Find(What:="*", LookIn:=xlValues, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If f Is Nothing Then
        LastDataRow = 0
    Else
        LastDataRow = f.Row
    End If
End Function



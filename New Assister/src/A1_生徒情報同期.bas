Attribute VB_Name = "A1_生徒情報同期"
Option Explicit

Private Const SRC_FILE  As String = "Students.xlsm"
Private Const SRC_SHEET As String = "生徒情報一覧"
Private Const DST_SHEET As String = "Students from Students.xlsm"
Private Const COLS As Long = 14     ' A:N
Private Const KEY_COL As Long = 1   'StudentID


Public Sub 生徒情報同期()
    'Dim scr As Boolean, evt As Boolean, calc As XlCalculation, alerts As Boolean
    'scr = Application.ScreenUpdating: Application.ScreenUpdating = False
    'evt = Application.EnableEvents:   Application.EnableEvents = False
    'calc = Application.Calculation:   Application.Calculation = xlCalculationManual
    'alerts = Application.DisplayAlerts: Application.DisplayAlerts = False

    On Error GoTo Fin

    Dim wbThis As Workbook: Set wbThis = ThisWorkbook
    Dim wsDst As Worksheet: Set wsDst = wbThis.Worksheets(DST_SHEET)

    ' --- ソースを必要時のみ開く（読み取り専用） ---
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

    ' --- 実データ最終行（値ベース。罫線等の書式は無視） ---
    Dim lastSrc As Long: lastSrc = LastDataRow(wsSrc, "A:N")
    Dim lastDst As Long: lastDst = LastDataRow(wsDst, "A:N")
    If lastDst < 1 Then lastDst = 1 ' ヘッダのみでも安全

    ' --- Assister側のID→行番号辞書 ---
    Dim dstIdx As Object: Set dstIdx = CreateObject("Scripting.Dictionary")
    dstIdx.CompareMode = 1 ' TextCompare
    Dim r As Long, id As String
    For r = 2 To lastDst
        id = Trim$(CStr(wsDst.Cells(r, KEY_COL).Value))
        If Len(id) > 0 Then If Not dstIdx.Exists(id) Then dstIdx.Add id, r
    Next

    ' --- ソース側ID集合（削除判定用） ---
    Dim srcIDs As Object: Set srcIDs = CreateObject("Scripting.Dictionary")
    srcIDs.CompareMode = 1

    ' --- 追加・更新（行単位。A:N全列を比較→差分あれば上書き） ---
    For r = 2 To lastSrc
        id = Trim$(CStr(wsSrc.Cells(r, KEY_COL).Value))
        If Len(id) = 0 Then GoTo NextR

        srcIDs(id) = True ' 存在マーク

        If dstIdx.Exists(id) Then
            ' 既存 → 差分チェックして上書き
            Dim dstRow As Long: dstRow = CLng(dstIdx(id))
            If Not RowsEqual(wsSrc.Cells(r, 1).Resize(1, COLS), wsDst.Cells(dstRow, 1).Resize(1, COLS)) Then
                wsDst.Cells(dstRow, 1).Resize(1, COLS).Value2 = wsSrc.Cells(r, 1).Resize(1, COLS).Value2
            End If
        Else
            ' 新規 → 末尾に追記
            lastDst = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).Row
            If lastDst < 1 Then lastDst = 1
            Dim newRow As Long: newRow = lastDst + 1
            wsDst.Cells(newRow, 1).Resize(1, COLS).Value2 = wsSrc.Cells(r, 1).Resize(1, COLS).Value2
            dstIdx.Add id, newRow
        End If
NextR:
    Next

    ' --- 削除（ソースにいないID行を下から削除） ---
    Dim i As Long
    For i = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        id = Trim$(CStr(wsDst.Cells(i, KEY_COL).Value))
        If Len(id) = 0 Then
            wsDst.Rows(i).Delete
        ElseIf Not srcIDs.Exists(id) Then
            wsDst.Rows(i).Delete
        End If
    Next

    ' （任意）見た目
    ' wsDst.Columns("A:N").AutoFit

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
End Sub

' --- 2つの1行範囲（A:N）を厳密比較。1セルでも違えば False ---
Private Function RowsEqual(r1 As Range, r2 As Range) As Boolean
    Dim c As Long
    For c = 1 To COLS
        If NormalizeCell(r1.Cells(1, c).Value2) <> NormalizeCell(r2.Cells(1, c).Value2) Then
            RowsEqual = False
            Exit Function
        End If
    Next
    RowsEqual = True
End Function

' 値の正規化（Null/Empty→空文字、数値はそのまま、日付はCDblで比較しても良い）
Private Function NormalizeCell(v As Variant) As String
    If IsError(v) Then
        NormalizeCell = "#ERR!"
    ElseIf IsEmpty(v) Or v = "" Then
        NormalizeCell = ""
    ElseIf IsDate(v) Then
        ' 日付はシリアル値で比較（フォーマット差を吸収）
        NormalizeCell = CStr(CDbl(CDate(v)))
    Else
        NormalizeCell = CStr(v)
    End If
End Function

' 値ベースでA:Nの最後の行（罫線等の書式は無視）
Private Function LastDataRow(ws As Worksheet, ByVal addr As String) As Long
    Dim f As Range
    Set f = ws.Range(addr).Find(What:="*", LookIn:=xlValues, LookAt:=xlPart, _
                                SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If f Is Nothing Then LastDataRow = 0 Else LastDataRow = f.Row
End Function


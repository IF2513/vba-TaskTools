Attribute VB_Name = "B3_印刷用整形・印刷"
Option Explicit

' 週ターゲット定数（Long）
Public Const WK_AUTO As Long = 0      ' 土日→来週 / 平日→今週
Public Const WK_THIS As Long = 1      ' 常に今週
Public Const WK_NEXT As Long = 2      ' 常に来週

Public Const TARGET_WEEK_MODE As Long = WK_AUTO

Private Sub ComputeWeekRange(runDate As Date, mode As Long, _
                             ByRef weekStart As Date, ByRef weekEnd As Date)
    Dim base As Date, wd As Long
    wd = Weekday(runDate, vbMonday)   ' 月=1 … 日=7
    base = runDate - (wd - 1)         ' その週の月曜

    Select Case mode
        Case WK_THIS: weekStart = base
        Case WK_NEXT: weekStart = base + 7
        Case Else    ' WK_AUTO
            If wd >= 6 Then           ' 土(6)・日(7)は来週
                weekStart = base + 7
            Else
                weekStart = base
            End If
    End Select
    weekEnd = weekStart + 5           ' 月曜＋5＝土曜
End Sub

Public Sub アシストシート印刷()
    Dim wsS As Worksheet, wsO As Worksheet, wsL As Worksheet
    Set wsS = ThisWorkbook.Worksheets("TaskStatus") ' Source
    Set wsO = ThisWorkbook.Worksheets("印刷用")     ' Output
    Set wsL = ThisWorkbook.Worksheets("Log")        ' Log（既存）

    Const PAGE_SIZE As Long = 40
    Const OUT_FIRST_COL As Long = 5   ' E
    Const OUT_LAST_COL  As Long = 18  ' R
    Const CHUNK_SIZE As Long = (OUT_LAST_COL - OUT_FIRST_COL + 1) ' 14
    Const WEEKDAYS_TEXT As String = "月・火・水・木・金・土"

    ' 色
    Dim DARK_GREY As Long:  DARK_GREY = RGB(174, 170, 174)   ' 値ありセル
    Dim LIGHT_GREY As Long: LIGHT_GREY = RGB(242, 242, 242)  ' 偶数行の空だったセル & 生徒の偶数行

    ' 行範囲
    Dim firstRow As Long, lastRow As Long
    firstRow = 6
    lastRow = wsS.Cells(wsS.Rows.Count, "A").End(xlUp).Row
    If lastRow < firstRow Then Exit Sub

    ' タスク最終列：1行目の右端
    Dim srcLastCol As Long
    srcLastCol = wsS.Cells(1, wsS.Columns.Count).End(xlToLeft).Column

    ' Log初期化（必要なら）
    wsL.Range("A:D").ClearContents

    Dim r As Long, blockStart As Long, blockEnd As Long, rowsInBlock As Long
    Dim blockcnt As Long: blockcnt = 0
    
    ' ===== ページカウンタ =====
    Dim totalPages As Long, pageNo As Long
    pageNo = 1

    ' === 総ページ数を事前に算出 ===
    Dim scanR As Long, scanStart As Long, scanEnd As Long, scanRows As Long
    Dim col As Long, colCnt As Long, blanks As Long

    totalPages = 0
    scanR = firstRow
    Do While scanR <= lastRow
        scanStart = scanR
        scanEnd = Application.Min(scanStart + PAGE_SIZE - 1, lastRow)
        scanRows = scanEnd - scanStart + 1

        ' このブロックで「全員完了ではない」タスク列数（＝表示対象）を数える
        colCnt = 0
        For col = 6 To srcLastCol
            blanks = Application.WorksheetFunction.CountBlank(wsS.Cells(scanStart, col).Resize(scanRows, 1))
            If blanks > 0 Then colCnt = colCnt + 1
        Next col

        If colCnt = 0 Then
            totalPages = totalPages + 1                  ' 生徒一覧だけで1ページ
        Else
            totalPages = totalPages + ((colCnt + CHUNK_SIZE - 1) \ CHUNK_SIZE) ' 14列ごと
        End If

        scanR = scanEnd + 1
    Loop

    
    ' 以下、印刷整形ループ開始・画面停止
    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Dim prevScreen As Boolean:     prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    
    r = firstRow
    Do While r <= lastRow
        blockStart = r
        blockEnd = Application.Min(blockStart + PAGE_SIZE - 1, lastRow)
        rowsInBlock = blockEnd - blockStart + 1

        ' ========== 1) 生徒一覧（B,A,C,D → A,B,C,D）を一括代入 ==========
        wsO.Range("A4:D43").ClearContents

        ' B列→A列
        wsO.Range("A4").Resize(rowsInBlock, 1).Value = wsS.Cells(blockStart, 2).Resize(rowsInBlock, 1).Value
        ' A列→B列
        wsO.Range("B4").Resize(rowsInBlock, 1).Value = wsS.Cells(blockStart, 1).Resize(rowsInBlock, 1).Value
        ' C列→C列
        wsO.Range("C4").Resize(rowsInBlock, 1).Value = wsS.Cells(blockStart, 3).Resize(rowsInBlock, 1).Value
        ' D列→D列
        wsO.Range("D4").Resize(rowsInBlock, 1).Value = wsS.Cells(blockStart, 4).Resize(rowsInBlock, 1).Value

        ' 生徒一覧 偶数行を #F2F2F2（40回以内の軽ループ）
        Dim rr As Long
        For rr = 5 To 3 + rowsInBlock Step 2         ' 5,7,9,...（2行目=偶数）
            wsO.Range(wsO.Cells(rr, 1), wsO.Cells(rr, 4)).Interior.Color = LIGHT_GREY
        Next rr

        ' Logへ積む（端数もOK）
        Dim logStart As Long
        logStart = blockcnt * PAGE_SIZE + 1
        wsL.Range(wsL.Cells(logStart, 1), wsL.Cells(logStart + rowsInBlock - 1, 4)).Value = _
            wsO.Range(wsO.Cells(4, 1), wsO.Cells(4 + rowsInBlock - 1, 4)).Value

        ' ========== 2) タスク列の抽出（全員完了は CountBlank=0 でスキップ） ==========
        Dim eligible() As Long, cnt As Long, srcCol As Long
        cnt = 0
        For srcCol = 6 To srcLastCol ' F→
            blanks = Application.WorksheetFunction.CountBlank(wsS.Cells(blockStart, srcCol).Resize(rowsInBlock, 1))
            If blanks > 0 Then
                cnt = cnt + 1
                If cnt = 1 Then ReDim eligible(1 To 1) Else ReDim Preserve eligible(1 To cnt)
                eligible(cnt) = srcCol
            End If
        Next srcCol

        ' ========== 3) タスク：14列ずつ（E?R）に分割して転記 → 各塊ごとに印刷 ==========
        Dim printedThisBlock As Boolean: printedThisBlock = False
        Dim tIdx As Long, tEnd As Long, k As Long, srcC As Long
        Dim outCol As Long, dstColRg As Range, dstValsRg As Range
        Dim rngNonBlank As Range, rngForm As Range, rngBlanks As Range, rngEven As Range, rngHit As Range

        tIdx = 1
        Do While cnt > 0 And tIdx <= cnt
            tEnd = Application.Min(tIdx + CHUNK_SIZE - 1, cnt)

            ' タスク領域だけクリア
            wsO.Range(wsO.Cells(2, OUT_FIRST_COL), wsO.Cells(3, OUT_LAST_COL)).ClearContents
            With wsO.Range(wsO.Cells(4, OUT_FIRST_COL), wsO.Cells(3 + PAGE_SIZE, OUT_LAST_COL))
                .ClearContents
                .Interior.Pattern = xlNone
            End With

            outCol = OUT_FIRST_COL
            For k = tIdx To tEnd
                srcC = eligible(k)

                ' ヘッダ（1→2行目、4→3行目）
                wsO.Cells(2, outCol).Value = wsS.Cells(1, srcC).Value
                wsO.Cells(3, outCol).Value = wsS.Cells(4, srcC).Value

                ' 状況（40行）を一括転記
                Set dstColRg = wsO.Cells(4, outCol).Resize(rowsInBlock, 1)
                dstColRg.Value = wsS.Cells(blockStart, srcC).Resize(rowsInBlock, 1).Value

                ' 値ありセルをまとめて濃グレー（定数／数式の両方）
                On Error Resume Next
                Set rngNonBlank = dstColRg.SpecialCells(xlCellTypeConstants)
                Set rngForm = dstColRg.SpecialCells(xlCellTypeFormulas)
                On Error GoTo 0
                If Not rngForm Is Nothing Then
                    If rngNonBlank Is Nothing Then
                        Set rngNonBlank = rngForm
                    Else
                        Set rngNonBlank = Union(rngNonBlank, rngForm)
                    End If
                End If
                If Not rngNonBlank Is Nothing Then rngNonBlank.Interior.Color = DARK_GREY

                ' 空白セルすべてに曜日文字列を一括代入（文字列書式）
                On Error Resume Next
                Set rngBlanks = dstColRg.SpecialCells(xlCellTypeBlanks)
                On Error GoTo 0
                If Not rngBlanks Is Nothing Then
                    rngBlanks.NumberFormatLocal = "@"
                    rngBlanks.Value = WEEKDAYS_TEXT

                    ' 偶数行（ページ内2,4,6…行）だったセルだけ薄グレー
                    Set rngEven = Nothing
                    For rr = 5 To 3 + rowsInBlock Step 2 ' 5,7,9,... = 偶数番目行
                        If rngEven Is Nothing Then
                            Set rngEven = wsO.Cells(rr, outCol)
                        Else
                            Set rngEven = Union(rngEven, wsO.Cells(rr, outCol))
                        End If
                    Next rr
                    If Not rngEven Is Nothing Then
                        Set rngHit = Intersect(rngBlanks, rngEven)
                        If Not rngHit Is Nothing Then rngHit.Interior.Color = LIGHT_GREY
                    End If
                End If

                outCol = outCol + 1
                If outCol > OUT_LAST_COL Then Exit For
            Next k
            
            ' ▼ F1：週範囲（例: 11月4日～11月9日）
            Dim wkStart As Date, wkEnd As Date
            Call ComputeWeekRange(Date, TARGET_WEEK_MODE, wkStart, wkEnd)
            With wsO.Range("G1")
                .NumberFormatLocal = "@"
                .Value = Format$(wkStart, "m""月""d""日""") & "～" & Format$(wkEnd, "m""月""d""日""")
            End With

            ' ページ数入力・印刷
            With wsO.Range("R1")
                .NumberFormatLocal = "@"
                .Value = CStr(pageNo) & "/" & CStr(totalPages)
            End With
            Worksheets("印刷用").PrintOut
            printedThisBlock = True
            pageNo = pageNo + 1

            tIdx = tEnd + 1
        Loop

        ' タスク塊が無かったブロックは、生徒一覧のみ印刷
        If Not printedThisBlock Then
            With wsO.Range("R1")
                .NumberFormatLocal = "@"
                .Value = CStr(pageNo) & "/" & CStr(totalPages)
            End With
            Worksheets("印刷用").PrintOut
            pageNo = pageNo + 1
        End If

        ' ========== 4) ブロック完了：生徒＋タスクをまとめてクリア ==========
        wsO.Range("A4:D43").ClearContents
        wsO.Range(wsO.Cells(2, OUT_FIRST_COL), wsO.Cells(3, OUT_LAST_COL)).ClearContents
        With wsO.Range(wsO.Cells(4, OUT_FIRST_COL), wsO.Cells(3 + PAGE_SIZE, OUT_LAST_COL))
            .ClearContents
            .Interior.Pattern = xlNone
        End With

        r = blockEnd + 1
        wsO.Range("G1, R1").ClearContents
        blockcnt = blockcnt + 1
    Loop
    
    ' ループ終了
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    
End Sub

Public Sub 印刷ボタン()
Dim ans As Long
    ans = MsgBox("アシストシートを印刷しますか？", vbQuestion + vbYesNo + vbSystemModal, "印刷の確認")
    If ans = vbYes Then
        アシストシート印刷
    Else
        Exit Sub
    End If
End Sub


Attribute VB_Name = "A4_ListtoStatus"
Option Explicit

' TaskListからTaskStatusに項目反映
' 反映条件：
'  ・今日が開始日～掲載終了日の範囲内（両端含む）  Start<=Today<=End
'  ・TaskListのJ列が1ではない（=対象者の未完了者がいる）
Public Sub 実行タスク反映toTaskStatus()
    Dim wsList As Worksheet, wsStatus As Worksheet
    Dim lastRow As Long, i As Long, j As Long, n As Long
    Dim src As Variant            ' TaskList A:J（2行目～）
    Dim work() As Variant         ' フィルタ後（出力候補：A～Eのみ持つ）
    Dim todayD As Date: todayD = Date
    
    Set wsList = ThisWorkbook.Sheets("TaskList")
    Set wsStatus = ThisWorkbook.Sheets("TaskStatus")
    
    ' TaskList の最終行
    lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row
    ' 出力先は毎回クリア
    wsStatus.Range("F:AZ").ClearContents
    If lastRow < 2 Then Exit Sub
    
    ' --- A:J を配列取得（2行目～最終行）---
    ' A:ID, B:タスク名, C:開始日, D:締切日, E:掲載終了日, ... , J:全員完了(=1なら全完了)
    src = wsList.Range("A2:J" & lastRow).Value
    
    ' --- 条件でフィルタ ---
    ReDim work(1 To UBound(src, 1), 1 To 5)   ' A～Eのみ持てば十分
    n = 0
    For i = 1 To UBound(src, 1)
        Dim keep As Boolean: keep = True
        
        ' 1) 実施期間内か（Start<=Today<=End）
        '    ・開始日(C)が日付で、かつ Today より後なら開始前 → 除外
        If IsDate(src(i, 3)) Then
            If CDate(src(i, 3)) > todayD Then keep = False
        End If
        '    ・掲載終了(E)が日付で、かつ Today より前なら終了済 → 除外
        If keep Then
            If IsDate(src(i, 5)) Then
                If CDate(src(i, 5)) < todayD Then keep = False
            End If
        End If
        
        ' 2) 全員完了タスクの除外（J列=1 は除外）
        '    Jが数値/文字列問わず "1" と等価なら除外
        If keep Then
            If Val(CStr(src(i, 10))) = 1 Then keep = False
        End If
        
        ' 採用なら A～E を work に積む
        If keep Then
            n = n + 1
            For j = 1 To 5
                work(n, j) = src(i, j)
            Next j
        End If
    Next i
    
    ' 出力対象なしなら終了（クリア済）
    If n = 0 Then Exit Sub
    
    ' --- 残ったタスクのみ簡易ソート（掲載終了日昇順。空/非日付は後ろ）---
    '     必要なければこのブロック削除可
    Dim swapped As Boolean, tmp(1 To 5)
    Dim dA As Double, dB As Double
    Dim keyA As Long, keyB As Long
    Do
        swapped = False
        For i = 1 To n - 1
            ' A側
            If IsDate(work(i, 5)) Then
                dA = CDbl(CDate(work(i, 5)))
                keyA = 0
            Else
                dA = CDbl(DateSerial(9999, 12, 31)) ' 非日付/空白は末尾寄せ
                keyA = 1
            End If
            ' B側
            If IsDate(work(i + 1, 5)) Then
                dB = CDbl(CDate(work(i + 1, 5)))
                keyB = 0
            Else
                dB = CDbl(DateSerial(9999, 12, 31))
                keyB = 1
            End If
            
            If (keyA > keyB) Or (keyA = keyB And dA > dB) Then
                For j = 1 To 5
                    tmp(j) = work(i, j)
                    work(i, j) = work(i + 1, j)
                    work(i + 1, j) = tmp(j)
                Next j
                swapped = True
            End If
        Next i
    Loop While swapped
    
    ' --- TaskStatus へ出力 ---
    ' 1=ID, 2=タスク名, 3=開始日, 4=締切日, 5=掲載終了日 を F列以降に
    Dim outCol As Long: outCol = 6 ' F
    For i = 1 To n
        wsStatus.Cells(1, outCol).Value = work(i, 1)
        wsStatus.Cells(2, outCol).Value = work(i, 2)
        wsStatus.Cells(3, outCol).Value = work(i, 3)
        wsStatus.Cells(4, outCol).Value = work(i, 4)
        wsStatus.Cells(5, outCol).Value = work(i, 5)
        outCol = outCol + 1
    Next i
End Sub


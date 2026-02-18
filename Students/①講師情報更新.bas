Attribute VB_Name = "①講師情報更新"
Sub 講師情報更新()
    Const SRC_BOOK As String = "Tutors.xlsm"
    Const SRC_SHEET As String = "講師一覧"
    Const DST_SHEET As String = "講師一覧(from Tutors.xlsm)"

    Dim wbSut As Workbook, wsTuList As Worksheet
    Dim wbTut As Workbook, wsTut As Worksheet
    Dim lastTut As Long
    Dim arr As Variant

    On Error GoTo ErrHandler

    ' --- 呼び元（Students.xlsm 側） ---
    Set wbSut = ThisWorkbook
    Set wsTuList = wbSut.Worksheets(DST_SHEET)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' --- Tutors.xlsm を（なければ）開く ---
    On Error Resume Next
    Set wbTut = Workbooks(SRC_BOOK)
    On Error GoTo ErrHandler

    If wbTut Is Nothing Then
        ' ★修正: wbSut.Path（誤: wbSup.Path）
        If Len(Dir$(wbSut.Path & "\" & SRC_BOOK)) = 0 Then
            Err.Raise vbObjectError + 100, , "ソースファイルが見つかりません: " & wbSut.Path & "\" & SRC_BOOK
        End If
        Set wbTut = Workbooks.Open(Filename:=wbSut.Path & "\" & SRC_BOOK, ReadOnly:=True, UpdateLinks:=0)
    End If

    ' --- シート取得 ---
    On Error Resume Next
    Set wsTut = wbTut.Worksheets(SRC_SHEET)
    On Error GoTo ErrHandler
    If wsTut Is Nothing Then Err.Raise vbObjectError + 101, , "Tutors.xlsm にシート'" & SRC_SHEET & "'がありません。"

    ' --- ソース最終行（A列） ---
    lastTut = wsTut.Cells(wsTut.rowS.Count, 1).End(xlUp).Row
    If lastTut < 2 Then
        ' データなし
        wsTuList.Range("A2:B" & wsTuList.rowS.Count).ClearContents
        GoTo CleanUp
    End If

    ' --- 配列で一括転記（A:B） ---
    arr = wsTut.Range("A2:B" & lastTut).Value

    ' 転記先をクリアしてから書き込み
    wsTuList.Range("A2:B" & wsTuList.rowS.Count).ClearContents
    wsTuList.Range("A2").Resize(UBound(arr, 1), 2).Value = arr

CleanUp:
    ' 画面更新等を元に戻す
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Tutors.xlsm を閉じる（このプロシージャで開いた場合のみ閉じる）
    If Not wbTut Is Nothing Then
        ' ほかのウィンドウ状態に依らず確実に閉じる
        If wbTut.ReadOnly Then
            wbTut.Close SaveChanges:=False
        Else
            wbTut.Close SaveChanges:=False
        End If
    End If
    Exit Sub

ErrHandler:
    ' エラー内容を表示して後始末へ
    MsgBox "講師情報更新でエラーが発生しました:" & vbCrLf & Err.Description, vbCritical
    Resume CleanUp
End Sub


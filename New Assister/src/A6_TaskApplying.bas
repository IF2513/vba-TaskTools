Attribute VB_Name = "A6_TaskApplying"
Option Explicit

'================ ユーティリティ ================
' キー用正規化（前後空白除去 + 全角→半角）
Private Function NormalizeKey(ByVal v As Variant) As String
    NormalizeKey = Trim$(StrConv(CStr(v), vbNarrow))
End Function

' 文字列リストを正規化して配列に（, 、 ， / ・ 空白を区切りとする）
Private Function SplitCondList(ByVal s As String) As Variant
    Dim t As String
    t = NormalizeKey(s)
    t = Replace(t, "、", ",")
    t = Replace(t, "，", ",")
    t = Replace(t, "・", ",")
    t = Replace(t, "／", "/")
    t = Replace(t, "/", ",")
    t = Replace(t, " ", ",")
    Do While InStr(t, ",,") > 0
        t = Replace(t, ",,", ",")
    Loop
    If Left$(t, 1) = "," Then t = Mid$(t, 2)
    If Right$(t, 1) = "," Then t = Left$(t, Len(t) - 1)
    If Len(t) = 0 Then
        SplitCondList = Array()
    Else
        SplitCondList = Split(t, ",")
    End If
End Function

' グループ内OR判定（候補が空ならワイルドカード）
Private Function MatchesAny(ByVal target As String, ByVal condList As String, _
                            Optional ByVal useContains As Boolean = True) As Boolean
    Dim arr As Variant, x As Variant, tg As String
    tg = NormalizeKey(target)
    arr = SplitCondList(condList)
    If UBound(arr) < LBound(arr) Then
        MatchesAny = True    ' 条件未指定＝ワイルドカード
        Exit Function
    End If
    For Each x In arr
        x = NormalizeKey(x)
        If useContains Then
            If Len(x) > 0 And InStr(1, tg, x, vbTextCompare) > 0 Then
                MatchesAny = True: Exit Function
            End If
        Else
            If StrComp(tg, x, vbTextCompare) = 0 Then
                MatchesAny = True: Exit Function
            End If
        End If
    Next
    MatchesAny = False
End Function


' すべて数字なら先頭ゼロを除去した文字列を返す（数字以外が混じれば空文字）
Private Function StripLeadingZerosDigits(ByVal s As String) As String
    Dim i As Long, ch As String
    s = NormalizeKey(s)
    If Len(s) = 0 Then StripLeadingZerosDigits = "": Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then
            StripLeadingZerosDigits = ""
            Exit Function
        End If
    Next
    i = 1
    Do While i < Len(s) And Mid$(s, i, 1) = "0"
        i = i + 1
    Loop
    StripLeadingZerosDigits = Mid$(s, i)
    If Len(StripLeadingZerosDigits) = 0 Then StripLeadingZerosDigits = "0"
End Function

' 学年一致（TaskListのG列はカンマ区切り／TaskStatusのB列は正規化済み）
Private Function GradeMatches(condGrades As String, stuGrade As String) As Boolean
    Dim arr() As String, g As Variant
    GradeMatches = False
    If Len(Trim$(condGrades)) = 0 Then
        GradeMatches = True: Exit Function
    End If
    arr = Split(condGrades, ",")
    For Each g In arr
        If StrComp(Trim$(CStr(g)), Trim$(CStr(stuGrade)), vbTextCompare) = 0 Then
            GradeMatches = True: Exit Function
        End If
    Next
End Function

'================ メイン ================
Public Sub Task条件を生徒に適用()
    ' TaskStatus 列・行の前提
    Const FIRST_STUDENT_ROW As Long = 6
    Const FIRST_TASK_COL    As Long = 6
    Const TS_COL_ID         As Long = 1 ' A 会員番号
    Const TS_COL_GRADE      As Long = 2 ' B 学年（正規化済み）
    Const TS_COL_NAME       As Long = 3 ' C 氏名

    ' Students シートの前提
    Const STU_COL_ID        As Long = 1 ' A 会員番号
    Const STU_COL_SCHOOLCD  As Long = 4 ' D 学校コード

    ' 学校情報 シートの前提
    Const SCH_COL_CODE      As Long = 1 ' A 学校コード
    Const SCH_COL_CATEG     As Long = 3 ' C 設置区分/種別
    Const SCH_COL_TERM      As Long = 4 ' D 学期制

    Dim wsList As Worksheet, wsStatus As Worksheet, wsSchool As Worksheet, wsLog As Worksheet, wsStu As Worksheet
    Dim vList As Variant, vStatus As Variant, vSchool As Variant, vStu As Variant, vLog As Variant
    Dim dictLog As Object, dictSchool As Object, dictTask As Object, dictStu2Sch As Object
    Dim lastRow As Long, lastCol As Long, lastTaskCol As Long, lastLogRow As Long
    Dim phase As String
    Dim i As Long, r As Long, col As Long

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    phase = "シート参照"
    Set wsList = ThisWorkbook.Sheets("TaskList")
    Set wsStatus = ThisWorkbook.Sheets("TaskStatus")
    Set wsSchool = ThisWorkbook.Sheets("学校情報 from Students.xlsm")
    Set wsStu = ThisWorkbook.Sheets("Students from Students.xlsm")
    Set wsLog = ThisWorkbook.Sheets("TaskLog")

    '--- A1～最終セルで読み込み（UsedRangeのズレ回避） ---
    phase = "TaskStatus読込"
    lastCol = wsStatus.Cells(1, wsStatus.Columns.Count).End(xlToLeft).Column
    lastRow = wsStatus.Cells(wsStatus.Rows.Count, "A").End(xlUp).Row
    If lastCol < FIRST_TASK_COL Or lastRow < FIRST_STUDENT_ROW Then GoTo CleanExit
    vStatus = wsStatus.Range(wsStatus.Cells(1, 1), wsStatus.Cells(lastRow, lastCol)).Value
    lastTaskCol = UBound(vStatus, 2)

    phase = "TaskList読込"
    Dim llr As Long, llc As Long
    llr = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row
    llc = wsList.Cells(1, wsList.Columns.Count).End(xlToLeft).Column
    vList = wsList.Range(wsList.Cells(1, 1), wsList.Cells(llr, llc)).Value

    phase = "学校情報読込"
    Dim lsr As Long, lsc As Long
    lsr = wsSchool.Cells(wsSchool.Rows.Count, "A").End(xlUp).Row
    lsc = wsSchool.Cells(1, wsSchool.Columns.Count).End(xlToLeft).Column
    vSchool = wsSchool.Range(wsSchool.Cells(1, 1), wsSchool.Cells(lsr, lsc)).Value

    phase = "Students読込"
    Dim lur As Long, luc As Long
    lur = wsStu.Cells(wsStu.Rows.Count, "A").End(xlUp).Row
    luc = wsStu.Cells(1, wsStu.Columns.Count).End(xlToLeft).Column
    vStu = wsStu.Range(wsStu.Cells(1, 1), wsStu.Cells(lur, luc)).Value

    phase = "TaskLog読込"
    Dim lgr As Long, lgc As Long
    lgr = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
    lgc = wsLog.Cells(1, wsLog.Columns.Count).End(xlToLeft).Column
    vLog = wsLog.Range(wsLog.Cells(1, 1), wsLog.Cells(lgr, lgc)).Value
    lastLogRow = lgr

    '--- TaskLog 辞書（StudentID|TaskID）---
    phase = "辞書作成:Log"
    Set dictLog = CreateObject("Scripting.Dictionary")
    If IsArray(vLog) Then
        For i = 2 To UBound(vLog, 1)
            If Len(vLog(i, 1)) > 0 And Len(vLog(i, 2)) > 0 Then
                dictLog(NormalizeKey(vLog(i, 2)) & "|" & NormalizeKey(vLog(i, 1))) = True
            End If
        Next
    End If

    '--- 学校情報 辞書：キー=学校コード（文字列/数字）→ {種別, 学期制} ---
    phase = "辞書作成:School"
    Set dictSchool = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(vSchool, 1)
        Dim scRaw As String, scKey As String, scNum As String
        scRaw = vSchool(i, SCH_COL_CODE)
        scKey = NormalizeKey(scRaw)
        If Len(scKey) > 0 Then
            Dim infoArr As Variant
            infoArr = Array(Trim$(CStr(vSchool(i, SCH_COL_CATEG))), Trim$(CStr(vSchool(i, SCH_COL_TERM))))
            dictSchool(scKey) = infoArr
            scNum = StripLeadingZerosDigits(scKey)
            If Len(scNum) > 0 Then dictSchool(scNum) = infoArr
        End If
    Next

    '--- Students 辞書：キー=会員番号（文字列/数字）→ {学校コード文字列キー, 学校コード数字キー} ---
    phase = "辞書作成:Students"
    Set dictStu2Sch = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(vStu, 1)
        Dim idRaw As String, idKey As String, idNumKey As String
        Dim schRaw As String, schKey As String, schNumKey As String
        idRaw = vStu(i, STU_COL_ID)
        idKey = NormalizeKey(idRaw)
        If Len(idKey) > 0 Then
            schRaw = vStu(i, STU_COL_SCHOOLCD)
            schKey = NormalizeKey(schRaw)
            schNumKey = StripLeadingZerosDigits(schKey)
            dictStu2Sch(idKey) = Array(schKey, schNumKey)
            idNumKey = StripLeadingZerosDigits(idKey)
            If Len(idNumKey) > 0 Then dictStu2Sch(idNumKey) = Array(schKey, schNumKey)
        End If
    Next

    '--- TaskList 辞書：TaskID → {学年(G), 設置区分(H), 学期制(I)} ---
    phase = "辞書作成:TaskList"
    Set dictTask = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(vList, 1)
        If Len(vList(i, 1)) > 0 Then
            dictTask(Trim$(CStr(vList(i, 1)))) = Array( _
                Trim$(CStr(vList(i, 7))), _
                Trim$(CStr(vList(i, 8))), _
                Trim$(CStr(vList(i, 9))) _
            )
        End If
    Next

    '--- 判定＆反映（「-」座標だけ記録） ---
    phase = "判定ループ"
    Dim paintPos() As Long, pCount As Long
    ReDim paintPos(1 To 2, 1 To 5000)

    Dim taskId As String, stuID As String, key As String
    Dim condGrade As String, condCateg As String, condTerm As String
    Dim stuGrade As String, stuCateg As String, stuTerm As String
    Dim schPair As Variant

    For col = FIRST_TASK_COL To lastTaskCol
        taskId = Trim$(CStr(vStatus(1, col)))
        If Len(taskId) = 0 Then GoTo ContinueTask
        If Not dictTask.Exists(taskId) Then GoTo ContinueTask

        condGrade = dictTask(taskId)(0)
        condCateg = NormalizeKey(dictTask(taskId)(1))
        condTerm = NormalizeKey(dictTask(taskId)(2))

        For r = FIRST_STUDENT_ROW To UBound(vStatus, 1)
            stuID = NormalizeKey(vStatus(r, TS_COL_ID))
            If Len(stuID) = 0 Then GoTo ContinueStudent

            stuGrade = Trim$(CStr(vStatus(r, TS_COL_GRADE)))  ' 正規化済
            stuCateg = "": stuTerm = ""

            ' 会員番号 → 学校コード（両形式キーを試す）
            If dictStu2Sch.Exists(stuID) Then
                schPair = dictStu2Sch(stuID)
            Else
                Dim tmpIDNum As String
                tmpIDNum = StripLeadingZerosDigits(stuID)
                If Len(tmpIDNum) > 0 And dictStu2Sch.Exists(tmpIDNum) Then
                    schPair = dictStu2Sch(tmpIDNum)
                Else
                    Erase schPair
                End If
            End If

            ' 学校コード → 学校情報
            If IsArray(schPair) Then
                Dim k1 As String, k2 As String
                k1 = schPair(0): k2 = schPair(1)
                If Len(k1) > 0 And dictSchool.Exists(k1) Then
                    stuCateg = NormalizeKey(dictSchool(k1)(0))
                    stuTerm = NormalizeKey(dictSchool(k1)(1))
                ElseIf Len(k2) > 0 And dictSchool.Exists(k2) Then
                    stuCateg = NormalizeKey(dictSchool(k2)(0))
                    stuTerm = NormalizeKey(dictSchool(k2)(1))
                End If
            End If

            ' 条件判定（学年=OR一致、種別=OR/部分一致、学期制=OR/部分一致）
            Dim valid As Boolean: valid = True
            If Not GradeMatches(condGrade, stuGrade) Then valid = False
            If Not MatchesAny(stuCateg, condCateg, True) Then valid = False
            If Not MatchesAny(stuTerm, condTerm, True) Then valid = False

            ' 対象=空白 / 非対象="-"
            If valid Then
                If vStatus(r, col) = "-" Then vStatus(r, col) = ""   ' “-”だけ解除。既存の値(日付等)は触らない
            Else
                If vStatus(r, col) = "" Then
                    vStatus(r, col) = "-"
                    pCount = pCount + 1
                    If pCount > UBound(paintPos, 2) Then ReDim Preserve paintPos(1 To 2, 1 To pCount + 5000)
                    paintPos(1, pCount) = r
                    paintPos(2, pCount) = col
                End If
            End If


            ' TaskLog 追記（未登録のみ）
            key = NormalizeKey(vStatus(r, TS_COL_ID)) & "|" & taskId
            If valid And Not dictLog.Exists(key) Then
                lastLogRow = lastLogRow + 1
                wsLog.Cells(lastLogRow, "A").Value = taskId
                wsLog.Cells(lastLogRow, "B").Value = vStatus(r, TS_COL_ID)
                wsLog.Cells(lastLogRow, "C").Value = vStatus(r, TS_COL_NAME)
                wsLog.Cells(lastLogRow, "D").Value = vStatus(r, TS_COL_GRADE)
                dictLog(key) = True
            End If

ContinueStudent:
        Next r
ContinueTask:
    Next col

    '--- まず書き戻し ---
    phase = "書き戻し"
    wsStatus.Range("A1").Resize(UBound(vStatus, 1), UBound(vStatus, 2)).Value = vStatus

    '--- タスク領域の塗りつぶしを一旦全解除（前回のグレーを消す）---
    phase = "塗り全解除"
    With wsStatus.Range( _
        wsStatus.Cells(FIRST_STUDENT_ROW, FIRST_TASK_COL), _
        wsStatus.Cells(UBound(vStatus, 1), UBound(vStatus, 2)) _
    )
        .Interior.Pattern = xlNone   ' または .Interior.ColorIndex = xlNone
    End With

    '--- 「-」だけをグレー塗り ---
    phase = "グレー塗り"
    If pCount > 0 Then
        Dim j As Long
        For j = 1 To pCount
            wsStatus.Cells(paintPos(1, j), paintPos(2, j)).Interior.Color = RGB(174, 170, 170)
        Next j
    End If

CleanExit:
    On Error Resume Next
    Set dictLog = Nothing: Set dictTask = Nothing
    Set dictSchool = Nothing: Set dictStu2Sch = Nothing
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    On Error GoTo 0
    Exit Sub

CleanFail:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    On Error GoTo 0
    MsgBox "Task条件を生徒に適用でエラー (" & CStr(Err.Number) & ") at [" & phase & "] : " & Err.Description, vbCritical
End Sub



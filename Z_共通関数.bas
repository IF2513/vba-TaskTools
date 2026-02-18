Attribute VB_Name = "Z_共通関数"
Option Explicit

'==== 条件一致系 ===='
Private Function splitList(text As String) As Variant
    Dim s As String
    s = Replace(Replace(Replace(Trim(text), "　", ""), "，", ","), " ", "")
    If Len(s) = 0 Then
        splitList = Array()
    Else
        splitList = Split(s, ",")
    End If
End Function

Public Function MatchesCondition(valueText As String, condText As String) As Boolean
    ' condText が空なら常に一致。カンマ区切りのいずれか一致でOK
    Dim arr As Variant, i As Long, v As String
    If Len(condText) = 0 Then
        MatchesCondition = True
        Exit Function
    End If
    arr = splitList(condText)
    For i = LBound(arr) To UBound(arr)
        v = CStr(arr(i))
        If StrComp(CStr(valueText), v, vbTextCompare) = 0 Then
            MatchesCondition = True
            Exit Function
        End If
    Next
    MatchesCondition = False
End Function

Public Function SchoolDivisionFromGrade(gradeText As String) As String
    ' 学年から学校区分を推定（G列が無い場合に使用）
    If Left$(gradeText, 1) = "高" Then
        SchoolDivisionFromGrade = "高校"
    ElseIf Left$(gradeText, 1) = "中" Then
        SchoolDivisionFromGrade = "中学"
    Else
        SchoolDivisionFromGrade = ""
    End If
End Function

Public Function IsWithinPublicationRange(startDate, endDate, targetDate As Date) As Boolean
    Dim okStart As Boolean, okEnd As Boolean
    okStart = (IsDate(startDate) = False) Or (CDate(startDate) <= targetDate)
    okEnd = (IsDate(endDate) = False) Or (targetDate <= CDate(endDate))
    IsWithinPublicationRange = (okStart And okEnd)
End Function

Public Function IsTargetStudent(studentRow As Range, taskRow As Range) As Boolean
    ' StudentList: A=ID, B=氏名, C=ふりがな, D=学校, E=学年, F=誕生日, (任意)G=学校区分, (任意)H=学期制
    ' TaskList:   G=対象学年, H=対象学校種別, I=対象学期制, J=対象区分
    Dim grade As String, schoolType As String, term As String, division As String
    Dim condGrade As String, condSchool As String, condTerm As String, condDiv As String

    grade = CStr(studentRow.Cells(1, "E").Value2)
    ' 学校種別は StudentList D列の学校名から直接一致させる運用ならここを変更
    schoolType = CStr(studentRow.Cells(1, "D").Value2)
    ' 学期制は H列がある場合のみ
    On Error Resume Next
    term = CStr(studentRow.Cells(1, "H").Value2)
    On Error GoTo 0
    ' 学校区分（G列が無ければ学年から推定）
    division = ""
    On Error Resume Next
    division = CStr(studentRow.Cells(1, "G").Value2)
    On Error GoTo 0
    If Len(division) = 0 Then division = SchoolDivisionFromGrade(grade)

    condGrade = CStr(taskRow.Cells(1, "G").Value2)
    condSchool = CStr(taskRow.Cells(1, "H").Value2)
    condTerm = CStr(taskRow.Cells(1, "I").Value2)
    condDiv = CStr(taskRow.Cells(1, "J").Value2)

    IsTargetStudent = _
        MatchesCondition(grade, condGrade) And _
        MatchesCondition(schoolType, condSchool) And _
        MatchesCondition(term, condTerm) And _
        MatchesCondition(division, condDiv)
End Function

'==== TaskStatus 辞書化 ===='
Public Function BuildStatusDict(wsStatus As Worksheet) As Object
    ' key: studentID & "|" & taskID  -> value: 実施日（Variant/Empty可）
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long, r As Long, key As String
    lastRow = wsStatus.Cells(wsStatus.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        key = CStr(wsStatus.Cells(r, 1).Value2) & "|" & CStr(wsStatus.Cells(r, 2).Value2)
        If Not dict.Exists(key) Then dict.Add key, wsStatus.Cells(r, 3).Value
    Next
    Set BuildStatusDict = dict
End Function

Public Sub EnsureStatusRow(wsStatus As Worksheet, studentID As String, taskId As String)
    ' 対象者だけ TaskStatus にレコードが無ければ追加（実施日は空）
    Dim lastRow As Long
    If WorksheetFunction.CountIfs(wsStatus.Columns(1), studentID, wsStatus.Columns(2), taskId) = 0 Then
        lastRow = wsStatus.Cells(wsStatus.Rows.Count, 1).End(xlUp).Row + 1
        wsStatus.Cells(lastRow, 1).Value = studentID
        wsStatus.Cells(lastRow, 2).Value = taskId
        wsStatus.Cells(lastRow, 3).Value = "" ' 未実施
        wsStatus.Cells(lastRow, 4).Value = False
    End If
End Sub

' 半角変換：全角数字や記号を半角にする
Public Function Hankaku(ByVal text As String) As String
    Hankaku = StrConv(text, vbNarrow)
End Function

' 日付パース：文字列を日付型に変換。無効 or 空欄は Empty を返す
Public Function DateOrEmpty(ByVal text As String) As Variant
    Dim t As String: t = Trim$(text)
    t = Replace(t, "年", "/")
    t = Replace(t, "月", "/")
    t = Replace(t, "日", "")
    t = Replace(t, "-", "/")

    If t = "" Then
        DateOrEmpty = Empty
    ElseIf IsDate(t) Then
        DateOrEmpty = CDate(t)
    Else
        DateOrEmpty = Empty
    End If
End Function

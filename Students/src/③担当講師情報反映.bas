Attribute VB_Name = "③担当講師情報反映"
Option Explicit

' === このブック内の実際のシート名を定義 ===
Private Const SH_TUTORS  As String = "講師一覧(from Tutors.xlsm)" ' B列=講師名
Private Const SH_STUINFO As String = "生徒情報一覧"               ' I～Nに講師ラベル
Private Const SH_ASSIGN  As String = "受講・担当講師情報"         ' A:会員番号, C:教科, H:講師名

' 教科 → 生徒情報一覧の列番号(I=9～N=14)を返す
Private Function CourseToStuinfoCol(ByVal course As String) As Long
    Select Case Trim$(course)
        Case "英語": CourseToStuinfoCol = 9
        Case "数学": CourseToStuinfoCol = 10
        Case "国語": CourseToStuinfoCol = 11
        Case "理科": CourseToStuinfoCol = 12
        Case "社会": CourseToStuinfoCol = 13
        Case Else: CourseToStuinfoCol = 14   ' その他
    End Select
End Function

' ---- 講師マスタ参照（シートは存在前提）----

' 講師一覧（同ブック）B列の講師名を全取得（空なら Empty を返す）
Private Function GetAllTutorNames() As Variant
    Dim ws As Worksheet, lastR As Long
    Set ws = ThisWorkbook.Worksheets(SH_TUTORS)
    lastR = ws.Cells(ws.rowS.Count, 2).End(xlUp).Row   ' ★Rowsに修正
    If lastR < 2 Then Exit Function
    GetAllTutorNames = ws.Range(ws.Cells(2, 2), ws.Cells(lastR, 2)).Value
End Function

' 全角/半角スペースで分割して「姓」「名」を返す
Private Sub SplitNameJP(ByVal fullName As String, ByRef family As String, ByRef given As String)
    Dim s As String: s = Trim$(fullName)
    s = Replace$(s, "　", " ")
    If InStr(s, " ") > 0 Then
        family = Trim$(Split(s, " ")(0))
        given = Trim$(Mid$(s, Len(family) + 1))
    Else
        family = s: given = ""
    End If
End Sub

' 表記ゆれを判定用に正規化（出力には使わない）
Private Function CanonicalSurname(ByVal family As String) As String
    Dim s As String: s = family
    s = Replace$(s, "齋", "斎")
    s = Replace$(s, "齊", "斎")
    s = Replace$(s, "斉", "斎")
    s = Replace$(s, "邊", "辺")
    s = Replace$(s, "邉", "辺")
    CanonicalSurname = s
End Function

' マスタ基準で「紛らわしい姓」か判定（表記ゆれ or 同姓多数）
Private Function NeedsGivenInitial(ByVal familyOriginal As String) As Boolean
    Dim names As Variant, r As Long, fam As String, given As String
    Dim famC As String: famC = CanonicalSurname(familyOriginal)
    Dim sameOrigCount As Long, hasVariant As Boolean

    names = GetAllTutorNames()
    If IsEmpty(names) Then Exit Function

    For r = LBound(names, 1) To UBound(names, 1)
        Call SplitNameJP(CStr(names(r, 1)), fam, given)
        If Len(fam) > 0 Then
            If StrComp(CanonicalSurname(fam), famC, vbTextCompare) = 0 Then
                If StrComp(fam, familyOriginal, vbBinaryCompare) = 0 Then
                    sameOrigCount = sameOrigCount + 1
                Else
                    hasVariant = True
                End If
            End If
        End If
        If hasVariant Then Exit For
    Next

    NeedsGivenInitial = (hasVariant Or sameOrigCount >= 2)
End Function

' ---- ラベル生成・追記 ----

Private Function TokenExists(ByVal currentVal As String, ByVal token As String) As Boolean
    Dim t As Variant
    For Each t In Split(currentVal, ",")
        If StrComp(Trim$(CStr(t)), token, vbTextCompare) = 0 Then
            TokenExists = True: Exit Function
        End If
    Next
End Function

' 出力はオリジナル表記。必要時のみ名頭を付け、列内でユニーク化。
Private Function BuildUniqueTutorLabel(ByVal currentVal As String, ByVal familyOriginal As String, ByVal given As String) As String
    Dim label As String, n As Long
    label = familyOriginal

    If NeedsGivenInitial(familyOriginal) And Len(given) > 0 Then
        For n = 1 To Len(given)
            label = familyOriginal & Left$(given, n)
            If Not TokenExists(currentVal, label) Then Exit For
        Next
    Else
        If TokenExists(currentVal, label) And Len(given) > 0 Then
            For n = 1 To Len(given)
                label = familyOriginal & Left$(given, n)
                If Not TokenExists(currentVal, label) Then Exit For
            Next
        End If
    End If
    BuildUniqueTutorLabel = label
End Function

' 列の値に追記（"1"=未定は置換）
Private Function AppendTutorLabel(ByVal currentVal As String, ByVal familyOriginal As String, ByVal given As String) As String
    Dim cur As String: cur = Trim$(currentVal)
    If Len(Trim$(familyOriginal)) = 0 Then
        AppendTutorLabel = IIf(cur = "", "1", cur)
        Exit Function
    End If

    Dim label As String: label = BuildUniqueTutorLabel(cur, familyOriginal, given)

    If cur = "" Or cur = "1" Then
        AppendTutorLabel = label
    ElseIf TokenExists(cur, label) Then
        AppendTutorLabel = cur
    Else
        AppendTutorLabel = cur & "," & label
    End If
End Function

' ---- 生徒行の取得/クリア ----

Private Function FindStudentRow(ByVal studentId As String) As Long
    Dim ws As Worksheet, f As Range
    Set ws = ThisWorkbook.Worksheets(SH_STUINFO)
    Set f = ws.Columns(1).Find(What:=Trim$(studentId), LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then FindStudentRow = f.Row Else FindStudentRow = 0
End Function

' 対象生徒の I～N を一括クリア（I=9 列から横に6列）
Private Sub ClearStudentSubjectCols(ByVal rowS As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_STUINFO)

    ' rowS が不正なら何もしない（1行目はヘッダ想定：2以上だけ）
    If rowS < 2 Then Exit Sub

    Dim rng As Range
    Set rng = ws.Cells(rowS, 9).Resize(1, 6) ' I～N

    On Error Resume Next
    rng.ClearContents
    If Err.Number <> 0 Then
        ' 保護や結合等で失敗するケースに備えてワンチャン再実行
        Dim wasProtected As Boolean
        wasProtected = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
        If wasProtected Then ws.Unprotect
        Err.Clear
        rng.ClearContents
        If wasProtected Then ws.Protect
    End If
    On Error GoTo 0
End Sub


' ---- 個別/全体の転記実行 ----

Public Sub UpdateTutorSummaryForStudent(ByVal studentId As String)
    Dim wsA As Worksheet, wsS As Worksheet
    Set wsA = ThisWorkbook.Worksheets(SH_ASSIGN)
    Set wsS = ThisWorkbook.Worksheets(SH_STUINFO)

    Dim rowS As Long: rowS = FindStudentRow(studentId)
    If rowS = 0 Then Exit Sub

    ClearStudentSubjectCols rowS

    Dim lastA As Long: lastA = wsA.Cells(wsA.rowS.Count, 1).End(xlUp).Row   ' ★Rowsに修正
    Dim r As Long, course As String, teacherFull As String, fam As String, giv As String, tgtCol As Long
    If lastA < 2 Then Exit Sub

    For r = 2 To lastA
        If CStr(wsA.Cells(r, 1).Value) = Trim$(studentId) Then
            course = CStr(wsA.Cells(r, 3).Value)        ' C=教科
            teacherFull = CStr(wsA.Cells(r, 8).Value)   ' H=講師名（オリジナル）
            Call SplitNameJP(teacherFull, fam, giv)
            tgtCol = CourseToStuinfoCol(course)
            wsS.Cells(rowS, tgtCol).Value = AppendTutorLabel(CStr(wsS.Cells(rowS, tgtCol).Value), fam, giv)
        End If
    Next
End Sub

Public Sub UpdateTutorSummaryAll()
    Dim wsA As Worksheet
    Set wsA = ThisWorkbook.Worksheets(SH_ASSIGN)

    Dim lastA As Long: lastA = wsA.Cells(wsA.rowS.Count, 1).End(xlUp).Row   ' ★Rowsに修正
    If lastA < 2 Then Exit Sub

    Dim r As Long, sid As String
    Dim ids As Object: Set ids = CreateObject("Scripting.Dictionary")
    For r = 2 To lastA
        sid = Trim$(CStr(wsA.Cells(r, 1).Value))
        If Len(sid) > 0 Then If Not ids.Exists(sid) Then ids.Add sid, 1
    Next

    Dim key As Variant
    For Each key In ids.Keys
        UpdateTutorSummaryForStudent CStr(key)
    Next
End Sub


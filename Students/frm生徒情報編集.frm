VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm生徒情報編集 
   Caption         =   "生徒情報編集"
   ClientHeight    =   11325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12210
   OleObjectBlob   =   "frm生徒情報編集.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frm生徒情報編集"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' === シート名（ThisWorkbook 内）===
Private Const SH_STUINFO As String = "生徒情報一覧"           ' A:会員番号 B:氏名(漢字) C:ふりがな D:学校コード E:学校名 F:学年 G:学期制
Private Const SH_ASSIGN  As String = "受講・担当講師情報"     ' A:会員番号 B:会員氏名 C:教科 D:科目 E:曜日 F:コマ G:講師番号 H:講師名
Private Const SH_SCHOOL  As String = "学校情報"               ' A:学校コード B:学校名 F:学期制
Private Const SH_TUTORS  As String = "講師一覧(from Tutors.xlsm)" ' A:講師番号 B:講師名

' === 内部状態 ===
Private schCode As String
Private schTerm As String
Private originalId As String   ' 読み込み時の会員番号（保存時の行特定に使用）



' ============ 初期化 ============
Private Sub UserForm_Initialize()

    ' ▼一覧（左）：登録済み生徒
    With lstStudents
        .Clear
        .ColumnCount = 2     ' 会員番号 / 漢字氏名 / 学年 or 学校名 お好みで
        .ColumnHeads = False
        .ColumnWidths = "70;100"
        LoadAllStudentsToList
    End With

    ' ▼学年
    With cmbGrade
        .Clear
        .AddItem ""
        .AddItem "小学校1年": .AddItem "小学校2年": .AddItem "小学校3年": .AddItem "小学校4年": .AddItem "小学校5年": .AddItem "小学校6年"
        .AddItem "中学校1年": .AddItem "中学校2年": .AddItem "中学校3年"
        .AddItem "高等学校1年": .AddItem "高等学校2年": .AddItem "高等学校3年"
        .AddItem "既卒"
        .ListIndex = 0
    End With

    ' ▼学校
    LoadSchoolsToCombo

    ' ▼曜日
    With cmbDay: .Clear: .AddItem "": .AddItem "月": .AddItem "火": .AddItem "水": .AddItem "木": .AddItem "金": .AddItem "土": .ListIndex = 0: End With
    ' ▼コマ
    With cmbPeriod: .Clear: .AddItem "": .AddItem "6": .AddItem "7": .AddItem "8": .ListIndex = 0: End With
    ' ▼教科
    With cmbCourse: .Clear: .AddItem "": .AddItem "英語": .AddItem "数学": .AddItem "国語": .AddItem "理科": .AddItem "社会": .AddItem "他": .ListIndex = 0: End With
    ' ▼科目
    cmbSubject.Clear

    ' ▼講師
    LoadTutorsToCombo

    ' ▼右下の一覧（受講・担当講師）
    With lstAssignments
        .Clear
        .ColumnCount = 6   ' 教科, 科目, 曜日, コマ, 講師番号, 講師名
        .ColumnHeads = False
        .ColumnWidths = "45 pt;85 pt;30 pt;30 pt;0 pt;85 pt"
    End With

    ' ボタン表示
    cmdSave.Caption = "保存"   ' このフォームでは「登録」ではなく保存
End Sub

' ===== 生徒一覧読み込み =====
' 左の生徒一覧を、番号/名前フィルタで再構築（名前は漢字+かなの両方で部分一致）
Private Sub LoadAllStudentsToList(Optional ByVal idFilter As String = "", _
                                  Optional ByVal nameFilter As String = "")
    Dim ws As Worksheet
    Dim lastR As Long
    Dim r As Long                     ' ★ 指定どおり明示宣言
    Dim idKey As String, nmKey As String

    idKey = NormalizeKey(idFilter)
    nmKey = NormalizeKey(nameFilter)

    Set ws = ThisWorkbook.Worksheets(SH_STUINFO)
    
    With lstStudents
        .Clear
        .ColumnCount = 2
        .ColumnHeads = False
        .ColumnWidths = "70;100"
    End With
    
    lastR = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub

    Dim sid As String, nameJP As String, nameKana As String
    Dim matchNamePool As String

    For r = 2 To lastR
        sid = CStr(ws.Cells(r, 1).Value)        ' A=会員番号
        nameJP = CStr(ws.Cells(r, 2).Value)     ' B=氏名（漢字）
        nameKana = CStr(ws.Cells(r, 3).Value)   ' C=ふりがな

        ' ★ 漢字＋かなを結合した検索対象文字列
        matchNamePool = nameJP & " " & nameKana

        If MatchesFilter(sid, matchNamePool, idKey, nmKey) Then
            lstStudents.AddItem
            lstStudents.List(lstStudents.ListCount - 1, 0) = sid
            lstStudents.List(lstStudents.ListCount - 1, 1) = nameJP
        End If
    Next r
End Sub

' 検索テキスト変更イベント（都度絞り込み）
Private Sub txtFindID_Change()
    LoadAllStudentsToList txtFindID.Text, txtFindName.Text
End Sub

Private Sub txtFindName_Change()
    LoadAllStudentsToList txtFindID.Text, txtFindName.Text
End Sub

Private Sub cmdClearFilter_Click()
    txtFindID.Text = ""
    txtFindName.Text = ""
    LoadAllStudentsToList
End Sub

' ----- フィルタ判定ユーティリティ -----

' 全角→半角（英数）、小文字化、前後空白除去で正規化
Private Function NormalizeKey(ByVal s As String) As String
    s = Trim$(s)
    If s = "" Then NormalizeKey = "": Exit Function
    s = StrConv(s, vbNarrow)  ' 簡易半角化（英数）
    s = LCase$(s)
    NormalizeKey = s
End Function

' idKey / nmKey ともに部分一致（AND 条件）
Private Function MatchesFilter(ByVal sid As String, ByVal namePool As String, _
                               ByVal idKey As String, ByVal nmKey As String) As Boolean
    Dim idN As String, nmN As String
    idN = NormalizeKey(sid)
    nmN = NormalizeKey(namePool)   ' ★ 漢字+かなの結合文字列で判定

    If idKey <> "" Then
        If InStr(idN, idKey) = 0 Then Exit Function
    End If
    If nmKey <> "" Then
        If InStr(nmN, nmKey) = 0 Then Exit Function
    End If
    MatchesFilter = True
End Function

' ============ 左リスト選択 → 右側へ展開 ============
Private Sub lstStudents_Click()
    Dim i As Long: i = lstStudents.ListIndex
    If i < 0 Then Exit Sub

    Dim sid As String: sid = CStr(lstStudents.List(i, 0))
    LoadOneStudent sid
End Sub

Private Sub LoadOneStudent(ByVal sid As String)
    Dim ws As Worksheet, r As Long, lastR As Long
    Set ws = ThisWorkbook.Worksheets(SH_STUINFO)

    ' 行検索
    Dim f As Range: Set f = ws.Columns(1).Find(What:=sid, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then Exit Sub

    originalId = sid

    ' 基本情報をUIへ
    txtID.Value = CStr(ws.Cells(f.Row, 1).Value)
    Dim nameJP As String: nameJP = CStr(ws.Cells(f.Row, 2).Value)
    Dim nameKana As String: nameKana = CStr(ws.Cells(f.Row, 3).Value)

    ' 氏名を姓/名に分割（半角スペース基準）
    Dim fam As String, first As String
    fam = SplitNameLeft(nameJP): first = SplitNameRight(nameJP)
    txtFamName.Value = fam: txtFirstName.Value = first

    fam = SplitNameLeft(nameKana): first = SplitNameRight(nameKana)
    txtFamKana.Value = fam: txtFirstKana.Value = first

    schCode = CStr(ws.Cells(f.Row, 4).Value)
    cmbSchool.Value = CStr(ws.Cells(f.Row, 5).Value)
    cmbGrade.Value = CStr(ws.Cells(f.Row, 6).Value)
    schTerm = CStr(ws.Cells(f.Row, 7).Value)

    ' 受講・担当講師一覧を右下にロード
    LoadAssignmentsForStudent sid
End Sub

' 氏名の左（姓）だけ
Private Function SplitNameLeft(ByVal full As String) As String
    Dim s As String: s = Trim$(Replace(full, "　", " "))
    If InStr(s, " ") > 0 Then SplitNameLeft = Split(s, " ")(0) Else SplitNameLeft = s
End Function
' 氏名の右（名）だけ
Private Function SplitNameRight(ByVal full As String) As String
    Dim s As String: s = Trim$(Replace(full, "　", " "))
    If InStr(s, " ") > 0 Then SplitNameRight = Trim$(Mid$(s, Len(Split(s, " ")(0)) + 1)) Else SplitNameRight = ""
End Function

' ============ 学校選択でコード等取得 ============
Private Sub cmbSchool_Change()
    Dim ws As Worksheet, m As Variant, schoolName As String
    schCode = "": schTerm = ""
    schoolName = Trim$(cmbSchool.Value)
    If Len(schoolName) = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SH_SCHOOL)
    m = Application.Match(schoolName, ws.Columns(2), 0)
    If Not IsError(m) Then
        schCode = CStr(ws.Cells(m, 1).Value)  ' 学校コード
        schTerm = CStr(ws.Cells(m, 6).Value)  ' 学期制
    End If
End Sub

' ============ コンボ類のロード ============
Private Sub LoadSchoolsToCombo()
    Dim ws As Worksheet, lastRow As Long, i As Long, nm As String
    Set ws = ThisWorkbook.Worksheets(SH_SCHOOL)
    With cmbSchool
        .Clear: .AddItem "": .ListIndex = 0
        lastRow = ws.Cells(ws.rowS.Count, 2).End(xlUp).Row
        For i = 2 To lastRow
            nm = CStr(ws.Cells(i, 2).Value)
            If Len(nm) > 0 Then .AddItem nm
        Next
    End With
End Sub

Private Sub LoadTutorsToCombo()
    Dim ws As Worksheet, lastR As Long, r As Long, tid As String, tnm As String, added As Long
    Set ws = ThisWorkbook.Worksheets(SH_TUTORS)
    With cmbTeacher
        .Clear
        .ColumnCount = 2: .BoundColumn = 2
        .ColumnWidths = "0;100"
        lastR = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
        For r = 2 To lastR
            tid = Trim$(CStr(ws.Cells(r, 1).Value))
            tnm = Trim$(CStr(ws.Cells(r, 2).Value))
            If Len(tnm) > 0 Then
                .AddItem
                .List(.ListCount - 1, 0) = tid
                .List(.ListCount - 1, 1) = tnm
                added = added + 1
            End If
        Next
        .ListIndex = -1
    End With
End Sub

' 教科→科目
Private Function SubjectMap() As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    m("英語") = Array("小学英語", "中学英語", "高校英語", "OC1", "ライティング", "リーディング", "英語1", "英語2", "英語長文", "英文法", "受験英語")
    m("数学") = Array("小学算数", "中学数学", "図形", "数量", "数学1A", "数学2B", "数学3C", "数学1", "数学2", "数学3", "数学A", "数学B", "数学C", "数学基礎", "受験数学", "高校数学", "理数数学1")
    m("国語") = Array("小学国語", "小学作文", "中学国語", "中学作文", "現代文", "古典", "古典文法", "漢文", "国語総合", "国語表現1", "国語表現2", "小論文", "受験国語", "高校国語")
    m("理科") = Array("小学理科", "中学理科", "物理基礎", "物理", "生物基礎", "生物", "化学基礎", "化学", "地学基礎", "地学", "理科総合A", "理科総合B", "理科基礎", "受験理科", "高校理科", "理数化学", "理数物理")
    m("社会") = Array("小学社会", "中学社会", "中学歴史", "中学地理", "日本史A", "日本史B", "世界史A", "世界史B", "地理A", "地理B", "政治・経済", "現代社会", "受験社会", "高校社会")
    m("他") = Array("小理小社", "小算小国", "小英小算", "中英中国", "中英中数", "中理中社", "中数中国", "高英高数", "高英高国", "高英高理", "高英高社", "高数高国", "高数高理", "高国高社", "高R高W", "家庭基礎", "情報A", "全般")
    Set SubjectMap = m
End Function

Private Sub cmbCourse_Change()
    Dim m As Object: Set m = SubjectMap()
    With cmbSubject
        .Clear
        If m.Exists(Trim$(cmbCourse.Value)) Then
            Dim i As Long
            For i = LBound(m(Trim$(cmbCourse.Value))) To UBound(m(Trim$(cmbCourse.Value)))
                .AddItem m(Trim$(cmbCourse.Value))(i)
            Next
            If .ListCount > 0 Then .ListIndex = 0
        End If
    End With
End Sub

' ============ 右下の一覧：追加/更新/削除/クリア ============
Private Function ExistsExactRow(course As String, subj As String, dayW As String, period As String, tid As String) As Boolean
    Dim i As Long
    For i = 0 To lstAssignments.ListCount - 1
        If lstAssignments.List(i, 0) = course _
        And lstAssignments.List(i, 1) = subj _
        And lstAssignments.List(i, 2) = dayW _
        And lstAssignments.List(i, 3) = period _
        And lstAssignments.List(i, 4) = tid Then
            ExistsExactRow = True: Exit Function
        End If
    Next
End Function

Private Sub cmdAddRow_Click()
    Dim course$, subj$, dayW$, period$, tid$, tname$
    course = Trim$(cmbCourse.Value): subj = Trim$(cmbSubject.Value)
    dayW = Trim$(cmbDay.Value): period = Trim$(cmbPeriod.Value)
    If course = "" Or subj = "" Or dayW = "" Or period = "" Then MsgBox "教科・科目・曜日・コマは必須です。", vbExclamation: Exit Sub

    If cmbTeacher.ListIndex >= 0 Then
        tid = CStr(cmbTeacher.List(cmbTeacher.ListIndex, 0))
        tname = CStr(cmbTeacher.List(cmbTeacher.ListIndex, 1))
    End If
    If ExistsExactRow(course, subj, dayW, period, tid) Then MsgBox "同一の行が既にあります。", vbInformation: Exit Sub

    lstAssignments.AddItem
    lstAssignments.List(lstAssignments.ListCount - 1, 0) = course
    lstAssignments.List(lstAssignments.ListCount - 1, 1) = subj
    lstAssignments.List(lstAssignments.ListCount - 1, 2) = dayW
    lstAssignments.List(lstAssignments.ListCount - 1, 3) = period
    lstAssignments.List(lstAssignments.ListCount - 1, 4) = tid
    lstAssignments.List(lstAssignments.ListCount - 1, 5) = tname
End Sub

Private Sub lstAssignments_Click()
    Dim i As Long: i = lstAssignments.ListIndex
    If i < 0 Then Exit Sub

    cmbCourse.Value = lstAssignments.List(i, 0)
    cmbCourse_Change
    cmbSubject.Value = lstAssignments.List(i, 1)
    cmbDay.Value = lstAssignments.List(i, 2)
    cmbPeriod.Value = lstAssignments.List(i, 3)

    Dim tid As String: tid = lstAssignments.List(i, 4)
    Dim idx As Long
    For idx = 0 To cmbTeacher.ListCount - 1
        If CStr(cmbTeacher.List(idx, 0)) = tid Then cmbTeacher.ListIndex = idx: Exit For
    Next
End Sub

Private Sub cmdUpdateRow_Click()
    Dim i As Long: i = lstAssignments.ListIndex
    If i < 0 Then MsgBox "更新する行を選択してください。", vbExclamation: Exit Sub

    Dim course$, subj$, dayW$, period$, tid$, tname$
    course = Trim$(cmbCourse.Value): subj = Trim$(cmbSubject.Value)
    dayW = Trim$(cmbDay.Value): period = Trim$(cmbPeriod.Value)
    If course = "" Or subj = "" Or dayW = "" Or period = "" Then MsgBox "教科・科目・曜日・コマは必須です。", vbExclamation: Exit Sub

    If cmbTeacher.ListIndex >= 0 Then
        tid = CStr(cmbTeacher.List(cmbTeacher.ListIndex, 0))
        tname = CStr(cmbTeacher.List(cmbTeacher.ListIndex, 1))
    End If

    ' 自分以外との完全一致の重複を禁止
    Dim j As Long
    For j = 0 To lstAssignments.ListCount - 1
        If j <> i Then
            If lstAssignments.List(j, 0) = course _
            And lstAssignments.List(j, 1) = subj _
            And lstAssignments.List(j, 2) = dayW _
            And lstAssignments.List(j, 3) = period _
            And lstAssignments.List(j, 4) = tid Then
                MsgBox "同一の行が既にあります。", vbInformation: Exit Sub
            End If
        End If
    Next

    lstAssignments.List(i, 0) = course
    lstAssignments.List(i, 1) = subj
    lstAssignments.List(i, 2) = dayW
    lstAssignments.List(i, 3) = period
    lstAssignments.List(i, 4) = tid
    lstAssignments.List(i, 5) = tname
End Sub

Private Sub cmdRemoveRow_Click()
    Dim i As Long: i = lstAssignments.ListIndex
    If i < 0 Then MsgBox "削除する行を選択してください。", vbExclamation: Exit Sub
    lstAssignments.RemoveItem i
End Sub

Private Sub cmdClearRow_Click()
    cmbCourse.ListIndex = 0
    cmbSubject.Clear
    cmbDay.ListIndex = 0
    cmbPeriod.ListIndex = 0
    If cmbTeacher.ListCount > 0 Then cmbTeacher.ListIndex = -1 Else cmbTeacher.Value = ""
End Sub

' ============ 保存（上書き） ============
Private Sub cmdSave_Click()
    ' 必須
    If Trim$(txtID.Value) = "" Then MsgBox "会員番号を入力してください。", vbExclamation: txtID.SetFocus: Exit Sub
    If Trim$(txtFamName.Value) = "" Or Trim$(txtFirstName.Value) = "" Then MsgBox "氏名（姓・名）を入力してください。", vbExclamation: txtFamName.SetFocus: Exit Sub
    If Trim$(txtFamKana.Value) = "" Or Trim$(txtFirstKana.Value) = "" Then MsgBox "ふりがな（せい・めい）を入力してください。", vbExclamation: txtFamKana.SetFocus: Exit Sub
    If Trim$(cmbGrade.Value) = "" Then MsgBox "学年を選択してください。", vbExclamation: cmbGrade.SetFocus: Exit Sub
    If Trim$(cmbSchool.Value) = "" Then MsgBox "学校を選択してください。", vbExclamation: cmbSchool.SetFocus: Exit Sub

    ' 氏名整形
    Dim nameJP As String, nameKana As String, s As String, keyId As String
    s = Trim$(Replace$(txtFamName.Value, "　", " ")) & " " & Trim$(Replace$(txtFirstName.Value, "　", " "))
    nameJP = Trim$(Replace$(s, "　", " ")): Do While InStr(nameJP, "  ") > 0: nameJP = Replace$(nameJP, "  ", " "): Loop
    s = Trim$(Replace$(txtFamKana.Value, "　", " ")) & " " & Trim$(Replace$(txtFirstKana.Value, "　", " "))
    nameKana = Trim$(Replace$(s, "　", " ")): Do While InStr(nameKana, "  ") > 0: nameKana = Replace$(nameKana, "  ", " "): Loop
    keyId = Trim$(txtID.Value)

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_STUINFO)

    ' 上書き対象行の特定：originalId（読み込み時のID）
    Dim f As Range: Set f = ws.Columns(1).Find(What:=originalId, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then
        MsgBox "対象の生徒行が見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    ' IDを変更した場合の重複チェック
    If StrComp(originalId, keyId, vbTextCompare) <> 0 Then
        Dim f2 As Range: Set f2 = ws.Columns(1).Find(What:=keyId, LookAt:=xlWhole, MatchCase:=False)
        If Not f2 Is Nothing Then
            MsgBox "この会員番号は既に使用されています。", vbCritical
            Exit Sub
        End If
    End If

    ' ===== 行を上書き =====
    ws.Cells(f.Row, 1).Value = keyId
    ws.Cells(f.Row, 2).Value = nameJP
    ws.Cells(f.Row, 3).Value = nameKana
    ws.Cells(f.Row, 4).Value = schCode
    ws.Cells(f.Row, 5).Value = Trim$(cmbSchool.Value)
    ws.Cells(f.Row, 6).Value = Trim$(cmbGrade.Value)
    ws.Cells(f.Row, 7).Value = schTerm

    ' ===== 受講・担当講師情報を完全差し替え =====
    SaveTutorAssignments keyId, nameJP    ' ←既存の関数（当該IDの行を一旦削除→ListBox内容で再作成）

    ' ===== I～N列の講師ラベル再構築（個別） =====
    UpdateTutorSummaryForStudent keyId

    ' 左の一覧を再読み込み＆選択維持
    Dim prevTop As Long
    prevTop = 0
    On Error Resume Next
    prevTop = lstStudents.TopIndex     ' ★スクロール位置を控える
    On Error GoTo 0

    LoadAllStudentsToList              ' ★再構築

    ' 変更後ID（keyId）を優先して再選択、無ければ originalId を試す
    Dim idx As Long
    idx = FindListIndexById(lstStudents, keyId)
    If idx < 0 Then idx = FindListIndexById(lstStudents, originalId)
    If idx >= 0 Then lstStudents.ListIndex = idx

    ' ★スクロール位置を復元（範囲ガード付き）
    If lstStudents.ListCount > 0 Then
        If prevTop > lstStudents.ListCount - 1 Then
            prevTop = lstStudents.ListCount - 1
        End If
        If prevTop < 0 Then prevTop = 0
        On Error Resume Next
        lstStudents.TopIndex = prevTop
        On Error GoTo 0
    End If

    originalId = keyId
    MsgBox "保存しました。", vbInformation

End Sub

' ============ 既存：担当講師情報の読込（登録フォームと同じロジック） ============
Private Sub LoadAssignmentsForStudent(ByVal studentId As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_ASSIGN)
    Dim lastR As Long, r As Long
    lastR = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
    lstAssignments.Clear
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value) = studentId Then
            lstAssignments.AddItem
            lstAssignments.List(lstAssignments.ListCount - 1, 0) = CStr(ws.Cells(r, 3).Value) ' 教科
            lstAssignments.List(lstAssignments.ListCount - 1, 1) = CStr(ws.Cells(r, 4).Value) ' 科目
            lstAssignments.List(lstAssignments.ListCount - 1, 2) = CStr(ws.Cells(r, 5).Value) ' 曜日
            lstAssignments.List(lstAssignments.ListCount - 1, 3) = CStr(ws.Cells(r, 6).Value) ' コマ
            lstAssignments.List(lstAssignments.ListCount - 1, 4) = CStr(ws.Cells(r, 7).Value) ' 講師番号
            lstAssignments.List(lstAssignments.ListCount - 1, 5) = CStr(ws.Cells(r, 8).Value) ' 講師名
        End If
    Next
End Sub

' 受講・担当講師情報を完全差し替え保存
' A:会員番号 B:会員氏名 C:教科 D:科目 E:曜日 F:コマ G:講師番号 H:講師名
Private Sub SaveTutorAssignments(ByVal studentId As String, ByVal studentName As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_ASSIGN)
    Dim lastR As Long, r As Long

    Application.ScreenUpdating = False

    ' 1) 既存レコードを削除（該当生徒の行を全部）
    lastR = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
    For r = lastR To 2 Step -1
        If CStr(ws.Cells(r, 1).Value) = studentId Then
            ws.rowS(r).Delete
        End If
    Next r

    ' 2) 右下リスト（lstAssignments）の内容を末尾に一括追加
    If lstAssignments.ListCount > 0 Then
        Dim outStart As Long, cnt As Long, i As Long
        outStart = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row + 1
        cnt = lstAssignments.ListCount

        ' 8列分の2次元配列を作ってから一括代入（高速）
        Dim arr As Variant
        ReDim arr(1 To cnt, 1 To 8)

        For i = 1 To cnt
            arr(i, 1) = studentId
            arr(i, 2) = studentName
            arr(i, 3) = lstAssignments.List(i - 1, 0) ' 教科
            arr(i, 4) = lstAssignments.List(i - 1, 1) ' 科目
            arr(i, 5) = lstAssignments.List(i - 1, 2) ' 曜日
            arr(i, 6) = lstAssignments.List(i - 1, 3) ' コマ
            arr(i, 7) = lstAssignments.List(i - 1, 4) ' 講師番号
            arr(i, 8) = lstAssignments.List(i - 1, 5) ' 講師名
        Next

        ws.Cells(outStart, 1).Resize(cnt, 8).Value = arr
    End If

    Application.ScreenUpdating = True
End Sub

' ▼右ペインを初期化（必要なら既存の同名Subは省略可）
Private Sub ClearStudentDetailFields()
    txtID.Value = ""
    txtFamName.Value = "": txtFirstName.Value = ""
    txtFamKana.Value = "": txtFirstKana.Value = ""
    cmbSchool.Value = "": cmbGrade.Value = ""
    cmbCourse.ListIndex = 0
    cmbSubject.Clear
    cmbDay.ListIndex = 0
    cmbPeriod.ListIndex = 0
    If cmbTeacher.ListCount > 0 Then cmbTeacher.ListIndex = -1 Else cmbTeacher.Value = ""
    lstAssignments.Clear
    schCode = "": schTerm = "": originalId = ""
End Sub

' ▼（任意）会員番号でListBoxのインデックスを探す
Private Function FindListIndexById(lb As MSForms.ListBox, ByVal sid As String) As Long
    Dim i As Long
    For i = 0 To lb.ListCount - 1
        If CStr(lb.List(i, 0)) = sid Then
            FindListIndexById = i: Exit Function
        End If
    Next
    FindListIndexById = -1
End Function

' ▼削除ボタン：完全削除（生徒情報一覧＆受講・担当講師情報）
Private Sub cmdDelete_Click()
    If lstStudents.ListIndex < 0 Then
        MsgBox "削除する生徒を左の一覧から選択してください。", vbExclamation
        Exit Sub
    End If

    Dim sid As String, sname As String
    sid = CStr(lstStudents.List(lstStudents.ListIndex, 0))
    sname = CStr(lstStudents.List(lstStudents.ListIndex, 1))

    Dim ans As VbMsgBoxResult
    ans = MsgBox( _
        "この生徒を完全に削除します。よろしいですか？" & vbCrLf & vbCrLf & _
        "会員番号: " & sid & vbCrLf & "氏名: " & sname & vbCrLf & vbCrLf & _
        "※『生徒情報一覧』と『受講・担当講師情報』から削除され、元に戻せません。", _
        vbQuestion + vbYesNo + vbDefaultButton2, "削除の確認")
    If ans <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo FAIL_SAFE

    ' スクロール位置保持
    Dim prevTop As Long: On Error Resume Next: prevTop = lstStudents.TopIndex: On Error GoTo 0

    ' --- シート参照 ---
    Dim wsStu As Worksheet, wsAsn As Worksheet
    Set wsStu = ThisWorkbook.Worksheets(SH_STUINFO)
    Set wsAsn = ThisWorkbook.Worksheets(SH_ASSIGN)

    ' --- 削除処理 ---
    Dim lastR As Long, r As Long
    Dim delStu As Long, delAsn As Long
    delStu = 0: delAsn = 0

    ' 受講・担当講師情報：該当IDの全行削除
    lastR = wsAsn.Cells(wsAsn.rowS.Count, 1).End(xlUp).Row
    For r = lastR To 2 Step -1
        If CStr(wsAsn.Cells(r, 1).Value) = sid Then
            wsAsn.rowS(r).Delete
            delAsn = delAsn + 1
        End If
    Next r

    ' 生徒情報一覧：該当IDの行（通常1行）を削除
    lastR = wsStu.Cells(wsStu.rowS.Count, 1).End(xlUp).Row
    For r = lastR To 2 Step -1
        If CStr(wsStu.Cells(r, 1).Value) = sid Then
            wsStu.rowS(r).Delete
            delStu = delStu + 1
        End If
    Next r

    ' --- UI更新：右ペインクリア、左リスト再構築（スクロール復元） ---
    ClearStudentDetailFields
    LoadAllStudentsToList

    If lstStudents.ListCount > 0 Then
        If prevTop > lstStudents.ListCount - 1 Then prevTop = lstStudents.ListCount - 1
        If prevTop < 0 Then prevTop = 0
        On Error Resume Next
        lstStudents.TopIndex = prevTop
        On Error GoTo 0
    End If
    lstStudents.ListIndex = -1

    Application.ScreenUpdating = True
    MsgBox "削除しました。" & vbCrLf & _
           "・生徒情報一覧: " & delStu & " 行" & vbCrLf & _
           "・受講・担当講師情報: " & delAsn & " 行", vbInformation
    Exit Sub

FAIL_SAFE:
    Application.ScreenUpdating = True
    MsgBox "削除処理でエラーが発生しました: " & Err.Description, vbExclamation
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

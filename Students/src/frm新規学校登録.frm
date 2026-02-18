VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm新規学校登録 
   Caption         =   "学校情報登録"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11340
   OleObjectBlob   =   "frm新規学校登録.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frm新規学校登録"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSearch_Click()
    Dim ws As Worksheet
    Set ws = Workbooks("Students.xlsm").Sheets("学校コードマスタ")
    
    Dim lastRow As Long: lastRow = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row
    Dim keyword As String: keyword = Trim$(txtSearch.Value)
    
    If keyword = "" Then
        MsgBox "学校名を入力してください。", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long, nameCell As String
    Dim resultList As Collection: Set resultList = New Collection
    
    For i = 2 To lastRow
        nameCell = ws.Cells(i, 6).Value
        If InStr(nameCell, keyword) > 0 Then
            ' 学校名を表示
            resultList.Add nameCell
        End If
    Next i
    
    ' 結果をListBoxに表示
    lstCandidates.Clear
    If resultList.Count = 0 Then
        lstCandidates.AddItem "該当する学校が見つかりませんでした。"
    Else
        For i = 1 To resultList.Count
            lstCandidates.AddItem resultList(i)
        Next i
    End If
End Sub

Private Sub lstCandidates_Click()
    If lstCandidates.ListIndex < 0 Then Exit Sub
    If lstCandidates.Value = "該当する学校が見つかりませんでした。" Then Exit Sub

    Dim schoolName As String: schoolName = CStr(lstCandidates.Value)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("学校コードマスタ")
    Dim lastRow As Long: lastRow = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If CStr(ws.Cells(i, 6).Value) = schoolName Then
            Dim code As String:      code = CStr(ws.Cells(i, 1).Value)
            Dim kindCode As String:  kindCode = CStr(ws.Cells(i, 2).Value)
            Dim prefCode As String:  prefCode = CStr(ws.Cells(i, 3).Value)
            Dim categCode As String: categCode = CStr(ws.Cells(i, 4).Value)

            ' ===== 種別（先頭2文字で判定・安全化）=====
            Dim kindCode2 As String
            kindCode2 = Left$(Trim$(kindCode), 2)
            Dim kind As String
            Select Case kindCode2
                Case "B1": kind = "小学校"
                Case "C1": kind = "中学校"
                Case "C2": kind = "義務教育学校"
                Case "D1": kind = "高等学校"
                Case "D2": kind = "中等教育学校"
                Case Else: kind = "その他"
            End Select

            ' ===== 都道府県 =====
            Dim prefCode2 As String
            prefCode2 = Right$("00" & Left$(Trim$(prefCode), 2), 2)

            Dim arr
            arr = Array("北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", _
                        "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県", _
                        "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県", "静岡県", _
                        "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県", _
                        "鳥取県", "島根県", "岡山県", "広島県", "山口県", "徳島県", "香川県", "愛媛県", _
                        "高知県", "福岡県", "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県")

            Dim pref As String
            If IsNumeric(prefCode2) Then
                Dim idx As Integer: idx = CInt(prefCode2)
                If idx >= 1 And idx <= 47 Then
                    pref = arr(idx - 1)
                Else
                    pref = "不明"
                End If
            Else
                pref = "不明"
            End If

            ' ===== 設置区分 =====
            Dim instCode2 As String: instCode2 = Left$(Trim$(categCode), 1)

            Dim categ As String
            If instCode2 = "1" Then
                categ = "国立"
            ElseIf instCode2 = "3" Then
                categ = "私立"
            Else
                ' instCode2 = "2"（公立）
                ' まず学校名から 市/区/町/村立 を優先抽出
                Dim catFromName As String: catFromName = ""
                Dim posK As Long, posS As Long, posT As Long, posM As Long, posP As Long
                posK = InStr(1, schoolName, "区立")
                posS = InStr(1, schoolName, "市立")
                posT = InStr(1, schoolName, "町立")
                posM = InStr(1, schoolName, "村立")
                posP = InStr(1, schoolName, "県立") ' 県立が学校名に含まれるか

                If posK > 0 Then
                    catFromName = Left$(schoolName, posK + Len("区立") - 1)
                ElseIf posS > 0 Then
                    catFromName = Left$(schoolName, posS + Len("市立") - 1)
                ElseIf posT > 0 Then
                    catFromName = Left$(schoolName, posT + Len("町立") - 1)
                ElseIf posM > 0 Then
                    catFromName = Left$(schoolName, posM + Len("村立") - 1)
                ElseIf posP > 0 Then
                    ' 「県立」が含まれる場合は県名付きに正規化
                    If pref <> "不明" Then
                        catFromName = pref & "立"       ' 例：「神奈川県立」
                    Else
                        catFromName = "県立"            ' 県が不明なら暫定
                    End If
                End If

                If LenB(catFromName) > 0 Then
                    categ = catFromName                 ' 例：「横浜市立」「世田谷区立」「神奈川県立」
                Else
                    ' 学校名に設置表記が無い場合のフォールバック
                    Dim head As String
                    Select Case True
                        Case pref Like "*東京都*": head = "都立"
                        Case pref Like "*北海道*": head = "道立"
                        Case pref Like "*大阪府*" Or pref Like "*京都府*": head = "府立"
                        Case Else: head = "県立"
                    End Select

                    ' ★要件：県立は県名付きにする
                    If head = "県立" And pref <> "不明" Then
                        categ = pref & "立"             ' 例：「神奈川県立」
                    Else
                        categ = head                    ' 都立/道立/府立は従来どおり
                    End If
                End If
            End If


            ' ===== 学校名整形 =====
            Dim fixedName As String
            fixedName = Trim$(Replace$(Replace$(schoolName, "　", ""), " ", "")) ' 全角/半角空白除去

            If instCode2 <> "3" Then
                ' 1) カテゴリ名そのものが先頭にあれば、まずそれを除去（横浜市立○○、都立○○、県立○○など）
                If Left$(fixedName, Len(categ)) = categ Then
                    fixedName = Mid$(fixedName, Len(categ) + 1)
                Else
                    ' 2) 先頭から最初の「立」までを除去（◯◯市立/県立/都立など広く拾う）
                    Dim pos As Long
                    pos = InStr(1, fixedName, "立")
                    If pos >= 2 And pos <= 10 Then
                        fixedName = Mid$(fixedName, pos + 1)
                    Else
                        ' 3) 念のための既知接頭辞チェック
                        Dim prefixes As Variant: prefixes = Array("国立", "道立", "府立", "県立", "市立", "区立", "町立", "村立", "都立")
                        Dim p As Variant
                        For Each p In prefixes
                            If Left$(fixedName, Len(p)) = p Then
                                fixedName = Mid$(fixedName, Len(p) + 1)
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If


            ' ===== フォームに反映 =====
            txtCode.Value = code
            txtKind.Value = kind
            txtPref.Value = pref
            txtCateg.Value = categ
            txtSchoolName.Value = fixedName
            
            Exit Sub
        End If
    Next i
End Sub

Private Sub UserForm_Initialize()
    With cmbTerm
        .Clear
        .List = Array("", "2学期制", "3学期制", "不明")
        .Value = ""                     ' 既定は空欄
        .MatchRequired = True
    End With
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRegister_Click()
    ' ===== 必須チェック =====
    If Trim$(txtSchoolName.Value) = "" Then
        MsgBox "学校名が空です。", vbExclamation
        Exit Sub
    End If
    If Trim$(txtCode.Value) = "" Then
        ' コード未設定でも続けたい場合はこのチェックを弱めてOK
        If MsgBox("学校コードが空です。続行しますか？", vbExclamation + vbYesNo) = vbNo Then Exit Sub
    End If
    If LenB(cmbTerm.Value) = 0 Then
        cmbTerm.Value = "不明" ' 未選択なら不明を入れる
    End If

    ' ===== 書き込み先 =====
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Workbooks("Students.xlsm").Worksheets("学校情報")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "書込先シート 'Students.xlsm' の '学校情報' が見つかりません。", vbCritical
        Exit Sub
    End If

    ' ===== 既存行検索（学校コード優先→学校名で代替）=====
    Dim keyCode As String: keyCode = Trim$(txtCode.Value)
    Dim keyName As String: keyName = Trim$(txtSchoolName.Value)

    Dim found As Range
    If keyCode <> "" Then
        Set found = ws.Columns(1).Find(What:=keyCode, LookAt:=xlWhole, MatchCase:=False) ' Col A: コード想定
    End If
    If found Is Nothing Then
        Set found = ws.Columns(2).Find(What:=keyName, LookAt:=xlWhole, MatchCase:=False) ' Col B: 学校名想定
    End If

    ' ===== 上書き確認 → 既存行を削除 =====
    If Not found Is Nothing Then
        If MsgBox("既に登録があります。上書き（末尾へ再登録）しますか？", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
        ws.rowS(found.Row).Delete
    End If

    ' ===== 追記 =====
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rowS.Count, 1).End(xlUp).Row + 1

    ' 列マッピング：A:コード / B:学校名 / C:都道府県 / D:種別 / E:設置区分 / F:学期制
    ws.Cells(lastRow, 1).Value = Trim$(txtCode.Value)
    ws.Cells(lastRow, 2).Value = Trim$(txtSchoolName.Value)
    ws.Cells(lastRow, 3).Value = Trim$(txtPref.Value)
    ws.Cells(lastRow, 4).Value = Trim$(txtKind.Value)
    ws.Cells(lastRow, 5).Value = Trim$(txtCateg.Value)
    ws.Cells(lastRow, 6).Value = Trim$(cmbTerm.Value)

    MsgBox "登録しました。", vbInformation

    ' ===== 入力クリア =====
    ClearFormFields
End Sub

Private Sub ClearFormFields()
    txtCode.Value = ""
    txtSchoolName.Value = ""
    txtPref.Value = ""
    txtKind.Value = ""
    txtCateg.Value = ""
    cmbTerm.Value = ""
    lstCandidates.Clear
End Sub

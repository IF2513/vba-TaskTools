Attribute VB_Name = "A3_生徒一覧toTaskStatus"
Option Explicit

Public Sub TaskStatusに生徒一覧反映()
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim lastRow As Long, writeRow As Long
    Dim fullGrade As String, shortGrade As String
    Dim r As Long
    Dim orderList As Variant, g As Variant
    
    ' === シート参照 ===
    Set wsSrc = ThisWorkbook.Sheets("Students from Students.xlsm")
    Set wsDst = ThisWorkbook.Sheets("TaskStatus")
    
    ' === 出力先初期化（A:Dまでクリアに拡張） ===
    wsDst.Range("A6:D" & wsDst.Rows.Count).ClearContents
    writeRow = 6
    
    ' === 並び順定義（降順） ===
    orderList = Array("高3", "他", "高2", "高1", _
                      "中3", "中2", "中1", _
                      "小6", "小5", "小4", "小3", "小2", "小1")
    
    ' === 生徒一覧最終行 ===
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    
    ' === 並び順で走査 ===
    For Each g In orderList
        For r = 2 To lastRow
            fullGrade = Trim$(wsSrc.Cells(r, "F").Value)
            shortGrade = ConvertGrade(fullGrade)
            
            If shortGrade = g Then
                ' A:会員番号 / B:学年(短縮) / C:氏名 / D:担当講師（I～Nを重複排除で結合）
                wsDst.Cells(writeRow, "A").Value = wsSrc.Cells(r, "A").Value
                wsDst.Cells(writeRow, "B").Value = shortGrade
                wsDst.Cells(writeRow, "C").Value = wsSrc.Cells(r, "B").Value
                wsDst.Cells(writeRow, "D").Value = BuildTutorList(wsSrc, r)  ' ★追加
                writeRow = writeRow + 1
            End If
        Next r
    Next g
    
    'MsgBox "TaskStatusに生徒一覧を学年順で反映しました。", vbInformation
End Sub

' === 学年表記変換 ===
Private Function ConvertGrade(ByVal fullGrade As String) As String
    Select Case True
        Case InStr(fullGrade, "高等学校3") > 0: ConvertGrade = "高3"
        Case InStr(fullGrade, "高等学校2") > 0: ConvertGrade = "高2"
        Case InStr(fullGrade, "高等学校1") > 0: ConvertGrade = "高1"
        Case InStr(fullGrade, "中学校3") > 0: ConvertGrade = "中3"
        Case InStr(fullGrade, "中学校2") > 0: ConvertGrade = "中2"
        Case InStr(fullGrade, "中学校1") > 0: ConvertGrade = "中1"
        Case InStr(fullGrade, "小学校6") > 0: ConvertGrade = "小6"
        Case InStr(fullGrade, "小学校5") > 0: ConvertGrade = "小5"
        Case InStr(fullGrade, "小学校4") > 0: ConvertGrade = "小4"
        Case InStr(fullGrade, "小学校3") > 0: ConvertGrade = "小3"
        Case InStr(fullGrade, "小学校2") > 0: ConvertGrade = "小2"
        Case InStr(fullGrade, "小学校1") > 0: ConvertGrade = "小1"
        Case InStr(fullGrade, "既卒") > 0 Or InStr(fullGrade, "その他") > 0: ConvertGrade = "他"
        Case Else: ConvertGrade = fullGrade ' 該当なしは原文ママ
    End Select
End Function

' === I～N列の担当講師を重複なしで結合（半角カンマ区切り） ===
' ・各セル内の複数講師「A,B」も分解
' ・全角カンマ（，・、）や全角/半角空白も正規化して扱う
Private Function BuildTutorList(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim c As Long, raw As String, parts As Variant, i As Long
    Dim nm As String
    
    ' I(9)～N(14) の6教科
    For c = 9 To 14
        raw = CStr(ws.Cells(rowNum, c).Value)
        If Len(raw) > 0 Then
            ' 全角カンマ・読点を半角カンマに、全角空白を半角に正規化
            raw = Replace(raw, "，", ",")
            raw = Replace(raw, "、", ",")
            raw = Replace(raw, "　", " ")
            ' 区切って走査
            parts = Split(raw, ",")
            For i = LBound(parts) To UBound(parts)
                nm = Trim$(parts(i))
                If Len(nm) > 0 Then
                    ' カンマ内にスペース混入対策（"A, B"→"A","B"）
                    nm = Trim$(nm)
                    If Not dict.Exists(nm) Then dict.Add nm, True
                End If
            Next i
        End If
    Next c
    
    If dict.Count = 0 Then
        BuildTutorList = ""
    Else
        ' 追加順のまま連結（半角カンマ区切り）
        Dim arr() As String, k As Long
        ReDim arr(0 To dict.Count - 1)
        k = 0
        Dim key As Variant
        For Each key In dict.Keys
            arr(k) = CStr(key)
            k = k + 1
        Next key
        BuildTutorList = Join(arr, ",")
    End If
End Function



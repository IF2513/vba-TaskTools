Attribute VB_Name = "A5_TaskLog反映"
Option Explicit

Public Sub TaskLog反映()
    Dim wsLog As Worksheet, wsStatus As Worksheet
    Dim lastRow As Long, r As Long
    Dim stuID As Variant
    Dim taskId As String
    Dim compDate As Variant
    
    ' === シート参照 ===
    Set wsLog = ThisWorkbook.Sheets("TaskLog")
    Set wsStatus = ThisWorkbook.Sheets("TaskStatus")
    
    ' === TaskStatusの範囲特定 ===
    Dim lastCol As Long, lastStuRow As Long
    lastCol = wsStatus.Cells(1, wsStatus.Columns.Count).End(xlToLeft).Column
    lastStuRow = wsStatus.Cells(wsStatus.Rows.Count, "A").End(xlUp).Row
    
    ' === TaskID辞書（TaskStatus 1行目） ===
    Dim dicTask As Object: Set dicTask = CreateObject("Scripting.Dictionary")
    Dim c As Long, idText As String
    For c = 6 To lastCol      ' TaskStatus のタスク列が6列目以降想定
        idText = Trim$(wsStatus.Cells(1, c).Value)
        If Len(idText) > 0 Then dicTask(idText) = c
    Next c
    
    ' === Student辞書（TaskStatus A列） ===
    Dim dicStu As Object: Set dicStu = CreateObject("Scripting.Dictionary")
    Dim rStu As Long
    For rStu = 6 To lastStuRow
        If Len(wsStatus.Cells(rStu, "A").Value) > 0 Then
            dicStu(CStr(wsStatus.Cells(rStu, "A").Value)) = rStu
        End If
    Next rStu
    
    ' === TaskLogの最終行 ===
    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
    
    ' === 転記＋色変更処理 ===
    Dim tgtCell As Range
    For r = 2 To lastRow
        taskId = Trim$(CStr(wsLog.Cells(r, "A").Value)) ' A列＝TaskID
        stuID = wsLog.Cells(r, "B").Value               ' B列＝StudentID
        compDate = wsLog.Cells(r, "E").Value            ' E列＝CompletedDate
        
        If dicStu.Exists(CStr(stuID)) And dicTask.Exists(taskId) Then
            Set tgtCell = wsStatus.Cells(dicStu(CStr(stuID)), dicTask(taskId))
            
            ' --- 日付反映 (mm/dd 表記) ---
            If IsDate(compDate) Then
                tgtCell.NumberFormatLocal = "mm/dd"
                tgtCell.Value = CDate(compDate)
                tgtCell.Interior.Color = RGB(174, 170, 170) ' 完了済セルをグレー
            End If
        End If
    Next r
End Sub


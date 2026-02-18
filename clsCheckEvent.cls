VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCheckEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents chk As MSForms.CheckBox
Attribute chk.VB_VarHelpID = -1
Private parentForm As Object

Public Sub Init(ByVal c As MSForms.CheckBox, ByVal frm As Object)
    Set chk = c
    Set parentForm = frm
End Sub

' ===== 改良版：クリックしたチェックだけ残す =====
Private Sub chk_Click()
    Dim tagValue As String
    tagValue = chk.Tag

    Dim ctl As Control
    Dim thisName As String
    thisName = chk.Name

    ' クリックしたCheckBoxを一時保存
    Dim wasChecked As Boolean
    wasChecked = chk.Value

    ' 同じ行の他をすべてOFF
    For Each ctl In parentForm.FrameScroll.Controls
        If TypeName(ctl) = "CheckBox" Then
            If ctl.Tag = tagValue And ctl.Name <> thisName Then
                ctl.Value = False
            End If
        End If
    Next ctl

    ' クリックしたチェックボックスが押されていたら再度ONにする
    If wasChecked Then chk.Value = True
End Sub


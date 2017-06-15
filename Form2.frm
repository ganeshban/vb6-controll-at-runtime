VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim WithEvents Cmd1 As CommandButton
Attribute Cmd1.VB_VarHelpID = -1
'
Private Sub Form_Load()
Dim i As Integer
For i = 1 To 4
    Create_Button CStr(i)
Next i
End Sub
'
Private Sub Cmd1_click()
  MsgBox "I have been Created Dynamically at Run-time", _
    , "Dynamic Controls"
End Sub


Private Sub Create_Button(BtnName As String)
  Set Cmd1 = Controls.Add("vb.commandbutton", "btn" & BtnName)
  Cmd1.Width = 2000
  Cmd1.Top = Me.Height / 2 - Cmd1.Height / 2 - 100
  Cmd1.Left = CInt(BtnName) * 2500
  Cmd1.Caption = BtnName
  Cmd1.Visible = True
End Sub


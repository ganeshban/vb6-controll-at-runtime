VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdButton(4) As CommandButton
 
Private Sub Form_Load()
 
    Dim i As Integer
 
    For i = 0 To 4
        Set cmdButton(i) = Me.Controls.Add("VB.CommandButton", "cmdButton" & Me.Controls.Count)
        With cmdButton(i)
            .Left = 750 * i
            .Top = 1000
            .Width = 700
            .Height = 500
            .Caption = "Hello"
            .Visible = True
        End With
    Next i
 
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
 
    Dim i As Integer
 
    For i = 0 To 4
        Set cmdButton(i) = Nothing
    Next i
     
End Sub

VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblTemplate 
      Caption         =   "Label1"
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
' load a new control into the control array
i = i + 1
Load lblTemplate(i)
' position the new control on the form and add a caption
lblTemplate(i).Left = lblTemplate(0).Left
lblTemplate(i).Top = lblTemplate(i - 1).Top + _
lblTemplate(i - 1).Height + 100
lblTemplate(i).Caption = "index = " _
& Str$(i)
' make the control visible
lblTemplate(i).Visible = True
End Sub

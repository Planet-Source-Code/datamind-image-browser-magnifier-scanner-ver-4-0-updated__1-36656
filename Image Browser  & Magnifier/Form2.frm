VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "confirmation"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Des = Text1.Text
If Des = "yes" Then
MsgBox ("Thank you for running IMAGE BROWSER & MAGNIFIER")
End
Else
If Des = "no" Then
Form2.Hide
End If
End If
End Sub

Private Sub Form_Load()
MsgBox ("type yes if you want to exit or no to abord")
End Sub

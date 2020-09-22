VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "My Browser"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "FORM1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "FORM1.frx":08CA
   MousePointer    =   4  'Icon
   Picture         =   "FORM1.frx":0BD4
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   480
      Pattern         =   "*.WMF;*.BMP;*.JPG;*.ICO;*.CUR"
      TabIndex        =   2
      Top             =   3120
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      DragIcon        =   "FORM1.frx":2E76
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   0
      Width           =   3375
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3015
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu scannerItem 
         Caption         =   "Scanner"
      End
      Begin VB.Menu mnuPrintItem 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuSaveItem 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuOpenItem 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuClock 
      Caption         =   "&Clock"
      Begin VB.Menu mnuDateItem 
         Caption         =   "&Date"
      End
      Begin VB.Menu mnuTimeItem 
         Caption         =   "&Time"
      End
   End
   Begin VB.Menu mnuPreferenses 
      Caption         =   "&Preferenses"
      Begin VB.Menu mnuItemWrite 
         Caption         =   "&Write a note..."
      End
      Begin VB.Menu mnuTextItem 
         Caption         =   "&Text"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowtoItem 
         Caption         =   "How to..."
      End
      Begin VB.Menu mnuAboutItem 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPhone 
      Caption         =   "&Phone manager"
      Begin VB.Menu mnudialItem 
         Caption         =   "&Dial"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveHandler
Dir1.Path = Drive1.Drive
DriveHandler:
Call DH
End Sub

Private Sub File1_Click()
SelectedFile = File1.Path & "\" & File1.FileName
Image1.Picture = LoadPicture(SelectedFile)
Image2.Picture = LoadPicture(SelectedFile)
End Sub

Private Sub Image2_Click()
Image2.Height = Image2.Height + 300
Image2.Width = Image2.Width + 300
End Sub

Private Sub mnuAboutItem_Click()
frmAbout.Show
End Sub

Private Sub mnuDateItem_Click()
Label1.Caption = Date
Label1.FontBold = True
End Sub

Private Sub mnudialItem_Click()
Shell ("c:\Windows\Dialer.exe")
End Sub

Private Sub mnuExitItem_Click()
Q = InputBox("Are you sure you want to exit the program?", "Info")
If Q = "yes" Then
MsgBox " Thank you for using this program ", vbInformation, "Info"
End
ElseIf Q = "no" Then
End If
End Sub

Private Sub mnuHowtoItem_Click()
MsgBox ("Magnify a picture:Select the picture and click on the second picture as many times as you wish")
MsgBox ("Make a phone call:Click dial in the Phone manager menu.Then the dialer will show up.")
End Sub

Private Sub mnuItemWrite_Click()
Shell ("C:\Windows\Notepad.exe")
End Sub

Private Sub mnuOpenItem_Click()
CommonDialog1.Filter = "Metafiles (*.WMF)|*.WMF|Bitmaps (*.BMP)|*.BMP|Jpegimage (*.JPG)|*.JPG"
CommonDialog1.ShowOpen
Image1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub mnuPrintItem_Click()
Form1.PrintForm
End Sub

Private Sub mnuSaveItem_Click()
SavePicture Image1.Picture, App.Path & "\yourpic.jpg"
MsgBox "The picture was saved in Image Browser Magnifier Scanner directory as yourpic.jpg "
End Sub

Private Sub mnuTextItem_Click()
CommonDialog1.Flags = &H1&
CommonDialog1.ShowColor
Label1.ForeColor = CommonDialog1.Color
End Sub

Private Sub mnuTimeItem_Click()
Label1.Caption = Time
Label1.FontBold = True
End Sub

Private Sub scannerItem_Click()
Form3.Show
End Sub

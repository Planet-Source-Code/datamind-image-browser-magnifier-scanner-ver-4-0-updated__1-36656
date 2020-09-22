VERSION 5.00
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.1#0"; "IMGEDIT.OCX"
Begin VB.Form Form3 
   Caption         =   "SCANNER"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ScanLibCtl.ImgScan ImgScan1 
      Left            =   360
      Top             =   8640
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   0
      DestImageControl=   "ImgEdit1"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   8640
      Width           =   1215
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   7815
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   10215
      _Version        =   131073
      _ExtentX        =   18018
      _ExtentY        =   13785
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit1"
      BeginProperty AnnotationFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UndoBufferSize  =   139964416
      OcrZoneVisibility=   -4044
      AnnotationOcrType=   127
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ScannerHandle
With ImgScan1
.ScannerAvailable
.OpenScanner
.StartScan
.CloseScanner
ImgScan1.Image = ImgEdit1.Image
ScannerHandle:
Call DH
End With
End Sub

Private Sub Form_Load()
Form3.Icon = Form1.Icon
End Sub

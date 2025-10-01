VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Window 
   Caption         =   "Window"
   ClientHeight    =   11400
   ClientLeft      =   4425
   ClientTop       =   2085
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   11400
   ScaleWidth      =   10845
   Begin MSComctlLib.ProgressBar renderProgress 
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   240
   End
   Begin VB.PictureBox paintBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   11220
      Left            =   0
      ScaleHeight     =   11190
      ScaleWidth      =   10785
      TabIndex        =   0
      Top             =   140
      Width           =   10815
   End
End
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub drawPixel(x, y, color As Long)
    ' 15 "Points" to a pixel
    adj = 15
    
    paintBox.PSet (x * adj, y * adj), color
End Sub

Private Sub Form_Resize()
    ' paintBox height 4140
    ' paintBox width 7455
    ' window hight 4725
    ' window width 7575
    
    paintBox.Height = Window.Height - 585
    paintBox.Width = Window.Width - 120
    
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Dim thisMain As main
    Set thisMain = New main
End Sub

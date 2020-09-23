VERSION 5.00
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Main"
   ClientHeight    =   5325
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7695
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctBorderRight 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   7680
      ScaleHeight     =   4875
      ScaleWidth      =   15
      TabIndex        =   3
      Top             =   435
      Width           =   15
   End
   Begin VB.PictureBox pctBorderLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   0
      ScaleHeight     =   4875
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   435
      Width           =   15
   End
   Begin VB.PictureBox pctBorderBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   7695
      TabIndex        =   1
      Top             =   5310
      Width           =   7695
   End
   Begin AppTemplate.TabHead TabHead1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   767
      TabString       =   "Control,About"
      SelectedBold    =   -1  'True
      InactiveColor   =   -2147483640
      BackColor       =   -2147483633
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "&Show"
      Begin VB.Menu mnuShowControl 
         Caption         =   "&Control"
      End
      Begin VB.Menu mnuShowAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    Me.Caption = App.Title
    gFrmExt.ShowMe frmControl
End Sub

Private Sub MDIForm_Resize()
    gFrmExt.DoResize gFrmExt.frmActive
End Sub

Private Sub mnuHelpAbout_Click()
    gFrmExt.ShowMe frmAbout
End Sub

Private Sub mnuShowAbout_Click()
    TabHead1.SelectedTab = 1
    gFrmExt.ShowMe frmAbout
End Sub

Private Sub mnuShowControl_Click()
    TabHead1.SelectedTab = 0
    gFrmExt.ShowMe frmControl
End Sub

Private Sub TabHead1_TabSelected(iIndex As Integer, sName As String)
    Select Case iIndex
        Case 0
            gFrmExt.ShowMe frmControl
        Case 1
            gFrmExt.ShowMe frmAbout
    End Select
End Sub

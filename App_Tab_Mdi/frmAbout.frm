VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "About"
      Height          =   3495
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5655
      Begin VB.Label Label4 
         Caption         =   $"frmAbout.frx":0000
         Height          =   1155
         Left            =   180
         TabIndex        =   4
         Top             =   2160
         Width           =   5115
      End
      Begin VB.Label Label3 
         Caption         =   "Licence"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   1860
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   $"frmAbout.frx":0155
         Height          =   1155
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Application Template : MDI Tab"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Frame1.Width = Me.Width - (Frame1.Left * 2)
    Frame1.Height = Me.Height - (Frame1.Top * 2)
End Sub

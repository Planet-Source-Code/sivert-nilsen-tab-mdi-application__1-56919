VERSION 5.00
Begin VB.Form frmControl 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Control this"
      Height          =   3195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5895
      Begin VB.Label Label5 
         Caption         =   "Error handling, and logging to file is included in the ""modError"" module."
         Height          =   435
         Left            =   300
         TabIndex        =   5
         Top             =   2220
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "Errors included."
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
         Left            =   300
         TabIndex        =   4
         Top             =   1920
         Width           =   3795
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   300
         Picture         =   "frmControl.frx":0000
         Top             =   1260
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   300
         Picture         =   "frmControl.frx":0102
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "The tab selector is located on the MDI parent window and is a usercontrol you can edit"
         Height          =   675
         Left            =   660
         TabIndex        =   3
         Top             =   1260
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   $"frmControl.frx":024C
         Height          =   675
         Left            =   660
         TabIndex        =   2
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "This is actually a form, not a Tab."
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3795
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Frame1.Width = Me.Width - (Frame1.Left * 2)
    Frame1.Height = Me.Height - (Frame1.Top * 2)
End Sub

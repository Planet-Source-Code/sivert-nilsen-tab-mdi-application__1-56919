VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormExtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public frmActive As Form

'
' Resize the window to fit the mdimain window
'
Public Sub DoResize(theform As Form)
    On Error GoTo errh
    
    Dim iBorder As Integer
    iBorder = 0
    If Not theform Is Nothing Then
        theform.Width = mdiMain.ScaleWidth
        theform.Height = mdiMain.ScaleHeight
        theform.Left = 0
        theform.Top = 0
    End If
    
    Exit Sub
errh:
    HandleError Err.Description, Err.Number, "FormExtender:DoResize()"
    Resume Next
End Sub

'
' hide current active form and show the one in the parameter
'
Public Sub ShowMe(theform As Form)
    On Error GoTo errh
    
    If Not frmActive Is Nothing Then frmActive.Visible = False
    DoResize theform
    theform.Visible = True
    Set frmActive = theform
    
    Exit Sub
errh:
    HandleError Err.Description, Err.Number, "FormExtender:ShowMe()"
    Resume Next
End Sub

'
' Translate all texts on form
'
Public Sub TranslateMe(theform As Form)
    On Error GoTo errh
    
    
    Exit Sub
errh:
    HandleError Err.Description, Err.Number, "FormExtender:TranslateMe()"
    Resume Next
End Sub



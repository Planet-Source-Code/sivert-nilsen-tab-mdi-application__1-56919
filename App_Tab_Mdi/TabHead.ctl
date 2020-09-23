VERSION 5.00
Begin VB.UserControl TabHead 
   Alignable       =   -1  'True
   BackColor       =   &H80000010&
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8460
   ControlContainer=   -1  'True
   ScaleHeight     =   630
   ScaleWidth      =   8460
   ToolboxBitmap   =   "TabHead.ctx":0000
   Begin VB.Timer tmrMouseOver 
      Interval        =   1000
      Left            =   7920
      Top             =   60
   End
   Begin VB.Shape shpFocus 
      BorderStyle     =   3  'Dot
      Height          =   255
      Left            =   300
      Top             =   180
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line lnTabShadow 
      BorderColor     =   &H80000015&
      X1              =   1080
      X2              =   1080
      Y1              =   120
      Y2              =   420
   End
   Begin VB.Label lblTabItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   11
      Left            =   7140
      MouseIcon       =   "TabHead.ctx":0312
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblTabItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   10
      Left            =   6540
      MouseIcon       =   "TabHead.ctx":0464
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblTabItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   9
      Left            =   5940
      MouseIcon       =   "TabHead.ctx":05B6
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblTabItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   8
      Left            =   5340
      MouseIcon       =   "TabHead.ctx":0708
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblTabItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   7
      Left            =   4740
      MouseIcon       =   "TabHead.ctx":085A
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblTabItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   6
      Left            =   4140
      MouseIcon       =   "TabHead.ctx":09AC
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblTabItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   3540
      MouseIcon       =   "TabHead.ctx":0AFE
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblTabItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   2940
      MouseIcon       =   "TabHead.ctx":0C50
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblTabItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   2340
      MouseIcon       =   "TabHead.ctx":0DA2
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblTabItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   1740
      MouseIcon       =   "TabHead.ctx":0EF4
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblTabItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   1080
      MouseIcon       =   "TabHead.ctx":1046
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   180
      Width           =   615
   End
   Begin VB.Line lnLeft 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblTabItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   300
      MouseIcon       =   "TabHead.ctx":1198
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   180
      Width           =   615
   End
   Begin VB.Line lnRight 
      BorderColor     =   &H80000014&
      X1              =   960
      X2              =   3360
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000014&
      Height          =   735
      Left            =   240
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "TabHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private iSelectedTab As Integer
Private iNumLoadedTabs As Integer
Private sTabString As String
Private iFocusTab As String

Event TabSelected(iIndex As Integer, sName As String)
'Default Property Values:
Const m_def_SelectedBold = 0
Const m_def_SelectedColor = &H80000012
Const m_def_InactiveBold = 0
Const m_def_InactiveColor = &H8000000E
'Property Variables:
Dim m_SelectedBold As Boolean
Dim m_SelectedColor As OLE_COLOR
Dim m_InactiveBold As Boolean
Dim m_InactiveColor As OLE_COLOR



'----------------------------------------------------------------------
' Property TabString
'----------------------------------------------------------------------
Public Property Let TabString(sItems As String)
    sTabString = sItems
    NameItems
    DrawTabs
End Property

Public Property Get TabString() As String
    TabString = sTabString
End Property

'----------------------------------------------------------------------
' Property TabSelect
'----------------------------------------------------------------------
Public Property Let SelectedTab(sItem As String)
    iSelectedTab = Val(sItem)
    'NameItems
    DrawTabs
End Property

Public Property Get SelectedTab() As String
    SelectedTab = iSelectedTab
End Property

Private Sub UserControl_EnterFocus()
    'UpdateFocusRect True
End Sub

'----------------------------------------------------------------------
' Property Initialize
'----------------------------------------------------------------------
Private Sub UserControl_InitProperties()
    sTabString = "Tab1,Tab2"
    iSelectedTab = 0
    iFocusTab = 0
    NameItems
    m_SelectedBold = m_def_SelectedBold
    m_SelectedColor = m_def_SelectedColor
    m_InactiveBold = m_def_InactiveBold
    m_InactiveColor = m_def_InactiveColor
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyRight
            'SelectNext
        Case vbKeyLeft
            'SelectPrevious
    End Select
End Sub

Private Sub UserControl_LostFocus()
    'UpdateFocusRect False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    sTabString = PropBag.ReadProperty("TabString", "Tab1,Tab2")
    iSelectedTab = Val(PropBag.ReadProperty("SelectedTab", "0"))
    NameItems
    m_SelectedBold = PropBag.ReadProperty("SelectedBold", m_def_SelectedBold)
    m_SelectedColor = PropBag.ReadProperty("SelectedColor", m_def_SelectedColor)
    m_InactiveBold = PropBag.ReadProperty("InactiveBold", m_def_InactiveBold)
    m_InactiveColor = PropBag.ReadProperty("InactiveColor", m_def_InactiveColor)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000010)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "TabString", sTabString, "Tab1,Tab2"
    PropBag.WriteProperty "SelectedTab", CStr(iSelectedTab), "0"
    Call PropBag.WriteProperty("SelectedBold", m_SelectedBold, m_def_SelectedBold)
    Call PropBag.WriteProperty("SelectedColor", m_SelectedColor, m_def_SelectedColor)
    Call PropBag.WriteProperty("InactiveBold", m_InactiveBold, m_def_InactiveBold)
    Call PropBag.WriteProperty("InactiveColor", m_InactiveColor, m_def_InactiveColor)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000010)
End Sub

'----------------------------------------------------------------------
' Events
'----------------------------------------------------------------------
Private Sub lblTabItem_Click(Index As Integer)
    RaiseEvent TabSelected(Index, lblTabItem(Index).Caption)
    iSelectedTab = Index
    iFocusTab = Index
    DrawTabs
    'UpdateFocusRect True
End Sub

Private Sub lblTabItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 0 To lblTabItem.Count - 1
        If Not i = Index Then
            If lblTabItem(i).FontUnderline = True Then lblTabItem(i).FontUnderline = False
        End If
    Next
    
    If lblTabItem(Index).FontUnderline = False Then lblTabItem(Index).FontUnderline = True
    tmrMouseOver.Enabled = True
    tmrMouseOver.Tag = Index
End Sub

Private Sub tmrMouseOver_Timer()
    If lblTabItem(Val(tmrMouseOver.Tag)).FontUnderline = True Then
        lblTabItem(Val(tmrMouseOver.Tag)).FontUnderline = False
    End If
    tmrMouseOver.Enabled = False
End Sub

Private Sub UserControl_Resize()
    'NameItems
    DrawTabs
End Sub

'----------------------------------------------------------------------
' DrawTabs
'----------------------------------------------------------------------
Private Sub DrawTabs()
    Dim i As Integer
    Dim iPadding As Integer
    Dim iWidthCounter As Integer
    Dim bSelectedMatched As Boolean
    iPadding = Screen.TwipsPerPixelX * 2
    bSelectedMatched = False
    
    'Set initial,left padding
    iWidthCounter = iPadding * 6
    
    For i = 0 To lblTabItem.Count - 1
        If i = iSelectedTab Then
            bSelectedMatched = True
            lblTabItem(i).Left = iWidthCounter '+ ((iPadding / 2) * i)
            lblTabItem(i).Top = Height - lblTabItem(i).Height
            shpBack.Left = lblTabItem(i).Left - iPadding
            shpBack.Width = lblTabItem(i).Width + (2 * iPadding / 2)
            shpBack.Top = IIf(lblTabItem(i).Top - iPadding * 3 < 0, 0, lblTabItem(i).Top - iPadding * 3)
            shpBack.Height = Height + IIf(shpBack.Top < 1, Screen.TwipsPerPixelY, shpBack.Top)
            lblTabItem(i).ForeColor = m_SelectedColor
            lblTabItem(i).FontBold = m_SelectedBold
            iWidthCounter = iWidthCounter + lblTabItem(i).Width + ((iPadding / 2) * 2)
            
        Else
            lblTabItem(i).Left = iWidthCounter '+ ((iPadding / 2) * i)
            lblTabItem(i).Top = Height - lblTabItem(i).Height
            lblTabItem(i).ForeColor = m_InactiveColor
            lblTabItem(i).FontBold = m_InactiveBold
            iWidthCounter = iWidthCounter + lblTabItem(i).Width + ((iPadding / 2) * 2)
        End If
    Next
    
    'Hide TabRectangle if none selected, or Tabstring is empty
    If bSelectedMatched And Not sTabString = "" Then
        shpBack.Visible = True
    Else
        shpBack.Visible = False
    End If
    
    'Draw Borderline
    lnLeft.X1 = 0
    lnLeft.X2 = IIf(shpBack.Visible, shpBack.Left, shpBack.Left + shpBack.Width)
    lnRight.X1 = shpBack.Left + shpBack.Width - Screen.TwipsPerPixelX
    lnRight.X2 = Width
    
    lnTabShadow.X1 = shpBack.Left + shpBack.Width - Screen.TwipsPerPixelX
    lnTabShadow.X2 = shpBack.Left + shpBack.Width - Screen.TwipsPerPixelX
    If lnTabShadow.X1 Mod Screen.TwipsPerPixelX > 5 Then
        lnTabShadow.X1 = lnTabShadow.X1 - Screen.TwipsPerPixelX
        lnTabShadow.X2 = lnTabShadow.X2 - Screen.TwipsPerPixelX
    End If
    lnLeft.Y1 = Height - Screen.TwipsPerPixelY
    lnLeft.Y2 = Height - Screen.TwipsPerPixelY
    lnRight.Y1 = lnLeft.Y1
    lnRight.Y2 = lnLeft.Y2
    lnTabShadow.Y1 = shpBack.Top + Screen.TwipsPerPixelY
    lnTabShadow.Y2 = lnRight.Y1
    lnTabShadow.BorderColor = vbButtonShadow
    
End Sub

'----------------------------------------------------------------------
' NameItems
'----------------------------------------------------------------------
Private Sub NameItems()
    Dim i As Integer
    Dim icnt As Integer
    Dim varr() As String
    
    varr = Split(sTabString, ",")
    icnt = UBound(varr)
    iNumLoadedTabs = icnt + 1
    For i = 0 To lblTabItem.Count - 1
        If i > icnt Then
            lblTabItem(i).Visible = False
        ElseIf varr(i) = "" Then
            lblTabItem(i).Caption = "..."
            lblTabItem(i).Width = 3 * 120
        Else
            lblTabItem(i).Caption = varr(i)
            lblTabItem(i).Width = Len(varr(i)) * 90 + 200
            lblTabItem(i).Visible = True
        End If
    Next
    
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get SelectedBold() As Boolean
    SelectedBold = m_SelectedBold
End Property

Public Property Let SelectedBold(ByVal New_SelectedBold As Boolean)
    m_SelectedBold = New_SelectedBold
    PropertyChanged "SelectedBold"
    DrawTabs
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SelectedColor() As OLE_COLOR
    SelectedColor = m_SelectedColor
End Property

Public Property Let SelectedColor(ByVal New_SelectedColor As OLE_COLOR)
    m_SelectedColor = New_SelectedColor
    PropertyChanged "SelectedColor"
    DrawTabs
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get InactiveBold() As Boolean
    InactiveBold = m_InactiveBold
End Property

Public Property Let InactiveBold(ByVal New_InactiveBold As Boolean)
    m_InactiveBold = New_InactiveBold
    PropertyChanged "InactiveBold"
    DrawTabs
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get InactiveColor() As OLE_COLOR
    InactiveColor = m_InactiveColor
End Property

Public Property Let InactiveColor(ByVal New_InactiveColor As OLE_COLOR)
    m_InactiveColor = New_InactiveColor
    PropertyChanged "InactiveColor"
    DrawTabs
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    DrawTabs
End Property


Private Sub UpdateFocusRect(bShow As Boolean)
    
    If bShow = True Then
        shpFocus.Left = shpBack.Left + Screen.TwipsPerPixelX * 4
        shpFocus.Width = shpBack.Width - Screen.TwipsPerPixelX * 8
        shpFocus.Top = shpBack.Top + Screen.TwipsPerPixelY * 4
    End If
    
    If Not shpFocus.Visible = bShow Then
        shpFocus.Visible = bShow
        Debug.Print "show focus"
    Else
        Debug.Print "already visible"
    End If
End Sub

Private Sub SelectNext()
    If Not iSelectedTab = iNumLoadedTabs - 1 Then
        iSelectedTab = iSelectedTab + 1
        DrawTabs
        'UpdateFocusRect True
        RaiseEvent TabSelected(iSelectedTab, lblTabItem(iSelectedTab).Caption)
    End If
End Sub

Private Sub SelectPrevious()
    If Not iSelectedTab = 0 Then
        iSelectedTab = iSelectedTab - 1
        DrawTabs
        'UpdateFocusRect True
        RaiseEvent TabSelected(iSelectedTab, lblTabItem(iSelectedTab).Caption)
    End If
End Sub

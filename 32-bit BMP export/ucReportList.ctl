VERSION 5.00
Begin VB.UserControl ucReportList 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   208
   Begin VB.ListBox lstReport 
      Height          =   1305
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1770
   End
End
Attribute VB_Name = "ucReportList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User Control:  ucReportList.ctl
' Author:        Carles P.V.
' Dependencies:  None
' Last revision: 2003.12.22
'================================================
'
' - 2003.12.12
'
'   * Added List() prop.
'   * Added <index> param. in AddItem
'   * Fixed Refresh (added listbox refresh)
'
' - 2003.12.15
'
'   * DrawText 'string lenght' param. has been changed from Len(<string>) to -1.
'     Extended char-sets now supported. Thanks to CodeClub.
'
' - 2003.12.17
'
'   * Fixed pvCalcPixelsPerDlgUnit(). Thanks to Vlad Vissoultchev.
'
' - 2003.12.22
'
'   * Fixed 'Property Get List()'

Option Explicit
Option Base 0

'-- API:

Private Const GWL_STYLE        As Long = (-16)
Private Const GWL_EXSTYLE      As Long = (-20)
Private Const WS_THICKFRAME    As Long = &H40000
Private Const WS_BORDER        As Long = &H800000
Private Const WS_EX_WINDOWEDGE As Long = &H100&
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_EX_STATICEDGE As Long = &H20000

Private Const LB_SETTABSTOPS   As Long = &H192

Private Const BDR_RAISEDOUTER  As Long = &H1
Private Const BF_BOTTOM        As Long = &H8
Private Const BF_RIGHT         As Long = &H4

Private Const DT_LEFT          As Long = &H0
Private Const DT_RIGHT         As Long = &H2
Private Const DT_SINGLELINE    As Long = &H20
Private Const DT_VCENTER       As Long = &H4

Private Const COLOR_BTNFACE    As Long = 15

Private Const WM_GETFONT       As Long = &H31

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT2, lpSourceRect As RECT2) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT2, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As Long) As Long
Private Declare Function GetDialogBaseUnits Lib "user32" () As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpString As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'//

'-- Public Enums.:
Public Enum eAlignmentCts
    [LeftJustify] = DT_LEFT
    [RightJustify] = DT_RIGHT
End Enum

'-- Private Types:
Private Type uColumn
    rHeader          As RECT2
    rHeaderCaption   As RECT2
    sHeaderCaption   As String
    eColumnAlignment As eAlignmentCts
End Type

'-- Private Variables:
Private m_uCols()       As uColumn
Private m_lTabs()       As Long
Private m_rHeaders      As RECT2
Private m_rHeadersPad   As RECT2
Private m_lHeaderHeight As Long

'-- Event declarations:
Public Event Click()



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()

  Dim lStyle As Long
    
    '-- Remove list box border
    lStyle = GetWindowLong(lstReport.hWnd, GWL_EXSTYLE)
    lStyle = lStyle And Not WS_EX_CLIENTEDGE
    Call SetWindowLong(lstReport.hWnd, GWL_EXSTYLE, lStyle)
        
    '-- Initialize arrays
    ReDim m_uCols(0)
    ReDim m_lTabs(0)
End Sub

Private Sub UserControl_Resize()

  Dim nIdx As Integer

    '-- Headers height
    m_lHeaderHeight = 1.5 * UserControl.TextHeight(vbNullString)
    
    '-- Update headers height
    For nIdx = 1 To UBound(m_uCols())
        m_uCols(nIdx).rHeader.y2 = m_lHeaderHeight
        m_uCols(nIdx).rHeaderCaption.y2 = m_lHeaderHeight
    Next nIdx
    
    '-- Set headers background and pad rects.
    Call SetRect(m_rHeaders, 0, 0, UserControl.ScaleWidth, m_lHeaderHeight)
    Call SetRect(m_rHeadersPad, m_uCols(UBound(m_uCols)).rHeader.x2, 0, UserControl.ScaleWidth, m_lHeaderHeight)
    
    '-- Resize list box
    Call lstReport.Move(0, m_lHeaderHeight, UserControl.ScaleWidth, UserControl.ScaleHeight - m_lHeaderHeight)
End Sub

Private Sub UserControl_Paint()
    
  Dim nIdx As Integer
    
    '-- Erase background
    Call FillRect(UserControl.hDC, m_rHeaders, GetSysColorBrush(COLOR_BTNFACE))
    
    '-- Draw headers
    For nIdx = 1 To UBound(m_uCols)
        With m_uCols(nIdx)
            Call DrawEdge(UserControl.hDC, .rHeader, BDR_RAISEDOUTER, BF_BOTTOM Or BF_RIGHT)
            Call DrawText(UserControl.hDC, .sHeaderCaption, -1, .rHeaderCaption, .eColumnAlignment + DT_SINGLELINE + DT_VCENTER)
        End With
    Next nIdx
    
    '-- Draw header's pad
    Call DrawEdge(UserControl.hDC, m_rHeadersPad, BDR_RAISEDOUTER, BF_BOTTOM Or BF_RIGHT)
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Sub AddHeader(ByVal Width As Long, _
                     Optional ByVal Alignment As eAlignmentCts = [LeftJustify], _
                     Optional ByVal Caption As String = vbNullString)
    
  Dim nIdx               As Integer
  Dim nCols              As Integer
  Dim snPixelsPerDlgUnit As Single
    
    '-- Current number of columns
    nCols = UBound(m_uCols)
    
    '-- Increase count
    ReDim Preserve m_uCols(0 To nCols + 1)
    ReDim Preserve m_lTabs(0 To nCols + 1)
    
    '-- Define header/column
    With m_uCols(nCols + 1)
        Call SetRect(.rHeader, m_uCols(nCols).rHeader.x2, 0, m_uCols(nCols).rHeader.x2 + Width, m_lHeaderHeight)
        Call CopyRect(.rHeaderCaption, .rHeader)
        Call InflateRect(.rHeaderCaption, -2, 0)
        Let .eColumnAlignment = Alignment
        Let .sHeaderCaption = Caption
    End With
    
    '-- Get pixels per dialog unit coeff. (tabs)
    snPixelsPerDlgUnit = pvCalcPixelsPerDlgUnit(lstReport.hWnd)
    
    '-- Readjust tabs
    For nIdx = 1 To UBound(m_uCols)
        With m_uCols(nIdx)
            Select Case .eColumnAlignment
                Case LeftJustify:  m_lTabs(nIdx) = m_uCols(nIdx).rHeaderCaption.x1 / snPixelsPerDlgUnit
                Case RightJustify: m_lTabs(nIdx) = -(m_uCols(nIdx).rHeaderCaption.x2 + 1) / snPixelsPerDlgUnit
            End Select
        End With
    Next nIdx
    Call SendMessage(lstReport.hWnd, LB_SETTABSTOPS, UBound(m_uCols), m_lTabs(1))

    '-- Refresh
    Call UserControl_Paint
End Sub

Public Sub ClearHeaders()
    
    '-- Clear arrays
    ReDim m_uCols(0)
    ReDim m_lTabs(0)
    
    '-- Refresh
    Call UserControl_Paint
End Sub

Public Sub AddItem(ByVal Item As String, _
                   Optional ByVal Index As Variant _
                   )
    
    '-- Add/Insert item
    If (IsMissing(Index)) Then
        Call lstReport.AddItem(Item, lstReport.ListCount)
      Else
        Call lstReport.AddItem(Item, Index)
    End If
End Sub

Public Sub RemoveItem(ByVal Index As Integer)
    Call lstReport.RemoveItem(Index)
End Sub

Public Sub ClearList()
    Call lstReport.Clear
End Sub

Public Sub Refresh()
    Call UserControl_Paint
    Call lstReport.Refresh
End Sub



'========================================================================================
' Events
'========================================================================================

Private Sub lstReport_Click()
    RaiseEvent Click
End Sub



'========================================================================================
' Properties
'========================================================================================

Public Property Get FontHeader() As StdFont
    Set FontHeader = UserControl.Font
End Property
Public Property Set FontHeader(ByVal New_FontHeader As StdFont)
    Set UserControl.Font = New_FontHeader ' Use Refresh method
    Call UserControl_Resize
End Property

Public Property Get Font() As StdFont
    Set Font = lstReport.Font
End Property
Public Property Set Font(ByVal New_Font As StdFont)
    Set lstReport.Font = New_Font
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = lstReport.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lstReport.BackColor() = New_BackColor
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lstReport.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lstReport.ForeColor() = New_ForeColor
End Property

Public Property Get HeaderCount() As Integer
    HeaderCount = UBound(m_uCols)
End Property

Public Property Get List(ByVal Index As Integer) As String
    List = lstReport.List(Index)
End Property
Public Property Let List(ByVal Index As Integer, ByVal New_Data As String)
    lstReport.List(Index) = New_Data
End Property

Public Property Get ListCount() As Integer
    ListCount = lstReport.ListCount
End Property

Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = lstReport.ListIndex
End Property
Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    lstReport.ListIndex() = New_ListIndex
End Property

Public Property Get Enabled() As Boolean
    Enabled = lstReport.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    lstReport.Enabled() = New_Enabled
End Property

'*

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get hListWnd() As Long
    hListWnd = lstReport.hWnd
End Property

'//

Private Sub UserControl_InitProperties()

    Set UserControl.Font = Ambient.Font
    Set lstReport.Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set UserControl.Font = PropBag.ReadProperty("FontHeader", Ambient.Font)
    Set lstReport.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lstReport.BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    lstReport.ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    lstReport.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("FontHeader", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Font", lstReport.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackColor", lstReport.BackColor, vbWindowBackground)
    Call PropBag.WriteProperty("ForeColor", lstReport.ForeColor, vbWindowText)
    Call PropBag.WriteProperty("Enabled", lstReport.Enabled, True)
End Sub



'========================================================================================
' Private
'========================================================================================

Private Function pvCalcPixelsPerDlgUnit(hWndLB As Long) As Single
' Returns the number of pixels-per-dialog
' unit for the given font.
'
' Provided to VBnet by Brad Martinez
' Thanks to Vlad Vissoultchev for the simplification

  Dim hFont      As Long
  Dim hFontOld   As Long
  Dim lhDC       As Long
  Dim sz         As Size
  Dim cxAvLBChar As Long ' average LB char width, in pixels
  Const sChars   As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"

    '-- Get the device contect of the listbox
    lhDC = GetDC(hWndLB)
  
    If (lhDC) Then
   
        '-- Select hWndLB's HFONT into its DC (VB
        '   does not select a control's Font into its DC)
        hFont = SendMessage(hWndLB, WM_GETFONT, 0, ByVal 0&)
        hFontOld = SelectObject(lhDC, hFont)
    
        If (GetTextExtentPoint32(lhDC, sChars, Len(sChars), sz)) Then
        
            '-- Get the list box average char width
            '   and the system's horizontal dialog
            '   base units
            cxAvLBChar = sz.cx / Len(sChars)
        
            '-- Calculate and return the number of
            '   pixels per dialog unit for the list
            pvCalcPixelsPerDlgUnit = cxAvLBChar / 4
        End If
    
        Call SelectObject(lhDC, hFontOld)
        Call ReleaseDC(hWndLB, lhDC)
    End If
End Function

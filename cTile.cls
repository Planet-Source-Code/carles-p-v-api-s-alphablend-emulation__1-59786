VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Class:         cTile.cls
' Author:        Carles P.V.
' Dependencies:
' Last revision: 2003.03.28
'================================================

Option Explicit

'-- API:

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const OBJ_BITMAP     As Long = 7

Private Const HS_HORIZONTAL  As Long = 0
Private Const HS_VERTICAL    As Long = 1
Private Const HS_FDIAGONAL   As Long = 2
Private Const HS_BDIAGONAL   As Long = 3
Private Const HS_CROSS       As Long = 4
Private Const HS_DIAGCROSS   As Long = 5

Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long
Private Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'//

'-- Public Enums.:
Public Enum HatchBrushStyleCts
    [brHorizontal] = 0
    [brVertival]
    [brDownwardDiagonal]
    [brUpwardDiagonal]
    [brCross]
    [brDiagonalCross]
End Enum

'-- Private Variables:
Private m_hBrush As Long ' Pattern brush



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Initialize()
    m_hBrush = 0
End Sub

Private Sub Class_Terminate()
    Call Me.DestroyPattern
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Function CreatePatternFromStdPicture(Image As StdPicture) As Boolean

  Dim uBI       As BITMAP
  Dim uBIH      As BITMAPINFOHEADER
  Dim aBuffer() As Byte ' Packed DIB
    
  Dim lhDC      As Long
  Dim lhOldBmp  As Long
    
    If (GetObjectType(Image.handle) = OBJ_BITMAP) Then
    
        '-- Get image info
        Call GetObject(Image.handle, Len(uBI), uBI)
        
        '-- Prepare DIB header and redim. buffer array
        With uBIH
            .biSize = Len(uBIH)
            .biPlanes = 1
            .biBitCount = 24
            .biWidth = uBI.bmWidth
            .biHeight = uBI.bmHeight
            .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
        End With
        ReDim aBuffer(1 To Len(uBIH) + uBIH.biSizeImage)
            
        '-- Create DIB brush
        lhDC = CreateCompatibleDC(0)
        If (lhDC <> 0) Then
            lhOldBmp = SelectObject(lhDC, Image.handle)
                    
            '-- Build packed DIB:
            '-  Merge Header
            Call CopyMemory(aBuffer(1), uBIH, Len(uBIH))
            '-  Get and merge DIB bits
            Call GetDIBits(lhDC, Image.handle, 0, uBI.bmHeight, aBuffer(Len(uBIH) + 1), uBIH, DIB_RGB_COLORS)
            
            Call SelectObject(lhDC, lhOldBmp)
            Call DeleteDC(lhDC)
            
            '-  Create brush from packed DIB
            Call Me.DestroyPattern
            m_hBrush = CreateDIBPatternBrushPt(aBuffer(1), DIB_RGB_COLORS)
        End If
    End If
    
    '-- Success
    CreatePatternFromStdPicture = (m_hBrush <> 0)
End Function

Public Function CreatePatternFromHatchBrush(ByVal BrushStyle As HatchBrushStyleCts, ByVal Color As OLE_COLOR) As Boolean

    '-- Create brush from system brush
    Call Me.DestroyPattern
    Call OleTranslateColor(Color, 0, Color)
    m_hBrush = CreateHatchBrush(BrushStyle, Color)
    
    '-- Success
    CreatePatternFromHatchBrush = (m_hBrush <> 0)
End Function

Public Function CreatePatternFromSolidColor(ByVal Color As OLE_COLOR) As Boolean

    '-- Create brush from solid color
    Call Me.DestroyPattern
    Call OleTranslateColor(Color, 0, Color)
    m_hBrush = CreateSolidBrush(Color)
    
    '-- Success
    CreatePatternFromSolidColor = (m_hBrush <> 0)
End Function

Public Sub Tile(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal ResetBrushOrigin As Boolean = True)

  Dim rTile As RECT2
  Dim ptOrg As POINTAPI
  
    If (m_hBrush <> 0) Then
        '-- Set brush origin
        If (ResetBrushOrigin) Then
            Call SetBrushOrgEx(hDC, x, y, ptOrg)
          Else
            Call SetBrushOrgEx(hDC, 0, 0, ptOrg)
        End If
        '-- Tile image
        Call SetRect(rTile, x, y, x + Width, y + Height)
        Call FillRect(hDC, rTile, m_hBrush)
    End If
End Sub

Public Sub DestroyPattern()
    If (m_hBrush <> 0) Then
        Call DeleteObject(m_hBrush)
        m_hBrush = 0
    End If
End Sub

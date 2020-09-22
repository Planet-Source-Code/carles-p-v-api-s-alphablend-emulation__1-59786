Attribute VB_Name = "mAlphaBlt"
'================================================
' Module:        mAlphaBlt.bas
' Author:        Carles P.V.
' Dependencies:
' Last revision: 2005.05.03
'
' 2005.04.01: First release
' 2005.05.03: Speed up: checked special alpha values
'             (full opaque and full transparent)
'================================================

Option Explicit

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

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

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds     As SAFEARRAYBOUND
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const OBJ_BITMAP     As Long = 7

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As Any, ByVal un As Long, lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long



Public Function AlphaBlt( _
                ByVal hDC As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal BackColor As OLE_COLOR, _
                ByVal hBitmap As Long _
                ) As Long
  
  Dim uBI       As BITMAP
  Dim uBIH      As BITMAPINFOHEADER
  
  Dim R         As Long
  Dim G         As Long
  Dim B         As Long
  Dim a1        As Long
  Dim a2        As Long
  
  Dim uSA       As SAFEARRAY1D
  Dim aBits()   As Byte
  
  Dim i         As Long
  Dim iIn       As Long
    
    '-- Check type (bitmap)
    If (GetObjectType(hBitmap) = OBJ_BITMAP) Then
        
        '-- Get bitmap info
        If (GetObject(hBitmap, Len(uBI), uBI)) Then
            
            '-- Check if source bitmap is 32-bit!
            If (uBI.bmBitsPixel = 32) Then
            
                With uBIH
                
                    '-- Define DIB info
                    .biSize = Len(uBIH)
                    .biPlanes = 1
                    .biBitCount = 32
                    .biWidth = uBI.bmWidth
                    .biHeight = uBI.bmHeight
                    .biSizeImage = (4 * .biWidth) * .biHeight
                        
                    '-- Get source (image) color data
                    ReDim aBits(.biSizeImage - 1)
                    Call CopyMemory(aBits(0), ByVal uBI.bmBits, .biSizeImage)
                    
                    '-- Translate OLE color
                    Call OleTranslateColor(BackColor, 0, BackColor)
                    R = (BackColor And &HFF&)
                    G = (BackColor And &HFF00&) \ &H100
                    B = (BackColor And &HFF0000) \ &H10000
                    
                    '-- Blend with BackColor
                    For i = 3 To .biSizeImage - 1 Step 4
                        a1 = aBits(i)
                        If (a1 = &H0) Then
                            '-- Dest. = Source (solid background)
                            aBits(i - 1) = R
                            aBits(i - 2) = G
                            aBits(i - 3) = B
                        ElseIf (a1 = &HFF) Then
                            '-- Do nothing
                        Else
                            '-- Blend
                            a2 = &HFF - a1
                            iIn = i - 1
                            aBits(iIn) = (a1 * aBits(iIn) + a2 * R) \ &HFF: iIn = iIn - 1
                            aBits(iIn) = (a1 * aBits(iIn) + a2 * G) \ &HFF: iIn = iIn - 1
                            aBits(iIn) = (a1 * aBits(iIn) + a2 * B) \ &HFF
                        End If
                    Next i
                    
                    '-- Paint alpha-blended
                    AlphaBlt = StretchDIBits(hDC, x, y, .biWidth, .biHeight, 0, 0, .biWidth, .biHeight, aBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
                End With
            End If
        End If
    End If
End Function

Public Function AlphaBlend( _
                ByVal hDC As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal hBitmap As Long _
                ) As Long
  
  Dim uBI       As BITMAP
  Dim uBIH      As BITMAPINFOHEADER
  
  Dim lhDC      As Long
  Dim lhDIB     As Long
  Dim lhDIBOld  As Long
  
  Dim a1        As Long
  Dim a2        As Long
  
  Dim uSSA      As SAFEARRAY1D
  Dim aSBits()  As Byte
  Dim uDSA      As SAFEARRAY1D
  Dim aDBits()  As Byte
  Dim lpData    As Long
  
  Dim i         As Long
  Dim iIn       As Long
    
    '-- Check type (bitmap)
    If (GetObjectType(hBitmap) = OBJ_BITMAP) Then
        
        '-- Get bitmap info
        If (GetObject(hBitmap, Len(uBI), uBI)) Then
        
            '-- Check if source bitmap is 32-bit!
            If (uBI.bmBitsPixel = 32) Then
            
                With uBIH
                
                    '-- Define DIB info
                    .biSize = Len(uBIH)
                    .biPlanes = 1
                    .biBitCount = 32
                    .biWidth = uBI.bmWidth
                    .biHeight = uBI.bmHeight
                    .biSizeImage = (4 * .biWidth) * .biHeight
                    
                    '-- Create a temporary DIB section, select into a DC, and
                    '   bitblt destination DC area
                    lhDC = CreateCompatibleDC(0)
                    lhDIB = CreateDIBSection(lhDC, uBIH, DIB_RGB_COLORS, lpData, 0, 0)
                    lhDIBOld = SelectObject(lhDC, lhDIB)
                    Call BitBlt(lhDC, 0, 0, uBI.bmWidth, uBI.bmHeight, hDC, x, y, vbSrcCopy)
                    
                    '-- Map destination color data
                    Call pvMapDIBits(uDSA, aDBits(), lpData, .biSizeImage)
                    
                    '-- Map source color data
                    Call pvMapDIBits(uSSA, aSBits(), uBI.bmBits, .biSizeImage)
                    
                    '-- Blend with destination
                    For i = 3 To .biSizeImage - 1 Step 4
                        a1 = aSBits(i)
                        If (a1 = &H0) Then
                            '-- Do nothing (dest. preserved)
                        ElseIf (a1 = &HFF) Then
                            '-- Dest. = Source
                            iIn = i - 1
                            aDBits(iIn) = aSBits(iIn): iIn = iIn - 1
                            aDBits(iIn) = aSBits(iIn): iIn = iIn - 1
                            aDBits(iIn) = aSBits(iIn)
                        Else
                            '-- Blend
                            a2 = &HFF - a1
                            iIn = i - 1
                            aDBits(iIn) = (a1 * aSBits(iIn) + a2 * aDBits(iIn)) \ &HFF: iIn = iIn - 1
                            aDBits(iIn) = (a1 * aSBits(iIn) + a2 * aDBits(iIn)) \ &HFF: iIn = iIn - 1
                            aDBits(iIn) = (a1 * aSBits(iIn) + a2 * aDBits(iIn)) \ &HFF
                        End If
                    Next i
                    
                    '-- Paint alpha-blended
                    AlphaBlend = StretchDIBits(hDC, x, y, .biWidth, .biHeight, 0, 0, .biWidth, .biHeight, ByVal lpData, uBIH, DIB_RGB_COLORS, vbSrcCopy)
                End With
                
                '-- Unmap
                Call pvUnmapDIBits(aDBits())
                Call pvUnmapDIBits(aSBits())
                
                '-- Clean up
                Call SelectObject(lhDC, lhDIBOld)
                Call DeleteObject(lhDIB)
                Call DeleteDC(lhDC)
            End If
        End If
    End If
End Function

Private Sub pvMapDIBits(uSA As SAFEARRAY1D, aBits() As Byte, ByVal lpData As Long, ByVal lSize As Long)
    
    With uSA
        .cbElements = 1
        .cDims = 1
        .Bounds.lLbound = 0
        .Bounds.cElements = lSize
        .pvData = lpData
    End With
    Call CopyMemory(ByVal VarPtrArray(aBits()), VarPtr(uSA), 4)
End Sub

Private Sub pvUnmapDIBits(aBits() As Byte)

    Call CopyMemory(ByVal VarPtrArray(aBits()), 0&, 4)
End Sub

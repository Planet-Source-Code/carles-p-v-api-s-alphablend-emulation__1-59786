VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Class:         cDIB32.cls (cDIB 32-bit version)
' Author:        Carles P.V.
' Dependencies:
' Last revision: 2004.9.12 (# 2005.04.02)
'================================================
'
' - 2004.08.18
'
'   * Fixed CreateFromBitmapFile() function:
'     m_uBIH structure was not initialized (filled) -> Save() function failed
'     Thanks to Paul Caton
'
'     Now, private variables 'm_Width, m_Height, m_BPP, m_BytesPerScanline and m_Size'
'     have been removed. Directly taken from m_uBIH structure.
'
'   * Also added CreateFromResourceBitmap() variation (by Paul Caton)
'
' - 2004.08.19
'
'   * Fixed Image() Property:
'     Going back to original routine. Probably problems for using a still selected DIB (?).
'     Anyway, old function worked well when setting/extracting Image to/from VB-Clipboard
'     using vbCFDIB format (but not using vbCFBitmap one) (?).
'     Also, curiously, Image was returned as 'down-top DIB'.
'
'   * Improved Save() function: Removed GlobalLock/GlobalUnlock. Redundant when allocating
'     fixed memory. The GlobalAlloc returns the memory pointer.
'
'   Thanks again, Paul.
'
' - 2004.09.12
'
'   * Fixed CreateFromStdPicture() function.
'   * Added StretchBltMode param. in Stretch() and Paint() functions.
'
'========================================================================================
'
' # 2005.04.02
'
'   * Current cDIB32:
'     - Simplified version: only accepting 32-bit bitmaps
'     - New CreateFromResourceBitmap() (uses LoadLibraryEx()). Thanks to redbird77.

Option Explicit

'-- API:

Private Type BITMAPFILEHEADER
    bfType      As Integer
    bfSize      As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits   As Long
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

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PICTDESC
    Size     As Long
    Type     As Long
    hBmp     As Long
    hPal     As Long
    Reserved As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const DIB_RGB_COLORS      As Long = 0
Private Const COLORONCOLOR        As Long = 3
Private Const HALFTONE            As Long = 4
Private Const OBJ_BITMAP          As Long = 7
Private Const LR_LOADFROMFILE     As Long = &H10
Private Const IMAGE_BITMAP        As Long = 0
Private Const LR_CREATEDIBSECTION As Long = &H2000

Private Declare Function CreateDIBSection32 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long

Private Declare Function FindResourceStr Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As Long) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long

Private Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2
Private Const RT_BITMAP                As Long = 2

'//

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)

'//

Private Const CF_DIB As Long = 8

Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

'//

Private Const GMEM_FIXED    As Long = &H0
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_DDESHARE As Long = &H2000
Private Const GMEM_ZEROINIT As Long = &H40

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long

'//

Private Const INVALID_HANDLE_VALUE  As Long = -1
Private Const CREATE_ALWAYS         As Long = 2
Private Const GENERIC_READ          As Long = &H80000000
Private Const GENERIC_WRITE         As Long = &H40000000
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'//

'-- Public Enums.:
Public Enum eBPPCts
    [01_bpp] = 1
    [04_bpp] = 4
    [08_bpp] = 8
    [16_bpp] = 16
    [24_bpp] = 24
    [32_bpp] = 32
End Enum

Public Enum eStretchBltModeCts
    [sbmColorOnColor] = COLORONCOLOR
    [sbmHalftone] = HALFTONE
End Enum

'-- Private Variables:
Private m_uBIH    As BITMAPINFOHEADER
Private m_hDC     As Long
Private m_hDIB    As Long
Private m_hOldDIB As Long
Private m_lpBits  As Long



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Terminate()

    '-- Destroy current DIB
    Call Me.Destroy
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Function Create(ByVal Width As Long, _
                       ByVal Height As Long _
                       ) As Long
    
    '-- Destroy previous
    Call Me.Destroy
    
    '-- Define DIB header
    With m_uBIH
        .biSize = Len(m_uBIH)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
        .biSizeImage = (4 * .biWidth) * .biHeight
    End With
    
    '-- Create DIB and select into a DC
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        m_hDIB = CreateDIBSection32(m_hDC, m_uBIH, DIB_RGB_COLORS, m_lpBits, 0, 0)
        If (m_hDIB <> 0) Then
            m_hOldDIB = SelectObject(m_hDC, m_hDIB)
          Else
            Call Me.Destroy
        End If
    End If
    
    '-- Success
    Create = m_hDIB
End Function

Public Function CreateFromBitmapFile(ByVal Filename As String _
                                    ) As Long
  
  Dim uBI As BITMAP
    
    '-- File exists [?]
    If (FileLen(Filename)) Then
    
        '-- Destroy previous
        Call Me.Destroy

        '-- Create DIB and select into a DC
        m_hDC = CreateCompatibleDC(0)
        If (m_hDC <> 0) Then
            m_hDIB = LoadImage(App.hInstance, Filename, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
            
            '-- Is there a DIB handle
            If (m_hDIB <> 0) Then
                Call GetObject(m_hDIB, Len(uBI), uBI)
                
                '-- 32-bit?
                If (uBI.bmBitsPixel = 32) Then
                    m_hOldDIB = SelectObject(m_hDC, m_hDIB)
                  Else
                    Call Me.Destroy
                End If
              
              Else
                Call Me.Destroy
            End If
        End If
        
        '-- Get DIB props.:
        If (m_hDIB <> 0) Then
            '-- Get props.:
            With m_uBIH
                .biSize = Len(m_uBIH)
                .biPlanes = 1
                .biBitCount = uBI.bmBitsPixel
                .biWidth = uBI.bmWidth
                .biHeight = uBI.bmHeight
                .biSizeImage = (4 * .biWidth) * .biHeight
            End With
            '-- Bits pointer
            m_lpBits = uBI.bmBits
            
            '-- Success
            CreateFromBitmapFile = m_hDIB
        End If
    End If
End Function

Public Function CreateFromResourceBitmap(ByVal Filename As String, _
                                         ByVal ResID As Variant _
                                         ) As Long

  Dim hInstance As Long
  Dim hInfo     As Long
  Dim hData     As Long
  Dim lSize     As Long

  Dim uBIH      As BITMAPINFOHEADER
  Dim lpResHDR  As Long
  Dim lpResBMP  As Long
  
    '-- File exists [?]
    If (FileLen(Filename)) Then

        '-- Get handle to the mapped executable module
        hInstance = LoadLibraryEx(Filename, 0, LOAD_LIBRARY_AS_DATAFILE)
        If (hInstance) Then
            
            '-- Get resource info handle
            hInfo = FindResourceStr(hInstance, IIf(IsNumeric(ResID), "#", vbNullString) & ResID, RT_BITMAP)
            If (hInfo) Then
            
                '-- Get handle to DIB data
                hData = LoadResource(hInstance, hInfo)
                If (hData) Then
                
                    '-- Get size of DIB data
                    lSize = SizeofResource(hInstance, hInfo)
        
                    '-- Get pointer to first byte of DIB data (header)
                    lpResHDR = LockResource(hData)
                    
                    '-- Extract DIB info header
                    Call CopyMemory(uBIH, ByVal lpResHDR, Len(uBIH))
                    
                    '-- 32-bit?
                    If (uBIH.biBitCount = 32) Then
                    
                        '-- Create DIB / fill data
                        If (Me.Create(uBIH.biWidth, uBIH.biHeight)) Then
                            
                            lpResBMP = lpResHDR + Len(m_uBIH)
                            With m_uBIH
                                Call CopyMemory(ByVal m_lpBits, ByVal lpResBMP, .biSizeImage)
                            End With
                            
                            '-- Success
                            CreateFromResourceBitmap = m_hDIB
                        End If
                    End If
                End If
            End If
            Call FreeLibrary(hInstance)
        End If
    End If
End Function

Public Function CreateFromClipboard() As Long
    
  Dim hClipMem  As Long
  Dim lpClipHDR As Long
  Dim lpClipBMP As Long
  Dim uBIH      As BITMAPINFOHEADER
  
    If (OpenClipboard(0)) Then
        
        '-- Get clipboard mem. handle
        hClipMem = GetClipboardData(CF_DIB)
        If (hClipMem <> 0) Then
            
            '-- Get pointer / access data
            lpClipHDR = GlobalLock(hClipMem)
            If (lpClipHDR <> 0) Then
                
                '-- Extract DIB info header
                Call CopyMemory(uBIH, ByVal lpClipHDR, Len(uBIH))
                
                '-- 32-bit?
                If (uBIH.biBitCount = 32) Then

                    '-- Create DIB / fill data
                    If (Me.Create(uBIH.biWidth, uBIH.biHeight)) Then
                        
                        lpClipBMP = lpClipHDR + Len(m_uBIH)
                        With m_uBIH
                            Call CopyMemory(ByVal m_lpBits, ByVal lpClipBMP, .biSizeImage)
                        End With
                            
                        '-- Success
                        CreateFromClipboard = m_hDIB
                    End If
                End If
                Call GlobalUnlock(hClipMem)
            End If
        End If
        Call CloseClipboard
    End If
End Function

Public Function CopyToClipboard() As Long

  Dim hMem      As Long
  Dim lpClipHDR As Long
  Dim lpClipBMP As Long
  Dim lSize     As Long
    
    If (m_hDIB <> 0) Then
    
        If (OpenClipboard(0)) Then
            
            '-- Size of packed DIB
            lSize = Len(m_uBIH) + m_uBIH.biSizeImage
            
            '-- Get mem. handle and its pointer
            hMem = GlobalAlloc(GMEM_MOVEABLE, lSize)
            If (hMem <> 0) Then
                lpClipHDR = GlobalLock(hMem)
                If (lpClipHDR <> 0) Then
                
                    lpClipBMP = lpClipHDR + Len(m_uBIH)
                    
                    '-- Copy data
                    Call CopyMemory(ByVal lpClipHDR, m_uBIH, Len(m_uBIH))
                    Call CopyMemory(ByVal lpClipBMP, ByVal lpBits, m_uBIH.biSizeImage)
                    Call GlobalUnlock(hMem)
                    
                    '-- Paste to Clipboard
                    Call EmptyClipboard
                    CopyToClipboard = SetClipboardData(CF_DIB, hMem)
                    Call CloseClipboard
                End If
            End If
        End If
    End If
End Function

Public Function CloneTo(oDIB32 As cDIB32 _
                       ) As Long
    
    If (m_hDIB <> 0) Then
    
        '-- Create dest. DIB
        If (oDIB32.Create(m_uBIH.biWidth, m_uBIH.biHeight)) Then
            
            '-- Copy color data
            Call CopyMemory(ByVal oDIB32.lpBits, ByVal m_lpBits, m_uBIH.biSizeImage)
            
            '-- Success
            CloneTo = oDIB32.hDIB
        End If
    End If
End Function

Public Function Save(ByVal Filename As String _
                    ) As Boolean

  Const BITMAPTYPE  As Integer = &H4D42
  
  Dim hFile         As Long
  Dim lRet          As Long
  Dim lBytesWritten As Long

  Dim hMem          As Long
  Dim lPtr          As Long
  Dim uBFH          As BITMAPFILEHEADER
  Dim aPal()        As Byte
  Dim lColors       As Long
  
    If (m_hDIB <> 0) Then
    
        '-- Get file handler
        hFile = CreateFile(Filename, GENERIC_READ Or GENERIC_WRITE, ByVal 0, ByVal 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
        
        If (hFile <> INVALID_HANDLE_VALUE) Then
            
            '-- Prepare file header
            With uBFH
                .bfType = BITMAPTYPE
                .bfSize = Len(m_uBIH) + m_uBIH.biSizeImage
                .bfOffBits = Len(uBFH) + Len(m_uBIH)
            End With
            
            '-- Not long alignment...
            hMem = GlobalAlloc(GMEM_FIXED + GMEM_ZEROINIT, Len(uBFH))
            lPtr = GlobalLock(hMem)
            Call CopyMemory(ByVal lPtr, uBFH.bfType, 2)
            Call CopyMemory(ByVal lPtr + 2, uBFH.bfSize, 4)
            Call CopyMemory(ByVal lPtr + 6, 0, 4)
            Call CopyMemory(ByVal lPtr + 10, uBFH.bfOffBits, 4)
            
            '-- Write file header
            lRet = WriteFile(hFile, ByVal lPtr, Len(uBFH), lBytesWritten, ByVal 0)
            Call GlobalUnlock(hMem)
            Call GlobalFree(hMem)
            
            If (lRet <> 0) Then
                
                '-- Write bitmap info header
                lRet = WriteFile(hFile, m_uBIH, Len(m_uBIH), lBytesWritten, ByVal 0)
                '-- Write bitmap color data
                If (lRet <> 0) Then
                    lRet = WriteFile(hFile, ByVal m_lpBits, m_uBIH.biSizeImage, lBytesWritten, ByVal 0)
                End If
            End If
            Call CloseHandle(hFile)
            
            '-- Success
            Save = (lRet <> 0)
        End If
    End If
End Function

Public Sub Destroy()

    '-- Destroy DIB
    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            Call SelectObject(m_hDC, m_hOldDIB)
            Call DeleteObject(m_hDIB)
        End If
        Call DeleteDC(m_hDC)
    End If
    
    '-- Reset BIH structure
    Call ZeroMemory(m_uBIH, Len(m_uBIH))
    
    '-- Reset DIB vars.
    m_hDC = 0
    m_hDIB = 0
    m_hOldDIB = 0
    m_lpBits = 0
End Sub

'//

Public Function Paint(ByVal hDC As Long, _
                      Optional ByVal x As Long = 0, _
                      Optional ByVal y As Long = 0, _
                      Optional ByVal ROP As RasterOpConstants = vbSrcCopy, _
                      Optional ByVal StretchBltMode As eStretchBltModeCts = [sbmColorOnColor] _
                      ) As Long
    
    Paint = Me.Stretch(hDC, x, y, m_uBIH.biWidth, m_uBIH.biHeight, , , , , ROP, StretchBltMode)
End Function

Public Function Stretch(ByVal hDC As Long, _
                        ByVal x As Long, _
                        ByVal y As Long, _
                        ByVal nWidth As Long, _
                        ByVal nHeight As Long, _
                        Optional ByVal xSrc As Long, _
                        Optional ByVal ySrc As Long, _
                        Optional ByVal nSrcWidth As Long, _
                        Optional ByVal nSrcHeight As Long, _
                        Optional ByVal ROP As RasterOpConstants = vbSrcCopy, _
                        Optional ByVal StretchBltMode As eStretchBltModeCts = [sbmColorOnColor] _
                        ) As Long
  
  Dim lOldMode As Long
    
    If (m_hDIB <> 0) Then
        
        If (nSrcWidth = 0) Then nSrcWidth = m_uBIH.biWidth
        If (nSrcHeight = 0) Then nSrcHeight = m_uBIH.biHeight
        
        lOldMode = SetStretchBltMode(hDC, StretchBltMode)
        Stretch = StretchBlt(hDC, x, y, nWidth, nHeight, m_hDC, xSrc, ySrc, nSrcWidth, nSrcHeight, ROP)
        Call SetStretchBltMode(hDC, lOldMode)
    End If
End Function

Public Sub Cls(Optional ByVal Color As OLE_COLOR = vbBlack)
  
  Dim uRect  As RECT2
  Dim hBrush As Long
   
    If (m_hDIB <> 0) Then
        Call SetRect(uRect, 0, 0, m_uBIH.biWidth, m_uBIH.biHeight)
        Call OleTranslateColor(Color, 0, Color)
        hBrush = CreateSolidBrush(Color)
        Call FillRect(m_hDC, uRect, hBrush)
        Call DeleteObject(hBrush)
    End If
End Sub

Public Sub Reset()
    
    If (m_hDIB <> 0) Then
        Call ZeroMemory(ByVal m_lpBits, m_uBIH.biSizeImage)
    End If
End Sub

'//

Public Sub GetBestFitInfo(ByVal SrcW As Long, _
                          ByVal SrcH As Long, _
                          ByVal DstW As Long, _
                          ByVal DstH As Long, _
                          bfx As Long, _
                          bfy As Long, _
                          bfW As Long, _
                          bfH As Long, _
                          Optional ByVal StretchFit As Boolean = False)
                          
  Dim cW As Single
  Dim cH As Single
    
    If ((SrcW > DstW Or SrcH > DstH) Or StretchFit) Then
        cW = DstW / SrcW
        cH = DstH / SrcH
        If (cW < cH) Then
            bfW = DstW
            bfH = SrcH * cW
          Else
            bfH = DstH
            bfW = SrcW * cH
        End If
      Else
        bfW = SrcW
        bfH = SrcH
    End If
    If (bfW < 1) Then bfW = 1
    If (bfH < 1) Then bfH = 1
    bfx = (DstW - bfW) \ 2
    bfy = (DstH - bfH) \ 2
End Sub



'========================================================================================
' Properties
'========================================================================================

Public Property Get hDC() As Long
    hDC = m_hDC
End Property

Public Property Get hDIB() As Long
    hDIB = m_hDIB
End Property

Public Property Get lpBits() As Long
    lpBits = m_lpBits
End Property

Public Property Get Width() As Long
    Width = m_uBIH.biWidth
End Property

Public Property Get Height() As Long
    Height = m_uBIH.biHeight
End Property

Public Property Get BytesPerScanline() As Long
    BytesPerScanline = 4 * m_uBIH.biWidth
End Property

Public Property Get Size() As Long
    Size = m_uBIH.biSizeImage
End Property

Public Property Get Image() As StdPicture
    
  Dim lhSrcDC       As Long
  Dim lhDC          As Long
  Dim lhBmp         As Long
  Dim lhOldBmp      As Long
  Dim lpPictDesc    As PICTDESC
  Dim IID_IDispatch As GUID

    If (m_hDIB <> 0) Then
    
        '-- Create screen compatible bitmap
        lhSrcDC = GetDC(0)
        lhDC = CreateCompatibleDC(lhSrcDC)
        lhBmp = CreateCompatibleBitmap(lhSrcDC, m_uBIH.biWidth, m_uBIH.biHeight)
        lhOldBmp = SelectObject(lhDC, lhBmp)
        Call ReleaseDC(0, lhSrcDC)
        
        '-- Paint from DIB
        Call Me.Paint(lhDC)
        '-- Clean up
        Call SelectObject(lhDC, lhOldBmp)
        Call DeleteDC(lhDC)
        
        '-- Fill image info
        With lpPictDesc
            .Size = Len(lpPictDesc)
            .Type = vbPicTypeBitmap
            .hBmp = lhBmp
            .hPal = 0
        End With
        '-- Fill GUID info
        With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        
        '-- Create picture from bitmap handle
        Call OleCreatePictureIndirect(lpPictDesc, IID_IDispatch, 1, Image)
    End If
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Class:         cIcon.cls (*)
' Author:        Carles P.V.
' Dependencies:  cDIB.cls, mAlphaBlt.bas
' Last revision: 2005.04.04
'================================================
' (*) Simplified version (RO/render version)
'================================================

Option Explicit

'-- API:

Private Type ICONDIR
    idReserved      As Integer ' Reserved
    idType          As Integer ' Resource type
    idCount         As Integer ' Image Count
End Type

Private Type ICONDIRENTRY
    bWidth          As Byte    ' Width of the image
    bHeight         As Byte    ' Height of the image (2 * Height)
    bColorCount     As Byte    ' Number of colors in image (0 when >= 8 bpp)
    bReserved       As Byte    ' Reserved
    wPlanes         As Integer ' Color planes   (-> xHotspot [Cursor])
    wBitCount       As Integer ' Bits per pixel (-> yHotspot [Cursor])
    dwBytesInRes    As Long    ' How many bytes in this resource?
    dwImageOffset   As Long    ' Where in the file is this image?
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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (lpDst As Any, ByVal Length As Long, ByVal Fill As Byte)

'//

'-- Public Enums.:

Public Enum eResourceTypeCts
    [rtIcon] = 1
    [rtCursor] = 2
End Enum

Public Enum eIconBPPCts
    [002_Colors] = 1
    [016_Colors] = 4
    [256_Colors] = 8
    [True_Color] = 24
    [ARGB_Color] = 32
End Enum

'-- Private Types:

Private Type DIBData
    XORDIB As cDIB                     ' XOR DIB section
    ANDDIB As cDIB                     ' AND DIB section
End Type

'-- Private Variables:

Private m_bLoaded      As Boolean      ' Icon <loaded> flag
Private m_sOrderKey()  As String * 8   ' Image format key
Private m_uDir         As ICONDIR      ' Icon file header
Private m_uIconEntry() As ICONDIRENTRY ' Icon image headers
Private m_uDIBData()   As DIBData      ' Icon data (DIBs)

 

'========================================================================================
' Class
'========================================================================================

Private Sub Class_Terminate()
    Call pvClear
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Sub Destroy()
    Call pvClear
End Sub

Public Function LoadFromFile(ByVal Filename As String, _
                             Optional ByVal SortByFormat As Boolean = True _
                             ) As Boolean
    
  Dim eResource  As eResourceTypeCts
  Dim uBIH       As BITMAPINFOHEADER
  Dim aXORBits() As Byte
  Dim aANDBits() As Byte
  Dim aXORPal()  As Byte
  Dim aANDPal(7) As Byte

  Dim hFile As Integer
  Dim nImg  As Integer
    
    '-- Store temp. resource type (->Err)
    eResource = m_uDir.idType
    
    '-- Clear / Activate error handling
    Call pvClear
    On Error GoTo ErrLoad
    
    '-- Open file
    hFile = FreeFile()
    Open Filename For Binary Access Read As hFile
      
        '-- Get icon header
        Get #hFile, , m_uDir
        '-- Basic check
        With m_uDir
            If (.idReserved <> 0 Or (.idType < 0 Or .idType > 1)) Then
                GoTo ErrLoad
            End If
        End With
        
        '-- Get icon entries
        ReDim m_uIconEntry(m_uDir.idCount - 1)
        Get #hFile, , m_uIconEntry()
        
        '-- Initialize arrays and monochrome palette
        ReDim m_sOrderKey(m_uDir.idCount - 1)
        ReDim m_uDIBData(m_uDir.idCount - 1)
        Call FillMemory(aANDPal(4), 3, &HFF)
        
        '-- Get images
        For nImg = 0 To m_uDir.idCount - 1
            
            '-- Move to begin of image data
            Seek #hFile, m_uIconEntry(nImg).dwImageOffset + 1
            
            '-- Load BITMAPINFOHEADER
            Get #hFile, , uBIH
            
            '-- Load XOR palette [?] (<= 8 bpp)
            If (uBIH.biBitCount <= 8) Then
                ReDim aXORPal(4 * 2 ^ uBIH.biBitCount - 1)
                Get #hFile, , aXORPal()
            End If
            
            With m_uDIBData(nImg)
                
                '-- Initialize XOR DIB
                Set .XORDIB = New cDIB
                Call .XORDIB.Create(uBIH.biWidth, uBIH.biHeight \ 2, uBIH.biBitCount)
                If (uBIH.biBitCount <= 8) Then
                    Call .XORDIB.SetPalette(aXORPal())
                End If
                
                '-- Initialize AND DIB
                Set .ANDDIB = New cDIB
                Call .ANDDIB.Create(uBIH.biWidth, uBIH.biHeight \ 2, [01_bpp])
                Call .ANDDIB.SetPalette(aANDPal())
                
                '-- Read DIB bits
                ReDim aXORBits(.XORDIB.Size - 1): Get #hFile, , aXORBits()
                ReDim aANDBits(.ANDDIB.Size - 1): Get #hFile, , aANDBits()
                
                '-- Assign DIB bits
                Call CopyMemory(ByVal .XORDIB.lpBits, aXORBits(0), .XORDIB.Size)
                Call CopyMemory(ByVal .ANDDIB.lpBits, aANDBits(0), .ANDDIB.Size)
                
                '-- Build image format key (sort entries)
                m_sOrderKey(nImg) = Format$(uBIH.biWidth, "000") & Format$(uBIH.biHeight \ 2, "000") & Format$(uBIH.biBitCount, "00")
            End With
        Next nImg

    Close #hFile
    If (SortByFormat) Then
        Call pvSortEntries
    End If
    
'-- Success
    m_bLoaded = True
    LoadFromFile = True
    Exit Function
    
'-- Err.
ErrLoad:
    Close #hFile
    Call pvClear
    m_uDir.idType = eResource
End Function

Public Function Paint(ByVal nIcon As Integer, _
                      ByVal hDC As Long, _
                      Optional ByVal x As Long = 0, _
                      Optional ByVal y As Long = 0 _
                      ) As Boolean

  Dim W As Long
  Dim H As Long

    If (m_bLoaded) Then

        If (nIcon >= 0 And nIcon < m_uDir.idCount) Then

            '-- Paint it
            With m_uDIBData(nIcon)

                W = .XORDIB.Width
                H = .XORDIB.Height
                
                If (.XORDIB.BPP = [32_bpp]) Then
                    '-- ARGB format
                    Call mAlphaBlt.AlphaBlend(hDC, x, y, .XORDIB.hDIB)
                  Else
                    '-- Any other format
                    Call .ANDDIB.Paint(hDC, x, y, vbSrcAnd)
                    Call .XORDIB.Paint(hDC, x, y, vbSrcPaint)
                End If
            End With
            Paint = True
        End If
    End If
End Function

'//

Public Sub SortByFormat()

    '-- Sort entries ###[W]###[H]##[BPP]
    Call pvSortEntries
End Sub

Public Function GetFormatIndex(ByVal Width As Integer, _
                               ByVal Height As Integer, _
                               ByVal BPP As eIconBPPCts _
                               ) As Integer

  Dim sKey As String
  Dim nImg As Integer
    
    '-- Get icon index from given format key
    sKey = Format$(Width, "000") & Format$(Height, "000") & Format$(BPP, "00")
    
    For nImg = 0 To m_uDir.idCount - 1
        If (sKey = m_sOrderKey(nImg)) Then
            '-- Success
            GetFormatIndex = nImg
            Exit Function
        End If
    Next nImg
    
    '-- Format not found
    GetFormatIndex = -1
End Function



'========================================================================================
' Properties
'========================================================================================

'-- Use it carefully!
Public Property Get oXORDIB(ByVal nIcon As Integer) As cDIB
    Set oXORDIB = m_uDIBData(nIcon).XORDIB
End Property

'-- Use it carefully!
Public Property Get oANDDIB(ByVal nIcon As Integer) As cDIB
    Set oANDDIB = m_uDIBData(nIcon).ANDDIB
End Property

'//

Public Property Get ResourceType() As eResourceTypeCts
    ResourceType = m_uDir.idType
End Property

Public Property Get Count() As Integer
    If (m_bLoaded) Then
        Count = m_uDir.idCount
    End If
End Property

Public Property Get Width(ByVal nIcon As Integer) As Integer
    If (m_bLoaded) Then
        Width = m_uDIBData(nIcon).XORDIB.Width
    End If
End Property

Public Property Get Height(ByVal nIcon As Integer) As Integer
    If (m_bLoaded) Then
        Height = m_uDIBData(nIcon).XORDIB.Height
    End If
End Property

Public Property Get BPP(ByVal nIcon As Integer) As eIconBPPCts
    If (m_bLoaded) Then
        BPP = m_uDIBData(nIcon).XORDIB.BPP
    End If
End Property

Public Property Get ImageSize(ByVal nIcon As Integer) As Long
    If (m_bLoaded) Then
        ImageSize = m_uIconEntry(nIcon).dwBytesInRes
    End If
End Property

Public Property Get ImagesSize() As Long

  Dim nImg As Integer
  
    If (m_bLoaded) Then
        For nImg = 0 To m_uDir.idCount - 1
            ImagesSize = ImagesSize + m_uIconEntry(nImg).dwBytesInRes
        Next nImg
    End If
End Property

'//

Public Property Get HotSpotX(ByVal nIcon As Integer) As Integer
    If (m_bLoaded) Then
        HotSpotX = m_uIconEntry(nIcon).wPlanes
    End If
End Property

Public Property Get HotSpotY(ByVal nIcon As Integer) As Integer
    If (m_bLoaded) Then
        HotSpotY = m_uIconEntry(nIcon).wBitCount
    End If
End Property



'========================================================================================
' Private
'========================================================================================

Private Sub pvClear()

  Dim nImg As Integer
  
    '-- Reset header/descriptors
    m_uDir.idCount = 0
    Erase m_uIconEntry()
    Erase m_sOrderKey()
    
    '-- Destroy DIB sections
    If (m_bLoaded) Then
        For nImg = 0 To m_uDir.idCount - 1
            Set m_uDIBData(nImg).XORDIB = Nothing
            Set m_uDIBData(nImg).ANDDIB = Nothing
        Next nImg
        Erase m_uDIBData()
    End If
    
    '-- Reset <loaded> flag
    m_bLoaded = False
End Sub

Private Sub pvSortEntries()
    
  Dim sOrderKey  As String * 8
  Dim uIconEntry As ICONDIRENTRY
  Dim uDIBData   As DIBData
  
  Dim n1         As Integer
  Dim n2         As Integer
    
    '-- Sort by Size, Color depth
    If (m_uDir.idCount > 1) Then
    
        For n1 = 0 To m_uDir.idCount - 2
        For n2 = n1 To m_uDir.idCount - 1
        
            If (m_sOrderKey(n2) < m_sOrderKey(n1)) Then
                
                '-- Swap key!
                sOrderKey = m_sOrderKey(n1)
                m_sOrderKey(n1) = m_sOrderKey(n2)
                m_sOrderKey(n2) = sOrderKey
                
                '-- Swap icon entry
                uIconEntry = m_uIconEntry(n1)
                m_uIconEntry(n1) = m_uIconEntry(n2)
                m_uIconEntry(n2) = uIconEntry
                
                '-- Swap icon data
                uDIBData = m_uDIBData(n1)
                m_uDIBData(n1) = m_uDIBData(n2)
                m_uDIBData(n2) = uDIBData
            End If
        Next n2, n1
    End If
End Sub

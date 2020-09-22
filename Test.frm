VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "mAlphaBlt test"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPaint 
      Caption         =   "&Paint"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1785
      TabIndex        =   0
      Top             =   3030
      Width           =   1005
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oDIB32 As cDIB32
Private m_oTile  As cTile
Private m_oT     As cTiming


Private Sub Form_Load()
    
    Set Me.Icon = Nothing
    
    Set m_oDIB32 = New cDIB32
    Set m_oTile = New cTile
    Set m_oT = New cTiming
    
    '-- Load 32-bit bitmap
    'Call m_oDIB32.CreateFromResourceBitmap(pvFixPath(App.Path) & "Test.exe", 101)
    Call m_oDIB32.CreateFromBitmapFile(pvFixPath(App.Path) & "Test32x32.bmp")
    
    '-- Load pattern
    Call m_oTile.CreatePatternFromStdPicture(LoadResPicture(102, vbResBitmap))
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set m_oDIB32 = Nothing
    Set m_oTile = Nothing
End Sub



Private Sub cmdPaint_Click()
  
  Dim i As Long
  Dim j As Long
  Dim k As Long
    
    'Call m_oTile.Tile(Me.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
    
    Screen.MousePointer = vbArrowHourglass
    
    Me.CurrentY = 0
    Me.Print "Rendering..." & Space$(100)
    
    Call m_oT.Reset
    For k = 1 To 250
        For j = 1 To 5
            For i = 1 To 8
                 Call AlphaBlt(Me.hDC, i * 32 - 10, j * 32 - 10, Me.BackColor, m_oDIB32.hDIB)
                 'Call AlphaBlend(Me.hDC, i * 32 - 10, j * 32 - 10, m_oDIB32.hDIB)
            Next i
        Next j
    Next k
    
    Me.CurrentY = 0
    Me.Print 250 * 5 * 8 & " alpha bitmaps rendered in " & Format$(m_oT.Elapsed / 1000, "0.000 sec.")
    
    Screen.MousePointer = vbDefault
End Sub



Private Function pvFixPath(ByVal sPath As String) As String
    pvFixPath = sPath & IIf(Right$(sPath, 1) = "\", vbNullString, "\")
End Function


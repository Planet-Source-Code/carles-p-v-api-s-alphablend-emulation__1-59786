VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BMP32Export"
   ClientHeight    =   4455
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   StartUpPosition =   2  'CenterScreen
   Begin BMP32Export.ucReportList ucReportList 
      Height          =   2040
      Left            =   360
      TabIndex        =   1
      Top             =   690
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3598
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picView 
      AutoRedraw      =   -1  'True
      Height          =   2040
      Left            =   2760
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   132
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   690
      Width           =   2040
   End
   Begin VB.Label lblFileInfo 
      BackColor       =   &H00808080&
      Caption         =   " File info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   360
      TabIndex        =   4
      Top             =   3045
      Width           =   4440
   End
   Begin VB.Label lblFormatPreview 
      BackColor       =   &H00808080&
      Caption         =   " Format preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   2040
   End
   Begin VB.Label lblFormatsList 
      BackColor       =   &H00808080&
      Caption         =   " Formats list"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblInfo 
      Height          =   825
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   4440
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open icon resource..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Export 32-bpp bitmap..."
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   3
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sPath As String
Private m_oIcon As cIcon
Private m_oTile As cTile



Private Sub Form_Load()
    
    '-- No default app. icon!
    Set Me.Icon = Nothing
    
    '-- Initialize Icon object
    Set m_oIcon = New cIcon
    
    '-- Initialize Tile object (load pattern)
    Set m_oTile = New cTile
    Call m_oTile.CreatePatternFromStdPicture(LoadResPicture(101, vbResBitmap))
    
    '-- Initialize report list
    Call ucReportList.AddHeader(0)
    Call ucReportList.AddHeader(60, [RightJustify], "Size")
    Call ucReportList.AddHeader(40, [RightJustify], "BPP")
End Sub

Private Sub Form_Paint()
    
    Me.Line (0, 0)-(Me.ScaleWidth, 0), vb3DShadow
    Me.Line (0, 1)-(Me.ScaleWidth, 1), vb3DHighlight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set m_oIcon = Nothing
    Set m_oTile = Nothing
    Set fMain = Nothing
End Sub



Private Sub mnuFile_Click(Index As Integer)
    
  Dim sPath As String
  
    Select Case Index
    
        Case 0 '-- Open
            
            sPath = mDialogFile.GetFileName(Me.hWnd, m_sPath, "Icon|*.ico|Cursor|*.cur", , "Open", OpenMode:=True, ViewMode:=ViewLargeIcon)
            If (sPath <> vbNullString) Then
                If (m_oIcon.LoadFromFile(sPath, SortByFormat:=True)) Then
                    m_sPath = sPath
                    Call pvShowFileInfo
                    Call pvFillFormatsList
                  Else
                    Call MsgBox("Incorrect file format or an unexpected error has occurred.", vbExclamation)
                    Call ucReportList.ClearList
                    Call picView.Cls
                    mnuFile(1).Enabled = False
                End If
            End If
        
        Case 1 '-- Export
            
            sPath = pvFixPath(m_sPath)
            sPath = mDialogFile.GetFileName(Me.hWnd, sPath, "Bitmap|*.bmp", , "Export", OpenMode:=False, ViewMode:=ViewLargeIcon)
            If (sPath <> vbNullString) Then
                If (m_oIcon.oXORDIB(ucReportList.ListIndex).Save(sPath)) Then
                  Else
                    Call MsgBox("Unexpected error exporting bitmap.", vbExclamation)
                End If
            End If
            
        Case 3 '-- Exit
            
            Call Unload(Me)
    End Select
End Sub

Private Sub ucReportList_Click()
    
    With picView
        
        '-- Paint background pattern
        Call m_oTile.Tile(.hDC, 0, 0, .ScaleWidth, .ScaleHeight)
        '-- Paint/render icon
        Screen.MousePointer = vbArrowHourglass
        Call m_oIcon.Paint(ucReportList.ListIndex, .hDC, 2, 2)
        Screen.MousePointer = vbArrow
        
        '-- Refresh view
        Call .Refresh
    End With
    
    '-- Can be exported (32-bit)?
    mnuFile(1).Enabled = (m_oIcon.BPP(ucReportList.ListIndex) = [ARGB_Color])
End Sub

Private Sub picView_DblClick()
    
    If (mnuFile(1).Enabled) Then
        Call mnuFile_Click(1) 'Export
    End If
End Sub



Private Sub pvFillFormatsList()

  Dim nIdx As Long
  
    With ucReportList
        Call .ClearList
        For nIdx = 0 To m_oIcon.Count - 1
            Call .AddItem(vbTab & m_oIcon.Width(nIdx) & "x" & m_oIcon.Height(nIdx) & vbTab & m_oIcon.BPP(nIdx))
        Next nIdx
        Let .ListIndex = 0
    End With
End Sub

Private Sub pvShowFileInfo()

    With lblInfo
        .Caption = m_sPath & vbCrLf
        .Caption = .Caption & "Image formats: " & m_oIcon.Count & vbCrLf
        .Caption = .Caption & "Images size: " & Format$(m_oIcon.ImagesSize, "#,# bytes")
    End With
End Sub

Private Function pvFixPath(ByVal sPath As String) As String
    
  Dim nIdx As Integer
    
    nIdx = ucReportList.ListIndex
    
    With m_oIcon
        pvFixPath = Left$(sPath, Len(sPath) - 4)
        pvFixPath = pvFixPath & "_" & .Width(nIdx) & "x" & .Height(nIdx) & "x32bpp"
        pvFixPath = pvFixPath & ".bmp"
    End With
End Function

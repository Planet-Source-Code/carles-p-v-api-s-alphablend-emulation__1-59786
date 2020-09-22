Attribute VB_Name = "mDialogFile"
'================================================
' Module:        mDialogFile.bas
' Author:
' Dependencies:  None
' Last revision: 2003.03.28
'================================================

Option Explicit
Option Compare Text

'-- API:

'-- Open & Save Dialog
Private Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    Flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type

'-- Hook and notification support
Private Type NMHDR
    hwndFrom As Long
    IDFrom   As Long
    Code     As Long
End Type

Private Type OFNOTIFYshort
    HDR   As NMHDR
    lpOFN As Long
End Type

Private Type LV_ITEM
    Mask       As Long
    iItem      As Long
    iSubItem   As Long
    State      As Long
    StateMask  As Long
    pszText    As String
    cchTextMax As Long
    iImage     As Long
    lParam     As Long
    iIndent    As Long
End Type
 
Private Const LVM_FIRST          As Long = &H1000
Private Const LVM_GETNEXTITEM    As Long = LVM_FIRST + 12
Private Const LVM_GETITEMTEXT    As Long = LVM_FIRST + 45
Private Const LVM_SETVIEW        As Long = LVM_FIRST + 142
Private Const LVNI_FOCUSED       As Long = &H1
Private Const LVNI_SELECTED      As Long = &H2
Private Const LV_VIEW_ICON       As Long = &H0
Private Const LV_VIEW_DETAILS    As Long = &H1
Private Const LV_VIEW_SMALLICON  As Long = &H2
Private Const LV_VIEW_LIST       As Long = &H3
Private Const LV_VIEW_MAX        As Long = &H4
Private Const LV_VIEW_TILE       As Long = &H4

Private Const ID_OPEN            As Long = &H1   ' Open or Save button
Private Const ID_CANCEL          As Long = &H2   ' Cancel Button
Private Const ID_HELP            As Long = &H40E ' Help Button
Private Const ID_READONLY        As Long = &H410 ' Read-only check box
Private Const ID_FILETYPELABEL   As Long = &H441 ' FileType label
Private Const ID_FILELABEL       As Long = &H442 ' FileName label
Private Const ID_FOLDERLABEL     As Long = &H443 ' Folder label
Private Const ID_LIST            As Long = &H461 ' Parent of file list
Private Const ID_FORMAT          As Long = &H470 ' FileType combo box
Private Const ID_FOLDER          As Long = &H471 ' Folder combo box
Private Const ID_FILETEXT        As Long = &H480 ' FileName text box

Private Const OFN_OPENFLAGS      As Long = &H881024
Private Const OFN_SAVEFLAGS      As Long = &H880026
    
Private Const WM_INITDIALOG      As Long = &H110
Private Const WM_COMMAND         As Long = &H111
Private Const WM_DESTROY         As Long = &H2
Private Const WM_NOTIFY          As Long = &H4E
Private Const WM_SETICON         As Long = &H80

Private Const WM_USER            As Long = &H400
Private Const MYWM_POSTINIT      As Long = WM_USER + 1
Private Const CDM_FIRST          As Long = (WM_USER + 100)
Private Const CDM_GETSPEC        As Long = (CDM_FIRST + &H0)
Private Const CDM_GETFILEPATH    As Long = (CDM_FIRST + &H1)
Private Const CDM_GETFOLDERPATH  As Long = (CDM_FIRST + &H2)
Private Const CDM_SETCONTROLTEXT As Long = (CDM_FIRST + &H4)
Private Const CDM_HIDECONTROL    As Long = (CDM_FIRST + &H5)
Private Const CDM_SETDEFEXT      As Long = (CDM_FIRST + &H6)
Private Const CB_GETCURSEL       As Long = &H147

Private Const CDN_FIRST          As Long = -601&
Private Const CDN_INITDONE       As Long = (CDN_FIRST)
Private Const CDN_SELCHANGE      As Long = (CDN_FIRST - &H1)
Private Const CDN_FOLDERCHANGE   As Long = (CDN_FIRST - &H2)
Private Const CDN_HELP           As Long = (CDN_FIRST - &H4)
Private Const CDN_FILEOK         As Long = (CDN_FIRST - &H5)
Private Const CDN_TYPECHANGE     As Long = (CDN_FIRST - &H6)

Private Const GW_HWNDFIRST       As Long = 0
Private Const GW_HWNDNEXT        As Long = 2
Private Const GW_CHILD           As Long = 5

Private Const MAX_PATH           As Long = 260

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'//

' -- http://mvps.org/vbvision/grouped_demos.htm#Common_Dialogs
' -- MESSAGES PROVIDED BY Brad Martinez
'    View menu command IDs fall within the FCIDM_SHVIEWFIRST/LAST
'    range defined in ShlObj.h:
  
Private Const FCIDM_SHVIEW_LARGEICON As Long = &H7029& ' 28713
Private Const FCIDM_SHVIEW_SMALLICON As Long = &H702A& ' 28714
Private Const FCIDM_SHVIEW_LIST      As Long = &H702B& ' 28715
Private Const FCIDM_SHVIEW_REPORT    As Long = &H702C& ' 28716
Private Const FCIDM_SHVIEW_THUMBNAIL As Long = &H702D& ' 28717
Private Const FCIDM_SHVIEW_TILE      As Long = &H702E& ' 28718

'//

'-- Public enums.:

Public Enum eViewMode
    [ViewLargeIcon] = FCIDM_SHVIEW_LARGEICON
    [ViewSmallIcon] = FCIDM_SHVIEW_SMALLICON
    [ViewList] = FCIDM_SHVIEW_LIST
    [ViewReport] = FCIDM_SHVIEW_REPORT
    [ViewThumbnail] = FCIDM_SHVIEW_THUMBNAIL
    [ViewTile] = FCIDM_SHVIEW_TILE
End Enum

'-- Private Variables:

Private m_hDlg       As Long
Private m_bOpenMode  As Boolean
Private m_eViewMode  As eViewMode
Private m_sDlgFilter As String
Private m_sCurExt    As String



'========================================================================================
' Methods
'========================================================================================

Public Function GetFileName(ByVal hwndOwner As Long, _
                            Optional sPath As String, _
                            Optional sFilter As String, _
                            Optional nFltIndex As Long, _
                            Optional sTitle As String, _
                            Optional OpenMode As Boolean = True, _
                            Optional ViewMode As eViewMode = [ViewList] _
                            ) As String
   
  Dim uOFN As OPENFILENAME
  Dim lRet As Long
  Dim lIdx As Long
  Dim sExt As String
 
    m_hDlg = 0
    m_bOpenMode = OpenMode
    m_eViewMode = ViewMode
   
    For lIdx = 1 To Len(sFilter)
        If (Mid$(sFilter, lIdx, 1) = "|") Then
            Mid$(sFilter, lIdx, 1) = vbNullChar
        End If
    Next lIdx
    
    If (Len(sFilter) < MAX_PATH) Then
        sFilter = sFilter & String$(MAX_PATH - Len(sFilter), 0)
      Else
        sFilter = sFilter & Chr$(0) & Chr$(0)
    End If
    m_sDlgFilter = sFilter
    
    If (sPath <> vbNullString And (nFltIndex = 0)) Then
        nFltIndex = pvGetFilterIndex(sPath)
    End If
        
    With uOFN
        .hwndOwner = hwndOwner
        .lStructSize = Len(uOFN)
        .lpstrTitle = sTitle
        .lpstrFile = sPath & String(MAX_PATH - Len(sPath), 0)
        .lpstrFilter = sFilter
        .lpfnHook = pvHookAddress(AddressOf pvDialogHookProcess)
        .hInstance = App.hInstance
        .nFilterIndex = nFltIndex
        .nMaxFile = MAX_PATH
    End With
   
    If (m_bOpenMode) Then
        uOFN.Flags = uOFN.Flags Or OFN_OPENFLAGS
        lRet = GetOpenFileName(uOFN)
      Else
        uOFN.Flags = uOFN.Flags Or OFN_SAVEFLAGS
        lRet = GetSaveFileName(uOFN)
    End If
    
    If (lRet) Then
        GetFileName = pvTrimNull(uOFN.lpstrFile)
        If (uOFN.nFileExtension = 0) And Len(m_sCurExt) > 2 Then
            GetFileName = GetFileName & Mid$(m_sCurExt, 2)
        End If
    End If
End Function

Public Property Get Extension() As String
    Extension = m_sCurExt
End Property



'========================================================================================
' Private
'========================================================================================

Private Function pvHookAddress(lPtr As Long) As Long
    pvHookAddress = lPtr
End Function

Private Function pvDialogHookProcess(ByVal hDlg As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     ByVal lParam As Long _
                                     ) As Long
   
  Dim uNMH  As NMHDR
  Dim hLV   As Long
  Dim sPath As String
  Dim sExt  As String
  Dim lPos  As Long
  
    Select Case wMsg
    
        Case WM_NOTIFY
        
            Call CopyMemory(uNMH, ByVal lParam, Len(uNMH))
        
            Select Case uNMH.Code
              
                Case CDN_FOLDERCHANGE
            
                    hLV = FindWindowEx(GetParent(hDlg), 0, "SHELLDLL_DefView", vbNullString)
                    If (hLV) Then
                        Call SendMessage(hLV, WM_COMMAND, m_eViewMode, ByVal 0&)
                    End If
                    
                    Call SendMessage(m_hDlg, CDM_SETCONTROLTEXT, ID_FILETEXT, ByVal vbNullString)
                
                Case CDN_SELCHANGE
                
                    sPath = pvGetSelItem
                    If (sPath <> vbNullString) Then
                        Call SendMessage(m_hDlg, CDM_SETCONTROLTEXT, ID_FILETEXT, ByVal sPath)
                    End If
                
                Case CDN_TYPECHANGE
              
                    If (Not m_bOpenMode) Then
                    
                        sPath = String(MAX_PATH, 0)
                        Call SendMessage(m_hDlg, CDM_GETSPEC, MAX_PATH, ByVal sPath)
                        sPath = pvTrimNull(sPath)
                        
                        If (Len(sPath) > 4) Then
                            sExt = Right$(sPath, 5)
                            lPos = InStr(1, sExt, ".")
                            If (lPos) Then
                                sPath = Left$(sPath, Len(sPath) - 6 + lPos)
                            End If
                        End If
                        
                        m_sCurExt = pvGetExtension()
                        If (Len(m_sCurExt) > 2) Then
                            Call SendMessage(m_hDlg, CDM_SETDEFEXT, 0, ByVal Mid$(m_sCurExt, 3))
                        End If
                        Call SendMessage(m_hDlg, CDM_SETCONTROLTEXT, ID_FILETEXT, ByVal sPath)
                    End If
                
                Case CDN_INITDONE
                
                    m_hDlg = GetParent(hDlg)
            End Select
               
        Case WM_DESTROY
        
            m_sCurExt = pvGetExtension()
   End Select
End Function

Private Function pvTrimNull(StartStr As String) As String
  
  Dim lPos As Long
  
    lPos = InStr(StartStr, Chr$(0))
    If (lPos) Then
        pvTrimNull = Left$(StartStr, lPos - 1)
      Else
        pvTrimNull = StartStr
    End If
End Function

Private Function pvGetSelItem() As String
  
  Static sOldPath As String
  
  Dim uLVI        As LV_ITEM
  Dim lRet        As Long
  Dim hFileList   As Long
  Dim sPath       As String
  Dim sNewPath    As String
   
    sNewPath = String(MAX_PATH, 0)
    Call SendMessage(m_hDlg, CDM_GETFILEPATH, MAX_PATH, ByVal sNewPath)
    sNewPath = pvTrimNull(sNewPath)
    
    If (sNewPath <> sOldPath) Then
        sOldPath = sNewPath
        Exit Function
    End If
    
    hFileList = GetDlgItem(GetDlgItem(m_hDlg, ID_LIST), 1)
    
    If (hFileList <> 0) Then
        lRet = SendMessage(hFileList, LVM_GETNEXTITEM, -1, ByVal LVNI_SELECTED)
        
        If (lRet <> -1) Then
            uLVI.cchTextMax = MAX_PATH
            uLVI.pszText = Space$(MAX_PATH)
            lRet = SendMessage(hFileList, LVM_GETITEMTEXT, lRet, uLVI)
            
            If (lRet > 1) Then
                sPath = Left$(uLVI.pszText, lRet)
            End If
            pvGetSelItem = sPath
            sOldPath = sPath
        End If
    End If
End Function

Private Function pvGetExtension() As String

  Dim lIdx    As Long
  Dim lFilter As Long
  Dim lStart  As Long
  Dim hCombo  As Long
  Dim sFilter As String
  Dim sTemp   As String
   
    hCombo = GetDlgItem(m_hDlg, ID_FORMAT)
    lFilter = SendMessage(hCombo, CB_GETCURSEL, 0, ByVal 0&)
    sFilter = m_sDlgFilter
   
    For lIdx = 1 To lFilter * 2 + 1
        lStart = InStr(1, sFilter, Chr$(0))
        If (lStart) Then
            sFilter = Mid$(sFilter, lStart + 1)
          Else
            Exit For
        End If
    Next lIdx
    
    sTemp = pvTrimNull(sFilter)
    If (Len(sTemp) <> 0) Then
        If (InStr(1, sTemp, ";") = 0) Then
            pvGetExtension = sTemp
        End If
    End If
End Function

Private Function pvGetFilterIndex(ByVal sPath As String) As Long

  Dim sExt   As String
  Dim lIdx   As Long
  Dim lStart As Long
  
    sExt = Right$(sPath, 4)
    
    If (Left$(sExt, 1) = ".") Then
        sExt = Mid$(sExt, 2)
    End If
    sExt = "*." & sExt & Chr$(0)
    
    lStart = 1
    Do While lStart
        lStart = InStr(lStart + 1, m_sDlgFilter, Chr$(0), vbTextCompare)
        If (Mid$(m_sDlgFilter, lStart + 1, Len(sExt)) = sExt) Then
            Exit Do
        End If
        lIdx = lIdx + 1
    Loop
    pvGetFilterIndex = lIdx \ 2 + 1
End Function

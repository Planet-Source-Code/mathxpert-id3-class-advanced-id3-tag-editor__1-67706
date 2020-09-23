Attribute VB_Name = "modBrowse"
Option Explicit

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_BROWSEINCLUDEURLS = 128
Private Const BIF_EDITBOX = 16
Private Const BIF_NEWDIALOGSTYLE = 64
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_STATUSTEXT = 4
Private Const BIF_VALIDATE = 32
Public Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_VALIDATEFAILEDA = 3
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Dim m_StartFolder As String
Dim bValidateFailed As Boolean

Public Function BrowseForFolder(ByVal hwndOwner As Long, ByVal Prompt As String, Optional ByVal StartFolder) As String
    Dim lNull As Long
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo

    With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = Prompt
        .ulFlags = BIF_BROWSEINCLUDEURLS Or BIF_NEWDIALOGSTYLE Or BIF_EDITBOX Or BIF_VALIDATE Or BIF_RETURNONLYFSDIRS Or BIF_STATUSTEXT
        If Not IsMissing(StartFolder) Then
            m_StartFolder = StartFolder
            If Right$(m_StartFolder, 1) <> Chr$(0) Then m_StartFolder = m_StartFolder & Chr$(0)
            .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
        End If
    End With

    bValidateFailed = False
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList And Not bValidateFailed Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        CoTaskMemFree lpIDList
        lNull = InStr(sPath, vbNullChar)
        If lNull Then
            sPath = Left$(sPath, lNull - 1)
        End If
    End If

    BrowseForFolder = sPath
End Function

Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    On Error Resume Next
    
    Dim lpIDList As Long
    Dim ret As Long
    Dim sBuffer As String
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            SendMessageA hWnd, BFFM_SETSELECTION, 1, m_StartFolder
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH)
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                SendMessageA hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer
            End If
        Case BFFM_VALIDATEFAILEDA
            bValidateFailed = True
    End Select
    
    BrowseCallbackProc = 0
End Function

Private Function GetAddressofFunction(Add As Long) As Long
    GetAddressofFunction = Add
End Function

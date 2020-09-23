Attribute VB_Name = "basCommonDialog"
Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_FILEMUSTEXIST = &H1000

Private Declare Function GetOpenFileNameA Lib "comdlg32" (pOpenfilename As OPENFILENAME) As Long

Public Function ShowOpenDialog(ByVal hwndOwner As Long, ByVal Filter As String, ByVal Title As String, Optional ByVal InitialDirectory) As String
    Dim lNull As Long
    Dim sFilter As String
    Dim sFile As String
    Dim lngOpenFileName As OPENFILENAME
    Dim lResult As Long
    
    With lngOpenFileName
        .lStructSize = Len(lngOpenFileName)
        .hwndOwner = hwndOwner
        .hInstance = App.hInstance
        sFilter = Filter
        If Right$(sFilter, 1) <> "|" Then sFilter = sFilter & "|"
        .lpstrFilter = Replace(sFilter, "|", Chr$(0))
        .lpstrFile = String$(MAX_PATH, 0)
        .nMaxFile = MAX_PATH
        .lpstrFileTitle = String$(MAX_PATH, 0)
        .nMaxFileTitle = MAX_PATH
        .lpstrInitialDir = IIf(IsMissing(InitialDirectory), vbNullString, InitialDirectory)
        .lpstrTitle = Title
        .flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        
        lResult = GetOpenFileNameA(lngOpenFileName)
        If lResult Then
            sFile = .lpstrFile
            lNull = InStr(sFile, vbNullChar)
            If lNull Then
                sFile = Left$(sFile, lNull - 1)
            End If
        End If
    End With
    
    ShowOpenDialog = sFile
End Function

Attribute VB_Name = "mMain"
Option Explicit

Private Const ICC_BAR_CLASSES As Long = &H4
Private Const ICC_LISTVIEW_CLASSES As Long = &H1
Private Const ICC_PROGRESS_CLASS As Long = &H20
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const ICC_TAB_CLASSES As Long = &H8
Private Const ICC_UPDOWN_CLASS As Long = &H10
Private Const ICC_USEREX_CLASSES As Long = &H200
Private Const ICC_WIN95_CLASSES As Long = &HFF&

Private Type InitCommonControlsEx
    dwSize As Long 'size of this structure
    dwICC As Long 'flags indicating which classes to be initialized
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (ByRef TLPINITCOMMONCONTROLSEX As InitCommonControlsEx) As Long

Public Sub Main()
    On Error Resume Next
    Dim iccex As InitCommonControlsEx
    With iccex
        .dwSize = Len(iccex)
        .dwICC = ICC_BAR_CLASSES Or ICC_LISTVIEW_CLASSES Or ICC_PROGRESS_CLASS Or ICC_STANDARD_CLASSES Or ICC_TAB_CLASSES Or ICC_UPDOWN_CLASS Or ICC_USEREX_CLASSES Or ICC_WIN95_CLASSES
    End With
    InitCommonControlsEx iccex
    On Error GoTo 0
    frmMain.Show
End Sub

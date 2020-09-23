Attribute VB_Name = "modSizeLimitAndWheel"
Option Explicit

Private Const GWL_WNDPROC = -4
Private Const WM_GETMINMAXINFO = &H24
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_MOVING As Long = &H216
Private Const WM_SIZE As Long = &H5
Private Const WM_SIZING As Long = &H214

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type COMBOBOXINFO
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton  As Long
   hwndCombo  As Long
   hwndEdit  As Long
   hwndList As Long
End Type

Global lpPrevWndProc As Long
Global gHW As Long

Public XPos As Long
Public YPos As Long

Public XSize As Long
Public YSize As Long

Private Declare Function DefWindowProc Lib "user32" Alias _
   "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias _
   "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
   "SetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemoryToMinMaxInfo Lib "kernel32" Alias _
   "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, _
    ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "kernel32" Alias _
   "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, _
    ByVal cbCopy As Long)

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function GetComboBoxInfo Lib "user32" _
  (ByVal hwndCombo As Long, _
   CBInfo As COMBOBOXINFO) As Long

Public Sub Hook()
    'Start subclassing.
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
       AddressOf WindowProc)
End Sub

Public Sub Unhook()
    Dim temp As Long

    'Cease subclassing.
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim MinMax As MINMAXINFO
    
    Dim pt As POINTAPI
    Dim Cmb As COMBOBOXINFO
    Dim lng As Long
    Dim bDefWndProc As Boolean
    
    bDefWndProc = False
    GetCursorPos pt
    lng = WindowFromPoint(pt.x, pt.Y)
    
    'Check for request for min/max window sizes.
    If uMsg = WM_GETMINMAXINFO Then
        bDefWndProc = True
        
        'Retrieve default MinMax settings
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)

        'Specify new minimum size for window.
        MinMax.ptMinTrackSize.x = 568
        MinMax.ptMinTrackSize.Y = 445

        'Specify new maximum size for window.
        'MinMax.ptMaxTrackSize.x = 0
        'MinMax.ptMaxTrackSize.y = 0

        'Copy local structure back.
        CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)
    ElseIf uMsg = WM_MOUSEWHEEL Then
        With frmMain
            If .VScroll1.Visible Then
                ' Obtain a handle based on the mouse position
                GetCursorPos pt
                lng = WindowFromPoint(pt.x, pt.Y)
                
                ' We need to obtain all the handles of the initial key combo box
                Cmb.cbSize = Len(Cmb)
                GetComboBoxInfo .cmbKey.hWnd, Cmb
                
                ' Does the handle match any handle related to the advanced view?
                Select Case lng
                    Case .txtComposer.hWnd, .txtBand.hWnd, .txtConductor.hWnd, .txtInterpretedBy.hWnd, _
                         .txtLyricist.hWnd, .txtOriginalArtist.hWnd, .txtOriginalAlbum.hWnd, _
                         .txtOriginalFileName.hWnd, .txtOriginalLyricist.hWnd, .txtOriginalReleaseYear.hWnd, _
                         .txtCopyright.hWnd, .txtFileOwner.hWnd, .txtPublisher.hWnd, _
                         .txtInternetRadioStationName.hWnd, .txtInternetRadioStationOwner.hWnd, _
                         .txtISRC.hWnd, .txtLanguages.hWnd, .txtCommercialInfo.hWnd, _
                         .txtCopyrightInfo.hWnd, .txtAudioURL.hWnd, .txtArtistURL.hWnd, .txtAudioSourceURL.hWnd, _
                         .txtInternetRadioStationURL.hWnd, .txtPaymentURL.hWnd, .txtPublisherURL.hWnd, _
                         .txtEncodedBy.hWnd, .txtBPM.hWnd, Cmb.hwndCombo, Cmb.hwndEdit, Cmb.hwndList, _
                         .txtDiscNumber.hWnd, .txtDiscsTotal.hWnd, .Frame3.hWnd, .VScroll1.hWnd
                            If wParam < 0 Then ' Scrolling down
                                If .VScroll1.Value < .VScroll1.Max Then
                                    .VScroll1.Value = .VScroll1.Value + 1
                                End If
                            ElseIf wParam > 0 Then ' Scrolling up
                                If .VScroll1.Value > .VScroll1.Min Then
                                    .VScroll1.Value = .VScroll1.Value - 1
                                End If
                            End If
                End Select
            End If
        End With
    ElseIf uMsg = WM_MOVING Or uMsg = WM_SIZE Then
        With frmMain
            If .WindowState = vbNormal And .Visible Then
                If XPos <> .Left Then _
                    XPos = .Left
                If YPos <> .Top Then _
                    YPos = .Top
            End If
        End With
    ElseIf uMsg = WM_SIZING Then
        With frmMain
            If .WindowState = vbNormal And .Visible Then
                If XSize <> .Width Then _
                    XSize = .Width
                If YSize <> .Height Then _
                    YSize = .Height
            End If
        End With
    End If
    
    If bDefWndProc Then
        WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
    Else
        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, _
           wParam, lParam)
    End If
End Function


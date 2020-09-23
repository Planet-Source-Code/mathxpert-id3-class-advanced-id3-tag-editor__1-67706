Attribute VB_Name = "modMultidata"
Option Explicit

Public cWCOM As Collection
Public cWOAR As Collection
Public cAPICData As Collection
Public cAPICIType As Collection
Public cAPICType As Collection
Public cAPIC0 As Collection

Public APICData As MultiFrameData

Public indWCOM As Long
Public totWCOM As Long

Public indWOAR As Long
Public totWOAR As Long

Public indAPIC As Long
Public totAPIC As Long

Public bWCOMBlank As Boolean
Public bWOARBlank As Boolean
Public bAPICBlank As Boolean

Public ID3Revision As Byte

Public Const S_AURL As String = "Artist URL"
Public Const S_CURL As String = "Commercial Info URL"
Public Const S_APIC As String = "Album Art"
Public Const S_APICTT As String = "Click here to browse... (this will change the current image)"

Public Const S_ADD As String = "Add "
Public Const S_PREV As String = "Previous "
Public Const S_NEXT As String = "Next "
Public Const S_DEL As String = "Delete "

Public Const I_ADD As String = "add"
Public Const I_ADDI As String = "addi"
Public Const I_PREV As String = "prev"
Public Const I_PREVI As String = "previ"
Public Const I_NEXT As String = "next"
Public Const I_DEL As String = "del"
Public Const I_DELI As String = "deli"

Public Sub SetBG(ByVal Black As Boolean)
    With frmMain.picArt
        If Black Then
            If .BackColor <> vbBlack Then .BackColor = vbBlack
        Else
            If .BackColor <> vbApplicationWorkspace Then .BackColor = vbApplicationWorkspace
        End If
    End With
End Sub

Public Sub StretchImage(ByVal Pic As StdPicture)
    On Error GoTo MyErr
    
    Dim Div As Double
    Dim Horizontal As Boolean
    With frmMain.imgArt
        Horizontal = (Pic.Width < Pic.Height)
        If Horizontal Then
            Div = CDbl(Pic.Width) / CDbl(Pic.Height)
            .Width = frmMain.picArt.ScaleWidth * Div
            .Height = frmMain.picArt.ScaleHeight
        Else
            Div = CDbl(Pic.Height) / CDbl(Pic.Width)
            .Width = frmMain.picArt.ScaleWidth
            .Height = frmMain.picArt.ScaleHeight * Div
        End If
        .Left = (frmMain.picArt.ScaleWidth - .Width) / 2
        .Top = (frmMain.picArt.ScaleHeight - .Height) / 2
        Exit Sub
MyErr:
        .Width = frmMain.picArt.ScaleWidth
        .Height = frmMain.picArt.ScaleHeight
        .Left = 0
        .Top = 0
    End With
End Sub

Public Sub LoadMultiData(Ctl As Object, Frame As MultiFrameData, ByVal Description As String, CountControl As Label, PrevControl As Image, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean)
    Dim sNext As String
    Dim sNextT As String
    Dim sDel As String
    Dim sDelT As String
    Dim i As Long
    Dim bPic As Boolean
    Dim bPicQualified As Boolean
    Dim APD As APicDecoder
    Dim GPC As GDIPlusCandy
    
    Dim MIMEType As String
    Dim PictureType As PictureType
    Dim Pic As StdPicture
    Dim OrigPicData As String
    
    bPic = (TypeName(Ctl) = "PictureBox")
    Set Col = Nothing
    Set Col = New Collection
    If bPic Then
        Set APICData = Nothing
        Set APICData = Frame
        Set cAPIC0 = Nothing
        Set cAPIC0 = New Collection
        Set cAPICIType = Nothing
        Set cAPICIType = New Collection
        Set cAPICType = Nothing
        Set cAPICType = New Collection
        With frmMain.cmbImageType
            If ID3Revision > 2 And .ListCount = 2 Then
                .AddItem "BMP", 0
                .AddItem "GIF", 1
            ElseIf ID3Revision <= 2 And .ListCount = 4 Then
                .RemoveItem 0
                .RemoveItem 0
            End If
        End With
    End If
    bPicQualified = (bPic And Frame.Count > 0)
    
    If bPicQualified Then Set APD = New APicDecoder
    For i = 1 To Frame.Count
        If Replace(Frame(i), Chr$(0), "") <> "" Then
            If bPic Then
                If APD.DecodeImage(Frame, i, MIMEType, PictureType, Pic, ID3Revision, OrigPicData) Then
                    Col.Add OrigPicData
                    cAPICIType.Add MIMEType
                    cAPICType.Add PictureType
                    cAPIC0.Add i
                End If
            Else
                Col.Add Frame(i)
            End If
        End If
    Next
    If bPicQualified Then Set APD = Nothing
    
    If Col.Count = 0 Then
        Index = 0
        Total = 0
        FrameBlank = True
        If bPic Then
            frmMain.cmbImageType.Enabled = False
            frmMain.cmbPictureType.Enabled = False
            SetBG False
            Set frmMain.imgArt.Picture = Nothing
            StretchImage frmMain.imgArt.Picture
            frmMain.imgArt.Visible = False
            frmMain.lblBrowse.Visible = True
            frmMain.cmbImageType.ListIndex = 2 * (frmMain.cmbImageType.ListCount \ 4)
            frmMain.cmbPictureType.ListIndex = 0
            Ctl.ToolTipText = ""
            frmMain.imgArt.ToolTipText = ""
        Else
            Ctl.Text = ""
        End If
        sNext = I_ADDI: sNextT = ""
        sDel = I_DELI: sDelT = ""
    Else
        Index = 1
        Total = Col.Count
        FrameBlank = (Col.Count = 1 And Col(1) = "")
        If bPic Then
            frmMain.cmbImageType.Enabled = True
            frmMain.cmbPictureType.Enabled = True
            frmMain.lblBrowse.Visible = False
            frmMain.imgArt.Visible = True
            Set GPC = New GDIPlusCandy
            Set Pic = GPC.DataToImage(Col(1))
            Set GPC = Nothing
            Set frmMain.imgArt.Picture = Nothing
            StretchImage Pic
            Set frmMain.imgArt.Picture = Pic
            SetBG True
            frmMain.cmbImageType.ListIndex = GetIndex(cAPICIType(1), ID3Revision)
            frmMain.cmbPictureType.ListIndex = cAPICType(1)
            Ctl.ToolTipText = S_APICTT
            frmMain.imgArt.ToolTipText = S_APICTT
        Else
            Ctl.Text = Col(1)
        End If
        If Col.Count > 1 Then
            sNext = I_NEXT: sNextT = S_NEXT & Description
        Else
            sNext = I_ADD: sNextT = S_ADD & Description
        End If
        If Col(1) = "" Then
            sDel = I_DELI: sDelT = ""
        Else
            sDel = I_DEL: sDelT = S_DEL & Description
        End If
    End If
    
    CountControl.Caption = CStr(Index) & "/" & CStr(Total)
    Set PrevControl.Picture = frmMain.Buttons.ListImages(I_PREVI).Picture
    PrevControl.ToolTipText = ""
    Set NextControl.Picture = frmMain.Buttons.ListImages(sNext).Picture
    NextControl.ToolTipText = sNextT
    Set DelControl.Picture = frmMain.Buttons.ListImages(sDel).Picture
    DelControl.ToolTipText = sDelT
End Sub

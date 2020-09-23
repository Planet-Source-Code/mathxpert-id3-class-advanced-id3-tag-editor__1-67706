VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APicDecoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PictureType
    OtherPicType
    FileIcon32x32
    OtherFileIcon
    FrontCover
    BackCover
    LeafletPage
    Media
    LeadArtist
    Artist
    Conductor
    Band
    Composer
    Lyricist
    RecordingLocation
    DuringRecording
    DuringPerformance
    VideoScreenCapture
    BrightColoredFish
    Illustration
    BandLogotype
    PublisherLogotype
End Enum

Public Function DecodeImage(ByVal MFDClass As MultiFrameData, ByVal Index As Long, MIMEType As String, PictureType As PictureType, Pic As StdPicture, ByVal ID3Revision As Byte, Optional OrigPicData As String) As Boolean
    On Error GoTo Err
    Dim bRet As Boolean
    Dim sMIMEType As String
    Dim tMIMEType As String
    Dim lPicType As PictureType
    Dim GPC As GDIPlusCandy
    Dim sPic As StdPicture
    Dim sValue As String
    Dim i As Long
    
    bRet = False
    MIMEType = ""
    PictureType = OtherPicType
    Set Pic = Nothing
    OrigPicData = ""
    
    sValue = MFDClass(Index)
    If ID3Revision > 2 Then i = InStr(sValue, Chr$(0)) Else i = 4
    If i > 0 Then
        sMIMEType = Left$(sValue, i - 1)
        If sMIMEType = "-->" Then GoTo Err ' Skip image URLs
        sValue = Mid$(sValue, i + 1 * Abs(ID3Revision > 2))
        lPicType = Asc(Left$(sValue, 1))
        If lPicType > PublisherLogotype Then lPicType = OtherPicType  ' Fix invalid picture type
        sValue = Mid$(sValue, 2)
        If Left$(sValue, 1) = Chr$(0) Then ' Skip frames with descriptions
            sValue = Mid$(sValue, 2)
            Set GPC = New GDIPlusCandy
            Set sPic = GPC.DataToImage(sValue)
            Set GPC = Nothing
            If sPic Is Nothing Then
                GoTo Err
            Else
                tMIMEType = DetermineImageType(sValue, ID3Revision)
                If sMIMEType <> tMIMEType And tMIMEType <> ImageUnsupported Then
                    sMIMEType = tMIMEType
                End If
                MIMEType = sMIMEType
                If lPicType = FileIcon32x32 Then
                    If sMIMEType <> ImagePNG Or HimetricToPixelsX(sPic.Width) <> 32 Or HimetricToPixelsY(sPic.Height) <> 32 Then
                        lPicType = OtherFileIcon
                    End If
                End If
                PictureType = lPicType
                Set Pic = sPic
                OrigPicData = sValue
                bRet = True
            End If
        End If
    End If
Err:
    DecodeImage = bRet
End Function

' Insert data WITHOUT encoding
Public Sub InsertImageData(MFDClass As MultiFrameData, ByVal Index As Long, ByVal MIMEType As String, ByVal PictureType As PictureType, ByVal Data As String, ByVal ID3Revision As Byte)
    On Error GoTo Err
    
    Dim lPicType As PictureType: lPicType = PictureType
    Dim sMIMEType As String
    Dim tMIMEType As String
    Dim GPC As GDIPlusCandy
    Dim sPic As StdPicture
    
    sMIMEType = MIMEType
    tMIMEType = DetermineImageType(Data, ID3Revision)
    If sMIMEType <> tMIMEType And tMIMEType <> ImageUnsupported Then
        sMIMEType = tMIMEType
    End If
    
    Set GPC = New GDIPlusCandy
    Set sPic = GPC.DataToImage(Data)
    Set GPC = Nothing
    
    If Not sPic Is Nothing Then  ' Do not insert the data if invalid
        If lPicType = FileIcon32x32 Then
            If MIME(sMIMEType, ID3Revision) <> ImagePNG Or HimetricToPixelsX(sPic.Width) <> 32 Or HimetricToPixelsY(sPic.Height) <> 32 Then
                lPicType = OtherFileIcon
            End If
        End If
        MFDClass(Index) = sMIMEType & IIf(ID3Revision > 2, Chr$(0), "") & Chr$(lPicType) & Chr$(0) & Data
    End If
Err:
End Sub

' Insert data WITH encoding
Public Sub InsertImage(MFDClass As MultiFrameData, ByVal Index As Long, ByVal MIMEType As String, ByVal PictureType As PictureType, ByVal Pic As StdPicture, ByVal ID3Revision As Byte)
    On Error GoTo Err
    
    Dim lPicType As PictureType: lPicType = PictureType
    Dim sMIMEType As String
    Dim tMIMEType As String
    Dim GPC As GDIPlusCandy
    Dim sData As String
    
    Set GPC = New GDIPlusCandy
    sData = GPC.ImageToData(Pic, MIME(MIMEType, ID3Revision))
    Set GPC = Nothing
    
    If sData <> "" Then  ' Do not insert the data if invalid
        If lPicType = FileIcon32x32 Then
            If MIME(MIMEType, ID3Revision) <> ImagePNG Or HimetricToPixelsX(Pic.Width) <> 32 Or HimetricToPixelsY(Pic.Height) <> 32 Then
                lPicType = OtherFileIcon
            End If
        End If
        MFDClass(Index) = MIMEType & IIf(ID3Revision > 2, Chr$(0), "") & Chr$(lPicType) & Chr$(0) & sData
    End If
Err:
End Sub

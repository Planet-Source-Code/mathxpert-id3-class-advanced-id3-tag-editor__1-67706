Attribute VB_Name = "modImageType"
Option Explicit

Public Const ImageBMP = "image/bmp"
Public Const ImageGIF = "image/gif"
Public Const ImageJPEG = "image/jpeg"
Public Const ImagePNG = "image/png"
Public Const ImageUnsupported = "[unsupported image type]"

Public Const ImageJPEGOld = "JPG"
Public Const ImagePNGOld = "PNG"

Public Function MIME(ByVal MIMEType As String, ByVal ID3Revision As Byte) As String
    Dim s As String
    If ID3Revision > 2 Then
        s = MIMEType
    Else
        Select Case MIMEType
            Case ImageJPEGOld: s = ImageJPEG
            Case ImagePNGOld: s = ImagePNG
        End Select
    End If
    MIME = s
End Function

Public Function GetIndex(ByVal MIMEType As String, ByVal ID3Revision As Byte) As Long
    If ID3Revision > 2 Then
        Select Case MIMEType
            Case ImageBMP: GetIndex = 0
            Case ImageGIF: GetIndex = 1
            Case ImageJPEG: GetIndex = 2
            Case ImagePNG: GetIndex = 3
            Case Else: GetIndex = -1
        End Select
    Else
        Select Case MIMEType
            Case ImageJPEGOld: GetIndex = 0
            Case ImagePNGOld: GetIndex = 1
            Case Else: GetIndex = -1
        End Select
    End If
End Function

Public Function ImageTypeFromIndex(ByVal Index As Long, ByVal ID3Revision As Byte) As String
    Dim t As String
    If ID3Revision > 2 Then
        Select Case Index
            Case 0: t = ImageBMP
            Case 1: t = ImageGIF
            Case 2: t = ImageJPEG
            Case 3: t = ImagePNG
            Case Else: t = ImageUnsupported
        End Select
    Else
        Select Case Index
            Case 0, 1, 3: t = ImagePNGOld
            Case 2: t = ImageJPEGOld
            Case Else: t = ImageUnsupported
        End Select
    End If
    ImageTypeFromIndex = t
End Function

Public Function DetermineImageType(ByVal sData As String, ByVal ID3Revision As Byte) As String
    On Error GoTo Err
    Dim t As String
    If Left$(sData, 2) = "BM" And Right$(sData, 1) = Chr$(0) Then
        t = ImageTypeFromIndex(0, ID3Revision)
    ElseIf (Left$(sData, 6) = "GIF87a" Or Left$(sData, 6) = "GIF89a") And Right$(sData, 2) = Chr$(0) & ";" Then
        t = ImageTypeFromIndex(1, ID3Revision)
    ElseIf Left$(sData, 12) = "ÿØÿà" & Chr$(0) & Chr$(&H10&) & "JFIF" & Chr$(0) & Chr$(1) And Right$(sData, 2) = "ÿÙ" Then
        t = ImageTypeFromIndex(2, ID3Revision)
    ElseIf Left$(sData, 16) = "‰PNG" & vbCrLf & Chr$(&H1A&) & vbLf & String$(3, 0) & vbCr & "IHDR" And Right$(sData, 12) = String$(4, 0) & "IEND®B`‚" Then
        t = ImageTypeFromIndex(3, ID3Revision)
    Else
Err:
        t = ImageUnsupported
    End If
    On Error GoTo 0
    
    DetermineImageType = t
End Function

Attribute VB_Name = "modWMRetrieval"
Option Explicit

Private Const META_TAB = "    "

Private Const META_TAB_1 = META_TAB
Private Const META_TAB_2 = META_TAB_1 & META_TAB
Private Const META_TAB_3 = META_TAB_2 & META_TAB
Private Const META_TAB_4 = META_TAB_3 & META_TAB

Public Const ART_HOST = "services.windowsmedia.com"
Public Const ART_PATH = "/cover/"

Public dBitRate As Double
Public ArtParam As String

Private Function SplitIntoWords(ByVal Expression As String) As String()
    Dim s() As String
    Dim t As String
    Dim t_0 As String
    Dim x As String
    Dim i As Long
    Dim j As Long
    
    s = Split(Expression, " ")
    
    For i = LBound(s) To UBound(s)
        t = s(i)
        t_0 = ""
        For j = 1 To Len(t)
            x = Mid$(t, j, 1)
            Select Case Asc(x)
                Case 45, 47 To 57, 65 To 90, 97 To 122, &HC0 To &HD6, &HD8 To &HF6, &HF8 To &HFF: t_0 = t_0 & x
            End Select
        Next
        s(i) = t_0
    Next
    
    SplitIntoWords = s
End Function

Private Function WordTags(ByVal Tabs As String, Words() As String) As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim s As String
    Dim subsplit() As String
    Dim subsubsplit() As String
    
    For i = LBound(Words) To UBound(Words)
        subsplit = Split(Words(i), "-")
        For j = LBound(subsplit) To UBound(subsplit)
            subsubsplit = Split(subsplit(j), "/")
            For k = LBound(subsubsplit) To UBound(subsubsplit)
                If subsubsplit(k) <> "" Then
                    s = s & Tabs & "<word>" & ANSItoUTF8(subsubsplit(k)) & "</word>" & vbCrLf
                End If
            Next
        Next
    Next
    
    WordTags = s
End Function

Public Function WGetTagData(ByVal Data As String, ByVal TagName As String, Optional ByVal ConvertHTML As Boolean = False) As String
    Dim i As Long
    Dim j As Long
    Dim td As String
    
    i = InStr(Data, "<" & TagName & ">")
    If i > 0 Then
        j = InStr(i + Len(TagName) + 2, Data, "</" & TagName & ">")
        If j > 0 Then
            td = Mid$(Data, i + Len(TagName) + 2, j - i - Len(TagName) - 2)
            If ConvertHTML Then td = ReplaceHTML(UTF8toANSI(td))
            WGetTagData = td
        End If
    End If
End Function

Public Function WGetTagData2(ByVal Data As String, ByVal TagName As String, ByVal TagName2 As String, ByVal Iterations As Long, ReachedEnd As Boolean, Optional ByVal ConvertHTML As Boolean = False)
    Dim i0 As Long
    Dim i As Long
    Dim j As Long
    Dim Snip As String
    Dim td As String
    
    For i0 = 1 To Iterations
        i = InStr(j + 1, Data, "<" & TagName & ">")
        If i = 0 Then
            ReachedEnd = True
            Exit Function
        End If
        
        j = InStr(i + Len(TagName) + 2, Data, "</" & TagName & ">")
        If j = 0 Then
            ReachedEnd = True
            Exit Function
        End If
    Next
    ReachedEnd = False
    
    Snip = Mid$(Data, i + Len(TagName) + 2, j - i - Len(TagName) - 2)
    i = InStr(Snip, "<" & TagName2 & ">")
    If i > 0 Then
        j = InStr(i + Len(TagName2) + 2, Snip, "</" & TagName2 & ">")
        If j > 0 Then
            td = Mid$(Snip, i + Len(TagName2) + 2, j - i - Len(TagName2) - 2)
            If ConvertHTML Then td = ReplaceHTML(UTF8toANSI(td))
            WGetTagData2 = td
        End If
    End If
End Function

Public Sub ConnectToWM(Socket As Socket)
    Socket.RemoteHost = "info.music.metaservices.microsoft.com"
    Socket.RemotePort = 80
    Socket.Connect
End Sub

Public Sub SendDataToWM(Socket As Socket, ByVal FileName As String, ByVal Title As String, ByVal Artist As String, ByVal Album As String, ByVal Band As String, ByVal TrackNumber As String, ByVal Duration As Double, ByVal BitRate As Double)
    Dim sMetaData As String
    Dim i As Long
    Dim stHex As String
    Dim sTitle As String
    Dim sTitleWords As String
    Dim sArtist As String
    Dim sArtistWords As String
    Dim sAlbum As String
    Dim sAlbumWords As String
    Dim sBand As String
    Dim sBandWords As String
    Dim bTitleIsFile As Boolean
    Dim sTrackNumber As String
    
    If Title = "" Then
        bTitleIsFile = True
        If Len(FileName) > 4 Then
            If LCase$(Right$(FileName, 4)) = ".mp3" Then
                sTitle = Left$(FileName, Len(FileName) - 4)
            Else
                sTitle = FileName
            End If
        Else
            sTitle = FileName
        End If
    Else
        sTitle = Title
    End If
    sTitleWords = WordTags(META_TAB_4, SplitIntoWords(sTitle))
    sTitle = ANSItoUTF8(ToHTML(sTitle, True))
    
    sArtist = Replace(Artist, "/", "; ")
    sArtistWords = WordTags(META_TAB_4, SplitIntoWords(sArtist))
    sArtist = ANSItoUTF8(ToHTML(sArtist, True))
    
    sAlbum = Album
    sAlbumWords = WordTags(META_TAB_4, SplitIntoWords(sAlbum))
    sAlbum = ANSItoUTF8(ToHTML(sAlbum, True))
    
    If Band = "" Then
        sBand = sArtist
        sBandWords = sArtistWords
    Else
        sBand = Band
        i = InStr(sBand, "/")
        If i > 0 Then sBand = Left$(sBand, i - 1)
        sBandWords = WordTags(META_TAB_4, SplitIntoWords(sBand))
        sBand = ANSItoUTF8(ToHTML(sBand, True))
    End If
    
    If IsNumeric(TrackNumber) Then sTrackNumber = CStr(Fix(CDbl(TrackNumber)))
    
    stHex = GenerateHexString(8) & "-" & GenerateHexString(4) & "-" & GenerateHexString(4) & "-" & GenerateHexString(4) & "-" & GenerateHexString(12)
    
    sMetaData = "<METADATA>" & vbCrLf & _
                META_TAB_1 & "<MDQ-CD>" & vbCrLf & _
                 META_TAB_2 & "<mdqRequestID>" & stHex & "</mdqRequestID>" & IIf(sAlbum = "", "", vbCrLf & _
                 META_TAB_2 & "<album>" & IIf(sAlbum = "", "", vbCrLf & _
                  META_TAB_3 & "<title>" & vbCrLf & _
                   META_TAB_4 & "<text>" & sAlbum & "</text>" & vbCrLf & _
                   sAlbumWords & _
                  META_TAB_3 & "</title>") & IIf(sBand = "", "", vbCrLf & _
                  META_TAB_3 & "<artist>" & vbCrLf & _
                   META_TAB_4 & "<text>" & sBand & "</text>" & vbCrLf & _
                   sBandWords & _
                  META_TAB_3 & "</artist>") & vbCrLf & _
                 META_TAB_2 & "</album>") & vbCrLf & _
                 META_TAB_2 & "<track>" & vbCrLf & _
                  META_TAB_3 & "<title>" & vbCrLf & _
                   META_TAB_4 & "<text>" & sTitle & "</text>" & vbCrLf & _
                   sTitleWords & IIf(bTitleIsFile, _
                   "<TitleIsFileName>1</TitleIsFileName>" & vbCrLf, "") & _
                  META_TAB_3 & "</title>" & IIf(sArtist = "", "", vbCrLf & _
                  META_TAB_3 & "<artist>" & vbCrLf & _
                   META_TAB_4 & "<text>" & sArtist & "</text>" & vbCrLf & _
                   sArtistWords & _
                  META_TAB_3 & "</artist>") & IIf(TrackNumber = "", "", vbCrLf & _
                  META_TAB_3 & "<trackNumber>" & TrackNumber & "</trackNumber>") & vbCrLf & _
                  META_TAB_3 & "<filename>" & ANSItoUTF8(ToHTML(FileName, True)) & "</filename>"
    sMetaData = sMetaData & IIf(Duration <= 0, "", vbCrLf & META_TAB_3 & "<trackDuration>" & CStr(Fix(Duration * 1000)) & "</trackDuration>") & IIf(BitRate <= 0, "", vbCrLf & _
                  META_TAB_3 & "<bitrate>" & CStr(Fix(BitRate)) & "</bitrate>") & vbCrLf & _
                  META_TAB_3 & "<drmProtected>0</drmProtected>" & vbCrLf & _
                  META_TAB_3 & "<trackRequestID>" & IIf(sAlbum = "", "1", "0") & "</trackRequestID>" & vbCrLf & _
                 META_TAB_2 & "</track>" & vbCrLf & _
                META_TAB_1 & "</MDQ-CD>" & vbCrLf & _
                "</METADATA>"
    
    Socket.SendData "POST /cdinfo/getmdrcd.aspx?Partner=&locale=409&geoid=f4&version=10.0.0.4036&userlocale=409&requestID=" & stHex & " HTTP/1.0" & vbCrLf & _
                    "Accept: */*" & vbCrLf & _
                    "User-Agent: Windows-Media-Player/10.00.00.4036" & vbCrLf & _
                    "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
                    "Host: info.music.metaservices.microsoft.com" & vbCrLf & _
                    "Content-Length: " & Len(sMetaData) & vbCrLf & _
                    "Connection: Keep-Alive" & vbCrLf & vbCrLf
    Socket.SendData sMetaData
End Sub

Public Sub GetAlbumArt(Socket As Socket, ByVal Host As String, ByVal PathBeginning As String, ByVal Pathend As String, Optional ByVal AsITunes As Boolean = False)
    Dim UserAgent As String
    If AsITunes Then
        UserAgent = "iTunes/7.0.2"
    Else
        UserAgent = "Windows-Media-Player/10.00.00.4036"
    End If
    Socket.SendData "GET " & PathBeginning & Pathend & " HTTP/1.0" & vbCrLf & _
                    "Accept: */*" & vbCrLf & _
                    "User-Agent: " & UserAgent & vbCrLf & _
                    "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
                    "Host: " & Host & vbCrLf & _
                    "Connection: Keep-Alive" & vbCrLf & vbCrLf
End Sub

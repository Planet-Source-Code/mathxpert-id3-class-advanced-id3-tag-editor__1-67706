Attribute VB_Name = "modITunesRetrieval"
Option Explicit

Public ReceivedXML As String
Public bRet As Boolean
Public bLaunchArtDL As Boolean

Public bConnect As Boolean
Public dDuration As Double
Public bItunes As Boolean

Public ArtHost As String
Public ArtPort As Long
Public ArtPath As String

Public Sub AnalyzeURL(ByVal URL As String, Host As String, Port As Long, Path As String)
    Dim i As Long
    Dim j As Long
    Dim sURL As String
    
    sURL = URL
    i = InStr(sURL, "://")
    If i > 0 Then sURL = Mid$(sURL, i + 3)
    
    i = InStr(sURL, "/")
    j = InStr(sURL, ":")
    
    If i > 0 Then
        If j < i And j > 0 Then
            Host = Left$(sURL, j - 1)
            Port = Mid(sURL, j + 1, i - j - 1)
        Else
            Host = Left$(sURL, i - 1)
            Port = 80
        End If
        Path = Mid$(sURL, i)
    Else
        If j > 0 Then
            Host = Left$(sURL, j - 1)
            Port = Mid(sURL, j + 1)
        Else
            Host = sURL
            Port = 80
        End If
        Path = "/"
    End If
End Sub

Public Function HTTPCode(ByVal sData As String, Optional Description As String) As Long
    On Error Resume Next
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Description = ""
    If Left$(sData, 5) = "HTTP/" Then
        i = InStr(sData, " ")
        j = InStr(i + 1, sData, " ")
        k = InStr(j + 1, sData, vbCrLf)
        
        If i > 0 And j > 0 Then
            HTTPCode = CLng(Mid$(sData, i + 1, j - i - 1))
            
            If k > 0 Then
                Description = Mid$(sData, j + 1, k - j - 1)
            End If
        End If
    End If
End Function

Public Function GetTagData(ByVal Data As String, ByVal TagName As String, ByVal vType As VbVarType) As String
    Dim i As Long
    Dim j As Long
    Dim Snip As String
    Dim sType As String
    
    i = InStr(Data, "<key>" & TagName & "</key>")
    If i > 0 Then
        Snip = Mid$(Data, i + Len(TagName) + 11)
        
        If vType = vbString Then
            sType = "string"
        ElseIf vType = vbInteger Then
            sType = "integer"
        End If
        
        If Left$(Snip, 2 + Len(sType)) = "<" & sType & ">" Then
            j = InStr(Len(sType) + 2, Snip, "</" & sType & ">")
            If j > 0 Then
                GetTagData = ReplaceHTML(UTF8toANSI(Mid$(Snip, Len(sType) + 3, j - Len(sType) - 3)))
            End If
        End If
    End If
End Function

Public Function GenerateHexString(ByVal Length As Long) As String
    Dim i As Long
    Dim r As Long
    Dim s As String
    
    For i = 1 To Length
        Randomize ' Make sure we don't get the "sameness"
        r = Int(Rnd * 16) ' Generate random value between 0 and 15  ===>  [Random Integer] = Int(Rnd * (Max - Min + 1)) + Min
        Randomize ' Make sure we randomized the seed very well
        
        Select Case r
            Case 0 To 9
                s = s & CStr(r)
            Case Else
                s = s & Chr$(55 + r)
        End Select
    Next
    
    GenerateHexString = s
End Function

Public Sub ConnectToItunes(Socket As Socket)
    Socket.RemoteHost = "phobos.apple.com"
    Socket.RemotePort = 80
    Socket.Connect
End Sub

Public Sub SendDataToItunes(Socket As Socket, ByVal FileName As String, ByVal Title As String, ByVal Artist As String, ByVal Album As String, ByVal Genre As String, ByVal Composer As String, ByVal Band As String, ByVal Year As String, ByVal TrackNumber As String, ByVal Duration As Double)
    Dim sXML As String
    Dim stHex As String
    Dim sTitle As String
    Dim sYear As String
    Dim sTrackNumber As String
    
    If Title = "" And Artist = "" And Album = "" And Composer = "" And Band = "" And Year = "" And TrackNumber = "" Then
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
    
    If IsNumeric(Year) Then sYear = CStr(Fix(CDbl(Year)))
    If IsNumeric(TrackNumber) Then sTrackNumber = CStr(Fix(CDbl(TrackNumber)))
    
    stHex = GenerateHexString(32)
    
    sXML = "--" & stHex & vbCrLf & _
           "Content-Disposition: form-data; name=""trackInfo""" & vbCrLf & vbCrLf & _
           "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
           "<!DOCTYPE plist PUBLIC ""-//Apple Computer//DTD PLIST 1.0//EN"" ""http://www.apple.com/DTDs/PropertyList-1.0.dtd"">" & vbCrLf & _
           "<plist version=""1.0"">" & vbCrLf & _
           "<dict>" & vbCrLf & _
           vbTab & "<key>viewField</key><string>songName</string>" & vbCrLf & _
           vbTab & "<key>trackInfo</key>" & vbCrLf & _
           vbTab & "<dict>" & IIf(Album = "", "", vbCrLf & _
           vbTab & vbTab & "<key>playlistName</key><string>" & ANSItoUTF8(ToHTML(Album)) & "</string>") & IIf(sTitle = "", "", vbCrLf & _
           vbTab & vbTab & "<key>songName</key><string>" & ANSItoUTF8(ToHTML(sTitle)) & "</string>") & IIf(Artist = "", "", vbCrLf & _
           vbTab & vbTab & "<key>artistName</key><string>" & ANSItoUTF8(ToHTML(Artist)) & "</string>") & IIf(Band = "", "", vbCrLf & _
           vbTab & vbTab & "<key>playlistArtistName</key><string>" & ANSItoUTF8(ToHTML(Band)) & "</string>") & IIf(Genre = "", "", vbCrLf & _
           vbTab & vbTab & "<key>genre</key><string>" & ANSItoUTF8(ToHTML(Genre)) & "</string>") & IIf(Composer = "", "", vbCrLf & _
           vbTab & vbTab & "<key>composerName</key><string>" & ANSItoUTF8(ToHTML(Composer)) & "</string>") & IIf(sYear = "", "", vbCrLf & _
           vbTab & vbTab & "<key>year</key><integer>" & sYear & "</integer>") & IIf(sTrackNumber = "", "", vbCrLf & _
           vbTab & vbTab & "<key>trackNumber</key><integer>" & sTrackNumber & "</integer>") & IIf(Duration <= 0, "", vbCrLf & _
           vbTab & vbTab & "<key>duration</key><integer>" & CStr(Fix(Duration * 1000)) & "</integer>") & vbCrLf & _
           vbTab & "</dict>" & vbCrLf & _
           "</dict>" & vbCrLf & _
           "</plist>" & vbCrLf & vbCrLf & _
           "--" & stHex & "--" & vbCrLf

    Socket.SendData "POST /WebObjects/MZSearch.woa/wa/DirectAction/libraryLink HTTP/1.1" & vbCrLf & _
                    "Content-Length: " & Len(sXML) & vbCrLf & _
                    "Accept-Language: en-us, en;q=0.50" & vbCrLf & _
                    "X-Apple-Tz: -21600" & vbCrLf & _
                    "Cookie: s_vi=[CS]v1|44C9763800007C6C-A000C56000061E7[CE]" & vbCrLf & _
                    "Content-Type: multipart/form-data; boundary=" & stHex & vbCrLf & _
                    "User-Agent: iTunes/7.0.2" & vbCrLf & _
                    "X-Apple-Validation: " & GenerateHexString(8) & "-" & GenerateHexString(32) & vbCrLf & _
                    "Accept-Encoding: gzip, x-aes-cbc" & vbCrLf & _
                    "X-Apple-Store-Front: 143441" & vbCrLf & _
                    "Host: phobos.apple.com" & vbCrLf & _
                    "Cache-Control: no-cache" & vbCrLf & vbCrLf
    Socket.SendData sXML
End Sub

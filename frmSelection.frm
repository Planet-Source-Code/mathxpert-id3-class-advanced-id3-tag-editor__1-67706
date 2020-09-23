VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelection 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5160
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4268
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2948
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Artist"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Album"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Genre"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Track No."
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Select an item that best fits the file:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AudioURL As Collection
Dim DiscCount As Collection
Dim DiscNumber As Collection
Dim Artist As Collection
Dim GenreID As Collection
Dim Genre As Collection
Dim SongName As Collection
Dim PlaylistName As Collection
Dim ReleaseYear As Collection
Dim TrackCount As Collection
Dim TrackNumber As Collection
Dim Year As Collection
Dim Duration As Collection
Dim Composer As Collection
Dim Conductor As Collection
Dim Band As Collection
Dim Label As Collection
Dim ReqID As Collection

Private Sub Command1_Click()
    Dim idx As Long
    Dim BlankWCOM As New MultiFrameData
    Dim BlankWOAR As New MultiFrameData
    Dim BlankAPIC As New MultiFrameData
    idx = ListView1.SelectedItem.Index
    
    With frmMain
        .txtTitle = SongName(idx)
        .txtArtist = Artist(idx)
        .txtAlbum = PlaylistName(idx)
        .cmbGenre = Genre(idx)
        .txtTrackNumber = TrackNumber(idx)
        If TrackCount.Count = 0 Then
            .txtTracksTotal = ""
        Else
            .txtTracksTotal = TrackCount(idx)
        End If
        .txtYear = Year(idx)
        .txtComments = ""
        .txtLyrics = ""
        .txtComposer = Composer(idx)
        .txtBand = Band(idx)
        If Conductor.Count = 0 Then
            .txtConductor = ""
        Else
            .txtConductor = Conductor(idx)
        End If
        .txtInterpretedBy = ""
        .txtLyricist = ""
        .txtOriginalArtist = ""
        .txtOriginalAlbum = ""
        .txtOriginalFileName = ""
        .txtOriginalLyricist = ""
        If ReleaseYear.Count = 0 Then
            .txtOriginalReleaseYear = ""
        Else
            .txtOriginalReleaseYear = ReleaseYear(idx)
        End If
        .txtCopyright = ""
        .txtFileOwner = ""
        If Label.Count = 0 Then
            .txtPublisher = ""
        Else
            .txtPublisher = Label(idx)
        End If
        .txtInternetRadioStationName = ""
        .txtInternetRadioStationOwner = ""
        .txtISRC = ""
        .txtLanguages = ""
        LoadMultiData .txtCommercialInfo, BlankWCOM, S_CURL, .countCommercialInfo, .prevCommercialInfo, .nextCommercialInfo, .delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
        .txtCopyrightInfo = ""
        If AudioURL.Count = 0 Then
            .txtAudioURL = ""
        Else
            .txtAudioURL = AudioURL(idx)
        End If
        LoadMultiData .txtArtistURL, BlankWOAR, S_AURL, .countArtistURL, .prevArtistURL, .nextArtistURL, .delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
        .txtAudioSourceURL = ""
        .txtInternetRadioStationURL = ""
        .txtPaymentURL = ""
        .txtPublisherURL = ""
        .txtEncodedBy = ""
        .txtBPM = ""
        .cmbKey = ""
        If DiscNumber.Count = 0 Then
            .txtDiscNumber = ""
            .txtDiscsTotal = ""
        Else
            .txtDiscNumber = DiscNumber(idx)
            .txtDiscsTotal = DiscCount(idx)
        End If
        If ArtParam = "" And ArtHost = "" And ArtPort = 80 And ArtPath = "" Then
            LoadMultiData .picArt, BlankAPIC, S_APIC, .countArt, .prevArt, .nextArt, .delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
        Else
            bLaunchArtDL = True
        End If
    End With
    
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Function AnalyzeName(ByVal name As String) As String
    Dim n As String
    Dim s As String
    Dim i As Long
    Dim t As String
    Dim a As Integer
    Dim j As Long
    
    n = LCase$(name)
    
    If Len(n) = 3 Then
        If n = "and" Then n = "&"
    ElseIf Len(n) > 3 Then
        If Left$(n, 4) = "and " Then
            n = "&" & Mid$(n, 4)
        ElseIf Right$(n, 4) = " and" Then
            n = Left$(n, Len(n) - 3) & "&"
        End If
    End If
    
    Do
        i = InStr(i + 1, n, "and")
        If i > 0 Then
            If Mid$(n, i - 1, 1) = " " And Mid$(n, i + 3, 1) = " " Then
                n = Left$(n, i - 1) & "&" & Mid$(n, i + 3)
            End If
        Else
            Exit Do
        End If
    Loop
    
    If Left$(n, 3) = "the" Then
        If Len(n) = 3 Then
            j = 1
            GoTo Chop
        Else
            Select Case Asc(Mid$(n, 4, 1))
                Case 48 To 57, 97 To 122, &HC0 To &HD6, &HD8 To &HF6, &HF8 To &HFF
                Case Else
Chop:
                    n = Mid$(n, 5 - j)
            End Select
        End If
    End If
    
    For i = 1 To Len(n)
        t = Mid$(n, i, 1)
        a = Asc(t)
        Select Case a
            Case 65 To 90: s = s & Chr$(a + 32)
            Case 38, 48 To 57, 97 To 122: s = s & t
            Case &HDF: s = s & "ss"
            Case &HC0 To &HC5, &HE0 To &HE5: s = s & "a"
            Case &HC6, &HE6: s = s & "ae"
            Case &HC7, &HE7: s = s & "c"
            Case &HC8 To &HCB, &HE8 To &HEB: s = s & "e"
            Case &HCC To &HCF, &HEC To &HEF: s = s & "i"
            Case &HD0, &HF0: s = s & "d"
            Case &HD1, &HF1: s = s & "n"
            Case &HD2 To &HD6, &HD8, &HF2 To &HF6, &HF8: s = s & "o"
            Case &HD9 To &HDC, &HF9 To &HFC: s = s & "u"
            Case &HDD, &HFD, &HFF: s = s & "y"
            Case &HDE, &HFE: s = s & "th"
        End Select
    Next
    
    AnalyzeName = s
End Function

Private Sub Form_Load()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long
    Dim Snip As String
    Dim SubSnip As String
    Dim sRel As String
    Dim bSelected As Boolean
    Dim bReachedEnd As Boolean
    Dim Iter As Long
    
    Dim sAudioURL As String
    Dim sDiscCount As String
    Dim sDiscNumber As String
    Dim sArtist As String
    Dim sGenreID As String
    Dim sGenre As String
    Dim sSongName As String
    Dim sPlaylistName As String
    Dim sReleaseYear As String
    Dim sTrackCount As String
    Dim sTrackNumber As String
    Dim sYear As String
    Dim sDuration As String
    Dim sComposer As String
    Dim sConductor As String
    Dim sBand As String
    Dim sLabel As String
    Dim sReqID As String
    Dim sURL As String
    
    bLaunchArtDL = False
    bSelected = False
    Command1.Enabled = False
    
    Set AudioURL = New Collection
    Set DiscCount = New Collection
    Set DiscNumber = New Collection
    Set Artist = New Collection
    Set GenreID = New Collection
    Set Genre = New Collection
    Set SongName = New Collection
    Set PlaylistName = New Collection
    Set ReleaseYear = New Collection
    Set TrackCount = New Collection
    Set TrackNumber = New Collection
    Set Year = New Collection
    Set Duration = New Collection
    Set Composer = New Collection
    Set Conductor = New Collection
    Set Band = New Collection
    Set Label = New Collection
    Set ReqID = New Collection
    
    Left = frmMain.Left + (frmMain.Width - Width) / 2
    Top = frmMain.Top + (frmMain.Height - Height) / 2
    
    ArtParam = ""
    ArtHost = ""
    ArtPort = 80
    ArtPath = ""
    sURL = ""
    
    If bItunes Then
        Caption = "iTunes Store Data Selection"
        
        With ListView1
            .ColumnHeaders.Add Text:="Tracks Total"
            .ColumnHeaders(6).Width = 1100
            .ColumnHeaders(6).Alignment = lvwColumnRight
            
            .ColumnHeaders.Add Text:="Year"
            .ColumnHeaders(7).Width = 600
            .ColumnHeaders(7).Alignment = lvwColumnRight
            
            .ColumnHeaders.Add Text:="Duration"
            .ColumnHeaders(8).Width = 800
            .ColumnHeaders(8).Alignment = lvwColumnRight
        End With
        
Beginning:
        i = InStr(i + 1, ReceivedXML, "<GotoURL ")
        If i > 0 Then
            j = InStr(i + 9, ReceivedXML, "</GotoURL>")
            If j > 0 Then
                k = InStr(i + 9, ReceivedXML, "<PictureView ")
                If k > 0 And k < j Then
                    l = InStr(k + 13, ReceivedXML, "url=""")
                    If l > 0 And l < j Then
                        m = InStr(l + 5, ReceivedXML, """")
                        If m > 0 And m < j Then
                            sURL = ReplaceHTML(UTF8toANSI(Mid$(ReceivedXML, l + 5, m - l - 5)))
                        Else
                            i = l + 4
                            GoTo Beginning
                        End If
                    Else
                        i = k + 12
                        GoTo Beginning
                    End If
                Else
                    i = i + 8
                    GoTo Beginning
                End If
            End If
        End If
        
        If sURL <> "" Then
            AnalyzeURL sURL, ArtHost, ArtPort, ArtPath
        End If
        
        i = InStr(ReceivedXML, "<TrackList>")
        If i > 0 Then
            j = InStr(i + 11, ReceivedXML, "</TrackList>")
            If j > 0 Then Snip = Mid$(ReceivedXML, i + 11, j - i - 11)
        End If
        
        If i > 0 And j > 0 Then
            i = InStr(Snip, "<dict>")
            If i > 0 Then
                j = InStrRev(Snip, "</dict>")
                If j > 0 Then
                    Snip = Mid$(Snip, i + 6, j - i - 6)
                    i = 0
                    Do
                        i = InStr(i + 1, Snip, "<dict>")
                        If i > 0 Then
                            j = InStr(i + 6, Snip, "</dict>")
                            If j > 0 Then
                                SubSnip = Mid$(Snip, i + 6, j - i - 6)
                                
                                sAudioURL = GetTagData(SubSnip, "previewURL", vbString)
                                sDiscCount = GetTagData(SubSnip, "discCount", vbInteger)
                                sDiscNumber = GetTagData(SubSnip, "discNumber", vbInteger)
                                sArtist = GetTagData(SubSnip, "artistName", vbString)
                                sGenreID = GetTagData(SubSnip, "genreID", vbInteger)
                                sGenre = GetTagData(SubSnip, "genre", vbString)
                                sSongName = GetTagData(SubSnip, "songName", vbString)
                                sPlaylistName = GetTagData(SubSnip, "playlistName", vbString)
                                
                                sRel = GetTagData(SubSnip, "releaseDate", vbString)
                                k = InStr(sRel, "-")
                                If k > 0 Then
                                    sRel = Left$(sRel, k - 1)
                                End If
                                sReleaseYear = sRel
                                
                                sTrackCount = GetTagData(SubSnip, "trackCount", vbInteger)
                                sTrackNumber = GetTagData(SubSnip, "trackNumber", vbInteger)
                                sYear = GetTagData(SubSnip, "year", vbInteger)
                                sDuration = GetTagData(SubSnip, "duration", vbInteger)
                                sComposer = GetTagData(SubSnip, "composerName", vbString)
                                sBand = GetTagData(SubSnip, "playlistArtistName", vbString)
                                
                                If sAudioURL <> "" Or _
                                   sDiscCount <> "" Or _
                                   sDiscNumber <> "" Or _
                                   sArtist <> "" Or _
                                   sGenreID <> "" Or _
                                   sGenre <> "" Or _
                                   sSongName <> "" Or _
                                   sPlaylistName <> "" Or _
                                   sReleaseYear <> "" Or _
                                   sTrackCount <> "" Or _
                                   sTrackNumber <> "" Or _
                                   sYear <> "" Or _
                                   sDuration <> "" Or _
                                   sComposer <> "" Or _
                                   sBand <> "" Then
                                   
                                    AudioURL.Add sAudioURL
                                    DiscCount.Add sDiscCount
                                    DiscNumber.Add sDiscNumber
                                    Artist.Add sArtist
                                    GenreID.Add sGenreID
                                    Genre.Add sGenre
                                    SongName.Add sSongName
                                    PlaylistName.Add sPlaylistName
                                    ReleaseYear.Add sReleaseYear
                                    TrackCount.Add sTrackCount
                                    TrackNumber.Add sTrackNumber
                                    Year.Add sYear
                                    Duration.Add sDuration
                                    Composer.Add sComposer
                                    Band.Add sBand
                                End If
                            Else
                                Exit Do
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                End If
                
                For i = 1 To SongName.Count
                    ListView1.ListItems.Add Text:=SongName(i)
                    ListView1.ListItems(i).SubItems(1) = Artist(i)
                    ListView1.ListItems(i).SubItems(2) = PlaylistName(i)
                    ListView1.ListItems(i).SubItems(3) = Genre(i)
                    ListView1.ListItems(i).SubItems(4) = TrackNumber(i)
                    ListView1.ListItems(i).SubItems(5) = TrackCount(i)
                    ListView1.ListItems(i).SubItems(6) = Year(i)
                    
                    If Duration(i) = "" Then
                        ListView1.ListItems(i).SubItems(7) = ""
                    Else
                        ListView1.ListItems(i).SubItems(7) = CStr(Duration(i) \ 1000 \ 60) & ":" & Format(Duration(i) \ 1000 Mod 60, "00")
                    End If
                    
                    If Not bSelected Then
                        With frmMain
                            If .txtTitle <> "" Then
                                If InStr(AnalyzeName(ListView1.ListItems(i).Text), AnalyzeName(.txtTitle)) = 1 Or InStr(AnalyzeName(.txtTitle), AnalyzeName(ListView1.ListItems(i).Text)) = 1 Then
                                    bSelected = True
                                    ListView1.ListItems(i).Selected = True
                                    ListView1.ListItems(i).EnsureVisible
                                End If
                            End If
                        End With
                    End If
                Next
            End If
        End If
    Else
        Caption = "Windows Media Metadata Selection"
        
        With ListView1
            .ColumnHeaders.Add Text:="Year"
            .ColumnHeaders(6).Width = 600
            .ColumnHeaders(6).Alignment = lvwColumnRight
        End With
        
        Snip = WGetTagData(ReceivedXML, "METADATA", True)
        If Snip <> "" Then
            sPlaylistName = WGetTagData(Snip, "albumTitle")
            sBand = WGetTagData(Snip, "albumArtist")
            
            sRel = WGetTagData(Snip, "releaseDate")
            k = InStr(sRel, "-")
            If k > 0 Then
                sRel = Left$(sRel, k - 1)
            End If
            sYear = sRel
            
            sLabel = WGetTagData(Snip, "label")
            sGenre = WGetTagData(Snip, "genre")
            ArtParam = WGetTagData(Snip, "largeCoverParams")
            
            Iter = 1
            Do
                sReqID = WGetTagData2(Snip, "track", "trackRequestID", Iter, bReachedEnd)
                If bReachedEnd Then Exit Do
                
                sSongName = WGetTagData2(Snip, "track", "trackTitle", Iter, bReachedEnd)
                sTrackNumber = WGetTagData2(Snip, "track", "trackNumber", Iter, bReachedEnd)
                sArtist = Replace(WGetTagData2(Snip, "track", "trackPerformer", Iter, bReachedEnd), "; ", "/")
                sComposer = Replace(WGetTagData2(Snip, "track", "trackComposer", Iter, bReachedEnd), "; ", "/")
                sConductor = Replace(WGetTagData2(Snip, "track", "trackConductor", Iter, bReachedEnd), "; ", "/")
                
                If sPlaylistName <> "" Or _
                   sBand <> "" Or _
                   sYear <> "" Or _
                   sLabel <> "" Or _
                   sGenre <> "" Or _
                   sSongName <> "" Or _
                   sTrackNumber <> "" Or _
                   sArtist <> "" Or _
                   sComposer <> "" Or _
                   sConductor <> "" Or _
                   sReqID <> "" Then
                    
                    PlaylistName.Add sPlaylistName
                    Band.Add sBand
                    Year.Add sYear
                    Label.Add sLabel
                    Genre.Add sGenre
                    SongName.Add sSongName
                    TrackNumber.Add sTrackNumber
                    Artist.Add sArtist
                    Composer.Add sComposer
                    Conductor.Add sConductor
                    ReqID.Add sReqID
                End If
                
                Iter = Iter + 1
            Loop
            
            For i = 1 To SongName.Count
                ListView1.ListItems.Add Text:=SongName(i)
                ListView1.ListItems(i).SubItems(1) = Artist(i)
                ListView1.ListItems(i).SubItems(2) = PlaylistName(i)
                ListView1.ListItems(i).SubItems(3) = Genre(i)
                ListView1.ListItems(i).SubItems(4) = TrackNumber(i)
                ListView1.ListItems(i).SubItems(5) = Year(i)
                
                If Not bSelected Then
                    If ReqID(i) <> "" Then
                        bSelected = True
                        ListView1.ListItems(i).Selected = True
                        ListView1.ListItems(i).EnsureVisible
                    End If
                End If
            Next
        End If
    End If
    
    If SelectedIndex(ListView1) <> -1 Then
        If Not bSelected Then
            ListView1.SelectedItem.Selected = False
        End If
    End If
    
    If SelectedIndex(ListView1) <> -1 Then _
        Command1.Enabled = True

    If ListView1.ListItems.Count = 0 Then
        bRet = True
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Do Until ListView1.ColumnHeaders.Count = 5
        ListView1.ColumnHeaders.Remove 6
    Loop
    
    Set AudioURL = Nothing
    Set DiscCount = Nothing
    Set DiscNumber = Nothing
    Set Artist = Nothing
    Set GenreID = Nothing
    Set Genre = Nothing
    Set SongName = Nothing
    Set PlaylistName = Nothing
    Set ReleaseYear = Nothing
    Set TrackCount = Nothing
    Set TrackNumber = Nothing
    Set Year = Nothing
    Set Duration = Nothing
    Set Composer = Nothing
    Set Conductor = Nothing
    Set Band = Nothing
    Set Label = Nothing
    Set ReqID = Nothing
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Not Command1.Enabled Then Command1.Enabled = True
End Sub

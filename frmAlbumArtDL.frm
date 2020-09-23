VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlbumArtDL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Downloading Album Art"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   1553
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   135
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   60000
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Downloading album art; please wait..."
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
      Width           =   4455
   End
End
Attribute VB_Name = "frmAlbumArtDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents SK As Socket
Attribute SK.VB_VarHelpID = -1

Dim s As String

Dim ContentLength As Long
Dim ContentStart As Long

Dim bContentLength As Boolean
Dim bStart As Boolean

Dim bClosed As Boolean

Private Sub EmptyAlbumArt()
    Dim BlankAPIC As New MultiFrameData
    SK.CloseSck
    With frmMain
        LoadMultiData .picArt, BlankAPIC, S_APIC, .countArt, .prevArt, .nextArt, .delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
    End With
    Unload Me
End Sub

Private Sub Command1_Click()
    If bConnect Then
        If SK.state <> sckClosed And SK.state <> sckClosing Then
            SK.CloseSck
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Set SK = New Socket
    
    Left = frmMain.Left + (frmMain.Width - Width) / 2
    Top = frmMain.Top + (frmMain.Height - Height) / 2
    
    s = ""
    bContentLength = False
    bStart = False
    bClosed = False
    PB.Value = 0
    If bItunes Then
        SK.RemoteHost = ArtHost
        SK.RemotePort = ArtPort
        SK.Connect
    Else
        SK.RemoteHost = ART_HOST
        SK.RemotePort = 80
        SK.Connect
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SK = Nothing
End Sub

Private Sub SK_CloseSck()
    Dim tData As String
    Dim lCode As Long
    Dim Desc As String
    Dim MIMEType As String
    Dim NewAPIC As MultiFrameData
    SK.CloseSck
    bClosed = True
    
    tData = Mid$(s, ContentStart)
    lCode = HTTPCode(s, Desc)
    If lCode = 200 Then ' OK
        MIMEType = DetermineImageType(tData, ID3Revision)
        If MIMEType = ImageUnsupported Then
            EmptyAlbumArt
        Else
            Set NewAPIC = New MultiFrameData
            NewAPIC.Add MIMEType & IIf(ID3Revision > 2, Chr$(0), "") & Chr$(3) & Chr$(0) & tData
            With frmMain
                LoadMultiData .picArt, NewAPIC, S_APIC, .countArt, .prevArt, .nextArt, .delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
            End With
            Unload Me
        End If
    Else
        EmptyAlbumArt
    End If
End Sub

Private Sub SK_Connect()
    If bItunes Then
        GetAlbumArt SK, ArtHost, ArtPath, "", True
    Else
        GetAlbumArt SK, ART_HOST, ART_PATH, ArtParam
    End If
End Sub

Private Sub SK_ConnectionRequest(ByVal requestID As Long)
    SK.Accept requestID
End Sub

Private Sub SK_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Dim i As Long
    Dim j As Long
    
    SK.GetData sData, vbString
    s = s & sData
    
    If Not bContentLength Then
        i = InStr(1, s, "Content-Length: ", vbTextCompare)
        If i > 0 Then
            j = InStr(i + 16, s, vbCrLf)
            If j > 0 Then
                ContentLength = Mid$(s, i + 16, j - i - 16)
                bContentLength = True
            End If
        End If
    End If
    
    If bStart Then
CheckEnd:
        If bContentLength Then
            PB.Value = CSng(CDbl(Len(s) - ContentStart + 1) / CDbl(ContentLength) * 60000)
            If Len(s) - ContentStart + 1 = ContentLength Then
                If Not bClosed Then SK_CloseSck
            End If
        End If
    Else
        i = InStr(s, vbCrLf & vbCrLf)
        If i > 0 Then
            ContentStart = i + 4
            bStart = True
            GoTo CheckEnd
        End If
    End If
End Sub

Private Sub SK_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    EmptyAlbumArt
End Sub

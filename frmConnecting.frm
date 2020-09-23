VERSION 5.00
Begin VB.Form frmConnecting 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   960
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
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
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
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
Attribute VB_Name = "frmConnecting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSG_RECV = "Receiving data..."

Dim WithEvents SK As Socket
Attribute SK.VB_VarHelpID = -1

Dim s As String
Dim ConnID As Long

Dim sHost As String
Dim sPath As String

Dim ContentLength As Long
Dim ContentEncoding As String
Dim ContentStart As Long

Dim bContentLength As Boolean
Dim bContentEncoding As Boolean
Dim bStart As Boolean

Dim bClosed As Boolean

Dim URL As String

Private Sub Command1_Click()
    If bConnect Then
        If SK.state <> sckClosed And SK.state <> sckClosing Then
            SK.CloseSck
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Cap As String
    
    Left = frmMain.Left + (frmMain.Width - Width) / 2
    Top = frmMain.Top + (frmMain.Height - Height) / 2
    
    Cap = "Connecting to "
    If bItunes Then
        Cap = Cap & "the iTunes Store"
    Else
        Cap = Cap & "Microsoft"
    End If
    Cap = Cap & "..."
    Caption = Cap
    
    If bConnect Then
        Set SK = New Socket
        
        Command1.Caption = "&Cancel"
        s = ""
        ConnID = 1
        sHost = ""
        sPath = ""
        URL = ""
        ReceivedXML = ""
        bContentLength = False
        bContentEncoding = False
        bStart = False
        bClosed = False
        Label1 = "Resolving host..."
        If bItunes Then
            ConnectToItunes SK
        Else
            ConnectToWM SK
        End If
    Else
        Command1.Caption = "&Close"
        Label1 = "Data not found."
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bConnect Then Set SK = Nothing
End Sub

Private Sub SK_CloseSck()
    Dim i As Long
    Dim j As Long
    Dim sPort As Long
    Dim tData As String
    Dim lCode As Long
    Dim Desc As String
    Dim zlib As New cZLIB
    
    SK.CloseSck
    bClosed = True
    
    tData = Mid$(s, ContentStart)
    If InStr(1, ContentEncoding, "gzip", vbTextCompare) > 0 Then zlib.UncompressString tData, Z_AUTO
    
    lCode = HTTPCode(s, Desc)
    If lCode = 200 Then ' OK
        If bItunes Then
            If ConnID = 1 Then
                ConnID = 2
                
                URL = GetTagData(tData, "url", vbString)
                
                If URL <> "" Then
                    URL = ReplaceHTML(URL)
                    AnalyzeURL URL, sHost, sPort, sPath
                    
                    s = ""
                    Label1 = "Resolving host..."
                    bContentLength = False
                    bContentEncoding = False
                    bStart = False
                    bClosed = False
                    SK.RemoteHost = sHost
                    SK.RemotePort = sPort
                    SK.Connect
                Else
                    Label1 = "Information unavailable"
                    Command1.Caption = "&Close"
                End If
            ElseIf ConnID = 2 Then
            
                If GetTagData(tData, "kind", vbString) = "Goto" Then
                
                    URL = GetTagData(tData, "url", vbString)
                    
                    If URL <> "" Then
                        URL = ReplaceHTML(URL)
                        AnalyzeURL URL, sHost, sPort, sPath
                        
                        s = ""
                        Label1 = "Resolving host..."
                        bContentLength = False
                        bContentEncoding = False
                        bStart = False
                        bClosed = False
                        SK.RemoteHost = sHost
                        SK.RemotePort = sPort
                        SK.Connect
                    Else
                        Label1 = "Information unavailable"
                        Command1.Caption = "&Close"
                    End If
                Else
                    ReceivedXML = tData
                    bRet = True
                    Unload Me
                End If
            End If
        Else
            ReceivedXML = tData
            bRet = True
            Unload Me
        End If
    Else
        Label1 = "HTTP " & CStr(lCode) & ": " & Desc
        Command1.Caption = "&Close"
    End If
End Sub

Private Sub SK_Connect()
    Label1 = "Sending data..."
    With frmMain
        If bItunes Then
            If ConnID = 1 Then
                SendDataToItunes SK, .ListView1.SelectedItem.Text, .txtTitle, .txtArtist, .txtAlbum, .cmbGenre, .txtComposer, .txtBand, .txtYear, .txtTrackNumber, dDuration
            Else
                SK.SendData "GET " & sPath & " HTTP/1.1" & vbCrLf & _
                            "Accept-Language: en-us, en;q=0.50" & vbCrLf & _
                            "X-Apple-Tz: -21600" & vbCrLf & _
                            "User-Agent: iTunes/7.0.2" & vbCrLf & _
                            "X-Apple-Validation: " & GenerateHexString(8) & "-" & GenerateHexString(32) & vbCrLf & _
                            "Accept-Encoding: gzip, x-aes-cbc" & vbCrLf & _
                            "X-Apple-Store-Front: 143441" & vbCrLf & _
                            "Host: " & sHost & vbCrLf & vbCrLf
            End If
        Else
            SendDataToWM SK, .ListView1.SelectedItem.Text, .txtTitle, .txtArtist, .txtAlbum, .txtBand, .txtTrackNumber, dDuration, dBitRate
        End If
    End With
End Sub

Private Sub SK_ConnectionRequest(ByVal requestID As Long)
    SK.Accept requestID
End Sub

Private Sub SK_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Dim i As Long
    Dim j As Long
    
    If Label1 <> MSG_RECV Then Label1 = MSG_RECV
    
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
    
    If Not bContentEncoding Then
        i = InStr(1, s, "Content-Encoding: ", vbTextCompare)
        If i > 0 Then
            j = InStr(i + 18, s, vbCrLf)
            If j > 0 Then
                ContentEncoding = Mid$(s, i + 18, j - i - 18)
                bContentEncoding = True
            End If
        End If
    End If
    
    If bStart Then
CheckEnd:
        If bContentLength Then
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
    SK.CloseSck
    Label1 = "ERROR: " & Description
    Command1.Caption = "&Close"
End Sub

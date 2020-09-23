VERSION 5.00
Begin VB.Form frmGZip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Required DLL Not Found"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Install vbzlib1.dll"
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
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   514
      Picture         =   "frmGZip.frx":0000
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "An error occurred while installing vbzlib.dll."
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
      Left            =   874
      TabIndex        =   3
      Top             =   1080
      Width           =   3053
   End
   Begin VB.Label Label1 
      Caption         =   "Advanced MP3 Info Editor did not find the DLL required for decompressing GZip-compressed data."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmGZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ORIG_HEIGHT = 1560
Private Const NEW_HEIGHT = 2160

Private Sub Command1_Click()
    On Error Resume Next
    
    Dim f As Integer
    Dim b() As Byte
    
    f = FreeFile
    b = LoadResData(101, "CUSTOM")
    
    Open GetSpecialFolderLocation(CSIDL_SYSTEM) & "\vbzlib1.dll" For Binary Access Write Shared As #f
        Put #f, , b
    Close #f
    
    If err Then
        Height = NEW_HEIGHT
        Top = frmMain.Top + (frmMain.Height - Height) / 2
        Command1.Caption = "Retry &installing vbzlib1.dll"
    Else
        bRet = True
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Height = ORIG_HEIGHT
    Left = frmMain.Left + (frmMain.Width - Width) / 2
    Top = frmMain.Top + (frmMain.Height - Height) / 2
    Command1.Caption = "&Install vbzlib1.dll"
End Sub

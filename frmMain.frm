VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Advanced MP3 Info Editor"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList Buttons 
      Left            =   1080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0077
            Key             =   "addi"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":00F5
            Key             =   "del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":01E6
            Key             =   "deli"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":02D7
            Key             =   "next"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":03C6
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":04B6
            Key             =   "previ"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9120
      TabIndex        =   2
      Top             =   90
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   7
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0688
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "MP3 ID Panel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   9735
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   9255
         Begin VB.TextBox txtTracksTotal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5640
            TabIndex        =   18
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtLyrics 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   5400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   1440
            Width           =   3855
         End
         Begin VB.TextBox txtYear 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6960
            TabIndex        =   20
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtComments 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   960
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   1440
            Width           =   3615
         End
         Begin VB.TextBox txtTrackNumber 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4800
            TabIndex        =   16
            Top             =   1080
            Width           =   495
         End
         Begin VB.ComboBox cmbGenre 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":075A
            Left            =   600
            List            =   "frmMain.frx":091D
            TabIndex        =   14
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtAlbum 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   12
            Top             =   720
            Width           =   8655
         End
         Begin VB.TextBox txtArtist 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   10
            Top             =   360
            Width           =   8655
         End
         Begin VB.TextBox txtTitle 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   8
            Top             =   0
            Width           =   8655
         End
         Begin VB.Label Label39 
            Caption         =   "of"
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
            Left            =   5400
            TabIndex        =   17
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label8 
            Caption         =   "Lyrics:"
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
            Left            =   4800
            TabIndex        =   23
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Year:"
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
            Left            =   6480
            TabIndex        =   19
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Comments:"
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
            Left            =   0
            TabIndex        =   21
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Track:"
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
            Left            =   4260
            TabIndex        =   15
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Genre:"
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
            Left            =   0
            TabIndex        =   13
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Album:"
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
            Left            =   0
            TabIndex        =   11
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Artist:"
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
            Left            =   0
            TabIndex        =   9
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Title:"
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
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
         Width           =   9255
         Begin VB.VScrollBar VScroll1 
            Height          =   2655
            LargeChange     =   5
            Left            =   9000
            Max             =   29
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   9825
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   9015
            Begin VB.TextBox txtInterpretedBy 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   34
               Top             =   1080
               Width           =   7095
            End
            Begin VB.TextBox txtPublisher 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   52
               Top             =   4320
               Width           =   7095
            End
            Begin VB.TextBox txtDiscsTotal 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6000
               TabIndex        =   88
               Top             =   9480
               Width           =   495
            End
            Begin VB.TextBox txtDiscNumber 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5160
               TabIndex        =   86
               Top             =   9480
               Width           =   495
            End
            Begin VB.TextBox txtInternetRadioStationURL 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   74
               Top             =   7920
               Width           =   7095
            End
            Begin VB.TextBox txtAudioSourceURL 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   72
               Top             =   7560
               Width           =   7095
            End
            Begin VB.TextBox txtArtistURL 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   69
               Top             =   7200
               Width           =   5850
            End
            Begin VB.TextBox txtCopyrightInfo 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   65
               Top             =   6480
               Width           =   7095
            End
            Begin VB.TextBox txtCommercialInfo 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   62
               Top             =   6120
               Width           =   5850
            End
            Begin VB.TextBox txtISRC 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   58
               Top             =   5400
               Width           =   7095
            End
            Begin VB.TextBox txtInternetRadioStationOwner 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   56
               Top             =   5040
               Width           =   7095
            End
            Begin VB.TextBox txtInternetRadioStationName 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   54
               Top             =   4680
               Width           =   7095
            End
            Begin VB.TextBox txtFileOwner 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   50
               Top             =   3960
               Width           =   7095
            End
            Begin VB.TextBox txtConductor 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   32
               Top             =   720
               Width           =   7095
            End
            Begin VB.TextBox txtBand 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   30
               Top             =   360
               Width           =   7095
            End
            Begin VB.TextBox txtPublisherURL 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   78
               Top             =   8640
               Width           =   7095
            End
            Begin VB.TextBox txtComposer 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   28
               Top             =   0
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalArtist 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   38
               Top             =   1800
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalAlbum 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   40
               Top             =   2160
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalFileName 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   42
               Top             =   2520
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalLyricist 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   44
               Top             =   2880
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalReleaseYear 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   46
               Top             =   3240
               Width           =   7095
            End
            Begin VB.TextBox txtCopyright 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   48
               Top             =   3600
               Width           =   7095
            End
            Begin VB.TextBox txtLanguages 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   60
               Top             =   5760
               Width           =   7095
            End
            Begin VB.TextBox txtLyricist 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   36
               Top             =   1440
               Width           =   7095
            End
            Begin VB.TextBox txtBPM 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   480
               TabIndex        =   82
               Top             =   9480
               Width           =   1095
            End
            Begin VB.ComboBox cmbKey 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmMain.frx":0EF2
               Left            =   2880
               List            =   "frmMain.frx":0F62
               TabIndex        =   84
               Top             =   9480
               Width           =   1455
            End
            Begin VB.TextBox txtAudioURL 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   67
               Top             =   6840
               Width           =   7095
            End
            Begin VB.TextBox txtPaymentURL 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   76
               Top             =   8280
               Width           =   7095
            End
            Begin VB.TextBox txtEncodedBy 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   80
               Top             =   9000
               Width           =   7095
            End
            Begin VB.Label countArtistURL 
               Alignment       =   1  'Right Justify
               Caption         =   "0/0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   135
               Left            =   7680
               TabIndex        =   70
               Top             =   7260
               Width           =   570
            End
            Begin VB.Label countCommercialInfo 
               Alignment       =   1  'Right Justify
               Caption         =   "0/0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   135
               Left            =   7680
               TabIndex        =   63
               Top             =   6180
               Width           =   570
            End
            Begin VB.Image delArtistURL 
               Height          =   285
               Left            =   8685
               Picture         =   "frmMain.frx":111A
               ToolTipText     =   "Delete Artist URL"
               Top             =   7200
               Width           =   210
            End
            Begin VB.Image nextArtistURL 
               Height          =   285
               Left            =   8475
               Picture         =   "frmMain.frx":11FB
               ToolTipText     =   "Next Artist URL"
               Top             =   7200
               Width           =   210
            End
            Begin VB.Image prevArtistURL 
               Height          =   285
               Left            =   8265
               Picture         =   "frmMain.frx":12DA
               ToolTipText     =   "Previous Artist URL"
               Top             =   7200
               Width           =   210
            End
            Begin VB.Image delCommercialInfo 
               Height          =   285
               Left            =   8685
               Picture         =   "frmMain.frx":13BA
               ToolTipText     =   "Delete Commercial Info URL"
               Top             =   6120
               Width           =   210
            End
            Begin VB.Image nextCommercialInfo 
               Height          =   285
               Left            =   8475
               Picture         =   "frmMain.frx":149B
               ToolTipText     =   "Next Commercial Info URL"
               Top             =   6120
               Width           =   210
            End
            Begin VB.Image prevCommercialInfo 
               Height          =   285
               Left            =   8265
               Picture         =   "frmMain.frx":157A
               ToolTipText     =   "Previous Commercial Info URL"
               Top             =   6120
               Width           =   210
            End
            Begin VB.Label Label40 
               Caption         =   "Interpreted by:"
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
               Left            =   0
               TabIndex        =   33
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label Label38 
               Caption         =   "Publisher:"
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
               Left            =   0
               TabIndex        =   51
               Top             =   4320
               Width           =   1695
            End
            Begin VB.Label Label37 
               Caption         =   "of"
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
               Left            =   5760
               TabIndex        =   87
               Top             =   9480
               Width           =   255
            End
            Begin VB.Label Label36 
               Caption         =   "Disc:"
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
               Left            =   4680
               TabIndex        =   85
               Top             =   9480
               Width           =   375
            End
            Begin VB.Label Label35 
               Caption         =   "Net Radio Station URL:"
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
               Left            =   0
               TabIndex        =   73
               Top             =   7920
               Width           =   1695
            End
            Begin VB.Label Label34 
               Caption         =   "Audio Source URL:"
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
               Left            =   0
               TabIndex        =   71
               Top             =   7560
               Width           =   1695
            End
            Begin VB.Label Label33 
               Caption         =   "Artist URL:"
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
               Left            =   0
               TabIndex        =   68
               Top             =   7200
               Width           =   1695
            End
            Begin VB.Label Label32 
               Caption         =   "Copyright Info URL:"
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
               Left            =   0
               TabIndex        =   64
               Top             =   6480
               Width           =   1695
            End
            Begin VB.Label Label31 
               Caption         =   "Commercial Info URL:"
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
               Left            =   0
               TabIndex        =   61
               Top             =   6120
               Width           =   1695
            End
            Begin VB.Label Label30 
               Caption         =   "ISRC:"
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
               Left            =   0
               TabIndex        =   57
               Top             =   5400
               Width           =   1695
            End
            Begin VB.Label Label29 
               Caption         =   "Net Radio Stn. Owner:"
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
               Left            =   0
               TabIndex        =   55
               Top             =   5040
               Width           =   1695
            End
            Begin VB.Label Label28 
               Caption         =   "Net Radio Stn. Name:"
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
               Left            =   0
               TabIndex        =   53
               Top             =   4680
               Width           =   1695
            End
            Begin VB.Label Label27 
               Caption         =   "File Owner/Licensee:"
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
               Left            =   0
               TabIndex        =   49
               Top             =   3960
               Width           =   1695
            End
            Begin VB.Label Label26 
               Caption         =   "Conductor:"
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
               Left            =   0
               TabIndex        =   31
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label25 
               Caption         =   "Band/Orchestra:"
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
               Left            =   0
               TabIndex        =   29
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label24 
               Caption         =   "Publisher URL:"
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
               Left            =   0
               TabIndex        =   77
               Top             =   8640
               Width           =   1695
            End
            Begin VB.Label Label9 
               Caption         =   "Composer:"
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
               Left            =   0
               TabIndex        =   27
               Top             =   0
               Width           =   1695
            End
            Begin VB.Label Label10 
               Caption         =   "Original Artist:"
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
               Left            =   0
               TabIndex        =   37
               Top             =   1800
               Width           =   1695
            End
            Begin VB.Label Label11 
               Caption         =   "Original Album:"
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
               Left            =   0
               TabIndex        =   39
               Top             =   2160
               Width           =   1695
            End
            Begin VB.Label Label12 
               Caption         =   "Original Filename:"
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
               Left            =   0
               TabIndex        =   41
               Top             =   2520
               Width           =   1695
            End
            Begin VB.Label Label13 
               Caption         =   "Original Lyricist:"
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
               Left            =   0
               TabIndex        =   43
               Top             =   2880
               Width           =   1695
            End
            Begin VB.Label Label14 
               Caption         =   "Original Release Year:"
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
               Left            =   0
               TabIndex        =   45
               Top             =   3240
               Width           =   1695
            End
            Begin VB.Label Label15 
               Caption         =   "Copyright:"
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
               Left            =   0
               TabIndex        =   47
               Top             =   3600
               Width           =   1695
            End
            Begin VB.Label Label16 
               Caption         =   "Languages:"
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
               Left            =   0
               TabIndex        =   59
               Top             =   5760
               Width           =   1695
            End
            Begin VB.Label Label17 
               Caption         =   "Lyricist:"
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
               Left            =   0
               TabIndex        =   35
               Top             =   1440
               Width           =   1695
            End
            Begin VB.Label Label19 
               Caption         =   "BPM:"
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
               Left            =   0
               TabIndex        =   81
               Top             =   9480
               Width           =   375
            End
            Begin VB.Label Label20 
               Caption         =   "Initial Key:"
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
               Left            =   1920
               TabIndex        =   83
               Top             =   9480
               Width           =   855
            End
            Begin VB.Label Label21 
               Caption         =   "Audio URL:"
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
               Left            =   0
               TabIndex        =   66
               Top             =   6840
               Width           =   1695
            End
            Begin VB.Label Label22 
               Caption         =   "Payment URL:"
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
               Left            =   0
               TabIndex        =   75
               Top             =   8280
               Width           =   1695
            End
            Begin VB.Label Label23 
               Caption         =   "Encoded by:"
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
               Left            =   0
               TabIndex        =   79
               Top             =   9000
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   2
         Left            =   240
         TabIndex        =   94
         Top             =   720
         Visible         =   0   'False
         Width           =   9255
         Begin VB.ComboBox cmbImageType 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":165A
            Left            =   3960
            List            =   "frmMain.frx":166A
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   1920
            Width           =   1275
         End
         Begin VB.PictureBox picArt 
            BackColor       =   &H8000000C&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   3960
            ScaleHeight     =   1755
            ScaleWidth      =   1755
            TabIndex        =   95
            Top             =   0
            Width           =   1815
            Begin VB.Image imgArt 
               Height          =   1755
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1755
            End
            Begin VB.Label lblBrowse 
               Alignment       =   2  'Center
               BackColor       =   &H8000000C&
               BackStyle       =   0  'Transparent
               Caption         =   "Click here to browse..."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   495
               Left            =   0
               TabIndex        =   96
               Top             =   720
               Width           =   1815
            End
         End
         Begin VB.ComboBox cmbPictureType 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":1683
            Left            =   3960
            List            =   "frmMain.frx":16C6
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   2280
            Width           =   2520
         End
         Begin VB.Label countArt 
            Alignment       =   1  'Right Justify
            Caption         =   "0/0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   5265
            TabIndex        =   101
            Top             =   1980
            Width           =   570
         End
         Begin VB.Image delArt 
            Height          =   285
            Left            =   6270
            Picture         =   "frmMain.frx":17F5
            Top             =   1920
            Width           =   210
         End
         Begin VB.Image nextArt 
            Height          =   285
            Left            =   6060
            Picture         =   "frmMain.frx":18D6
            Top             =   1920
            Width           =   210
         End
         Begin VB.Image prevArt 
            Height          =   285
            Left            =   5850
            Picture         =   "frmMain.frx":19B5
            Top             =   1920
            Width           =   210
         End
         Begin VB.Label Label43 
            Caption         =   "Picture type:"
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
            Left            =   2880
            TabIndex        =   99
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label41 
            Caption         =   "Image type:"
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
            Left            =   2880
            TabIndex        =   97
            Top             =   1920
            Width           =   975
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Search Info via &Microsoft"
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
         Left            =   2400
         TabIndex        =   91
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Search Info in iTunes Store"
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
         TabIndex        =   90
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Delete MP3 Tags"
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
         Left            =   8160
         TabIndex        =   93
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Update MP3 Info"
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
         Left            =   6600
         TabIndex        =   92
         Top             =   3600
         Width           =   1455
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5741
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Basic"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Advanced"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Album Art"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Artist"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Album"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Genre"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Track No."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Tracks Total"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Year"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Duration"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Bit Rate"
         Object.Width           =   2064
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Comments"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8640
      TabIndex        =   1
      Top             =   90
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
   Begin VB.Menu mnuArt 
      Caption         =   "ArtMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuArtItem 
         Caption         =   "&Copy"
         Index           =   0
      End
      Begin VB.Menu mnuArtItem 
         Caption         =   "&Paste"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

Private Const SW_SHOW As Long = 5
Private Const CF_BITMAP As Long = 2
Private Const IMAGE_BITMAP As Long = 0
Private Const LR_COPYRETURNORG As Long = &H4

Private Const S_OTHER As String = "Other"

Private Const FILTER_BMP As String = "*.bmp;*.dib"
Private Const FILTER_GIF As String = "*.gif"
Private Const FILTER_JPEG As String = "*.jpeg;*.jpg;*.jpe;*.jfif;*.jfi;*.jif"
Private Const FILTER_PNG As String = "*.png"
Private Const FILTER_SUPPORTED As String = FILTER_BMP & ";" & FILTER_GIF & ";" & FILTER_JPEG & ";" & FILTER_PNG

Private Const MNU_COPY As Long = 0
Private Const MNU_PASTE As Long = 1

Private Const PASTE_TXT_1 As String = "&Paste"
Private Const PASTE_TXT_2 As String = PASTE_TXT_1 & " (this will change the current image)"

Dim myWindowState As Integer
Dim bInitialized As Boolean

Private Function ValidateMenu() As Boolean
    On Error Resume Next
    
    Dim bVal1 As Boolean
    Dim tPic As StdPicture
    Dim bVal2 As Boolean
    Dim bBW As Boolean
    
    If ListView1.ListItems.Count > 0 Then
        bVal1 = imgArt.Visible
        Set tPic = Clipboard.GetData(CF_BITMAP)
        bVal2 = (Not tPic Is Nothing And tPic.handle <> 0)
    End If
    
    ' Apparently, order seems to matter when it comes to hiding certain items
    bBW = (Not bVal1 And bVal2)
    If bBW Then GoTo PasteItem
CopyItem:
    If mnuArtItem(MNU_COPY).Visible <> (Not bVal2 Or bVal1) Then mnuArtItem(MNU_COPY).Visible = (Not bVal2 Or bVal1)
    If bBW Then GoTo ContinueProc
PasteItem:
    If mnuArtItem(MNU_PASTE).Visible <> (Not bVal1 Or bVal2) Then mnuArtItem(MNU_PASTE).Visible = (Not bVal1 Or bVal2)
    If bBW Then GoTo CopyItem
    
ContinueProc:
    If bVal1 And bVal2 Then
        If mnuArtItem(MNU_PASTE).Caption <> PASTE_TXT_2 Then mnuArtItem(MNU_PASTE).Caption = PASTE_TXT_2
    Else
        If mnuArtItem(MNU_PASTE).Caption <> PASTE_TXT_1 Then mnuArtItem(MNU_PASTE).Caption = PASTE_TXT_1
    End If
    
    ValidateMenu = bVal1 Or bVal2
End Function

Private Function FormatGenre(ByVal ID3Class As clsID3, ByVal GenreID As GenreConstants, ByVal Genre As String) As String
    If (GenreID = OtherGenre Or GenreID = Unknown) And Genre <> "" Then
        FormatGenre = Genre
    Else
        FormatGenre = ID3Class.GenreName(GenreID)
    End If
End Function

Private Function FormatTime(ByVal TimeVal As Double, Optional ByVal StoreTime As Boolean = False) As String
    On Error Resume Next
    
    Dim tv As Double
    Dim hr As Double
    Dim min As Double
    Dim sec As Double
    Dim ts As String
    
    tv = TimeVal
    If tv <= 0 Then
        If StoreTime Then dDuration = 0
    Else
        If StoreTime Then dDuration = tv
        
        tv = Fix(tv)
        min = Fix(tv / 60)
        sec = tv - 60 * min
        hr = Fix(min / 60)
        min = min - 60 * hr
        
        ts = ":" & Format$(sec, "00")
        If hr > 0 Then
            ts = CStr(hr) & ":" & Format$(min, "00") & ts
        Else
            ts = CStr(min) & ts
        End If
        
        FormatTime = ts
    End If
End Function

Private Function FormatBitRate(ByVal BitRate As Double, ByVal Encoding As EncodingEnum, Optional ByVal StoreBitRate As Boolean = False) As String
    On Error Resume Next
    
    Dim br As Double
    br = BitRate
    If br <= 0 Then
        If StoreBitRate Then dBitRate = 0
    Else
        If StoreBitRate Then dBitRate = br
        FormatBitRate = CStr(Fix(br / 1000)) & " kbps " & IIf(Encoding = CBR, "CBR", "VBR")
    End If
End Function

Private Sub RemoveAPICItem(ByVal Index As Long)
    Dim i As Long
    
    cAPICIType.Remove Index
    cAPICType.Remove Index
    cAPICData.Remove Index
    APICData.Remove cAPIC0(Index)
    cAPIC0.Remove Index
    
    For i = Index To cAPIC0.Count
        SetItem cAPIC0, i, cAPIC0(i) - 1
    Next
End Sub

Private Sub AddAPICItem(ByVal MIMEType As String, ByVal PictureType As PictureType, ByVal Data As String)
    Dim APD As APicDecoder
    
    If Data = "" Then
        cAPICIType.Add ""
        cAPICType.Add ""
        cAPICData.Add ""
    Else
        cAPICIType.Add MIMEType
        cAPICType.Add PictureType
        cAPICData.Add Data
    End If
    APICData.Add ""
    
    If Data <> "" Then
        Set APD = New APicDecoder
        APD.InsertImageData APICData, APICData.Count, MIMEType, PictureType, Data, ID3Revision
        Set APD = Nothing
    End If
    
    cAPIC0.Add APICData.Count
End Sub

Private Sub MakeNecessaryChanges(ByVal Index As Long)
    Dim APD As APicDecoder
    Dim GPC As GDIPlusCandy
    
    Dim MIMEType As String
    Dim PictureType As PictureType
    Dim Pic As StdPicture
    Dim PicData As String
    
    Set APD = New APicDecoder
    APD.DecodeImage APICData, cAPIC0(Index), MIMEType, PictureType, Pic, ID3Revision
    If MIMEType = cAPICIType(Index) Then
        PicData = cAPICData(Index)
    Else
        Set GPC = New GDIPlusCandy
        PicData = GPC.ImageToData(Pic, cAPICIType(Index))
        Set GPC = Nothing
    End If
    APD.InsertImageData APICData, cAPIC0(Index), cAPICIType(Index), cAPICType(Index), PicData, ID3Revision
    Set APD = Nothing
End Sub

Private Function FilterEntry(ByVal Description As String, ByVal Filter As String) As String
    FilterEntry = Description & "|" & Filter & "|"
End Function

Private Sub ImageBrowse()
    Dim fn As String
    Dim f As Integer
    Dim st As String
    Dim sMIMEType As String
    Dim tMIMEType As String
    Dim sExt As String
    Dim GPC As GDIPlusCandy
    Dim sPic As StdPicture
    Dim i As Long
    Dim idx As Long
    Dim bConvertImage As Boolean
    
    If ListView1.ListItems.Count > 0 Then
        fn = ShowOpenDialog(hWnd, FilterEntry("All Supported Formats", FILTER_SUPPORTED) & FilterEntry("Windows Bitmap", FILTER_BMP) & FilterEntry("Graphics Interchange Format", FILTER_GIF) & FilterEntry("JPEG File Interchange Format", FILTER_JPEG) & FilterEntry("Portable Network Graphics", FILTER_PNG), "Select Image")
        If fn <> "" Then
            i = InStrRev(fn, ".")
            If i > 0 Then
                sExt = Mid$(LCase$(fn), i + 1)
                Select Case sExt
                    Case "bmp", "dib"
                        sMIMEType = ImageTypeFromIndex(0, ID3Revision)
                        idx = 1
                        If ID3Revision > 2 Then idx = 0
                        bConvertImage = (ID3Revision <= 2)
                    Case "gif"
                        sMIMEType = ImageTypeFromIndex(1, ID3Revision)
                        idx = 1
                        bConvertImage = (ID3Revision <= 2)
                    Case "jpeg", "jpg", "jpe", "jfif", "jfi", "jif"
                        sMIMEType = ImageTypeFromIndex(2, ID3Revision)
                        idx = 0
                        If ID3Revision > 2 Then idx = 2
                    Case "png"
                        sMIMEType = ImageTypeFromIndex(3, ID3Revision)
                        idx = 1
                        If ID3Revision > 2 Then idx = 3
                    Case Else
                        sMIMEType = ""
                        idx = -1
                End Select
                If idx <> -1 Then cmbImageType.ListIndex = idx
            Else
                sMIMEType = ""
            End If
            
            f = FreeFile
            Open fn For Binary Access Read Shared As #f
                st = Space$(LOF(f))
                Get #f, , st
            Close #f
            
            Set GPC = New GDIPlusCandy
            Set sPic = GPC.DataToImage(st)
            Set GPC = Nothing
            
            If Not sPic Is Nothing Then
                cmbImageType.Enabled = True
                cmbPictureType.Enabled = True
                
                ' As ID3v2.0 and ID3v2.2 allow only JPEG and PNG images, do the necessary conversion for BMP and GIF images
                If bConvertImage Then
                    Set GPC = New GDIPlusCandy
                    st = GPC.ImageToData(sPic, ImagePNG)
                    Set sPic = GPC.DataToImage(st) ' Show the converted image
                    Set GPC = Nothing
                End If
                
                tMIMEType = DetermineImageType(st, ID3Revision)
                If sMIMEType <> tMIMEType And tMIMEType <> ImageUnsupported Then
                    sMIMEType = tMIMEType
                    cmbImageType.ListIndex = GetIndex(sMIMEType, ID3Revision)
                End If
                ArtAddProc sMIMEType, sPic, st
            End If
        End If
    End If
End Sub

Private Sub TextProc(Ctl As Object, ByVal Description As String, CountControl As Label, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean)
    If Index = Total Then
        If Ctl = "" Then
            If Index > 0 Then
                FrameBlank = True
                Col.Remove Index
                Index = Index - 1
                Total = Total - 1
                Set NextControl.Picture = Buttons.ListImages(I_ADDI).Picture
                NextControl.ToolTipText = ""
                If Index = 0 Then
                    Set DelControl.Picture = Buttons.ListImages(I_DELI).Picture
                    DelControl.ToolTipText = ""
                End If
            End If
        Else
            If FrameBlank Then
                FrameBlank = False
                Col.Add Ctl.Text
                Index = Index + 1
                Total = Total + 1
                Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
                NextControl.ToolTipText = S_ADD & Description
                Set DelControl.Picture = Buttons.ListImages(I_DEL).Picture
                DelControl.ToolTipText = S_DEL & Description
            Else
                If Index > 0 Then SetItem Col, Index, Ctl.Text
            End If
        End If
    Else
        If Index > 0 Then SetItem Col, Index, Ctl.Text
    End If
    CountControl = CStr(Index) & "/" & CStr(Total)
End Sub

Private Sub ArtAddProc(ByVal MIMEType As String, ByVal Pic As StdPicture, ByVal Data As String)
    picArt.ToolTipText = S_APICTT
    imgArt.ToolTipText = S_APICTT
    imgArt.Visible = True
    Set imgArt.Picture = Nothing
    StretchImage Pic
    Set imgArt.Picture = Pic
    SetBG True
    lblBrowse.Visible = False
    
    If indAPIC = totAPIC Then
        If bAPICBlank Then
            bAPICBlank = False
            AddAPICItem MIMEType, cmbPictureType.ListIndex, Data
            indAPIC = indAPIC + 1
            totAPIC = totAPIC + 1
            Set nextArt.Picture = Buttons.ListImages(I_ADD).Picture
            nextArt.ToolTipText = S_ADD & S_APIC
            Set delArt.Picture = Buttons.ListImages(I_DEL).Picture
            delArt.ToolTipText = S_DEL & S_APIC
        Else
            If indAPIC > 0 Then
                SetItem cAPICData, indAPIC, Data
                SetItem cAPICIType, indAPIC, MIMEType
                SetItem cAPICType, indAPIC, cmbPictureType.ListIndex
            End If
        End If
    Else
        If indAPIC > 0 Then
            SetItem cAPICData, indAPIC, Data
            SetItem cAPICIType, indAPIC, MIMEType
            SetItem cAPICType, indAPIC, cmbPictureType.ListIndex
        End If
    End If
    countArt = CStr(indAPIC) & "/" & CStr(totAPIC)
End Sub

Private Sub PrevProc(Ctl As Object, ByVal Description As String, CountControl As Label, PrevControl As Image, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean, Optional ByVal DeleteMode As Boolean = False)
    Dim bPic As Boolean, GDP As GDIPlusCandy, Pic As StdPicture
    bPic = (TypeName(Ctl) = "PictureBox")
    If (Index > 0 And DeleteMode) Or (Index > 1 And Not DeleteMode) Or (Index > 0 And Not DeleteMode And FrameBlank) Then
        If Index = Total Then
            Set DelControl.Picture = Buttons.ListImages(I_DEL).Picture
            DelControl.ToolTipText = S_DEL & Description
            If FrameBlank Then
IsBlank:
                Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
                NextControl.ToolTipText = S_ADD & Description
                FrameBlank = False
            Else
                Index = Index - 1
                If Index = 0 Then
                    GoTo IsBlank
                Else
                    Set NextControl.Picture = Buttons.ListImages(I_NEXT).Picture
                    NextControl.ToolTipText = S_NEXT & Description
                End If
            End If
        Else
            Index = Index - 1
        End If
        If Index <= 1 Then
            Set PrevControl.Picture = Buttons.ListImages(I_PREVI).Picture
            PrevControl.ToolTipText = ""
        End If
        If Index = 0 Then
            If bPic Then
                SetBG False
                Set imgArt.Picture = Nothing
                StretchImage imgArt.Picture
                imgArt.Visible = False
                cmbImageType.Enabled = False
                cmbPictureType.Enabled = False
                cmbImageType.ListIndex = 2 * (cmbImageType.ListIndex \ 4)
                cmbPictureType.ListIndex = 0
                lblBrowse.Visible = True
                Ctl.ToolTipText = ""
                imgArt.ToolTipText = ""
            Else
                Ctl = ""
            End If
            FrameBlank = True
        Else
            If bPic Then
                lblBrowse.Visible = False
                imgArt.Visible = True
                cmbImageType.Enabled = True
                cmbPictureType.Enabled = True
                Set GDP = New GDIPlusCandy
                Set Pic = GDP.DataToImage(Col(Index))
                Set GDP = Nothing
                Set imgArt.Picture = Nothing
                StretchImage Pic
                Set imgArt.Picture = Pic
                SetBG True
                cmbImageType.ListIndex = GetIndex(cAPICIType(Index), ID3Revision)
                cmbPictureType.ListIndex = cAPICType(Index)
                Ctl.ToolTipText = S_APICTT
                imgArt.ToolTipText = S_APICTT
            Else
                Ctl = Col(Index)
            End If
        End If
        CountControl = CStr(Index) & "/" & CStr(Total)
    End If
End Sub

Private Sub NextProc(Ctl As Object, ByVal Description As String, CountControl As Label, PrevControl As Image, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean)
    Dim bPic As Boolean, bBlank As Boolean, GDP As GDIPlusCandy, Pic As StdPicture, lIType As Long, vType As Variant
    Dim bRefresh As Boolean
    Dim blFrameBlank As Boolean
    bPic = (TypeName(Ctl) = "PictureBox")
    If Index < Total Then
        bRefresh = True
        Index = Index + 1
        If Index = Total Then
            Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
            NextControl.ToolTipText = S_ADD & Description
        End If
    Else
        If bPic Then
            If imgArt.Visible Then
                bRefresh = True
                blFrameBlank = True
                bBlank = True
                Set DelControl.Picture = Buttons.ListImages(I_DEL).Picture
                DelControl.ToolTipText = S_DEL & Description
            End If
        Else
            If Ctl <> "" Then
                bRefresh = True
                blFrameBlank = True
                Index = Index + 1
                Col.Add ""
                Total = Col.Count
                Set DelControl.Picture = Buttons.ListImages(I_DEL).Picture
                DelControl.ToolTipText = S_DEL & Description
            End If
        End If
    End If
    If bRefresh Then
        Set PrevControl.Picture = Buttons.ListImages(I_PREV).Picture
        PrevControl.ToolTipText = S_PREV & Description
        If bPic Then
            If bBlank Then
                cmbImageType.Enabled = False
                cmbPictureType.Enabled = False
                imgArt.Visible = False
                SetBG False
                Set imgArt.Picture = Nothing
                StretchImage imgArt.Picture
                cmbImageType.ListIndex = 2 * (cmbImageType.ListCount \ 4)
                cmbPictureType.ListIndex = 0
                lblBrowse.Visible = True
                Ctl.ToolTipText = ""
                imgArt.ToolTipText = ""
                Set NextControl.Picture = Buttons.ListImages(I_ADDI).Picture
                NextControl.ToolTipText = ""
            Else
                cmbImageType.Enabled = True
                cmbPictureType.Enabled = True
                imgArt.Visible = True
                Set GDP = New GDIPlusCandy
                Set Pic = GDP.DataToImage(Col(Index))
                Set GDP = Nothing
                Set imgArt.Picture = Nothing
                StretchImage Pic
                Set imgArt.Picture = Pic
                SetBG True
                cmbImageType.ListIndex = GetIndex(cAPICIType(Index), ID3Revision)
                cmbPictureType.ListIndex = cAPICType(Index)
                lblBrowse.Visible = False
                Ctl.ToolTipText = S_APICTT
                imgArt.ToolTipText = S_APICTT
                Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
                NextControl.ToolTipText = S_ADD & Description
            End If
        Else
            Ctl = Col(Index)
        End If
        If blFrameBlank Then FrameBlank = True
        CountControl = CStr(Index) & "/" & CStr(Total)
    End If
End Sub

Private Sub DelProc(Ctl As Object, ByVal Description As String, CountControl As Label, PrevControl As Image, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean)
    Dim bPic As Boolean, GDP As GDIPlusCandy, Pic As StdPicture
    bPic = (TypeName(Ctl) = "PictureBox")
    If Total > 0 Then
        If FrameBlank Then
            PrevProc Ctl, Description, CountControl, PrevControl, NextControl, DelControl, Col, Index, Total, FrameBlank, True
        Else
            If bPic Then
                RemoveAPICItem Index
            Else
                Col.Remove Index
            End If
            Total = Col.Count
            If Index > Total Then Index = Total
            If Index = 0 Then
                If bPic Then
                    cmbImageType.Enabled = False
                    cmbPictureType.Enabled = False
                    SetBG False
                    Set imgArt.Picture = Nothing
                    StretchImage imgArt.Picture
                    cmbImageType.ListIndex = 2 * (cmbImageType.ListIndex \ 4)
                    cmbPictureType.ListIndex = 0
                    imgArt.Visible = False
                    lblBrowse.Visible = True
                    Ctl.ToolTipText = ""
                    imgArt.ToolTipText = ""
                Else
                    Ctl = ""
                End If
                FrameBlank = True
            Else
                If bPic Then
                    cmbImageType.Enabled = True
                    cmbPictureType.Enabled = True
                    imgArt.Visible = True
                    Set GDP = New GDIPlusCandy
                    Set Pic = GDP.DataToImage(Col(Index))
                    Set GDP = Nothing
                    Set imgArt.Picture = Nothing
                    StretchImage Pic
                    Set imgArt.Picture = Pic
                    SetBG True
                    cmbImageType.ListIndex = GetIndex(cAPICIType(Index), ID3Revision)
                    cmbPictureType.ListIndex = cAPICType(Index)
                    lblBrowse.Visible = False
                    Ctl.ToolTipText = S_APICTT
                    imgArt.ToolTipText = S_APICTT
                Else
                    Ctl = Col(Index)
                End If
            End If
            If Index = Total Then
                If Index = 0 Then
                    Set NextControl.Picture = Buttons.ListImages(I_ADDI).Picture
                    NextControl.ToolTipText = ""
                Else
                    Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
                    NextControl.ToolTipText = S_ADD & Description
                End If
            End If
            If Index <= 1 Then
                Set PrevControl.Picture = Buttons.ListImages(I_PREVI).Picture
                PrevControl.ToolTipText = ""
                If Total = 0 Then
                    Set DelControl.Picture = Buttons.ListImages(I_DELI).Picture
                    DelControl.ToolTipText = ""
                End If
            End If
            CountControl = CStr(Index) & "/" & CStr(Total)
        End If
    End If
End Sub

Private Function WithinBounds(ByVal Obj As Object, ByVal x As Single, ByVal Y As Single) As Boolean
    Dim oWidth As Single, oHeight As Single
    If TypeName(Obj) = "PictureBox" Then
        oWidth = Obj.ScaleWidth
        oHeight = Obj.ScaleHeight
    Else
        oWidth = Obj.Width
        oHeight = Obj.Height
    End If
    WithinBounds = (x >= 0 And x <= oWidth And Y >= 0 And Y <= oHeight)
End Function

Private Sub ChangeFields(ByVal bEnabled As Boolean)

    Dim BG As Long
    BG = IIf(bEnabled, vbWindowBackground, vbButtonFace)

    If txtTitle.Enabled <> bEnabled Then
        txtTitle.Enabled = bEnabled
        txtTitle.BackColor = BG
    End If

    If txtArtist.Enabled <> bEnabled Then
        txtArtist.Enabled = bEnabled
        txtArtist.BackColor = BG
    End If

    If txtAlbum.Enabled <> bEnabled Then
        txtAlbum.Enabled = bEnabled
        txtAlbum.BackColor = BG
    End If

    If cmbGenre.Enabled <> bEnabled Then
        cmbGenre.Enabled = bEnabled
        cmbGenre.BackColor = BG
    End If

    If txtTrackNumber.Enabled <> bEnabled Then
        txtTrackNumber.Enabled = bEnabled
        txtTrackNumber.BackColor = BG
    End If

    If txtTracksTotal.Enabled <> bEnabled Then
        txtTracksTotal.Enabled = bEnabled
        txtTracksTotal.BackColor = BG
    End If

    If txtYear.Enabled <> bEnabled Then
        txtYear.Enabled = bEnabled
        txtYear.BackColor = BG
    End If

    If txtComments.Enabled <> bEnabled Then
        txtComments.Enabled = bEnabled
        txtComments.BackColor = BG
    End If

    If txtLyrics.Enabled <> bEnabled Then
        txtLyrics.Enabled = bEnabled
        txtLyrics.BackColor = BG
    End If

    If txtComposer.Enabled <> bEnabled Then
        txtComposer.Enabled = bEnabled
        txtComposer.BackColor = BG
    End If

    If txtBand.Enabled <> bEnabled Then
        txtBand.Enabled = bEnabled
        txtBand.BackColor = BG
    End If

    If txtConductor.Enabled <> bEnabled Then
        txtConductor.Enabled = bEnabled
        txtConductor.BackColor = BG
    End If

    If txtInterpretedBy.Enabled <> bEnabled Then
        txtInterpretedBy.Enabled = bEnabled
        txtInterpretedBy.BackColor = BG
    End If

    If txtLyricist.Enabled <> bEnabled Then
        txtLyricist.Enabled = bEnabled
        txtLyricist.BackColor = BG
    End If

    If txtOriginalArtist.Enabled <> bEnabled Then
        txtOriginalArtist.Enabled = bEnabled
        txtOriginalArtist.BackColor = BG
    End If

    If txtOriginalAlbum.Enabled <> bEnabled Then
        txtOriginalAlbum.Enabled = bEnabled
        txtOriginalAlbum.BackColor = BG
    End If

    If txtOriginalFileName.Enabled <> bEnabled Then
        txtOriginalFileName.Enabled = bEnabled
        txtOriginalFileName.BackColor = BG
    End If

    If txtOriginalLyricist.Enabled <> bEnabled Then
        txtOriginalLyricist.Enabled = bEnabled
        txtOriginalLyricist.BackColor = BG
    End If

    If txtOriginalReleaseYear.Enabled <> bEnabled Then
        txtOriginalReleaseYear.Enabled = bEnabled
        txtOriginalReleaseYear.BackColor = BG
    End If

    If txtCopyright.Enabled <> bEnabled Then
        txtCopyright.Enabled = bEnabled
        txtCopyright.BackColor = BG
    End If

    If txtFileOwner.Enabled <> bEnabled Then
        txtFileOwner.Enabled = bEnabled
        txtFileOwner.BackColor = BG
    End If

    If txtPublisher.Enabled <> bEnabled Then
        txtPublisher.Enabled = bEnabled
        txtPublisher.BackColor = BG
    End If

    If txtInternetRadioStationName.Enabled <> bEnabled Then
        txtInternetRadioStationName.Enabled = bEnabled
        txtInternetRadioStationName.BackColor = BG
    End If

    If txtInternetRadioStationOwner.Enabled <> bEnabled Then
        txtInternetRadioStationOwner.Enabled = bEnabled
        txtInternetRadioStationOwner.BackColor = BG
    End If

    If txtISRC.Enabled <> bEnabled Then
        txtISRC.Enabled = bEnabled
        txtISRC.BackColor = BG
    End If

    If txtLanguages.Enabled <> bEnabled Then
        txtLanguages.Enabled = bEnabled
        txtLanguages.BackColor = BG
    End If

    If txtCommercialInfo.Enabled <> bEnabled Then
        txtCommercialInfo.Enabled = bEnabled
        txtCommercialInfo.BackColor = BG
    End If

    If txtCopyrightInfo.Enabled <> bEnabled Then
        txtCopyrightInfo.Enabled = bEnabled
        txtCopyrightInfo.BackColor = BG
    End If

    If txtAudioURL.Enabled <> bEnabled Then
        txtAudioURL.Enabled = bEnabled
        txtAudioURL.BackColor = BG
    End If

    If txtArtistURL.Enabled <> bEnabled Then
        txtArtistURL.Enabled = bEnabled
        txtArtistURL.BackColor = BG
    End If

    If txtAudioSourceURL.Enabled <> bEnabled Then
        txtAudioSourceURL.Enabled = bEnabled
        txtAudioSourceURL.BackColor = BG
    End If

    If txtInternetRadioStationURL.Enabled <> bEnabled Then
        txtInternetRadioStationURL.Enabled = bEnabled
        txtInternetRadioStationURL.BackColor = BG
    End If

    If txtPaymentURL.Enabled <> bEnabled Then
        txtPaymentURL.Enabled = bEnabled
        txtPaymentURL.BackColor = BG
    End If

    If txtPublisherURL.Enabled <> bEnabled Then
        txtPublisherURL.Enabled = bEnabled
        txtPublisherURL.BackColor = BG
    End If

    If txtEncodedBy.Enabled <> bEnabled Then
        txtEncodedBy.Enabled = bEnabled
        txtEncodedBy.BackColor = BG
    End If

    If txtBPM.Enabled <> bEnabled Then
        txtBPM.Enabled = bEnabled
        txtBPM.BackColor = BG
    End If

    If cmbKey.Enabled <> bEnabled Then
        cmbKey.Enabled = bEnabled
        cmbKey.BackColor = BG
    End If

    If txtDiscNumber.Enabled <> bEnabled Then
        txtDiscNumber.Enabled = bEnabled
        txtDiscNumber.BackColor = BG
    End If

    If txtDiscsTotal.Enabled <> bEnabled Then
        txtDiscsTotal.Enabled = bEnabled
        txtDiscsTotal.BackColor = BG
    End If

End Sub

Private Sub VMove(ByVal By As Long, ParamArray Objs() As Variant)
    Dim i As Long
    For i = LBound(Objs) To UBound(Objs)
        Objs(i).Top = Objs(i).Top + By
    Next
End Sub

Private Sub AdjustVScrollProps()
    Dim VMax As Integer
    VMax = Ceiling((Frame3.Height - Frame2(1).Height) / 360) - 1
    If VScroll1.Max <> VMax Then
        VScroll1.Max = VMax
        If VScroll1.Max = 0 Then
            VScroll1.Visible = False
        Else
            VScroll1.Visible = True
        End If
    End If
    If VScroll1.Max = 0 Then
        Frame3.Width = Frame2(1).Width
    Else
        Frame3.Width = Frame2(1).Width - 255
    End If
    VScroll1_Change
End Sub

Private Sub ShowOrHideNecessaryFields()
    Dim bShow As Boolean
    Dim lAdd As Long
    
    If Frame2(1).Visible And Frame3.Visible Then
        If ID3Revision > 2 Then
            bShow = True
            lAdd = 360
        Else
            bShow = False
            lAdd = -360
        End If
        
        ' Hide the text fields not supported by ID3v2.0 and ID3v2.2
        
        If txtFileOwner.Visible <> bShow Then
            Label27.Visible = bShow
            txtFileOwner.Visible = bShow
            VMove lAdd, Label38, txtPublisher, _
                        Label28, txtInternetRadioStationName, _
                        Label29, txtInternetRadioStationOwner, _
                        Label30, txtISRC, _
                        Label16, txtLanguages, _
                        Label31, txtCommercialInfo, _
                            countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, _
                        Label32, txtCopyrightInfo, _
                        Label21, txtAudioURL, _
                        Label33, txtArtistURL, _
                            countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, _
                        Label34, txtAudioSourceURL, _
                        Label35, txtInternetRadioStationURL, _
                        Label22, txtPaymentURL, _
                        Label24, txtPublisherURL, _
                        Label23, txtEncodedBy, _
                        Label19, txtBPM, _
                        Label20, cmbKey, _
                        Label36, txtDiscNumber, _
                        Label37, txtDiscsTotal
            Frame3.Height = Frame3.Height + lAdd
        End If
        
        If txtInternetRadioStationName.Visible <> bShow Then
            Label28.Visible = bShow
            txtInternetRadioStationName.Visible = bShow
            VMove lAdd, Label29, txtInternetRadioStationOwner, _
                        Label30, txtISRC, _
                        Label16, txtLanguages, _
                        Label31, txtCommercialInfo, _
                            countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, _
                        Label32, txtCopyrightInfo, _
                        Label21, txtAudioURL, _
                        Label33, txtArtistURL, _
                            countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, _
                        Label34, txtAudioSourceURL, _
                        Label35, txtInternetRadioStationURL, _
                        Label22, txtPaymentURL, _
                        Label24, txtPublisherURL, _
                        Label23, txtEncodedBy, _
                        Label19, txtBPM, _
                        Label20, cmbKey, _
                        Label36, txtDiscNumber, _
                        Label37, txtDiscsTotal
            Frame3.Height = Frame3.Height + lAdd
        End If
        
        If txtInternetRadioStationOwner.Visible <> bShow Then
            Label29.Visible = bShow
            txtInternetRadioStationOwner.Visible = bShow
            VMove lAdd, Label30, txtISRC, _
                        Label16, txtLanguages, _
                        Label31, txtCommercialInfo, _
                            countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, _
                        Label32, txtCopyrightInfo, _
                        Label21, txtAudioURL, _
                        Label33, txtArtistURL, _
                            countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, _
                        Label34, txtAudioSourceURL, _
                        Label35, txtInternetRadioStationURL, _
                        Label22, txtPaymentURL, _
                        Label24, txtPublisherURL, _
                        Label23, txtEncodedBy, _
                        Label19, txtBPM, _
                        Label20, cmbKey, _
                        Label36, txtDiscNumber, _
                        Label37, txtDiscsTotal
            Frame3.Height = Frame3.Height + lAdd
        End If
        
        If txtPaymentURL.Visible <> bShow Then
            Label22.Visible = bShow
            txtPaymentURL.Visible = bShow
            VMove lAdd, Label24, txtPublisherURL, _
                        Label23, txtEncodedBy, _
                        Label19, txtBPM, _
                        Label20, cmbKey, _
                        Label36, txtDiscNumber, _
                        Label37, txtDiscsTotal
            Frame3.Height = Frame3.Height + lAdd
        End If
        
        AdjustVScrollProps
    End If
End Sub

Private Sub LoadFileEntries(ByVal Path As String)
    On Error Resume Next
    
    Dim ID3 As New clsID3
    Dim sPath As String
    Dim d As String
    Dim HourPart As String
    Dim BlankWCOM As New MultiFrameData
    Dim BlankWOAR As New MultiFrameData
    Dim BlankAPIC As New MultiFrameData
    
    sPath = Path
    If Right$(Path, 1) <> "\" Then sPath = sPath & "\"
    
    d = Dir$(sPath)
    ListView1.ListItems.Clear
    
    ID3Revision = 3
    ShowOrHideNecessaryFields
    txtTitle = ""
    txtArtist = ""
    txtAlbum = ""
    cmbGenre = ""
    txtTrackNumber = ""
    txtTracksTotal = ""
    txtYear = ""
    txtComments = ""
    txtLyrics = ""
    txtComposer = ""
    txtBand = ""
    txtConductor = ""
    txtInterpretedBy = ""
    txtLyricist = ""
    txtOriginalArtist = ""
    txtOriginalAlbum = ""
    txtOriginalFileName = ""
    txtOriginalLyricist = ""
    txtOriginalReleaseYear = ""
    txtCopyright = ""
    txtFileOwner = ""
    txtPublisher = ""
    txtInternetRadioStationName = ""
    txtInternetRadioStationOwner = ""
    txtISRC = ""
    txtLanguages = ""
    LoadMultiData txtCommercialInfo, BlankWCOM, S_CURL, countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
    txtCopyrightInfo = ""
    txtAudioURL = ""
    LoadMultiData txtArtistURL, BlankWOAR, S_AURL, countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
    txtAudioSourceURL = ""
    txtInternetRadioStationURL = ""
    txtPaymentURL = ""
    txtPublisherURL = ""
    txtEncodedBy = ""
    txtBPM = ""
    cmbKey = ""
    txtDiscNumber = ""
    txtDiscsTotal = ""
    LoadMultiData picArt, BlankAPIC, S_APIC, countArt, prevArt, nextArt, delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
    lblBrowse.Visible = False
    
    ChangeFields False
    If Command1.Enabled Then Command1.Enabled = False
    If Command2.Enabled Then Command2.Enabled = False
    If Command3.Enabled Then Command3.Enabled = False
    If Command4.Enabled Then Command4.Enabled = False
    If Command5.Enabled Then Command5.Enabled = False
    If Command6.Enabled Then Command6.Enabled = False
    
    Do Until d = ""
        If d <> "." And d <> ".." Then
            If LCase$(Right$(d, 4)) = ".mp3" Then
                With ListView1
                    If MousePointer = vbDefault Then
                        MousePointer = vbHourglass
                        DoEvents
                    End If
                    .ListItems.Add Text:=d
                    ID3.FileName = sPath & d
                    With .ListItems(.ListItems.Count)
                        .SubItems(1) = ID3.Title
                        .SubItems(2) = ID3.Artist
                        .SubItems(3) = ID3.Album
                        .SubItems(4) = FormatGenre(ID3, ID3.GenreID, ID3.Genre)
                        .SubItems(5) = ID3.TrackNumber
                        .SubItems(6) = ID3.TracksTotal
                        .SubItems(7) = ID3.Year
                        .SubItems(8) = FormatTime(ID3.length)
                        .SubItems(9) = FormatBitRate(ID3.BitRate, ID3.Encoding)
                        .SubItems(10) = ID3.Comments
                    End With
                End With
            End If
        End If
        d = Dir$
    Loop
    
    Resort = True
    SortLvwOnLong ListView1, ListView1.SortKey + 1
    Resort = False
    
    If MousePointer = vbHourglass Then _
       MousePointer = vbDefault
    
    If Not Command1.Enabled Then Command1.Enabled = True
    If Not Command5.Enabled Then Command5.Enabled = True
    
    If ListView1.ListItems.Count > 0 Then
        ChangeFields True
        If Not Command2.Enabled Then Command2.Enabled = True
        If Not Command3.Enabled Then Command3.Enabled = True
        If Not Command4.Enabled Then Command4.Enabled = True
        If Not Command6.Enabled Then Command6.Enabled = True
        ListView1.ListItems(1).Selected = True
        ListView1_ItemClick ListView1.ListItems(1)
    Else
        ChangeFields False
        If Command2.Enabled Then Command2.Enabled = False
        If Command3.Enabled Then Command3.Enabled = False
        If Command4.Enabled Then Command4.Enabled = False
        If Command6.Enabled Then Command6.Enabled = False
    End If
End Sub

Private Sub cmbImageType_Change()
    cmbImageType_Click
End Sub

Private Sub cmbImageType_Click()
    Dim MIMEType As String
    Dim PNGIndex As Long
    If cmbImageType.Enabled Then
        If cmbImageType.ListCount = 4 Then
            Select Case cmbImageType.ListIndex
                Case 0: MIMEType = ImageBMP
                Case 1: MIMEType = ImageGIF
                Case 2: MIMEType = ImageJPEG
                Case 3: MIMEType = ImagePNG
            End Select
            PNGIndex = 3
        Else
            Select Case cmbImageType.ListIndex
                Case 0: MIMEType = ImageJPEGOld
                Case 1: MIMEType = ImagePNGOld
            End Select
            PNGIndex = 1
        End If
        SetItem cAPICIType, indAPIC, MIMEType
        If cmbPictureType.ListIndex = 1 And cmbImageType.ListIndex <> PNGIndex Then
            cmbPictureType.ListIndex = 2
            SetItem cAPICType, indAPIC, cmbImageType.ListIndex
        End If
    End If
End Sub

Private Sub cmbPictureType_Change()
    cmbPictureType_Click
End Sub

Private Sub cmbPictureType_Click()
    If cmbPictureType.Enabled Then
        If cmbPictureType.ListIndex = 1 Then
            If cmbImageType.ListIndex <> (1 + 2 * (cmbImageType.ListCount \ 4)) Or HimetricToPixelsX(imgArt.Picture.Width) <> 32 Or HimetricToPixelsY(imgArt.Picture.Height) <> 32 Then
                cmbPictureType.ListIndex = 2
            End If
        End If
        SetItem cAPICType, indAPIC, cmbPictureType.ListIndex
    End If
End Sub

Private Sub Command1_Click()
    Dim Folder As String
    Dim sExistingFolder As String
    
    If Right$(Text1, 1) = "\" Then
        sExistingFolder = Text1
    Else
        sExistingFolder = Text1 & "\"
    End If
    
    Folder = BrowseForFolder(hWnd, "Select a folder:", sExistingFolder)
    If Folder <> "" Then
        Text1 = Folder
        LoadFileEntries Folder
    End If
End Sub

Private Sub Command2_Click()
    Dim ID3 As New clsID3
    Dim i As Long
    
    With ID3
        .FileName = Text1 & "\" & ListView1.SelectedItem.Text
        .Title = txtTitle
        .Artist = txtArtist
        .Album = txtAlbum
        .Genre = cmbGenre.Text
        .GenreID = .ToGenreID(.Genre)
        .TrackNumber = txtTrackNumber
        .TracksTotal = txtTracksTotal
        .Year = txtYear
        .Comments = txtComments
        .Lyrics = txtLyrics
        .Composer = txtComposer
        .Band = txtBand
        .Conductor = txtConductor
        .InterpretedBy = txtInterpretedBy
        .Lyricist = txtLyricist
        .OriginalArtist = txtOriginalArtist
        .OriginalAlbum = txtOriginalAlbum
        .OriginalFileName = txtOriginalFileName
        .OriginalLyricist = txtOriginalLyricist
        .OriginalReleaseYear = txtOriginalReleaseYear
        .Copyright = txtCopyright
        .FileOwner = txtFileOwner
        .Publisher = txtPublisher
        .InternetRadioStationName = txtInternetRadioStationName
        .InternetRadioStationOwner = txtInternetRadioStationOwner
        .ISRC = txtISRC
        .Languages = txtLanguages
        .CommercialInfo.Clear
        For i = 1 To cWCOM.Count
            .CommercialInfo.Add cWCOM(i)
        Next
        .CopyrightInfo = txtCopyrightInfo
        .AudioURL = txtAudioURL
        .ArtistURL.Clear
        For i = 1 To cWOAR.Count
            .ArtistURL.Add cWOAR(i)
        Next
        .AudioSourceURL = txtAudioSourceURL
        .InternetRadioURL = txtInternetRadioStationURL
        .PaymentURL = txtPaymentURL
        .PublisherURL = txtPublisherURL
        .EncodedBy = txtEncodedBy
        .BeatsPerMinute = txtBPM
        .InitialKey = cmbKey
        .DiscNumber = txtDiscNumber
        .DiscsTotal = txtDiscsTotal
        For i = 1 To cAPICData.Count
            MakeNecessaryChanges i
        Next
        Set .AttachedPictures = APICData
        
        If MousePointer = vbDefault Then
            MousePointer = vbHourglass
            DoEvents
        End If
        
        .UpdateID3Tags
        
        If MousePointer = vbHourglass Then _
           MousePointer = vbDefault
        
        ListView1_ItemClick ListView1.SelectedItem
    End With
End Sub

Private Sub Command3_Click()
    Dim ID3 As New clsID3
    
    With ID3
        .FileName = Text1 & "\" & ListView1.SelectedItem.Text
        
        If MousePointer = vbDefault Then
            MousePointer = vbHourglass
            DoEvents
        End If
        
        .DeleteID3Tags
        
        If MousePointer = vbHourglass Then _
           MousePointer = vbDefault
        
        ListView1_ItemClick ListView1.SelectedItem
    End With
End Sub

Private Function Ceiling(ByVal num As Double) As Double
    Dim d As Double
    d = num
    If num <> Fix(num) Then d = d + 1
    Ceiling = d
End Function

Private Sub Command4_Click()
    On Error Resume Next
    
    bConnect = True
    If Dir$(GetSpecialFolderLocation(CSIDL_SYSTEM) & "\vbzlib1.dll") = "" Then
        bRet = False
        frmGZip.Show 1
        If bRet Then GoTo NormProc
    Else
NormProc:
        bRet = False
        bItunes = True
        frmConnecting.Show 1
        If bRet Then
            bRet = False
            frmSelection.Show 1
            If bRet Then
                bRet = False
                bConnect = False
                frmConnecting.Show 1
            ElseIf bLaunchArtDL Then
                frmAlbumArtDL.Show 1
            End If
        End If
    End If
End Sub

Private Sub Command5_Click()
    LoadFileEntries Text1
End Sub

Private Sub Command6_Click()
    On Error Resume Next
    
    bConnect = True
    bRet = False
    bItunes = False
    frmConnecting.Show 1
    If bRet Then
        bRet = False
        frmSelection.Show 1
        If bRet Then
            bRet = False
            bConnect = False
            frmConnecting.Show 1
        ElseIf bLaunchArtDL Then
            frmAlbumArtDL.Show 1
        End If
    End If
End Sub

Private Sub delArt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(delArt, x, Y) Then
        DelProc picArt, S_APIC, countArt, prevArt, nextArt, delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
    End If
End Sub

Private Sub delArtistURL_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(delArtistURL, x, Y) Then
        DelProc txtArtistURL, S_AURL, countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
    End If
End Sub

Private Sub delCommercialInfo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(delCommercialInfo, x, Y) Then
        DelProc txtCommercialInfo, S_CURL, countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim hBitmap As Long
    
    Dim i As Long
    Dim t As Long
    Dim strT As String
    
    bInitialized = False
    
    XSize = GetSetting(Caption, "Window", "XSize", Width)
    If Err Then
        XSize = Width
        Err.Clear
    End If
    
    YSize = GetSetting(Caption, "Window", "YSize", Height)
    If Err Then
        YSize = Height
        Err.Clear
    End If
    
    XPos = GetSetting(Caption, "Window", "XPos", Left)
    If Err Then
        XPos = Left
        Err.Clear
    End If
    
    YPos = GetSetting(Caption, "Window", "YPos", Top)
    If Err Then
        YPos = Top
        Err.Clear
    End If
    
    myWindowState = GetSetting(Caption, "Window", "State", vbNormal)
    If Err Then
        myWindowState = vbNormal
        Err.Clear
    End If
    
    If XSize < 568 * Screen.TwipsPerPixelX Then XSize = 568 * Screen.TwipsPerPixelX
    If XSize > Screen.Width Then XSize = Screen.Width
    
    If YSize < 445 * Screen.TwipsPerPixelY Then YSize = 445 * Screen.TwipsPerPixelY
    If YSize > Screen.Height Then YSize = Screen.Height
    
    If XPos < 0 Then XPos = 0
    If XPos > Screen.Width - Width Then XPos = Screen.Width - Width
    
    If YPos < 0 Then YPos = 0
    If YPos > Screen.Height - Height Then YPos = Screen.Height - Height
    
    If myWindowState <> vbNormal And myWindowState <> vbMaximized Then _
        myWindowState = vbNormal
    
    If Width <> XSize Then _
        Width = XSize
    If Height <> YSize Then _
        Height = YSize
    
    If Left <> XPos Then _
        Left = XPos
    If Top <> YPos Then _
        Top = YPos
    
    If WindowState <> myWindowState Then _
        WindowState = myWindowState
    
    bInitialized = True
    
    With ListView1
        .ColumnHeaderIcons = ImageList1
        
        .SortKey = GetSetting(Caption, "Columns", "SortKey", 0)
        If Err Then
            Err.Clear
            .SortKey = 0
        End If
        
        .SortOrder = GetSetting(Caption, "Columns", "SortOrder", lvwAscending)
        If Err Then
            Err.Clear
            .SortOrder = lvwAscending
        End If
        Resort = False
        ShowListViewColumnHeaderSortIcon ListView1
        
        For i = 1 To .ColumnHeaders.Count
            t = GetSetting(Caption, "ColumnPos", Format$(i, "00"), CStr(i))
            If Err Then
                Err.Clear
                t = i
            End If
            
            .ColumnHeaders(t).Position = CInt(i)
            If Err Then Err.Clear
            
            .ColumnHeaders(i).Width = GetSetting(Caption, "Columns", Format$(i, "00"), .ColumnHeaders(i).Width)
            If Not Err Then
                If .ColumnHeaders(i).Width < 200 Then
                    .ColumnHeaders(i).Width = 200
                End If
            Else
                Err.Clear
            End If
        Next
    End With
    
    gHW = hWnd
    Hook
    
    strT = GetSetting(Caption, "MP3s", "Directory", GetSpecialFolderLocation(CSIDL_PERSONAL))
    If Right$(strT, 1) = "\" Then strT = Left$(strT, Len(strT) - 1)
    
    If Dir$(strT & "\") = "" Then
        strT = GetSpecialFolderLocation(CSIDL_PERSONAL)
    End If
    
    Text1 = strT
    Timer1.Enabled = True
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        If bInitialized Then myWindowState = WindowState
        Text1.Width = ScaleWidth - 1560
        Command1.Left = ScaleWidth - 1335
        Command5.Left = ScaleWidth - 840
        ListView1.Width = ScaleWidth - 225
        ListView1.Height = 4 * ScaleHeight \ 5 - 2961
        Frame1.Width = ScaleWidth - 225
        Frame1.Top = 4 * ScaleHeight \ 5 - 2376
        Frame1.Height = ScaleHeight \ 5 + 2271
        Command2.Left = Frame1.Width - 3135
        Command2.Top = Frame1.Height - 495
        Command3.Left = Frame1.Width - 1575
        Command3.Top = Frame1.Height - 495
        Command4.Top = Frame1.Height - 495
        Command6.Top = Frame1.Height - 495
        TabStrip1.Width = Frame1.Width - 240
        TabStrip1.Height = Frame1.Height - 840
        Frame2(0).Width = Frame1.Width - 480
        Frame2(0).Height = Frame1.Height - 1440
        Frame2(1).Width = Frame1.Width - 480
        Frame2(1).Height = Frame1.Height - 1440
        Frame2(2).Width = Frame1.Width - 480
        Frame2(2).Height = Frame1.Height - 1440
        VScroll1.Height = Frame2(1).Height
        VScroll1.Left = Frame2(1).Width - 255
        AdjustVScrollProps
        txtTitle.Width = Frame2(0).Width - 600
        txtArtist.Width = Frame2(0).Width - 600
        txtAlbum.Width = Frame2(0).Width - 600
        txtComments.Width = Frame2(0).Width \ 2 - 1012
        txtComments.Height = Frame2(0).Height - 1440
        Label8.Left = Frame2(0).Width \ 2 + 173
        txtLyrics.Left = Frame2(0).Width \ 2 + 773
        txtLyrics.Height = Frame2(0).Height - 1440
        txtLyrics.Width = Frame2(0).Width \ 2 - 772
        txtComposer.Width = Frame3.Width - 1920
        txtBand.Width = Frame3.Width - 1920
        txtConductor.Width = Frame3.Width - 1920
        txtInterpretedBy.Width = Frame3.Width - 1920
        txtLyricist.Width = Frame3.Width - 1920
        txtOriginalArtist.Width = Frame3.Width - 1920
        txtOriginalAlbum.Width = Frame3.Width - 1920
        txtOriginalFileName.Width = Frame3.Width - 1920
        txtOriginalLyricist.Width = Frame3.Width - 1920
        txtOriginalReleaseYear.Width = Frame3.Width - 1920
        txtCopyright.Width = Frame3.Width - 1920
        txtFileOwner.Width = Frame3.Width - 1920
        txtPublisher.Width = Frame3.Width - 1920
        txtInternetRadioStationName.Width = Frame3.Width - 1920
        txtInternetRadioStationOwner.Width = Frame3.Width - 1920
        txtISRC.Width = Frame3.Width - 1920
        txtLanguages.Width = Frame3.Width - 1920
        txtCommercialInfo.Width = Frame3.Width - 3165
        countCommercialInfo.Left = Frame3.Width - 1335
        prevCommercialInfo.Left = Frame3.Width - 750
        nextCommercialInfo.Left = Frame3.Width - 540
        delCommercialInfo.Left = Frame3.Width - 330
        txtCopyrightInfo.Width = Frame3.Width - 1920
        txtAudioURL.Width = Frame3.Width - 1920
        txtArtistURL.Width = Frame3.Width - 3165
        countArtistURL.Left = Frame3.Width - 1335
        prevArtistURL.Left = Frame3.Width - 750
        nextArtistURL.Left = Frame3.Width - 540
        delArtistURL.Left = Frame3.Width - 330
        txtAudioSourceURL.Width = Frame3.Width - 1920
        txtInternetRadioStationURL.Width = Frame3.Width - 1920
        txtPaymentURL.Width = Frame3.Width - 1920
        txtPublisherURL.Width = Frame3.Width - 1920
        txtEncodedBy.Width = Frame3.Width - 1920
        picArt.Left = Frame2(2).Width \ 2 - (Frame2(2).Height - 840) \ 2
        picArt.Width = Frame2(2).Height - 840 ' Most album art is square
        picArt.Height = Frame2(2).Height - 840
        StretchImage imgArt.Picture
        lblBrowse.Width = picArt.ScaleWidth
        lblBrowse.Top = picArt.Height \ 2 - 247
        Label41.Left = Frame2(2).Width \ 2 - 1800
        Label41.Top = Frame2(2).Height - 735
        cmbImageType.Left = Frame2(2).Width \ 2 - 720
        cmbImageType.Top = Frame2(2).Height - 735
        countArt.Left = Frame2(2).Width \ 2 + 585
        countArt.Top = Frame2(2).Height - 675
        prevArt.Left = Frame2(2).Width \ 2 + 1170
        prevArt.Top = Frame2(2).Height - 735
        nextArt.Left = Frame2(2).Width \ 2 + 1380
        nextArt.Top = Frame2(2).Height - 735
        delArt.Left = Frame2(2).Width \ 2 + 1590
        delArt.Top = Frame2(2).Height - 735
        Label43.Left = Frame2(2).Width \ 2 - 1800
        Label43.Top = Frame2(2).Height - 375
        cmbPictureType.Left = Frame2(2).Width \ 2 - 720
        cmbPictureType.Top = Frame2(2).Height - 375
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    SaveSetting Caption, "Window", "XSize", XSize
    SaveSetting Caption, "Window", "YSize", YSize
    SaveSetting Caption, "Window", "XPos", XPos
    SaveSetting Caption, "Window", "YPos", YPos
    SaveSetting Caption, "Window", "State", myWindowState
    
    With ListView1
        SaveSetting Caption, "Columns", "SortKey", .SortKey
        SaveSetting Caption, "Columns", "SortOrder", .SortOrder
        
        For i = 1 To .ColumnHeaders.Count
            SaveSetting Caption, "ColumnPos", Format$(.ColumnHeaders(i).Position, "00"), i
            SaveSetting Caption, "Columns", Format$(i, "00"), .ColumnHeaders(i).Width
        Next
    End With
    
    SaveSetting Caption, "MP3s", "Directory", Text1.Text
    Unhook
End Sub

Private Sub imgArt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If WithinBounds(imgArt, x, Y) Then
        Select Case Button
            Case vbLeftButton: ImageBrowse
            Case vbRightButton: If ValidateMenu Then PopupMenu mnuArt
        End Select
    End If
End Sub

Private Sub lblBrowse_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If WithinBounds(picArt, x + lblBrowse.Left, Y + lblBrowse.Top) Then
        Select Case Button
            Case vbLeftButton: ImageBrowse
            Case vbRightButton: If ValidateMenu Then PopupMenu mnuArt
        End Select
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim i As Long
    Dim idx As Long
    
    SortLvwOnLong ListView1, ColumnHeader.Index
    ShowListViewColumnHeaderSortIcon ListView1
    EnsureSelVisible ListView1
End Sub

Private Sub ListView1_DblClick()
    If SelectedIndex(ListView1) <> -1 Then
        ShellExecute 0&, "open", Text1 & "\" & ListView1.SelectedItem.Text, vbNullString, vbNullString, SW_SHOW
    End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim ID3 As New clsID3
    Dim HourPart As String
    Dim tempStr As String
    Dim sGenreID As String
    Dim bResort As Boolean
    Dim sItem As ListItem
    Dim bRefresh As Boolean
    Dim idx As Long
    
    bResort = False
    bRefresh = False
    Set sItem = ListView1.SelectedItem
    
    If Dir$(Text1 & "\" & sItem.Text) = "" Then
        ListView1.ListItems.Remove SelectedItemIdx(ListView1) + 1
        Exit Sub
    End If
    
    With ID3
        .FileName = Text1 & "\" & sItem.Text
        ID3Revision = .ID3RevisionV2
        ShowOrHideNecessaryFields
        txtTitle = .Title
        txtArtist = .Artist
        txtAlbum = .Album
        cmbGenre = FormatGenre(ID3, .GenreID, .Genre)
        txtTrackNumber = .TrackNumber
        txtTracksTotal = .TracksTotal
        txtYear = .Year
        txtComments = .Comments
        txtLyrics = .Lyrics
        txtComposer = .Composer
        txtBand = .Band
        txtConductor = .Conductor
        txtInterpretedBy = .InterpretedBy
        txtLyricist = .Lyricist
        txtOriginalArtist = .OriginalArtist
        txtOriginalAlbum = .OriginalAlbum
        txtOriginalFileName = .OriginalFileName
        txtOriginalLyricist = .OriginalLyricist
        txtOriginalReleaseYear = .OriginalReleaseYear
        txtCopyright = .Copyright
        txtFileOwner = .FileOwner
        txtPublisher = .Publisher
        txtInternetRadioStationName = .InternetRadioStationName
        txtInternetRadioStationOwner = .InternetRadioStationOwner
        txtISRC = .ISRC
        txtLanguages = .Languages
        LoadMultiData txtArtistURL, .ArtistURL, S_AURL, countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
        txtCopyrightInfo = .CopyrightInfo
        txtAudioURL = .AudioURL
        LoadMultiData txtCommercialInfo, .CommercialInfo, S_CURL, countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
        txtAudioSourceURL = .AudioSourceURL
        txtInternetRadioStationURL = .InternetRadioURL
        txtPaymentURL = .PaymentURL
        txtPublisherURL = .PublisherURL
        txtEncodedBy = .EncodedBy
        txtBPM = .BeatsPerMinute
        cmbKey = .InitialKey
        txtDiscNumber = .DiscNumber
        txtDiscsTotal = .DiscsTotal
        LoadMultiData picArt, .AttachedPictures, S_APIC, countArt, prevArt, nextArt, delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
    End With
    
    With sItem
        If .SubItems(1) <> ID3.Title Then
            bRefresh = True
            If ListView1.SortKey = 1 Then bResort = True
            .SubItems(1) = ID3.Title
        End If
        
        If .SubItems(2) <> ID3.Artist Then
            bRefresh = True
            If ListView1.SortKey = 2 Then bResort = True
            .SubItems(2) = ID3.Artist
        End If
        
        If .SubItems(3) <> ID3.Album Then
            bRefresh = True
            If ListView1.SortKey = 3 Then bResort = True
            .SubItems(3) = ID3.Album
        End If
        
        tempStr = FormatGenre(ID3, ID3.GenreID, ID3.Genre)
        If .SubItems(4) <> tempStr Then
            bRefresh = True
            If ListView1.SortKey = 4 Then bResort = True
            .SubItems(4) = tempStr
        End If
        
        If .SubItems(5) <> ID3.TrackNumber Then
            bRefresh = True
            If ListView1.SortKey = 5 Then bResort = True
            .SubItems(5) = ID3.TrackNumber
        End If
        
        If .SubItems(6) <> ID3.TracksTotal Then
            bRefresh = True
            If ListView1.SortKey = 6 Then bResort = True
            .SubItems(6) = ID3.TracksTotal
        End If
        
        If .SubItems(7) <> ID3.Year Then
            bRefresh = True
            If ListView1.SortKey = 7 Then bResort = True
            .SubItems(7) = ID3.Year
        End If
        
        tempStr = FormatTime(ID3.length, True)
        If .SubItems(8) <> tempStr Then
            bRefresh = True
            If ListView1.SortKey = 8 Then bResort = True
            .SubItems(8) = tempStr
        End If
        
        tempStr = FormatBitRate(ID3.BitRate, ID3.Encoding, True)
        If .SubItems(9) <> tempStr Then
            bRefresh = True
            If ListView1.SortKey = 9 Then bResort = True
            .SubItems(9) = tempStr
        End If
        
        If .SubItems(10) <> ID3.Comments Then
            bRefresh = True
            If ListView1.SortKey = 10 Then bResort = True
            .SubItems(10) = ID3.Comments
        End If
    End With
    
    If bResort Then
        Resort = True
        SortLvwOnLong ListView1, ListView1.SortKey + 1
        Resort = False
    End If
    
    If bRefresh Then EnsureSelVisible ListView1, True
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        ListView1_DblClick
    End If
End Sub

Private Sub mnuArtItem_Click(Index As Integer)
    On Error Resume Next
    Dim hMem As Long
    Dim mPic As StdPicture
    Dim GPC As GDIPlusCandy
    Dim st As String
    Select Case Index
        Case MNU_COPY
            If OpenClipboard(0) Then
                EmptyClipboard
                hMem = CopyImage(imgArt.Picture.handle, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
                SetClipboardData CF_BITMAP, hMem
                DeleteObject hMem
                CloseClipboard
            End If
        Case MNU_PASTE
            Set mPic = Clipboard.GetData(CF_BITMAP)
            If Not mPic Is Nothing And mPic.handle <> 0 Then
                cmbImageType.ListIndex = 1 + 2 * (cmbImageType.ListCount \ 4)
                cmbImageType.Enabled = True
                cmbPictureType.Enabled = True
                Set GPC = New GDIPlusCandy
                st = GPC.ImageToData(mPic, ImagePNG)
                ArtAddProc ImagePNG, GPC.DataToImage(st), st
                Set GPC = Nothing
            End If
    End Select
End Sub

Private Sub nextArt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(nextArt, x, Y) Then
        NextProc picArt, S_APIC, countArt, prevArt, nextArt, delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
    End If
End Sub

Private Sub nextArtistURL_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(nextArtistURL, x, Y) Then
        NextProc txtArtistURL, S_AURL, countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
    End If
End Sub

Private Sub nextCommercialInfo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(nextCommercialInfo, x, Y) Then
        NextProc txtCommercialInfo, S_CURL, countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
    End If
End Sub

Private Sub picArt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If WithinBounds(picArt, x, Y) Then
        Select Case Button
            Case vbLeftButton: ImageBrowse
            Case vbRightButton: If ValidateMenu Then PopupMenu mnuArt
        End Select
    End If
End Sub

Private Sub prevArt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(prevArt, x, Y) Then
        PrevProc picArt, S_APIC, countArt, prevArt, nextArt, delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
    End If
End Sub

Private Sub prevArtistURL_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(prevArtistURL, x, Y) Then
        PrevProc txtArtistURL, S_AURL, countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
    End If
End Sub

Private Sub prevCommercialInfo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(prevCommercialInfo, x, Y) Then
        PrevProc txtCommercialInfo, S_CURL, countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
    End If
End Sub

Private Sub TabStrip1_Click()
    On Error Resume Next
    Dim i As Long
    For i = 1 To TabStrip1.Tabs.Count
        If Frame2(i - 1).Visible <> TabStrip1.Tabs(i).Selected Then
            Frame2(i - 1).Visible = TabStrip1.Tabs(i).Selected
        End If
    Next
    ShowOrHideNecessaryFields
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    DoEvents
    LoadFileEntries Text1
End Sub

Private Sub txtArtistURL_Change()
    TextProc txtArtistURL, S_AURL, countArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
End Sub

Private Sub txtCommercialInfo_Change()
    TextProc txtCommercialInfo, S_CURL, countCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
End Sub

Private Sub VScroll1_Change()
    Dim FTop As Single: FTop = -CSng(VScroll1.Value) * 360
    Dim FTopMax As Single: FTopMax = -Frame3.Height + Frame2(1).Height
    
    If (VScroll1.Value = VScroll1.Max And FTop > FTopMax) Or FTop < FTopMax Then
        If Frame3.Top <> FTopMax Then Frame3.Top = FTopMax
    Else
        If Frame3.Top <> FTop Then Frame3.Top = FTop
    End If
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub

VERSION 5.00
Begin VB.Form frm_mapa 
   Caption         =   "Mapa de Mesas"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   14655
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   39
         Left            =   13200
         TabIndex        =   49
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   39
            Left            =   360
            Picture         =   "frm_mapa.frx":0000
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   38
         Left            =   11760
         TabIndex        =   48
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   38
            Left            =   360
            Picture         =   "frm_mapa.frx":0C42
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   37
         Left            =   10320
         TabIndex        =   47
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   37
            Left            =   360
            Picture         =   "frm_mapa.frx":1884
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   36
         Left            =   8880
         TabIndex        =   46
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   36
            Left            =   360
            Picture         =   "frm_mapa.frx":24C6
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   35
         Left            =   7440
         TabIndex        =   45
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   35
            Left            =   360
            Picture         =   "frm_mapa.frx":3108
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   34
         Left            =   6000
         TabIndex        =   44
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   34
            Left            =   360
            Picture         =   "frm_mapa.frx":3D4A
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   33
         Left            =   4560
         TabIndex        =   43
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   33
            Left            =   360
            Picture         =   "frm_mapa.frx":498C
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   32
         Left            =   3120
         TabIndex        =   42
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   32
            Left            =   360
            Picture         =   "frm_mapa.frx":55CE
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   31
         Left            =   1680
         TabIndex        =   41
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   31
            Left            =   360
            Picture         =   "frm_mapa.frx":6210
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   30
         Left            =   240
         TabIndex        =   40
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   30
            Left            =   360
            Picture         =   "frm_mapa.frx":6E52
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   29
         Left            =   13200
         TabIndex        =   39
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   29
            Left            =   360
            Picture         =   "frm_mapa.frx":7A94
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   28
         Left            =   11760
         TabIndex        =   38
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   28
            Left            =   360
            Picture         =   "frm_mapa.frx":86D6
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   27
         Left            =   10320
         TabIndex        =   37
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   27
            Left            =   360
            Picture         =   "frm_mapa.frx":9318
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   26
         Left            =   8880
         TabIndex        =   36
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   26
            Left            =   360
            Picture         =   "frm_mapa.frx":9F5A
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   25
         Left            =   7440
         TabIndex        =   35
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   25
            Left            =   360
            Picture         =   "frm_mapa.frx":AB9C
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   24
         Left            =   6000
         TabIndex        =   34
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   24
            Left            =   360
            Picture         =   "frm_mapa.frx":B7DE
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   23
         Left            =   4560
         TabIndex        =   33
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   23
            Left            =   360
            Picture         =   "frm_mapa.frx":C420
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   22
         Left            =   3120
         TabIndex        =   32
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   22
            Left            =   360
            Picture         =   "frm_mapa.frx":D062
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   21
         Left            =   1680
         TabIndex        =   31
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   21
            Left            =   360
            Picture         =   "frm_mapa.frx":DCA4
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   20
         Left            =   240
         TabIndex        =   30
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   20
            Left            =   360
            Picture         =   "frm_mapa.frx":E8E6
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   19
         Left            =   13200
         TabIndex        =   29
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   19
            Left            =   360
            Picture         =   "frm_mapa.frx":F528
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   18
         Left            =   11760
         TabIndex        =   28
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   18
            Left            =   360
            Picture         =   "frm_mapa.frx":1016A
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   17
         Left            =   10320
         TabIndex        =   27
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   17
            Left            =   360
            Picture         =   "frm_mapa.frx":10DAC
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   16
         Left            =   8880
         TabIndex        =   26
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   16
            Left            =   360
            Picture         =   "frm_mapa.frx":119EE
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   15
         Left            =   7440
         TabIndex        =   25
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   15
            Left            =   360
            Picture         =   "frm_mapa.frx":12630
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   14
         Left            =   6000
         TabIndex        =   24
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   14
            Left            =   360
            Picture         =   "frm_mapa.frx":13272
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   13
         Left            =   4560
         TabIndex        =   23
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   13
            Left            =   360
            Picture         =   "frm_mapa.frx":13EB4
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   12
         Left            =   3120
         TabIndex        =   22
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   12
            Left            =   360
            Picture         =   "frm_mapa.frx":14AF6
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   11
         Left            =   1680
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   11
            Left            =   360
            Picture         =   "frm_mapa.frx":15738
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   10
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   10
            Left            =   360
            Picture         =   "frm_mapa.frx":1637A
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   9
         Left            =   13200
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   9
            Left            =   360
            Picture         =   "frm_mapa.frx":16FBC
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   8
         Left            =   11760
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   8
            Left            =   360
            Picture         =   "frm_mapa.frx":17BFE
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   7
         Left            =   10320
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   7
            Left            =   360
            Picture         =   "frm_mapa.frx":18840
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   6
         Left            =   8880
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   6
            Left            =   360
            Picture         =   "frm_mapa.frx":19482
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   5
         Left            =   7440
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   5
            Left            =   360
            Picture         =   "frm_mapa.frx":1A0C4
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   4
         Left            =   6000
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   360
            Picture         =   "frm_mapa.frx":1AD06
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   3
         Left            =   4560
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   360
            Picture         =   "frm_mapa.frx":1B948
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   2
         Left            =   3120
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   360
            Picture         =   "frm_mapa.frx":1C58A
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   1
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   360
            Picture         =   "frm_mapa.frx":1D1CC
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Apt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   360
            Picture         =   "frm_mapa.frx":1DE0E
            Top             =   360
            Width           =   480
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   14655
      Begin VB.OptionButton op_avulsas 
         Caption         =   "Avulsas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton bt_mesa_Avulsa 
         Caption         =   "Mesa Avulsa"
         Height          =   495
         Left            =   11400
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmb_grupo 
         DataField       =   "tipo"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   3855
      End
      Begin VB.CommandButton Bt_Sair 
         Cancel          =   -1  'True
         Caption         =   "Fechar"
         Height          =   495
         Left            =   13080
         TabIndex        =   6
         ToolTipText     =   "Fechar esta Janela"
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton op_livres 
         Caption         =   "Livres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton op_ocup 
         Caption         =   "Ocupadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Op_Todos 
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Tipo :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   6240
         TabIndex        =   8
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5535
      Index           =   1
      Left            =   360
      Top             =   1680
      Width           =   14655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   360
      Top             =   360
      Width           =   14655
   End
End
Attribute VB_Name = "frm_mapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database
Dim Mapa_Tab1 As Recordset  'Mesas
Dim Mapa_Tab2 As Recordset  'AUXILIAR TIPO DE Mesas
Dim Mapa_Tab4 As Recordset  'AUXILIAR Mesa

Dim Apt_index As Byte

Private Sub Apt_Click(Index As Integer)

If Apt(Index).Tag = "" Then Exit Sub

Mesa = Apt(Index).Tag
frm_extrato.Show
End Sub

Private Sub bt_mesa_Avulsa_Click()

mesa_avulsa = InputBox("Informe detalhe sobre Mesa Avulsa", "Atenção")
If mesa_avulsa = "" Then Exit Sub

Set Mapa_Tab4 = db1.OpenRecordset("select * from [tbl_Mesas] order by [numero] desc")
mesa_avulsa_num = Mapa_Tab4!Numero + 1

'cria nova mesa
With Mapa_Tab4
    .AddNew
    !Numero = mesa_avulsa_num
    !Tipo = "AVULSA"
    !Status = "O"
    !Observações = mesa_avulsa
    .Update
End With
Mapa_Tab4.Close

Mesa = mesa_avulsa_num
frm_extrato.Show

End Sub

Private Sub Bt_Sair_Click()
Unload Me
If Rotina = "MENU" Then frm_mnu.barramenu.Visible = True
End Sub

Private Sub cmb_grupo_Click()

Call Limpar_Mapa

If cmb_grupo = "TODAS" Then
    Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] order by [Numero]")
Else
    Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] where [Tipo] ='" & cmb_grupo & "' " _
        & "order by [Numero]")
End If

Call Montar_Mapa

End Sub

Private Sub Form_Load()

Set db1 = OpenDatabase(Caminho_Rede & "\dados.mdb")

'atualiza combo de tipos de Mesas
Set Mapa_Tab2 = db1.OpenRecordset("SELECT Tbl_Mesas.tipo From Tbl_Mesas " _
    & "GROUP BY Tbl_Mesas.tipo;")
cmb_grupo.AddItem ("TODAS")
Do While Not Mapa_Tab2.EOF
    cmb_grupo.AddItem ("" & Mapa_Tab2!Tipo)
    Mapa_Tab2.MoveNext
Loop

Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] order by [Numero]")
If Mapa_Tab1.EOF Then bt_mesa_Avulsa.Enabled = False
    
Call Montar_Mapa

End Sub

Sub Limpar_Mapa()

'limpar tela
Apt_index = 0
Do While Apt_index <= 39
    Apt(Apt_index).Caption = ""
    'Apt(Apt_index).BackColor = &H80000005
    Apt(Apt_index).Visible = False
    Apt(Apt_index).Tag = ""
    Apt_index = Apt_index + 1
Loop

End Sub

Sub Montar_Mapa()

On Error Resume Next

Apt_index = 0
Do While Not Mapa_Tab1.EOF

    If Mapa_Tab1!Status = "L" Then Apt(Apt_index).ForeColor = &HC000&: Image1(Apt_index).Picture = LoadPicture(Caminho_Rede & "\midias\mesa.bmp")
    If Mapa_Tab1!Status = "O" Then Apt(Apt_index).ForeColor = &HFF&: Image1(Apt_index).Picture = LoadPicture(Caminho_Rede & "\midias\mesa_red.bmp")
    
    If Mapa_Tab1!Tipo = "AVULSA" Then
        Apt(Apt_index).Caption = Mapa_Tab1!Observações
    Else
        Apt(Apt_index).Caption = Mapa_Tab1!Numero
    End If
    
    Apt(Apt_index).Tag = Mapa_Tab1!Numero
        
    Apt(Apt_index).Visible = True
    
    Mapa_Tab1.MoveNext
    Apt_index = Apt_index + 1
    
    If Apt_index > 39 Then Exit Do

Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_mnu.barramenu.Visible = True
End Sub

Private Sub Image1_Click(Index As Integer)
Call Apt_Click(Index)
End Sub

Private Sub op_avulsas_Click()

'somente mesas avulsas
Call Limpar_Mapa
Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] where [tipo] ='AVULSA' order by [Numero]")
Call Montar_Mapa

End Sub

Private Sub op_livres_Click()

'somente os livres
Call Limpar_Mapa
Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] where [status] ='L' order by [Numero]")
Call Montar_Mapa

End Sub

Private Sub op_ocup_Click()

'somente os ocupados
Call Limpar_Mapa
Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] where [status] ='O' order by [Numero]")
Call Montar_Mapa

End Sub

Private Sub Op_TODOS_Click()

'TODAS as mesas
Call Limpar_Mapa
Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] order by [Numero]")
Call Montar_Mapa

End Sub


Sub Atualiza_Mapa()

'TODAS os aptos
Call Limpar_Mapa
Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] order by [Numero]")
Call Montar_Mapa

End Sub

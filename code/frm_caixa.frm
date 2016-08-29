VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frm_caixa 
   Caption         =   "Caixa"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   14925
   WindowState     =   2  'Maximized
   Begin VB.CommandButton bt_caixa_abrir 
      Caption         =   "Abrir Caixa"
      Enabled         =   0   'False
      Height          =   855
      Left            =   13200
      Picture         =   "frm_caixa.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Abrir Caixa"
      Top             =   7800
      Width           =   1815
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   12938
      _Version        =   327680
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Detalhes"
      TabPicture(0)   =   "frm_caixa.frx":0442
      Tab(0).ControlCount=   5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DBGrid7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Data7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frame_fechar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frame_abertura"
      Tab(0).Control(4).Enabled=   0   'False
      TabCaption(1)   =   "Saídas"
      TabPicture(1)   =   "frm_caixa.frx":045E
      Tab(1).ControlCount=   8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Data1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "bt_saida_excluir"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "bt_saida_novo"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "txt_saida_descricao"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "TXT_saida_valor"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "DBGrid1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblLabels(7)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblLabels(2)"
      Tab(1).Control(7).Enabled=   0   'False
      TabCaption(2)   =   "Recebimentos"
      TabPicture(2)   =   "frm_caixa.frx":047A
      Tab(2).ControlCount=   7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DBGrid2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Data2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "op_receb_din"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "op_receb_cart"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "op_receb_todos"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "op_receb_fiado"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "bt_imprimir_recebimentos"
      Tab(2).Control(6).Enabled=   -1  'True
      TabCaption(3)   =   "Garçons"
      TabPicture(3)   =   "frm_caixa.frx":0496
      Tab(3).ControlCount=   10
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Data8"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "bt_garcon_del"
      Tab(3).Control(1).Enabled=   -1  'True
      Tab(3).Control(2)=   "DBGrid8"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "bt_garcon_ad"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "cmb_garcon"
      Tab(3).Control(4).Enabled=   -1  'True
      Tab(3).Control(5)=   "bt_imprimir_gorjetas"
      Tab(3).Control(5).Enabled=   -1  'True
      Tab(3).Control(6)=   "Data3"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "DBGrid3"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "lblLabels(3)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "lblLabels(1)"
      Tab(3).Control(9).Enabled=   0   'False
      TabCaption(4)   =   "Exclusões"
      TabPicture(4)   =   "frm_caixa.frx":04B2
      Tab(4).ControlCount=   3
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DBGrid5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Data5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "bt_imprimir_exclusões"
      Tab(4).Control(2).Enabled=   -1  'True
      TabCaption(5)   =   "Preços Alt."
      TabPicture(5)   =   "frm_caixa.frx":04CE
      Tab(5).ControlCount=   3
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Data6"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "bt_precos_alt"
      Tab(5).Control(1).Enabled=   -1  'True
      Tab(5).Control(2)=   "DBGrid6"
      Tab(5).Control(2).Enabled=   0   'False
      TabCaption(6)   =   "Caixas Anter."
      TabPicture(6)   =   "frm_caixa.frx":04EA
      Tab(6).ControlCount=   3
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "bt_exibir_caixa"
      Tab(6).Control(0).Enabled=   -1  'True
      Tab(6).Control(1)=   "Data4"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "DBGrid4"
      Tab(6).Control(2).Enabled=   0   'False
      Begin VB.Frame frame_abertura 
         Caption         =   "ABERTURA DE CAIXA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   6495
         Left            =   3720
         TabIndex        =   36
         Top             =   2280
         Visible         =   0   'False
         Width           =   17655
         Begin VB.ComboBox cmb_garcon_abertura 
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
            Left            =   360
            TabIndex        =   70
            Top             =   720
            Width           =   6855
         End
         Begin VB.CommandButton bt_garcon_ad_abertura 
            Caption         =   "Adicionar"
            Height          =   420
            Left            =   7440
            TabIndex        =   69
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton bt_abrir 
            Caption         =   "Confirmar"
            Height          =   495
            Left            =   9120
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Confirmar Abertura de Caixa"
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CommandButton bt_caixa_cancelar 
            Caption         =   "Cancelar"
            Height          =   495
            Left            =   11280
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Confirmar Abertura de Caixa"
            Top             =   1440
            Width           =   1455
         End
         Begin MSMask.MaskEdBox txt_saldo_inicial 
            Height          =   540
            Left            =   9120
            TabIndex        =   39
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   953
            _Version        =   327680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,##0.00;(##,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSDBGrid.DBGrid DBGrid10 
            Bindings        =   "frm_caixa.frx":0506
            Height          =   4815
            Left            =   360
            OleObjectBlob   =   "frm_caixa.frx":0516
            TabIndex        =   68
            Top             =   1440
            Width           =   8415
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Garçons Habilitados :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   17
            Left            =   360
            TabIndex        =   71
            Top             =   480
            Width           =   1515
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Inicial:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   9120
            TabIndex        =   40
            Top             =   480
            Width           =   900
         End
      End
      Begin VB.Frame frame_fechar 
         Caption         =   "ENCERRAMENTO DE CAIXA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   6495
         Left            =   1080
         TabIndex        =   47
         Top             =   960
         Visible         =   0   'False
         Width           =   17655
         Begin VB.Data Data9 
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   240
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2040
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.CommandButton bt_fechar_cancel 
            Caption         =   "Cancelar"
            Height          =   855
            Left            =   15480
            Picture         =   "frm_caixa.frx":0EE9
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Fechar Caixa"
            Top             =   4200
            Width           =   1815
         End
         Begin VB.CommandButton bt_conf 
            Caption         =   "Fechar Caixa"
            Height          =   855
            Left            =   13320
            Picture         =   "frm_caixa.frx":132B
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Fechar Caixa"
            Top             =   4200
            Width           =   1815
         End
         Begin VB.Frame Frame3 
            Caption         =   "RESUMO"
            ForeColor       =   &H000000FF&
            Height          =   3495
            Left            =   9240
            TabIndex        =   49
            Top             =   480
            Width           =   8055
            Begin VB.TextBox txt_din_saldo_resumo 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   5520
               Locked          =   -1  'True
               TabIndex        =   66
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   2640
               Width           =   2295
            End
            Begin VB.TextBox txt_saidas_resumo_total 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   5520
               Locked          =   -1  'True
               TabIndex        =   64
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox txt_saidas_resumo 
               Alignment       =   1  'Right Justify
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
               Left            =   2880
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox txt_total_vendas_resumo 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   5520
               Locked          =   -1  'True
               TabIndex        =   60
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   720
               Width           =   2295
            End
            Begin VB.TextBox txt_garcons_resumo 
               Alignment       =   1  'Right Justify
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
               Left            =   360
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox txt_din_total_resumo 
               Alignment       =   1  'Right Justify
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
               Left            =   360
               Locked          =   -1  'True
               TabIndex        =   51
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   720
               Width           =   2295
            End
            Begin VB.TextBox txt_Card_Total_resumo 
               Alignment       =   1  'Right Justify
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
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   50
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   720
               Width           =   2295
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Saldo final DINHEIRO"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   16
               Left            =   5520
               TabIndex        =   67
               Top             =   2400
               Width           =   1575
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Total de SAIDAS"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   15
               Left            =   5520
               TabIndex        =   65
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Saídas do Caixa:"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   10
               Left            =   2880
               TabIndex        =   63
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Total de VENDAS"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   4
               Left            =   5520
               TabIndex        =   61
               Top             =   480
               Width           =   1290
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Garçons (Diárias e Gorjetas)"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   13
               Left            =   360
               TabIndex        =   55
               Top             =   1440
               Width           =   1980
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Total DINHEIRO"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   12
               Left            =   360
               TabIndex        =   53
               Top             =   480
               Width           =   1200
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Total CARTÃO"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   11
               Left            =   2880
               TabIndex        =   52
               Top             =   480
               Width           =   1065
            End
         End
         Begin MSDBGrid.DBGrid DBGrid9 
            Bindings        =   "frm_caixa.frx":176D
            Height          =   5655
            Left            =   480
            OleObjectBlob   =   "frm_caixa.frx":177D
            TabIndex        =   48
            Top             =   600
            Width           =   8415
         End
      End
      Begin VB.Data Data8 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   -65880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   1250
      End
      Begin VB.CommandButton bt_garcon_del 
         Caption         =   "Excluir"
         Height          =   420
         Left            =   -59040
         TabIndex        =   46
         Top             =   6600
         Width           =   1695
      End
      Begin MSDBGrid.DBGrid DBGrid8 
         Bindings        =   "frm_caixa.frx":2327
         Height          =   4815
         Left            =   -65760
         OleObjectBlob   =   "frm_caixa.frx":2337
         TabIndex        =   45
         Top             =   1560
         Width           =   8415
      End
      Begin VB.CommandButton bt_garcon_ad 
         Caption         =   "Adicionar"
         Height          =   420
         Left            =   -59040
         TabIndex        =   44
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox cmb_garcon 
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
         Left            =   -65760
         TabIndex        =   42
         Top             =   840
         Width           =   6375
      End
      Begin VB.Data Data7 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3240
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBGrid.DBGrid DBGrid7 
         Bindings        =   "frm_caixa.frx":2D09
         Height          =   4935
         Left            =   240
         OleObjectBlob   =   "frm_caixa.frx":2D19
         TabIndex        =   35
         Top             =   2040
         Width           =   17655
      End
      Begin VB.Data Data6 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6480
         Width           =   1250
      End
      Begin VB.CommandButton bt_precos_alt 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   -59280
         TabIndex        =   33
         Top             =   6480
         Width           =   1935
      End
      Begin VB.CommandButton bt_exibir_caixa 
         Caption         =   "Exibir Caixa"
         Height          =   495
         Left            =   -58920
         TabIndex        =   31
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Data Data4 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6600
         Width           =   1250
      End
      Begin VB.CommandButton bt_imprimir_recebimentos 
         Caption         =   "Imprimir"
         Height          =   420
         Left            =   -58920
         TabIndex        =   25
         Top             =   6480
         Width           =   1575
      End
      Begin VB.OptionButton op_receb_fiado 
         Caption         =   "Cred. Próprio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68400
         TabIndex        =   24
         Top             =   6600
         Width           =   2415
      End
      Begin MSDBGrid.DBGrid DBGrid5 
         Bindings        =   "frm_caixa.frx":3A4F
         Height          =   5655
         Left            =   -74640
         OleObjectBlob   =   "frm_caixa.frx":3A5F
         TabIndex        =   23
         Top             =   600
         Width           =   17415
      End
      Begin VB.Data Data5 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6600
         Width           =   1250
      End
      Begin VB.CommandButton bt_imprimir_exclusões 
         Caption         =   "Imprimir"
         Height          =   420
         Left            =   -58800
         TabIndex        =   22
         Top             =   6480
         Width           =   1575
      End
      Begin VB.OptionButton op_receb_todos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -73080
         TabIndex        =   20
         Top             =   6600
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton op_receb_cart 
         Caption         =   "Cartão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69840
         TabIndex        =   19
         Top             =   6600
         Width           =   1815
      End
      Begin VB.OptionButton op_receb_din 
         Caption         =   "Dinheiro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -71520
         TabIndex        =   18
         Top             =   6600
         Width           =   1815
      End
      Begin VB.CommandButton bt_imprimir_gorjetas 
         Caption         =   "Imprimir"
         Height          =   420
         Left            =   -68040
         TabIndex        =   17
         Top             =   6600
         Width           =   1695
      End
      Begin VB.Data Data3 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6600
         Width           =   1250
      End
      Begin VB.Data Data2 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6600
         Width           =   1250
      End
      Begin VB.Data Data1 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6600
         Width           =   1250
      End
      Begin VB.CommandButton bt_saida_excluir 
         Caption         =   "Excluir"
         Enabled         =   0   'False
         Height          =   420
         Left            =   -59040
         TabIndex        =   12
         Top             =   6600
         Width           =   1695
      End
      Begin VB.CommandButton bt_saida_novo 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   420
         Left            =   -63840
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txt_saida_descricao 
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
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   840
         Width           =   9135
      End
      Begin VB.TextBox TXT_saida_valor 
         Alignment       =   1  'Right Justify
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
         Left            =   -65280
         MultiLine       =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Width           =   1335
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_caixa.frx":4771
         Height          =   4935
         Left            =   -74640
         OleObjectBlob   =   "frm_caixa.frx":4781
         TabIndex        =   9
         Top             =   1440
         Width           =   17295
      End
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   17655
         Begin VB.TextBox txt_data_caixa 
            Alignment       =   1  'Right Justify
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
            Left            =   13920
            Locked          =   -1  'True
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   600
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.TextBox txt_total_vendas 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txt_din_saldo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   10440
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txt_saidas 
            Alignment       =   1  'Right Justify
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
            Left            =   7920
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txt_caixa 
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
            Left            =   12480
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   -120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txt_Card_Total 
            Alignment       =   1  'Right Justify
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
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txt_din_total 
            Alignment       =   1  'Right Justify
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
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label lbl_data_caixa 
            AutoSize        =   -1  'True
            Caption         =   "Data do Caixa"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   13920
            TabIndex        =   73
            Top             =   360
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Total de VENDAS"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   14
            Left            =   5280
            TabIndex        =   59
            Top             =   360
            Width           =   1290
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Total em DINHEIRO"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   8
            Left            =   10440
            TabIndex        =   29
            Top             =   360
            Width           =   1905
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Saídas do Caixa:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   9
            Left            =   7920
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Total CARTÃO"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   6
            Left            =   2760
            TabIndex        =   15
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Total DINHEIRO"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1200
         End
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frm_caixa.frx":514F
         Height          =   5655
         Left            =   -74640
         OleObjectBlob   =   "frm_caixa.frx":515F
         TabIndex        =   13
         Top             =   600
         Width           =   17295
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "frm_caixa.frx":5CD1
         Height          =   5775
         Left            =   -74640
         OleObjectBlob   =   "frm_caixa.frx":5CE1
         TabIndex        =   16
         ToolTipText     =   "Duplo clique para selecionar Garçon"
         Top             =   600
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Bindings        =   "frm_caixa.frx":6857
         Height          =   5655
         Left            =   -74640
         OleObjectBlob   =   "frm_caixa.frx":6867
         TabIndex        =   32
         Top             =   600
         Width           =   17295
      End
      Begin MSDBGrid.DBGrid DBGrid6 
         Bindings        =   "frm_caixa.frx":7A8D
         Height          =   5655
         Left            =   -74640
         OleObjectBlob   =   "frm_caixa.frx":7A9D
         TabIndex        =   34
         ToolTipText     =   "Duplo clique para selecionar Garçon"
         Top             =   600
         Width           =   17295
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Garçons Habilitados :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   -65760
         TabIndex        =   43
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "* Duplo clique para imprimir individual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   -73200
         TabIndex        =   41
         Top             =   6600
         Width           =   3360
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   7
         Left            =   -74640
         TabIndex        =   11
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   -65280
         TabIndex        =   10
         Top             =   600
         Width           =   405
      End
   End
   Begin VB.CommandButton bt_conf_ini 
      Caption         =   "Fechar Caixa"
      Enabled         =   0   'False
      Height          =   855
      Left            =   15360
      Picture         =   "frm_caixa.frx":87B7
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Fechar Caixa"
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton Bt_Sair 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   855
      Left            =   17520
      Picture         =   "frm_caixa.frx":8BF9
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Fechar esta Janela"
      Top             =   7800
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   1
      Left            =   13320
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   7335
      Index           =   3
      Left            =   360
      Top             =   360
      Width           =   18135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   0
      Left            =   15480
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   6
      Left            =   17640
      Top             =   7920
      Width           =   855
   End
End
Attribute VB_Name = "frm_caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database
Dim Tab1 As Recordset           'AUXILIAR
Dim Tab2 As Recordset           'caixas
Dim TabErr As Recordset         'auxiliar conexão MySQL
Dim Tab_garcon As Recordset     'garçons
Dim Tab_Gorjetas As Recordset   'auxiliar - gorjetas

'extrato
Dim Tab11 As Recordset   'auxiliar impressao de extrato (Daruma)
Dim extrato_subtotal As Currency
Dim extrato_opcional As Currency
Dim extrato_total As Currency
Dim extrato_valorindiv As Currency

Dim DataCaixa As String
Dim DataFecham As String

Private Sub bt_abrir_Click()

'VALIDAÇÕES
If Not IsNumeric(txt_saldo_inicial) Then MsgBox "Saldo Inicial Inválido", vbExclamation, "Atenção": txt_saldo_inicial.SetFocus: Exit Sub
If Data8.Recordset.EOF Then MsgBox "Informe Garçons habilitados", vbExclamation, "Atenção": Exit Sub

If Conf("Confirma Abertura do Caixa, " & Usuário & " ?", "Atenção") = 7 Then Exit Sub

Me.MousePointer = 11

'cria novo caixa em tabela local
DataCaixa = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
With Tab2
    .AddNew
    !Data = DataCaixa
    !Cod_Operador = Cod_Operador
    !Operador = Usuário
    .Update
    .MoveLast
    NumCaixa = !código
End With

'atualiza número de caixa em tabela de garçons habilitados
db1.Execute "UPDATE tbl_Garcons_habilitados SET tbl_Garcons_habilitados.Caixa = " & NumCaixa & " WHERE (((tbl_Garcons_habilitados.Caixa)=0));"

'lança dados em tabela cloud MySQL
'ConnMySQL_Executar_Instrucao ("INSERT INTO database_foodcontrol.tbl_caixas (ID_Caixa, Data_Caixa, Cod_Operador, Operador, " _
    & "Total_Dinheiro, Total_Cartoes, Total_Saidas, Total_Saldo, Fechado) VALUES ('" & NumCaixa _
    & "', '" & DataCaixa & "', '" & Cod_Operador & "', '" & Usuário & "', '0', '0', '0', '0', '0');")


If CCur(txt_saldo_inicial) <> 0 Then
    'lança saldo inicial
    Set Tab1 = db1.OpenRecordset("select * from [tbl_lancamentos]")
    Dim DataLanc As String
    DataLanc = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
    With Tab1
        .AddNew
        !Data = DataLanc
        !Cod_Operador = Cod_Operador
            
        !Descrição = "SALDO INICIAL DO CAIXA"
        
        !Valor = CCur(txt_saldo_inicial)
        !Quant = 1
        !total = CCur(txt_saldo_inicial)
        !Encerrada = True
        
        !Forma_Pagam = "D"
        !Tipo = "C"
        !caixa = NumCaixa
        
        Dim DataReceb As String
        DataReceb = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
        !Recebimento = Date
        .Update
    End With
    
    'lança pagamento em tabela cloud MySQL
    Dim vtotal As String
    vtotal = LTrim(Str(txt_saldo_inicial))
    'Call ConnMySQL_InserirLançamento(DataLanc, "0", "0", "SALDO INICIAL DO CAIXA", vtotal, "1", vtotal, "", "C", Cod_Operador, "D", NumCaixa, "0", "", DataReceb)

End If

frame_abertura.Visible = False
Me.MousePointer = 0

txt_caixa = NumCaixa
MsgBox "Caixa de número: " & NumCaixa & " aberto para o operador : " & Usuário, vbInformation, "Ok"

Call Carrega_Caixa(NumCaixa)

bt_caixa_abrir.Enabled = False
bt_conf_ini.Enabled = False
bt_saida_novo.Enabled = True
bt_saida_excluir.Enabled = True
Bt_Sair.SetFocus

End Sub

Private Sub bt_caixa_abrir_Click()

Set Tab1 = db1.OpenRecordset("select [código] from [tbl_Caixas] order by [código] desc")
If Not Tab1.EOF Then If Limite_Caixas <> 0 Then If Tab1!código >= Limite_Caixas Then MsgBox "Prazo de Avaliação do Software Atingido!", vbExclamation, "Considere adquirir a versão completa do Software : (71)9341-6896": Exit Sub

txt_saldo_inicial = 0
frame_abertura.Left = 240
frame_abertura.Top = 480
frame_abertura.Visible = True
txt_saldo_inicial.SetFocus

End Sub

Private Sub bt_caixa_cancelar_Click()

db1.Execute "delete * from [tbl_Garcons_habilitados] where [caixa]=0"
Data8.Refresh

txt_saldo_inicial = 0
frame_abertura.Visible = False

End Sub

Private Sub bt_conf_Click()

If CCur(txt_din_saldo_resumo) < 0 Then MsgBox "Saldo insuficente para fechar o caixa", vbExclamation, "Atenção": Exit Sub

If Conf("Confirma Fechamento do Caixa ?", "Atenção") = 7 Then Exit Sub

Me.MousePointer = 11

'lança saidas de diária e gorjetas
With Data9.Recordset
    If Not .EOF Then .MoveFirst
    Do While Not .EOF
        With Data1.Recordset
            .AddNew
            !caixa = NumCaixa
            !Data = Date
            !Valor = Data9.Recordset!Diaria + Data9.Recordset!Gorjetas
            !Descrição = "GARÇON : " & Data9.Recordset!Nome_Garcon
            !Cod_Operador = Cod_Operador
            .Update
        End With
        .MoveNext
    Loop
End With

'registra encerramento de caixa em banco de dados CLOUD mysql
DataFecham = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
Dim vdin, vcard, vsaidas, vsaldo As String
vdin = LTrim(Str(CCur(txt_din_total)))
vcard = LTrim(Str(CCur(txt_Card_Total)))
vsaidas = LTrim(Str(CCur(txt_saidas_resumo_total)))
vsaldo = LTrim(Str(CCur(txt_din_saldo)))
'ConnMySQL_Executar_Instrucao ("UPDATE database_foodcontrol.tbl_caixas SET Cod_Operador='" & Cod_Operador & "', Operador='" & Usuário & "', " _
    & "Total_Dinheiro='" & vdin & "', Total_Cartoes='" & vcard & "', Total_Saidas='" & vsaidas & "', " _
    & "Total_Saldo='" & vsaldo & "', Fechado='1', Data_Fechamento='" & DataFecham & "' WHERE ID_Caixa='" & NumCaixa & "';")

'registra encerramento de caixa em banco de dados local
With Tab2
    .Edit
    !Cod_Operador = Cod_Operador
    !Operador = Usuário
    !Dinheiro = txt_din_total
    !Cartão = txt_Card_Total
    !Saidas = txt_saidas_resumo_total
    !Saldo_Dinheiro = txt_din_saldo
    !fechado = True
    !data_fechamento = DataFecham
    .Update
End With

Me.MousePointer = 0
MsgBox "Caixa Encerrado", vbInformation, "Ok"
NumCaixa = 0
Unload Me
If Rotina = "MENU" Then frm_mnu.barramenu.Visible = True

End Sub

Private Sub bt_conf_ini_Click()

'VERIFICA SE EXISTEM MESAS EM ABERTO
Set Tab1 = db1.OpenRecordset("select * from [tbl_lancamentos] where [Encerrada]=false")
If Not Tab1.EOF Then MsgBox "Existem lançamentos em Aberto. Verifique!", vbExclamation, "Atenção": Exit Sub

'GARÇONS HABILITADOS
With Data8.Recordset
    Do While Not .EOF
        total_gorjetas = 0
        'total de gorjetas do garçon
        Set Tab_Gorjetas = db1.OpenRecordset("SELECT Sum(tbl_Garcons_Gorjetas.Valor) AS SomaDeValor From tbl_Garcons_Gorjetas " _
            & "WHERE (((tbl_Garcons_Gorjetas.Caixa)=" & NumCaixa & ") AND ((tbl_Garcons_Gorjetas.Nome_Garcon)='" & Data8.Recordset!Nome_Garcon & "'));")
        If IsNumeric(Tab_Gorjetas!SomaDeValor) Then total_gorjetas = Tab_Gorjetas!SomaDeValor
        
        marca = .Bookmark
        .Edit
        !Gorjetas = total_gorjetas
        .Update
        .Bookmark = marca
        .MoveNext
    Loop
End With

'GRID com total de diárias e gorjetas
Data9.RecordSource = "select * from [tbl_Garcons_habilitados] where [caixa]=" & NumCaixa & " order by [Nome_Garcon]"
Data9.Refresh

'total geral garçons
Set Tab_Gorjetas = db1.OpenRecordset("SELECT Sum(tbl_Garcons_habilitados.Diaria) AS SomaDeDiaria, Sum(tbl_Garcons_habilitados.Gorjetas) " _
    & "AS SomaDeGorjetas From tbl_Garcons_habilitados WHERE (((tbl_Garcons_habilitados.Caixa)=" & NumCaixa & " ));")
txt_garcons_resumo = Format(Tab_Gorjetas!SomaDeDiaria + Tab_Gorjetas!SomaDeGorjetas, "currency")
If Not IsNumeric(txt_garcons_resumo) Then txt_garcons_resumo = 0

txt_saidas_resumo_total = Format(CCur(txt_saidas_resumo) + CCur(txt_garcons_resumo), "currency")
txt_din_saldo = Format(CCur(txt_din_total) - CCur(txt_saidas_resumo_total), "currency")
txt_din_saldo_resumo = txt_din_saldo

frame_fechar.Left = 240
frame_fechar.Top = 480
frame_fechar.Visible = True

End Sub

Private Sub bt_exibir_caixa_Click()

If Data4.Recordset.EOF Then Exit Sub

Call Carrega_Caixa(Data4.Recordset!código)
txt_caixa = Data4.Recordset!código

bt_caixa_abrir.Enabled = False
bt_conf_ini.Enabled = False
bt_saida_excluir.Enabled = False

DBGrid8.AllowUpdate = False
bt_garcon_del.Enabled = False
bt_garcon_ad.Enabled = False

txt_data_caixa.Text = Data4.Recordset!Data
lbl_data_caixa.Visible = True
txt_data_caixa.Visible = True

SSTab1.Tab = 0

End Sub

Private Sub bt_fechar_cancel_Click()

frame_fechar.Visible = False

End Sub

Private Sub bt_garcon_ad_abertura_Click()

Tab_garcon.FindFirst ("Nome_Garcon = '" & cmb_garcon_abertura & "'")
If Tab_garcon.NoMatch Then MsgBox "Selecione um Garçon já Cadastrado", vbExclamation, "Atenção": cmb_garcon_abertura.SetFocus: Exit Sub

With Data8.Recordset
    .AddNew
    !ID_garcon = Tab_garcon!ID
    !Nome_Garcon = Tab_garcon!Nome_Garcon
    !caixa = NumCaixa
    !Diaria = Tab_garcon!Diaria
    .Update
    .MoveLast
End With

End Sub

Private Sub bt_garcon_ad_Click()

Tab_garcon.FindFirst ("Nome_Garcon = '" & cmb_garcon & "'")
If Tab_garcon.NoMatch Then MsgBox "Selecione um Garçon já Cadastrado", vbExclamation, "Atenção": cmb_garcon.SetFocus: Exit Sub

With Data8.Recordset
    .AddNew
    !ID_garcon = Tab_garcon!ID
    !Nome_Garcon = Tab_garcon!Nome_Garcon
    !caixa = NumCaixa
    !Diaria = Tab_garcon!Diaria
    .Update
End With

End Sub

Private Sub bt_garcon_del_Click()

Call Excluir(Data8)

End Sub

Private Sub bt_imprimir_exclusões_Click()

If Data5.Recordset.EOF Then Exit Sub

With CrystalReport1
    .ReportFileName = Caminho_Rede & "\exclusoes.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .Formulas(2) = ""
    .SelectionFormula = "{tbl_lancamentos_exclusoes.ID_Caixa} = " & txt_caixa
    .Action = 1
End With

End Sub

Private Sub bt_imprimir_gorjetas_Click()

If Data3.Recordset.EOF Then Exit Sub

'With CrystalReport1
'    .ReportFileName = Caminho_Rede & "\gorjetas.rpt"
''    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
'    .Formulas(1) = "endereco = '" & Empresa_End & "'"
'    .Formulas(2) = ""
'    .SelectionFormula = "{tbl_Garcons_Gorjetas.Caixa} = " & txt_caixa
'    .Action = 1
'End With

Call Imprimir_Gorjetas_DARUMA

End Sub

Sub Imprimir_Gorjetas_DARUMA()

'=================== CABEÇALHo

iRetorno = iImprimirTexto_DUAL_DarumaFramework("<e><ce><b>" + Empresa_Nome + "</b></ce></e>", 0)    'expandido,centralizado,negrito
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<ce><b>" + Empresa_End + "</b></ce>", 0)            'centraliado, negrito
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<l></l>", 0)                                        'salta 1 linha

Texto = "RELATORIO DE GORJETAS"
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<b>" + Texto + "</b>", 0)

Texto = "Data / Hora  : " + "<dt></dt><sp>4</sp><hr></hr>"
iRetorno = iImprimirTexto_DUAL_DarumaFramework(Texto, 0)

iRetorno = iImprimirTexto_DUAL_DarumaFramework("<tc>=</tc>", 0)                                     'linha tracejada

'====================  CORPO

gorjetas_total = 0

Set Tab11 = db1.OpenRecordset("select * from [tbl_Garcons_Gorjetas] where [caixa]=" & txt_caixa.Text & " order by [Data],[Nome_Garcon]")

Do While Not Tab11.EOF
    
    Texto = Tab11!Nome_Garcon & " - Valor : " & Format(Tab11!Valor, "Fixed")
    iRetorno = iImprimirTexto_DUAL_DarumaFramework(Texto, 0)

    gorjetas_total = gorjetas_total + Tab11!Valor
        
    Tab11.MoveNext
Loop

'====================  TOTALIZADORES
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<tc>-</tc>", 0)                                     'linha tracejada

Texto = "Total                           => " + Format(gorjetas_total, "fixed")
iRetorno = iImprimirTexto_DUAL_DarumaFramework(Texto, 0)

'====================  RODAPÉ
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<l></l>", 0)                                        'salta 1 linha
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<tc>=</tc>", 0)                                     'linha tracejada
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<ce><c>Desenvolvido por : www.naturaltecnologia.com</c></ce>", 0)                                     'linha tracejada

iRetorno = iImprimirTexto_DUAL_DarumaFramework("<sl>2</sl>", 0)                                     'salta 2 linhas
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<gui></gui>", 0)                                    'aciona guilhotina

End Sub



Private Sub bt_imprimir_recebimentos_Click()

If Data2.Recordset.EOF Then Exit Sub

With CrystalReport1
    .ReportFileName = Caminho_Rede & "\recebimentos.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .Formulas(2) = "Periodo = 'Caixa : " & txt_caixa & "'"
    
    If op_receb_todos.Value = True Then .SelectionFormula = "{tbl_lancamentos.tipo}='C'  and {tbl_lancamentos.Caixa} = " & txt_caixa
    If op_receb_din.Value = True Then .SelectionFormula = "{tbl_lancamentos.Descrição} = 'FECHAMENTO: DINHEIRO' and {tbl_lancamentos.tipo}='C'  and {tbl_lancamentos.Caixa} = " & txt_caixa
    If op_receb_cart.Value = True Then .SelectionFormula = "{tbl_lancamentos.Descrição} like '*CARTÃO*' and {tbl_lancamentos.tipo}='C' and {tbl_lancamentos.Caixa} = " & txt_caixa
    If op_receb_fiado.Value = True Then .SelectionFormula = "{tbl_lancamentos.Descrição} like '*CRED.PROPRIO*' and {tbl_lancamentos.tipo}='C' and {tbl_lancamentos.Caixa} = " & txt_caixa

    .Action = 1
End With

End Sub

Private Sub bt_precos_alt_Click()

If Data6.Recordset.EOF Then Exit Sub

With CrystalReport1
    .ReportFileName = Caminho_Rede & "\precos_alterados.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .Formulas(2) = ""
    .SelectionFormula = "{tbl_lancamentos.caixa} = " & txt_caixa & " and {tbl_lancamentos.preço_alterado}"
    .Action = 1
End With

End Sub

Private Sub bt_saida_excluir_Click()

On Error GoTo Trata_erro

'validações
If Data1.Recordset.EOF Then Exit Sub
If Nível > 1 Then MsgBox "Usuário não autorizado para esta operação", vbExclamation, "Atenção": Exit Sub

If Conf("Confirma Exclusão de Saida: " & Data1.Recordset!Descrição & " ?", "Atenção") = 7 Then Exit Sub

Me.MousePointer = 11

'exclui lançamento de base de dados cloud MYSQL
Dim vsaida As String
vsaida = LTrim(Str(CCur(Data1.Recordset!Valor)))
'ConnMySQL_Executar_Instrucao ("delete from tbl_caixas_saidas where ID_Caixa=" & NumCaixa & " and Valor=" & vsaida _
    & " and Descricao='" & Data1.Recordset!Descrição & "'")

'TOTAL DE DESPESAS e saldo
txt_saidas = Format(CCur(txt_saidas) - Data1.Recordset!Valor, "currency")
txt_din_saldo = Format(CCur(txt_din_total) - CCur(txt_saidas), "currency")

Data1.Recordset.Delete
Me.MousePointer = 0

Exit Sub
Trata_erro:
Exit Sub

End Sub

Private Sub bt_saida_novo_Click()

'validações
If NumCaixa = 0 Then MsgBox "Necessário abrir caixa", vbExclamation, "Atenção": Exit Sub
If Not IsNumeric(TXT_saida_valor) Then MsgBox "Valor Inválido", vbExclamation, "Atenção": TXT_saida_valor.SetFocus: Exit Sub
If txt_saida_descricao = "" Then MsgBox "Informe descrição", vbExclamation, "Atenção": txt_saida_descricao.SetFocus: Exit Sub

Me.MousePointer = 11

'lança saida em base de dados cloud MYSQL
DataCaixa = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
Dim vsaida As String
vsaida = LTrim(Str(CCur(TXT_saida_valor)))
'ConnMySQL_Executar_Instrucao ("INSERT INTO database_foodcontrol.tbl_caixas_saidas (ID_Caixa, Data_Saida, Valor, " _
    & "Descricao, Cod_Operador) VALUES ('" & NumCaixa & "', '" & DataCaixa & "', '" & vsaida & "', '" & txt_saida_descricao & "', '" & Cod_Operador & "');")


'lança saida em base de dados local
With Data1.Recordset
    .AddNew
    !caixa = NumCaixa
    !Data = Date
    !Valor = TXT_saida_valor
    !Descrição = txt_saida_descricao
    !Cod_Operador = Cod_Operador
    .Update
End With


'TOTAL DE saidas e saldo
txt_saidas = Format(CCur(txt_saidas) + CCur(TXT_saida_valor), "currency")
txt_saidas_resumo = txt_saidas

txt_din_saldo = Format(CCur(txt_din_total) - CCur(TXT_saida_valor), "currency")

txt_saida_descricao = ""
TXT_saida_valor = ""
TXT_saida_valor.SetFocus

Me.MousePointer = 0

End Sub

Private Sub Bt_Sair_Click()

Unload Me

End Sub

Private Sub DBGrid3_DblClick()

If Data3.Recordset.EOF Then Exit Sub

With CrystalReport1
    .ReportFileName = Caminho_Rede & "\gorjetas.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .Formulas(2) = ""
    .SelectionFormula = "{tbl_Garcons_Gorjetas.Caixa} = " & txt_caixa & " and {tbl_Garcons_Gorjetas.Id_Garcon} = " & Data3.Recordset!ID_garcon
    .Action = 1
End With

End Sub

Private Sub DBGrid4_DblClick()
Call bt_exibir_caixa_Click
End Sub

Private Sub Form_Load()

Set db1 = OpenDatabase(Caminho_Rede & "\dados.mdb")

'combo garçons
Set Tab_garcon = db1.OpenRecordset("select * from [tbl_Garcons] order by [Nome_Garcon]")
Do While Not Tab_garcon.EOF
    cmb_garcon.AddItem ("" & Tab_garcon!Nome_Garcon)
    cmb_garcon_abertura.AddItem ("" & Tab_garcon!Nome_Garcon)
    Tab_garcon.MoveNext
Loop

txt_caixa = NumCaixa

'-====================='

'GRID RESUMO DE VENDAS
Call Abrir_BD_Data(Data7, "tbl_lancamentos", "Mesa", "[caixa]=0")

'GRID RECEBIMENTOS
Call Abrir_BD_Data(Data2, "tbl_lancamentos", "Mesa", "[caixa]=" & numcx & " and [tipo]='C'")

'GRID SAIDAS
Call Abrir_BD_Data(Data1, "Tbl_Caixas_Saidas", "Data", "[caixa]=" & numcx)

'GRID GORJETAS
Call Abrir_BD_Data(Data3, "tbl_Garcons_Gorjetas", "[Data],[Nome_Garcon]", "[caixa]=" & numcx)

'GRID EXCLUSÕES
Call Abrir_BD_Data(Data5, "tbl_lancamentos_exclusoes", "[Data]", "[id_caixa]=" & numcx)

'GRID PREÇOS ALTERADOS
Call Abrir_BD_Data(Data6, "tbl_lancamentos", "[Data]", "[id_caixa]=" & numcx)

'GRID GARÇONS HABILITADOS
Call Abrir_BD_Data(Data8, "tbl_Garcons_habilitados", "[caixa]", "[caixa]=0")

'total de diárias e gorjetas
Call Abrir_BD_Data(Data9, "tbl_Garcons_habilitados", "[caixa]", "[caixa]=0")

Set Tab2 = db1.OpenRecordset("select * from [Tbl_Caixas] where [fechado]=false")
If Tab2.EOF Then
    bt_caixa_abrir.Enabled = True
    bt_conf_ini.Enabled = False
    bt_saida_novo.Enabled = False
    bt_saida_excluir.Enabled = False
Else
    bt_caixa_abrir.Enabled = False
    bt_conf_ini.Enabled = True
    bt_saida_novo.Enabled = True
    bt_saida_excluir.Enabled = True
End If

'caixas anteriores
Call Abrir_BD_Data(Data4, "tbl_caixas", "[Código] desc", "Fechado=true")

Call Carrega_Caixa(NumCaixa)

End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_mnu.barramenu.Visible = True
End Sub

Sub ConnMySQL_Executar_Instrucao(str_instrucao As String)

On Error GoTo Trata_erro

'declara e inicia conexão
Set conn = New ADODB.Connection
conn.ConnectionString = StringConexao
conn.CursorLocation = adUseClient
conn.Open

'executa instrucao
conn.Execute str_instrucao

'fecha conexão
conn.Close


Exit Sub
Trata_erro:
'-------------------------------------------------------------------------------------------------------------

Me.MousePointer = 0

'dados do erro de envio
cloud_erro = Str$(Err.Number)
Cloud_erro_desc = Err.Description

DataLanc = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")

'salva detalhes do erro bem como a instrução à ser executada
Set TabErr = db1.OpenRecordset("select * from [Tbl_lancamentos_instrucoes] where [Executada]=false")
With TabErr
    .AddNew
    !cloud_erro = cloud_erro
    !Cloud_erro_desc = Cloud_erro_desc
    !Instrucao = str_instrucao
    !Data = DataLanc
    .Update
End With
TabErr.Close
Exit Sub

End Sub

Private Sub op_receb_cart_Click()

'recebimentos em cartão
Data2.RecordSource = "select * from [tbl_lancamentos] where [caixa]=" & txt_caixa & " and [tipo]='C' and [Descrição] like '*CARTÃO*' order by [data]"
Data2.Refresh

End Sub

Private Sub op_receb_din_Click()

'recebimentos em dinheiro
Data2.RecordSource = "select * from [tbl_lancamentos] where [caixa]=" & txt_caixa & " and [tipo]='C' and [forma_pagam] = 'D' order by [data]"
Data2.Refresh

End Sub

Private Sub op_receb_fiado_Click()

'recebimentos CRED.PROPRIO
Data2.RecordSource = "select * from [tbl_lancamentos] where [caixa]=" & txt_caixa & " and [tipo]='C' and [Descrição]='FECHAMENTO: CRED.PROPRIO' order by [data]"
Data2.Refresh

End Sub

Private Sub op_receb_todos_Click()

'todos os recebimentos
Data2.RecordSource = "select * from [tbl_lancamentos] where [caixa]=" & txt_caixa & " and [tipo]='C' order by [data]"
Data2.Refresh

End Sub

Sub Carrega_Caixa(numcx As Long)

'TOTAL RECEBIMENTOS EM DINHEIRO
Set Tab1 = db1.OpenRecordset("SELECT Sum(tbl_lancamentos.Total) AS SomaDeTotal From tbl_lancamentos " _
    & "WHERE (((tbl_lancamentos.Caixa)=" & numcx & ") AND ((tbl_lancamentos.Mesa)<>999) AND ((tbl_lancamentos.Forma_Pagam)='D'));")
If Not IsNumeric(Tab1!SomaDeTotal) Then Txt_Din_Hosp = 0 Else Txt_Din_Hosp = Format(Tab1!SomaDeTotal, "currency")
txt_din_total = Format(CCur(Txt_Din_Hosp), "currency")
txt_din_total_resumo = txt_din_total
'-----

'TOTAL RECEBIMENTOS EM CARTÃO
Set Tab1 = db1.OpenRecordset("SELECT Sum(tbl_lancamentos.Total) AS SomaDeTotal From tbl_lancamentos " _
    & "WHERE (((tbl_lancamentos.Caixa)=" & numcx & ") AND ((tbl_lancamentos.Mesa)<>999) AND ((tbl_lancamentos.Forma_Pagam)='C'));")
If Not IsNumeric(Tab1!SomaDeTotal) Then Txt_card_hosp = 0 Else Txt_card_hosp = Format(Tab1!SomaDeTotal, "currency")
txt_Card_Total = Format(CCur(Txt_card_hosp), "currency")
txt_Card_Total_resumo = txt_Card_Total
'----

'TOTAL DE VENDAS
txt_total_vendas = Format(CCur(txt_din_total) + CCur(txt_Card_Total), "currency")
txt_total_vendas_resumo = txt_total_vendas

'Total de SAIDas
Set Tab1 = db1.OpenRecordset("SELECT Sum(Tbl_Caixas_Saidas.Valor) AS SomaDeValor From Tbl_Caixas_Saidas " _
    & "WHERE (((Tbl_Caixas_Saidas.Caixa)=" & numcx & "));")
If Not IsNumeric(Tab1!SomaDeValor) Then txt_saidas = Format(0, "currency") Else txt_saidas = Format(Tab1!SomaDeValor, "currency")
txt_saidas_resumo = txt_saidas

txt_din_saldo = Format(CCur(txt_din_total) - CCur(txt_saidas), "currency")

'-====================='
'GRID ITENS VENDIDOS
Data7.RecordSource = "SELECT tbl_lancamentos.Descrição, tbl_lancamentos.Valor, Sum(tbl_lancamentos.Quant) AS SomaDeQuant, " _
    & "Sum(tbl_lancamentos.Total) AS SomaDeTotal From tbl_lancamentos Where (((tbl_lancamentos.Caixa) = " & numcx & ") " _
    & "And ((tbl_lancamentos.Tipo) = 'D')) GROUP BY tbl_lancamentos.Descrição, tbl_lancamentos.Valor;"
Data7.Refresh


'GRID RECEBIMENTOS
Data2.RecordSource = "select * from [tbl_lancamentos] where [caixa]=" & numcx & " and [tipo]='C'"
Data2.Refresh

'GRID SAIDAS
Data1.RecordSource = "select * from [Tbl_Caixas_Saidas] where [caixa]=" & numcx & " order by [Data]"
Data1.Refresh

'GRID GORJETAS
Data3.RecordSource = "select * from [tbl_Garcons_Gorjetas] where [caixa]=" & numcx & " order by [Data],[Nome_Garcon]"
Data3.Refresh

'GRID exclusões
Data5.RecordSource = "select * from [tbl_lancamentos_exclusoes] where [id_caixa]=" & numcx & " order by [Data]"
Data5.Refresh

'GRID PREÇOS ALTERADOS
Data6.RecordSource = "select * from [tbl_lancamentos] where [caixa]=" & numcx & " and [preço_alterado]=true order by [Data]"
Data6.Refresh

'GRID GARÇONS HABILITADOS
Data8.RecordSource = "select * from [tbl_Garcons_habilitados] where [caixa]=" & numcx & " order by [Nome_Garcon]"
Data8.Refresh

End Sub

Private Sub txt_saldo_inicial_GotFocus()
Call Selecionar(txt_saldo_inicial)
End Sub


Sub ConnMySQL_InserirLançamento(lancData As String, lancMesa As Long, lancIDPRod As String, lancDesc As String, lancValor As String, lancQuant As String, lancTotal As String, lancObs As String, lancTipo As String, lancCodOperador As Long, lancFormapag As String, lancNumCaixa As Long, lancIDGarcon As String, lancGarcon As String, lancRecebimento As String)

On Error GoTo Trata_erro

cloud_erro = 0
Cloud_erro_desc = ""

'declara e inicia conexão
Set conn = New ADODB.Connection
conn.ConnectionString = StringConexao
conn.CursorLocation = adUseClient
conn.Open

'abre tabela de lançamentos
Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM tbl_lancamentos where ID_caixa=0", conn, adOpenStatic, adLockOptimistic

'insere registro
rs.AddNew
rs!Data_Lancamento = lancData
rs!ID_Mesa = lancMesa
rs!ID_Produto = lancIDPRod
rs!Descricao = lancDesc
rs!Valor = lancValor
rs!Quant = lancQuant
rs!total = lancTotal
rs!Observacoes = lancObs
rs!Tipo = lancTipo
rs!Forma_Pag = lancFormapag
rs!id_caixa = lancNumCaixa
rs!ID_Operador = lancCodOperador
rs!ID_garcon = lancIDGarcon
rs!Garcon = lancGarcon
rs!Recebimento = lancRecebimento
rs!Encerrada = False
rs.Update

'fecha recordset
rs.Close
'fecha conexão
conn.Close

Exit Sub
Trata_erro:
'-------------------------------------------------------------------------------------------------------------

Me.MousePointer = 0

'dados do erro de envio
cloud_erro = Str$(Err.Number)
Cloud_erro_desc = Err.Description

'monta string com instrução a ser executada após ser resolvido problema com conexao
str_instrucao = "INSERT INTO database_foodcontrol.tbl_lancamentos (Data_Lancamento, ID_Mesa,ID_Produto, Descricao, Valor, Quant, Total, " _
    & "Observacoes, Tipo, Forma_Pag, ID_Caixa, ID_Operador, Recebimento, ID_Garcon, Garcon) VALUES " _
    & "('" & lancData & "', '" & lancMesa & "', '" & lancIDPRod & "', '" & lancDesc & "', '" & lancValor & "', '" & lancQuant & "', '" _
    & lancTotal & "', '" & lancObs & "', '" & lancTipo & "', '" & lancFormapag & "', '" & lancNumCaixa & "', '" & lancCodOperador _
    & "', '" & lancRecebimento & "', '" & lancIDGarcon & "', '" & lancGarcon & "');"

'salva detalhes do erro bem como a instrução à ser executada
Set TabErr = db1.OpenRecordset("select * from [Tbl_lancamentos_instrucoes] where [Executada]=false")
With TabErr
    .AddNew
    !cloud_erro = cloud_erro
    !Cloud_erro_desc = Cloud_erro_desc
    !Instrucao = str_instrucao
    !Data = DataLanc
    .Update
End With
TabErr.Close
Exit Sub

End Sub



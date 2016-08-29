VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frm_extrato 
   Caption         =   "Lançar Consumo"
   ClientHeight    =   8565
   ClientLeft      =   255
   ClientTop       =   345
   ClientWidth     =   18360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   18360
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   18375
      _ExtentX        =   32411
      _ExtentY        =   14843
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Mesas"
      TabPicture(0)   =   "frm_extrato.frx":0000
      Tab(0).ControlCount=   8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "bt_mesa_Avulsa"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Bt_Sair"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "op_mesas_todas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "op_mesas_ocupadas"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "op_mesas_livres"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "op_mesas_avulsas"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame5"
      Tab(0).Control(7).Enabled=   0   'False
      TabCaption(1)   =   "Consumo"
      TabPicture(1)   =   "frm_extrato.frx":001C
      Tab(1).ControlCount=   8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frame_Excluir"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "bt_avulsa_del"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "bt_trocar"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "bt_excluir_item"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "Data1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "DBGrid1"
      Tab(1).Control(7).Enabled=   0   'False
      TabCaption(2)   =   "Conta"
      TabPicture(2)   =   "frm_extrato.frx":0038
      Tab(2).ControlCount=   6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frame_parcial"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frame_fiado"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "frame_encerramento2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "frame_encerramento"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "DBGrid4"
      Tab(2).Control(5).Enabled=   0   'False
      Begin VB.Frame frame_parcial 
         Caption         =   "Pagamento Parcial de Conta"
         Height          =   6015
         Left            =   -60120
         TabIndex        =   191
         Top             =   2760
         Visible         =   0   'False
         Width           =   8655
         Begin VB.Frame Frame6 
            Caption         =   "Forma de Pagamento:"
            ForeColor       =   &H00FF0000&
            Height          =   2415
            Left            =   600
            TabIndex        =   196
            Top             =   1680
            Width           =   2775
            Begin VB.OptionButton op_parcial_din 
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
               Height          =   255
               Left            =   360
               TabIndex        =   200
               Top             =   360
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton op_parcial_visa 
               Caption         =   "VISA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   199
               Top             =   840
               Width           =   1335
            End
            Begin VB.OptionButton op_parcial_master 
               Caption         =   "MASTER"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   198
               Top             =   1320
               Width           =   1335
            End
            Begin VB.OptionButton op_parcial_hiper 
               Caption         =   "HIPER"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   197
               Top             =   1800
               Width           =   1335
            End
         End
         Begin VB.CommandButton bt_parcial_calcelar 
            Caption         =   "Cancelar"
            Height          =   615
            Left            =   600
            TabIndex        =   195
            Top             =   4680
            Width           =   1815
         End
         Begin VB.CommandButton bt_parcial_ok 
            Caption         =   "Confirmar"
            Height          =   615
            Left            =   6120
            TabIndex        =   194
            Top             =   4680
            Width           =   1695
         End
         Begin MSMask.MaskEdBox txt_parcial_valor 
            Height          =   540
            Left            =   600
            TabIndex        =   192
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   953
            _Version        =   327680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Valor :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   27
            Left            =   600
            TabIndex        =   193
            Top             =   600
            Width           =   450
         End
      End
      Begin VB.Frame frame_Excluir 
         Caption         =   "EXCLUSÃO DE ITEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   7695
         Left            =   -60000
         TabIndex        =   42
         Top             =   3000
         Visible         =   0   'False
         Width           =   17655
         Begin VB.CommandButton bt_excluir_limpar 
            Caption         =   "X"
            Height          =   420
            Left            =   12000
            TabIndex        =   63
            ToolTipText     =   "Limpar"
            Top             =   6480
            Width           =   495
         End
         Begin VB.TextBox txt_produto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   720
            Width           =   6495
         End
         Begin VB.TextBox txt_quant_Excluir 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   11280
            MaxLength       =   4
            TabIndex        =   47
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton bt_conf_excluir 
            Caption         =   "Confirmar"
            Height          =   495
            Left            =   9000
            TabIndex        =   46
            Top             =   7080
            Width           =   1335
         End
         Begin VB.CommandButton bt_cancelar_Excluir 
            Caption         =   "Cancelar"
            Height          =   495
            Left            =   10560
            TabIndex        =   45
            Top             =   7080
            Width           =   1335
         End
         Begin VB.TextBox txt_motivo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   4680
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   6480
            Width           =   7215
         End
         Begin VB.Data Data2 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   420
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3240
            Visible         =   0   'False
            Width           =   1250
         End
         Begin MSDBGrid.DBGrid DBGrid2 
            Bindings        =   "frm_extrato.frx":0054
            Height          =   4695
            Left            =   4680
            OleObjectBlob   =   "frm_extrato.frx":0064
            TabIndex        =   43
            Top             =   1560
            Width           =   7815
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Produto"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   4680
            TabIndex        =   51
            Top             =   480
            Width           =   555
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Quant. à excluir :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   5
            Left            =   11280
            TabIndex        =   50
            Top             =   480
            Width           =   1200
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Motivo"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   6
            Left            =   4680
            TabIndex        =   49
            Top             =   1320
            Width           =   480
         End
      End
      Begin VB.CommandButton bt_avulsa_del 
         Caption         =   "Excluir Mesa Avulsa"
         Height          =   375
         Left            =   -61560
         TabIndex        =   189
         ToolTipText     =   "Excluir Mesa Avulsa"
         Top             =   7680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame Frame5 
         Height          =   1815
         Left            =   7560
         TabIndex        =   181
         Top             =   6480
         Width           =   6735
         Begin VB.CommandButton bt_rapida 
            Caption         =   "Confirmar Venda Rápida"
            Height          =   420
            Left            =   3120
            TabIndex        =   188
            Top             =   1200
            Width           =   3375
         End
         Begin VB.ComboBox cmb_rapida_formapag 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frm_extrato.frx":0882
            Left            =   240
            List            =   "frm_extrato.frx":0895
            TabIndex        =   186
            Text            =   "DINHEIRO"
            Top             =   1200
            Width           =   2655
         End
         Begin VB.TextBox txt_rapida_quant 
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
            Left            =   5520
            MaxLength       =   3
            TabIndex        =   184
            Text            =   "1"
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cmb_rapida 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            TabIndex        =   182
            Top             =   480
            Width           =   5055
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pagamento"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   26
            Left            =   240
            TabIndex        =   187
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Quant."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   25
            Left            =   5520
            TabIndex        =   185
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "VENDA RÁPIDA"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   24
            Left            =   240
            TabIndex        =   183
            Top             =   240
            Width           =   1200
         End
      End
      Begin VB.OptionButton op_mesas_avulsas 
         Caption         =   "AVULSAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   6720
         Width           =   1575
      End
      Begin VB.OptionButton op_mesas_livres 
         Caption         =   "LIVRES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   6720
         Width           =   1215
      End
      Begin VB.OptionButton op_mesas_ocupadas 
         Caption         =   "OCUPADAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   6720
         Width           =   1575
      End
      Begin VB.OptionButton op_mesas_todas 
         Caption         =   "TODAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   0
         Top             =   6720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton bt_trocar 
         Caption         =   "Transferir Mesa"
         Height          =   375
         Left            =   -63600
         TabIndex        =   106
         ToolTipText     =   "Trocar de Mesa"
         Top             =   7680
         Width           =   1695
      End
      Begin VB.Frame frame_fiado 
         Caption         =   "CREDIÁRIO PRÓPRIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   7695
         Left            =   -59280
         TabIndex        =   52
         Top             =   4320
         Visible         =   0   'False
         Width           =   17655
         Begin VB.TextBox txt_fiado_valor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txt_fiado_cli 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   600
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   1680
            Width           =   7815
         End
         Begin VB.CommandButton bt_fiado_ok 
            Caption         =   "Confirmar"
            Height          =   495
            Left            =   5640
            TabIndex        =   57
            Top             =   3120
            Width           =   1335
         End
         Begin VB.CommandButton bt_fiado_cancel 
            Caption         =   "Cancelar"
            Height          =   495
            Left            =   7080
            TabIndex        =   58
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox txt_fiado_contato 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   420
            Left            =   600
            TabIndex        =   56
            Top             =   2400
            Width           =   7815
         End
         Begin MSMask.MaskEdBox txt_data_receb 
            Height          =   420
            Left            =   2160
            TabIndex        =   54
            Top             =   960
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   741
            _Version        =   327680
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   16
            Left            =   2160
            TabIndex        =   62
            Top             =   720
            Width           =   885
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Valor :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   14
            Left            =   600
            TabIndex        =   61
            Top             =   720
            Width           =   450
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Responsável"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   15
            Left            =   600
            TabIndex        =   60
            Top             =   1440
            Width           =   930
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Contato:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   7
            Left            =   600
            TabIndex        =   59
            Top             =   2160
            Width           =   600
         End
      End
      Begin VB.CommandButton bt_excluir_item 
         Caption         =   "Excluir Item"
         Height          =   375
         Left            =   -65640
         TabIndex        =   99
         ToolTipText     =   "Exclui o Item Selecionado"
         Top             =   7680
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Height          =   2535
         Left            =   -74640
         TabIndex        =   64
         Top             =   600
         Width           =   8655
         Begin VB.CheckBox op_op 
            Caption         =   "10% Opcional"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   360
            TabIndex        =   100
            Top             =   2040
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.TextBox txt_garcon_tab2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   420
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   69
            Top             =   600
            Width           =   6135
         End
         Begin VB.TextBox txt_obs_tab2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   585
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   67
            Top             =   1320
            Width           =   6135
         End
         Begin VB.TextBox txt_mesa_tab2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1260
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Garçon:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   19
            Left            =   2160
            TabIndex        =   70
            Top             =   360
            Width           =   570
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Observações:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   18
            Left            =   2160
            TabIndex        =   68
            Top             =   1080
            Width           =   990
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Mesa :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   66
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.CommandButton Bt_Sair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   975
         Left            =   16800
         Picture         =   "frm_extrato.frx":08C0
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Fechar esta Janela"
         Top             =   6840
         Width           =   1215
      End
      Begin VB.CommandButton bt_mesa_Avulsa 
         Caption         =   "Mesa Avulsa"
         Height          =   975
         Left            =   15240
         Picture         =   "frm_extrato.frx":0D02
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5655
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   17775
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
            Height          =   855
            Index           =   74
            Left            =   16680
            TabIndex        =   180
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   74
               Left            =   200
               Picture         =   "frm_extrato.frx":1144
               Top             =   240
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
            Height          =   855
            Index           =   73
            Left            =   15502
            TabIndex        =   179
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   73
               Left            =   200
               Picture         =   "frm_extrato.frx":1D86
               Top             =   240
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
            Height          =   855
            Index           =   72
            Left            =   14328
            TabIndex        =   178
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   72
               Left            =   200
               Picture         =   "frm_extrato.frx":29C8
               Top             =   240
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
            Height          =   855
            Index           =   71
            Left            =   13154
            TabIndex        =   177
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   71
               Left            =   200
               Picture         =   "frm_extrato.frx":360A
               Top             =   240
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
            Height          =   855
            Index           =   70
            Left            =   11980
            TabIndex        =   176
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   70
               Left            =   200
               Picture         =   "frm_extrato.frx":424C
               Top             =   240
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
            Height          =   855
            Index           =   69
            Left            =   10806
            TabIndex        =   175
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   69
               Left            =   200
               Picture         =   "frm_extrato.frx":4E8E
               Top             =   240
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
            Height          =   855
            Index           =   68
            Left            =   9632
            TabIndex        =   174
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   68
               Left            =   200
               Picture         =   "frm_extrato.frx":5AD0
               Top             =   240
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
            Height          =   855
            Index           =   67
            Left            =   8458
            TabIndex        =   173
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   67
               Left            =   200
               Picture         =   "frm_extrato.frx":6712
               Top             =   240
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
            Height          =   855
            Index           =   66
            Left            =   7284
            TabIndex        =   172
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   66
               Left            =   200
               Picture         =   "frm_extrato.frx":7354
               Top             =   240
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
            Height          =   855
            Index           =   65
            Left            =   6110
            TabIndex        =   171
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   65
               Left            =   200
               Picture         =   "frm_extrato.frx":7F96
               Top             =   240
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
            Height          =   855
            Index           =   64
            Left            =   4936
            TabIndex        =   170
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   64
               Left            =   200
               Picture         =   "frm_extrato.frx":8BD8
               Top             =   240
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
            Height          =   855
            Index           =   63
            Left            =   3762
            TabIndex        =   169
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   63
               Left            =   200
               Picture         =   "frm_extrato.frx":981A
               Top             =   240
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
            Height          =   855
            Index           =   62
            Left            =   2588
            TabIndex        =   168
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   62
               Left            =   200
               Picture         =   "frm_extrato.frx":A45C
               Top             =   240
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
            Height          =   855
            Index           =   61
            Left            =   1414
            TabIndex        =   167
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   61
               Left            =   200
               Picture         =   "frm_extrato.frx":B09E
               Top             =   240
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
            Height          =   855
            Index           =   60
            Left            =   240
            TabIndex        =   166
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   60
               Left            =   200
               Picture         =   "frm_extrato.frx":BCE0
               Top             =   240
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
            Height          =   855
            Index           =   53
            Left            =   9632
            TabIndex        =   165
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   53
               Left            =   200
               Picture         =   "frm_extrato.frx":C922
               Top             =   240
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
            Height          =   855
            Index           =   59
            Left            =   16680
            TabIndex        =   164
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   59
               Left            =   200
               Picture         =   "frm_extrato.frx":D564
               Top             =   240
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
            Height          =   855
            Index           =   58
            Left            =   15502
            TabIndex        =   163
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   58
               Left            =   200
               Picture         =   "frm_extrato.frx":E1A6
               Top             =   240
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
            Height          =   855
            Index           =   57
            Left            =   14328
            TabIndex        =   162
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   57
               Left            =   200
               Picture         =   "frm_extrato.frx":EDE8
               Top             =   240
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
            Height          =   855
            Index           =   56
            Left            =   13154
            TabIndex        =   161
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   56
               Left            =   200
               Picture         =   "frm_extrato.frx":FA2A
               Top             =   240
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
            Height          =   855
            Index           =   55
            Left            =   11980
            TabIndex        =   160
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   55
               Left            =   200
               Picture         =   "frm_extrato.frx":1066C
               Top             =   240
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
            Height          =   855
            Index           =   54
            Left            =   10806
            TabIndex        =   159
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   54
               Left            =   200
               Picture         =   "frm_extrato.frx":112AE
               Top             =   240
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
            Height          =   855
            Index           =   52
            Left            =   8458
            TabIndex        =   158
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   52
               Left            =   200
               Picture         =   "frm_extrato.frx":11EF0
               Top             =   240
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
            Height          =   855
            Index           =   51
            Left            =   7284
            TabIndex        =   157
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   51
               Left            =   200
               Picture         =   "frm_extrato.frx":12B32
               Top             =   240
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
            Height          =   855
            Index           =   50
            Left            =   6110
            TabIndex        =   156
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   50
               Left            =   200
               Picture         =   "frm_extrato.frx":13774
               Top             =   240
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
            Height          =   855
            Index           =   49
            Left            =   4936
            TabIndex        =   155
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   49
               Left            =   200
               Picture         =   "frm_extrato.frx":143B6
               Top             =   240
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
            Height          =   855
            Index           =   48
            Left            =   3762
            TabIndex        =   154
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   48
               Left            =   200
               Picture         =   "frm_extrato.frx":14FF8
               Top             =   240
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
            Height          =   855
            Index           =   47
            Left            =   2588
            TabIndex        =   153
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   47
               Left            =   200
               Picture         =   "frm_extrato.frx":15C3A
               Top             =   240
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
            Height          =   855
            Index           =   46
            Left            =   1414
            TabIndex        =   152
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   46
               Left            =   200
               Picture         =   "frm_extrato.frx":1687C
               Top             =   240
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
            Height          =   855
            Index           =   45
            Left            =   240
            TabIndex        =   151
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   45
               Left            =   200
               Picture         =   "frm_extrato.frx":174BE
               Top             =   240
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
            Height          =   855
            Index           =   44
            Left            =   16680
            TabIndex        =   150
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   44
               Left            =   200
               Picture         =   "frm_extrato.frx":18100
               Top             =   240
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
            Height          =   855
            Index           =   43
            Left            =   15502
            TabIndex        =   149
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   43
               Left            =   200
               Picture         =   "frm_extrato.frx":18D42
               Top             =   240
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
            Height          =   855
            Index           =   42
            Left            =   14328
            TabIndex        =   148
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   42
               Left            =   200
               Picture         =   "frm_extrato.frx":19984
               Top             =   240
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
            Height          =   855
            Index           =   41
            Left            =   13154
            TabIndex        =   147
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   41
               Left            =   200
               Picture         =   "frm_extrato.frx":1A5C6
               Top             =   240
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
            Height          =   855
            Index           =   40
            Left            =   11980
            TabIndex        =   146
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   40
               Left            =   200
               Picture         =   "frm_extrato.frx":1B208
               Top             =   240
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
            Height          =   855
            Index           =   39
            Left            =   10806
            TabIndex        =   145
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   39
               Left            =   200
               Picture         =   "frm_extrato.frx":1BE4A
               Top             =   240
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
            Height          =   855
            Index           =   38
            Left            =   9632
            TabIndex        =   144
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   38
               Left            =   200
               Picture         =   "frm_extrato.frx":1CA8C
               Top             =   240
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
            Height          =   855
            Index           =   37
            Left            =   8458
            TabIndex        =   143
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   37
               Left            =   200
               Picture         =   "frm_extrato.frx":1D6CE
               Top             =   240
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
            Height          =   855
            Index           =   36
            Left            =   7284
            TabIndex        =   142
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   36
               Left            =   200
               Picture         =   "frm_extrato.frx":1E310
               Top             =   240
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
            Height          =   855
            Index           =   35
            Left            =   6110
            TabIndex        =   141
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   35
               Left            =   200
               Picture         =   "frm_extrato.frx":1EF52
               Top             =   240
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
            Height          =   855
            Index           =   34
            Left            =   4936
            TabIndex        =   140
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   34
               Left            =   200
               Picture         =   "frm_extrato.frx":1FB94
               Top             =   240
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
            Height          =   855
            Index           =   33
            Left            =   3762
            TabIndex        =   139
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   33
               Left            =   200
               Picture         =   "frm_extrato.frx":207D6
               Top             =   240
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
            Height          =   855
            Index           =   32
            Left            =   2588
            TabIndex        =   138
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   32
               Left            =   200
               Picture         =   "frm_extrato.frx":21418
               Top             =   240
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
            Height          =   855
            Index           =   31
            Left            =   1414
            TabIndex        =   137
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   31
               Left            =   200
               Picture         =   "frm_extrato.frx":2205A
               Top             =   240
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
            Height          =   855
            Index           =   30
            Left            =   240
            TabIndex        =   136
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   30
               Left            =   200
               Picture         =   "frm_extrato.frx":22C9C
               Top             =   240
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
            Height          =   855
            Index           =   29
            Left            =   16680
            TabIndex        =   135
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   29
               Left            =   200
               Picture         =   "frm_extrato.frx":238DE
               Top             =   240
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
            Height          =   855
            Index           =   28
            Left            =   15502
            TabIndex        =   134
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   28
               Left            =   200
               Picture         =   "frm_extrato.frx":24520
               Top             =   240
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
            Height          =   855
            Index           =   27
            Left            =   14328
            TabIndex        =   133
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   27
               Left            =   200
               Picture         =   "frm_extrato.frx":25162
               Top             =   240
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
            Height          =   855
            Index           =   26
            Left            =   13154
            TabIndex        =   132
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   26
               Left            =   200
               Picture         =   "frm_extrato.frx":25DA4
               Top             =   240
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
            Height          =   855
            Index           =   25
            Left            =   11980
            TabIndex        =   131
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   25
               Left            =   200
               Picture         =   "frm_extrato.frx":269E6
               Top             =   240
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
            Height          =   855
            Index           =   24
            Left            =   10806
            TabIndex        =   130
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   24
               Left            =   200
               Picture         =   "frm_extrato.frx":27628
               Top             =   240
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
            Height          =   855
            Index           =   23
            Left            =   9632
            TabIndex        =   129
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   23
               Left            =   200
               Picture         =   "frm_extrato.frx":2826A
               Top             =   240
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
            Height          =   855
            Index           =   22
            Left            =   8458
            TabIndex        =   128
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   22
               Left            =   200
               Picture         =   "frm_extrato.frx":28EAC
               Top             =   240
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
            Height          =   855
            Index           =   21
            Left            =   7284
            TabIndex        =   127
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   21
               Left            =   200
               Picture         =   "frm_extrato.frx":29AEE
               Top             =   240
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
            Height          =   855
            Index           =   20
            Left            =   6110
            TabIndex        =   126
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   20
               Left            =   200
               Picture         =   "frm_extrato.frx":2A730
               Top             =   240
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
            Height          =   855
            Index           =   19
            Left            =   4936
            TabIndex        =   125
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   19
               Left            =   200
               Picture         =   "frm_extrato.frx":2B372
               Top             =   240
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
            Height          =   855
            Index           =   18
            Left            =   3762
            TabIndex        =   124
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   18
               Left            =   200
               Picture         =   "frm_extrato.frx":2BFB4
               Top             =   240
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
            Height          =   855
            Index           =   17
            Left            =   2588
            TabIndex        =   123
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   17
               Left            =   200
               Picture         =   "frm_extrato.frx":2CBF6
               Top             =   240
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
            Height          =   855
            Index           =   16
            Left            =   1414
            TabIndex        =   122
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   16
               Left            =   200
               Picture         =   "frm_extrato.frx":2D838
               Top             =   240
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
            Height          =   855
            Index           =   15
            Left            =   240
            TabIndex        =   121
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   15
               Left            =   200
               Picture         =   "frm_extrato.frx":2E47A
               Top             =   240
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
            Height          =   855
            Index           =   14
            Left            =   16680
            TabIndex        =   120
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   14
               Left            =   200
               Picture         =   "frm_extrato.frx":2F0BC
               Top             =   240
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
            Height          =   855
            Index           =   13
            Left            =   15502
            TabIndex        =   119
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   13
               Left            =   200
               Picture         =   "frm_extrato.frx":2FCFE
               Top             =   240
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
            Height          =   855
            Index           =   12
            Left            =   14328
            TabIndex        =   118
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   12
               Left            =   200
               Picture         =   "frm_extrato.frx":30940
               Top             =   240
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
            Height          =   855
            Index           =   11
            Left            =   13154
            TabIndex        =   117
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   11
               Left            =   200
               Picture         =   "frm_extrato.frx":31582
               Top             =   240
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
            Height          =   855
            Index           =   10
            Left            =   11980
            TabIndex        =   116
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   10
               Left            =   200
               Picture         =   "frm_extrato.frx":321C4
               Top             =   240
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
            Height          =   855
            Index           =   9
            Left            =   10806
            TabIndex        =   115
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   9
               Left            =   200
               Picture         =   "frm_extrato.frx":32E06
               Top             =   240
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
            Height          =   855
            Index           =   8
            Left            =   9632
            TabIndex        =   114
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   8
               Left            =   200
               Picture         =   "frm_extrato.frx":33A48
               Top             =   240
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
            Height          =   855
            Index           =   7
            Left            =   8458
            TabIndex        =   113
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   7
               Left            =   200
               Picture         =   "frm_extrato.frx":3468A
               Top             =   240
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
            Height          =   855
            Index           =   6
            Left            =   7284
            TabIndex        =   112
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   6
               Left            =   200
               Picture         =   "frm_extrato.frx":352CC
               Top             =   240
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
            Height          =   855
            Index           =   5
            Left            =   6110
            TabIndex        =   111
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   5
               Left            =   200
               Picture         =   "frm_extrato.frx":35F0E
               Top             =   240
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
            Height          =   855
            Index           =   4
            Left            =   4936
            TabIndex        =   110
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   4
               Left            =   200
               Picture         =   "frm_extrato.frx":36B50
               Top             =   240
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
            Height          =   855
            Index           =   3
            Left            =   3762
            TabIndex        =   109
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   3
               Left            =   200
               Picture         =   "frm_extrato.frx":37792
               Top             =   240
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
            Height          =   855
            Index           =   2
            Left            =   2588
            TabIndex        =   108
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   2
               Left            =   200
               Picture         =   "frm_extrato.frx":383D4
               Top             =   240
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
            Height          =   855
            Index           =   1
            Left            =   1414
            TabIndex        =   107
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   1
               Left            =   200
               Picture         =   "frm_extrato.frx":39016
               Top             =   240
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
            Height          =   855
            Index           =   0
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            Begin VB.Image Image1 
               Height          =   480
               Index           =   0
               Left            =   200
               Picture         =   "frm_extrato.frx":39C58
               Top             =   240
               Width           =   480
            End
         End
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   -65160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4800
         Visible         =   0   'False
         Width           =   1250
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   7695
         Left            =   -74640
         TabIndex        =   20
         Top             =   480
         Width           =   8535
         Begin VB.OptionButton op_grupo4 
            Caption         =   "AGUA"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   6960
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton op_grupo3 
            Caption         =   "TIRA-GOSTO"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   5280
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton op_grupo2 
            Caption         =   "REFRIGERANTE"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3240
            TabIndex        =   6
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton op_grupo1 
            Caption         =   "BEBIDAS"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1680
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton op_todas 
            Caption         =   "TODOS"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.CommandButton bt_quant_exc 
            Caption         =   "-"
            Height          =   255
            Left            =   5040
            TabIndex        =   102
            Top             =   7320
            Width           =   255
         End
         Begin VB.CommandButton bt_quant_ad 
            Caption         =   "+"
            Height          =   255
            Left            =   5040
            TabIndex        =   101
            Top             =   6960
            Width           =   255
         End
         Begin VB.CommandButton bt_desc 
            Caption         =   "D-"
            Height          =   420
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Desconto"
            Top             =   7080
            Width           =   420
         End
         Begin VB.CommandButton bt_acrescimo 
            Caption         =   "A+"
            Height          =   420
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Acréscimo"
            Top             =   7080
            Width           =   420
         End
         Begin VB.ComboBox cmb_prod 
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
            TabIndex        =   27
            Top             =   6360
            Width           =   6495
         End
         Begin VB.CommandButton bt_lançar 
            Caption         =   "Lançar"
            Height          =   1155
            Left            =   6960
            Picture         =   "frm_extrato.frx":3A89A
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Lançar Consumo"
            Top             =   6360
            Width           =   1335
         End
         Begin VB.TextBox txt_quant 
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
            Left            =   4080
            MaxLength       =   4
            TabIndex        =   25
            Top             =   7080
            Width           =   975
         End
         Begin VB.TextBox txt_und 
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
            MaxLength       =   5
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   7080
            Width           =   975
         End
         Begin VB.TextBox txt_preço 
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
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   7080
            Width           =   1335
         End
         Begin VB.TextBox txt_total 
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
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   7080
            Width           =   1215
         End
         Begin VB.Data Data3 
            Caption         =   "Data3"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSDBGrid.DBGrid DBGrid3 
            Bindings        =   "frm_extrato.frx":3ACDC
            Height          =   5415
            Left            =   240
            OleObjectBlob   =   "frm_extrato.frx":3ACEC
            TabIndex        =   21
            Top             =   600
            Width           =   8055
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Produto"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   34
            Top             =   6120
            Width           =   555
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Quant."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   12
            Left            =   4080
            TabIndex        =   33
            Top             =   6840
            Width           =   480
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Preço:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   11
            Left            =   1560
            TabIndex        =   32
            Top             =   6840
            Width           =   465
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Unidade:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   31
            Top             =   6840
            Width           =   645
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   8
            Left            =   5520
            TabIndex        =   30
            Top             =   6840
            Width           =   405
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3135
         Left            =   -65640
         TabIndex        =   12
         Top             =   480
         Width           =   8655
         Begin VB.TextBox txt_opcional_resumo 
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
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   3960
            Locked          =   -1  'True
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   103
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txt_saldo_taxa 
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
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   5400
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1575
         End
         Begin VB.TextBox txt_saldo 
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
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1575
         End
         Begin VB.CommandButton bt_mapa 
            Caption         =   "Mesas"
            Height          =   975
            Left            =   7200
            Picture         =   "frm_extrato.frx":3B85A
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Todas as Mesas"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CommandButton bt_conta 
            Caption         =   "Pagamento"
            Height          =   975
            Left            =   240
            Picture         =   "frm_extrato.frx":3C49C
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Encerrar Conta"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txt_id_garcon 
            Height          =   495
            Left            =   4080
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   735
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
            Left            =   2160
            TabIndex        =   15
            Top             =   480
            Width           =   6135
         End
         Begin VB.TextBox txt_obs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   2160
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   1200
            Width           =   6135
         End
         Begin VB.TextBox txt_mesa 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1260
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Total da Conta:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   9
            Left            =   5400
            TabIndex        =   91
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "10% Gorjeta"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   20
            Left            =   3960
            TabIndex        =   89
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   2160
            TabIndex        =   88
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Garçon:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   19
            Top             =   240
            Width           =   570
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Observações:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   2160
            TabIndex        =   18
            Top             =   960
            Width           =   990
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Mesa :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame frame_encerramento2 
         Height          =   1335
         Left            =   -74640
         TabIndex        =   11
         Top             =   6840
         Width           =   8655
         Begin VB.TextBox txt_saldo_taxa_resumo 
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
            ForeColor       =   &H80000012&
            Height          =   555
            Left            =   6480
            Locked          =   -1  'True
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   95
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txt_saldo_resumo 
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
            ForeColor       =   &H80000012&
            Height          =   555
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   93
            Top             =   600
            Width           =   1815
         End
         Begin MSMask.MaskEdBox txt_pagam_parcial 
            Height          =   555
            Left            =   4320
            TabIndex        =   201
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   979
            _Version        =   327680
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt_opcional 
            Height          =   555
            Left            =   2280
            TabIndex        =   104
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   979
            _Version        =   327680
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "(+) GORJETA:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   22
            Left            =   2280
            TabIndex        =   203
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "(-) PAG. PARCIAL"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   28
            Left            =   4320
            TabIndex        =   202
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "TOTAL DA CONTA"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   23
            Left            =   6480
            TabIndex        =   94
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "SUB-TOTAL"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   92
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame frame_encerramento 
         Caption         =   "Forma de Pagamento :"
         Height          =   7575
         Left            =   -65640
         TabIndex        =   10
         Top             =   600
         Width           =   8655
         Begin VB.CommandButton bt_pag_parcial 
            Caption         =   "Pagamento Parcial"
            Height          =   540
            Left            =   1320
            TabIndex        =   190
            ToolTipText     =   "Pagamento Parcial"
            Top             =   4560
            Width           =   2415
         End
         Begin VB.CommandButton bt_mesas_tab2 
            Caption         =   "Mesas"
            Height          =   975
            Left            =   5280
            Picture         =   "frm_extrato.frx":3D0DE
            Style           =   1  'Graphical
            TabIndex        =   98
            ToolTipText     =   "Todas as Mesas"
            Top             =   6240
            Width           =   1815
         End
         Begin VB.CommandButton bt_encerram_ok 
            Caption         =   "Encerrar"
            Height          =   975
            Left            =   3240
            Picture         =   "frm_extrato.frx":3D520
            Style           =   1  'Graphical
            TabIndex        =   97
            ToolTipText     =   "Encerrar Conta"
            Top             =   6240
            Width           =   1815
         End
         Begin VB.CommandButton bt_imprimir 
            Caption         =   "Extrato"
            Height          =   975
            Left            =   1320
            Picture         =   "frm_extrato.frx":3E162
            Style           =   1  'Graphical
            TabIndex        =   96
            ToolTipText     =   "Imprimir Extrato"
            Top             =   6240
            Width           =   1695
         End
         Begin VB.TextBox txt_total_pago 
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
            ForeColor       =   &H00000000&
            Height          =   555
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   4680
            Width           =   1815
         End
         Begin VB.TextBox txt_troco 
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
            ForeColor       =   &H00000000&
            Height          =   555
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   5400
            Width           =   1815
         End
         Begin VB.CommandButton bt_din 
            Caption         =   "Dinheiro "
            Height          =   540
            Left            =   1320
            TabIndex        =   76
            Top             =   600
            Width           =   3735
         End
         Begin VB.CommandButton bt_visa 
            Caption         =   "VISA"
            Height          =   540
            Left            =   1320
            TabIndex        =   75
            Top             =   2025
            Width           =   3735
         End
         Begin VB.CommandButton bt_master 
            Caption         =   "MASTER"
            Height          =   540
            Left            =   1320
            TabIndex        =   74
            Top             =   2610
            Width           =   3735
         End
         Begin VB.CommandButton bt_hiper 
            Caption         =   "HIPER"
            Height          =   540
            Left            =   1320
            TabIndex        =   73
            Top             =   3195
            Width           =   3735
         End
         Begin VB.CommandButton bt_debito 
            Caption         =   "Cartão de DÉBITO"
            Height          =   540
            Left            =   1320
            TabIndex        =   72
            Top             =   1185
            Width           =   3735
         End
         Begin VB.CommandButton bt_fiado 
            Caption         =   "Cred. Próprio"
            Height          =   540
            Left            =   1320
            TabIndex        =   71
            Top             =   3780
            Width           =   3735
         End
         Begin MSMask.MaskEdBox txt_din 
            Height          =   540
            Left            =   5280
            TabIndex        =   77
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   953
            _Version        =   327680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TXT_VISA 
            Height          =   540
            Left            =   5280
            TabIndex        =   78
            Top             =   2025
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   953
            _Version        =   327680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt_master 
            Height          =   540
            Left            =   5280
            TabIndex        =   79
            Top             =   2610
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   953
            _Version        =   327680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt_hiper 
            Height          =   540
            Left            =   5280
            TabIndex        =   80
            Top             =   3195
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   953
            _Version        =   327680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt_debito 
            Height          =   540
            Left            =   5280
            TabIndex        =   81
            Top             =   1185
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   953
            _Version        =   327680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt_fiado 
            Height          =   540
            Left            =   5280
            TabIndex        =   82
            Top             =   3780
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   953
            _Version        =   327680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TROCO:"
            Height          =   195
            Left            =   4560
            TabIndex        =   105
            Top             =   5400
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TOTAL PAGO"
            Height          =   195
            Left            =   4080
            TabIndex        =   85
            Top             =   4680
            Width           =   1020
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_extrato.frx":3EDA4
         Height          =   3855
         Left            =   -65640
         OleObjectBlob   =   "frm_extrato.frx":3EDB4
         TabIndex        =   35
         Top             =   3720
         Width           =   8655
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Bindings        =   "frm_extrato.frx":3FADA
         Height          =   3375
         Left            =   -74640
         OleObjectBlob   =   "frm_extrato.frx":3FAEA
         TabIndex        =   86
         Top             =   3360
         Width           =   8655
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
End
Attribute VB_Name = "frm_extrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'conexão banco MySQL
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cloud_erro As Long
Dim Cloud_erro_desc As String
Dim TabErr As Recordset
Dim str_instrucao As String

Dim db1 As Database
Dim Tab1 As Recordset   'auxiliar Mesas
Dim Tab2 As Recordset   'auxiliar hospedes
Dim Tab3 As Recordset   'auxiliar totais
Dim Tab4 As Recordset   'auxiliar produtos
Dim Tab5 As Recordset   'auxiliar Lançamentos
Dim Tab6 As Recordset   'auxiliar HORAS adicionais

Dim Tab_FichaTec As Recordset      'auxiliar ficha técnica
Dim Tab_FichaTec_insumo As Recordset      'auxiliar insumos da ficha técnica

Dim Tab7 As Recordset   'auxiliar Troca de MEsa
Dim Tab8 As Recordset   'tempo de ocupação da mesa - arquivo

Dim Tab9 As Recordset   'garcons
Dim Tab10 As Recordset   'gorjeta garcons

Dim Tab12 As Recordset   'registro de exclusões de lançamentos

Dim Tab11 As Recordset   'auxiliar impressao de extrato (Daruma)
Dim extrato_subtotal As Currency
Dim extrato_opcional As Currency
Dim extrato_total As Currency
Dim extrato_valorindiv As Currency

Dim vtemp As Currency   'variavel temporaria

'variaveis temporária para exclusao de itens
Dim vanterior As Currency
Dim vnovo As Currency

Dim Total_CR As Single
Dim Total_DB As Single
Dim Total_PARCIAL As Single


Dim Texto As String
Dim DataLanc As String
Dim DataReceb As String

Dim Preço_Alterado As Boolean

Dim Mapa_Tab1 As Recordset  'Mesas
Dim Mapa_Tab2 As Recordset  'AUXILIAR TIPO DE Mesas
Dim Mapa_Tab4 As Recordset  'AUXILIAR Mesa

Dim Apt_index As Byte
Dim Carregando As Boolean

Private Sub bt_acrescimo_Click()

altera = InputBox("Valor a acrescentar:", "Acréscimo no preço")
If altera = "" Then Exit Sub
If Not IsNumeric(altera) Then MsgBox "Valor Inválido", vbExclamation, "Atenção": Exit Sub

txt_preço = Format(CCur(txt_preço) + CCur(altera), "fixed")
Preço_Alterado = True
Call txt_quant_LostFocus
bt_lançar.SetFocus

End Sub

Private Sub bt_avulsa_del_Click()

If Not Data1.Recordset.EOF Then MsgBox "Já existem itens lançados", vbExclamation, "Atenção": Exit Sub

'apaga mesa avulsa
db1.Execute "delete * from [tbl_Mesas] where [Numero] = " & Mesa

MsgBox "Mesa Excluida com Sucesso", vbInformation, "Ok"
Call bt_mapa_Click

End Sub

Private Sub bt_cancelar_Excluir_Click()

txt_motivo.Text = ""
frame_Excluir.Visible = False
cmb_prod.SetFocus

End Sub

Private Sub bt_conf_excluir_Click()

'validações
If Not IsNumeric(txt_quant_Excluir) Then MsgBox "Quantidade Inválida", vbExclamation, "Atenção": Exit Sub
If Val(txt_quant_Excluir) > Data1.Recordset!SomaDeQuant Then MsgBox "Quantidade superior ao lançado", vbExclamation, "Atenção. Não é possível excluir": Exit Sub
If txt_motivo.Text = "" Then MsgBox "Informe Motivo da Exclusão", vbExclamation, "Atenção. Não é possível excluir": txt_motivo.SetFocus: Exit Sub

If Conf("Confirma Exclusão", "Atenção") = 7 Then Exit Sub

'diminui a quantidade informada
If Data1.Recordset!SomaDeQuant = 1 Then
    db1.Execute "delete * from [tbl_lancamentos] where [mesa] = " & Mesa & " and [Descrição] ='" & Data1.Recordset!Descrição & "' " _
        & "and [encerrada]=false and [quant]=1"
Else
    Tab5.FindFirst ("Descrição = '" & Data1.Recordset!Descrição & "'")
    If Not Tab5.NoMatch Then
        If Val(txt_quant_Excluir) > Tab5!Quant Then
            MsgBox "Excluir itens individualmente", vbExclamation, "Não é possível excluir esta quantidade de uma só vez"
            Exit Sub
        Else
            If Val(txt_quant_Excluir) = Tab5!Quant Then
                Tab5.Delete
            Else
                With Tab5
                    vquant = !Quant - Val(txt_quant_Excluir)
                    vnovo = vquant * !valor
                    .Edit
                    !Quant = vquant
                    !total = vquant * !valor
                    .Update
                End With
            End If
        End If
    End If
End If

'registra fato em tabela de exclusões
With Tab12
    .AddNew
    !Data = Date
    !Descrição = Data1.Recordset!Descrição
    !Quant = txt_quant_Excluir
    !motivo = txt_motivo
    !Usuário = Usuário
    !id_caixa = NumCaixa
    
    !Mesa = Mesa
    !ID_garcon = txt_id_garcon
    !Garcon = cmb_garcon
    
    .Update
End With

'estorno em ESTOQUE
With Tab4
    .FindFirst ("Descrição = '" & Data1.Recordset!Descrição & "'")
    If .NoMatch Then
        MsgBox "Problemas no Estorno do Estoque", vbExclamation, "Atenção"
    Else
        .Edit
        !estoque = !estoque + txt_quant_Excluir
        .Update
    End If
End With

'recalcula totais
txt_saldo = Format(CCur(txt_saldo) - vanterior + vnovo, "fixed")
txt_saldo_resumo = txt_saldo

If op_op.Value = 1 Then txt_opcional.Text = Format(CCur(txt_saldo) * 0.1, "fixed") Else txt_opcional.Text = 0
txt_opcional_resumo = txt_opcional

txt_saldo_taxa = Format(CCur(txt_saldo) + CCur(txt_opcional), "fixed")
txt_saldo_taxa_resumo = txt_saldo_taxa

txt_motivo.Text = ""

Data1.Refresh
Data2.Refresh
Data3.Refresh

frame_Excluir.Visible = False
cmb_prod.SetFocus

End Sub

Private Sub bt_conta_Click()
SSTab1.Tab = 2
End Sub

Private Sub bt_debito_Click()

vtemp = CCur(txt_saldo_taxa) - (CCur(TXT_VISA) + CCur(txt_master) + CCur(txt_hiper) + CCur(txt_din) + CCur(txt_fiado))
If vtemp < 0 Then vtemp = 0
txt_debito.Text = vtemp

End Sub

Private Sub bt_desc_Click()

altera = InputBox("Valor do Desconto:", "Desconto no preço")
If altera = "" Then Exit Sub
If Not IsNumeric(altera) Then MsgBox "Valor Inválido", vbExclamation, "Atenção": Exit Sub

txt_preço = Format(CCur(txt_preço) - CCur(altera), "fixed")
Preço_Alterado = True
Call txt_quant_LostFocus
bt_lançar.SetFocus

End Sub

Private Sub bt_din_Click()

txt_troco.Text = 0
vtemp = CCur(txt_saldo_taxa) - (CCur(TXT_VISA) + CCur(txt_master) + CCur(txt_hiper) + CCur(txt_debito) + CCur(txt_fiado))
If vtemp < 0 Then vtemp = 0
txt_din.Text = vtemp

End Sub

Private Sub bt_encerram_ok_Click()

'validações
If txt_total_pago <> txt_saldo_taxa Then MsgBox "Conferir Valores", vbExclamation, "Atenção": txt_din.SetFocus: Exit Sub
If Data1.Recordset.EOF Then Exit Sub
If txt_opcional <> 0 Then If txt_id_garcon.Text = 0 Then MsgBox "Informe Garçon", vbExclamation, "Atenção": cmb_garcon.SetFocus: Exit Sub

If Conf("Confirma Encerramento ?", "Atenção") = 7 Then Exit Sub

Me.MousePointer = 11

Dim FormaPagGorj As String

'lança pagamento em tabela de lançamentos
If CCur(txt_din) <> 0 Then FormaPagGorj = "DINHEIRO": Call Lança_Pagamento("DINHEIRO", CCur(txt_din) - CCur(txt_troco), "D")
If CCur(TXT_VISA) <> 0 Then FormaPagGorj = "CARTAO VISA": Call Lança_Pagamento("CARTÃO VISA", CCur(TXT_VISA), "C")
If CCur(txt_master) <> 0 Then FormaPagGorj = "CARTAO MASTER": Call Lança_Pagamento("CARTÃO MASTER", CCur(txt_master), "C")
If CCur(txt_hiper) <> 0 Then FormaPagGorj = "CARTAO HIPER": Call Lança_Pagamento("CARTÃO HIPER", CCur(txt_hiper), "C")
If CCur(txt_debito) <> 0 Then FormaPagGorj = "CARTAO DE DÉBITO": Call Lança_Pagamento("CARTÃO DE DÉBITO", CCur(txt_debito), "C")
If CCur(txt_fiado) <> 0 Then FormaPagGorj = "CRED.PROPRIO": Call Lança_Pagamento("CRED.PROPRIO", CCur(txt_fiado), "C")

hsaida = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")

'tempo de duração de ocupação
With Tab8
    .AddNew
    !ID_Mesa = Tab1!Numero
    !abertura = Tab1!abertura
    !Encerramento = hsaida
    !Duracao = DateDiff("n", Tab1!abertura, CDate(hsaida))
    .Update
End With

'gorjeta garçon
If txt_opcional > 0 Then
    With Tab10
        .AddNew
        !Data = Date
        !Mesa = txt_mesa
        !ID_garcon = txt_id_garcon
        !Nome_Garcon = cmb_garcon
        
        If FormaPagGorj = "DINHEIRO" Then
            !valor = txt_opcional
        Else
            !valor = txt_opcional - (txt_opcional * 0.1)
        End If
        
        !caixa = NumCaixa
        !Forma_Pag = FormaPagGorj
        .Update
    End With
    
    'lança gorjeta em tabela cloud MySQL
    Dim vopcional As String
    vopcional = LTrim(Str(CCur(txt_opcional)))
    'Call ConnMySQL_Executar_Instrucao("INSERT INTO database_foodcontrol.tbl_garcons_gorjetas (ID_Garcon, ID_Mesa, ID_Caixa, Data_Gorjeta, Nome_Garcon, " _
        & "Valor, Forma_Pag) VALUES ('" & txt_id_garcon & "', '" & txt_mesa & "', '" & NumCaixa & "', '" & DataLanc & "', '" & cmb_garcon _
        & "', '" & vopcional & "', '" & FormaPagGorj & "');")
    
End If

'apaga mesa caso seja avulsa, caso não seja marca como livre
If Tab1!Tipo = "AVULSA" Then
    Tab1.Delete
Else
    'Altera Status do Mesa
    With Tab1
        .Edit
        !Status = "L"
        !Observações = ""
        .Update
    End With
End If

'marca lançamentos como encerrados
db1.Execute "UPDATE tbl_lancamentos SET tbl_lancamentos.Encerrada = True " _
    & "WHERE (((tbl_lancamentos.Encerrada)=False) AND ((tbl_lancamentos.Mesa)=" & Mesa & "));"

'marca lançamentos como encerrados (Tabela Cloud MySQL)
'Call ConnMySQL_Executar_Instrucao("UPDATE tbl_lancamentos set Encerrada = true where Encerrada=false and ID_Mesa = " & txt_mesa)

Me.MousePointer = 0

MsgBox "Encerramento Efetuado com Sucesso", vbInformation, "Ok"
Call Carrega_Mesa
Call bt_mapa_Click

End Sub

Private Sub bt_excluir_item_Click()

If Data1.Recordset.EOF Then Exit Sub

frame_Excluir.Top = 480
frame_Excluir.Left = 360
frame_Excluir.Visible = True

txt_produto.Text = Data1.Recordset!Descrição

vanterior = Data1.Recordset!SomaDeTotal
txt_quant_Excluir.Text = 1
txt_quant_Excluir.SetFocus

End Sub

Private Sub bt_excluir_limpar_Click()
txt_motivo.Text = ""
txt_motivo.SetFocus
End Sub

Private Sub bt_fiado_cancel_Click()

txt_fiado.Text = 0
txt_fiado_valor.Text = 0
Call Calcula_Total_Pago

txt_fiado_cli.Text = ""
txt_fiado_contato.Text = ""

bt_encerram_ok.Enabled = True
bt_imprimir.Enabled = True
frame_fiado.Visible = False

End Sub

Private Sub bt_fiado_Click()

'If CCur(txt_din) < CCur(txt_opcional) Then MsgBox "Retirar taxa de Gorjeta", vbExclamation, "Atenção": Exit Sub

vtemp = CCur(txt_saldo_taxa) - (CCur(TXT_VISA) + CCur(txt_master) + CCur(txt_hiper) + CCur(txt_debito) + CCur(txt_din))
If vtemp < 0 Then vtemp = 0
txt_fiado.Text = vtemp

'responsável pelo CRED.PROPRIO
bt_encerram_ok.Enabled = False
bt_imprimir.Enabled = False

frame_fiado.Left = 360
frame_fiado.Top = 480
frame_fiado.Visible = True

txt_fiado_cli.Text = ""
txt_fiado_contato.Text = ""
txt_fiado_valor.Text = Format(vtemp, "fixed")

txt_data_receb.Text = Format(Date, "dd/mm/yyyy")
txt_data_receb.SetFocus

End Sub

Private Sub bt_fiado_ok_Click()

If txt_fiado_cli.Text = "" Then MsgBox "Informe Responsável", vbExclamation, "Atenção": txt_fiado_cli.SetFocus: Exit Sub
If txt_fiado_contato.Text = "" Then MsgBox "Informe contato do Responsável. Telefone, celular, etc.", vbExclamation, "Atenção": txt_fiado_contato.SetFocus: Exit Sub
If Not IsDate(txt_data_receb) Then MsgBox "Data de Vencimento Inválida", vbExclamation, "Atenção": txt_data_receb.SetFocus: Exit Sub

bt_encerram_ok.Enabled = True
bt_imprimir.Enabled = True
frame_fiado.Visible = False
Call Calcula_Total_Pago
Call bt_encerram_ok_Click

End Sub

Private Sub bt_hiper_Click()

vtemp = CCur(txt_saldo_taxa) - (CCur(TXT_VISA) + CCur(txt_master) + CCur(txt_din) + CCur(txt_debito) + CCur(txt_fiado))
If vtemp < 0 Then vtemp = 0
txt_hiper.Text = vtemp

End Sub

Private Sub bt_imprimir_Click()

If Data1.Recordset.EOF Then Exit Sub

If ModeloPRinter = "1" Then Call Imprimir_Extrato_DARUMA
If ModeloPRinter = "4" Then Call Imprimir_Extrato_impressora_windows

End Sub

Private Sub bt_lançar_Click()

'validações
If NumCaixa = 0 Then MsgBox "Necessário abrir caixa", vbExclamation, "Atenção": Exit Sub
If cmb_prod.Text = "" Then MsgBox "Informe Produto", vbExclamation, "Atenção": cmb_prod.SetFocus: Exit Sub
If Not IsNumeric(txt_quant) Then MsgBox "Quantidade Incorreta", vbExclamation, "Atenção": txt_quant.SetFocus: Exit Sub
If Not IsNumeric(txt_total) Then MsgBox "Total Incorreto", vbExclamation, "Atenção": Exit Sub
If txt_id_garcon = 0 Then MsgBox "Informe Garcon", vbExclamation, "Atenção": cmb_garcon.SetFocus: Exit Sub

'dados do produto
With Tab4
    .FindFirst ("Descrição = '" & cmb_prod & "'")
    If .NoMatch Then MsgBox "Selecione um produto da lista", vbExclamation, "Atenção": cmb_prod.SetFocus: Exit Sub
    'If !estoque < 0 Then MsgBox "Saldo Insuficiente", vbExclamation, "Atenção"
End With

DataLanc = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")

'desabilita controles enquanto efetua lançamentos em bancos de dados
Call Controles_Desabilita

Me.MousePointer = 11

Dim vpreco As String
Dim vtotal As String

vpreco = LTrim(Str(CCur(txt_preço)))
vtotal = LTrim(Str(CCur(txt_total)))

'lança registros em banco de dados MYSQL
'Call ConnMySQL_InserirLançamento(DataLanc, Mesa, Tab4!código, cmb_prod, vpreco, txt_quant, vtotal, txt_obs, "D", Cod_Operador, "T", NumCaixa, txt_id_garcon, cmb_garcon, DataLanc)

'Altera Status do Mesa
If Total_DB = 0 Then
    With Tab1
        .Edit
        !Status = "O"
        !abertura = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
        .Update
    End With
End If

'cadastra lançamento
With Tab5
    .AddNew
    !Data = DataLanc
    !Mesa = Mesa
    !ID_Produto = Tab4!código
    !Descrição = cmb_prod
    
    !valor = txt_preço
    !Quant = txt_quant
    !total = txt_total
    !obs = txt_obs
    
    !Tipo = "D"
    
    !caixa = NumCaixa
    !Forma_Pagam = "T"
    !Preço_Alterado = Preço_Alterado
    
    !Cod_Operador = Cod_Operador
    !ID_garcon = txt_id_garcon
    !Garcon = cmb_garcon
        
    .Update
End With

'baixa em ESTOQUE
With Tab4
    If !Ficha_Tec = False Then
        .Edit
        !estoque = !estoque - txt_quant
        .Update
    Else
        'itens da ficha técnica
        Set Tab_FichaTec = db1.OpenRecordset("select * from [tbl_Produtos_FichaTec] where [ID_Produto]=" & Tab4!código)
        Do While Not Tab_FichaTec.EOF
            'localiza insumo
            Set Tab_FichaTec_insumo = db1.OpenRecordset("select * from [tbl_Produtos] where [código]=" & Tab_FichaTec!id_insumo)
            If Not Tab_FichaTec_insumo.NoMatch Then
                'baixa no ESTOQUE do insumo
                Tab_FichaTec_insumo.Edit
                Tab_FichaTec_insumo!estoque = Tab_FichaTec_insumo!estoque - Tab_FichaTec!Quant
                Tab_FichaTec_insumo.Update
            End If
            Tab_FichaTec.MoveNext
        Loop
    End If
End With

'TOTAIS
Total_DB = Total_DB + CCur(txt_total)
txt_saldo = Format(Total_DB, "fixed")
txt_saldo_resumo = txt_saldo

If op_op.Value = 1 Then txt_opcional.Text = Format(CCur(txt_saldo) * 0.1, "fixed") Else txt_opcional.Text = 0
txt_opcional_resumo = txt_opcional

txt_saldo_taxa = Format(CCur(txt_saldo) + CCur(txt_opcional), "fixed")
txt_saldo_taxa_resumo = txt_saldo_taxa

txt_und.Text = ""
txt_preço.Text = ""
txt_quant.Text = ""
txt_total.Text = ""

Data1.Refresh
Data3.Refresh

Call Controles_Habilita
Me.MousePointer = 0

If Conf("Lançar OUTRO item para esta Mesa ?", "Lançado: " & cmb_prod.Text & " com Sucesso!") = 7 Then Call bt_mapa_Click

End Sub

Private Sub bt_mapa_Click()

cmb_garcon = ""
txt_id_garcon = ""
cmb_prod.Text = ""
Call cmb_prod_Click

If op_mesas_todas.Value = True Then Call Atualiza_Mapa
If op_mesas_ocupadas.Value = True Then op_mesas_ocupadas_Click
If op_mesas_livres.Value = True Then op_mesas_livres_Click
If op_mesas_avulsas.Value = True Then op_mesas_avulsas_Click

SSTab1.Tab = 0

End Sub

Private Sub bt_master_Click()

vtemp = CCur(txt_saldo_taxa) - (CCur(TXT_VISA) + CCur(txt_din) + CCur(txt_hiper) + CCur(txt_debito) + CCur(txt_fiado))
If vtemp < 0 Then vtemp = 0
txt_master.Text = vtemp

End Sub

Private Sub bt_mesas_tab2_Click()
Call bt_mapa_Click
End Sub

Private Sub bt_pag_parcial_Click()

frame_parcial.Left = 9360
frame_parcial.Top = 600
frame_parcial.Visible = True
txt_parcial_valor.SetFocus

End Sub

Private Sub bt_parcial_calcelar_Click()
frame_parcial.Visible = False
End Sub

Private Sub bt_parcial_ok_Click()

If Not IsNumeric(txt_parcial_valor.Text) Then MsgBox "Valor Inválido", vbExclamation, "Atenção": Exit Sub

If Conf("Confirme Lançamento de Pagamento Parcial", "Atenção") = 7 Then Exit Sub

Me.MousePointer = 11

If op_parcial_din.Value = True Then FormaPagGorj = "DINHEIRO": Call Lança_Pagamento("DINHEIRO", CCur(txt_parcial_valor), "D")
If op_parcial_visa.Value = True Then FormaPagGorj = "CARTAO VISA": Call Lança_Pagamento("CARTÃO VISA", CCur(txt_parcial_valor), "C")
If op_parcial_master.Value = True Then FormaPagGorj = "CARTAO MASTER": Call Lança_Pagamento("CARTÃO MASTER", CCur(txt_parcial_valor), "C")
If op_parcial_hiper.Value = True Then FormaPagGorj = "CARTAO HIPER": Call Lança_Pagamento("CARTÃO HIPER", CCur(txt_parcial_valor), "C")

txt_parcial_valor.Text = 0
frame_parcial.Visible = False

Call Controles_Habilita
Me.MousePointer = 0

MsgBox "Pagamento Parcial Lançado com Sucesso", vbInformation, "Ok"
Call bt_mapa_Click

End Sub

Private Sub bt_quant_ad_Click()

If Not IsNumeric(txt_quant) Then Exit Sub
txt_quant = Val(txt_quant) + 1
Call txt_quant_LostFocus

End Sub

Private Sub bt_quant_exc_Click()

If Not IsNumeric(txt_quant) Then Exit Sub
If Val(txt_quant) > 1 Then txt_quant = Val(txt_quant) - 1
Call txt_quant_LostFocus

End Sub

Private Sub bt_rapida_Click()

'validações
If cmb_rapida.Text = "" Then MsgBox "Informe Produto", vbExclamation, "Atenção": cmb_prod.SetFocus: Exit Sub
If Not IsNumeric(txt_rapida_quant) Then MsgBox "Quantidade Incorreta", vbExclamation, "Atenção": txt_rapida_quant.SetFocus: Exit Sub

If Conf("Confirme Venda Rápida de : " & cmb_rapida, "Atenção") = 7 Then Exit Sub

'dados do produto
With Tab4
    .FindFirst ("Descrição = '" & cmb_rapida & "'")
    If .NoMatch Then MsgBox "Selecione um produto da lista", vbExclamation, "Atenção": cmb_rapida.SetFocus: Exit Sub
    Dim rapida_preco As Currency
    rapida_preco = Tab4!Preço
End With

DataLanc = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")

Me.MousePointer = 11

'cadastra lançamento
Set Tab5 = db1.OpenRecordset("select * from [tbl_lancamentos] where [caixa]=0")
With Tab5
    .AddNew
    !Data = DataLanc
    !Mesa = 0
    !ID_Produto = Tab4!código
    !Descrição = cmb_rapida
    
    !valor = rapida_preco
    !Quant = txt_rapida_quant
    !total = (rapida_preco * txt_rapida_quant)
    !obs = "VENDA RAPIDA"
    
    !Tipo = "D"
    
    !caixa = NumCaixa
    !Forma_Pagam = "T"
    !Encerrada = True
    
    !Cod_Operador = Cod_Operador
        
    .Update
End With

'baixa em ESTOQUE
With Tab4
    If !Ficha_Tec = False Then
        .Edit
        !estoque = !estoque - txt_rapida_quant
        .Update
    Else
        'itens da ficha técnica
        Set Tab_FichaTec = db1.OpenRecordset("select * from [tbl_Produtos_FichaTec] where [ID_Produto]=" & Tab4!código)
        Do While Not Tab_FichaTec.EOF
            'localiza insumo
            Set Tab_FichaTec_insumo = db1.OpenRecordset("select * from [tbl_Produtos] where [código]=" & Tab_FichaTec!id_insumo)
            If Not Tab_FichaTec_insumo.NoMatch Then
                'baixa no ESTOQUE do insumo
                Tab_FichaTec_insumo.Edit
                Tab_FichaTec_insumo!estoque = Tab_FichaTec_insumo!estoque - Tab_FichaTec!Quant
                Tab_FichaTec_insumo.Update
            End If
            Tab_FichaTec.MoveNext
        Loop
    End If
End With

'lança pagamento em tabela de lançamentos
txt_id_garcon = 0
cmb_garcon = ""

If cmb_rapida_formapag = "DINHEIRO" Then Call Lança_Pagamento("DINHEIRO", rapida_preco, "D")
If cmb_rapida_formapag = "DÉBITO" Then Call Lança_Pagamento("CARTÃO DE DÉBITO", rapida_preco, "D")
If cmb_rapida_formapag = "VISA" Then Call Lança_Pagamento("CARTÃO VISA", rapida_preco, "D")
If cmb_rapida_formapag = "MASTER" Then Call Lança_Pagamento("CARTÃO MASTER", rapida_preco, "D")
If cmb_rapida_formapag = "HIPER" Then Call Lança_Pagamento("CARTÃO MASTER", rapida_preco, "D")

txt_rapida_quant.Text = "1"
cmb_rapida_formapag = "DINHEIRO"
cmb_rapida = ""
cmb_rapida.SetFocus

Me.MousePointer = 0
MsgBox "Venda Rápida Lançada com Sucesso", vbInformation, "Ok"


End Sub

Private Sub Bt_Sair_Click()

db1.Close
Unload Me
'If Rotina = "MENU" Then frm_mnu.barramenu.Visible = True
frm_mnu.barramenu.Visible = True

End Sub

Private Sub bt_trocar_Click()

If Data1.Recordset.EOF Then Exit Sub

'validações
novamesa = InputBox("Nova Mesa", "Mudança de Mesa")
If novamesa = "" Then Exit Sub
If novamesa = Mesa Then Exit Sub

Set Tab7 = db1.OpenRecordset("select * from [Tbl_Mesas] where [numero]=" & novamesa)
If Tab7.EOF Then MsgBox "Mesa Inválida: " & novamesa, vbExclamation, "Atenção": Exit Sub

Me.MousePointer = 11

'altera mesa na tabela de Lançamentos
db1.Execute "UPDATE tbl_lancamentos SET tbl_lancamentos.Mesa = " & novamesa _
    & " WHERE (((tbl_lancamentos.Mesa)=" & Mesa & ") AND ((tbl_lancamentos.Encerrada)=False));"

'altera mesa na tabela de Lançamentos (cloud MySQL)
'ConnMySQL_Executar_Instrucao ("UPDATE tbl_lancamentos SET tbl_lancamentos.ID_Mesa = " & novamesa _
    & " WHERE ID_Mesa=" & Mesa & " AND Encerrada=False")


'Altera Status do Mesa para ocupado
db1.Execute "UPDATE Tbl_Mesas SET Tbl_Mesas.Status = 'O' WHERE (((Tbl_Mesas.Numero)=" & novamesa & "));"

'Altera Status do Mesa para LIVRE
With Tab1
    .Edit
    !Status = "L"
    .Update
End With

Me.MousePointer = 0
MsgBox "Troca Efetuada", vbInformation, "Ok"

Mesa = novamesa
Carrega_Mesa
Call Atualiza_Mapa
SSTab1.Tab = 0

End Sub

Private Sub bt_visa_Click()

vtemp = CCur(txt_saldo_taxa) - (CCur(txt_din) + CCur(txt_master) + CCur(txt_hiper) + CCur(txt_debito) + CCur(txt_fiado))
If vtemp < 0 Then vtemp = 0
TXT_VISA.Text = vtemp

End Sub

Private Sub cmb_garcon_LostFocus()

Tab9.FindFirst ("Nome_Garcon = '" & cmb_garcon & "'")
If Not Tab9.NoMatch Then
    txt_id_garcon = Tab9!ID_garcon
Else
    txt_id_garcon = 0
End If

With Tab1
        .Edit
        !ID_garcon = txt_id_garcon.Text
        !Garcon = cmb_garcon.Text
        .Update
End With


End Sub

Private Sub cmb_prod_Click()

'dados do produto
If cmb_prod.Text = "" Then
    txt_quant.Text = 1
    txt_preço.Text = ""
    txt_und.Text = ""
    Exit Sub
End If
Preço_Alterado = False

With Tab4
    .FindFirst ("Descrição = '" & cmb_prod & "'")
    If .NoMatch Then MsgBox "Selecione um produto da lista", vbExclamation, "Atenção": cmb_prod.SetFocus: Exit Sub
    
    txt_und = "" & !unidade
    txt_preço = Format(!Preço, "fixed")
End With

txt_quant.Text = 1
txt_quant.SetFocus
Call txt_quant_LostFocus

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Data2_Reposition()

If Data2.Recordset.EOF Then Exit Sub
txt_motivo = Data2.Recordset!motivo

End Sub

Private Sub Data3_Reposition()

If Carregando = True Then Exit Sub
If Data3.Recordset.EOF Then Exit Sub
cmb_prod.Text = "" & Data3.Recordset!Descrição
Call cmb_prod_Click

End Sub

Private Sub DBGrid3_Click()

Carregando = False

End Sub

Private Sub Form_Load()

Carregando = True

Set db1 = OpenDatabase(Caminho_Rede & "\dados.mdb")

'LANÇAMENTOS do Mesa
Call Abrir_BD_Data(Data1, "tbl_lancamentos", "[data] desc", "[Mesa]=" & Mesa)

'GRID produtos
Call Abrir_BD_Data(Data3, "Tbl_Produtos", "[Grupo],[Descrição]", "[insumo]='N'")

'motivos de exclusão
Data2.DatabaseName = Caminho_Rede & "\dados.mdb"
Data2.RecordSource = "SELECT tbl_lancamentos_exclusoes.Motivo From tbl_lancamentos_exclusoes " _
    & "GROUP BY tbl_lancamentos_exclusoes.Motivo;"

'tempo de ocupação da mesa
Set Tab8 = db1.OpenRecordset("select * from [tbl_Mesas_Duracao]")

'combo Produto
Set Tab4 = db1.OpenRecordset("select * from [Tbl_Produtos] where [insumo]='N' order by [Grupo],[Descrição]")
Do While Not Tab4.EOF
    cmb_prod.AddItem ("" & Tab4!Descrição)
    cmb_rapida.AddItem ("" & Tab4!Descrição)
    Tab4.MoveNext
Loop

'combo Garçons
Set Tab9 = db1.OpenRecordset("select * from [tbl_Garcons_habilitados] WHERE [caixa]=" & NumCaixa & " order by [Nome_Garcon]")
Do While Not Tab9.EOF
    cmb_garcon.AddItem ("" & Tab9!Nome_Garcon)
    Tab9.MoveNext
Loop

'gorjetas
Set Tab10 = db1.OpenRecordset("select * from [tbl_Garcons_Gorjetas] where [Id_Garcon]=0")

'gorjetas
Set Tab10 = db1.OpenRecordset("select * from [tbl_Garcons_Gorjetas] where [Id_Garcon]=0")

'registro de exclusão de lançamentos
Set Tab12 = db1.OpenRecordset("select * from [tbl_lancamentos_exclusoes] where [Quant]=0")

'monta mapa de mesas
Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] order by [Numero]")
If Mapa_Tab1.EOF Then bt_mesa_Avulsa.Enabled = False
Call Montar_Mapa

Total_DB = 0

End Sub

Private Sub op_grupo1_Click()

If op_grupo1.Value = True Then Data3.RecordSource = "select * from [tbl_produtos] where [Grupo]= 'BEBIDAS' order by [descrição]": Data3.Refresh

End Sub

Private Sub op_grupo2_Click()

If op_grupo2.Value = True Then Data3.RecordSource = "select * from [tbl_produtos] where [Grupo]= 'REFRIGERANTE' order by [descrição]": Data3.Refresh

End Sub

Private Sub op_grupo3_Click()

If op_grupo3.Value = True Then Data3.RecordSource = "select * from [tbl_produtos] where [Grupo]= 'TIRA-GOSTO' order by [descrição]": Data3.Refresh

End Sub

Private Sub op_grupo4_Click()

If op_grupo4.Value = True Then Data3.RecordSource = "select * from [tbl_produtos] where [Grupo]= 'AGUA' order by [descrição]": Data3.Refresh

End Sub

Private Sub op_mesas_avulsas_Click()

'somente os livres
Call Limpar_Mapa
Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] where [Tipo] = 'AVULSA' order by [Numero]")
Call Montar_Mapa

End Sub

Private Sub op_mesas_livres_Click()

'somente os livres
Call Limpar_Mapa
Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] where [status] ='L' order by [Numero]")
Call Montar_Mapa

End Sub

Private Sub op_mesas_ocupadas_Click()

'somente os ocupados
Call Limpar_Mapa
Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] where [status] ='O' order by [Numero]")
Call Montar_Mapa


End Sub

Private Sub op_mesas_todas_Click()

'TODAS as mesas
Call Limpar_Mapa
Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] order by [Numero]")
Call Montar_Mapa

End Sub

Private Sub op_op_Click()

If op_op.Value = 1 Then
    txt_opcional.Text = CCur(txt_saldo) * 0.1
Else
    txt_opcional.Text = 0
End If
txt_opcional_resumo = txt_opcional

If IsNumeric(txt_saldo) Then
    txt_saldo_taxa = Format(CCur(txt_saldo) + CCur(txt_opcional), "fixed")
Else
    txt_saldo_taxa = ""
End If
txt_saldo_taxa_resumo = txt_saldo_taxa

Call Calcula_Total_Pago

End Sub

Private Sub op_todas_Click()

If op_todas.Value = True Then Data3.RecordSource = "select * from [tbl_produtos] order by [descrição]": Data3.Refresh

End Sub

Private Sub txt_data_receb_GotFocus()

Call Mask_Data(txt_data_receb)
Call Selecionar(txt_data_receb)

End Sub

Private Sub txt_data_receb_LostFocus()
txt_data_receb.Mask = ""
End Sub

Private Sub txt_debito_Change()
Call Calcula_Total_Pago
End Sub

Private Sub txt_debito_GotFocus()
Call Selecionar(txt_debito)
End Sub

Private Sub txt_din_Change()
Call Calcula_Total_Pago
End Sub

Private Sub txt_din_GotFocus()
Call Selecionar(txt_din)
End Sub

Private Sub txt_fiado_GotFocus()
Call bt_fiado_Click
End Sub

Private Sub txt_hiper_Change()
Call Calcula_Total_Pago
End Sub

Private Sub txt_hiper_GotFocus()
Call Selecionar(txt_hiper)
End Sub

Private Sub txt_master_Change()
Call Calcula_Total_Pago
End Sub

Private Sub txt_master_GotFocus()
Call Selecionar(txt_master)
End Sub

Private Sub txt_motivo_GotFocus()
Call Selecionar(txt_motivo)
End Sub

Private Sub txt_obs_LostFocus()

With Tab1
    .Edit
    !Observações = txt_obs
    .Update
End With

End Sub

Private Sub txt_opcional_Change()

If Not IsNumeric(txt_opcional.Text) Then txt_opcional.Text = 0
If Not IsNumeric(txt_saldo) Then txt_saldo = 0
If Not IsNumeric(txt_opcional) Then txt_opcional = 0

txt_saldo_taxa = Format(CCur(txt_saldo) + CCur(txt_opcional), "fixed")
txt_saldo_taxa_resumo = txt_saldo_taxa

End Sub

Private Sub txt_opcional_GotFocus()
Call Selecionar(txt_opcional)
End Sub

Private Sub txt_quant_Excluir_GotFocus()
Call Selecionar(txt_quant_Excluir)
End Sub

Private Sub txt_quant_GotFocus()
Call Selecionar(txt_quant)
End Sub

Private Sub txt_quant_LostFocus()

If Not IsNumeric(txt_preço) Then Exit Sub
If txt_quant = "" Then Exit Sub

If Not IsNumeric(txt_quant) Then MsgBox "Quantidade Incorreta", vbExclamation, "Atenção": txt_quant.SetFocus: Exit Sub
txt_total = Format(CCur(txt_quant) * CCur(txt_preço), "fixed")

End Sub


Sub Calcula_Total_Pago()

If Not IsNumeric(txt_din) Then txt_din.Text = 0
If Not IsNumeric(TXT_VISA) Then TXT_VISA.Text = 0
If Not IsNumeric(txt_master) Then txt_master.Text = 0
If Not IsNumeric(txt_hiper) Then txt_hiper.Text = 0
If Not IsNumeric(txt_debito) Then txt_debito.Text = 0
If Not IsNumeric(txt_opcional) Then txt_opcional.Text = 0
If Not IsNumeric(txt_troco) Then txt_troco.Text = 0
If Not IsNumeric(txt_saldo_taxa) Then txt_saldo_taxa.Text = 0
If Not IsNumeric(txt_fiado) Then txt_fiado.Text = 0

If CCur(txt_din) = 0 Then txt_troco.Text = 0
If CCur(txt_din) > CCur(txt_saldo_taxa) Then
    txt_troco.Text = Format(CCur(txt_din) - CCur(txt_saldo_taxa), "fixed")
Else
    txt_troco.Text = Format(0, "fixed")
End If

txt_total_pago = Format(CCur(txt_din) + CCur(TXT_VISA) + CCur(txt_master) + CCur(txt_hiper) + CCur(txt_debito) + CCur(txt_fiado) - CCur(txt_troco), "fixed")

End Sub

Private Sub TXT_VISA_Change()
Call Calcula_Total_Pago
End Sub

Private Sub TXT_VISA_GotFocus()
Call Selecionar(TXT_VISA)
End Sub

Sub Lança_Pagamento(DescriPag As String, ValorPag As Currency, FormaPag As String)

DataLanc = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")

With Tab5
    .AddNew
    !Data = DataLanc
    !Mesa = Mesa
    !Cod_Operador = Cod_Operador
        
    !Descrição = "FECHAMENTO: " & DescriPag
    
    !valor = ValorPag
    !Quant = 1
    !total = ValorPag
    
    !Forma_Pagam = FormaPag
    !Tipo = "C"
    !caixa = NumCaixa
    
    !ID_garcon = txt_id_garcon
    !Garcon = cmb_garcon
    
    If DescriPag = "DINHEIRO" Then
        !Recebimento = Date
        DataReceb = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
        
    ElseIf DescriPag = "CRED.PROPRIO" Then
        !Recebimento = txt_data_receb
        DataReceb = Format(Date + 1, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
        !Responsavel = txt_fiado_cli.Text
        !Contato = txt_fiado_contato.Text
        
    ElseIf DescriPag = "CARTÃO DE DÉBITO" Then
        !Recebimento = Date + 2
        DataReceb = Format(Date + 2, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
        
    Else
        !Recebimento = Date + 30
        DataReceb = Format(Date + 30, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
        
    End If
        
    .Update
End With

'lança pagamento em tabela cloud MySQL
Dim vtotal As String
vtotal = LTrim(Str(ValorPag))
'Call ConnMySQL_InserirLançamento(DataLanc, Mesa, "0", "FECHAMENTO: " & DescriPag, vtotal, "1", vtotal, "", "C", Cod_Operador, FormaPag, NumCaixa, txt_id_garcon, cmb_garcon, DataReceb)

End Sub

Sub Carrega_Mesa()

txt_saldo = ""
txt_saldo_resumo = ""

txt_opcional.Text = ""
txt_opcional_resumo.Text = ""

txt_saldo_taxa = ""
txt_saldo_taxa_resumo = ""

Total_DB = 0
Total_PARCIAL = 0
txt_id_garcon = 0
 
txt_din.Text = 0
TXT_VISA.Text = 0
txt_master.Text = 0
txt_hiper.Text = 0
txt_debito.Text = 0
txt_saldo_taxa.Text = 0
txt_fiado.Text = 0

'dados do Mesa
Set Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] where [Numero] =" & Mesa)
If Tab1!Tipo = "AVULSA" Then bt_avulsa_del.Visible = True Else bt_avulsa_del.Visible = False

'LANÇAMENTOS da Mesa (Agrupados por quantidade)
Data1.RecordSource = "SELECT tbl_lancamentos.Descrição, Sum(tbl_lancamentos.Quant) AS SomaDeQuant, " _
    & "tbl_lancamentos.Valor, Sum(tbl_lancamentos.Total) AS SomaDeTotal From tbl_lancamentos " _
    & "Where (((tbl_lancamentos.Mesa) = " & Mesa & ") And ((tbl_lancamentos.Encerrada) = False) and (tbl_lancamentos.Tipo) <> 'C' ) " _
    & "GROUP BY tbl_lancamentos.Descrição, tbl_lancamentos.Valor;"
Data1.Refresh

'AUXILIAR lançamentos
Set Tab5 = db1.OpenRecordset("select * from [tbl_lancamentos] where [Mesa]=" & Mesa & " and [Encerrada]=false")

On Error Resume Next

'pagamentos parciais
Set Tab3 = db1.OpenRecordset("SELECT Sum(tbl_lancamentos.Total) AS PARCIAL From tbl_lancamentos " _
    & "WHERE (((tbl_lancamentos.Tipo)='C') AND ((tbl_lancamentos.Mesa)=" & Mesa & " and [encerrada]=false ));")
If Not Tab3.EOF Then Total_PARCIAL = Format(Tab3!PARCIAL, "fixed")
txt_pagam_parcial = Format(Total_PARCIAL, "fixed")


'totais de débito
Set Tab3 = db1.OpenRecordset("SELECT Sum(tbl_lancamentos.Total) AS DB From tbl_lancamentos " _
    & "WHERE (((tbl_lancamentos.Tipo)='D') AND ((tbl_lancamentos.Mesa)=" & Mesa & " and [encerrada]=false ));")
If Not Tab3.EOF Then Total_DB = Format(Tab3!DB, "fixed")

txt_saldo = Format(Total_DB, "fixed")
txt_saldo_resumo = txt_saldo

If op_op.Value = 1 Then txt_opcional.Text = Format(CCur(txt_saldo) * 0.1, "fixed") Else txt_opcional.Text = 0
txt_opcional_resumo = txt_opcional

txt_saldo_taxa = Format(CCur(txt_saldo) + CCur(txt_opcional) - CCur(txt_pagam_parcial), "fixed")
txt_saldo_taxa_resumo = txt_saldo_taxa

txt_mesa.Text = Mesa
txt_mesa_tab2.Text = Mesa

txt_obs = Tab1!Observações
txt_obs_tab2.Text = Tab1!Observações

txt_id_garcon = 0 & Tab1!ID_garcon
cmb_garcon.Text = Tab1!Garcon
txt_garcon_tab2.Text = Tab1!Garcon

End Sub

Sub Imprimir_Extrato_DARUMA()

pessoas = InputBox("Quantidade de Pagantes", "Divisão de Conta", 1)
If pessoas = "" Then Exit Sub

'=================== CABEÇALHo

iRetorno = iImprimirTexto_DUAL_DarumaFramework("<e><ce><b>" + Empresa_Nome + "</b></ce></e>", 0)    'expandido,centralizado,negrito
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<ce><b>" + Empresa_End + "</b></ce>", 0)            'centraliado, negrito
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<l></l>", 0)                                        'salta 1 linha

Texto = "Extrato Mesa : " & txt_mesa
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<b>" + Texto + "</b>", 0)

Texto = "Data / Hora  : " + "<dt></dt><sp>4</sp><hr></hr>"
iRetorno = iImprimirTexto_DUAL_DarumaFramework(Texto, 0)

iRetorno = iImprimirTexto_DUAL_DarumaFramework("<tc>=</tc>", 0)                                     'linha tracejada
Texto = "Descricao<tb></tb><tb></tb><sp>4</sp>Quant<tb></tb><sp>2</sp>Valor<tb></tb>Total"
iRetorno = iImprimirTexto_DUAL_DarumaFramework(Texto, 0)
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<tc>=</tc>", 0)                                     'linha tracejada


'====================  CORPO

Set Tab11 = db1.OpenRecordset("SELECT tbl_lancamentos.Descrição, tbl_lancamentos.Valor, tbl_lancamentos.Tipo, Sum(tbl_lancamentos.Quant) AS SomaDeQuant, " _
    & "Sum(tbl_lancamentos.Total) AS Valor_Total From tbl_lancamentos Where (((tbl_lancamentos.Mesa) = " & txt_mesa & ") And " _
    & "((tbl_lancamentos.Encerrada) = False)) GROUP BY tbl_lancamentos.Descrição, tbl_lancamentos.Valor, tbl_lancamentos.Tipo ;")

extrato_subtotal = 0
extrato_opcional = 0
extrato_total = 0

Do While Not Tab11.EOF
    If Len(Tab11!Descrição) > 22 Then
        Texto = Left(Tab11!Descrição, 22)
    Else
        Texto = Tab11!Descrição & String(22 - Len(Tab11!Descrição), " ")
    End If

    Texto = Texto & "    " & Tab11!SomaDeQuant & "    " & Format(Tab11!valor, "Fixed") & "    " & Format(Tab11!Valor_Total, "Fixed")
    iRetorno = iImprimirTexto_DUAL_DarumaFramework(Texto, 0)
    
    If Tab11!Tipo = "D" Then
        extrato_subtotal = extrato_subtotal + Tab11!Valor_Total
    Else
        extrato_subtotal = extrato_subtotal - Tab11!Valor_Total
    End If
    
    Tab11.MoveNext
Loop

'====================  TOTALIZADORES
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<tc>-</tc>", 0)                                     'linha tracejada

Texto = "Sub-Total                           => " + Format(extrato_subtotal, "fixed")
iRetorno = iImprimirTexto_DUAL_DarumaFramework(Texto, 0)

If txt_opcional <> 0 Then extrato_opcional = txt_opcional
Texto = "Taxa de Servico (Opcional)      => " + Format(extrato_opcional, "fixed")
iRetorno = iImprimirTexto_DUAL_DarumaFramework(Texto, 0)

extrato_total = extrato_subtotal + extrato_opcional
Texto = "Total                               => " + Format(extrato_total, "fixed")
iRetorno = iImprimirTexto_DUAL_DarumaFramework(Texto, 0)

iRetorno = iImprimirTexto_DUAL_DarumaFramework("<l></l>", 0)                                        'salta 1 linha

Texto = "Quantidade de Pagantes              => " + pessoas
iRetorno = iImprimirTexto_DUAL_DarumaFramework(Texto, 0)


extrato_valorindiv = extrato_total / Val(pessoas)
Texto = "Valor Individual                    => " + Format(extrato_valorindiv, "fixed")
iRetorno = iImprimirTexto_DUAL_DarumaFramework(Texto, 0)

'====================  RODAPÉ
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<l></l>", 0)                                        'salta 1 linha
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<tc>=</tc>", 0)                                     'linha tracejada
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<ce><c>Desenvolvido por : www.naturaltecnologia.com</c></ce>", 0)                                     'linha tracejada

iRetorno = iImprimirTexto_DUAL_DarumaFramework("<sl>2</sl>", 0)                                     'salta 2 linhas
iRetorno = iImprimirTexto_DUAL_DarumaFramework("<gui></gui>", 0)                                    'aciona guilhotina

End Sub

Sub Imprimir_Extrato_impressora_windows()

pessoas = InputBox("Quantidade de Pagantes", "Divisão de Conta", 1)
If pessoas = "" Then Exit Sub

If txt_opcional <> 0 Then taxaOp = 1 Else taxaOp = 0
With CrystalReport1
    .ReportFileName = Caminho_Rede & "\EXTRATO.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .Formulas(2) = "pagantes = " & pessoas
    .Formulas(3) = "taxa = " & taxaOp
    .SelectionFormula = "not {tbl_lancamentos.Encerrada} and {tbl_lancamentos.Mesa} = " & Mesa
    .Action = 1
End With

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
rs!valor = lancValor
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

Sub Controles_Desabilita()

'desabilita controles
cmb_prod.Enabled = False
bt_lançar.Enabled = False
bt_imprimir.Enabled = False
bt_encerram_ok.Enabled = False
Bt_Sair.Enabled = False

End Sub
Sub Controles_Habilita()

'habilita controles
cmb_prod.Enabled = True
bt_lançar.Enabled = True
bt_imprimir.Enabled = True
bt_encerram_ok.Enabled = True
Bt_Sair.Enabled = True

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


Private Sub Apt_Click(Index As Integer)

If Apt(Index).Tag = "" Then Exit Sub
Mesa = Apt(Index).Tag
Call Carrega_Mesa
SSTab1.Tab = 1

End Sub

Private Sub Image1_Click(Index As Integer)
Call Apt_Click(Index)
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
Call Carrega_Mesa
SSTab1.Tab = 1

End Sub

Sub Limpar_Mapa()

'limpar tela
Apt_index = 0
Do While Apt_index <= 74
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
    
    If Apt_index > 74 Then Exit Do

Loop

End Sub

Sub Atualiza_Mapa()

'TODAS os aptos
Call Limpar_Mapa
Set Mapa_Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] order by [Numero]")
Call Montar_Mapa

End Sub


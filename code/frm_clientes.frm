VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frm_clientes 
   BackColor       =   &H80000018&
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   7425
   ClientLeft      =   1110
   ClientTop       =   585
   ClientWidth     =   8235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   19200
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Bt_Alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   7200
      Picture         =   "frm_clientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Alterar"
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Bt_Excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   8280
      Picture         =   "frm_clientes.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Excluir"
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Bt_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   6120
      Picture         =   "frm_clientes.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Novo"
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton bt_confirmar 
      Caption         =   "Confirmar"
      Height          =   855
      Left            =   9720
      Picture         =   "frm_clientes.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Confirmar Inclusões/Alterações"
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton bt_desistir 
      Caption         =   "Desistir"
      Height          =   855
      Left            =   10800
      Picture         =   "frm_clientes.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Desistir da última Inclusão/Alteração"
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Bt_Sair 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   855
      Left            =   12360
      Picture         =   "frm_clientes.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Fechar esta Janela"
      Top             =   6480
      Width           =   855
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   350
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Width           =   1250
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   10610
      _Version        =   327680
      Style           =   1
      TabHeight       =   520
      BackColor       =   -2147483624
      TabCaption(0)   =   "Todos"
      TabPicture(0)   =   "frm_clientes.frx":198C
      Tab(0).ControlCount=   4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "bt_loc_nome1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "bt_loc_cod1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "bt_loc_doc"
      Tab(0).Control(3).Enabled=   0   'False
      TabCaption(1)   =   "Ficha Individual"
      TabPicture(1)   =   "frm_clientes.frx":19A8
      Tab(1).ControlCount=   4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "bt_voltar"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "bt_loc_nome"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "bt_loc_cod"
      Tab(1).Control(3).Enabled=   -1  'True
      TabCaption(2)   =   "Log"
      TabPicture(2)   =   "frm_clientes.frx":19C4
      Tab(2).ControlCount=   3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DBGrid2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Data2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Text3"
      Tab(2).Control(2).Enabled=   0   'False
      Begin VB.TextBox Text3 
         DataField       =   "Nome"
         DataSource      =   "Data1"
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
         Left            =   -74280
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   30
         Top             =   600
         Width           =   11895
      End
      Begin VB.Data Data2 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   -74760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5160
         Visible         =   0   'False
         Width           =   1250
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frm_clientes.frx":19E0
         Height          =   4215
         Left            =   -74280
         OleObjectBlob   =   "frm_clientes.frx":19F0
         TabIndex        =   29
         Top             =   1200
         Width           =   11895
      End
      Begin VB.CommandButton bt_loc_doc 
         Caption         =   "D"
         Height          =   315
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Localizar por CNPJ/CPF"
         Top             =   1560
         Width           =   315
      End
      Begin VB.CommandButton bt_voltar 
         Caption         =   "<"
         Height          =   315
         Left            =   -74760
         TabIndex        =   27
         ToolTipText     =   "Voltar"
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton bt_loc_cod1 
         Caption         =   "C"
         Height          =   315
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Localizar por Código"
         Top             =   1080
         Width           =   315
      End
      Begin VB.Frame frame1 
         Enabled         =   0   'False
         Height          =   4935
         Left            =   -74280
         TabIndex        =   20
         Top             =   480
         Width           =   12015
         Begin VB.PictureBox Picture1 
            Height          =   2295
            Left            =   9360
            ScaleHeight     =   2235
            ScaleWidth      =   2235
            TabIndex        =   38
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox Text11 
            DataField       =   "Telefone"
            DataSource      =   "Data1"
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
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   10
            Top             =   2640
            Width           =   7935
         End
         Begin VB.TextBox Text10 
            DataField       =   "UF"
            DataSource      =   "Data1"
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
            Left            =   7800
            MaxLength       =   2
            TabIndex        =   9
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox Text9 
            DataField       =   "Cidade"
            DataSource      =   "Data1"
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
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   8
            Top             =   2160
            Width           =   5775
         End
         Begin VB.TextBox Text8 
            DataField       =   "CEP"
            DataSource      =   "Data1"
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
            Left            =   7800
            MaxLength       =   10
            TabIndex        =   7
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Text6 
            DataField       =   "Número"
            DataSource      =   "Data1"
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
            Left            =   8160
            MaxLength       =   5
            TabIndex        =   5
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox Text4 
            DataField       =   "Endereço"
            DataSource      =   "Data1"
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
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   4
            Top             =   1200
            Width           =   5775
         End
         Begin VB.TextBox Text2 
            DataField       =   "RG_DOC"
            DataSource      =   "Data1"
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
            Left            =   5640
            MaxLength       =   20
            TabIndex        =   3
            Top             =   720
            Width           =   3495
         End
         Begin VB.TextBox TXT_CPF 
            DataField       =   "CPF"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txt_nome 
            DataField       =   "Nome"
            DataSource      =   "Data1"
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
            Left            =   3120
            MaxLength       =   100
            TabIndex        =   1
            Top             =   240
            Width           =   6015
         End
         Begin VB.TextBox txt_cod 
            DataField       =   "Código"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text5 
            DataField       =   "email"
            DataSource      =   "Data1"
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
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   11
            Top             =   3120
            Width           =   7935
         End
         Begin VB.TextBox Text7 
            DataField       =   "Bairro"
            DataSource      =   "Data1"
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
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1680
            Width           =   5775
         End
         Begin VB.TextBox Text1 
            DataField       =   "obs"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1155
            Left            =   1200
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   3600
            Width           =   10455
         End
         Begin MSMask.MaskEdBox CGC_CPF 
            Height          =   375
            Left            =   1200
            TabIndex        =   2
            Top             =   720
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   327680
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Telefone:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   37
            Top             =   2640
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   10
            Left            =   7320
            TabIndex        =   36
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   35
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   7
            Left            =   7320
            TabIndex        =   34
            Top             =   1680
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   4
            Left            =   7320
            TabIndex        =   33
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   32
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "RG / DOC :"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   1
            Left            =   4680
            TabIndex        =   31
            Top             =   720
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Obs.:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   26
            Top             =   3600
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nome :"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   3
            Left            =   2400
            TabIndex        =   25
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "e-mail:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   23
            Top             =   3120
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   22
            Top             =   1680
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CPF:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   21
            Top             =   720
            Width           =   345
         End
      End
      Begin VB.CommandButton bt_loc_nome 
         Caption         =   "?"
         Height          =   315
         Left            =   -68100
         TabIndex        =   19
         Top             =   1080
         Width           =   315
      End
      Begin VB.CommandButton bt_loc_cod 
         Caption         =   "?"
         Height          =   315
         Left            =   -72480
         TabIndex        =   18
         Top             =   720
         Width           =   315
      End
      Begin VB.CommandButton bt_loc_nome1 
         Caption         =   "N"
         Height          =   315
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "Localizar por Nome"
         Top             =   600
         Width           =   315
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_clientes.frx":2566
         Height          =   4815
         Left            =   720
         OleObjectBlob   =   "frm_clientes.frx":2596
         TabIndex        =   17
         Top             =   600
         Width           =   12015
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   7
      Left            =   10920
      Top             =   6600
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   6
      Left            =   12480
      Top             =   6600
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   5
      Left            =   9840
      Top             =   6600
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   4
      Left            =   8400
      Top             =   6600
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   3
      Left            =   7320
      Top             =   6600
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   2
      Left            =   6240
      Top             =   6600
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5775
      Index           =   0
      Left            =   360
      Top             =   480
      Width           =   12975
   End
End
Attribute VB_Name = "frm_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Confirmando As Boolean
Dim Adicionando As Boolean
Dim Editar As Boolean

Dim Desistir As Boolean
Dim Carregando As Boolean

Dim LOCALIZADO As Boolean

Dim db1 As Database
Dim Tab1 As Recordset   'funcionarios
Dim Tab2 As Recordset   'log

Private Sub Bt_Alterar_Click()

If Data1.Recordset.EOF Then Exit Sub

Bt_novo.Enabled = False
Bt_Excluir.Enabled = False
Bt_Alterar.Enabled = False
Bt_Sair.Enabled = False
   
Editar = True

frame1.Enabled = True

SSTab1.Tab = 1
txt_nome.SetFocus

Data1.Recordset.Edit

End Sub

Private Sub bt_confirmar_Click()
If Editar = False Then Exit Sub Else Editar = False
If Adicionando = True Then Confirmando = True
Call Data1_Validate(-1, -1)
End Sub

Private Sub bt_desistir_Click()
If Editar = False Then Exit Sub Else Editar = False
Desistir = True
Call Data1_Validate(-1, -1)
Desistir = False
End Sub

Private Sub bt_loc_cod_Click()
Call Loc_number(Data1, "Informe o Código do Cliente", "Código")
End Sub

Private Sub bt_loc_cod1_Click()

Call bt_loc_cod_Click

If LOCALIZADO Then
    SSTab1.Tab = 1
    bt_voltar.SetFocus
End If

End Sub

Private Sub bt_loc_doc_Click()

Call Loc_like(Data1, "Informe o número do CPF do Cliente (Completo ou Parcial)", "CPF")

If LOCALIZADO Then
    SSTab1.Tab = 1
    bt_voltar.SetFocus
End If

End Sub

Private Sub bt_loc_nome_Click()
Call Loc_like(Data1, "Informe o Nome (Completo ou Parcial) do Cliente", "nome")
End Sub

Private Sub bt_loc_nome1_Click()

Call bt_loc_nome_Click

If LOCALIZADO Then
    SSTab1.Tab = 1
    bt_voltar.SetFocus
End If

End Sub

Private Sub Bt_Voltar_Click()
SSTab1.Tab = 0
bt_loc_nome1.SetFocus
End Sub

Private Sub Data1_Reposition()

If Carregando Then Data1.Recordset.LockEdits = False: Carregando = False
If Data1.Recordset.EOF Then Exit Sub

If Len(Data1.Recordset!CPF) = 11 Then
    CGC_CPF.Mask = "###.###.###-##"
    CGC_CPF = "" & Data1.Recordset!CPF
Else
    CGC_CPF.Mask = "##.###.###/####-##"
    CGC_CPF = "" & Data1.Recordset!CPF
End If

'log operações
Data2.RecordSource = "select * from [tbl_cad_log] where [Cod_Hóspede]=" & Data1.Recordset!Código _
    & " ORDER BY [data] desc"
Data2.Refresh

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

On Error GoTo Trata_erro

If Save And Desistir = False Then
            
    If txt_nome.Text = "" Then
        MsgBox "Informe o Nome do Cliente!"
        txt_nome.SetFocus
        Editar = True
        Save = False
        Action = vbDataActionCancel
        Exit Sub
    End If
        
    If CGC_CPF.Text = "" Then
        CGC_CPF.Text = "11111111111"
        'MsgBox "Informe o CPF/CNPJ!"
        'CGC_CPF.SetFocus
        'Editar = True
        'Save = False
        'Action = vbDataActionCancel
        'Exit Sub
    End If
        
    TXT_CPF = "" & CGC_CPF.Text
        
    'LOG DE OPERAÇÕES
    With Tab2
        .AddNew
        !Data = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
        !Cod_Operador = Cod_Operador
        If IsNumeric(txt_cod.Text) Then !Cod_Hóspede = txt_cod.Text
        !Operador = Usuário
        If Adicionando Then !operação = "CADASTRO" Else !operação = "ALTERAÇÃO"
        .Update
    End With
   
ElseIf Save And Desistir = True Then

    Data1.Recordset.CancelUpdate
            
End If

'desabilita campos para edição e habilita botões
   
Bt_novo.Enabled = True
Bt_Excluir.Enabled = True
Bt_Alterar.Enabled = True
Bt_Sair.Enabled = True
    
frame1.Enabled = False

If Confirmando = True Then
    Confirmando = False
    Adicionando = False
    Data1.Recordset.MoveLast
End If

Exit Sub
Trata_erro:
MsgBox "Bloqueio de Segurança ! Tente Novamente. Cod.: " & Str$(Err.Number) & "  /  Descrição : " & Err.Description
Exit Sub

End Sub

Private Sub Bt_Excluir_Click()
Call Excluir(Data1)
End Sub

Private Sub Bt_Novo_Click()

On Error GoTo Trata_erro

'auxiliar funcionarios
Set Tab1 = db1.OpenRecordset("select [Código],[CPF] from [Tbl_Clientes] order by [Código] desc")

Data1.Recordset.AddNew

'Abilita campos para edição e desabilita botões
frame1.Enabled = True

Bt_novo.Enabled = False
Bt_Excluir.Enabled = False
Bt_Alterar.Enabled = False
Bt_Sair.Enabled = False

SSTab1.Tab = 1
txt_nome.SetFocus

Editar = True
Adicionando = True

Exit Sub
Trata_erro:
MsgBox "Bloqueio de Segurança ! Tente Novamente. Cod.: " & Str$(Err.Number) & "  /  Descrição : " & Err.Description
Exit Sub

End Sub

Private Sub Bt_Sair_Click()
Unload Me
frm_mnu.barramenu.Visible = True
End Sub

Private Sub DBGrid1_DblClick()
SSTab1.Tab = 1
End Sub

Private Sub Form_Load()

Call Abrir_BD_Data(Data1, "Tbl_Clientes", "nome", "")
Call Abrir_BD_Data(Data2, "tbl_cad_log", "data", "cod_operador=0")

Set db1 = OpenDatabase(Caminho_Rede & "\dados.mdb")
Set Tab2 = db1.OpenRecordset("select * from [tbl_cad_log]")

End Sub

Private Sub cgc_CPF_GotFocus()
CGC_CPF.Mask = "##############"
End Sub

Private Sub CGC_CPF_LostFocus()

If Len(CGC_CPF.Text) > 0 Then

    Select Case Len(CGC_CPF.Text)
       Case Is = 11
         CGC_CPF.Mask = "###.###.###-##"
         If Not calculaCPF(CGC_CPF.Text) Then
            MsgBox "CPF Incorreto !!!", vbExclamation, "Cadastro"
            CGC_CPF = ""
            CGC_CPF.Mask = "###########"
            CGC_CPF.SetFocus
         End If
         
    Case Is = 14
         CGC_CPF.Mask = "##.###.###/####-##"
         If Not ValidaCGC(CGC_CPF.Text) Then
            MsgBox "CGC Incorreto !!! ", vbExclamation, "Cadastro"
            CGC_CPF = ""
            CGC_CPF.Mask = "##############"
            CGC_CPF.SetFocus
         End If
    
    Case Else
            MsgBox "CGC Incorreto !!! ", vbExclamation, "Cadastro"
            CGC_CPF = ""
            CGC_CPF.Mask = "###########"
            CGC_CPF.SetFocus
       
    End Select
End If

'verificar se cliente já existe
If Adicionando Then
    If CGC_CPF.Text = "" Then Exit Sub
    Tab1.FindFirst ("CPF = '" & CGC_CPF & "'")
    If Tab1.NoMatch = False Then
        MsgBox "CPF Já Cadastrado", vbExclamation, "Atenção"
        
        Desistir = True
        Call Data1_Validate(-1, -1)
        Desistir = False
        
        Data1.Recordset.FindFirst ("Código = " & Tab1!Código)
        
        Exit Sub
    End If
End If

End Sub

Public Function CalculaCGC(Numero As String) As String

Dim I As Integer
Dim prod As Integer
Dim mult As Integer
Dim digito As Integer

If Not IsNumeric(Numero) Then
   CalculaCGC = ""
   Exit Function
End If

mult = 2
For I = Len(Numero) To 1 Step -1
  prod = prod + Val(Mid(Numero, I, 1)) * mult
  mult = IIf(mult = 9, 2, mult + 1)
Next

digito = 11 - Int(prod Mod 11)
digito = IIf(digito = 10 Or digito = 11, 0, digito)

CalculaCGC = Trim(Str(digito))

End Function

Function calculaCPF(CPF As String) As Boolean

On Error GoTo Err_CPF

Dim I As Integer        'utilizada nos FOR... NEXT
Dim strcampo As String  'armazena do CPF que será utilizada para o cálculo
Dim strCaracter As String   'armazena os dígitos do CPF da direita para a esquerda
Dim intNumero As Integer    'armazena o digito separado para cálculo (uma a um)
Dim intMais As Integer  'armazena o digito específico multiplicado pela sua base
Dim lngSoma As Long     'armazena a soma dos dígitos multiplicados pela sua base(intmais)
Dim dblDivisao As Double    'armazena a divisão dos dígitos * base por 11
Dim lngInteiro As Long  'armazena inteiro da divisão
Dim intResto As Integer     'armazena o resto
Dim intDig1 As Integer  'armazena o 1º digito verificador
Dim intDig2 As Integer  'armazena o 2º digito verificador
Dim strConf As String   'armazena o digito verificador

lngSoma = 0
intNumero = 0
intMais = 0
strcampo = Left(CPF, 9)

'Inicia cálculos do 1º dígito
For I = 2 To 10
    strCaracter = Right(strcampo, I - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * I
    lngSoma = lngSoma + intMais
Next I
dblDivisao = lngSoma / 11

lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
Else
    intDig1 = 11 - intResto
End If

strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
lngSoma = 0
intNumero = 0
intMais = 0
'Inicia cálculos do 2º dígito
For I = 2 To 11
    strCaracter = Right(strcampo, I - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * I
    lngSoma = lngSoma + intMais
Next I
dblDivisao = lngSoma / 11
lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig2 = 0
Else
    intDig2 = 11 - intResto
End If
strConf = intDig1 & intDig2
'Caso o CPF esteja errado dispara a mensagem
If strConf <> Right(CPF, 2) Then
    calculaCPF = False
Else
    calculaCPF = True
End If
Exit Function

Exit_CPF:
    Exit Function
Err_CPF:
    'MsgBox Error$
    Resume Exit_CPF
End Function

Public Function ValidaCGC(CGC As String) As Boolean

If CalculaCGC(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
   ValidaCGC = False
   Exit Function
End If

If CalculaCGC(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
   ValidaCGC = False
   Exit Function
End If

ValidaCGC = True

End Function

Private Sub Form_Unload(Cancel As Integer)
frm_mnu.barramenu.Visible = True
End Sub

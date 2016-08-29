VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_produtos 
   BackColor       =   &H80000018&
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   8595
   ClientLeft      =   180
   ClientTop       =   510
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton bt_import 
      Caption         =   "&Importar"
      Height          =   855
      Left            =   4200
      Picture         =   "frm_produtos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Importar dados da Internet"
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton bt_repor 
      Caption         =   "&Reposição"
      Height          =   855
      Left            =   1920
      Picture         =   "frm_produtos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Reposição de Estoque"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton bt_ajuste 
      Caption         =   "Ajuste"
      Height          =   855
      Left            =   3000
      Picture         =   "frm_produtos.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Ajuste de Estoque"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Bt_Sair 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   855
      Left            =   13440
      Picture         =   "frm_produtos.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Fechar esta Janela"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton bt_desistir 
      Caption         =   "Desistir"
      Height          =   855
      Left            =   12360
      Picture         =   "frm_produtos.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Desistir da última Inclusão/Alteração"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton bt_confirmar 
      Caption         =   "Confirmar"
      Height          =   855
      Left            =   11400
      Picture         =   "frm_produtos.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Confirmar Inclusões/Alterações"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Bt_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   8400
      Picture         =   "frm_produtos.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Novo"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Bt_Excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   10320
      Picture         =   "frm_produtos.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Excluir"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Bt_Alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   9360
      Picture         =   "frm_produtos.frx":2210
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Alterar"
      Top             =   6360
      Width           =   855
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Width           =   1250
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   240
      TabIndex        =   18
      Top             =   240
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   12726
      _Version        =   327680
      TabHeight       =   520
      BackColor       =   -2147483624
      TabCaption(0)   =   "Listagem"
      TabPicture(0)   =   "frm_produtos.frx":2652
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Detalhes"
      TabPicture(1)   =   "frm_produtos.frx":266E
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "Ficha Técnica"
      TabPicture(2)   =   "frm_produtos.frx":268A
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Begin VB.Frame Frame2 
         Height          =   5295
         Left            =   -74760
         TabIndex        =   27
         Top             =   480
         Width           =   13815
         Begin VB.Data Data2 
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3120
            Visible         =   0   'False
            Width           =   1250
         End
         Begin VB.CommandButton bt_insumo_ad 
            Caption         =   "Adicionar"
            Height          =   420
            Left            =   9600
            TabIndex        =   31
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txt_insumo_quant 
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
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   29
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txt_insumo_und 
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
            Left            =   8400
            MaxLength       =   5
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1200
            Width           =   975
         End
         Begin VB.ComboBox cmb_insumo 
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
            TabIndex        =   28
            Top             =   1200
            Width           =   6255
         End
         Begin MSDBGrid.DBGrid DBGrid2 
            Bindings        =   "frm_produtos.frx":26A6
            Height          =   2895
            Left            =   240
            OleObjectBlob   =   "frm_produtos.frx":26B6
            TabIndex        =   32
            Top             =   1800
            Width           =   13215
         End
         Begin VB.Label lbl_produto 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Descrição"
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
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   13215
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   9
            Left            =   6720
            TabIndex        =   35
            Top             =   960
            Width           =   870
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Unidade:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   8
            Left            =   8400
            TabIndex        =   33
            Top             =   960
            Width           =   645
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Insumo:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   34
            Top             =   960
            Width           =   555
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_produtos.frx":3214
         Height          =   5295
         Left            =   240
         OleObjectBlob   =   "frm_produtos.frx":3224
         TabIndex        =   17
         Top             =   600
         Width           =   13815
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   5415
         Left            =   -74760
         TabIndex        =   19
         Top             =   480
         Width           =   13815
         Begin VB.CheckBox op_ficha 
            Caption         =   "Ficha Técnica"
            DataField       =   "Ficha_Tec"
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
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   4080
            Width           =   2535
         End
         Begin VB.TextBox txt_insumo 
            DataField       =   "Insumo"
            DataSource      =   "Data1"
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   600
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txt_Estoque_min 
            Alignment       =   1  'Right Justify
            DataField       =   "Estoque_min"
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
            Left            =   240
            MaxLength       =   5
            TabIndex        =   5
            Top             =   3480
            Width           =   1455
         End
         Begin VB.TextBox txt_estoque 
            Alignment       =   1  'Right Justify
            DataField       =   "Estoque"
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
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   3480
            Width           =   1455
         End
         Begin VB.TextBox txt_cod 
            DataField       =   "Código"
            DataSource      =   "Data1"
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
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   0
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cmb_grupo 
            DataField       =   "Grupo"
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
            Left            =   7080
            TabIndex        =   2
            Top             =   1560
            Width           =   6495
         End
         Begin MSMask.MaskEdBox txt_preço 
            DataField       =   "Preço"
            DataSource      =   "Data1"
            Height          =   420
            Left            =   2040
            TabIndex        =   4
            Top             =   2520
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   741
            _Version        =   327680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txt_und 
            DataField       =   "Unidade"
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
            Left            =   240
            MaxLength       =   3
            ScrollBars      =   1  'Horizontal
            TabIndex        =   3
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txt_produto 
            DataField       =   "Descrição"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   240
            MaxLength       =   100
            TabIndex        =   1
            Top             =   1560
            Width           =   6495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Estoque Mínimo:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   26
            Top             =   3240
            Width           =   1200
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Estoque atual :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   5
            Left            =   2040
            TabIndex        =   25
            Top             =   3240
            Width           =   1065
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Preço de Venda:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   2040
            TabIndex        =   23
            Top             =   2280
            Width           =   1200
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Unidade:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Grupo:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   7080
            TabIndex        =   21
            Top             =   1320
            Width           =   480
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   20
            Top             =   1320
            Width           =   600
         End
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   7215
      Index           =   1
      Left            =   360
      Top             =   360
      Width           =   14295
   End
End
Attribute VB_Name = "frm_produtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Confirmando As Boolean
Dim Adicionando As Boolean
Dim Desistir As Boolean
Dim Carregando As Boolean
Dim Editar As Boolean

Dim db1 As Database
Dim Tab1 As Recordset   'auxiliar grupos
Dim Tab2 As Recordset   'auxiliar reposição de estoque
Dim Tab3 As Recordset   'insumos

Dim CodInsumo As Long

Private Sub bt_ajuste_Click()

If Data1.Recordset.EOF Then Exit Sub

quant_Ajuste = InputBox("Quantidade a Ajustar", "Estoque Atual =" & Data1.Recordset!estoque)
If Not IsNumeric(quant_Ajuste) Then MsgBox "Quantidade Inválida", vbExclamation, "Atenção": Exit Sub

If Conf("Estoque com Ajuste = " & Data1.Recordset!estoque + quant_Ajuste, "Confirma Ajuste ?") = 7 Then Exit Sub

With Data1.Recordset
    .Edit
    !estoque = !estoque + quant_Ajuste
    .Update
End With

'lança ajuste de estoque na tabela de histórico de reposições
Set Tab2 = db1.OpenRecordset("select * from [Tbl_Produtos_Reposições]")
With Tab2
    .AddNew
    !Cod_Produto = Data1.Recordset!código
    !Data = Date
    !FORNECEDOR = "Ajuste de Estoque"
    !Quant = quant_Ajuste
    !Operador = Usuário
    .Update
End With
Tab2.Close

End Sub

Private Sub Bt_Alterar_Click()

If Data1.Recordset.EOF Then Exit Sub

Bt_novo.Enabled = False
Bt_Excluir.Enabled = False
Bt_Alterar.Enabled = False
Bt_Sair.Enabled = False
bt_repor.Enabled = False
bt_ajuste.Enabled = False
    
Editar = True

Frame1.Enabled = True

SSTab1.Tab = 1
txt_produto.SetFocus

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


Private Sub Bt_Excluir_Click()

Call Excluir(Data1)

End Sub

Private Sub bt_import_Click()

Set Tab1 = db1.OpenRecordset("select [Encerrada] from [tbl_lancamentos] where [encerrada]=false")
If Not Tab1.EOF Then MsgBox "Existem lançamentos em aberto. Necessário ENCERRAR contas primeiro", vbExclamation, "Atenção": Exit Sub

If Conf("O cadastro atual será APAGADO! Importar dados lançados pela Internet?", "Atenção!") = 7 Then Exit Sub

Exit Sub
'desativado até definição de prioridade na definição do ID
Call Atualiza_Cad_Produtos

End Sub

Private Sub bt_insumo_ad_Click()

'validações
If Data1.Recordset.EOF Then Exit Sub
If CodInsumo = 0 Then MsgBox "Selecione um Insumo da Lista", vbExclamation, "Atenção": cmb_insumo.SetFocus: Exit Sub
If Not IsNumeric(txt_insumo_quant) Then MsgBox "Quantidade Inválida", vbExclamation, "Atenção": txt_insumo_quant.SetFocus: Exit Sub

'adiciona ficha técnica
With Data2.Recordset
    .AddNew
    !ID_Produto = Data1.Recordset!código
    !id_insumo = CodInsumo
    !Insumo = Tab3!Descrição
    !Und = Tab3!unidade
    !Quant = txt_insumo_quant
    .Update
End With

cmb_insumo.Text = ""
txt_insumo_quant.Text = ""
txt_insumo_und.Text = ""
cmb_insumo.SetFocus

End Sub

Private Sub Bt_Novo_Click()

On Error GoTo Trata_erro

'atualiza combo grupo de produtos
'Set Tab1 = db1.OpenRecordset("SELECT Tbl_Produtos.Grupo From Tbl_Produtos " _
'    & "where [Insumo]='" & Rotina & "' GROUP BY Tbl_Produtos.Grupo;")
'cmb_grupo.Clear
'Do While Not Tab1.EOF
'    cmb_grupo.AddItem ("" & Tab1!grupo)
'    Tab1.MoveNext
'Loop

Data1.Recordset.AddNew

'Abilita campos para edição e desabilita botões
Frame1.Enabled = True

Bt_novo.Enabled = False
Bt_Excluir.Enabled = False
Bt_Alterar.Enabled = False
Bt_Sair.Enabled = False
bt_repor.Enabled = False
bt_ajuste.Enabled = False

SSTab1.Tab = 1
txt_produto.SetFocus

Editar = True
Adicionando = True

Exit Sub
Trata_erro:
MsgBox "Bloqueio de Segurança ! Tente Novamente. Cod.: " & Str$(Err.Number) & "  /  Descrição : " & Err.Description
Exit Sub

End Sub

Private Sub bt_repor_Click()

If Data1.Recordset.EOF Then Exit Sub

Cod_Produto_Repor = Data1.Recordset!código
frm_produtos_Repor.Show

End Sub

Private Sub Bt_Sair_Click()
Unload Me
frm_mnu.barramenu.Visible = True
End Sub

Private Sub cmb_insumo_LostFocus()

txt_insumo_und.Text = ""
CodInsumo = 0

If cmb_insumo.Text = "" Then Exit Sub
Tab3.FindFirst ("Descrição = '" & cmb_insumo.Text & "'")
If Not Tab3.NoMatch Then
    txt_insumo_und.Text = Tab3!unidade
    CodInsumo = Tab3!código
End If

End Sub

Private Sub Data1_Reposition()

If Carregando Then Data1.Recordset.LockEdits = False: Carregando = False

If Data1.Recordset.EOF Then Exit Sub

'ficha técnica
Data2.RecordSource = "select * from [tbl_Produtos_FichaTec] where [id_produto]= " & Data1.Recordset!código
Data2.Refresh

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

On Error GoTo Trata_erro

If Save And Desistir = False Then

    If txt_produto.Text = "" Then
        MsgBox "Informe o nome do Produto!"
        txt_produto.SetFocus
        Editar = True
        Save = False
        Action = vbDataActionCancel
        Exit Sub
    End If
    
    txt_insumo.Text = Rotina
    
ElseIf Save And Desistir = True Then

    Data1.Recordset.CancelUpdate
            
End If

'desabilita campos para edição e habilita botões
   
Bt_novo.Enabled = True
Bt_Excluir.Enabled = True
Bt_Alterar.Enabled = True
Bt_Sair.Enabled = True
bt_repor.Enabled = True
bt_ajuste.Enabled = True
    
Frame1.Enabled = False

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

Private Sub DBGrid2_DblClick()

Call Excluir(Data2)

End Sub

Private Sub Form_Load()

'produtos/insumos
Call Abrir_BD_Data(Data1, "Tbl_Produtos", "[grupo],[Descrição]", "[insumo]='" & Rotina & "'")
'ficha técnica (insumos)
Call Abrir_BD_Data(Data2, "tbl_Produtos_FichaTec", "[ID_Insumo]", "[ID_Insumo]=0")

Carregando = True

Set db1 = OpenDatabase(Caminho_Rede & "\dados.mdb")

'combo grupo de produtos
Set Tab1 = db1.OpenRecordset("SELECT Tbl_Produtos.Grupo From Tbl_Produtos " _
    & "where [Insumo]='" & Rotina & "' GROUP BY Tbl_Produtos.Grupo;")
Do While Not Tab1.EOF
    cmb_grupo.AddItem ("" & Tab1!grupo)
    Tab1.MoveNext
Loop

'combo insumos
Set Tab3 = db1.OpenRecordset("select * from [tbl_produtos] where [insumo]='S' order by [Descrição]")
Do While Not Tab3.EOF
    cmb_insumo.AddItem ("" & Tab3!Descrição)
    Tab3.MoveNext
Loop

If Rotina = "S" Then
    txt_preço.Visible = False
    lblLabels(4).Visible = False
    op_ficha.Visible = False
End If

End Sub

Private Sub DBGrid1_DblClick()
SSTab1.Tab = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_mnu.barramenu.Visible = True
End Sub

Private Sub txt_Estoque_min_GotFocus()
Call Selecionar(txt_Estoque_min)
End Sub

Private Sub txt_preço_GotFocus()
Call Selecionar(txt_preço)
End Sub

Private Sub txt_produto_GotFocus()
Call Selecionar(txt_produto)
End Sub

Sub Atualiza_Cad_Produtos()

On Error GoTo Trata_erro

Me.MousePointer = 11

'declara e inicia conexão
Set conn = New ADODB.Connection
conn.ConnectionString = StringConexao
conn.CursorLocation = adUseClient
conn.Open

Set rs = New ADODB.Recordset
rs.Open "Select * from Tbl_Produtos", conn

If rs.EOF Then
    MsgBox "Não existem dados para importar", vbExclamation, "Atenção"
    'fecha conexão
    conn.Close
    Me.MousePointer = 0
    Exit Sub
End If

'apaga registros atuais
db1.Execute "delete * from [Tbl_Produtos]"

'abre arquivo para receber novos dados
Set Tab1 = db1.OpenRecordset("select * from Tbl_Produtos")

Do While Not rs.EOF
    With Tab1
        .AddNew
        !código = rs!ID_Produto
        !grupo = rs!grupo
        !Descrição = rs!Descricao
        !unidade = rs!unidade
        !Estoque_min = rs!Estoque_min
        !Preço = rs!preco
        !Insumo = rs!Insumo
        .Update
    End With
    rs.MoveNext
Loop

'fecha conexão
conn.Close

'atualiza grid
Data1.Refresh
bt_import.Enabled = False
Me.MousePointer = 0
MsgBox "Cadastro Atualizado", vbInformation, "Ok"

Exit Sub
Trata_erro:
'-------------------------------------------------------------------------------------------------------------
'dados do erro
cloud_erro = Str$(Err.Number)
Cloud_erro_desc = Err.Description
Erro_Reenvio = True
Me.MousePointer = 0
MsgBox "Atenção : " & Cloud_erro_desc, vbInformation, "Erro: " & cloud_erro
Exit Sub

End Sub

Private Sub txt_und_GotFocus()
Call Selecionar(txt_und)
End Sub

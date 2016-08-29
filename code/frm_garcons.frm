VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_garcons 
   BackColor       =   &H80000018&
   Caption         =   "Cadastro de Garçons"
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
   Begin VB.CommandButton Bt_Sair 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   855
      Left            =   13440
      Picture         =   "frm_garcons.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Fechar esta Janela"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton bt_desistir 
      Caption         =   "Desistir"
      Height          =   855
      Left            =   12120
      Picture         =   "frm_garcons.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Desistir da última Inclusão/Alteração"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton bt_confirmar 
      Caption         =   "Confirmar"
      Height          =   855
      Left            =   11040
      Picture         =   "frm_garcons.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Confirmar Inclusões/Alterações"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Bt_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   7080
      Picture         =   "frm_garcons.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Novo"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Bt_Excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   9600
      Picture         =   "frm_garcons.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Excluir"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Bt_Alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   8280
      Picture         =   "frm_garcons.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   5
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
      TabIndex        =   10
      Top             =   240
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   12726
      _Version        =   327680
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   -2147483624
      TabCaption(0)   =   "Garçons"
      TabPicture(0)   =   "frm_garcons.frx":198C
      Tab(0).ControlCount=   3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "bt_import"
      Tab(0).Control(2).Enabled=   0   'False
      Begin VB.CommandButton bt_import 
         Caption         =   "&Importar"
         Height          =   855
         Left            =   1800
         Picture         =   "frm_garcons.frx":19A8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Importar dados da Internet"
         Top             =   6120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   2175
         Left            =   240
         TabIndex        =   11
         Top             =   3720
         Width           =   13815
         Begin VB.TextBox txt_nome 
            DataField       =   "Nome_Garcon"
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
            MaxLength       =   20
            TabIndex        =   1
            Top             =   1320
            Width           =   8655
         End
         Begin VB.TextBox txt_num 
            DataField       =   "id"
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
            Top             =   480
            Width           =   1575
         End
         Begin MSMask.MaskEdBox txt_diaria 
            DataField       =   "Diaria"
            DataSource      =   "Data1"
            Height          =   420
            Left            =   9240
            TabIndex        =   2
            Top             =   1320
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
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Valor Diária"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   9240
            TabIndex        =   15
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   465
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_garcons.frx":1DEA
         Height          =   3315
         Left            =   240
         OleObjectBlob   =   "frm_garcons.frx":1DFA
         TabIndex        =   3
         Top             =   300
         Width           =   13815
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
Attribute VB_Name = "frm_garcons"
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
Dim Tab1 As Recordset

Private Sub Bt_Alterar_Click()


If Data1.Recordset.EOF Then Exit Sub

'SSTab1.Tab = 1

Bt_novo.Enabled = False
Bt_Excluir.Enabled = False
Bt_Alterar.Enabled = False
Bt_Sair.Enabled = False
    
Editar = True

Frame1.Enabled = True

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


Private Sub Bt_Excluir_Click()

Call Excluir(Data1)

End Sub

Private Sub bt_import_Click()

Set Tab1 = db1.OpenRecordset("select [Encerrada] from [tbl_lancamentos] where [encerrada]=false")
If Not Tab1.EOF Then MsgBox "Existem lançamentos em aberto. Necessário ENCERRAR contas primeiro", vbExclamation, "Atenção": Exit Sub

If Conf("O cadastro atual será APAGADO! Importar dados lançados pela Internet?", "Atenção!") = 7 Then Exit Sub

Exit Sub
'desativado até definição de prioridade na definição do ID
Call Atualiza_Cad_Garcons

End Sub

Private Sub Bt_Novo_Click()

On Error GoTo Trata_erro

Data1.Recordset.AddNew

'Abilita campos para edição e desabilita botões
Frame1.Enabled = True

Bt_novo.Enabled = False
Bt_Excluir.Enabled = False
Bt_Alterar.Enabled = False
Bt_Sair.Enabled = False

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

Private Sub Data1_Reposition()

If Carregando Then Data1.Recordset.LockEdits = False: Carregando = False

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

On Error GoTo Trata_erro

If Save And Desistir = False Then

    If txt_nome = "" Then
        MsgBox "Informe Nome do Garçon!"
        txt_nome.SetFocus
        Editar = True
        Save = False
        Action = vbDataActionCancel
        Exit Sub
    End If
        
ElseIf Save And Desistir = True Then

    Data1.Recordset.CancelUpdate
            
End If

'desabilita campos para edição e habilita botões
   
Bt_novo.Enabled = True
Bt_Excluir.Enabled = True
Bt_Alterar.Enabled = True
Bt_Sair.Enabled = True
    
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

Private Sub Form_Load()

Call Abrir_BD_Data(Data1, "tbl_Garcons", "[Nome_Garcon]", "")
Carregando = True

Set db1 = OpenDatabase(Caminho_Rede & "\dados.mdb")

End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_mnu.barramenu.Visible = True
End Sub

Sub Atualiza_Cad_Garcons()

On Error GoTo Trata_erro

Me.MousePointer = 11

'declara e inicia conexão
Set conn = New ADODB.Connection
conn.ConnectionString = StringConexao
conn.CursorLocation = adUseClient
conn.Open

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Garcons", conn

If rs.EOF Then
    MsgBox "Não existem dados para importar", vbExclamation, "Atenção"
    'fecha conexão
    conn.Close
    Me.MousePointer = 0
    Exit Sub
End If

'apaga registros atuais
db1.Execute "delete * from [tbl_Garcons]"

'abre arquivo para receber novos dados
Set Tab1 = db1.OpenRecordset("select * from tbl_Garcons")

Do While Not rs.EOF
    With Tab1
        .AddNew
        !ID = rs!ID_garcon
        !Nome_Garcon = rs!Nome
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

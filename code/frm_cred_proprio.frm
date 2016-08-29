VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frm_cred_proprio 
   Caption         =   "Acompanhamento Crediario Próprio"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame frame_RECEBER 
      Caption         =   "RECEBIMENTO"
      ForeColor       =   &H000000FF&
      Height          =   2895
      Left            =   3480
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   8415
      Begin VB.TextBox txt_receber 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txt_fiado_contato 
         DataField       =   "Contato"
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
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   7815
      End
      Begin VB.CommandButton bt_fiado_cancel 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   6840
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton bt_fiado_ok 
         Caption         =   "Confirmar"
         Height          =   495
         Left            =   5160
         TabIndex        =   7
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txt_fiado_cli 
         DataField       =   "Responsavel"
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
         Left            =   360
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   7815
      End
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
         Left            =   360
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Saldo A RECEBER:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   14
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Responsável"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   15
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Valor RECEBIDO:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   1275
      End
   End
   Begin VB.CommandButton bt_receb 
      Caption         =   "Recebimento"
      Height          =   855
      Left            =   11520
      Picture         =   "frm_cred_proprio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Lançar Recebimento de Crediário Especial"
      Top             =   6720
      Width           =   1455
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   13800
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.CommandButton bt_imprimir 
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   13320
      Picture         =   "frm_cred_proprio.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Reposição de Estoque"
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Bt_Sair 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   855
      Left            =   14640
      Picture         =   "frm_cred_proprio.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Fechar esta Janela"
      Top             =   6720
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   -360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1250
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_cred_proprio.frx":0CC6
      Height          =   6255
      Left            =   240
      OleObjectBlob   =   "frm_cred_proprio.frx":0CD6
      TabIndex        =   0
      Top             =   240
      Width           =   15255
   End
End
Attribute VB_Name = "frm_cred_proprio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database

'conexão banco MySQL
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cloud_erro As Long
Dim Cloud_erro_desc As String
Dim TabErr As Recordset
Dim str_instrucao As String

Private Sub bt_fiado_cancel_Click()

frame_RECEBER.Visible = False

End Sub

Private Sub bt_fiado_ok_Click()

'validações
If Not IsNumeric(txt_fiado_valor) Then MsgBox "Valor Inválido", vbExclamation, "Atenção": txt_fiado_valor.SetFocus: Exit Sub

Dim resp As String
resp = Data1.Recordset!Responsavel

'saldo a pagar
With Data1.Recordset
    .Edit
    !total = Data1.Recordset!total - CCur(txt_fiado_valor)
    .Update
End With

Dim DataLanc As String
DataLanc = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")

'lança recebimento
With Data1.Recordset
    .AddNew
    !Data = DataLanc
    !Cod_Operador = Cod_Operador
        
    !Descrição = "RECEB.CRED.PROPRIO: " & resp
    
    !valor = CCur(txt_fiado_valor)
    !Quant = 1
    !total = CCur(txt_fiado_valor)
    
    !Forma_Pagam = "D"
    !Tipo = "C"
    !caixa = NumCaixa
    
    !Recebimento = Date
    Dim DataReceb As String
    DataReceb = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
    .Update
End With

'lança pagamento em tabela cloud MySQL
Dim vtotal As String
vtotal = LTrim(Str(ValorPag))
'Call ConnMySQL_InserirLançamento(DataLanc, Mesa, "0", "RECEBIMENTO: CRED.PROPRIO", vtotal, "1", vtotal, "", "C", Cod_Operador, "D", NumCaixa, "0", "", DataReceb)

Data1.Refresh
frame_RECEBER.Visible = False

End Sub

Private Sub bt_imprimir_Click()

If Data1.Recordset.EOF Then Exit Sub

With CrystalReport1
    .ReportFileName = Caminho_Rede & "\crediario.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .SelectionFormula = "{tbl_lancamentos.Descrição} = 'FECHAMENTO: CRED.PROPRIO'"
    .Action = 1
End With

End Sub

Private Sub bt_receb_Click()

If Data1.Recordset.EOF Then Exit Sub

txt_fiado_valor.Text = Data1.Recordset!total
frame_RECEBER.Visible = True
txt_fiado_valor.SetFocus

End Sub

Private Sub Bt_Sair_Click()

Unload Me
frm_mnu.barramenu.Visible = True

End Sub

Private Sub Form_Load()


Set db1 = OpenDatabase(Caminho_Rede & "\dados.mdb")

Call Abrir_BD_Data(Data1, "tbl_lancamentos", "[Recebimento],[Garcon],[Responsavel]", "Descrição='FECHAMENTO: CRED.PROPRIO' and [total]>0")

End Sub

Private Sub txt_fiado_valor_GotFocus()
Call Selecionar(txt_fiado_valor)
End Sub

Private Sub txt_fiado_valor_LostFocus()

If CCur(txt_fiado_valor) > Data1.Recordset!total Then MsgBox "Valor Maior que valor do débito", vbExclamation, "Atenção": txt_fiado_valor.SetFocus: Exit Sub
txt_receber = Format(Data1.Recordset!total - CCur(txt_fiado_valor), "fixed")

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



VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frm_extrato 
   Caption         =   "Extrato"
   ClientHeight    =   8565
   ClientLeft      =   255
   ClientTop       =   345
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   14940
   WindowState     =   2  'Maximized
   Begin VB.CheckBox op_ocupadas 
      Caption         =   "Ocupadas"
      Height          =   255
      Left            =   360
      TabIndex        =   41
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   1250
   End
   Begin VB.Frame frame_encerramento2 
      Height          =   2055
      Left            =   11160
      TabIndex        =   21
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txt_saldo 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txt_mesa 
         Alignment       =   1  'Right Justify
         DataField       =   "Numero"
         DataSource      =   "Data2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton bt_trocar 
         Caption         =   "#"
         Height          =   360
         Left            =   1680
         TabIndex        =   27
         ToolTipText     =   "Trocar de Mesa"
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox op_op 
         Caption         =   "10% Opcional"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txt_saldo_taxa 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txt_opcional 
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         Caption         =   "Sub-Total:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Mesa:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Total da Conta:"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame frame_encerramento1 
      Height          =   2895
      Left            =   11160
      TabIndex        =   20
      Top             =   5280
      Width           =   3495
      Begin VB.CommandButton Bt_Sair 
         Cancel          =   -1  'True
         Caption         =   "Fechar"
         Height          =   495
         Left            =   240
         TabIndex        =   32
         ToolTipText     =   "Fechar esta Janela"
         Top             =   2160
         Width           =   3015
      End
      Begin VB.CommandButton bt_imprimir 
         Caption         =   "Extrato"
         Height          =   975
         Left            =   480
         Picture         =   "frm_extrato_foodcontrol.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Fechar esta Janela"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton bt_encerram_ok 
         Caption         =   "Encerrar Conta"
         Height          =   975
         Left            =   2040
         Picture         =   "frm_extrato_foodcontrol.frx":0C42
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5640
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.Frame frame_encerramento 
      Height          =   2895
      Left            =   11160
      TabIndex        =   19
      Top             =   2280
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "TOTAL PAGO"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton bt_amex 
         Caption         =   "Cartão AMEX"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton bt_hiper 
         Caption         =   "Cartão HIPER"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton bt_master 
         Caption         =   "Cartão MASTER"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton bt_visa 
         Caption         =   "Cartão VISA"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton bt_din 
         Caption         =   "Dinheiro "
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txt_total_pago 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txt_din 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox TXT_VISA 
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox txt_master 
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox txt_amex 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox txt_hiper 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Height          =   1935
      Left            =   2520
      TabIndex        =   13
      Top             =   120
      Width           =   8415
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
         TabIndex        =   0
         Top             =   480
         Width           =   7815
      End
      Begin VB.CommandButton bt_lançar 
         Caption         =   "Lançar"
         Height          =   495
         Left            =   6360
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txt_quant 
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
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txt_und 
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
         MaxLength       =   5
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txt_preço 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txt_total 
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Quant."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   12
         Left            =   3240
         TabIndex        =   17
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Preço:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   11
         Left            =   1560
         TabIndex        =   16
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Unidade:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   8
         Left            =   4680
         TabIndex        =   14
         Top             =   1080
         Width           =   405
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
      Left            =   8520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   1250
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_extrato_foodcontrol.frx":1884
      Height          =   6015
      Left            =   2520
      OleObjectBlob   =   "frm_extrato_foodcontrol.frx":1894
      TabIndex        =   6
      ToolTipText     =   "Duplo clique para excluir ítem"
      Top             =   2160
      Width           =   8415
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frm_extrato_foodcontrol.frx":25AE
      Height          =   7575
      Left            =   360
      OleObjectBlob   =   "frm_extrato_foodcontrol.frx":25BE
      TabIndex        =   34
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frm_extrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database
Dim Tab1 As Recordset   'auxiliar Mesas
Dim Tab2 As Recordset   'auxiliar hospedes
Dim Tab3 As Recordset   'auxiliar totais
Dim Tab4 As Recordset   'auxiliar produtos
Dim Tab5 As Recordset   'auxiliar Lançamentos
Dim Tab6 As Recordset   'auxiliar HORAS adicionais

Dim Tab7 As Recordset   'auxiliar Troca de MEsa
Dim Tab8 As Recordset   'tempo de ocupação da mesa - arquivo

Dim vtemp As Currency   'variavel temporaria

Dim Total_CR As Single
Dim Total_DB As Single

Private Sub bt_amex_Click()

vtemp = CCur(txt_saldo_taxa) - (CCur(TXT_VISA) + CCur(txt_master) + CCur(txt_hiper) + CCur(txt_din))
If vtemp < 0 Then vtemp = 0
txt_amex.Text = vtemp

End Sub

Private Sub bt_din_Click()

vtemp = CCur(txt_saldo_taxa) - (CCur(TXT_VISA) + CCur(txt_master) + CCur(txt_hiper) + CCur(txt_amex))
If vtemp < 0 Then vtemp = 0
txt_din.Text = vtemp

End Sub

Private Sub bt_encerram_ok_Click()

'validações
If txt_total_pago <> txt_saldo_taxa Then MsgBox "Conferir Valores", vbExclamation, "Atenção": txt_din.SetFocus: Exit Sub
If Data1.Recordset.EOF Then Exit Sub

'lança pagamento em tabela de lançamentos
If CCur(txt_din) <> 0 Then Call Lança_Pagamento("DINHEIRO", CCur(txt_din), "D")
If CCur(TXT_VISA) <> 0 Then Call Lança_Pagamento("CARTÃO VISA", CCur(TXT_VISA), "C")
If CCur(txt_master) <> 0 Then Call Lança_Pagamento("CARTÃO MASTER", CCur(txt_master), "C")
If CCur(txt_hiper) <> 0 Then Call Lança_Pagamento("CARTÃO HIPER", CCur(txt_hiper), "C")
If CCur(txt_amex) <> 0 Then Call Lança_Pagamento("CARTÃO AMEX", CCur(txt_amex), "C")

hsaida = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")

'tempo de duração de ocupação
With Tab8
    .AddNew
    !ID_mesa = Tab1!Numero
    !abertura = Tab1!abertura
    !Encerramento = hsaida
    !Duracao = DateDiff("n", Tab1!abertura, CDate(hsaida))
    .Update
End With

'Altera Status do Mesa
With Tab1
    .Edit
    !Status = "L"
    .Update
End With

'marca lançamentos como encerrados
db1.Execute "UPDATE tbl_lancamentos SET tbl_lancamentos.Encerrada = True " _
    & "WHERE (((tbl_lancamentos.Encerrada)=False) AND ((tbl_lancamentos.Mesa)=" & Mesa & "));"

MsgBox "Encerramento Efetuado com Sucesso", vbInformation, "Ok"
Data2.Recordset.MoveFirst

End Sub

Private Sub bt_encerrar_Click()

End Sub

Private Sub bt_hiper_Click()

vtemp = CCur(txt_saldo_taxa) - (CCur(TXT_VISA) + CCur(txt_master) + CCur(txt_din) + CCur(txt_amex))
If vtemp < 0 Then vtemp = 0
txt_hiper.Text = vtemp

End Sub

Private Sub bt_imprimir_Click()

If Data1.Recordset.EOF Then Exit Sub

pessoas = InputBox("Quantidade de Pagantes", "Divisão de Conta", 1)
If pessoas = "" Then Exit Sub
'If Not IsNumeric(pessoas) = "" Then MsgBox "Quantidade inválida", vbExclamation, "Atenção": Exit Sub

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

Private Sub bt_lançar_Click()

'validações
If cmb_prod.Text = "" Then MsgBox "Informe Produto", vbExclamation, "Atenção": cmb_prod.SetFocus: Exit Sub
If Not IsNumeric(txt_quant) Then MsgBox "Quantidade Incorreta", vbExclamation, "Atenção": txt_quant.SetFocus: Exit Sub
If Not IsNumeric(txt_total) Then MsgBox "Total Incorreto", vbExclamation, "Atenção": Exit Sub

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
With Data1.Recordset
    .AddNew
    
    !Data = Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss")
    !Mesa = Mesa
    !Descrição = cmb_prod
    
    !valor = txt_preço
    !quant = txt_quant
    !Total = txt_total
    !Obs = txt_obs
    
    !TIPO = "D"
    
    !Cod_Operador = Cod_Operador
    !Forma_Pagam = "T"
    
    .Update
End With

'baixa em estoque
With Tab4
    .FindFirst ("Descrição = '" & cmb_prod & "'")
    If .NoMatch Then MsgBox "Selecione um produto da lista", vbExclamation, "Atenção": cmb_prod.SetFocus: Exit Sub
    
    .Edit
    !Estoque = !Estoque - txt_quant
    .Update
End With

'TOTAIS
Total_DB = Total_DB + CCur(txt_total)
txt_saldo = Format(Total_DB, "CURRENCY")

If op_op.Value = 1 Then txt_opcional.Text = CCur(txt_saldo) * 0.1 Else txt_opcional.Text = 0
txt_saldo_taxa = Format(CCur(txt_saldo) + CCur(txt_opcional), "currency")


cmb_prod.Text = ""
txt_und.Text = ""
txt_preço.Text = ""
txt_quant.Text = ""
txt_total.Text = ""

Data1.Refresh

End Sub

Private Sub bt_master_Click()

vtemp = CCur(txt_saldo_taxa) - (CCur(TXT_VISA) + CCur(txt_din) + CCur(txt_hiper) + CCur(txt_amex))
If vtemp < 0 Then vtemp = 0
txt_master.Text = vtemp

End Sub

Private Sub Bt_Sair_Click()
'Call frm_mapa.Atualiza_Mapa
Unload Me
End Sub

Private Sub bt_trocar_Click()

If Data1.Recordset.EOF Then Exit Sub

'validações
novamesa = InputBox("Nova Mesa", "Mudança de Mesa")
If novamesa = "" Then Exit Sub
If novamesa = Mesa Then Exit Sub

Set Tab7 = db1.OpenRecordset("select * from [Tbl_Mesas] where [numero]=" & novamesa)
If Tab7.EOF Then MsgBox "Mesa Inválida: " & novamesa, vbExclamation, "Atenção": Exit Sub

'altera mesa em Lançamentos
db1.Execute "UPDATE tbl_lancamentos SET tbl_lancamentos.Mesa = " & novamesa _
    & " WHERE (((tbl_lancamentos.Mesa)=" & Mesa & ") AND ((tbl_lancamentos.Encerrada)=False));"

db1.Execute "UPDATE Tbl_Mesas SET Tbl_Mesas.Status = 'O' WHERE (((Tbl_Mesas.Numero)=" & novamesa & "));"

'Altera Status do Mesa para LIVRE
With Tab1
    .Edit
    !Status = "L"
    .Update
End With

Data2.Recordset.FindFirst ("Numero =" & novamesa)

End Sub

Private Sub bt_visa_Click()

vtemp = CCur(txt_saldo_taxa) - (CCur(txt_din) + CCur(txt_master) + CCur(txt_hiper) + CCur(txt_amex))
If vtemp < 0 Then vtemp = 0
TXT_VISA.Text = vtemp

End Sub

Private Sub cmb_prod_Click()

'dados do produto

If cmb_prod.Text = "" Then Exit Sub

With Tab4
    .FindFirst ("Descrição = '" & cmb_prod & "'")
    If .NoMatch Then MsgBox "Selecione um produto da lista", vbExclamation, "Atenção": cmb_prod.SetFocus: Exit Sub
    
    txt_und = "" & !Unidade
    txt_preço = Format(!Preço, "currency")
End With

txt_quant.Text = 1
txt_quant.SetFocus
Call txt_quant_LostFocus

End Sub

Private Sub Data2_Reposition()

If Data2.Recordset.EOF Then
    txt_saldo = ""
    txt_opcional.Text = ""
    txt_saldo_taxa = ""
    Mesa = 0
Else
    Mesa = Data2.Recordset!Numero
End If

cmb_prod.Text = ""
txt_quant.Text = 0

Call Carrega_Mesa

End Sub

Private Sub DBGrid1_DblClick()

On Error GoTo Trata_erro

'validações
If Data1.Recordset.EOF Then Exit Sub
If Nível > 1 Then MsgBox "Usuário não autorizado para esta operação", vbExclamation, "Atenção": Exit Sub

If Conf("Confirma Exclusão de Item : " & Data1.Recordset!Descrição & " ?", "Atenção") = 7 Then Exit Sub

'TOTAL DE DESPESAS e saldo
Total_DB = Total_DB - Data1.Recordset!Total
txt_saldo = Format(Total_DB, "currency")

If op_op.Value = 1 Then txt_opcional.Text = CCur(txt_saldo) * 0.1 Else txt_opcional.Text = 0
txt_saldo_taxa = Format(CCur(txt_saldo) + CCur(txt_opcional), "currency")


'estorno em estoque
With Tab4
    .FindFirst ("Descrição = '" & Data1.Recordset!Descrição & "'")
    If Not .NoMatch Then
        .Edit
        !Estoque = !Estoque + Data1.Recordset!quant
        .Update
    End If
End With

Data1.Recordset.Delete

Exit Sub
Trata_erro:
Exit Sub

End Sub

Private Sub Form_Load()

Set db1 = OpenDatabase(Caminho_Rede & "\dados.mdb")

'Mesas
Call Abrir_BD_Data(Data2, "tbl_mesas", "[Numero]", "")

'LANÇAMENTOS do Mesa
Call Abrir_BD_Data(Data1, "tbl_lancamentos", "[data] desc", "[Mesa]=0")

'tempo de ocupação da mesa
Set Tab8 = db1.OpenRecordset("select * from [tbl_Mesas_Duracao]")

'combo Produto
Set Tab4 = db1.OpenRecordset("select * from [Tbl_Produtos] order by [Grupo],[Descrição]")
Do While Not Tab4.EOF
    cmb_prod.AddItem ("" & Tab4!Descrição)
    Tab4.MoveNext
Loop

Total_DB = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_mnu.barramenu.Visible = True
End Sub

Private Sub op_ocupadas_Click()

If op_ocupadas.Value = 1 Then
    Data2.RecordSource = "select * from [Tbl_Mesas] where [status] ='O' order by [Numero]"
Else
    Data2.RecordSource = "select * from [Tbl_Mesas] where [status] ='L' order by [Numero]"
End If
Data2.Refresh

End Sub

Private Sub op_op_Click()

If op_op.Value = 1 Then
    txt_opcional.Text = CCur(txt_saldo) * 0.1
Else
    txt_opcional.Text = 0
End If

If IsNumeric(txt_saldo) Then
    txt_saldo_taxa = Format(CCur(txt_saldo) + CCur(txt_opcional), "currency")
Else
    txt_saldo_taxa = ""
End If

End Sub

Private Sub txt_amex_Change()
Call Calcula_Total_Pago
End Sub

Private Sub txt_amex_GotFocus()
Call Selecionar(txt_amex)
End Sub

Private Sub txt_din_Change()
Call Calcula_Total_Pago
End Sub

Private Sub txt_din_GotFocus()
Call Selecionar(txt_din)
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

Private Sub txt_opcional_LostFocus()

If Not IsNumeric(txt_opcional.Text) Then txt_opcional.Text = 0
txt_saldo_taxa = Format(CCur(txt_saldo) + CCur(txt_opcional), "currency")

End Sub

Private Sub txt_quant_GotFocus()
Call Selecionar(txt_quant)
End Sub

Private Sub txt_quant_LostFocus()

If Not IsNumeric(txt_preço) Then Exit Sub
If txt_quant = "" Then Exit Sub

If Not IsNumeric(txt_quant) Then MsgBox "Quantidade Incorreta", vbExclamation, "Atenção": txt_quant.SetFocus: Exit Sub
txt_total = Format(CCur(txt_quant) * CCur(txt_preço), "currency")

End Sub


Sub Calcula_Total_Pago()

If Not IsNumeric(txt_din) Then txt_din.Text = 0
If Not IsNumeric(TXT_VISA) Then TXT_VISA.Text = 0
If Not IsNumeric(txt_master) Then txt_master.Text = 0
If Not IsNumeric(txt_hiper) Then txt_hiper.Text = 0
If Not IsNumeric(txt_amex) Then txt_amex.Text = 0
If Not IsNumeric(txt_opcional) Then txt_opcional.Text = 0

txt_total_pago = Format(CCur(txt_din) + CCur(TXT_VISA) + CCur(txt_master) + CCur(txt_hiper) + CCur(txt_amex), "currency")

End Sub

Private Sub TXT_VISA_Change()
Call Calcula_Total_Pago
End Sub

Private Sub TXT_VISA_GotFocus()
Call Selecionar(TXT_VISA)
End Sub

Sub Lança_Pagamento(DescriPag As String, ValorPag As Currency, FormaPag As String)

With Data1.Recordset
    .AddNew
    !Data = Date
    !Mesa = Mesa
    !Cod_Operador = Cod_Operador
        
    !Descrição = "FECHAMENTO: " & DescriPag
    
    !valor = ValorPag
    !quant = 1
    !Total = ValorPag
    
    !Forma_Pagam = FormaPag
    !TIPO = "C"
        
    .Update
End With

End Sub

Sub Carrega_Mesa()

 txt_saldo = ""
 txt_opcional.Text = ""
 txt_saldo_taxa = ""
 Total_DB = 0

txt_din.Text = 0
TXT_VISA.Text = 0
txt_master.Text = 0
txt_hiper.Text = 0
txt_amex.Text = 0

'dados do Mesa
Set Tab1 = db1.OpenRecordset("select * from [Tbl_Mesas] where [Numero] =" & Mesa)

'LANÇAMENTOS do Mesa
Data1.RecordSource = "select * from [tbl_lancamentos] where [Mesa]=" & Mesa & " and [Encerrada]=false"
Data1.Refresh


'AUXILIAR lançamentos
Set Tab5 = db1.OpenRecordset("select * from [tbl_lancamentos] where [Mesa]=" & Mesa & " and [Encerrada]=false")

On Error Resume Next

'totais de débito
Set Tab3 = db1.OpenRecordset("SELECT Sum(tbl_lancamentos.Total) AS DB From tbl_lancamentos " _
    & "WHERE (((tbl_lancamentos.Tipo)='D') AND ((tbl_lancamentos.Mesa)=" & Mesa & " and [encerrada]=false ));")
If Not Tab3.EOF Then Total_DB = Format(Tab3!DB, "currency")

txt_saldo = Format(Total_DB, "currency")
If op_op.Value = 1 Then txt_opcional.Text = CCur(txt_saldo) * 0.1 Else txt_opcional.Text = 0
txt_saldo_taxa = Format(CCur(txt_saldo) + CCur(txt_opcional), "currency")

End Sub

VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.MDIForm frm_mnu 
   BackColor       =   &H8000000C&
   Caption         =   "Food Control"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frm_menu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8640
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.PictureBox barramenu 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   4620
      TabIndex        =   1
      Top             =   0
      Width           =   4680
      Begin VB.CommandButton bt2 
         Height          =   615
         Left            =   1080
         Picture         =   "frm_menu.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Caixa"
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton bt_mapa 
         Height          =   615
         Left            =   240
         Picture         =   "frm_menu.frx":09FE
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Mapa de Mesas"
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Menu mnu_oper 
      Caption         =   "&Operacional"
      Begin VB.Menu mnu_oper_mapa 
         Caption         =   "Mapa de Mesas"
      End
      Begin VB.Menu mnu_oper_caixa 
         Caption         =   "Caixa"
      End
      Begin VB.Menu BARCAD 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_cred 
         Caption         =   "Controle de Crediário Próprio"
      End
   End
   Begin VB.Menu mnu_cad 
      Caption         =   "&Cadastros"
      Begin VB.Menu mnu_cad_mesas 
         Caption         =   "Mesas"
      End
      Begin VB.Menu mnu_cad_garçons 
         Caption         =   "Garçons"
      End
      Begin VB.Menu barcad2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_cad_insumos 
         Caption         =   "Insumos"
      End
      Begin VB.Menu mnu_cad_Prod 
         Caption         =   "Produtos"
      End
   End
   Begin VB.Menu op_ger 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnu_caixa_per 
         Caption         =   "Mov. de Caixa no Período"
      End
      Begin VB.Menu mnu_rel_prod 
         Caption         =   "Produtos Vendidos no Periodo"
      End
      Begin VB.Menu barrel1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_rel_cartões 
         Caption         =   "Cartões à Receber"
      End
      Begin VB.Menu mnu_rel_cred_proprio 
         Caption         =   "Crediário Próprio"
      End
   End
   Begin VB.Menu mnu_sist 
      Caption         =   "&Sistema"
      Begin VB.Menu mnu_sist_refresh 
         Caption         =   "Atualizar Dados"
      End
      Begin VB.Menu barsist2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Sist_testeDaruma 
         Caption         =   "Testar Impressora (DARUMA)"
      End
      Begin VB.Menu mnu_sist_usu 
         Caption         =   "Usuários e Senhas"
      End
      Begin VB.Menu barsist 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sair 
         Caption         =   "&Sair"
      End
   End
End
Attribute VB_Name = "frm_mnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bt_mapa_Click()
Call mnu_oper_mapa_Click
End Sub

Private Sub bt2_Click()
Call mnu_oper_caixa_Click
End Sub

Private Sub MDIForm_Load()
Me.Caption = Me.Caption & " - Operador: " & Usuário
End Sub

Private Sub mnu_cad_garçons_Click()

barramenu.Visible = False
frm_garcons.Show

End Sub

Private Sub mnu_cad_insumos_Click()

Rotina = "S"   'insumo = SIM
barramenu.Visible = False
frm_produtos.Show

End Sub

Private Sub mnu_cad_mesas_Click()

barramenu.Visible = False
frm_mesas.Show

End Sub

Private Sub mnu_cad_Prod_Click()

Rotina = "N"   'insumo = NÂO
barramenu.Visible = False
frm_produtos.Show

End Sub

Private Sub mnu_caixa_per_Click()

per1 = InputBox("Data Inicial", "Relatório de Movimento de Caixa", "01/" & Format(Date, "mm/yy"))
If Not IsDate(per1) Then MsgBox "Data Inválida", vbExclamation, "Atenção": Exit Sub

per2 = InputBox("Data Final", "Relatório de Movimento de Caixa", Format(Date, "dd/mm/yy"))
If Not IsDate(per2) Then MsgBox "Data Inválida", vbExclamation, "Atenção": Exit Sub

With CrystalReport1
    .ReportFileName = Caminho_Rede & "\caixas.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .Formulas(2) = "periodo = 'Período: " & per1 & " a " & per2 & "'"
    .SelectionFormula = "{Tbl_Caixas.Data} in Date (" & Format(CDate(per1), "yyyy,mm,dd") & ") to Date (" & Format(CDate(per2), "yyyy,mm,dd") & ")"
    .Action = 1
End With
    
End Sub

Private Sub mnu_cred_Click()

barramenu.Visible = False
frm_cred_proprio.Show

End Sub

Private Sub mnu_oper_caixa_Click()
If Nível > 1 Then MsgBox "Usuário não autorizado para esta operação", vbExclamation, "Atenção": Exit Sub
barramenu.Visible = False
frm_caixa.Show
End Sub

Private Sub mnu_oper_mapa_Click()

If NumCaixa = 0 Then MsgBox "Necessário abrir caixa", vbExclamation, "Atenção": Exit Sub

barramenu.Visible = False
'frm_mapa.Show
frm_extrato.Show

End Sub


Private Sub mnu_rel_cartões_Click()

per1 = InputBox("Data Inicial", "Relatório de Cartões à Receber", Format(Date, "dd/mm/yy"))
If Not IsDate(per1) Then MsgBox "Data Inválida", vbExclamation, "Atenção": Exit Sub

per2 = InputBox("Data Final", "Relatório de Cartões à Receber", Format(Date + 31, "dd/mm/yy"))
If Not IsDate(per2) Then MsgBox "Data Inválida", vbExclamation, "Atenção": Exit Sub

With CrystalReport1
    .ReportFileName = Caminho_Rede & "\cartoes.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .Formulas(2) = "periodo = 'Período: " & per1 & " a " & per2 & "'"
    .SelectionFormula = "{tbl_lancamentos.recebimento} in Date (" & Format(CDate(per1), "yyyy,mm,dd") & ") to Date (" & Format(CDate(per2), "yyyy,mm,dd") & ")"
    .Action = 1
End With

End Sub

Private Sub mnu_rel_cred_proprio_Click()

With CrystalReport1
    .ReportFileName = Caminho_Rede & "\crediario.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .Formulas(2) = ""
    .SelectionFormula = "{tbl_lancamentos.Descrição} = 'FECHAMENTO: CRED.PROPRIO' and {tbl_lancamentos.Tipo} = 'C'"
    .Action = 1
End With

End Sub

Private Sub mnu_rel_prod_Click()

d_ini = "01/" & Format(Date, "mm/yyyy")

data_ini = InputBox("Data Inicial", "Produtos vendidos no período", d_ini)
If data_ini = "" Then Exit Sub
If Not IsDate(data_ini) Then MsgBox "Data inválida", vbExclamation, "Atenção": Exit Sub

data_fim = InputBox("Data Final", "Produtos vendidos no período", Date)
If data_fim = "" Then Exit Sub
If Not IsDate(data_fim) Then MsgBox "Data inválida", vbExclamation, "Atenção": Exit Sub

per1 = Format(CDate(data_ini), "yyyy,mm,dd")
per2 = Format(CDate(data_fim), "yyyy,mm,dd")

With CrystalReport1
    .ReportFileName = Caminho_Rede & "\produtos.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .Formulas(2) = "periodo = 'Período: " & data_ini & " a " & data_fim & "'"
    .SelectionFormula = "{tbl_lancamentos.data} in Date (" & Format(CDate(per1), "yyyy,mm,dd") & ") to Date (" & Format(CDate(per2), "yyyy,mm,dd") & ")"
    .Action = 1
End With

End Sub

Private Sub mnu_sair_Click()
Call Sair
End Sub

Private Sub mnu_sist_refresh_Click()

barramenu.Visible = False
frm_cloud_refresh.Show

End Sub

Private Sub mnu_Sist_testeDaruma_Click()

Dim iRetorno As Integer
    
iRetorno = rStatusImpressora_DUAL_DarumaFramework()
  
Select Case (iRetorno)
    Case 0:     MsgBox "[0] - Impressora desligada"
    Case 1:     MsgBox "[1] - Impressora OK"
    Case -27:     MsgBox "[-27] - Erro generico"
    Case -50:     MsgBox "[-50] - Impressora OFFLINE"
    Case -51:     MsgBox "[-51] - Impressora sem papel!"
    Case -52:     MsgBox "[-52] - Impressora inicializando"
End Select

End Sub

Private Sub mnu_sist_usu_Click()

If Nível > 1 Then MsgBox "Usuário não autorizado para esta operação", vbExclamation, "Atenção": Exit Sub
frm_Usuários.Show

End Sub

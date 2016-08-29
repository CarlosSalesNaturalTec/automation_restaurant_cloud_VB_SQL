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
         Caption         =   "Controle de Credi�rio Pr�prio"
      End
   End
   Begin VB.Menu mnu_cad 
      Caption         =   "&Cadastros"
      Begin VB.Menu mnu_cad_mesas 
         Caption         =   "Mesas"
      End
      Begin VB.Menu mnu_cad_gar�ons 
         Caption         =   "Gar�ons"
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
      Caption         =   "&Relat�rios"
      Begin VB.Menu mnu_caixa_per 
         Caption         =   "Mov. de Caixa no Per�odo"
      End
      Begin VB.Menu mnu_rel_prod 
         Caption         =   "Produtos Vendidos no Periodo"
      End
      Begin VB.Menu barrel1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_rel_cart�es 
         Caption         =   "Cart�es � Receber"
      End
      Begin VB.Menu mnu_rel_cred_proprio 
         Caption         =   "Credi�rio Pr�prio"
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
         Caption         =   "Usu�rios e Senhas"
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
Me.Caption = Me.Caption & " - Operador: " & Usu�rio
End Sub

Private Sub mnu_cad_gar�ons_Click()

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

Rotina = "N"   'insumo = N�O
barramenu.Visible = False
frm_produtos.Show

End Sub

Private Sub mnu_caixa_per_Click()

per1 = InputBox("Data Inicial", "Relat�rio de Movimento de Caixa", "01/" & Format(Date, "mm/yy"))
If Not IsDate(per1) Then MsgBox "Data Inv�lida", vbExclamation, "Aten��o": Exit Sub

per2 = InputBox("Data Final", "Relat�rio de Movimento de Caixa", Format(Date, "dd/mm/yy"))
If Not IsDate(per2) Then MsgBox "Data Inv�lida", vbExclamation, "Aten��o": Exit Sub

With CrystalReport1
    .ReportFileName = Caminho_Rede & "\caixas.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .Formulas(2) = "periodo = 'Per�odo: " & per1 & " a " & per2 & "'"
    .SelectionFormula = "{Tbl_Caixas.Data} in Date (" & Format(CDate(per1), "yyyy,mm,dd") & ") to Date (" & Format(CDate(per2), "yyyy,mm,dd") & ")"
    .Action = 1
End With
    
End Sub

Private Sub mnu_cred_Click()

barramenu.Visible = False
frm_cred_proprio.Show

End Sub

Private Sub mnu_oper_caixa_Click()
If N�vel > 1 Then MsgBox "Usu�rio n�o autorizado para esta opera��o", vbExclamation, "Aten��o": Exit Sub
barramenu.Visible = False
frm_caixa.Show
End Sub

Private Sub mnu_oper_mapa_Click()

If NumCaixa = 0 Then MsgBox "Necess�rio abrir caixa", vbExclamation, "Aten��o": Exit Sub

barramenu.Visible = False
'frm_mapa.Show
frm_extrato.Show

End Sub


Private Sub mnu_rel_cart�es_Click()

per1 = InputBox("Data Inicial", "Relat�rio de Cart�es � Receber", Format(Date, "dd/mm/yy"))
If Not IsDate(per1) Then MsgBox "Data Inv�lida", vbExclamation, "Aten��o": Exit Sub

per2 = InputBox("Data Final", "Relat�rio de Cart�es � Receber", Format(Date + 31, "dd/mm/yy"))
If Not IsDate(per2) Then MsgBox "Data Inv�lida", vbExclamation, "Aten��o": Exit Sub

With CrystalReport1
    .ReportFileName = Caminho_Rede & "\cartoes.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .Formulas(2) = "periodo = 'Per�odo: " & per1 & " a " & per2 & "'"
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
    .SelectionFormula = "{tbl_lancamentos.Descri��o} = 'FECHAMENTO: CRED.PROPRIO' and {tbl_lancamentos.Tipo} = 'C'"
    .Action = 1
End With

End Sub

Private Sub mnu_rel_prod_Click()

d_ini = "01/" & Format(Date, "mm/yyyy")

data_ini = InputBox("Data Inicial", "Produtos vendidos no per�odo", d_ini)
If data_ini = "" Then Exit Sub
If Not IsDate(data_ini) Then MsgBox "Data inv�lida", vbExclamation, "Aten��o": Exit Sub

data_fim = InputBox("Data Final", "Produtos vendidos no per�odo", Date)
If data_fim = "" Then Exit Sub
If Not IsDate(data_fim) Then MsgBox "Data inv�lida", vbExclamation, "Aten��o": Exit Sub

per1 = Format(CDate(data_ini), "yyyy,mm,dd")
per2 = Format(CDate(data_fim), "yyyy,mm,dd")

With CrystalReport1
    .ReportFileName = Caminho_Rede & "\produtos.rpt"
    .Formulas(0) = "empresa = '" & Empresa_Nome & "'"
    .Formulas(1) = "endereco = '" & Empresa_End & "'"
    .Formulas(2) = "periodo = 'Per�odo: " & data_ini & " a " & data_fim & "'"
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

If N�vel > 1 Then MsgBox "Usu�rio n�o autorizado para esta opera��o", vbExclamation, "Aten��o": Exit Sub
frm_Usu�rios.Show

End Sub

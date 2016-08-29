VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_produtos_Repor 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reposição de Estoque"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9195
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bt_repor 
      Caption         =   "Lançar"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton bt_fechar 
      Caption         =   "Fechar"
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   4320
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6588
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lançar Reposição"
      TabPicture(0)   =   "frm_produtos_reposição.frx":0000
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Reposições Lançadas"
      TabPicture(1)   =   "frm_produtos_reposição.frx":001C
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Data1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DBGrid1"
      Tab(1).Control(1).Enabled=   0   'False
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   -74760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2880
         Visible         =   0   'False
         Width           =   1250
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_produtos_reposição.frx":0038
         Height          =   2895
         Left            =   -74760
         OleObjectBlob   =   "frm_produtos_reposição.frx":0048
         TabIndex        =   10
         Top             =   600
         Width           =   8175
      End
      Begin VB.Frame Frame1 
         Height          =   3015
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   8055
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
            Left            =   240
            MaxLength       =   5
            TabIndex        =   2
            Top             =   2160
            Width           =   1335
         End
         Begin VB.ComboBox cmb_forn 
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
            TabIndex        =   1
            Top             =   1320
            Width           =   7455
         End
         Begin MSMask.MaskEdBox txt_data 
            Height          =   420
            Left            =   240
            TabIndex        =   0
            Top             =   600
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
            PromptChar      =   "_"
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   9
            Top             =   1920
            Width           =   870
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   390
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Fornecedor:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   7
            Top             =   1080
            Width           =   855
         End
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   5400
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   7440
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3735
      Index           =   1
      Left            =   360
      Top             =   360
      Width           =   8655
   End
End
Attribute VB_Name = "frm_produtos_Repor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database
Dim Tab1 As Recordset   'histórico de reposições
Dim Tab2 As Recordset   'fornecedores
Dim Tab3 As Recordset   'produtos

Private Sub bt_fechar_Click()
Unload Me
End Sub

Private Sub bt_repor_Click()

'validações
If Not IsDate(txt_data) Then MsgBox "Data Inválida", vbExclamation, "Atenção": txt_data.SetFocus: Exit Sub
If cmb_forn.Text = "" Then MsgBox "Informe nome do Fornecedor (ou digite: Estoque Inicial)", vbExclamation, "Atenção": cmb_forn.SetFocus: Exit Sub
If Not IsNumeric(txt_quant) Then MsgBox "Quantidade Inválida", vbExclamation, "Atenção": txt_quant.SetFocus: Exit Sub

'cadastro do produto
Set Tab3 = db1.OpenRecordset("select [estoque],[Código] from [tbl_produtos] where [Código]= " & Cod_Produto_Repor)
If Tab3.EOF Then MsgBox "Produto Não localizado", vbExclamation, "Atenção": Exit Sub

'lança em histórico de reposição
With Tab1
    .AddNew
    !Cod_Produto = Cod_Produto_Repor
    !Data = txt_data
    !FORNECEDOR = cmb_forn
    !Quant = txt_quant
    !Operador = Usuário
    .Update
End With

'registra entrada no estoque
With Tab3
    .Edit
    !estoque = !estoque + txt_quant
    .Update
End With

'atualiza form
With frm_produtos.Data1
    .Refresh
    .Recordset.FindFirst ("Código = " & Cod_Produto_Repor)
End With

MsgBox "Reposição Registrada com Sucesso", vbInformation, "Ok"
Unload Me

End Sub

Private Sub Form_Load()

Set db1 = OpenDatabase(Caminho_Rede & "\dados.mdb")
Set Tab1 = db1.OpenRecordset("select * from [Tbl_Produtos_Reposições]")

'histórico de reposições do produto
Call Abrir_BD_Data(Data1, "Tbl_Produtos_Reposições", "[Data],[fornecedor]", "Cod_Produto = " & Cod_Produto_Repor)

'combo fornecedores
Set Tab2 = db1.OpenRecordset("SELECT Tbl_Produtos_Reposições.Fornecedor From Tbl_Produtos_Reposições " _
    & "GROUP BY Tbl_Produtos_Reposições.Fornecedor;")
If Tab2.EOF Then cmb_forn.AddItem ("ESTOQUE INICIAL")
Do While Not Tab2.EOF
    If Tab2!FORNECEDOR <> "Ajuste de Estoque" Then cmb_forn.AddItem ("" & Tab2!FORNECEDOR)
    Tab2.MoveNext
Loop

txt_data = Format(Date, "dd/mm/yy")

End Sub

Private Sub Form_Unload(Cancel As Integer)
db1.Close
End Sub

Private Sub txt_data_GotFocus()
Call Mask_Data(txt_data)
Call Selecionar(txt_data)
End Sub


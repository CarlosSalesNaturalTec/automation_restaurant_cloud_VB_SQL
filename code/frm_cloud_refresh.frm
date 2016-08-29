VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_cloud_refresh 
   Caption         =   "Sincronizar Dados"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   14208
      _Version        =   327680
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Envio de Dados"
      TabPicture(0)   =   "frm_cloud_refresh.frx":0000
      Tab(0).ControlCount=   8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape1(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DBGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt_status"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Data1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cloud_erro"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "bt_envio"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Bt_Sair"
      Tab(0).Control(7).Enabled=   0   'False
      Begin VB.CommandButton Bt_Sair 
         Caption         =   "Fechar"
         Height          =   855
         Left            =   8400
         Picture         =   "frm_cloud_refresh.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar esta Janela"
         Top             =   6900
         Width           =   855
      End
      Begin VB.CommandButton bt_envio 
         Caption         =   "Enviar Dados"
         Height          =   855
         Left            =   6480
         Picture         =   "frm_cloud_refresh.frx":045E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Enviar"
         Top             =   6900
         Width           =   1695
      End
      Begin VB.Frame cloud_erro 
         Caption         =   "Detalhes"
         Height          =   2295
         Left            =   240
         TabIndex        =   3
         Top             =   4440
         Width           =   9135
         Begin VB.TextBox Text2 
            DataField       =   "Instrucao"
            DataSource      =   "Data1"
            Height          =   615
            Left            =   1200
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   1440
            Width           =   7695
         End
         Begin VB.TextBox Text1 
            DataField       =   "cloud_erro_desc"
            DataSource      =   "Data1"
            Height          =   495
            Left            =   1200
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   840
            Width           =   7695
         End
         Begin VB.TextBox txt_cloud_erro 
            DataField       =   "cloud_erro"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txt_data 
            DataField       =   "data"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Instrução:"
            Height          =   195
            Left            =   360
            TabIndex        =   11
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   360
            TabIndex        =   10
            Top             =   840
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cód.:"
            Height          =   195
            Left            =   3600
            TabIndex        =   9
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Width           =   390
         End
      End
      Begin VB.Data Data1 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   8040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1260
         Visible         =   0   'False
         Width           =   1250
      End
      Begin VB.TextBox txt_status 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   9135
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_cloud_refresh.frx":08A0
         Height          =   3015
         Left            =   240
         OleObjectBlob   =   "frm_cloud_refresh.frx":08B0
         TabIndex        =   2
         Top             =   1200
         Width           =   9135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   855
         Index           =   6
         Left            =   8520
         Top             =   7020
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   855
         Index           =   0
         Left            =   6600
         Top             =   7020
         Width           =   1695
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   8055
      Index           =   1
      Left            =   360
      Top             =   360
      Width           =   9615
   End
End
Attribute VB_Name = "frm_cloud_refresh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database
Dim Tab1 As Recordset 'auxiliar importações

Dim Erro_Reenvio As Boolean

Private Sub bt_envio_Click()

If Data1.Recordset.EOF Then Exit Sub

If Conf("Confirma Envio de Dados (Pedidos & Caixa)?", "Atenção") = 7 Then Exit Sub
Erro_Reenvio = False

txt_status.Text = "ENVIANDO Movimentação..."

Me.MousePointer = 11

With Data1.Recordset
    .MoveFirst
    Do While Not .EOF
        Call Reenviar
        If Erro_Reenvio = True Then Exit Sub
        
        'marca como executado
        marca = .Bookmark
        .Edit
        !Executada = True
        .Update
        .Bookmark = marca
        
        .MoveNext
    Loop
End With

Me.MousePointer = 0
Data1.Refresh
txt_status.Text = txt_status.Text & " - DADOS ENVIADOS!"

End Sub

Private Sub Bt_Sair_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Call Bt_Sair_Click
End Sub

Private Sub Form_Load()

Call Abrir_BD_Data(Data1, "Tbl_lancamentos_instrucoes", "Data", "Executada=false")

Set db1 = OpenDatabase(Caminho_Rede & "\dados.mdb")

End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_mnu.barramenu.Visible = True
End Sub

Sub Reenviar()

If Data1.Recordset.EOF Then Exit Sub

On Error GoTo Trata_erro

'declara e inicia conexão
Set conn = New ADODB.Connection
conn.ConnectionString = StringConexao
conn.CursorLocation = adUseClient
conn.Open

'executa instrucao
conn.Execute Data1.Recordset!Instrucao

'fecha conexão
conn.Close

Exit Sub
Trata_erro:
'-------------------------------------------------------------------------------------------------------------
'dados do erro
cloud_erro = Str$(Err.Number)
Cloud_erro_desc = Err.Description
Erro_Reenvio = True
MsgBox "Atenção : " & Cloud_erro_desc, vbInformation, "Erro: " & cloud_erro
Exit Sub

End Sub

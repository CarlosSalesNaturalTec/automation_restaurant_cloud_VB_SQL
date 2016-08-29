VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_reservas 
   BackColor       =   &H80000018&
   Caption         =   "Controle de Reservas"
   ClientHeight    =   8595
   ClientLeft      =   180
   ClientTop       =   510
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Bt_Sair 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   855
      Left            =   13440
      Picture         =   "frm_reservas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Fechar esta Janela"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton bt_desistir 
      Caption         =   "Desistir"
      Height          =   855
      Left            =   12120
      Picture         =   "frm_reservas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Desistir da última Inclusão/Alteração"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton bt_confirmar 
      Caption         =   "Confirmar"
      Height          =   855
      Left            =   11040
      Picture         =   "frm_reservas.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Confirmar Inclusões/Alterações"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Bt_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   7080
      Picture         =   "frm_reservas.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Novo"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Bt_Excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   9600
      Picture         =   "frm_reservas.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Excluir"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Bt_Alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   8280
      Picture         =   "frm_reservas.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   7
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
      TabIndex        =   12
      Top             =   240
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   12726
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483624
      TabCaption(0)   =   "Pesquisa"
      TabPicture(0)   =   "frm_reservas.frx":198C
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Ficha Detalhada"
      TabPicture(1)   =   "frm_reservas.frx":19A8
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_reservas.frx":19C4
         Height          =   5355
         Left            =   240
         OleObjectBlob   =   "frm_reservas.frx":19D4
         TabIndex        =   5
         Top             =   600
         Width           =   13815
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   5535
         Left            =   -74640
         TabIndex        =   13
         Top             =   480
         Width           =   13815
         Begin VB.TextBox cmb_hospede 
            DataField       =   "Hóspede"
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
            Height          =   450
            Left            =   240
            MaxLength       =   100
            TabIndex        =   0
            Top             =   480
            Width           =   10695
         End
         Begin VB.TextBox Text1 
            DataField       =   "obs"
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
            Height          =   2940
            Left            =   240
            MaxLength       =   20
            TabIndex        =   4
            Top             =   2280
            Width           =   13215
         End
         Begin MSMask.MaskEdBox txt_chegada 
            DataField       =   "chegada"
            DataSource      =   "Data1"
            Height          =   495
            Left            =   240
            TabIndex        =   2
            Top             =   1320
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   873
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
         Begin VB.TextBox txt_num 
            DataField       =   "Numero"
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
            Left            =   11880
            MaxLength       =   20
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   480
            Width           =   1575
         End
         Begin MSMask.MaskEdBox txt_saida 
            DataField       =   "Saída"
            DataSource      =   "Data1"
            Height          =   495
            Left            =   2760
            TabIndex        =   3
            Top             =   1320
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   873
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
            Caption         =   "Observações"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   18
            Top             =   2040
            Width           =   945
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Data da Saida:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   2760
            TabIndex        =   17
            Top             =   1080
            Width           =   1065
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Nome do Hóspede"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Número da Reserva:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   11880
            TabIndex        =   15
            Top             =   240
            Width           =   1470
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Data de Chegada:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   1080
            Width           =   1305
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
Attribute VB_Name = "frm_reservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Confirmando As Boolean
Dim Adicionando As Boolean
Dim Desistir As Boolean
Dim Carregando As Boolean
Dim Editar As Boolean

Private Sub Bt_Alterar_Click()

If Data1.Recordset.EOF Then Exit Sub

Bt_novo.Enabled = False
Bt_Excluir.Enabled = False
Bt_Alterar.Enabled = False
Bt_Sair.Enabled = False
    
Editar = True

Frame1.Enabled = True

SSTab1.Tab = 1
cmb_hospede.SetFocus

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

Private Sub Bt_Novo_Click()

On Error GoTo Trata_erro

Data1.Recordset.AddNew

'Abilita campos para edição e desabilita botões
Frame1.Enabled = True

Bt_novo.Enabled = False
Bt_Excluir.Enabled = False
Bt_Alterar.Enabled = False
Bt_Sair.Enabled = False

SSTab1.Tab = 1
cmb_hospede.SetFocus

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
            
    If cmb_hospede.Text = "" Then
        MsgBox "Informe o nome do Hóspede!"
        cmb_hospede.SetFocus
        Editar = True
        Save = False
        Action = vbDataActionCancel
        Exit Sub
    End If
    
    If Not Valida_D(txt_chegada, "Data de Chegada Inválida") Then txt_chegada.SetFocus: Editar = True: Exit Sub
    If Not Valida_D(txt_saida, "Data de Chegada Inválida") Then txt_saida.SetFocus: Editar = True: Exit Sub
    
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

Call Abrir_BD_Data(Data1, "Tbl_reservas", "[chegada]", "")
Carregando = True

End Sub

Private Sub DBGrid1_DblClick()
SSTab1.Tab = 1
End Sub

Private Sub txt_chegada_GotFocus()
Call Mask_Data(txt_chegada)
Call Selecionar(txt_chegada)
End Sub

Private Sub txt_saida_GotFocus()
Call Mask_Data(txt_saida)
Call Selecionar(txt_saida)
End Sub


Private Sub cmb_hospede_GotFocus()
Call Selecionar(cmb_hospede)
End Sub


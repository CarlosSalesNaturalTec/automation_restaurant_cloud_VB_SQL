VERSION 5.00
Begin VB.Form frm_login 
   Caption         =   "Food Control"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3720
   Icon            =   "frm_login_cad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txt_usuário 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   960
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Senha :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   555
      End
   End
   Begin VB.CommandButton Bt_Sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2280
      Picture         =   "frm_login_cad.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Bt_Ok 
      Caption         =   "&Entrar"
      Default         =   -1  'True
      Height          =   855
      Left            =   1080
      Picture         =   "frm_login_cad.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database
Dim Tab1 As Recordset
Dim Tab2 As Recordset   'dados da empresa
Dim Tab4 As Recordset   'parametros

Private Sub Bt_Ok_Click()

On Error GoTo Trata_erro
            
Caminho_Rede = App.Path
NumCaixa = 0

Set db1 = OpenDatabase(Caminho_Rede & "\sec.mdb", False, False, ";PWD=0305ca")
Set Tab1 = db1.OpenRecordset("select * from dados")

Tab1.FindFirst ("usuário  = '" & txt_usuário.Text & "'")
If Not Tab1.NoMatch Then

    If Text1.Text = Tab1!senha Then
        Nível = Tab1!Nível
        Usuário = "" & Tab1!Usuário
        Cod_Operador = Tab1!código
        
        'dados da empresa
        Set Tab2 = db1.OpenRecordset("select * from [tbl_empresa]")
        Empresa_Nome = "" & Tab2!Empresa
        Empresa_End = "" & Tab2!Endereço
        Empresa_CNPJ = "" & Tab2!cnpj
                
        'parametros - string de conexão
        Set Tab4 = db1.OpenRecordset("select * from [tbl_parametros]")
        If Tab4.EOF Then
            StringConexao = ""
            ModeloPRinter = ""
            MsgBox "Parâmetros incompletos", vbExclamation, "Atenção"
        Else
            StringConexao = RTrim(Tab4!StringConexao)
            ModeloPRinter = Tab4!ModeloPRinter
            Limite_Caixas = Tab4!Limite_Caixas
        End If
        
        'verifica se existe caixa aberto
        Dim Db2 As Database
        Dim Tab3 As Recordset
        Set Db2 = OpenDatabase(Caminho_Rede & "\dados.mdb")
        Set Tab3 = Db2.OpenRecordset("select * from [tbl_caixas] where [Fechado]=false")
        If Not Tab3.EOF Then NumCaixa = Tab3!código
        
        db1.Close
        Db2.Close
        
        frm_mnu.Show
        Unload Me
        Exit Sub
    End If
    
End If
    
continua = MsgBox("Usuário ou Senha Incorretos!", vbExclamation + vbRetryCancel, "Acesso ao Sistema")
If continua = 2 Then End

Text1.Text = ""
Text1.SetFocus

Exit Sub

Trata_erro:
MsgBox "Bloqueio de Segurança ! Tente Novamente. " & Str$(Err.Number) & "  /  Descrição : " & Err.Description
Exit Sub

End Sub

Private Sub Bt_Sair_Click()
Unload Me
End Sub

Private Sub Form_Load()

If App.PrevInstance Then
    MsgBox "O Sistema Já Está Aberto !", vbExclamation + vbOKOnly, "ATENÇÃO"
    End
End If

End Sub

Private Sub Text1_Change()
If Len(Text1.Text) = 6 Then Bt_Ok.SetFocus
End Sub

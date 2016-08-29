VERSION 5.00
Begin VB.Form frm_Usuários 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6735
   Begin VB.CommandButton Bt_Voltar 
      Caption         =   "< Voltar"
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton bt_avançar 
      Caption         =   "Avançar >"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton bt_confirmar 
      Caption         =   "&Confirmar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   3480
      Picture         =   "frm_usuarios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton bt_desistir 
      Caption         =   "&Desistir"
      Enabled         =   0   'False
      Height          =   855
      Left            =   4440
      Picture         =   "frm_usuarios.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   480
      TabIndex        =   9
      Top             =   240
      Width           =   5895
      Begin VB.TextBox txt_usuário 
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   2
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Senha :"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nível (1/2) :"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.CommandButton Bt_Sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5520
      Picture         =   "frm_usuarios.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Bt_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   480
      Picture         =   "frm_usuarios.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Bt_Excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   2400
      Picture         =   "frm_usuarios.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Bt_Alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1440
      Picture         =   "frm_usuarios.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
End
Attribute VB_Name = "frm_Usuários"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database
Dim usu_tab1 As Recordset

Dim Desistir As Boolean
Dim Carregando As Boolean
Private Sub Bt_Alterar_Click()

On Error GoTo Trata_erro

If usu_tab1.EOF Then Exit Sub

Bt_novo.Enabled = False
Bt_Excluir.Enabled = False
Bt_Alterar.Enabled = False
Bt_Sair.Enabled = False
bt_avançar.Enabled = False
Bt_Voltar.Enabled = False
    
bt_confirmar.Enabled = True
bt_desistir.Enabled = True

Frame1.Enabled = True
txt_usuário.SetFocus

usu_tab1.Edit

Exit Sub
Trata_erro:
MsgBox "Bloqueio de Segurança ! Tente Novamente. Cod.: " & Str$(Err.Number) & "  /  Descrição : " & Err.Description
Exit Sub


End Sub

Private Sub bt_avançar_Click()

If Not usu_tab1.EOF Then usu_tab1.MoveNext
If usu_tab1.EOF Then usu_tab1.MoveLast
Call Mostra_dados

End Sub

Private Sub bt_confirmar_Click()

On Error GoTo Trata_erro

If Not IsNumeric(Text2.text) Then
    MsgBox "Nível Incorreto"
    Text2.SetFocus
    Exit Sub
End If
    
usu_tab1!Nível = Text2.text
usu_tab1!senha = Text3.text
usu_tab1!Usuário = txt_usuário

usu_tab1.Update
            
'desabilita campos para edição e habilita botões
   
Bt_novo.Enabled = True
Bt_Excluir.Enabled = True
Bt_Alterar.Enabled = True
Bt_Sair.Enabled = True
bt_avançar.Enabled = True
Bt_Voltar.Enabled = True

bt_confirmar.Enabled = False
bt_desistir.Enabled = False
  
Frame1.Enabled = False
    
Exit Sub
Trata_erro:
MsgBox "Bloqueio de Segurança ! Tente Novamente. Cod.: " & Str$(Err.Number) & "  /  Descrição : " & Err.Description
Exit Sub

End Sub

Private Sub bt_desistir_Click()

If Not usu_tab1.EOF Then usu_tab1.MoveLast
Call Mostra_dados

Bt_novo.Enabled = True
Bt_Excluir.Enabled = True
Bt_Alterar.Enabled = True
Bt_Sair.Enabled = True
bt_avançar.Enabled = True
Bt_Voltar.Enabled = True

bt_confirmar.Enabled = False
bt_desistir.Enabled = False


End Sub

Private Sub Bt_Excluir_Click()

On Error GoTo Trata_erro

If usu_tab1.EOF Then Exit Sub
confirma = MsgBox("Confirma ?", 4, "Excluir Registro")
If confirma = 7 Then Exit Sub

usu_tab1.Delete
usu_tab1.MoveFirst
Call Mostra_dados

Exit Sub
Trata_erro:
MsgBox "Bloqueio de Segurança ! Tente Novamente. Cod.: " & Str$(Err.Number) & "  /  Descrição : " & Err.Description
Exit Sub

End Sub

Private Sub Bt_Novo_Click()

On Error GoTo Trata_erro

usu_tab1.AddNew
Text2.text = ""
Text3.text = ""
txt_usuário.text = ""

'Abilita campos para edição e desabilita botões
Frame1.Enabled = True

Bt_novo.Enabled = False
Bt_Excluir.Enabled = False
Bt_Alterar.Enabled = False
Bt_Sair.Enabled = False
bt_avançar.Enabled = False
Bt_Voltar.Enabled = False
  
bt_confirmar.Enabled = True
bt_desistir.Enabled = True

txt_usuário.SetFocus

Exit Sub
Trata_erro:
MsgBox "Bloqueio de Segurança ! Tente Novamente. Cod.: " & Str$(Err.Number) & "  /  Descrição : " & Err.Description
Exit Sub

End Sub

Private Sub Bt_Sair_Click()
Unload Me
End Sub

Private Sub Bt_Voltar_Click()
If Not usu_tab1.BOF Then usu_tab1.MovePrevious
If usu_tab1.BOF Then usu_tab1.MoveFirst
Call Mostra_dados
End Sub

Private Sub Form_Load()

Set db1 = OpenDatabase(Caminho_Rede & "\sec.mdb", False, False, ";PWD=0305ca")
Set usu_tab1 = db1.OpenRecordset("dados")

Call Mostra_dados

End Sub

Sub Mostra_dados()

If usu_tab1.EOF Or usu_tab1.BOF Then Exit Sub

Text2.text = "" & usu_tab1!Nível
Text3.text = "" & usu_tab1!senha
txt_usuário.text = "" & usu_tab1!Usuário

End Sub

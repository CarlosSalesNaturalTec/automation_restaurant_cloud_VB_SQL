Attribute VB_Name = "Gerais"
'**************** INDICE ****************
'=============================================================================================

' --- Formulário Menu
' SAIR          : Sair do sistema
' BACKUP        : Cópia de Segurança
' RESTORE       : Restaurar Cópia de Segurança
' COMPACTAR     : Compactar BD
' CONFIG_REDE   : Configurar Rede
' IMPRIME       : Imprime relatórios (somente a partir de frm_menu)
' CARREGAR_MDI  : Inicia Frm_Menu com definição do nome do arquivo de parâmeto de path

' --- Tratamento de Objetos
' SELECIONAR    : Seleção automática de campo / Ex : call selecionar(txt_valor)
' MASK_DATA     : Máscara de Data 99/99/99 (não esquecer o lostfocus : ojeto.mask = "")
' - Caixas
' CONF          : Caixa de Confirmação Sim/Não  / Ex : Conf("Confirma Valor", "Atenção")
' SEGDAT        : Solicitar Data e verificar consistência de valor digitado / Ex.: If Not Segdat("Data Inicial", "Contas a Pagar no Período", Date) Then Exit Sub

' --- Banco de Dados
' ABRIR_BD_DATA : Abre banco de dados em objeto Data (data1.databasename...)
' VALIDA_N      : Validação de campo do tipo Numérico em data1.validate
' VALIDA_D      : Validação de campo do tipo Data em data1.validate
' -- Registros
' EXCLUIR       : Excluir registro de objeto data
' LOC_LIKE      : Localizar registro por qualquer parte da string
' LOC_NUMBER    : Localizar registro numérico
' LOC_DATE      : Localizar registro tipo data

' --- Diversos
' VALIDA_CGC    : Valida CGC (obs. ver gotfocus e lostfocus no formulário)
' CALCULA_CGC   : Valida CGC (obs. ver gotfocus e lostfocus no formulário)

'==========================================================================================================

'----- Variáveis de trabalho

Global Rotina As String
Global Nível As Single
Global Usuário As String

Global Caminho_Rede As String
Global Arq_path As String
Global Segue_data As Date

Global Exibir_lembrete As Boolean

Public Sub Sair()

'---- SAIR DO SISTEMA

confirma = MsgBox("Confirma ?", vbQuestion + vbYesNo, "Sair do Sistema")
If confirma = 7 Then Exit Sub
End

End Sub

Public Sub Backup()

'BACKUP
Call compactar

disco = MsgBox("Pressione OK para Continuar", vbInformation + vbOKCancel, "Copiando")

'If disco <> 2 Then
'    Dim X As Variant
'    Y = Caminho_Rede & "\arj a c:\copias\copia C:\delivery_new\dados.mdb"
'    X = Shell(Y, 3)
'End If

FileCopy Caminho_Rede & "\dados.mdb", "C:\copias\copia.mdb"

MsgBox "Processo Concluído", vbInformation, "ok"

End Sub

Public Sub Restore()

'RESTORE
confirma = MsgBox("Confirma ?", vbQuestion + vbYesNo, "Restaurar Cópia de Segurança")
If confirma = 7 Then Exit Sub
    
Dim X As Variant
Y = Caminho_Rede & "\arj x -y -v1440 C:copia.arj"
X = Shell(Y, 3)

End Sub


'COMPACTAR
Public Sub compactar()

On Error GoTo okerror

Dim Db1 As String
Dim Db2 As String

Db1 = Caminho_Rede & "\dados.mdb"
Db2 = Caminho_Rede & "\dados2.mdb"

DBEngine.CompactDatabase Db1, Db2
Kill Db1
Name Db2 As Db1

Exit Sub

okerror:
Screen.MousePointer = 0
MsgBox "Saia do Sistema (inclusive nas estações) e Tente Novamente. " _
    & Err.Description, vbExclamation, "NÃO foi possível Compactar Banco de Dados"

End Sub

Public Sub Config_Rede(Arq_path)

'****** ROTINA DE ALTERAÇÃO DE CAMINHO DE REDE
Open Arq_path For Random As #1
Get #1, 1, ddrive
Caminho_Rede = ddrive
Close #1

caminho = InputBox("Informe Endereço Rede ", "Atual : " & Caminho_Rede)
If caminho = "" Then Exit Sub

Open Arq_path For Random As #1
Put #1, 1, caminho
Close #1

Caminho_Rede = caminho

MsgBox "Parâmetros Salvos com Sucesso !"

End Sub

'******** SELEÇÃO AUTOMÁTICA DE CAMPO

Public Sub Selecionar(objeto As Object)

objeto.SelStart = 0
objeto.SelLength = Len(objeto.Text)

End Sub

Public Function Conf(mensagem, aviso)

'confirmação (sim ou não)
Conf = MsgBox(mensagem, vbQuestion + vbYesNo, aviso)

End Function

Public Function Segdat(pergunta, titulo, valor_padrão) As Boolean

'solicitar data e verificar consistência de valor digitado

'exemplo de chamada da função :
'If Not Segdat("Data Inicial", "Contas a Pagar no Período", Date) Then Exit Sub

v1 = InputBox(pergunta, titulo, valor_padrão)
If v1 = "" Then
    Segdat = False
    Exit Function
End If

If Not IsDate(v1) Then
    MsgBox "Data Inválida", vbExclamation, "Atenção"
    Segdat = False
    Exit Function
End If

Segdat = True
Segue_data = CDate(v1)

End Function


Public Sub Imprime(nome_rel, periodo1, periodo2, campo)

'=================== IMPRIMIR RELATÓRIOS (a partir do menu)

ano1 = Year(periodo1)
mes1 = Month(periodo1)
dia1 = Day(periodo1)

ano2 = Year(periodo2)
mes2 = Month(periodo2)
dia2 = Day(periodo2)

With frm_menu.CrystalReport1
    .ReportFileName = Caminho_Rede & "\" & nome_rel
    If campo = "" Then
        .SelectionFormula = ""
    Else
        .SelectionFormula = campo & " in Date (" & ano1 & ", " & mes1 & ", " & dia1 & ") to Date (" & ano2 & "," & mes2 & "," & dia2 & ")"
    End If
    .Formulas(0) = "periodo = 'Período : " & periodo1 & " a " & periodo2 & "'"
    .Action = 1
End With

End Sub

Public Sub Abrir_BD_Data(objeto As Object, Tabela, Campo_Order, Filtro)

On Error GoTo Trata_erro

objeto.DatabaseName = Caminho_Rede & "\dados.mdb"
If Filtro = "" Then
    If Campo_Order = "" Then
        objeto.RecordSource = "Select * from [" & Tabela & "]"
    Else
        objeto.RecordSource = "Select * from [" & Tabela & "] order by " & Campo_Order
    End If
Else
    If Campo_Order = "" Then
        objeto.RecordSource = "Select * from [" & Tabela & "] where " & Filtro
    Else
        objeto.RecordSource = "Select * from [" & Tabela & "] where " & Filtro & " order by " & Campo_Order
    End If
End If

Exit Sub
Trata_erro:
MsgBox "Bloqueio de Segurança ! Tente Novamente. Cod.: " & Str$(Err.Number) & "  /  Descrição : " & Err.Description
Exit Sub

End Sub

Public Sub Carregar_Mdi(Arq_path)

Open Arq_path For Random As #1
Get #1, 1, ddrive
Caminho_Rede = ddrive
Close #1

End Sub

Public Sub Mask_Data(objeto As Object)

If IsDate(objeto.Text) Then
    anterior = CDate(objeto.Text)
    objeto.Mask = "99/99/99"
    objeto.Text = Format(anterior, "dd/mm/yy")
Else
    objeto.Mask = "99/99/99"
End If

End Sub

Public Sub Mask_CEP(objeto As Object)

If objeto.Text <> "" Then
    anterior = objeto.Text
    objeto.Mask = "99.999-999"
    objeto.Text = anterior
Else
    objeto.Mask = "99.999-999"
End If

End Sub


Public Function Valida_N(Objeto1 As Object, mensagem) As Boolean

Valida_N = True

If Not IsNumeric(Objeto1) Then
    MsgBox mensagem, vbExclamation, "Atenção"
    Save = False
    Action = vbDataActionCancel
    Valida_N = False
End If

End Function

Public Function Valida_D(Objeto2 As Object, mensagem) As Boolean

Valida_D = True

If Not IsDate(Objeto2) Then
    MsgBox mensagem, vbExclamation, "Atenção"
    Save = False
    Action = vbDataActionCancel
    Valida_D = False
End If

End Function

Public Function CalculaCGC(Numero As String) As String

Dim I As Integer
Dim prod As Integer
Dim mult As Integer
Dim digito As Integer

If Not IsNumeric(Numero) Then
   CalculaCGC = ""
   Exit Function
End If

mult = 2
For I = Len(Numero) To 1 Step -1
  prod = prod + Val(Mid(Numero, I, 1)) * mult
  mult = IIf(mult = 9, 2, mult + 1)
Next

digito = 11 - Int(prod Mod 11)
digito = IIf(digito = 10 Or digito = 11, 0, digito)

CalculaCGC = Trim(Str(digito))

End Function

Public Function ValidaCGC(CGC As String) As Boolean

If CalculaCGC(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
   ValidaCGC = False
   Exit Function
End If

If CalculaCGC(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
   ValidaCGC = False
   Exit Function
End If

ValidaCGC = True

End Function

Public Sub Excluir(objeto As Object)

On Error GoTo Trata_erro

If objeto.Recordset.EOF Then Exit Sub
confirma = MsgBox("Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, "Excluir Registro")
If confirma = 7 Then Exit Sub

objeto.Recordset.Delete
objeto.Recordset.MoveLast

Exit Sub
Trata_erro:
MsgBox "Bloqueio de Segurança ! Tente Novamente. Cod.: " & Str$(Err.Number) & "  /  Descrição : " & Err.Description
Exit Sub

End Sub

Public Sub Loc_like(objeto As Object, pergunta, campo)

localiz = InputBox(pergunta, "Localizar")
If localiz = "" Then Exit Sub

objeto.Recordset.FindFirst (campo & " like '*" & localiz & "*'")
If objeto.Recordset.NoMatch Then MsgBox "Não Localizado", vbInformation, "Atenção"

End Sub

Public Sub Loc_number(objeto As Object, pergunta, campo)

localiz = InputBox(pergunta, "Localizar")
If localiz = "" Then Exit Sub
If Not IsNumeric(localiz) Then Exit Sub

objeto.Recordset.FindFirst (campo & " = " & localiz)
If objeto.Recordset.NoMatch Then MsgBox "Não Localizado", vbInformation, "Atenção"

End Sub

Public Sub Loc_Date(objeto As Object, pergunta, campo)

localiz = InputBox(pergunta, "Localizar")
If localiz = "" Then Exit Sub
If Not IsDate(localiz) Then MsgBox "Data Incorreta", vbInformation, "Atenção": Exit Sub

objeto.Recordset.FindFirst (campo & " = #" & Format(CDate(localiz), "mm/dd/yy") & "#")
If objeto.Recordset.NoMatch Then MsgBox "Não Localizado", vbInformation, "Atenção"

End Sub



Function Num_Mes(nome_mes As String)

Select Case nome_mes
    Case "janeiro"
        Num_Mes = 1
    Case "fevereiro"
        Num_Mes = 2
    Case "março"
        Num_Mes = 3
    Case "abril"
        Num_Mes = 4
    Case "maio"
        Num_Mes = 5
    Case "junho"
        Num_Mes = 6
    Case "julho"
        Num_Mes = 7
    Case "agosto"
        Num_Mes = 8
    Case "setembro"
        Num_Mes = 9
    Case "outubro"
        Num_Mes = 10
    Case "novembro"
        Num_Mes = 11
    Case "dezembro"
        Num_Mes = 12
End Select
        
End Function

Sub Log_erros(Num_erro As String, Desc_erro As String)

'****** LOG DE ERROS
nome_arquivo = Caminho_Rede & "\log_erros\log_erro_dlvry_" & Format(Date, "dd_mm") & "_" & Format(Time, "hh_mm_ss") & "_" & Usuário & ".txt"
Open nome_arquivo For Output As #2
Write #2, Num_erro, Desc_erro
Close #2

End Sub

Attribute VB_Name = "DarumaFramework_VB6"
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public iRetorno As Integer
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'================================DECLARACOES DARUMA FRAMEWORK ================================'
'===========                           IMPRESSORAS DUAL                            ==========='

Public Declare Function iEnviarBMP_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stArqOrigem As String) As Integer
Public Declare Function iAcionarGaveta_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" () As Integer
Public Declare Function iImprimirArquivo_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stPath As String) As Integer
Public Declare Function eBuscarPortaVelocidade_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" () As Integer
Public Declare Function rStatusGuilhotina_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" () As Integer
Public Declare Function iConfigurarGuilhotina_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal iHabilitar As Integer, ByVal iQtdeLinha As Integer) As Integer
Public Declare Function eGerarQrCodeArquivo_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stPath As String, ByVal stDados As String) As Integer
Public Declare Function iImprimirBMP_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stPath As String) As Integer
Public Declare Function rStatusGaveta_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByRef iStatusGaveta As Integer) As Integer
Public Declare Function rStatusDocumento_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" () As Integer
Public Declare Function rStatusImpressora_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" () As Integer
Public Declare Function regVelocidade_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stParametro As String) As Integer
Public Declare Function regTermica_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stParametro As String) As Integer
Public Declare Function regTabulacao_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stParametro As String) As Integer
Public Declare Function regPortaComunicacao_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stParametro As String) As Integer
Public Declare Function regModoGaveta_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stParametro As String) As Integer
Public Declare Function regLinhasGuilhotina_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stParametro As String) As Integer
Public Declare Function regEnterFinal_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stParametro As String) As Integer
Public Declare Function regAguardarProcesso_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stParametro As String) As Integer
Public Declare Function iImprimirTexto_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stTexto As String, ByVal iTam As Integer) As Integer
Public Declare Function iAutenticarDocumento_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stTexto As String, ByVal stLocal As String, ByVal stTimeOut As String) As Integer
Public Declare Function regCodePageAutomatico_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stParametro As String) As Integer
Public Declare Function regZeroCortado_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stParametro As String) As Integer
Public Declare Function rConsultaStatusImpressora_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal stIndice As String, ByVal stTipo As String, ByVal stRetorno As String) As Integer


'================================DECLARACOES DARUMA FRAMEWORK ================================'
'===========                               GENERICO                                ==========='

Public Declare Function eAbrirSerial_Daruma Lib "DarumaFrameWork.dll" (ByVal pszPorta As String, ByVal pszVelocidade As String) As Integer
Public Declare Function eFecharSerial_Daruma Lib "DarumaFrameWork.dll" () As Integer
Public Declare Function tEnviarDados_Daruma Lib "DarumaFrameWork.dll" (ByVal pszBytes As String, ByVal iTamBytes As Integer) As Integer
Public Declare Function rReceberDados_Daruma Lib "DarumaFrameWork.dll" (ByVal pszBufferEntrada As String) As Integer

'================================DECLARACOES DARUMA FRAMEWORK ================================'
'===========                          DARUMAFRAMEWORK                              ==========='

Public Declare Function eVerificarVersaoDLL_Daruma Lib "DarumaFrameWork.dll" (ByVal sVersaoDLL As String) As Integer
Public Declare Function eDefinirProduto_Daruma Lib "DarumaFrameWork.dll" (ByVal sProduto As String) As Integer
Public Declare Function eBuscarPortaVelocidade_ECF_Daruma Lib "DarumaFrameWork.dll" () As Integer
Public Declare Function eAcionarGuilhotina_ECF_Daruma Lib "DarumaFrameWork.dll" (ByVal sTipoCorte As String) As Integer
Public Declare Function eAbrirGaveta_ECF_Daruma Lib "DarumaFrameWork.dll" () As Integer
Public Declare Function eInterpretarRetorno_ECF_Daruma Lib "DarumaFrameWork.dll" (ByVal iErro As Integer, ByVal sMsg_Erro As String) As Integer
Public Declare Function eInterpretarErro_ECF_Daruma Lib "DarumaFrameWork.dll" (ByVal iErro As Integer, ByVal sMsg_Erro As String) As Integer
Public Declare Function eInterpretarAviso_ECF_Daruma Lib "DarumaFrameWork.dll" (ByVal iAviso As Integer, ByVal sMsg_Aviso As String) As Integer

'================================    FUNÇÕES GLOBAIS    ================================'
'===========            TRATAMENTO DE RETORNO IMPRESSORA FISCAL              ==========='

Public Function DarumaFramework_Mostrar_Retorno_ECF(iRetorno As Integer)
        
        Dim Str_Msg_NumRetorno As String
        Dim Str_Msg_NumErro As String
        Dim Str_Msg_NumAviso As String
        Dim Int_NumRetorno As Integer
        Dim Int_NumErro As Integer
        Dim Int_NumAviso As Integer
        
        
            Str_Msg_NumRetorno = Space(200)
            Str_Msg_NumErro = Space(200)
            Str_Msg_NumAviso = Space(200)
            
            Int_NumRetorno = 0
            Int_NumErro = 0
            Int_NumAviso = 0
                                
           FR_MostraAvisoErro.lblRetorno.Caption = "Retorno do Método:  " + Str_Msg_NumRetorno
           FR_MostraAvisoErro.lblErro.Caption = "Mensagem de Erro:  " + Str_Msg_NumErro
           FR_MostraAvisoErro.lblAviso.Caption = "Mensagem de Aviso:  " + Str_Msg_NumAviso
           FR_MostraAvisoErro.Show (1)
       
End Function

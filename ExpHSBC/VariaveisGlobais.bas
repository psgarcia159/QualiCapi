Attribute VB_Name = "VariaveisGlobais"
Option Explicit
'Padrões
Global PCodEmpresa                  As String   'Código da empresa
Global PEmpresa                     As String   'Nome da Empresa
Global PFantasia                    As String   'Nome Fantasia
Global PEndereco                    As String   'Endereço
Global PCidade                      As String   'Cidade
Global XGI_CONTACORRENTE            As Integer  'Código da Conta Corrente da Empresa
Global XGT_CGC                      As String   'CGC da empresa
Global XGT_CEP                      As String   'CEP da empresa
Global XGT_NUMPORTA                 As String   'Nº porta
Global XGT_BAIRRO                   As String   'Bairro da empresa
Global XGT_ESTADO                   As String   'Estado da empresa
Global XGT_TIPO                     As String   'Importação/Exportação do Itaú
Global XGT_TIPOGRID                 As String   'Indica o tipo do grid (B-Bound/U-UnBound)
Global XGT_CAMPOFILTRO              As String   'Nome do campo que será filtrado
Global XGB_PQUALIFIN                As Boolean  'VERIFICA SE O SISTEMA QUALIFIN ESTA INTEGRADO
Global XGB_PQUALICONT               As Boolean  'VERIFICA SE O SISTEMA QUALICONT ESTA INTEGRADO
Global XGT_EXPCOMPROMCAPI           As String   'Indica se os compromissos serão exp. p/ Qualifin

'GLOBAIS
Global pTipoCopia                   As String   'Indica o tipo da cópia (Mono=Mono-Empresa e Multi=Multi-Empresa)
Global XGT_ARQUIVO                  As String   'Nome do arquivo para exportação
Global XGT_LOCALARQ                 As String * 254 'Caminho para gravação do arquivo
Global func_cd_operador             As String  ' Código do operador do sistema

Global XGT_NomeAplic                'Nome da Aplicação
Global XGT_IdAplic                  'Identificação da Aplicação
Global XGT_DirPadrao                'Diretorio Padrao da Segurança do Sistema
Global XGT_DirSistema               'Diretorio Padrao do Sistema
Global XGI_NSerie                   'Numero de Serie do Aplicativo
Global XGT_SQL As String            'STRING PARA CONTRUÇÃO DE QUERIES
Global Chave As Variant             'armazena chave

Global XInserir As String        'Flag para identificar se entrou na rotina para inserir ou alterar ('I'Inserir, 'A'Alterar, 'C'Consultar)
Global XOrigem As String         'Identifica a origem (que formulario) chamou o formulário ativo

'Variáveis para ajustar a tela
Global ResolucaoX As Long
Global Resolucaoy As Long

'VARIAVEIS DE RELATORIO
Global XGR_APLICACAO    As New CRAXDRT.Application
Global XGR_RELATORIO    As New CRAXDRT.Report
Global XGO_QUERYREL     As Object 'QUERY DO RELATORIO
Global XGO_QUERYCAB     As Object 'QUERY DO CABECALHO
Global XGT_SELECAO      As String
Global XGT_CONJUNCAO    As String
Global XGO_RSRELATORIO  As New ADODB.Recordset
Global XGO_RSCABECALHO  As New ADODB.Recordset

Global XControle As String 'Colocada como global para poder limpar a tabela ProcPagto_Temp no Un load do Form que é global

'/* Descricao do Registro que contem as Informacoes da Protecao */

Type T_PROT_INFO
     CodRet         As Long
     CodDiag        As Long
     CodDebug       As Long
     NroInstRest    As Long
     NroMaxInst     As Long
     NroExecsRest   As Long
     NroMaxExecs    As Long
     NroDiasRest    As Long
     NroMaxDias     As Long
     NroExecsEmRede As Long
     NroResetsAtual As Long
     NroResetsInv   As Long
     InstEmergRest  As Long
     MaxInstEmerg   As Long
     DiasEmerg      As Long
     CampoInt       As Long
     CampoIntExtra  As Long
     Aplicacao      As String * 9
     MsgRet         As String * 51
     Usuario        As String * 71
     Empresa        As String * 71
     NroSerie       As String * 15
     DirInst        As String * 171
     VersaoApl      As String * 9
     DiaInstalacao  As String * 11
     DiaLimite      As String * 11
     DiaUltAcesso   As String * 11
     CampoStr       As String * 201
     Protegida      As String * 1
     Personalizacao As String * 1
     Personalizar   As String * 1
     TipoProt       As String * 1
     CDInstOpt      As String * 1
     RegistrarData  As String * 1
     CpuLock        As String * 1
     NtCompativel   As String * 1
End Type


Declare Function hl_leprot_vb Lib "HP32V301.DLL" (ByVal nome_apl As String, ByVal senha_id As String, ByVal dir_inst As String, ByVal atualiza_limites As Long, ByVal verifica_uso_em_rede As Long, info As Any) As Long
Declare Function hl_move_vb Lib "HP32V301.DLL" (ByVal Source As String, ByVal dest As String, ByVal sit_espec As Long, ByVal empr As String, ByVal usr As String) As Long

Global ProtResult As T_PROT_INFO  'Vetor com informações da proteção


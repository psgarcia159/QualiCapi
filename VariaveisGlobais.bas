Attribute VB_Name = "VariaveisGlobais"
Option Explicit

Declare Function hl_leprot_vb Lib "HP32V301.DLL" (ByVal nome_apl As String, ByVal senha_id As String, ByVal dir_inst As String, ByVal atualiza_limites As Long, ByVal verifica_uso_em_rede As Long, info As Any) As Long
Declare Function hl_move_vb Lib "HP32V301.DLL" (ByVal Source As String, ByVal dest As String, ByVal sit_espec As Long, ByVal empr As String, ByVal usr As String) As Long


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
Global ProtResult As T_PROT_INFO  'Vetor com informações da proteção

Type Rect
  Topo As Single
  Esquerda As Single
  Altura As Single
  Largura As Single
End Type

Private Type DptoCompras
  Responsavel As String
  Telefone As String
  Fax As String
  Email As String
End Type
Global XGO_DPTOCOMPRAS As DptoCompras 'Informações sobre o dpto de compras

'Padrões
Global PCodEmpresa                  As Integer  'Código da empresa
Global PControleOrcamentoObra       As Boolean  'Indica se a empresa utiliza o módulo Orçamento Obra
Global PEmpresa                     As String   'Nome da Empresa
Global PCliente                     As String   'Nome do Cliente
Global PFantasia                    As String   'Nome Fantasia
Global PEndereco                    As String   'Endereço
Global PCidade                      As String   'Cidade
Global PLogo                        As String   'Logotipo
Global XGI_CONTACORRENTE            As Integer  'Código da Conta Corrente da Empresa
Global XGT_TIPOEXPORTACAO           As String   'Tipo de Exportação para Contabilidade (PC-ProvisaoCaixa/CX-Caixa)
Global XGI_QTDDIASPROVISAO          As Integer  'QTd de dias definido pela empresa para realizar a provisão
Global XGT_EXPCOMPROMCAPI           As String   'Indica se os compromissos serão exp. p/ Qualifin
Global XGT_IMPRIMEPP                As String   'Indica se utiliza o pp do Qualifin
Global XGT_ORDEMDESCONTO            As String   'Indica se o combo de desconto será oredenado por descrição ou cod. reduzido
Global XGT_CGC                      As String   'CGC da empresa
Global XGT_CEP                      As String   'CEP da empresa
Global XGT_NUMPORTA                 As String   'Nº porta
Global XGT_BAIRRO                   As String   'Bairro da empresa
Global XGT_ESTADO                   As String   'Estado da empresa
Global XGT_CAMPOFILTRO              As String   'Nome do campo que será filtrado
Global XGT_INCLUIRACRESCIMOS        As String   'Incluir acrescimos no valores indexado nos relatórios do Capi
Global XGT_MODELORECIBO             As String   'Deine o tipo da pro-rata
Global XGT_MSKCODREDUZIDO           As String * 254  'máscara do código reduzido
Global XGT_MSKCATEGORIA             As String * 254
Global XGB_CNPJOBRIGATORIO          As Boolean
Global XGT_CONTAGLOSA               As String * 6 'Código reduzido da conta de Glosa a Recuperar
Global XGT_CONTAGLOSASUS            As String * 6 'Código reduzido da conta de Glosa para o cliente SUS
Global XGI_QTDDIASFILTRO            As Integer  'Indica a qtd de dias para o filtro nas telas do qualifin
Global XGB_INCLUICANCELADOS         As Boolean  'Indica se os registros cancelados serão apresentados
Global XGB_EMAIL                    As Boolean  'Indica se a empresa usará a opção de enviar e-mail
Global XGB_AUTENTICAEMAIL           As Boolean  'Indica se o provedor de e-mail vai autenticar
Global XGT_CNAE                     As String   'Indica qual a tela do Cnae
Global XGB_APROPRIACRATIVOPASS      As Boolean  'Indica se será atribuido CR aos lançamentos com conta de ativo/passivo
Global XGT_NUMDIAS                  As String * 254 ' Quantidade de dias a ser utilizado no cálculo do número de registros a serem vizualizados nas telas de Duplicatas e Contas a Pagar\Receber
Global XGB_EXPMOVPATCONT            As Boolean  'Indica se os movimentos realizados no QualiPat deverão ser exportados para a contabilidade

Global PRaizBanco                   As String   'Diretorio Raiz do banco
Global pPCRmascara                  As String   'Mascara do CR
Global pPCRnivel                    As Integer  'Nivel do CR
Global pPlanoCmascara               As String   'Mascara do Plano
Global pPlanoCNivel                 As Integer  'Nivel do Plano de Contas
Global pCategoriaMascara            As String   'Mascara da Categoria
Global pCrAtivoPassivo                As String   'Apropria CR a contas de Ativo e Passivo
Global pCategoriaNivel              As Integer  'Nivel da Categoria
Global pVinculaCrCc                 As Boolean  'Informa se a empresa habilitou ou não a função de vinculação CR x CC
Global XGI_PCATEGORIANVFINANCEIRO   As Integer  'NIVEL PARA CONTA DAS CATEGORIAS
Global pTipoCR                      As String   'Tipo do CR
Global PadCaminho                   As String   'Armazena o caminho do banco
Global pTipoCopia                   As String   'Indica o tipo da cópia (Mono=Mono-Empresa e Multi=Multi-Empresa)
Global XGB_PQUALIFIN                As Boolean  'VERIFICA SE O SISTEMA QUALIFIN ESTA INTEGRADO
Global XGB_PQUALIEST                As Boolean  'VERIFICA SE O SISTEMA QUALIEST ESTA INTEGRADO
Global XGB_PCOMPRAS                 As Boolean  'VERIFICA SE O SISTEMA COMPRA ESTA INTEGRADO
Global XGB_PCAPI                    As Boolean  'VERIFICA SE O SISTEMA CAPI ESTA INTEGRADO
Global XGB_PQUALICONT               As Boolean  'VERIFICA SE O SISTEMA QUALICONT ESTA INTEGRADO
Global XGB_PQUALIFATS               As Boolean  'VERIFICA SE O SISTEMA QUALIFATS ESTA INTEGRADO
'Global XGB_PQUALIPAT                As Boolean  'VERIFICA SE O SISTEMA QUALIPAT ESTA INTEGRADO
Global XGT_ARQUIVO                  As String   'Nome do arquivo para exportação
Global XGT_LOCALARQ                 As String   'Caminho para gravação do arquivo

Global XGT_ALMOXARIFADOCENTRAL      As String   'Define o almoxarifado central para a empresa que se está trabalhando
Global XGT_ALMOXARIFADOUSUARIO      As String   'Define o almoxarifado do usuário para a empresa que se está trabalhando
Global XLT_CONTADOACAO              As String   'Armazena a conta de Doação de Receita
Global XGV_MATERIAIS                As Variant  'Armazena código do material

'globais
Global VGV_Chave As New XArray      'Vetor de chave para o caso de ter mais de uma chave
Global Chave As Variant             'armazena chave
Global xChave As Variant            'Armazena o valor do bookmark
Global Sql As String                'Armazena a string Sql da query
Global XGT_NomeAplic                'Nome da Aplicação
Global XGT_IdAplic                  'Identificação da Aplicação
Global XGT_DirPadrao                'Diretorio Padrao da Segurança do Sistema
Global XGT_DirSistema               'Diretorio Padrao do Sistema
Global XGI_NSerie                   'Numero de Serie do Aplicativo
Global XGT_SQL As String            'STRING PARA CONTRUÇÃO DE QUERIES
Global XGI_INDICE As Integer        'VALOR DE INDICE
Global XGT_CODIGO As String

'Dados do Usuário
Global XGT_USUARIO As String  'Nome do funcionário
Global XGL_CODPCR As Long     'Código do CR do funcionário
Global XGT_PCR As String      'Nome do CR do funcionário
Global func_cd_operador As Integer ' Código do operador do sistema
Global XGT_CARGO As String    'Nome do Cargo do funcionário

'Constantes referentes a rotina
'no travamento dos registros
Global Const NumberTravaRegistro = 40002
Global Const DescriptionTravaRegistro = "S1T00: [Microsoft][ODBC SQL Server Driver]Timeout expired"
Global Const NumberRegistroExcluido = 40009
Global Const DescriptionRegistroExcluido = "No current row."

'Variáveis para ajustar a tela
Global ResolucaoX As Long
Global Resolucaoy As Long

Global XInserir As String        'Flag para identificar se entrou na rotina para inserir ou alterar ('I'Inserir, 'A'Alterar, 'C'Consultar)
Global XCopiar As String
Global XOrigem As String         'Identifica a origem (que formulario) chamou o formulário ativo
Global VetorJanelas() As Variant
                                       
'Variaveis globais de chamada de formularios
Global XGV_FornNotaFiscal As Variant ' variavel que armazena o codigo do fornecedor
Global XGV_CodContabil As Variant

'VARIAVEIS DE RELATORIO
Global XGR_APLICACAO    As New CRAXDRT.Application
Global XGR_RELATORIO    As New CRAXDRT.Report
Global XGO_QUERYREL     As Object 'QUERY DO RELATORIO
Global XGO_QUERYCAB     As Object 'QUERY DO CABECALHO
Global XGT_SELECAO      As String
Global XGT_CONJUNCAO    As String
Global XGO_RSRELATORIO  As New adodb.Recordset
Global XGO_RSCABECALHO  As New adodb.Recordset

Global XControle As String 'Colocada como global para poder limpar a tabela ProcPagto_Temp no Un load do Form que é global
Global XGT_TIPOEXPFOLHA As String 'Indica se a exportação será por CR ou Geral

'Início - Larissa Moura 29/05/08
'Variáveis utilizadas na geração de log
 Global XGB_REGISTRARLOG As Boolean 'Define se é para registrar no log a operação que acabou de ser executada.
 Global XGM_MATRIZLOG() As Variant 'Guarda os nomes dos campos possíveis de serem alterados (do registro), os valores (que constam no banco) dos campos do registro possíveis de serem alterados e os valores (que constam no banco) dos campos do registro possíveis de serem alterados.
'Fim - Larissa Moura 29/05/08


Global Data_VinculoCRCC As String 'Data de corte para execução da rotina de vinculação CR x CC - Larissa, 30/07/08
Global XGB_ContinaPagamento As Boolean 'Usada para impedir a efetivação de títulos cujo desconto não possui conta contábil cadastrada - Larissa, 14/08/08
Global XGB_EntrouDupl As Boolean 'Usada para verificar se prssionou o botão alterar na tela de duplicatas
Global XGB_AlteraTransfCusto As Boolean 'Usada para verificar se está alterando uma transferência com Custo
Global XGB_AlertarSaldoAdiantamento As Boolean 'Usada para definir alerta de saldo do adiantamento em contas a pagar







Global XGB_CONTABILIZA As Boolean



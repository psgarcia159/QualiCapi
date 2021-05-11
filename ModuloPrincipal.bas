Attribute VB_Name = "ModuloPrincipal"
Option Explicit
Global FornClie As String         'Variável que guarda a classe C:Cliente ou F:Fornecedor
Global Pessoa As String * 1       'Variável que Guarda o tipo de Pessoa: F:Fisica,J:Juridica,3:Jurídica Produtos,4:Jurídica Produtos e Serviços
Global XFormForn As String        'indica qual formulário chamou o formulário de fornecedores
Global xLocalDescricao As String  'Variavel que definirá o label de Localizar
Global xLocalCampo As String      'Variável que definirá o campo que será mostrado na lista do combo de Localizar
Global XFormulario As String      'Indica a partir de que formulario a tela foi chamada
Global Finaliza As Boolean        'Variavel que determina se a aplicação será ou não finalizada

Public Sub Main()
    'Esta rotina inicia o sistema setando suas principais
    
    'variaveis globais e fazendo a chamada para o formulário principal do sistem
    Dim ResPadrao As ADODB.Recordset
    Dim XLT_DirPadrao As String * 254
    Dim XLT_CATEGORIA As String * 254
    Dim XLT_MSGERRO As String
    
    On Error GoTo RotuloErro
    
    'Seleciona o Diretorio Padrao do Sistema
    XGT_DirSistema = App.path
    
    'seleciona o diretorio padrão de segurança
    If GetPrivateProfileString("BancoDeDados", "DirPadrao", "", XLT_DirPadrao, 255, XGT_DirSistema + "\QualiAdmFin.INI") = 0 Then
        MsgBox "Problemas na Leitura do Arquivo de configuração", vbCritical, "Configuração"
        End
    End If
    XGT_DirPadrao = Trim(XLT_DirPadrao)
    
    'Indica o tipo de bloqueio a ser usado no sistema
    'Esta opção só é utilizada para o RDO
    TipoBloqueio = adLockPessimistic 'Lock Pessimista
    'TipoBloqueio = adLockReadOnly 'Lock Leitura
    'TipoBloqueio = adLockOptimistic 'Lock Otimista
    'TipoBloqueio = adLockBatchOptimistic 'Lock Otimista Baseadas no Valor de Linhas
          
    'Estabelece a conexão com o banco de dados selecionado
    While Not funAbreConexao(XLT_MSGERRO)
      MsgBox "Não foi possível se conectar ao banco de dados. Será iniciada a rotina de " & _
        "configuraçao para que você possa indicar os dados corretos para a conexão com o banco." & _
        Chr(13) & Chr(13) & "DESCRIÇÃO DO ERRO: " & XLT_MSGERRO, _
        vbCritical, "Conexão com o Banco de Dados"
      FrmConfBanco.Show 1
      If XGB_RESPOSTA = False Then End
    Wend
    
    'Estabelece a conexão com o banco de dados selecionado para a emissao dos Relatórios
    If Not funAbreConexaoRelatorio() Then
      MsgBox "Houve problemas com a conexão com o Banco de Dados."
      End 'Finaliza o programa
    End If
    
    GetPrivateProfileString "BancoDeDados", "Mascara", " ", XLT_CATEGORIA, 255, XGT_DirSistema + "\QualiAdmFin.INI"
    pCategoriaMascara = FunStrArqIni(XLT_CATEGORIA)
    XLT_CATEGORIA = ""
    GetPrivateProfileString "BancoDeDados", "Nivel", " ", XLT_CATEGORIA, 255, XGT_DirSistema + "\QualiAdmFin.INI"
    pCategoriaNivel = FunStrArqIni(XLT_CATEGORIA)
    
    SubQOpenRecordset ResPadrao, "SELECT * from Empresas where empr_tx_atual = 'S'", Estatico
  
    'INSERIDO EM 08/09/09 - LARISSA
    If Not (ResPadrao.EOF Or ResPadrao.BOF) Then
        PCodEmpresa = ResPadrao!empr_cd_empresa
        PFantasia = FunNulo(ResPadrao!empr_tx_fantasia)
        PEmpresa = FunNulo(ResPadrao!empr_tx_razaosocial)
        PCliente = FunNulo(ResPadrao!empr_tx_cliente)
        PEndereco = FunNulo(ResPadrao!empr_tx_endereco)
        PCidade = FunNulo(ResPadrao!empr_tx_cidade)
        pPCRmascara = FunNulo(ResPadrao!empr_tx_pcrmascara)
'        If IsNull(ResPadrao!empr_bl_ControleOrcamentoObra) Then
'            PControleOrcamentoObra = False
'        Else
'            PControleOrcamentoObra = ResPadrao!empr_bl_ControleOrcamentoObra
'        End If
        
        'Inserido em 31/07/08 - Larissa
        If FunNulo(ResPadrao!empr_tx_VinculaCrCc) = "S" Then
          pVinculaCrCc = True
        Else
          pVinculaCrCc = False
        End If
        
        XGI_QTDDIASFILTRO = FunNuloVal(ResPadrao!empr_nr_QtdDiasFiltro)
        
        If ResPadrao!empr_tx_EnviarEmail = "S" Then
          XGB_EMAIL = True
        Else
          XGB_EMAIL = False
        End If
          
        If ResPadrao!empr_tx_IncluiCancelados = "S" Then
          XGB_INCLUICANCELADOS = True
        Else
          XGB_INCLUICANCELADOS = False
        End If
        If ResPadrao!empr_tx_CrAtivoPassivo = "S" Then
          XGB_APROPRIACRATIVOPASS = True
        Else
          XGB_APROPRIACRATIVOPASS = False
        End If
        XGI_CONTACORRENTE = FunNuloVal(ResPadrao!Empr_cd_codcontacorrente)
        XGI_QTDDIASPROVISAO = FunNuloVal(ResPadrao!Empr_nr_qtdDiasProvisao)
        XGT_TIPOEXPORTACAO = FunNulo(ResPadrao!Empr_tx_ExpContabil)
        XGT_EXPCOMPROMCAPI = FunNulo(ResPadrao!empr_tx_ExpCompromissoCapi)
        XGT_INCLUIRACRESCIMOS = FunNulo(ResPadrao!empr_tx_IncluiAcrescimos)
        XGT_MODELORECIBO = FunNulo(ResPadrao!empr_tx_ModeloRecibo)
        XGB_CNPJOBRIGATORIO = FunNuloVal(ResPadrao!empr_tx_CnpjObrigatorio)
        XGT_IMPRIMEPP = FunNulo(ResPadrao!empr_tx_ImprimePP)
        XGT_ORDEMDESCONTO = FunNulo(ResPadrao!empr_tx_OrdemDescontos)
        XGT_CGC = FunNuloVal(ResPadrao!empr_tx_CGC)
        XGT_CEP = FunNuloVal(ResPadrao!empr_tx_Cep)
        XGT_NUMPORTA = FunNuloVal(ResPadrao!empr_nr_Porta)
        XGT_BAIRRO = FunNuloVal(ResPadrao!empr_tx_Bairro)
        XGT_ESTADO = FunNuloVal(ResPadrao!empr_tx_Estado)
        pPlanoCmascara = FunNuloVal(ResPadrao!empr_tx_planocmascara)
        pPCRnivel = FunNuloVal(ResPadrao!empr_nr_pcrnivel)
        pPlanoCNivel = FunNuloVal(ResPadrao!empr_nr_planocnivel)
        pCrAtivoPassivo = FunNuloVal(ResPadrao!empr_tx_CrAtivoPassivo)
        PLogo = FunNulo(ResPadrao!empr_im_caminhoLogo)
        XGI_PCATEGORIANVFINANCEIRO = FunNuloVal(ResPadrao!empr_nr_CategoriaNvContabil)
        'Lendo informações do dpto de compras
        With XGO_DPTOCOMPRAS
          .Responsavel = FunNulo(ResPadrao("empr_tx_ComResp").Value)
          .Telefone = FunNulo(ResPadrao("empr_tx_ComTelefone").Value)
          .Fax = FunNulo(ResPadrao("empr_tx_ComFax").Value)
          .Email = FunNulo(ResPadrao("empr_tx_ComEmail").Value)
        End With
        
        If ResPadrao!empr_tx_ExpMovPatCont = "S" Then
            XGB_EXPMOVPATCONT = True
        Else
            XGB_EXPMOVPATCONT = False
        End If
        
        If ResPadrao!empr_tx_AlertarSaldoAdiantamento = "S" Then
            XGB_AlertarSaldoAdiantamento = True
        Else
            XGB_AlertarSaldoAdiantamento = False
        End If
        
          'If XGB_PQUALIEST Then 'Pegar a conta de doação
              XLT_CONTADOACAO = FunNulo(ResPadrao("plco_cd_ReceitaDoacao").Value)
          'End If
        
          'O beócio que fez esta merda esqueceu de fechar o RecordSet. Logo, estava abrindo milhares de conexões na Costa Andrade
          'e gerando DeadLocks
          ResPadrao.Close
          Set ResPadrao = Nothing
    End If
    
    'Inicia o formulário principal do sistema
    Load MdiPrincipal
    MdiPrincipal.Label1.Left = 8745 + (MdiPrincipal.SysInfo1.WorkAreaWidth - 9600)
    MdiPrincipal.Picture1.Left = 7245 + (MdiPrincipal.SysInfo1.WorkAreaWidth - 9600)
    MdiPrincipal.BarraStatus.Panels(2).Text = PFantasia
    
    ResolucaoX = MdiPrincipal.SysInfo1.WorkAreaWidth
    Resolucaoy = MdiPrincipal.SysInfo1.WorkAreaHeight
    
    Call SubImagemFundo(ResolucaoX)
    
    MdiPrincipal.Show
    Exit Sub

RotuloErro:
  'Chama a rotina de tratamento de erros
  funTrataErros (ComMensagem)
  Resume Next
End Sub

Private Sub SubVerificaAtualizacao()

  Dim XLO_ATUALIZACAO As New clsUpdateFiles
  Dim XLI_QTD As Integer
  Dim XLT_TEMP As String * 255
  Dim XLT_PASTA As String
  
  On Error GoTo Erro:
  
  If GetPrivateProfileString("Atualizacao", "Pasta", "", XLT_TEMP, 255, App.path + "\QualiAdmFin.INI") <> 0 Then
    XLT_PASTA = Left$(XLT_TEMP, InStr(XLT_TEMP, Chr$(0)) - 1)
    
    XLI_QTD = XLO_ATUALIZACAO.SetListFiles(XLT_PASTA, App.path)
    If XLI_QTD > -1 Then
      If MsgBox("Existem novas atualizações a serem feitas. Deseja continuar?", vbInformation + vbYesNo, "Atualizações") = vbYes Then
        Shell App.path & "\..\QualiUpdate\QualiUpdate.exe " & XLT_PASTA & " " & App.path & " " & App.EXEName, vbNormalFocus
        End
      End If
    End If
  End If
  
Erro:
      
  If Err.Number <> 0 Then
    If Err.Number = 76 Then
      MsgBox "O sistema de atualização não está na pasta correta!", vbCritical, "ERRO"
    Else
      MsgBox Err.Description, vbCritical, "ERRO"
    End If
  End If

End Sub

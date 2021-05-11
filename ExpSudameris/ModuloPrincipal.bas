Attribute VB_Name = "ModuloPrincipal"
Option Explicit
'Indica a partir de que formulario a tela foi chamada
Global XFormulario As String

'Esta rotina inicia o sistema setando suas principais variaveis globais e fazendo a chamada para o formulário principal do sistema.
Public Sub Main()
    
    Dim ResPadrao As Object
    Dim XLT_DirPadrao As String * 254
    Dim XLT_CATEGORIA As String * 254
    Dim XLT_PARAMETROS() As String
    Dim XLO_DESCONTOS As New ADODB.Recordset
    
    'Seleciona o Diretorio Padrao do Sistema
    XGT_DirSistema = App.Path
    
    'seleciona o diretorio padrão de segurança
    If GetPrivateProfileString("BancoDeDados", "DirPadrao", "", XLT_DirPadrao, 255, XGT_DirSistema + "\QualiAdmFin.INI") = 0 Then
        MsgBox "Problemas na Leitura do Arquivo de configuração", vbCritical, "Configuração"
        End
    End If
    XGT_DirPadrao = Trim(XLT_DirPadrao)
    
    'PEGA O caminho do arquivo
    GetPrivateProfileString "BancoDeDados", "ExpBancos", "", XGT_LOCALARQ, 255, App.Path + "\QualiAdmFin.INI"
    
    'Indica o tipo de bloqueio a ser usado no sistema
    'Esta opção só é utilizada para o RDO
    TipoBloqueio = adLockPessimistic 'Lock Pessimista
    'TipoBloqueio = adLockReadOnly 'Lock Leitura
    'TipoBloqueio = adLockOptimistic 'Lock Otimista
    'TipoBloqueio = adLockBatchOptimistic 'Lock Otimista Baseadas no Valor de Linhas
            
    'Estabelece a conexão com o banco de dados selecionado
    If Not funAbreConexao() Then
       MsgBox "Houve problemas com a conexão com o Banco de Dados."
       End 'Finaliza o programa
    End If
    
    'Estabelece a conexão com o banco de dados selecionado para a emissao dos Relatórios
    If Not funAbreConexaoRelatorio() Then
       MsgBox "Houve problemas com a conexão com o Banco de Dados."
       End 'Finaliza o programa
    End If
        
    'Parametros passados pelo sistema que chama (código do funcionário e da empresa)
    XLT_PARAMETROS = Split(Command$(), ",")
    
    If UBound(XLT_PARAMETROS) > -1 Then
        PCodEmpresa = XLT_PARAMETROS(0)
        func_cd_operador = XLT_PARAMETROS(1)
        
        If func_cd_operador = "" Then
            MsgBox "Não foi possível definir o operador do sistema."
            End
        End If
    Else
        If PCodEmpresa = "" Then
           PCodEmpresa = 1
        End If
        If func_cd_operador = "" Then
            func_cd_operador = 1
        End If
    End If
    
    XGT_TIPOGRID = "B"
        
    SubQOpenRecordset ResPadrao, "Select * from Empresas where empr_cd_empresa=" & PCodEmpresa & "", Dinamico
    
    PFantasia = FunNulo(ResPadrao!empr_tx_fantasia)
    PEmpresa = FunNulo(ResPadrao!empr_tx_razaosocial)
    PEndereco = FunNulo(ResPadrao!empr_tx_endereco)
    PCidade = FunNulo(ResPadrao!empr_tx_cidade)
    XGI_CONTACORRENTE = FunNuloVal(ResPadrao!Empr_cd_codcontacorrente)
    XGT_CGC = FunNulo(ResPadrao!empr_tx_CGC)
    XGT_CEP = FunNulo(ResPadrao!empr_tx_Cep)
    XGT_NUMPORTA = FunNulo(ResPadrao!empr_nr_Porta)
    XGT_BAIRRO = FunNulo(ResPadrao!empr_tx_Bairro)
    XGT_ESTADO = FunNulo(ResPadrao!empr_tx_Estado)
    XGT_EXPCOMPROMCAPI = FunNulo(ResPadrao!empr_tx_ExpCompromissoCapi)

      
    '***************************INICIO - PEGA O ID DOS TIPOS DE DESCONTOS******************************
    SubQOpenRecordset XLO_DESCONTOS, "SELECT deac_cd_descacresc, deac_tx_Descricao FROM DescontosAcrescimos WHERE deac_tx_Descricao IN ('Multa por Atraso - Cliente', 'Juros por Atraso - Cliente', 'Deságio - Cliente', 'Descontos - Cliente', 'Seguro - Cliente', 'Outros - Cliente')", Estatico
    
    If Not XLO_DESCONTOS.EOF Then
        XLO_DESCONTOS.MoveFirst
        While Not XLO_DESCONTOS.EOF
        
            If (XLO_DESCONTOS!deac_tx_Descricao = "Multa por Atraso - Cliente") Then
                XGI_CODMULTA = XLO_DESCONTOS!deac_cd_descacresc
            ElseIf (XLO_DESCONTOS!deac_tx_Descricao = "Juros por Atraso - Cliente") Then
                XGI_CODJUROS = XLO_DESCONTOS!deac_cd_descacresc
            ElseIf (XLO_DESCONTOS!deac_tx_Descricao = "Deságio - Cliente") Then
                XGI_CODDESAGIO = XLO_DESCONTOS!deac_cd_descacresc
            ElseIf (XLO_DESCONTOS!deac_tx_Descricao = "Descontos - Cliente") Then
                XGI_CODDESCONTO = XLO_DESCONTOS!deac_cd_descacresc
            ElseIf (XLO_DESCONTOS!deac_tx_Descricao = "Outros - Cliente") Then
                XGI_CODSEGURO = XLO_DESCONTOS!deac_cd_descacresc
            ElseIf (XLO_DESCONTOS!deac_tx_Descricao = "Seguro - Cliente") Then
                XGI_CODOUTROS = XLO_DESCONTOS!deac_cd_descacresc
            End If
        
            XLO_DESCONTOS.MoveNext
        Wend
    End If
    
    XLO_DESCONTOS.Close
    Set XLO_DESCONTOS = Nothing
    '***************************FIM - PEGA O ID DOS TIPOS DE DESCONTOS******************************
        
      
    'Inicia o formulário principal do sistema
    ResolucaoX = TelaInicial.SysInfo1.WorkAreaWidth
    Resolucaoy = TelaInicial.SysInfo1.WorkAreaHeight
    
    TelaInicial.Show
Exit Sub

RotuloErro:
    
    'Chama a rotina de tratamento de erros
    funTrataErros (ComMensagem)
    Resume Next
End Sub

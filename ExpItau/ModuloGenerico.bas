Attribute VB_Name = "ModuloGenerico"
Option Explicit
Sub subCarregaVariaveis(Sql As String)
    'Atribui os parâmetros às variáveis globais
    '*Ordem = OrdemTabela
    '*Campos_Obrig = CamposObrigatorios
    FiltroOrig = funTrataSql(Sql)
    FiltroAtual = FiltroOrig
    
    'Inicializa as variáveis de controle de filtro da tabela
    filtra_especial = False ' Indica que não há filtros pré-definidos
    Filtrou = False 'Indica que não há filtros definidos pelo usuário
    Tem_Filtro = False 'Indica que não há filtros definidos pelo usuário

    Quantid = 0 'Zera a quantidade de registros da tabela

    'Inicializa o número de Filtros e o array de Filtros.
    ContFil = 0
    ReDim Filtros(3, 0)

    'Inicializa o número de Ordens e o vetor de ordens
    ConTelaOrdem = 1
    ReDim Ordens(1) As Integer
    Ordens(1) = 1
End Sub
'Rotina para teste de proteção
Sub subTestaProtecao(ByVal NomeApl As String, ByVal Id As String, ByVal DirPadrao)
    Dim XLI_ResSeg As Integer

    'XLI_ResSeg = hl_leprot_vb(NomeApl, Id, DirPadrao, 1, 1, ProtResult)
  '  If XLI_ResSeg <> 0 Then
  '     MsgBox Str(ProtResult.CodRet) & " - " & RTrim(ProtResult.MsgRet), 16, "Verificação da Proteção"
  '     Beep
  '     End
  '  End If

'Armazena o Nº de Serie do aplicativo
    XGI_NSerie = Left(ProtResult.NroSerie, 6)
    
'Verifica se é Multi ou Mono Empresa
    'If Mid(XGI_NSerie, 1, 1) = "1" Then '1- Multi / 0- Mono Empresa
        pTipoCopia = "Multi"
    'Else
        pTipoCopia = "Mono"
    'End If
    
    'Verifica se Possui o Sistema QualiFin
    'If Mid(XGI_NSerie, 2, 1) = "1" Then '1- Sim / 0- Não
        XGB_PQUALIFIN = True
    'Else
    '    XGB_PQUALIFIN = False
    'End If
    
    'Verifica se Possui o Sistema de Contabilidade
    'If Mid(XGI_NSerie, 6, 1) = "1" Then '1- Sim / 0- Não
        XGB_PQUALICONT = True
    'Else
    '    XGB_PQUALICONT = False
    'End If

End Sub

Function funDefinirCodigoLanc(XLD_DATA As Date, XLI_NrLote As Integer) As Integer
  '**********************************************************************
  'Define o cod do Lancamento para a Data e o Lote Atual
  'chamdo pelo CmdGravar
  '*****************************************************************************
  Dim XLO_RSLANC As New ADODB.Recordset

  XLO_RSLANC.Open _
    "SELECT max(lanc_cd_lanc) as lanc_cd_lanc " & _
    "FROM Lancamentos " & _
    "WHERE lanc_dt_data = " & FunNuloData(XLD_DATA, NomeSgbd) & _
    " AND lanc_nr_Lote = " & XLI_NrLote & _
    " AND empr_cd_empresa = " & PCodEmpresa, Conexao
      
  If Not IsNull(XLO_RSLANC("lanc_cd_Lanc").Value) Then
    funDefinirCodigoLanc = XLO_RSLANC("lanc_cd_Lanc").Value + 1
  Else
    funDefinirCodigoLanc = 1
  End If
End Function


Function FunDefinirUltimoLote(Data As Date) As Integer
  '*****************************************************************************
  'Define o numero do maior Lote existente para a data do lancamento
  'Chamado qnd abre o Formulario de Lancamento e qnd Muda a Data
  '*****************************************************************************
  Dim XLO_RSLANC As New ADODB.Recordset
  Dim XLT_SQL As String
        
  XLT_SQL = "SELECT Max(lanc_nr_lote) as Lanc_nr_Lote " & _
              "FROM Lancamentos " & _
              "WHERE lanc_dt_data = " & FunNuloData(Data, NomeSgbd) & _
              "  AND empr_cd_empresa = " & PCodEmpresa
  SubQOpenRecordset XLO_RSLANC, XLT_SQL, Estatico
  
  'se o retorno do select for vazio eh pq não houve nenhum lancamento para
  'essa data, entao o lote vai para 1. se houver retorno entao afuncao
  'retorna o numero do maior lote para a data
  If IsNull(XLO_RSLANC("lanc_nr_Lote")) Then
    FunDefinirUltimoLote = 1
  Else
    FunDefinirUltimoLote = XLO_RSLANC("Lanc_nr_Lote").Value + 1
  End If
End Function



Sub subImprimeListagemGRIDUnBound(xOrientacao As Integer, xGrid As TrueDBGrid70.PrintInfo, Texto As String)
    With xGrid
        ' Setar margens a depender do tipo de impressão
        If xOrientacao = 1 Then
            .SettingsOrientation = 1
            .SettingsMarginTop = 567
            '.SettingsMarginBottom = 630
            .SettingsMarginBottom = 750
            '.SettingsMarginLeft = 1134
            .SettingsMarginLeft = 750
            .SettingsMarginRight = 567
        Else
            .SettingsOrientation = 2
            .SettingsMarginTop = 1134
            '.SettingsMarginBottom = 630
            .SettingsMarginBottom = 750
            .SettingsMarginLeft = 567
            '.SettingsMarginRight = 567
            .SettingsMarginRight = 850
       End If
        
        'Determina se o cabeçalho será definido pelo usuário e o seu tamanho
        'Neste caso o cabeçalho será definido pelo usuário através de API
        '.PageHeaderOwnerDraw = True
        '.PageHeaderHeight = 4000
        
        ' Setar o cabeçalho
        .PageHeaderFont.Name = "Arial"
        .PageHeaderFont.Size = 12
        .PageHeaderFont.Bold = True
        .PageHeader = Texto + " \t\t" + CStr(Now)
        
        ' O Cabeçalho e o rodapé devem aparecer em todas as páginas
        .RepeatColumnHeaders = True
        .RepeatColumnFooters = True
        
        ' Imprimir texto no rodapé
        .PageFooter = "QUALIDADOS\t\tPágina: \p/\P"
        
        ' Chamar o Preview de impressão
        .PageSetup
        '.PrintPreview
    End With
    If Not xGrid.PageSetupCancelled Then
        xGrid.PrintPreview
    End If
    

End Sub








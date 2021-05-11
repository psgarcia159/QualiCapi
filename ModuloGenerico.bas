Attribute VB_Name = "ModuloGenerico"
Option Explicit

Public Declare Function GetVersionExA Lib "kernel32" _
               (lpVersionInformation As OSVERSIONINFO) As Integer
 
Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
 
Public Function getVersion() As String
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer

   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)

   With osinfo
   Select Case .dwPlatformId

    Case 1
    
        Select Case .dwMinorVersion
            Case 0
                getVersion = "Windows 95"
            Case 10
                getVersion = "Windows 98"
            Case 90
                getVersion = "Windows Millennium"
        End Select

    Case 2
        Select Case .dwMajorVersion
            Case 3
                getVersion = "Windows NT 3.51"
            Case 4
                getVersion = "Windows NT 4.0"
            Case 5
                If .dwMinorVersion = 0 Then
                    getVersion = "Windows 2000"
                Else
                    getVersion = "Windows XP"
                End If
            Case 6
               
               getVersion = "Windows 7"
               
        End Select

    Case Else
       getVersion = "Failed"
End Select
 
   End With
End Function
            
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

Sub subChecaStatusRM(XLL_ITEMRM As Long)
    Dim XLO_ITENSRM         As ADODB.Recordset
    Dim XLL_CODRM           As Long
    Dim XLL_CODSAIDA        As Long
    Dim XLL_CODSOLICITACAO  As Long
    Dim XLB_CONCLUIU        As Boolean
    
    XLL_CODRM = -1
    XLL_CODSAIDA = -1
    XLL_CODSOLICITACAO = -1
    XGT_SQL = "SELECT * FROM ItensRM WHERE itrm_cd_codigo = " & CStr(XLL_ITEMRM)
    SubQOpenRecordset XLO_ITENSRM, XGT_SQL, Estatico
    If Not XLO_ITENSRM.EOF Then
        XLL_CODRM = XLO_ITENSRM!rema_cd_Codigo
    End If
    XLO_ITENSRM.Close
    XGT_SQL = "SELECT * FROM ItensRMSolicitacoes WHERE itrm_cd_codigo = " & CStr(XLL_ITEMRM)
    SubQOpenRecordset XLO_ITENSRM, XGT_SQL, Estatico
    If Not XLO_ITENSRM.EOF Then
        XLL_CODSOLICITACAO = XLO_ITENSRM!soli_cd_codigo
    End If
    XLO_ITENSRM.Close
    XGT_SQL = "SELECT * FROM ItensRMSaidaMateriais WHERE itrm_cd_codigo = " & CStr(XLL_ITEMRM)
    SubQOpenRecordset XLO_ITENSRM, XGT_SQL, Estatico
    If Not XLO_ITENSRM.EOF Then
        XLL_CODSAIDA = XLO_ITENSRM!sama_nr_Ordem
    End If
    XLO_ITENSRM.Close
    
    XLB_CONCLUIU = True
    XGT_SQL = "SELECT * FROM ItensRMSolicitacoes WHERE soli_cd_codigo = " & CStr(XLL_CODSOLICITACAO)
    SubQOpenRecordset XLO_ITENSRM, XGT_SQL, Estatico
    Do While Not XLO_ITENSRM.EOF
        If IsNull(XLO_ITENSRM!irms_dt_Atendimento) Then
            XLB_CONCLUIU = False
        End If
        XLO_ITENSRM.MoveNext
    Loop
    XLO_ITENSRM.Close
    XGT_SQL = "SELECT * FROM ItensRMSaidaMateriais WHERE sama_nr_ordem = " & CStr(XLL_CODSAIDA)
    SubQOpenRecordset XLO_ITENSRM, XGT_SQL, Estatico
    Do While Not XLO_ITENSRM.EOF
        If IsNull(XLO_ITENSRM!irsm_dt_Atendimento) Then
            XLB_CONCLUIU = False
        End If
        XLO_ITENSRM.MoveNext
    Loop
    XLO_ITENSRM.Close
    Set XLO_ITENSRM = Nothing
    
    If XLB_CONCLUIU Then
        XGT_SQL = "UPDATE RequisicoesMateriais SET rema_dt_fechamento = " & FunNuloData(Date, NomeSgbd) & " WHERE rema_cd_codigo = " & CStr(XLL_CODRM)
        Conexao.Execute XGT_SQL
    Else
        XGT_SQL = "UPDATE RequisicoesMateriais SET rema_dt_fechamento = Null WHERE rema_cd_codigo = " & CStr(XLL_CODRM)
        Conexao.Execute XGT_SQL
    End If
End Sub

'a descobrir
Function FunLocalizaJanelasAtivas(NomeFormulario As String) As Long
    Dim x As Long
    FunLocalizaJanelasAtivas = -1
    For x = 0 To UBound(VetorJanelas)
        If NomeFormulario = VetorJanelas(x) Then
            FunLocalizaJanelasAtivas = x
        End If
    Next
End Function

'PosicaoJanelaVetor: A Posicao do formulário no vetor,
'se o formulário não existir o valor retornado será -1
Sub subManutencaoJanelasAtivas(Tipo As String, NomeFormulario As String)
    Dim PosicaoJanelaVetor As Long, x As Long
    
    'Tipo: (I)nserir ou (R)emover item da lista de janelas ativas
    'NomeFormulario: Formulário ativo
    
    PosicaoJanelaVetor = FunLocalizaJanelasAtivas(NomeFormulario)
    If PosicaoJanelaVetor <> -1 Then
        If Tipo = "R" Then
            If UBound(VetorJanelas) = 0 Then
                VetorJanelas(0) = ""
            ElseIf UBound(VetorJanelas) = PosicaoJanelaVetor Then
                ReDim Preserve VetorJanelas(PosicaoJanelaVetor - 1)
            Else
                For x = PosicaoJanelaVetor To UBound(VetorJanelas) - 1
                    VetorJanelas(x) = VetorJanelas(x + 1)
                Next
                ReDim Preserve VetorJanelas(x - 1)
            End If
        End If
    Else
        If Tipo = "I" Then
            If UBound(VetorJanelas) = 0 Then
                If FunNulo(VetorJanelas(0)) <> "" Then
                    x = UBound(VetorJanelas) + 1
                    ReDim Preserve VetorJanelas(x)
                End If
            Else
                x = UBound(VetorJanelas) + 1
                ReDim Preserve VetorJanelas(x)
            End If
            VetorJanelas(x) = NomeFormulario
        End If
    End If
End Sub

'Rotina para teste de proteção
Sub subTestaProtecao(ByVal NomeApl As String, ByVal Id As String, ByVal DirPadrao)
    Dim XLI_ResSeg As Integer
    Dim XLT_EMPRESA As String
    Dim XLO_EMPRESA As ADODB.Recordset

    'Inclusão em 24/04/09 - Patrícia
    SubQOpenRecordset XLO_EMPRESA, "SELECT TOP 1 empr_tx_cliente FROM empresas WHERE empr_tx_cliente IS NOT NULL", Estatico
    If Not XLO_EMPRESA.EOF Then
        XLT_EMPRESA = XLO_EMPRESA!empr_tx_cliente
    End If
    XLO_EMPRESA.Close

    Set XLO_EMPRESA = Nothing
    
    'Inserido em 08/09/09
    If XLT_EMPRESA = "" Then
        pTipoCopia = "Mono"
        XGB_PQUALIFIN = True
    ElseIf XLT_EMPRESA <> "Temae" Then

'            XLI_ResSeg = hl_leprot_vb(NomeApl, Id, DirPadrao, 1, 1, ProtResult)
'
'            'Erro 3 "O programa não pode operar nesse disco.",
'            'acontece quando altera o compartilhamento do hd onde está a proteção
'            'ou o usuário nao tenha permissao sobre a pasta
'            If XLI_ResSeg <> 0 And XLI_ResSeg <> 3 Then
'               MsgBox Str(ProtResult.CodRet) & " - " & RTrim(ProtResult.MsgRet), 16, "Verificação da Proteção"
'               Beep
'               End
'            End If
'
''            'Armazena o Nº de Serie do aplicativo
'            XGI_NSerie = Left(ProtResult.NroSerie, 7)
'
''            'Verifica se é Multi ou Mono Empresa
'             If Mid(XGI_NSerie, 1, 1) = "1" Then '1- Multi / 0- Mono Empresa
                pTipoCopia = "Multi"
'             Else
'                pTipoCopia = "Mono"
'             End If
''
''            'Verifica se Possui o Sistema QualiFin
'            If Mid(XGI_NSerie, 2, 1) = "1" Then '1- Sim / 0- Não
                XGB_PQUALIFIN = True
'            Else
'                XGB_PQUALIFIN = False
'            End If
''
''            'Verifica se Possui o Sistema QualiEst
'            If Mid(XGI_NSerie, 3, 1) = "1" Then '1- Sim / 0- Não
                XGB_PQUALIEST = True
'            Else
'                XGB_PQUALIEST = False
'            End If
''
''            'Verifica se Possui o Sistema Compras
'            If Mid(XGI_NSerie, 4, 1) = "1" Then '1- Sim / 0- Não'
                XGB_PCOMPRAS = True
'            Else
'                XGB_PCOMPRAS = False
'            End If
''
''            'Verifica se Possui o Sistema Capi
'            If Mid(XGI_NSerie, 5, 1) = "1" Then '1- Sim / 0- Não
                XGB_PCAPI = True
'            Else
'                XGB_PCAPI = False
'            End If
''
''            'Verifica se Possui o Sistema de Contabilidade
'            If Mid(XGI_NSerie, 6, 1) = "1" Then '1- Sim / 0- Não
               XGB_PQUALICONT = True
'            Else
'                XGB_PQUALICONT = False
'            End If
''
''            'Verifica se Possui o Sistema de Qualifats
'            If Mid(XGI_NSerie, 7, 1) = "1" Then '1- Sim / 0- Não
                XGB_PQUALIFATS = True
'            Else
'                XGB_PQUALIFATS = False
'            End If

    Else
            pTipoCopia = "Multi"
            XGB_PQUALIFIN = True
    End If
End Sub


Function FunVerificaVersao() As Boolean
    
    Dim XLO_VERIFICA As ADODB.Recordset
    Dim versao As Double
    Dim versaobanco As Double
    
    
    versao = CDbl(Format(App.Major, "#0") & "," & Format(App.Revision, "00"))
    
    'Verifica a versão do cliente no banco de dados
    SubQOpenRecordset XLO_VERIFICA, "SELECT sist_vl_versao FROM Sistema WHERE sist_tx_nome = '" & XGT_NomeAplic & "'", Estatico
    
    versaobanco = CDbl(XLO_VERIFICA!sist_vl_versao)
    
    XLO_VERIFICA.Close
    Set XLO_VERIFICA = Nothing
    
    
    'Se a versão atual for maior que a existente atualiza o banco
    If versao > versaobanco Then
        Conexao.Execute "UPDATE Sistema SET sist_vl_versao = " & Replace(CStr(versao), ",", ".") & " WHERE sist_tx_nome = '" & XGT_NomeAplic & "'"
        FunVerificaVersao = True
    'Senão retorna falso para solicitar atualização do sistema
    ElseIf versao = versaobanco Then
        FunVerificaVersao = True
    Else
        FunVerificaVersao = False
    End If
    
    
End Function
'Ajustar a Tela à resolução do monitor
Sub subAjustaTelaLoc()
    Formulario.Top = 0
    Formulario.Height = Resolucaoy - 1380
    Formulario.Width = ResolucaoX - 60
    Formulario.Left = 5
    Formulario.TDBGrid1.Top = 20
    Formulario.TDBGrid1.Left = 75
    Formulario.TDBGrid1.Width = Formulario.Width - 300
    Formulario.TDBGrid1.Height = Formulario.Height - Formulario.TDBGrid1.Top - Formulario.PanBotoes.Height - Formulario.cboLocalizar.Height - 150
    Formulario.lblLocalizar.Top = Formulario.TDBGrid1.Top + Formulario.TDBGrid1.Height + 50
    Formulario.cboLocalizar.Top = Formulario.TDBGrid1.Top + Formulario.TDBGrid1.Height + 50
    Formulario.Adodc1.Top = Formulario.TDBGrid1.Top + Formulario.TDBGrid1.Height + 50
    Formulario.PanBotoes.Top = Formulario.TDBGrid1.Top + Formulario.TDBGrid1.Height + Formulario.cboLocalizar.Height + 100
    Formulario.PanBotoes.Left = 75 + (Formulario.TDBGrid1.Width - Formulario.PanBotoes.Width) / 2
End Sub

'Ajustar a Tela à resolução do monitor
Sub SubAjustaTela()
    Formulario.Top = 0
    Formulario.Height = Resolucaoy - 1380
    Formulario.Width = ResolucaoX - 60
    Formulario.Left = 5
    Formulario.TDBGrid1.Top = 20
    Formulario.TDBGrid1.Left = 75
    Formulario.TDBGrid1.Width = Formulario.Width - 300
    Formulario.TDBGrid1.Height = Formulario.Height - Formulario.TDBGrid1.Top - Formulario.PanBotoes.Height - 100
    Formulario.PanBotoes.Top = Formulario.TDBGrid1.Top + Formulario.TDBGrid1.Height + 50
    Formulario.PanBotoes.Left = 75 + (Formulario.TDBGrid1.Width - Formulario.PanBotoes.Width) / 2
End Sub

Sub subAjustaTelaGeral(Formulario As Form)
  'Ajusta a tela à resolução do monitor
  'A tela deve possuir os seguintes paineis: PanBotoes, PanLocalizar, PanGeral
  'PanBotoes: Painel que contém todos os botões da tela
  'PanLocalizar: Painel que contém a descrição e o combo de localizar
  'PanGeral: Painel que contém todos os outros painéis. Este painel deve estar
  '  alinhado ao topo da tela.
  'O grid deve estar alinhado ao todo da janela
  'Este procedimento deve ser usado no Load e Resize do formulário
  
  With Formulario
    If (.Width >= 9180) And (.Height >= 5025) Then
      .TDBGrid1.Height = .ScaleHeight - .panGeral.Height
      .panLocalizar.Left = (.panGeral.Width - .panLocalizar.Width) / 2
      .PanBotoes.Left = (.panGeral.Width - .PanBotoes.Width) / 2
    Else
      On Error GoTo RotuloErro
      If .Width < 9180 Then
        .Width = 9180
      Else
        .Height = 5025
      End If
    End If
  End With
  
RotuloErro:
  
End Sub

' Verifica a quantidade de registros no recordset de ControleDados
' e atribui ao painel na barra de status do MdiPrincipal.
Public Sub subContaReg(ControleDados As Object)
   Dim cont As Integer
    
   cont = 0
   If Not ControleDados.Recordset.EOF Then
      ControleDados.Recordset.MoveLast
      cont = ControleDados.Recordset.RecordCount
      ControleDados.Recordset.MoveFirst
   End If
   
   MdiPrincipal.BarraStatus.Panels(3).Text = Str(cont) + "/" + Str(Quantid)
   
End Sub

'efetua copiar de registro
Public Sub subCopiaRegistro(ControleDados As Object, Registro() As Variant)
    Dim cont As Integer
    Dim Max As Integer
    
    If TipoAcesso = "Dao" Then
        Max = ControleDados.Recordset.Fields.Count
        ControleDados.Recordset.Edit
        For cont = 0 To Max - 1
            Registro(cont) = ControleDados.Recordset.Fields.Item(cont)
        Next cont
        ControleDados.Recordset.Update
    Else
        Max = ControleDados.Recordset.rdoColumns.Count
        For cont = 0 To Max - 1
            Registro(cont) = ControleDados.Resultset.rdoColumns.Item(cont)
        Next cont
        ControleDados.Resultset.Update
    End If
End Sub

Public Sub subGuardaUltimo(CampoChave As String, valor As Long)

    ' Aramazena o valor obtido em funPegaUltimo

    Dim tb_ultimos As ADODB.Recordset
    
    SubQOpenRecordset tb_ultimos, "select * from ULTIMOS", Dinamico
    
    If tb_ultimos.EOF Then
       tb_ultimos.AddNew
       letQRecordsetFields(tb_ultimos, CampoChave) = 1
    Else
       'tb_ultimos.EditMode
       If IsNull(getQRecordsetFields(tb_ultimos, CampoChave)) Then
          letQRecordsetFields(tb_ultimos, CampoChave) = 1
       Else
          letQRecordsetFields(tb_ultimos, CampoChave) = valor
       End If
    End If
    
    tb_ultimos.Update
    tb_ultimos.Close
    Set tb_ultimos = Nothing
    
    
End Sub

Public Sub subRestauraRegistro(ControleDados As Object, Registro() As Variant)
    Dim cont As Integer
    Dim Max As Integer
    
    If TipoAcesso = "Dao" Then
        Max = ControleDados.Recordset.Fields.Count
        ControleDados.Recordset.Edit
        For cont = 0 To Max - 1
            ControleDados.Recordset.Fields.Item(cont) = Registro(cont)
        Next cont
        ControleDados.Recordset.Update
    Else
        Max = ControleDados.Resultset.rdoColumns.Count
        For cont = 0 To Max - 1
            ControleDados.Resultset.rdoColumns.Item(cont) = Registro(cont)
        Next cont
        ControleDados.Resultset.Update
    End If

End Sub

'Seleciona conteudo do Maskedit
Sub subSelecionaMSK(XObjeto As MaskEdBox)
    XObjeto.SelStart = 0
    XObjeto.SelLength = XObjeto.MaxLength
End Sub

'Seleciona conteudo de TextBox
Sub subSelecionaTXT(XObjeto As TextBox)
    XObjeto.SelStart = 0
    If XObjeto.MaxLength > 0 Then
        XObjeto.SelLength = XObjeto.MaxLength
    Else
        XObjeto.SelLength = Len(Trim(XObjeto.Text))
    End If
End Sub

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
  
  XLO_RSLANC.Close
  Set XLO_RSLANC = Nothing
  
End Function

Function funDefinirCodigoLanc(XLD_DATA As Date, XLI_NrLote As Integer) As Integer
  '**********************************************************************
  'Encontra o próximo cod. livre de lancamento para a data e o lote atual
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
  
  XLO_RSLANC.Close
  Set XLO_RSLANC = Nothing
End Function

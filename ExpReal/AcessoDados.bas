Attribute VB_Name = "AcessoDados"
Option Explicit
'*Definição de constantes para os tipos de recordset ou resultset a serem utilizados
Global Const Dinamico = "Dinamico"
Global Const Direto = "Direto"
Global Const Estatico = "Estatico"

'*Indicam os procedimentos de inicialização do Banco de Dados
Global Conexao As New ADODB.Connection          ' Representa a conexão com o banco de dados.
Global ConexaoRelatorio As New ADODB.Connection ' Representa a conexão com o banco de dados para Relatórios.

Global NomeSgbd As String     ' Indica o nome do banco de dados a ser utilizado ex: Access

Global TipoBloqueio As Integer         ' Indica o tipo de bloqueio a ser utilizado ex: Pessimista ou Otimista
                                       ' O Bloqueio Pessimista bloqueia o registro a partir do momento que
                                       ' é realizado o edit e libera após a realização do update.
                                       ' O Bloqueio Otimista bloqueia o registro apenas no momento que está
                                       ' sendo realizado o update.

'Conecta o controle de dados
Public Sub subConectarControleDadosNV(ControleDados As Object, Sql As String, TipoRecordSet As String)
    Dim ObjRecordset As Object
    
    Set ObjRecordset = New ADODB.Recordset
    
    'Seleciona o tipo de acesso
    Select Case TipoRecordSet
        Case "Dinamico"  ' Recordset que pode ser alterado
            ObjRecordset.Open funTrataSql(Sql), Conexao, adOpenDynamic, TipoBloqueio, adCmdText
            'ControleDados.CursorType = adOpenDynamic
        Case "Direto"       ' Recordset de leitura com movimentacao em apenas uma direcao
            ObjRecordset.Open funTrataSql(Sql), Conexao, adOpenForwardOnly, TipoBloqueio, adCmdText
            'ControleDados.CursorType = adOpenForwardOnly
        Case "Estatico"    ' Recordset de leitura com livre movimentacao
            ObjRecordset.CursorLocation = adUseClient
            ObjRecordset.Open funTrataSql(Sql), Conexao, adOpenStatic, TipoBloqueio, adCmdText
'            ControleDados.CursorType = adOpenStatic
    End Select
    Set ControleDados.Recordset = ObjRecordset
'    ControleDados.ConnectionString = Conexao.ConnectionString
'    ControleDados.LockType = TipoBloqueio
'    ControleDados.RecordSource = funTrataSql(Sql)
'    ControleDados.Refresh
End Sub



' Abre um conjunto de registros de acordo com o Tipo de Acesso
Public Sub SubQOpenRecordset(ObjRecordset As Object, Sql As String, TipoRecordSet As String)
    Set ObjRecordset = New ADODB.Recordset
    Select Case TipoRecordSet
        Case "Dinamico"  ' Recordset que pode ser alterado
            'Set ObjRecordset = Conexao.OpenRecordset(funTrataSql(sql), adOpenDynamic)
            
            'Se usar o CursorType = adOpenDynamic, o propriedade Bookmark não é setada,
            'consequentemente não podemos pegar a chave para gravação de tabelas relacionadas
            ObjRecordset.CursorType = adOpenKeyset
            ObjRecordset.LockType = TipoBloqueio
            ObjRecordset.Open funTrataSql(Sql), Conexao, , , adCmdText
            'ObjRecordset.Open funTrataSql(sql), Conexao, adOpenDynamic, TipoBloqueio, adCmdText
        Case "Direto"       ' Recordset de leitura com movimentacao em apenas uma direcao
            'Set ObjRecordset = Conexao.OpenRecordset(funTrataSql(sql), adOpenForwardOnly)
            ObjRecordset.Open funTrataSql(Sql), Conexao, adOpenForwardOnly, TipoBloqueio, adCmdText
        Case "Estatico"    ' Recordset de leitura com livre movimentacao
            'Set ObjRecordset = Conexao.OpenRecordset(funTrataSql(sql), adOpenStatic)
            ObjRecordset.CursorLocation = adUseClient
            ObjRecordset.Open funTrataSql(Sql), Conexao, adOpenStatic, TipoBloqueio, adCmdText
    End Select
End Sub

'Abre a conexao com o banco
Public Function funAbreConexao() As Boolean
    Dim XLT_STRINGCONEXAO   As String * 254
    Dim XLT_TIPOBANCO       As String * 254
    Dim XLT_NOMEBANCO       As String * 254
    
    On Error GoTo RotuloErro
    Screen.MousePointer = vbHourglass
    'Indica o banco de dados a ser utilizado
    If GetPrivateProfileString("BancoDeDados", "TipoBanco", "", XLT_TIPOBANCO, 255, App.Path + "\QualiAdmFin.INI") = 0 Then
        MsgBox "Problemas na Leitura do Arquivo de configuração", vbCritical, "Configuração"
    End If
    
    'PEGA O ENDEREÇOÇ DO BANCO
    GetPrivateProfileString "BancoDeDados", "Endereco", "", XLT_STRINGCONEXAO, 255, App.Path + "\QualiAdmFin.INI"
    
    'PEGA O NOME DO BANCO
    GetPrivateProfileString "BancoDeDados", "NomeBanco", "", XLT_NOMEBANCO, 255, App.Path + "\QualiAdmFin.INI"
    
    Select Case Val(XLT_TIPOBANCO)
        Case 1
            NomeSgbd = "Access"
            'Leitura da String de conexao do banco
            Conexao.Provider = "Microsoft.Jet.OLEDB.4.0"
            Conexao.Open FunStrArqIni(XLT_STRINGCONEXAO), "admin", ""
        Case 2
            NomeSgbd = "Sql Server"
            Conexao.Provider = "sqloledb"
            XGT_SQL = "Data Source=" & FunStrArqIni(XLT_STRINGCONEXAO) & _
                      ";Initial Catalog=" & FunStrArqIni(XLT_NOMEBANCO) & _
                      ";User Id=qualiadmfin;Password=qd; "
            Conexao.Open XGT_SQL
    End Select
    Screen.MousePointer = vbDefault
    funAbreConexao = True
Exit Function

RotuloErro:

    Select Case funTrataErros(ComMensagem)
    Case ErrFalhaConexao
        funAbreConexao = False
        Screen.MousePointer = vbDefault
        Exit Function
    End Select

End Function
'Abre a conexao com o banco
Public Function funAbreConexaoRelatorio() As Boolean
    Dim XLT_STRINGCONEXAO   As String * 254
    Dim XLT_TIPOBANCO       As String * 254
    Dim XLT_NOMEBANCO       As String * 254
    
    On Error GoTo RotuloErro
    Screen.MousePointer = vbHourglass
    'Indica o banco de dados a ser utilizado
    If GetPrivateProfileString("BancoDeDados", "TipoBanco", "", XLT_TIPOBANCO, 255, App.Path + "\QualiAdmFin.INI") = 0 Then
        MsgBox "Problemas na Leitura do Arquivo de configuração", vbCritical, "Configuração"
    End If
    
    'PEGA O ENDEREÇOÇ DO BANCO
    GetPrivateProfileString "BancoDeDados", "Endereco", "", XLT_STRINGCONEXAO, 255, App.Path + "\QualiAdmFin.INI"
    
    'PEGA O NOME DO BANCO
    GetPrivateProfileString "BancoDeDados", "NomeBanco", "", XLT_NOMEBANCO, 255, App.Path + "\QualiAdmFin.INI"
    
    Select Case Val(XLT_TIPOBANCO)
        Case 1
            NomeSgbd = "Access"
            'Leitura da String de conexao do banco
            ConexaoRelatorio.Provider = "Microsoft.Jet.OLEDB.4.0"
            ConexaoRelatorio.Open FunStrArqIni(XLT_STRINGCONEXAO), "admin", ""
        Case 2
            NomeSgbd = "Sql Server"
            ConexaoRelatorio.Provider = "sqloledb"
            XGT_SQL = "Data Source=" & FunStrArqIni(XLT_STRINGCONEXAO) & _
                      ";Initial Catalog=" & FunStrArqIni(XLT_NOMEBANCO) & _
                      ";User Id=qualiadmfin;Password=qd; "
            ConexaoRelatorio.Open XGT_SQL
    End Select
    Screen.MousePointer = vbDefault
    funAbreConexaoRelatorio = True

Exit Function

RotuloErro:

    Select Case funTrataErros(ComMensagem)
    Case ErrFalhaConexao
        funAbreConexaoRelatorio = False
        Screen.MousePointer = vbDefault
        Exit Function
    End Select

End Function

'Carrega os Valores da Tabela Atual inicializando o seu controle de dados.
Public Sub subrecarregadadosNV(ControleDados As Object, TabelaAtiva As String, _
                                         OrdemTabela As String, CamposObrigatorios As String, _
                                         FiltroOriginal As String)
    Dim Cont As Long
    'Atribui os parâmetros às variáveis globais
    Ordem = OrdemTabela
    Campos_Obrig = CamposObrigatorios
    'FiltroOrig = funTrataSql(FiltroOriginal)
    'FiltroAtual = FiltroOrig
    FiltroAtual = funTrataSql(FiltroOriginal)

    'Inicializa o Controle de Dados.
    If Ordem <> "" Then
        subConectarControleDadosNV ControleDados, FiltroAtual + " ORDER BY " + Ordem, Estatico
    Else
        subConectarControleDadosNV ControleDados, FiltroAtual, Estatico
    End If

    Cont = 0 'Zera a quantidade de registros da tabela
    
    If Not ControleDados.Recordset.EOF Then
       ControleDados.Recordset.MoveLast
       Cont = ControleDados.Recordset.RecordCount 'Informa a quantidade de registros da tabela.
    End If

    ' Mostra a quantidade de registros na barra de status
   'MdiPrincipal.BarraStatus.Panels(3).Text = Str(Cont) + "/" + Str(Quantid)

    ' Aponta para o primeiro registro da tabela
    If Cont <> 0 Then
       ControleDados.Recordset.MoveFirst
    End If
End Sub

'Carrega os Valores da Tabela Atual inicializando o seu controle de dados.
Public Sub subCarregaDadosNV(ControleDados As Object, TabelaAtiva As String, _
                                         OrdemTabela As String, CamposObrigatorios As String, _
                                         FiltroOriginal As String)
    'Atribui os parâmetros às variáveis globais
    Ordem = OrdemTabela
    Campos_Obrig = CamposObrigatorios
    FiltroOrig = funTrataSql(FiltroOriginal)
    FiltroAtual = FiltroOrig

    'Inicializa as variáveis de controle de filtro da tabela
    filtra_especial = False ' Indica que não há filtros pré-definidos
    Filtrou = False 'Indica que não há filtros definidos pelo usuário
    Tem_Filtro = False 'Indica que não há filtros definidos pelo usuário
            
    'Inicializa o Controle de Dados.
    If Ordem <> "" Then
        subConectarControleDadosNV ControleDados, FiltroOrig + " ORDER BY " + Ordem, Estatico
    Else
        subConectarControleDadosNV ControleDados, FiltroOrig, Estatico
    End If

    Quantid = 0 'Zera a quantidade de registros da tabela
    
    If Not ControleDados.Recordset.EOF Then
       ControleDados.Recordset.MoveLast
       Quantid = ControleDados.Recordset.RecordCount 'Informa a quantidade de registros da tabela.
    End If

    ' Mostra a quantidade de registros na barra de status
    'MdiPrincipal.BarraStatus.Panels(3).Text = Str(Quantid) + "/" + Str(Quantid)

    ' Aponta para o primeiro registro da tabela
    If Quantid <> 0 Then
       ControleDados.Recordset.MoveFirst
    End If

    'Inicializa o número de Filtros e o array de Filtros.
    ContFil = 0
    ReDim Filtros(3, 0)

    'Inicializa o número de Ordens e o vetor de ordens
    ConTelaOrdem = 1
    ReDim Ordens(1) As Integer
    Ordens(1) = 1
End Sub


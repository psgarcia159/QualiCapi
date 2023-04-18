Attribute VB_Name = "AcessoDados"
Option Explicit

'*Definição de constantes para os tipos de recordset ou resultset a serem utilizados
Global Const Dinamico = "Dinamico"
Global Const Direto = "Direto"
Global Const Estatico = "Estatico"

'*Indicam os procedimentos de inicialização do Banco de Dados
Global Conexao As New ADODB.Connection          ' Representa a conexão com o banco de dados.
Global ConexaoRelatorio As New ADODB.Connection ' Representa a conexão com o banco de dados para Relatórios.
Global ConexaoMfin As New ADODB.Connection      ' Representa a conexão com o banco de dados.
Global ConexaoPadroes As New ADODB.Connection   ' Representa a conexão com o banco de dados.
Global Comando As New ADODB.Command             ' Representa a conexão com o banco, podendo ser parametrizado

Global ConexaoADO As New ADODB.Connection      ' Representa a conexão com o banco de dados.
Global ConexaoADOPadrao As New ADODB.Connection      ' Representa a conexão com o banco de dados.

Global TipoAcesso As String  ' Indica o tipo de acesso a dados que será utilizado ex: DAO, RDO
Global NomeSgbd As String     ' Indica o nome do banco de dados a ser utilizado ex: Access

Global TipoBloqueio As Integer         ' Indica o tipo de bloqueio a ser utilizado ex: Pessimista ou Otimista
                                       ' O Bloqueio Pessimista bloqueia o registro a partir do momento que
                                       ' é realizado o edit e libera após a realização do update.
                                       ' O Bloqueio Otimista bloqueia o registro apenas no momento que está
                                       ' sendo realizado o update.

Public Sub subConectarControleDadosNV(ControleDados As Object, Sql As String, TipoRecordSet As String)
'Conecta o controle de dados
    Dim ObjRecordset As ADODB.Recordset
    
    Set ObjRecordset = New ADODB.Recordset
    
    'Seleciona o tipo de acesso
    
    Select Case TipoRecordSet
        Case "Dinamico"  ' Recordset que pode ser alterado
            ObjRecordset.Open funTrataSql(Sql), Conexao, adOpenDynamic, TipoBloqueio, adCmdText
        Case "Direto"       ' Recordset de leitura com movimentacao em apenas uma direcao
            ObjRecordset.Open funTrataSql(Sql), Conexao, adOpenForwardOnly, adLockReadOnly, adCmdText
        Case "Estatico"    ' Recordset de leitura com livre movimentacao
            ObjRecordset.CursorLocation = adUseClient
            If (UCase(XGT_NomeAplic) = "QUALIORC") Then
                ObjRecordset.Open funTrataSql(Sql), Conexao, adOpenDynamic, TipoBloqueio, adCmdText
            Else
                ObjRecordset.Open funTrataSql(Sql), Conexao, adOpenStatic, adLockReadOnly, adCmdText
            End If
    End Select
    Set ControleDados.Recordset = ObjRecordset
End Sub

Public Sub subConectarControleDados(ControleDados As Object, Sql As String, TipoRecordSet As String)
'Conecta o controle de dados
    'Seleciona o tipo de acesso
    Select Case TipoRecordSet
        Case "Dinamico"  ' Recordset que pode ser alterado
            ControleDados.CursorType = adOpenDynamic
            ControleDados.LockType = TipoBloqueio
        Case "Direto"       ' Recordset de leitura com movimentacao em apenas uma direcao
            ControleDados.CursorType = adOpenForwardOnly
            ControleDados.LockType = adLockReadOnly
        Case "Estatico"    ' Recordset de leitura com livre movimentacao
            ControleDados.CursorType = adOpenStatic
            ControleDados.LockType = adLockReadOnly
    End Select
    ControleDados.ConnectionString = Conexao.ConnectionString
    
    ControleDados.RecordSource = funTrataSql(Sql)
    ControleDados.Refresh
End Sub

Public Sub subLocalizaData(DataGeral As Object, NomeCampo As String, ValorCampo As Variant)
'Localizar um registro no controle de dados
    DataGeral.Recordset.MoveFirst
    Do While Not DataGeral.Recordset.EOF
       If DataGeral.Recordset(NomeCampo) = ValorCampo Then
            Exit Do
       End If
       DataGeral.Recordset.MoveNext
    Loop
End Sub

Public Sub subrecarregadados(ControleDados As Object, TabelaAtiva As String, _
  OrdemTabela As String, CamposObrigatorios As String, FiltroOriginal As String)
'Carrega os Valores da Tabela Atual inicializando o seu controle de dados.
    'Atribui os parâmetros às variáveis globais
    Ordem = OrdemTabela
    Campos_Obrig = CamposObrigatorios
    FiltroOrig = funTrataSql(FiltroOriginal)
    FiltroAtual = FiltroOrig

    'Inicializa o Controle de Dados.
    If Ordem <> "" Then
        subConectarControleDados ControleDados, FiltroOrig + " ORDER BY " + Ordem, Estatico
    Else
        subConectarControleDados ControleDados, FiltroOrig, Estatico
    End If

    Quantid = 0 'Zera a quantidade de registros da tabela
    
    If Not ControleDados.Recordset.EOF Then
       ControleDados.Recordset.MoveLast
       Quantid = ControleDados.Recordset.RecordCount 'Informa a quantidade de registros da tabela.
    End If

    ' Aponta para o primeiro registro da tabela
    If Quantid <> 0 Then
       ControleDados.Recordset.MoveFirst
    End If
End Sub

Public Sub subrecarregadadosNV(ControleDados As Object, TabelaAtiva As String, _
  OrdemTabela As String, CamposObrigatorios As String, FiltroOriginal As String)
    'Carrega os Valores da Tabela Atual inicializando o seu controle de dados.
    Dim Cont As Long
    
    'Atribui os parâmetros às variáveis globais
    Ordem = OrdemTabela
    Campos_Obrig = CamposObrigatorios
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
   MdiPrincipal.BarraStatus.Panels(3).Text = Str(Cont) + "/" + Str(Quantid)

    ' Aponta para o primeiro registro da tabela
    If Cont <> 0 Then
       ControleDados.Recordset.MoveFirst
    End If
End Sub

Public Sub subCarregaDados(ControleDados As Object, TabelaAtiva As String, _
  OrdemTabela As String, CamposObrigatorios As String, FiltroOriginal As String)
'Carrega os Valores da Tabela Atual inicializando o seu controle de dados.
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
        subConectarControleDados ControleDados, FiltroOrig + " ORDER BY " + Ordem, Estatico
    Else
        subConectarControleDados ControleDados, FiltroOrig, Estatico
    End If

    Quantid = 0 'Zera a quantidade de registros da tabela
    
    If Not ControleDados.Recordset.EOF Then
       ControleDados.Recordset.MoveLast
       Quantid = ControleDados.Recordset.RecordCount 'Informa a quantidade de registros da tabela.
    End If

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

Public Sub subCarregaDadosNV(ControleDados As Object, TabelaAtiva As String, _
  OrdemTabela As String, CamposObrigatorios As String, FiltroOriginal As String)
'Carrega os Valores da Tabela Atual inicializando o seu controle de dados.
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
    MdiPrincipal.BarraStatus.Panels(3).Text = Str(Quantid) + "/" + Str(Quantid)

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

Public Sub SubQOpenRecordset(ObjRecordset As ADODB.Recordset, Sql As String, TipoRecordSet As String)
    ' Abre um conjunto de registros de acordo com o Tipo de Acesso
    Set ObjRecordset = New ADODB.Recordset
    Select Case TipoRecordSet
        Case "Dinamico"  ' Recordset que pode ser alterado
            'Se usar o CursorType = adOpenDynamic, o propriedade Bookmark não é setada, consequentemente não podemos pegar a chave para gravação de tabelas relacionadas
            ObjRecordset.CursorType = adOpenKeyset
            ObjRecordset.LockType = TipoBloqueio
            ObjRecordset.Open funTrataSql(Sql), Conexao, , TipoBloqueio, adCmdText
        Case "Direto"       ' Recordset de leitura com movimentacao em apenas uma direcao
            ObjRecordset.Open funTrataSql(Sql), Conexao, adOpenForwardOnly, adLockReadOnly, adCmdText
        Case "Estatico"    ' Recordset de leitura com livre movimentacao
            ObjRecordset.CursorLocation = adUseClient
            ObjRecordset.Open funTrataSql(Sql), Conexao, adOpenStatic, adLockReadOnly, adCmdText
    End Select
End Sub

Public Sub subQOpenRecordset2(ObjRecordset As ADODB.Recordset, Sql As String, TipoRecordSet As String)
    'Abre um conjunto de registros de acordo com o Tipo de Acesso
    Set ObjRecordset = New ADODB.Recordset
    Select Case TipoRecordSet
        Case "Dinamico"  ' Recordset que pode ser alterado
            'Se usar o CursorType = adOpenDynamic, o propriedade Bookmark não é setada, consequentemente não podemos pegar a chave para gravação de tabelas relacionadas
            ObjRecordset.CursorType = adOpenKeyset
            ObjRecordset.LockType = TipoBloqueio
            ObjRecordset.Open funTrataSql(Sql), ConexaoMfin, , TipoBloqueio, adCmdText
        Case "Direto"       ' Recordset de leitura com movimentacao em apenas uma direcao
            ObjRecordset.Open funTrataSql(Sql), ConexaoMfin, adOpenForwardOnly, adLockReadOnly, adCmdText
        Case "Estatico"    ' Recordset de leitura com livre movimentacao
            ObjRecordset.Open funTrataSql(Sql), ConexaoMfin, adOpenStatic, adLockReadOnly, adCmdText
    End Select
End Sub

Public Function funAbreConexao(XLT_MSGERRO As String) As Boolean
'Abre a conexao com o banco
    Dim XLT_STRINGCONEXAO   As String * 254
    Dim XLT_TIPOBANCO       As String * 254
    Dim XLT_NOMEBANCO       As String * 254
    Dim XLT_USUARIO         As String * 254
    Dim XLT_SENHA           As String * 254
    Dim ErroOdbc            As String
    
    On Error GoTo RotuloErro
    
    Screen.MousePointer = vbHourglass
    'Indica o banco de dados a ser utilizado
    If GetPrivateProfileString("BancoDeDados", "TipoBanco", "", XLT_TIPOBANCO, 255, App.Path + "\QualiAdmFin.INI") = 0 Then
        MsgBox "Problemas na Leitura do Arquivo de configuração", vbCritical, "Configuração"
    End If
    
    'PEGA O ENDEREÇO DO BANCO
    GetPrivateProfileString "BancoDeDados", "Endereco", "", XLT_STRINGCONEXAO, 255, App.Path + "\QualiAdmFin.INI"
    
    'PEGA O NOME DO BANCO
    GetPrivateProfileString "BancoDeDados", "NomeBanco", "", XLT_NOMEBANCO, 255, App.Path + "\QualiAdmFin.INI"
    
    'PEGA O USUARIO
    GetPrivateProfileString "BancoDeDados", "Usuario", "", XLT_USUARIO, 255, App.Path + "\QualiAdmFin.INI"
    
    If IsNull(XLT_USUARIO) Or IsEmpty(XLT_USUARIO) Or XLT_USUARIO = "" Then
        XLT_USUARIO = "qualiadmfin"
    End If
        
    'PEGA A SENHA
    GetPrivateProfileString "BancoDeDados", "Senha", "", XLT_SENHA, 255, App.Path + "\QualiAdmFin.INI"
    
    If IsNull(XLT_SENHA) Or IsEmpty(XLT_SENHA) Or XLT_SENHA = "" Then
        XLT_SENHA = "qd"
    End If
    
    'PEGA MÁSCARA DO CÓDIGO REDUZIDO
    GetPrivateProfileString "BancoDeDados", "MskCodReduzido", "", XGT_MSKCODREDUZIDO, 255, App.Path + "\QualiAdmFin.INI"
    
    'PEGA MÁSCARA DO CÓDIGO DA CATEGORIA
    GetPrivateProfileString "BancoDeDados", "Mascara", "", XGT_MSKCATEGORIA, 255, App.Path + "\QualiAdmFin.INI"
    
    'PEGA CÓDIGO REDUZIDO DA CONTA DE GLOSA A RECUPERAR
    XGT_CONTAGLOSA = ""
    GetPrivateProfileString "BancoDeDados", "ContaGlosa", "", XGT_CONTAGLOSA, 7, App.Path + "\QualiAdmFin.INI"
    XGT_CONTAGLOSA = Trim(FunStrArqIni(XGT_CONTAGLOSA))

    'Patrícia - PEGA O NÚMERO DE DIAS PARA SER UTILIZADO NO CÁLCULO DO NÚMERO DE REGISTROS A SEREM MOSTRADOS NAS TELAS DE DUPLICATAS E CONTAS A PAGAR\RECEBER.
    XGT_NUMDIAS = ""
    GetPrivateProfileString "BancoDeDados", "NumDias", "", XGT_NUMDIAS, 7, App.Path + "\QualiAdmFin.INI"
    XGT_NUMDIAS = Trim(FunStrArqIni(XGT_NUMDIAS))

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
                ";Persist Security Info=True;Initial Catalog=" & FunStrArqIni(XLT_NOMEBANCO) & _
                ";User Id=" & FunStrArqIni(XLT_USUARIO) & ";Password=" & FunStrArqIni(XLT_SENHA) & "; "
        Conexao.CommandTimeout = 1200
        Conexao.Open XGT_SQL

    End Select
    Screen.MousePointer = vbDefault
    funAbreConexao = True
Exit Function

RotuloErro:
  XLT_MSGERRO = Err.Number & " - " & Err.Description
  funAbreConexao = False
  Screen.MousePointer = vbDefault

End Function

Public Function funAbreConexaoRelatorio() As Boolean
'Abre a conexao com o banco
    Dim XLT_STRINGCONEXAO   As String * 254
    Dim XLT_TIPOBANCO       As String * 254
    Dim XLT_NOMEBANCO       As String * 254
    Dim XLT_USUARIO         As String * 254
    Dim XLT_SENHA           As String * 254
    Dim ErroOdbc            As String
    
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
    
    'PEGA O USUARIO
    GetPrivateProfileString "BancoDeDados", "Usuario", "", XLT_USUARIO, 255, App.Path + "\QualiAdmFin.INI"
    
    If IsNull(XLT_USUARIO) Or IsEmpty(XLT_USUARIO) Or XLT_USUARIO = "" Then
        XLT_USUARIO = "qualiadmfin"
    End If
        
    'PEGA A SENHA
    GetPrivateProfileString "BancoDeDados", "Senha", "", XLT_SENHA, 255, App.Path + "\QualiAdmFin.INI"
    
    If IsNull(XLT_SENHA) Or IsEmpty(XLT_SENHA) Or XLT_SENHA = "" Then
        XLT_SENHA = "qd"
    End If
    
    Select Case Val(XLT_TIPOBANCO)
        Case 1
            NomeSgbd = "Access"
            'Leitura da String de conexao do banco
            ConexaoRelatorio.Provider = "Microsoft.Jet.OLEDB.4.0"
            ConexaoRelatorio.CommandTimeout = 0
            ConexaoRelatorio.Open FunStrArqIni(XLT_STRINGCONEXAO), "admin", ""
        Case 2
            NomeSgbd = "Sql Server"
            ConexaoRelatorio.Provider = "sqloledb"
            ConexaoRelatorio.CommandTimeout = 1200
            ConexaoRelatorio.CursorLocation = adUseClient
            XGT_SQL = "Data Source=" & FunStrArqIni(XLT_STRINGCONEXAO) & _
                      ";Persist Security Info=True;Initial Catalog=" & FunStrArqIni(XLT_NOMEBANCO) & _
                      ";User Id=" & FunStrArqIni(XLT_USUARIO) & ";Password=" & FunStrArqIni(XLT_SENHA) & ";OLE DB Services = -2; "
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

Public Sub subQOpenRecordsetPadroes(ObjRecordset As ADODB.Recordset, Sql As String, TipoRecordSet As String)
' Abre um conjunto de registros de acordo com o Tipo de Acesso
    Set ObjRecordset = New ADODB.Recordset
    Select Case TipoRecordSet
        Case "Dinamico"  ' Recordset que pode ser alterado
            'Se usar o CursorType = adOpenDynamic, o propriedade Bookmark não é setada, consequentemente não podemos pegar a chave para gravação de tabelas relacionadas
            ObjRecordset.CursorType = adOpenKeyset
            ObjRecordset.LockType = TipoBloqueio
            ObjRecordset.Open funTrataSql(Sql), ConexaoPadroes, , TipoBloqueio, adCmdText
        Case "Direto"       ' Recordset de leitura com movimentacao em apenas uma direcao
            ObjRecordset.Open funTrataSql(Sql), ConexaoPadroes, adOpenForwardOnly, adLockReadOnly, adCmdText
        Case "Estatico"    ' Recordset de leitura com livre movimentacao
            ObjRecordset.Open funTrataSql(Sql), ConexaoPadroes, adOpenStatic, adLockReadOnly, adCmdText
    End Select
End Sub

Public Property Let letQRecordsetFields(Tabela As Object, Field As String, ByVal vNewValue As Variant)
    
    ' Altera o conteúdo do campo em um conjunto de registros de acordo com o Tipo de Acesso, atribuindo ao campo o valor contido em vNewValue
    Tabela(Field) = vNewValue
    
End Property

Public Property Get getQRecordsetFields(Tabela As Object, Field As String) As Variant
    
    ' Retorna a conteúdo do campo em um conjunto de registros de acordo com o Tipo de Acesso
    getQRecordsetFields = Tabela(Field)
    
End Property

Function funTipoCampo(ControleDados As Object, Grid As TrueOleDBGrid70.Columns, coluna As Integer) As String
' Retorna o tipo do campo usado em uma coluna de um grid.
    Dim Tipo As Variant
    ' Verifica o tipo de acesso e utiliza a propriedade type para saber qual o tipo daquela coluna do grid.
    Tipo = ControleDados.Recordset.Fields(Grid(coluna).DataField).Type
    
    ' Verifica o tipo encontrado para classificar de acordo com os padrões pré-definidos.
    Select Case Tipo
       Case adBoolean
            funTipoCampo = "B"  'campo Booleano
       Case adTinyInt
            funTipoCampo = "Y"  'campo Byte
       Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, adSingle, adSmallInt, adNumeric
            funTipoCampo = "N"  'campo Numerico
       Case adDate, adDBDate, adDBTime, adDBTimeStamp
            funTipoCampo = "D"  'campo Data
       Case adChar, adVarChar
            funTipoCampo = "T"  'campo Texto
       Case adLongVarChar
            funTipoCampo = "M"  'campo Memo
    End Select
End Function



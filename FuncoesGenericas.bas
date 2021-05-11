Attribute VB_Name = "FuncoesGenericas"
Option Explicit
Global TabelaAtiva As String 'Informa o nome da Tabela do Banco de Dados Ativa

Global Filtros() As String ' Array para armazenar os Filtros da Tabela
Global ContFil As Integer ' Contador que armazena o número de filtros ativos
Global Filtrou As Boolean ' Informa se Algum campo foi filtrado
Global Ordem As String ' Nome do Campo que tem a chave da Tabela ativa
Global Campos_Obrig As String ' Lista dos campos que são obrigatórios na tabela
Global Formulario As Form ' Indica o Formulario Ativo
Global TipoVar As String ' Tipo do campo a ser filtrado
Global FiltroOrig As String ' Indica os filtros determinados pelo próprio sistema
Global FiltroAtual As String ' Indica o filtro atual da tabela corrente
Global filtra_especial As Integer ' Indica se a tabela utilizada possui filtros pré-definidos pelo sistema em FiltroOrig
                                               
Global BarraFerramentasPressionada As Boolean
Global VariasVezes As Boolean    'Indica se vai executar a insercao na TelaPadrao varias ou uma vez
                                               
Global Tem_Filtro As Boolean ' Informa se já foi relizado um filtro na tabela pelo usuário
Global ConTelaOrdem As Integer ' Indica o número de chaves utilizados para ordenação
Global Ordens() As Integer    ' Vetor que mantem a ordem dos campos
Global Vetor() As Variant ' Vetor para auxílio na criação da tabela padrao
Global TabelaUltimos As String ' Indica o nome da tabela de últimos que está sendo utilizada
Global Registro(100) As Variant ' Vetor para guardar cópias de registros
Global Quantid As Long  ' Quantidade de registros da tabela

Global XGB_RESPOSTA As Boolean       'Indica se o formulário modal foi fechado com OK (True) ou Cancelar (False)
Global XGI_GRAFICO As Integer        'Número do tipo do gráfico selecionado
Global XGI_GRAFICO_NRSUMS As Integer 'Indica o número de somatórios que o gráfico pode suportar

'* Funcoes de manipulacao de arquivos INI
' Leitura de arquivo INI
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
' Gravação em arquivo INI
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Function funFormataCCusto(CCusto As String) As String
  Dim XLI_X As Integer
  Dim Custo As String
  Dim XLI_CONT As Integer
  
  If pPCRnivel = 1 Then
    Custo = CCusto
    For XLI_X = (Len(pPCRmascara) - Len(Custo)) To 1 Step -1
      If Mid(pPCRmascara, XLI_X, 1) = "." Then
        Custo = "." & Custo
      Else
        Custo = "0" & Custo
      End If
    Next
  Else
    For XLI_X = 1 To Len(pPCRmascara)
        If Mid(pPCRmascara, XLI_X, 1) <> "." Then
            XLI_CONT = XLI_CONT + 1
            Custo = Custo + Mid(CCusto, XLI_CONT, 1)
        Else
            Custo = Custo + "."
        End If
    Next
  End If
  funFormataCCusto = Custo
End Function

Public Function funChecaCategoria(XLT_CODESTRUT As String, _
                                  XLT_CONTAESTOQUE As String, _
                                  XLT_CONTAAPLICACAO As String, _
                                  XLT_TIPOCODIGO As String) As Boolean
    Dim XLI_X As Integer
    Dim XLT_CATEGORIA As String
    Dim XLT_CARACTER As String
    Dim XLO_CATEGORIA   As ADODB.Recordset
    Dim XLB_ACHOU As Boolean
    Dim XLB_EXISTE As Boolean
    
    XLB_EXISTE = False
    If XLT_TIPOCODIGO = "E" Then
        XGT_SQL = "SELECT ctpl_cd_ContaEstoque, ctpl_cd_ContaAplicacao, cate_cd_Estruturado FROM ConsGENCategorias " & _
                    "WHERE cate_cd_Estruturado='" & XLT_CODESTRUT & "' AND " & _
                          "empr_cd_Empresa=" & PCodEmpresa
    Else
        XGT_SQL = "SELECT ctpl_cd_ContaEstoque, ctpl_cd_ContaAplicacao, cate_cd_Estruturado FROM ConsGENCategorias " & _
                    "WHERE cate_cd_Categoria = " & XLT_CODESTRUT & " AND " & _
                          "empr_cd_Empresa=" & PCodEmpresa
    End If
    SubQOpenRecordset XLO_CATEGORIA, XGT_SQL, Estatico
    XLT_CODESTRUT = FunNulo(XLO_CATEGORIA!cate_cd_estruturado)
    If IsNull(XLO_CATEGORIA!ctpl_cd_ContaEstoque) Or IsNull(XLO_CATEGORIA!ctpl_cd_ContaAplicacao) Then
        'Achar o nivel anterior da categoria que se está checando
        XLT_CATEGORIA = XLT_CODESTRUT
        XLB_ACHOU = False
        For XLI_X = Len(XLT_CODESTRUT) To 1 Step -1
            XLT_CARACTER = Mid(XLT_CODESTRUT, XLI_X, 1)
            If XLT_CARACTER <> " " And XLT_CARACTER <> "." And Not XLB_ACHOU Then
                XLT_CATEGORIA = Left(XLT_CATEGORIA, XLI_X - 1) & " " & Right(XLT_CATEGORIA, (Len(XLT_CODESTRUT) - XLI_X))
                If XLI_X - 1 = 0 Then
                    XLB_ACHOU = True
                ElseIf Mid(XLT_CODESTRUT, XLI_X - 1, 1) = "." Then
                    XLB_ACHOU = True
                End If
            ElseIf XLB_ACHOU = True Then
                'Procurar Conta neste nível
                XGT_SQL = "SELECT ctpl_cd_ContaEstoque, ctpl_cd_ContaAplicacao FROM ConsGENCategorias " & _
                            "WHERE cate_cd_Estruturado='" & XLT_CATEGORIA & "' AND " & _
                                  "empr_cd_Empresa=" & PCodEmpresa
                SubQOpenRecordset XLO_CATEGORIA, XGT_SQL, Estatico
                'Quando base caixa é obrigatório ter conta de aplicação
                'Quando base prov e caixa é obrigatório ter as duas contas
                If XGT_TIPOEXPORTACAO = "PC" And IsNull(XLO_CATEGORIA!ctpl_cd_ContaEstoque) Or IsNull(XLO_CATEGORIA!ctpl_cd_ContaAplicacao) Then
                    XLB_ACHOU = False
                ElseIf XGT_TIPOEXPORTACAO = "CX" And IsNull(XLO_CATEGORIA!ctpl_cd_ContaAplicacao) Then
                    XLB_ACHOU = False
                Else
                    'Setar valores da Conta Aplicação e Conta Estoque
                    XLT_CONTAESTOQUE = FunNulo(XLO_CATEGORIA!ctpl_cd_ContaEstoque)
                    XLT_CONTAAPLICACAO = XLO_CATEGORIA!ctpl_cd_ContaAplicacao
                    XLB_EXISTE = True
                    
                    Exit For
                End If

            End If
        Next
    Else
        'Setar valores da Conta Aplicação e Conta Estoque
        XLT_CONTAESTOQUE = XLO_CATEGORIA!ctpl_cd_ContaEstoque
        XLT_CONTAAPLICACAO = XLO_CATEGORIA!ctpl_cd_ContaAplicacao
        XLB_EXISTE = True
    End If
    If Not XLB_EXISTE Then
        MsgBox "Este Grupo de Categoria não possui Conta de Estoque ou Aplicação cadastradas para o seu nível ou qualquer nível superior ao seu, Favor cadastrar.", vbCritical, "ATENÇÃO"
    End If
    XLO_CATEGORIA.Close
    Set XLO_CATEGORIA = Nothing
    funChecaCategoria = XLB_EXISTE
End Function

' Realiza a chamada do formulário de Filtro, de acordo com o tipo de dados da coluna do grid.
Function FunExecutaFiltroUnbound(VetorCampos As Variant, Vetor As XArray, Grid As TDBGrid, XForm As Form, Posicaox As Single, XLT_SQL As String)
    Dim CurCol As Integer 'Receberá o numero da coluna clicada
    
    CurCol = XForm.TDBGrid1.ColContaining(Posicaox)   'A variável "curcol" contém o numero da coluna clicada
    
    If CurCol >= 0 Then 'Corresponde a um campo qualquer da tabela
       XForm.TDBGrid1.Col = CurCol
       
       If XForm.TDBGrid1.Columns(CurCol).HeadFont.Italic Then
          MsgBox "Campo já se encontra filtrado", 32, "ERRO"
          Exit Function
       End If
       
       XGT_CAMPOFILTRO = XForm.TDBGrid1.Columns(CurCol).FooterText
       
       If XGT_CAMPOFILTRO = "N" Or XGT_CAMPOFILTRO = "D" Or XGT_CAMPOFILTRO = "Y" Then
           FiltroNumericoUnBound.Label2.Caption = XForm.TDBGrid1.Columns(CurCol).Caption
           FiltroNumericoUnBound.Grid() = XForm.TDBGrid1.Columns
           FiltroNumericoUnBound.ColunaGrid() = CurCol
           FiltroNumericoUnBound.Show 1
       Else
           FiltroCaracterUnBound.Label2.Caption = XForm.TDBGrid1.Columns(CurCol).Caption
           FiltroCaracterUnBound.Grid() = XForm.TDBGrid1.Columns
           FiltroCaracterUnBound.ColunaGrid() = CurCol
           FiltroCaracterUnBound.Show 1
       End If

       If Filtrou = True Then
         subCarregaVetor FiltroAtual, VetorCampos, Vetor, Grid
       End If
    End If

End Function

Function funMesFechado(XLD_DATALANC As Date) As String
    Dim XLT_SQL As String        'Armazena SELECT
    Dim XLO_CONTABIL As ADODB.Recordset 'Recordset para abrir tabela de fechamento
    
    XLT_SQL = "SELECT * FROM Fechamento " & _
                " WHERE month(fech_dt_mes) = " & Month(XLD_DATALANC) & "" & _
                " AND year(fech_dt_mes) = " & Year(XLD_DATALANC) & "" & _
                " AND empr_cd_empresa = " & PCodEmpresa & ""
    SubQOpenRecordset XLO_CONTABIL, XLT_SQL, Estatico
    
    If Not XLO_CONTABIL.EOF Then
        funMesFechado = XLO_CONTABIL!fech_tx_status
        XLO_CONTABIL.Close
        Set XLO_CONTABIL = Nothing
    Else
        XLO_CONTABIL.Close
        Set XLO_CONTABIL = Nothing
        'Se não existir o mês, verificar se ele é posterior ao último fechado
        XLT_SQL = "SELECT * FROM Fechamento " & _
                " WHERE fech_tx_Status='F' " & _
                " AND empr_cd_empresa = " & PCodEmpresa & " ORDER BY fech_dt_mes desc"
        SubQOpenRecordset XLO_CONTABIL, XLT_SQL, Estatico
        
        If Not XLO_CONTABIL.EOF Then
            If (Month(XLD_DATALANC) > Month(XLO_CONTABIL!fech_dt_mes) And Year(XLD_DATALANC) = Year(XLO_CONTABIL!fech_dt_mes)) _
                Or Year(XLD_DATALANC) > Year(XLO_CONTABIL!fech_dt_mes) Then
                Conexao.Execute "INSERT INTO Fechamento(fech_dt_Mes,empr_cd_empresa,fech_tx_Status) " & _
                                "VALUES(" & FunNuloData(Format(XLD_DATALANC, "\01/MM/yyyy"), NomeSgbd) & "," & _
                                "" & PCodEmpresa & ",'A')"
                funMesFechado = "A"
            Else
                funMesFechado = "N" 'Não pode cadastrar, pois existe um mês posterior fechado
            End If
        Else 'se não existe nenhum fechado criar o mês
            Conexao.Execute "INSERT INTO Fechamento(fech_dt_Mes,empr_cd_empresa,fech_tx_Status) " & _
                            "VALUES(" & FunNuloData(Format(XLD_DATALANC, "\01/MM/yyyy"), NomeSgbd) & "," & _
                            "" & PCodEmpresa & ",'A')"
            funMesFechado = "A"
        End If
        XLO_CONTABIL.Close
        Set XLO_CONTABIL = Nothing
    End If
End Function

'RETORNA A UNIDADE DE CONTROLE PADRAO
Public Function funUnPadrao(XLL_CODMATERIAL)
    
    Dim XLO_UNIDADE     As ADODB.Recordset
    Dim XLT_UNPADRAO    As String
        
    XGT_SQL = "SELECT * FROM UnidadesFornecimento " & _
              " WHERE mate_cd_Reduzido = " & XLL_CODMATERIAL & _
              " AND unfo_bl_UnidadePadrao='S'"
    SubQOpenRecordset XLO_UNIDADE, XGT_SQL, Estatico
    
    'PEGA AS INFORMAÇÕES DA UNIDADE PADRAO
    If Not XLO_UNIDADE.EOF Then
        XLT_UNPADRAO = XLO_UNIDADE.Fields("unfo_cd_UnidadeFornecimento")
    Else
        XLT_UNPADRAO = ""
    End If
    
    funUnPadrao = XLT_UNPADRAO
    XLO_UNIDADE.Close
    Set XLO_UNIDADE = Nothing
End Function

'RETORNA A UNIDADE DE CONTROLE PADRAO
Public Function FunUnForn(XLL_CODMATERIAL)
    
    Dim XLO_UNIDADE     As ADODB.Recordset
    Dim XLT_UNFORNC     As String
        
    XGT_SQL = "SELECT * FROM UnidadesFornecimento " & _
              " WHERE mate_cd_Reduzido = " & XLL_CODMATERIAL & _
              " AND unfo_bl_UnidadeFornecimento='S'"
    SubQOpenRecordset XLO_UNIDADE, XGT_SQL, Estatico
    
    'PEGA AS INFORMAÇÕES DA UNIDADE PADRAO
    If Not XLO_UNIDADE.EOF Then
        XLT_UNFORNC = XLO_UNIDADE.Fields("unfo_cd_UnidadeFornecimento")
    Else
        XLT_UNFORNC = ""
    End If
    
    FunUnForn = XLT_UNFORNC
    XLO_UNIDADE.Close
End Function

'FORMATA O CODIGO REDUZIDO DO MATERIAL
Public Function FunCodigoReduzido(ByVal XLT_CODIGO As String) As String
    Dim XLI_POS As Integer
    Dim XLT_MAT As String, XLT_DIG As String
    
    XLI_POS = InStr(XLT_CODIGO, "-")
    XLT_MAT = Trim(Mid(XLT_CODIGO, 1, XLI_POS - 1))
    XLT_DIG = Trim(Mid(XLT_CODIGO, XLI_POS + 1))
    
    FunCodigoReduzido = Format(XLT_MAT, "00000") & "-" & XLT_DIG
End Function

'retorna String limpa do arquivo INI
Function FunStrArqIni(XLT_TEXTO As String) As String
    Dim XLI_INDICE As Integer
    Dim XLT_RESULTADO As String
    
    XLT_RESULTADO = ""
    
    For XLI_INDICE = 1 To Len(XLT_TEXTO)
        If Asc(Mid(XLT_TEXTO, XLI_INDICE, 1)) <> 0 Then
            XLT_RESULTADO = XLT_RESULTADO + Mid(XLT_TEXTO, XLI_INDICE, 1)
        End If
    Next
    
    FunStrArqIni = XLT_RESULTADO
End Function

'TRATA ASPA SIMPLES NO TEXTO
Public Function FunTrataString(XVariavel As String) As String
    Dim Cont As Integer
    Dim XResult As String
    
    Cont = 1
    While (Cont <= Len(Trim(XVariavel)))
        If (Mid(XVariavel, Cont, 1) <> "'") Then
            XResult = XResult + Mid(XVariavel, Cont, 1)
        Else
            XResult = XResult + "''"
        End If

        Cont = Cont + 1
    Wend
    
    FunTrataString = XResult
End Function

'TRATA TIRA ASPAS DO TEXTO
Public Function FunTiraAspas(XVariavel As String) As String
    Dim Cont As Integer
    Dim XResult As String
    
    Cont = 1
    While (Cont <= Len(Trim(XVariavel)))
        If Asc(Mid(XVariavel, Cont, 1)) <> 39 And Asc(Mid(XVariavel, Cont, 1)) <> 34 Then
            XResult = XResult + Mid(XVariavel, Cont, 1)
        End If

        Cont = Cont + 1
    Wend
    
    FunTiraAspas = XResult
End Function

Public Sub subCentralizaFormulario(XLO_FORMULARIO As Form)
  XLO_FORMULARIO.Left = Int((ResolucaoX - XLO_FORMULARIO.ScaleWidth) / 2)
  XLO_FORMULARIO.Top = Int((Resolucaoy - XLO_FORMULARIO.ScaleHeight) / 2)
End Sub

Public Sub subCentralizaPainel(XLO_PAINEL As SSPanel, XLO_FORMULARIO As Form)
  XLO_PAINEL.Left = Int((XLO_FORMULARIO.ScaleWidth - XLO_PAINEL.Width) / 2)
  XLO_PAINEL.Top = Int((XLO_FORMULARIO.ScaleHeight - XLO_PAINEL.Height) / 2)
End Sub

Public Function TrataString(XVariavel As String) As String
    Dim Cont As Integer
    Dim XResult As String
    
    Cont = 1
    While (Cont <= Len(Trim(XVariavel)))
        If (Mid(XVariavel, Cont, 1) <> "'") Then
            XResult = XResult + Mid(XVariavel, Cont, 1)
        Else
            XResult = XResult + "''"
        End If

        Cont = Cont + 1
    Wend
    
    TrataString = XResult
    
End Function

'Transforma uma string formatada com "," para uma formatacao com "." ex: "10,2" ficará "10.2"
Public Function FunTrataFloat(ByVal XVariavel As String) As String
    If XVariavel = "" Then
        FunTrataFloat = "0"
    ElseIf Len(XVariavel) > 25 Then
        XVariavel = Mid$(XVariavel, 33)
        FunTrataFloat = Str(CDbl(XVariavel))
        Else
        FunTrataFloat = Str(CDbl(XVariavel))
    End If
End Function

' Verifica se a o recordset / resultset contém dados
Public Function FunTabelaVazia(ControleDados As Object) As Boolean
    If ((ControleDados.Recordset.EOF) And (ControleDados.Recordset.BOF)) Then
       FunTabelaVazia = True
    Else
       FunTabelaVazia = False
    End If
End Function

' Converte um SQL para os padrões pré-definidos
Function FunArruma_SQL(Tabela As String) As String
   Dim a, Posicao, Intervalo As Integer
   Dim DataStr, StrSQL As String
    
   StrSQL = ""
   
   For a = 1 To ContFil
       
       Posicao = InStr(1, Filtros(1, a), ".")
       
       If Posicao <> 0 Then
          Tabela = ""
       Else
          If InStr(1, Tabela, ".") = 0 Then
            Tabela = Tabela + "."
          End If
       End If
       
       If Filtros(3, a) = "1" Then
          
          If a = 1 Then
             StrSQL = "uppercase({" + Tabela + Filtros(1, a) + "}) LIKE '" + UCase$(Filtros(2, a)) + "*'"
          Else
             StrSQL = StrSQL + " AND uppercase({" + Tabela + Filtros(1, a) + "}) LIKE '" + UCase$(Filtros(2, a)) + "*'"
          End If
       
       Else
          DataStr = Filtros(2, a)
          Posicao = InStr(1, DataStr, "AND") 'Guarda a posição do AND
          
          If Posicao <> 0 Then    'Tira o AND
             DataStr = Left$(DataStr, Posicao - 1) + Right$(DataStr, Len(DataStr) - Posicao - 3)
             Posicao = InStr(1, DataStr, Filtros(1, a))
             If Posicao <> 0 Then
                DataStr = Left$(DataStr, Posicao - 1) + Right$(DataStr, Len(DataStr) - Posicao - Len(Filtros(1, a)))
             End If
          End If
          
          Posicao = InStr(1, DataStr, "#")
          If Posicao = 0 Then
             StrSQL = "{" + Tabela + Filtros(1, a) + "} " + DataStr
          Else
                Do While Posicao <> 0     'Tira os # da Data
                    If Posicao = Len(DataStr) Then
                       DataStr = Left$(DataStr, Len(DataStr) - 1)
                       Exit Do
                    Else
                       DataStr = Left$(DataStr, Posicao - 1) + Mid$(DataStr, Posicao + 1, Len(DataStr) - Posicao)
                    End If
                    Posicao = InStr(1, DataStr, "#")
                Loop
                DataStr = Trim$(DataStr)
                If Len(Trim$(Left(DataStr, 2))) = 1 Then
                    If a = 1 Then
                       StrSQL = "{" + Tabela + Filtros(1, a) + "} " + Left$(DataStr, 1) + " Date(19" + Right$(DataStr, 2) + "," + Mid$(DataStr, 3, 2) + "," + Mid$(DataStr, 6, 2) + ")"
                    Else
                       StrSQL = StrSQL + " AND {" + Tabela + Filtros(1, a) + "} " + Left$(DataStr, 1) + " Date(19" + Right$(DataStr, 2) + "," + Mid$(DataStr, 3, 2) + "," + Mid$(DataStr, 6, 2) + ")"
                    End If
                Else
                    If a = 1 Then
                       StrSQL = "{" + Tabela + Filtros(1, a) + "} " + Left$(DataStr, 2) + " Date(19" + Mid$(DataStr, 10, 2) + "," + Mid$(DataStr, 4, 2) + "," + Mid$(DataStr, 7, 2) + ") AND {" + Tabela + Filtros(1, a) + "} " + Mid$(DataStr, 13, 2) + " Date(19" + Right$(DataStr, 2) + "," + Mid$(DataStr, 16, 2) + "," + Mid$(DataStr, 19, 2) + ")"
                    Else
                       StrSQL = StrSQL + " AND {" + Tabela + Filtros(1, a) + "} " + Left$(DataStr, 2) + " Date(19" + Mid$(DataStr, 10, 2) + "," + Mid$(DataStr, 4, 2) + "," + Mid$(DataStr, 7, 2) + ") AND {" + Tabela + Filtros(1, a) + "} " + Mid$(DataStr, 13, 2) + " Date(19" + Right$(DataStr, 2) + "," + Mid$(DataStr, 16, 2) + "," + Mid$(DataStr, 19, 2) + ")"
                    End If
                End If
          End If
       End If
   Next
    
   FunArruma_SQL = funTrataSql(StrSQL)

End Function

'Retorna o último valor da chave primária de uma tabela
Public Function FunPegaUltimo(CampoChave As String) As Long
  Dim tb_ultimos As Object
    
    SubQOpenRecordset tb_ultimos, "select * from Ultimos", Dinamico
   
    If tb_ultimos.EOF Then
       FunPegaUltimo = 1
    Else
       If IsNull(getQRecordsetFields(tb_ultimos, CampoChave)) Then
          FunPegaUltimo = 1
       Else
          FunPegaUltimo = getQRecordsetFields(tb_ultimos, CampoChave) + 1
       End If
    End If
    
    tb_ultimos.Close

End Function

'Verifica se KeyAscii é um número
Function funValidaNumero(KeyAscii As Integer) As Boolean
    funValidaNumero = False
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        funValidaNumero = True
    End If
End Function

'Verifica se KeyAscii é um número
Function FunDigito(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        FunDigito = False
    Else
        FunDigito = True
    End If
End Function

'Verifica se Datax é uma data válida
Function FunVerificaData(ByVal Datax As String) As Integer
    Dim Mes, Dia, Ano, DiaAux As Integer
    
    Datax = Trim$(Datax)
    Dia = Trim$(Left$(Datax, 2))
    Mes = Trim$(Mid$(Datax, 4, 2))
    Ano = Trim$(Mid$(Datax, 7, 2))
    
    If Len(Dia) <> 2 Or Len(Mes) <> 2 Or Len(Ano) <> 2 Then 'Se tamanho diferente de 8
       MsgBox "Formato Correto: DD/MM/AA", 16, "DATA INVÁLIDA"
       FunVerificaData = False
       Exit Function
    End If
    
    If Dia < "01" Or Dia > "31" Then
       MsgBox "Dia Inválido !", 16, "DATA INVÁLIDA"
       FunVerificaData = False
       Exit Function
    End If
    
    Select Case Mes
           Case "01", "03", "05", "07", "08", "10", "12"
                If Dia < "01" Or Dia > "31" Then
                   MsgBox "Dia deve ser de 01 a 31", 16, "DATA INVÁLIDA"
                   FunVerificaData = False
                   Exit Function
                End If
           Case "04", "06", "09", "11"
                If Dia < "01" Or Dia > "30" Then
                   MsgBox "Dia deve ser de 01 a 30", 16, "DATA INVÁLIDA"
                   FunVerificaData = False
                   Exit Function
                End If
           Case "02"
                DiaAux = 28
                'Verifica se o ano é bissexto
                If (Val(Ano) Mod 4) = 0 Then
                   DiaAux = 29
                End If
                If Dia < "01" Or Dia > Format(DiaAux, "00") Then
                   MsgBox "Dia deve ser de 01 a " + Format(DiaAux, "00"), 16, "DATA INVÁLIDA"
                   FunVerificaData = False
                   Exit Function
                End If
           Case Else
                MsgBox "Mês deve ser de 01 a 12", 16, "DATA INVÁLIDA"
                FunVerificaData = False
                Exit Function
    End Select
    
    If Ano < "00" Or Ano > "99" Then
       MsgBox "Ano Inválido !", 16, "DATA INVÁLIDA"
       FunVerificaData = False
       Exit Function
    End If
    
    FunVerificaData = True

End Function

Function FunUltDiaMes(XLI_MES As Integer, XLI_ANO As Integer) As Integer
    Dim XLI_DIA As Integer
    
    Select Case XLI_MES
           Case 1, 3, 5, 7, 8, 10, 12
                FunUltDiaMes = 31
           Case 4, 6, 9, 11
                FunUltDiaMes = 30
           Case 2
                XLI_DIA = 28
                'Verifica se o ano é bissexto
                If (Val(XLI_ANO) Mod 4) = 0 Then
                   XLI_DIA = 29
                End If
                FunUltDiaMes = XLI_DIA
    End Select
End Function

'Verifica se a tecla está entre os valores válidos para o campo de acordo com Tipo_Campo
Function FunVerifica_Tecla(Tipo_Campo As String, Tecla As Integer) As Integer
FunVerifica_Tecla = Tecla

Select Case Tipo_Campo
    Case "N"
         If Not FunDigito(Tecla) And Tecla <> 8 Then
            FunVerifica_Tecla = 0
         End If
    Case "D"
         If Not FunDigito(Tecla) And Tecla <> 47 And Tecla <> 8 Then
            FunVerifica_Tecla = 0
         End If
End Select

End Function

' Verifica se o usuário possui permisssão para acessar a rotina determinada em Programa e Ordem. Esta rotina emite uma mensagem caso o usuário não tenha habilitação
Function FunVerifica_Permissao(ByVal Programa As String, Ordem As String) As Boolean
    Dim Tb_Habilit As ADODB.Recordset
    
    FunVerifica_Permissao = True
    
    If func_cd_operador = 9999 Then
       Exit Function
    End If

    SubQOpenRecordset Tb_Habilit, "select * from Habilitacoes where func_cd_operador = " & func_cd_operador & " and  roti_tx_programa  = '" & Programa & "' and roti_tx_ordem = '" & Ordem & "'", Estatico

    If Tb_Habilit.EOF Then
       Beep
       MsgBox "Você não possui permissão para acessar esta rotina.", 16, "AVISO"
       FunVerifica_Permissao = False
    ElseIf Tb_Habilit!habi_bl_habilitado = "N" Then
        Beep
        MsgBox "Você não possui permissão para acessar esta rotina.", 16, "AVISO"
        FunVerifica_Permissao = False
    End If

    Tb_Habilit.Close
    Set Tb_Habilit = Nothing

End Function

' Verifica se o usuário possui permisssão para acessar a rotina determinada em Programa e Ordem. Esta rotina NÃO emite mensagem caso o usuário não tenha habilitação
Function FunVerifica_Permissao_SemMsg(ByVal Programa As String, Ordem As String) As Boolean
    Dim Tb_Habilit As ADODB.Recordset
    
    SubQOpenRecordset Tb_Habilit, _
        "select habi_bl_habilitado from Habilitacoes " & _
        "WHERE func_cd_operador = " & func_cd_operador & _
        "  AND roti_tx_programa = '" & Programa & "'" & _
        "  AND roti_tx_ordem = '" & Ordem & "'", _
        Estatico
  
    If Tb_Habilit.EOF Then
        FunVerifica_Permissao_SemMsg = False
    ElseIf Tb_Habilit!habi_bl_habilitado = "N" Then
        FunVerifica_Permissao_SemMsg = False
    Else
        FunVerifica_Permissao_SemMsg = True
    End If
    
    Tb_Habilit.Close
    Set Tb_Habilit = Nothing
  
End Function

Function funTrataCabecalho(XVariavel As String) As String
    Dim Cont As Integer
    Dim XResult As String
    
    Cont = 1
    While (Cont <= Len(Trim(XVariavel)))
        If (Mid(XVariavel, Cont, 1) <> "*") And (Mid(XVariavel, Cont, 1) <> " ") Then
            XResult = XResult + Mid(XVariavel, Cont, 1)
        End If
        Cont = Cont + 1
    Wend
    
    funTrataCabecalho = XResult
    
End Function

' Realiza a chamada do formulário de Filtro, de acordo com o tipo de dados da coluna do grid.
Function FunExecutaFiltro(ControleDados As Object, XForm As Form, Posicaox As Single)
    Dim CurCol As Integer 'Receberá o numero da coluna clicada
    
    'Verifica se tabela está vazia
    If FunTabelaVazia(ControleDados) Then
       Exit Function
    End If
    
    'A variável "curcol" contém o numero da coluna clicada
    CurCol = XForm.TDBGrid1.ColContaining(Posicaox)
    
    'Corresponde a um campo qualquer da tabela
    If CurCol >= 0 Then
       XForm.TDBGrid1.Col = CurCol
       
       If XForm.TDBGrid1.Columns(CurCol).HeadFont.Italic Then
          MsgBox "Campo já se encontra filtrado", 32, "ERRO"
          Exit Function
       End If
       
       TipoVar = funTipoCampo(ControleDados, XForm.TDBGrid1.Columns, CurCol)
       If TipoVar = "N" Or TipoVar = "D" Or TipoVar = "Y" Then
           FiltroNumerico.Label2.Caption = XForm.TDBGrid1.Columns(CurCol).Caption
           FiltroNumerico.ControleDados() = ControleDados
           FiltroNumerico.Grid() = XForm.TDBGrid1.Columns
           FiltroNumerico.ColunaGrid() = CurCol
           FiltroNumerico.Show 1
       Else
           FiltroCaracter.Label2.Caption = XForm.TDBGrid1.Columns(CurCol).Caption
           FiltroCaracter.ControleDados() = ControleDados
           FiltroCaracter.Grid() = XForm.TDBGrid1.Columns
           FiltroCaracter.ColunaGrid() = CurCol
           FiltroCaracter.Show 1
       End If

       If Filtrou = True Then
         subrecarregadadosNV ControleDados, "", Ordem, "", FiltroAtual
       End If
    End If

End Function

Function FunNulo(Texto As Variant) As String
    If IsNull(Texto) Or IsEmpty(Texto) Or Texto = "" Then
        FunNulo = ""
    Else
        FunNulo = Texto
    End If
End Function

Function funNull(XLV_Valor As Variant) As Variant
  'Retorna NULL caso o valor passado seja zero ou vazio.
  If (XLV_Valor = "") Or (XLV_Valor = 0) Then
    funNull = Null
  Else
    funNull = XLV_Valor
  End If
End Function

'RETORNA A COTAÇÃO DE UMA MOEDA
Public Function FunCotacaoMoedaPadrao(XLD_DTMOVIMENTO As String, XLB_DTESPECIFICA As Boolean) As Double
    'VARIAVEL OBJETO
    Dim XLO_MOEDA As ADODB.Recordset 'Object
    
    'SELECIONA INDEXADOR PADRAO
    XGT_SQL = "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_tx_Padrao='S'"

    'PEGA A COTAÇÃO DE UMA DATA
    If XLB_DTESPECIFICA Then
        XGT_SQL = XGT_SQL & " AND cota_dt_datacotacao = '" & Trim(Replace(XLD_DTMOVIMENTO, "'", "")) & "'"
    End If
    
    XGT_SQL = XGT_SQL & " ORDER BY cota_dt_datacotacao DESC"
    SubQOpenRecordset XLO_MOEDA, XGT_SQL, Estatico
    
    If Not XLO_MOEDA.EOF Then
        FunCotacaoMoedaPadrao = XLO_MOEDA!cota_vl_valor
    Else
        FunCotacaoMoedaPadrao = 1
    End If
    
    XLO_MOEDA.Close
    Set XLO_MOEDA = Nothing
    
    
End Function

'Retorna o Digito Verifica do Codigo Reduzido do Material
Function FunDigMaterial(CodMaterial As String) As Integer
    Dim VarLIntPosicao As Integer
    Dim VarLIntSoma As Integer
    Dim VarLIntMulti As Integer
    Dim VarLIntTotalSoma As Integer
    Dim VarLDblDivisao As Double
    Dim VarLIntResto As Integer
    
    VarLIntTotalSoma = 0
    VarLIntMulti = 2
    
    For VarLIntPosicao = 1 To 6
        VarLIntSoma = Val(Mid(CodMaterial, VarLIntPosicao, 1)) * VarLIntMulti
        VarLIntTotalSoma = VarLIntTotalSoma + VarLIntSoma
        VarLIntMulti = VarLIntMulti + 1
    Next
    
    VarLDblDivisao = VarLIntTotalSoma / 11
    VarLIntResto = Int((VarLDblDivisao - Int(VarLDblDivisao)) * 10)
    
    If VarLIntResto <> 0 Then VarLIntResto = Int(11 - VarLIntResto)
    If VarLIntResto = 10 Then VarLIntResto = 0
    
    FunDigMaterial = VarLIntResto
End Function

'Pega o caminho padrão do Banco
Function FunPadraoCaminho() As String
    Dim RecPadroes As ADODB.Recordset

    Conexao.Provider = "Microsoft.Jet.OLEDB.4.0"
    Conexao.Open App.Path & "\Padroes.mdb", "admin", ""
    
    SubQOpenRecordset RecPadroes, "Select * from Padroes", Dinamico
    
    RecPadroes.MoveFirst
    If RecPadroes!raizbanco = "" Or IsNull(RecPadroes!raizbanco) Then
          
            RecPadroes!raizbanco = App.Path & "\QualiAdmfin.mdb"
            RecPadroes!Logomarca = App.Path & "\Logo.bmp"
        RecPadroes.Update
        RecPadroes.MoveFirst
    End If
    
    FunPadraoCaminho = RecPadroes!raizbanco
    
    RecPadroes.Close
    Set RecPadroes = Nothing
    Conexao.Close
End Function

Function FunObrigatorioTXT(Text As TextBox, Texto As String) As Boolean
    If Trim$(Text) = "" Then
        MsgBox Texto, vbCritical, "CAMPO OBRIGATÓRIO"
        FunObrigatorioTXT = True
        If Text.Enabled = True Then Text.SetFocus
    End If
End Function

'Torna o preenchimento do combo obrigatório
Function FunObrigatorioCBO(combo As DataCombo, Texto As String) As Boolean
    If combo.Text = "" Then
        MsgBox Texto, vbInformation, "CAMPO OBRIGATÓRIO"
        FunObrigatorioCBO = True
        If combo.Enabled = True Then combo.SetFocus
    End If
End Function

'Torna o preenchimento do combo obrigatório
Function FunObrigatorioCBOX(combo As ComboBox, Texto As String) As Boolean
    If combo.Text = "" Then
        MsgBox Texto, vbInformation, "CAMPO OBRIGATÓRIO"
        FunObrigatorioCBOX = True
        combo.SetFocus
    End If
End Function

Function FunObrigatorioDTP(dtp As DTPicker, Texto As String) As Boolean
    If IsNull(dtp.Value) Then
        MsgBox Texto, vbInformation, "CAMPO OBRIGATÓRIO"
        FunObrigatorioDTP = True
        dtp.SetFocus
    End If
End Function

Function FunNuloData(Texto As Variant, Banco As String) As String
    If Not IsDate(Texto) Then
        FunNuloData = "Null"
    ElseIf Banco = "Access" Then
        FunNuloData = "#" & Format(CVDate(Texto), "MM/DD/YYYY") & "#"
    ElseIf Banco = "Sql Server" Then
        FunNuloData = "'" & Format(CVDate(Texto), "MM/DD/YYYY") & "'"
    Else
        FunNuloData = "'" & Format(CVDate(Texto), "MM/DD/YYYY") & "'"
    End If
End Function


Function FunNuloDataComHora(Texto As Variant, Banco As String) As String
    If Not IsDate(Texto) Then
        FunNuloDataComHora = "Null"
    ElseIf Banco = "Access" Then
        FunNuloDataComHora = "#" & Format(CVDate(Texto), "MM/DD/YYYY HH:MM") & "#"
    ElseIf Banco = "Sql Server" Then
        FunNuloDataComHora = "'" & Format(CVDate(Texto), "MM/DD/YYYY HH:MM") & "'"
    Else
        FunNuloDataComHora = "'" & Format(CVDate(Texto), "MM/DD/YYYY HH:MM") & "'"
    End If
End Function

Function funNuloDataHora(Texto As Variant, Banco As String, Tipo As String) As String
    If Not IsDate(Texto) Then
        funNuloDataHora = "Null"
    ElseIf Banco = "Access" Then
            funNuloDataHora = "#" & Format(CVDate(Texto), "MM/DD/YYYY")
            If Tipo = "I" Then
                funNuloDataHora = funNuloDataHora & " 00:00#"
            Else
                funNuloDataHora = funNuloDataHora & " 23:59#"
            End If
        ElseIf Banco = "Sql Server" Then
            funNuloDataHora = "'" & Format(CVDate(Texto), "MM/DD/YYYY")
            If Tipo = "I" Then
                funNuloDataHora = funNuloDataHora & " 00:00'"
            Else
                funNuloDataHora = funNuloDataHora & " 23:59'"
            End If
    End If
End Function

Function FunNuloData2(Texto As Variant, Banco As String) As String
    If Not IsDate(Texto) Then
        FunNuloData2 = "Null"
    ElseIf Banco = "Access" Then
        FunNuloData2 = "#" & Format(CVDate(Texto), "MM/DD/YYYY") & "#"
    ElseIf Banco = "Sql Server" Then
        FunNuloData2 = "" & Format(CVDate(Texto), "MM/DD/YYYY") & ""
    Else
        FunNuloData2 = "" & Format(CVDate(Texto), "MM/DD/YYYY") & ""
    End If
End Function

'verificação de plano de contas
Function FunVerificaContaAnt(valor As String) As String

    Dim Cont As Integer
    Dim Tam As Integer
    Dim resultado As String
    Dim achei As Boolean
    
    Cont = 1
    Tam = Len(valor)
    resultado = valor
    achei = False
       
    'Verificar se o código só possui um nivel
    While Cont < Tam
        If Mid(resultado, Cont, 1) = "." Then
            achei = True
            Cont = Tam
        End If
        Cont = Cont + 1
    Wend
    If Not achei Then
        FunVerificaContaAnt = ""
        Exit Function
    End If
    
    Cont = 1
    Tam = Len(valor)
    resultado = valor
    achei = False
    
    'Verificar se o código cadastro é o 1º nível
    While Cont < Tam
        If Mid(resultado, Cont, 1) = "." Then
            achei = True
        ElseIf Not IsNumeric(Mid(resultado, Cont, 1)) And achei Then
            FunVerificaContaAnt = ""
            Exit Function
        ElseIf IsNumeric(Mid(resultado, Cont, 1)) And achei Then
            Cont = Tam
        End If
        Cont = Cont + 1
    Wend
       
    Cont = 1
    Tam = Len(valor)
    resultado = valor
    achei = False

    'Verificar se o último caracter é um valor ou um espaço
    If (Right(valor, Tam)) = "" Then
        While Tam > Cont
            If Mid(resultado, Tam, 1) = "." Then
                If Not achei Then
                    Tam = Tam - 1
                Else
                    Tam = Cont
                End If
            Else
                Mid(resultado, Tam, 1) = " "
                achei = True
                Tam = Tam - 1
            End If
        Wend
    Else
        While Tam > Cont
            If Mid(resultado, Tam, 1) = "." Then
                If Not achei Then
                    Tam = Tam - 1
                Else
                    Tam = Cont
                End If
            ElseIf Mid(resultado, Tam, 1) = " " Then
                Tam = Tam - 1
            ElseIf IsNumeric(Mid(resultado, Tam, 1)) Then
                Mid(resultado, Tam, 1) = " "
                Tam = Tam - 1
                achei = True
            End If
        Wend
    End If
    
    FunVerificaContaAnt = resultado

End Function

'Define ultimo nivel do plano de contas
Function FunDefineUltimoNivel(Mascara As String) As Integer
    Dim x As Integer
    Dim xCaracter As String
    Dim xNivel As Integer
    
    xNivel = 0
    For x = 1 To Len(Mascara)
        xCaracter = Mid(Mascara, x, 1)
        If xCaracter = "." Then
            xNivel = xNivel + 1
        End If
        If xCaracter = " " Then
            Exit For
        End If
    Next
    If xNivel = 0 Then
        FunDefineUltimoNivel = 1
    Else
        If xCaracter = " " Then
            FunDefineUltimoNivel = xNivel
        Else
            FunDefineUltimoNivel = xNivel + 1
        End If
    End If
End Function

Function FunObrigatorioMSK(Mask As MaskEdBox, Texto As String) As Boolean
    If Mask.PromptInclude Then
        Mask.PromptInclude = False
        If Mask.Text = "" Then
            Mask.PromptInclude = True
            MsgBox Texto, vbCritical, "CAMPO OBRIGATÓRIO"
            FunObrigatorioMSK = True
            Mask.SetFocus
            Exit Function
        End If
        Mask.PromptInclude = True
    Else
        If Mask.Text = "" Then
            MsgBox Texto, vbCritical, "CAMPO OBRIGATÓRIO"
            FunObrigatorioMSK = True
            Mask.SetFocus
            Exit Function
        End If
    End If
End Function

'Soma ou subtrai unidade de Quantid operando também FRM_Quantid
Function funOperacaoQuantid(XLT_OPERACAO As String) As Long
    If XLT_OPERACAO = "+" Then
        Quantid = Quantid + 1
    ElseIf XLT_OPERACAO = "-" Then
        Quantid = Quantid - 1
    End If
    funOperacaoQuantid = Quantid
End Function

Function FunPegaGuardaUltimo(Tipo As String) As Double
    Dim ResUltimoDoc As ADODB.Recordset
    Dim Cont As Integer
    
    SubQOpenRecordset ResUltimoDoc, "SELECT * FROM Ultimos", Dinamico
    
    Select Case Tipo
        Case "D" 'Número Documento
            FunPegaGuardaUltimo = ResUltimoDoc!ulti_nr_numdoc + 1
            ResUltimoDoc!ulti_nr_numdoc = (ResUltimoDoc!ulti_nr_numdoc + 1)
        Case "F" 'Número da Fatura
            FunPegaGuardaUltimo = ResUltimoDoc!ulti_nr_fatura + 1
            ResUltimoDoc!ulti_nr_fatura = (ResUltimoDoc!ulti_nr_fatura + 1)
        Case "C" 'Número de Controle
            FunPegaGuardaUltimo = ResUltimoDoc!ulti_nr_controle + 1
            ResUltimoDoc!ulti_nr_controle = (ResUltimoDoc!ulti_nr_controle + 1)
        Case "B" 'Número de Borderô
            FunPegaGuardaUltimo = ResUltimoDoc!ulti_nr_bordero + 1
            ResUltimoDoc!ulti_nr_bordero = (ResUltimoDoc!ulti_nr_bordero + 1)
        Case "E" 'Número de movimento do Estoque
            FunPegaGuardaUltimo = ResUltimoDoc!ulti_nr_MovEstoque + 1
            ResUltimoDoc!ulti_nr_MovEstoque = (ResUltimoDoc!ulti_nr_MovEstoque + 1)
        Case "R" 'Número de Recibo de Receita
            FunPegaGuardaUltimo = ResUltimoDoc!ulti_nr_ReciboReceita + 1
            ResUltimoDoc!ulti_nr_ReciboReceita = (ResUltimoDoc!ulti_nr_ReciboReceita + 1)
        Case "PR" 'Documento Pagto/Rec
            FunPegaGuardaUltimo = ResUltimoDoc!Ulti_nr_DocPagRec + 1
            ResUltimoDoc!Ulti_nr_DocPagRec = (ResUltimoDoc!Ulti_nr_DocPagRec + 1)
        Case "RS" ' Pega o último número da remessa do ExpBrasil
            FunPegaGuardaUltimo = ResUltimoDoc!Ulti_nr_Remessa + 1
            ResUltimoDoc!Ulti_nr_Remessa = (ResUltimoDoc!Ulti_nr_Remessa + 1)
        Case "P" ' Pega o último número da provisao
            FunPegaGuardaUltimo = ResUltimoDoc!Ulti_nr_provisao + 1
            ResUltimoDoc!Ulti_nr_provisao = (ResUltimoDoc!Ulti_nr_provisao + 1)

    End Select
    
    ResUltimoDoc.Update
    ResUltimoDoc.Close
    Set ResUltimoDoc = Nothing
End Function

Function FunNuloMSKdata(Texto As Variant) As String
    If IsNull(Texto) Then
        FunNuloMSKdata = "  /  /  "
    Else
        FunNuloMSKdata = Format(CVDate(Texto), "DD/MM/YYYY")
    End If
End Function

Function FunNuloValor(combo As DataCombo) As Integer
    If combo.BoundText <> "" Then
       FunNuloValor = combo.BoundText
    Else
        FunNuloValor = 0
    End If
End Function

Function FunNuloVal(Texto As Variant) As String
    If IsNull(Texto) Or IsEmpty(Texto) Or Texto = "" Then
        FunNuloVal = 0
    Else
        FunNuloVal = Texto
    End If
End Function

Function FunNuloBancoVal(Texto As Variant) As String
    If IsNull(Texto) Or Texto = Empty Then
        FunNuloBancoVal = "Null"
    Else
        FunNuloBancoVal = Trim(Str(Texto))
    End If
End Function

Function FunNuloBancoZero(Texto As Variant) As String
    If Texto = 0 Then
        FunNuloBancoZero = "Null"
    Else
        FunNuloBancoZero = Trim(Str(Texto))
    End If
End Function

Function FunNuloBanco(Texto As Variant) As String
    If IsNull(Texto) Or Texto = Empty Then
        FunNuloBanco = "Null"
    Else
        FunNuloBanco = Trim(Str(Texto))
    End If
End Function

Function FunEnganaAccess(Texto As String) As String

  FunEnganaAccess = Texto
  
End Function

Function FunTrataZero(Texto As Variant) As String
   On Error GoTo TrataErro
   
   Dim XLD_RETVAL As Double
   Dim XLB_ERRO As Boolean
   
   XLB_ERRO = False
   
   XLD_RETVAL = CDbl(Texto)
      
      
Continua:
   
   If XLD_RETVAL = 0 And XLB_ERRO = False Then
     FunTrataZero = ""
   Else
     FunTrataZero = Texto
   End If
   
   Exit Function
   
TrataErro:
    XLB_ERRO = True
   GoTo Continua
End Function

' conferências dos dígitos verificadores  do CPF
Function FunConfereCPF(CPF As String) ' retorna True quando OK
    Dim Soma As Integer
    Dim Dígito1, Dígito2 As Integer
    Dim Resto As Integer
    
    'Considera errado a priori
    FunConfereCPF = False
            
    If Not IsNumeric(CPF) Or Len(Trim(CPF)) <> 11 Or _
        CPF = "11111111111" Or CPF = "22222222222" Or CPF = "33333333333" Or _
        CPF = "44444444444" Or CPF = "55555555555" Or CPF = "66666666666" Or _
        CPF = "77777777777" Or CPF = "88888888888" Or CPF = "99999999999" Or _
        CPF = "00000000000" Then
       Exit Function
    End If
    
    'calcula o primeiro dígito soma até o oitavo algarismo
    Soma = 0
    Soma = Soma + Val(Mid(CPF, 9, 1)) * 2
    Soma = Soma + Val(Mid(CPF, 8, 1)) * 3
    Soma = Soma + Val(Mid(CPF, 7, 1)) * 4
    Soma = Soma + Val(Mid(CPF, 6, 1)) * 5
    Soma = Soma + Val(Mid(CPF, 5, 1)) * 6
    Soma = Soma + Val(Mid(CPF, 4, 1)) * 7
    Soma = Soma + Val(Mid(CPF, 3, 1)) * 8
    Soma = Soma + Val(Mid(CPF, 2, 1)) * 9
    Soma = Soma + Val(Mid(CPF, 1, 1)) * 10
     
    Resto = Soma Mod 11
    Dígito1 = IIf(Resto = 0 Or Resto = 1, 0, 11 - Resto)
    
    'calcula o segundo dígito
    Soma = 0
    Soma = Soma + Val(Mid(CPF, 10, 1)) * 2
    Soma = Soma + Val(Mid(CPF, 9, 1)) * 3
    Soma = Soma + Val(Mid(CPF, 8, 1)) * 4
    Soma = Soma + Val(Mid(CPF, 7, 1)) * 5
    Soma = Soma + Val(Mid(CPF, 6, 1)) * 6
    Soma = Soma + Val(Mid(CPF, 5, 1)) * 7
    Soma = Soma + Val(Mid(CPF, 4, 1)) * 8
    Soma = Soma + Val(Mid(CPF, 3, 1)) * 9
    Soma = Soma + Val(Mid(CPF, 2, 1)) * 10
    Soma = Soma + Val(Mid(CPF, 1, 1)) * 11
     
    Resto = Soma Mod 11
    Dígito2 = IIf(Resto = 0 Or Resto = 1, 0, 11 - Resto)
       
    If Dígito1 = Val(Mid(CPF, 10, 1)) And Dígito2 = Val(Mid(CPF, 11, 1)) Then
        FunConfereCPF = True
    End If
End Function

' Confere o digito verificador dod GCG
Function FunConfereCGC(CGC As String)  ' retorna True quando OK
    Dim Soma As Integer
    Dim Dígito1, Dígito2 As Integer
    Dim Resto As Integer
    
    FunConfereCGC = False ' considera errado a priori
        
    If Len(Trim(CGC)) <> 14 Then
        Exit Function
    End If
    
    ' calcula o primeiro dígito
    Soma = 0
    Soma = Soma + Val(Mid(CGC, 1, 1)) * 5
    Soma = Soma + Val(Mid(CGC, 2, 1)) * 4
    Soma = Soma + Val(Mid(CGC, 3, 1)) * 3
    Soma = Soma + Val(Mid(CGC, 4, 1)) * 2
    Soma = Soma + Val(Mid(CGC, 5, 1)) * 9
    Soma = Soma + Val(Mid(CGC, 6, 1)) * 8
    Soma = Soma + Val(Mid(CGC, 7, 1)) * 7
    Soma = Soma + Val(Mid(CGC, 8, 1)) * 6
    Soma = Soma + Val(Mid(CGC, 9, 1)) * 5
    Soma = Soma + Val(Mid(CGC, 10, 1)) * 4
    Soma = Soma + Val(Mid(CGC, 11, 1)) * 3
    Soma = Soma + Val(Mid(CGC, 12, 1)) * 2
    
    
    Resto = Soma Mod 11
    Dígito1 = IIf(Resto = 0 Or Resto = 1, 0, 11 - Resto)

    ' calcula o segundo dígito
    Soma = 0
    Soma = Soma + Val(Mid(CGC, 1, 1)) * 6
    Soma = Soma + Val(Mid(CGC, 2, 1)) * 5
    Soma = Soma + Val(Mid(CGC, 3, 1)) * 4
    Soma = Soma + Val(Mid(CGC, 4, 1)) * 3
    Soma = Soma + Val(Mid(CGC, 5, 1)) * 2
    Soma = Soma + Val(Mid(CGC, 6, 1)) * 9
    Soma = Soma + Val(Mid(CGC, 7, 1)) * 8
    Soma = Soma + Val(Mid(CGC, 8, 1)) * 7
    Soma = Soma + Val(Mid(CGC, 9, 1)) * 6
    Soma = Soma + Val(Mid(CGC, 10, 1)) * 5
    Soma = Soma + Val(Mid(CGC, 11, 1)) * 4
    Soma = Soma + Val(Mid(CGC, 12, 1)) * 3
    Soma = Soma + Dígito1 * 2
    
    Resto = Soma Mod 11
    Dígito2 = IIf(Resto = 0 Or Resto = 1, 0, 11 - Resto)
    
    If Dígito1 = Val(Mid(CGC, 13, 1)) And Dígito2 = Val(Mid(CGC, 14, 1)) Then
        FunConfereCGC = True
    End If
End Function

'Transforma a vígula que separa as casas decimais em ponto
Function funTransfVirgPonto(XLS_VALOR As Single) As String
    Dim Cont As Integer
    Dim saldo As String
    Dim resultado As String
    
    saldo = CStr(XLS_VALOR)
    Cont = 1
    
    While (Cont <= Len(Trim$(saldo)))
        If Mid(saldo, Cont, 1) = "," Then
            resultado = resultado + "."
        Else
            resultado = resultado + Mid$(saldo, Cont, 1)
        End If
        Cont = Cont + 1
    Wend
    funTransfVirgPonto = resultado
End Function

'RETORNA O TAMANHO DO NIVEL
Public Function FunTamanhoNivel(XLT_CODESTRUTURADO As String, XLI_NIVEL As Integer)
    Dim XLI_POSICAO As Integer
    XLI_POSICAO = InStr(1, XLT_CODESTRUTURADO, " ")
    If XLI_POSICAO = 0 Then
        XLI_POSICAO = Len(XLT_CODESTRUTURADO)
    Else
        XLI_POSICAO = XLI_POSICAO - 1
    End If
    FunTamanhoNivel = XLI_POSICAO
End Function

Sub SubPreencheComZero(XLO_MASKBOX As MaskEdBox)
   
   Dim XLS_Temp As String
   Dim XLI_CONT As Integer
   
   If Len(XLO_MASKBOX.Text) < 4 Then
     XLS_Temp = XLO_MASKBOX.Text
     XLO_MASKBOX.Text = ""
     For XLI_CONT = 1 To Len(XLO_MASKBOX.Mask) - Len(XLS_Temp)
       XLO_MASKBOX.Text = XLO_MASKBOX.Text + "0"
     Next
     XLO_MASKBOX.Text = XLO_MASKBOX.Text + XLS_Temp
   End If
 
End Sub

Sub SubImagemFundo(XLL_RESOLUCAO As Long)
   
    If XLL_RESOLUCAO < 10000 Then   '640x480
        MdiPrincipal.Picture = LoadPicture(App.Path & "\ImagemFundo640.jpg", vbLPLarge, vbLPColor)
    ElseIf XLL_RESOLUCAO > 10000 And XLL_RESOLUCAO < 13000 Then  '800x600
        MdiPrincipal.Picture = LoadPicture(App.Path & "\ImagemFundo800.jpg", vbLPLarge, vbLPColor)
    ElseIf XLL_RESOLUCAO > 13000 And XLL_RESOLUCAO < 16000 Then  '1024 x 728
        MdiPrincipal.Picture = LoadPicture(App.Path & "\ImagemFundo1024.jpg", vbLPLarge, vbLPColor)
    Else        '1024x728 ou superior
        MdiPrincipal.Picture = LoadPicture(App.Path & "\ImagemFundo1280.jpg", vbLPLarge, vbLPColor)
    End If
    
End Sub

Function FunMesExtenso(XLI_MES As Integer) As String
  Select Case XLI_MES
    Case 1
      FunMesExtenso = "Janeiro"
    Case 2
      FunMesExtenso = "Fevereiro"
    Case 3
      FunMesExtenso = "Março"
    Case 4
      FunMesExtenso = "Abril"
    Case 5
      FunMesExtenso = "Maio"
    Case 6
      FunMesExtenso = "Junho"
    Case 7
      FunMesExtenso = "Julho"
    Case 8
      FunMesExtenso = "Agosto"
    Case 9
      FunMesExtenso = "Setembro"
    Case 10
      FunMesExtenso = "Outubro"
    Case 11
      FunMesExtenso = "Novembro"
    Case Else
      FunMesExtenso = "Dezembro"
  End Select
End Function

'Esta função recebe uma string e a varre retirando os acentos e caracteres especiais os caracteres especiais recebem " " ( vazio ) e os acentuados ficam sem acento
'o retorno desta função é a String tratada
Function funTiraAcento(XLT_TEXTO As String, XLT_UPCASE As Boolean) As String
  
  Dim Cont As Long
  Dim XLT_Result As String
  Dim XLT_LETRA As String
  Dim XLL_ASCIILETRA As Long
  
  Cont = 1
  While (Cont <= Len(XLT_TEXTO))
    XLT_LETRA = (Mid(XLT_TEXTO, Cont, 1))
    XLL_ASCIILETRA = Asc(XLT_LETRA)
    If ((XLL_ASCIILETRA >= 224) And (XLL_ASCIILETRA <= 229)) _
    Or ((XLL_ASCIILETRA >= 192) And (XLL_ASCIILETRA <= 197)) Then
      XLT_Result = XLT_Result + "A"
    Else
      If ((XLL_ASCIILETRA >= 200) And (XLL_ASCIILETRA <= 203)) _
      Or ((XLL_ASCIILETRA >= 232) And (XLL_ASCIILETRA <= 235)) Then
        XLT_Result = XLT_Result + "E"
      Else
        If ((XLL_ASCIILETRA >= 204) And (XLL_ASCIILETRA <= 207)) _
        Or ((XLL_ASCIILETRA >= 236) And (XLL_ASCIILETRA <= 239)) Then
          XLT_Result = XLT_Result + "I"
        Else
          If ((XLL_ASCIILETRA >= 210) And (XLL_ASCIILETRA <= 214)) _
          Or ((XLL_ASCIILETRA >= 242) And (XLL_ASCIILETRA <= 246)) Then
            XLT_Result = XLT_Result + "O"
          Else
            If ((XLL_ASCIILETRA >= 217) And (XLL_ASCIILETRA <= 220)) _
            Or ((XLL_ASCIILETRA >= 249) And (XLL_ASCIILETRA <= 252)) Then
              XLT_Result = XLT_Result + "U"
            Else
              If ((XLL_ASCIILETRA = 199) Or (XLL_ASCIILETRA = 231)) Then
                XLT_Result = XLT_Result + "C"
              Else
                If ((XLL_ASCIILETRA >= 33) And (XLL_ASCIILETRA <= 47) And (XLL_ASCIILETRA <> 46)) _
                Or ((XLL_ASCIILETRA >= 58) And (XLL_ASCIILETRA <= 63)) _
                Or ((XLL_ASCIILETRA >= 91) And (XLL_ASCIILETRA <= 96)) _
                Or ((XLL_ASCIILETRA >= 123) And (XLL_ASCIILETRA <= 126)) _
                Or ((XLL_ASCIILETRA >= 161) And (XLL_ASCIILETRA <= 191)) _
                Or ((XLL_ASCIILETRA >= 215) And (XLL_ASCIILETRA <= 216)) _
                Or ((XLL_ASCIILETRA >= 222) And (XLL_ASCIILETRA <= 223)) _
                Or ((XLL_ASCIILETRA >= 217) And (XLL_ASCIILETRA <= 220)) _
                Or ((XLL_ASCIILETRA >= 247) And (XLL_ASCIILETRA <= 248)) _
                Or ((XLL_ASCIILETRA >= 217) And (XLL_ASCIILETRA <= 220)) _
                Or (XLL_ASCIILETRA = 198) Or (XLL_ASCIILETRA = 208) Or (XLL_ASCIILETRA = 230) Or (XLL_ASCIILETRA = 241) Then
                  XLT_Result = XLT_Result + " " 'caracteres especiais
                Else
                  If XLT_UPCASE = True Then
                    XLT_Result = XLT_Result & UCase(XLT_LETRA)
                  Else
                    XLT_Result = XLT_Result & UCase(XLT_LETRA)
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End If
    Cont = Cont + 1
  Wend
  funTiraAcento = XLT_Result
End Function

Function FunZeros(XLI_QTD As Integer) As String
  
   Dim XLT_ZEROS As String
   Dim XLI_CONT As Integer
      
   For XLI_CONT = 1 To XLI_QTD
     XLT_ZEROS = XLT_ZEROS & "0"
   Next
   
   FunZeros = XLT_ZEROS
   
End Function

Function FunPreencheComNove(XLI_QTD As Integer) As String
  
   Dim XLT_NOVES As String
   Dim XLI_CONT As Integer
      
   For XLI_CONT = 1 To XLI_QTD
     XLT_NOVES = XLT_NOVES & "9"
   Next
   
   FunPreencheComNove = XLT_NOVES
   
End Function

Public Function funSubstituiCaracter(Texto As String, Caracter As String, NovoCaracter As String) As String
  Dim Pos1 As Integer
  Pos1 = 1
  While Pos1 <> 0
    Pos1 = InStr(Pos1, Texto, Caracter, vbTextCompare)
    If Pos1 > 0 Then
      Texto = Left$(Texto, Pos1 - 1) & NovoCaracter & Right$(Texto, Len(Texto) - Pos1)
    End If
  Wend
  
  funSubstituiCaracter = Texto

End Function

Public Function funTemProtecao(XLT_LOCALINI As String, XLT_NOMEAPL As String, XLT_ID As String) As Boolean
  Dim XLO_PROTECAO As T_PROT_INFO  'Vetor com informações da proteção
  Dim XLT_LOCALPROT As String * 255
  Dim XLI_RESULT As Integer

    funTemProtecao = True

End Function

Public Function funEscolherGrafico(XLI_PADRAO As Integer, XLI_NRSUMS As Integer) As Integer
  XGI_GRAFICO_NRSUMS = XLI_NRSUMS
  FrmTiposGraficos.Show 1
  If XGB_RESPOSTA = True Then
    funEscolherGrafico = XGI_GRAFICO
  Else
    funEscolherGrafico = XLI_PADRAO
  End If
End Function

Public Function funNomeGrafico(XLI_GRAFICO As Integer) As String
  Select Case XLI_GRAFICO
    Case crSideBySideBarGraph
      funNomeGrafico = "Colunas"
    Case crFaked3DSideBySideBarGraph
      funNomeGrafico = "Colunas 3D"
    Case crStackedBarGraph
      funNomeGrafico = "Empilhado"
    Case crFaked3DStackedBarGraph
      funNomeGrafico = "Empilhado 3D"
    Case crPercentBarGraph
      funNomeGrafico = "Percentual"
    Case crFaked3DPercentBarGraph
      funNomeGrafico = "Percentual 3D"
    Case crRegularLineGraph
      funNomeGrafico = "Linha"
    Case 120
      funNomeGrafico = "Área"
    Case crThreeDRegularGraph
      funNomeGrafico = "Blocos 3D"
    Case crRegularPieGraph
      funNomeGrafico = "Pizza"
    Case crMultiplePieGraph
      funNomeGrafico = "Pizzas Múltiplas"
    Case crMultipleProportionalPieGraph
      funNomeGrafico = "Pizzas Ponderadas"
    Case Else
      funNomeGrafico = "Desconhecido"
  End Select
End Function

Function funInteiroPositivo(Numero As Variant) As Variant
  If Numero < 0 Then
    funInteiroPositivo = 0
  Else
    funInteiroPositivo = Numero
  End If
End Function

Public Function ExtractFilePath(caminho As String) As String
  Dim pos As Integer
  Dim S As String
  
  pos = InStrRev(caminho, "\")
  S = Left(caminho, pos)
  ExtractFilePath = S
End Function

Public Function funUltimoDiaMes(Data As Date) As Date
  funUltimoDiaMes = DateSerial(Year(Data), Month(Data) + 1, 0)
End Function

Public Function funPrimeiroDiaMes(Data As Date) As Date
  funPrimeiroDiaMes = DateSerial(Year(Data), Month(Data), 1)
End Function

Function funNuloString(XLT_TEXTO As String) As String
  If XLT_TEXTO = "" Then
    funNuloString = "NULL"
  Else
    funNuloString = "'" & XLT_TEXTO & "'"
  End If
End Function

Function funTiraCaracteresNulos(XLT_TEXTO As String) As String
  
  funTiraCaracteresNulos = Left$(XLT_TEXTO, InStr(XLT_TEXTO, Chr$(0)) - 1)

End Function


Function funTiraMascara(XLT_TEXTOMASCARADO As String) As String
  Dim XLI_X As Integer
  Dim XLT_GRUPO As String
  Dim XLT_CARACTER As String
    
  XLT_GRUPO = ""
  For XLI_X = 1 To Len(XLT_TEXTOMASCARADO)
    XLT_CARACTER = Mid(XLT_TEXTOMASCARADO, XLI_X, 1)
    If XLT_CARACTER <> " " Then
      XLT_GRUPO = XLT_GRUPO + XLT_CARACTER
    Else
      XLI_X = Len(XLT_TEXTOMASCARADO) + 1
    End If
  Next
  funTiraMascara = XLT_GRUPO
End Function

Sub subGravaLog(XLI_CODUSUARIO As Integer, XLT_ROTINA As String, XLT_ORDEMROTINA As String, XLD_HORA As Date, XLT_DESCRICAO As String)

  Dim XLO_LOG As New ADODB.Recordset
  
  SubQOpenRecordset XLO_LOG, "SELECT * FROM Log WHERE func_cd_Operador=-1", Dinamico
  
  XLO_LOG.AddNew
  XLO_LOG!func_cd_operador = XLI_CODUSUARIO
  XLO_LOG!roti_tx_Programa = XLT_ROTINA
  XLO_LOG!roti_tx_Ordem = XLT_ORDEMROTINA
  XLO_LOG!log_dt_Hora = XLD_HORA
  XLO_LOG!log_tx_Descricao = XLT_DESCRICAO
  XLO_LOG.Update
  
  XLO_LOG.Close
  
  Set XLO_LOG = Nothing
  
End Sub
Function funTransfVirgPontoDouble(XLS_VALOR As Double) As String

    Dim Cont As Integer
    Dim saldo As String
    Dim resultado As String
    
    saldo = CStr(XLS_VALOR)
    Cont = 1
    
    While (Cont <= Len(Trim$(saldo)))
        If Mid(saldo, Cont, 1) = "," Then
            resultado = resultado + "."
        Else
            resultado = resultado + Mid$(saldo, Cont, 1)
        End If
        Cont = Cont + 1
    Wend
    funTransfVirgPontoDouble = resultado
End Function
Function funTransfVirgPontoVariant(XLS_VALOR As Variant) As String


    Dim Cont As Integer
    Dim saldo As String
    Dim resultado As String
    
    saldo = XLS_VALOR
    Cont = 1
    
    While (Cont <= Len(Trim$(saldo)))
        If Mid(saldo, Cont, 1) = "," Then
            resultado = resultado + "."
        Else
            resultado = resultado + Mid$(saldo, Cont, 1)
        End If
        Cont = Cont + 1
    Wend
    funTransfVirgPontoVariant = resultado
End Function

Function funTransfPontoVirg(XLS_VALOR As String) As String
    Dim Cont As Integer
    Dim saldo As String
    Dim resultado As String
    
    saldo = CStr(XLS_VALOR)
    Cont = 1
    
    While (Cont <= Len(Trim$(saldo)))
        If Mid(saldo, Cont, 1) = "." Then
            resultado = resultado + ","
        Else
            resultado = resultado + Mid$(saldo, Cont, 1)
        End If
        Cont = Cont + 1
    Wend
    funTransfPontoVirg = resultado
End Function

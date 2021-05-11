Attribute VB_Name = "ModGridComVetor"
Option Explicit
Sub subGravaManutencaoVetor(XLS_TABELA As String, VLV_CAMPOSTABELA As Variant, VLV_TIPOS As Variant, XLI_QTDCAMPOS As Byte, VLV_COLUNAS As Variant, XLT_COMPLEMENTO As String, VLV_VETOR As XArray, Grid As TDBGrid, XLV_CAMPOSITUACAO As Variant, VLV_CHAVES As Variant)
    Dim XLS_VALORES As String, XLS_ALTERACAO As String, XLS_COMPLEMENTOANT As String
    Dim XLI_X As Integer, XLI_Y As Integer, XLS_STATUS As String, XLS_CONDICAO As String
    Dim XLS_VALORCHAVE As String, XLS_SQL As String
    Dim XLO_CHECK As Object, XLV_VALORCAMPO As Variant, XLT_CAMPOS As String
    Dim XLS_CONJUNCAO As String
    'XLS_TABELA - nome da tabela onde serão inseridos os registros do grid
    'VLV_CAMPOSTABELA - string simples com os campos na respectiva ordem do grid
    
    'VLV_TIPOS - vetor simples com os respectivos tipos de dados das colunas do grid
    'N: Número; D: Data; T: Texto
    
    'XLI_QTDCAMPOS - quantidade de colunas do grid que serão gravadas
    'VLV_COLUNAS - vetor simples com o índice das respectivas colunas que serão gravadas
    'XLT_COMPLEMENTO - Valores de campos fixos que deseje gravar e não existe no grid para evitar
    'repetição Ex.: Todos os campos possuem o mesmo código de serviço
    
    'VLV_VETOR - vetor que será preenchido
    'Grid - grid que será preenchido
    
    'XLV_CAMPOSITUACAO - indica a posicao do vetor que se encontra o campo situação
    'VLV_CHAVES - Vetor indicando a posição das chaves, no vetor, que serão usadas para fazer as deleções e alterações
    XLS_COMPLEMENTOANT = XLT_COMPLEMENTO
    If VLV_VETOR(0, 0) <> Empty Then
        If XLT_COMPLEMENTO <> "" Then
            XLT_COMPLEMENTO = ", " & XLT_COMPLEMENTO
        End If
        
        For XLI_X = 0 To VLV_VETOR.UpperBound(1)
            XLS_CONDICAO = ""
            XLS_VALORES = ""
            XLS_ALTERACAO = ""
            XLT_CAMPOS = ""
            XLS_STATUS = VLV_VETOR(XLI_X, XLV_CAMPOSITUACAO)
            XLS_CONJUNCAO = ""
            'Identificar condições para Alteração e Deleção
            'XLS_CONDICAO = "true"
            For XLI_Y = 0 To UBound(VLV_CHAVES, 1)
                'XLS_CONDICAO = XLS_CONDICAO & " AND "
                If VLV_CHAVES(XLI_Y) > UBound(VLV_COLUNAS, 1) Then
                    XLS_VALORCHAVE = XLS_COMPLEMENTOANT
                Else
                    XLS_VALORCHAVE = VLV_VETOR(XLI_X, VLV_COLUNAS(VLV_CHAVES(XLI_Y)))
                End If
                If VLV_TIPOS(VLV_CHAVES(XLI_Y)) = "N" Then
                    XLS_CONDICAO = XLS_CONDICAO & XLS_CONJUNCAO & _
                      VLV_CAMPOSTABELA(VLV_CHAVES(XLI_Y)) & " = " & Str(XLS_VALORCHAVE)
                    XLS_CONJUNCAO = " AND "
                ElseIf VLV_TIPOS(VLV_CHAVES(XLI_Y)) = "T" Then
                    XLS_CONDICAO = XLS_CONDICAO & XLS_CONJUNCAO & _
                      VLV_CAMPOSTABELA(VLV_CHAVES(XLI_Y)) & " = '" & Str(XLS_VALORCHAVE) & "'"
                    XLS_CONJUNCAO = " AND "
                ElseIf VLV_TIPOS(VLV_CHAVES(XLI_Y)) = "D" Then
                    XLS_CONDICAO = XLS_CONDICAO & XLS_CONJUNCAO & _
                      VLV_CAMPOSTABELA(VLV_CHAVES(XLI_Y)) & " = " & FunNuloData(XLS_VALORCHAVE, NomeSgbd)
                    XLS_CONJUNCAO = " AND "
                End If
            Next
            'Formatar string de formação do SQL
            For XLI_Y = 0 To XLI_QTDCAMPOS - 1
                If XLS_VALORES <> "" Then
                    XLS_VALORES = XLS_VALORES & ", "
                    XLS_ALTERACAO = XLS_ALTERACAO & ", "
                    XLT_CAMPOS = XLT_CAMPOS & ", "
                End If
                
                If XLI_Y > UBound(VLV_COLUNAS, 1) Then
                    XLV_VALORCAMPO = XLS_COMPLEMENTOANT
                Else
                    XLV_VALORCAMPO = VLV_VETOR(XLI_X, VLV_COLUNAS(XLI_Y))
                End If
                XLT_CAMPOS = XLT_CAMPOS & VLV_CAMPOSTABELA(XLI_Y)
                If VLV_TIPOS(XLI_Y) = "N" Then
                    XLS_VALORES = XLS_VALORES & Str(XLV_VALORCAMPO)
                    XLS_ALTERACAO = XLS_ALTERACAO & VLV_CAMPOSTABELA(XLI_Y) & " = " & Str(XLV_VALORCAMPO)
                ElseIf VLV_TIPOS(XLI_Y) = "T" Then
                    XLS_VALORES = XLS_VALORES & "'" & XLV_VALORCAMPO & "'"
                    XLS_ALTERACAO = XLS_ALTERACAO & VLV_CAMPOSTABELA(XLI_Y) & " = '" & XLV_VALORCAMPO & "'"
                ElseIf VLV_TIPOS(XLI_Y) = "D" Then
                    XLS_VALORES = XLS_VALORES & FunNuloData((XLV_VALORCAMPO), NomeSgbd)
                    XLS_ALTERACAO = XLS_ALTERACAO & VLV_CAMPOSTABELA(XLI_Y) & " = " & FunNuloData(XLV_VALORCAMPO, NomeSgbd)
                End If
            Next
    
            If XLS_STATUS = "E" Then   'Exclusão do registro
                XLS_SQL = "DELETE FROM " & XLS_TABELA & " WHERE " & XLS_CONDICAO
            Else
                'Checar existência do registro - se sim , alterar, se não, cadastrar
                XLS_SQL = "SELECT * FROM " & XLS_TABELA & " WHERE " & XLS_CONDICAO
                SubQOpenRecordset XLO_CHECK, XLS_SQL, Estatico
                If XLO_CHECK.EOF Then
                    XLS_SQL = "INSERT INTO " & XLS_TABELA & " (" & XLT_CAMPOS & ") VALUES (" & XLS_VALORES & ")"
                Else
                    XLS_SQL = "UPDATE " & XLS_TABELA & " SET " & XLS_ALTERACAO & " WHERE " & XLS_CONDICAO
                End If
            End If
            
            Conexao.Execute XLS_SQL
        Next
    End If
End Sub


'Verifica se o registro está cadastrado no vetor
Function funExisteChave(Numchaves As Byte, coluna As Variant, valor As Variant, Vetor As XArray, Grid As TDBGrid) As Boolean
    Dim cont As Integer, Achou As Boolean
    Dim x As Integer, y As Integer
    
    funExisteChave = False
    
    While x <= Vetor.UpperBound(1)
        If Vetor(x, coluna(0)) = CStr(valor(0)) Then
            If Numchaves = 1 Then
                funExisteChave = True
                x = Vetor.UpperBound(1)
            Else
                Achou = True
                y = 1
                While Achou Or y <= Numchaves - 1
                    If coluna(y) <> valor(y) Then
                        Achou = False
                    End If
                    y = y + 1
                Wend
                If Achou Then
                    funExisteChave = True
                    x = Vetor.UpperBound(1)
                End If
            End If
        End If
        x = x + 1
    Wend
End Function
Sub subGravaVetor(Tabela As String, CamposTabela As String, Tipos As Variant, QtdCampos As Byte, Colunas As Variant, Complemento As String, Vetor As XArray, Grid As TDBGrid)
    Dim Valores As String
    Dim x As Integer, y As Integer
    'Tabela - nome da tabela onde serão inseridos os registros do grid
    'CamposTabela - string simples com os campos na respectiva ordem do grid
    
    'Tipos - vetor simples com os respectivos tipos de dados das colunas do grid
    'N: Número; D: Data; T: Texto
    
    'QtdCampos - quantidade de colunas do grid que serão gravadas
    'Colunas - vetor simples com o índice das respectivas colunas que serão gravadas
    'Complemento - Valores de campos fixos que deseje gravar e não existe no grid para evitar
    'repetição Ex.: Todos os campos possuem o mesmo código de serviço
    
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    If Vetor(0, 0) <> Empty Then
        If Complemento <> "" Then
            Complemento = ", " & Complemento
        End If
        
        For x = 0 To Vetor.UpperBound(1)
            Valores = ""
            For y = 0 To QtdCampos - 1
                If Valores <> "" Then
                    If Tipos(y) = "N" Then
                        'Valores = Valores & ", " & Str(Vetor(x, Colunas(y)))
                        'Qdo o valor é "" ocorre Type Mismatch com o uso da função Str
                        Valores = Valores & ", " & FunNuloBancoVal(Vetor(x, Colunas(y)))
                    ElseIf Tipos(y) = "T" Then
                        Valores = Valores & ", '" & FunTrataString(Vetor(x, Colunas(y))) & "'"
                    ElseIf Tipos(y) = "D" Then
                        Valores = Valores & "," & FunNuloData(Vetor(x, Colunas(y)), NomeSgbd) & ""
                    End If
                Else
                    If Tipos(y) = "N" Then
                        'Valores = Valores & ", " & Str(Vetor(x, Colunas(y)))
                        'Qdo o valor é "" ocorre Type Mismatch com o uso da função Str
                        Valores = FunNuloBancoVal(Vetor(x, Colunas(y)))
                    ElseIf Tipos(y) = "T" Then
                        Valores = "'" & Vetor(x, Colunas(y)) & "'"
                    ElseIf Tipos(y) = "D" Then
                        Valores = "" & FunNuloData(Vetor(x, Colunas(y)), NomeSgbd) & ""
                    End If
                End If
            Next
    
            Conexao.Execute "INSERT INTO " & Tabela & " (" & CamposTabela & ") VALUES (" & Valores & Complemento & ")"
        Next
    End If
End Sub
'Retorna a posição do elemento no vetor
Function funRetornaPos(coluna As Variant, valor As Variant, Vetor As XArray) As Integer
    Dim x As Integer
    funRetornaPos = False
    
    While x <= Vetor.UpperBound(1)
        If Vetor(x, coluna(0)) = CStr(valor(0)) Then
                funRetornaPos = x
        End If
        x = x + 1
    Wend
End Function
'Essa função atualiza os valores do vetor nos seus respectivos campos, em uma tabela do banco de dados
Sub subAtualizaTabela(Tabela As String, CamposTabela As Variant, Tipos As Variant, QtdCampos As Byte, Colunas As Variant, Complemento As String, Vetor As XArray, Grid As TDBGrid, NumRest As Integer, campos As Variant, Chaves As Variant, PosChave As Variant, NumCh As Byte, CamposAlt As Integer)
    Dim Valores As String
    Dim x As Integer, y As Integer, w As Integer, p As Integer
    Dim Restricao As String
    Dim VlChave As Variant
    
    'Tabela - nome da tabela onde serão inseridos os registros do grid
    'CamposTabela - vetor simples com os campos que serão atualizados, na respectiva ordem do grid
    
    'Tipos - vetor simples com os respectivos tipos de dados das colunas do grid
    'N: Número; D: Data; T: Texto
    
    'QtdCampos - quantidade de colunas do grid que serão gravadas
    
    'Colunas - vetor simples com o índice das respectivas colunas que serão gravadas
    'Complemento - Valores de campos fixos que deseje gravar e não existe no grid para evitar
    'repetição Ex.: Todos os campos possuem o mesmo código de serviço
    
    
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    
    'NumRes - número de restrições que serão feitas na tabela
    'Campos - vetor simples contendo os nomes dos campos chaves
    'Chaves - vetor simples contendo os valores dos campos chaves, se o valor do campo chave estiver no grid, preencher com ""
    'PosChave - vetor simples com a posição das chaves no grid, caso elas estejam armazenadas no vetor
    'NumCH - número de chaves que estão armazenadas no grid/vetor
    'CamposAlt - contém o número de linhas que serão alteradas
    
    If Vetor(0, 0) <> Empty Then
        If Complemento <> "" Then
            Complemento = ", " & Complemento
        End If
        
        For x = 0 To CamposAlt
            Valores = ""
            For y = 0 To QtdCampos - 1
                If Valores <> "" Then
                    If Tipos(y) = "N" Then
                        Valores = Valores + ", " & CamposTabela(y) & "=" & Str(Vetor(x, Colunas(y)))
                    ElseIf Tipos(y) = "T" Then
                        Valores = Valores + ", " & CamposTabela(y) & "='" & Vetor(x, Colunas(y)) & "'"
                    ElseIf Tipos(y) = "D" Then
                        Valores = Valores + ", " & CamposTabela(y) & "='" & Format$(Vetor(x, Colunas(y)), "DD/MM/YY") & "'"
                    End If
                    
                Else
                    If Tipos(y) = "N" Then
                        Valores = CamposTabela(y) + "=" + Str(Vetor(x, Colunas(y)))
                    ElseIf Tipos(y) = "T" Then
                        Valores = CamposTabela(y) + "=" + "'" & Vetor(x, Colunas(y)) & "'"
                    ElseIf Tipos(y) = "D" Then
                        Valores = CamposTabela(y) + "=" + "'" & Format$(Vetor(x, Colunas(y)), "DD/MM/YY") & "'"
                    End If
                    
                End If
            Next
            
            w = 0
            p = 0
            While w < NumRest
                If Chaves(w) = "" Then
                    VlChave = Vetor(x, PosChave(p))
                    If NumCh > 1 Then
                        p = p + 1
                    End If
                Else
                    VlChave = Chaves(w)
                End If
                If w = 0 Then
                    Restricao = campos(w) + "=" & VlChave
                ElseIf w >= 1 Then
                    Restricao = Restricao + " AND " + campos(w) + "=" & VlChave
                End If
                w = w + 1
            Wend
            
            Conexao.Execute "UPDATE " & Tabela & " SET " & Valores + Complemento + " WHERE " + Restricao & ""
        Next
    End If
End Sub
'Insere um regitro no vetor
Sub subInsereRegistroVetor(VetorCampos As Variant, Vetor As XArray, Grid As TDBGrid)  'Carrega o vetores
    Dim cont As Integer
    Dim x As Integer
    
    'VetorCampos - vetor simples com os valores a inserir na respectiva ordem do grid
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    
    If Vetor.UpperBound(1) <> 0 Or Vetor(0, 0) <> Empty Then
        Vetor.Insert 1, Vetor.UpperBound(1) + 1
    End If
    cont = Vetor.UpperBound(1)
    For x = 0 To Grid.Columns.Count - 1
        subSetaVal Vetor, cont, x, FunNulo(VetorCampos(x))
    Next
    With Grid
        .ReBind
    End With
End Sub
'Altera registros do vetor
Sub subAlteraRegistroVetor(VetorCampos As Variant, Vetor As XArray, Grid As TDBGrid, cont As Integer)
    'Dim cont As Integer
    Dim x As Integer
    
    
    
    'VetorCampos - vetor simples com os valores a inserir na respectiva ordem do grid
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    
    'cont = Vetor.UpperBound(1)
    For x = 0 To Grid.Columns.Count - 1
        subSetaVal Vetor, cont, x, FunNulo(VetorCampos(x))
    Next
    With Grid
        .ReBind
    End With
End Sub
'Carrega o valor da variavel campo no vetor, na posicao determinada pela linha x e coluna y
Sub subSetaVal(Vetor As XArray, x As Integer, y As Integer, Campo As String)

    If Not IsNull(Campo) Then
        Vetor(x, y) = Campo
    End If
End Sub
Sub subCarregaVetor(Sql As String, VetorCampos As Variant, Vetor As XArray, Grid As TDBGrid)   'Carrega o vetores
'Carrega os valores no vetor
  Dim ResAvaliacao As adodb.Recordset
  Dim cont As Integer
  Dim x As Integer
  
  'sql - string sql com a tabela ou consulta para carregar o vetor e o grid
  'VetorCampos - vetor simples com os campos na respectiva ordem do grid
  'Vetor - vetor que será preenchido
  'Grid - grid que será preenchido
  
  SubQOpenRecordset ResAvaliacao, Sql, Estatico
  Vetor.ReDim 0, 0, 0, Grid.Columns.Count - 1
  Vetor.Clear
  Grid.Array = Vetor
  If Not (ResAvaliacao.EOF And ResAvaliacao.BOF) Then
      ResAvaliacao.MoveFirst
      Do While Not ResAvaliacao.EOF
          If Vetor.UpperBound(1) <> 0 Or Vetor(0, 0) <> Empty Then
              Vetor.Insert 1, Vetor.UpperBound(1) + 1
          End If
          cont = Vetor.UpperBound(1)
          For x = 0 To Grid.Columns.Count - 1
              If VetorCampos(x) <> "" Then
                  subSetaVal Vetor, cont, x, FunNulo(ResAvaliacao(VetorCampos(x)))
              End If
          Next
          ResAvaliacao.MoveNext
      Loop
  End If
  Grid.ReBind
  ResAvaliacao.Close
  Set ResAvaliacao = Nothing
End Sub


Sub subCarregaVetorSemGrid(Sql As String, XVO_DESTINO As XArray)      'Carrega o vetores
'    Dim XVO_DESTINO As New XArray
    Dim XLO_RS As adodb.Recordset
    Dim j As Integer
  
    
    SubQOpenRecordset XLO_RS, Sql, Estatico
    'DIMENSIONA O VETOR COM 0 LINHAS E QUANTIDADE DE COLUNAS TANTO QUANTO O QUE VEIO NA CONSULTA
    
    'XVO_DESTINO.ReDim lowerRow, upperRow, lowerColumn, upperColumn
    XVO_DESTINO.ReDim 0, 0, 0, XLO_RS.Fields.Count
    XVO_DESTINO.Clear
    
    If Not (XLO_RS.EOF And XLO_RS.BOF) Then
        XLO_RS.MoveFirst
        
        Do While Not XLO_RS.EOF
            'INSERE UMA LINHA NO VETOR
            XVO_DESTINO.Insert 1, XVO_DESTINO.UpperBound(1) + 1
            
            'PARA CADA COLUNA DA LINHA DO RECORDSET
            For j = 0 To XLO_RS.Fields.Count - 1
                'SE A LINHA NAO FOR NULA
                If Not IsNull(XLO_RS(j)) Then
                    'ATRIBUE O VALOR DO RECORDSET NO ARRAY
                    XVO_DESTINO(XVO_DESTINO.UpperBound(1), j) = FunNulo(XLO_RS(j))
                End If
            Next
            XLO_RS.MoveNext
        Loop
    End If
    
  XLO_RS.Close
End Sub


'Calcula a soma dos valores de uma determinada coluna do vetor
Function funSomaColuna(Vetor As XArray, coluna As Byte) As String
    Dim row As Long, Col As Long
    Dim Acumulador As Double
    
    Acumulador = 0
    For row = Vetor.LowerBound(1) To Vetor.UpperBound(1)
        If IsNumeric(Vetor(row, coluna)) Then
            Acumulador = Acumulador + CDbl(Vetor(row, coluna))
        End If
    Next row
    funSomaColuna = Format$(Acumulador, "standard")
End Function
'Limpa os registros do vetor, caso ele não esteja vazio
Sub subLimpaVetor(Vetor As XArray, Grid As TDBGrid)

    While Vetor(0, 0) <> Empty
        Grid.MoveFirst
        subRemoveRegistroVetor Vetor, Grid
    Wend
End Sub
'Remove os registros do vetor
Sub subRemoveRegistroVetor(Vetor As XArray, Grid As TDBGrid)
    Dim cont As Integer
    Dim x As Integer
    
    If Grid.ApproxCount = 1 Then
        For x = 0 To Grid.Columns.Count - 1
            Vetor(0, x) = Empty
        Next
    Else
        Grid.Delete
        If Grid.EOF Then
            Grid.MoveLast
        End If
    End If
    With Grid
        .ReBind
    End With
End Sub



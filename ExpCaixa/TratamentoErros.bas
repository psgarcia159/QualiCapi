Attribute VB_Name = "TratamentoErros"
Option Explicit

'Declara constantes a serem usadas como parâmetro _
'da chamada  da função de tratamento de erros.

Global Const ComMensagem = True
Global Const SemMensagem = False

'Declara as contantes de Tratamento de Erros

Global Const ErrOutros = 1
Global Const ErrFalhaConexao = 2
Global Const ErrChaveDuplicada = 3
Global Const ErrRegistroAlterado = 4
Global Const ErrRegistroExcluido = 5
Global Const ErrTipoInvalido = 6
Global Const ErrRegistroBloqueado = 7
Global Const ErrChaveNula = 8
Global Const ErrCamposObrigatorios = 9
Global Const ErrExclusaoRelacionamento = 10
Global Const ErrTabelaRelacionada = 11
Global Const ErrRegistroJaAlterado = 12
Public Function funTrataErrosAccess(TipoTratamento As String) As Integer
    
    ' Esta rotina verifica os erros ocorridos
    ' durante a execução com o banco de dados Access
    
    Select Case Errors(0).Number
    
    
    Case 3201, 3200
        If TipoTratamento Then
            MsgBox "O registro não pode ser excluído pois está relacionado com outras tabelas.", vbCritical, "ATENÇÃO"
        End If
        funTrataErrosAccess = ErrTabelaRelacionada
    
    Case 3022
        If TipoTratamento Then
            MsgBox "Este registro já foi cadastrado.", vbCritical, "ATENÇÃO"
        End If
        funTrataErrosAccess = ErrChaveDuplicada
    
    Case 3024
        If TipoTratamento Then
            MsgBox "Banco de Dados " & NomeSgbd & " não encontrado", vbCritical, "ATENÇÃO"
        End If
        funTrataErrosAccess = ErrFalhaConexao
    
    Case 3006
        If TipoTratamento Then
            MsgBox "O Banco de Dados " & NomeSgbd & " está indisponível", vbCritical, "ATENÇÃO"
        End If
        funTrataErrosAccess = ErrFalhaConexao
    
    Case 3197
        If TipoTratamento Then
            MsgBox "Este registro foi alterado ou excluído por outro usuário", vbCritical, "ATENÇÃO"
        End If
        funTrataErrosAccess = ErrRegistroExcluido
    
    Case 3421
        If TipoTratamento Then
            MsgBox "Tipo de dado inválido", vbCritical, "ATENÇÃO"
        End If
        funTrataErrosAccess = ErrTipoInvalido
    
    Case 3058
        If TipoTratamento Then
            MsgBox "Chave primária ou índice com valor nulo", vbCritical, "ATENÇÃO"
        End If
        funTrataErrosAccess = ErrChaveNula
    
    Case 3314
        If TipoTratamento Then
            MsgBox "Favor preencher campos obrigatórios", vbCritical, "ATENÇÃO"
        End If
        funTrataErrosAccess = ErrCamposObrigatorios
    
    Case 3315
        If TipoTratamento Then
            MsgBox "Favor preencher campos obrigatórios", vbCritical, "ATENÇÃO"
        End If
        funTrataErrosAccess = ErrCamposObrigatorios
    
      Case Else 'Outros erros não listados
        If TipoTratamento Then
            MsgBox "Erro nº " & Err & ": " & Err.Description, vbCritical, "ATENÇÃO"
        End If
        funTrataErrosAccess = ErrOutros
    
    End Select

End Function

Public Function funTrataErrosSqlServer(TipoTratamento As String) As Integer
   
   ' Esta rotina verifica os erros ocorridos
   ' durante a execução com o banco de dados SQL Server
   
   Dim ErroOdbc As String
   Dim nuErro As Integer
   Dim dsSQLstate As String
   Dim Descricao As String
   
'======================================================================================================================
' TRATAMENTO DE ERRO PARA ADO
    Select Case Conexao.Errors(0).NativeError
        Case 2627
            If TipoTratamento Then
                MsgBox "Não é possivel concluir a gravação." & Chr(10) & _
                       "A tabela já possui registros cadastrados", _
                       vbCritical + vbApplicationModal + vbOKOnly, "ATENÇÃO"
            End If
            funTrataErrosSqlServer = ErrTabelaRelacionada
        Case -2147467262
            If TipoTratamento Then
                MsgBox "Arquivo não Encontrado.", vbCritical + vbOKOnly, "ATENÇÃO"
            End If
            funTrataErrosSqlServer = ErrTabelaRelacionada
        Case Else   'quando for um novo erro
            If TipoTratamento Then
                MsgBox "Ocorreu o seguinte Erro" & Chr(10) & Chr(10) & _
                       "Erro nº..:" & Err.Number & Chr(10) & Chr(10) & _
                       "Descrição:" & Err.Description, vbCritical, "ATENÇÃO"
            End If
            funTrataErrosSqlServer = ErrOutros
    End Select
'   '* Erro no RDO
'   If rdoEngine.rdoErrors.Count <> 0 Then
'
'    If (rdoEngine.rdoErrors(0).SQLRetcode <> rdSQLSuccess And _
'        rdoEngine.rdoErrors(0).SQLRetcode <> rdSQLSuccessWithInfo) Then
'
'        '* Erro na ODBC
'
'            dsSQLstate = rdoEngine.rdoErrors(0).SQLState
'            nuErro = rdoEngine.rdoErrors(0).Number
'            Descricao = rdoEngine.rdoErrors(0).Description
'
'        Select Case dsSQLstate
'
'        Case "01000"
'
'            '* Chave duplicada ou tentativa de excluir registro pai;
'            ' comando abortado pelo SQL Server
'
'            If (nuErro = 3621) Then
'
'                ' Tentativa de excluir um registro que é chave primária em
'                ' outra tabela
'
'                If InStr(rdoEngine.rdoErrors(1).Description, "DELETE") <> 0 Then
'                    If TipoTratamento Then
'                        MsgBox "Este registro não pode ser excluído, pois existem tabelas a ele associadas ", vbCritical, "ATENÇÃO"
'                    End If
'                    funTrataErrosSqlServer = ErrExclusaoRelacionamento
'                    rdoErrors.Clear
'                End If
'
'                ' Tentativa de alterar a chave primária do registro para
'                ' um valor já existente na tabela
'
'                If InStr(rdoEngine.rdoErrors(1).Description, "UPDATE") <> 0 Or _
'                    rdoEngine.rdoErrors(1).Number = 2627 Then
'                    If TipoTratamento Then
'                        MsgBox "Este registro já foi cadastrado.", vbCritical, "ATENÇÃO"
'                    End If
'                    funTrataErrosSqlServer = ErrChaveDuplicada
'                    rdoErrors.Clear
'                End If
'
'            Else 'Outros erros não listados
'
'                If TipoTratamento Then
'                    MsgBox "Erro nº " & nuErro & ": " & Descricao, vbCritical, "ATENÇÃO"
'                End If
'                funTrataErrosSqlServer = ErrOutros
'
'            End If
'
'        Case "01S03"
'            '* Tentativa de atualizar ou excluir registro que
'            '  ja foi excluido por outro usuario
'            If TipoTratamento Then
'                MsgBox "Este registro foi excluído por outro usuário", vbCritical, "ATENÇÃO"
'            End If
'            funTrataErrosSqlServer = ErrRegistroExcluido
'
'        Case "37000"
'            '* Tentativa de atualizar registro que
'            '  ja foi alterado por outro usuario
'            If (nuErro = 16934) Then
'                '* Tentativa de atualizar ou excluir registro que
'                '  ja foi excluido por outro usuario
'                If TipoTratamento Then
'                    MsgBox "Este registro foi alterado por outro usuário", vbCritical, "ATENÇÃO"
'                End If
'                funTrataErrosSqlServer = ErrRegistroAlterado
'            Else 'Outros erros não listados
'                If TipoTratamento Then
'                    MsgBox "Erro nº " & nuErro & ": " & Descricao, vbCritical, "ATENÇÃO"
'                End If
'                funTrataErrosSqlServer = ErrOutros
'            End If
'
'          Case "S1T00"
'            '* A tabela esta bloqueada por outro usuario,
'            '  nao conseguiu criar o ResultSet
'            If TipoTratamento Then
'                MsgBox "A tabela esta bloqueada por outro usuario. Tente mais tarde.", vbCritical, "ATENÇÃO"
'            End If
'            funTrataErrosSqlServer = ErrChaveDuplicada
'
'        Case Else
'            '* Outro erro ODBC
'            If TipoTratamento Then
'                MsgBox "Erro nº " & nuErro & ": " & Descricao, vbCritical, "ATENÇÃO"
'            End If
'            funTrataErrosSqlServer = ErrOutros
'        End Select
'
'    Else
'        'O tratamento abaixo foi comentado para evitar uma repetição do tratamento de erros
'        'no grid, pois este é feito também pelo controle de dados
'        '* Erro de Execucao do programa
'         'MsgBox "Erro nº " & Err & ": " & Err.Description, vbCritical, "ATENÇÃO"
'    End If
'
'End If
    
End Function
Public Function funTrataErros(TipoTratamento As String) As Integer

    'Verifica qual o tipo de acesso que está sendo utilizado

    Select Case NomeSgbd
    Case "Access"
            funTrataErros = funTrataErrosAccess(TipoTratamento)
    Case "Sql Server"
            funTrataErros = funTrataErrosSqlServer(TipoTratamento)
    End Select

End Function


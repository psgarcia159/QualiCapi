Attribute VB_Name = "TratamentoErros"
Option Explicit

'Declara constantes a serem usadas como parâmetro 'da chamada  da função de tratamento de erros.

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
    
    Select Case Conexao.Errors(0).Number
    
    
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

Public Sub subSetTimeouts(ByVal lngComponentBusy As _
Long, ByVal lngRequestPending As Long)
    App.OleServerBusyTimeout = lngComponentBusy
    App.OleRequestPendingTimeout = lngRequestPending
End Sub

Public Sub subResetTimeouts()
    App.OleServerBusyTimeout = 10000
    App.OleRequestPendingTimeout = 5000
End Sub

Public Function funTrataErrosSqlServer(TipoTratamento As String) As Integer
   
   ' Esta rotina verifica os erros ocorridos
   ' durante a execução com o banco de dados SQL Server
   
   Dim ErroOdbc As String
   Dim nuErro As Integer
   Dim dsSQLstate As String
   Dim Descricao As String
   
'======================================================================================================================
' TRATAMENTO DE ERRO PARA ADO
    If Conexao.Errors.Count > 0 Then
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
    Else
      MsgBox "Ocorreu o seguinte Erro" & Chr(13) & _
        Err.Number & " - " & Err.Description, vbCritical, "ATENÇÃO"
      funTrataErrosSqlServer = ErrOutros
    End If

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


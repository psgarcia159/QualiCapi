Attribute VB_Name = "TratamentoSql"
Option Explicit
Global SQLAtivo As String    'Atualizar o SQL Ativo após modificações de filtro e ordem
Public Function funTrataSql(sql As String) As String
                   
    ' funTrataSql tem como função verificar o banco de dados utilizado
    ' para que seja feito a conversão da cláusula SQL para o
    ' a sintaxe utilizada por este banco.
    
    ' O código abaixo Converte as aspas duplas por aspas simples para
    ' padronizar futuras pesquisas por cadeias de caracteres.
    
    
    'Verifica Banco de Dados
    
    Select Case NomeSgbd
    Case "Access"
        funTrataSql = sql
    Case "Sql Server"
        funTrataSql = funTrataSql_SQLServer(sql)
    End Select
    
End Function
Public Function funTrataSql_SQLServer(sql As String) As String
    
    'Tratamento da cláusula SQL para compatibilizar com o
    'banco de dados SQL Server
    
    '1º Caso: Cláusula "Like"
    '              Caracteres coringa devem ser: % para um ou mais caracteres,
    '                                                            _ para um caracter
    subConverteLike sql, "*", "%"
    subConverteLike sql, "?", "_"
        
    funTrataSql_SQLServer = sql
    
End Function
Public Sub subConverteLike(sql As String, CoringaAtual As String, CoringaNovo As String)
    
    ' Substitui o caracter coringa da cadeia de caracteres que
    ' segue a cláusula Like, pelo novo caracter coringa.
    
    Dim Posicao, Inicio, Fim As Integer
    Dim StrLike As String
    
    
    Posicao = InStr(UCase(sql), "LIKE")
    
    While Posicao <> 0 'Enquanto houver cláusula Like a ser tratada, faça:
        
        Posicao = Posicao + 1
        
        'Encontra o início da cadeia de caracteres da cláusula Like
        While Mid(sql, Posicao, 1) <> "'"
            Posicao = Posicao + 1
        Wend
        Posicao = Posicao + 1
        Inicio = Posicao
        
        ' Encontra o fim da cadeia de caracteres da cláusula Like
        ' e armazena em StrLike
        While Mid(sql, Posicao, 1) <> "'"
            Posicao = Posicao + 1
        Wend
        Fim = Posicao
        StrLike = Mid(sql, Inicio, Fim - Inicio)
        
        ' Substitui os caracterer coringa atual pelo novo
        subConverteCaracter StrLike, CoringaAtual, CoringaNovo
                
        ' Atualiza a cadeia SQL com as alterações realizadas
        Mid(sql, Inicio, Fim) = StrLike
                   
        Posicao = InStr(Posicao, UCase(sql), "LIKE")
                   
    Wend

End Sub
Public Sub subConverteCaracter(Cadeia As String, CaracterAtual As String, CaracterNovo As String)
        
        'Converte o "Caracter Atual" pelo novo caracter na cadeia
        
        While InStr(Cadeia, CaracterAtual) <> 0
            Mid(Cadeia, InStr(Cadeia, CaracterAtual)) = CaracterNovo
        Wend
        
End Sub


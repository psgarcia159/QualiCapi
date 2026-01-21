Attribute VB_Name = "ModuloLog"
'Enumarator com as operações que terão log
Enum ENUM_OPERACOES

    'QualiFin
    ALTERAR
    ALTERAR_DUPLICATA
    ALTERAR_DESCONTO
    ALTERAR_DISTRIBUICAO
    BAIXA_PARCIAL  'Fazer pagamento parcial de uma duplicata
    CADASTRAR
    CADASTRAR_DUPLICATA
    CADASTRAR_DESCONTO
    CADASTRAR_DISTRIBUICAO
    CANCELAR 'Cancelar uma conta a pagar (nota fiscal) ou duplicata compromissada
    CANCELAR_ESTORNO
    DEVOLVER 'Devolução do valor total ou parcial de uma conta a pagar compromissada ou realizada: gera uma conta a receber realizada
    EFETIVAR 'Tornar uma duplicata realizada. Pode ser pela rotina de duplicatas ou de pagamentos.
    EFETIVAR_CONSOLIDADO
    ESTORNO 'Estorno de conta a pagar realizada.
    EXCLUIR
    EXCLUIR_DUPLICATA
    EXCLUIR_DESCONTO
    EXCLUIR_DISTRIBUICAO
    GLOSA 'Glosa de conta a receber compromissada.
    PENDENTE 'Cancelar o pagamento de uma duplicata, torná-la compromissada.
    RECOLHER 'Recolher impostos de notas fiscais.
    TRANSFERENCIA
    PRESTAR_CONTAS
    RESTITUIR_ADIANTAMENTO 'Restituição de adiantamento pela rotina de prestação de contas
    DEVOLVER_ADIANTAMENTO 'Devolução de adiantamento pela rotina de prestação de contas
    EXPORTACAO_FIN
    IMPORTACAO_FIN
    EXPORTACAO_CAPI
    IMPORTACAO_CAPI
    IMPORTAR_FOLHA_NORMAL
    IMPORTAR_FOLHA_CPF
    IMPORTAR_FOLHA_CR
    IMPORTAR_FOLHA_PIS_FGTS
    IMPORTACAO_FOLHA_FUNCIONARIOS
    
    'QualiCapi
    ALTERACAO_MULTIPLA
    ATIVAR_CONTRATO
    CANCELAR_REPASSE_TITULO
    COPIAR
    CESSAO_CONTRATO
    DISTRATAR_CONTRATO
    EXCLUSAO_MULTIPLA
    RECALCULAR
    REPASSE_TITULO
    TRANSFERENCIA_ENTRE_EMPRESAS
    PAGTO_TITULO_DTDEPOSITO
            
    'ALTERAR_EMPREENDIMENTO
    'ALTERAR_IMOVEL
    'ALTERAR_TITULO
    'CADASTRAR_EMPREENDIMENTO
    'CADASTRAR_IMOVEL
    'CADASTRAR_TITULO
    'CADASTRAR_PLANO
    'EXCLUIR_EMPREENDIMENTO
    'EXCLUIR_IMOVEL
    
    PAGAMENTO_TITULO
    CANCELAR_PAGAMENTO_TITULO
    
    EMITIR_BOLETO
End Enum

'Enumarator com as entidades que terão log
Enum ENUM_ENTIDADES
    
    'QualiFin
    IMPOSTO
    DISTRIBUICAO
    DUPLICATA
    MULTA_JURO
    NOTAFISCAL_RECEBER
    NOTAFISCAL_PAGAR
    FATURAS
    TRANSFERENCIA_BANCARIA
    TRANSFERENCIA_CUSTO
    PRESTACAO_CONTAS
    PAGAMENTO_CHEQUE
    PAGAMENTO_DEBITO_CREDITO
    PAGAMENTO_BORDERO
    IMPORTACAO_FOLHA
    
    'Exportação Bancária
    PAGAMENTO_ELETRONICO_ITAU
    PAGAMENTO_ELETRONICO_REAL
    PAGAMENTO_ELETRONICO_SANTANDER
    PAGAMENTO_ELETRONICO_SUDAMERIS
    PAGAMENTO_ELETRONICO_BRASIL
    PAGAMENTO_ELETRONICO_UNIBANCO
    PAGAMENTO_ELETRONICO_HSBC
    PAGAMENTO_ELETRONICO_CAIXA
    
    'QualiCapi
    CONTRATO_CAPI
    EMPREENDIMENTO_CAPI
    IMOVEL_CAPI
    PLANO_CAPI
    TITULO_CAPI
    PAGAMENTO_MULTIPLO_CAPI
    PAGAMENTO_UNICO_CAPI
    
End Enum

'Rotina para gravar log
'XLT_ROTINA: Nome da tela
'XLT_ORDEMROTINA: ordem da rotina na tabela Rotinas
Sub subRegistraLog(XLT_ROTINA As String, XLT_ORDEMROTINA As String, XLT_DESCRICAO As String)

    Dim XLO_COMANDO As New ADODB.Command 'Comando para stored procedure
    
    With XLO_COMANDO
        .CommandText = "sp_ins_Log"
        .CommandType = adCmdStoredProc
        Set XLO_COMANDO.ActiveConnection = ConexaoRelatorio
        'Cria automaticamente todos os parâmetros dentro do objeto Comando
        .Parameters.Refresh
    End With
  
    'Adiciona os parâmetros a procedure
    XLO_COMANDO.Parameters(1).Value = func_cd_operador
    XLO_COMANDO.Parameters(2).Value = XLT_ROTINA
    XLO_COMANDO.Parameters(3).Value = XLT_ORDEMROTINA
    XLO_COMANDO.Parameters(4).Value = XLT_DESCRICAO
     
    XLO_COMANDO.Execute
  
End Sub

'Rotina para criar descrição do log
'XLM_MATRIZLOG: contém todos os dados utilizados para criar a descrição do log. Col0: nome do campo; Col1: valores originais; Col2: valores atuais.
'XLT_OPERACAO: operação que será executada (inclusão, efetivar...).
'XLT_ENTIDADE: entidade que será alterada (nota fiscal, fatura, etc).
Function funCriaDescricaoLog(ByVal XLM_MATRIZLOG As Variant, XLT_OPERACAO As ENUM_OPERACOES, XLT_ENTIDADE As ENUM_ENTIDADES) As String
    Dim Cont As Integer
    Dim registroAnterior As Object
    Dim Descricao As String
     
    'Informa a empresa, a operação, a entidade e a identificação da entidade
    Descricao = "Emp: " & PEmpresa & "; OPERAÇÃO: " & funNomeOperacao(XLT_OPERACAO) & "; TELA: " & funNomeEntidade(XLT_ENTIDADE) & ";"
        
    'Descreve os campos apenas para alteração e exclusão.
    If XLT_OPERACAO = ALTERAR Or XLT_OPERACAO = ALTERAR_DESCONTO Or XLT_OPERACAO = ALTERAR_DISTRIBUICAO Or XLT_OPERACAO = ALTERAR_DUPLICATA Or XLT_OPERACAO = ALTERACAO_MULTIPLA Then
        'Preenche identificação da entidade antes de informar os campos alterados
        For Cont = 0 To 2
            Descricao = Descricao & " " + UCase(XLM_MATRIZLOG(Cont, 0)) & ": " & XLM_MATRIZLOG(Cont, 1) & ";"
        Next
        
        For Cont = 0 To (UBound(XLM_MATRIZLOG) - 1)
            If (XLM_MATRIZLOG(Cont, 1) <> XLM_MATRIZLOG(Cont, 2) And XLM_MATRIZLOG(Cont, 2) <> Empty) Then
                Descricao = Descricao & " " & UCase(XLM_MATRIZLOG(Cont, 0)) & " alterado de: " & UCase(XLM_MATRIZLOG(Cont, 1)) & " para: " & UCase(XLM_MATRIZLOG(Cont, 2)) & ";"
            End If
        Next
    Else
        For Cont = 0 To (UBound(XLM_MATRIZLOG) - 1)
            Descricao = Descricao & " " + UCase(XLM_MATRIZLOG(Cont, 0)) & ": " & XLM_MATRIZLOG(Cont, 1) & ";"
        Next
    End If
  
  funCriaDescricaoLog = Descricao
End Function

'Função para retornar o nome da operação como string
Function funNomeOperacao(enumerator As ENUM_OPERACOES)
    Select Case enumerator
        Case ALTERAR
            funNomeOperacao = "ALTERAR"
        Case ALTERAR_DESCONTO
            funNomeOperacao = "ALTERAR DESCONTO"
        Case ALTERAR_DISTRIBUICAO
            funNomeOperacao = "ALTERAR DISTRIBUICAO"
        Case ALTERAR_DUPLICATA
            funNomeOperacao = "ALTERAR DUPLICATA"
        Case BAIXA_PARCIAL
            funNomeOperacao = "BAIXA PARCIAL"
        Case CADASTRAR
            funNomeOperacao = "CADASTRAR"
        Case CADASTRAR_DESCONTO
            funNomeOperacao = "CADASTRAR DESCONTO"
        Case CADASTRAR_DISTRIBUICAO
            funNomeOperacao = "CADASTRAR DISTRIBUIÇÃO"
        Case CADASTRAR_DUPLICATA
            funNomeOperacao = "CADASTRAR DUPLICATA"
        Case CANCELAR
            funNomeOperacao = "CANCELAR"
        Case CANCELAR_ESTORNO
            funNomeOperacao = "CANCELAR ESTORNO FINANCEIRO"
        Case DEVOLVER
            funNomeOperacao = "DEVOLUCAO"
        Case EFETIVAR
            funNomeOperacao = "EFETIVAR"
        Case EFETIVAR_CONSOLIDADO
            funNomeOperacao = "EFETIVAR CONSOLIDADO"
        Case ESTORNO
            funNomeOperacao = "ESTORNO FINANCEIRO"
        Case EXCLUIR
            funNomeOperacao = "EXCLUIR"
        Case EXCLUIR_DESCONTO
            funNomeOperacao = "EXCLUIR DESCONTO"
        Case EXCLUIR_DISTRIBUICAO
            funNomeOperacao = "EXCLUIR DISTRIBUIÇÃO"
        Case EXCLUIR_DUPLICATA
            funNomeOperacao = "EXCLUIR DUPLICATA"
        Case GLOSA
            funNomeOperacao = "GLOSA"
        Case PENDENTE
            funNomeOperacao = "CANCELAR PAGAMENTO"
        Case RECOLHER
            funNomeOperacao = "RECOLHER"
        Case TRANSFERENCIA
            funNomeOperacao = "TRANSFERÊNCIA"
        Case PRESTAR_CONTAS
            funNomeOperacao = "PRESTAR CONTAS"
        Case RESTITUIR_ADIANTAMENTO
            funNomeOperacao = "RESTITUIR ADIANTAMENTO"
        Case DEVOLVER_ADIANTAMENTO
            funNomeOperacao = "DEVOLVER ADIANTAMENTO"
        Case EXPORTACAO_FIN
            funNomeOperacao = "EXPORTAÇÃO BANCÁRIA QUALIFIN"
        Case EXPORTACAO_CAPI
            funNomeOperacao = "EXPORTAÇÃO BANCÁRIA QUALICAPI"
        Case IMPORTACAO_FIN
            funNomeOperacao = "IMPORTAÇÃO BANCÁRIA QUALIFIN"
        Case IMPORTACAO_CAPI
            funNomeOperacao = "IMPORTAÇÃO BANCÁRIA QUALICAPI"
        Case IMPORTAR_FOLHA_NORMAL
            funNomeOperacao = "IMPORTAÇÃO DE FOLHA DE PAGAMENTO NORMAL"
        Case IMPORTAR_FOLHA_CPF
            funNomeOperacao = "IMPORTAÇÃO DE FOLHA DE PAGAMENTO POR CPF"
        Case IMPORTAR_FOLHA_CR
            funNomeOperacao = "IMPORTAÇÃO DE FOLHA DE PAGAMENTO POR CR"
        Case IMPORTAR_FOLHA_PIS_FGTS
            funNomeOperacao = "IMPORTAÇÃO DE FOLHA PIS/FGTS"
        Case IMPORTAR_FOLHA_FUNCIONARIOS
            funNomeOperacao = "IMPORTAÇÃO DE FUNCIONÁRIOS"
        Case ATIVAR_CONTRATO
            funNomeOperacao = "ATIVAR CONTRATO INATIVO"
        Case CESSAO_CONTRATO
            funNomeOperacao = "CESSÃO DE CONTRATO"
        Case TRANSFERENCIA_ENTRE_EMPRESAS
            funNomeOperacao = "TRANSFERÊNCIA ENTRE EMPRESAS"
        Case EXCLUSAO_MULTIPLA
            funNomeOperacao = "EXCLUSÃO MÚLTIPLA"
        Case RECALCULAR
            funNomeOperacao = "RECALCULAR"
        Case REPASSE_TITULO
            funNomeOperacao = "REPASSE DE TÍTULO"
        Case CANCELAR_REPASSE_TITULO
            funNomeOperacao = "CANCELAR REPASSE DE TÍTULO"
        Case ALTERACAO_MULTIPLA
            funNomeOperacao = "ALTERAÇÃO MÚLTIPLA"
        Case DISTRATAR_CONTRATO
            funNomeOperacao = "DISTRATO DE CONTRATO"
        Case COPIAR
            funNomeOperacao = "COPIAR"
        Case PAGTO_TITULO_DTDEPOSITO
            funNomeOperacao = "PAGTO DE TÍTULO COM DT DÉPOSITO MAIOR QUE DT PAGAMENTO"
        Case CANCELAR_PAGAMENTO_TITULO
            funNomeOperacao = "CANCELAMENTO DE PAGTO DE TÍTULO"
        Case PAGAMENTO_TITULO
            funNomeOperacao = "PAGAMENTO DE TÍTULO"
        Case EMITIR_BOLETO
            funNomeOperacao = "EMISSÃO DE BOLETO"
    End Select
End Function

Function funNomeEntidade(enumerator As ENUM_ENTIDADES)
    Select Case enumerator
        Case IMPOSTO
            funNomeEntidade = "DESCONTO"
        Case DISTRIBUICAO
            funNomeEntidade = "DISTRIBUICAO"
        Case DUPLICATA
            funNomeEntidade = "DUPLICATA"
        Case MULTA_JURO
            funNomeEntidade = "MULTA JURO"
        Case NOTAFISCAL_RECEBER
            funNomeEntidade = "CONTAS A RECEBER"
        Case NOTAFISCAL_PAGAR
            funNomeEntidade = "CONTAS A PAGAR"
        Case FATURAS
            funNomeEntidade = "FATURA"
        Case TRANSFERENCIA_BANCARIA
            funNomeEntidade = "TRANSFERÊNCIA BANCÁRIA"
        Case TRANSFERENCIA_CUSTO
            funNomeEntidade = "TRANSFERÊNCIA DE CUSTO"
        Case PRESTACAO_CONTAS
            funNomeEntidade = "PRESTAÇÃO DE CONTAS"
        Case PAGAMENTO_CHEQUE
            funNomeEntidade = "PAGAMENTO CHEQUE"
        Case PAGAMENTO_DEBITO_CREDITO
            funNomeEntidade = "PAGAMENTO DÉBITO/CRÉDITO EM CONTA"
        Case PAGAMENTO_BORDERO
            funNomeEntidade = "PAGAMENTO BORDERÔ"
        Case PAGAMENTO_ELETRONICO_BRADESCO
            funNomeEntidade = "PAGAMENTO ELETRÔNICO BANCO BRADESCO"
        Case PAGAMENTO_ELETRONICO_ITAU
            funNomeEntidade = "PAGAMENTO ELETRÔNICO BANCO ITAÚ"
        Case PAGAMENTO_ELETRONICO_REAL
            funNomeEntidade = "PAGAMENTO ELETRÔNICO BANCO REAL"
        Case PAGAMENTO_ELETRONICO_SANTANDER
            funNomeEntidade = "PAGAMENTO ELETRÔNICO BANCO SANTANDER"
        Case PAGAMENTO_ELETRONICO_SUDAMERIS
            funNomeEntidade = "PAGAMENTO ELETRÔNICO BANCO SUDAMERIS"
        Case PAGAMENTO_ELETRONICO_BRASIL
            funNomeEntidade = "PAGAMENTO ELETRÔNICO BANCO DO BRASIL"
        Case PAGAMENTO_ELETRONICO_UNIBANCO
            funNomeEntidade = "PAGAMENTO ELETRÔNICO BANCO UNIBANCO"
        Case PAGAMENTO_ELETRONICO_HSBC
            funNomeEntidade = "PAGAMENTO ELETRÔNICO BANCO HSBC"
        Case IMPORTACAO_FOLHA
            funNomeEntidade = "IMPORTAÇÃO DE FOLHA DE PAGAMENTO"
        Case EMPREENDIMENTO_CAPI
            funNomeEntidade = "EMPREENDIMENTOS"
        Case IMOVEL_CAPI
            funNomeEntidade = "IMÓVEIS"
        Case CONTRATO_CAPI
            funNomeEntidade = "CONTRATOS"
        Case TITULO_CAPI
            funNomeEntidade = "TÍTULOS"
    End Select
End Function


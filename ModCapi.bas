Attribute VB_Name = "ModCapi"
Option Explicit

Global XGB_RESULTADO As Boolean   'Indica se o usuário clicou em Gravar em um formulário moldal.
Global XGB_GERARESIDUO As Boolean 'Indica se o resíduo deve ser gerado ou não
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Sub SubContabilizarPromissoria(XLL_CODDUPLICATA As Long, XLL_FATURA As Long, XLL_CODCCORRENTE As Long, XLT_DOCREC As String, XLD_PAGTOREC As Date, XLB_CONSOLIDADO As Boolean)
 
    '********************DECRIÇÃO DA PROCEDURE***********************************
    'Contabilização das Notas Promissorias - Apenas Base Caixa
    '****************************************************************************
    
    '**************PARÂMETROS*************************
    'XLL_CODDUPLICATA (Código da Duplicata)
    'XLL_FATURA (Nº da fatura da Parcela)
    'XLL_CODCCORRENTE (Código da conta corrente)
    'XLT_DOCREC (Nº do doc pag/rec)
    'XLD_PAGTOREC (Data de PagtoRec da Parcela)
    'XLB_CONSOLIDADO (Indica se o pagto é consolidado)
    '*************************************************
    
    Dim XLI_CODLANC As Integer      'Código para a tabela de Lançamentos
    Dim XLI_LOTE As Integer         'Nº do Lote
    Dim XLI_IDENTIFICADOR As Integer 'Código do identificador
    Dim XLO_PARCELA As Object       'Recordset com os dados das parcelas
    Dim XLO_VLCONSOLIDADO As Object 'Recordset para obter o valor do pagto consolidado
    Dim XLO_PARCESPECIAL As Object  'Recordset para tipos de documentos como TB, DEV, REST etc.
    Dim XLO_CHECK As Object         'Recordset para  definir o lote
    Dim XLB_PROVISIONADO As Boolean 'Indica se o lançamento foi provisionado
    Dim XLT_SQL As String           'Prepara Select
    Dim XLT_TIPOPAGTO As String     'Tipo de Pagto/rec (se cheque, dinheiro...)
    Dim XLT_PCRDEBITO As String, XLT_PCRCREDITO As String         'C. Custo de débito e crédito
    Dim XLT_CONTADEBITO As String, XLT_CONTACREDITO As String     'C. Contábil de débito e crédito
    Dim XLT_CONTADEBITOCC As String, XLT_CONTACREDITOCC As String 'C. Contábil (Banco ou forn) de débito e crédito
    Dim XLT_HISTORICO As String     'Histórico Contábil
    Dim XLT_HISTCUSTO As String     'Histórico do Custos = financeiro
    Dim XLT_TIPO As String          'Tipo do lançamento(crédito ou débito)
    Dim XLF_VALOR As Double         'Valor do lançamento
    Dim XLF_VALORBANCO As Double    'Valor da parte do Banco
    Dim XLF_VALORDESC As Double     'Valor de Desconto
    Dim XLF_VALORDESCACRESC As Double  'Valor do desconto acrescimo
    Dim XLF_VLCONSOLIDADO As Double 'Total de valores consolidados
    Dim XLF_VALORDUPL As Double     'Valor da parcela
    Dim XLV_POSICAO As Variant      'Armazena a posição do grid
    
    'Pegar o número do Lote e Lancamento para cadastrar no mesmo lote do dia
    'Origem = C - Capi
    XLT_SQL = "SELECT * FROM Lancamentos " & _
                "WHERE lanc_dt_data = " & FunNuloData(XLD_PAGTOREC, NomeSgbd) & " AND " & _
                    "lanc_tx_origem LIKE 'C%' AND " & _
                    "empr_cd_empresa = " & PCodEmpresa & _
                 " ORDER BY lanc_nr_lote DESC, lanc_cd_lanc DESC"
    SubQOpenRecordset XLO_CHECK, XLT_SQL, Estatico
    If XLO_CHECK.EOF Then
        XLI_LOTE = (FunDefinirUltimoLote(XLD_PAGTOREC))
    Else
        XLI_LOTE = XLO_CHECK!lanc_nr_Lote
    End If
    XLO_CHECK.Close
    
    XLT_TIPOPAGTO = "5" 'Crédito em Conta
    'pegar o valor total para posterior contabilização da parte do banco
    XLT_SQL = "SELECT sum(dupl_vl_valor) as ValorChCons, " & _
            " count(dupl_cd_duplicata) as TotalCheques FROM Duplicatas " & _
            " where dupl_tx_tipopagrec ='" & XLT_TIPOPAGTO & "' and coco_cd_codigo = " & XLL_CODCCORRENTE & " " & _
            " and dupl_tx_docpagrec= '" & XLT_DOCREC & "'"
    SubQOpenRecordset XLO_VLCONSOLIDADO, XLT_SQL, Estatico
    XLF_VLCONSOLIDADO = XLO_VLCONSOLIDADO!ValorChCons
    XLO_VLCONSOLIDADO.Close
    
    XLT_SQL = "SELECT * FROM ConsFINExpContabilidade2 " & _
                    " WHERE coco_cd_codigo = " & XLL_CODCCORRENTE & "" & _
                    " AND dupl_tx_tipopagrec ='" & XLT_TIPOPAGTO & "'" & _
                    " AND dupl_tx_docpagrec= '" & XLT_DOCREC & "'" & _
                    " AND empr_cd_empresa = " & PCodEmpresa & ""
    SubQOpenRecordset XLO_PARCELA, XLT_SQL, Estatico
    
    If Not XLO_PARCELA.EOF Then
        XLT_CONTADEBITOCC = XLO_PARCELA!PlcoContaCC
        While Not XLO_PARCELA.EOF
            '**********************NOTA PROMISSÓRIA***********************************
            XLT_PCRCREDITO = XLO_PARCELA!cere_cd_Pcr ''XLO_PARCESPECIAL!cere_cd_Pcr
            XLT_CONTACREDITO = XLO_PARCELA!ContaRecDesp ''XLO_PARCESPECIAL!ContaRecDesp
            XLT_PCRDEBITO = ""
            XLT_CONTADEBITO = "" 'Debitar depois a soma total do dia
            If XGB_PQUALICONT Then 'possui contabilidade, o histórico pode ser maior
                XLT_HISTORICO = Mid(XLO_PARCELA!nofi_tx_historico + "-" + XLO_PARCELA!focl_tx_RazaoSocial + "-Doc: " + XLO_PARCELA!nofi_nr_documento, 1, 120)
            Else
                XLT_HISTORICO = Mid(XLO_PARCELA!nofi_tx_historico + "-" + XLO_PARCELA!focl_tx_RazaoSocial + "-Doc: " + XLO_PARCELA!nofi_nr_documento, 1, 60)
            End If
            XLT_HISTCUSTO = XLO_PARCELA!nofi_tx_historico
            XLI_IDENTIFICADOR = FunNuloVal(XLO_PARCELA!iden_cd_Projeto)
            XLD_PAGTOREC = XLO_PARCELA!dupl_dt_pagrec
            XLF_VALORDESC = XLO_PARCELA!desconto + XLO_PARCELA!Acrescimo
            XLF_VALORDUPL = XLF_VALORDESC + XLO_PARCELA!distvalor
            XLF_VALORBANCO = 0
            
            XLI_CODLANC = funDefinirCodigoLanc(XLO_PARCELA!dupl_dt_pagrec, XLI_LOTE)
            
            XLT_SQL = "INSERT INTO Lancamentos(lanc_nr_Lote, lanc_cd_Lanc, " & _
                            " lanc_dt_Data, " & _
                            " cere_cd_pcrCreditar, cere_cd_pcrDebitar, " & _
                            " plco_cd_Credito, plco_cd_Debito," & _
                            " lanc_tx_historico, lanc_vl_valor," & _
                            " empr_cd_empresa, lanc_tx_DescCustos, " & _
                            " iden_cd_Projeto, lanc_cd_OrigemFin, lanc_tx_Origem) " & _
                        " VALUES (" & XLI_LOTE & ", " & XLI_CODLANC & "," & _
                            "" & FunNuloData(XLD_PAGTOREC, NomeSgbd) & "," & _
                            "" & FunNuloBanco(XLT_PCRCREDITO) & "," & FunNuloBanco(XLT_PCRDEBITO) & "," & _
                            "" & FunNuloBanco(XLT_CONTACREDITO) & "," & FunNuloBanco(XLT_CONTADEBITO) & "," & _
                            "'" & FunNulo(FunTrataString(XLT_HISTORICO)) & "'," & FunTrataFloat(XLF_VALORDUPL) & "," & _
                            "" & PCodEmpresa & ",'" & FunNulo(FunTrataString(XLT_HISTCUSTO)) & "'," & _
                            "" & FunNuloBanco(XLI_IDENTIFICADOR) & "," & XLL_CODDUPLICATA & ",'C')"
            Conexao.Execute XLT_SQL
            
            XLO_PARCELA.MoveNext
        Wend
        XLO_PARCELA.Close

        '*********************PARTE DO BANCO DA NOTA PROMISSÓRIA********************
        XLT_HISTORICO = ""
        XLT_CONTACREDITO = ""
        XLT_PCRCREDITO = ""
        XLT_PCRDEBITO = ""
        XLT_HISTORICO = "Total diário de Notas Promissória " + Format(CStr(XLD_PAGTOREC), "dd/MM/yy")
                
        XLT_SQL = "SELECT * FROM Lancamentos " & _
                    " WHERE plco_cd_Debito = '" & XLT_CONTADEBITOCC & "'" & _
                    " AND lanc_dt_Data =" & FunNuloData(XLD_PAGTOREC, NomeSgbd) & "" & _
                    " AND empr_cd_empresa = " & PCodEmpresa & ""
        SubQOpenRecordset XLO_PARCELA, XLT_SQL, Dinamico
        If XLO_PARCELA.EOF Then
            XLO_PARCELA.AddNew
            XLI_CODLANC = funDefinirCodigoLanc(XLD_PAGTOREC, XLI_LOTE)
            XLO_PARCELA!lanc_cd_Lanc = XLI_CODLANC
        End If
        
        XLO_PARCELA!lanc_nr_Lote = XLI_LOTE
        XLO_PARCELA!lanc_dt_Data = XLD_PAGTOREC
        XLO_PARCELA!cere_cd_pcrCreditar = Null 'XLT_PCRCREDITO
        XLO_PARCELA!cere_cd_pcrDebitar = Null 'XLT_PCRDEBITO
        XLO_PARCELA!plco_cd_Credito = Null 'XLT_CONTACREDITO
        XLO_PARCELA!plco_cd_Debito = XLT_CONTADEBITOCC
        XLO_PARCELA!lanc_tx_historico = XLT_HISTORICO
        XLO_PARCELA!lanc_vl_valor = CDbl(XLO_PARCELA!lanc_vl_valor) + XLF_VLCONSOLIDADO
        XLO_PARCELA!empr_cd_Empresa = PCodEmpresa
        XLO_PARCELA!lanc_tx_DescCustos = XLT_HISTCUSTO
        XLO_PARCELA!iden_cd_Projeto = XLI_IDENTIFICADOR
        XLO_PARCELA!lanc_cd_OrigemFin = XLL_CODDUPLICATA
        XLO_PARCELA!lanc_tx_Origem = "C"
                                        
        XLO_PARCELA.Update
        XLO_PARCELA.Close
    End If
End Sub




'Retorna o TipoPag
Function funGeraTipoPag(XLT_CORRIGE As String, XLF_VALORPAGO As Double, XLF_TOTALTITULO As Double, XLF_VLCORRIGIDO As Double, XLD_VENCIMENTO As Date, XLD_PAGAMENTO As Date) As String
    If XLT_CORRIGE = "1" Then
        If CDbl(XLF_VALORPAGO) < XLF_TOTALTITULO Then 'Parcial
            If XLD_VENCIMENTO = XLD_PAGAMENTO Then
                funGeraTipoPag = "2" 'Pagto Parcial Corrigido no vencimento
            ElseIf XLD_VENCIMENTO < XLD_PAGAMENTO Then
                funGeraTipoPag = "4" 'Pagto Parcial Corrigido vencido
            ElseIf XLD_VENCIMENTO > XLD_PAGAMENTO Then
                funGeraTipoPag = "6" 'Pagto Parcial Corrigido antecipado
            End If
        Else
            If XLD_VENCIMENTO = XLD_PAGAMENTO Then
                funGeraTipoPag = "1" 'Pagto Total Corrigido no vencimento
            ElseIf XLD_VENCIMENTO < XLD_PAGAMENTO Then
                funGeraTipoPag = "3" 'Pagto Total Corrigido vencido
            ElseIf XLD_VENCIMENTO > XLD_PAGAMENTO Then
                funGeraTipoPag = "5" 'Pagto Total Corrigido antecipado
            End If
        End If
    Else
        If CDbl(XLF_VALORPAGO) < XLF_TOTALTITULO Then 'Parcial
            If XLD_VENCIMENTO = XLD_PAGAMENTO Then
                funGeraTipoPag = "8" 'Pagto Parcial sem Correçao no vencimento
            ElseIf XLD_VENCIMENTO < XLD_PAGAMENTO Then
                funGeraTipoPag = "10" 'Pagto Parcial sem Correçao vencido
            ElseIf XLD_VENCIMENTO > XLD_PAGAMENTO Then
                funGeraTipoPag = "12" 'Pagto Parcial sem Correçao antecipado
            End If
        Else
            If XLD_VENCIMENTO = XLD_PAGAMENTO Then
                funGeraTipoPag = "7" 'Pagto Total sem Correçao no vencimento
            ElseIf XLD_VENCIMENTO < XLD_PAGAMENTO Then
                funGeraTipoPag = "9" 'Pagto Total sem Correçao vencido
            ElseIf XLD_VENCIMENTO > XLD_PAGAMENTO Then
                funGeraTipoPag = "11" 'Pagto Total sem Correçao antecipado
            End If
        End If
    End If
End Function

Public Function FunCotacao(XLD_DATA As Date, XLL_CODMOEDA As Long, XLT_INDMOEDA As String, XLF_COTACAO As Double) As Boolean
    
'PARAMETROS
    'XLD_DATA - DATA DA COTACAO DESEJADA
    'XLL_CODMOEDA - CÓDIGO DA MOEDA
    'XLT_INDMOEDA - INDEXAÇÃO DA MOEDA
    
'VARIAVEIS
    Dim XLO_COTACAO     As ADODB.Recordset
    Dim XLT_SQL         As String
    Dim XLB_Cotacao     As Boolean
    Dim XLT_ORDEM       As String
    Dim XLL_MOUSE       As Long
    
    XLT_ORDEM = Ordem
    
'PEGA FORMATO DO MOUSE
    XLL_MOUSE = Screen.MousePointer
    
Inicio_Funcao:
    'seleciona cotacao
    XLB_Cotacao = False
    If XLT_INDMOEDA = "D" Then
        XLT_SQL = "SELECT * FROM ConsGENCotacaoMoeda " & _
                  "WHERE cota_dt_DataCotacao =" & FunNuloData(XLD_DATA, NomeSgbd) & _
                  "  AND moed_cd_CodMoeda = " & XLL_CODMOEDA & _
                 " ORDER BY cota_dt_DataCotacao"
        SubQOpenRecordset XLO_COTACAO, XLT_SQL, Estatico
        If Not XLO_COTACAO.EOF Then
            XLB_Cotacao = True
            XLF_COTACAO = XLO_COTACAO!cota_vl_Valor
        End If
    ElseIf XLT_INDMOEDA = "M" Then
        XLT_SQL = "SELECT * FROM ConsGENCotacaoMoeda " & _
                  "WHERE cota_dt_DataCotacao <=" & FunNuloData(XLD_DATA, NomeSgbd) & _
                  " AND moed_cd_CodMoeda = " & XLL_CODMOEDA & _
                  " ORDER BY cota_dt_DataCotacao"
        SubQOpenRecordset XLO_COTACAO, XLT_SQL, Estatico
        If XLO_COTACAO.RecordCount > 0 Then
            XLO_COTACAO.MoveLast
            If Month(XLO_COTACAO!cota_dt_DataCotacao) = Month(XLD_DATA) And Year(XLO_COTACAO!cota_dt_DataCotacao) = Year(XLD_DATA) Then
                XLB_Cotacao = True
                XLF_COTACAO = XLO_COTACAO!cota_vl_Valor
            End If
        End If
        XLD_DATA = "01/" + Trim(Str(Month(XLD_DATA))) + "/" + Trim(Str(Year(XLD_DATA)))
    End If
    XLO_COTACAO.Close
    If Not XLB_Cotacao Then
        Screen.MousePointer = vbDefault
        If MsgBox("Indexador não esta Atualizado nesta data: " & Str(XLD_DATA) & " ,  Deseja Atualizar Agora", _
        vbYesNo + vbDefaultButton1 + vbInformation, "COTAÇÃO") = vbYes Then
            'XFormulario = "TelaEntradaMateriais"
            TelaCotacoes.LblChave.Caption = XLL_CODMOEDA
            TelaCotacoes.LblData.Caption = Format(XLD_DATA, "dd/mm/yy")
            TelaCotacoes.Show 1
            Screen.MousePointer = vbHourglass
            GoTo Inicio_Funcao
        Else
            Screen.MousePointer = vbDefault
            MsgBox "Sem a Cotação atualizada, não é possível efetuar os Cálculos.", vbCritical, "Atenção"
        End If
    Else
        Screen.MousePointer = XLL_MOUSE
    End If
    FunCotacao = XLB_Cotacao
    Ordem = XLT_ORDEM
End Function

Sub subGravaFinanceiro(XLO_TITULO As Object, XLT_STATUS As String, XLI_CODCLIENTE As Integer, XLT_TIPOPLANO As String, XLT_NATUREZAPLANO As String)
    'Esta função grava os lançamentos compromissados e/ou realizados no financeiro

    '********PARÂMETRO*********
    'XLO_TITULO 'Recordset com a tabela dos titulos
    'XLT_STATUS 'Indicar se está pagando, cancelando, ou cadastrando um título
                '(Pagar, Cancelar, Compromissar, Alterar)
    'XLI_CODCLIENTE 'Código do cliente
    'XLT_TIPOPLANO  'Tipo Plano
    'XLT_NATUREZAPLANO  'Natureza plano
    '**************************
    Dim XLT_SQL As String           'Prepara Select
    Dim XLT_NUMDOC As String        'Nº do doc. da NF (gerado pelo sistema)
    Dim XLT_STATUSGRAVAR As String  'Sttaus da Nota e da parcela para gravação no banco
    Dim XLL_FATURA As Long          'Nº da Fatura que será gravado em NF e Duplicatas
    Dim XLI_TIPODOC As Integer      'Código do tipo de documento = Nota Promissória
    Dim XLO_REGISTRO As ADODB.Recordset 'Recordset para os acessos às tabelas do Qualifin
    Dim XLT_DOCREC As String        'Nº do doc. recebimento. Foi gerado pelo sistema posi no capi vem tudo junto
    Dim XLT_TITULO As String        'Código completo do título
    Dim XLF_VALOR As Currency       'Valor A Receber(compromisso) Recebido (pago)
    Dim XLL_CODDUPL As Long         'Código da duplicat no financeiro
    Dim XLI_CODJUROS As Integer     'Código do Acrécimo Juros no Financeiro
    Dim XLI_CODMULTA As Integer     'Código do Acrécimo Multa no Financeiro
    Dim XLI_CODDESAGIO As Integer   'Código do Desconto Deságio no Financeiro
    Dim XLI_CODDESCONTO As Integer  'Código do Desconto no Financeiro
    Dim XLI_CODSEGURO As Integer    'Código do Acréscimo Seguro no Financeiro
    Dim XLI_CODOUTROS As Integer    'Código do Acréscimo Outros no Financeiro
    Dim XLD_REGCONTRATO As Date     'Data do Reg. Contrato Para ser exp. como dt. emissão p/ Qualifin qdo. pago
    Dim XLD_VENDACONTRATO As Date   'Dt. Venda do Imovel P/ ser exp. como dt. compet. e emissão p/ Qualifin qdo. não pago
    
    'Obter código do tipo de doc cuja sigla = NP (Nota Promissória)
    SubQOpenRecordset XLO_REGISTRO, "SELECT tido_cd_TipoDoc FROM TipoDocumentos WHERE tido_tx_Sigla= 'NP'", Estatico
    XLI_TIPODOC = XLO_REGISTRO!tido_cd_TipoDoc
    XLO_REGISTRO.Close
    
    'Obter Dt de registro do contrato e dt. venda do imovel
    SubQOpenRecordset XLO_REGISTRO, "SELECT cont_dt_RegistroContrato, cont_dt_Venda FROM Contratos " & _
                " WHERE empd_cd_Empreendimento= '" & XLO_TITULO!empd_cd_Empreendimento & "' AND " & _
                " imov_cd_Imovel='" & XLO_TITULO!imov_cd_Imovel & "' AND cont_cd_Contrato='" & XLO_TITULO!cont_cd_Contrato & "' AND " & _
                " empr_cd_Empresa = " & PCodEmpresa & "", Estatico
    XLD_REGCONTRATO = XLO_REGISTRO!cont_dt_RegistroContrato
    XLD_VENDACONTRATO = XLO_REGISTRO!cont_dt_Venda
    XLO_REGISTRO.Close
    
    XLT_TITULO = XLO_TITULO!empd_cd_Empreendimento + "." + XLO_TITULO!imov_cd_Imovel + "." + XLO_TITULO!cont_cd_Contrato + "." + XLO_TITULO!titu_cd_Plano + "." + XLO_TITULO!titu_cd_Parcela + "." + XLO_TITULO!titu_cd_Residuo
    If (XLT_STATUS = "COMPROMISSAR") _
        Or (XGT_EXPCOMPROMCAPI = "N" And XLT_STATUS = "PAGAR") Then
        SubQOpenRecordset XLO_REGISTRO, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal= -1", Dinamico
        XLO_REGISTRO.AddNew
        XLL_FATURA = FunPegaGuardaUltimo("F")
        XLT_NUMDOC = Format(FunPegaGuardaUltimo("D"), "A000000000")
        XLO_REGISTRO!nofi_nr_documento = XLT_NUMDOC
        XLO_REGISTRO!nofi_nr_fatura = XLL_FATURA
    ElseIf (XGT_EXPCOMPROMCAPI = "S" And XLT_STATUS = "PAGAR") _
        Or (XLT_STATUS = "ALTERAR") Then
        SubQOpenRecordset XLO_REGISTRO, "SELECT * FROM NotasFiscais WHERE nofi_nr_TituloCapi= '" & XLT_TITULO & "'", Dinamico
        If Not XLO_REGISTRO.EOF Then
            XLT_NUMDOC = XLO_REGISTRO!nofi_nr_documento
            XLL_FATURA = XLO_REGISTRO!nofi_nr_fatura
        End If
    End If
    
    If XLO_REGISTRO.EOF Then 'Se não tiver sido compromissado no Qualifin
        XLO_REGISTRO.AddNew
        XLL_FATURA = FunPegaGuardaUltimo("F")
        XLT_NUMDOC = Format(FunPegaGuardaUltimo("D"), "A000000000")
        XLO_REGISTRO!nofi_nr_documento = XLT_NUMDOC
        XLO_REGISTRO!nofi_nr_fatura = XLL_FATURA
    End If
    
    If XLT_STATUS = "COMPROMISSAR" Or XLT_STATUS = "ALTERAR" Then
        XLT_STATUSGRAVAR = "C"
        XLF_VALOR = XLO_TITULO!titu_vl_Parcela
        XLO_REGISTRO!nofi_dt_entrada = Format$(XLD_VENDACONTRATO, "DD/MM/YY")
        XLO_REGISTRO!nofi_dt_emissao = Format$(XLD_VENDACONTRATO, "DD/MM/YY")
    ElseIf XLT_STATUS = "PAGAR" Then
        XLT_STATUSGRAVAR = "R"
        'XLF_VALOR = XLO_TITULO!titu_vl_ValorPago
        XLF_VALOR = Format(XLO_TITULO!titu_vl_ValorPago - XLO_TITULO!titu_vl_Seguro - XLO_TITULO!titu_vl_Juros - XLO_TITULO!titu_vl_Multa - XLO_TITULO!titu_vl_Outros + XLO_TITULO!titu_vl_Desagio + XLO_TITULO!titu_vl_Desconto, "standard")
        XLO_REGISTRO!nofi_dt_entrada = Format$(XLO_TITULO!titu_dt_Pagamento, "DD/MM/YY")
        XLO_REGISTRO!nofi_dt_emissao = Format$(XLD_REGCONTRATO, "DD/MM/YY")
    End If
    
    XLO_REGISTRO!nofi_tx_tipomov = "E" 'Entrada
    XLO_REGISTRO!nofi_tx_status = XLT_STATUSGRAVAR
    XLO_REGISTRO!focl_cd_FornCli = XLI_CODCLIENTE
    XLO_REGISTRO!nofi_tx_serie = "U" 'Única
    XLO_REGISTRO!tido_cd_TipoDoc = XLI_TIPODOC
    XLO_REGISTRO!nofi_tx_historico = "PARC. " + XLT_TITULO + " - " + XLT_TIPOPLANO + " - " + XLT_NATUREZAPLANO
    XLO_REGISTRO!nofi_vl_valor = XLF_VALOR
    XLO_REGISTRO!nofi_tx_distribuir = "N" 'Sempre 1 conta para 1 CCusto
    XLO_REGISTRO!nofi_tx_tipofatura = "1" 'Uma nota para uma duplicata
    XLO_REGISTRO!nofi_nr_TituloCapi = XLT_TITULO 'Código do título
    XLO_REGISTRO!empr_cd_Empresa = PCodEmpresa

    XLO_REGISTRO.Update

    Chave = XLO_REGISTRO!nofi_cd_notaFiscal
    XLL_FATURA = XLO_REGISTRO!nofi_nr_fatura
    XLO_REGISTRO.Close
        
    '*****************************DUPLICATAS*******************************
    'dupl_tx_debcred = (C - Crédito)
    'dupl_nr_duplicata=(sempre parcela U - Única)
    'DocPagrec 'Não está separado. Gera automat.
    If XLT_STATUS = "PAGAR" Then
        XLT_DOCREC = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
    End If
    If (XLT_STATUS = "COMPROMISSAR") _
        Or (XGT_EXPCOMPROMCAPI = "N" And XLT_STATUS = "PAGAR") Then
        SubQOpenRecordset XLO_REGISTRO, "SELECT * FROM Duplicatas WHERE dupl_cd_Duplicata = -1", Dinamico
        XLO_REGISTRO.AddNew
    ElseIf (XGT_EXPCOMPROMCAPI = "S" And XLT_STATUS = "PAGAR") _
        Or (XLT_STATUS = "ALTERAR") Then
        SubQOpenRecordset XLO_REGISTRO, "SELECT * FROM Duplicatas WHERE dupl_nr_Fatura = " & XLL_FATURA, Dinamico
    End If
        
    If XLO_REGISTRO.EOF Then 'Se não foi cadastrado no Qualifin
        XLO_REGISTRO.AddNew
    End If
    XLO_REGISTRO!dupl_nr_duplicata = "U"
    XLO_REGISTRO!dupl_dt_vencimento = XLO_TITULO!titu_dt_Prorrogacao
    XLO_REGISTRO!dupl_vl_valor = XLF_VALOR
    XLO_REGISTRO!dupl_tx_status = XLT_STATUSGRAVAR
    If XLO_TITULO!titu_tx_LocalPagto = "3" Then 'Dinheiro
        XLO_REGISTRO!dupl_tx_tipopagrec = "6" 'Dinheiro
    Else
        XLO_REGISTRO!dupl_tx_tipopagrec = "5" 'Crédito em Conta
    End If
    XLO_REGISTRO!dupl_tx_DocPagRec = XLT_DOCREC
    XLO_REGISTRO!coco_cd_Codigo = XLO_TITULO!coco_cd_Codigo
    XLO_REGISTRO!dupl_tx_debcred = "C"
    XLO_REGISTRO!dupl_nr_fatura = XLL_FATURA
    XLO_REGISTRO!dupl_dt_pagrec = XLO_TITULO!titu_dt_Deposito
    
    XLO_REGISTRO.Update
    
    XLL_CODDUPL = XLO_REGISTRO!dupl_cd_Duplicata
    
    XLO_REGISTRO.Close
                
    '***************************DESCONTO/ACRÉSCIMOS***************************
    SubQOpenRecordset XLO_REGISTRO, "SELECT * FROM DescontosAcrescimos", Estatico
    If XLO_TITULO!titu_vl_Multa <> 0 Then
        XLO_REGISTRO.MoveFirst
        XLO_REGISTRO.Find "deac_tx_Descricao = 'Multa por Atraso - Cliente'"
        XLI_CODMULTA = XLO_REGISTRO!deac_cd_descacresc
        
        Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata,deac_cd_descacresc,dude_vl_descacresc)" & _
                        " VALUES (" & XLL_CODDUPL & "," & XLI_CODMULTA & "," & FunTrataFloat(XLO_TITULO!titu_vl_Multa) & ")"
    End If
    If XLO_TITULO!titu_vl_Juros <> 0 Then
        XLO_REGISTRO.MoveFirst
        XLO_REGISTRO.Find "deac_tx_Descricao ='Juros por Atraso - Cliente'"
        XLI_CODJUROS = XLO_REGISTRO!deac_cd_descacresc
        
        Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata,deac_cd_descacresc,dude_vl_descacresc)" & _
                        " VALUES (" & XLL_CODDUPL & "," & XLI_CODJUROS & "," & FunTrataFloat(XLO_TITULO!titu_vl_Juros) & ")"
    End If
    If XLO_TITULO!titu_vl_Desagio <> 0 Then
        XLO_REGISTRO.MoveFirst
        XLO_REGISTRO.Find "deac_tx_Descricao ='Deságio - Cliente'"
        XLI_CODDESAGIO = XLO_REGISTRO!deac_cd_descacresc
        
        Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata,deac_cd_descacresc,dude_vl_descacresc)" & _
                        " VALUES (" & XLL_CODDUPL & "," & XLI_CODDESAGIO & "," & FunTrataFloat(XLO_TITULO!titu_vl_Desagio) & ")"
    End If
    If XLO_TITULO!titu_vl_Desconto <> 0 Then
        XLO_REGISTRO.MoveFirst
        XLO_REGISTRO.Find "deac_tx_Descricao ='Descontos - Cliente'"
        XLI_CODDESCONTO = XLO_REGISTRO!deac_cd_descacresc
        
        Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata,deac_cd_descacresc,dude_vl_descacresc)" & _
                        " VALUES (" & XLL_CODDUPL & "," & XLI_CODDESCONTO & "," & FunTrataFloat(XLO_TITULO!titu_vl_Desconto) & ")"
    End If
    If XLO_TITULO!titu_vl_Seguro <> 0 Then
        XLO_REGISTRO.MoveFirst
        XLO_REGISTRO.Find "deac_tx_Descricao ='Seguro - Cliente'"
        XLI_CODSEGURO = XLO_REGISTRO!deac_cd_descacresc
        
        Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata,deac_cd_descacresc,dude_vl_descacresc)" & _
                        " VALUES (" & XLL_CODDUPL & "," & XLI_CODSEGURO & "," & FunTrataFloat(XLO_TITULO!titu_vl_Seguro) & ")"
    End If
    If XLO_TITULO!titu_vl_Outros <> 0 Then
        XLO_REGISTRO.MoveFirst
        XLO_REGISTRO.Find "deac_tx_Descricao ='Outros - Cliente'"
        XLI_CODOUTROS = XLO_REGISTRO!deac_cd_descacresc
        
        Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata,deac_cd_descacresc,dude_vl_descacresc)" & _
                        " VALUES (" & XLL_CODDUPL & "," & XLI_CODOUTROS & "," & FunTrataFloat(XLO_TITULO!titu_vl_Outros) & ")"
    End If
'    If XLO_TITULO.titu_vl_multa <> 0 Then
'        XLT_SQL = "SELECT *, (SELECT deac_cd_descacresc FROM DescontosAcrescimos" & _
'                    " WHERE deac_tx_Descricao LIKE 'Multa por Atraso - Cliente') as CodDescAcresc" & _
'                    " FROM DuplDescAcresc "
'        SubQOpenRecordset XLO_REGISTRO, XLT_SQL, Dinamico
'        XLI_CODMULTA = XLO_REGISTRO!deac_cd_DescAcresc
'        XLO_REGISTRO.AddNew
'
'        XLO_REGISTRO!dupl_cd_Duplicata = XLL_CODDUPL
'        XLO_REGISTRO!deac_cd_DescAcrec = XLI_CODMULTA
'        XLO_REGISTRO!dude_vl_DescAcrec = XLO_TITULO!titu_vl_multa
'
'        XLO_REGISTRO.Update
'    End If
'    XLO_REGISTRO.Close

    '*****************************DISTRIBUIÇÃO*******************************
    'Com identificador
    'XLT_SQL = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta,dist_vl_valor," & _
                    "nofi_cd_notafiscal, iden_cd_Projeto)" & _
                " VALUES ('" & XLO_TITULO!cere_cd_pcr & "'," & _
                    "'" & XLO_TITULO!plco_cd_conta & "'," & _
                    "" & FunTrataFloat(xlo_tiulo!titu_vl_ValorPago) & "," & Chave & "," & _
                    "" & FunNuloBancoVal(CboIdentificador2.BoundText) & ")"
    If (XLT_STATUS = "COMPROMISSAR") _
        Or (XGT_EXPCOMPROMCAPI = "N" And XLT_STATUS = "PAGAR") Then
        XLT_SQL = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta, " & _
                        "dist_vl_valor,iden_cd_Projeto,nofi_cd_notafiscal)" & _
                  " VALUES ('" & XLO_TITULO!cere_cd_Pcr & "'," & _
                        "'" & XLO_TITULO!plco_cd_Conta & "'," & _
                        "" & FunTrataFloat(XLF_VALOR) & "," & _
                        "" & FunNuloBancoVal(XLO_TITULO!iden_cd_Projeto) & "," & Chave & ")"
        Conexao.Execute XLT_SQL
    ElseIf (XGT_EXPCOMPROMCAPI = "S" And XLT_STATUS = "PAGAR") _
        Or (XLT_STATUS = "ALTERAR") Then
        XLT_SQL = "UPDATE Distribuicao SET cere_cd_pcr='" & XLO_TITULO!cere_cd_Pcr & "'," & _
                        "plco_cd_conta='" & XLO_TITULO!plco_cd_Conta & "'," & _
                        "dist_vl_valor=" & FunTrataFloat(XLF_VALOR) & "," & _
                        "iden_cd_Projeto= " & FunNuloBancoVal(XLO_TITULO!iden_cd_Projeto) & "" & _
                    " WHERE nofi_cd_NotaFiscal = " & Chave & ""
        Conexao.Execute XLT_SQL
    End If
    'XLO_TITULO.Close
    If XLT_STATUS = "PAGAR" Then
        'TESTAR SubContabilizarPromissoria XLL_CODDUPL, XLL_FATURA, XLO_TITULO!coco_cd_Codigo, XLT_DOCREC, XLO_TITULO!titu_dt_Deposito, True
    End If
End Sub

'Retorna o valor indexado por uma determinada moeda
'Function FunCalculaValorIndexadoTEMP(XLF_VALORREAL As Double, XLD_DTBASE As Date, XLI_CODMOEDA As Integer, XLT_INDEXACAO As String, XLB_PRORATA As Boolean)
'
'  Dim XLO_COTACAO As New ADODB.Recordset
'  Dim XLO_COTACAOANTERIOR As New ADODB.Recordset
'  Dim XLF_COTACAO As Double
'  Dim XLF_COTACAOANTERIOR As Double
'  Dim XLF_PERCENT As Double
'
'  If XLT_INDEXACAO = "D" Then
'    SubQOpenRecordset XLO_COTACAO, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData(XLD_DTBASE, NomeSgbd) & "", Estatico
'  Else
'    SubQOpenRecordset XLO_COTACAO, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData("01/" & Month(XLD_DTBASE) & "/" & Year(XLD_DTBASE), NomeSgbd) & "", Estatico
'  End If
'  If Not XLO_COTACAO.EOF Then
'    XLF_COTACAO = XLO_COTACAO!cota_vl_Valor
'    'Se for diário ou mensal e não prorata
'    If XLT_INDEXACAO = "D" Or (XLT_INDEXACAO = "M" And Not XLB_PRORATA) Then
'      FunCalculaValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
'    'se for mensal prorata
'    Else
'      XLD_DTBASE = DateAdd("m", -1, XLD_DTBASE)
'      SubQOpenRecordset XLO_COTACAOANTERIOR, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData("01/" & Month(XLD_DTBASE) & "/" & Year(XLD_DTBASE), NomeSgbd) & "", Estatico
'      If Not XLO_COTACAOANTERIOR.EOF Then
'        XLF_COTACAOANTERIOR = XLO_COTACAOANTERIOR!cota_vl_Valor
'        XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
'        XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASE) - 1)) * XLF_COTACAO)
'        FunCalculaValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
'      Else
'        TelaCotacoes.LblChave.Caption = XLI_CODMOEDA
'        If XLT_INDEXACAO = "D" Then
'          TelaCotacoes.LblData.Caption = Format(XLD_DTBASE, "DD/MM/YY")
'        Else
'          TelaCotacoes.LblData.Caption = "01/" & Format(XLD_DTBASE, "MM/YY")
'        End If
'        TelaCotacoes.Show 1
'        XLO_COTACAOANTERIOR.Requery
'        If Not XLO_COTACAOANTERIOR.EOF Then
'          XLF_COTACAOANTERIOR = XLO_COTACAOANTERIOR!cota_vl_Valor
'          XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
'          XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASE) - 1)) * XLF_COTACAO)
'          FunCalculaValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
'        End If
'      End If
'    End If
'  Else
'      TelaCotacoes.LblChave.Caption = XLI_CODMOEDA
'      If XLT_INDEXACAO = "D" Then
'        TelaCotacoes.LblData.Caption = Format(XLD_DTBASE, "DD/MM/YY")
'      Else
'        TelaCotacoes.LblData.Caption = "01/" & Format(XLD_DTBASE, "MM/YY")
'      End If
'      TelaCotacoes.Show 1
'      XLO_COTACAO.Requery
'      If Not XLO_COTACAO.EOF Then
'        XLF_COTACAO = XLO_COTACAO!cota_vl_Valor
'        'Se for diário ou mensal e não prorata
'        If XLT_INDEXACAO = "D" Or (XLT_INDEXACAO = "M" And Not XLB_PRORATA) Then
'          FunCalculaValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
'        'se for mensal prorata
'        Else
'          XLD_DTBASE = DateAdd("m", -1, XLD_DTBASE)
'          SubQOpenRecordset XLO_COTACAOANTERIOR, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData("01/" & Month(XLD_DTBASE) & "/" & Year(XLD_DTBASE), NomeSgbd) & "", Estatico
'          If Not XLO_COTACAOANTERIOR.EOF Then
'            XLF_COTACAOANTERIOR = XLO_COTACAOANTERIOR!cota_vl_Valor
'            XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
'            XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASE) - 1)) * XLF_COTACAO)
'            FunCalculaValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
'          Else
'            TelaCotacoes.LblChave.Caption = XLI_CODMOEDA
'            If XLT_INDEXACAO = "D" Then
'              TelaCotacoes.LblData.Caption = XLD_DTBASE
'            Else
'              TelaCotacoes.LblData.Caption = "01/" & Format(XLD_DTBASE, "MM/YY")
'            End If
'            TelaCotacoes.Show 1
'            XLO_COTACAOANTERIOR.Requery
'            If Not XLO_COTACAOANTERIOR.EOF Then
'              XLF_COTACAOANTERIOR = XLO_COTACAOANTERIOR!cota_vl_Valor
'              XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
'              XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASE) - 1)) * XLF_COTACAO)
'              FunCalculaValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
'            End If
'          End If
'        End If
'      End If
'  End If
'  XLO_COTACAO.Close
'  Set XLO_COTACAO = Nothing
'
'End Function

'Retorna o valor indexado por uma determinada moeda
Function FunCalculaValorIndexado(XLF_VALORREAL As Double, ByVal XLD_DTBASE As Date, XLI_CODMOEDA As Integer, XLT_INDEXACAO As String, XLT_PRORATA As String)
        
  Dim XLO_COTACAO As New ADODB.Recordset
  Dim XLO_COTACAOANTERIOR As New ADODB.Recordset
  Dim XLF_COTACAO As Double
  Dim XLF_COTACAOANTERIOR As Double
  Dim XLF_PERCENT As Double
 
  If XLT_INDEXACAO = "D" Then
    SubQOpenRecordset XLO_COTACAO, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData(XLD_DTBASE, NomeSgbd) & "", Estatico
  Else
    SubQOpenRecordset XLO_COTACAO, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData("01/" & Month(XLD_DTBASE) & "/" & Year(XLD_DTBASE), NomeSgbd) & "", Estatico
  End If
  If Not XLO_COTACAO.EOF Then
    XLF_COTACAO = XLO_COTACAO!cota_vl_Valor
    'Se for diário ou mensal e não prorata
    If XLT_INDEXACAO = "D" Or (XLT_INDEXACAO = "M" And XLT_PRORATA = "0") Then
      FunCalculaValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
    'se for mensal prorata
    Else
      
      If XLT_PRORATA = "A" Then
        XLD_DTBASE = DateAdd("m", -1, XLD_DTBASE)
      Else
        XLD_DTBASE = DateAdd("m", 1, XLD_DTBASE)
      End If
        
      SubQOpenRecordset XLO_COTACAOANTERIOR, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData("01/" & Month(XLD_DTBASE) & "/" & Year(XLD_DTBASE), NomeSgbd) & "", Estatico
      If Not XLO_COTACAOANTERIOR.EOF Then
        XLF_COTACAOANTERIOR = XLO_COTACAOANTERIOR!cota_vl_Valor
        
        If XLT_PRORATA = "A" Then
          XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
        Else
          XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAO
        End If
        
        XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASE) - 1)) * XLF_COTACAO)
        FunCalculaValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
      Else
        TelaCotacoes.LblChave.Caption = XLI_CODMOEDA
        If XLT_INDEXACAO = "D" Then
          TelaCotacoes.LblData.Caption = Format(XLD_DTBASE, "DD/MM/YY")
        Else
          TelaCotacoes.LblData.Caption = "01/" & Format(XLD_DTBASE, "MM/YY")
        End If
        TelaCotacoes.Show 1
        XLO_COTACAOANTERIOR.Requery
        If Not XLO_COTACAOANTERIOR.EOF Then
          XLF_COTACAOANTERIOR = XLO_COTACAOANTERIOR!cota_vl_Valor
          
          If XLT_PRORATA = "A" Then
            XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
          Else
            XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAO
          End If
          
          XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASE) - 1)) * XLF_COTACAO)
          FunCalculaValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
        End If
      End If
    End If
  Else
    TelaCotacoes.LblChave.Caption = XLI_CODMOEDA
    If XLT_INDEXACAO = "D" Then
      TelaCotacoes.LblData.Caption = Format(XLD_DTBASE, "DD/MM/YY")
    Else
      TelaCotacoes.LblData.Caption = "01/" & Format(XLD_DTBASE, "MM/YY")
    End If
    TelaCotacoes.Show 1
    XLO_COTACAO.Requery
    If Not XLO_COTACAO.EOF Then
      XLF_COTACAO = XLO_COTACAO!cota_vl_Valor
      'Se for diário ou mensal e não prorata
      If XLT_INDEXACAO = "D" Or (XLT_INDEXACAO = "M" And XLT_PRORATA = "0") Then
        FunCalculaValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
      'se for mensal prorata
      Else
        If XLT_PRORATA = "A" Then
          XLD_DTBASE = DateAdd("m", -1, XLD_DTBASE)
        Else
          XLD_DTBASE = DateAdd("m", 1, XLD_DTBASE)
        End If

        SubQOpenRecordset XLO_COTACAOANTERIOR, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData("01/" & Month(XLD_DTBASE) & "/" & Year(XLD_DTBASE), NomeSgbd) & "", Estatico
        If Not XLO_COTACAOANTERIOR.EOF Then
          XLF_COTACAOANTERIOR = XLO_COTACAOANTERIOR!cota_vl_Valor
          
          If XLT_PRORATA = "A" Then
            XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
          Else
            XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAO
          End If

          XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASE) - 1)) * XLF_COTACAO)
          FunCalculaValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
        Else
          TelaCotacoes.LblChave.Caption = XLI_CODMOEDA
          If XLT_INDEXACAO = "D" Then
            TelaCotacoes.LblData.Caption = XLD_DTBASE
          Else
            TelaCotacoes.LblData.Caption = "01/" & Format(XLD_DTBASE, "MM/YY")
          End If
          TelaCotacoes.Show 1
          XLO_COTACAOANTERIOR.Requery
          If Not XLO_COTACAOANTERIOR.EOF Then
            XLF_COTACAOANTERIOR = XLO_COTACAOANTERIOR!cota_vl_Valor
            
            If XLT_PRORATA = "A" Then
              XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
            Else
              XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAO
            End If

            XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASE) - 1)) * XLF_COTACAO)
            FunCalculaValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
          End If
        End If
      End If
    End If
  End If
  XLO_COTACAO.Close
  Set XLO_COTACAO = Nothing

End Function




'Retorna o valor calculado do jurosMulta
Function funCalculaJurosMulta(XLF_JUROS As Double, XLF_MULTA As Double, XLT_CODTITULO As String, XLD_DTBASEPAGTO As Date, XLF_VALOR As Double) As Double
    Dim XLT_CODEMPREENDIMENTO As String 'Armazena código do empreendimento
    Dim XLT_CODIMOVEL As String         'Armazena código do Imóvel
    Dim XLT_CODCONTRATO As String       'Armazena código do contrato
    Dim XLT_CODPLANO As String          'Armazena código do plano
    Dim XLT_CODPARCELA As String        'Armazena código da Parcela
    Dim XLT_CODRESIDUO As String        'Armazena código do Residuo
    Dim XLO_TITULO As Object            'Recordset para leitura dos dados do título
    Dim XLT_SQL As String               'Prepara o SQL para o Recordset
    Dim XLF_PERCJUROS As Double         'Armazena percentual do juros gravado no título
    Dim XLI_NUMDIAS As Integer          'Armazena o nº de dias em atraso
    Dim XLI_NUMMESES As Integer         'Armazena o nº de meses em atraso
    Dim XLF_PERCMULTA As Double         'Armazena percentual da multa gravado no título
    
    XLT_CODEMPREENDIMENTO = Left$(XLT_CODTITULO, 4)
    XLT_CODIMOVEL = Mid$(XLT_CODTITULO, 6, 4)
    XLT_CODCONTRATO = Mid$(XLT_CODTITULO, 11, 2)
    XLT_CODPLANO = Mid$(XLT_CODTITULO, 14, 2)
    XLT_CODPARCELA = Mid$(XLT_CODTITULO, 17, 3)
    XLT_CODRESIDUO = Right$(XLT_CODTITULO, 2)
   
    XLT_SQL = "SELECT * FROM ConsCapBaseCalculo WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
            " AND empd_cd_Empreendimento= '" & XLT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XLT_CODIMOVEL & "'" & _
            " AND cont_cd_Contrato='" & XLT_CODCONTRATO & "' AND titu_cd_Plano='" & XLT_CODPLANO & "'" & _
            " AND titu_cd_Parcela='" & XLT_CODPARCELA & "' AND titu_cd_Residuo='" & XLT_CODRESIDUO & "'"
    SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
    
    XLF_PERCJUROS = (XLO_TITULO!titu_vl_PercJurosMora / 100)
    XLF_PERCMULTA = (XLO_TITULO!titu_vl_PercMultaMora / 100)
    'XLI_NUMMESES = (XLD_DTBASEPAGTO - XLO_TITULO!titu_dt_Prorrogacao) \ 30
    XLI_NUMMESES = (XLD_DTBASEPAGTO - XLO_TITULO!titu_dt_vencimento) \ 30
    'XLI_NUMDIAS = XLD_DTBASEPAGTO - XLO_TITULO!titu_dt_Prorrogacao
    XLI_NUMDIAS = XLD_DTBASEPAGTO - XLO_TITULO!titu_dt_vencimento
    
    'JUROS
    Select Case XLO_TITULO!cont_tx_TipoJurosMora
        Case "D" 'Juros Diário
            XLF_JUROS = Format((((1 + XLF_PERCJUROS) ^ XLI_NUMDIAS) - 1) * XLF_VALOR, "standard")
        Case "M" 'Juros Mensal
            'S=P(1+i) elavado a n, onde S=montante, P=Vl. Principal, i=taxa, n=período
            XLF_JUROS = Format(((((1 + XLF_PERCJUROS) ^ XLI_NUMMESES) - 1) * XLF_VALOR), "standard")
        Case "C" 'Juros Cheio
            'If ((XLD_DTBASEPAGTO - XLO_TITULO!titu_dt_Prorrogacao) Mod 30) > 0 Then
            If ((XLD_DTBASEPAGTO - XLO_TITULO!titu_dt_vencimento) Mod 30) > 0 Then
                XLI_NUMMESES = XLI_NUMMESES + 1
            End If
            'S=P(1+i) elavado a n, onde S=montante, P=Vl. Principal, i=taxa, n=período
            'funCalculaJuros = ((((1 + XLF_PERCJUROS) ^ XLI_NUMMESES) - 1) * XLF_VALOR)
            'Qdo. Juros Cheio o Valor deve ser sempre sobre o valor histórico
            XLF_JUROS = Format(((((1 + XLF_PERCJUROS) ^ XLI_NUMMESES) - 1) * XLO_TITULO!titu_vl_Parcela), "standard")
            XLF_MULTA = Format((XLF_PERCMULTA * XLO_TITULO!titu_vl_Parcela), "standard")
            XLO_TITULO.Close
            Exit Function
    End Select
                
    'MULTA (Se o tipo de juros for = cheio, a multa já foi calculada acima com base no valor histórico)
    If XLO_TITULO!cont_tx_MultaProRata = "1" And XLI_NUMDIAS <= 30 Then 'se for Pro_rata e o juros não é cheio
        XLF_MULTA = Format((((((1 + XLF_PERCMULTA) ^ (1 / 30)) ^ XLI_NUMDIAS) - 1) * XLF_VALOR), "standard")
    Else 'não é pro-rata or o tipo de juros é cheio
        XLF_MULTA = Format((XLF_PERCMULTA * XLF_VALOR), "standard")
    End If
    
    XLO_TITULO.Close
End Function


'Retorna o valor calculado do deságio
Function funCalculaDesagio(XLT_CODTITULO As String, XLD_DTBASEPAGTO As Date, XLF_VALOR As Double) As Double
    Dim XLT_CODEMPREENDIMENTO As String 'Armazena código do empreendimento
    Dim XLT_CODIMOVEL As String         'Armazena código do Imóvel
    Dim XLT_CODCONTRATO As String       'Armazena código do contrato
    Dim XLT_CODPLANO As String          'Armazena código do plano
    Dim XLT_CODPARCELA As String        'Armazena código da Parcela
    Dim XLT_CODRESIDUO As String        'Armazena código do Residuo
    Dim XLO_TITULO As Object            'Recordset para leitura dos dados do título
    Dim XLT_SQL As String               'Prepara o SQL para o Recordset
    Dim XLF_PERCJUROS As Double         'Armazena percentual do juros gravado no título
    Dim XLI_NUMDIAS As Integer          'Armazena o nº de dias em atraso
    Dim XLI_NUMMESES As Integer         'Armazena o nº de meses em atraso
       
    XLT_CODEMPREENDIMENTO = Left$(XLT_CODTITULO, 4)
    XLT_CODIMOVEL = Mid$(XLT_CODTITULO, 6, 4)
    XLT_CODCONTRATO = Mid$(XLT_CODTITULO, 11, 2)
    XLT_CODPLANO = Mid$(XLT_CODTITULO, 14, 2)
    XLT_CODPARCELA = Mid$(XLT_CODTITULO, 17, 3)
    XLT_CODRESIDUO = Right$(XLT_CODTITULO, 2)
   
    XLT_SQL = "SELECT * FROM ConsCapBaseCalculo WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
            " AND empd_cd_Empreendimento= '" & XLT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XLT_CODIMOVEL & "'" & _
            " AND cont_cd_Contrato='" & XLT_CODCONTRATO & "' AND titu_cd_Plano='" & XLT_CODPLANO & "'" & _
            " AND titu_cd_Parcela='" & XLT_CODPARCELA & "' AND titu_cd_Residuo='" & XLT_CODRESIDUO & "'"
    SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
    
    'XLI_NUMDIAS = XLO_TITULO!titu_dt_Prorrogacao - XLD_DTBASEPAGTO
    XLI_NUMDIAS = XLO_TITULO!titu_dt_vencimento - XLD_DTBASEPAGTO
    'O percentual do deságio sempre é gravado no título como ao mês, o tipo é que pode ser diário ou mensal
    If XLO_TITULO!cont_tx_Desagio = "M" Then 'Mensal
        XLF_PERCJUROS = XLO_TITULO!titu_vl_PercJurosFin
        If XLI_NUMDIAS >= 30 Then
           'XLI_NUMMESES = (XLO_TITULO!titu_dt_Prorrogacao - XLD_DTBASEPAGTO) \ 30
           XLI_NUMMESES = (XLO_TITULO!titu_dt_vencimento - XLD_DTBASEPAGTO) \ 30
           funCalculaDesagio = (XLF_VALOR - (XLF_VALOR / ((1 + (XLF_PERCJUROS) / 100) ^ XLI_NUMMESES)))
        End If
    Else 'Diário
        XLF_PERCJUROS = (1 + (XLO_TITULO!titu_vl_PercJurosFin / 100)) ^ (1 / 30)
        XLF_PERCJUROS = (XLF_PERCJUROS - 1) * 100
        funCalculaDesagio = XLF_VALOR - (XLF_VALOR / ((1 + XLF_PERCJUROS / 100) ^ XLI_NUMDIAS))
    End If
    XLO_TITULO.Close
End Function

'Retorna o valor calculado da Correção Monetária
Function funCalculaCorrecaoMonetaria(XLT_CODTITULO As String, XLD_DTBASEPAGTO As Date, XLT_TIPO As String) As Double
    Dim XLT_CODEMPREENDIMENTO As String 'Armazena código do empreendimento
    Dim XLT_CODIMOVEL As String         'Armazena código do Imóvel
    Dim XLT_CODCONTRATO As String       'Armazena código do contrato
    Dim XLT_CODPLANO As String          'Armazena código do plano
    Dim XLT_CODPARCELA As String        'Armazena código da Parcela
    Dim XLT_CODRESIDUO As String        'Armazena código do Residuo
    Dim XLO_TITULO As Object            'Recordset para leitura dos dados do título
    Dim XLO_MOEDA As Object             'Recordset para leitura dos dados da cotação
    Dim XLT_SQL As String               'Prepara o SQL para o Recordset
    Dim XLF_PERCJUROS As Double         'Armazena percentual do juros gravado no título
    Dim XLI_NUMDIAS As Integer          'Armazena o nº de dias em atraso
    Dim XLI_NUMMESES As Integer         'Armazena o nº de meses em atraso
    Dim XLI_MESANTERIOR As Integer      'Armazena o mês anterior
    Dim XLI_ANO As Integer              'Armazena o ano do caso do mês anterior = 12
    Dim XLF_COTACAO As Double           'Armazena a cotacao na data base do pagto
    Dim XLF_COTACAOANTERIOR As Double   'Armazena a cotacao na data 1 mês antes do mês da dt base do pagto
    Dim XLF_COTACAOTITULO As Double     'Armazena a cotacao da geracao do título
    Dim XLF_PERCENT As Double           'Divisão entre a cotacao dt. base e o mês anterior
    Dim XLT_INDEXACAO As String         'M - Mensal D - Diário
    Dim XLL_CODMOEDA As Long            'Código da moeda do título
       
    XLT_CODEMPREENDIMENTO = Left$(XLT_CODTITULO, 4)
    XLT_CODIMOVEL = Mid$(XLT_CODTITULO, 6, 4)
    XLT_CODCONTRATO = Mid$(XLT_CODTITULO, 11, 2)
    XLT_CODPLANO = Mid$(XLT_CODTITULO, 14, 2)
    XLT_CODPARCELA = Mid$(XLT_CODTITULO, 17, 3)
    XLT_CODRESIDUO = Right$(XLT_CODTITULO, 2)
    
    XLT_SQL = "SELECT * FROM ConsCapBaseCalculo WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
            " AND empd_cd_Empreendimento= '" & XLT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XLT_CODIMOVEL & "'" & _
            " AND cont_cd_Contrato='" & XLT_CODCONTRATO & "' AND titu_cd_Plano='" & XLT_CODPLANO & "'" & _
            " AND titu_cd_Parcela='" & XLT_CODPARCELA & "' AND titu_cd_Residuo='" & XLT_CODRESIDUO & "'"
    SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
    
    'Verificar o tipo de indexação
    If XLT_TIPO = "1" Then 'Indexador 1
        XLF_COTACAOTITULO = XLO_TITULO!titu_vl_ParcelaIndex1
        XLL_CODMOEDA = XLO_TITULO!moed_cd_Moeda1
        XLT_INDEXACAO = XLO_TITULO!IndexMoeda1
    Else 'Indexador 2
        XLF_COTACAOTITULO = XLO_TITULO!titu_vl_ParcelaIndex2
        XLL_CODMOEDA = XLO_TITULO!moed_cd_Moeda2
        XLT_INDEXACAO = XLO_TITULO!IndexMoeda2
    End If
    
    'Pegar a Cotação
    If XLT_INDEXACAO = "D" Then 'Diária
        SubQOpenRecordset XLO_MOEDA, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLL_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData(XLD_DTBASEPAGTO, NomeSgbd) & "", Estatico
    Else 'Mensal
        SubQOpenRecordset XLO_MOEDA, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLL_CODMOEDA & " and month(cota_dt_datacotacao) =" & Month(XLD_DTBASEPAGTO) & " AND year(cota_dt_datacotacao) =" & Year(XLD_DTBASEPAGTO) & "", Estatico
    End If
    If Not XLO_MOEDA.EOF Then
        XLF_COTACAO = XLO_MOEDA!cota_vl_Valor
    Else
        'Verificar se existe cotação para a data base
        If FunCotacao(XLD_DTBASEPAGTO, XLL_CODMOEDA, XLT_INDEXACAO, XLF_COTACAO) = False Then
            XLO_TITULO.Close
            funCalculaCorrecaoMonetaria = 0
            Exit Function
        End If
    End If
    XLO_MOEDA.Close
        
    If XLT_INDEXACAO = "D" Or (XLT_INDEXACAO = "M" And XLO_TITULO!cont_tx_CorrecProRata = 0) Then
        'Indexação diária or Indexacao Mensal sem Correção monetária pro-rata
        funCalculaCorrecaoMonetaria = (XLF_COTACAO * XLF_COTACAOTITULO)
    ElseIf Day(XLD_DTBASEPAGTO) = 1 Then  'Indexação Mensal, Pro-Rata Dia dtBase = 1
        funCalculaCorrecaoMonetaria = (XLF_COTACAO * XLF_COTACAOTITULO)
    Else
        'Pegar a Cotação do mês anterior
'        If Month(XLD_DTBASEPAGTO) = 1 Then
'            XLI_MESANTERIOR = 12
'            XLI_ANO = Year(XLD_DTBASEPAGTO) - 1
'        Else
'            XLI_MESANTERIOR = Month(XLD_DTBASEPAGTO) - 1
'            XLI_ANO = Year(XLD_DTBASEPAGTO)
'        End If
           
        If XLO_TITULO!cont_tx_CorrecProRata = "A" Then
          XLD_DTBASEPAGTO = DateAdd("m", -1, XLD_DTBASEPAGTO)
        Else
          XLD_DTBASEPAGTO = DateAdd("m", 1, XLD_DTBASEPAGTO)
        End If
        XLI_MESANTERIOR = Month(XLD_DTBASEPAGTO)
        XLI_ANO = Year(XLD_DTBASEPAGTO)
        
        If XLT_TIPO = "1" Then 'Indexador 1
            XLT_SQL = "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLO_TITULO!moed_cd_Moeda1 & "" & _
                      " and month(cota_dt_datacotacao) =" & XLI_MESANTERIOR & " and year(cota_dt_datacotacao) =" & XLI_ANO & ""
        Else 'Indexador 2
            XLT_SQL = "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLO_TITULO!moed_cd_Moeda2 & "" & _
                      " and month(cota_dt_datacotacao) =" & XLI_MESANTERIOR & " and year(cota_dt_datacotacao) =" & XLI_ANO & ""
        End If
        SubQOpenRecordset XLO_MOEDA, XLT_SQL, Estatico
        
        If Not XLO_MOEDA.EOF Then
            XLF_COTACAOANTERIOR = XLO_MOEDA!cota_vl_Valor
        Else
            'Verificar se existe cotação para a data base
            'If FunCotacao(XLD_DTBASEPAGTO, XLL_CODMOEDA, XLT_INDEXACAO, XLF_COTACAO) = False Then
            If FunCotacao(XLD_DTBASEPAGTO, XLL_CODMOEDA, XLT_INDEXACAO, XLF_COTACAOANTERIOR) = False Then
                XLO_TITULO.Close
                funCalculaCorrecaoMonetaria = 0
                Exit Function
            End If
        End If
        XLO_MOEDA.Close
        
        If XLO_TITULO!cont_tx_CorrecProRata = "A" Then
          XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
        Else
          XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAO
        End If
        
        XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASEPAGTO) - 1)) * XLF_COTACAO)
        funCalculaCorrecaoMonetaria = (XLF_COTACAO * XLF_COTACAOTITULO)
    End If
    
    XLO_TITULO.Close
End Function

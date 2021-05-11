Attribute VB_Name = "ModQualiCapiPagamento"
Option Explicit

Function FunCalculaJurosCEF(XLO_TITULO As adodb.Recordset, XLD_DATABASE As Date, XLF_VALOR As Double)
   
  If XLO_TITULO!cont_tx_Desagio = "S" Then
    FunCalculaJurosCEF = FunJurosCEF(XLF_VALOR, XLO_TITULO!cont_vl_PercJurosFin, _
             XLO_TITULO!titu_vl_PercJurosFin, _
             XLD_DATABASE, _
             XLO_TITULO!titu_dt_Vencimento)
  Else
    FunCalculaJurosCEF = 0
  End If
    
End Function

Function funCalculaCorrecaoMonetaria(XLO_TITULO As adodb.Recordset, XLI_INDEXADOR As Byte, XLD_DATABASE As Date) As Double
    
  If XLI_INDEXADOR = 1 And Not IsNull(XLO_TITULO!moed_cd_Moeda1) Then 'Indexador 1
    funCalculaCorrecaoMonetaria = FunCorrecaoMonetaria(XLO_TITULO!titu_vl_ParcelaIndex1, _
      FunNuloVal(XLO_TITULO!Cotacao1), FunNuloVal(XLO_TITULO!Cotacao1_MesAnterior), _
      XLO_TITULO!Indexacao1, XLO_TITULO!cont_tx_CorrecProRata, XLD_DATABASE, XLO_TITULO!titu_dt_Base, XLO_TITULO!moed_cd_Moeda1)
  ElseIf XLI_INDEXADOR = 2 And Not IsNull(XLO_TITULO!moed_cd_Moeda2) Then 'Indexador 2
    funCalculaCorrecaoMonetaria = FunCorrecaoMonetaria(XLO_TITULO!titu_vl_ParcelaIndex2, _
      FunNuloVal(XLO_TITULO!Cotacao2), FunNuloVal(XLO_TITULO!Cotacao2_MesAnterior), _
      XLO_TITULO!Indexacao2, XLO_TITULO!cont_tx_CorrecProRata, XLD_DATABASE, XLO_TITULO!titu_dt_Base, XLO_TITULO!moed_cd_Moeda2)
  Else
    funCalculaCorrecaoMonetaria = XLO_TITULO!titu_vl_Parcela
End If
'  If XLI_INDEXADOR = 1 And Not IsNull(XLO_TITULO!moed_cd_Moeda1) Then 'Indexador 1
'    funCalculaCorrecaoMonetaria = FunCorrecaoMonetaria(XLO_TITULO!titu_vl_ParcelaIndex1, _
'      FunNuloVal(XLO_TITULO!Cotacao1), FunNuloVal(XLO_TITULO!Cotacao1_MesAnterior), _
'      XLO_TITULO!Indexacao1, XLO_TITULO!cont_tx_CorrecProRata, XLD_DATABASE)
'  ElseIf XLI_INDEXADOR = 2 And Not IsNull(XLO_TITULO!moed_cd_Moeda2) Then 'Indexador 2
'    funCalculaCorrecaoMonetaria = FunCorrecaoMonetaria(XLO_TITULO!titu_vl_ParcelaIndex2, _
'      FunNuloVal(XLO_TITULO!Cotacao2), FunNuloVal(XLO_TITULO!Cotacao2_MesAnterior), _
'      XLO_TITULO!Indexacao2, XLO_TITULO!cont_tx_CorrecProRata, XLD_DATABASE)
'  Else
'    funCalculaCorrecaoMonetaria = XLO_TITULO!titu_vl_Parcela
'  End If

End Function

Function FunCalculaValorIndexado(XLO_TITULO As adodb.Recordset, XLI_INDEXADOR As Byte, XLD_DATABASE As Date) As Double
    
  'Verifica o indexador
  If XLI_INDEXADOR = 1 Then 'Indexador 1
    FunCalculaValorIndexado = FunValorIndexado(XLO_TITULO!titu_vl_ParcelaIndex1, _
      FunNuloVal(XLO_TITULO!Cotacao1), FunNuloVal(XLO_TITULO!Cotacao1_MesAnterior), _
      XLO_TITULO!Indexacao1, XLO_TITULO!cont_tx_CorrecProRata, XLD_DATABASE)
  Else 'Indexador 2
    FunCalculaValorIndexado = FunValorIndexado(XLO_TITULO!titu_vl_ParcelaIndex2, _
      FunNuloVal(XLO_TITULO!Cotacao2), FunNuloVal(XLO_TITULO!Cotacao2_MesAnterior), _
      XLO_TITULO!Indexacao2, XLO_TITULO!cont_tx_CorrecProRata, XLD_DATABASE)
  End If
  
End Function

Function FunCalculaValorIndexadoResiduo(XLO_TITULO As adodb.Recordset, XLI_INDEXADOR As Byte, XLD_DATABASE As Date, XLF_VALORRESIDUO As Double) As Double
    
  'Verifica o indexador
  If XLI_INDEXADOR = 1 Then 'Indexador 1
    FunCalculaValorIndexadoResiduo = FunValorIndexado(XLF_VALORRESIDUO, _
      FunNuloVal(XLO_TITULO!Cotacao1), FunNuloVal(XLO_TITULO!Cotacao1_MesAnterior), _
      XLO_TITULO!Indexacao1, XLO_TITULO!cont_tx_CorrecProRata, XLD_DATABASE)
  Else 'Indexador 2
    FunCalculaValorIndexadoResiduo = FunValorIndexado(XLF_VALORRESIDUO, _
      FunNuloVal(XLO_TITULO!Cotacao2), FunNuloVal(XLO_TITULO!Cotacao2_MesAnterior), _
      XLO_TITULO!Indexacao2, XLO_TITULO!cont_tx_CorrecProRata, XLD_DATABASE)
  End If
  
End Function

Function FunCalculaJuros(XLO_TITULO As adodb.Recordset, XLD_DATABASE As Date, XLF_VALOR As Double)
 
  FunCalculaJuros = 0
  'Verifica se a data base é maior que a do vencimento e se não venceu no fim de semana
  If FunVerificaVencimento(XLD_DATABASE, XLO_TITULO!titu_dt_Vencimento) Then
      
    FunCalculaJuros = FunJuros(XLF_VALOR, XLO_TITULO!titu_vl_PercJurosMora, _
      XLO_TITULO!cont_tx_TipoJurosMora, XLD_DATABASE, _
      XLO_TITULO!titu_dt_Vencimento, _
      XLO_TITULO!titu_vl_Parcela)
  End If
  
End Function

Function FunCalculaMulta(XLO_TITULO As adodb.Recordset, XLD_DATABASE As Date, XLF_VALOR As Double)
  
  FunCalculaMulta = 0
  'Verifica se a data base é maior que a do vencimento e se não venceu no fim de semana
  If FunVerificaVencimento(XLD_DATABASE, XLO_TITULO!titu_dt_Vencimento) Then
    FunCalculaMulta = FunMulta(XLF_VALOR, XLO_TITULO!titu_vl_PercMultaMora, _
      XLO_TITULO!cont_tx_MultaProRata, _
      XLO_TITULO!cont_tx_TipoJurosMora, XLD_DATABASE, _
      XLO_TITULO!titu_dt_Vencimento, _
      XLO_TITULO!titu_vl_Parcela, _
      XLO_TITULO!cont_nr_LimiteDias)
   End If
   
End Function

'Chama a função de cálculo do deságio
Function funCalculaDesagio(XLO_TITULO As adodb.Recordset, XLD_DATABASE As Date, XLF_VALOR As Double) As Double
    
  If XLO_TITULO!cont_tx_Desagio <> "S" Then
    funCalculaDesagio = FunDesagio(XLF_VALOR, XLO_TITULO!titu_vl_PercJurosFin, _
      XLO_TITULO!cont_tx_Desagio, XLD_DATABASE, _
      XLO_TITULO!titu_dt_Vencimento)
  Else
    funCalculaDesagio = 0
  End If
    
End Function

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

Sub SubContabilizarPromissoria(XLL_CODDUPLICATA As Long, XLL_FATURA As Long, XLL_CODCCORRENTE As Long, XLT_DOCREC As String, XLD_PAGTOREC As Date, XLT_TIPOPAGTO As String, XLB_CONSOLIDADO As Boolean)
 
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
    Dim XLO_VLCONSOLIDADO As adodb.Recordset  'Recordset para obter o valor do pagto consolidado
    Dim XLO_PARCESPECIAL As adodb.Recordset  'Recordset para tipos de documentos como TB, DEV, REST etc.
    Dim XLO_CHECK As adodb.Recordset         'Recordset para  definir o lote
    Dim XLO_DESCONTOS As adodb.Recordset      'Recordset para a tabela de descontoas/acrescimos
    Dim XLB_PROVISIONADO As Boolean 'Indica se o lançamento foi provisionado
    Dim XLT_SQL As String           'Prepara Select
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
    Dim XLF_TOTALPORCODDESC As Double
    Dim XLL_CODDESC As Long
    Dim XLL_CODDESCPROX As Long
    
    'Pegar o número do Lote e Lancamento para cadastrar no mesmo lote do dia
    'Origem = C - Capi
    XLT_SQL = "SELECT lanc_nr_Lote FROM Lancamentos " & _
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
    Set XLO_CHECK = Nothing
    
    'pegar o valor total para posterior contabilização da parte do banco
    XLT_SQL = "SELECT sum(dupl_vl_valor) AS ValorChCons, " & _
            " COUNT (dupl_cd_duplicata) AS TotalCheques FROM Duplicatas " & _
            " WHERE dupl_tx_tipopagrec = '" & XLT_TIPOPAGTO & "' AND coco_cd_codigo = " & XLL_CODCCORRENTE & " " & _
            " AND dupl_tx_docpagrec= '" & XLT_DOCREC & "'"
    SubQOpenRecordset XLO_VLCONSOLIDADO, XLT_SQL, Estatico
    XLF_VLCONSOLIDADO = XLO_VLCONSOLIDADO!ValorChCons
    XLO_VLCONSOLIDADO.Close
    Set XLO_VLCONSOLIDADO = Nothing
    
    XLT_SQL = "SELECT PlcoContaCC, cere_cd_Pcr, ContaRecDesp, nofi_tx_historico, focl_tx_RazaoSocial, nofi_nr_documento, iden_cd_Projeto, dupl_dt_pagrec, desconto, Acrescimo, distvalor, dupl_cd_Duplicata FROM ConsFINExpContabilidade2 " & _
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
                XLT_HISTORICO = Mid(FunNulo(XLO_PARCELA!nofi_tx_historico) + "-" + XLO_PARCELA!focl_tx_RazaoSocial + "-Doc: " + XLO_PARCELA!nofi_nr_documento, 1, 120)
            Else
                XLT_HISTORICO = Mid(FunNulo(XLO_PARCELA!nofi_tx_historico) + "-" + XLO_PARCELA!focl_tx_RazaoSocial + "-Doc: " + XLO_PARCELA!nofi_nr_documento, 1, 60)
            End If
            
            XLT_HISTCUSTO = FunNulo(XLO_PARCELA!nofi_tx_historico)
            XLI_IDENTIFICADOR = FunNuloVal(XLO_PARCELA!iden_cd_Projeto)
            XLD_PAGTOREC = XLO_PARCELA!dupl_dt_pagrec
            XLF_VALORDESC = XLO_PARCELA!desconto + XLO_PARCELA!Acrescimo
            XLF_VALORDUPL = XLO_PARCELA!distvalor
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
            
            '******************************DESCONTOS/ACRÉSCIMOS******************************
            If XLO_PARCELA!desconto <> 0 Or XLO_PARCELA!Acrescimo <> 0 Then
                XLT_SQL = "SELECT deac_cd_descacresc, deac_tx_Descricao, deac_tx_Descricao, natu_cd_Natureza, plco_cd_Conta, cere_cd_Pcr, dude_vl_descacresc, TOTALDESCACRESC, deac_tx_tipo FROM ConsFINExpDescacresc WHERE dupl_cd_duplicata = " & XLO_PARCELA!dupl_cd_Duplicata & " and empr_cd_empresa = " & PCodEmpresa & ""
                SubQOpenRecordset XLO_DESCONTOS, XLT_SQL, Estatico
                
                If XLO_DESCONTOS.EOF Then
                    MsgBox "O Desconto/Acréscimo da parcela não possui conta contábil. Peça ao seu administrador para cadastrar a conta contábil pelo sistema QualiFin.", vbCritical, "ATENÇÃO!"
'                    XGB_ContinaPagamento = False 'Cancela o pagamento do título
                    XLO_DESCONTOS.Close
                    Set XLO_DESCONTOS = Nothing
                    Exit Sub
                Else
                    If IsNull(XLO_DESCONTOS!plco_cd_Conta) Or XLO_DESCONTOS!plco_cd_Conta = "" Then
                        MsgBox "O Desconto/Acréscimo '" & XLO_DESCONTOS!deac_tx_Descricao & "' não possui conta contábil. Peça ao seu administrador para cadastrar a conta contábil pelo sistema QualiFin.", vbCritical, "ATENÇÃO!"
'                        XGB_ContinaPagamento = False 'Cancela o pagamento do título
                        
                        XLO_DESCONTOS.Close
                        Set XLO_DESCONTOS = Nothing
                        Exit Sub
                    End If
                End If
                
                XLF_TOTALPORCODDESC = 0
                While Not XLO_DESCONTOS.EOF
                    XLL_CODDESC = XLO_DESCONTOS!deac_cd_descacresc
                    XLL_CODDESCPROX = XLO_DESCONTOS!deac_cd_descacresc
                    If XGB_PQUALICONT Then 'possui contabilidade, o histórico pode ser maior
                        If Mid(XLO_PARCELA!nofi_nr_documento, 1, 1) = "A" Then
                            XLT_HISTORICO = Mid(XLO_DESCONTOS!deac_tx_Descricao + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 120)
                        Else
                            XLT_HISTORICO = Mid(XLO_DESCONTOS!deac_tx_Descricao + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial) + "-Doc: " + XLO_PARCELA!nofi_nr_documento, 1, 120)
                        End If
                    Else
                        If Mid(XLO_PARCELA!nofi_nr_documento, 1, 1) = "A" Then
                            XLT_HISTORICO = Mid(XLO_DESCONTOS!deac_tx_Descricao + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial), 1, 60)
                        Else
                            XLT_HISTORICO = Mid(XLO_DESCONTOS!deac_tx_Descricao + "-" + FunNulo(XLO_PARCELA!focl_tx_RazaoSocial) + "-Doc: " + XLO_PARCELA!nofi_nr_documento, 1, 60)
                        End If
                    End If
                   
                    'Se for receita ou despesa, deve informar o Centro de custo
                    If XLO_DESCONTOS!natu_cd_Natureza = 3 Or XLO_DESCONTOS!natu_cd_Natureza = 4 Then
                        If XLO_DESCONTOS!deac_tx_tipo = "A" Then 'Acréscimo
                            XLT_CONTACREDITO = XLO_DESCONTOS!plco_cd_Conta
                            XLT_PCRCREDITO = XLO_DESCONTOS!cere_cd_Pcr
                            XLT_PCRDEBITO = ""
                            XLT_CONTADEBITO = ""
                        Else 'Desconto
                            XLT_CONTADEBITO = XLO_DESCONTOS!plco_cd_Conta
                            XLT_PCRDEBITO = XLO_DESCONTOS!cere_cd_Pcr
                            XLT_PCRCREDITO = ""
                            XLT_CONTACREDITO = ""
                        End If
                        XLF_VALOR = ((XLO_DESCONTOS!dude_vl_descacresc * 100) \ 1) / 100
                        XLF_TOTALPORCODDESC = XLF_TOTALPORCODDESC + XLF_VALOR
                        
                        XLO_DESCONTOS.MoveNext
                        If Not XLO_DESCONTOS.EOF Then
                            If XLL_CODDESC <> XLO_DESCONTOS!deac_cd_descacresc Then
                                XLO_DESCONTOS.MovePrevious
                                If XLF_TOTALPORCODDESC > XLO_DESCONTOS!TOTALDESCACRESC Then
                                    XLF_VALOR = XLF_VALOR - 0.01
                                ElseIf XLF_TOTALPORCODDESC < XLO_DESCONTOS!TOTALDESCACRESC Then
                                    XLF_VALOR = XLF_VALOR + (XLO_DESCONTOS!TOTALDESCACRESC - XLF_TOTALPORCODDESC)
                                End If
                                XLF_TOTALPORCODDESC = 0
                            Else
                                XLO_DESCONTOS.MovePrevious
                            End If
                        Else
                            XLO_DESCONTOS.MovePrevious
                            If XLF_TOTALPORCODDESC > XLO_DESCONTOS!TOTALDESCACRESC Then
                                XLF_VALOR = XLF_VALOR - 0.01
                            Else
                                XLF_VALOR = XLF_VALOR + (XLO_DESCONTOS!TOTALDESCACRESC - XLF_TOTALPORCODDESC)
                            End If
                            XLF_TOTALPORCODDESC = 0
                        End If
                    Else
                        '***********
                        If XLO_DESCONTOS!deac_tx_tipo = "A" Then 'Acréscimo
                            XLT_CONTADEBITO = XLO_DESCONTOS!plco_cd_Conta
                            XLT_PCRDEBITO = ""
                            XLT_PCRCREDITO = ""
                            XLT_CONTACREDITO = ""
                        Else 'Desconto
                            XLT_CONTACREDITO = XLO_DESCONTOS!plco_cd_Conta
                            XLT_PCRCREDITO = ""
                            XLT_PCRDEBITO = ""
                            XLT_CONTADEBITO = ""
                        End If
                        '*************
                        XLF_VALOR = 0
                        While XLL_CODDESC = XLL_CODDESCPROX And Not XLO_DESCONTOS.EOF
                            XLF_VALOR = XLF_VALOR + ((XLO_DESCONTOS!dude_vl_descacresc * 100) \ 1) / 100
                            XLO_DESCONTOS.MoveNext
                            If Not XLO_DESCONTOS.EOF Then
                                XLL_CODDESCPROX = XLO_DESCONTOS!deac_cd_descacresc
                                If XLL_CODDESC = XLL_CODDESCPROX Then
                                    XLF_VALORDESC = XLF_VALORDESC + ((XLO_DESCONTOS!dude_vl_descacresc * 100) \ 1) / 100
                                End If
                            End If
                        Wend
                        XLO_DESCONTOS.MovePrevious
                        
                        If XLF_VALOR > XLO_DESCONTOS!TOTALDESCACRESC Then
                            XLF_VALOR = XLF_VALOR - 0.01
                        Else
                            XLF_VALOR = XLF_VALOR + (XLO_DESCONTOS!TOTALDESCACRESC - XLF_VALOR)
                        End If
                        XLF_TOTALPORCODDESC = 0
                    End If
                    
                    XLI_CODLANC = funDefinirCodigoLanc(XLD_PAGTOREC, XLI_LOTE)
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
                            "'" & FunNulo(FunTrataString(XLT_HISTORICO)) & "'," & FunTrataFloat(XLF_VALOR) & "," & _
                            "" & PCodEmpresa & ",'" & FunTrataString(FunNulo(XLO_PARCELA!nofi_tx_historico)) & "'," & _
                            "" & FunNuloBanco(XLO_PARCELA!iden_cd_Projeto) & "," & XLO_PARCELA!dupl_cd_Duplicata & ",'C')"
                    Conexao.Execute XLT_SQL
                    XLO_DESCONTOS.MoveNext
                Wend
                
                XLO_DESCONTOS.Close
                Set XLO_DESCONTOS = Nothing
            End If
'////////////////////
            XLO_PARCELA.MoveNext
        Wend
        XLO_PARCELA.Close
        Set XLO_PARCELA = Nothing
        

        '*********************PARTE DO BANCO DA NOTA PROMISSÓRIA********************
        XLT_HISTORICO = ""
        XLT_CONTACREDITO = ""
        XLT_PCRCREDITO = ""
        XLT_PCRDEBITO = ""
        XLT_HISTORICO = "Total diário de Notas Promissória " + Format(CStr(XLD_PAGTOREC), "dd/MM/yy")
        
        'Obter o total de NotasPromissórias
        XLT_SQL = "SELECT sum(titu_vl_ValorPago) AS TotalNP " & _
                    " FROM titulos INNER JOIN" & _
                    " Contratos ON Titulos.cont_cd_Contrato = Contratos.cont_cd_Contrato AND " & _
                    " Titulos.imov_cd_Imovel = Contratos.imov_cd_Imovel AND " & _
                    " Titulos.empd_cd_Empreendimento = Contratos.empd_cd_Empreendimento AND " & _
                    " Titulos.empr_cd_Empresa = Contratos.empr_cd_Empresa " & _
                    " WHERE Contratos.cont_tx_Status = 'A' AND " & _
                    " titu_dt_deposito = " & FunNuloData(XLD_PAGTOREC, NomeSgbd) & " " & _
                    " AND Titulos.empr_cd_empresa = " & PCodEmpresa & " AND Titulos.coco_cd_codigo = " & XLL_CODCCORRENTE & " "
        SubQOpenRecordset XLO_PARCELA, XLT_SQL, Dinamico
        
        If Not XLO_PARCELA.EOF Then
            XLF_VLCONSOLIDADO = Format(XLO_PARCELA!TotalNP, "standard")
        End If
        
        XLO_PARCELA.Close
        Set XLO_PARCELA = Nothing
        
        XLT_SQL = "SELECT * FROM Lancamentos " & _
                    " WHERE plco_cd_Debito = '" & XLT_CONTADEBITOCC & "'" & _
                    " AND lanc_dt_Data =" & FunNuloData(XLD_PAGTOREC, NomeSgbd) & "" & _
                    " AND empr_cd_empresa = " & PCodEmpresa & "  and lanc_tx_origem = 'C'"
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
        XLO_PARCELA!lanc_vl_valor = XLF_VLCONSOLIDADO
        XLO_PARCELA!empr_cd_empresa = PCodEmpresa
        XLO_PARCELA!lanc_tx_DescCustos = XLT_HISTCUSTO
        XLO_PARCELA!iden_cd_Projeto = funNull(XLI_IDENTIFICADOR)
        XLO_PARCELA!lanc_cd_OrigemFin = XLL_CODDUPLICATA
        XLO_PARCELA!lanc_tx_Origem = "C"
                                        
        XLO_PARCELA.Update
        XLO_PARCELA.Close
        
        Set XLO_PARCELA = Nothing
    End If
End Sub

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
    Dim XLO_REGISTRO As adodb.Recordset 'Recordset para os acessos às tabelas do Qualifin
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
    Dim XLT_TIPOPAGTO As String     'Tipo de Pagto/rec (se cheque, dinheiro...)
    Dim XLI_REGATUALIZADOS As Integer 'Indica quantos registros foram atualizados na tabela de Distribuição
    
    
    
    'Obter código do tipo de doc cuja sigla = NP (Nota Promissória)
    SubQOpenRecordset XLO_REGISTRO, "SELECT tido_cd_TipoDoc FROM TipoDocumentos WHERE tido_tx_Sigla= 'NP'", Estatico
    XLI_TIPODOC = XLO_REGISTRO!tido_cd_TipoDoc
    XLO_REGISTRO.Close
    
    'Obter Dt de registro do contrato e dt. venda do imovel
    SubQOpenRecordset XLO_REGISTRO, "SELECT cont_dt_RegistroContrato, cont_dt_Venda FROM Contratos WITH (ROWLOCK)  " & _
                " WHERE empd_cd_Empreendimento= '" & XLO_TITULO!empd_cd_Empreendimento & "' AND " & _
                " imov_cd_Imovel='" & XLO_TITULO!imov_cd_Imovel & "' AND cont_cd_Contrato='" & XLO_TITULO!cont_cd_Contrato & "' AND " & _
                " empr_cd_Empresa = " & PCodEmpresa & "", Estatico
    XLD_REGCONTRATO = XLO_REGISTRO!cont_dt_RegistroContrato
    XLD_VENDACONTRATO = XLO_REGISTRO!cont_dt_Venda
    XLO_REGISTRO.Close
    
    XLT_TITULO = XLO_TITULO!empd_cd_Empreendimento + "." + XLO_TITULO!imov_cd_Imovel + "." + XLO_TITULO!cont_cd_Contrato + "." + XLO_TITULO!titu_cd_plano + "." + XLO_TITULO!titu_cd_Parcela + "." + XLO_TITULO!titu_cd_residuo
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
        SubQOpenRecordset XLO_REGISTRO, "SELECT * FROM NotasFiscais WHERE nofi_nr_TituloCapi= '" & XLT_TITULO & "' AND empr_cd_Empresa = " & PCodEmpresa & "", Dinamico
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
        XLF_VALOR = Format(XLO_TITULO!titu_vl_ValorPago - FunNuloVal(XLO_TITULO!titu_vl_Seguro) - FunNuloVal(XLO_TITULO!titu_vl_Juros) - FunNuloVal(XLO_TITULO!titu_vl_Multa) - FunNuloVal(XLO_TITULO!titu_vl_Outros) + FunNuloVal(XLO_TITULO!titu_vl_Desagio) + FunNuloVal(XLO_TITULO!titu_vl_Desconto), "standard")
        XLO_REGISTRO!nofi_dt_entrada = Format$(XLO_TITULO!titu_dt_Pagamento, "DD/MM/YY")
        XLO_REGISTRO!nofi_dt_emissao = Format$(XLD_REGCONTRATO, "DD/MM/YY")
    End If
    
    XLO_REGISTRO!nofi_tx_tipomov = "E"   'Entrada
    XLO_REGISTRO!nofi_tx_status = XLT_STATUSGRAVAR
    XLO_REGISTRO!focl_cd_FornCli = XLI_CODCLIENTE
    XLO_REGISTRO!nofi_tx_serie = "U" 'Única
    XLO_REGISTRO!tido_cd_TipoDoc = XLI_TIPODOC
    XLO_REGISTRO!nofi_tx_historico = Mid("PARC. " + XLT_TITULO + " - " + XLT_TIPOPLANO + " - " + XLT_NATUREZAPLANO, 1, 60)
    XLO_REGISTRO!nofi_vl_valor = XLF_VALOR
    XLO_REGISTRO!nofi_tx_distribuir = "N" 'Sempre 1 conta para 1 CCusto
    XLO_REGISTRO!nofi_tx_tipofatura = "1" 'Uma nota para uma duplicata
    XLO_REGISTRO!nofi_nr_TituloCapi = XLT_TITULO 'Código do título
    XLO_REGISTRO!empr_cd_empresa = PCodEmpresa

    XLO_REGISTRO.Update

    Chave = XLO_REGISTRO!nofi_cd_NotaFiscal
    XLL_FATURA = XLO_REGISTRO!nofi_nr_fatura
    XLO_REGISTRO.Close
        
    '*****************************TITULOS E NOTAS FISCAL*******************************
    'Gravar a nofi_cd_notafiscal na tabela de títulos
    Conexao.Execute "UPDATE Titulos SET nofi_cd_notafiscal = " & Chave & " WHERE empd_cd_Empreendimento= '" & XLO_TITULO!empd_cd_Empreendimento & "' AND " & _
                " imov_cd_Imovel='" & XLO_TITULO!imov_cd_Imovel & "' AND cont_cd_Contrato='" & XLO_TITULO!cont_cd_Contrato & "' AND " & _
                " titu_cd_plano = '" & XLO_TITULO!titu_cd_plano & "' AND titu_cd_Parcela = '" & XLO_TITULO!titu_cd_Parcela & "' AND " & _
                " titu_cd_residuo = '" & XLO_TITULO!titu_cd_residuo & "' AND empr_cd_Empresa = " & PCodEmpresa & ""
    
        
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
        XLT_TIPOPAGTO = "6"
    Else
        XLO_REGISTRO!dupl_tx_tipopagrec = "5" 'Crédito em Conta
        XLT_TIPOPAGTO = "5"
    End If
    XLO_REGISTRO!dupl_tx_DocPagRec = XLT_DOCREC
    XLO_REGISTRO!coco_cd_codigo = XLO_TITULO!coco_cd_codigo
    XLO_REGISTRO!dupl_tx_debcred = "C"
    XLO_REGISTRO!dupl_nr_fatura = XLL_FATURA
    XLO_REGISTRO!dupl_dt_pagrec = XLO_TITULO!titu_dt_Deposito
    
    XLO_REGISTRO.Update
    
    XLL_CODDUPL = XLO_REGISTRO!dupl_cd_Duplicata
    
    XLO_REGISTRO.Close
            
    '*********************************GRAVA OS DESCONTOS*************************************
    If XLO_TITULO!titu_vl_Multa <> 0 Then
        Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc) VALUES (" & XLL_CODDUPL & "," & XGI_CODMULTA & "," & FunTrataFloat(XLO_TITULO!titu_vl_Multa) & ")"
    End If
    
    If XLO_TITULO!titu_vl_Juros <> 0 Then
        Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc) VALUES (" & XLL_CODDUPL & "," & XGI_CODJUROS & "," & FunTrataFloat(XLO_TITULO!titu_vl_Juros) & ")"
    End If
    
    If XLO_TITULO!titu_vl_Desagio <> 0 Then
        Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc) VALUES (" & XLL_CODDUPL & "," & XGI_CODDESAGIO & "," & FunTrataFloat(XLO_TITULO!titu_vl_Desagio) & ")"
    End If

    If XLO_TITULO!titu_vl_Desconto <> 0 Then
        Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc) VALUES (" & XLL_CODDUPL & "," & XGI_CODDESCONTO & "," & FunTrataFloat(XLO_TITULO!titu_vl_Desconto) & ")"
    End If

    If XLO_TITULO!titu_vl_Seguro <> 0 Then
        Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc) VALUES (" & XLL_CODDUPL & "," & XGI_CODSEGURO & "," & FunTrataFloat(XLO_TITULO!titu_vl_Seguro) & ")"
    End If
    
    If XLO_TITULO!titu_vl_Outros <> 0 Then
        Conexao.Execute "INSERT INTO DuplDescAcresc(dupl_cd_duplicata, deac_cd_descacresc, dude_vl_descacresc) VALUES (" & XLL_CODDUPL & "," & XGI_CODOUTROS & "," & FunTrataFloat(XLO_TITULO!titu_vl_Outros) & ")"
    End If
        
    
    '*****************************DISTRIBUIÇÃO*******************************
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
        
        XLI_REGATUALIZADOS = 0
        XLT_SQL = "UPDATE Distribuicao SET cere_cd_pcr='" & XLO_TITULO!cere_cd_Pcr & "'," & _
                        "plco_cd_conta='" & XLO_TITULO!plco_cd_Conta & "'," & _
                        "dist_vl_valor=" & FunTrataFloat(XLF_VALOR) & "," & _
                        "iden_cd_Projeto= " & FunNuloBancoVal(XLO_TITULO!iden_cd_Projeto) & "" & _
                    " WHERE nofi_cd_NotaFiscal = " & Chave & ""
        Conexao.Execute XLT_SQL, XLI_REGATUALIZADOS
        
        'Se o compromisso não foi exportado anteriormente, será necessário inserir a distribuição
        If XLI_REGATUALIZADOS = 0 Then
            XLT_SQL = "INSERT INTO Distribuicao(cere_cd_pcr,plco_cd_conta, " & _
                        "dist_vl_valor,iden_cd_Projeto,nofi_cd_notafiscal)" & _
                    " VALUES ('" & XLO_TITULO!cere_cd_Pcr & "'," & _
                        "'" & XLO_TITULO!plco_cd_Conta & "'," & _
                        "" & FunTrataFloat(XLF_VALOR) & "," & _
                        "" & FunNuloBancoVal(XLO_TITULO!iden_cd_Projeto) & "," & Chave & ")"
            Conexao.Execute XLT_SQL
        End If
    End If
    'XLO_TITULO.Close
    If XLT_STATUS = "PAGAR" Then
        SubContabilizarPromissoria XLL_CODDUPL, XLL_FATURA, XLO_TITULO!coco_cd_codigo, XLT_DOCREC, XLO_TITULO!titu_dt_Deposito, XLT_TIPOPAGTO, True
    End If
End Sub



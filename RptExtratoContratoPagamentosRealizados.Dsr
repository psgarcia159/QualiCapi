VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptExtratoContratoPagamentosRealizados 
   ClientHeight    =   13440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14010
   OleObjectBlob   =   "RptExtratoContratoPagamentosRealizados.dsx":0000
End
Attribute VB_Name = "RptExtratoContratoPagamentosRealizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Dim VFV_SOMATORIO As New XArray ' Vetor que armazena os somatórios
Dim VFV_VALORES As New XArray  ' Vetor que armazena os valores que precisarão fazer somatório
Dim VFV_TOTALMOEDAPREVISTO As New XArray
Dim VFV_TOTALMOEDAORIGINAL As New XArray
Dim VFV_TOTALMOEDALIQUIDADO As New XArray
Dim VFV_TOTALMOEDADISTRATO As New XArray

Dim XFB_IMPRIMIU As Boolean 'Indica se chegou no final do relatório na primeira vez
Dim XFT_TITULOANTERIOR As String

Private Sub Report_Initialize()

  AddReportVariable crRVNumber, "SOMAVALORPREVISTO"
  AddReportVariable crRVNumber, "SOMAVALORORIGINAL"
  AddReportVariable crRVNumber, "SOMADISTRATO"
  AddReportVariable crRVNumber, "TOTALPREVISTOATUALIZADO" 'Usado para somar o valor previsto atualizado dos titulos
  AddReportVariable crRVNumber, "TOTALORIGINALATUALIZADO" 'Usado para somar o valor original atualizado dos titulos
  AddReportVariable crRVNumber, "TOTALLIQUIDADOATUALIZADO" 'Usado para somar o valor pago atualizado dos titulos
  AddReportVariable crRVNumber, "TOTALDISTRATOATUALIZADO" 'Usado para somar o valor p/distarto dos titulos
  
  XFB_IMPRIMIU = False
  VFV_TOTALMOEDAPREVISTO.ReDim 0, 0, 0, 3
  VFV_TOTALMOEDAORIGINAL.ReDim 0, 0, 0, 3
  VFV_TOTALMOEDALIQUIDADO.ReDim 0, 0, 0, 3
  VFV_TOTALMOEDADISTRATO.ReDim 0, 0, 0, 3
  VFV_SOMATORIO.ReDim 0, 0, 0, 7  'colunas = quantidade de campos para soma
  VFV_VALORES.ReDim 0, 0, 0, 6 'colunas = quantidade de campos para soma - 1

End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SectHeaderReport_Format(ByVal pFormattingInfo As Object)
    
  'Zera as variáveis quando passa pelo Header do relatório
  SetReportVariableValue "SOMAVALORPREVISTO", 0
  SetReportVariableValue "SOMAVALORORIGINAL", 0
  SetReportVariableValue "SOMADISTRATO", 0
  SetReportVariableValue "TOTALPREVISTOATUALIZADO", 0
  SetReportVariableValue "TOTALORIGINALATUALIZADO", 0
  SetReportVariableValue "TOTALLIQUIDADOATUALIZADO", 0
  SetReportVariableValue "TOTALDISTRATOATUALIZADO", 0
  
  
End Sub

Private Sub SecDetalhe_Format(ByVal pFormattingInfo As Object)
  
  Dim XLF_VALORCORRIGIDO As Double  'Valor corrigido monetariamente
  Dim XLF_DESAGIO As Double  'Valor do deságio
  Dim XLF_MULTA As Double  'Valor da Multa
  Dim XLF_JUROS As Double  'Valor do Juros
  
  Dim XLF_VALORTITULO As Double 'Valor do Titulo
  Dim XLF_VALORPREVISTO As Double 'Valor Previsto
  Dim XLF_VALORORIGINAL As Double 'Valor Original
  Dim XLF_VALORPAGO As Double 'Valor pago
  Dim XLF_VALORDISTRATO As Double 'Valor do distrato
  
  Dim XLF_VALORPAGOINDEXADO As Double 'Valor pago indexado
  Dim XLF_VALORORIGINALINDEXADO As Double 'Valor original indexado
  Dim XLF_VALORDISTRATOINDEXADO As Double 'Valor original indexado
  
  Dim XLF_VALORPREVISTOATUALIZADO As Double 'Valor previsto atualizado na correção da data base
  Dim XLF_VALORPAGOATUALIZADO As Double 'Valor pago atualizado na correção da data base
  Dim XLF_VALORORIGINALATUALIZADO As Double 'Valor original atualizado na correção da data base
  Dim XLF_VALORDISTRATOATUALIZADO As Double 'Valor distrato atualizado na correção da data base
  
  Dim XLI_MOEDA As Integer 'Código da moeda de acordo com o indexador escolhido
  Dim XLT_MOEDA As String 'Decrição da moeda de acordo com o indexador escolhido
  
  'Vetor auxiliar para armazenar o valor do titulo, a cotação e a cotação do mês anterior,
  'tudo de acordo com o indexador escolhido. Também indica se tem indexador ou não.
  Dim XLO_VETOR As New XArray
  Dim XLO_VETOR2 As New XArray
     
  Dim XLD_DATABASE As Date  'Data base passada por parâmetro para relatório
  Dim XLI_INDEXADOR As Byte  ' Indexador passado por parâmetro para o relatório
  Dim XLT_FILTROMULTA As String
  Dim XLT_FILTROJUROS As String
  Dim XLT_FILTRODESAGIO As String
  Dim XLT_FILTROTITULOSPAGOS As String
  
  Dim XLO_CAMPOS As CRAXDDRT.DatabaseFieldDefinitions  ' Objeto que representa os campos do relatório
   
  Set XLO_CAMPOS = Me.Database.Tables(1).Fields
  
  XLI_INDEXADOR = 1 'Só para compor os parâmetros das fórmulas
  
  'Pega o código da moeda e a sua descrição
  SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  XLT_MOEDA = XLO_VETOR(0, 0)
  XLI_MOEDA = XLO_VETOR(1, 0)
  
  'Pega as cotações do mês da data base e do mês anterior
  SubCotacao XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  
  XLD_DATABASE = XLO_CAMPOS.GetItemByName("titu_dt_BasePagto").Value
  
  XLF_VALORCORRIGIDO = Format(FunCalculoCorrecaoMonetaria(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLI_MOEDA), "Standard")
  
  'Juros da CEF
  XLF_VALORCORRIGIDO = XLF_VALORCORRIGIDO + Format(FunCalculoJurosCEF(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  
  XLF_JUROS = Format(FunCalculoJuros(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  XLF_MULTA = Format(FunCalculoMulta(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  XLF_DESAGIO = Format(FunCalculoDesagio(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
    
  XLF_VALORPREVISTO = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA - XLF_DESAGIO
  
  If XGT_INCLUIRACRESCIMOS = "S" Then
    XLF_VALORTITULO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORPREVISTO), "##,##0.0000")
  Else
    XLF_VALORTITULO = Format(XLO_VETOR(0, 0), "##,##0.0000")
  End If
  
  XLF_VALORORIGINAL = Format(FunCalculoValorOriginalPago(XLO_CAMPOS), "standard")
  XLF_VALORPAGO = Format(FunCalculoValorPago(XLO_CAMPOS, "1"), "standard")
  XLF_VALORDISTRATO = Format(FunCalculoValorDistrato(XLO_CAMPOS), "standard")
   
  'Pega a cotação para calcular o valor indexado na data do pagamento
  If FormulaFields.GetItemByName("IndexadorExtra").Value <> "" Then
    XLO_VETOR(0, 0) = 0
    XLO_VETOR(1, 0) = CDbl(XLO_CAMPOS.GetItemByName("Cotacao5").Value)
    XLO_VETOR(2, 0) = CDbl(XLO_CAMPOS.GetItemByName("Cotacao5_MesAnterior").Value)
    XLO_VETOR(3, 0) = ""
    XLO_VETOR(4, 0) = 0
  Else
    SubCotacao3 XLO_CAMPOS, XLO_VETOR
  End If
  
  XLF_VALORORIGINALINDEXADO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORORIGINAL), "##,##0.0000")
  XLF_VALORPAGOINDEXADO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORPAGO), "##,##0.0000")
  XLF_VALORDISTRATOINDEXADO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORDISTRATO), "##,##0.0000")
  
  'Pega a cotação para calcular o valor atualizado com base no valor indexado
  If FormulaFields.GetItemByName("IndexadorExtra").Value <> "" Then
    XLO_VETOR(0, 0) = 0
    XLO_VETOR(1, 0) = CDbl(XLO_CAMPOS.GetItemByName("Cotacao4").Value)
    XLO_VETOR(2, 0) = CDbl(XLO_CAMPOS.GetItemByName("Cotacao4_MesAnterior").Value)
    XLO_VETOR(3, 0) = ""
    XLO_VETOR(4, 0) = 0
  Else
    SubCotacao2 XLO_CAMPOS, XLO_VETOR
  End If
  
  XLF_VALORPREVISTOATUALIZADO = Format(FunCalculoValorAtualizado(XLO_CAMPOS, XLO_VETOR, CDate(FormulaFields.GetItemByName("FiltroDataBase").Value), XLF_VALORTITULO, XLI_MOEDA), "standard")
  XLF_VALORORIGINALATUALIZADO = Format(FunCalculoValorAtualizado(XLO_CAMPOS, XLO_VETOR, CDate(FormulaFields.GetItemByName("FiltroDataBase").Value), XLF_VALORORIGINALINDEXADO, XLI_MOEDA), "standard")
  XLF_VALORPAGOATUALIZADO = Format(FunCalculoValorAtualizado(XLO_CAMPOS, XLO_VETOR, CDate(FormulaFields.GetItemByName("FiltroDataBase").Value), XLF_VALORPAGOINDEXADO, XLI_MOEDA), "standard")
  XLF_VALORDISTRATOATUALIZADO = Format(FunCalculoValorAtualizado(XLO_CAMPOS, XLO_VETOR, CDate(FormulaFields.GetItemByName("FiltroDataBase").Value), XLF_VALORDISTRATOINDEXADO, XLI_MOEDA), "standard")
     
  ' Se for o mesmo titulo que o anterior não soma no vetor
  If XFT_TITULOANTERIOR <> XLO_CAMPOS.GetItemByName("Titulo").Value And XFB_IMPRIMIU = False Then
    If FormulaFields.GetItemByName("IndexadorExtra").Value <> "" Then
      SubSomaMoeda2 FormulaFields.GetItemByName("IndexadorExtra").Value, FormulaFields.GetItemByName("NomeIndexadorExtra").Value, VFV_TOTALMOEDAPREVISTO, XLF_VALORTITULO, Left(XLO_CAMPOS.GetItemByName("Titulo").Value, 12)
      SubSomaMoeda2 FormulaFields.GetItemByName("IndexadorExtra").Value, FormulaFields.GetItemByName("NomeIndexadorExtra").Value, VFV_TOTALMOEDAORIGINAL, XLF_VALORORIGINALINDEXADO, Left(XLO_CAMPOS.GetItemByName("Titulo").Value, 12)
      SubSomaMoeda2 FormulaFields.GetItemByName("IndexadorExtra").Value, FormulaFields.GetItemByName("NomeIndexadorExtra").Value, VFV_TOTALMOEDALIQUIDADO, XLF_VALORPAGOINDEXADO, Left(XLO_CAMPOS.GetItemByName("Titulo").Value, 12)
      SubSomaMoeda2 FormulaFields.GetItemByName("IndexadorExtra").Value, FormulaFields.GetItemByName("NomeIndexadorExtra").Value, VFV_TOTALMOEDADISTRATO, XLF_VALORDISTRATOINDEXADO, Left(XLO_CAMPOS.GetItemByName("Titulo").Value, 12)
    Else
      SubSomaMoeda2 XLI_MOEDA, XLT_MOEDA, VFV_TOTALMOEDAPREVISTO, XLF_VALORTITULO, Left(XLO_CAMPOS.GetItemByName("Titulo").Value, 12)
      SubSomaMoeda2 XLI_MOEDA, XLT_MOEDA, VFV_TOTALMOEDAORIGINAL, XLF_VALORORIGINALINDEXADO, Left(XLO_CAMPOS.GetItemByName("Titulo").Value, 12)
      SubSomaMoeda2 XLI_MOEDA, XLT_MOEDA, VFV_TOTALMOEDALIQUIDADO, XLF_VALORPAGOINDEXADO, Left(XLO_CAMPOS.GetItemByName("Titulo").Value, 12)
      SubSomaMoeda2 XLI_MOEDA, XLT_MOEDA, VFV_TOTALMOEDADISTRATO, XLF_VALORDISTRATOINDEXADO, Left(XLO_CAMPOS.GetItemByName("Titulo").Value, 12)
    End If
  End If
  
  SetReportVariableValue "TOTALPREVISTOATUALIZADO", (FunNuloVal(GetReportVariableValue("TOTALPREVISTOATUALIZADO")) + XLF_VALORPREVISTOATUALIZADO)
  SetReportVariableValue "TOTALORIGINALATUALIZADO", (FunNuloVal(GetReportVariableValue("TOTALORIGINALATUALIZADO")) + XLF_VALORORIGINALATUALIZADO)
  SetReportVariableValue "TOTALLIQUIDADOATUALIZADO", (FunNuloVal(GetReportVariableValue("TOTALLIQUIDADOATUALIZADO")) + XLF_VALORPAGOATUALIZADO)
  SetReportVariableValue "TOTALDISTRATOATUALIZADO", (FunNuloVal(GetReportVariableValue("TOTALDISTRATOATUALIZADO")) + XLF_VALORDISTRATOATUALIZADO)
       
  '----------------------------- Campos Texto para pagos e não pagos ----------------
  
  TxtValorTitulo.SetText Format(XLF_VALORTITULO, "##,##0.0000")
  TxtMoeda.SetText XLT_MOEDA
  TxtValorPrevisto.SetText Format(XLF_VALORPREVISTO, "Standard")
  TxtValorOriginal.SetText Format(XLF_VALORORIGINAL, "Standard")
  TxtValorDistrato.SetText Format(XLF_VALORDISTRATO, "Standard")
  
  '------------------------------- Somatorio ----------------------------------------
  
  If (Mid(CStr(XLO_CAMPOS.GetItemByName("Titulo").Value), Len(CStr(XLO_CAMPOS.GetItemByName("Titulo").Value)) - 2, 2) = "00") Then
    SetReportVariableValue "SOMAVALORPREVISTO", (FunNuloVal(GetReportVariableValue("SOMAVALORPREVISTO")) + CDbl(Format(XLF_VALORPREVISTO, "Standard")))
  End If
  SetReportVariableValue "SOMAVALORORIGINAL", (FunNuloVal(GetReportVariableValue("SOMAVALORORIGINAL")) + CDbl(Format(XLF_VALORORIGINAL, "Standard")))
  SetReportVariableValue "SOMADISTRATO", (FunNuloVal(GetReportVariableValue("SOMADISTRATO")) + CDbl(Format(XLF_VALORDISTRATO, "Standard")))
    
  '------------------------------- Formatar Cor ----------------------------------------
  If XLF_VALORPREVISTO = 0 Or XLO_VETOR(0, 0) = -1 Then
    TxtAtencao.SetText "*"
    TxtAtencao.TextColor = vbRed
  
    ForTitulo.TextColor = vbRed
    FldDataVencimento.TextColor = vbRed
    FldDataPagamento.TextColor = vbRed
    FldDataBasePagamento.TextColor = vbRed
    FldSeguro.TextColor = vbRed
    FldMulta.TextColor = vbRed
    FldJuros.TextColor = vbRed
    FldDesconto.TextColor = vbRed
    FldDataPagamento.TextColor = vbRed
    FldValorPago.TextColor = vbRed
    
    TxtValorTitulo.TextColor = vbRed
    TxtMoeda.TextColor = vbRed
    TxtValorPrevisto.TextColor = vbRed
    TxtValorOriginal.TextColor = vbRed
    TxtValorDistrato.TextColor = vbRed
    
  End If
  
  Set XLO_VETOR = Nothing
  Set XLO_VETOR2 = Nothing
  XFT_TITULOANTERIOR = XLO_CAMPOS.GetItemByName("Titulo").Value
  
End Sub

Private Sub SecFooterCliente1_Format(ByVal pFormattingInfo As Object)
      
  Dim XFI_CONT As Integer
  
  'Enquanto não chegou no final do relatorio guarda o soma por moeda no vetor
  If XFB_IMPRIMIU = False Then
    VFV_VALORES(0, 0) = GetReportVariableValue("SOMAVALORPREVISTO")
    VFV_VALORES(0, 1) = GetReportVariableValue("SOMAVALORORIGINAL")
    VFV_VALORES(0, 2) = GetReportVariableValue("SOMADISTRATO")
    VFV_VALORES(0, 3) = GetReportVariableValue("TOTALPREVISTOATUALIZADO")
    VFV_VALORES(0, 4) = GetReportVariableValue("TOTALORIGINALATUALIZADO")
    VFV_VALORES(0, 5) = GetReportVariableValue("TOTALLIQUIDADOATUALIZADO")
    VFV_VALORES(0, 6) = GetReportVariableValue("TOTALDISTRATOATUALIZADO")
    
    SubSomatorioVetor VFV_SOMATORIO, VFV_VALORES, Left(XFT_TITULOANTERIOR, 12)
  
  'Depois que já passou pelo fim do reltório descarrega o vetor com os somatórios
  Else
    For XFI_CONT = 0 To VFV_SOMATORIO.Count(1) - 1
      If VFV_SOMATORIO(XFI_CONT, 0) = Left(XFT_TITULOANTERIOR, 12) Then
        While XFI_CONT < VFV_SOMATORIO.Count(1)
          If VFV_SOMATORIO(XFI_CONT, 0) <> Left(XFT_TITULOANTERIOR, 12) Then
            XFI_CONT = VFV_SOMATORIO.Count(1) + 1
          Else
            TxtSomaValorPrevisto.SetText Format(VFV_SOMATORIO(XFI_CONT, 1), "standard")
            TxtSomaValorOriginal.SetText Format(VFV_SOMATORIO(XFI_CONT, 2), "standard")
            TxtSomaDistrato.SetText FunTrataZero(Format(VFV_SOMATORIO(XFI_CONT, 3), "standard"))
            XFI_CONT = XFI_CONT + 1
          End If
        Wend
      End If
    Next
  End If
  
  SetReportVariableValue "SOMAVALORPREVISTO", 0
  SetReportVariableValue "SOMAVALORORIGINAL", 0
  SetReportVariableValue "SOMADISTRATO", 0
  SetReportVariableValue "TOTALPREVISTOATUALIZADO", 0
  SetReportVariableValue "TOTALORIGINALATUALIZADO", 0
  SetReportVariableValue "TOTALLIQUIDADOATUALIZADO", 0
  SetReportVariableValue "TOTALDISTRATOATUALIZADO", 0
  
End Sub

Private Sub SecFooterCliente2_Format(ByVal pFormattingInfo As Object)
  
  Dim XFI_CONT As Integer
  
  'Só preenche o vetor quando já passou pelo final do relalório
  If XFB_IMPRIMIU = True Then
    
    TxtTotalPrevisto.SetText ""
    TxtTotalOriginal.SetText ""
    TxtTotalLiquidado.SetText ""
    TxtTotalDistrato.SetText ""
    
    TxtMoedaPrevisto.SetText ""
    TxtMoedaOriginal.SetText ""
    TxtMoedaLiquidado.SetText ""
    TxtMoedaDistrato.SetText ""
    
    'Descarrega os vetores com os totais das moedas do valor previsto
    For XFI_CONT = 0 To VFV_TOTALMOEDAPREVISTO.Count(1) - 1
      If VFV_TOTALMOEDAPREVISTO(XFI_CONT, 3) = Left(XFT_TITULOANTERIOR, 12) Then
        While XFI_CONT < VFV_TOTALMOEDAPREVISTO.Count(1)
          If VFV_TOTALMOEDAPREVISTO(XFI_CONT, 3) <> Left(XFT_TITULOANTERIOR, 12) Then
            XFI_CONT = VFV_TOTALMOEDAPREVISTO.Count(1) + 1
          Else
            TxtTotalPrevisto.SetText TxtTotalPrevisto.Text & vbCrLf & Format(VFV_TOTALMOEDAPREVISTO(XFI_CONT, 2), "##,##0.0000")
            TxtMoedaPrevisto.SetText TxtMoedaPrevisto.Text & vbCrLf & VFV_TOTALMOEDAPREVISTO(XFI_CONT, 0)
            XFI_CONT = XFI_CONT + 1
          End If
        Wend
      End If
    Next
    
    'Descarrega os vetores com os totais das moedas do valor original
    For XFI_CONT = 0 To VFV_TOTALMOEDAORIGINAL.Count(1) - 1
      If VFV_TOTALMOEDAORIGINAL(XFI_CONT, 3) = Left(XFT_TITULOANTERIOR, 12) Then
        While XFI_CONT < VFV_TOTALMOEDAORIGINAL.Count(1)
          If VFV_TOTALMOEDAORIGINAL(XFI_CONT, 3) <> Left(XFT_TITULOANTERIOR, 12) Then
            XFI_CONT = VFV_TOTALMOEDAORIGINAL.Count(1) + 1
          Else
            TxtTotalOriginal.SetText TxtTotalOriginal.Text & vbCrLf & Format(VFV_TOTALMOEDAORIGINAL(XFI_CONT, 2), "##,##0.0000")
            TxtMoedaOriginal.SetText TxtMoedaOriginal.Text & vbCrLf & VFV_TOTALMOEDAORIGINAL(XFI_CONT, 0)
            XFI_CONT = XFI_CONT + 1
          End If
        Wend
      End If
    Next
      
    'Descarrega os vetores com os totais das moedas liquidadas
    For XFI_CONT = 0 To VFV_TOTALMOEDALIQUIDADO.Count(1) - 1
      If VFV_TOTALMOEDALIQUIDADO(XFI_CONT, 3) = Left(XFT_TITULOANTERIOR, 12) Then
        While XFI_CONT < VFV_TOTALMOEDALIQUIDADO.Count(1)
          If VFV_TOTALMOEDALIQUIDADO(XFI_CONT, 3) <> Left(XFT_TITULOANTERIOR, 12) Then
            XFI_CONT = VFV_TOTALMOEDALIQUIDADO.Count(1) + 1
          Else
            TxtTotalLiquidado.SetText TxtTotalLiquidado.Text & vbCrLf & Format(VFV_TOTALMOEDALIQUIDADO(XFI_CONT, 2), "##,##0.0000")
            TxtMoedaLiquidado.SetText TxtMoedaLiquidado.Text & vbCrLf & VFV_TOTALMOEDALIQUIDADO(XFI_CONT, 0)
            XFI_CONT = XFI_CONT + 1
          End If
        Wend
      End If
    Next
    
    'Descarrega os vetores com os totais das moedas liquidadas
    For XFI_CONT = 0 To VFV_TOTALMOEDADISTRATO.Count(1) - 1
      If VFV_TOTALMOEDADISTRATO(XFI_CONT, 3) = Left(XFT_TITULOANTERIOR, 12) Then
        While XFI_CONT < VFV_TOTALMOEDADISTRATO.Count(1)
          If VFV_TOTALMOEDADISTRATO(XFI_CONT, 3) <> Left(XFT_TITULOANTERIOR, 12) Then
            XFI_CONT = VFV_TOTALMOEDADISTRATO.Count(1) + 1
          Else
            TxtTotalDistrato.SetText TxtTotalDistrato.Text & vbCrLf & Format(VFV_TOTALMOEDADISTRATO(XFI_CONT, 2), "##,##0.0000")
            TxtMoedaDistrato.SetText TxtMoedaDistrato.Text & vbCrLf & VFV_TOTALMOEDADISTRATO(XFI_CONT, 0)
            XFI_CONT = XFI_CONT + 1
          End If
        Wend
      End If
    Next
  
  End If
  
End Sub

Private Sub SecFooterCliente3_Format(ByVal pFormattingInfo As Object)
    
  Dim XFI_CONT As Integer
  
  'Descarrega os vetores com os totais atualizados só qaundo já passou pelo final do relatório
  If XFB_IMPRIMIU = True Then
    For XFI_CONT = 0 To VFV_SOMATORIO.Count(1) - 1
      If VFV_SOMATORIO(XFI_CONT, 0) = Left(XFT_TITULOANTERIOR, 12) Then
        While XFI_CONT < VFV_SOMATORIO.Count(1)
          If VFV_SOMATORIO(XFI_CONT, 0) <> Left(XFT_TITULOANTERIOR, 12) Then
            XFI_CONT = VFV_SOMATORIO.Count(1) + 1
          Else
            TxtTotalPrevistoAualizado.SetText FunTrataZero(Format(FunNuloVal(VFV_SOMATORIO(XFI_CONT, 4)), "standard"))
            TxtTotalOriginalAtualizado.SetText FunTrataZero(Format(FunNuloVal(VFV_SOMATORIO(XFI_CONT, 5)), "standard"))
            TxtTotalLiquidadoAtualizado.SetText FunTrataZero(Format(FunNuloVal(VFV_SOMATORIO(XFI_CONT, 6)), "standard"))
            TxtTotalDistratoAtualizado.SetText FunTrataZero(Format(FunNuloVal(VFV_SOMATORIO(XFI_CONT, 7)), "standard"))
            XFI_CONT = XFI_CONT + 1
          End If
        Wend
      End If
    Next
  End If
  
End Sub

Private Sub SecFooterReport_Format(ByVal pFormattingInfo As Object)
  
  XFB_IMPRIMIU = True
  
End Sub


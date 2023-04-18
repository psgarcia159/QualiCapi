VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptExtratoContratoAnaliticoDestacadoCliente 
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15300
   OleObjectBlob   =   "RptExtratoContratoAnaliticoDestacadoCliente.dsx":0000
End
Attribute VB_Name = "RptExtratoContratoAnaliticoDestacadoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Dim VFV_SOMATORIO As New XArray ' Vetor que armazena os somatórios
Dim VFV_VALORES As New XArray  ' Vetor que armazena os valores que precisarão fazer somatório
Dim VFV_TOTALMOEDADEBITO As New XArray 'Vetor que armazena o total em debito por moeda
Dim VFV_TOTALMOEDALIQUIDADO As New XArray 'Vetor que armazena o total liquidado por moeda
Dim XFB_IMPRIMIU As Boolean 'Indica se chegou no final do relatório na primeira vez
Dim XFB_PRIMEIRAVEZ As Boolean

'Armazena o número do título para se comparado com o próximo número do título.
Dim XFT_TITULOANTERIOR As String

Private Sub Report_Initialize()

  AddReportVariable crRVNumber, "SOMAVALORPREVISTO"
  AddReportVariable crRVNumber, "SOMAVALORORIGINAL"
  AddReportVariable crRVNumber, "SOMAJUROS"
  AddReportVariable crRVNumber, "SOMAMULTA"
  AddReportVariable crRVNumber, "SOMADESAGIO"
  AddReportVariable crRVNumber, "SOMAVALORPAGO"
  AddReportVariable crRVNumber, "TOTALDEBITOATUALIZADO" 'Usado para somar o valor atualizado dos titulos em débito
  AddReportVariable crRVNumber, "TOTALLIQUIDADOATUALIZADO" 'Usado para somar o valor atualizado dos titulos liquidados
  
  XFB_IMPRIMIU = False
  
  VFV_TOTALMOEDADEBITO.ReDim 0, 0, 0, 3 '4 colunas para armazenar a chave do titulo o codigo da moeda e valor
  VFV_TOTALMOEDALIQUIDADO.ReDim 0, 0, 0, 3 '4 colunas para armazenar a chave do titulo o codigo da moeda e valor
  VFV_SOMATORIO.ReDim 0, 0, 0, 8  'colunas = quantidade de campos para soma
  VFV_VALORES.ReDim 0, 0, 0, 7 'colunas = quantidade de campos para soma -1
  
End Sub

Private Sub SectHeaderReport_Format(ByVal pFormattingInfo As Object)
  
  
  'Zera as variáveis quando passa pelo Header do relatório
  SetReportVariableValue "SOMAVALORPREVISTO", 0
  SetReportVariableValue "SOMAVALORORIGINAL", 0
  SetReportVariableValue "SOMAJUROS", 0
  SetReportVariableValue "SOMAMULTA", 0
  SetReportVariableValue "SOMADESAGIO", 0
  SetReportVariableValue "SOMAVALORPAGO", 0
  SetReportVariableValue "TOTALDEBITOATUALIZADO", 0
  SetReportVariableValue "TOTALLIQUIDADOATUALIZADO", 0
  
End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
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
  Dim XLF_VALORPAGOINDEXADO As Double 'Valor pago indexado
  Dim XLF_VALORPAGOATUALIZADO As Double 'Valor pago atualizado na correção da data base
  Dim XLI_MOEDA As Integer 'Código da moeda de acordo com o indexador escolhido
  Dim XLT_MOEDA As String 'Decrição da moeda de acordo com o indexador escolhido
  
  'Vetor auxiliar para armazenar o valor do titulo, a cotação e a cotação do mês anterior,
  'tudo de acordo com o indexador escolhido. Também indica se tem indexador ou não.
  Dim XLO_VETOR As New XArray
     
  Dim XLD_DATABASE As Date  'Data base passada por parâmetro para relatório
  Dim XLI_INDEXADOR As Byte  ' Indexador passado por parâmetro para o relatório
  Dim XLT_FILTROMULTA As String
  Dim XLT_FILTROJUROS As String
  Dim XLT_FILTRODESAGIO As String
  Dim XLT_FILTROTITULOSPAGOS As String
  Dim XLO_CAMPOS As CRAXDDRT.DatabaseFieldDefinitions  ' Objeto que representa os campos do relatório
  
  Set XLO_CAMPOS = Me.Database.Tables(1).Fields
  
  XLI_INDEXADOR = CByte(FormulaFields.GetItemByName("Indexador").Value)
   
  SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  XLT_MOEDA = XLO_VETOR(0, 0)
  XLI_MOEDA = XLO_VETOR(1, 0)
  
  'Pega as cotações do mês da data base e do mês anterior
  SubCotacao XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  
  'se não foi pago
  If XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value = "" Then
    
    XLD_DATABASE = CDate(FormulaFields.GetItemByName("FiltroDataBase").Value)
    
    XLF_VALORCORRIGIDO = Format(FunCalculoCorrecaoMonetaria(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLI_MOEDA), "Standard")
       
    'Juros da CEF
    XLF_VALORCORRIGIDO = XLF_VALORCORRIGIDO + Format(FunCalculoJurosCEF(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
    
    If FormulaFields.GetItemByName("FiltroJuros").Value = "1" Then
      XLF_JUROS = Format(FunCalculoJuros(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
    End If
    If FormulaFields.GetItemByName("FiltroMulta").Value = "1" Then
      XLF_MULTA = Format(FunCalculoMulta(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
    End If
    If FormulaFields.GetItemByName("FiltroDesagio").Value = "1" Then
      XLF_DESAGIO = Format(FunCalculoDesagio(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
    End If
        
    XLF_VALORPREVISTO = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA - XLF_DESAGIO
    XLF_VALORORIGINAL = XLF_VALORCORRIGIDO
    
    If XGT_INCLUIRACRESCIMOS = "S" Then
      XLF_VALORTITULO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORPREVISTO), "##,##0.0000")
    Else
      XLF_VALORTITULO = Format(XLO_VETOR(0, 0), "##,##0.0000")
    End If
          
    ' Se for o mesmo titulo que o anterior não soma no vetor
    If XFT_TITULOANTERIOR <> XLO_CAMPOS.GetItemByName("Titulo").Value And XFB_IMPRIMIU = False Then
      SubSomaMoeda2 XLI_MOEDA, XLT_MOEDA, VFV_TOTALMOEDADEBITO, XLF_VALORTITULO, Left(XLO_CAMPOS.GetItemByName("Titulo").Value, 12)
    End If
    SetReportVariableValue "TOTALDEBITOATUALIZADO", (FunNuloVal(GetReportVariableValue("TOTALDEBITOATUALIZADO")) + XLF_VALORPREVISTO)
       
  'se foi pago
  Else
    XLD_DATABASE = XLO_CAMPOS.GetItemByName("titu_dt_BasePagto").Value
    
    XLF_VALORCORRIGIDO = Format(FunCalculoCorrecaoMonetaria(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLI_MOEDA), "Standard")
    
    'Juros da CEF
    XLF_VALORCORRIGIDO = XLF_VALORCORRIGIDO + Format(FunCalculoJurosCEF(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
    
    
    If FormulaFields.GetItemByName("FiltroTitulosPagos").Value = "1" Then
      XLF_JUROS = Format(FunCalculoJuros(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
      XLF_MULTA = Format(FunCalculoMulta(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
    End If
    
    XLF_DESAGIO = Format(FunCalculoDesagio(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
    
    XLF_VALORPREVISTO = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA - XLF_DESAGIO
    
    If XGT_INCLUIRACRESCIMOS = "S" Then 'Variável global indicada no campo da empresa
      XLF_VALORTITULO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORPREVISTO), "##,##0.0000")
    Else
      XLF_VALORTITULO = Format(XLO_VETOR(0, 0), "##,##0.0000")
    End If
    
    XLF_VALORORIGINAL = Format(FunCalculoValorOriginalPago(XLO_CAMPOS), "standard")
    'XLF_VALORPAGO = Format(FunCalculoValorPago(XLO_CAMPOS, FormulaFields.GetItemByName("FiltroTitulosPagos").Value), "standard")
    
    'Solicitado por Albetino para permitir que o valor pago mostrado inclua multas e juros - Em 03/2007
    XLF_VALORPAGO = Format(FunCalculoValorPago(XLO_CAMPOS, 1), "standard")
    
    
    
    'Multa, juros e deságio pagos
    If FormulaFields.GetItemByName("FiltroTitulosPagos").Value = "1" Then
      XLF_JUROS = Format(XLO_CAMPOS.GetItemByName("titu_vl_Juros").Value, "standard")
      XLF_MULTA = Format(XLO_CAMPOS.GetItemByName("titu_vl_Multa").Value, "standard")
    End If
    XLF_DESAGIO = Format(XLO_CAMPOS.GetItemByName("titu_vl_Desagio").Value, "standard")
    
   If FormulaFields.GetItemByName("IndexadorExtra").Value <> "" Then
      XLO_VETOR(0, 0) = 0
      XLO_VETOR(1, 0) = CDbl(XLO_CAMPOS.GetItemByName("Cotacao5").Value)
      XLO_VETOR(2, 0) = CDbl(XLO_CAMPOS.GetItemByName("Cotacao5_MesAnterior").Value)
      XLO_VETOR(3, 0) = FormulaFields.GetItemByName("IndexacaoExtra").Value
      XLO_VETOR(4, 0) = 0
    Else
      SubCotacao3 XLO_CAMPOS, XLO_VETOR
    End If
    
    If XGT_INCLUIRACRESCIMOS = "S" Then
      XLF_VALORPAGOINDEXADO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORPAGO), "##,##0.0000")
    Else
      XLF_VALORPAGOINDEXADO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, Format(FunCalculoValorPago(XLO_CAMPOS, "0"), "standard")), "##,##0.0000")
    End If
    
   If FormulaFields.GetItemByName("IndexadorExtra").Value <> "" Then
      XLO_VETOR(0, 0) = 0
      XLO_VETOR(1, 0) = CDbl(XLO_CAMPOS.GetItemByName("Cotacao4").Value)
      XLO_VETOR(2, 0) = CDbl(XLO_CAMPOS.GetItemByName("Cotacao4_MesAnterior").Value)
      XLO_VETOR(3, 0) = FormulaFields.GetItemByName("IndexacaoExtra").Value
      XLO_VETOR(4, 0) = 0
    Else
      SubCotacao2 XLO_CAMPOS, XLO_VETOR
    End If
    XLF_VALORPAGOATUALIZADO = Format(FunCalculoValorAtualizado(XLO_CAMPOS, XLO_VETOR, CDate(FormulaFields.GetItemByName("FiltroDataBase").Value), XLF_VALORPAGOINDEXADO, XLI_MOEDA), "standard")
       
    ' Se for o mesmo titulo que o anterior não soma no vetor
    If XFT_TITULOANTERIOR <> XLO_CAMPOS.GetItemByName("Titulo").Value And XFB_IMPRIMIU = False Then
      If FormulaFields.GetItemByName("IndexadorExtra").Value <> "" Then
        SubSomaMoeda2 FormulaFields.GetItemByName("IndexadorExtra").Value, FormulaFields.GetItemByName("NomeIndexadorExtra").Value, VFV_TOTALMOEDALIQUIDADO, XLF_VALORPAGOINDEXADO, Left(XLO_CAMPOS.GetItemByName("Titulo").Value, 12)
      Else
        SubSomaMoeda2 XLI_MOEDA, XLT_MOEDA, VFV_TOTALMOEDALIQUIDADO, XLF_VALORPAGOINDEXADO, Left(XLO_CAMPOS.GetItemByName("Titulo").Value, 12)
      End If
    End If
    
    SetReportVariableValue "TOTALLIQUIDADOATUALIZADO", (FunNuloVal(GetReportVariableValue("TOTALLIQUIDADOATUALIZADO")) + XLF_VALORPAGOATUALIZADO)
    
    TxtValorPago.SetText Format(XLF_VALORPAGO, "Standard")
    TxtValorPagoIndexado.SetText Format(XLF_VALORPAGOINDEXADO, "##,##0.0000")
      
  End If
  
  '----------------------------- Campos Texto para pagos e não pagos ----------------
    
  ' -- TxtMoeda.SetText XLT_MOEDA
  TxtValorTitulo.SetText Format(XLF_VALORTITULO, "##,##0.0000")
  TxtValorPrevisto.SetText Format(XLF_VALORPREVISTO, "Standard")
  TxtValorOriginal.SetText Format(XLF_VALORORIGINAL, "Standard")
  TxtJuros.SetText FunTrataZero(Format(XLF_JUROS, "Standard"))
  TxtMulta.SetText FunTrataZero(Format(XLF_MULTA, "Standard"))
  TxtDesagio.SetText FunTrataZero(Format(XLF_DESAGIO, "standard"))
  TxtObservacao.SetText FunObservacao(XLO_CAMPOS, XLD_DATABASE)
  
  '------------------------------- Somatorio ----------------------------------------
  
  If (Mid(CStr(XLO_CAMPOS.GetItemByName("Titulo").Value), Len(XLO_CAMPOS.GetItemByName("Titulo").Value) - 1, 2) = "00") Then
      SetReportVariableValue "SOMAVALORPREVISTO", (FunNuloVal(GetReportVariableValue("SOMAVALORPREVISTO")) + CDbl(Format(XLF_VALORPREVISTO, "Standard")))
  End If
  
  SetReportVariableValue "SOMAVALORORIGINAL", (FunNuloVal(GetReportVariableValue("SOMAVALORORIGINAL")) + CDbl(Format(XLF_VALORORIGINAL, "Standard")))
  SetReportVariableValue "SOMAJUROS", (FunNuloVal(GetReportVariableValue("SOMAJUROS")) + CDbl(Format(XLF_JUROS, "Standard")))
  SetReportVariableValue "SOMAMULTA", (FunNuloVal(GetReportVariableValue("SOMAMULTA")) + CDbl(Format(XLF_MULTA, "Standard")))
  SetReportVariableValue "SOMADESAGIO", (FunNuloVal(GetReportVariableValue("SOMADESAGIO")) + CDbl(Format(XLF_DESAGIO, "Standard")))
  SetReportVariableValue "SOMAVALORPAGO", (FunNuloVal(GetReportVariableValue("SOMAVALORPAGO")) + CDbl(Format(XLF_VALORPAGO, "Standard")))
    
  '------------------------------- Formatar Cor ----------------------------------------
  If XLF_VALORPREVISTO = 0 Or XLO_VETOR(0, 0) = -1 Then
    
    TxtAtencao.SetText "*"
    TxtAtencao.TextColor = vbRed
    
    ForTitulo.TextColor = vbRed
    FldDataPagamento.TextColor = vbRed
    FldDataVencimento.TextColor = vbRed
    TxtValorTitulo.TextColor = vbRed
    ' -- TxtMoeda.TextColor = vbRed
    TxtValorPrevisto.TextColor = vbRed
    TxtValorOriginal.TextColor = vbRed
    TxtJuros.TextColor = vbRed
    TxtMulta.TextColor = vbRed
    TxtDesagio.TextColor = vbRed
    TxtObservacao.TextColor = vbRed
    
    TxtValorPago.TextColor = vbRed
    TxtValorPagoIndexado.TextColor = vbRed
    FldDataPagamento.TextColor = vbRed
    ForMoeda.TextColor = vbRed
    
  End If
  
  Set XLO_VETOR = Nothing
  XFT_TITULOANTERIOR = XLO_CAMPOS.GetItemByName("Titulo").Value
  
End Sub

Private Sub SecFooterCliente1_Format(ByVal pFormattingInfo As Object)
      
  Dim XFI_CONT As Integer
  
  'Enquanto não chegou no final do relatorio guarda o soma por moeda no vetor
  If XFB_IMPRIMIU = False Then
    VFV_VALORES(0, 0) = GetReportVariableValue("SOMAVALORPREVISTO")
    VFV_VALORES(0, 1) = GetReportVariableValue("SOMAVALORORIGINAL")
    VFV_VALORES(0, 2) = GetReportVariableValue("SOMAJUROS")
    VFV_VALORES(0, 3) = GetReportVariableValue("SOMAMULTA")
    VFV_VALORES(0, 4) = GetReportVariableValue("SOMADESAGIO")
    VFV_VALORES(0, 5) = GetReportVariableValue("SOMAVALORPAGO")
    VFV_VALORES(0, 6) = GetReportVariableValue("TOTALDEBITOATUALIZADO")
    VFV_VALORES(0, 7) = GetReportVariableValue("TOTALLIQUIDADOATUALIZADO")
    
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
            TxtSomaJuros.SetText FunTrataZero(Format(VFV_SOMATORIO(XFI_CONT, 3), "standard"))
            TxtSomaMulta.SetText FunTrataZero(Format(VFV_SOMATORIO(XFI_CONT, 4), "standard"))
            TxtSomaDesagio.SetText FunTrataZero(Format(VFV_SOMATORIO(XFI_CONT, 5), "standard"))
            TxtSomaValorPago.SetText FunTrataZero(Format(VFV_SOMATORIO(XFI_CONT, 6), "standard"))
            XFI_CONT = XFI_CONT + 1
          End If
        Wend
      End If
    Next
  End If
  
  'quando passa pelo footer zera as variáveis
  SetReportVariableValue "SOMAVALORPREVISTO", 0
  SetReportVariableValue "SOMAVALORORIGINAL", 0
  SetReportVariableValue "SOMAJUROS", 0
  SetReportVariableValue "SOMAMULTA", 0
  SetReportVariableValue "SOMADESAGIO", 0
  SetReportVariableValue "SOMAVALORPAGO", 0
  SetReportVariableValue "TOTALDEBITOATUALIZADO", 0
  SetReportVariableValue "TOTALLIQUIDADOATUALIZADO", 0
  
End Sub

Private Sub SecFooterCliente2_Format(ByVal pFormattingInfo As Object)
  
  Dim XFI_CONT As Integer
  
  'Só preenche o vetor quando já passou pelo final do relalório
  If XFB_IMPRIMIU = True Then
    
    TxtTotalDebito.SetText ""
    ' -- TxtMoedaDebito.SetText ""
    TxtTotalLiquidado.SetText ""
    ' -- TxtMoedaLiquidado.SetText ""
    
    'Descarrega os vetores com os totais das moedas em debito
    For XFI_CONT = 0 To VFV_TOTALMOEDADEBITO.Count(1) - 1
      If VFV_TOTALMOEDADEBITO(XFI_CONT, 3) = Left(XFT_TITULOANTERIOR, 12) Then
        While XFI_CONT < VFV_TOTALMOEDADEBITO.Count(1)
          If VFV_TOTALMOEDADEBITO(XFI_CONT, 3) <> Left(XFT_TITULOANTERIOR, 12) Then
            XFI_CONT = VFV_TOTALMOEDADEBITO.Count(1) + 1
          Else
            TxtTotalDebito.SetText TxtTotalDebito.Text & vbCrLf & Format(VFV_TOTALMOEDADEBITO(XFI_CONT, 2), "##,##0.0000")
            ' -- TxtMoedaDebito.SetText TxtMoedaDebito.Text & vbCrLf & VFV_TOTALMOEDADEBITO(XFI_CONT, 0)
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
            ' -- TxtMoedaLiquidado.SetText TxtMoedaLiquidado.Text & vbCrLf & VFV_TOTALMOEDALIQUIDADO(XFI_CONT, 0)
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
            TxtTotalDebitoAualizado.SetText FunTrataZero(Format(FunNuloVal(VFV_SOMATORIO(XFI_CONT, 7)), "standard"))
            TxtTotalLiquidadoAtualizado.SetText FunTrataZero(Format(FunNuloVal(VFV_SOMATORIO(XFI_CONT, 8)), "standard"))
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


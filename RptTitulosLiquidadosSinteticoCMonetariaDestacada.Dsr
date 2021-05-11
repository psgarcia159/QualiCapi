VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptTitulosLiquidadosSinteticoCMonetariaDestacada 
   ClientHeight    =   13665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19140
   OleObjectBlob   =   "RptTitulosLiquidadosSinteticoCMonetariaDestacada.dsx":0000
End
Attribute VB_Name = "RptTitulosLiquidadosSinteticoCMonetariaDestacada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Private Sub Report_Initialize()

AddReportVariable crRVNumber, "SOMAVALORPREVISTO"
AddReportVariable crRVNumber, "SOMAVALORORIGINAL"
AddReportVariable crRVNumber, "SOMAVALORMULTA"
AddReportVariable crRVNumber, "SOMAVALORJUROS"
AddReportVariable crRVNumber, "SOMAVALORDESAGIO"
AddReportVariable crRVNumber, "SOMAVALOROUTROS"
AddReportVariable crRVNumber, "SOMAVALORDESCONTO"
AddReportVariable crRVNumber, "SOMAVALORSEGURO"
AddReportVariable crRVNumber, "SOMAVALORPAGO"

AddReportVariable crRVNumber, "TOTALVALORPREVISTO"
AddReportVariable crRVNumber, "TOTALVALORORIGINAL"
AddReportVariable crRVNumber, "TOTALVALORMULTA"
AddReportVariable crRVNumber, "TOTALVALORJUROS"
AddReportVariable crRVNumber, "TOTALVALORDESAGIO"
AddReportVariable crRVNumber, "TOTALVALOROUTROS"
AddReportVariable crRVNumber, "TOTALVALORDESCONTO"
AddReportVariable crRVNumber, "TOTALVALORSEGURO"
AddReportVariable crRVNumber, "TOTALVALORPAGO"



End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SecDetalhe2_Format(ByVal pFormattingInfo As Object)
  
  Dim XLF_VALORCORRIGIDO As Double  'Valor corrigido monetariamente
  Dim XLF_VALORHISTORICO As Double 'Valor histórico do título
  Dim XLF_VALORCORRECAOMONETARIA As Double 'Valor da Correção Monetária (separado do valor do título)
  Dim XLF_DESAGIO As Double  'Valor do deságio
  Dim XLF_MULTA As Double  'Valor da Multa
  Dim XLF_JUROS As Double  'Valor do Juros
  
  Dim XLF_OUTROS As Double  'Valor dos Outros descontos - 09/11/09
  Dim XLF_DESCONTO As Double  'Valor do desconto - 09/11/09
  Dim XLF_SEGURO As Double  'Valor do seguro - 09/11/09
  
  Dim XLF_VALORPREVISTO As Double 'Valor Previsto
  Dim XLF_VALORORIGINAL As Double 'Valor Original
  Dim XLI_MOEDA As Integer 'Código da moeda de acordo com o indexador escolhido
  
  'Vetor auxiliar para armazenar o valor do titulo, a cotação e a cotação do mês anterior,
  'tudo de acordo com o indexador escolhido. Também indica se tem indexador ou não.
  Dim XLO_VETOR As New XArray
     
  Dim XLD_DATABASE As Date  'Data base passada por parâmetro para relatório
  Dim XLI_INDEXADOR As Byte  ' Indexador passado por parâmetro para o relatório
  
  Dim XLO_CAMPOS As CRAXDDRT.DatabaseFieldDefinitions  ' Objeto que representa os campos do relatório
  
  'se passou pelo footer zera as variáveis
  Set XLO_CAMPOS = Me.Database.Tables(1).Fields
  
  'XLI_INDEXADOR = 1 'Só para compor os parâmetros das fórmulas
  
  XLI_INDEXADOR = CByte(FormulaFields.GetItemByName("Indexador").Value)
  
  'Pega o código da moeda e a sua descrição
  If XLO_CAMPOS.GetItemByName("titu_vl_ValorPago").Value = "" Then
    SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, True
  Else
    SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  End If
  
  XLI_MOEDA = XLO_VETOR(1, 0)

  
  'Pega as cotações do mês da data base e do mês anterior
  SubCotacao XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  
  XLD_DATABASE = XLO_CAMPOS.GetItemByName("titu_dt_BasePagto").Value
  
  XLF_VALORCORRIGIDO = Format(FunCalculoCorrecaoMonetaria(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLI_MOEDA), "Standard")
  XLF_VALORHISTORICO = XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value
  
  XLF_JUROS = XLO_CAMPOS.GetItemByName("titu_vl_Juros").Value 'Format(FunCalculoJuros(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  XLF_MULTA = Format(FunCalculoMulta(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
 
    
    
  XLF_DESAGIO = Format(FunCalculoDesagio(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  XLF_MULTA = XLO_CAMPOS.GetItemByName("titu_vl_Multa").Value  ' Format(FunCalculoMulta(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  XLF_OUTROS = XLO_CAMPOS.GetItemByName("titu_vl_Outros").Value
  XLF_DESCONTO = XLO_CAMPOS.GetItemByName("titu_vl_Desconto").Value
  XLF_SEGURO = XLO_CAMPOS.GetItemByName("titu_vl_Seguro").Value
  'Alterar o nome para algo mais claro
  'XLF_VALORPREVISTO = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA - XLF_DESAGIO
  XLF_VALORPREVISTO = Format(FunCalculoValorOriginalPago(XLO_CAMPOS), "standard")
  'XLF_VALORPREVISTO = Format(FunCalculoValorOriginalPago(XLO_CAMPOS), "standard")
  
'  XLF_VALORPREVISTO = XLF_VALORHISTORICO
  
  XLF_VALORCORRECAOMONETARIA = XLF_VALORCORRIGIDO - XLF_VALORHISTORICO - XLF_JUROS - XLF_MULTA + XLF_DESAGIO
  '------------------------------- Somatorio ----------------------------------------
    
  SetReportVariableValue "SOMAVALORPREVISTO", (FunNuloVal(GetReportVariableValue("SOMAVALORPREVISTO")) + CDbl(Format(XLF_VALORPREVISTO, "Standard")))
  SetReportVariableValue "SOMAVALORORIGINAL", (FunNuloVal(GetReportVariableValue("SOMAVALORORIGINAL")) + CDbl(Format(XLF_VALORCORRECAOMONETARIA, "Standard")))
  
  SetReportVariableValue "SOMAVALORSEGURO", (FunNuloVal(GetReportVariableValue("SOMAVALORSEGURO")) + CDbl(Format(XLF_SEGURO, "Standard")))
  SetReportVariableValue "SOMAVALORMULTA", (FunNuloVal(GetReportVariableValue("SOMAVALORMULTA")) + CDbl(Format(XLF_MULTA, "Standard")))
  SetReportVariableValue "SOMAVALORJUROS", (FunNuloVal(GetReportVariableValue("SOMAVALORJUROS")) + CDbl(Format(XLF_JUROS, "Standard")))
  SetReportVariableValue "SOMAVALOROUTROS", (FunNuloVal(GetReportVariableValue("SOMAVALOROUTROS")) + CDbl(Format(XLF_OUTROS, "Standard")))
  SetReportVariableValue "SOMAVALORDESAGIO", (FunNuloVal(GetReportVariableValue("SOMAVALORDESAGIO")) + CDbl(Format(XLF_DESAGIO, "Standard")))
  SetReportVariableValue "SOMAVALORDESCONTO", (FunNuloVal(GetReportVariableValue("SOMAVALORDESCONTO")) + CDbl(Format(XLF_DESCONTO, "Standard")))


  '23/11/09
  SetReportVariableValue "SOMAVALORPAGO", (FunNuloVal(GetReportVariableValue("SOMAVALORPAGO")) + CDbl(Format(XLF_VALORPREVISTO, "Standard")) + _
  CDbl(Format(XLF_VALORORIGINAL, "Standard")) + CDbl(Format(XLF_SEGURO, "Standard")) + CDbl(Format(XLF_MULTA, "Standard")) + _
  CDbl(Format(XLF_JUROS, "Standard")) + CDbl(Format(XLF_OUTROS, "Standard")) - CDbl(Format(XLF_DESAGIO, "Standard")) - _
  CDbl(Format(XLF_DESCONTO, "Standard")))
     
  Set XLO_VETOR = Nothing
End Sub

Private Sub SecFooterEmpreendimento1_Format(ByVal pFormattingInfo As Object)

  'Anterior
  TxtSomaValorPrevisto.SetText Format(GetReportVariableValue("SOMAVALORPREVISTO"), "standard")
  TxtSomaValorOriginal.SetText Format(GetReportVariableValue("SOMAVALORORIGINAL"), "standard")
  TxtSomaValorSeguro.SetText Format(GetReportVariableValue("SOMAVALORSEGURO"), "standard")
  TxtSomaValorMulta.SetText Format(GetReportVariableValue("SOMAVALORMULTA"), "standard")
  TxtSomaValorJuros.SetText Format(GetReportVariableValue("SOMAVALORJUROS"), "standard")
  TxtSomaValorOutros.SetText Format(GetReportVariableValue("SOMAVALOROUTROS"), "standard")
  TxtSomaValorDesagio.SetText Format(GetReportVariableValue("SOMAVALORDESAGIO"), "standard")
  TxtSomaValorDesconto.SetText Format(GetReportVariableValue("SOMAVALORDESCONTO"), "standard")
  TxtSomaValorPago.SetText Format(GetReportVariableValue("SOMAVALORPAGO"), "standard")
  
  SetReportVariableValue "TOTALVALORPREVISTO", (FunNuloVal(GetReportVariableValue("TOTALVALORPREVISTO")) + GetReportVariableValue("SOMAVALORPREVISTO"))
  SetReportVariableValue "TOTALVALORORIGINAL", (FunNuloVal(GetReportVariableValue("TOTALVALORORIGINAL")) + GetReportVariableValue("SOMAVALORORIGINAL"))
  SetReportVariableValue "TOTALVALORSEGURO", (FunNuloVal(GetReportVariableValue("TOTALVALORSEGURO")) + GetReportVariableValue("SOMAVALORSEGURO"))
  SetReportVariableValue "TOTALVALORMULTA", (FunNuloVal(GetReportVariableValue("TOTALVALORMULTA")) + GetReportVariableValue("SOMAVALORMULTA"))
  SetReportVariableValue "TOTALVALORJUROS", (FunNuloVal(GetReportVariableValue("TOTALVALORJUROS")) + GetReportVariableValue("SOMAVALORJUROS"))
  SetReportVariableValue "TOTALVALOROUTROS", (FunNuloVal(GetReportVariableValue("TOTALVALOROUTROS")) + GetReportVariableValue("SOMAVALOROUTROS"))
  SetReportVariableValue "TOTALVALORDESAGIO", (FunNuloVal(GetReportVariableValue("TOTALVALORDESAGIO")) + GetReportVariableValue("SOMAVALORDESAGIO"))
  SetReportVariableValue "TOTALVALORDESCONTO", (FunNuloVal(GetReportVariableValue("TOTALVALORDESCONTO")) + GetReportVariableValue("SOMAVALORDESCONTO"))
  
  SetReportVariableValue "TOTALVALORPAGO", (FunNuloVal(GetReportVariableValue("TOTALVALORPAGO")) + GetReportVariableValue("SOMAVALORPAGO"))
    
  '------------------------------- Formatar Cor ----------------------------------------
  If GetReportVariableValue("SOMAVALORPREVISTO") = 0 Then
    TxtSomaValorPrevisto.TextColor = vbRed
    TxtSomaValorOriginal.TextColor = vbRed
  End If
  
  SetReportVariableValue "SOMAVALORPREVISTO", 0
  SetReportVariableValue "SOMAVALORORIGINAL", 0
  SetReportVariableValue "SOMAVALORSEGURO", 0
  SetReportVariableValue "SOMAVALORMULTA", 0
  SetReportVariableValue "SOMAVALORJUROS", 0
  SetReportVariableValue "SOMAVALOROUTROS", 0
  SetReportVariableValue "SOMAVALORDESAGIO", 0
  SetReportVariableValue "SOMAVALORDESCONTO", 0
  SetReportVariableValue "SOMAVALORDESCONTO", 0
  SetReportVariableValue "SOMAVALORPAGO", 0



End Sub

Private Sub SecFooterReport_Format(ByVal pFormattingInfo As Object)
  
  TxtTotalValorPrevisto.SetText Format(GetReportVariableValue("TOTALVALORPREVISTO"), "standard")
  TxtTotalValorOriginal.SetText Format(GetReportVariableValue("TOTALVALORORIGINAL"), "standard")
  
  TxtTotalValorSeguro.SetText Format(GetReportVariableValue("TOTALVALORSEGURO"), "standard")
  TxtTotalValorMulta.SetText Format(GetReportVariableValue("TOTALVALORMULTA"), "standard")
  TxtTotalValorJuros.SetText Format(GetReportVariableValue("TOTALVALORJUROS"), "standard")
  TxtTotalValorOutros.SetText Format(GetReportVariableValue("TOTALVALOROUTROS"), "standard")
  TxtTotalValorDesagio.SetText Format(GetReportVariableValue("TOTALVALORDESAGIO"), "standard")
  TxtTotalValorOrDesconto.SetText Format(GetReportVariableValue("TOTALVALORDESCONTO"), "standard")
  TxtTotalGeralPago.SetText Format(GetReportVariableValue("TOTALVALORPAGO"), "standard")
      
End Sub

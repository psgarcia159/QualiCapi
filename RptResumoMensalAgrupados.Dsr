VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptResumoMensalAgrupados 
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11460
   OleObjectBlob   =   "RptResumoMensalAgrupados.dsx":0000
End
Attribute VB_Name = "RptResumoMensalAgrupados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Dim XFB_FOOTER1 As Boolean 'indica se passou pelo footer
Dim XFB_FOOTER2 As Boolean 'indica se passou pelo footer
Dim XFT_TITULOANTERIOR As String

Private Sub Report_Initialize()

AddReportVariable crRVNumber, "SOMAVENCIDO"
AddReportVariable crRVNumber, "SOMAVENCERCOMAGIO"
AddReportVariable crRVNumber, "SOMAVENCERSEMAGIO"
AddReportVariable crRVNumber, "SOMATOTAL12"
AddReportVariable crRVNumber, "SOMATOTAL13"
AddReportVariable crRVNumber, "SOMALIQUIDADOHISTORICO"
AddReportVariable crRVNumber, "SOMALIQUIDADOATUALIZADO"

AddReportVariable crRVNumber, "TOTALVENCIDO"
AddReportVariable crRVNumber, "TOTALVENCERCOMAGIO"
AddReportVariable crRVNumber, "TOTALVENCERSEMAGIO"
AddReportVariable crRVNumber, "TOTALTOTAL12"
AddReportVariable crRVNumber, "TOTALTOTAL13"
AddReportVariable crRVNumber, "TOTALLIQUIDADOHISTORICO"
AddReportVariable crRVNumber, "TOTALLIQUIDADOATUALIZADO"

XFB_FOOTER1 = False
XFB_FOOTER2 = False

End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SectHeaderReport_Format(ByVal pFormattingInfo As Object)
  
  SetReportVariableValue "SOMAVENCIDO", 0
  SetReportVariableValue "SOMAVENCERCOMAGIO", 0
  SetReportVariableValue "SOMAVENCERSEMAGIO", 0
  SetReportVariableValue "SOMATOTAL12", 0
  SetReportVariableValue "SOMATOTAL13", 0
  SetReportVariableValue "SOMALIQUIDADOHISTORICO", 0
  SetReportVariableValue "SOMALIQUIDADOATUALIZADO", 0
  
  SetReportVariableValue "TOTALVENCIDO", 0
  SetReportVariableValue "TOTALVENCERCOMAGIO", 0
  SetReportVariableValue "TOTALVENCERSEMAGIO", 0
  SetReportVariableValue "TOTALTOTAL12", 0
  SetReportVariableValue "TOTALTOTAL13", 0
  SetReportVariableValue "TOTALLIQUIDADOHISTORICO", 0
  SetReportVariableValue "TOTALLIQUIDADOATUALIZADO", 0
  
End Sub

Private Sub SecDetalhe2_Format(ByVal pFormattingInfo As Object)
  
  Dim XLF_VALORCORRIGIDO As Double  'Valor corrigido monetariamente
  Dim XLF_DESAGIO As Double  'Valor do deságio
  Dim XLF_MULTA As Double  'Valor da Multa
  Dim XLF_JUROS As Double  'Valor do Juros
    
  Dim XLF_VENCIDO As Double
  Dim XLF_VENCERCOMAGIO As Double
  Dim XLF_VENCERSEMAGIO As Double
  Dim XLF_TOTAL12 As Double
  Dim XLF_TOTAL13 As Double
  Dim XLF_VALORLIQUIDADOHISTORICO As Double
  Dim XLF_VALORLIQUIDADOINDEXADO As Double
  Dim XLF_VALORLIQUIDADOATUALIZADO As Double
    
  'Vetor auxiliar para armazenar o valor do titulo, a cotação e a cotação do mês anterior,
  'tudo de acordo com o indexador escolhido. Também indica se tem indexador ou não.
  Dim XLO_VETOR As New XArray
     
  Dim XLD_DATABASE As Date  'Data base passada por parâmetro para relatório
  Dim XLI_INDEXADOR As Byte  ' Indexador passado por parâmetro para o relatório
  Dim XLT_FILTROMULTA As String
  Dim XLT_FILTROJUROS As String
  Dim XLT_FILTRODESAGIO As String
  Dim XLT_FILTROTITULOSPAGOS As String
  Dim XLI_MOEDA As Integer 'Código da moeda de acordo com o indexador escolhido
  
  Dim XLO_CAMPOS As CRAXDDRT.DatabaseFieldDefinitions  ' Objeto que representa os campos do relatório
     
  Set XLO_CAMPOS = Me.Database.Tables(1).Fields
  
  XLI_INDEXADOR = CByte(FormulaFields.GetItemByName("Indexador").Value)
    
  'Pega o código da moeda e a sua descrição
  'If XLO_CAMPOS.GetItemByName("Pago").Value = "N" Then
    SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, True
  'Else
  '  SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  'End If
  
  XLI_MOEDA = XLO_VETOR(1, 0)
    
  'Pega as cotações do mês da data base e do mês anterior
  If FormulaFields.GetItemByName("ConsiderarTitulosPagos").Value = "1" Then
    'Se escolher considerar titulos pagos, pega a cotação como se não estivesse pago
    SubCotacao XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, True
  Else
    SubCotacao XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  End If
  
  'se não foi pago
  If (XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value = "") Or _
    (XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value <> "" And _
    XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value >= XLD_DATABASE And _
    FormulaFields.GetItemByName("ConsiderarTitulosPagos").Value = "1") Then
    
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
    XLF_DESAGIO = Format(FunCalculoDesagio(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
        
    If XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value < XLD_DATABASE Then
      XLF_VENCIDO = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA
    Else
      XLF_VENCERCOMAGIO = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA
      XLF_VENCERSEMAGIO = XLF_VENCERCOMAGIO - XLF_DESAGIO
    End If
    
    '------------------------------- Somatorio ----------------------------------------
  
    SetReportVariableValue "SOMAVENCIDO", (FunNuloVal(GetReportVariableValue("SOMAVENCIDO")) + XLF_VENCIDO)
    SetReportVariableValue "SOMAVENCERCOMAGIO", (FunNuloVal(GetReportVariableValue("SOMAVENCERCOMAGIO")) + XLF_VENCERCOMAGIO)
    SetReportVariableValue "SOMAVENCERSEMAGIO", (FunNuloVal(GetReportVariableValue("SOMAVENCERSEMAGIO")) + XLF_VENCERSEMAGIO)
    SetReportVariableValue "SOMATOTAL12", (FunNuloVal(GetReportVariableValue("SOMATOTAL12")) + XLF_VENCIDO + XLF_VENCERCOMAGIO)
    SetReportVariableValue "SOMATOTAL13", (FunNuloVal(GetReportVariableValue("SOMATOTAL13")) + XLF_VENCIDO + XLF_VENCERSEMAGIO)
    
  'se foi pago
  Else
    XLD_DATABASE = XLO_CAMPOS.GetItemByName("titu_dt_BasePagto").Value
       
    XLF_VALORLIQUIDADOHISTORICO = Format(FunCalculoValorPago(XLO_CAMPOS, FormulaFields.GetItemByName("FiltroTitulosPagos").Value), "standard")
        
    'Pegar a cotação da moeda3 ignorando se for indexador zero
    SubCotacao3 XLO_CAMPOS, XLO_VETOR
    If XGT_INCLUIRACRESCIMOS = "S" Then
      XLF_VALORLIQUIDADOINDEXADO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORLIQUIDADOHISTORICO), "##,##0.0000")
    Else
      XLF_VALORLIQUIDADOINDEXADO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, Format(FunCalculoValorPago(XLO_CAMPOS, "0"), "standard")), "##,##0.0000")
    End If
    
    
    'Pega a cotação 1 ou 2
    SubCotacao2 XLO_CAMPOS, XLO_VETOR
    XLF_VALORLIQUIDADOATUALIZADO = Format(FunCalculoValorAtualizado(XLO_CAMPOS, XLO_VETOR, CDate(FormulaFields.GetItemByName("FiltroDataBase").Value), XLF_VALORLIQUIDADOINDEXADO, XLI_MOEDA), "standard")
    
    '------------------------------- Somatorio ----------------------------------------
    SetReportVariableValue "SOMALIQUIDADOHISTORICO", (FunNuloVal(GetReportVariableValue("SOMALIQUIDADOHISTORICO")) + XLF_VALORLIQUIDADOHISTORICO)
    SetReportVariableValue "SOMALIQUIDADOATUALIZADO", (FunNuloVal(GetReportVariableValue("SOMALIQUIDADOATUALIZADO")) + XLF_VALORLIQUIDADOATUALIZADO)
             
  End If
  
    
  '------------------------------- Formatar Cor ----------------------------------------
'  If XLF_VALORPREVISTO = 0 Then
'    ForTitulo.TextColor = vbRed
'    FldDataPagamento.TextColor = vbRed
'    FldDataVencimento.TextColor = vbRed
'    TxtValorTitulo.TextColor = vbRed
'    TxtMoeda.TextColor = vbRed
'    TxtValorPrevisto.TextColor = vbRed
'    TxtValorOriginal.TextColor = vbRed
'    TxtJuros.TextColor = vbRed
'    TxtMulta.TextColor = vbRed
'    TxtDesagio.TextColor = vbRed
'    TxtObservacao.TextColor = vbRed
'  End If

  Set XLO_VETOR = Nothing
  
End Sub

Private Sub SecFooterMes1_Format(ByVal pFormattingInfo As Object)
  
  TxtVencido.SetText Format(FunNuloVal(GetReportVariableValue("SOMAVENCIDO")), "Standard")
  TxtVencerComAgio.SetText Format(FunNuloVal(GetReportVariableValue("SOMAVENCERCOMAGIO")), "Standard")
  TxtVencerSemAgio.SetText FunTrataZero(Format(FunNuloVal(GetReportVariableValue("SOMAVENCERSEMAGIO")), "Standard"))
  TxtTotal12.SetText FunTrataZero(Format(FunNuloVal(GetReportVariableValue("SOMATOTAL12")), "Standard"))
  TxtTotal13.SetText FunTrataZero(Format(FunNuloVal(GetReportVariableValue("SOMATOTAL13")), "standard"))
  TxtLiquidadoHistorico.SetText FunTrataZero(Format(FunNuloVal(GetReportVariableValue("SOMALIQUIDADOHISTORICO")), "standard"))
  TxtLiquidadoAtualizado.SetText FunTrataZero(Format(FunNuloVal(GetReportVariableValue("SOMALIQUIDADOATUALIZADO")), "standard"))
  
  SetReportVariableValue "TOTALVENCIDO", GetReportVariableValue("TOTALVENCIDO") + GetReportVariableValue("SOMAVENCIDO")
  SetReportVariableValue "TOTALVENCERCOMAGIO", GetReportVariableValue("TOTALVENCERCOMAGIO") + GetReportVariableValue("SOMAVENCERCOMAGIO")
  SetReportVariableValue "TOTALVENCERSEMAGIO", GetReportVariableValue("TOTALVENCERSEMAGIO") + GetReportVariableValue("SOMAVENCERSEMAGIO")
  SetReportVariableValue "TOTALTOTAL12", GetReportVariableValue("TOTALTOTAL12") + GetReportVariableValue("SOMATOTAL12")
  SetReportVariableValue "TOTALTOTAL13", GetReportVariableValue("TOTALTOTAL13") + GetReportVariableValue("SOMATOTAL13")
  SetReportVariableValue "TOTALLIQUIDADOHISTORICO", GetReportVariableValue("TOTALLIQUIDADOHISTORICO") + GetReportVariableValue("SOMALIQUIDADOHISTORICO")
  SetReportVariableValue "TOTALLIQUIDADOATUALIZADO", GetReportVariableValue("TOTALLIQUIDADOATUALIZADO") + GetReportVariableValue("SOMALIQUIDADOATUALIZADO")
  
  SetReportVariableValue "SOMAVENCIDO", 0
  SetReportVariableValue "SOMAVENCERCOMAGIO", 0
  SetReportVariableValue "SOMAVENCERSEMAGIO", 0
  SetReportVariableValue "SOMATOTAL12", 0
  SetReportVariableValue "SOMATOTAL13", 0
  SetReportVariableValue "SOMALIQUIDADOHISTORICO", 0
  SetReportVariableValue "SOMALIQUIDADOATUALIZADO", 0
  
End Sub

Private Sub SecFooterReport_Format(ByVal pFormattingInfo As Object)
  
  TxtSomaVencido.SetText FunNuloVal(Format(GetReportVariableValue("TOTALVENCIDO"), "standard"))
  TxtSomaVencerComAgio.SetText FunNuloVal(Format(GetReportVariableValue("TOTALVENCERCOMAGIO"), "standard"))
  TxtSomaVencerSemAgio.SetText FunNuloVal(FunTrataZero(Format(GetReportVariableValue("TOTALVENCERSEMAGIO"), "standard")))
  TxtSomaTotal12.SetText FunNuloVal(FunTrataZero(Format(GetReportVariableValue("TOTALTOTAL12"), "standard")))
  TxtSomaTotal13.SetText FunNuloVal(FunTrataZero(Format(GetReportVariableValue("TOTALTOTAL13"), "standard")))
  TxtSomaLiquidadoHistorico.SetText FunNuloVal(Format(GetReportVariableValue("TOTALLIQUIDADOHISTORICO"), "standard"))
  TxtSomaLiquidadoAtualizado.SetText FunNuloVal(Format(GetReportVariableValue("TOTALLIQUIDADOATUALIZADO"), "standard"))
  
End Sub

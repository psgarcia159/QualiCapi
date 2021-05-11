VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptFluxoGeralPagamento 
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12105
   OleObjectBlob   =   "RptFluxoGeralPagamento.dsx":0000
End
Attribute VB_Name = "RptFluxoGeralPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Dim VFV_TOTALMOEDALIQUIDADO As New XArray

Dim XFB_FOOTER As Boolean 'indica se passou pelo footer
Dim XFT_TITULOANTERIOR As String

Private Sub Report_Initialize()

AddReportVariable crRVNumber, "SOMAVALORPREVISTO"
AddReportVariable crRVNumber, "SOMAVALORPAGO"
AddReportVariable crRVNumber, "TOTALLIQUIDADOATUALIZADO" 'Usado para somar o valor atualizado dos titulos liquidados

AddReportVariable crRVNumber, "VALORPAGOINDEXADO"

XFB_FOOTER = False
VFV_TOTALMOEDALIQUIDADO.ReDim 0, 0, 0, 2

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
  
  'se passou pelo footer zera as variáveis
  If XFB_FOOTER Then
  
    VFV_TOTALMOEDALIQUIDADO.Clear
    VFV_TOTALMOEDALIQUIDADO.ReDim 0, 0, 0, 2
  
    SetReportVariableValue "SOMAVALORPREVISTO", 0
    SetReportVariableValue "SOMAVALORPAGO", 0
    
    SetReportVariableValue "TOTALLIQUIDADOATUALIZADO", 0
    
    XFB_FOOTER = False
  End If
  
  Set XLO_CAMPOS = Me.Database.Tables(1).Fields
  
  XLI_INDEXADOR = CByte(FormulaFields.GetItemByName("Indexador").Value)
  
  'Pega o código da moeda e a sua descrição
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
    
    If XGT_INCLUIRACRESCIMOS = "S" Then
      XLF_VALORTITULO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORPREVISTO), "##,##0.0000")
    Else
      XLF_VALORTITULO = Format(XLO_VETOR(0, 0), "##,##0.0000")
    End If
    
    XLF_VALORORIGINAL = Format(FunCalculoValorOriginalPago(XLO_CAMPOS), "standard")
    XLF_VALORPAGO = Format(FunCalculoValorPago(XLO_CAMPOS, FormulaFields.GetItemByName("FiltroTitulosPagos").Value), "standard")
    
    'Multa, juros e deságio pagos
    If FormulaFields.GetItemByName("FiltroTitulosPagos").Value = "1" Then
      XLF_JUROS = Format(XLO_CAMPOS.GetItemByName("titu_vl_Juros").Value, "standard")
      XLF_MULTA = Format(XLO_CAMPOS.GetItemByName("titu_vl_Multa").Value, "standard")
    End If
    XLF_DESAGIO = Format(XLO_CAMPOS.GetItemByName("titu_vl_Desagio").Value, "standard")
    
    'Pegar a cotação da moeda3 ignorando se for indexador zero
    SubCotacao3 XLO_CAMPOS, XLO_VETOR
    If XGT_INCLUIRACRESCIMOS = "S" Then
      XLF_VALORPAGOINDEXADO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORPAGO), "##,##0.0000")
    Else
      XLF_VALORPAGOINDEXADO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, Format(FunCalculoValorPago(XLO_CAMPOS, "0"), "standard")), "##,##0.0000")
    End If
    
    'Pega a cotação 1 ou 2
    SubCotacao2 XLO_CAMPOS, XLO_VETOR
    XLF_VALORPAGOATUALIZADO = Format(FunCalculoValorAtualizado(XLO_CAMPOS, XLO_VETOR, CDate(FormulaFields.GetItemByName("FiltroDataBase").Value), XLF_VALORPAGOINDEXADO, XLI_MOEDA), "standard")
       
    ' Se for o mesmo titulo que o anterior não soma no vetor
    If XFT_TITULOANTERIOR <> XLO_CAMPOS.GetItemByName("Titulo").Value Then
      SubSomaMoeda XLI_MOEDA, XLT_MOEDA, VFV_TOTALMOEDALIQUIDADO, XLF_VALORPAGOINDEXADO
    End If
    SetReportVariableValue "TOTALLIQUIDADOATUALIZADO", (FunNuloVal(GetReportVariableValue("TOTALLIQUIDADOATUALIZADO")) + XLF_VALORPAGOATUALIZADO)
    
    TxtValorPago.SetText Format(XLF_VALORPAGO, "Standard")
    'TxtValorPagoIndexado.SetText Format(XLF_VALORPAGOINDEXADO, "##,##0.0000")
      
  End If
  
  '----------------------------- Campos Texto para pagos e não pagos ----------------
  
  TxtValorTitulo.SetText Format(XLF_VALORTITULO, "##,##0.0000")
  TxtMoeda.SetText XLT_MOEDA
  TxtValorPrevisto.SetText Format(XLF_VALORPREVISTO, "Standard")
  TxtObservacao.SetText FunObservacao(XLO_CAMPOS, XLD_DATABASE)
  
  '------------------------------- Somatorio ----------------------------------------
  
  SetReportVariableValue "SOMAVALORPREVISTO", (FunNuloVal(GetReportVariableValue("SOMAVALORPREVISTO")) + CDbl(Format(XLF_VALORPREVISTO, "Standard")))
  SetReportVariableValue "SOMAVALORPAGO", (FunNuloVal(GetReportVariableValue("SOMAVALORPAGO")) + CDbl(Format(XLF_VALORPAGO, "Standard")))
    
  '------------------------------- Formatar Cor ----------------------------------------
  If XLF_VALORPREVISTO = 0 Or XLO_VETOR(0, 0) = -1 Then
    
    TxtAtencao.SetText "*"
    TxtAtencao.TextColor = vbRed
    
    FldTitulo.TextColor = vbRed
    FldDataPagamento.TextColor = vbRed
    FldDataVencimento.TextColor = vbRed
    TxtValorTitulo.TextColor = vbRed
    TxtMoeda.TextColor = vbRed
    TxtValorPrevisto.TextColor = vbRed
    TxtObservacao.TextColor = vbRed
    TxtValorPago.TextColor = vbRed
    FldRazaoSocial.TextColor = vbRed
  End If
   
  Set XLO_VETOR = Nothing
  XFT_TITULOANTERIOR = XLO_CAMPOS.GetItemByName("Titulo").Value
  
End Sub

Private Sub SecFooterPagamento1_Format(ByVal pFormattingInfo As Object)
  
  TxtSomaValorPrevisto.SetText Format(GetReportVariableValue("SOMAVALORPREVISTO"), "standard")
  TxtSomaValorPago.SetText FunTrataZero(Format(GetReportVariableValue("SOMAVALORPAGO"), "standard"))
  
End Sub

Private Sub SecFooterPagamento2_Format(ByVal pFormattingInfo As Object)
  
  Dim XFI_CONT As Integer
  
  TxtTotalLiquidado.SetText ""
  TxtMoedaLiquidado.SetText ""
    
  For XFI_CONT = 0 To VFV_TOTALMOEDALIQUIDADO.Count(1) - 1
    TxtTotalLiquidado.SetText TxtTotalLiquidado.Text & vbCrLf & Format(VFV_TOTALMOEDALIQUIDADO(XFI_CONT, 2), "##,##0.0000")
    TxtMoedaLiquidado.SetText TxtMoedaLiquidado.Text & vbCrLf & VFV_TOTALMOEDALIQUIDADO(XFI_CONT, 0)
  Next
  
  XFB_FOOTER = True
  
End Sub

Private Sub SecFooterPagamento3_Format(ByVal pFormattingInfo As Object)
    
  TxtTotalLiquidadoAtualizado.SetText FunTrataZero(Format(GetReportVariableValue("TOTALLIQUIDADOATUALIZADO"), "standard"))
  
End Sub

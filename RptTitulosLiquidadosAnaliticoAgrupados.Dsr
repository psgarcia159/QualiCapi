VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptTitulosLiquidadosAnaliticoAgrupados 
   ClientHeight    =   13230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19140
   OleObjectBlob   =   "RptTitulosLiquidadosAnaliticoAgrupados.dsx":0000
End
Attribute VB_Name = "RptTitulosLiquidadosAnaliticoAgrupados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Private Sub Report_Initialize()

AddReportVariable crRVNumber, "SOMAVALORPREVISTO"
AddReportVariable crRVNumber, "SOMAVALORORIGINAL"
AddReportVariable crRVNumber, "SOMADISTRATO"

End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SecDetalhe_Format(ByVal pFormattingInfo As Object)
  
  Dim XLF_VALORCORRIGIDO As Double  'Valor corrigido monetariamente
  Dim XLF_DESAGIO As Double  'Valor do deságio
  Dim XLF_MULTA As Double  'Valor da Multa
  Dim XLF_JUROS As Double  'Valor do Juros
  Dim XLT_FINALTITULO As String
  
  
  Dim XLF_VALORPREVISTO As Double 'Valor Previsto
  Dim XLF_VALORORIGINAL As Double 'Valor Original
  Dim XLF_VALORPREVISTOIMP As Double 'Valor Previsto
  Dim XLF_VALORINICIAL As Double 'Valor Histórico
  
  'Vetor auxiliar para armazenar o valor do titulo, a cotação e a cotação do mês anterior,
  'tudo de acordo com o indexador escolhido. Também indica se tem indexador ou não.
  Dim XLO_VETOR As New XArray
     
  Dim XLD_DATABASE As Date  'Data base passada por parâmetro para relatório
  Dim XLI_INDEXADOR As Byte  ' Indexador passado por parâmetro para o relatório
  Dim XLI_MOEDA As Integer 'Código da moeda de acordo com o indexador escolhido
  Dim XLB_COMRESIDUO As Boolean
  
  XLB_COMRESIDUO = False
  
  
  Dim XLO_CAMPOS As CRAXDDRT.DatabaseFieldDefinitions  ' Objeto que representa os campos do relatório
  
  'se passou pelo footer zera as variáveis
  Set XLO_CAMPOS = Me.Database.Tables(1).Fields
  
  XLI_INDEXADOR = 1 'Só para compor os parâmetros das fórmulas
  
    
  'XLI_INDEXADOR = CByte(FormulaFields.GetItemByName("Indexador").Value)
  
  'Pega o código da moeda e a sua descrição
  'If XLO_CAMPOS.GetItemByName("Pago").Value = "N" Then
  '  SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, True
  'Else
    SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  'End If
  
  XLI_MOEDA = XLO_VETOR(1, 0)
  
  
  'Pega as cotações do mês da data base e do mês anterior
  SubCotacao XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  
  'XLD_DATABASE = XLO_CAMPOS.GetItemByName("titu_dt_BasePagto").Value
  
  'XLF_VALORCORRIGIDO = Format(FunCalculoCorrecaoMonetaria(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLI_MOEDA), "Standard")
  
  'Juros da CEF
  'XLF_VALORCORRIGIDO = XLF_VALORCORRIGIDO + Format(FunCalculoJurosCEF(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  

   
  'Verifica se o título possue resíduo
  If Me.FormulaFields.GetItemByName("ComResiduo").Value = "1" Then
    XLB_COMRESIDUO = True
  End If
    
  'XLF_JUROS = Format(FunCalculoJuros(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  'XLF_MULTA = Format(FunCalculoMulta(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  'XLF_DESAGIO = Format(FunCalculoDesagio(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  
  
  'XLF_VALORPREVISTO = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA - XLF_DESAGIO
  XLF_VALORCORRIGIDO = 0
  
  If Not XLB_COMRESIDUO Then
      
       'valor histórico
        XLF_VALORINICIAL = XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value

        If Not IsNull(Me.FormulaFields.GetItemByName("ValorCorrigido").Value) And Not IsEmpty(Me.FormulaFields.GetItemByName("ValorCorrigido").Value) Then
            XLF_VALORCORRIGIDO = Format(Me.FormulaFields.GetItemByName("ValorCorrigido").Value, "standard") 'Valor corrigido
        End If

        If XLF_VALORCORRIGIDO = 0 Then
            XLF_VALORCORRIGIDO = XLF_VALORINICIAL
        End If

        XLF_VALORORIGINAL = XLF_VALORCORRIGIDO
  Else
    XLF_VALORORIGINAL = Format(FunCalculoValorOriginalPago(XLO_CAMPOS), "standard")

  End If
      
  XLF_VALORPREVISTO = XLF_VALORORIGINAL
  
  'XLT_FINALTITULO = Mid(XLO_CAMPOS.GetItemByName("Titulo").Value, Len(XLO_CAMPOS.GetItemByName("Titulo").Value) - 1, 2)
  
  'XLF_VALORPREVISTOIMP = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA - XLF_DESAGIO
  
  'If XLT_FINALTITULO = "00" Then
    'XLF_VALORPREVISTO = XLF_VALORPREVISTOIMP
  'End If
       
  '----------------------------- Campos Texto para pagos e não pagos ----------------
  
  TxtValorPrevisto.SetText Format(XLF_VALORPREVISTO, "Standard")
  TxtValorOriginal.SetText Format(XLF_VALORORIGINAL, "Standard")
  
  
  '------------------------------- Somatorio ----------------------------------------
  
  SetReportVariableValue "SOMAVALORPREVISTO", (FunNuloVal(GetReportVariableValue("SOMAVALORPREVISTO")) + CDbl(Format(XLF_VALORPREVISTO, "Standard")))
  SetReportVariableValue "SOMAVALORORIGINAL", (FunNuloVal(GetReportVariableValue("SOMAVALORORIGINAL")) + CDbl(Format(XLF_VALORORIGINAL, "Standard")))
    
  '------------------------------- Formatar Cor ----------------------------------------
  If XLF_VALORPREVISTO = 0 Then
    FldTitulo.TextColor = vbRed
    FldDataVencimento.TextColor = vbRed
    FldDataPagamento.TextColor = vbRed
    FldDataBasePagamento.TextColor = vbRed
    TxtValorPrevisto.TextColor = vbRed
    TxtValorOriginal.TextColor = vbRed
  End If
  
  Set XLO_VETOR = Nothing
End Sub

Private Sub SecFooterReport_Format(ByVal pFormattingInfo As Object)
  
  TxtSomaValorPrevisto.SetText FunTrataZero(Format(GetReportVariableValue("SOMAVALORPREVISTO"), "standard"))
  TxtSomaValorOriginal.SetText FunTrataZero(Format(GetReportVariableValue("SOMAVALORORIGINAL"), "standard"))
  
End Sub

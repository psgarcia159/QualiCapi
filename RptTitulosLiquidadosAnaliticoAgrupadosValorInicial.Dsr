VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptTitulosLiquidadosAnaliticoAgrupadosValorInicial 
   ClientHeight    =   13230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19140
   OleObjectBlob   =   "RptTitulosLiquidadosAnaliticoAgrupadosValorInicial.dsx":0000
End
Attribute VB_Name = "RptTitulosLiquidadosAnaliticoAgrupadosValorInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Private Sub Report_Initialize()

AddReportVariable crRVNumber, "SOMAVALORINICIAL" 'Valor inicial do histórico
AddReportVariable crRVNumber, "SOMAVALORORIGINAL"
AddReportVariable crRVNumber, "SOMADISTRATO"
AddReportVariable crRVNumber, "SOMAVALORCORRECAO"
AddReportVariable crRVNumber, "SOMAVALORATUAL"

End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SecDetalhe_Format(ByVal pFormattingInfo As Object)
  
  Dim XLF_VALORCORRIGIDO As Double  'Valor corrigido monetariamente
  Dim XLF_VALORINICIAL As Double 'Valor do histórico
  Dim XLF_VALORORIGINAL As Double 'Valor original pago
  Dim XLF_VALORCORRECAO As Double 'Valor da correção monetária
  Dim XLF_VALORPAGOINDEXADO As Double
  Dim XLF_VALORPAGO As Double
  Dim XLF_COTACAO4 As Double
  Dim XLF_VALORTITULO As Double
  
  Dim XLO_VETOR As New XArray
  
  Dim XLD_DATABASE As Date  'Data base passada por parâmetro para relatório
  Dim XLI_INDEXADOR As Byte  ' Indexador passado por parâmetro para o relatório
  Dim XLO_CAMPOS As CRAXDDRT.DatabaseFieldDefinitions  ' Objeto que representa os campos do relatório
  Dim XLI_MOEDA As Integer 'Código da moeda de acordo com o indexador escolhido
  Dim XLT_MOEDA As String 'Decrição da moeda de acordo com o indexador escolhido
  Dim XLB_COMRESIDUO As Boolean
  
  Set XLO_CAMPOS = Me.Database.Tables(1).Fields
  
  XLI_INDEXADOR = 1
  XLB_COMRESIDUO = False
  
  'Pega o código da moeda e a sua descrição
  SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  XLT_MOEDA = XLO_VETOR(0, 0)
  XLI_MOEDA = XLO_VETOR(1, 0)
  
  'Pega as cotações do mês da data base e do mês anterior
  SubCotacao XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  
  XLF_VALORCORRECAO = 0
  XLF_COTACAO4 = 0
  XLF_VALORCORRIGIDO = 0
  
  'Verifica se o título possue resíduo
  If Me.FormulaFields.GetItemByName("ComResiduo").Value = "1" Then
    XLB_COMRESIDUO = True
  End If

  If Not XLB_COMRESIDUO Then
        
        'valor histórico
        XLF_VALORINICIAL = XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value

        If Not IsNull(Me.FormulaFields.GetItemByName("ValorCorrigido").Value) And Not IsEmpty(Me.FormulaFields.GetItemByName("ValorCorrigido").Value) Then
            XLF_VALORCORRIGIDO = Format(Me.FormulaFields.GetItemByName("ValorCorrigido").Value, "standard") 'Valor corrigido
        End If

        If XLF_VALORCORRIGIDO > 0 And XLF_VALORCORRIGIDO <> XLF_VALORINICIAL Then

            XLF_VALORCORRECAO = XLF_VALORCORRIGIDO - XLF_VALORINICIAL 'Correção
        Else
            XLF_VALORCORRIGIDO = XLF_VALORINICIAL
        End If

        XLF_VALORORIGINAL = XLF_VALORCORRIGIDO

   Else

        XLF_VALORORIGINAL = Format(FunCalculoValorOriginalPago(XLO_CAMPOS), "standard") 'Valor corrigido pago

        XLF_VALORCORRIGIDO = XLF_VALORORIGINAL

        'Data base do pagamento
        XLD_DATABASE = XLO_CAMPOS.GetItemByName("titu_dt_BasePagto").Value

        'Valor do título indexado
        XLF_VALORTITULO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORCORRIGIDO), "##,##0.0000")

        'Cotação da moeda paga
        SubCotacao3 XLO_CAMPOS, XLO_VETOR

        'Valor pago indexado
        XLF_VALORPAGOINDEXADO = Format(FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORORIGINAL), "##,##0.0000")

        If Not IsNull(Me.FormulaFields.GetItemByName("Cotacao4").Value) And Not IsEmpty(Me.FormulaFields.GetItemByName("Cotacao4").Value) Then
             XLF_COTACAO4 = Me.FormulaFields.GetItemByName("Cotacao4").Value
        End If

        If XLF_COTACAO4 <> 0 Then

             XLF_VALORCORRECAO = Format(((XLF_VALORCORRIGIDO / XLF_VALORTITULO) - XLF_COTACAO4) * XLF_VALORPAGOINDEXADO, "standard")

        End If

        XLF_VALORINICIAL = XLF_VALORORIGINAL - XLF_VALORCORRECAO

   End If
     
  '----------------------------- Campos Texto para pagos e não pagos ----------------
  
  TxtValorInicial.SetText Format(XLF_VALORINICIAL, "Standard")
  TxtValorOriginal.SetText Format(XLF_VALORORIGINAL, "Standard")
  TxtValorCorrecao.SetText Format(XLF_VALORCORRECAO, "Standard")
  TxtValorAtual.SetText Format(XLF_VALORCORRIGIDO, "Standard")
  
  '------------------------------- Somatorio ----------------------------------------
  
  SetReportVariableValue "SOMAVALORINICIAL", (FunNuloVal(GetReportVariableValue("SOMAVALORINICIAL")) + CDbl(Format(XLF_VALORINICIAL, "Standard")))
  SetReportVariableValue "SOMAVALORORIGINAL", (FunNuloVal(GetReportVariableValue("SOMAVALORORIGINAL")) + CDbl(Format(XLF_VALORORIGINAL, "Standard")))
  SetReportVariableValue "SOMAVALORCORRECAO", (FunNuloVal(GetReportVariableValue("SOMAVALORCORRECAO")) + CDbl(Format(XLF_VALORCORRECAO, "Standard")))
  SetReportVariableValue "SOMAVALORATUAL", (FunNuloVal(GetReportVariableValue("SOMAVALORATUAL")) + CDbl(Format(XLF_VALORCORRIGIDO, "Standard")))
    
  '------------------------------- Formatar Cor ----------------------------------------
  If XLF_VALORINICIAL = 0 Then
    FldTitulo.TextColor = vbRed
    FldDataVencimento.TextColor = vbRed
    FldDataPagamento.TextColor = vbRed
    FldDataBasePagamento.TextColor = vbRed
    TxtValorInicial.TextColor = vbRed
    TxtValorOriginal.TextColor = vbRed
  End If
  
End Sub

Private Sub SecFooterReport_Format(ByVal pFormattingInfo As Object)
  
  TxtSomaValorInicial.SetText FunTrataZero(Format(GetReportVariableValue("SOMAVALORINICIAL"), "standard"))
  TxtSomaValorOriginal.SetText FunTrataZero(Format(GetReportVariableValue("SOMAVALORORIGINAL"), "standard"))
  TxtSomaValorCorrecao.SetText FunTrataZero(Format(GetReportVariableValue("SOMAVALORCORRECAO"), "standard"))
  TxtValorAtualTotal.SetText FunTrataZero(Format(GetReportVariableValue("SOMAVALORATUAL"), "standard"))

End Sub


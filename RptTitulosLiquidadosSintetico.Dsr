VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptTitulosLiquidadosSintetico 
   ClientHeight    =   11370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12990
   OleObjectBlob   =   "RptTitulosLiquidadosSintetico.dsx":0000
End
Attribute VB_Name = "RptTitulosLiquidadosSintetico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Private Sub Report_Initialize()

AddReportVariable crRVNumber, "SOMAVALORPREVISTO"
AddReportVariable crRVNumber, "SOMAVALORORIGINAL"
AddReportVariable crRVNumber, "TOTALVALORPREVISTO"
AddReportVariable crRVNumber, "TOTALVALORORIGINAL"

End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SecDetalhe2_Format(ByVal pFormattingInfo As Object)
  
  Dim XLF_VALORCORRIGIDO As Double  'Valor corrigido monetariamente
  Dim XLF_DESAGIO As Double  'Valor do deságio
  Dim XLF_MULTA As Double  'Valor da Multa
  Dim XLF_JUROS As Double  'Valor do Juros
  Dim XLT_FINALTITULO As String 'Informação que identifica se título possui resíduo ou não
  
  Dim XLF_VALORPREVISTO As Double 'Valor Previsto sem resíduos
  Dim XLF_VALORORIGINAL As Double 'Valor Original
  Dim XLF_VALORPREVISTOIMP As Double 'Valor Previsto
  
  'Vetor auxiliar para armazenar o valor do titulo, a cotação e a cotação do mês anterior,
  'tudo de acordo com o indexador escolhido. Também indica se tem indexador ou não.
  Dim XLO_VETOR As New XArray
     
  Dim XLD_DATABASE As Date  'Data base passada por parâmetro para relatório
  Dim XLI_INDEXADOR As Byte  ' Indexador passado por parâmetro para o relatório
  Dim XLI_MOEDA As Integer 'Código da moeda de acordo com o indexador escolhido
  Dim XLB_COMRESIDUO As Boolean
  
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
  
  XLD_DATABASE = XLO_CAMPOS.GetItemByName("titu_dt_BasePagto").Value
  
  If Not IsNull(Me.FormulaFields.GetItemByName("ValorCorrigido").Value) And Not IsEmpty(Me.FormulaFields.GetItemByName("ValorCorrigido").Value) Then
     XLF_VALORCORRIGIDO = Format(Me.FormulaFields.GetItemByName("ValorCorrigido").Value, "standard") 'Valor corrigido
  End If
   
  'Verifica se o título possue resíduo
  If Me.FormulaFields.GetItemByName("ComResiduo").Value = "1" Then
    XLB_COMRESIDUO = True
  End If
  
'  XLF_VALORCORRIGIDO = Format(FunCalculoCorrecaoMonetaria(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLI_MOEDA), "Standard")
'  XLF_JUROS = Format(FunCalculoJuros(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
'  XLF_MULTA = Format(FunCalculoMulta(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
'  XLF_DESAGIO = Format(FunCalculoDesagio(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
    
  'XLF_VALORPREVISTO = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA - XLF_DESAGIO
  XLF_VALORORIGINAL = Format(FunCalculoValorOriginalPago(XLO_CAMPOS), "standard")
  
  'If XLB_COMRESIDUO Then
    XLF_VALORPREVISTO = XLF_VALORORIGINAL
  'Else
  '  XLF_VALORPREVISTO = XLF_VALORCORRIGIDO
  'End If
  
'  XLT_FINALTITULO = Mid(XLO_CAMPOS.GetItemByName("Titulo").Value, Len(XLO_CAMPOS.GetItemByName("Titulo").Value) - 1, 2)
  
'  XLF_VALORPREVISTOIMP = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA - XLF_DESAGIO
'
'  If XLT_FINALTITULO = "00" Then
'    XLF_VALORPREVISTO = XLF_VALORPREVISTOIMP
'  End If
  
  '------------------------------- Somatorio ----------------------------------------
  
  SetReportVariableValue "SOMAVALORPREVISTO", (FunNuloVal(GetReportVariableValue("SOMAVALORPREVISTO")) + CDbl(Format(XLF_VALORPREVISTO, "Standard")))
  SetReportVariableValue "SOMAVALORORIGINAL", (FunNuloVal(GetReportVariableValue("SOMAVALORORIGINAL")) + CDbl(Format(XLF_VALORORIGINAL, "Standard")))
    
  Set XLO_VETOR = Nothing
End Sub

Private Sub SecFooterEmpreendimento1_Format(ByVal pFormattingInfo As Object)
      
  TxtSomaValorPrevisto.SetText Format(GetReportVariableValue("SOMAVALORPREVISTO"), "standard")
  TxtSomaValorOriginal.SetText Format(GetReportVariableValue("SOMAVALORORIGINAL"), "standard")
    
  SetReportVariableValue "TOTALVALORPREVISTO", (FunNuloVal(GetReportVariableValue("TOTALVALORPREVISTO")) + GetReportVariableValue("SOMAVALORPREVISTO"))
  SetReportVariableValue "TOTALVALORORIGINAL", (FunNuloVal(GetReportVariableValue("TOTALVALORORIGINAL")) + GetReportVariableValue("SOMAVALORORIGINAL"))
  
  '------------------------------- Formatar Cor ----------------------------------------
  If GetReportVariableValue("SOMAVALORPREVISTO") = 0 Then
    TxtSomaValorPrevisto.TextColor = vbRed
    TxtSomaValorOriginal.TextColor = vbRed
  End If
  
  SetReportVariableValue "SOMAVALORPREVISTO", 0
  SetReportVariableValue "SOMAVALORORIGINAL", 0
  
End Sub

Private Sub SecFooterReport_Format(ByVal pFormattingInfo As Object)
  
  TxtTotalValorPrevisto.SetText Format(GetReportVariableValue("TOTALVALORPREVISTO"), "standard")
  TxtTotalValorOriginal.SetText Format(GetReportVariableValue("TOTALVALORORIGINAL"), "standard")

End Sub

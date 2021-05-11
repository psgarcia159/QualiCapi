VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptPrevisaoRecebimentoSinteticoDiario 
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11460
   OleObjectBlob   =   "RptPrevisaoRecebimentoSinteticoDiario.dsx":0000
End
Attribute VB_Name = "RptPrevisaoRecebimentoSinteticoDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Private Sub Report_Initialize()

  AddReportVariable crRVNumber, "SOMAVALORPREVISTO"
  AddReportVariable crRVNumber, "TOTALVALORPREVISTO"
  AddReportVariable crRVNumber, "ACUMULADO"

End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SectHeaderReport_Format(ByVal pFormattingInfo As Object)
  
  SetReportVariableValue "SOMAVALORPREVISTO", 0
  SetReportVariableValue "TOTALVALORPREVISTO", 0
  SetReportVariableValue "ACUMULADO", 0
  
End Sub

Private Sub SecDetalhe_Format(ByVal pFormattingInfo As Object)
  
  Dim XLF_VALORCORRIGIDO As Double  'Valor corrigido monetariamente
  Dim XLF_DESAGIO As Double  'Valor do deságio
  Dim XLF_MULTA As Double  'Valor da Multa
  Dim XLF_JUROS As Double  'Valor do Juros
   
  Dim XLF_VALORPREVISTO As Double 'Valor Previsto
  
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
  
  'Pega o código da moeda e a sua descrição
  SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, True
  XLT_MOEDA = XLO_VETOR(0, 0)
  XLI_MOEDA = XLO_VETOR(1, 0)
  
  'Pega as cotações do mês da data base e do mês anterior
  SubCotacao XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, True
    
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
   
  '------------------------------- Somatorio ----------------------------------------
   
  SetReportVariableValue "SOMAVALORPREVISTO", (FunNuloVal(GetReportVariableValue("SOMAVALORPREVISTO")) + XLF_VALORPREVISTO)
  SetReportVariableValue "ACUMULADO", (FunNuloVal(GetReportVariableValue("ACUMULADO")) + XLF_VALORPREVISTO)
    
  Set XLO_VETOR = Nothing
  
End Sub

Private Sub SecFooterProorogacao1_Format(ByVal pFormattingInfo As Object)
  
  TxtValorPrevisto.SetText Format(GetReportVariableValue("SOMAVALORPREVISTO"), "Standard")
  TxtValorAcumulado.SetText Format(GetReportVariableValue("ACUMULADO"), "Standard")
    
  '------------------------------- Formatar Cor ----------------------------------------
  If GetReportVariableValue("SOMAVALORPREVISTO") = 0 Then
    FldDataVencimento.TextColor = vbRed
    FldDiaSemana.TextColor = vbRed
    TxtValorPrevisto.TextColor = vbRed
    TxtValorAcumulado.TextColor = vbRed
  End If
  
  SetReportVariableValue "TOTALVALORPREVISTO", (FunNuloVal(GetReportVariableValue("TOTALVALORPREVISTO")) + GetReportVariableValue("SOMAVALORPREVISTO"))
  SetReportVariableValue "SOMAVALORPREVISTO", 0
  
End Sub
Private Sub SecFooterEmpreendimento1_Format(ByVal pFormattingInfo As Object)
  
  TxtTotalValorPrevisto.SetText Format(GetReportVariableValue("TOTALVALORPREVISTO"), "standard")

  SetReportVariableValue "TOTALVALORPREVISTO", 0
  SetReportVariableValue "ACUMULADO", 0
  
End Sub

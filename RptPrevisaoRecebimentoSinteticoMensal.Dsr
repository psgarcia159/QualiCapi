VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptPrevisaoRecebimentoSinteticoMensal 
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11460
   OleObjectBlob   =   "RptPrevisaoRecebimentoSinteticoMensal.dsx":0000
End
Attribute VB_Name = "RptPrevisaoRecebimentoSinteticoMensal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Private Sub Report_Initialize()

  AddReportVariable crRVNumber, "SOMAVALORPREVISTO"
  AddReportVariable crRVNumber, "SOMAAVISTA"
  AddReportVariable crRVNumber, "SOMASINAL"
  AddReportVariable crRVNumber, "SOMAPOUPANCA"
  AddReportVariable crRVNumber, "SOMAFINANCIAMENTO"
  AddReportVariable crRVNumber, "SOMAFGTS"
  AddReportVariable crRVNumber, "SOMAOUTROS"
  
  AddReportVariable crRVNumber, "TOTALVALORPREVISTO"
  AddReportVariable crRVNumber, "TOTALAVISTA"
  AddReportVariable crRVNumber, "TOTALSINAL"
  AddReportVariable crRVNumber, "TOTALPOUPANCA"
  AddReportVariable crRVNumber, "TOTALFINANCIAMENTO"
  AddReportVariable crRVNumber, "TOTALFGTS"
  AddReportVariable crRVNumber, "TOTALOUTROS"
  
  AddReportVariable crRVNumber, "ACUMULADO"
  
End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SectHeaderReport_Format(ByVal pFormattingInfo As Object)
   
  SetReportVariableValue "SOMAVALORPREVISTO", 0
  SetReportVariableValue "SOMAAVISTA", 0
  SetReportVariableValue "SOMASINAL", 0
  SetReportVariableValue "SOMAPOUPANCA", 0
  SetReportVariableValue "SOMAFINANCIAMENTO", 0
  SetReportVariableValue "SOMAFGTS", 0
  SetReportVariableValue "SOMAOUTROS", 0
   
  SetReportVariableValue "TOTALVALORPREVISTO", 0
  SetReportVariableValue "TOTALAVISTA", 0
  SetReportVariableValue "TOTALSINAL", 0
  SetReportVariableValue "TOTALPOUPANCA", 0
  SetReportVariableValue "TOTALFINANCIAMENTO", 0
  SetReportVariableValue "TOTALFGTS", 0
  SetReportVariableValue "TOTALOUTROS", 0
  
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
  
  Select Case XLO_CAMPOS.GetItemByName("titu_tx_NaturezaPlano").Value
    Case "V":
      SetReportVariableValue "SOMAAVISTA", (FunNuloVal(GetReportVariableValue("SOMAAVISTA")) + XLF_VALORPREVISTO)
    Case "S":
      SetReportVariableValue "SOMASINAL", (FunNuloVal(GetReportVariableValue("SOMASINAL")) + XLF_VALORPREVISTO)
    Case "P":
      SetReportVariableValue "SOMAPOUPANCA", (FunNuloVal(GetReportVariableValue("SOMAPOUPANCA")) + XLF_VALORPREVISTO)
    Case "F":
      SetReportVariableValue "SOMAFINANCIAMENTO", (FunNuloVal(GetReportVariableValue("SOMAFINANCIAMENTO")) + XLF_VALORPREVISTO)
    Case "G":
      SetReportVariableValue "SOMAFGTS", (FunNuloVal(GetReportVariableValue("SOMAFGTS")) + XLF_VALORPREVISTO)
    Case "O":
      SetReportVariableValue "SOMAOUTROS", (FunNuloVal(GetReportVariableValue("SOMAOUTROS")) + XLF_VALORPREVISTO)
  End Select
  
   
  SetReportVariableValue "SOMAVALORPREVISTO", (FunNuloVal(GetReportVariableValue("SOMAVALORPREVISTO")) + XLF_VALORPREVISTO)
  SetReportVariableValue "ACUMULADO", (FunNuloVal(GetReportVariableValue("ACUMULADO")) + XLF_VALORPREVISTO)
   
  Set XLO_VETOR = Nothing
  
End Sub

Private Sub SecFooterProrrogacao1_Format(ByVal pFormattingInfo As Object)
   
  TxtValorPrevisto.SetText Format(GetReportVariableValue("SOMAVALORPREVISTO"), "Standard")
  TxtValorAcumulado.SetText Format(GetReportVariableValue("ACUMULADO"), "Standard")
  
  TxtAvista.SetText FunTrataZero(Format(GetReportVariableValue("SOMAAVISTA"), "Standard"))
  TxtSinal.SetText FunTrataZero(Format(GetReportVariableValue("SOMASINAL"), "Standard"))
  TxtPoupanca.SetText FunTrataZero(Format(GetReportVariableValue("SOMAPOUPANCA"), "Standard"))
  TxtFinanciamento.SetText FunTrataZero(Format(GetReportVariableValue("SOMAFINANCIAMENTO"), "Standard"))
  TxtFgts.SetText FunTrataZero(Format(GetReportVariableValue("SOMAFGTS"), "Standard"))
  TxtOutros.SetText FunTrataZero(Format(GetReportVariableValue("SOMAOUTROS"), "Standard"))
  
  '------------------------------- Formatar Cor ----------------------------------------
  If GetReportVariableValue("SOMAVALORPREVISTO") = 0 Then
    FldMesVencimento.TextColor = vbRed
    TxtValorPrevisto.TextColor = vbRed
    TxtValorAcumulado.TextColor = vbRed
  End If
  
  SetReportVariableValue "TOTALVALORPREVISTO", (FunNuloVal(GetReportVariableValue("TOTALVALORPREVISTO")) + GetReportVariableValue("SOMAVALORPREVISTO"))
  SetReportVariableValue "TOTALAVISTA", (FunNuloVal(GetReportVariableValue("TOTALAVISTA")) + GetReportVariableValue("SOMAAVISTA"))
  SetReportVariableValue "TOTALSINAL", (FunNuloVal(GetReportVariableValue("TOTALSINAL")) + GetReportVariableValue("SOMASINAL"))
  SetReportVariableValue "TOTALPOUPANCA", (FunNuloVal(GetReportVariableValue("TOTALPOUPANCA")) + GetReportVariableValue("SOMAPOUPANCA"))
  SetReportVariableValue "TOTALFINANCIAMENTO", (FunNuloVal(GetReportVariableValue("TOTALFINANCIAMENTO")) + GetReportVariableValue("SOMAFINANCIAMENTO"))
  SetReportVariableValue "TOTALFGTS", (FunNuloVal(GetReportVariableValue("TOTALFGTS")) + GetReportVariableValue("SOMAFGTS"))
  SetReportVariableValue "TOTALOUTROS", (FunNuloVal(GetReportVariableValue("TOTALOUTROS")) + GetReportVariableValue("SOMAOUTROS"))
    
  SetReportVariableValue "SOMAVALORPREVISTO", 0
  SetReportVariableValue "SOMAAVISTA", 0
  SetReportVariableValue "SOMASINAL", 0
  SetReportVariableValue "SOMAPOUPANCA", 0
  SetReportVariableValue "SOMAFINANCIAMENTO", 0
  SetReportVariableValue "SOMAFGTS", 0
  SetReportVariableValue "SOMAOUTROS", 0
  
End Sub

Private Sub SecFooterEmpreendimento1_Format(ByVal pFormattingInfo As Object)
  
  TxtTotalValorPrevisto.SetText FunTrataZero(Format(GetReportVariableValue("TOTALVALORPREVISTO"), "standard"))
  TxtTotalAvista.SetText FunTrataZero(Format(GetReportVariableValue("TOTALAVISTA"), "standard"))
  TxtTotalSinal.SetText FunTrataZero(Format(GetReportVariableValue("TOTALSINAL"), "standard"))
  TxtTotalPoupanca.SetText FunTrataZero(Format(GetReportVariableValue("TOTALPOUPANCA"), "standard"))
  TxtTotalFinanciamento.SetText FunTrataZero(Format(GetReportVariableValue("TOTALFINANCIAMENTO"), "standard"))
  TxtTotalFgts.SetText FunTrataZero(Format(GetReportVariableValue("TOTALFGTS"), "standard"))
  TxtTotalOutros.SetText FunTrataZero(Format(GetReportVariableValue("TOTALOUTROS"), "standard"))

  SetReportVariableValue "TOTALVALORPREVISTO", 0
  SetReportVariableValue "TOTALAVISTA", 0
  SetReportVariableValue "TOTALSINAL", 0
  SetReportVariableValue "TOTALPOUPANCA", 0
  SetReportVariableValue "TOTALFINANCIAMENTO", 0
  SetReportVariableValue "TOTALFGTS", 0
  SetReportVariableValue "TOTALOUTROS", 0
  
  SetReportVariableValue "ACUMULADO", 0
  
End Sub

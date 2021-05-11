VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptSegurosPagos 
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11460
   OleObjectBlob   =   "RptSegurosPagos.dsx":0000
End
Attribute VB_Name = "RptSegurosPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Private Sub Report_Initialize()

  AddReportVariable crRVNumber, "VALORDEBITO"
  AddReportVariable crRVNumber, "SOMAVALORDEBITO"
  
End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SecDetalhe_Format(ByVal pFormattingInfo As Object)
  
  Dim XLF_VALORCORRIGIDO As Double  'Valor corrigido monetariamente
  Dim XLF_DESAGIO As Double  'Valor do deságio
  Dim XLF_VALORDEBITO As Double 'Valor do debito
      
  'Vetor auxiliar para armazenar o valor do titulo, a cotação e a cotação do mês anterior,
  'tudo de acordo com o indexador escolhido. Também indica se tem indexador ou não.
  Dim XLO_VETOR As New XArray
     
  Dim XLD_DATABASE As Date  'Data base passada por parâmetro para relatório
  Dim XLI_INDEXADOR As Byte  ' Indexador passado por parâmetro para o relatório
  Dim XLD_DATABASEDESAGIO As Date
  Dim XLI_MOEDA As Integer 'Código da moeda de acordo com o indexador escolhido
    
  Dim XLO_CAMPOS As CRAXDDRT.DatabaseFieldDefinitions  ' Objeto que representa os campos do relatório
      
  Set XLO_CAMPOS = Me.Database.Tables(1).Fields
  
  XLI_INDEXADOR = CByte(FormulaFields.GetItemByName("Indexador").Value)

  'Pega o código da moeda e a sua descrição
  If XLO_CAMPOS.GetItemByName("Pago").Value = "N" Then
    SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, True
  Else
    SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  End If
  
  XLI_MOEDA = XLO_VETOR(1, 0)

  'Pega as cotações do mês da data base e do mês anterior
  SubCotacao XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, True
    
  XLD_DATABASE = CDate(FormulaFields.GetItemByName("FiltroDataBase").Value)
  
  XLF_VALORCORRIGIDO = Format(FunCalculoCorrecaoMonetaria(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLI_MOEDA), "Standard")
  
  'Juros da CEF
  XLF_VALORCORRIGIDO = XLF_VALORCORRIGIDO + Format(FunCalculoJurosCEF(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
    
  XLF_DESAGIO = Format(FunCalculoDesagio(XLO_CAMPOS, XLO_CAMPOS.GetItemByName("DataBaseDesagio").Value, XLF_VALORCORRIGIDO), "standard")
  XLF_VALORDEBITO = XLF_VALORCORRIGIDO - XLF_DESAGIO
     
  '------------------------------- Somatorio ----------------------------------------
  SetReportVariableValue "VALORDEBITO", (FunNuloVal(GetReportVariableValue("VALORDEBITO")) + XLF_VALORDEBITO)
  
  Set XLO_VETOR = Nothing
  
End Sub

Private Sub SecFooterCliente1_Format(ByVal pFormattingInfo As Object)
  
  Dim XLO_CAMPOS As CRAXDDRT.DatabaseFieldDefinitions  ' Objeto que representa os campos do relatório
  Dim XLF_SEGUROCOMPARCELA As Double
  Dim XLF_SEGUROSEMPARCELA As Double
  Dim XLF_PARCELASEGURO As Double
      
  Set XLO_CAMPOS = Me.Database.Tables(1).Fields
  
  TxtValorDebito.SetText Format(FunNuloVal(GetReportVariableValue("VALORDEBITO")), "standard")
  SetReportVariableValue "SOMAVALORDEBITO", (FunNuloVal(GetReportVariableValue("SOMAVALORDEBITO")) + CDbl(FunNuloVal(GetReportVariableValue("VALORDEBITO"))))
  SetReportVariableValue "VALORDEBITO", 0
  
End Sub

Private Sub SecFooterReport_Format(ByVal pFormattingInfo As Object)
  
  TxtSomaValorDebito.SetText Format(FunNuloVal(GetReportVariableValue("SOMAVALORDEBITO")), "standard")
  SetReportVariableValue "SOMAVALORDEBITO", 0
  
End Sub

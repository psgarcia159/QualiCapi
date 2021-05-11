VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptTitulosVencidos 
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11460
   OleObjectBlob   =   "RptTitulosVencidos.dsx":0000
End
Attribute VB_Name = "RptTitulosVencidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Private Sub Report_Initialize()
  
  AddReportVariable crRVNumber, "SOMAVALORORIGINAL"
  AddReportVariable crRVNumber, "SOMAJUROS"
  AddReportVariable crRVNumber, "SOMAMULTA"
  AddReportVariable crRVNumber, "SOMAVALORPREVISTO"

End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SecDetalhe_Format(ByVal pFormattingInfo As Object)
  
  Dim XLF_VALORCORRIGIDO As Double  'Valor corrigido monetariamente
  Dim XLF_MULTA As Double  'Valor da Multa
  Dim XLF_JUROS As Double  'Valor do Juros
  
  Dim XLF_VALORORIGINAL As Double 'Valor Original
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
    
  XLF_JUROS = Format(FunCalculoJuros(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  XLF_MULTA = Format(FunCalculoMulta(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
      
  XLF_VALORORIGINAL = XLF_VALORCORRIGIDO
  XLF_VALORPREVISTO = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA
   
  '------------------------------- Somatorio ----------------------------------------
  
  SetReportVariableValue "SOMAVALORORIGINAL", (FunNuloVal(GetReportVariableValue("SOMAVALORORIGINAL")) + XLF_VALORORIGINAL)
  SetReportVariableValue "SOMAJUROS", (FunNuloVal(GetReportVariableValue("SOMAJUROS")) + XLF_JUROS)
  SetReportVariableValue "SOMAMULTA", (FunNuloVal(GetReportVariableValue("SOMAMULTA")) + XLF_MULTA)
  SetReportVariableValue "SOMAVALORPREVISTO", (FunNuloVal(GetReportVariableValue("SOMAVALORPREVISTO")) + XLF_VALORPREVISTO)
  
  '----------------------------- Campos Texto para pagos e não pagos ----------------
   
  TxtValorOriginal.SetText Format(XLF_VALORORIGINAL, "Standard")
  TxtJuros.SetText Format(XLF_JUROS, "Standard")
  TxtMulta.SetText Format(XLF_MULTA, "Standard")
  TxtTotal.SetText Format(XLF_VALORPREVISTO, "Standard")
  TxtObservacao.SetText FunObservacao(XLO_CAMPOS, XLD_DATABASE)
    
  '------------------------------- Formatar Cor ----------------------------------------
  If XLF_VALORPREVISTO = 0 Then
    FldDataVencimento.TextColor = vbRed
    FldTitulo.TextColor = vbRed
    FldCliente.TextColor = vbRed
    ForStatus.TextColor = vbRed
    TxtValorOriginal.TextColor = vbRed
    TxtTotal.TextColor = vbRed
    TxtObservacao.TextColor = vbRed
  End If
  
  Set XLO_VETOR = Nothing
  
End Sub

Private Sub SecFooterEmpreendimento1_Format(ByVal pFormattingInfo As Object)
  
  TxtSomaValorOriginal.SetText Format(GetReportVariableValue("SOMAVALORORIGINAL"), "standard")
  TxtSomaJuros.SetText Format(GetReportVariableValue("SOMAJUROS"), "standard")
  TxtSomaMulta.SetText Format(GetReportVariableValue("SOMAMULTA"), "standard")
  TxtSomaTotal.SetText Format(GetReportVariableValue("SOMAVALORPREVISTO"), "standard")

  SetReportVariableValue "SOMAVALORORIGINAL", 0
  SetReportVariableValue "SOMAVALORPREVISTO", 0
  SetReportVariableValue "SOMAJUROS", 0
  SetReportVariableValue "SOMAMULTA", 0
  
End Sub


VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptCartasCobranca 
   ClientHeight    =   10020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
   OleObjectBlob   =   "RptCartasCobranca.dsx":0000
End
Attribute VB_Name = "RptCartasCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SecDetalhe_Format(ByVal pFormattingInfo As Object)
  
  Dim XLF_VALORCORRIGIDO As Double  'Valor corrigido monetariamente
  Dim XLF_DESAGIO As Double  'Valor do deságio
  Dim XLF_MULTA As Double  'Valor da Multa
  Dim XLF_JUROS As Double  'Valor do Juros
  Dim XLF_VALORATUAL As Double 'Valor Atual
  Dim XLF_VALORTOTAL As Double 'Valor Total
  
  'Vetor auxiliar para armazenar o valor do titulo, a cotação e a cotação do mês anterior,
  'tudo de acordo com o indexador escolhido. Também indica se tem indexador ou não.
  Dim XLO_VETOR As New XArray
     
  Dim XLD_DATABASE As Date  'Data base passada por parâmetro para relatório
  Dim XLI_INDEXADOR As Byte  ' Indexador passado por parâmetro para o relatório
  Dim XLI_MOEDA As Integer 'Código da moeda de acordo com o indexador escolhido
  
  Dim XLO_CAMPOS As CRAXDDRT.DatabaseFieldDefinitions  ' Objeto que representa os campos do relatório
      
  Set XLO_CAMPOS = Me.Database.Tables(1).Fields
  
  XLI_INDEXADOR = CByte(FormulaFields.GetItemByName("Indexador").Value)
    
  'Pega o código da moeda e a sua descrição
  If XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value = "" Then
    SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, True
  Else
    SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  End If
  
  XLI_MOEDA = XLO_VETOR(1, 0)
    
  'Pega as cotações do mês da data base e do mês anterior
  SubCotacao XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, True
    
  XLD_DATABASE = CDate(FormulaFields.GetItemByName("FiltroDataBase").Value)
  
  XLF_VALORCORRIGIDO = Format(FunCalculoCorrecaoMonetaria(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLI_MOEDA), "Standard")
  XLF_JUROS = Format(FunCalculoJuros(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  XLF_MULTA = Format(FunCalculoMulta(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
  XLF_VALORATUAL = XLF_VALORCORRIGIDO
  XLF_VALORTOTAL = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA
  
  '----------------------------- Campos Texto para pagos e não pagos ----------------
  
  TxtValorAtual.SetText Format(XLF_VALORATUAL, "Standard")
  TxtMulta.SetText Format(XLF_MULTA, "Standard")
  TxtJuros.SetText Format(XLF_JUROS, "Standard")
  TxtTotal.SetText Format(XLF_VALORTOTAL, "Standard")
  TxtObservacao.SetText FunObservacao(XLO_CAMPOS, XLD_DATABASE)
    
  '------------------------------- Formatar Cor ----------------------------------------
  If XLF_VALORATUAL = 0 Then
    FldTitulo.TextColor = vbRed
    FldProrrogacao.TextColor = vbRed
    TxtValorAtual.TextColor = vbRed
    TxtMulta.TextColor = vbRed
    TxtJuros.TextColor = vbRed
    TxtTotal.TextColor = vbRed
    TxtObservacao.TextColor = vbRed
  End If
  
  Set XLO_VETOR = Nothing
  
End Sub

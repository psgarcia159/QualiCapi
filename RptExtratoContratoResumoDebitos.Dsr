VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptExtratoContratoResumoDebitos 
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12105
   OleObjectBlob   =   "RptExtratoContratoResumoDebitos.dsx":0000
End
Attribute VB_Name = "RptExtratoContratoResumoDebitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Dim VFV_SOMATORIO As New XArray ' Vetor que armazena os somatórios
Dim VFV_SOMATORIO2 As New XArray ' Vetor que armazena os somatórios pelo segundo agrupamento
Dim VFV_VALORES As New XArray  ' Vetor que armazena os valores que precisarão fazer somatório
Dim XFB_IMPRIMIU As Boolean 'Indica se chegou no final do relatório na primeira vez`

Dim XFT_TITULOANTERIOR As String

Private Sub Report_Initialize()

AddReportVariable crRVNumber, "SOMASALDODEVEDOR"
AddReportVariable crRVNumber, "SOMAVENCIDOS"
AddReportVariable crRVNumber, "SOMAVENCIDOS360"
AddReportVariable crRVNumber, "SOMAVENCIDOS361"

AddReportVariable crRVNumber, "TOTALSALDODEVEDOR"
AddReportVariable crRVNumber, "TOTALVENCIDOS"
AddReportVariable crRVNumber, "TOTALVENCIDOS360"
AddReportVariable crRVNumber, "TOTALVENCIDOS361"

VFV_SOMATORIO.ReDim 0, 0, 0, 4  'colunas = quantidade de campos para soma
VFV_SOMATORIO2.ReDim 0, 0, 0, 4  'colunas = quantidade de campos para soma
VFV_VALORES.ReDim 0, 0, 0, 3 'colunas = quantidade de campos para soma -1

End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SectHeaderReport_Format(ByVal pFormattingInfo As Object)
  
  SetReportVariableValue "SOMASALDODEVEDOR", 0
  SetReportVariableValue "SOMAVENCIDOS", 0
  SetReportVariableValue "SOMAVENCIDOS360", 0
  SetReportVariableValue "SOMAVENCIDOS361", 0
  
  SetReportVariableValue "TOTALSALDODEVEDOR", 0
  SetReportVariableValue "TOTALVENCIDOS", 0
  SetReportVariableValue "TOTALVENCIDOS360", 0
  SetReportVariableValue "TOTALVENCIDOS361", 0
  
End Sub


Private Sub SecDetalhe_Format(ByVal pFormattingInfo As Object)
  
  Dim XLF_VALORCORRIGIDO As Double  'Valor corrigido monetariamente
  Dim XLF_DESAGIO As Double  'Valor do deságio
  Dim XLF_MULTA As Double  'Valor da Multa
  Dim XLF_JUROS As Double  'Valor do Juros
  
  Dim XLO_VETOR As New XArray
  
  Dim XLF_SALDODEVEDOR As Double 'Valor Previsto
  Dim XLF_VENCIDOS As Double 'Valor Previsto
  Dim XLF_VENCIDOS360 As Double 'Valor Previsto
  Dim XLF_VENCIDOS361 As Double 'Valor Previsto
     
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
  SubCotacao XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
  
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
      
  XLF_SALDODEVEDOR = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA - XLF_DESAGIO
  If XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value < XLD_DATABASE Then
    XLF_VENCIDOS = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA
  End If
  
  If XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value >= XLD_DATABASE And _
    XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value <= DateAdd("d", 360, XLD_DATABASE) Then
    XLF_VENCIDOS360 = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA - XLF_DESAGIO
  End If
  
  If XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value >= XLD_DATABASE And _
    XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value >= DateAdd("d", 361, XLD_DATABASE) Then
    XLF_VENCIDOS361 = XLF_VALORCORRIGIDO + XLF_JUROS + XLF_MULTA - XLF_DESAGIO
  End If
  
  '------------------------------- Somatorio ----------------------------------------
  
  SetReportVariableValue "SOMASALDODEVEDOR", (FunNuloVal(GetReportVariableValue("SOMASALDODEVEDOR")) + XLF_SALDODEVEDOR)
  SetReportVariableValue "SOMAVENCIDOS", (FunNuloVal(GetReportVariableValue("SOMAVENCIDOS")) + XLF_VENCIDOS)
  SetReportVariableValue "SOMAVENCIDOS360", (FunNuloVal(GetReportVariableValue("SOMAVENCIDOS360")) + XLF_VENCIDOS360)
  SetReportVariableValue "SOMAVENCIDOS361", (FunNuloVal(GetReportVariableValue("SOMAVENCIDOS361")) + XLF_VENCIDOS361)
 
  Set XLO_VETOR = Nothing
  XFT_TITULOANTERIOR = XLO_CAMPOS.GetItemByName("Titulo").Value
  
End Sub

Private Sub SecFooterCliente1_Format(ByVal pFormattingInfo As Object)
 
  Dim XFI_CONT As Integer
  
  'Enquanto não chegou no final do relatorio guarda o soma por moeda no vetor
  If XFB_IMPRIMIU = False Then
    VFV_VALORES(0, 0) = GetReportVariableValue("SOMASALDODEVEDOR")
    VFV_VALORES(0, 1) = GetReportVariableValue("SOMAVENCIDOS")
    VFV_VALORES(0, 2) = GetReportVariableValue("SOMAVENCIDOS360")
    VFV_VALORES(0, 3) = GetReportVariableValue("SOMAVENCIDOS361")
    
    SubSomatorioVetor VFV_SOMATORIO, VFV_VALORES, Left(XFT_TITULOANTERIOR, 12)
  
  'Depois que já passou pelo fim do reltório descarrega o vetor com os somatórios
  Else
    For XFI_CONT = 0 To VFV_SOMATORIO.Count(1) - 1
      If VFV_SOMATORIO(XFI_CONT, 0) = Left(XFT_TITULOANTERIOR, 12) Then
        While XFI_CONT < VFV_SOMATORIO.Count(1)
          If VFV_SOMATORIO(XFI_CONT, 0) <> Left(XFT_TITULOANTERIOR, 12) Then
            XFI_CONT = VFV_SOMATORIO.Count(1) + 1
          Else
            TxtSaldoDevedor.SetText Format(VFV_SOMATORIO(XFI_CONT, 1), "standard")
            TxtVencidos.SetText Format(VFV_SOMATORIO(XFI_CONT, 2), "standard")
            TxtVencidos360.SetText FunTrataZero(Format(VFV_SOMATORIO(XFI_CONT, 3), "standard"))
            TxtVencidos361.SetText FunTrataZero(Format(VFV_SOMATORIO(XFI_CONT, 4), "standard"))
            XFI_CONT = XFI_CONT + 1
          End If
        Wend
      End If
    Next
  End If
  
  SetReportVariableValue "TOTALSALDODEVEDOR", (FunNuloVal(GetReportVariableValue("TOTALSALDODEVEDOR")) + GetReportVariableValue("SOMASALDODEVEDOR"))
  SetReportVariableValue "TOTALVENCIDOS", (FunNuloVal(GetReportVariableValue("TOTALVENCIDOS")) + GetReportVariableValue("SOMAVENCIDOS"))
  SetReportVariableValue "TOTALVENCIDOS360", (FunNuloVal(GetReportVariableValue("TOTALVENCIDOS360")) + GetReportVariableValue("SOMAVENCIDOS360"))
  SetReportVariableValue "TOTALVENCIDOS361", (FunNuloVal(GetReportVariableValue("TOTALVENCIDOS361")) + GetReportVariableValue("SOMAVENCIDOS361"))
  
  SetReportVariableValue "SOMASALDODEVEDOR", 0
  SetReportVariableValue "SOMAVENCIDOS", 0
  SetReportVariableValue "SOMAVENCIDOS360", 0
  SetReportVariableValue "SOMAVENCIDOS361", 0
  
End Sub

Private Sub SecFooterEmpreendimento1_Format(ByVal pFormattingInfo As Object)
  
  Dim XFI_CONT As Integer
  
  'Enquanto não chegou no final do relatorio guarda o soma por moeda no vetor
  If XFB_IMPRIMIU = False Then
    VFV_VALORES(0, 0) = GetReportVariableValue("TOTALSALDODEVEDOR")
    VFV_VALORES(0, 1) = GetReportVariableValue("TOTALVENCIDOS")
    VFV_VALORES(0, 2) = GetReportVariableValue("TOTALVENCIDOS360")
    VFV_VALORES(0, 3) = GetReportVariableValue("TOTALVENCIDOS361")
    
    SubSomatorioVetor VFV_SOMATORIO2, VFV_VALORES, Left(XFT_TITULOANTERIOR, 12)
  
  'Depois que já passou pelo fim do reltório descarrega o vetor com os somatórios
  Else
    For XFI_CONT = 0 To VFV_SOMATORIO2.Count(1) - 1
      If VFV_SOMATORIO2(XFI_CONT, 0) = Left(XFT_TITULOANTERIOR, 12) Then
        While XFI_CONT < VFV_SOMATORIO2.Count(1)
          If VFV_SOMATORIO2(XFI_CONT, 0) <> Left(XFT_TITULOANTERIOR, 12) Then
            XFI_CONT = VFV_SOMATORIO2.Count(1) + 1
          Else
            TxtSomaSaldoDevedor.SetText Format(VFV_SOMATORIO2(XFI_CONT, 1), "standard")
            TxtSomaVencidos.SetText Format(VFV_SOMATORIO2(XFI_CONT, 2), "standard")
            TxtSomaVencidos360.SetText FunTrataZero(Format(VFV_SOMATORIO2(XFI_CONT, 3), "standard"))
            TxtSomaVencidos361.SetText FunTrataZero(Format(VFV_SOMATORIO2(XFI_CONT, 4), "standard"))
            XFI_CONT = XFI_CONT + 1
          End If
        Wend
      End If
    Next
  End If
  
  SetReportVariableValue "TOTALSALDODEVEDOR", 0
  SetReportVariableValue "TOTALVENCIDOS", 0
  SetReportVariableValue "TOTALVENCIDOS360", 0
  SetReportVariableValue "TOTALVENCIDOS361", 0
  
End Sub

Private Sub SecFooterReport_Format(ByVal pFormattingInfo As Object)
  
  XFB_IMPRIMIU = True
  
End Sub

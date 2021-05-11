VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptMapaContabil 
   ClientHeight    =   11925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14250
   OleObjectBlob   =   "RptMapaContabil.dsx":0000
End
Attribute VB_Name = "RptMapaContabil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Dim XLT_DISTRATO As Boolean

Private Sub Report_Initialize()
  
  AddReportVariable crRVNumber, "SALDOANTERIOR"
  AddReportVariable crRVNumber, "VENDASMES"
  AddReportVariable crRVNumber, "PAGAMENTOSMES"
  AddReportVariable crRVNumber, "DISTRATOSMES"
  AddReportVariable crRVNumber, "CORRECAOMONETARIA"
  AddReportVariable crRVNumber, "SALDOFINAL"
  
  AddReportVariable crRVNumber, "SOMASALDOANTERIOR"
  AddReportVariable crRVNumber, "SOMAVENDASMES"
  AddReportVariable crRVNumber, "SOMAPAGAMENTOSMES"
  AddReportVariable crRVNumber, "SOMADISTRATOSMES"
  AddReportVariable crRVNumber, "SOMACORRECAOMONETARIA"
  AddReportVariable crRVNumber, "SOMASALDOFINAL"
   
End Sub

Private Sub SectHeaderReport_Format(ByVal pFormattingInfo As Object)
    
  'Zera as variáveis quando passa pelo Header do relatório
  SetReportVariableValue "SALDOANTERIOR", 0
  SetReportVariableValue "VENDASMES", 0
  SetReportVariableValue "PAGAMENTOSMES", 0
  SetReportVariableValue "DISTRATOSMES", 0
  SetReportVariableValue "CORRECAOMONETARIA", 0
  SetReportVariableValue "SALDOFINAL", 0
  
  SetReportVariableValue "SOMASALDOANTERIOR", 0
  SetReportVariableValue "SOMAVENDASMES", 0
  SetReportVariableValue "SOMAPAGAMENTOSMES", 0
  SetReportVariableValue "SOMADISTRATOSMES", 0
  SetReportVariableValue "SOMACORRECAOMONETARIA", 0
  SetReportVariableValue "SOMASALDOFINAL", 0
  
End Sub

Private Sub Report_Terminate()
  
  Unload Me
  
End Sub

Private Sub SecDetalhe_Format(ByVal pFormattingInfo As Object)
  
    Dim XLF_VALORCORRIGIDO As Double  'Valor corrigido monetariamente
    Dim XLF_SALDO As Double
    Dim XLF_SALDOANTERIOR As Double
    Dim XLF_CORRECAOMONETARIA As Double
    Dim XLF_VENDAMES As Double
    Dim XLF_DISTRATOMES As Double
    Dim XLF_PAGAMENTOMES As Double
    
    Dim XLF_VALORPAGOORIGINAL As Double
    Dim XLF_VALORPAGOINDEXADO As Double 'Valor pago indexado
    Dim XLF_VALORPAGOATUALIZADO As Double 'Valor pago atualizado na correção da data base
    Dim XLF_VALORPAGOATUALIZADOANTERIOR As Double 'Valor pago atualizado na correção da data base menos um mês
      
    Dim XLF_VALORTITULO As Double 'Valor do Titulo
    Dim XLF_VALORPREVISTO As Double 'Valor Previsto
    Dim XLF_VALORORIGINAL As Double 'Valor Original
    Dim XLF_VALORPAGO As Double 'Valor pago
    
    Dim XLF_VALORPAGOINDEXADO2 As Double 'Valor pago indexado
    Dim XLF_VALORPAGOATUALIZADO2 As Double 'Valor pago atualizado na correção da data base
    
    'Vetor auxiliar para armazenar o valor do titulo, a cotação e a cotação do mês anterior, tudo de acordo com o indexador escolhido. Também indica se tem indexador ou não.
    Dim XLO_VETOR As New XArray
    
    'As cotaçãoe armazenadas correspondem a de um mês a menos que o de XLO_VETOR
    Dim XLO_VETORANTERIOR As New XArray
    Dim XLD_DATABASE As Date  'Data base passada por parâmetro para relatório
    Dim XLI_INDEXADOR As Byte  ' Indexador passado por parâmetro para o relatório
    Dim XLO_CAMPOS As CRAXDDRT.DatabaseFieldDefinitions  ' Objeto que representa os campos do relatório
    Dim XLI_MOEDA As Integer 'Código da moeda de acordo com o indexador escolhido
  
    Set XLO_CAMPOS = Me.Database.Tables(1).Fields
    
    XLI_INDEXADOR = CByte(FormulaFields.GetItemByName("Indexador").Value)
    
    'Pega o código da moeda e a sua descrição
    SubUnidadeMonetaria XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, True
  
    XLI_MOEDA = XLO_VETOR(1, 0)
    
    'Pega as cotações do mês da data base e do mês anterior
    SubCotacao XLO_CAMPOS, XLI_INDEXADOR, XLO_VETOR, False
    SubCotacaoMapaContabil XLO_CAMPOS, XLI_INDEXADOR, XLO_VETORANTERIOR, False
    
    'se não foi pago
    If XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value = "" Then
      
        XLD_DATABASE = CDate(FormulaFields.GetItemByName("FiltroDataBase").Value)
        XLF_VALORCORRIGIDO = Format(FunCalculoCorrecaoMonetaria(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLI_MOEDA), "Standard")
        
        'Juros da CEF
        XLF_VALORCORRIGIDO = XLF_VALORCORRIGIDO + Format(FunCalculoJurosCEF(XLO_CAMPOS, XLD_DATABASE, XLF_VALORCORRIGIDO), "standard")
        
        XLF_SALDO = XLF_VALORCORRIGIDO
      
        If XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value < XLD_DATABASE Then
            XLD_DATABASE = DateAdd("m", -1, XLD_DATABASE)
            XLF_VALORCORRIGIDO = Format(FunCalculoCorrecaoMonetaria(XLO_CAMPOS, XLO_VETORANTERIOR, XLD_DATABASE, XLI_MOEDA), "Standard")
            XLF_SALDOANTERIOR = XLF_VALORCORRIGIDO
            XLF_CORRECAOMONETARIA = XLF_SALDO - XLF_SALDOANTERIOR
        ElseIf Month(XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value) = Month(XLD_DATABASE) And Year(XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value) = Year(XLD_DATABASE) Then
            'Inserido em 04/12/2008, Larissa
            'A condição foi inserida p/ que o relatório não incluísse na soma os resíduos
            If (Right(XLO_CAMPOS.GetItemByName("Titulo").Value, 2) = "00") Then
                XLF_VENDAMES = XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value + XLF_VENDAMES
            End If
        End If
       
    'se foi pago
    Else
        XLF_VALORPAGOORIGINAL = Format(FunCalculoValorOriginalPago(XLO_CAMPOS), "standard")
        XLD_DATABASE = XLO_CAMPOS.GetItemByName("titu_dt_BasePagto").Value
          
        'Pegar a cotação da moeda3 ignorando se for indexador zero
        SubCotacao3 XLO_CAMPOS, XLO_VETOR
        
        XLF_VALORPAGOINDEXADO = FunCalculoValorIndexado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORPAGOORIGINAL)
'        XLF_VALORPAGOINDEXADO = XLO_CAMPOS.GetItemByName("titu_vl_ParcelaIndex1").Value
           
        XLD_DATABASE = CDate(FormulaFields.GetItemByName("FiltroDataBase").Value)
        
        If Month(XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value) = Month(XLD_DATABASE) And Year(XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value) = Year(XLD_DATABASE) Then
            XLF_PAGAMENTOMES = XLF_VALORPAGOORIGINAL
        End If
        
        'Pega a cotação 1 ou 2
        SubCotacao2 XLO_CAMPOS, XLO_VETOR
        XLF_VALORPAGOATUALIZADO = Format(FunCalculoValorAtualizado(XLO_CAMPOS, XLO_VETOR, XLO_CAMPOS.GetItemByName("titu_dt_BasePagto").Value, XLF_VALORPAGOINDEXADO, XLI_MOEDA), "standard")
        XLF_SALDO = XLF_VALORPAGOATUALIZADO
    
        If XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value < XLD_DATABASE Then
           
            XLD_DATABASE = DateAdd("m", -1, XLO_CAMPOS.GetItemByName("titu_dt_BasePagto").Value)
            
            'Pega a cotação para o mês anterior
            SubCotacao2MapaContabil XLO_CAMPOS, XLO_VETOR
            XLF_VALORPAGOATUALIZADOANTERIOR = Format(FunCalculoValorAtualizado(XLO_CAMPOS, XLO_VETOR, XLD_DATABASE, XLF_VALORPAGOINDEXADO, XLI_MOEDA), "standard")
            
            XLF_SALDOANTERIOR = XLF_VALORPAGOATUALIZADOANTERIOR
            
            XLF_CORRECAOMONETARIA = XLF_SALDO - XLF_SALDOANTERIOR
          
        ElseIf Month(XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value) = Month(XLD_DATABASE) And Year(XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value) = Year(XLD_DATABASE) Then
            'Inserido em 04/12/2008, Larissa
            'A condição foi inserida p/ que o relatório não incluísse na soma os resíduos
            If (Right(XLO_CAMPOS.GetItemByName("Titulo").Value, 2) = "00") Then
                XLF_VENDAMES = XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value + XLF_VENDAMES
            End If
        End If
    
    End If
  
    If Month(XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value) = Month(CDate(FormulaFields.GetItemByName("FiltroDataBase").Value)) And Year(XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value) = Year(CDate(FormulaFields.GetItemByName("FiltroDataBase").Value)) Then
        XLF_CORRECAOMONETARIA = XLF_SALDO - XLF_VENDAMES
    End If

  
    '------------------------------- Somatorio ----------------------------------------
    
    SetReportVariableValue "SALDOANTERIOR", (FunNuloVal(GetReportVariableValue("SALDOANTERIOR")) + XLF_SALDOANTERIOR)
    SetReportVariableValue "CORRECAOMONETARIA", (FunNuloVal(GetReportVariableValue("CORRECAOMONETARIA")) + XLF_CORRECAOMONETARIA)
    SetReportVariableValue "VENDASMES", (FunNuloVal(GetReportVariableValue("VENDASMES")) + XLF_VENDAMES)
    SetReportVariableValue "PAGAMENTOSMES", (FunNuloVal(GetReportVariableValue("PAGAMENTOSMES")) + XLF_PAGAMENTOMES)
    
    'Quando tem distrato no mês
    XLD_DATABASE = CDate(FormulaFields.GetItemByName("FiltroDataBase").Value)
    If XLO_CAMPOS.GetItemByName("cont_dt_Distrato").Value <> "" Then
        If Month(XLO_CAMPOS.GetItemByName("cont_dt_Distrato").Value) = Month(XLD_DATABASE) And Year(XLO_CAMPOS.GetItemByName("cont_dt_Distrato").Value) = Year(XLD_DATABASE) Then
            XLT_DISTRATO = True
        Else
            XLT_DISTRATO = False
        End If
    Else
        XLT_DISTRATO = False
    End If
       
    Set XLO_VETOR = Nothing
    Set XLO_VETORANTERIOR = Nothing
  
End Sub

Private Sub SecFooterContrato_Format(ByVal pFormattingInfo As Object)
         
    'Quando tem distrato no mês
    If XLT_DISTRATO = True Then
        SetReportVariableValue "CORRECAOMONETARIA", 0
        SetReportVariableValue "DISTRATOSMES", (FunNuloVal(GetReportVariableValue("DISTRATOSMES")) + GetReportVariableValue("SALDOANTERIOR") - FunNuloVal(GetReportVariableValue("PAGAMENTOSMES")))
    End If
         
    SetReportVariableValue "SALDOFINAL", GetReportVariableValue("SALDOANTERIOR") + GetReportVariableValue("VENDASMES") _
      - GetReportVariableValue("PAGAMENTOSMES") - GetReportVariableValue("DISTRATOSMES") _
      + GetReportVariableValue("CORRECAOMONETARIA")
  
    TxtSaldoAnterior.SetText Format(GetReportVariableValue("SALDOANTERIOR"), "standard")
    TxtVendasMes.SetText Format(GetReportVariableValue("VENDASMES"), "standard")
    TxtPagamentosMes.SetText Format(GetReportVariableValue("PAGAMENTOSMES"), "standard")
    TxtDistratosMes.SetText Format(GetReportVariableValue("DISTRATOSMES"), "standard")
    TxtCorrecaoMonetaria.SetText Format(GetReportVariableValue("CORRECAOMONETARIA"), "standard")
    TxtSaldoFinal2.SetText Format(GetReportVariableValue("SALDOFINAL"), "standard")
  
    SetReportVariableValue "SOMASALDOANTERIOR", GetReportVariableValue("SOMASALDOANTERIOR") + GetReportVariableValue("SALDOANTERIOR")
    SetReportVariableValue "SOMAVENDASMES", GetReportVariableValue("SOMAVENDASMES") + GetReportVariableValue("VENDASMES")
    SetReportVariableValue "SOMAPAGAMENTOSMES", GetReportVariableValue("SOMAPAGAMENTOSMES") + GetReportVariableValue("PAGAMENTOSMES")
    SetReportVariableValue "SOMADISTRATOSMES", GetReportVariableValue("SOMADISTRATOSMES") + GetReportVariableValue("DISTRATOSMES")
    SetReportVariableValue "SOMACORRECAOMONETARIA", GetReportVariableValue("SOMACORRECAOMONETARIA") + GetReportVariableValue("CORRECAOMONETARIA")
    SetReportVariableValue "SOMASALDOFINAL", GetReportVariableValue("SOMASALDOFINAL") + GetReportVariableValue("SALDOFINAL")
    
    'quando passa pelo footer zera as variáveis
    SetReportVariableValue "SALDOANTERIOR", 0
    SetReportVariableValue "VENDASMES", 0
    SetReportVariableValue "PAGAMENTOSMES", 0
    SetReportVariableValue "DISTRATOSMES", 0
    SetReportVariableValue "CORRECAOMONETARIA", 0
    SetReportVariableValue "SALDOFINAL", 0
    
End Sub

Private Sub SecFooterEmpreendimento1_Format(ByVal pFormattingInfo As Object)

  TxtSomaSaldoAnterior.SetText Format(GetReportVariableValue("SOMASALDOANTERIOR"), "standard")
  TxtSomaVendasMes.SetText Format(GetReportVariableValue("SOMAVENDASMES"), "standard")
  TxtSomaPagamentosMes.SetText Format(GetReportVariableValue("SOMAPAGAMENTOSMES"), "standard")
  TxtSomaDistratosMes.SetText Format(GetReportVariableValue("SOMADISTRATOSMES"), "standard")
  TxtSomaCorrecaoMonetaria.SetText Format(GetReportVariableValue("SOMACORRECAOMONETARIA"), "standard")
  TxtSomaSaldoFinal.SetText Format(GetReportVariableValue("SOMASALDOFINAL"), "standard")
  
  SetReportVariableValue "SOMASALDOANTERIOR", 0
  SetReportVariableValue "SOMAVENDASMES", 0
  SetReportVariableValue "SOMAPAGAMENTOSMES", 0
  SetReportVariableValue "SOMADISTRATOSMES", 0
  SetReportVariableValue "SOMACORRECAOMONETARIA", 0
  SetReportVariableValue "SOMASALDOFINAL", 0
  
End Sub

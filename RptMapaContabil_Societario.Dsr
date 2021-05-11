VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} RptMapaContabil_Societario 
   ClientHeight    =   10065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15015
   OleObjectBlob   =   "RptMapaContabil_Societario.dsx":0000
End
Attribute VB_Name = "RptMapaContabil_Societario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit

Private Sub Report_Initialize()
  
    'É o valor histórico dos títulos de contratos com data de venda menor que a data base informada.
    '(é o saldo devedor inicial do contrato)
    AddReportVariable crRVNumber, "SALDOANTERIOR"
    
    'É o valor histórico dos títulos de contratos com data de venda no mês/ano da data base informada.
    AddReportVariable crRVNumber, "VENDASMES"
    
    'Se o contrato teve distrato, o valor é o saldo anterior mais o distrato no mês
    AddReportVariable crRVNumber, "DISTRATOSMES"
    
    'Valor pago do título, sem correção, juros, multas ou deságio
    AddReportVariable crRVNumber, "VALORPAGOLIQUIDO"
    
    'É o valor pago dos títulos com a correção monetária, sem juros, multas e deságio.
    AddReportVariable crRVNumber, "VALORPAGOCORRIGIDO"
        
    'É o valor corrigido do título menos seu valor histórico
    AddReportVariable crRVNumber, "CORRECAOMONETARIA"
        
    'Percentual de juros mora de cada contrato
    AddReportVariable crRVNumber, "JUROSCONTRATO"
    
    'Valor dos juros pago do título (se houver)
    AddReportVariable crRVNumber, "JUROS"
    
    'Valor das multas paga do título (se houver)
    AddReportVariable crRVNumber, "MULTAS"
    
    'Valor do deságio do título (se houver)
    AddReportVariable crRVNumber, "DESAGIO"
    
    'É o saldo anterior menos o valor histórico pago (é o saldo devedor final do contrato)
    AddReportVariable crRVNumber, "SALDOFINAL"
    
        
    'Total do saldo anterior para um contrato
    AddReportVariable crRVNumber, "SOMASALDOANTERIOR"
    
    'Total de vendas no mês para um contrato
    AddReportVariable crRVNumber, "SOMAVENDASMES"
    
    'Total do valor líquido pago para um contrato
    AddReportVariable crRVNumber, "SOMAVALORPAGOLIQUIDO"
    
    'Total de juros pago para um contrato
    AddReportVariable crRVNumber, "SOMAJUROS"
    
    'Total de multas paga para um contrato
    AddReportVariable crRVNumber, "SOMAMULTAS"
    
    'Total de deságio concedido para um contrato
    AddReportVariable crRVNumber, "SOMADESAGIO"
    
    'Total de correção paga para um contrato
    AddReportVariable crRVNumber, "SOMACORRECAOMONETARIA"
    
    'Total pago no mês incluindo correção para um contrato
    AddReportVariable crRVNumber, "SOMAVALORPAGOCORRIGIDO"
    
    'Total do distrato do mês para um contrato
    AddReportVariable crRVNumber, "SOMADISTRATOSMES"
    
    'Total devedor para um contrato
    AddReportVariable crRVNumber, "SOMASALDOFINAL"
   
End Sub

Private Sub SectHeaderReport_Format(ByVal pFormattingInfo As Object)
    
    'Zera as variáveis quando passa pelo Header do relatório
    SetReportVariableValue "SALDOANTERIOR", 0
    SetReportVariableValue "VENDASMES", 0
    SetReportVariableValue "VALORPAGOCORRIGIDO", 0
    SetReportVariableValue "DISTRATOSMES", 0
    SetReportVariableValue "CORRECAOMONETARIA", 0
    SetReportVariableValue "SALDOFINAL", 0
    SetReportVariableValue "JUROSCONTRATO", 0
    SetReportVariableValue "JUROS", 0
    SetReportVariableValue "MULTAS", 0
    SetReportVariableValue "DESAGIO", 0
    
    SetReportVariableValue "SOMASALDOANTERIOR", 0
    SetReportVariableValue "SOMAVENDASMES", 0
    SetReportVariableValue "SOMAVALORPAGOLIQUIDO", 0
    SetReportVariableValue "SOMAJUROS", 0
    SetReportVariableValue "SOMAMULTAS", 0
    SetReportVariableValue "SOMADESAGIO", 0
    SetReportVariableValue "SOMAVALORPAGOCORRIGIDO", 0
    SetReportVariableValue "SOMADISTRATOSMES", 0
    SetReportVariableValue "SOMACORRECAOMONETARIA", 0
    SetReportVariableValue "SOMASALDOFINAL", 0
      
End Sub

Private Sub Report_Terminate()
    Unload Me
End Sub

Private Sub SecDetalhe_Format(ByVal pFormattingInfo As Object)
  
    Dim XLF_VALORCORRIGIDO As Double        'Valor corrigido do título pago
    Dim XLF_SALDOANTERIOR As Double         'Valor histórico do título em aberto ou pago em data maior ou igual a data base informada. Compõe o saldo devedor inicial do contrato.
    Dim XLF_CORRECAOMONETARIA As Double     'Valor da correção monetária do título: é o valor corrigido menos o valor histórico
    Dim XLF_VENDAMES As Double              'Valor histórico dos títulos de contratos cuja data de venda é no mês/ano da data base informada.
    Dim XLF_VALORPAGOCORRIGIDO As Double    'Valor pago do título com correção, sem juros, multas ou deságio.
    Dim XLF_VALORPAGOLIQUIDO As Double      'Valor pago do título sem correção, juros, multas ou deságio.
    Dim XLF_VALORJUROSCONTRATO As Double    'Valor dos juros do contrato (informado durante o cadastro do contrato).
    Dim XLF_JUROS As Double                 'Valor dos juros do título pago
    Dim XLF_MULTAS As Double                'Valor da multa do título pago
    Dim XLF_DESAGIO As Double               'Valor do deságio concedido ao título pago
    Dim XLD_DATABASE As Date                'Data base passada por parâmetro para relatório
    Dim XLO_Contrato As String              'Número do contrato: xxxx.xxxx.xx (empreendimento.imovel.contrato).
    Dim XLO_CAMPOS As CRAXDDRT.DatabaseFieldDefinitions  ' Objeto que representa os campos do relatório
    
    Set XLO_CAMPOS = Me.Database.Tables(1).Fields
    
    'Pega o percentual dos juros contratuais
    XLO_Contrato = ""
    If XLO_Contrato <> XLO_CAMPOS.GetItemByName("Contrato").Value Then '22/12/09
        XLO_Contrato = XLO_CAMPOS.GetItemByName("Contrato").Value
        XLF_VALORJUROSCONTRATO = XLO_CAMPOS.GetItemByName("titu_vl_PercJurosFin").Value
    End If
        
    'Pega a data base informada pelo usuário
    XLD_DATABASE = CDate(FormulaFields.GetItemByName("FiltroDataBase").Value)
    
    'Se o título não estiver pago
    If XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value = "" Then
        If XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value < XLD_DATABASE Then
            XLF_SALDOANTERIOR = Format(XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value, "standard")
        ElseIf Month(XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value) = Month(XLD_DATABASE) And Year(XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value) = Year(XLD_DATABASE) Then
            'Condição inserida p/ que o relatório não inclua na soma os resíduos
            If (Right(XLO_CAMPOS.GetItemByName("Titulo").Value, 2) = "00") Then
                XLF_VENDAMES = Format(XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value, "standard")
            End If
        End If
    'Se o título estiver pago
    Else
        'Só inclui no cálculo os títulos pagos no mês/ano da data base informada
        If Month(XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value) = Month(XLD_DATABASE) And Year(XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value) = Year(XLD_DATABASE) Then
        
           'Valor total pago no mês (incluindo apenas a correção)
            XLF_VALORPAGOCORRIGIDO = Format(FunCalculoValorOriginalPago(XLO_CAMPOS), "standard")
            
            'Valor pago corrigido (sem juros, multas e deságio)
            If XLO_CAMPOS.GetItemByName("titu_tx_IndicePagamento").Value = 1 Then
                XLF_VALORCORRIGIDO = Format(XLO_CAMPOS.GetItemByName("titu_vl_Corrigido1").Value, "standard")
            ElseIf XLO_CAMPOS.GetItemByName("titu_tx_IndicePagamento").Value = 2 Then
                XLF_VALORCORRIGIDO = Format(XLO_CAMPOS.GetItemByName("titu_vl_Corrigido2").Value, "standard")
            End If
                        
            'Valor da correção monetária: é o valor pago corrigido menos o valor histórico do título
             XLF_CORRECAOMONETARIA = XLF_VALORCORRIGIDO - XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value
                        
            'Valor pago líquido do título (valor pago sem juros, multas, deságio e correção)
            'Não é necessariamente o valor histórico do título por causa dos casos em que são gerados resíduos
            XLF_VALORPAGOLIQUIDO = XLF_VALORPAGOCORRIGIDO - XLF_CORRECAOMONETARIA
            
            'valor dos juros pagos do título
            XLF_JUROS = Format(XLO_CAMPOS.GetItemByName("titu_vl_Juros").Value, "standard")
            
            'Valor da multa paga do título
            XLF_MULTAS = Format(XLO_CAMPOS.GetItemByName("titu_vl_Multa").Value, "standard")
            
            'Valor do deságio do título
            XLF_DESAGIO = Format(XLO_CAMPOS.GetItemByName("titu_vl_Desagio").Value, "standard")
            
        End If
                        
        If XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value < XLD_DATABASE Then
            XLF_SALDOANTERIOR = Format(XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value, "standard")
        ElseIf Month(XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value) = Month(XLD_DATABASE) And Year(XLO_CAMPOS.GetItemByName("cont_dt_Venda").Value) = Year(XLD_DATABASE) Then
            'Condição inserida p/ que o relatório não inclua na soma os resíduos
            If (Right(XLO_CAMPOS.GetItemByName("Titulo").Value, 2) = "00") Then
                XLF_VENDAMES = Format(XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value, "standard")
            End If
        End If
    End If
    
    
    '------------------------------- Somatorio ----------------------------------------
    
    SetReportVariableValue "SALDOANTERIOR", (FunNuloVal(GetReportVariableValue("SALDOANTERIOR")) + XLF_SALDOANTERIOR)
    SetReportVariableValue "CORRECAOMONETARIA", (FunNuloVal(GetReportVariableValue("CORRECAOMONETARIA")) + XLF_CORRECAOMONETARIA)
    SetReportVariableValue "VENDASMES", (FunNuloVal(GetReportVariableValue("VENDASMES")) + XLF_VENDAMES)
    SetReportVariableValue "VALORPAGOCORRIGIDO", (FunNuloVal(GetReportVariableValue("VALORPAGOCORRIGIDO")) + XLF_VALORPAGOCORRIGIDO)
    SetReportVariableValue "JUROSCONTRATO", (XLF_VALORJUROSCONTRATO)
    SetReportVariableValue "JUROS", (FunNuloVal(GetReportVariableValue("JUROS")) + XLF_JUROS)
    SetReportVariableValue "MULTAS", (FunNuloVal(GetReportVariableValue("MULTAS")) + XLF_MULTAS)
    SetReportVariableValue "DESAGIO", (FunNuloVal(GetReportVariableValue("DESAGIO")) + XLF_DESAGIO)
    SetReportVariableValue "VALORPAGOLIQUIDO", (FunNuloVal(GetReportVariableValue("VALORPAGOLIQUIDO")) + XLF_VALORPAGOLIQUIDO)
   
    
    'Quando tem distrato no mês
    XLD_DATABASE = CDate(FormulaFields.GetItemByName("FiltroDataBase").Value)
    If XLO_CAMPOS.GetItemByName("cont_dt_Distrato").Value <> "" Then
        If Month(XLO_CAMPOS.GetItemByName("cont_dt_Distrato").Value) = Month(XLD_DATABASE) And _
            Year(XLO_CAMPOS.GetItemByName("cont_dt_Distrato").Value) = Year(XLD_DATABASE) Then
            SetReportVariableValue "CORRECAOMONETARIA", 0
            SetReportVariableValue "VALORPAGOCORRIGIDO", 0
            SetReportVariableValue "JUROS", 0
            SetReportVariableValue "MULTAS", 0
            SetReportVariableValue "DESAGIO", 0
            SetReportVariableValue "VALORPAGOLIQUIDO", 0
            SetReportVariableValue "DISTRATOSMES", (FunNuloVal(GetReportVariableValue("DISTRATOSMES")) + XLF_SALDOANTERIOR)
        End If
    End If

End Sub

Private Sub SecFooterContrato_Format(ByVal pFormattingInfo As Object)
     
    SetReportVariableValue "SALDOFINAL", GetReportVariableValue("SALDOANTERIOR") + GetReportVariableValue("VENDASMES") _
      - GetReportVariableValue("VALORPAGOCORRIGIDO") + GetReportVariableValue("CORRECAOMONETARIA") - GetReportVariableValue("DISTRATOSMES") '21/12/09
    
    TxtSaldoAnterior.SetText Format(GetReportVariableValue("SALDOANTERIOR"), "standard")
    TxtVendasMes.SetText Format(GetReportVariableValue("VENDASMES"), "standard")
    TxtValorPagoLiquido.SetText Format(GetReportVariableValue("VALORPAGOLIQUIDO"), "standard")
    TxtJuros.SetText Format(GetReportVariableValue("JUROS"), "standard")
    TxtMultas.SetText Format(GetReportVariableValue("MULTAS"), "standard")
    TxtDesagio.SetText Format(GetReportVariableValue("DESAGIO"), "standard")
    TxtCorrecao.SetText Format(GetReportVariableValue("CORRECAOMONETARIA"), "standard")
    TxtValorPagoCorrigido.SetText Format(GetReportVariableValue("VALORPAGOCORRIGIDO"), "standard")
    TxtDistratosMes.SetText Format(GetReportVariableValue("DISTRATOSMES"), "standard")
    TxtJurosContratuais.SetText Format(GetReportVariableValue("JUROSCONTRATO"), "standard") '21/12/09
    TxtSaldoFinal2.SetText Format(GetReportVariableValue("SALDOFINAL"), "standard")
            
    SetReportVariableValue "SOMASALDOANTERIOR", GetReportVariableValue("SOMASALDOANTERIOR") + GetReportVariableValue("SALDOANTERIOR")
    SetReportVariableValue "SOMAVENDASMES", GetReportVariableValue("SOMAVENDASMES") + GetReportVariableValue("VENDASMES")
    SetReportVariableValue "SOMAVALORPAGOLIQUIDO", GetReportVariableValue("SOMAVALORPAGOLIQUIDO") + GetReportVariableValue("VALORPAGOLIQUIDO")
    SetReportVariableValue "SOMAJUROS", GetReportVariableValue("SOMAJUROS") + GetReportVariableValue("JUROS")
    SetReportVariableValue "SOMAMULTAS", GetReportVariableValue("SOMAMULTAS") + GetReportVariableValue("MULTAS")
    SetReportVariableValue "SOMADESAGIO", GetReportVariableValue("SOMADESAGIO") + GetReportVariableValue("DESAGIO")
    SetReportVariableValue "SOMAVALORPAGOCORRIGIDO", GetReportVariableValue("SOMAVALORPAGOCORRIGIDO") + GetReportVariableValue("VALORPAGOCORRIGIDO")
    SetReportVariableValue "SOMADISTRATOSMES", GetReportVariableValue("SOMADISTRATOSMES") + GetReportVariableValue("DISTRATOSMES")
    SetReportVariableValue "SOMACORRECAOMONETARIA", GetReportVariableValue("SOMACORRECAOMONETARIA") + GetReportVariableValue("CORRECAOMONETARIA")
    SetReportVariableValue "SOMASALDOFINAL", GetReportVariableValue("SOMASALDOFINAL") + GetReportVariableValue("SALDOFINAL")
        
    'quando passa pelo footer zera as variáveis
    SetReportVariableValue "SALDOANTERIOR", 0
    SetReportVariableValue "VENDASMES", 0
    SetReportVariableValue "VALORPAGOLIQUIDO", 0
    SetReportVariableValue "JUROS", 0
    SetReportVariableValue "MULTAS", 0
    SetReportVariableValue "DESAGIO", 0
    SetReportVariableValue "VALORPAGOCORRIGIDO", 0
    SetReportVariableValue "DISTRATOSMES", 0
    SetReportVariableValue "CORRECAOMONETARIA", 0
    SetReportVariableValue "SALDOFINAL", 0
    
End Sub

Private Sub SecFooterEmpreendimento1_Format(ByVal pFormattingInfo As Object)

    TxtSomaSaldoAnterior.SetText Format(GetReportVariableValue("SOMASALDOANTERIOR"), "standard")
    TxtSomaVendasMes.SetText Format(GetReportVariableValue("SOMAVENDASMES"), "standard")
    TxtSomaValorPagoLiquido.SetText Format(GetReportVariableValue("SOMAVALORPAGOLIQUIDO"), "standard")
    TxtSomaJuros.SetText Format(GetReportVariableValue("SOMAJUROS"), "standard")
    TxtSomaMultas.SetText Format(GetReportVariableValue("SOMAMULTAS"), "standard")
    TxtSomaDesagio.SetText Format(GetReportVariableValue("SOMADESAGIO"), "standard")
    TxtSomaCorrecao.SetText Format(GetReportVariableValue("SOMACORRECAOMONETARIA"), "standard")
    TxtSomaValorPagoCorrigido.SetText Format(GetReportVariableValue("SOMAVALORPAGOCORRIGIDO"), "standard")
    TxtSomaDistratosMes.SetText Format(GetReportVariableValue("SOMADISTRATOSMES"), "standard")
    TxtSomaSaldoFinal.SetText Format(GetReportVariableValue("SOMASALDOFINAL"), "standard")
    
    SetReportVariableValue "SOMASALDOANTERIOR", 0
    SetReportVariableValue "SOMAVENDASMES", 0
    SetReportVariableValue "SOMAJUROS", 0
    SetReportVariableValue "SOMAMULTAS", 0
    SetReportVariableValue "SOMADESAGIO", 0
    SetReportVariableValue "SOMAVALORPAGOCORRIGIDO", 0
    SetReportVariableValue "SOMADISTRATOSMES", 0
    SetReportVariableValue "SOMACORRECAOMONETARIA", 0
    SetReportVariableValue "SOMASALDOFINAL", 0
  
End Sub

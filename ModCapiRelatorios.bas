Attribute VB_Name = "ModCapiRelatorios"
Option Explicit

Sub SubSomatorioVetor(XLV_SOMA As XArray, XLV_VALORES As XArray, XLT_CHAVE As String)
  
  Dim XLI_CONT1 As Integer 'Contador da chave
  Dim XLI_CONT2 As Integer 'Contador da chave
   
  XLI_CONT1 = XLV_SOMA.Count(1) - 1
  XLV_SOMA(XLI_CONT1, 0) = XLT_CHAVE
  For XLI_CONT2 = 0 To XLV_VALORES.Count(2) - 1
    XLV_SOMA(XLI_CONT1, XLI_CONT2 + 1) = XLV_VALORES(0, XLI_CONT2)
  Next
  XLV_SOMA.ReDim 0, XLI_CONT1 + 1, 0, XLV_VALORES.Count(2)

End Sub

'Soma o total de cada moeda em cada chave
Function SubSomaMoeda2(XLI_MOEDA As Integer, XLT_MOEDA As String, XLV_TOTALMOEDA As XArray, XLF_VALOR As Double, XLT_CHAVE As String)
  
  Dim XLI_CONT1 As Integer 'Contador da chave
  Dim XLI_CONT2 As Integer 'Contador da moeda
  
  Dim XLB_ACHOU As Boolean
  
  XLI_CONT1 = 0
  XLB_ACHOU = False
  
  If XLV_TOTALMOEDA(0, 0) = "" Then
    XLV_TOTALMOEDA(0, 0) = XLT_MOEDA
    XLV_TOTALMOEDA(0, 1) = XLI_MOEDA
    XLV_TOTALMOEDA(0, 2) = XLF_VALOR
    XLV_TOTALMOEDA(0, 3) = XLT_CHAVE
  Else
    
    For XLI_CONT1 = 0 To XLV_TOTALMOEDA.Count(1) - 1
      If XLV_TOTALMOEDA(XLI_CONT1, 3) = XLT_CHAVE Then
        While (XLV_TOTALMOEDA(XLI_CONT1, 3) = XLT_CHAVE) And (XLI_CONT2 < XLV_TOTALMOEDA.Count(1))
          For XLI_CONT2 = XLI_CONT1 To XLV_TOTALMOEDA.Count(1) - 1
            If XLV_TOTALMOEDA(XLI_CONT2, 1) = XLI_MOEDA Then
              XLV_TOTALMOEDA(XLI_CONT2, 2) = XLV_TOTALMOEDA(XLI_CONT2, 2) + XLF_VALOR
              XLI_CONT2 = XLV_TOTALMOEDA.Count(1)
              XLB_ACHOU = True
            End If
          Next
        Wend
      End If
    Next
    
    If XLI_CONT2 = 0 Then
      XLI_CONT2 = XLI_CONT1
    End If
      
    If XLB_ACHOU = False Then
      XLV_TOTALMOEDA.ReDim 0, XLI_CONT2, 0, 3
      XLV_TOTALMOEDA(XLI_CONT2, 0) = XLT_MOEDA
      XLV_TOTALMOEDA(XLI_CONT2, 1) = XLI_MOEDA
      XLV_TOTALMOEDA(XLI_CONT2, 2) = XLF_VALOR
      XLV_TOTALMOEDA(XLI_CONT2, 3) = XLT_CHAVE
    End If
  
  End If

End Function

'Soma o total de cada moeda
Function SubSomaMoeda(XLI_MOEDA As Integer, XLT_MOEDA As String, XLV_TOTALMOEDA As XArray, XLF_VALOR As Double)
  
  Dim XLI_CONT As Integer
  Dim XLB_ACHOU As Boolean
  
  XLI_CONT = 0
  XLB_ACHOU = False
  
  If XLV_TOTALMOEDA(0, 0) = "" Then
    XLV_TOTALMOEDA(0, 0) = XLT_MOEDA
    XLV_TOTALMOEDA(0, 1) = XLI_MOEDA
    XLV_TOTALMOEDA(0, 2) = XLF_VALOR
  Else
    
    For XLI_CONT = 0 To XLV_TOTALMOEDA.Count(1) - 1
      If XLV_TOTALMOEDA(XLI_CONT, 1) = XLI_MOEDA Then
        XLV_TOTALMOEDA(XLI_CONT, 2) = XLV_TOTALMOEDA(XLI_CONT, 2) + XLF_VALOR
        XLI_CONT = XLV_TOTALMOEDA.Count(1)
        XLB_ACHOU = True
      End If
    Next
    
    If XLB_ACHOU = False Then
      XLV_TOTALMOEDA.ReDim 0, XLI_CONT, 0, 2
      XLV_TOTALMOEDA(XLI_CONT, 0) = XLT_MOEDA
      XLV_TOTALMOEDA(XLI_CONT, 1) = XLI_MOEDA
      XLV_TOTALMOEDA(XLI_CONT, 2) = XLF_VALOR
    End If
  End If

End Function

' Pega a cotação do mês e do mês anterior, pega o valor do titulo na moeda escolhida
' e verifica se tem indexador. Usada para o calculo do valor do titulo, do valor corrigido
' e para o valor em debito atualizado

' Usado pela mapa contábil, a única diferença é que busca os campos cotação 4 ,cotação 5
' , cotacao4_mesanterior e cotacao5_mesanterior
' e que correspondem à cotação 1 , cotação 2, cotacao1_mesanterior e cotacao2_mesanterior
' só que com um mês a menos
Sub SubCotacaoMapaContabil(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLI_INDEXADOR As Byte, XLO_VETOR As XArray, XLB_IGNORARPAGO As Boolean)
  'Parâmetros
  'XLO_CAMPOS - Campos da consulta
  'XLI_INDEXADOR - Indexador escolhido (0,1,2)
  'XLO_VETOR - Vetor para ser preenchido com cotaões
  
  'XLB_IGNORARPAGO - Ignorar se o titulo está pago, tratando-o como se não estivesse pago.
  '  Usado quando é escolhido a opção "considerar tiulos pagos"

  Dim XLF_COTACAO As Double           'Armazena a cotacao na data base do pagto
  Dim XLF_COTACAOANTERIOR As Double   'Armazena a cotacao na data 1 mês antes do mês da dt base do pagto
  Dim XLF_VALORTITULO As Double     'Armazena o valor do título de acordo indexador escolhido
  Dim XLT_INDEXACAO As String         'M - Mensal D - Diário
  Dim XLB_SEMINDEXADOR As Boolean
  
  XLO_VETOR.ReDim 0, 4, 0, 1
    
  'não está pago
  If XLB_IGNORARPAGO = True Or (XLB_IGNORARPAGO = False And XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value = "") Then
  
    If XLI_INDEXADOR = 1 Then
      If XLO_CAMPOS.GetItemByName("moed_cd_Moeda1").Value = "" Then
        XLB_SEMINDEXADOR = True
      End If
    ElseIf XLI_INDEXADOR = 2 Then
      If XLO_CAMPOS.GetItemByName("moed_cd_Moeda2").Value = "" Then
        XLB_SEMINDEXADOR = True
      End If
    Else
      XLB_SEMINDEXADOR = True 'Sem indexador
    End If
    
    If XLB_SEMINDEXADOR = True Then
      XLF_VALORTITULO = XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value
    End If
    
    If XLB_SEMINDEXADOR = False Then
    
      'Verificar o tipo de indexação
      If XLI_INDEXADOR = 1 Then 'Indexador 1
        'Pegar o valor da parcela indexado pelo indexador 1
        XLF_VALORTITULO = XLO_CAMPOS.GetItemByName("titu_vl_ParcelaIndex1").Value
        
        'Pegar a cotacao da moeda 1
        XLF_COTACAO = XLO_CAMPOS.GetItemByName("Cotacao4").Value
        
        If XLO_CAMPOS.GetItemByName("Indexacao1").Value = "M" Then
          'Pegar a cotacao da moeda 1 no mês anterior
          XLF_COTACAOANTERIOR = XLO_CAMPOS.GetItemByName("Cotacao4_MesAnterior").Value
          XLT_INDEXACAO = "M"
        Else
          XLT_INDEXACAO = "D"
        End If
          
      Else 'Indexador 2
        'Pegar o valor da parcela indexado pelo indexador 2
        XLF_VALORTITULO = XLO_CAMPOS.GetItemByName("titu_vl_ParcelaIndex2").Value
        'Pegar a cotacao da moeda 2
        XLF_COTACAO = XLO_CAMPOS.GetItemByName("Cotacao5").Value
        
        If XLO_CAMPOS.GetItemByName("Indexacao2").Value = "M" Then
          'Pegar a cotacao da moeda 2 no mês anterior
          XLF_COTACAOANTERIOR = XLO_CAMPOS.GetItemByName("Cotacao5_MesAnterior").Value
          XLT_INDEXACAO = "M"
        Else
          XLT_INDEXACAO = "D"
        End If
      End If
    End If
  Else 'Se foi pago
    
    If XLI_INDEXADOR = 0 Or XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value = "" Then
      XLB_SEMINDEXADOR = True
      XLF_VALORTITULO = XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value
    End If
    
    If XLB_SEMINDEXADOR = False Then
    
      XLB_SEMINDEXADOR = True
      If XLO_CAMPOS.GetItemByName("moed_cd_Moeda1").Value = XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value Then
        'Pegar o valor da parcela indexado pelo indexador 1
        XLF_VALORTITULO = XLO_CAMPOS.GetItemByName("titu_vl_ParcelaIndex1").Value
        XLB_SEMINDEXADOR = False
      End If
        
      'Se o indexador 2 não é nulo e moeda3 é diferente de moeda1
      If XLB_SEMINDEXADOR = True Then
        If XLO_CAMPOS.GetItemByName("moed_cd_Moeda2").Value = XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value Then
          'Pegar o valor da parcela indexado pelo indexador 2
          XLF_VALORTITULO = XLO_CAMPOS.GetItemByName("titu_vl_ParcelaIndex2").Value
          XLB_SEMINDEXADOR = False
        End If
      End If
      
      'Pegar a cotacao da moeda do pagamento
      XLF_COTACAO = XLO_CAMPOS.GetItemByName("Cotacao3").Value
  
      If XLO_CAMPOS.GetItemByName("Indexacao3").Value = "M" Then
       'Pegar a cotacao da moeda 1 no mês anterior
        XLF_COTACAOANTERIOR = XLO_CAMPOS.GetItemByName("Cotacao3_MesAnterior").Value
        XLT_INDEXACAO = "M"
      Else
        XLT_INDEXACAO = "D"
      End If
    End If
  End If
  
  XLO_VETOR(0, 0) = CStr(XLF_VALORTITULO)
  XLO_VETOR(1, 0) = CStr(XLF_COTACAO)
  XLO_VETOR(2, 0) = CStr(XLF_COTACAOANTERIOR)
  XLO_VETOR(3, 0) = XLT_INDEXACAO
  If XLB_SEMINDEXADOR = True Then
    XLO_VETOR(4, 0) = 1
  Else
    XLO_VETOR(4, 0) = 0
  End If
    
End Sub

' Pega a cotação do mês e do mês anterior, pega o valor do titulo na moeda escolhida
' e verifica se tem indexador. Usada para o calculo do valor do titulo, do valor corrigido
' e para o valor em debito atualizado
Sub SubCotacao(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLI_INDEXADOR As Byte, XLO_VETOR As XArray, XLB_IGNORARPAGO As Boolean)
  'Parâmetros
  'XLO_CAMPOS - Campos da consulta
  'XLI_INDEXADOR - Indexador escolhido (0,1,2)
  'XLO_VETOR - Vetor para ser preenchido com cotaões
  
  'XLB_IGNORARPAGO - Ignorar se o titulo está pago, tratando-o como se não estivesse pago.
  '  Usado quando é escolhido a opção "considerar tiulos pagos"

  Dim XLF_COTACAO As Double           'Armazena a cotacao na data base do pagto
  Dim XLF_COTACAOANTERIOR As Double   'Armazena a cotacao na data 1 mês antes do mês da dt base do pagto
  Dim XLF_VALORTITULO As Double     'Armazena o valor do título de acordo indexador escolhido
  Dim XLT_INDEXACAO As String         'M - Mensal D - Diário
  Dim XLB_SEMINDEXADOR As Boolean
  
  XLO_VETOR.ReDim 0, 4, 0, 1
    
  'não está pago
  If XLB_IGNORARPAGO = True Or (XLB_IGNORARPAGO = False And XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value = "") Then
  
    If XLI_INDEXADOR = 1 Then
      If XLO_CAMPOS.GetItemByName("moed_cd_Moeda1").Value = "" Then
        XLB_SEMINDEXADOR = True
      End If
    ElseIf XLI_INDEXADOR = 2 Then
      If XLO_CAMPOS.GetItemByName("moed_cd_Moeda2").Value = "" Then
        XLB_SEMINDEXADOR = True
      End If
    Else
      XLB_SEMINDEXADOR = True 'Sem indexador
    End If
    
    If XLB_SEMINDEXADOR = True Then
      XLF_VALORTITULO = XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value
    End If
    
    If XLB_SEMINDEXADOR = False Then
    
      'Verificar o tipo de indexação
      If XLI_INDEXADOR = 1 Then 'Indexador 1
        'Pegar o valor da parcela indexado pelo indexador 1
        XLF_VALORTITULO = XLO_CAMPOS.GetItemByName("titu_vl_ParcelaIndex1").Value
        
        'Pegar a cotacao da moeda 1
        XLF_COTACAO = XLO_CAMPOS.GetItemByName("Cotacao1").Value
        
        If XLO_CAMPOS.GetItemByName("Indexacao1").Value = "M" Then
          'Pegar a cotacao da moeda 1 no mês anterior
          XLF_COTACAOANTERIOR = XLO_CAMPOS.GetItemByName("Cotacao1_MesAnterior").Value
          XLT_INDEXACAO = "M"
        Else
          XLT_INDEXACAO = "D"
        End If
          
      Else 'Indexador 2
        'Pegar o valor da parcela indexado pelo indexador 2
        XLF_VALORTITULO = XLO_CAMPOS.GetItemByName("titu_vl_ParcelaIndex2").Value
        'Pegar a cotacao da moeda 2
        XLF_COTACAO = XLO_CAMPOS.GetItemByName("Cotacao2").Value
        
        If XLO_CAMPOS.GetItemByName("Indexacao2").Value = "M" Then
          'Pegar a cotacao da moeda 2 no mês anterior
          XLF_COTACAOANTERIOR = XLO_CAMPOS.GetItemByName("Cotacao2_MesAnterior").Value
          XLT_INDEXACAO = "M"
        Else
          XLT_INDEXACAO = "D"
        End If
      End If
    End If
  Else 'Se foi pago
    
    If XLI_INDEXADOR = 0 Or XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value = "" Then
      XLB_SEMINDEXADOR = True
      XLF_VALORTITULO = XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value
    End If
    
    If XLB_SEMINDEXADOR = False Then
    
      XLB_SEMINDEXADOR = True
      If XLO_CAMPOS.GetItemByName("moed_cd_Moeda1").Value = XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value Then
        'Pegar o valor da parcela indexado pelo indexador 1
        XLF_VALORTITULO = XLO_CAMPOS.GetItemByName("titu_vl_ParcelaIndex1").Value
        XLB_SEMINDEXADOR = False
      End If
        
      'Se o indexador 2 não é nulo e moeda3 é diferente de moeda1
      If XLB_SEMINDEXADOR = True Then
        If XLO_CAMPOS.GetItemByName("moed_cd_Moeda2").Value = XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value Then
          'Pegar o valor da parcela indexado pelo indexador 2
          XLF_VALORTITULO = XLO_CAMPOS.GetItemByName("titu_vl_ParcelaIndex2").Value
          XLB_SEMINDEXADOR = False
        End If
      End If
      
      'Pegar a cotacao da moeda do pagamento
      XLF_COTACAO = XLO_CAMPOS.GetItemByName("Cotacao3").Value
  
      If XLO_CAMPOS.GetItemByName("Indexacao3").Value = "M" Then
       'Pegar a cotacao da moeda 1 no mês anterior
        XLF_COTACAOANTERIOR = XLO_CAMPOS.GetItemByName("Cotacao3_MesAnterior").Value
        XLT_INDEXACAO = "M"
      Else
        XLT_INDEXACAO = "D"
      End If
    End If
  End If
  
  XLO_VETOR(0, 0) = CStr(XLF_VALORTITULO)
  XLO_VETOR(1, 0) = CStr(XLF_COTACAO)
  XLO_VETOR(2, 0) = CStr(XLF_COTACAOANTERIOR)
  XLO_VETOR(3, 0) = XLT_INDEXACAO
  If XLB_SEMINDEXADOR = True Then
    XLO_VETOR(4, 0) = 1
  Else
    XLO_VETOR(4, 0) = 0
  End If
    
End Sub

' Pega a cotação do mês e do mês anterior, pega o valor do titulo na moeda escolhida
' ignorando se o indexador escolhido for 0(REAL). Usada para o calculo do valor liquidado atualizado.

' Usado pela mapa contábil, a única diferença é que busca os campos cotação 4 ,cotação 5
' , cotacao4_mesanterior e cotacao5_mesanterior
' e que correspondem à cotação 1 , cotação 2, cotacao1_mesanterior e cotacao2_mesanterior
' só que com um mês a menos

Sub SubCotacao2MapaContabil(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLO_VETOR As XArray)
  
  Dim XLF_COTACAO As Double           'Armazena a cotacao na data base do pagto
  Dim XLF_COTACAOANTERIOR As Double   'Armazena a cotacao na data 1 mês antes do mês da dt base do pagto
  Dim XLT_INDEXACAO As String         'M - Mensal D - Diário
  Dim XLB_SEMINDEXADOR As Boolean
  
  XLO_VETOR.ReDim 0, 4, 0, 1
  
  If XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value = "" Then
    XLB_SEMINDEXADOR = True
  End If
  
  If XLB_SEMINDEXADOR = False Then
  
    XLB_SEMINDEXADOR = True
    If XLO_CAMPOS.GetItemByName("moed_cd_Moeda1").Value = XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value Then
      'Pegar a cotacao da moeda 1
      XLF_COTACAO = XLO_CAMPOS.GetItemByName("Cotacao4").Value
      
      If XLO_CAMPOS.GetItemByName("Indexacao3").Value = "M" Then
        'Pegar a cotacao da moeda 1 no mês anterior
        XLF_COTACAOANTERIOR = XLO_CAMPOS.GetItemByName("Cotacao4_MesAnterior").Value
        XLT_INDEXACAO = "M"
      Else
        XLT_INDEXACAO = "D"
      End If
      XLB_SEMINDEXADOR = False
    End If
      
    'Se o indexador 2 não é nulo e moeda3 é diferente de moeda1
    If XLB_SEMINDEXADOR = True Then
      If XLO_CAMPOS.GetItemByName("moed_cd_Moeda2").Value = XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value Then
        'Pegar a cotacao da moeda 2
        XLF_COTACAO = XLO_CAMPOS.GetItemByName("Cotacao5").Value
        
        If XLO_CAMPOS.GetItemByName("Indexacao3").Value = "M" Then
          'Pegar a cotacao da moeda 2 no mês anterior
          XLF_COTACAOANTERIOR = XLO_CAMPOS.GetItemByName("Cotacao5_MesAnterior").Value
          XLT_INDEXACAO = "M"
        Else
          XLT_INDEXACAO = "D"
        End If
        XLB_SEMINDEXADOR = False
      End If
    End If
  End If
  
  XLO_VETOR(0, 0) = 0
  XLO_VETOR(1, 0) = CStr(XLF_COTACAO)
  XLO_VETOR(2, 0) = CStr(XLF_COTACAOANTERIOR)
  XLO_VETOR(3, 0) = XLT_INDEXACAO
  If XLB_SEMINDEXADOR = True Then
    XLO_VETOR(4, 0) = 1
  Else
    XLO_VETOR(4, 0) = 0
  End If

End Sub



' Pega a cotação do mês e do mês anterior, pega o valor do titulo na moeda escolhida
' ignorando se o indexador escolhido for 0(REAL). Usada para o calculo do valor liquidado atualizado.
Sub SubCotacao2(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLO_VETOR As XArray)
  
  Dim XLF_COTACAO As Double           'Armazena a cotacao na data base do pagto
  Dim XLF_COTACAOANTERIOR As Double   'Armazena a cotacao na data 1 mês antes do mês da dt base do pagto
  Dim XLT_INDEXACAO As String         'M - Mensal D - Diário
  Dim XLB_SEMINDEXADOR As Boolean
  
  XLO_VETOR.ReDim 0, 4, 0, 1
  
  If XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value = "" Then
    XLB_SEMINDEXADOR = True
  End If
  
  If XLB_SEMINDEXADOR = False Then
  
    XLB_SEMINDEXADOR = True
    If XLO_CAMPOS.GetItemByName("moed_cd_Moeda1").Value = XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value Then
      'Pegar a cotacao da moeda 1
      XLF_COTACAO = XLO_CAMPOS.GetItemByName("Cotacao1").Value
      
      If XLO_CAMPOS.GetItemByName("Indexacao3").Value = "M" Then
        'Pegar a cotacao da moeda 1 no mês anterior
        XLF_COTACAOANTERIOR = XLO_CAMPOS.GetItemByName("Cotacao1_MesAnterior").Value
        XLT_INDEXACAO = "M"
      Else
        XLT_INDEXACAO = "D"
      End If
      XLB_SEMINDEXADOR = False
    End If
      
    'Se o indexador 2 não é nulo e moeda3 é diferente de moeda1
    If XLB_SEMINDEXADOR = True Then
      If XLO_CAMPOS.GetItemByName("moed_cd_Moeda2").Value = XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value Then
        'Pegar a cotacao da moeda 2
        XLF_COTACAO = XLO_CAMPOS.GetItemByName("Cotacao2").Value
        
        If XLO_CAMPOS.GetItemByName("Indexacao3").Value = "M" Then
          'Pegar a cotacao da moeda 2 no mês anterior
          XLF_COTACAOANTERIOR = XLO_CAMPOS.GetItemByName("Cotacao2_MesAnterior").Value
          XLT_INDEXACAO = "M"
        Else
          XLT_INDEXACAO = "D"
        End If
        XLB_SEMINDEXADOR = False
      End If
    End If
  End If
  
  XLO_VETOR(0, 0) = 0
  XLO_VETOR(1, 0) = CStr(XLF_COTACAO)
  XLO_VETOR(2, 0) = CStr(XLF_COTACAOANTERIOR)
  XLO_VETOR(3, 0) = XLT_INDEXACAO
  If XLB_SEMINDEXADOR = True Then
    XLO_VETOR(4, 0) = 1
  Else
    XLO_VETOR(4, 0) = 0
  End If

End Sub
'Usado para pegar a cotação3 não importando o indexador
Sub SubCotacao3(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLO_VETOR As XArray)
  Dim XLF_COTACAO As Double           'Armazena a cotacao na data base do pagto
  Dim XLF_COTACAOANTERIOR As Double   'Armazena a cotacao na data 1 mês antes do mês da dt base do pagto
  Dim XLT_INDEXACAO As String         'M - Mensal D - Diário
  Dim XLB_SEMINDEXADOR As Boolean
  
  XLO_VETOR.ReDim 0, 4, 0, 1
  
  If XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value = "" Then
    XLB_SEMINDEXADOR = True
  End If
  
  If XLB_SEMINDEXADOR = False Then
    'Pegar a cotacao da moeda do pagamento
    XLF_COTACAO = XLO_CAMPOS.GetItemByName("Cotacao3").Value
  
    If XLO_CAMPOS.GetItemByName("Indexacao3").Value = "M" Then
     'Pegar a cotacao da moeda 1 no mês anterior
      XLF_COTACAOANTERIOR = XLO_CAMPOS.GetItemByName("Cotacao3_MesAnterior").Value
      XLT_INDEXACAO = "M"
    Else
      XLT_INDEXACAO = "D"
    End If
  End If
    
  XLO_VETOR(0, 0) = 0
  XLO_VETOR(1, 0) = CStr(XLF_COTACAO)
  XLO_VETOR(2, 0) = CStr(XLF_COTACAOANTERIOR)
  XLO_VETOR(3, 0) = XLT_INDEXACAO
  If XLB_SEMINDEXADOR = True Then
    XLO_VETOR(4, 0) = 1
  Else
    XLO_VETOR(4, 0) = 0
  End If

End Sub

Function FunCalculoValorOriginalPago(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions) As Double

FunCalculoValorOriginalPago = XLO_CAMPOS.GetItemByName("titu_vl_ValorPago").Value _
      - XLO_CAMPOS.GetItemByName("titu_vl_Multa").Value _
      - XLO_CAMPOS.GetItemByName("titu_vl_Juros").Value _
      - XLO_CAMPOS.GetItemByName("titu_vl_Seguro").Value _
      - XLO_CAMPOS.GetItemByName("titu_vl_Outros").Value _
      + XLO_CAMPOS.GetItemByName("titu_vl_Desagio").Value _
      + XLO_CAMPOS.GetItemByName("titu_vl_Desconto").Value

End Function

Function FunCalculoValorPago(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLT_FILTROTITULOSPAGOS) As Double

  If XLT_FILTROTITULOSPAGOS <> "1" Then
    FunCalculoValorPago = XLO_CAMPOS.GetItemByName("titu_vl_ValorPago").Value _
      - XLO_CAMPOS.GetItemByName("titu_vl_Multa").Value _
      - XLO_CAMPOS.GetItemByName("titu_vl_Juros").Value _
      - XLO_CAMPOS.GetItemByName("titu_vl_Seguro").Value _
      - XLO_CAMPOS.GetItemByName("titu_vl_Outros").Value
  Else
    FunCalculoValorPago = XLO_CAMPOS.GetItemByName("titu_vl_ValorPago").Value
  End If
  
End Function

Function FunCalculoValorDistrato(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions) As Double

  FunCalculoValorDistrato = XLO_CAMPOS.GetItemByName("titu_vl_ValorPago").Value _
    - XLO_CAMPOS.GetItemByName("titu_vl_Multa").Value _
    - XLO_CAMPOS.GetItemByName("titu_vl_Juros").Value _
    - XLO_CAMPOS.GetItemByName("titu_vl_Seguro").Value _
    - XLO_CAMPOS.GetItemByName("titu_vl_Outros").Value
  
End Function

Function FunCalculoCorrecaoMonetaria(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLO_VETOR As XArray, XLD_DATABASE As Date) As Double
    
  'Verifica se tem indexador
  If XLO_VETOR(4, 0) = 0 Then
     FunCalculoCorrecaoMonetaria = FunCorrecaoMonetaria(XLO_VETOR(0, 0), XLO_VETOR(1, 0), _
      XLO_VETOR(2, 0), XLO_VETOR(3, 0), XLO_CAMPOS.GetItemByName("cont_tx_CorrecProRata").Value, _
      XLD_DATABASE)
  Else
    FunCalculoCorrecaoMonetaria = XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value
  End If
  
End Function

Function FunCorrecaoMonetaria(XLF_VALORTITULO As Double, XLF_COTACAO As Double, XLF_COTACAOANTERIOR As Double, XLT_INDEXACAO As String, XLT_PRORATA As String, XLD_DATABASE As Date) As Double
  
  Dim XLF_PERCENT As Double   'Divisão entre a cotacao do mês e do mês anterior
  Dim XLF_CORRECAO As Double
  
  If XLT_INDEXACAO = "D" Or (XLT_INDEXACAO = "M" And XLT_PRORATA = "0") Then
    'Indexação diária or Indexacao Mensal sem Correção monetária pro-rata
    XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
  ElseIf Day(XLD_DATABASE) = 1 Then  'Indexação Mensal, Pro-Rata Dia dtBase = 1
    XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
  Else
    If XLT_PRORATA = "A" Then 'Prorata com base no mês anterior
      If XLF_COTACAOANTERIOR = 0 Then
        XLF_CORRECAO = 0
      Else
        XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
        XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DATABASE) - 1)) * XLF_COTACAO)
        XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
      End If
    Else 'Prorata com base no mês posterior. O nome da variável apesar de ser
         'XLF_COTACAOMESANTERIOR refere-se à cotação do mês posterior
      If XLF_COTACAO = 0 Then
        XLF_CORRECAO = 0
      Else
        XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAO
        XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DATABASE) - 1)) * XLF_COTACAO)
        XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
      End If
    End If
  End If
  FunCorrecaoMonetaria = XLF_CORRECAO

End Function

'Function FunCorrecaoMonetariaTEMP(XLF_VALORTITULO As Double, XLF_COTACAO As Double, XLF_COTACAOANTERIOR As Double, XLT_INDEXACAO As String, XLI_PRORATA As Byte, XLD_DATABASE As Date) As Double
'
'  Dim XLF_PERCENT As Double   'Divisão entre a cotacao do mês e do mês anterior
'  Dim XLF_CORRECAO As Double
'
'  If XLT_INDEXACAO = "D" Or (XLT_INDEXACAO = "M" And XLI_PRORATA = 0) Then
'    'Indexação diária or Indexacao Mensal sem Correção monetária pro-rata
'    XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
'  ElseIf Day(XLD_DATABASE) = 1 Then  'Indexação Mensal, Pro-Rata Dia dtBase = 1
'    XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
'  Else
'    If XLF_COTACAOANTERIOR = 0 Then
'      XLF_CORRECAO = 0
'    Else
'      XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
'      XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DATABASE) - 1)) * XLF_COTACAO)
'      XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
'    End If
'  End If
'  FunCorrecaoMonetaria = XLF_CORRECAO
'
'End Function

Function FunCalculoJuros(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLD_DATABASE As Date, XLF_VALOR As Double)
 
  Dim XLD_PRORROGACAO As Date
    
  XLD_PRORROGACAO = XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value
    
  If XLD_DATABASE > XLD_PRORROGACAO Then
    If Weekday(XLD_DATABASE) = 2 Then
      If (Weekday(XLD_PRORROGACAO) = 1 And (DateDiff("y", XLD_PRORROGACAO, XLD_DATABASE) = 1) _
        Or Weekday(XLD_PRORROGACAO) = 7 And DateDiff("y", XLD_PRORROGACAO, XLD_DATABASE) = 2) Then
          FunCalculoJuros = 0
      Else
        FunCalculoJuros = FunJuros(XLF_VALOR, XLO_CAMPOS.GetItemByName("titu_vl_PercJurosMora").Value, _
           XLO_CAMPOS.GetItemByName("cont_tx_TipoJurosMora").Value, XLD_DATABASE, _
           XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value, _
           XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value)
      End If
    Else
      FunCalculoJuros = FunJuros(XLF_VALOR, XLO_CAMPOS.GetItemByName("titu_vl_PercJurosMora").Value, _
        XLO_CAMPOS.GetItemByName("cont_tx_TipoJurosMora").Value, XLD_DATABASE, _
        XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value, _
        XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value)
    End If
  Else
    FunCalculoJuros = 0
  End If
    
End Function

Function FunJuros(XLF_VALOR As Double, XLF_TAXA As Double, XLT_TIPOJUROS As String, XLD_DATABASE As Date, XLD_VENCIMENTO As Date, XLF_VALORHISTORICO As Double) As Double
    
  Dim XLI_NUMDIAS As Integer
  Dim XLI_NUMMESES As Integer
  Dim XLF_VALORJUROS As Double
    
  XLF_TAXA = XLF_TAXA / 100
  XLI_NUMMESES = (XLD_DATABASE - XLD_VENCIMENTO) \ 30
  Select Case XLT_TIPOJUROS
    Case "D"
      XLI_NUMDIAS = XLD_DATABASE - XLD_VENCIMENTO
      XLF_VALORJUROS = (((1 + XLF_TAXA) ^ XLI_NUMDIAS) - 1) * XLF_VALOR
    Case "M"
      XLF_VALORJUROS = (((1 + XLF_TAXA) ^ XLI_NUMMESES) - 1) * XLF_VALOR
    Case "C"
     If ((XLD_DATABASE - XLD_VENCIMENTO) Mod 30) > 0 Then
        XLI_NUMMESES = XLI_NUMMESES + 1
     End If
      XLF_VALORJUROS = (((1 + XLF_TAXA) ^ XLI_NUMMESES) - 1) * XLF_VALORHISTORICO
  End Select
  
  FunJuros = XLF_VALORJUROS

End Function

Function FunCalculoJurosCEF(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLD_DATABASE As Date, XLF_VALOR As Double)
   
  If XLO_CAMPOS.GetItemByName("cont_tx_Desagio").Value = "S" Then
    FunCalculoJurosCEF = FunJurosCEF(XLF_VALOR, XLO_CAMPOS.GetItemByName("cont_vl_PercJurosFin").Value, _
             XLO_CAMPOS.GetItemByName("titu_vl_PercJurosFin").Value, _
             XLD_DATABASE, _
             XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value)
  Else
    FunCalculoJurosCEF = XLF_VALOR
  End If
    
End Function

'Juros de Financiamento da Caixa Econômica Federal
Function FunJurosCEF(XLF_VALOR As Double, XLF_TAXAMENSAL As Double, XLF_TAXATITULO As Double, XLD_DATABASE As Date, XLD_VENCIMENTO As Date) As Double
    
  Dim XLI_NUMDIAS As Integer
  Dim XLI_NUMMESES As Integer
  
  Dim XLF_TAXA As Double
    
  XLF_TAXATITULO = XLF_TAXATITULO / 100
  XLF_TAXAMENSAL = XLF_TAXAMENSAL / 100
  
  If XLD_DATABASE = XLD_VENCIMENTO Then
    XLF_TAXA = XLF_TAXATITULO
  Else
    XLI_NUMMESES = DateDiff("m", XLD_VENCIMENTO, XLD_DATABASE)
    XLF_TAXA = XLF_TAXATITULO + (XLF_TAXAMENSAL * XLI_NUMMESES)
    
    If Day(XLD_DATABASE) <> Day(XLD_VENCIMENTO) Then
      XLI_NUMDIAS = Day(XLD_DATABASE) - Day(XLD_VENCIMENTO)
      XLF_TAXA = XLF_TAXA + ((XLF_TAXAMENSAL / 30) * XLI_NUMDIAS)
    End If
  End If
    
  FunJurosCEF = XLF_VALOR * (1 + XLF_TAXA)

End Function


Function FunCalculoMulta(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLD_DATABASE As Date, XLF_VALOR As Double)
  
  Dim XLD_PRORROGACAO As Date
    
  XLD_PRORROGACAO = XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value
  
  If XLD_DATABASE > XLD_PRORROGACAO Then
    If Weekday(XLD_DATABASE) = 2 Then
      If (Weekday(XLD_PRORROGACAO) = 1 And (DateDiff("y", XLD_PRORROGACAO, XLD_DATABASE) = 1) _
        Or Weekday(XLD_PRORROGACAO) = 7 And DateDiff("y", XLD_PRORROGACAO, XLD_DATABASE) = 2) Then
          FunCalculoMulta = 0
      Else
        FunCalculoMulta = FunMulta(XLF_VALOR, XLO_CAMPOS.GetItemByName("titu_vl_PercMultaMora").Value, _
          XLO_CAMPOS.GetItemByName("cont_tx_MultaProRata").Value, _
          XLO_CAMPOS.GetItemByName("cont_tx_TipoJurosMora").Value, XLD_DATABASE, _
          XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value, _
          XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value)
      End If
    Else
      FunCalculoMulta = FunMulta(XLF_VALOR, XLO_CAMPOS.GetItemByName("titu_vl_PercMultaMora").Value, _
        XLO_CAMPOS.GetItemByName("cont_tx_MultaProRata").Value, _
        XLO_CAMPOS.GetItemByName("cont_tx_TipoJurosMora").Value, XLD_DATABASE, _
        XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value, _
        XLO_CAMPOS.GetItemByName("titu_vl_Parcela").Value)
    End If
  Else
    FunCalculoMulta = 0
  End If
  
End Function

Function FunMulta(XLF_VALOR As Double, XLF_TAXA As Double, XLI_PRORATA As Byte, XLT_TIPOJUROS As String, XLD_DATABASE As Date, XLD_VENCIMENTO As Date, XLF_VALORHISTORICO As Double) As Double

  Dim XLI_NUMDIAS As Integer
  Dim XLI_NUMMESES As Integer
  Dim XLF_VALORMULTA As Double
  
  XLF_TAXA = XLF_TAXA / 100
  XLI_NUMDIAS = XLD_DATABASE - XLD_VENCIMENTO
  'Se for Pro-rata e juros não é cheio
  If XLI_PRORATA = "1" And XLI_NUMDIAS <= 30 And XLT_TIPOJUROS <> "C" Then
    XLF_VALORMULTA = (((1 + XLF_TAXA) ^ (1 / 30) ^ XLI_NUMDIAS) - 1) * XLF_VALOR
  Else
    If XLT_TIPOJUROS = "C" Then
      'Quando o juros é cheio o valor deve ser sempre sobre o valor histórico
      XLF_VALORMULTA = XLF_TAXA * XLF_VALORHISTORICO
    Else
      XLF_VALORMULTA = XLF_TAXA * XLF_VALOR
    End If
  End If
  
  FunMulta = XLF_VALORMULTA
  
End Function

Function FunCalculoDesagio(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLD_DATABASE As Date, XLF_VALOR As Double)
    
  If XLO_CAMPOS.GetItemByName("titu_vl_PercJurosFin").Value <> "S" Then
    FunCalculoDesagio = FunDesagio(XLF_VALOR, XLO_CAMPOS.GetItemByName("titu_vl_PercJurosFin").Value, _
      XLO_CAMPOS.GetItemByName("cont_tx_Desagio").Value, XLD_DATABASE, _
      XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value)
  Else
   FunCalculoDesagio = 0
  End If

End Function

Function FunDesagio(XLF_VALOR As Double, XLF_TAXA As Double, XLT_TIPO As String, XLD_DATABASE As Date, XLD_VENCIMENTO As Date) As Double

  Dim XLI_NUMDIAS As Integer          'Armazena o nº de dias em atraso
  Dim XLI_NUMMESES As Integer         'Armazena o nº de meses em atraso
  
  If XLD_DATABASE < XLD_VENCIMENTO Then
  
    XLI_NUMDIAS = XLD_VENCIMENTO - XLD_DATABASE
    
    If XLT_TIPO = "M" Then 'Mensal
      If XLI_NUMDIAS >= 30 Then
        XLI_NUMMESES = (XLD_VENCIMENTO - XLD_DATABASE) \ 30
        FunDesagio = (XLF_VALOR - (XLF_VALOR / ((1 + (XLF_TAXA) / 100) ^ XLI_NUMMESES)))
      End If
    Else 'Diário
      XLF_TAXA = (1 + (XLF_TAXA / 100)) ^ (1 / 30)
      XLF_TAXA = (XLF_TAXA - 1) * 100
      FunDesagio = XLF_VALOR - (XLF_VALOR / ((1 + XLF_TAXA / 100) ^ XLI_NUMDIAS))
    End If
 Else
   FunDesagio = 0
 End If
  
End Function

Function FunCalculoTitulo(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLO_VETOR As XArray, XLD_DATABASE As Date, XLO_FILTROS As CRAXDDRT.FormulaFieldDefinitions) As Double
  
  Dim XLF_COTACAO As Double           'Armazena a cotacao na data base do pagto
  Dim XLF_COTACAOANTERIOR As Double   'Armazena a cotacao na data 1 mês antes do mês da dt base do pagto
  Dim XLT_INDEXACAO As String         'M - Mensal D - Diário
  Dim XLB_SEMINDEXADOR As Boolean
  Dim XLF_JUROS As Double
  Dim XLF_MULTA As Double
  Dim XLF_DESAGIO As Double
  Dim XLF_VALOR As Double
  'Dim XLO_VETOR As XArray
  
  XLF_VALOR = XLO_VETOR(0, 0)
  
  If XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value = "" Then
    
    If XLO_FILTROS.GetItemByName("FiltroJuros").Value = "1" Then
       XLF_JUROS = FunCalculoJuros(XLO_CAMPOS, XLD_DATABASE, XLF_VALOR)
    End If
    
    If XLO_FILTROS.GetItemByName("FiltroMulta").Value = "1" Then
      XLF_MULTA = FunCalculoMulta(XLO_CAMPOS, XLD_DATABASE, XLF_VALOR)
    End If
    
    If XLO_FILTROS.GetItemByName("FiltroDesagio").Value = "1" Then
      XLF_DESAGIO = FunCalculoDesagio(XLO_CAMPOS, XLD_DATABASE, XLF_VALOR)
    End If
    
 Else
    If XLO_FILTROS.GetItemByName("FiltroTitulosPagos").Value = "1" Then
      XLF_JUROS = FunCalculoJuros(XLO_CAMPOS, XLD_DATABASE, XLF_VALOR)
      XLF_MULTA = FunCalculoMulta(XLO_CAMPOS, XLD_DATABASE, XLF_VALOR)
    End If
    XLF_DESAGIO = FunCalculoDesagio(XLO_CAMPOS, XLD_DATABASE, XLF_VALOR)
  End If
  
  ' quando é cheio é calculado a multa e o juros sobre valor da parcela e
  ' depois tem que transformar para o valor indexado pela moeda atual se tiver indexador
  If XLO_CAMPOS.GetItemByName("cont_tx_TipoJurosMora").Value = "C" And XLO_VETOR(4, 0) = 0 Then
            
        XLF_JUROS = FunValorIndexado(XLF_JUROS, XLO_VETOR(1, 0), _
        XLO_VETOR(2, 0), XLO_VETOR(3, 0), XLO_CAMPOS.GetItemByName("cont_tx_CorrecProRata").Value, _
        XLD_DATABASE)
      
      XLF_MULTA = FunValorIndexado(XLF_MULTA, XLO_VETOR(1, 0), _
        XLO_VETOR(2, 0), XLO_VETOR(3, 0), XLO_CAMPOS.GetItemByName("cont_tx_CorrecProRata").Value, _
        XLD_DATABASE)
  End If
  FunCalculoTitulo = XLF_VALOR + XLF_JUROS + XLF_MULTA - XLF_DESAGIO
  
End Function

'Function FunValorIndexadoTEMP(XLF_VALORTITULO As Double, XLF_COTACAO As Double, XLF_COTACAOANTERIOR As Double, XLT_INDEXACAO As String, XLI_PRORATA As Byte, XLD_DATABASE As Date) As Double
'
'  Dim XLF_PERCENT As Double   'Divisão entre a cotacao do mês e do mês anterior
'  Dim XLF_VALORINDEXADO As Double
'
'  If XLF_COTACAO = 0 Then
'    FunValorIndexado = 0
'  Else
'    If XLT_INDEXACAO = "D" Or (XLT_INDEXACAO = "M" And XLI_PRORATA = 0) Then
'      'Indexação diária or Indexacao Mensal sem Correção monetária pro-rata
'      XLF_VALORINDEXADO = XLF_VALORTITULO / XLF_COTACAO
'    ElseIf Day(XLD_DATABASE) = 1 Then  'Indexação Mensal, Pro-Rata Dia dtBase = 1
'      XLF_VALORINDEXADO = XLF_VALORTITULO / XLF_COTACAO
'    Else
'      XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
'      XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DATABASE) - 1)) * XLF_COTACAO)
'      XLF_VALORINDEXADO = XLF_VALORTITULO / XLF_COTACAO
'    End If
'  End If
'  FunValorIndexado = XLF_VALORINDEXADO
'
'End Function

Function FunValorIndexado(XLF_VALORTITULO As Double, XLF_COTACAO As Double, XLF_COTACAOANTERIOR As Double, XLT_INDEXACAO As String, XLT_PRORATA As String, XLD_DATABASE As Date) As Double
  
  Dim XLF_PERCENT As Double   'Divisão entre a cotacao do mês e do mês anterior
  Dim XLF_VALORINDEXADO As Double
  
  If XLF_COTACAO = 0 Then
    FunValorIndexado = 0
  Else
    If XLT_INDEXACAO = "D" Or (XLT_INDEXACAO = "M" And XLT_PRORATA = "0") Then
      'Indexação diária or Indexacao Mensal sem Correção monetária pro-rata
      XLF_VALORINDEXADO = XLF_VALORTITULO / XLF_COTACAO
    ElseIf Day(XLD_DATABASE) = 1 Then  'Indexação Mensal, Pro-Rata Dia dtBase = 1
      XLF_VALORINDEXADO = XLF_VALORTITULO / XLF_COTACAO
    Else
      If XLF_COTACAOANTERIOR = 0 Then
        XLF_VALORINDEXADO = 0
      Else
        If XLT_PRORATA = "A" Then 'Prorata com base no mês anterior
          XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
          XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DATABASE) - 1)) * XLF_COTACAO)
          XLF_VALORINDEXADO = XLF_VALORTITULO / XLF_COTACAO
        Else  'Prorata com base no mês posterior
          XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAO
          XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DATABASE) - 1)) * XLF_COTACAO)
          XLF_VALORINDEXADO = XLF_VALORTITULO / XLF_COTACAO
        End If
      End If
    End If
  End If
  FunValorIndexado = XLF_VALORINDEXADO

End Function


Function FunCalculoValorIndexado(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLO_VETOR As XArray, XLD_DATABASE As Date, XLF_VALOR As Double)
    
  'Verifica se tem indexador
  If XLO_VETOR(4, 0) = 0 Then
    FunCalculoValorIndexado = FunValorIndexado(XLF_VALOR, XLO_VETOR(1, 0), _
      XLO_VETOR(2, 0), XLO_VETOR(3, 0), XLO_CAMPOS.GetItemByName("cont_tx_CorrecProRata").Value, _
      XLD_DATABASE)
  Else
    FunCalculoValorIndexado = XLF_VALOR
  End If
  
End Function

Function FunCalculoValorAtualizado(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLO_VETOR As XArray, XLD_DATABASE As Date, XLF_VALORINDEXADO As Double)
  
  'Verifica se tem indexador
  If XLO_VETOR(4, 0) = 0 Then
     FunCalculoValorAtualizado = FunCorrecaoMonetaria(XLF_VALORINDEXADO, XLO_VETOR(1, 0), _
      XLO_VETOR(2, 0), XLO_VETOR(3, 0), XLO_CAMPOS.GetItemByName("cont_tx_CorrecProRata").Value, _
      XLD_DATABASE)
  Else
    FunCalculoValorAtualizado = XLF_VALORINDEXADO
  End If

End Function

'Temporário até acrescentar XLB_IGNORARPAGO em SubUnidadeMonetaria
Sub SubUnidadeMonetariaTemp(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLI_INDEXADOR As Byte, XLO_VETOR As XArray, XLB_IGNORARPAGO As Boolean)
'Seleciona o simbolo da moeda

  Dim XLT_MOEDA As String 'Descrição do moeda
  Dim XLI_MOEDA As Integer 'Código da moeda
  
  XLO_VETOR.Clear
  XLO_VETOR.ReDim 0, 1, 0, 0
  
  If XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value = "" Or XLB_IGNORARPAGO = True Then
      If XLI_INDEXADOR = 1 Then
        'se o indexador1 for null retorna Real
        If XLO_CAMPOS.GetItemByName("moed_cd_Moeda1").Value = "" Then
          XLT_MOEDA = "REAL"
          XLI_MOEDA = 0
        Else
          XLT_MOEDA = XLO_CAMPOS.GetItemByName("Moeda1").Value
          XLI_MOEDA = XLO_CAMPOS.GetItemByName("moed_cd_Moeda1").Value
        End If
      ElseIf XLI_INDEXADOR = 2 Then
        'se o indexador2 for null retorna Real
        If XLO_CAMPOS.GetItemByName("moed_cd_Moeda2").Value = "" Then
          XLT_MOEDA = "REAL"
          XLI_MOEDA = 0
        Else
          XLT_MOEDA = XLO_CAMPOS.GetItemByName("Moeda2").Value
          XLI_MOEDA = XLO_CAMPOS.GetItemByName("moed_cd_Moeda2").Value
        End If
      ElseIf XLI_INDEXADOR = 0 Then
        XLT_MOEDA = "REAL"
        XLI_MOEDA = 0
      End If
  Else
    If XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value = "" Then
      XLT_MOEDA = "REAL"
      XLI_MOEDA = 0
    Else
      XLT_MOEDA = XLO_CAMPOS.GetItemByName("Moeda3").Value
      XLI_MOEDA = XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value
    End If
  End If
  
  XLO_VETOR(0, 0) = XLT_MOEDA
  XLO_VETOR(1, 0) = XLI_MOEDA

End Sub

Sub SubUnidadeMonetaria(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLI_INDEXADOR As Byte, XLO_VETOR As XArray, XLB_IGNORARPAGO As Boolean)
'Seleciona o simbolo da moeda

  Dim XLT_MOEDA As String 'Descrição do moeda
  Dim XLI_MOEDA As Integer 'Código da moeda
  
  XLO_VETOR.Clear
  XLO_VETOR.ReDim 0, 1, 0, 0
  
  If XLB_IGNORARPAGO = True Or (XLB_IGNORARPAGO = False And XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value = "") Then
    If XLI_INDEXADOR = 1 Then
      'se o indexador1 for null retorna Real
      If XLO_CAMPOS.GetItemByName("moed_cd_Moeda1").Value = "" Then
        XLT_MOEDA = "REAL"
        XLI_MOEDA = 0
      Else
        XLT_MOEDA = XLO_CAMPOS.GetItemByName("Moeda1").Value
        XLI_MOEDA = XLO_CAMPOS.GetItemByName("moed_cd_Moeda1").Value
      End If
    ElseIf XLI_INDEXADOR = 2 Then
      'se o indexador2 for null retorna Real
      If XLO_CAMPOS.GetItemByName("moed_cd_Moeda2").Value = "" Then
        XLT_MOEDA = "REAL"
        XLI_MOEDA = 0
      Else
        XLT_MOEDA = XLO_CAMPOS.GetItemByName("Moeda2").Value
        XLI_MOEDA = XLO_CAMPOS.GetItemByName("moed_cd_Moeda2").Value
      End If
    ElseIf XLI_INDEXADOR = 0 Then
      XLT_MOEDA = "REAL"
      XLI_MOEDA = 0
    End If
  Else
    If XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value = "" Then
      XLT_MOEDA = "REAL"
      XLI_MOEDA = 0
    Else
      XLT_MOEDA = XLO_CAMPOS.GetItemByName("Moeda3").Value
      XLI_MOEDA = XLO_CAMPOS.GetItemByName("moed_cd_Moeda3").Value
    End If
  End If
  
  XLO_VETOR(0, 0) = XLT_MOEDA
  XLO_VETOR(1, 0) = XLI_MOEDA

End Sub

Function FunObservacao(XLO_CAMPOS As CRAXDRT.DatabaseFieldDefinitions, XLD_DATABASE As Date) As String

  Dim XLT_OBSERVACAO As String
  Dim XLD_PAGAMENTO As Date
  
  If XLO_CAMPOS.GetItemByName("titu_dt_Pagamento").Value = "" Then
    If XLO_CAMPOS.GetItemByName("titu_dt_Prorrogacao").Value <> XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value Then
      XLT_OBSERVACAO = "Título Prorrogado: " & XLO_CAMPOS.GetItemByName("titu_dt_Prorrogacao").Value
    Else
      If XLO_CAMPOS.GetItemByName("obse_tx_Observacao").Value <> "" Then
        XLT_OBSERVACAO = XLO_CAMPOS.GetItemByName("obse_tx_Observacao").Value
      Else
        If XLO_CAMPOS.GetItemByName("titu_dt_Vencimento").Value < CDate(XLD_DATABASE) Then
          XLT_OBSERVACAO = "Vencido"
        Else
          XLT_OBSERVACAO = "A Vencer"
        End If
      End If
    End If
  Else
    Select Case XLO_CAMPOS.GetItemByName("titu_tx_TipoPag").Value
      Case "1"
        XLT_OBSERVACAO = "Pagto Total Corrigido no Vencimento"
      Case "2"
        XLT_OBSERVACAO = "Pagto Parcial Corrigido no Vencimento"
      Case "3"
        XLT_OBSERVACAO = "Pagto Total Corrigido Vencido"
      Case "4"
        XLT_OBSERVACAO = "Pagto Parcial Corrigido Vencido"
      Case "5"
        XLT_OBSERVACAO = "Pagto Total Corrigido Antecipado"
      Case "6"
        XLT_OBSERVACAO = "Pagto Parcial Corrigido Antecipado"
      Case "7"
        XLT_OBSERVACAO = "Pagto Total Sem Correção no Vencimento"
      Case "8"
        XLT_OBSERVACAO = "Pagto Parcial Sem Correção no Vencimento"
      Case "9"
        XLT_OBSERVACAO = "Pagto Total Sem Correção Vencido"
      Case "10"
       XLT_OBSERVACAO = "Pagto Parcial Sem Correção Vencido"
      Case "11"
        XLT_OBSERVACAO = "Pagto Total Sem Correção Antecipado"
      Case "12"
        XLT_OBSERVACAO = "Pagto Parcial Sem Correção Antecipado"
    End Select
  End If
  FunObservacao = XLT_OBSERVACAO
  
End Function

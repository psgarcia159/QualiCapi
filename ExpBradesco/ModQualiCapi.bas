Attribute VB_Name = "ModQualiCapi"
Option Explicit

Global XGF_COTACAOANTERIOR As Double 'Indica o valor da cotação do mes anterior para basear os jutos para o pagamento do título
Global XGB_RESULTADO As Boolean   'Indica se o usuário clicou em Gravar em um formulário moldal.
Global XGB_GERARESIDUO As Boolean 'Indica se o resíduo deve ser gerado ou não

Global XGI_CODMULTA As Integer 'INDICA QUAL O ID DA TABELA DESCONTOSACRESCIMOS QUE REPRESENTA Multa por Atraso - Cliente
Global XGI_CODJUROS  As Integer 'INDICA QUAL O ID DA TABELA DESCONTOSACRESCIMOS QUE REPRESENTA Juros por Atraso - Cliente
Global XGI_CODDESAGIO  As Integer 'INDICA QUAL O ID DA TABELA DESCONTOSACRESCIMOS QUE REPRESENTA Deságio - Cliente
Global XGI_CODDESCONTO  As Integer 'INDICA QUAL O ID DA TABELA DESCONTOSACRESCIMOS QUE REPRESENTA Descontos - Cliente
Global XGI_CODSEGURO  As Integer 'INDICA QUAL O ID DA TABELA DESCONTOSACRESCIMOS QUE REPRESENTA Seguro - Cliente
Global XGI_CODOUTROS  As Integer 'INDICA QUAL O ID DA TABELA DESCONTOSACRESCIMOS QUE REPRESENTA Outros - Cliente


Function FunZeros(XLI_QTD As Integer) As String
  
   Dim XLT_ZEROS As String
   Dim XLI_CONT As Integer
      
   For XLI_CONT = 1 To XLI_QTD
     XLT_ZEROS = XLT_ZEROS & "0"
   Next
   
   FunZeros = XLT_ZEROS
   
End Function

Function funVerirficaCotacao(XLO_TITULO As adodb.Recordset, XLT_INDEXADOR As Byte, XLD_DATABASE As Date, XLB_FORMULARIOCOTACAO As Boolean, XLO_TELACOTACAO As Form) As Boolean
' XLB_FORMULARIOCOTACAO Indica se apresentará a tela de cotações quando não tiver cotação
' XLO_TELACOTACAO è a tela de cotações que é passada como referência
    
  funVerirficaCotacao = True
  
  If XLT_INDEXADOR = 1 And Not IsNull(XLO_TITULO!moed_cd_Moeda1) Then
     
    If IsNull(XLO_TITULO!Cotacao1) Then
      
      XLO_TELACOTACAO.LblChave.Caption = XLO_TITULO!moed_cd_Moeda1
      
      funVerirficaCotacao = False
      
      'Abre a tela de cotacao
      If XLB_FORMULARIOCOTACAO = True Then
        If XLO_TITULO!Indexacao1 = "D" Then
          If MsgBox("Deseja cadastrar a cotação para o indexaor " & XLO_TITULO!moeda1 & " na data " & Format(XLD_DATABASE, "dd/mm/yy") & "?", vbInformation + vbYesNo, "Atualização de cotação") = vbYes Then
            XLO_TELACOTACAO.LblData.Caption = Format(XLD_DATABASE, "dd/mm/yy")
            XLO_TELACOTACAO.Show vbModal
          End If
        Else
          If MsgBox("Deseja cadastrar a cotação para o indexaor " & XLO_TITULO!moeda1 & " na data " & Format(XLD_DATABASE, "\01/mm/yy") & "?", vbInformation + vbYesNo, "Atualização de cotação") = vbYes Then
            XLO_TELACOTACAO.LblData.Caption = Format(XLD_DATABASE, "\01/mm/yy")
            XLO_TELACOTACAO.Show vbModal
          End If
        End If
      End If
    End If
    
    If XLO_TITULO!cont_tx_CorrecProRata = "M" Then 'Mista
        funVerirficaCotacao = True
    Else
        If XLO_TITULO!Indexacao1 = "M" And XLO_TITULO!cont_tx_CorrecProRata <> "0" And Day(XLD_DATABASE) <> 1 Then
          XLO_TELACOTACAO.LblChave.Caption = XLO_TITULO!moed_cd_Moeda1
          If IsNull(XLO_TITULO!Cotacao1_MesAnterior) Then
            
            funVerirficaCotacao = False
            
            'Abre a tela de cotacao
            If XLB_FORMULARIOCOTACAO = True Then
              XLD_DATABASE = Format(XLD_DATABASE, "\01/mm/yy")
              If XLO_TITULO!cont_tx_CorrecProRata = "A" Then
                XLD_DATABASE = DateAdd("m", -1, XLD_DATABASE)
              Else
                XLD_DATABASE = DateAdd("m", 1, XLD_DATABASE)
              End If
              
              If MsgBox("Deseja cadastrar a cotação para o indexaor " & XLO_TITULO!moeda1 & " na data " & Format(XLD_DATABASE, "dd/mm/yy") & "?", vbInformation + vbYesNo, "Atualização de cotação") = vbYes Then
                XLO_TELACOTACAO.LblData.Caption = XLD_DATABASE
                XLO_TELACOTACAO.Show vbModal
              End If
            End If
          End If
        End If
    End If
  ElseIf XLT_INDEXADOR = 2 And Not IsNull(XLO_TITULO!moed_cd_Moeda2) Then
    If IsNull(XLO_TITULO!Cotacao2) Then
      
      funVerirficaCotacao = False
      
      'Abre a tela de cotacao
      If XLB_FORMULARIOCOTACAO = True Then
        XLO_TELACOTACAO.LblChave.Caption = XLO_TITULO!moed_cd_Moeda2
        If XLO_TITULO!Indexacao2 = "D" Then
          If MsgBox("Deseja cadastrar a cotação para o indexaor " & XLO_TITULO!moeda2 & " na data " & Format(XLD_DATABASE, "dd/mm/yy") & "?", vbInformation + vbYesNo, "Atualização de cotação") = vbYes Then
            XLO_TELACOTACAO.LblData.Caption = Format(XLD_DATABASE, "dd/mm/yy")
            XLO_TELACOTACAO.Show vbModal
          End If
        Else
          If MsgBox("Deseja cadastrar a cotação para o indexaor " & XLO_TITULO!moeda2 & " na data " & Format(XLD_DATABASE, "\01/mm/yy") & "?", vbInformation + vbYesNo, "Atualização de cotação") = vbYes Then
            XLO_TELACOTACAO.LblData.Caption = Format(XLD_DATABASE, "\01/mm/yy")
            XLO_TELACOTACAO.Show vbModal
          End If
        End If
      End If
    End If
    
    If XLO_TITULO!cont_tx_CorrecProRata = "M" Then 'Mista
        funVerirficaCotacao = True
    Else
        If XLO_TITULO!Indexacao2 = "M" And XLO_TITULO!cont_tx_CorrecProRata <> "0" And Day(XLD_DATABASE) <> 1 Then
          If IsNull(XLO_TITULO!Cotacao2_MesAnterior) Then
            
            funVerirficaCotacao = False
            
            'Abre a tela de cotacao
            If XLB_FORMULARIOCOTACAO = True Then
              XLO_TELACOTACAO.LblChave.Caption = XLO_TITULO!moed_cd_Moeda2
              If IsNull(XLO_TITULO!Cotacao1_MesAnterior) Then
                XLD_DATABASE = Format(XLD_DATABASE, "\01/mm/yy")
                If XLO_TITULO!cont_tx_CorrecProRata = "A" Then
                  XLD_DATABASE = DateAdd("m", -1, XLD_DATABASE)
                Else
                  XLD_DATABASE = DateAdd("m", 1, XLD_DATABASE)
                End If
                If MsgBox("Deseja cadastrar a cotação para o indexador " & XLO_TITULO!moeda2 & " na data " & Format(XLD_DATABASE, "dd/mm/yy") & "?", vbInformation + vbYesNo, "Atualização de cotação") = vbYes Then
                  XLO_TELACOTACAO.LblData.Caption = XLD_DATABASE
                  XLO_TELACOTACAO.Show vbModal
                End If
              End If
            End If
          End If
        End If
    End If
  End If
 
End Function

Function FunCriaConsultaBase(XLD_DATABASE As Date, XLT_TIPOBANCO As String, XLT_NOMECONSULTA As String, XLI_TIPOCONSULTA As Byte) As String

  Dim XLT_SQL As String
  
  If XLI_TIPOCONSULTA = 1 Then 'Usado quando não precisa do indexador 3, relatórios de previsão recebimento, titulos vencidos, mapa geral e etc
  
      XLT_SQL = "" & _
      "SELECT " & XLT_NOMECONSULTA & ".*," & _
      "CASE Indexacao1   when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(XLD_DATABASE) & "/" & Year(XLD_DATABASE), XLT_TIPOBANCO) & ")" & _
        " when 'D' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(XLD_DATABASE, XLT_TIPOBANCO) & ")" & _
      " end as Cotacao1," & _
      "CASE  Indexacao2 when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(XLD_DATABASE) & "/" & Year(XLD_DATABASE), XLT_TIPOBANCO) & ")" & _
        " when 'D'  then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(XLD_DATABASE, XLT_TIPOBANCO) & ")" & _
      " end as Cotacao2,"
      
      XLT_SQL = XLT_SQL & _
      "CASE  Indexacao1 when 'M' then " & _
        "CASE cont_tx_CorrecProrata when 'A' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(XLD_DATABASE) & "/" & Year(XLD_DATABASE), XLT_TIPOBANCO) & "))" & _
        "when 'P' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData("01/" & Month(XLD_DATABASE) & "/" & Year(XLD_DATABASE), XLT_TIPOBANCO) & "))" & _
        "end " & _
      " when 'D' then Null " & _
      " end as Cotacao1_MesAnterior,"
  
      XLT_SQL = XLT_SQL & _
      "CASE Indexacao2 when 'M' then " & _
        "CASE cont_tx_CorrecProrata when 'A' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(XLD_DATABASE) & "/" & Year(XLD_DATABASE), XLT_TIPOBANCO) & "))" & _
        "when 'P' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData("01/" & Month(XLD_DATABASE) & "/" & Year(XLD_DATABASE), XLT_TIPOBANCO) & "))" & _
        "end " & _
      " when 'D' then Null " & _
      " end As Cotacao2_MesAnterior" & _
      " From " & XLT_NOMECONSULTA & " WHERE "
      
  ElseIf XLI_TIPOCONSULTA = 3 Then 'Exportacao bancária
    
      XLT_SQL = "" & _
      "SELECT " & XLT_NOMECONSULTA & ".*," & _
      "CASE Indexacao1   when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=convert(varchar,month(titu_dt_Vencimento))+'/01/' + convert(varchar,year(titu_dt_Vencimento)))" & _
        " when 'D' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=titu_dt_Vencimento )" & _
      " end as Cotacao1," & _
      "CASE  Indexacao2 when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=convert(varchar,month(titu_dt_Vencimento))+'/01/' + convert(varchar,year(titu_dt_Vencimento)))" & _
        " when 'D'  then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=titu_dt_Vencimento)" & _
      " end as Cotacao2,"
      
      XLT_SQL = XLT_SQL & _
      "CASE  Indexacao1 when 'M' then " & _
        "CASE cont_tx_CorrecProrata when 'A' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1,convert(varchar,month(titu_dt_Vencimento))+'/01/' + convert(varchar,year(titu_dt_Vencimento))))" & _
        "when 'P' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1,convert(varchar,month(titu_dt_Vencimento))+'/01/' + convert(varchar,year(titu_dt_Vencimento))))" & _
        "end " & _
      " when 'D' then Null " & _
      " end as Cotacao1_MesAnterior,"
  
      XLT_SQL = XLT_SQL & _
      "CASE Indexacao2 when 'M' then " & _
        "CASE cont_tx_CorrecProrata when 'A' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1,convert(varchar,month(titu_dt_Vencimento))+'/01/' + convert(varchar,year(titu_dt_Vencimento))))" & _
        "when 'P' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1,convert(varchar,month(titu_dt_Vencimento))+'/01/' + convert(varchar,year(titu_dt_Vencimento))))" & _
        "end " & _
      " when 'D' then Null " & _
      " end As Cotacao2_MesAnterior" & _
      " From " & XLT_NOMECONSULTA & " WHERE "
  
  ElseIf XLI_TIPOCONSULTA = 4 Then 'Entrega de chaves
    XLT_SQL = "" & _
      "SELECT " & XLT_NOMECONSULTA & ".*," & _
      "CASE Indexacao1   when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(XLD_DATABASE) & "/" & Year(XLD_DATABASE), XLT_TIPOBANCO) & ")" & _
        " when 'D' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(XLD_DATABASE, XLT_TIPOBANCO) & ")" & _
      " end as Cotacao1," & _
      "CASE  Indexacao2 when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(XLD_DATABASE) & "/" & Year(XLD_DATABASE), XLT_TIPOBANCO) & ")" & _
        " when 'D'  then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(XLD_DATABASE, XLT_TIPOBANCO) & ")" & _
      " end as Cotacao2,"
      
      XLT_SQL = XLT_SQL & _
      "CASE  Indexacao1 when 'M' then " & _
        "CASE cont_tx_CorrecProrata when 'A' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(XLD_DATABASE) & "/" & Year(XLD_DATABASE), XLT_TIPOBANCO) & "))" & _
        "when 'P' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData("01/" & Month(XLD_DATABASE) & "/" & Year(XLD_DATABASE), XLT_TIPOBANCO) & "))" & _
        "end " & _
      " when 'D' then Null " & _
      " end as Cotacao1_MesAnterior,"
  
      XLT_SQL = XLT_SQL & _
      "CASE Indexacao2 when 'M' then " & _
        "CASE cont_tx_CorrecProrata when 'A' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(XLD_DATABASE) & "/" & Year(XLD_DATABASE), XLT_TIPOBANCO) & "))" & _
        "when 'P' then" & _
          " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData("01/" & Month(XLD_DATABASE) & "/" & Year(XLD_DATABASE), XLT_TIPOBANCO) & "))" & _
        "end " & _
      " when 'D' then Null " & _
      " end As Cotacao2_MesAnterior," & _
      " datediff(m," & FunNuloData(XLD_DATABASE, XLT_TIPOBANCO) & ",titu_dt_Vencimento) as DiferencaMes" & _
      " From " & XLT_NOMECONSULTA & " WHERE "
      
  End If
  
  FunCriaConsultaBase = XLT_SQL

End Function

Function FunCorrecaoMonetaria(ByVal XLF_VALORTITULO As Double, ByVal XLF_COTACAO As Double, ByVal XLF_COTACAOANTERIOR As Double, ByVal XLT_INDEXACAO As String, ByVal XLT_PRORATA As String, ByVal XLD_DATABASE As Date, ByVal XLD_DATABASETITULO As Date, XLI_MOEDA As Integer) As Double
    'Function FunCorrecaoMonetaria(ByVal XLF_VALORTITULO As Double, ByVal XLF_COTACAO As Double, ByVal XLF_COTACAOANTERIOR As Double, ByVal XLT_INDEXACAO As String, ByVal XLT_PRORATA As String, ByVal XLD_DATABASE As Date) As Double
    
    Dim XLF_PERCENT As Double   'Divisão entre a cotacao do mês e do mês anterior
    Dim XLF_CORRECAO As Double
    Dim XLD_DATA As Date
    Dim XLO_COTACAO As adodb.Recordset
    
    'Se a moeda for 0 significa que ´deve ser informado o valor histórico, a correcao é 0
    If XLT_PRORATA = "M" And XLI_MOEDA = 0 Then
       FunCorrecaoMonetaria = XLF_CORRECAO
       Exit Function
    End If
    
    'Indexação diária or Indexacao Mensal sem Correção monetária pro-rata
    If XLT_INDEXACAO = "D" Or (XLT_INDEXACAO = "M" And XLT_PRORATA = "0") Then
        XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
    
    'Indexação Mensal, Pro-Rata Dia dtBase = 1, Mês Anterior ou Posterior
    ElseIf Day(XLD_DATABASE) = 1 And (XLT_PRORATA = "A" Or XLT_PRORATA = "P") Then
      XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
      
    Else
        'Prorata com base no mês anterior
        If XLT_PRORATA = "A" Then
          If XLF_COTACAOANTERIOR = 0 Then
                XLF_CORRECAO = 0
          Else
                XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
                XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DATABASE) - 1)) * XLF_COTACAO)
                XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
          End If
        
        'Prorata com base no mês posterior. O nome da variável apesar de ser XLF_COTACAOMESANTERIOR refere-se à cotação do mês posterior
        ElseIf XLT_PRORATA = "P" Then
          If XLF_COTACAO = 0 Then
                XLF_CORRECAO = 0
          Else
                XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAO
                XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DATABASE) - 1)) * XLF_COTACAO)
                XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
          End If
          
        'Prorata Mista - Franisa. Se for, não calcula pro-rata. Será o vl indexado * cotação cheia
        ElseIf XLT_PRORATA = "M" Then
            'Verificar se o dia da dt base do contrato é = ao dia da dt base do título
            If Day(XLD_DATABASE) = Day(XLD_DATABASETITULO) Then
                XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
                
            'Senão calcula pro-rata apenas dos dias excedentes. Ex. se a dt base do contrato é 20 e a do título é 30. Calcular cheia até 20 e pro-rata de 10 dias apenas
            ElseIf Day(XLD_DATABASE) > Day(XLD_DATABASETITULO) Then
                XLF_CORRECAO = (XLF_COTACAO * XLF_VALORTITULO)
                
                'Pegar cotação do mês anterior
                SubQOpenRecordset XLO_COTACAO, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_MOEDA & " and cota_dt_datacotacao =" & FunNuloData("01/" & CStr(Format(DateAdd("m", -1, XLD_DATABASE), "MM/yyyy")), NomeSgbd) & "", Estatico
                                
                If Not XLO_COTACAO.EOF Then XLF_COTACAOANTERIOR = XLO_COTACAO!cota_vl_valor
                
                XLO_COTACAO.Close
                Set XLO_COTACAO = Nothing
                XGF_COTACAOANTERIOR = XLF_COTACAOANTERIOR
                
                '11/04/08 Se não existir cotação do mês anterior então a percentagem é zero
                If XLF_COTACAOANTERIOR = 0 Then
                    XLF_PERCENT = 0
                Else
                    XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
                End If
                
                XLF_COTACAO = Round((((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DATABASE) - Day(XLD_DATABASETITULO))) - 1) * 100, 4)
                XLF_CORRECAO = XLF_CORRECAO + ((XLF_COTACAO * XLF_CORRECAO) / 100)
            
            Else
                Dim XLF_COTACAOANTERIOR2 As Double
                
                'Pegar cotação do mês anterior
                SubQOpenRecordset XLO_COTACAO, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_MOEDA & " and cota_dt_datacotacao =" & FunNuloData("01/" & CStr(Format(DateAdd("m", -1, XLD_DATABASE), "MM/yyyy")), NomeSgbd) & "", Estatico
                                
                If Not XLO_COTACAO.EOF Then XLF_COTACAOANTERIOR = XLO_COTACAO!cota_vl_valor
                
                XLO_COTACAO.Close
                Set XLO_COTACAO = Nothing
                'Verificar se o dia da dt base do contrato é = ao dia da dt base do título
                XLF_CORRECAO = (XLF_COTACAOANTERIOR * XLF_VALORTITULO)
                
                'Pegar cotação de 2 meses atrás
                SubQOpenRecordset XLO_COTACAO, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_MOEDA & " and cota_dt_datacotacao =" & FunNuloData("01/" & CStr(Format(DateAdd("m", -2, XLD_DATABASE), "MM/yyyy")), NomeSgbd) & "", Estatico
                                
                If Not XLO_COTACAO.EOF Then XLF_COTACAOANTERIOR2 = XLO_COTACAO!cota_vl_valor
                        
                XLO_COTACAO.Close
                Set XLO_COTACAO = Nothing
                '11/11/08 Se não existir cotação do mês anterior então a percentagem é zero
                If XLF_COTACAOANTERIOR2 = 0 Then
                    XLF_PERCENT = 0
                Else
                    XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAOANTERIOR2
                End If
                
                'Qtd de dias tem que ser 10/04-20/03 (qtd de dias entre a dt base do mês anterior e a dt base informada)
                If Not IsDate(CStr(Day(XLD_DATABASETITULO)) + "/" + CStr(Format(DateAdd("m", -1, XLD_DATABASE), "MM/yyyy"))) Then
                    XLD_DATA = "01/" + CStr(Format(XLD_DATABASE, "MM/yyyy"))
                    XLD_DATA = DateAdd("d", -1, XLD_DATA)
                Else
                    XLD_DATA = CStr(Day(XLD_DATABASETITULO)) + "/" + CStr(Format(DateAdd("m", -1, XLD_DATABASE), "MM/yyyy"))
                End If
                
                XLF_COTACAO = Round((((XLF_PERCENT ^ (1 / 30)) ^ (DateDiff("d", XLD_DATA, XLD_DATABASE))) - 1) * 100, 4)
                XLF_CORRECAO = XLF_CORRECAO + ((XLF_COTACAO * XLF_CORRECAO) / 100)
            End If
        End If
  End If
  
  FunCorrecaoMonetaria = XLF_CORRECAO

End Function

Function FunJuros(ByVal XLF_VALOR As Double, ByVal XLF_TAXA As Double, ByVal XLT_TIPOJUROS As String, ByVal XLD_DATABASE As Date, ByVal XLD_VENCIMENTO As Date, ByVal XLF_VALORHISTORICO As Double) As Double
    
  Dim XLI_NUMDIAS As Integer
  Dim XLI_NUMMESES As Integer
  Dim XLF_VALORJUROS As Double
    
  XLF_TAXA = XLF_TAXA / 100
  XLI_NUMMESES = (XLD_DATABASE - XLD_VENCIMENTO) \ 30
  
  'Alteração proposto por maurício:
  XLI_NUMMESES = DateDiff("m", XLD_VENCIMENTO, XLD_DATABASE)
  '*If DatePart("d", XLD_DATABASE) > DatePart("d", XLD_VENCIMENTO) Then
  '*  If XLI_NUMMESES > 0 Then
  '*    XLI_NUMMESES = XLI_NUMMESES - 1
  '*  End If
  '*End If
  
  If DatePart("d", XLD_DATABASE) - DatePart("d", XLD_VENCIMENTO) >= 30 Then
    'If XLI_NUMMESES > 0 Then
      XLI_NUMMESES = XLI_NUMMESES + 1
    'End If
  End If
  
  Select Case XLT_TIPOJUROS
    Case "D" 'Diário
      XLI_NUMDIAS = XLD_DATABASE - XLD_VENCIMENTO
      XLF_VALORJUROS = (((1 + XLF_TAXA) ^ XLI_NUMDIAS) - 1) * XLF_VALOR
    Case "M" 'Mensal
      XLF_VALORJUROS = (((1 + XLF_TAXA) ^ XLI_NUMMESES) - 1) * XLF_VALOR
    Case "C" 'Cheio com juros compostos
      'XLI_NUMMESES = DateDiff("m", XLD_VENCIMENTO, XLD_DATABASE) 'Verificar com Maurício. O IF incluido por sua orientação, não atende a todos os casos. Elane
      'If ((XLD_DATABASE - XLD_VENCIMENTO) Mod 30) > 0 Then
      '  XLI_NUMMESES = XLI_NUMMESES + 1
      'End If
      If XLI_NUMMESES = 0 And DatePart("d", XLD_DATABASE) - DatePart("d", XLD_VENCIMENTO) > 0 Then
          XLI_NUMMESES = XLI_NUMMESES + 1
      End If
      XLF_VALORJUROS = (((1 + XLF_TAXA) ^ XLI_NUMMESES) - 1) * XLF_VALORHISTORICO
    Case "S" 'Cheio com juros simples
      'XLI_NUMMESES = DateDiff("m", XLD_VENCIMENTO, XLD_DATABASE) 'Verificar com Maurício. O IF incluido por sua orientação, não atende a todos os casos. Elane
      'If ((XLD_DATABASE - XLD_VENCIMENTO) Mod 30) > 0 Then
      '  XLI_NUMMESES = XLI_NUMMESES + 1
      'End If
      If XLI_NUMMESES = 0 And DatePart("d", XLD_DATABASE) - DatePart("d", XLD_VENCIMENTO) > 0 Then
          XLI_NUMMESES = XLI_NUMMESES + 1
      End If
      XLF_VALORJUROS = (XLF_TAXA * XLI_NUMMESES) * XLF_VALOR
  End Select
  
  FunJuros = XLF_VALORJUROS

End Function

Function FunMulta(ByVal XLF_VALOR As Double, ByVal XLF_TAXA As Double, ByVal XLI_TIPOMULTA As Byte, ByVal XLT_TIPOJUROS As String, ByVal XLD_DATABASE As Date, ByVal XLD_VENCIMENTO As Date, ByVal XLF_VALORHISTORICO As Double, ByVal XLI_LIMITEDIAS As Integer) As Double

  Dim XLI_NUMDIAS As Integer
  Dim XLF_VALORMULTA As Double
  
  XLF_TAXA = XLF_TAXA / 100
  XLI_NUMDIAS = XLD_DATABASE - XLD_VENCIMENTO
  'Se for Pro-rata e juros não é cheio
  If XLI_TIPOMULTA = "1" And XLI_NUMDIAS <= 30 And XLT_TIPOJUROS <> "C" Then
    XLF_VALORMULTA = (((1 + XLF_TAXA) ^ (1 / 30) ^ XLI_NUMDIAS) - 1) * XLF_VALOR
  Else
    If XLI_TIPOMULTA = 0 Then 'Multa Normal
      If XLT_TIPOJUROS = "C" Then
        'Quando o juros é cheio o valor deve ser sempre sobre o valor histórico
        XLF_VALORMULTA = XLF_TAXA * XLF_VALORHISTORICO
      Else
        XLF_VALORMULTA = XLF_TAXA * XLF_VALOR
      End If
    Else 'Multa crescente
      If XLI_NUMDIAS < XLI_LIMITEDIAS Then 'Se for menor que o limite de dias de crecimento
        XLF_TAXA = XLF_TAXA * XLI_NUMDIAS
      Else 'se for maior ou igual ao limite de dias
        XLF_TAXA = XLF_TAXA * XLI_LIMITEDIAS
      End If
      XLF_VALORMULTA = XLF_TAXA * XLF_VALOR
    End If
  End If
  
  FunMulta = XLF_VALORMULTA
  
End Function

Function FunDesagio(ByVal XLF_VALOR As Double, ByVal XLF_TAXA As Double, ByVal XLT_TIPO As String, ByVal XLD_DATABASE As Date, ByVal XLD_VENCIMENTO As Date) As Double

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

'Juros de Financiamento da Caixa Econômica Federal
Function FunJurosCEF(ByVal XLF_VALOR As Double, ByVal XLF_TAXAMENSAL As Double, ByVal XLF_TAXATITULO As Double, ByVal XLD_DATABASE As Date, ByVal XLD_VENCIMENTO As Date) As Double
    
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
    
  FunJurosCEF = XLF_VALOR * XLF_TAXA

End Function

'Retorna o valor indexado
Function FunValorIndexado(ByVal XLF_VALORTITULO As Double, ByVal XLF_COTACAO As Double, ByVal XLF_COTACAOANTERIOR As Double, ByVal XLT_INDEXACAO As String, ByVal XLT_PRORATA As String, ByVal XLD_DATABASE As Date) As Double
  
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

'Retorna o valor indexado por uma determinada moeda
Function FunCalcularValorIndexado(ByVal XLF_VALORREAL As Double, ByVal XLD_DTBASE As Date, ByVal XLI_CODMOEDA As Integer, ByVal XLT_INDEXACAO As String, ByVal XLT_PRORATA As String, XLO_TELACOTACAO As Form)
         
    Dim XLO_COTACAO As New adodb.Recordset
    Dim XLO_COTACAOANTERIOR As New adodb.Recordset
    Dim XLO_COTACAOPOSTERIOR As New adodb.Recordset
    Dim XLF_COTACAO As Double
    Dim XLF_COTACAOANTERIOR As Double
    Dim XLF_COTACAOPOSTERIOR As Double
    Dim XLF_PERCENT As Double
    Dim XLD_DATA As Date
    
    If XLT_INDEXACAO = "D" Then
        SubQOpenRecordset XLO_COTACAO, "SELECT cota_vl_Valor FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData(XLD_DTBASE, NomeSgbd) & "", Estatico
        XLD_DATA = Format(XLD_DTBASE, "dd/mm/yy")
    Else
        SubQOpenRecordset XLO_COTACAO, "SELECT cota_vl_Valor FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData("01/" & Month(XLD_DTBASE) & "/" & Year(XLD_DTBASE), NomeSgbd) & "", Estatico
        XLD_DATA = Format(XLD_DTBASE, "\01/mm/yy")
    End If
    
    If Not XLO_COTACAO.EOF Then
        XLF_COTACAO = XLO_COTACAO!cota_vl_valor
            
        'Se for diário ou mensal e não prorata
        If XLT_INDEXACAO = "D" Or (XLT_INDEXACAO = "M" And XLT_PRORATA = "0") Then
            XLO_COTACAO.Close
            Set XLO_COTACAO = Nothing

            FunCalcularValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
        'se for mensal prorata
        Else
            If XLT_PRORATA = "M" Then 'Pro-Rata Mista
                XLO_COTACAO.Close
                Set XLO_COTACAO = Nothing
    
                FunCalcularValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
            Else
                If XLT_PRORATA = "A" Then
                    XLD_DTBASE = DateAdd("m", -1, XLD_DTBASE)
                Else
                    XLD_DTBASE = DateAdd("m", 1, XLD_DTBASE)
                End If
              
                SubQOpenRecordset XLO_COTACAOANTERIOR, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData("01/" & Month(XLD_DTBASE) & "/" & Year(XLD_DTBASE), NomeSgbd) & "", Estatico
                If Not XLO_COTACAOANTERIOR.EOF Then
                    XLF_COTACAOANTERIOR = XLO_COTACAOANTERIOR!cota_vl_valor
                  
                    If XLT_PRORATA = "A" Then
                       XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
                    Else
                       XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAO
                    End If
                  
                    XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASE) - 1)) * XLF_COTACAO)
                    
                    XLO_COTACAO.Close
                    Set XLO_COTACAO = Nothing
                    XLO_COTACAOANTERIOR.Close
                    Set XLO_COTACAOANTERIOR = Nothing
                    FunCalcularValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
                Else
                  
                    If MsgBox("Deseja cadastrar a cotação para a data " & XLD_DATA & "?", vbInformation + vbYesNo, "Atualização de cotação") = vbYes Then
                  
                        XLO_TELACOTACAO.LblChave.Caption = XLI_CODMOEDA
                        If XLT_INDEXACAO = "D" Then
                            XLO_TELACOTACAO.LblData.Caption = Format(XLD_DTBASE, "DD/MM/YY")
                        Else
                            XLO_TELACOTACAO.LblData.Caption = "01/" & Format(XLD_DTBASE, "MM/YY")
                        End If
                    
                        XLO_TELACOTACAO.Show 1
                        XLO_COTACAOANTERIOR.Requery
                        If Not XLO_COTACAOANTERIOR.EOF Then
                            XLF_COTACAOANTERIOR = XLO_COTACAOANTERIOR!cota_vl_valor
                          
                            If XLT_PRORATA = "A" Then
                                XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
                            Else
                                XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAO
                            End If
                          
                            XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASE) - 1)) * XLF_COTACAO)
                            
                            XLO_COTACAO.Close
                            Set XLO_COTACAO = Nothing
                            XLO_COTACAOANTERIOR.Close
                            Set XLO_COTACAOANTERIOR = Nothing
                            
                            FunCalcularValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
                        End If
                    Else
                        'FunCalcularValorIndexado = 0
                        Exit Function
                    End If
                End If
            End If
        End If
    Else
        XLO_TELACOTACAO.LblChave.Caption = XLI_CODMOEDA
        If XLT_INDEXACAO = "D" Then
            XLO_TELACOTACAO.LblData.Caption = Format(XLD_DTBASE, "DD/MM/YY")
        Else
            XLO_TELACOTACAO.LblData.Caption = "01/" & Format(XLD_DTBASE, "MM/YY")
        End If
      
        If MsgBox("Deseja cadastrar a cotação para a data " & XLD_DATA & "?", vbInformation + vbYesNo, "Atualização de cotação") = vbYes Then
            XLO_TELACOTACAO.Show 1
            XLO_COTACAO.Requery
            If Not XLO_COTACAO.EOF Then
                XLF_COTACAO = XLO_COTACAO!cota_vl_valor
                'Se for diário ou mensal e não prorata
                If XLT_INDEXACAO = "D" Or (XLT_INDEXACAO = "M" And XLT_PRORATA = "0") Then
                    XLO_COTACAO.Close
                    Set XLO_COTACAO = Nothing
                    FunCalcularValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
                'se for mensal prorata
                Else
                    If XLT_PRORATA = "A" Then
                        XLD_DTBASE = DateAdd("m", -1, XLD_DTBASE)
                    Else
                        XLD_DTBASE = DateAdd("m", 1, XLD_DTBASE)
                    End If
        
                    SubQOpenRecordset XLO_COTACAOANTERIOR, "SELECT * FROM ConsGENCotacaoMoeda WHERE moed_cd_codmoeda = " & XLI_CODMOEDA & " and cota_dt_datacotacao =" & FunNuloData("01/" & Month(XLD_DTBASE) & "/" & Year(XLD_DTBASE), NomeSgbd) & "", Estatico
                    If Not XLO_COTACAOANTERIOR.EOF Then
                        XLF_COTACAOANTERIOR = XLO_COTACAOANTERIOR!cota_vl_valor
                      
                        If XLT_PRORATA = "A" Then
                          XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
                        Else
                          XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAO
                        End If
            
                        XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASE) - 1)) * XLF_COTACAO)
                                            
                        XLO_COTACAO.Close
                        Set XLO_COTACAO = Nothing
                        XLO_COTACAOANTERIOR.Close
                        Set XLO_COTACAOANTERIOR = Nothing

                        FunCalcularValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
                    Else
                        If MsgBox("Deseja cadastrar a cotação para a data " & XLD_DATA & "?", vbInformation + vbYesNo, "Atualização de cotação") = vbYes Then
                      
                            XLO_TELACOTACAO.LblChave.Caption = XLI_CODMOEDA
                            If XLT_INDEXACAO = "D" Then
                                XLO_TELACOTACAO.LblData.Caption = XLD_DTBASE
                            Else
                                XLO_TELACOTACAO.LblData.Caption = "01/" & Format(XLD_DTBASE, "MM/YY")
                            End If
                            XLO_TELACOTACAO.Show 1
                            XLO_COTACAOANTERIOR.Requery
                            If Not XLO_COTACAOANTERIOR.EOF Then
                                XLF_COTACAOANTERIOR = XLO_COTACAOANTERIOR!cota_vl_valor
                            
                                If XLT_PRORATA = "A" Then
                                    XLF_PERCENT = XLF_COTACAO / XLF_COTACAOANTERIOR
                                Else
                                    XLF_PERCENT = XLF_COTACAOANTERIOR / XLF_COTACAO
                                End If
                
                                XLF_COTACAO = (((XLF_PERCENT ^ (1 / 30)) ^ (Day(XLD_DTBASE) - 1)) * XLF_COTACAO)
                                
                                XLO_COTACAO.Close
                                Set XLO_COTACAO = Nothing
                                XLO_COTACAOANTERIOR.Close
                                Set XLO_COTACAOANTERIOR = Nothing

                                FunCalcularValorIndexado = Format(XLF_VALORREAL / XLF_COTACAO, "##0.0000")
                            End If
                        Else
                            XLO_COTACAO.Close
                            Set XLO_COTACAO = Nothing
                            'FunCalcularValorIndexado = 0
                            Exit Function
                        End If
                    End If
                End If
            End If
        Else
            'FunCalcularValorIndexado = 0
            XLO_COTACAO.Close
            Set XLO_COTACAO = Nothing
            Exit Function
        End If
    End If
    

End Function

Function FunVerificaVencimento(XLD_DATABASE As Date, XLD_VENCIMENTO As Date) As Boolean
    
  FunVerificaVencimento = True
  If XLD_DATABASE > XLD_VENCIMENTO Then
    If Weekday(XLD_DATABASE) = 2 Then
      If (Weekday(XLD_VENCIMENTO) = 1 And (DateDiff("d", XLD_VENCIMENTO, XLD_DATABASE) = 1) _
        Or Weekday(XLD_VENCIMENTO) = 7 And DateDiff("d", XLD_VENCIMENTO, XLD_DATABASE) = 2) Then
          FunVerificaVencimento = False
      End If
    End If
  Else
    FunVerificaVencimento = False
  End If
  
End Function

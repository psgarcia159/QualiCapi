Attribute VB_Name = "FuncoesGenericas"
Option Explicit
Global Filtros() As String ' Array para armazenar os Filtros da Tabela
Global ContFil As Integer ' Contador que armazena o número de filtros ativos
Global Filtrou As Boolean ' Informa se Algum campo foi filtrado
Global Ordem As String ' Nome do Campo que tem a chave da Tabela ativa
Global Campos_Obrig As String ' Lista dos campos que são obrigatórios na tabela
Global Formulario As Form ' Indica o Formulario Ativo
'Global TipoVar As String ' Tipo do campo a ser filtrado
Global FiltroOrig As String ' Indica os filtros determinados pelo próprio sistema
Global FiltroAtual As String ' Indica o filtro atual da tabela corrente
Global filtra_especial As Integer ' Indica se a tabela utilizada possui filtros pré-definidos
                                               ' pelo sistema em FiltroOrig
                                                                                
Global Tem_Filtro As Boolean ' Informa se já foi relizado um filtro na tabela pelo usuário
Global ConTelaOrdem As Integer ' Indica o número de chaves utilizados para ordenação
Global Ordens() As Integer    ' Vetor que mantem a ordem dos campos
Global Vetor() As Variant ' Vetor para auxílio na criação da tabela padrao
Global Quantid As Long  ' Quantidade de registros da tabela

'* Funcoes de manipulacao de arquivos INI
' Leitura de arquivo INI
Declare Function GetPrivateProfileString _
   Lib "kernel32" _
   Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                     ByVal lpKeyName As Any, _
                                     ByVal lpDefault As String, _
                                     ByVal lpReturnedString As String, _
                                     ByVal nSize As Long, _
                                     ByVal lpFileName As String) As Long
' Gravação em arquivo INI
Declare Function WritePrivateProfileString _
   Lib "kernel32" _
   Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                       ByVal lpKeyName As Any, _
                                       ByVal lpString As Any, _
                                       ByVal lpFileName As String) As Long




Public Function FunTrataString(XVariavel As String) As String
    Dim Cont As Integer
    Dim XResult As String
    
    Cont = 1
    While (Cont <= Len(Trim(XVariavel)))
        If (Mid(XVariavel, Cont, 1) <> "'") Then
            XResult = XResult + Mid(XVariavel, Cont, 1)
        Else
            XResult = XResult + "''"
        End If

        Cont = Cont + 1
    Wend
    
    FunTrataString = XResult
End Function


Function FunPegaGuardaUltimo(Tipo As String) As Double
    Dim ResUltimoDoc As Object
    Dim Cont As Integer
    
    SubQOpenRecordset ResUltimoDoc, "SELECT * FROM Ultimos", Dinamico
    
    Select Case Tipo
        Case "D" 'Número Documento
            FunPegaGuardaUltimo = ResUltimoDoc!ulti_nr_numdoc + 1
            ResUltimoDoc!ulti_nr_numdoc = (ResUltimoDoc!ulti_nr_numdoc + 1)
        Case "F" 'Número da Fatura
            FunPegaGuardaUltimo = ResUltimoDoc!ulti_nr_fatura + 1
            ResUltimoDoc!ulti_nr_fatura = (ResUltimoDoc!ulti_nr_fatura + 1)
        Case "C" 'Número de Controle
            FunPegaGuardaUltimo = ResUltimoDoc!ulti_nr_controle + 1
            ResUltimoDoc!ulti_nr_controle = (ResUltimoDoc!ulti_nr_controle + 1)
        Case "B" 'Número de Borderô
            FunPegaGuardaUltimo = ResUltimoDoc!ulti_nr_bordero + 1
            ResUltimoDoc!ulti_nr_bordero = (ResUltimoDoc!ulti_nr_bordero + 1)
        Case "E" 'Número de movimento do Estoque
            FunPegaGuardaUltimo = ResUltimoDoc!ulti_nr_MovEstoque + 1
            ResUltimoDoc!ulti_nr_MovEstoque = (ResUltimoDoc!ulti_nr_MovEstoque + 1)
        Case "R" 'Número de Recibo de Receita
            FunPegaGuardaUltimo = ResUltimoDoc!ulti_nr_ReciboReceita + 1
            ResUltimoDoc!ulti_nr_ReciboReceita = (ResUltimoDoc!ulti_nr_ReciboReceita + 1)
        Case "PR" 'Documento Pagto/Rec
            FunPegaGuardaUltimo = ResUltimoDoc!Ulti_nr_DocPagRec + 1
            ResUltimoDoc!Ulti_nr_DocPagRec = (ResUltimoDoc!Ulti_nr_DocPagRec + 1)
        Case "NN" 'Nosso_Numero do Boleto
            FunPegaGuardaUltimo = ResUltimoDoc!Ulti_Nr_BoletoNossoNumero + 1
            ResUltimoDoc!Ulti_Nr_BoletoNossoNumero = (ResUltimoDoc!Ulti_Nr_BoletoNossoNumero + 1)
'        Case "RM" 'Número da Requisição (Estoque)
'            FunPegaGuardaUltimo = ResUltimoDoc!Ulti_nr_Requisicao + 1
'            ResUltimoDoc!Ulti_nr_Requisicao = (ResUltimoDoc!Ulti_nr_Requisicao + 1)
    End Select
    
    ResUltimoDoc.Update
    ResUltimoDoc.Close
    
End Function


' Realiza a chamada do formulário de Filtro, de acordo com o tipo de dados da coluna do grid.
Function FunExecutaFiltroUnbound(VetorCampos As Variant, Vetor As XArray, Grid As TDBGrid, XForm As Form, Posicaox As Single, XLT_SQL As String)
    
    Dim CurCol As Integer 'Receberá o numero da coluna clicada
           
    'If FunTabelaVazia(ControleDados) Then       'Verifica se tabela está vazia
    '   Exit Function
    'End If
    
    CurCol = XForm.TDBGrid1.ColContaining(Posicaox)   'A variável "curcol" contém o numero da coluna clicada
    
    If CurCol >= 0 Then 'Corresponde a um campo qualquer da tabela
       XForm.TDBGrid1.Col = CurCol
       
       If XForm.TDBGrid1.Columns(CurCol).HeadFont.Italic Then
          MsgBox "Campo já se encontra filtrado", 32, "ERRO"
          Exit Function
       End If
       
       'TipoVar = funTipoCampo(ControleDados, XForm.TDBGrid1.Columns, CurCol)
       XGT_CAMPOFILTRO = XForm.TDBGrid1.Columns(CurCol).FooterText
       'If TipoVar = "N" Or TipoVar = "D" Or TipoVar = "Y" Then
       If XGT_CAMPOFILTRO = "N" Or XGT_CAMPOFILTRO = "D" Or XGT_CAMPOFILTRO = "Y" Then
           FiltroNumerico.Label2.Caption = XForm.TDBGrid1.Columns(CurCol).Caption
           'FiltroNumerico.ControleDados() = ControleDados
           FiltroNumerico.Grid() = XForm.TDBGrid1.Columns
           FiltroNumerico.ColunaGrid() = CurCol
           FiltroNumerico.Show 1
       Else
           FiltroCaracter.Label2.Caption = XForm.TDBGrid1.Columns(CurCol).Caption
           'FiltroCaracter.ControleDados() = ControleDados
           FiltroCaracter.Grid() = XForm.TDBGrid1.Columns
           FiltroCaracter.ColunaGrid() = CurCol
           FiltroCaracter.Show 1
       End If

       If Filtrou = True Then
         subCarregaVetor FiltroAtual, VetorCampos, Vetor, Grid
       End If
    End If

End Function


'retorna String limpa do arquivo INI
Function FunStrArqIni(XLT_TEXTO As String) As String
    Dim XLI_INDICE As Integer
    Dim XLT_RESULTADO As String
    XLT_RESULTADO = ""
    For XLI_INDICE = 1 To Len(XLT_TEXTO)
        If Asc(Mid(XLT_TEXTO, XLI_INDICE, 1)) <> 0 Then
            XLT_RESULTADO = XLT_RESULTADO + Mid(XLT_TEXTO, XLI_INDICE, 1)
        End If
    Next
    FunStrArqIni = XLT_RESULTADO
End Function

'Transforma uma string formatada com "," para uma formatacao com "." ex: "10,2" ficará "10.2"
Public Function FunTrataFloat(ByVal XVariavel As String) As String
    If XVariavel = "" Then
        FunTrataFloat = "0"
    Else
        FunTrataFloat = str(CDbl(XVariavel))
    End If
End Function

'Verifica se KeyAscii é um número
Function FunDigito(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        FunDigito = False
    Else
        FunDigito = True
    End If
End Function

'Trata variavel para incluir valor nulo em campo de tabelas
Function FunNulo(texto As Variant) As String
    If IsNull(texto) Or IsEmpty(texto) Or texto = "" Then
        FunNulo = ""
    Else
        FunNulo = texto
    End If
End Function

Function funNull(XLV_Valor As Variant) As Variant
  'Retorna NULL caso o valor passado seja zero ou vazio.
  If (XLV_Valor = "") Or (XLV_Valor = 0) Then
    funNull = Null
  Else
    funNull = XLV_Valor
  End If
End Function

Function funNuloDataHora(texto As Variant, Banco As String, Tipo As String) As String
    'xTipo define se a hora é de inicio ou de fim
    If Not IsDate(texto) Then
        funNuloDataHora = "Null"
    ElseIf Banco = "Access" Then
            funNuloDataHora = "#" & Format(CVDate(texto), "MM/DD/YY")
            If Tipo = "I" Then
                funNuloDataHora = funNuloDataHora & " 00:00#"
            Else
                funNuloDataHora = funNuloDataHora & " 23:59#"
            End If
        ElseIf Banco = "Sql Server" Then
            funNuloDataHora = "'" & Format(CVDate(texto), "MM/DD/YY")
            If Tipo = "I" Then
                funNuloDataHora = funNuloDataHora & " 00:00'"
            Else
                funNuloDataHora = funNuloDataHora & " 23:59'"
            End If
    End If
End Function

'Verifica se Datax é uma data válida
Function FunVerificaData(ByVal Datax As String) As Integer
    Dim Mes, Dia, Ano, DiaAux As Integer
    
    Datax = Trim$(Datax)
    Dia = Trim$(Left$(Datax, 2))
    Mes = Trim$(Mid$(Datax, 4, 2))
    Ano = Trim$(Mid$(Datax, 7, 2))
    
    If Len(Dia) <> 2 Or Len(Mes) <> 2 Or Len(Ano) <> 2 Then 'Se tamanho diferente de 8
       MsgBox "Formato Correto: DD/MM/AA", 16, "DATA INVÁLIDA"
       FunVerificaData = False
       Exit Function
    End If
    
    If Dia < "01" Or Dia > "31" Then
       MsgBox "Dia Inválido !", 16, "DATA INVÁLIDA"
       FunVerificaData = False
       Exit Function
    End If
    
    Select Case Mes
           Case "01", "03", "05", "07", "08", "10", "12"
                If Dia < "01" Or Dia > "31" Then
                   MsgBox "Dia deve ser de 01 a 31", 16, "DATA INVÁLIDA"
                   FunVerificaData = False
                   Exit Function
                End If
           Case "04", "06", "09", "11"
                If Dia < "01" Or Dia > "30" Then
                   MsgBox "Dia deve ser de 01 a 30", 16, "DATA INVÁLIDA"
                   FunVerificaData = False
                   Exit Function
                End If
           Case "02"
                DiaAux = 28
                'Verifica se o ano é bissexto
                If (Val(Ano) Mod 4) = 0 Then
                   DiaAux = 29
                End If
                If Dia < "01" Or Dia > Format(DiaAux, "00") Then
                   MsgBox "Dia deve ser de 01 a " + Format(DiaAux, "00"), 16, "DATA INVÁLIDA"
                   FunVerificaData = False
                   Exit Function
                End If
           Case Else
                MsgBox "Mês deve ser de 01 a 12", 16, "DATA INVÁLIDA"
                FunVerificaData = False
                Exit Function
    End Select
    
    If Ano < "00" Or Ano > "99" Then
       MsgBox "Ano Inválido !", 16, "DATA INVÁLIDA"
       FunVerificaData = False
       Exit Function
    End If
    
    FunVerificaData = True

End Function


'Torna o preenchimento do combo obrigatório
Function FunObrigatorioCBO(combo As DataCombo, texto As String) As Boolean
    If combo.Text = "" Then
        MsgBox texto, vbCritical, "CAMPO OBRIGATÓRIO"
        FunObrigatorioCBO = True
        If combo.Enabled = True Then combo.SetFocus
    End If
End Function

Function FunNuloData(texto As Variant, Banco As String) As String
    If Not IsDate(texto) Then
        FunNuloData = "Null"
    ElseIf Banco = "Access" Then
        FunNuloData = "#" & Format(CVDate(texto), "MM/DD/YY") & "#"
    ElseIf Banco = "Sql Server" Then
        FunNuloData = "'" & Format(CVDate(texto), "MM/DD/YY") & "'"
    Else
        FunNuloData = "'" & Format(CVDate(texto), "MM/DD/YY") & "'"
    End If
End Function

Function FunNuloVal(texto As Variant) As String
    If IsNull(texto) Or IsEmpty(texto) Or texto = "" Then
        FunNuloVal = 0
    Else
        FunNuloVal = texto
    End If
End Function

Function FunNuloBancoVal(texto As Variant) As String
    If IsNull(texto) Or texto = Empty Then
        FunNuloBancoVal = "Null"
    Else
        FunNuloBancoVal = Trim(str(texto))
    End If
End Function

Function FunNuloBanco(texto As Variant) As String
    If IsNull(texto) Or texto = Empty Then
        FunNuloBanco = "Null"
    Else
        FunNuloBanco = Trim(str(texto))
    End If
End Function

' conferências dos dígitos verificadores  do CPF
Function FunConfereCPF(CPF As String) As Boolean ' retorna True quando OK
    Dim Soma As Integer
    Dim Dígito1, Dígito2 As Integer
    Dim Resto As Integer
    
    FunConfereCPF = False ' considera errado a priori
            
    If Not IsNumeric(CPF) Or Len(CPF) <> 11 Or _
        CPF = "11111111111" Or CPF = "22222222222" Or CPF = "33333333333" Or _
        CPF = "44444444444" Or CPF = "55555555555" Or CPF = "66666666666" Or _
        CPF = "77777777777" Or CPF = "88888888888" Or CPF = "99999999999" Or _
        CPF = "00000000000" Then
       Exit Function
    End If
    'calcula o primeiro dígito
    'soma até o oitavo algarismo
    Soma = 0
    Soma = Soma + Val(Mid(CPF, 9, 1)) * 2
    Soma = Soma + Val(Mid(CPF, 8, 1)) * 3
    Soma = Soma + Val(Mid(CPF, 7, 1)) * 4
    Soma = Soma + Val(Mid(CPF, 6, 1)) * 5
    Soma = Soma + Val(Mid(CPF, 5, 1)) * 6
    Soma = Soma + Val(Mid(CPF, 4, 1)) * 7
    Soma = Soma + Val(Mid(CPF, 3, 1)) * 8
    Soma = Soma + Val(Mid(CPF, 2, 1)) * 9
    Soma = Soma + Val(Mid(CPF, 1, 1)) * 10
     
    Resto = Soma Mod 11
    Dígito1 = IIf(Resto = 0 Or Resto = 1, 0, 11 - Resto)
    
    'calcula o segundo dígito
    Soma = 0
    Soma = Soma + Val(Mid(CPF, 10, 1)) * 2
    Soma = Soma + Val(Mid(CPF, 9, 1)) * 3
    Soma = Soma + Val(Mid(CPF, 8, 1)) * 4
    Soma = Soma + Val(Mid(CPF, 7, 1)) * 5
    Soma = Soma + Val(Mid(CPF, 6, 1)) * 6
    Soma = Soma + Val(Mid(CPF, 5, 1)) * 7
    Soma = Soma + Val(Mid(CPF, 4, 1)) * 8
    Soma = Soma + Val(Mid(CPF, 3, 1)) * 9
    Soma = Soma + Val(Mid(CPF, 2, 1)) * 10
    Soma = Soma + Val(Mid(CPF, 1, 1)) * 11
     
    Resto = Soma Mod 11
    Dígito2 = IIf(Resto = 0 Or Resto = 1, 0, 11 - Resto)
       
    If Dígito1 = Val(Mid(CPF, 10, 1)) And Dígito2 = Val(Mid(CPF, 11, 1)) Then
        FunConfereCPF = True
    End If
End Function

' Confere o digito verificador dod GCG
Function FunConfereCGC(CGC As String) As Boolean  ' retorna True quando OK
    Dim Soma As Integer
    Dim Dígito1, Dígito2 As Integer
    Dim Resto As Integer
    
    FunConfereCGC = False ' considera errado a priori
        
    If Len(CGC) <> 14 Then
        Exit Function
    End If
    ' calcula o primeiro dígito
    Soma = 0
    Soma = Soma + Val(Mid(CGC, 1, 1)) * 5
    Soma = Soma + Val(Mid(CGC, 2, 1)) * 4
    Soma = Soma + Val(Mid(CGC, 3, 1)) * 3
    Soma = Soma + Val(Mid(CGC, 4, 1)) * 2
    Soma = Soma + Val(Mid(CGC, 5, 1)) * 9
    Soma = Soma + Val(Mid(CGC, 6, 1)) * 8
    Soma = Soma + Val(Mid(CGC, 7, 1)) * 7
    Soma = Soma + Val(Mid(CGC, 8, 1)) * 6
    Soma = Soma + Val(Mid(CGC, 9, 1)) * 5
    Soma = Soma + Val(Mid(CGC, 10, 1)) * 4
    Soma = Soma + Val(Mid(CGC, 11, 1)) * 3
    Soma = Soma + Val(Mid(CGC, 12, 1)) * 2
    
    
    Resto = Soma Mod 11
    Dígito1 = IIf(Resto = 0 Or Resto = 1, 0, 11 - Resto)

    ' calcula o segundo dígito
    Soma = 0
    Soma = Soma + Val(Mid(CGC, 1, 1)) * 6
    Soma = Soma + Val(Mid(CGC, 2, 1)) * 5
    Soma = Soma + Val(Mid(CGC, 3, 1)) * 4
    Soma = Soma + Val(Mid(CGC, 4, 1)) * 3
    Soma = Soma + Val(Mid(CGC, 5, 1)) * 2
    Soma = Soma + Val(Mid(CGC, 6, 1)) * 9
    Soma = Soma + Val(Mid(CGC, 7, 1)) * 8
    Soma = Soma + Val(Mid(CGC, 8, 1)) * 7
    Soma = Soma + Val(Mid(CGC, 9, 1)) * 6
    Soma = Soma + Val(Mid(CGC, 10, 1)) * 5
    Soma = Soma + Val(Mid(CGC, 11, 1)) * 4
    Soma = Soma + Val(Mid(CGC, 12, 1)) * 3
    Soma = Soma + Dígito1 * 2
    
    Resto = Soma Mod 11
    Dígito2 = IIf(Resto = 0 Or Resto = 1, 0, 11 - Resto)
    
    If Dígito1 = Val(Mid(CGC, 13, 1)) And Dígito2 = Val(Mid(CGC, 14, 1)) Then
        FunConfereCGC = True
    End If
End Function

'Esta função recebe uma string e a varre retirando os acentos e caracteres especiais
'os caracteres especiais recebem " " ( vazio ) e os acentuados ficam sem acento
'o retorno desta função é a String tratada
Function funTiraAcento(XLT_TEXTO As String, XLT_UPCASE As Boolean) As String
  
  Dim Cont As Integer
  Dim XLT_Result As String
  Dim XLT_LETRA As String
  Dim XLL_ASCIILETRA As Long
  
  Cont = 1
  While (Cont <= Len(XLT_TEXTO))
    XLT_LETRA = (Mid(XLT_TEXTO, Cont, 1))
    XLL_ASCIILETRA = Asc(XLT_LETRA)
    If ((XLL_ASCIILETRA >= 224) And (XLL_ASCIILETRA <= 229)) _
    Or ((XLL_ASCIILETRA >= 192) And (XLL_ASCIILETRA <= 197)) Then
      XLT_Result = XLT_Result + "A"
    Else
      If ((XLL_ASCIILETRA >= 200) And (XLL_ASCIILETRA <= 203)) _
      Or ((XLL_ASCIILETRA >= 232) And (XLL_ASCIILETRA <= 235)) Then
        XLT_Result = XLT_Result + "E"
      Else
        If ((XLL_ASCIILETRA >= 204) And (XLL_ASCIILETRA <= 207)) _
        Or ((XLL_ASCIILETRA >= 236) And (XLL_ASCIILETRA <= 239)) Then
          XLT_Result = XLT_Result + "I"
        Else
          If ((XLL_ASCIILETRA >= 210) And (XLL_ASCIILETRA <= 214)) _
          Or ((XLL_ASCIILETRA >= 242) And (XLL_ASCIILETRA <= 246)) Then
            XLT_Result = XLT_Result + "O"
          Else
            If ((XLL_ASCIILETRA >= 217) And (XLL_ASCIILETRA <= 220)) _
            Or ((XLL_ASCIILETRA >= 249) And (XLL_ASCIILETRA <= 252)) Then
              XLT_Result = XLT_Result + "U"
            Else
              If ((XLL_ASCIILETRA = 199) Or (XLL_ASCIILETRA = 231)) Then
                XLT_Result = XLT_Result + "C"
              Else
                If ((XLL_ASCIILETRA >= 33) And (XLL_ASCIILETRA <= 47) And (XLL_ASCIILETRA <> 46)) _
                Or ((XLL_ASCIILETRA >= 58) And (XLL_ASCIILETRA <= 64)) _
                Or ((XLL_ASCIILETRA >= 91) And (XLL_ASCIILETRA <= 96)) _
                Or ((XLL_ASCIILETRA >= 123) And (XLL_ASCIILETRA <= 126)) _
                Or ((XLL_ASCIILETRA >= 161) And (XLL_ASCIILETRA <= 191)) _
                Or ((XLL_ASCIILETRA >= 215) And (XLL_ASCIILETRA <= 216)) _
                Or ((XLL_ASCIILETRA >= 222) And (XLL_ASCIILETRA <= 223)) _
                Or ((XLL_ASCIILETRA >= 217) And (XLL_ASCIILETRA <= 220)) _
                Or ((XLL_ASCIILETRA >= 247) And (XLL_ASCIILETRA <= 248)) _
                Or ((XLL_ASCIILETRA >= 217) And (XLL_ASCIILETRA <= 220)) _
                Or (XLL_ASCIILETRA = 198) Or (XLL_ASCIILETRA = 208) Or (XLL_ASCIILETRA = 230) Or (XLL_ASCIILETRA = 241) Then
                  XLT_Result = XLT_Result + " " 'caracteres especiais
                Else
                  If XLT_UPCASE = True Then
                    XLT_Result = XLT_Result & UCase(XLT_LETRA)
                  Else
                    XLT_Result = XLT_Result & UCase(XLT_LETRA)
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End If
    Cont = Cont + 1
  Wend
  funTiraAcento = XLT_Result
End Function

Function FunZeros(XLI_QTD As Integer) As String
  
   Dim XLT_ZEROS As String
   Dim XLI_CONT As Integer
      
   For XLI_CONT = 1 To XLI_QTD
     XLT_ZEROS = XLT_ZEROS & "0"
   Next
   
   FunZeros = XLT_ZEROS
   
End Function


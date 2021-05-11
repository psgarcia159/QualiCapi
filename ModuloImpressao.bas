Attribute VB_Name = "ModuloImpressao"
Option Explicit
Global AltRodape As Single       'Variável usada para calcular tamanho da página (refere-se a altura ocupada pelo rodapé)
Global AltCabecalho As Single       'Variável usada para calcular tamanho da página (refere-se a altura ocupada pelo cabeçalho)

'Constantes para a função GetDeviceCaps
Global Const PHYSICALOFFSETX As Single = 112
Global Const PHYSICALOFFSETY As Single = 113
Global Const PHYSICALWIDTH As Single = 110
Global Const PHYSICALHEIGHT As Single = 111

'Função que pega as propriedades da impressora
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Dim a As CRAXDRT.FieldObject


Sub subImprimeListagemGRID(xOrientacao As Integer, xGrid As TrueOleDBGrid70.PrintInfo, Texto As String)
    With xGrid
        ' Setar margens a depender do tipo de impressão
        If xOrientacao = 1 Then
            .SettingsOrientation = 1
            .SettingsMarginTop = 567
            '.SettingsMarginBottom = 630
            .SettingsMarginBottom = 750
            '.SettingsMarginLeft = 1134
            .SettingsMarginLeft = 750
            .SettingsMarginRight = 567
        Else
            .SettingsOrientation = 2
            .SettingsMarginTop = 1134
            '.SettingsMarginBottom = 630
            .SettingsMarginBottom = 750
            .SettingsMarginLeft = 567
            '.SettingsMarginRight = 567
            .SettingsMarginRight = 850
       End If
        
        'Determina se o cabeçalho será definido pelo usuário e o seu tamanho
        'Neste caso o cabeçalho será definido pelo usuário através de API
        '.PageHeaderOwnerDraw = True
        '.PageHeaderHeight = 4000
        
        ' Setar o cabeçalho
        .PageHeaderFont.Name = "Arial"
        .PageHeaderFont.Size = 12
        .PageHeaderFont.Bold = True
        .PageHeader = Texto + " \t\t" + CStr(Now)
        
        ' O Cabeçalho e o rodapé devem aparecer em todas as páginas
        .RepeatColumnHeaders = True
        .RepeatColumnFooters = True
        
        ' Imprimir texto no rodapé
        .PageFooter = "QUALIDADOS\t\tPágina: \p/\P"
        
        ' Chamar o Preview de impressão
        .PageSetup
        '.PrintPreview
    End With
    If Not xGrid.PageSetupCancelled Then
        xGrid.PrintPreview
    End If
    

End Sub

'Sub subImprimeListagemGRIDUnBound(xOrientacao As Integer, xGrid As TrueDBGrid70.PrintInfo, Texto As String)
Sub subImprimeListagemGRIDUnBound(xOrientacao As Integer, xGrid As Object, Texto As String)
    With xGrid
        ' Setar margens a depender do tipo de impressão
        If xOrientacao = 1 Then
            .SettingsOrientation = 1
            .SettingsMarginTop = 567
            '.SettingsMarginBottom = 630
            .SettingsMarginBottom = 750
            '.SettingsMarginLeft = 1134
            .SettingsMarginLeft = 750
            .SettingsMarginRight = 567
        Else
            .SettingsOrientation = 2
            .SettingsMarginTop = 1134
            '.SettingsMarginBottom = 630
            .SettingsMarginBottom = 750
            .SettingsMarginLeft = 567
            '.SettingsMarginRight = 567
            .SettingsMarginRight = 850
       End If
        
        'Determina se o cabeçalho será definido pelo usuário e o seu tamanho
        'Neste caso o cabeçalho será definido pelo usuário através de API
        '.PageHeaderOwnerDraw = True
        '.PageHeaderHeight = 4000
        
        ' Setar o cabeçalho
        .PageHeaderFont.Name = "Arial"
        .PageHeaderFont.Size = 12
        .PageHeaderFont.Bold = True
        .PageHeader = Texto + " \t\t" + CStr(Now)
        
        ' O Cabeçalho e o rodapé devem aparecer em todas as páginas
        .RepeatColumnHeaders = True
        .RepeatColumnFooters = True
        
        ' Imprimir texto no rodapé
        .PageFooter = "QUALIDADOS\t\tPágina: \p/\P"
        
        ' Chamar o Preview de impressão
        .PageSetup
        '.PrintPreview
    End With
    If Not xGrid.PageSetupCancelled Then
        xGrid.PrintPreview
    End If
    

End Sub

Function funTamanhoTextoPrinter(Texto As String, Escala As Integer, Fonte As String, TamLetra As Long) As Single
    Printer.ScaleMode = Escala
    Printer.FontName = Fonte
    Printer.FontSize = TamLetra
    funTamanhoTextoPrinter = Printer.TextWidth(Texto)
End Function
Sub subQuebraTexto(Texto As String, XInicio As Single, TamCampo As Single, TamLetra As Long, Fonte As String, Negrito As Boolean, Italico As Boolean)
    Dim Resto As String
    Dim ProcuraEspaco As Integer
        
    'Para Garantir Que Vai Verificar o Tamanho Com a Formatação Correta
    Printer.Font.Name = Fonte
    Printer.Font.Size = TamLetra
    Printer.Font.Bold = Negrito
    Printer.Font.Italic = Italico
    
    Resto = ""
    Texto = Trim$(Texto)
    'Texto é uma variável que indica a string a ser impressa, após a impressão o seu valor é
    'substituído pelo valor de Resto, se não existir Resto significa que já foi impresso toda a
    'string, consequentemente o valor de texto também será vazio.
    
    'Se Texto for vazio imprime vazio e sai
    If Texto = "" Then
        subImprimeTexto Texto, XInicio, Printer.CurrentY, TamLetra, Fonte, Negrito, Italico
        'subImprimeTexto Texto, XInicio, YInicio, TamLetra, Fonte, Negrito, Italico
        Exit Sub
    End If
    
    While Texto <> ""
        'Enquanto o tamanho do Texto for maior que o tamanho do espaço reservado para ele
        'será retirado o último caractere do Texto e este será acrescido a variável Resto
        While Printer.TextWidth(Texto) > TamCampo
            Resto = Mid$(Texto, Len(Texto), 1) + Resto
            Texto = Mid$(Texto, 1, (Len(Texto) - 1))
        Wend
        
'*********************************************************************************************************
'Esta parte da sub é para impedir que se quebre uma palavra no meio
'*********************************************************************************************************
        'Se Resto for igual a vazio significa que não existirá uma próxima linha
        If Resto <> "" Then
            If Mid$(Resto, 1, 1) = " " Then 'Resto possui espaço no início quebrou uma palavra completa
                Resto = LTrim$(Resto)
            ElseIf Mid$(Texto, Len(Texto), 1) = " " Then 'Texto possui espaço no fim
                Texto = RTrim$(Texto)                           'quebrou uma palavra completa
            Else 'Quebrou metade de uma palavra
                ProcuraEspaco = Len(Texto)
                'Verifica se Texto possui algum espaço para quebrar a palavra completa
                While Mid$(Texto, ProcuraEspaco, 1) <> " " And ProcuraEspaco <> 1
                    ProcuraEspaco = ProcuraEspaco - 1
                Wend
                If ProcuraEspaco <> 1 Then
                    Resto = Right$(Texto, Len(Texto) - ProcuraEspaco) & Resto
                    Texto = RTrim$(Left(Texto, ProcuraEspaco))
                End If
            End If
        End If
'*********************************************************************************************************
        subImprimeTexto Texto, XInicio, Printer.CurrentY, TamLetra, Fonte, Negrito, Italico
        Texto = Resto
        Resto = ""
    Wend
End Sub
'Sub subQuebraTextoLaser(Texto As String, XInicio As Single, TamCampo As Single, TamLetra As Long, Fonte As String, Negrito As Boolean, Italico As Boolean)
'    Dim Resto As String
'    Dim ProcuraEspaco As Integer
'
'    'Para Garantir Que Vai Verificar o Tamanho Com a Formatação Correta
'    a.Font.Name = Fonte
'    a.Font.Size = TamLetra
'    a.Font.Bold = Negrito
'    a.Font.Italic = Italico
'
'    Resto = ""
'    Texto = Trim$(Texto)
'    'Texto é uma variável que indica a string a ser impressa, após a impressão o seu valor é
'    'substituído pelo valor de Resto, se não existir Resto significa que já foi impresso toda a
'    'string, consequentemente o valor de texto também será vazio.
'
'    'Se Texto for vazio imprime vazio e sai
'    If Texto = "" Then
'        subImprimeTextoLaser Texto, XInicio, a.Left, TamLetra, Fonte, Negrito, Italico
'        'subImprimeTexto Texto, XInicio, YInicio, TamLetra, Fonte, Negrito, Italico
'        Exit Sub
'    End If
'
'    While Texto <> ""
'        'Enquanto o tamanho do Texto for maior que o tamanho do espaço reservado para ele
'        'será retirado o último caractere do Texto e este será acrescido a variável Resto
'        While a.Width(Texto) > TamCampo
'            Resto = Mid$(Texto, Len(Texto), 1) + Resto
'            Texto = Mid$(Texto, 1, (Len(Texto) - 1))
'        Wend
'
''*********************************************************************************************************
''Esta parte da sub é para impedir que se quebre uma palavra no meio
''*********************************************************************************************************
'        'Se Resto for igual a vazio significa que não existirá uma próxima linha
'        If Resto <> "" Then
'            If Mid$(Resto, 1, 1) = " " Then 'Resto possui espaço no início quebrou uma palavra completa
'                Resto = LTrim$(Resto)
'            ElseIf Mid$(Texto, Len(Texto), 1) = " " Then 'Texto possui espaço no fim
'                Texto = RTrim$(Texto)                           'quebrou uma palavra completa
'            Else 'Quebrou metade de uma palavra
'                ProcuraEspaco = Len(Texto)
'                'Verifica se Texto possui algum espaço para quebrar a palavra completa
'                While Mid$(Texto, ProcuraEspaco, 1) <> " " And ProcuraEspaco <> 1
'                    ProcuraEspaco = ProcuraEspaco - 1
'                Wend
'                If ProcuraEspaco <> 1 Then
'                    Resto = Right$(Texto, Len(Texto) - ProcuraEspaco) & Resto
'                    Texto = RTrim$(Left(Texto, ProcuraEspaco))
'                End If
'            End If
'        End If
''*********************************************************************************************************
'        subImprimeTextoLaser Texto, XInicio, a.Top, TamLetra, Fonte, Negrito, Italico
'        Texto = Resto
'        Resto = ""
'    Wend
'End Sub

'*************************

Function funTamanhoTexto(Texto As String) As Single
    Dim XSeguidas As Integer, x As Integer, ContLinhas As Integer
    x = 1
    XSeguidas = 1
    ContLinhas = 1
    While x <> Len(Texto)
        If Mid$(Texto, x, 1) = Chr$(13) Then
            ContLinhas = ContLinhas + 1
            XSeguidas = 0
        End If
        If XSeguidas \ 80 <> 0 Then
            ContLinhas = ContLinhas + 1
            XSeguidas = 0
        End If
        x = x + 1
        XSeguidas = XSeguidas + 1
    Wend
    funTamanhoTexto = ContLinhas * 0.4
End Function

Sub subImprimeBox(XInicio As Single, YInicio As Single, XFim As Single, YFim As Single)
    Printer.DrawWidth = 1
    Printer.Line (XInicio, YInicio)-(XFim, YFim), QBColor(0), B
End Sub

Function subRetornaImpNum(Texto As String, XInicio As Single, TamCampo As Single, TamLetra As Long, Fonte As String, Negrito As Boolean, Italico As Boolean) As Single
    
    'Para Garantir Que Vai Verificar o Tamanho Com a Formatação Correta
    Printer.Font.Name = Fonte
    Printer.Font.Size = TamLetra
    Printer.Font.Bold = Negrito
    Printer.Font.Italic = Italico
    subRetornaImpNum = XInicio + (TamCampo - Printer.TextWidth(Texto))
End Function
Sub subRodape(XInicio As Single, YInicio As Single, ByVal NomeRelatorio As String)
    Dim IniciaImp As Single
    AltRodape = 2.5
    Printer.CurrentY = YInicio - AltRodape
    
    If Printer.Orientation = 1 Then
        subImprimeLinha XInicio, 20, 6
        IniciaImp = Printer.CurrentY + 0.2
'        subImprimeTexto NomeRelatorio, XInicio, IniciaImp, 8, "Times New Roman", False, False
        subImprimeTexto "DATA: " & Format(Date, "dd/mm/yy"), 13, IniciaImp, 8, "Times New Roman", False, False
        subImprimeTexto "HORA: " & Format(Time, "HH:MM:SS"), 15.2, IniciaImp, 8, "Times New Roman", False, False
        subImprimeTexto "PÁGINA: " & Printer.Page, 17.5, IniciaImp, 8, "Times New Roman", False, False
    Else
        subImprimeLinha XInicio, 29, 6
        IniciaImp = Printer.CurrentY + 0.2
'        subImprimeTexto NomeRelatorio, XInicio, IniciaImp, 8, "Times New Roman", False, False
        subImprimeTexto "DATA: " & Format(Date, "dd/mm/yy"), 21, IniciaImp, 8, "Times New Roman", False, False
        subImprimeTexto "HORA: " & Format(Time, "HH:MM:SS"), 23.2, IniciaImp, 8, "Times New Roman", False, False
        subImprimeTexto "PÁGINA: " & Printer.Page, 25.5, IniciaImp, 8, "Times New Roman", False, False
    End If
End Sub
Sub subImprimeCelula(XData As Object, IniCaixaX As Single, FimCaixaX As Single, IniCaixaY As Single, Largura As Single, Titulo As String, DistanciaTitCampo As Single, xcampo As String, Negrito As Boolean)
'    Dim FimCaixaY As Single, GuardaY As Single, RecebeCampo As Variant
'
'    FimCaixaY = IniCaixaY + Largura
'    'subImprimeCelula(IniCaixaX, FimCaixaX, IniCaixaY, FimCaixaY, Titulo, Negrito)
'    subImprimeBox IniCaixaX, IniCaixaY, FimCaixaX, FimCaixaY
'    Printer.CurrentY = IniCaixaY + 0.1
'    GuardaY = Printer.CurrentY
'
'    subImprimeTexto Titulo, (IniCaixaX + 0.2), GuardaY, 5, "Times New Roman", False, False
'
'    GuardaY = Printer.CurrentY + DistanciaTitCampo
'    If xcampo = "Titulo" Then
'        Printer.CurrentY = GuardaY
'        RecebeCampo = QFields(XData, "Titulo")
'        CriaVetorObs RecebeCampo, (IniCaixaX + 0.2), 58, 8, Negrito, False
'    ElseIf xcampo = "Devolvido" Then
'        If QFields(XData, "Devolvido") = "S" Then
'            subImprimeTexto "Devolvido", (IniCaixaX + 0.2), GuardaY, 8, "Times New Roman", Negrito, False
'        Else
'            subImprimeTexto "Não Devolvido", (IniCaixaX + 0.2), GuardaY, 8, "Times New Roman", Negrito, False
'        End If
'    Else
'        RecebeCampo = QFields(XData, xcampo)
'        subImprimeTexto funNulo(RecebeCampo), (IniCaixaX + 0.2), GuardaY, 8, "Times New Roman", Negrito, False
'    End If
    
End Sub

Sub subCabecalhoListagemRelatorio(Titulo As String)
    Printer.CurrentY = 1
'    If PLogo <> "" Then
'        Printer.PaintPicture LoadPicture(PLogo), 2, 1, 30, 5.5, , , 38.36, 7.28
'    End If
    subImprimeTexto PEmpresa, 6, Printer.CurrentY, 16, "Times New Roman", True, True
    'subImprimeTexto PEndereco, 6, Printer.CurrentY, 10, "Times New Roman", False, False
    subImprimeTexto Titulo, 6, Printer.CurrentY, 15, "Times New Roman", True, False
    Printer.CurrentY = 4
    AltCabecalho = 4
End Sub

Sub subCabecalho_Listagem(Titulo As String)
    Printer.CurrentY = 1.5
    Printer.Line (2.1, 1.5)-((Len(Titulo) / 2.5) + 0.1, 2.2), QBColor(0), BF
    Printer.FillStyle = 0
    Printer.DrawStyle = 5
    Printer.FillColor = RGB(240, 240, 240)
    Printer.Line (2, 1.4)-(Len(Titulo) / 2.5, 2.1), , B
    Printer.FillColor = RGB(250, 250, 250)
    Printer.DrawStyle = 0
    Printer.FillStyle = 1
    Printer.CurrentY = 1.5
    subImprimeTexto Titulo, 2.2, Printer.CurrentY, 13, "Times New Roman", True, False
    Printer.CurrentY = 3
    AltCabecalho = 3
End Sub

Sub subCabecalho_Listagem_Relatorio(Titulo As String)
    Printer.CurrentY = 1
'    If PLogo <> "" Then
'        Printer.PaintPicture LoadPicture(PLogo), 2, 1, 30, 5.5, , , 38.36, 7.28
'    End If
    subImprimeTexto PEmpresa, 6, Printer.CurrentY, 16, "Times New Roman", True, True
    subImprimeTexto PEndereco, 6, Printer.CurrentY, 10, "Times New Roman", False, False
    subImprimeTexto Titulo, 6, Printer.CurrentY, 15, "Times New Roman", True, False
    Printer.CurrentY = 4
    AltCabecalho = 4
End Sub

Sub subImprimeLinha(XInicio As Single, XFim As Single, Espessura As Double)
    Printer.DrawWidth = Espessura     'Definida em Pixels
    Printer.Line (XInicio, Printer.CurrentY)-(XFim, Printer.CurrentY), QBColor(0) ', BF
End Sub

Sub subImprimeTexto(cabecalho As String, ByVal x As Single, ByVal y As Single, Tamanho As Long, Font As String, Negrito As Boolean, Italico As Boolean)
    'Printer.ScaleMode = vbCentimeters
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Font.Name = Font
    Printer.Font.Size = Tamanho
    Printer.Font.Bold = Negrito
    Printer.Font.Italic = Italico
    'Printer.FontTransparent = False
    
    Printer.Print cabecalho
End Sub

Sub subImprimeTextoLaser(cabecalho As String, ByVal x As Single, ByVal y As Single, Tamanho As Long, Font As String, Negrito As Boolean, Italico As Boolean)
    'Printer.ScaleMode = vbCentimeters
    a.Left = x
    a.Left = y
    a.Font.Name = Font
    a.Font.Size = Tamanho
    a.Font.Bold = Negrito
    a.Font.Italic = Italico
    'Printer.FontTransparent = False
    
    'Print cabecalho
End Sub
Sub subCompletaCampo(RecebeCampo As String, XFim As Single, Negrito As Boolean)
    'Valor do cheque por extenso
      
'      GuardaY = Printer.CurrentY

      'Dim XFim As Single,
      Dim final As Integer, x As Integer
      Dim linha1 As String
      Dim linha2 As String
  
      'XFim = 17 - ResPadrao!banc_vl_colext 'em centimetros
      final = Len(RecebeCampo)
      If Printer.TextWidth(RecebeCampo) > XFim Then
         While Printer.TextWidth(RecebeCampo) > XFim
            linha2 = RecebeCampo
            RecebeCampo = Left(RecebeCampo, 50)
            linha2 = Right(linha2, final - XFim)
         Wend '
         'Do While Mid(RecebeCampo, final, 1) <> " " And (Printer.TextWidth(Left(RecebeCampo , (Len(RecebeCampo) - final))) > XFim)
         '   final = final - 1
         'Loop
         linha1 = Mid(RecebeCampo, 1, final)
         linha2 = Mid(RecebeCampo, final + 1, Len(RecebeCampo) - final)
         
      Else
         linha1 = RecebeCampo
         linha2 = ""
      End If
      linha1 = Trim(linha1)
      linha2 = Trim(linha2)
      If Printer.TextWidth(linha1) < XFim And linha2 = "" Then
         While Printer.TextWidth(linha1) < XFim
           linha1 = linha1 + "*"
         Wend
      End If
      If Printer.TextWidth(linha2) < (XFim + 2) Then
         While Printer.TextWidth(linha2) < (XFim + 2)
           linha2 = linha2 + "*"
         Wend
      End If
      subImprimeTexto linha1, Printer.CurrentX, Printer.CurrentY, 9, "arial", Negrito, False
      If linha1 <> linha2 Then
       subImprimeTexto linha2, Printer.CurrentX, Printer.CurrentY, 9, "arial", Negrito, False
      End If
End Sub
Function subImprimePromissoria(Cont As Long, XLS_Obs As String, XLI_Clausula As Integer, XLO_RSTITU As ADODB.Recordset, XLO_RSCONFIG As ADODB.Recordset, XLD_DATA As Date, XLT_FONT As String)
    Dim TamanhoPapel As Single
    Dim LargPapel As Single, AltPapel As Single
    Dim RecebeCampo As String
    Dim XLB_Negrito As Boolean
    Dim XLS_Campo As String
    Dim aux As Boolean
    Dim LargPapelClausula As Long, AltPapelClausula As Long
    
    
        
    'Cont - seta o numero do documento a ser impresso. recebe o acumulo da altura do documento

    aux = True
    TamanhoPapel = Printer.PaperSize
    LargPapel = 17.5 / 567 'Convertendo Twips para Cms
    AltPapel = 7.6 / 567 'Convertendo Twips para Cms
    While Not (XLO_RSCONFIG.EOF)
      If Not (XLO_RSCONFIG!form_nr_negrito = -1) Then
        XLB_Negrito = False
      Else
        XLB_Negrito = True
      End If
      Printer.Font.Bold = XLB_Negrito
      Printer.Font.Size = XLO_RSCONFIG!form_nr_tamfonte
      Printer.Font.Name = XLT_FONT
      XLS_Campo = XLO_RSCONFIG!form_tx_nomecampo

    
            
      Select Case (XLS_Campo)
      'campos que nao estao na consulta ( datas e valores por extenso )
          Case "ValorTitulo"
            If Not (XLO_RSCONFIG!form_tx_Descricao = "Valor do Título") Then
              RecebeCampo = funExtenso(XLO_RSTITU.Fields(XLS_Campo))
                If XLI_Clausula = 1 Then
                  Printer.CurrentY = Cont + XLO_RSCONFIG!form_nr_altura
                  Printer.CurrentX = XLO_RSCONFIG!form_nr_largura
                  RecebeCampo = RecebeCampo & " REAJUSTADOS CONFORME O PACTUADO EM CONTRATO"
                  RecebeCampo = RecebeCampo & " *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** "
                  subQuebraTexto RecebeCampo, XLO_RSCONFIG!form_nr_largura, 15, XLO_RSCONFIG!form_nr_tamfonte, XLT_FONT, XLO_RSCONFIG!form_nr_negrito, False
                Else
                  Printer.CurrentY = Cont + XLO_RSCONFIG!form_nr_altura
                  Printer.CurrentX = XLO_RSCONFIG!form_nr_largura
                  RecebeCampo = RecebeCampo & " *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** ***  *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** ***"
                  subQuebraTexto RecebeCampo, XLO_RSCONFIG!form_nr_largura, 15, XLO_RSCONFIG!form_nr_tamfonte, XLT_FONT, XLO_RSCONFIG!form_nr_negrito, False
                  LargPapelClausula = XLO_RSCONFIG!form_nr_largura
                  AltPapelClausula = Cont + XLO_RSCONFIG!form_nr_altura + 1.5
                End If
            Else
              RecebeCampo = Format$(XLO_RSTITU.Fields(XLS_Campo), "##,##0.00")
            End If
          Case "Observacao"
            If Not (XLS_Obs = "" Or IsNull(XLS_Obs)) Then
              RecebeCampo = XLS_Obs
            Else
              RecebeCampo = ""
            End If
            
          Case "DataVenc"
            Select Case (XLO_RSCONFIG!form_tx_Descricao)
              Case "Dia do Vencimento - 2 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " dd")
              Case "Mês do Vencimento - Extenso"
                RecebeCampo = UCase(Format(XLO_RSTITU.Fields(XLS_Campo), "mmmm "))
              Case "Ano do Vencimento - 4 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "yyyy ")
              Case "Data do Vencimento - Exenso"
                'RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd             mmmm             yyyy")
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd \dia\s \do \mê\s \de mmmm \do a\no \de yyyy")
              Case "Ano do Vencimento - 2 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "yy ")
              Case "Data do Vencimento - Exenso 2"
                RecebeCampo = UCase(Format(XLO_RSTITU.Fields(XLS_Campo), "dd \de mmmm \de yyyy"))
              Case "Dia do Vencimento - Extenso"
                RecebeCampo = funExtenso2(Format(XLO_RSTITU.Fields(XLS_Campo), "dd"))
              Case Else
                RecebeCampo = FunNulo(XLO_RSTITU.Fields(XLS_Campo))
            End Select
          
          Case "DataVenda"
            Select Case (XLO_RSCONFIG!form_tx_Descricao)
              Case "Dia da venda - 2 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " dd")
              Case "Mês da venda - Extenso"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " mmmm ")
              Case "Ano da venda - 4 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "yyyy ")
              Case "Data da venda - Exenso"
                'RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " dd             mmmm             yyyy")
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd \dia\s \do \mê\s \de mmmm \do a\no \de yyyy")
              Case "Ano da venda - 2 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "yy ")
            Case Else
              RecebeCampo = FunNulo(XLO_RSTITU.Fields(XLS_Campo))
            End Select
          
          Case "DataRegCont"
            Select Case (XLO_RSCONFIG!form_tx_Descricao)
              Case "Dia do Registro Cont. - 2 Dig"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd")
              Case "Mês do Registro Cont - Extenso"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " mmmm ")
              Case "Ano Registro Cont. - 4 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " yyyy ")
              Case "Data Registro Cont. - Extenso"
                'RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd         mmmm             yyyy")
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd \dia\s \do \mê\s \de mmmm \do a\no \de yyyy")
              Case "Ano Registro Cont. - 2 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " yy ")
            Case Else
              RecebeCampo = FunNulo(XLO_RSTITU.Fields(XLS_Campo))
            End Select
            
            Case "DataAtual"
            Select Case (XLO_RSCONFIG!form_tx_Descricao)
              Case "Dia Atual"
                RecebeCampo = Format(XLD_DATA, "dd")
              Case "Mês Atual - Extenso"
                RecebeCampo = Format(XLD_DATA, "mmmm")
              Case "Ano Atual - 4 Dígitos"
                RecebeCampo = Format(XLD_DATA, "yyyy")
              Case "Data Atual - Extenso"
                'RecebeCampo = Format(XLD_DATA, "dd             mmmm                     yyyy")
                RecebeCampo = Format(XLD_DATA, "dd \dia\s \do \mê\s \de mmmm \do a\no \de yyyy")
              Case "Ano Atual - 2 Dígitos"
                RecebeCampo = Format(XLD_DATA, "yy")
              Case Else
                RecebeCampo = FunNulo(Format(XLD_DATA, "dd/mm/yyyy"))
            End Select
            
          Case "JurosFin"
            RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "##0.0000")
          
          Case "Cotacao"
            RecebeCampo = Format(FunNuloVal(XLO_RSTITU.Fields(XLS_Campo)), "##0.0000")
            
          Case "ValorIndexado"
            RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "##0.0000")
            
        Case Else
          RecebeCampo = FunNulo(XLO_RSTITU.Fields(XLS_Campo))
      End Select
      'bufferiza o docuento a ser impresso
      If aux Then
        subImprimeTexto RecebeCampo, XLO_RSCONFIG!form_nr_largura, Cont + XLO_RSCONFIG!form_nr_altura, XLO_RSCONFIG!form_nr_tamfonte, XLT_FONT, XLO_RSCONFIG!form_nr_negrito, False
        aux = True
      End If
      XLO_RSCONFIG.MoveNext
        
    Wend
    
    If (Mid(PEmpresa, 1, 7) = "FRANISA" Or Mid(PEmpresa, 1, 7) = "PLENA") Then 'COlocar esta condição nos padrões da empresa
        RecebeCampo = " EM MOEDA CORRENTE NACIONAL, REAJUSTADOS CONFORME O PACTUADO EM CONTRATO,"
        RecebeCampo = RecebeCampo & " NA PRAÇA DE SALVADOR - BA *** *** *** *** *** *** *** *** *** *** "
        'Printer.CurrentY = AltPapelClausula
        Printer.CurrentY = 5.5
        Printer.CurrentX = LargPapelClausula
        subQuebraTexto RecebeCampo, CSng(LargPapelClausula), 15, 8, XLT_FONT, 1, False
        subImprimeTexto RecebeCampo, LargPapelClausula, 5.5, 8, XLT_FONT, 1, False
    End If
    
'If Not XLO_RsConfig.EOF Then
  XLO_RSCONFIG.MoveFirst
'End If

End Function
Function subImprimePromissoriaLaser(Cont As Long, XLS_Obs As String, XLI_Clausula As Integer, XLO_RSTITU As ADODB.Recordset, XLO_RSCONFIG As ADODB.Recordset, XLD_DATA As Date, XLT_FONT As String)
    Dim TamanhoPapel As Single
    Dim LargPapel As Single, AltPapel As Single
    Dim RecebeCampo As String
    Dim XLB_Negrito As Boolean
    Dim XLS_Campo As String
    Dim aux As Boolean
    Dim LargPapelClausula As Long, AltPapelClausula As Long
  
    
    'Cont - seta o numero do documento a ser impresso. recebe o acumulo da altura do documento

    aux = True
    TamanhoPapel = a.Width
    LargPapel = 17.5 / 567 'Convertendo Twips para Cms
    AltPapel = 7.6 / 567 'Convertendo Twips para Cms
    While Not (XLO_RSCONFIG.EOF)
      If Not (XLO_RSCONFIG!form_nr_negrito = -1) Then
        XLB_Negrito = False
      Else
        XLB_Negrito = True
      End If
      a.Font.Bold = XLB_Negrito
      a.Font.Size = XLO_RSCONFIG!form_nr_tamfonte
      a.Font.Name = XLT_FONT
      XLS_Campo = XLO_RSCONFIG!form_tx_nomecampo
      
      Select Case (XLS_Campo)
      'campos que nao estao na consulta ( datas e valores por extenso )
          Case "ValorTitulo"
            If Not (XLO_RSCONFIG!form_tx_Descricao = "Valor do Título") Then
              RecebeCampo = funExtenso(XLO_RSTITU.Fields(XLS_Campo))
                If XLI_Clausula = 1 Then
                  a.Left = Cont + XLO_RSCONFIG!form_nr_altura
                  a.Top = XLO_RSCONFIG!form_nr_largura
                  RecebeCampo = RecebeCampo & " REAJUSTADOS CONFORME O PACTUADO EM CONTRATO"
                  RecebeCampo = RecebeCampo & " *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** "
'                  subQuebraTextoLaser RecebeCampo, XLO_RSCONFIG!form_nr_largura, 15, XLO_RSCONFIG!form_nr_tamfonte, XLT_FONT, XLO_RSCONFIG!form_nr_negrito, False
                Else
                  a.Left = Cont + XLO_RSCONFIG!form_nr_altura
                  a.Top = XLO_RSCONFIG!form_nr_largura
                  RecebeCampo = RecebeCampo & " *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** ***  *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** ***"
                  'subQuebraTextoLaserLaser RecebeCampo, XLO_RSCONFIG!form_nr_largura, 15, XLO_RSCONFIG!form_nr_tamfonte, XLT_FONT, XLO_RSCONFIG!form_nr_negrito, False
                  LargPapelClausula = XLO_RSCONFIG!form_nr_largura
                  AltPapelClausula = Cont + XLO_RSCONFIG!form_nr_altura + 1.5
                End If
            Else
              RecebeCampo = Format$(XLO_RSTITU.Fields(XLS_Campo), "##,##0.00")
            End If
          Case "Observacao"
            If Not (XLS_Obs = "" Or IsNull(XLS_Obs)) Then
              RecebeCampo = XLS_Obs
            Else
              RecebeCampo = ""
            End If
            
          Case "DataVenc"
            Select Case (XLO_RSCONFIG!form_tx_Descricao)
              Case "Dia do Vencimento - 2 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " dd")
              Case "Mês do Vencimento - Extenso"
                RecebeCampo = UCase(Format(XLO_RSTITU.Fields(XLS_Campo), "mmmm "))
              Case "Ano do Vencimento - 4 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "yyyy ")
              Case "Data do Vencimento - Exenso"
                'RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd             mmmm             yyyy")
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd \dia\s \do \mê\s \de mmmm \do a\no \de yyyy")
              Case "Ano do Vencimento - 2 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "yy ")
              Case "Data do Vencimento - Exenso 2"
                RecebeCampo = UCase(Format(XLO_RSTITU.Fields(XLS_Campo), "dd \de mmmm \de yyyy"))
              Case "Dia do Vencimento - Extenso"
                RecebeCampo = funExtenso2(Format(XLO_RSTITU.Fields(XLS_Campo), "dd"))
              Case Else
                RecebeCampo = FunNulo(XLO_RSTITU.Fields(XLS_Campo))
            End Select
          
          Case "DataVenda"
            Select Case (XLO_RSCONFIG!form_tx_Descricao)
              Case "Dia da venda - 2 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " dd")
              Case "Mês da venda - Extenso"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " mmmm ")
              Case "Ano da venda - 4 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "yyyy ")
              Case "Data da venda - Exenso"
                'RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " dd             mmmm             yyyy")
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd \dia\s \do \mê\s \de mmmm \do a\no \de yyyy")
              Case "Ano da venda - 2 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "yy ")
            Case Else
              RecebeCampo = FunNulo(XLO_RSTITU.Fields(XLS_Campo))
            End Select
          
          Case "DataRegCont"
            Select Case (XLO_RSCONFIG!form_tx_Descricao)
              Case "Dia do Registro Cont. - 2 Dig"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd")
              Case "Mês do Registro Cont - Extenso"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " mmmm ")
              Case "Ano Registro Cont. - 4 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " yyyy ")
              Case "Data Registro Cont. - Extenso"
                'RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd         mmmm             yyyy")
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "dd \dia\s \do \mê\s \de mmmm \do a\no \de yyyy")
              Case "Ano Registro Cont. - 2 Dígitos"
                RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), " yy ")
            Case Else
              RecebeCampo = FunNulo(XLO_RSTITU.Fields(XLS_Campo))
            End Select
            
            Case "DataAtual"
            Select Case (XLO_RSCONFIG!form_tx_Descricao)
              Case "Dia Atual"
                RecebeCampo = Format(XLD_DATA, "dd")
              Case "Mês Atual - Extenso"
                RecebeCampo = Format(XLD_DATA, "mmmm")
              Case "Ano Atual - 4 Dígitos"
                RecebeCampo = Format(XLD_DATA, "yyyy")
              Case "Data Atual - Extenso"
                'RecebeCampo = Format(XLD_DATA, "dd             mmmm                     yyyy")
                RecebeCampo = Format(XLD_DATA, "dd \dia\s \do \mê\s \de mmmm \do a\no \de yyyy")
              Case "Ano Atual - 2 Dígitos"
                RecebeCampo = Format(XLD_DATA, "yy")
              Case Else
                RecebeCampo = FunNulo(Format(XLD_DATA, "dd/mm/yyyy"))
            End Select
            
          Case "JurosFin"
            RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "##0.0000")
          
          Case "Cotacao"
            RecebeCampo = Format(FunNuloVal(XLO_RSTITU.Fields(XLS_Campo)), "##0.0000")
            
          Case "ValorIndexado"
            RecebeCampo = Format(XLO_RSTITU.Fields(XLS_Campo), "##0.0000")
            
        Case Else
          RecebeCampo = FunNulo(XLO_RSTITU.Fields(XLS_Campo))
      End Select
       'bufferiza o docuento a ser impresso
        If aux Then
            subImprimeTextoLaser RecebeCampo, XLO_RSCONFIG!form_nr_largura, Cont + XLO_RSCONFIG!form_nr_altura, XLO_RSCONFIG!form_nr_tamfonte, XLT_FONT, XLO_RSCONFIG!form_nr_negrito, False
        
        aux = True
      End If
      XLO_RSCONFIG.MoveNext
        
    Wend
    
    If (Mid(PEmpresa, 1, 7) = "FRANISA" Or Mid(PEmpresa, 1, 7) = "PLENA") Then 'COlocar esta condição nos padrões da empresa
        RecebeCampo = " EM MOEDA CORRENTE NACIONAL, REAJUSTADOS CONFORME O PACTUADO EM CONTRATO,"
        RecebeCampo = RecebeCampo & " NA PRAÇA DE SALVADOR - BA *** *** *** *** *** *** *** *** *** *** "
        'Printer.CurrentY = AltPapelClausula
        a.Left = 5.5
        a.Top = LargPapelClausula
        'subQuebraTextoLaser RecebeCampo, CSng(LargPapelClausula), 15, 8, XLT_FONT, 1, False
        subImprimeTexto RecebeCampo, LargPapelClausula, 5.5, 8, XLT_FONT, 1, False
    End If
    
'If Not XLO_RsConfig.EOF Then
  XLO_RSCONFIG.MoveFirst
'End If

End Function


Function subImprimeNFS(XLO_RSNOTA As ADODB.Recordset, XLO_RSCONFIG As ADODB.Recordset, CAMINHODORELATORIO As String, XLT_CAPTION As String)
' Esta funcao recebe dois RecordSets como entrada o primeiro contem os dados a serem impressos
'o segundo a posicao. lembrar de tratar a excecao gerada pela falta de um desses componentes antes de chamar a funcao
    Dim TamanhoPapel As Single
    Dim LargPapel As Single, AltPapel As Single
    Dim RecebeCampo As String
    Dim XLB_Negrito As Boolean
    Dim XLS_Campo As String
  
    Dim XLO_CRSECTIONS As CRAXDRT.Sections
    Dim XLO_CRSECTION As CRAXDRT.Section
    Dim XLO_CRRELOBJ As CRAXDRT.ReportObjects
    Dim XLO_CRSUBOBJ As CRAXDRT.FieldObject
    Dim XLO_CRSUBREPORT As CRAXDRT.Report
    Dim x As Integer
    Dim y As Integer
    Dim XLB_ACHOU As Boolean

    'Carrega os relatórios
    FrmVisRelatorios.Caption = XLT_CAPTION
    
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.path + CAMINHODORELATORIO)
  
    'Passa o RecordSet com as Notas a serem Impressas para o relatório
    XGR_RELATORIO.Database.SetDataSource XLO_RSNOTA, 3, 1

    TamanhoPapel = Printer.PaperSize
    LargPapel = 17.5 / 567 'Convertendo Twips para Cms ( 1 cm = 567 twips )
    AltPapel = 7.6 / 567 'Convertendo Twips para Cms
   ' AltIni = XLO_RSCONFIG!form_nr_altura / 567

  'Set XGR_RELATORIO = ConexaoRelatorio.Execute("SELECT * from consfinnfservico")
  
  Set XLO_CRSECTIONS = XGR_RELATORIO.Sections
  For x = 1 To XLO_CRSECTIONS.Count
    Set XLO_CRSECTION = XLO_CRSECTIONS.Item(x)
    Set XLO_CRRELOBJ = XLO_CRSECTION.ReportObjects
'    Laco dos objetos de cada seção
    y = 1
    'Percorre o recordSer de objetos do cristal
    While (y <= XLO_CRRELOBJ.Count)
      XLB_ACHOU = False
      XLO_RSCONFIG.MoveFirst
      'Percorre o Recordset com as confiurações de posição dos campos
      While (Not XLO_RSCONFIG.EOF And Not XLB_ACHOU)
        'Testa se eh uma campo do relatorio a ser impresso
        If XLO_CRRELOBJ.Item(y).Kind = crFieldObject Then
          'testa se o campo vai imprimir ou nao
          If XLO_RSCONFIG!form_nr_imprime = -1 Then
            'seta o objeto que esta esta sendo percorrido para um SubObj
            Set XLO_CRSUBOBJ = XLO_CRRELOBJ.Item(y)
            'Compara se o nome do Ojeto = nome do campo do relatorio
            XLS_Campo = funSubstituiCaracter(XLO_RSCONFIG!form_tx_nomecampo, "_", "")
            If UCase(XLO_CRSUBOBJ.Name) = UCase(XLS_Campo) Then
              XLB_ACHOU = True 'forca sair da iteracao internar
              XLO_CRSUBOBJ.Top = (XLO_RSCONFIG!form_nr_altura * 567)
              XLO_CRSUBOBJ.Left = (XLO_RSCONFIG!form_nr_largura * 567)
              XLO_CRSUBOBJ.Font.Size = XLO_RSCONFIG!form_nr_tamfonte
              XLO_CRSUBOBJ.Suppress = False
              If XLO_RSCONFIG!form_nr_negrito = -1 Then
                XLO_CRSUBOBJ.Font.Bold = True
              End If
            Else
              XLO_RSCONFIG.MoveNext
            End If
          Else
            XLO_RSCONFIG.MoveNext
          End If
        Else
          'y = y + 1
          XLB_ACHOU = True
        End If
      Wend
      y = y + 1 'so troca o obj depois de percorrer todo o RS de configuracoes
    Wend
  Next

  'Ativa os Relatórios
  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
End Function
Function funExtenso(XNum As Double) As String

'Esta função retorna um valor por extenso

Dim T(4) As String
Dim V(4) As String
Dim p(4) As String
Dim S(4) As String
Dim N(900) As String
Dim Texto As String
Dim xNumero, num1, num2, num3, num4 As String
Dim x As Integer
Dim Status, Primeiro As Boolean
Texto = ""
N(1) = "UM "
N(2) = "DOIS "
N(3) = "TRES "
N(4) = "QUATRO "
N(5) = "CINCO "
N(6) = "SEIS "
N(7) = "SETE "
N(8) = "OITO "
N(9) = "NOVE "
N(10) = "DEZ "
N(11) = "ONZE "
N(12) = "DOZE "
N(13) = "TREZE "
N(14) = "QUATORZE "
N(15) = "QUINZE "
N(16) = "DEZESSEIS "
N(17) = "DEZESSETE "
N(18) = "DEZOITO "
N(19) = "DEZENOVE "
N(20) = "VINTE "
N(30) = "TRINTA "
N(40) = "QUARENTA "
N(50) = "CINQUENTA "
N(60) = "SESSENTA "
N(70) = "SETENTA "
N(80) = "OITENTA "
N(90) = "NOVENTA "
N(0) = "CEM "
N(100) = "CENTO "
N(200) = "DUZENTOS "
N(300) = "TREZENTOS "
N(400) = "QUATROCENTOS "
N(500) = "QUINHENTOS "
N(600) = "SEISCENTOS "
N(700) = "SETECENTOS "
N(800) = "OITOCENTOS "
N(900) = "NOVECENTOS "
T(4) = "BILHAO "
T(3) = "MILHAO "
T(2) = "MIL "
T(1) = ""
V(4) = "BILHOES "
V(3) = "MILHOES "
V(2) = "MIL "
V(1) = ""
xNumero = Right((Space(12) + Trim(Str(Int(XNum)))), 12)
p(4) = Mid(xNumero, 1, 3)
p(3) = Mid(xNumero, 4, 3)
p(2) = Mid(xNumero, 7, 3)
p(1) = Mid(xNumero, 10, 3)
'p(0) = Trim(Str(Round(XNum, 2) * 100 - Round(XNum) * 100))
p(0) = Right(Trim(Str((Round(XNum, 2) * 100))), 2)
S(4) = Val(p(4) + "000000000.00")
S(3) = Val(p(4) + p(3) + "000000.00")
S(2) = Val(p(4) + p(3) + p(2) + "000.00")
S(1) = Val(p(4) + p(3) + p(2) + p(1) + ".00")
For x = 4 To 1 Step -1
    If Val(p(x)) <> 0 Then
       num1 = Trim(Mid(p(x), 1, 1))
       num2 = Trim(Mid(p(x), 2, 1))
       num3 = Trim(Mid(p(x), 3, 1))
       Status = IIf(num2 <> "1", True, False)
       Primeiro = True
       If num1 <> "" And num1 <> "0" Then
          If num1 = "1" And num2 + num3 = "00" Then
             num1 = "000"
          Else
             num1 = num1 + "00"
          End If
          If XNum > S(x) Then
             Texto = IIf(Texto <> "", Trim(Texto) + ", " + N(Val(num1)), Texto + N(Val(num1)))
          Else
             Texto = IIf(Texto <> "", Texto + "E " + N(Val(num1)), Texto + N(Val(num1)))
          End If
          Primeiro = False
       End If
       If num2 <> "" And num2 <> "0" Then
          num2 = IIf(Status, num2 + "0", num2 + num3)
          If Primeiro And XNum > S(x) Then
             Texto = IIf(Texto <> "", Trim(Texto) + ", " + N(Val(num2)), Texto + N(Val(num2)))
          Else
             Texto = IIf(Texto <> "", Texto + "E " + N(Val(num2)), Texto + N(Val(num2)))
          End If
          Primeiro = False
       End If
       If num3 <> "" And num3 <> "0" And Status Then
          If Primeiro And XNum > S(x) Then
             Texto = IIf(Texto <> "", Trim(Texto) + ", " + N(Val(num3)), Texto + N(Val(num3)))
          Else
             Texto = IIf(Texto <> "", Texto + "E " + N(Val(num3)), Texto + N(Val(num3)))
          End If
       End If
       Texto = IIf(Val(p(x)) = 1, Texto + T(x), Texto + V(x))
    End If
Next
If Val(xNumero) <> 0 Then
   If Val(xNumero) = 1 Then
      Texto = Texto + "REAL "
   ElseIf Val(p(2)) = 0 And Val(p(1)) = 0 Then
      Texto = Texto + "DE REAIS "
   Else
      Texto = Texto + "REAIS "
   End If
End If
If Val(p(0)) <> 0 Then
   num1 = Mid(p(0), 1, 1)
   num2 = Mid(p(0), 2, 1)
   If num1 <> "" And num1 <> "0" Then
      num4 = IIf(num1 <> "1", num1 + "0", num1 + num2)
      Texto = IIf(Texto <> "", Texto + "E " + N(Val(num4)), Texto + N(Val(num4)))
   End If
   If num2 <> "" And num2 <> "0" And num1 <> "1" Then
      Texto = IIf(Texto <> "", Texto + "E " + N(Val(num2)), Texto + N(Val(num2)))
   End If
   If Val(p(0)) = 1 Then
      Texto = Texto + "CENTAVO"
   Else
      Texto = Texto + "CENTAVOS"
   End If
End If
funExtenso = Texto

End Function


Function funExtenso2(XNum As Double) As String

'Esta função retorna um valor por extenso

Dim T(4) As String
Dim V(4) As String
Dim p(4) As String
Dim S(4) As String
Dim N(900) As String
Dim Texto As String
Dim xNumero, num1, num2, num3, num4 As String
Dim x As Integer
Dim Status, Primeiro As Boolean
Texto = ""
N(1) = "UM "
N(2) = "DOIS "
N(3) = "TRES "
N(4) = "QUATRO "
N(5) = "CINCO "
N(6) = "SEIS "
N(7) = "SETE "
N(8) = "OITO "
N(9) = "NOVE "
N(10) = "DEZ "
N(11) = "ONZE "
N(12) = "DOZE "
N(13) = "TREZE "
N(14) = "QUATORZE "
N(15) = "QUINZE "
N(16) = "DEZESSEIS "
N(17) = "DEZESSETE "
N(18) = "DEZOITO "
N(19) = "DEZENOVE "
N(20) = "VINTE "
N(30) = "TRINTA "
V(1) = ""
xNumero = Right((Space(12) + Trim(Str(Int(XNum)))), 12)
p(4) = Mid(xNumero, 1, 3)
p(3) = Mid(xNumero, 4, 3)
p(2) = Mid(xNumero, 7, 3)
p(1) = Mid(xNumero, 10, 3)
'p(0) = Trim(Str(Round(XNum, 2) * 100 - Round(XNum) * 100))
p(0) = Right(Trim(Str((Round(XNum, 2) * 100))), 2)
S(4) = Val(p(4) + "000000000.00")
S(3) = Val(p(4) + p(3) + "000000.00")
S(2) = Val(p(4) + p(3) + p(2) + "000.00")
S(1) = Val(p(4) + p(3) + p(2) + p(1) + ".00")

For x = 4 To 1 Step -1
    If Val(p(x)) <> 0 Then
       num1 = Trim(Mid(p(x), 1, 1))
       num2 = Trim(Mid(p(x), 2, 1))
       num3 = Trim(Mid(p(x), 3, 1))
       Status = IIf(num2 <> "1", True, False)
       Primeiro = True
       If num1 <> "" And num1 <> "0" Then
          If num1 = "1" And num2 + num3 = "00" Then
             num1 = "000"
          Else
             num1 = num1 + "00"
          End If
          If XNum > S(x) Then
             Texto = IIf(Texto <> "", Trim(Texto) + ", " + N(Val(num1)), Texto + N(Val(num1)))
          Else
             Texto = IIf(Texto <> "", Texto + "E " + N(Val(num1)), Texto + N(Val(num1)))
          End If
          Primeiro = False
       End If
       If num2 <> "" And num2 <> "0" Then
          num2 = IIf(Status, num2 + "0", num2 + num3)
          If Primeiro And XNum > S(x) Then
             Texto = IIf(Texto <> "", Trim(Texto) + ", " + N(Val(num2)), Texto + N(Val(num2)))
          Else
             Texto = IIf(Texto <> "", Texto + "E " + N(Val(num2)), Texto + N(Val(num2)))
          End If
          Primeiro = False
       End If
       If num3 <> "" And num3 <> "0" And Status Then
          If Primeiro And XNum > S(x) Then
             Texto = IIf(Texto <> "", Trim(Texto) + ", " + N(Val(num3)), Texto + N(Val(num3)))
          Else
             Texto = IIf(Texto <> "", Texto + "E " + N(Val(num3)), Texto + N(Val(num3)))
          End If
       End If
       Texto = IIf(Val(p(x)) = 1, Texto + T(x), Texto + V(x))
    End If
Next

funExtenso2 = Texto

End Function


Sub subQuebraLinha(Texto As String, XInicio As Single, TamCampo As Single, TamLetra As Long, Fonte As String, Negrito As Boolean, Italico As Boolean)
    'feita por Daniel
    Dim Resto As String
    Dim ProcuraEspaco As Integer
        
    'Para Garantir Que Vai Verificar o Tamanho Com a Formatação Correta
    Printer.Font.Name = Fonte
    Printer.Font.Size = TamLetra
    Printer.Font.Bold = Negrito
    Printer.Font.Italic = Italico
    
    Resto = ""
    Texto = Trim$(Texto)
    'Texto é uma variável que indica a string a ser impressa, após a impressão o seu valor é
    'substituído pelo valor de Resto, se não existir Resto significa que já foi impresso toda a
    'string, consequentemente o valor de texto também será vazio.
    
    'Se Texto for vazio imprime vazio e sai
    If Texto = "" Then
        subImprimeTexto Texto, XInicio, Printer.CurrentY, TamLetra, Fonte, Negrito, Italico
        Exit Sub
    End If
    
    While Texto <> ""
        'Enquanto o tamanho do Texto for maior que o tamanho do espaço reservado para ele
        'será retirado o último caractere do Texto e este será acrescido a variável Resto
        While Printer.TextWidth(Texto) > TamCampo
            Resto = Mid$(Texto, Len(Texto), 1) + Resto
            Texto = Mid$(Texto, 1, (Len(Texto) - 1))
        Wend
        
'*********************************************************************************************************
'Esta parte da sub é para impedir que se quebre uma palavra no meio
'*********************************************************************************************************
        'Se Resto for igual a vazio significa que não existirá uma próxima linha
        If Resto <> "" Then
            If Mid$(Resto, 1, 1) = " " Then 'Resto possui espaço no início quebrou uma palavra completa
                Resto = LTrim$(Resto)
            ElseIf Mid$(Texto, Len(Texto), 1) = " " Then 'Texto possui espaço no fim
                Texto = RTrim$(Texto)                           'quebrou uma palavra completa
            Else 'Quebrou metade de uma palavra
                ProcuraEspaco = Len(Texto)
                'Verifica se Texto possui algum espaço para quebrar a palavra completa
                While Mid$(Texto, ProcuraEspaco, 1) <> " " And ProcuraEspaco <> 1
                    ProcuraEspaco = ProcuraEspaco - 1
                Wend
                If ProcuraEspaco <> 1 Then
                    Resto = Right$(Texto, Len(Texto) - ProcuraEspaco) & Resto
                    Texto = RTrim$(Left(Texto, ProcuraEspaco))
                End If
            End If
        End If
'*********************************************************************************************************
        subImprimeTexto Texto, XInicio, Printer.CurrentY, TamLetra, Fonte, Negrito, Italico
        Texto = Resto
        Resto = ""
    Wend
End Sub




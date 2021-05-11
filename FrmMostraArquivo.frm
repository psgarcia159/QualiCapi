VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMostraArquivo 
   Caption         =   "Visualização do arquivo a ser importado"
   ClientHeight    =   6660
   ClientLeft      =   3720
   ClientTop       =   1365
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleMode       =   0  'User
   ScaleWidth      =   12420
   StartUpPosition =   1  'CenterOwner
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   5655
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9975
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Linha"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Centro de Custo"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Cod.Desc/Prov"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Conta Contábil"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Data Pagamento"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Histórico"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Valor Líquido"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Tipo"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Status"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=794"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2143"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2064"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1931"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1852"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1958"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1879"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2355"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2275"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=5159"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=5080"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2037"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1958"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=1005"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=926"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(72)  =   "Named:id=33:Normal"
      _StyleDefs(73)  =   ":id=33,.parent=0"
      _StyleDefs(74)  =   "Named:id=34:Heading"
      _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(76)  =   ":id=34,.wraptext=-1"
      _StyleDefs(77)  =   "Named:id=35:Footing"
      _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(79)  =   "Named:id=36:Selected"
      _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=37:Caption"
      _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(83)  =   "Named:id=38:HighlightRow"
      _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=39:EvenRow"
      _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(87)  =   "Named:id=40:OddRow"
      _StyleDefs(88)  =   ":id=40,.parent=33"
      _StyleDefs(89)  =   "Named:id=41:RecordSelector"
      _StyleDefs(90)  =   ":id=41,.parent=34"
      _StyleDefs(91)  =   "Named:id=42:FilterBar"
      _StyleDefs(92)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   10800
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   375
      Left            =   9360
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog DlgAbrirArquivo 
      Left            =   8640
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmMostraArquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const COLUNASGRID_LINHA As Integer = 0
Const COLUNASGRID_CENTROCUSTO As Integer = 1
Const COLUNASGRID_CODDESC_PROV As Integer = 2
Const COLUNASGRID_CCONTABIL As Integer = 3
Const COLUNASGRID_DATAPGTO As Integer = 4
Const COLUNASGRID_HISTORICO As Integer = 5
Const COLUNASGRID_VALLIQUIDO As Integer = 6
Const COLUNASGRID_TIPO As Integer = 7
Const COLUNASGRID_STATUS As Integer = 8
Private XLB_CONFIRMA As Boolean
Dim marca As BookmarkEnum

Public Property Get ConfirmaImportacao() As Boolean
    ConfirmaImportacao = XLB_CONFIRMA
End Property

Function VerificaString(texto As String) As Boolean
    Dim Cont As Integer
      
    For Cont = 1 To Len(texto)
        
        If (((Asc(Mid(texto, Cont, 1)) >= 32) And (Asc(Mid(texto, Cont, 1)) <= 126))) Or (((Asc(Mid(texto, Cont, 1)) >= 128) And (Asc(Mid(texto, Cont, 1)) <= 255))) Then
            VerificaString = True
            
        Else
            
            VerificaString = False ' se achou o primeiro caracter estranho a função já retorna falso
            GoTo Fim
            
        End If
    Next Cont
Fim:
End Function

Function AbreArquivo() As Boolean
    Dim XGT_MSKCODREDUZIDO              As String * 254
    Dim XLT_NOMEARQUIVO                 As String
    Dim XLT_TEXTOLINHA                  As String
    Dim XLI_ARQ                         As Integer
    Dim XLI_NUMLINHA                    As Integer
    Dim XLT_PCR                         As String
    Dim XLT_CCONTABIL                   As String
    Dim XLT_CODRM                       As String
    Dim XLD_DATA                        As String
    Dim XLT_HISTORICO                   As String
    Dim XLT_STATUS                      As String
    Dim XLF_VALORLIQUIDO                As String
    Dim XLT_DEBCRED                     As String
    Dim VLT_DATASOURCE                  As New XArray
    Dim Cont                            As Integer
    Dim linha_col                       As String
    Dim linha                           As String
    Dim coluna                          As String
      
    XLT_NOMEARQUIVO = XGT_LOCALARQ
    XLI_NUMLINHA = 0
        
    'Verificando se o arquivo existe:
    If Dir$(XLT_NOMEARQUIVO) = "" Then
        'O arquivo NÃO existe.
        MsgBox "O arquivo não foi informado!", vbExclamation, "ATENÇÃO"
        Exit Function
    End If
    XLI_ARQ = FreeFile
    
    Open XLT_NOMEARQUIVO For Input As XLI_ARQ
    
    'Verificando se o arquivo está em branco:
    If LOF(XLI_ARQ) = 0 Then
        'O arquivo está em branco.
        MsgBox "O arquivo está em branco!", vbExclamation, "ATENÇÃO"
        'Fecha o arquivo
        Close XLI_ARQ
        Exit Function
    End If


    'Vetor - variavel global do tipo vetor que será preenchido para popular o grid
    VLT_DATASOURCE.ReDim 0, 0, 0, TDBGrid1.Columns.Count - 1
    VLT_DATASOURCE.Clear
    TDBGrid1.Array = VLT_DATASOURCE
    
    
    Do While Not EOF(XLI_ARQ) 'Enquanto ñ chegar ao fim do arq...
        Line Input #XLI_ARQ, XLT_TEXTOLINHA
        
         If VerificaString(XLT_TEXTOLINHA) = False Then
                XLT_STATUS = "Linha com erro"
                
         Else
                XLT_STATUS = "OK"
         End If
            
        'REDIMENSIONA O VETOR NOVAMENTE PARA INSERIR MAIS UMA LINHA
        If VLT_DATASOURCE.UpperBound(1) <> 0 Or VLT_DATASOURCE(0, 0) <> Empty Then
            VLT_DATASOURCE.Insert 1, VLT_DATASOURCE.UpperBound(1) + 1
        End If
        'XLI_TAMANHOVETOR = Vetor.UpperBound(1)
               
        'CENTRO DE RESULTADO
        XLT_PCR = funFormataCCusto(Trim(Mid(XLT_TEXTOLINHA, 1, 10)))
        If Val(XLT_PCR) = True Then
            XLT_STATUS = "ERRO"
        End If
        
        
            
        If FrmImportaFolha.ChkMascara.Value = 1 Then
            'CODIGO DE DESCONTO RM
            XLT_CODRM = Trim(Mid(XLT_TEXTOLINHA, 11, 7))
            'CONTA CONTABIL
            XLT_CCONTABIL = FunTrataZero(Trim(Mid(XLT_TEXTOLINHA, 18, 10)))
        
        Else
            'CODIGO DE DESCONTO RM
            XLT_CODRM = Trim(Mid(XLT_TEXTOLINHA, 11, 7))
            'CONTA CONTABIL
            XLT_CCONTABIL = Format(Trim(Mid(XLT_TEXTOLINHA, 18, 10)), XGT_MSKCODREDUZIDO)
        End If
        
        'DATA
        XLD_DATA = (Mid(XLT_TEXTOLINHA, 28, 2) + "/" + Mid(XLT_TEXTOLINHA, 30, 2) + "/" + Mid(XLT_TEXTOLINHA, 32, 4))
        'If (CDate(XLD_DATA)) = False Then
           ' XLT_STATUS = "ERRO"


        'End If
        'HISTORICO
        XLT_HISTORICO = Trim(Mid(XLT_TEXTOLINHA, 36, 44))
        'VALOR LIQUIDO
        XLF_VALORLIQUIDO = Format((Mid(XLT_TEXTOLINHA, 81, 13) + "," + Mid(XLT_TEXTOLINHA, 94, 2)), "standard")
        'SE É UM DEBITO OU CREDITO
        XLT_DEBCRED = Mid(XLT_TEXTOLINHA, 96, 1)
        
        'ATRIBUE AO VETOR NA LINHA CORRESPONDENTE OS VALORES LIDOS
        VLT_DATASOURCE(XLI_NUMLINHA, COLUNASGRID_LINHA) = XLI_NUMLINHA + 1
        VLT_DATASOURCE(XLI_NUMLINHA, COLUNASGRID_CENTROCUSTO) = XLT_PCR
        VLT_DATASOURCE(XLI_NUMLINHA, COLUNASGRID_CODDESC_PROV) = XLT_CODRM
        VLT_DATASOURCE(XLI_NUMLINHA, COLUNASGRID_CCONTABIL) = XLT_CCONTABIL
        VLT_DATASOURCE(XLI_NUMLINHA, COLUNASGRID_DATAPGTO) = XLD_DATA
        VLT_DATASOURCE(XLI_NUMLINHA, COLUNASGRID_HISTORICO) = XLT_HISTORICO
        VLT_DATASOURCE(XLI_NUMLINHA, COLUNASGRID_VALLIQUIDO) = XLF_VALORLIQUIDO
        VLT_DATASOURCE(XLI_NUMLINHA, COLUNASGRID_TIPO) = XLT_DEBCRED
        VLT_DATASOURCE(XLI_NUMLINHA, COLUNASGRID_STATUS) = XLT_STATUS
        
        'INCREMENTA UMA LINHA
        XLI_NUMLINHA = XLI_NUMLINHA + 1
    Loop
    
    TDBGrid1.ReBind
    
    

        
    'Fecha o arquivo
    Close XLI_ARQ
    
    AbreArquivo = True
    
End Function

Private Sub cmdCancela_Click()
    XLB_CONFIRMA = False
    Unload Me
End Sub

Private Sub CmdConfirmar_Click()
    XLB_CONFIRMA = True
    Unload Me
End Sub
Private Sub Form_Load()
    
    AbreArquivo
    
End Sub

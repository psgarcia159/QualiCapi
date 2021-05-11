VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form TelaCartas 
   Caption         =   "Tabela de Observações"
   ClientHeight    =   7860
   ClientLeft      =   1830
   ClientTop       =   2220
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   8910
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanObservacao 
      Height          =   5475
      Left            =   1200
      TabIndex        =   12
      Top             =   1020
      Visible         =   0   'False
      Width           =   6780
      _Version        =   65536
      _ExtentX        =   11959
      _ExtentY        =   9657
      _StockProps     =   15
      Caption         =   " Cadastro de Cartas"
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   0
      Begin RichTextLib.RichTextBox RTBox 
         Height          =   3555
         Left            =   240
         TabIndex        =   18
         Top             =   1380
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6271
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"TelaCartas.frx":0000
      End
      Begin VB.CommandButton cmdAbrirCarta 
         Caption         =   "A&brir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7020
         TabIndex        =   17
         Top             =   1740
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5280
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdAbrirEditor 
         Caption         =   "Abrir &Editor"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   1
         Top             =   4980
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox TxtDescricao 
         Height          =   315
         Left            =   240
         MaxLength       =   30
         TabIndex        =   0
         Top             =   840
         Width           =   6300
      End
      Begin VB.CommandButton CmdConfirmar 
         Caption         =   "&Cadastrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4380
         TabIndex        =   2
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton CmdRetornar 
         Caption         =   "C&ancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5460
         TabIndex        =   3
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label LblDescricao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   915
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4320
      Top             =   6780
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaCartas.frx":00AE
      Height          =   6690
      Left            =   60
      TabIndex        =   14
      Top             =   15
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   11800
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Descrição"
      Columns(0).DataField=   "cart_tx_descricao"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=14684"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=14605"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   2
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(21)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(22)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7,.wraptext=-1"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.locked=0"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15,.alignment=3"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17,.alignment=3"
      _StyleDefs(36)  =   "Named:id=33:Normal"
      _StyleDefs(37)  =   ":id=33,.parent=0"
      _StyleDefs(38)  =   "Named:id=34:Heading"
      _StyleDefs(39)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(40)  =   ":id=34,.wraptext=-1"
      _StyleDefs(41)  =   "Named:id=35:Footing"
      _StyleDefs(42)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(43)  =   "Named:id=36:Selected"
      _StyleDefs(44)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(45)  =   "Named:id=37:Caption"
      _StyleDefs(46)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(47)  =   "Named:id=38:HighlightRow"
      _StyleDefs(48)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=39:EvenRow"
      _StyleDefs(50)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(51)  =   "Named:id=40:OddRow"
      _StyleDefs(52)  =   ":id=40,.parent=33"
      _StyleDefs(53)  =   "Named:id=41:RecordSelector"
      _StyleDefs(54)  =   ":id=41,.parent=34"
      _StyleDefs(55)  =   "Named:id=42:FilterBar"
      _StyleDefs(56)  =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   360
      Left            =   -60
      TabIndex        =   11
      Top             =   7260
      Width           =   8760
      _Version        =   65536
      _ExtentX        =   15452
      _ExtentY        =   635
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin VB.CommandButton CmdInserir 
         BackColor       =   &H00000000&
         Caption         =   "&Inserir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7560
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Im&primir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6360
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton CmdOrdem 
         Caption         =   "&Ordem"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5160
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton CmdRemoverFiltro 
         Caption         =   "Se&leções"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3960
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton CmdExcluir 
         Caption         =   "&Excluir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2760
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton CmdAlterar 
         BackColor       =   &H00000000&
         Caption         =   "&Alterar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaCartas.frx":00C3
      Height          =   315
      Left            =   2640
      TabIndex        =   15
      Top             =   6780
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "obse_tx_observacao"
      BoundColumn     =   "obse_cd_observacao"
      Text            =   ""
   End
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Descrição:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   6780
      Width           =   2670
   End
End
Attribute VB_Name = "TelaCartas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************************
'Variaveis que guardarão valores específicos de cada formulário
Dim PrimeiraVez As Boolean

Dim FRM_ContFil As Integer ' Contador que armazena o número de filtros ativos
Dim FRM_Ordem As String ' Nome do Campo que tem a chave da Tabela ativa
Dim FRM_Campos_Obrig As String ' Lista dos campos que são obrigatórios na tabela
Dim FRM_TipoVar As String ' Tipo do campo a ser filtrado
Dim FRM_FiltroOrig As String ' Indica os filtros determinados pelo próprio sistema
Dim FRM_FiltroAtual As String ' Indica o filtro atual da tabela corrente
Dim FRM_filtra_especial As Integer ' Indica se a tabela utilizada possui filtros pré-definidos pelo sistema em FiltroOrig
Dim FRM_Tem_Filtro As Boolean ' Informa se já foi relizado um filtro na tabela pelo usuário
Dim FRM_ConTelaOrdem As Integer ' Indica o número de chaves utilizados para ordenação
Dim FRM_Quantid As Long  ' Quantidade de registros da tabela
Dim FRM_Filtrou As Boolean ' Informa se Algum campo foi filtrado
Dim FRM_Ordens() As Integer    ' Vetor que mantem a ordem dos campos
Dim FRM_Filtros() As String ' Array para armazenar os Filtros da Tabela

Dim FRM_BarraFerramentasPressionada As Boolean
Dim FRM_VariasVezes As Boolean    'Indica se vai executar a insercao na TelaPadrao varias ou uma vez
Dim FRM_XInserir As String        'Flag para identificar se entrou na rotina para inserir ou alterar ('I'Inserir, 'A'Alterar, 'C'Consultar)
Dim FRM_XOrigem As String         'Identifica a origem (que formulario) chamou o formulário ativo
Dim FRM_SQLAtivo As String        'Atualizar o SQL Ativo após modificações de filtro e ordem
Dim XFS_PathCarta As String       'indica o caminho da carta que foi carregada

Sub subTelaValoresGlobais(GuardaOUPegaValores As String)
    Dim Count As Integer, Count2 As Integer
    If GuardaOUPegaValores = "G" Then 'Guardando variáveis de definição da tela
        FRM_ContFil = ContFil
        FRM_Ordem = Ordem
        FRM_Campos_Obrig = Campos_Obrig
        FRM_TipoVar = TipoVar
        FRM_FiltroOrig = FiltroOrig
        FRM_FiltroAtual = FiltroAtual
        FRM_filtra_especial = filtra_especial
        FRM_Tem_Filtro = Tem_Filtro
        FRM_ConTelaOrdem = ConTelaOrdem
        FRM_Quantid = Quantid
        FRM_Filtrou = Filtrou
        FRM_BarraFerramentasPressionada = BarraFerramentasPressionada
        FRM_VariasVezes = VariasVezes
        FRM_XInserir = XInserir
        FRM_XOrigem = XOrigem
        FRM_SQLAtivo = SQLAtivo
        
        ReDim FRM_Filtros(UBound(Filtros, 1), UBound(Filtros, 2))
        ReDim FRM_Ordens(UBound(Ordens, 1)) As Integer
        
        For Count = LBound(Filtros, 1) To UBound(Filtros, 1)
            For Count2 = LBound(Filtros, 2) To UBound(Filtros, 2)
                FRM_Filtros(Count, Count2) = Filtros(Count, Count2)
            Next Count2
        Next Count
        For Count = LBound(Ordens) To UBound(Ordens)
            FRM_Ordens(Count) = Ordens(Count)
        Next Count
    ElseIf GuardaOUPegaValores = "P" Then  'Regravando os valores relativos a tela atual variáveis de definição da tela
        ContFil = FRM_ContFil
        Ordem = FRM_Ordem
        Campos_Obrig = FRM_Campos_Obrig
        TipoVar = FRM_TipoVar
        FiltroOrig = FRM_FiltroOrig
        FiltroAtual = FRM_FiltroAtual
        filtra_especial = FRM_filtra_especial
        Tem_Filtro = FRM_Tem_Filtro
        ConTelaOrdem = FRM_ConTelaOrdem
        Quantid = FRM_Quantid
        Filtrou = FRM_Filtrou
        BarraFerramentasPressionada = FRM_BarraFerramentasPressionada
        VariasVezes = FRM_VariasVezes
        XInserir = FRM_XInserir
        XOrigem = FRM_XOrigem
        SQLAtivo = FRM_SQLAtivo

        ReDim Filtros(UBound(FRM_Filtros, 1), UBound(FRM_Filtros, 2))
        ReDim Ordens(UBound(FRM_Ordens, 1))
        
        For Count = LBound(FRM_Filtros, 1) To UBound(FRM_Filtros, 1)
            For Count2 = LBound(FRM_Filtros, 2) To UBound(FRM_Filtros, 2)
                Filtros(Count, Count2) = FRM_Filtros(Count, Count2)
            Next Count2
        Next Count
        For Count = LBound(FRM_Ordens) To UBound(FRM_Ordens)
            Ordens(Count) = FRM_Ordens(Count)
        Next Count
    End If
End Sub

Private Sub SubHabilitaBotoes()

    'Objetivo: Habilita os Botões da Tela
    CmdInserir.Enabled = True
    CmdAlterar.Enabled = True
    CmdExcluir.Enabled = True
    CmdSair.Enabled = True
    CmdImprimir.Enabled = True
    CmdOrdem.Enabled = True

    'Verifica se tem Filtro Ativo
    If ContFil = 0 Then
       CmdRemoverFiltro.Enabled = False
    Else
       CmdRemoverFiltro.Enabled = True
    End If
End Sub
Private Sub SubDesabilitaBotoes()
    'Objetivo: Desabilita os Botões da Tela
    CmdInserir.Enabled = False
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    CmdRemoverFiltro.Enabled = False
    CmdSair.Enabled = False
    CmdImprimir.Enabled = False
    CmdOrdem.Enabled = False
End Sub

Private Sub cboLocalizar_Click(Area As Integer)
    If Area = 2 Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
End Sub

Private Sub cboLocalizar_KeyUp(KeyCode As Integer, Shift As Integer)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
End Sub

Private Sub cboLocalizar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
End Sub


Private Sub chkItalico_Click()
  Rich.SelItalic = chkItalico.Value
End Sub

Private Sub chkNegrito_Click()
  Rich.SelBold = chkNegrito.Value
End Sub

Private Sub chkSubilinhado_Click()
  Rich.SelUnderline = chksublinhado.Value
End Sub

Private Sub cmdAbrirWP_Click()
    
End Sub

Private Sub cmdAbrirCarta_Click()
  
  With CommonDialog1
    .CancelError = True
    .Filter = "*.rtf"
    .ShowOpen
    XFS_PathCarta = .FileName
    txtPathCarta.Text = XFS_PathCarta
  End With
End Sub

Private Sub cmdAbrirEditor_Click()
  
  Dim aux As Double
  Dim XLO_RSCARTA As New ADODB.Recordset
  Dim XLT_SQL As String
  
  Dim XLI_ARQUIVO As Integer 'Identificacao do arquivo #XLI_ARQUIVO
  Dim XLL_ARQUIVO As Long
  Dim XFM_Arquivo As Variant
  Dim XLT_LINHA
  
  If Not (PanObservacao.Tag = "I") Then
      XLT_SQL = "SELECT * FROM cartas WHERE cart_cd_carta = " & Adodc1.Recordset.Fields("cart_cd_Carta")
      SubQOpenRecordset XLO_RSCARTA, XLT_SQL, Dinamico
      XFM_Arquivo = XLO_RSCARTA!cart_vb_Arquivo 'carrega o arquivo do banco para uma variavel variant no vb
      XLI_ARQUIVO = FreeFile
      Open App.Path & "\carta.txt" For Append As XLI_ARQUIVO
         XLL_ARQUIVO = LOF(XLI_ARQUIVO)
         'testa se o tamanho do arquivo = 0 e sai da funcao
         Print #XLI_ARQUIVO, XFM_Arquivo 'passa da variavel para o arquivo
      Close XLI_ARQUIVO
      aux = ShellExecute(Me.hwnd, "open", "wordpad.exe", App.Path & "\carta.txt", "", vbNormalFocus)
  Else
      If IsNull(txtPathCarta.Text) Then
         aux = ShellExecute(Me.hwnd, "open", "wordpad.exe", "", "", vbNormalFocus)
      Else
         aux = ShellExecute(Me.hwnd, "open", "wordpad.exe", txtPathCarta.Text, "", vbNormalFocus)
      End If
  End If
  XFM_Arquivo = ""

  
End Sub

Private Sub CmdConfirmar_Click()
    Dim XLT_SQL As String
    Dim XLO_RSCARTA As ADODB.Recordset
    Dim XLI_ARQUIVO As Integer

    Dim XLL_ARQUIVO As Long
    Dim XFSO_ARQUIVO As New FileSystemObject
    Dim XFF_ARQUIVO As File
    Dim XFTS_TS As TextStream
    Dim XFM_Arquivo As Variant
'
' '   On Error GoTo MostraErro
'
'    'Função para verificar preenchimento dos campos obrigatórios
'    If (FunObrigatorioTXT(TxtDescricao, "O preenchimento do campo Descrição é obrigatório.")) Then
'       Exit Sub
'    End If
'    If PanObservacao.Tag = "I" Then
'      If (FunObrigatorioTXT(txtPathCarta, "Um arquivo deve ser escolhido")) Then
'      Exit Sub
'      End If
'    '/888888888888888888888888888888888888888888888888888888888888888888888888
'                     ' utilizacao de FSO
'      Set XFSO_ARQUIVO = CreateObject("Scripting.FileSystemObject")
'      XFSO_ARQUIVO.CreateTextFile (App.Path & "carta.txt")
'
'      XFTS_TS = XFF_ARQUIVO.OpenAsTextStream(ForWriting)
'      XFTS_TS.Write ("Conteudo da carta")
'
'      Set XFF_ARQUIVO = XFSO_ARQUIVO.GetFile((App.Path & "carta.txt"))
'      Set XFTS_TS = XFF_ARQUIVO.OpenAsTextStream(ForReading)
'
'     '/888888888888888888888888888888888888888888888888888888888888888
'         'gravar o arquivo no banco
'      XLT_SQL = "SELECT * FROM Cartas"
'      SubQOpenRecordset XLO_RSCARTA, XLT_SQL, Dinamico
'      Conexao.BeginTrans
'         If PanObservacao.Tag = "I" Then
'           XLO_RSCARTA.AddNew
'         End If
'         XLO_RSCARTA.Fields("cart_tx_Descricao") = TxtDescricao.Text
'         XLO_RSCARTA.Fields("cart_vb_arquivo").Value = XFF_ARQUIVO 'XFM_ARQUIVO
'         XLO_RSCARTA.Update
'      Conexao.CommitTrans
'    Else 'ALteracao
'      XLT_SQL = "SELECT * FROM cartas WHERE cart_cd_carta = " & Adodc1.Recordset.Fields("cart_cd_Carta")
'      SubQOpenRecordset XLO_RSCARTA, XLT_SQL, Dinamico
'      XFM_Arquivo = XLO_RSCARTA!cart_vb_Arquivo 'carrega o arquivo do banco para uma variavel variant no vb
'      XLI_ARQUIVO = FreeFile
'      Open App.Path & "\carta.doc" For Append As XLI_ARQUIVO
'         XLL_ARQUIVO = LOF(XLI_ARQUIVO)
'         Print #XLI_ARQUIVO, XFM_Arquivo 'passa da variavel para o arquivo
'      Close XLI_ARQUIVO
'                              'FSO
'      Set XFSO_ARQUIVO = CreateObject("Scripting.FileSystemObject")
'      XFSO_ARQUIVO.CreateTextFile ("carta.doc")
'      Set XFF_ARQUIVO = XFSO_ARQUIVO.GetFile(App.Path & "\carta.doc")
'      Set XFTS_TS = XFF_ARQUIVO.OpenAsTextStream(ForReading)
'                              'Gravando no banco
'      Conexao.BeginTrans
'         XLO_RSCARTA.Fields("cart_tx_Descricao") = TxtDescricao.Text
'         XLO_RSCARTA.Fields("cart_vb_arquivo").Value = XFF_ARQUIVO
'         XLO_RSCARTA.Update
'      Conexao.CommitTrans
'    End If
'    Close XLI_ARQUIVO
'    XLO_RSCARTA.Close
'    Set XLO_RSCARTA = Nothing
'
'    PanObservacao.Visible = False
'    PanBotoes.Enabled = True
'    SubHabilitaBotoes
'
'    subrecarregadados Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
'    Adodc1.Refresh
' '   If PanObservacao.Tag <> "I" Then
''        cboLocalizar.BoundText = PanObservacao.Tag
''        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
''    End If
'    'Limpa o arquivo temporario
'    Dim XFM As Variant
'    XFM = ""
'    XLI_ARQUIVO = FreeFile
'    Open App.Path & "\carta.doc" For Append As XLI_ARQUIVO
'      XLL_ARQUIVO = LOF(XLI_ARQUIVO)
'      'testa se o tamanho do arquivo = 0 e sai da funcao
'      Print #XLI_ARQUIVO, XFM 'passa da variavel para o arquivo
'    Close XLI_ARQUIVO
'    TDBGrid1.Enabled = True
'    TDBGrid1.SetFocus
'
'    Exit Sub
'
'MostraErro:
'
'   'Enumera a coleção de erros e apresenta
'   ' as propriedades de cada erro ocorrido
'   'For Each errLoop In Conexao.Errors
'    '  StrError = "Error #" & errLoop.Number & vbCr & _
'         "   " & errLoop.Description & vbCr & _
'         "   (Source: " & errLoop.Source & ")" & vbCr & _
'         "   (SQL State: " & errLoop.SQLState & ")" & vbCr & _
'         "   (NativeError: " & errLoop.NativeError & ")" & vbCr
'
''   MsgBox StrError, vbCritical + vbOKOnly, "ATENÇÃO !!!"
''   Next
''
''   Resume Next
''     MsgBox "Este código reduzido já foi atribuído a outro item do plano de contas.", vbCritical + vbOKOnly, "ATENÇÃO"
''     Exit Sub

    'Função para verificar preenchimento dos campos obrigatórios
    If (FunObrigatorioTXT(TxtDescricao, "O preenchimento do campo Descrição é obrigatório.")) Then
       Exit Sub
    End If
    If PanObservacao.Tag = "I" Then
      'If (FunObrigatorioTXT(txtPathCarta, "Um arquivo deve ser escolhido")) Then
      '   Exit Sub
      'End If
      'gravar o arquivo no banco
      XLT_SQL = "SELECT * FROM Cartas"
      SubQOpenRecordset XLO_RSCARTA, XLT_SQL, Dinamico
      Conexao.BeginTrans
         XLO_RSCARTA.AddNew
         XLO_RSCARTA.Fields("cart_tx_Descricao") = TxtDescricao.Text
         XLO_RSCARTA.Fields("cart_tx_texto") = RTBox.TextRTF
         XLO_RSCARTA.Update
      Conexao.CommitTrans
    Else 'ALteracao
      XLT_SQL = "SELECT * FROM cartas WHERE cart_cd_carta = " & Adodc1.Recordset.Fields("cart_cd_Carta")
      SubQOpenRecordset XLO_RSCARTA, XLT_SQL, Dinamico
      Conexao.BeginTrans
         XLO_RSCARTA.Fields("cart_tx_Descricao") = TxtDescricao.Text
         XLO_RSCARTA.Fields("cart_tx_texto").Value = RTBox.TextRTF
         XLO_RSCARTA.Update
      Conexao.CommitTrans
    End If
    XLO_RSCARTA.Close
    Set XLO_RSCARTA = Nothing

    PanObservacao.Visible = False
    PanBotoes.Enabled = True
    SubHabilitaBotoes

    subrecarregadados Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    Adodc1.Refresh

     
End Sub

Private Sub cmdCriarCarta_Click()
  With CommonDialog1
    .CancelError = True
    .Filter = "*.rtf"
    .ShowOpen
    XFS_PathCarta = .FileName
    txtPathCarta = XFS_PathCarta
  End With
  
End Sub

Private Sub CmdImprimir_Click()
    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Observações")
End Sub

Private Sub CmdRemoverFiltro_Click()
    Dim XLI_FIL As Integer
    
    SubDesabilitaBotoes

    Set Formulario = TelaManTabCartas
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"

    
    If Filtrou = True Then
        XLI_FIL = ContFil
        subContaReg Adodc1
        Adodc1.Refresh
        subrecarregadados Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        ContFil = XLI_FIL
    End If

    SubHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdOrdem.Enabled = False
       CmdImprimir.Enabled = False
    End If
    
    TDBGrid1.SetFocus

End Sub

Private Sub CmdRemoverFiltro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Mostra os filtros ativos da tabela"
End Sub

Private Sub CmdAlterar_Click()
    PanBotoes.Enabled = False
    PanObservacao.Visible = True
    PanObservacao.Enabled = True
    PanObservacao.Left = (TDBGrid1.Width - PanObservacao.Width) / 2
    PanObservacao.Top = (TDBGrid1.Height - PanObservacao.Height) / 2
    PanObservacao.Tag = Adodc1.Recordset.Fields("cart_cd_carta")
    PanObservacao.Caption = "Alteração de Cartas"
    SubDesabilitaBotoes
    
    TxtDescricao.Text = Adodc1.Recordset.Fields("cart_tx_Descricao")
    RTBox.TextRTF = Adodc1.Recordset.Fields("cart_tx_texto")
    'txtPathCarta.Text = ""
    'cmdAbrirEditor.SetFocus
    TDBGrid1.Enabled = False
    
    
End Sub

Private Sub CmdRetornar_Click()
    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.CancelUpdate
    PanBotoes.Enabled = True
    PanObservacao.Visible = False
    PanObservacao.Enabled = False
    SubHabilitaBotoes
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    'Limpa o arquivo temporario
    Dim XFM As Variant
    Dim XLI_ARQUIVO As Integer
    Dim XLL_ARQUIVO As Long
    XFM = ""
    XLI_ARQUIVO = FreeFile
    Open App.Path & "\carta.doc" For Append As XLI_ARQUIVO
      XLL_ARQUIVO = LOF(XLI_ARQUIVO)
      'testa se o tamanho do arquivo = 0 e sai da funcao
      Print #XLI_ARQUIVO, XFM 'passa da variavel para o arquivo
    Close XLI_ARQUIVO
End Sub

Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Imprime os dados da tabela"
End Sub

Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    Unload Me
End Sub

Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub
Private Sub CmdOrdem_Click()
    
    SubDesabilitaBotoes
    Set Formulario = TelaManTabObservacoes
    subTelaValoresGlobais "P"
    TelaOrdem.Show 1
    subTelaValoresGlobais "G"
    SubHabilitaBotoes
    
    lblLocalizar.Caption = "Localizar " & xLocalDescricao & ": "
    cboLocalizar.ListField = xLocalCampo
    cboLocalizar.Text = ""
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus

End Sub

Private Sub CmdOrdem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Ordena os registros atuais da tabela"
End Sub

Private Sub CmdExcluir_Click()
    Dim XLI_OPCAO As Integer
    Dim XLI_HOUVEERRO As Integer

    XLI_HOUVEERRO = False
    
    SubDesabilitaBotoes
    
    XLI_OPCAO = MsgBox("Confirma Remoção do Registro ?", 273, "CUIDADO")
    
    If XLI_OPCAO = 1 Then  'Confirma exclusão
       On Error GoTo RotuloErro
       Conexao.BeginTrans
       Conexao.Execute "DELETE FROM cartas WHERE cart_cd_carta=" & Adodc1.Recordset.Fields("cart_cd_Carta")
       Conexao.CommitTrans
       If XLI_HOUVEERRO = False Then
          subConectarControleDados Adodc1, FRM_FiltroAtual, "Estatico"
          Adodc1.Refresh
          Quantid = Quantid - 1
          MdiPrincipal.BarraStatus.Panels(3).Text = Str(Quantid) + "/" + Str(Quantid)
          If ContFil <> 0 Then
             subContaReg Adodc1
          End If
       End If
       
       SubHabilitaBotoes
       
       If FunTabelaVazia(Adodc1) = True Then
          CmdAlterar.Enabled = False
          CmdExcluir.Enabled = False
          CmdImprimir.Enabled = False
          CmdOrdem.Enabled = False
       End If
       
    Else
       SubHabilitaBotoes
    End If
       
    cboLocalizar.Text = ""
    TDBGrid1.SetFocus
    TDBGrid1.Col = 0
    
    Exit Sub
    
RotuloErro:
    XLI_HOUVEERRO = True
    funTrataErros (ComMensagem)
    Resume Next
End Sub
Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
End Sub

Private Sub CmdInserir_Click()
    PanBotoes.Enabled = False
    PanObservacao.Visible = True
    PanObservacao.Enabled = True
    PanObservacao.Left = (TDBGrid1.Width - PanObservacao.Width) / 2
    PanObservacao.Top = (TDBGrid1.Height - PanObservacao.Height) / 2
    PanObservacao.Caption = "Cadastro de Nova Carta"
    'OLE.AutoVerbMenu
    PanObservacao.Tag = "I"
    SubDesabilitaBotoes
    TxtDescricao.Text = ""
    cboLocalizar.Text = ""
    TDBGrid1.Enabled = False
    TxtDescricao.SetFocus
    'txtPathCarta.Text = ""
End Sub

Private Sub CmdInserir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Insere um novo registro na tabela"
End Sub
  
Private Sub Adodc1_Error(ByVal Number As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Select Case funTrataErros(ComMensagem)
        Case ErrChaveDuplicada
            CancelDisplay = False
        Case ErrRegistroExcluido
            CancelDisplay = False
            Adodc1.Refresh
            CmdInserir.Caption = "&Inserir"
            CmdExcluir.Caption = "&Excluir"
            SubHabilitaBotoes
        Case ErrExclusaoRelacionamento
            CancelDisplay = False
    End Select
End Sub

Private Sub TDBGrid1_DblClick()
    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaManTabObservacoes
        subTelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, X
        subTelaValoresGlobais "G"
    End If
    
    SubHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) = True Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False 'Desabilita o botão Excluir
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If
End Sub


Private Sub Form_Activate()
    Set Formulario = TelaCartas
    If Not PrimeiraVez Then
        subTelaValoresGlobais "P"
    End If
    
    PrimeiraVez = False

End Sub

Private Sub Form_Load()
  Dim XFS_PathCarta As String
    Set Formulario = TelaCartas
    
    Call subAjustaTelaLoc
    
    subManutencaoJanelasAtivas "I", "TelaCartas"
         
    'Associa um banco de dados à tabela Observações
    subCarregaDados Adodc1, "Cartas", "cart_tx_Descricao", _
                            "", "SELECT * FROM Cartas"
                               
    SubHabilitaBotoes
    
    filtra_especial = True
    
    subTelaValoresGlobais "G"
    PrimeiraVez = True
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaManTabCartas"
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

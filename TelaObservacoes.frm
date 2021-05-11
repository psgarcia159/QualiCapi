VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaManTabObservacoes 
   Caption         =   "Tabela de Observações"
   ClientHeight    =   4545
   ClientLeft      =   1770
   ClientTop       =   3855
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   8910
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanObservacao 
      Height          =   1515
      Left            =   1635
      TabIndex        =   11
      Top             =   870
      Visible         =   0   'False
      Width           =   5280
      _Version        =   65536
      _ExtentX        =   9313
      _ExtentY        =   2672
      _StockProps     =   15
      Caption         =   " Cadastro de Observação"
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
      Begin VB.TextBox TxtObservacao 
         Height          =   315
         Left            =   1545
         MaxLength       =   30
         TabIndex        =   0
         Top             =   570
         Width           =   3600
      End
      Begin VB.CommandButton CmdConfirmar 
         Caption         =   "&Confirmar"
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
         Left            =   3120
         TabIndex        =   8
         Top             =   1065
         Width           =   975
      End
      Begin VB.CommandButton CmdRetornar 
         Caption         =   "&Retornar"
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
         Left            =   4170
         TabIndex        =   9
         Top             =   1065
         Width           =   975
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
         Left            =   120
         TabIndex        =   12
         Top             =   585
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4590
      Top             =   3285
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
      Bindings        =   "TelaObservacoes.frx":0000
      Height          =   3690
      Left            =   45
      TabIndex        =   13
      Top             =   15
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   6509
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Descrição"
      Columns(0).DataField=   "obse_tx_observacao"
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
      PrintInfos(0)._StateFlags=   3
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
      Left            =   120
      TabIndex        =   10
      Top             =   4125
      Width           =   8760
      _Version        =   65536
      _ExtentX        =   15452
      _ExtentY        =   635
      _StockProps     =   15
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
         TabIndex        =   1
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton CmdOrdem 
         Caption         =   "&Ordem"
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
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton CmdRemoverFiltro 
         Caption         =   "Se&leções"
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaObservacoes.frx":0015
      Height          =   315
      Left            =   2760
      TabIndex        =   14
      Top             =   3750
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
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
      TabIndex        =   15
      Top             =   3780
      Width           =   2670
   End
End
Attribute VB_Name = "TelaManTabObservacoes"
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

Private Sub subHabilitaBotoes()

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
Private Sub subDesabilitaBotoes()
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

Private Sub cboLocalizar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
End Sub


Private Sub CmdConfirmar_Click()
    
    Dim XLT_SQL As String
    Dim XLO_OBSERVACAO As New ADODB.Recordset
    
    
    'Função para verificar preenchimento dos campos obrigatórios
    If TxtObservacao.Text = "" Then
        MsgBox "O preenchimento do campo Observação é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
    
    If XInserir = "I" Then
      'se for inserir procura um registro já existentente
      SubQOpenRecordset XLO_OBSERVACAO, "SELECT * FROM Observacoes WHERE obse_tx_Observacao='" & TxtObservacao.Text & "'", Dinamico
      If Not XLO_OBSERVACAO.EOF Then
        MsgBox "Observação já existente", vbCritical, "ATENÇÃO"
        XLO_OBSERVACAO.Close
        Set XLO_OBSERVACAO = Nothing
        Exit Sub
      Else
      'se não encontrar insere e pega a a chave autonumérica
        XLO_OBSERVACAO.AddNew
        XLO_OBSERVACAO!obse_tx_Observacao = TxtObservacao.Text
        XLO_OBSERVACAO.Update
        Chave = XLO_OBSERVACAO!obse_cd_Observacao
        XLO_OBSERVACAO.Close
        Set XLO_OBSERVACAO = Nothing
      End If
    Else
      If PanObservacao.Tag <> TxtObservacao.Text Then
        'se for alterar e tiver alterado procura um registro já existentente
        SubQOpenRecordset XLO_OBSERVACAO, "SELECT * FROM Observacoes WHERE obse_tx_Observacao='" & TxtObservacao.Text & "'", Estatico
        If Not XLO_OBSERVACAO.EOF Then
          MsgBox "Observação já existente", vbCritical, "ATENÇÃO"
          XLO_OBSERVACAO.Close
          Set XLO_OBSERVACAO = Nothing
          Exit Sub
        Else
          XLO_OBSERVACAO.Close
          Set XLO_OBSERVACAO = Nothing
          ' se não encontrar o registro
          XLT_SQL = "UPDATE Observacoes SET obse_tx_observacao = '" & TxtObservacao.Text & "'" & _
            " WHERE obse_cd_observacao = " & Chave
          Conexao.Execute XLT_SQL
        End If
      Else
        PanObservacao.Visible = False
        PanBotoes.Enabled = True
        subHabilitaBotoes
        Exit Sub
      End If
    End If
    
    
'    Conexao.BeginTrans
'    If PanObservacao.Tag = "I" Then
'        XLT_SQL = "INSERT INTO Observacoes (obse_tx_observacao) VALUES ('" & TxtObservacao.Text & "')"
'              Conexao.Execute XLT_SQL
'    Else    'Alteração
'        XLT_SQL = "UPDATE Observacoes SET obse_tx_observacao = '" & TxtObservacao.Text & "'" & _
'                " WHERE obse_cd_observacao = " & PanObservacao.Tag
'        Conexao.Execute XLT_SQL
'    End If
'    Conexao.CommitTrans
    
    PanObservacao.Visible = False
    PanBotoes.Enabled = True
    subHabilitaBotoes
    
    FRM_Quantid = funOperacaoQuantid("+")
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    cboLocalizar.BoundText = Chave
    Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    Exit Sub

End Sub

Private Sub CmdImprimir_Click()
    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Observações")
End Sub

Private Sub CmdRemoverFiltro_Click()
    Dim XLI_FIL As Integer
    
    subDesabilitaBotoes

    Set Formulario = TelaManTabObservacoes
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"
    
    If Filtrou = True Then
        XLI_FIL = ContFil
        subContaReg Adodc1
        Adodc1.Refresh
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        ContFil = XLI_FIL
    End If

    subHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdOrdem.Enabled = False
       CmdImprimir.Enabled = False
    End If
    
    TDBGrid1.SetFocus

End Sub

Private Sub CmdRemoverFiltro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Mostra os filtros ativos da tabela"
End Sub

Private Sub CmdAlterar_Click()
    PanBotoes.Enabled = False
    PanObservacao.Visible = True
    PanObservacao.Enabled = True
    PanObservacao.Left = (TDBGrid1.Width - PanObservacao.Width) / 2
    PanObservacao.Top = (TDBGrid1.Height - PanObservacao.Height) / 2
    PanObservacao.Tag = Adodc1.Recordset.Fields("obse_cd_observacao")
    PanObservacao.Caption = " Alteração de Observação"
    subDesabilitaBotoes
    XInserir = "A"
    PanObservacao.Tag = Adodc1.Recordset.Fields("obse_tx_observacao")
    TxtObservacao.Text = Adodc1.Recordset.Fields("obse_tx_observacao")
    Chave = Adodc1.Recordset.Fields("obse_cd_observacao")
    TDBGrid1.Enabled = False
    TxtObservacao.SetFocus
End Sub

Private Sub CmdRetornar_Click()
    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.CancelUpdate
    PanBotoes.Enabled = True
    PanObservacao.Visible = False
    PanObservacao.Enabled = False
    subHabilitaBotoes
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub

Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Imprime os dados da tabela"
End Sub

Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    Unload Me
End Sub

Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub
Private Sub CmdOrdem_Click()
    
    subDesabilitaBotoes
    Set Formulario = TelaManTabObservacoes
    subTelaValoresGlobais "P"
    TelaOrdem.Show 1
    subTelaValoresGlobais "G"
    subHabilitaBotoes
    
    lblLocalizar.Caption = "Localizar " & xLocalDescricao & ": "
    cboLocalizar.ListField = xLocalCampo
    cboLocalizar.Text = ""
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus

End Sub

Private Sub CmdOrdem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Ordena os registros atuais da tabela"
End Sub

Private Sub CmdExcluir_Click()
    Dim XLI_OPCAO As Integer
    Dim XLI_BOOKMARK As Integer
    
    subDesabilitaBotoes
    
    XLI_OPCAO = MsgBox("Confirma Remoção do Registro ?", 273, "CUIDADO")
    
    If XLI_OPCAO = 1 Then  'Confirma exclusão
      On Error GoTo RotuloErro
      Conexao.BeginTrans
      Conexao.Execute "DELETE FROM observacoes WHERE obse_cd_observacao=" & Adodc1.Recordset.Fields("obse_cd_observacao")
      Conexao.CommitTrans
       
      FRM_Quantid = funOperacaoQuantid("-")
      XLI_BOOKMARK = TDBGrid1.bookmark
      subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
      If Not Adodc1.Recordset.EOF Then
         If XLI_BOOKMARK <= Adodc1.Recordset.RecordCount Then
           TDBGrid1.bookmark = XLI_BOOKMARK
         Else
           TDBGrid1.bookmark = XLI_BOOKMARK - 1
         End If
       End If
       
       subHabilitaBotoes
       
       If FunTabelaVazia(Adodc1) = True Then
          CmdAlterar.Enabled = False
          CmdExcluir.Enabled = False
          CmdImprimir.Enabled = False
          CmdOrdem.Enabled = False
       End If
    Else
       subHabilitaBotoes
    End If
       
    cboLocalizar.Text = ""
    TDBGrid1.SetFocus
    TDBGrid1.Col = 0
    
    Exit Sub
    
RotuloErro:
    Conexao.RollbackTrans
    If Err.Number = -2147467259 Then
      MsgBox "Essa observação não pode ser excluída pois já possui algum título relacionado!", vbCritical, "ATENÇÃO"
    End If
    subHabilitaBotoes
End Sub
Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
End Sub

Private Sub CmdInserir_Click()
    
    PanBotoes.Enabled = False
    PanObservacao.Visible = True
    PanObservacao.Enabled = True
    PanObservacao.Left = (TDBGrid1.Width - PanObservacao.Width) / 2
    PanObservacao.Top = (TDBGrid1.Height - PanObservacao.Height) / 2
    PanObservacao.Caption = " Inserção de Observação"
    XInserir = "I"
    PanObservacao.Tag = ""
    subDesabilitaBotoes
    TxtObservacao.Text = ""
    cboLocalizar.Text = ""
    TDBGrid1.Enabled = False
    TxtObservacao.SetFocus
    
End Sub

Private Sub CmdInserir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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
            subHabilitaBotoes
        Case ErrExclusaoRelacionamento
            CancelDisplay = False
    End Select
End Sub

Private Sub TDBGrid1_DblClick()
    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    cboLocalizar.Text = ""
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaManTabObservacoes
        subTelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, x
        subTelaValoresGlobais "G"
    End If
    
    subHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) = True Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False 'Desabilita o botão Excluir
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If
End Sub


Private Sub Form_Activate()
    
    Set Formulario = TelaManTabObservacoes
    If Not PrimeiraVez Then
        subTelaValoresGlobais "P"
    End If
    
    PrimeiraVez = False

End Sub

Private Sub Form_Load()
    
    Set Formulario = TelaManTabObservacoes
    
    Call subAjustaTelaLoc
    
    subManutencaoJanelasAtivas "I", "TelaManTabObservacoes"
         
    'Associa um banco de dados à tabela Observações
    subCarregaDadosNV Adodc1, "observacoes", "obse_tx_observacao", _
                            "", "SELECT * FROM observacoes"
                               
    subHabilitaBotoes
    
    filtra_especial = False
    
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
    subManutencaoJanelasAtivas "R", "TelaManTabObservacoes"
End Sub

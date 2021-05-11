VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaFornecedoresRamosAtividades 
   Caption         =   "Fornecedores x Ramos de Atividades"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   1245
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   10350
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanTipo 
      Height          =   4920
      Left            =   1035
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   7560
      _Version        =   65536
      _ExtentX        =   13335
      _ExtentY        =   8678
      _StockProps     =   15
      Caption         =   "Cadastro de Fornecedores x Ramos de Atividades"
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
      Begin VB.ListBox lsbRequisitosLegais 
         Height          =   2985
         ItemData        =   "TelaFornecedoresRamosAtividades.frx":0000
         Left            =   2220
         List            =   "TelaFornecedoresRamosAtividades.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   1200
         Width           =   5170
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
         Height          =   375
         Left            =   5355
         TabIndex        =   1
         Top             =   4365
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelar 
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
         Height          =   375
         Left            =   6405
         TabIndex        =   2
         Top             =   4365
         Width           =   975
      End
      Begin MSDataListLib.DataCombo CboRamoAtv 
         Bindings        =   "TelaFornecedoresRamosAtividades.frx":0004
         Height          =   330
         Left            =   2220
         TabIndex        =   13
         Top             =   720
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "raat_tx_descricao"
         BoundColumn     =   "raat_cd_RamoAtividade"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblRequisitosLegais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Requisitos Legais:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label LblNomeFantasia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ramos de Atividades:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaFornecedoresRamosAtividades.frx":001F
      Height          =   3690
      Left            =   45
      TabIndex        =   12
      Top             =   0
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   6509
      _LayoutType     =   4
      _RowHeight      =   25
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Ramos de Atividades"
      Columns(0).DataField=   "raat_tx_descricao"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Requisitos Legais"
      Columns(1).DataField=   "qtdRequisitos"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=26458"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=26379"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=20"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=873"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=794"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=529"
      Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(1)._MinWidth=56554156"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaFornecedoresRamosAtividades.frx":0034
      Height          =   315
      Left            =   2925
      TabIndex        =   11
      Top             =   3765
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "raat_tx_Descricao"
      BoundColumn     =   "raat_cd_ramoatividade"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc AdodcRamoAtv 
      Height          =   375
      Left            =   8175
      Top             =   3750
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "AdodcRamoAtv"
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
   Begin Threed.SSPanel PanBotoes 
      Height          =   360
      Left            =   360
      TabIndex        =   8
      Top             =   4140
      Width           =   7635
      _Version        =   65536
      _ExtentX        =   13467
      _ExtentY        =   635
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
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
         Left            =   1150
         TabIndex        =   17
         Top             =   15
         Width           =   1020
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
         Left            =   6555
         TabIndex        =   7
         Top             =   15
         Width           =   1020
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
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
         Left            =   5475
         TabIndex        =   6
         Top             =   15
         Width           =   1020
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
         Left            =   4395
         TabIndex        =   5
         Top             =   15
         Width           =   1020
      End
      Begin VB.CommandButton Botao_Fil 
         Caption         =   "&Seleções"
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
         Left            =   3315
         TabIndex        =   4
         Top             =   15
         Width           =   1020
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
         Left            =   2235
         TabIndex        =   3
         Top             =   15
         Width           =   1020
      End
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
         Left            =   60
         TabIndex        =   0
         Top             =   15
         Width           =   1020
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8160
      Top             =   4080
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
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
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Descrição:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   165
      TabIndex        =   10
      Top             =   3810
      Width           =   2670
   End
End
Attribute VB_Name = "TelaFornecedoresRamosAtividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Tipos() As String * 1
Dim Quantid As Long
Dim Fornecedor As Long
Dim ResFormulario As Object
Dim KeyAscii As Integer
Dim errLoop As ADODB.Error
Dim StrError As String
Dim xAlterar As String

Dim Chave As Integer
'*****************************************************************************************
'Variaveis que guardarão valores específicos de cada formulário
Dim PRIMEIRAVEZ As Boolean

Dim FRM_CONTFIL As Integer ' Contador que armazena o número de filtros ativos
Dim FRM_ORDEM As String ' Nome do Campo que tem a chave da Tabela ativa
Dim FRM_CAMPOS_OBRIG As String ' Lista dos campos que são obrigatórios na tabela
Dim FRM_TIPOVAR As String ' Tipo do campo a ser filtrado
Dim FRM_FILTROORIG As String ' Indica os filtros determinados pelo próprio sistema
Dim FRM_FILTROATUAL As String ' Indica o filtro atual da tabela corrente
Dim FRM_FILTRA_ESPECIAL As Integer ' Indica se a tabela utilizada possui filtros pré-definidos pelo sistema em FiltroOrig
Dim FRM_TEM_FILTRO As Boolean ' Informa se já foi relizado um filtro na tabela pelo usuário
Dim FRM_CONTELAORDEM As Integer ' Indica o número de chaves utilizados para ordenação
Dim FRM_QUANTID As Long  ' Quantidade de registros da tabela
Dim FRM_FILTROU As Boolean ' Informa se Algum campo foi filtrado
Dim FRM_ORDENS() As Integer    ' Vetor que mantem a ordem dos campos
Dim FRM_FILTROS() As String ' Array para armazenar os Filtros da Tabela

Dim FRM_BARRAFERRAMENTASPRESSIONADA As Boolean
Dim FRM_VARIASVEZES As Boolean    'Indica se vai executar a insercao na TelaPadrao varias ou uma vez
Dim FRM_XINSERIR As String        'Flag para identificar se entrou na rotina para inserir ou alterar ('I'Inserir, 'A'Alterar, 'C'Consultar)
Dim FRM_XORIGEM As String         'Identifica a origem (que formulario) chamou o formulário ativo
Dim FRM_SQLATIVO As String        'Atualizar o SQL Ativo após modificações de filtro e ordem

Sub subTelaValoresGlobais(GuardaOUPegaValores As String)
    Dim Count As Integer, Count2 As Integer
    If GuardaOUPegaValores = "G" Then 'Guardando variáveis de definição da tela
        FRM_CONTFIL = ContFil
        FRM_ORDEM = Ordem
        FRM_CAMPOS_OBRIG = Campos_Obrig
        FRM_TIPOVAR = TipoVar
        FRM_FILTROORIG = FiltroOrig
        FRM_FILTROATUAL = FiltroAtual
        FRM_FILTRA_ESPECIAL = filtra_especial
        FRM_TEM_FILTRO = Tem_Filtro
        FRM_CONTELAORDEM = ConTelaOrdem
        FRM_QUANTID = Quantid
        FRM_FILTROU = Filtrou
        FRM_BARRAFERRAMENTASPRESSIONADA = BarraFerramentasPressionada
        FRM_VARIASVEZES = VariasVezes
        FRM_XINSERIR = XInserir
        FRM_XORIGEM = XOrigem
        FRM_SQLATIVO = SQLAtivo
        
        ReDim FRM_FILTROS(UBound(Filtros, 1), UBound(Filtros, 2))
        ReDim FRM_ORDENS(UBound(Ordens, 1)) As Integer
        
        For Count = LBound(Filtros, 1) To UBound(Filtros, 1)
            For Count2 = LBound(Filtros, 2) To UBound(Filtros, 2)
                FRM_FILTROS(Count, Count2) = Filtros(Count, Count2)
            Next Count2
        Next Count
        For Count = LBound(Ordens) To UBound(Ordens)
            FRM_ORDENS(Count) = Ordens(Count)
        Next Count
    ElseIf GuardaOUPegaValores = "P" Then  'Regravando os valores relativos a tela atual variáveis de definição da tela
        ContFil = FRM_CONTFIL
        Ordem = FRM_ORDEM
        Campos_Obrig = FRM_CAMPOS_OBRIG
        TipoVar = FRM_TIPOVAR
        FiltroOrig = FRM_FILTROORIG
        FiltroAtual = FRM_FILTROATUAL
        filtra_especial = FRM_FILTRA_ESPECIAL
        Tem_Filtro = FRM_TEM_FILTRO
        ConTelaOrdem = FRM_CONTELAORDEM
        Quantid = FRM_QUANTID
        Filtrou = FRM_FILTROU
        BarraFerramentasPressionada = FRM_BARRAFERRAMENTASPRESSIONADA
        VariasVezes = FRM_VARIASVEZES
        XInserir = FRM_XINSERIR
        XOrigem = FRM_XORIGEM
        SQLAtivo = FRM_SQLATIVO

        ReDim Filtros(UBound(FRM_FILTROS, 1), UBound(FRM_FILTROS, 2))
        ReDim Ordens(UBound(FRM_ORDENS, 1))
        
        For Count = LBound(FRM_FILTROS, 1) To UBound(FRM_FILTROS, 1)
            For Count2 = LBound(FRM_FILTROS, 2) To UBound(FRM_FILTROS, 2)
                Filtros(Count, Count2) = FRM_FILTROS(Count, Count2)
            Next Count2
        Next Count
        For Count = LBound(FRM_ORDENS) To UBound(FRM_ORDENS)
            Ordens(Count) = FRM_ORDENS(Count)
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
       Botao_Fil.Enabled = False
    Else
       Botao_Fil.Enabled = True
    End If
End Sub
Private Sub SubDesabilitaBotoes()
    'Objetivo: Desabilita os Botões da Tela
    CmdInserir.Enabled = False
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    CmdSair.Enabled = False
    CmdImprimir.Enabled = False
    CmdOrdem.Enabled = False
End Sub

Private Sub Botao_Fil_Click()
    Dim Fil As Integer
    SubDesabilitaBotoes
    
    Set Formulario = TelaFornecedoresRamosAtividades
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"

    If Filtrou = True Then
        Fil = ContFil
        subContaReg Adodc1
        subrecarregadadosNV Adodc1, "", FRM_ORDEM, "", FRM_FILTROATUAL
        ContFil = Fil
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

Private Sub Botao_Fil_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Mostra os filtros ativos da tabela"
End Sub

Private Sub cboLocalizar_Click(Area As Integer)
    If Area = 2 Then
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If

End Sub

Private Sub cboLocalizar_KeyUp(KeyCode As Integer, Shift As Integer)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If

End Sub


Private Sub cboLocalizar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If

End Sub

Private Sub CboRamoAtv_Change()
    Dim XLO_RsRequisitosLegais As New ADODB.Recordset
    Dim XLO_RsRequisitosLegaisRamosAtividades As New ADODB.Recordset
    Dim XLI_I As Integer
    
    If CboRamoAtv.SelectedItem > 0 Then
        SubQOpenRecordset XLO_RsRequisitosLegais, "SELECT * FROM ConsCOMPRamosAtividadesRequiLegais WHERE raat_cd_ramoatividade = " & CboRamoAtv.BoundText & " ORDER BY rele_tx_descricao ", "Estatico"
        SubQOpenRecordset XLO_RsRequisitosLegaisRamosAtividades, "SELECT * FROM FornCli_Ramos_Requisitos WHERE raat_cd_ramoatividade = " & CboRamoAtv.BoundText & " AND focl_cd_FornCli = " & Fornecedor & "", "Estatico"
    
        lsbRequisitosLegais.Clear
        'Preenche o listbox com os requitos legais
        If Not XLO_RsRequisitosLegais.EOF Then
    
            XLO_RsRequisitosLegais.MoveFirst
            While Not XLO_RsRequisitosLegais.EOF
            
              lsbRequisitosLegais.AddItem (XLO_RsRequisitosLegais.Fields("rele_cd_requisitolegal") & " - " & XLO_RsRequisitosLegais.Fields("rele_tx_descricao"))
              XLO_RsRequisitosLegais.MoveNext
              
            Wend
        End If


        If xAlterar = "A" And Not XLO_RsRequisitosLegaisRamosAtividades.EOF Then
    
            If Not XLO_RsRequisitosLegaisRamosAtividades.EOF Then
    
                XLO_RsRequisitosLegaisRamosAtividades.MoveFirst
    
                While Not XLO_RsRequisitosLegaisRamosAtividades.EOF
    
                    For XLI_I = 0 To lsbRequisitosLegais.ListCount - 1
    
                        'Seleciona os requisitos legais dos ramos de atividades
                        If XLO_RsRequisitosLegaisRamosAtividades!rele_cd_requisitolegal = CInt(RTrim(LTrim(Split(lsbRequisitosLegais.List(XLI_I), "-")(0)))) Then
                             lsbRequisitosLegais.Selected(XLI_I) = True
                        End If
    
                    Next XLI_I
    
                  XLO_RsRequisitosLegaisRamosAtividades.MoveNext
    
                Wend
    
            End If
    
        End If
    
        XLO_RsRequisitosLegais.Close
        XLO_RsRequisitosLegaisRamosAtividades.Close
End If
End Sub

Private Sub CmdAlterar_Click()
    TelaFornecedoresRamosAtividades.MousePointer = vbHourglass
    PanBotoes.Enabled = False
    PanTipo.Enabled = True
    PanTipo.Tag = Adodc1.Recordset.Fields("raat_cd_ramoatividade")
    PanTipo.Top = (TDBGrid1.Height - PanTipo.Height) / 2
    PanTipo.Left = (TDBGrid1.Width - PanTipo.Width) / 2
    PanTipo.Caption = "Alteração de Requisitos Legais"
    PanTipo.Visible = True
    CboRamoAtv.BoundText = Adodc1.Recordset.Fields("raat_cd_RamoAtividade")
    xAlterar = "A"
    CboRamoAtv_Change
    CboRamoAtv.Enabled = False
    SubDesabilitaBotoes
    TDBGrid1.Enabled = False
    TelaFornecedoresRamosAtividades.MousePointer = vbDefault
End Sub

Private Sub CmdCancelar_Click()
    If Not FunTabelaVazia(Adodc1) Then Adodc1.Recordset.CancelUpdate
    PanBotoes.Enabled = True
    PanTipo.Visible = False
    SubHabilitaBotoes
    CmdConfirmar.Caption = "Confirmar"
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub


Private Sub CmdConfirmar_Click()
    Dim Sql As String
    Dim SqlRequisitos As String
    Dim XLI_I As Integer
    Dim Count As Integer
    Dim ResExiste As New ADODB.Recordset
    
    On Error GoTo MostraErro
    
'    Função para verificar preenchimento dos campos obrigatórios
    
    If CboRamoAtv.Text = "" Then
        MsgBox "O Ramo de Atividade é um campo obrigatório.", vbInformation + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
    
    For XLI_I = 0 To lsbRequisitosLegais.ListCount - 1

        If lsbRequisitosLegais.Selected(XLI_I) = True Then
             Count = Count + 1
        End If

    Next XLI_I

    SubQOpenRecordset ResExiste, "SELECT count(raat_cd_ramoatividade) as RamosAtividade FROM FornCli_Ramos_Requisitos WHERE raat_cd_ramoatividade = " & CboRamoAtv.BoundText & " AND focl_cd_FornCli = " & Fornecedor & "", Dinamico
    If (PanTipo.Tag = "I" And ResExiste!RamosAtividade > 0) Then
        MsgBox "Este Ramos de Atividade já foi cadastrado. Tente outro.", vbInformation + vbOKOnly, "ATENÇÃO"
        CboRamoAtv.SetFocus
        ResExiste.Close
        Exit Sub
    End If
    ResExiste.Close
        
    Conexao.BeginTrans
    Screen.MousePointer = vbHourglass
    
    If PanTipo.Tag = "I" Then
        Sql = "INSERT INTO FornCliente_RamosAtividades (focl_cd_FornCli, raat_cd_RamoAtividade) " & _
               "VALUES (" & Fornecedor & " , " & CboRamoAtv.BoundText & ")"
               Conexao.Execute Sql
               
         MdiPrincipal.BarraStatus.Panels(3).Text = CStr((Adodc1.Recordset.RecordCount + 1)) + "/" + CStr((Adodc1.Recordset.RecordCount + 1))
    End If
    
    If Count > 0 Then
        'Cadastro de Ramos de Atividades
        Conexao.Execute "DELETE FROM FornCli_Ramos_Requisitos WHERE focl_cd_forncli = " & Fornecedor & _
           " AND raat_cd_ramoatividade = " & CboRamoAtv.BoundText & ""
        
        For XLI_I = 0 To lsbRequisitosLegais.ListCount - 1
        
           'Cadastra os ramos de atividades do fornecedor
           If lsbRequisitosLegais.Selected(XLI_I) = True Then
               Conexao.Execute "INSERT INTO FornCli_Ramos_Requisitos (focl_cd_forncli, raat_cd_ramoatividade, rele_cd_requisitolegal) VALUES (" & Fornecedor & ", " & CboRamoAtv.BoundText & ", " & CInt(RTrim(LTrim(Split(lsbRequisitosLegais.List(XLI_I), "-")(0)))) & ")"
           End If
        
        Next XLI_I
    End If
     
    Conexao.CommitTrans
    Screen.MousePointer = Default
     
    PanTipo.Visible = False
    PanTipo.Enabled = False
    PanBotoes.Enabled = True
    SubHabilitaBotoes
    
    subrecarregadadosNV Adodc1, "", FRM_ORDEM, "", FRM_FILTROATUAL
    If PanTipo.Tag <> "I" Then
        cboLocalizar.BoundText = PanTipo.Tag
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    
    Exit Sub
    
MostraErro:

   ' Enumera a coleção de erros e apresenta
   ' as propriedades de cada erro ocorrido
   For Each errLoop In Conexao.Errors
      StrError = "Error #" & errLoop.Number & vbCr & _
         "   " & errLoop.Description & vbCr & _
         "   (Source: " & errLoop.Source & ")" & vbCr & _
         "   (SQL State: " & errLoop.SQLState & ")" & vbCr & _
         "   (NativeError: " & errLoop.NativeError & ")" & vbCr

   MsgBox StrError, vbCritical + vbOKOnly, "ATENÇÃO !!!"
   Next

   Resume Next

    
End Sub

Private Sub cmdImprimir_Click()
    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Planos Alternativos")
End Sub

Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Imprime os dados da tabela"
End Sub


Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    MdiPrincipal.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub
Private Sub CmdOrdem_Click()
    
    SubDesabilitaBotoes
    
    Set Formulario = TelaFornecedoresRamosAtividades
    subTelaValoresGlobais "P"
    TelaOrdem.Show 1
    subTelaValoresGlobais "G"
    SubHabilitaBotoes
    
    lblLocalizar.Caption = "Localizar " & xLocalDescricao & ": "
    cboLocalizar.ListField = xLocalCampo
    cboLocalizar.Text = ""
    
    TDBGrid1.SetFocus
    
End Sub

Private Sub CmdOrdem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Ordena os registros atuais da tabela"
End Sub

Private Sub CmdExcluir_Click()
    Dim opcao As Integer
    Dim Houve_Erro As Integer

    Houve_Erro = False
    
    SubDesabilitaBotoes
    
    opcao = MsgBox("Confirma Remoção do Registro ?", vbQuestion, "ATENÇÃO")
    
    If opcao = 1 Then  'Confirma exclusão
       On Error GoTo RotuloErro
       Conexao.BeginTrans
       Conexao.Execute "DELETE FROM FornCliente_RamosAtividades WHERE raat_cd_RamoAtividade = " & Adodc1.Recordset.Fields("raat_cd_RamoAtividade") & " AND focl_cd_FornCli = " & Fornecedor & _
       " DELETE FROM FornCli_Ramos_Requisitos WHERE raat_cd_RamoAtividade = " & Adodc1.Recordset.Fields("raat_cd_RamoAtividade") & " AND focl_cd_FornCli = " & Fornecedor & ""
       Conexao.CommitTrans
       
       If Houve_Erro = False Then
          subrecarregadadosNV Adodc1, "", FRM_ORDEM, "", FRM_FILTROATUAL
          MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
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

   MsgBox "O Ramo de Atividade não pode ser apagado, pois está relacionado com lançamentos da empresa.", vbCritical + vbOKOnly, "ATENÇÃO"
   Houve_Erro = True
   Resume Next
End Sub
Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If CmdInserir.Caption = "&Gravar" Then
       MdiPrincipal.BarraStatus.Panels(1).Text = "Cancela as novas informações"
    Else
       MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
    End If
End Sub

Private Sub CmdInserir_Click()
    TelaFornecedoresRamosAtividades.MousePointer = vbHourglass
    PanBotoes.Enabled = False
    PanTipo.Enabled = True
    PanTipo.Tag = "I"
    PanTipo.Top = (TDBGrid1.Height - PanTipo.Height) / 2
    PanTipo.Left = (TDBGrid1.Width - PanTipo.Width) / 2
    PanTipo.Visible = True
    PanTipo.Caption = "Cadastro de Fornecedores x Ramos de Atividades"
    SubDesabilitaBotoes
    cboLocalizar.BoundText = ""
    CboRamoAtv.Text = ""
    lsbRequisitosLegais.Clear
    CboRamoAtv.Enabled = True
    TDBGrid1.Enabled = False
    xAlterar = "I"
    TelaFornecedoresRamosAtividades.MousePointer = vbDefault
End Sub

Private Sub CmdInserir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If CmdInserir.Caption = "&Gravar" Then
       MdiPrincipal.BarraStatus.Panels(1).Text = "Grava na tabela as novas informações"
    Else
       MdiPrincipal.BarraStatus.Panels(1).Text = "Insere um novo registro na tabela"
    End If
End Sub

Private Sub TDBGrid1_DblClick()
    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    cboLocalizar.Text = ""
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaFornecedoresRamosAtividades
        subTelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, x
        subTelaValoresGlobais "G"
    End If
    
    SubHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) = True Then
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If
   
End Sub

Private Sub Form_Activate()
    Set Formulario = TelaFornecedoresRamosAtividades
    If Not PRIMEIRAVEZ Then
        subTelaValoresGlobais "P"
    End If
    
    PRIMEIRAVEZ = False
    
    MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
End Sub

Private Sub Form_Load()
    Set Formulario = TelaFornecedoresRamosAtividades
    
    Call subAjustaTelaLoc
    
    TelaFornecedoresRamosAtividades.Caption = "Fornecedor: " & TelaFornecedores.Adodc1.Recordset.Fields("focl_cd_FornCli") & " - " & TelaFornecedores.Adodc1.Recordset.Fields("focl_tx_RazaoSocial")
    Fornecedor = TelaFornecedores.Adodc1.Recordset.Fields("focl_cd_FornCli")
    
    Set Formulario = TelaFornecedoresRamosAtividades
    
    subManutencaoJanelasAtivas "I", "TelaFornecedoresRamosAtividades"
    subCarregaDadosNV Adodc1, "", "focl_cd_FornCli, raat_cd_ramoatividade", "", "SELECT * FROM ConsCOMFornCli_Ramos_Requisitos WHERE focl_cd_FornCli = " & Fornecedor & ""
    
    subConectarControleDadosNV AdodcRamoAtv, "SELECT * FROM RamosAtividades ORDER BY raat_tx_descricao", Estatico
        
    SubHabilitaBotoes
    filtra_especial = True
    subTelaValoresGlobais "G"
    PRIMEIRAVEZ = True
    
    If FunTabelaVazia(Adodc1) Then
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If
    
    End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaFornecedoresRamosAtividades"
End Sub

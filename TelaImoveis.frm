VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaManImoveis 
   Caption         =   "Tabela de Imóveis"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   1470
   ClientWidth     =   8910
   Icon            =   "TelaImoveis.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   8910
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6360
      Top             =   2820
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
      Bindings        =   "TelaImoveis.frx":058A
      Height          =   3240
      Left            =   60
      TabIndex        =   10
      Top             =   15
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   5715
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Imóvel"
      Columns(0).DataField=   "Imovel"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Cliente"
      Columns(1).DataField=   "Cliente"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Empreendimento"
      Columns(2).DataField=   "Empreendimento"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Status"
      Columns(3).DataField=   "Status"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Valor do Imóvel"
      Columns(4).DataField=   "imov_vl_AVista"
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Data da Venda"
      Columns(5).DataField=   "DataVenda"
      Columns(5).NumberFormat=   "dd/mm/yy"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Entrega Chaves"
      Columns(6).DataField=   "imov_dt_EntregaChaves"
      Columns(6).NumberFormat=   "dd/mm/yy"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Tipo"
      Columns(7).DataField=   "Tipo"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Área"
      Columns(8).DataField=   "imov_nr_area"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2223"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2143"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=5398"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=5318"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=2097680"
      Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(16)=   "Column(2).Width=6403"
      Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=6324"
      Splits(0)._ColumnProps(19)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=2097680"
      Splits(0)._ColumnProps(21)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(23)=   "Column(3).Width=2143"
      Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=2064"
      Splits(0)._ColumnProps(26)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=2097680"
      Splits(0)._ColumnProps(28)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(30)=   "Column(4).Width=2540"
      Splits(0)._ColumnProps(31)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._WidthInPix=2461"
      Splits(0)._ColumnProps(33)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(34)=   "Column(4)._ColStyle=2097682"
      Splits(0)._ColumnProps(35)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(36)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(37)=   "Column(5).Width=1905"
      Splits(0)._ColumnProps(38)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(5)._WidthInPix=1826"
      Splits(0)._ColumnProps(40)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(5)._ColStyle=2097681"
      Splits(0)._ColumnProps(42)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(43)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(44)=   "Column(6).Width=1905"
      Splits(0)._ColumnProps(45)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(6)._WidthInPix=1826"
      Splits(0)._ColumnProps(47)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(48)=   "Column(6)._ColStyle=2097681"
      Splits(0)._ColumnProps(49)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(50)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(51)=   "Column(7).Width=2487"
      Splits(0)._ColumnProps(52)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(7)._WidthInPix=2408"
      Splits(0)._ColumnProps(54)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(55)=   "Column(7)._ColStyle=2097680"
      Splits(0)._ColumnProps(56)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(57)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(58)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(59)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(61)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(62)=   "Column(8)._ColStyle=2097681"
      Splits(0)._ColumnProps(63)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(64)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(65)=   "Column(8)._MinWidth=179855952"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(27)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(28)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7,.wraptext=-1"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.locked=0"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15,.alignment=3"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17,.alignment=3"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=0"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=32,.parent=13,.alignment=1"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=66,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=0"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(74)  =   "Named:id=33:Normal"
      _StyleDefs(75)  =   ":id=33,.parent=0,.alignment=2"
      _StyleDefs(76)  =   "Named:id=34:Heading"
      _StyleDefs(77)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   ":id=34,.wraptext=-1"
      _StyleDefs(79)  =   "Named:id=35:Footing"
      _StyleDefs(80)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(81)  =   "Named:id=36:Selected"
      _StyleDefs(82)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(83)  =   "Named:id=37:Caption"
      _StyleDefs(84)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(85)  =   "Named:id=38:HighlightRow"
      _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(87)  =   "Named:id=39:EvenRow"
      _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(89)  =   "Named:id=40:OddRow"
      _StyleDefs(90)  =   ":id=40,.parent=33"
      _StyleDefs(91)  =   "Named:id=41:RecordSelector"
      _StyleDefs(92)  =   ":id=41,.parent=34"
      _StyleDefs(93)  =   "Named:id=42:FilterBar"
      _StyleDefs(94)  =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   825
      Left            =   45
      TabIndex        =   12
      Top             =   3690
      Width           =   8850
      _Version        =   65536
      _ExtentX        =   15610
      _ExtentY        =   1455
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
      Begin VB.CommandButton CmdCopiar 
         BackColor       =   &H00000000&
         Caption         =   "&Copiar"
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
         Left            =   5175
         TabIndex        =   3
         Top             =   60
         Width           =   1020
      End
      Begin VB.CommandButton CmdReserva 
         Caption         =   "&Reserva"
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
         Left            =   6420
         TabIndex        =   4
         Top             =   60
         Width           =   1020
      End
      Begin VB.CommandButton CmdContratos 
         Caption         =   "Co&ntratos"
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
         Left            =   1440
         TabIndex        =   5
         Top             =   480
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
         Left            =   1440
         TabIndex        =   0
         Top             =   60
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
         Left            =   6420
         TabIndex        =   9
         Top             =   480
         Width           =   1020
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
         Left            =   5175
         TabIndex        =   8
         Top             =   480
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
         Left            =   3930
         TabIndex        =   7
         Top             =   480
         Width           =   1020
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
         Left            =   2685
         TabIndex        =   6
         Top             =   480
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
         Left            =   3930
         TabIndex        =   2
         Top             =   60
         Width           =   1020
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
         Left            =   2685
         TabIndex        =   1
         Top             =   60
         Width           =   1020
      End
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaImoveis.frx":059F
      Height          =   315
      Left            =   2730
      TabIndex        =   11
      Top             =   3330
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Imovel"
      BoundColumn     =   "Imovel"
      Text            =   ""
   End
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Imóvel:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   -30
      TabIndex        =   13
      Top             =   3360
      Width           =   2670
   End
End
Attribute VB_Name = "TelaManImoveis"
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
    CmdContratos.Enabled = True
    CmdReserva.Enabled = True
    CmdSair.Enabled = True
    CmdImprimir.Enabled = True
    CmdOrdem.Enabled = True
    CmdCopiar.Enabled = True

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
    CmdContratos.Enabled = False
    CmdReserva.Enabled = False
    CmdRemoverFiltro.Enabled = False
    CmdSair.Enabled = False
    CmdImprimir.Enabled = False
    CmdOrdem.Enabled = False
    CmdCopiar.Enabled = False
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

Private Sub CmdContratos_Click()
   
   If Not FunVerifica_Permissao("TelaManContratos", "1") Then
        Exit Sub
   End If
   
   'Constrói a chave concatenando o código do empreendimento com o código do imóvel
   Chave = Adodc1.Recordset.Fields("Imovel")
   subDesabilitaBotoes
   TelaManContratos.Show
   subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
   cboLocalizar.BoundText = Chave
   Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
   subHabilitaBotoes
   
End Sub

Private Sub CmdCopiar_Click()
  
  If Not FunVerifica_Permissao(Me.Name, "2") Then
        Exit Sub
  End If

  Dim XLI_QTDINSERIDO As Integer
  Dim XLI_CONT As Integer
  
  Chave = Adodc1.Recordset.Fields("empd_cd_Empreendimento") + Adodc1.Recordset.Fields("imov_cd_Imovel")
  XInserir = "C"
  subDesabilitaBotoes
  FrmManImoveis.Show 1
  
  XLI_QTDINSERIDO = TDBGrid1.ApproxCount
  subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual

  'Como pode ser inserido mais de um imóvel na chamada de inserir então
  'é preciso saber quantos registros foram inseridos e assim atualizar o panel
  XLI_QTDINSERIDO = TDBGrid1.ApproxCount - XLI_QTDINSERIDO
  If XLI_QTDINSERIDO > 0 Then
    For XLI_CONT = 1 To XLI_QTDINSERIDO
      FRM_Quantid = funOperacaoQuantid("+")
    Next
  End If
  MdiPrincipal.BarraStatus.Panels(3).Text = Str(TDBGrid1.ApproxCount) + "/" + Str(Quantid)
        
  
  cboLocalizar.BoundText = Chave
  If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
  subHabilitaBotoes
  
End Sub

Private Sub CmdImprimir_Click()
    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Imóveis")
End Sub


Private Sub CmdRemoverFiltro_Click()
    Dim XLI_FIL As Integer
    
    subDesabilitaBotoes

    Set Formulario = TelaManImoveis
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
       CmdContratos.Enabled = False
       CmdReserva.Enabled = False
       CmdOrdem.Enabled = False
       CmdImprimir.Enabled = False
       CmdCopiar.Enabled = False
    End If
    
    TDBGrid1.SetFocus

End Sub

Private Sub CmdRemoverFiltro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Mostra os filtros ativos da tabela"
End Sub

Private Sub CmdAlterar_Click()
    
    If Not FunVerifica_Permissao(Me.Name, "3") Then
        Exit Sub
    End If
    
    'Constrói a chave concatenando o código do empreendimento com o código do imóvel
    Chave = Adodc1.Recordset.Fields("empd_cd_Empreendimento") + Adodc1.Recordset.Fields("imov_cd_Imovel")
    
    XInserir = "A"
    xChave = Adodc1.Recordset.bookmark
    subDesabilitaBotoes
    FrmManImoveis.Show 1
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    cboLocalizar.BoundText = Chave
    If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    subHabilitaBotoes
    
End Sub

Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Imprime os dados da tabela"
End Sub

Private Sub CmdReserva_Click()
   
   If Not FunVerifica_Permissao("FrmManReserva", "1") Then
        Exit Sub
   End If
   
   Dim XLO_Contrato As New ADODB.Recordset
   
   'Primeiro verifica se o status do imóvel é disponível
   'se for disponível verifica se já existe reserva
   If TDBGrid1.Columns("Status") = "Disponível" Then
     SubQOpenRecordset XLO_Contrato, "SELECT Contrato FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "' AND imov_cd_Imovel= '" & Adodc1.Recordset.Fields("imov_cd_Imovel") & "' AND cont_tx_Status='I'", Estatico
     If Not XLO_Contrato.EOF Then
       Chave = Adodc1.Recordset.Fields("Imovel")
     Else
       Chave = Adodc1.Recordset.Fields("Imovel")
     End If
     XInserir = "I"
     XLO_Contrato.Close
     Set XLO_Contrato = Nothing
   
   ElseIf TDBGrid1.Columns("Status") = "Reservado" Then
     SubQOpenRecordset XLO_Contrato, "SELECT Contrato FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "' AND imov_cd_Imovel= '" & Adodc1.Recordset.Fields("imov_cd_Imovel") & "' AND cont_tx_Status='R'", Estatico
     If Not XLO_Contrato.EOF Then
       Chave = XLO_Contrato!contrato
     End If
     XInserir = "A"
     XLO_Contrato.Close
     Set XLO_Contrato = Nothing
   Else
     MsgBox "Esse imóvel já está vendido.", vbCritical, "ATENÇÃO !"
     Exit Sub
   End If
  
   subDesabilitaBotoes
   FrmManReserva.Show vbModal
   subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
   cboLocalizar.BoundText = Chave
   Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
   subHabilitaBotoes
   
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
    Set Formulario = TelaManImoveis
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
    
    If Not FunVerifica_Permissao(Me.Name, "4") Then
        Exit Sub
    End If
    
    Dim XLI_OPCAO As Integer
    Dim XLI_BOOKMARK As Integer
    
    If TDBGrid1.Columns("Status") = "Reservado" Then
        MsgBox "Esse imóvel não pode ser excluído pois já possui contratos relacionados!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
    On Error GoTo RotuloErro
    
    subDesabilitaBotoes
    
    XLI_OPCAO = MsgBox("Confirma Remoção do Registro ?", 273, "CUIDADO")
    
    If XLI_OPCAO = 1 Then  'Confirma exclusão
    
      Conexao.Execute "DELETE FROM Imoveis WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel ='" & Adodc1.Recordset.Fields("imov_cd_imovel") & "' AND empd_cd_Empreendimento= '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "'"
      XLI_BOOKMARK = TDBGrid1.bookmark
      FRM_Quantid = funOperacaoQuantid("-")
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
         CmdContratos.Enabled = False
         CmdReserva.Enabled = False
         CmdImprimir.Enabled = False
         CmdOrdem.Enabled = False
         CmdCopiar.Enabled = False
      End If
       
    Else
       subHabilitaBotoes
    End If
       
    cboLocalizar.Text = ""
    TDBGrid1.SetFocus
    TDBGrid1.Col = 0
    
    Exit Sub
    
RotuloErro:
    
    If Err.Number = -2147467259 Or Err.Number = -2147217873 Then
      MsgBox "Esse imóvel não pode ser excluído pois já possui contratos relacionados!", vbCritical, "ATENÇÃO"
    End If
    subHabilitaBotoes
    
End Sub

Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If CmdInserir.Caption = "&Gravar" Then
       MdiPrincipal.BarraStatus.Panels(1).Text = "Cancela as novas informações"
    Else
       MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
    End If
End Sub

Private Sub CmdInserir_Click()
    
    If Not FunVerifica_Permissao(Me.Name, "2") Then
        Exit Sub
    End If
    
    Dim XLI_QTDINSERIDO As Integer
    Dim XLI_CONT As Integer
    
    Chave = -1
    XInserir = "I"
    cboLocalizar.Text = ""
    subDesabilitaBotoes
    FrmManImoveis.Show 1
    If Chave <> -1 Then
        
        XLI_QTDINSERIDO = TDBGrid1.ApproxCount
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        
        'Como pode ser inserido mais de um imóvel na chamada de inserir então
        'é preciso saber quantos registros foram inseridos e assim atualizar o panel
        XLI_QTDINSERIDO = TDBGrid1.ApproxCount - XLI_QTDINSERIDO
        If XLI_QTDINSERIDO > 0 Then
          For XLI_CONT = 1 To XLI_QTDINSERIDO
            FRM_Quantid = funOperacaoQuantid("+")
          Next
        End If
        MdiPrincipal.BarraStatus.Panels(3).Text = Str(TDBGrid1.ApproxCount) + "/" + Str(Quantid)
        
        cboLocalizar.BoundText = Chave
        If Not IsNull(cboLocalizar.SelectedItem) Then
          Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
        End If
    End If
    subHabilitaBotoes
    
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
   If TDBGrid1.ColContaining(x) <> TDBGrid1.Columns("Data da Venda").ColIndex Then
      If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
          Set Formulario = TelaManImoveis
          subTelaValoresGlobais "P"
          FunExecutaFiltro Adodc1, Formulario, x
          subTelaValoresGlobais "G"
      End If
   Else
     MsgBox "Essa coluna não pode ser filtrada", vbInformation, "ATENÇÃO"
   End If
    
    subHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) = True Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False 'Desabilita o botão Excluir
       CmdCopiar.Enabled = False
       CmdReserva.Enabled = False
       CmdContratos.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If
End Sub


Private Sub Form_Activate()
    
    Set Formulario = TelaManImoveis
    If Not PrimeiraVez Then
       subTelaValoresGlobais "P"
    End If
    
    filtra_especial = True
 
    If Not PrimeiraVez Then
      subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
      cboLocalizar.BoundText = Chave
      If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
    
    PrimeiraVez = False

End Sub

Private Sub Form_Load()
    
    Set Formulario = TelaManImoveis
    
    Call subAjustaTelaLoc
    
    subManutencaoJanelasAtivas "I", "TelaManImoveis"
    
    If XOrigem = "TelaManEmpreendimentos" Then
      subCarregaDadosNV Adodc1, "", "Imovel", _
                        "", "SELECT * FROM ConsCAPImoveis WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & Chave & "'"
    Else
      subCarregaDadosNV Adodc1, "", "Imovel", _
                       "", "SELECT * FROM ConsCAPImoveis WHERE empr_cd_empresa = " & PCodEmpresa
    End If
    subHabilitaBotoes
    
    filtra_especial = True
    
    subTelaValoresGlobais "G"
    PrimeiraVez = True
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdContratos.Enabled = False
       CmdReserva.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
       CmdCopiar.Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaManImoveis"
End Sub

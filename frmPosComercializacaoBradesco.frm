VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRelPosComercializacaoBradesco 
   Caption         =   "Posição de Comercialização Bradesco"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   8850
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   7680
      TabIndex        =   13
      Top             =   4320
      Width           =   1035
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Height          =   345
      Left            =   6480
      TabIndex        =   12
      Top             =   4320
      Width           =   1035
   End
   Begin VB.Frame FraPrincipal 
      Height          =   4245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.Frame Frame2 
         Caption         =   "Correção"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   120
         TabIndex        =   3
         Top             =   2680
         Width           =   4335
         Begin VB.OptionButton optValorHistorico 
            Caption         =   "Valor Histórico"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1340
         End
         Begin VB.OptionButton optIndexador1 
            Caption         =   "Indexador 1"
            Height          =   255
            Left            =   1680
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton optIndexador2 
            Caption         =   "Indexador 2"
            Height          =   255
            Left            =   3000
            TabIndex        =   6
            Top             =   360
            Width           =   1140
         End
      End
      Begin MSComCtl2.DTPicker DtBase 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   3680
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         _Version        =   393216
         Format          =   101974017
         CurrentDate     =   39700
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Height          =   3705
         Left            =   4560
         TabIndex        =   10
         Top             =   360
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   6535
         _LayoutType     =   4
         _RowHeight      =   24
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Planos"
         Columns(0).DataField=   "tipl_tx_Descricao"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   68
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Chaves"
         Columns(1).DataField=   "PlanosChaves"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "CodigoPlano"
         Columns(2).DataField=   "tipl_cd_TipoPlano"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   4
         Splits(0).SizeMode=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4868"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4789"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=20"
         Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=20"
         Splits(0)._ColumnProps(20)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(1)._UserFlags=   0
         Splits(1).ExtendRightColumn=   -1  'True
         Splits(1).MarqueeStyle=   4
         Splits(1).SizeMode=   2
         Splits(1).Size  =   6
         Splits(1).Size.vt=   2
         Splits(1).RecordSelectors=   0   'False
         Splits(1).RecordSelectorWidth=   503
         Splits(1).ScrollBars=   2
         Splits(1).DividerColor=   12632256
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=3"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=6535"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=6456"
         Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=20"
         Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(7)=   "Column(0).WrapText=1"
         Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(9)=   "Column(1).Width=1244"
         Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=1164"
         Splits(1)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=17"
         Splits(1)._ColumnProps(14)=   "Column(1).WrapText=1"
         Splits(1)._ColumnProps(15)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(16)=   "Column(2).Width=2725"
         Splits(1)._ColumnProps(17)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(18)=   "Column(2)._WidthInPix=2646"
         Splits(1)._ColumnProps(19)=   "Column(2)._EditAlways=0"
         Splits(1)._ColumnProps(20)=   "Column(2)._ColStyle=20"
         Splits(1)._ColumnProps(21)=   "Column(2).Visible=0"
         Splits(1)._ColumnProps(22)=   "Column(2).WrapText=1"
         Splits(1)._ColumnProps(23)=   "Column(2).Order=3"
         Splits.Count    =   2
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
         DataMode        =   4
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
         _StyleDefs(24)  =   "Splits(0).Style:id=47,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=88,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=48,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=49,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=52,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=51,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=53,.parent=8,.bgcolor=&H800000&"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=54,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=87,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=89,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=90,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=94,.parent=47"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=91,.parent=48"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=92,.parent=49,.bold=0,.fontsize=825"
         _StyleDefs(39)  =   ":id=92,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(40)  =   ":id=92,.fontname=MS Sans Serif"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=93,.parent=51"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=16,.parent=47"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=48"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=49"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=51"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=24,.parent=47"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=21,.parent=48"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=22,.parent=49"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=23,.parent=51"
         _StyleDefs(50)  =   "Splits(1).Style:id=103,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(51)  =   "Splits(1).CaptionStyle:id=112,.parent=4"
         _StyleDefs(52)  =   "Splits(1).HeadingStyle:id=104,.parent=2"
         _StyleDefs(53)  =   "Splits(1).FooterStyle:id=105,.parent=3"
         _StyleDefs(54)  =   "Splits(1).InactiveStyle:id=106,.parent=5"
         _StyleDefs(55)  =   "Splits(1).SelectedStyle:id=108,.parent=6"
         _StyleDefs(56)  =   "Splits(1).EditorStyle:id=107,.parent=7"
         _StyleDefs(57)  =   "Splits(1).HighlightRowStyle:id=109,.parent=8,.bgcolor=&H800000&"
         _StyleDefs(58)  =   "Splits(1).EvenRowStyle:id=110,.parent=9"
         _StyleDefs(59)  =   "Splits(1).OddRowStyle:id=111,.parent=10"
         _StyleDefs(60)  =   "Splits(1).RecordSelectorStyle:id=113,.parent=11"
         _StyleDefs(61)  =   "Splits(1).FilterBarStyle:id=114,.parent=12"
         _StyleDefs(62)  =   "Splits(1).Columns(0).Style:id=118,.parent=103"
         _StyleDefs(63)  =   "Splits(1).Columns(0).HeadingStyle:id=115,.parent=104"
         _StyleDefs(64)  =   "Splits(1).Columns(0).FooterStyle:id=116,.parent=105"
         _StyleDefs(65)  =   "Splits(1).Columns(0).EditorStyle:id=117,.parent=107"
         _StyleDefs(66)  =   "Splits(1).Columns(1).Style:id=20,.parent=103,.alignment=2"
         _StyleDefs(67)  =   "Splits(1).Columns(1).HeadingStyle:id=17,.parent=104"
         _StyleDefs(68)  =   "Splits(1).Columns(1).FooterStyle:id=18,.parent=105"
         _StyleDefs(69)  =   "Splits(1).Columns(1).EditorStyle:id=19,.parent=107"
         _StyleDefs(70)  =   "Splits(1).Columns(2).Style:id=28,.parent=103"
         _StyleDefs(71)  =   "Splits(1).Columns(2).HeadingStyle:id=25,.parent=104"
         _StyleDefs(72)  =   "Splits(1).Columns(2).FooterStyle:id=26,.parent=105"
         _StyleDefs(73)  =   "Splits(1).Columns(2).EditorStyle:id=27,.parent=107"
         _StyleDefs(74)  =   "Named:id=33:Normal"
         _StyleDefs(75)  =   ":id=33,.parent=0"
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
      Begin MSComctlLib.TreeView TvwEmpreendimentos 
         Height          =   2205
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3889
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ComctlLib.ProgressBar pgbImpressao 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   4200
         Visible         =   0   'False
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   873
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label LblEmpreendimento 
         Alignment       =   1  'Right Justify
         Caption         =   "Empreendimento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Data Base:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   3780
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Classificação dos Planos:"
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
         Left            =   4560
         TabIndex        =   9
         Top             =   120
         Width           =   2415
      End
   End
   Begin MSComDlg.CommonDialog dlgSalvar 
      Left            =   2400
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatEmpreendimento2 
      Height          =   330
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
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
      Caption         =   "DatEmpreendimento"
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
   Begin MSAdodcLib.Adodc DatEmpreendimento1 
      Height          =   330
      Left            =   120
      Top             =   4680
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
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
      Caption         =   "DatEmpreendimento"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   120
      Top             =   4920
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
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
      Caption         =   "DatEmpreendimento"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmRelPosComercializacaoBradesco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VFV_VETORPLANOS As New XArray           'Vetor que armazena a classificação dos planos
Public XFO_ExlObj As Object 'Excel.Application      ' Create excel object
Public XFI_QtdEmpreendimento As Integer     'Quantidade de empreendimentos
Public XFI_UltimaLinha  As Integer         'Guarda a última linha preenchida na planilha, para o caso de vários empreendimentos selecionados
Public XFI_Bloco As Integer

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelPosComercializacaoBradesco
End Sub

Private Sub Form_Load()
    Sql = "SELECT tipl_tx_Descricao, '' AS PlanosChaves, tipl_cd_TipoPlano FROM TiposPlanos WHERE tipl_tx_Sigla IS NULL OR tipl_tx_Sigla <> 'R' ORDER BY tipl_tx_Descricao"
    subCarregaVetor Sql, Array("tipl_tx_Descricao", "PlanosChaves", "tipl_cd_TipoPlano"), VFV_VETORPLANOS, TDBGrid1
    
    subConectarControleDadosNV DatEmpreendimento1, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_cd_empreendimento", Estatico
    'subConectarControleDadosNV DatEmpreendimento2, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_cd_empreendimento", Estatico
    DtBase.Value = Date
    
    XFI_QtdEmpreendimento = 0
    PreencheTreeViewEmpreendimentos
End Sub

'Carrega o treeview com os empreendimentos da empresa ativa
Private Sub PreencheTreeViewEmpreendimentos()
    Dim nodeEmpd As MSComctlLib.Node
            
    TvwEmpreendimentos.Nodes.Clear
        
    'Cria os nós principais
    DatEmpreendimento1.Recordset.MoveFirst
    
    While Not DatEmpreendimento1.Recordset.EOF
        Set nodeEmpd = TvwEmpreendimentos.Nodes.Add(, , "N" & DatEmpreendimento1.Recordset.Fields("empd_cd_Empreendimento"), _
            DatEmpreendimento1.Recordset.Fields("empd_cd_Empreendimento") & " - " & DatEmpreendimento1.Recordset.Fields("empd_tx_Nome"))
        
        nodeEmpd.Tag = "Folha"
        
        'Cria as folhas do nó criado
'        CarregaRegistros nodeEmpd
        
        'Atualiza a quantidade de empreendimentos
        XFI_QtdEmpreendimento = XFI_QtdEmpreendimento + 1
        
        DatEmpreendimento1.Recordset.MoveNext
    Wend
End Sub

'Preenche a árvore do nó selecionado
Private Sub CarregaRegistros(ByVal nodeNo As MSComctlLib.Node)
    Dim nodeFolha As MSComctlLib.Node
    
    'Se os registros do nó já foram carregados não faz nada
    If nodeNo.Children > 0 Then Exit Sub
    
    DatEmpreendimento2.Recordset.MoveFirst
    
    While Not DatEmpreendimento2.Recordset.EOF
        If (nodeNo.Text <> DatEmpreendimento2.Recordset.Fields("empd_cd_Empreendimento") & " - " & DatEmpreendimento2.Recordset.Fields("empd_tx_Nome")) Then
            Set nodeFolha = TvwEmpreendimentos.Nodes.Add(nodeNo, tvwChild, , _
                            DatEmpreendimento2.Recordset.Fields("empd_cd_Empreendimento") & " - " & DatEmpreendimento2.Recordset.Fields("empd_tx_Nome"))
            
            nodeFolha.Tag = "Folha"
        End If
        DatEmpreendimento2.Recordset.MoveNext
    Wend
End Sub

'Inicializa a barra de progresso
Private Sub InicializarProgresso(valorMinimo As Integer, valorMaximo As Integer)
    'Reajusta a altura do formulário
    FrmRelPosComercializacaoBradesco.Height = 5895
    CmdImprimir.Top = 4920
    CmdCancelar.Top = 4920
    FraPrincipal.Height = 4845
        
    pgbImpressao.Min = valorMinimo
    pgbImpressao.Max = valorMaximo
    pgbImpressao.Value = 0
    pgbImpressao.Visible = True
End Sub

'Finaliza a barra de progresso
Private Sub FinalizarProgresso()
    pgbImpressao.Value = 0
    pgbImpressao.Visible = False
    
    'Reajusta a altura do formulário
    FraPrincipal.Height = 4245
    CmdImprimir.Top = 4320
    CmdCancelar.Top = 4320
    FrmRelPosComercializacaoBradesco.Height = 5250
End Sub

'Avança unidades na barra de progresso
Private Sub AvancarProgresso(Valor As Integer)
    If (pgbImpressao.Value + Valor <= pgbImpressao.Max) Then
        pgbImpressao.Value = pgbImpressao.Value + Valor
    End If
End Sub

'Imprime planilha
Private Sub CmdImprimir_Click()
  
    Dim XLO_COMANDO             As New ADODB.Command
    Dim rs                      As New ADODB.Recordset
    Dim XLI_INDEXADOR           As Integer
    Dim XLT_CHAVES              As String   'Armazena o código dos planos classificados como Chaves
    Dim XLI_QTDEMPREENDIMENTOS  As Integer  'Auxiliar
    Dim j                       As Integer  'Auxiliar
    Dim sheet                   As Integer  'Quantidade de planilhas no workbook
    Dim XLT_EMPREENDIMENTOS     As String
    Dim XLO_EMPREENDIMENTO()    As String
        
    
    On Error GoTo TrataErro
    
    'Cria string com os planos considerados "Chaves"
    TDBGrid1.MoveFirst
    While Not TDBGrid1.EOF
        If TDBGrid1.Columns("PlanosChaves").Value = "-1" Then
            XLT_CHAVES = XLT_CHAVES + ", " + TDBGrid1.Columns("CodigoPlano")
        End If
        TDBGrid1.MoveNext
    Wend
    
    'Remove a vírgula do início
    XLT_CHAVES = Mid(XLT_CHAVES, 2)
    
    
    If XLT_CHAVES = "" Then
        MsgBox "É necessário informar os planos que são chaves.", vbInformation, "Atenção"
        FrmRelRelacaoClientes.MousePointer = vbDefault
        Exit Sub
    End If
                                  
    
    'Seta o filtro das extensões de arquivos que será usada
    dlgSalvar.Filter = "Documento do Execl(*.xls)|*.xls"
    
    'Indica qual sera o filtro padrão
    dlgSalvar.FilterIndex = 1
                
    dlgSalvar.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
            
    'Define o diretório onde o arquivo será salvo e exclui o que existe
    dlgSalvar.FileName = "C:\COMPRADORES.xls"
    DeleteFile (dlgSalvar.FileName)
        
        
    'Define o indexador para o cálculo da correção dos títulos
    If optValorHistorico.Value = True Then
      XLI_INDEXADOR = 0
    ElseIf optIndexador1.Value = True Then
      XLI_INDEXADOR = 1
    ElseIf optIndexador2.Value = True Then
      XLI_INDEXADOR = 2
    End If
    
    'Altera o cursor do mouse
    FrmRelPosComercializacaoSantander.MousePointer = vbHourglass
    
    'Inicializa o objeto excel
    Set XFO_ExlObj = CreateObject("excel.application")
    
    'Adiciona um WorkBook
    XFO_ExlObj.Workbooks.Add
    XFO_ExlObj.Sheets.Add

    'Impede que o usuário veja a planilha enquanto ela está sendo preenchida
    XFO_ExlObj.Visible = False
    
    'Verifica quantos empreendimentos foram selecionados
    i = 1
    XLI_QTDEMPREENDIMENTOS = 0
    While i <= TvwEmpreendimentos.Nodes.Count
        If (TvwEmpreendimentos.Nodes.Item(i).Tag = "Folha") And (TvwEmpreendimentos.Nodes.Item(i).Checked = True) Then
            XLI_QTDEMPREENDIMENTOS = XLI_QTDEMPREENDIMENTOS + 1
        End If
        i = i + 1
    Wend
    
    'Inicializa a barra de progresso
    InicializarProgresso 0, XLI_QTDEMPREENDIMENTOS
                    
                
    i = 1
    sheet = 1
    XLT_EMPREENDIMENTOS = ""
    XFI_UltimaLinha = 1
    XFI_Bloco = 1
    
    While i <= TvwEmpreendimentos.Nodes.Count
        If (TvwEmpreendimentos.Nodes.Item(i).Tag = "Folha") And (TvwEmpreendimentos.Nodes.Item(i).Checked = True) Then
            
            
            XLO_EMPREENDIMENTO = Split(TvwEmpreendimentos.Nodes.Item(i).Text, "-")
            XLT_EMPREENDIMENTOS = Trim(XLO_EMPREENDIMENTO(0))
                          
            For j = 1 To i - 1
                If (TvwEmpreendimentos.Nodes.Item(j).Checked = True) Then
                    XLO_EMPREENDIMENTO = Split(TvwEmpreendimentos.Nodes.Item(j).Text, "-")
                    XLT_EMPREENDIMENTOS = XLT_EMPREENDIMENTOS + "," & Trim(XLO_EMPREENDIMENTO(0))
                End If
            Next
            
            
            AvancarProgresso 1
            
            '------------------
            'Imprime relatório
            '------------------
            If (XLT_EMPREENDIMENTOS <> "") Then
                With XLO_COMANDO
                    .CommandTimeout = 200
                    .CommandText = "spCAPI_sel_posicaoComercializacaoBRADESCO"
                    .CommandType = adCmdStoredProc
                    Set XLO_COMANDO.ActiveConnection = Conexao
                    'Cria automaticamente todos os parâmetros dentro do objeto Comando
                    .Parameters.Refresh
                End With
                
                XLO_COMANDO.Parameters(1).Value = PCodEmpresa           'Empresa
                XLO_COMANDO.Parameters(2).Value = DtBase.Value          'Data base
                XLO_COMANDO.Parameters(3).Type = adBSTR
                XLO_COMANDO.Parameters(3).Value = XLT_EMPREENDIMENTOS   'Empreendimento
                XLO_COMANDO.Parameters(4).Value = XLI_INDEXADOR         'Indexador
                XLO_COMANDO.Parameters(5).Value = XLT_CHAVES            'Planos chaves
                
                Set rs = XLO_COMANDO.Execute
                     
                'Verifica se a consulta retornou dados, caso contrário finaliza a rotina
                If rs.EOF Then
                    MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados """
                    FrmRelRelacaoClientes.MousePointer = vbDefault
                    Exit Sub
                End If
        
                AvancarProgresso 3
                
                ExportaExcel rs, sheet
                
                'Incrementa o número do bloco (que agrupa os imóveis por empreendimento)
                XFI_Bloco = XFI_Bloco + 1
            End If
        End If
        
        XLT_EMPREENDIMENTOS = ""
        i = i + 1
    Wend

    
    'Salva a planilha no caminho selecionado pelo usuário
    XFO_ExlObj.Workbooks(1).SaveAs dlgSalvar.FileName
    XFO_ExlObj.Workbooks(1).Close
    XFO_ExlObj.Application.Quit
    Set XFO_ExlObj = Nothing
    
    AvancarProgresso 1
                
    'Limpa o caminho do arquivo
    dlgSalvar.FileName = ""
    FrmRelPosComercializacaoSantander.MousePointer = vbDefault
    MsgBox "Relatório Exportado com sucesso!", vbInformation, "Informação"
         
    FinalizarProgresso
      
Exit Sub

TrataErro:
    If Err.Number = cdlCancel Then
        Exit Sub
    End If
    FrmRelPosComercializacaoBradesco.MousePointer = vbDefault
    MsgBox "Ocorreu um erro na exportação do relatório. Contate o suporte. Detalhe do Erro: " & Err.Description, vbCritical, "Erro"
    DeleteFile (dlgSalvar.FileName)

End Sub

Private Sub ExportaExcel(dados As ADODB.Recordset, sheet As Integer)
    
    Dim i As Integer
        
    'Seleciona a próxima planilha vazia para ser preenchida
    XFO_ExlObj.Sheets(sheet).Select
         
    'Preenchendo as células
    With XFO_ExlObj.Sheets(sheet)
        'Seta o nome da aba da planilha
        .Name = "Import"
        XFO_ExlObj.ActiveSheet.Range("A1").Select
    End With
                
    'Preenche os títulos das colunas
    With XFO_ExlObj.Sheets(sheet)
        .Cells(1, 1) = "BLOCO"
        .Cells(1, 2) = "UNIDADE"
        .Cells(1, 3) = "DEPOSITO"
        .Cells(1, 4) = "G1"
        .Cells(1, 5) = "G2"
        .Cells(1, 6) = "G3"
        .Cells(1, 7) = "G4"
        .Cells(1, 8) = "COMPRADOR"
        .Cells(1, 9) = "DDD"
        .Cells(1, 10) = "NRO_TEL"
        .Cells(1, 11) = "DDD"
        .Cells(1, 12) = "NRO_TEL"
        .Cells(1, 13) = "NRO_CPF-CNPJ"
        .Cells(1, 14) = "NRO_FILIAL"
        .Cells(1, 15) = "NRO_DIG"
        .Cells(1, 16) = "LOG_TIPO"
        .Cells(1, 17) = "LOG_ENDR"
        .Cells(1, 18) = "LOG_NRO"
        .Cells(1, 19) = "LOG_COMPL"
        .Cells(1, 20) = "LOG_BAIRRO"
        .Cells(1, 21) = "LOG_CIDADE"
        .Cells(1, 22) = "LOG_UF"
        .Cells(1, 23) = "CEP_NRO"
        .Cells(1, 24) = "CEP_COMPL"
        .Cells(1, 25) = "DT_VENDA"
        .Cells(1, 26) = "VLR_VENDA"
        .Cells(1, 27) = "VLR_PRE-RECEBD"
        .Cells(1, 28) = "VLR_PRE-RECEBR"
        .Cells(1, 29) = "VLR_POS-CH"
        .Cells(1, 30) = "VLR_SDO-DV"
        .Cells(1, 31) = "DT_REF_SDODV"
    End With
    
    AvancarProgresso 1
    
    i = XFI_UltimaLinha + 1
        
    'Preenche a planilha com os dados
    While Not dados.EOF
        
        With XFO_ExlObj.Sheets(sheet)
            .Cells(i, 1) = "0" + Trim(Str(XFI_Bloco))
            If (dados!Unidade <> "") Then .Cells(i, 2) = dados!Unidade
            If (dados!Deposito <> "") Then .Cells(i, 3) = dados!Deposito
            If (dados!G1 <> "") Then .Cells(i, 4) = dados!G1
            If (dados!G2 <> "") Then .Cells(i, 5) = dados!G2
            If (dados!G3 <> "") Then .Cells(i, 6) = dados!G3
            If (dados!G4 <> "") Then .Cells(i, 7) = dados!G4
            If (dados!Comprador <> "") Then .Cells(i, 8) = dados!Comprador
            If (dados!DDD1 <> "") Then .Cells(i, 9) = dados!DDD1
            If (dados!Telefone1 <> "") Then .Cells(i, 10) = dados!Telefone1
            If (dados!DDD2 <> "") Then .Cells(i, 11) = dados!DDD2
            If (dados!Telefone2 <> "") Then .Cells(i, 12) = dados!Telefone2
            If (dados!NumeroCGC <> "") Then .Cells(i, 13) = dados!NumeroCGC
            If (dados!FilialCGC <> "") Then .Cells(i, 14) = dados!FilialCGC
            If (dados!DigitoCGC <> "") Then .Cells(i, 15) = dados!DigitoCGC
            If (dados!TipoLogradouro <> "") Then .Cells(i, 16) = dados!TipoLogradouro
            If (dados!Endereco <> "") Then .Cells(i, 17) = dados!Endereco
            If (dados!NumeroEndereco <> "") Then .Cells(i, 18) = dados!NumeroEndereco
            If (dados!ComplementoEndereco <> "") Then .Cells(i, 19) = dados!ComplementoEndereco
            If (dados!Bairro <> "") Then .Cells(i, 20) = dados!Bairro
            If (dados!Cidade <> "") Then .Cells(i, 21) = dados!Cidade
            If (dados!Estado <> "") Then .Cells(i, 22) = dados!Estado
            If (dados!NumeroCEP <> "") Then .Cells(i, 23) = dados!NumeroCEP
            If (dados!ComplementoCEP <> "") Then .Cells(i, 24) = dados!ComplementoCEP
            If (dados!DataVenda <> "") Then .Cells(i, 25) = dados!DataVenda
            If (dados!ValorVenda <> "") Then .Cells(i, 26) = dados!ValorVenda
            If (dados!ValorPreChavesPago <> "") Then .Cells(i, 27) = dados!ValorPreChavesPago
            If (dados!ValorPreChavesReceber <> "") Then .Cells(i, 28) = dados!ValorPreChavesReceber
            If (dados!ValorPosChavesReceber <> "") Then .Cells(i, 29) = dados!ValorPosChavesReceber
            If (dados!SaldoDevedorAtual <> "") Then .Cells(i, 30) = dados!SaldoDevedorAtual
            If (dados!DataSaldoDevedor <> "") Then .Cells(i, 31) = dados!DataSaldoDevedor
        End With
               
        i = i + 1
        dados.MoveNext
    Wend
    
    XFI_UltimaLinha = i - 1
    
    AvancarProgresso 3
    
End Sub

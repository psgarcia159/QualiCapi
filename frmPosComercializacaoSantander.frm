VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRelPosComercializacaoSantander 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Posição de Comercialização Santander"
   ClientHeight    =   4545
   ClientLeft      =   3330
   ClientTop       =   3615
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgSalvar 
      Left            =   2760
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraPrincipal 
      Height          =   4125
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11055
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
         Height          =   660
         Left            =   4040
         TabIndex        =   5
         Top             =   3340
         Width           =   5295
         Begin VB.OptionButton optIndexador2 
            Caption         =   "Indexador 2"
            Height          =   255
            Left            =   3720
            TabIndex        =   8
            Top             =   240
            Width           =   1140
         End
         Begin VB.OptionButton optIndexador1 
            Caption         =   "Indexador 1"
            Height          =   255
            Left            =   2160
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton optValorHistorico 
            Caption         =   "Valor Histórico"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   240
            Width           =   1340
         End
      End
      Begin MSComCtl2.DTPicker DtBase 
         Height          =   375
         Left            =   9420
         TabIndex        =   10
         Top             =   3615
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         _Version        =   393216
         Format          =   21102593
         CurrentDate     =   39700
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Height          =   2985
         Left            =   4040
         TabIndex        =   4
         Top             =   360
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   5265
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
         Columns(1).Caption=   "Mensal"
         Columns(1).DataField=   "PlanosMensais"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   68
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Semestral"
         Columns(2).DataField=   "PlanosSemestrais"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   68
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Anual"
         Columns(3).DataField=   "PlanosAnuais"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   68
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Outros"
         Columns(4).DataField=   "PlanosOutros"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   68
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Chaves"
         Columns(5).DataField=   "PlanosChaves"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   68
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Repasse"
         Columns(6).DataField=   "PlanosRepasse"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "CodigoPlano"
         Columns(7).DataField=   "tipl_cd_TipoPlano"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3810"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3731"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1482"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1402"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=17"
         Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(16)=   "Column(2).Width=1535"
         Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=1455"
         Splits(0)._ColumnProps(19)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=17"
         Splits(0)._ColumnProps(21)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(22)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(24)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=17"
         Splits(0)._ColumnProps(29)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(30)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(31)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(32)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(33)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(35)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(4)._ColStyle=20"
         Splits(0)._ColumnProps(37)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(38)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(39)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(40)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(42)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(43)=   "Column(5)._ColStyle=20"
         Splits(0)._ColumnProps(44)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(45)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(46)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(47)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(49)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(50)=   "Column(6)._ColStyle=20"
         Splits(0)._ColumnProps(51)=   "Column(6).WrapText=1"
         Splits(0)._ColumnProps(52)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(53)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(54)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(56)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(57)=   "Column(7)._ColStyle=20"
         Splits(0)._ColumnProps(58)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(59)=   "Column(7).WrapText=1"
         Splits(0)._ColumnProps(60)=   "Column(7).Order=8"
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
         Splits(1)._ColumnProps(0)=   "Columns.Count=8"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=6535"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=6456"
         Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=20"
         Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(7)=   "Column(0).WrapText=1"
         Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(9)=   "Column(1).Width=1164"
         Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=1085"
         Splits(1)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=17"
         Splits(1)._ColumnProps(14)=   "Column(1).WrapText=1"
         Splits(1)._ColumnProps(15)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(16)=   "Column(2).Width=1402"
         Splits(1)._ColumnProps(17)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(18)=   "Column(2)._WidthInPix=1323"
         Splits(1)._ColumnProps(19)=   "Column(2)._EditAlways=0"
         Splits(1)._ColumnProps(20)=   "Column(2)._ColStyle=17"
         Splits(1)._ColumnProps(21)=   "Column(2).WrapText=1"
         Splits(1)._ColumnProps(22)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(23)=   "Column(3).Width=1191"
         Splits(1)._ColumnProps(24)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(25)=   "Column(3)._WidthInPix=1111"
         Splits(1)._ColumnProps(26)=   "Column(3)._EditAlways=0"
         Splits(1)._ColumnProps(27)=   "Column(3)._ColStyle=17"
         Splits(1)._ColumnProps(28)=   "Column(3).WrapText=1"
         Splits(1)._ColumnProps(29)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(30)=   "Column(4).Width=1191"
         Splits(1)._ColumnProps(31)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(32)=   "Column(4)._WidthInPix=1111"
         Splits(1)._ColumnProps(33)=   "Column(4)._EditAlways=0"
         Splits(1)._ColumnProps(34)=   "Column(4)._ColStyle=131601"
         Splits(1)._ColumnProps(35)=   "Column(4).WrapText=1"
         Splits(1)._ColumnProps(36)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(37)=   "Column(5).Width=1217"
         Splits(1)._ColumnProps(38)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(39)=   "Column(5)._WidthInPix=1138"
         Splits(1)._ColumnProps(40)=   "Column(5)._EditAlways=0"
         Splits(1)._ColumnProps(41)=   "Column(5)._ColStyle=17"
         Splits(1)._ColumnProps(42)=   "Column(5).WrapText=1"
         Splits(1)._ColumnProps(43)=   "Column(5).Order=6"
         Splits(1)._ColumnProps(44)=   "Column(6).Width=1323"
         Splits(1)._ColumnProps(45)=   "Column(6).DividerColor=0"
         Splits(1)._ColumnProps(46)=   "Column(6)._WidthInPix=1244"
         Splits(1)._ColumnProps(47)=   "Column(6)._EditAlways=0"
         Splits(1)._ColumnProps(48)=   "Column(6)._ColStyle=17"
         Splits(1)._ColumnProps(49)=   "Column(6).WrapText=1"
         Splits(1)._ColumnProps(50)=   "Column(6).Order=7"
         Splits(1)._ColumnProps(51)=   "Column(7).Width=2725"
         Splits(1)._ColumnProps(52)=   "Column(7).DividerColor=0"
         Splits(1)._ColumnProps(53)=   "Column(7)._WidthInPix=2646"
         Splits(1)._ColumnProps(54)=   "Column(7)._EditAlways=0"
         Splits(1)._ColumnProps(55)=   "Column(7)._ColStyle=20"
         Splits(1)._ColumnProps(56)=   "Column(7).Visible=0"
         Splits(1)._ColumnProps(57)=   "Column(7).WrapText=1"
         Splits(1)._ColumnProps(58)=   "Column(7).Order=8"
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
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=98,.parent=47,.alignment=2,.valignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=48"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=49"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=51"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=150,.parent=47,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=147,.parent=48"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=148,.parent=49"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=149,.parent=51"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=154,.parent=47,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=151,.parent=48"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=152,.parent=49"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=153,.parent=51"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=32,.parent=47"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=48"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=49"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=51"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=16,.parent=47"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=48"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=49"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=51"
         _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=24,.parent=47"
         _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=48"
         _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=49"
         _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=51"
         _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=158,.parent=47"
         _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=155,.parent=48"
         _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=156,.parent=49"
         _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=157,.parent=51"
         _StyleDefs(70)  =   "Splits(1).Style:id=103,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(71)  =   "Splits(1).CaptionStyle:id=112,.parent=4"
         _StyleDefs(72)  =   "Splits(1).HeadingStyle:id=104,.parent=2"
         _StyleDefs(73)  =   "Splits(1).FooterStyle:id=105,.parent=3"
         _StyleDefs(74)  =   "Splits(1).InactiveStyle:id=106,.parent=5"
         _StyleDefs(75)  =   "Splits(1).SelectedStyle:id=108,.parent=6"
         _StyleDefs(76)  =   "Splits(1).EditorStyle:id=107,.parent=7"
         _StyleDefs(77)  =   "Splits(1).HighlightRowStyle:id=109,.parent=8,.bgcolor=&H800000&"
         _StyleDefs(78)  =   "Splits(1).EvenRowStyle:id=110,.parent=9"
         _StyleDefs(79)  =   "Splits(1).OddRowStyle:id=111,.parent=10"
         _StyleDefs(80)  =   "Splits(1).RecordSelectorStyle:id=113,.parent=11"
         _StyleDefs(81)  =   "Splits(1).FilterBarStyle:id=114,.parent=12"
         _StyleDefs(82)  =   "Splits(1).Columns(0).Style:id=118,.parent=103"
         _StyleDefs(83)  =   "Splits(1).Columns(0).HeadingStyle:id=115,.parent=104"
         _StyleDefs(84)  =   "Splits(1).Columns(0).FooterStyle:id=116,.parent=105"
         _StyleDefs(85)  =   "Splits(1).Columns(0).EditorStyle:id=117,.parent=107"
         _StyleDefs(86)  =   "Splits(1).Columns(1).Style:id=122,.parent=103,.alignment=2,.valignment=2"
         _StyleDefs(87)  =   "Splits(1).Columns(1).HeadingStyle:id=119,.parent=104"
         _StyleDefs(88)  =   "Splits(1).Columns(1).FooterStyle:id=120,.parent=105"
         _StyleDefs(89)  =   "Splits(1).Columns(1).EditorStyle:id=121,.parent=107"
         _StyleDefs(90)  =   "Splits(1).Columns(2).Style:id=166,.parent=103,.alignment=2"
         _StyleDefs(91)  =   "Splits(1).Columns(2).HeadingStyle:id=163,.parent=104"
         _StyleDefs(92)  =   "Splits(1).Columns(2).FooterStyle:id=164,.parent=105"
         _StyleDefs(93)  =   "Splits(1).Columns(2).EditorStyle:id=165,.parent=107"
         _StyleDefs(94)  =   "Splits(1).Columns(3).Style:id=174,.parent=103,.alignment=2,.bold=-1"
         _StyleDefs(95)  =   ":id=174,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(96)  =   ":id=174,.fontname=MS Sans Serif"
         _StyleDefs(97)  =   "Splits(1).Columns(3).HeadingStyle:id=171,.parent=104,.bold=0,.fontsize=825"
         _StyleDefs(98)  =   ":id=171,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(99)  =   ":id=171,.fontname=MS Sans Serif"
         _StyleDefs(100) =   "Splits(1).Columns(3).FooterStyle:id=172,.parent=105,.bold=0,.fontsize=825"
         _StyleDefs(101) =   ":id=172,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(102) =   ":id=172,.fontname=MS Sans Serif"
         _StyleDefs(103) =   "Splits(1).Columns(3).EditorStyle:id=173,.parent=107"
         _StyleDefs(104) =   "Splits(1).Columns(4).Style:id=46,.parent=103,.alignment=2"
         _StyleDefs(105) =   "Splits(1).Columns(4).HeadingStyle:id=43,.parent=104,.alignment=2"
         _StyleDefs(106) =   "Splits(1).Columns(4).FooterStyle:id=44,.parent=105,.alignment=2"
         _StyleDefs(107) =   "Splits(1).Columns(4).EditorStyle:id=45,.parent=107"
         _StyleDefs(108) =   "Splits(1).Columns(5).Style:id=20,.parent=103,.alignment=2"
         _StyleDefs(109) =   "Splits(1).Columns(5).HeadingStyle:id=17,.parent=104"
         _StyleDefs(110) =   "Splits(1).Columns(5).FooterStyle:id=18,.parent=105"
         _StyleDefs(111) =   "Splits(1).Columns(5).EditorStyle:id=19,.parent=107"
         _StyleDefs(112) =   "Splits(1).Columns(6).Style:id=28,.parent=103,.alignment=2"
         _StyleDefs(113) =   "Splits(1).Columns(6).HeadingStyle:id=25,.parent=104,.bold=0,.fontsize=825"
         _StyleDefs(114) =   ":id=25,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(115) =   ":id=25,.fontname=MS Sans Serif"
         _StyleDefs(116) =   "Splits(1).Columns(6).FooterStyle:id=26,.parent=105"
         _StyleDefs(117) =   "Splits(1).Columns(6).EditorStyle:id=27,.parent=107"
         _StyleDefs(118) =   "Splits(1).Columns(7).Style:id=178,.parent=103"
         _StyleDefs(119) =   "Splits(1).Columns(7).HeadingStyle:id=175,.parent=104"
         _StyleDefs(120) =   "Splits(1).Columns(7).FooterStyle:id=176,.parent=105"
         _StyleDefs(121) =   "Splits(1).Columns(7).EditorStyle:id=177,.parent=107"
         _StyleDefs(122) =   "Named:id=33:Normal"
         _StyleDefs(123) =   ":id=33,.parent=0"
         _StyleDefs(124) =   "Named:id=34:Heading"
         _StyleDefs(125) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(126) =   ":id=34,.wraptext=-1"
         _StyleDefs(127) =   "Named:id=35:Footing"
         _StyleDefs(128) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(129) =   "Named:id=36:Selected"
         _StyleDefs(130) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(131) =   "Named:id=37:Caption"
         _StyleDefs(132) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(133) =   "Named:id=38:HighlightRow"
         _StyleDefs(134) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(135) =   "Named:id=39:EvenRow"
         _StyleDefs(136) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(137) =   "Named:id=40:OddRow"
         _StyleDefs(138) =   ":id=40,.parent=33"
         _StyleDefs(139) =   "Named:id=41:RecordSelector"
         _StyleDefs(140) =   ":id=41,.parent=34"
         _StyleDefs(141) =   "Named:id=42:FilterBar"
         _StyleDefs(142) =   ":id=42,.parent=33"
      End
      Begin MSComctlLib.TreeView TvwEmpreendimentos 
         Height          =   3645
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3860
         _ExtentX        =   6800
         _ExtentY        =   6429
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
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   4140
         Visible         =   0   'False
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
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
         Left            =   3960
         TabIndex        =   3
         Top             =   120
         Width           =   2415
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
         Left            =   9420
         TabIndex        =   9
         Top             =   3360
         Width           =   975
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
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Height          =   345
      Left            =   9000
      TabIndex        =   11
      Top             =   4155
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc DatEmpreendimento1 
      Height          =   330
      Left            =   360
      Top             =   4800
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
   Begin MSAdodcLib.Adodc DatEmpreendimento2 
      Height          =   330
      Left            =   360
      Top             =   5160
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
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   10125
      TabIndex        =   12
      Top             =   4155
      Width           =   1035
   End
End
Attribute VB_Name = "FrmRelPosComercializacaoSantander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VFV_VETORPLANOS As New XArray           'Vetor que armazena a classificação dos planos
Public XFO_ExlObj As Excel.Application      ' Create excel object
Public XFI_QtdEmpreendimento As Integer     'Quantidade de empreendimentos

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelPosComercializacaoSantander
End Sub

Private Sub Form_Load()
    Sql = "SELECT tipl_tx_Descricao, '' AS PlanosMensais, '' AS PlanosSemestrais, '' AS PlanosAnuais, '' AS PlanosOutros, '' AS PlanosChaves, '' AS PlanosRepasse, tipl_cd_TipoPlano FROM TiposPlanos WHERE tipl_tx_Sigla <> 'R'"
    subCarregaVetor Sql, Array("tipl_tx_Descricao", "PlanosMensais", "PlanosSemestrais", "PlanosAnuais", "PlanosOutros", "PlanosChaves", "PlanosRepasse", "tipl_cd_TipoPlano"), VFV_VETORPLANOS, TDBGrid1
    
    subConectarControleDadosNV DatEmpreendimento1, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_cd_empreendimento", Estatico
    subConectarControleDadosNV DatEmpreendimento2, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_cd_empreendimento", Estatico
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
        
        nodeEmpd.Tag = "Nó"
        
        'Cria as folhas do nó criado
        CarregaRegistros nodeEmpd
        
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
    FrmRelPosComercializacaoSantander.Height = 5490
    CmdImprimir.Top = 4635
    CmdCancelar.Top = 4635
    FraPrincipal.Height = 4605
        
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
    FraPrincipal.Height = 4125
    CmdImprimir.Top = 4155
    CmdCancelar.Top = 4155
    FrmRelPosComercializacaoSantander.Height = 5020
End Sub

'Avança unidades na barra de progresso
Private Sub AvancarProgresso(valor As Integer)
    If (pgbImpressao.Value + valor <= pgbImpressao.Max) Then
        pgbImpressao.Value = pgbImpressao.Value + valor
    End If
End Sub

'Imprime planilha
Private Sub CmdImprimir_Click()
  
    Dim XLO_COMANDO     As New ADODB.Command
    Dim rs              As New ADODB.Recordset
    Dim XLI_INDEXADOR   As Integer
    
    Dim XLT_MENSAIS             As String   'Armazena o código dos planos classificados como Mensais
    Dim XLT_SEMESTRAIS          As String   'Armazena o código dos planos classificados como Semestrais
    Dim XLT_ANUAIS              As String   'Armazena o código dos planos classificados como Anuais
    Dim XLT_CHAVES              As String   'Armazena o código dos planos classificados como Chaves
    Dim XLT_REPASSE             As String   'Armazena o código dos planos classificados como Repasse
    Dim XLT_OUTROS              As String   'Os títulos desses planos serão incluídos no cálculo dos valores recebidos e a receber, mas não entram na contagem de parcelas.
     
    Dim i                       As Integer  'Auxiliar
    Dim j                       As Integer  'Auxiliar
    Dim sheet                   As Integer  'Quantidade de planilhas no workbook
    Dim XLT_EMPREENDIMENTOS     As String
    Dim XLO_EMPREENDIMENTO()    As String
    Dim XLT_NOMEEMPREENDIMENTO  As String
        
    
    On Error GoTo TrataErro
    
    'Pega classificação dos planos
    TDBGrid1.MoveFirst
    While Not TDBGrid1.EOF
        If TDBGrid1.Columns("PlanosMensais").Value = "-1" Then
            XLT_MENSAIS = XLT_MENSAIS + ", " + TDBGrid1.Columns("CodigoPlano")
        End If
        
        If TDBGrid1.Columns("PlanosSemestrais").Value = "-1" Then
            XLT_SEMESTRAIS = XLT_SEMESTRAIS + ", " + TDBGrid1.Columns("CodigoPlano")
        End If
        
        If TDBGrid1.Columns("PlanosAnuais").Value = "-1" Then
            XLT_ANUAIS = XLT_ANUAIS + ", " + TDBGrid1.Columns("CodigoPlano")
        End If
        
        If TDBGrid1.Columns("PlanosOutros").Value = "-1" Then
            XLT_OUTROS = XLT_OUTROS + ", " + TDBGrid1.Columns("CodigoPlano")
        End If
        
        If TDBGrid1.Columns("PlanosChaves").Value = "-1" Then
            XLT_CHAVES = XLT_CHAVES + ", " + TDBGrid1.Columns("CodigoPlano")
        End If
        
        If TDBGrid1.Columns("PlanosRepasse").Value = "-1" Then
            XLT_REPASSE = XLT_REPASSE + ", " + TDBGrid1.Columns("CodigoPlano")
        End If
        
        TDBGrid1.MoveNext
    Wend
    
    'Remove a vírgula do início
    XLT_MENSAIS = Mid(XLT_MENSAIS, 2)
    XLT_SEMESTRAIS = Mid(XLT_SEMESTRAIS, 2)
    XLT_ANUAIS = Mid(XLT_ANUAIS, 2)
    XLT_OUTROS = Mid(XLT_OUTROS, 2)
    XLT_CHAVES = Mid(XLT_CHAVES, 2)
    XLT_REPASSE = Mid(XLT_REPASSE, 2)
    
    If XLT_MENSAIS = "" Or XLT_SEMESTRAIS = "" Or XLT_ANUAIS = "" Or XLT_CHAVES = "" Or XLT_REPASSE = "" Then
        MsgBox "É necessário informar pelo menos um plano para cada classificação (mensal, semestral, anual, chaves e repasse).", vbInformation, "Atenção"
        FrmRelRelacaoClientes.MousePointer = vbDefault
        Exit Sub
    End If
                                  
    'Seta o Titulo da Tela de Dialogo salvar
    dlgSalvar.DialogTitle = "Salvar Arquivo"
    
    'Seta o filtro das extensões de arquivos que será usada
    dlgSalvar.Filter = "Documento do Execl(*.xls)|*.xls"
    
    'Indica qual sera o filtro padrão
    dlgSalvar.FilterIndex = 1
                
    dlgSalvar.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
            
    'Exibe a tela do dialogo para o usuário
    dlgSalvar.ShowSave
        
    'Verifica se o usuário selecionou o caminho do arquivo
    If (dlgSalvar.FileName <> "") Then
        
        'Indexador
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
                
        'Verifica quantas planilhas serão criadas
        i = 1
        j = 0
        While i < TvwEmpreendimentos.Nodes.Count
            If (TvwEmpreendimentos.Nodes.Item(i).Tag = "Nó") And (TvwEmpreendimentos.Nodes.Item(i).Checked = True) Then
                j = j + 1
            End If
            i = i + 1
        Wend
                
        'Adiciona as planilhas necessárias
        If (j > 2) Then
            For i = 2 To j
                XFO_ExlObj.Sheets.Add
            Next
        End If
                
        'Impede que o usuário veja a planilha enquanto ela está sendo preenchida
        XFO_ExlObj.Visible = False
        
        'Inicializa a barra de progresso
        InicializarProgresso 0, j * 12 + 1
                        
        AvancarProgresso 1
                
        i = 1
        sheet = 1
        XLT_EMPREENDIMENTOS = ""
        
        While i < TvwEmpreendimentos.Nodes.Count
            If (TvwEmpreendimentos.Nodes.Item(i).Tag = "Nó") And (TvwEmpreendimentos.Nodes.Item(i).Checked = True) Then
                
                XLO_EMPREENDIMENTO = Split(TvwEmpreendimentos.Nodes.Item(i).Text, "-")
                XLT_EMPREENDIMENTOS = Trim(XLO_EMPREENDIMENTO(0))
                XLT_NOMEEMPREENDIMENTO = Trim(XLO_EMPREENDIMENTO(1))
                              
                For j = i + 1 To i + XFI_QtdEmpreendimento - 1
                    If (TvwEmpreendimentos.Nodes.Item(j).Checked = True) Then
                        XLO_EMPREENDIMENTO = Split(TvwEmpreendimentos.Nodes.Item(j).Text, "-")
                        XLT_EMPREENDIMENTOS = XLT_EMPREENDIMENTOS + "," & Trim(XLO_EMPREENDIMENTO(0))
                    End If
                Next
                
            
                '------------------
                'Imprime relatório
                '------------------
                If (XLT_EMPREENDIMENTOS <> "") Then
                    With XLO_COMANDO
                        .CommandTimeout = 300
                        .CommandText = "spCAPI_sel_posicaoComercializacaoSANTANDER"
                        .CommandType = adCmdStoredProc
                        Set XLO_COMANDO.ActiveConnection = Conexao
                        'Cria automaticamente todos os parâmetros dentro do objeto Comando
                        .Parameters.Refresh
                    End With
                    
                    XLO_COMANDO.Parameters(1).Value = PCodEmpresa 'Empresa
                    XLO_COMANDO.Parameters(2).Value = DtBase.Value 'Data base
                    XLO_COMANDO.Parameters(3).Type = adBSTR
                    XLO_COMANDO.Parameters(3).Value = XLT_EMPREENDIMENTOS 'Empreendimento
                    XLO_COMANDO.Parameters(4).Value = XLI_INDEXADOR     'Indexador
                    XLO_COMANDO.Parameters(5).Value = XLT_MENSAIS       'Planos mensais
                    XLO_COMANDO.Parameters(6).Value = XLT_SEMESTRAIS    'Planos semestrais
                    XLO_COMANDO.Parameters(7).Value = XLT_ANUAIS        'Planos anuais
                    XLO_COMANDO.Parameters(8).Value = XLT_OUTROS        'Planos outros
                    XLO_COMANDO.Parameters(9).Value = XLT_CHAVES        'Planos chaves
                    XLO_COMANDO.Parameters(10).Value = XLT_REPASSE       'Planos repasse
                    
                    Set rs = XLO_COMANDO.Execute
                         
                    'Verifica se a consulta retornou dados, caso contrário finaliza a rotina
                    If rs.EOF Then
                        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
                        FrmRelRelacaoClientes.MousePointer = vbDefault
                        Exit Sub
                    End If
            
                    AvancarProgresso 2
                    
                    ExportaExcel rs, XLT_NOMEEMPREENDIMENTO, sheet
                    
                    'Incrementa a quantidade de planilhas
                    sheet = sheet + 1
                End If
            End If
            
            XLT_EMPREENDIMENTOS = ""
            i = i + XFI_QtdEmpreendimento
        Wend
        
        'Salva a planilha no caminho selecionado pelo usuário
        XFO_ExlObj.Workbooks(1).SaveAs dlgSalvar.FileName
        XFO_ExlObj.Workbooks(1).Close
        XFO_ExlObj.Application.Quit
        Set XFO_ExlObj = Nothing
  
                    
        'Limpa o caminho do arquivo
        dlgSalvar.FileName = ""
        FrmRelPosComercializacaoSantander.MousePointer = vbDefault
        MsgBox "Relatório Exportado com sucesso!", vbInformation, "Informação"
             
        FinalizarProgresso
    End If
      
Exit Sub

TrataErro:
    If Err.Number = cdlCancel Then
        Exit Sub
    End If
    FrmRelPosComercializacaoSantander.MousePointer = vbDefault
    MsgBox "Ocorreu um erro na exportação do relatório. Contate o suporte. Detalhe do Erro: " & Err.Description, vbCritical, "Erro"
    DeleteFile (dlgSalvar.FileName)

End Sub

Private Sub ExportaExcel(dados As ADODB.Recordset, nomeEmpreendimento As String, sheet As Integer)
    
    'Seleciona a próxima planilha vazia para ser preenchida
    XFO_ExlObj.Sheets(sheet).Select
    
    
    'Configurando margem da página
    With XFO_ExlObj.Sheets(sheet).PageSetup
        .LeftMargin = Application.InchesToPoints(0.64)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.54)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .Orientation = xlLandscape
        .PrintTitleRows = "$1:$12" 'Repete o cabeçalho em todas as páginas
        .Zoom = 56
    End With
    
                                                   
    'Preenchendo as células
    With XFO_ExlObj.Sheets(sheet)
        'Seta o nome da aba da planilha
        .Name = nomeEmpreendimento
        
        XFO_ExlObj.ActiveSheet.Range("A1").Select
        .Shapes.AddPicture "c:\QualiAdmFin\QualiCapi\Relatorios\Imagens\imgPosComercializacaoSantander.png", True, True, 5, 5, 200, 30
        .Cells(5, 20).Value = "DATA BASE"
        .Cells(5, 20).Font.Bold = True
        .Cells(5, 21).Value = DtBase.Value
        .Cells(5, 21).Interior.ColorIndex = 15
        .Cells(5, 21).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Cells(5, 21).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        .Cells(5, 21).Borders(xlEdgeTop).Weight = xlMedium
        .Cells(5, 21).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Cells(5, 21).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Cells(5, 21).Borders(xlEdgeBottom).Weight = xlMedium
        .Cells(5, 21).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Cells(5, 21).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        .Cells(5, 21).Borders(xlEdgeLeft).Weight = xlMedium
        .Cells(5, 21).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Cells(5, 21).Borders(xlEdgeRight).ColorIndex = xlAutomatic
        .Cells(5, 21).Borders(xlEdgeRight).Weight = xlMedium
        .Range("A1:U1").Merge
        .Range("A1:U1").HorizontalAlignment = xlHAlignCenter
        .Cells(1, 1).Value = "SUPERINTENDÊNCIA DE CRÉDITO IMOBILIÁRIO - PJ"
        .Cells(1, 1).Font.Bold = True
        .Range("A3:U3").Merge
        .Range("A3:U3").HorizontalAlignment = xlHAlignCenter
        .Cells(3, 1).Value = "RESUMO DAS VENDAS REALIZADAS"
        .Cells(3, 1).Font.Bold = True
        .Cells(4, 1).Value = "EMPRESA: " & PEmpresa
        .Cells(4, 1).Font.Bold = True
        .Cells(5, 1).Value = "EMPREENDIMENTO: " & nomeEmpreendimento
        .Cells(5, 1).Font.Bold = True
        .Cells(6, 1).Value = "ENDEREÇO DO EMPREENDIMENTO":
        .Cells(6, 1).Font.Bold = True
   
        .Range("A7:U9").Font.Bold = True
        .Range("A7:U9").HorizontalAlignment = xlHAlignCenter
       
        'N° da unidade
        .Cells(7, 1).Value = "N.°"
        .Cells(8, 1).Value = "da"
        .Cells(9, 1).Value = "Unid."
        .Range("A7:A9").Interior.ColorIndex = 46
        .Range("A7:A9").Font.ColorIndex = 2
        
        'Tipo de Venda
        .Cells(7, 2).Value = "Tipo"
        .Cells(8, 2).Value = "de"
        .Cells(9, 2).Value = "Venda"
        .Range("B7:B9").Interior.ColorIndex = 46
        .Range("B7:B9").Font.ColorIndex = 2
        
        'Comprador
        .Range("C7:G7").Merge
        .Cells(7, 3).Value = "Comprador"
        .Range("C7:G7").Interior.ColorIndex = 15
                      
        'Nome
        .Range("C8:C9").Merge
        .Cells(8, 3).Value = "Nome"
        .Range("C8:C9").HorizontalAlignment = xlHAlignCenter
        .Range("C8:C9").VerticalAlignment = xlCenter
        .Range("C8:C9").Interior.ColorIndex = 3
        .Cells(8, 3).Font.ColorIndex = 2
        
        'CPF
        .Range("D8:D9").Merge
        .Cells(8, 4).Value = "CPF"
        .Range("D8:D9").HorizontalAlignment = xlHAlignCenter
        .Range("D8:D9").VerticalAlignment = xlVAlignCenter
        .Range("D8:D9").Interior.ColorIndex = 3
        .Cells(8, 4).Font.ColorIndex = 2
             
        'Endereço
        .Range("E8:E9").Merge
        .Cells(8, 5).Value = "Endereço"
        .Range("E8:E9").HorizontalAlignment = xlHAlignCenter
        .Range("E8:E9").VerticalAlignment = xlVAlignCenter
        .Range("E8:E9").Interior.ColorIndex = 3
        .Cells(8, 5).Font.ColorIndex = 2
        
        'CEP
        .Range("F8:F9").Merge
        .Cells(8, 6).Value = "CEP"
        .Range("F8:F9").HorizontalAlignment = xlHAlignCenter
        .Range("F8:F9").VerticalAlignment = xlVAlignCenter
        .Range("F8:F9").Interior.ColorIndex = 3
        .Cells(8, 6).Font.ColorIndex = 2
        
        'Telefone
        .Range("G8:G9").Merge
        .Cells(8, 7).Value = "Telefone"
        .Range("G8:G9").HorizontalAlignment = xlHAlignCenter
        .Range("G8:G9").VerticalAlignment = xlVAlignCenter
        .Range("G8:G9").Interior.ColorIndex = 3
        .Cells(8, 7).Font.ColorIndex = 2
        
        'Data da venda
        .Cells(7, 8).Value = "Data"
        .Cells(8, 8).Value = "da"
        .Cells(9, 8).Value = "Venda"
        .Range("H7:H9").Interior.ColorIndex = 3
        .Range("H7:H9").Font.ColorIndex = 2
          
        'Valor de venda
        .Cells(7, 9).Value = "Valor"
        .Cells(8, 9).Value = "de"
        .Cells(9, 9).Value = "Venda"
        .Range("I7:I9").Interior.ColorIndex = 3
        .Range("I7:I9").Font.ColorIndex = 2
                 
        'Valores já recebidos
        .Cells(7, 10).Value = "Valores"
        .Cells(8, 10).Value = "já"
        .Cells(9, 10).Value = "Recebidos"
        .Range("J7:J9").Interior.ColorIndex = 3
        .Range("J7:J9").Font.ColorIndex = 2
        
        'Valores a Receber
        .Range("K6:N6").Merge
        .Cells(6, 11).Value = "Valores devem ser corrigidos mensalmente"
        .Range("K6:N6").HorizontalAlignment = xlHAlignCenter
        .Range("K6:N6").Interior.ColorIndex = 36
        .Cells(6, 11).Font.Bold = True
        
        'Durante a Construção
        .Cells(7, 11).Value = "Valores a Receber"
        .Cells(8, 11).Value = "Durante a"
        .Cells(9, 11).Value = "Construção"
        .Range("K7:K9").Interior.ColorIndex = 3
        .Range("K7:K9").Font.ColorIndex = 2
            
        'Parcela de Chaves
        .Cells(7, 12).Value = "Valores da"
        .Cells(8, 12).Value = "Parcela "
        .Cells(9, 12).Value = "de Chaves"
        .Range("L7:L9").Interior.ColorIndex = 3
        .Range("L7:L9").Font.ColorIndex = 2
        
        'Pós Chaves
        .Cells(7, 13).Value = "Valores a"
        .Cells(8, 13).Value = "Receber "
        .Cells(9, 13).Value = "Pós Chaves"
        .Range("M7:M9").Interior.ColorIndex = 3
        .Range("M7:M9").Font.ColorIndex = 2
                
        'Parcela a Repassar
        .Cells(7, 14).Value = "Valor da"
        .Cells(8, 14).Value = "Parcela a"
        .Cells(9, 14).Value = "Repassar"
        .Range("N7:N9").Interior.ColorIndex = 3
        .Range("N7:N9").Font.ColorIndex = 2
        
        'Total a receber
        .Range("O7:O9").Merge
        .Range("O7:O9").WrapText = True
        .Cells(7, 15).Value = "Total a Receber"
        .Cells(7, 15).Font.Bold = True
        .Range("O7:O9").HorizontalAlignment = xlHAlignCenter
        .Range("O7:O9").VerticalAlignment = xlVAlignCenter
        .Range("O7:O9").Interior.ColorIndex = 3
        .Range("O7:O9").Font.ColorIndex = 2
        
        'Forma de pagamento
        .Range("P7:U7").Merge
        .Cells(7, 16).Value = "FORMA DE PAGAMENTO"
        .Range("P7:P7").Interior.ColorIndex = 15
        
        'Tipo parcela
        .Cells(8, 16).Value = "Tipo"
        .Cells(9, 16).Value = "Parcela"
        .Range("P8:P9").Interior.ColorIndex = 3
        .Range("P8:P9").Font.ColorIndex = 2
                
        'Tipo Indexador
        .Cells(8, 17).Value = "Tipo"
        .Cells(9, 17).Value = "Indexador"
        .Range("Q8:Q9").Interior.ColorIndex = 3
        .Range("Q8:Q9").Font.ColorIndex = 2
        
        'Quantidade parcela
        .Cells(8, 18).Value = "Quantidade"
        .Cells(9, 18).Value = "Parcela"
        .Range("R8:R9").Interior.ColorIndex = 3
        .Range("R8:R9").Font.ColorIndex = 2
        
        'Valor parcela
        .Cells(8, 19).Value = "Valor"
        .Cells(9, 19).Value = "Parcela"
        .Range("S8:S9").Interior.ColorIndex = 3
        .Range("S8:S9").Font.ColorIndex = 2
        
        'Valor total
        .Cells(8, 20).Value = "Valor"
        .Cells(9, 20).Value = "Total"
        .Range("T8:T9").Interior.ColorIndex = 3
        .Range("T8:T9").Font.ColorIndex = 2
        
        '1º Vencimento da parcela
        .Cells(8, 21).Value = "1º Vencto."
        .Cells(9, 21).Value = "Parcela"
        .Range("U8:U9").Interior.ColorIndex = 3
        .Range("U8:U9").Font.ColorIndex = 2
     End With
     
     
     AvancarProgresso 1
     
                              
    'Formata a largura das colunas
    With XFO_ExlObj.Sheets(sheet)
        .Range("A7").ColumnWidth = 6
        .Range("B7").ColumnWidth = 6
        .Range("C7").ColumnWidth = 35.14
        .Range("D7").ColumnWidth = 12.57
        .Range("E7").ColumnWidth = 35.14
        .Range("F7").ColumnWidth = 12.57
        .Range("G7").ColumnWidth = 13
        .Range("H7").ColumnWidth = 10
        .Range("I7").ColumnWidth = 12.57
        .Range("J7").ColumnWidth = 12.57
        .Range("K7").ColumnWidth = 15.57
        .Range("L7").ColumnWidth = 13.57
        .Range("M7").ColumnWidth = 13.57
        .Range("N7").ColumnWidth = 13.57
        .Range("O7").ColumnWidth = 15
        .Range("P7").ColumnWidth = 8
        .Range("Q7").ColumnWidth = 10
        .Range("R7").ColumnWidth = 11
        .Range("S7").ColumnWidth = 14.14
        .Range("T7").ColumnWidth = 16.14
        .Range("U7").ColumnWidth = 16.86
    End With
    

    'Adiciona a borda na planilha
    With XFO_ExlObj.Sheets(sheet).Range("A7:U9").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlMedium
    End With
        
    With XFO_ExlObj.Sheets(sheet).Range("A7:U9").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlMedium
    End With
        
    With XFO_ExlObj.Sheets(sheet).Range("A7:U9").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlMedium
    End With
         
    With XFO_ExlObj.Sheets(sheet).Range("A7:U9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlMedium
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("A7:A9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
        
    With XFO_ExlObj.Sheets(sheet).Range("B7:B9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
        
    With XFO_ExlObj.Sheets(sheet).Range("C7:G9").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
        
    With XFO_ExlObj.Sheets(sheet).Range("C7:G9").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("H7:H9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("H7:H9").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("I7:I9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("J7:J9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("K7:K9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("L7:L9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
 
    With XFO_ExlObj.Sheets(sheet).Range("M7:M9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("N7:N9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("O7:O9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
            
    With XFO_ExlObj.Sheets(sheet).Range("P7:U7").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("P8:P9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("Q8:Q9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("R8:R9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("S8:S9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("T8:T9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("U8:U9").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    
    AvancarProgresso 1
    
    
    Dim cont, i As Integer
    
    i = 10
    cont = i
    
    'Preenche a planilha com os dados
    While Not dados.EOF
        XFO_ExlObj.Sheets(sheet).Range("Q" & i & ":Q" & i + 4).NumberFormat = "@"
        
        With XFO_ExlObj.Sheets(sheet)
            .Cells(i, 1) = dados!Imovel
            .Cells(i, 2) = dados!TipoVenda
            .Cells(i, 3) = dados!Comprador
            .Cells(i, 4) = dados!CPF
            .Cells(i, 5) = dados!Endereco
            .Cells(i, 6) = dados!CEP
            .Cells(i, 7) = dados!Telefone
            .Cells(i, 8) = dados!DataVenda
            .Cells(i, 9) = dados!ValorVenda
            .Cells(i, 10) = dados!ValorRecebido
            .Cells(i, 11) = dados!DuranteConstrucao
            .Cells(i, 12) = dados!chaves
            .Cells(i, 13) = dados!PosChave
            .Cells(i, 14) = dados!repasse
            .Cells(i, 15).Formula = "=SUM(K" & i & ":N" & i & ")"
            
            'Tipo parcela
            .Cells(i, 16) = "M" 'mensal
            .Cells(i + 1, 16) = "A" 'anual
            .Cells(i + 2, 16) = "S" 'semestral
            .Cells(i + 3, 16) = "Chaves"
            .Cells(i + 4, 16) = "Repasse"
            
            'Tipo indexador
            .Cells(i, 17) = "40"
            .Cells(i + 1, 17) = "00"
            .Cells(i + 2, 17) = "40"
            .Cells(i + 3, 17) = "40"
            .Cells(i + 4, 17) = "04"
            
            'Quantidade de parcelas com vencimento antes da entrega das chaves
            .Cells(i, 18) = dados!QtdeParcelaMensal
            .Cells(i + 1, 18) = dados!QtdeParcelaAnual
            .Cells(i + 2, 18) = dados!QtdeParcelaSemestral
            .Cells(i + 3, 18) = dados!QtdeParcelaChaves
            .Cells(i + 4, 18) = dados!QtdeParcelaRepasse
            
            'Valor da parcela
            .Cells(i, 19) = dados!ValorParcelaMensal
            .Cells(i + 1, 19) = dados!ValorParcelaAnual
            .Cells(i + 2, 19) = dados!ValorParcelaSemestral
            .Cells(i + 3, 19) = dados!ValorParcelaChaves
            .Cells(i + 4, 19) = dados!ValorParcelaRepasse
            
            'Valor total das parcelas com vencimento antes da entrega das chaves
            .Cells(i, 20) = dados!ValorTotalMensal
            .Cells(i + 1, 20) = dados!ValorTotalAnual
            .Cells(i + 2, 20) = dados!ValorTotalSemestral
            .Cells(i + 3, 20) = dados!ValorTotalChaves
            .Cells(i + 4, 20) = dados!ValorTotalRepasse
            
            'Data do primeiro vencimento das parcelas
            .Cells(i, 21) = dados!PrimeiroVenctoParcelaMensal
            .Cells(i + 1, 21) = dados!PrimeiroVenctoParcelaAnual
            .Cells(i + 2, 21) = dados!PrimeiroVenctoParcelaSemestral
            .Cells(i + 3, 21) = dados!PrimeiroVenctoParcelaChaves
            .Cells(i + 4, 21) = dados!PrimeiroVenctoParcelaRepasse
        End With
        
        'Coloca bordas
        With XFO_ExlObj.Sheets(sheet).Range("A" & cont & ":U" & cont).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
        End With
    
        With XFO_ExlObj.Sheets(sheet).Range("A" & cont & ":U" & cont).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
        End With
    
        With XFO_ExlObj.Sheets(sheet).Range("A" & cont & ":U" & cont).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
        End With
       
        With XFO_ExlObj.Sheets(sheet).Range("A" & cont & ":U" & cont).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
        End With
    
        With XFO_ExlObj.Sheets(sheet).Range("A" & cont & ":U" & cont).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
        End With
        
        With XFO_ExlObj.Sheets(sheet).Range("U" & cont & ":U" & cont).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .Weight = 3
        End With
        
        dados.MoveNext
        i = i + 5
        cont = cont + 5
    Wend
    
    
    AvancarProgresso 4
    
    
    'Define o formato dos números
    XFO_ExlObj.Sheets(sheet).Range("I10:O" & cont).NumberFormat = "#,##0.00"
    XFO_ExlObj.Sheets(sheet).Range("S10:T" & cont).NumberFormat = "#,##0.00"
    
    'Define o formato do campo data da venda
    XFO_ExlObj.Sheets(sheet).Range("H10:H" & cont).NumberFormat = "m/d/yyyy"
    XFO_ExlObj.Sheets(sheet).Range("U10:U" & cont).NumberFormat = "m/d/yyyy"
                
    
    cont = cont + 1
        
    With XFO_ExlObj.Sheets(sheet).Range("P11:U" & cont - 2).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.Sheets(sheet).Range("P11:U" & cont - 2).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.Sheets(sheet).Range("P11:U" & cont - 2).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
   
    With XFO_ExlObj.Sheets(sheet).Range("P11:U" & cont - 2).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.Sheets(sheet).Range("P11:U" & cont - 2).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.Sheets(sheet).Range("U11:U" & cont - 2).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = 3
    End With
        
    With XFO_ExlObj.Sheets(sheet)
        'Formata a fonte e o tamanho da fonte dos campos nome, endereço e Tipo parcela
        .Range("C10:C" & cont - 1).Font.Size = 8
        .Range("E10:E" & cont).Font.Size = 8
        .Range("P10:P" & cont).Font.Size = 10
    
        'Formata o alinhamento do conteúdo
        .Range("A10:B" & cont).HorizontalAlignment = xlHAlignCenter
        .Range("H10:H" & cont).HorizontalAlignment = xlHAlignCenter
        .Range("P10:R" & cont).HorizontalAlignment = xlHAlignCenter
        .Range("A" & cont & ":C" & cont).HorizontalAlignment = xlHAlignCenter
    End With
    
    AvancarProgresso 1
        
    With XFO_ExlObj.Sheets(sheet)
        .Range("H10:O" & cont).Font.Bold = True
        .Range("R10:R" & cont).Font.Bold = True
        .Range("A" & cont & ":U" & cont).Font.Bold = True
        .Range("A" & cont & ":U" & cont).Interior.ColorIndex = 3
        .Range("A" & cont & ":U" & cont).Font.ColorIndex = 2
        .Range("A" & cont & ":B" & cont).Merge
        .Cells(cont, 1).Formula = "=COUNTA(A10:A" & cont - 1 & ")"
        .Cells(cont, 3).Formula = "=COUNTA(C10:C" & cont - 1 & ")"
        .Range("D" & cont & ":H" & cont).Merge
        .Cells(cont, 4).Value = "TOTAL GERAL"
        .Cells(cont, 4).HorizontalAlignment = xlHAlignCenter
        .Cells(cont, 4).Font.Bold = True
        .Cells(cont, 9).Formula = "=SUM(I10:I" & cont - 1 & ")"
        .Cells(cont, 10).Formula = "=SUM(J10:J" & cont - 1 & ")"
        .Cells(cont, 11).Formula = "=SUM(K10:K" & cont - 1 & ")"
        .Cells(cont, 12).Formula = "=SUM(L10:L" & cont - 1 & ")"
        .Cells(cont, 13).Formula = "=SUM(M10:M" & cont - 1 & ")"
        .Cells(cont, 14).Formula = "=SUM(N10:N" & cont - 1 & ")"
        .Cells(cont, 15).Formula = "=SUM(O10:O" & cont - 1 & ")"
        .Cells(cont, 20).Formula = "=SUM(T10:T" & cont - 1 & ")"
    End With
    
    AvancarProgresso 1
                           
    With XFO_ExlObj.Sheets(sheet).Range("A" & cont & ":U" & cont).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = 3
    End With

    With XFO_ExlObj.Sheets(sheet).Range("A" & cont & ":U" & cont).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = 3
    End With
   
    With XFO_ExlObj.Sheets(sheet).Range("A" & cont & ":U" & cont).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = 3
    End With

    With XFO_ExlObj.Sheets(sheet).Range("A" & cont & ":U" & cont).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = 3
    End With

    With XFO_ExlObj.Sheets(sheet).Range("A" & cont & ":U" & cont).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = 3
    End With
    
    AvancarProgresso 1
    
    cont = cont + 2
    
    With XFO_ExlObj.Sheets(sheet)
        .Cells(cont, 1).Value = "Tipo de Venda"
        .Cells(cont, 3).Value = "P = Unidade Permutada"
        .Cells(cont, 11).Value = "Declaramos de que as informações aqui prestadas refletem fielmente"
        cont = cont + 1
        .Cells(cont, 3).Value = "Q = Unidade Quitada"
        .Cells(cont, 11).Value = "a situação atual dos contratos de venda das unidades relacionadas"
        cont = cont + 1
        .Cells(cont, 3).Value = "V.D. = Venda Direta (sem financiamento bancário)"
        .Cells(cont, 14).Value = "Salvador, 09 de junho de 2008."
        .Cells(cont, 14).HorizontalAlignment = xlHAlignRight
        cont = cont + 1
        .Cells(cont, 3).Value = "R = Venda com promessa de repasse (com financiamento bancário)"
        cont = cont + 1
        .Cells(cont, 3).Value = "E = Estoque (unidade ainda não comercializada e disponível para venda)"
        
        .Range("L" & cont & ":N" & cont).Merge
        .Range("L" & cont & ":N" & cont).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("L" & cont & ":N" & cont).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        cont = cont + 1
        .Range("L" & cont & ":N" & cont).Merge
        .Cells(cont, 12).Value = "Antonio Carlos Costa Andrade"
        .Cells(cont, 12).HorizontalAlignment = xlHAlignCenter
    End With
    
    'Formata para não exibir as linhas de grade da planilha
    XFO_ExlObj.ActiveWindow.DisplayGridlines = False
    
    XFO_ExlObj.ActiveSheet.Range("E10").Select
    XFO_ExlObj.ActiveWindow.FreezePanes = True

    AvancarProgresso 1
    
End Sub



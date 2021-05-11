VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRelAnexoFluxoRecebiveisItau 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Anexo de Fluxo Recebido"
   ClientHeight    =   5760
   ClientLeft      =   3330
   ClientTop       =   3615
   ClientWidth     =   9975
   Icon            =   "FrmAnexoFluxoRecebido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraPrincipal 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9810
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   6520
         TabIndex        =   9
         Top             =   4080
         Width           =   3135
         Begin VB.TextBox txtNumContrato 
            Height          =   375
            Left            =   1320
            TabIndex        =   11
            Top             =   140
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DtBase 
            Height          =   375
            Left            =   1320
            TabIndex        =   13
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   63504385
            CurrentDate     =   39692
         End
         Begin VB.Label lblDataBase 
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
            TabIndex        =   12
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblNumContrato 
            Caption         =   "Nº Contrato:"
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
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
      End
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
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   4080
         Width           =   6135
         Begin VB.OptionButton optIndexador2 
            Caption         =   "Indexador 2"
            Height          =   195
            Left            =   4320
            TabIndex        =   8
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton optIndexador1 
            Caption         =   "Indexador 1"
            Height          =   255
            Left            =   2400
            TabIndex        =   7
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optvlHistorico 
            Caption         =   "Valor Histórico"
            Height          =   195
            Left            =   480
            TabIndex        =   6
            Top             =   480
            Width           =   1455
         End
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Height          =   3645
         Left            =   4080
         TabIndex        =   4
         Top             =   360
         Width           =   5600
         _ExtentX        =   9869
         _ExtentY        =   6429
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
         Columns(1).Caption=   "Mensais"
         Columns(1).DataField=   "PlanoMensais"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   68
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Intermediárias"
         Columns(2).DataField=   "PlanoIntermediarias"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   68
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Chaves"
         Columns(3).DataField=   "PlanoChaves"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   68
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Repasse"
         Columns(4).DataField=   "PlanoRepasse"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "CodigoPlano"
         Columns(5).DataField=   "tipl_cd_TipoPlano"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3810"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3731"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1561"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1482"
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
         Splits(0)._ColumnProps(44)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(45)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(46)=   "Column(5).Order=6"
         Splits(1)._UserFlags=   0
         Splits(1).ExtendRightColumn=   -1  'True
         Splits(1).MarqueeStyle=   4
         Splits(1).SizeMode=   2
         Splits(1).Size  =   4
         Splits(1).Size.vt=   2
         Splits(1).RecordSelectors=   0   'False
         Splits(1).RecordSelectorWidth=   503
         Splits(1).ScrollBars=   2
         Splits(1).DividerColor=   12632256
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=6"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=6800"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=6720"
         Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=20"
         Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(7)=   "Column(0).WrapText=1"
         Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(9)=   "Column(1).Width=1191"
         Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=1111"
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
         Splits(1)._ColumnProps(23)=   "Column(3).Width=1244"
         Splits(1)._ColumnProps(24)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(25)=   "Column(3)._WidthInPix=1164"
         Splits(1)._ColumnProps(26)=   "Column(3)._EditAlways=0"
         Splits(1)._ColumnProps(27)=   "Column(3)._ColStyle=17"
         Splits(1)._ColumnProps(28)=   "Column(3).WrapText=1"
         Splits(1)._ColumnProps(29)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(30)=   "Column(4).Width=1244"
         Splits(1)._ColumnProps(31)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(32)=   "Column(4)._WidthInPix=1164"
         Splits(1)._ColumnProps(33)=   "Column(4)._EditAlways=0"
         Splits(1)._ColumnProps(34)=   "Column(4)._ColStyle=17"
         Splits(1)._ColumnProps(35)=   "Column(4).WrapText=1"
         Splits(1)._ColumnProps(36)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(37)=   "Column(5).Width=2725"
         Splits(1)._ColumnProps(38)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(39)=   "Column(5)._WidthInPix=2646"
         Splits(1)._ColumnProps(40)=   "Column(5)._EditAlways=0"
         Splits(1)._ColumnProps(41)=   "Column(5)._ColStyle=20"
         Splits(1)._ColumnProps(42)=   "Column(5).Visible=0"
         Splits(1)._ColumnProps(43)=   "Column(5).WrapText=1"
         Splits(1)._ColumnProps(44)=   "Column(5).Order=6"
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
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=92,.parent=49"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=93,.parent=51"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=98,.parent=47,.alignment=2,.valignment=2"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=48"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=49"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=51"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=150,.parent=47,.alignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=147,.parent=48"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=148,.parent=49"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=149,.parent=51"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=154,.parent=47,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=151,.parent=48"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=152,.parent=49"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=153,.parent=51"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=16,.parent=47"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=13,.parent=48"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=14,.parent=49"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=15,.parent=51"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=158,.parent=47"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=155,.parent=48"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=156,.parent=49"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=157,.parent=51"
         _StyleDefs(60)  =   "Splits(1).Style:id=103,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(61)  =   "Splits(1).CaptionStyle:id=112,.parent=4"
         _StyleDefs(62)  =   "Splits(1).HeadingStyle:id=104,.parent=2"
         _StyleDefs(63)  =   "Splits(1).FooterStyle:id=105,.parent=3"
         _StyleDefs(64)  =   "Splits(1).InactiveStyle:id=106,.parent=5"
         _StyleDefs(65)  =   "Splits(1).SelectedStyle:id=108,.parent=6"
         _StyleDefs(66)  =   "Splits(1).EditorStyle:id=107,.parent=7"
         _StyleDefs(67)  =   "Splits(1).HighlightRowStyle:id=109,.parent=8,.bgcolor=&H800000&"
         _StyleDefs(68)  =   "Splits(1).EvenRowStyle:id=110,.parent=9"
         _StyleDefs(69)  =   "Splits(1).OddRowStyle:id=111,.parent=10"
         _StyleDefs(70)  =   "Splits(1).RecordSelectorStyle:id=113,.parent=11"
         _StyleDefs(71)  =   "Splits(1).FilterBarStyle:id=114,.parent=12"
         _StyleDefs(72)  =   "Splits(1).Columns(0).Style:id=118,.parent=103"
         _StyleDefs(73)  =   "Splits(1).Columns(0).HeadingStyle:id=115,.parent=104"
         _StyleDefs(74)  =   "Splits(1).Columns(0).FooterStyle:id=116,.parent=105"
         _StyleDefs(75)  =   "Splits(1).Columns(0).EditorStyle:id=117,.parent=107"
         _StyleDefs(76)  =   "Splits(1).Columns(1).Style:id=122,.parent=103,.alignment=2,.valignment=2"
         _StyleDefs(77)  =   "Splits(1).Columns(1).HeadingStyle:id=119,.parent=104"
         _StyleDefs(78)  =   "Splits(1).Columns(1).FooterStyle:id=120,.parent=105"
         _StyleDefs(79)  =   "Splits(1).Columns(1).EditorStyle:id=121,.parent=107"
         _StyleDefs(80)  =   "Splits(1).Columns(2).Style:id=166,.parent=103,.alignment=2"
         _StyleDefs(81)  =   "Splits(1).Columns(2).HeadingStyle:id=163,.parent=104"
         _StyleDefs(82)  =   "Splits(1).Columns(2).FooterStyle:id=164,.parent=105"
         _StyleDefs(83)  =   "Splits(1).Columns(2).EditorStyle:id=165,.parent=107"
         _StyleDefs(84)  =   "Splits(1).Columns(3).Style:id=174,.parent=103,.alignment=2,.bold=-1"
         _StyleDefs(85)  =   ":id=174,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(86)  =   ":id=174,.fontname=MS Sans Serif"
         _StyleDefs(87)  =   "Splits(1).Columns(3).HeadingStyle:id=171,.parent=104,.bold=0,.fontsize=825"
         _StyleDefs(88)  =   ":id=171,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(89)  =   ":id=171,.fontname=MS Sans Serif"
         _StyleDefs(90)  =   "Splits(1).Columns(3).FooterStyle:id=172,.parent=105,.bold=0,.fontsize=825"
         _StyleDefs(91)  =   ":id=172,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(92)  =   ":id=172,.fontname=MS Sans Serif"
         _StyleDefs(93)  =   "Splits(1).Columns(3).EditorStyle:id=173,.parent=107"
         _StyleDefs(94)  =   "Splits(1).Columns(4).Style:id=20,.parent=103,.alignment=2"
         _StyleDefs(95)  =   "Splits(1).Columns(4).HeadingStyle:id=17,.parent=104"
         _StyleDefs(96)  =   "Splits(1).Columns(4).FooterStyle:id=18,.parent=105"
         _StyleDefs(97)  =   "Splits(1).Columns(4).EditorStyle:id=19,.parent=107"
         _StyleDefs(98)  =   "Splits(1).Columns(5).Style:id=178,.parent=103"
         _StyleDefs(99)  =   "Splits(1).Columns(5).HeadingStyle:id=175,.parent=104"
         _StyleDefs(100) =   "Splits(1).Columns(5).FooterStyle:id=176,.parent=105"
         _StyleDefs(101) =   "Splits(1).Columns(5).EditorStyle:id=177,.parent=107"
         _StyleDefs(102) =   "Named:id=33:Normal"
         _StyleDefs(103) =   ":id=33,.parent=0"
         _StyleDefs(104) =   "Named:id=34:Heading"
         _StyleDefs(105) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(106) =   ":id=34,.wraptext=-1"
         _StyleDefs(107) =   "Named:id=35:Footing"
         _StyleDefs(108) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(109) =   "Named:id=36:Selected"
         _StyleDefs(110) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(111) =   "Named:id=37:Caption"
         _StyleDefs(112) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(113) =   "Named:id=38:HighlightRow"
         _StyleDefs(114) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(115) =   "Named:id=39:EvenRow"
         _StyleDefs(116) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(117) =   "Named:id=40:OddRow"
         _StyleDefs(118) =   ":id=40,.parent=33"
         _StyleDefs(119) =   "Named:id=41:RecordSelector"
         _StyleDefs(120) =   ":id=41,.parent=34"
         _StyleDefs(121) =   "Named:id=42:FilterBar"
         _StyleDefs(122) =   ":id=42,.parent=33"
      End
      Begin MSComctlLib.TreeView TvwEmpreendimentos 
         Height          =   3645
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3855
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
         TabIndex        =   16
         Top             =   5300
         Visible         =   0   'False
         Width           =   9540
         _ExtentX        =   16828
         _ExtentY        =   661
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
         Left            =   4080
         TabIndex        =   3
         Top             =   120
         Width           =   2415
      End
   End
   Begin MSComDlg.CommonDialog dlgSalvar 
      Left            =   4200
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   8880
      TabIndex        =   15
      Top             =   5340
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc DatEmpreendimento1 
      Height          =   330
      Left            =   240
      Top             =   5640
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
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Height          =   345
      Left            =   7680
      TabIndex        =   14
      Top             =   5340
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc DatEmpreendimento2 
      Height          =   330
      Left            =   240
      Top             =   5880
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
End
Attribute VB_Name = "FrmRelAnexoFluxoRecebiveisItau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XFO_ExlObj               As Excel.Application    'Create excel object
Public XFI_QtdEmpreendimento    As Integer              'Quantidade de empreendimentos
Dim VFV_VETORPLANOS             As New XArray           'Vetor que armazena a classificação dos planos

Private Sub Form_Activate()
    Set Formulario = FrmRelAnexoFluxoRecebiveisItau
End Sub

Private Sub Form_Load()
    Dim Sql As String
    
    Sql = "SELECT tipl_tx_Descricao, '' AS PlanoMensais, '' AS PlanoIntermediarias, '' AS PlanoChaves, '' AS PlanoRepasse, tipl_cd_TipoPlano FROM TiposPlanos WHERE tipl_tx_Sigla <> 'R'"
    subCarregaVetor Sql, Array("tipl_tx_Descricao", "PlanoMensais", "PlanoIntermediarias", "PlanoChaves", "PlanoRepasse", "tipl_cd_TipoPlano"), VFV_VETORPLANOS, TDBGrid1
    
    subConectarControleDadosNV DatEmpreendimento1, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_cd_empreendimento", Estatico
    subConectarControleDadosNV DatEmpreendimento2, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_cd_empreendimento", Estatico

    DtBase.Value = Date
    
    XFI_QtdEmpreendimento = 0
    PreencheTreeViewEmpreendimentos
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

'Carrega o treeview com os empreendimentos da empresa ativa
Private Sub PreencheTreeViewEmpreendimentos()
    Dim nodeEmpd As MSComctlLib.node
            
    TvwEmpreendimentos.Nodes.Clear
    
    If Not DatEmpreendimento1.Recordset.EOF Then
        
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
    
    End If
    
End Sub

'Preenche a árvore do nó selecionado
Private Sub CarregaRegistros(ByVal nodeNo As MSComctlLib.node)
    Dim nodeFolha As MSComctlLib.node
    
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
    FrmRelAnexoFluxoRecebiveisItau.Height = 6720
    CmdImprimir.Top = 5820
    CmdCancelar.Top = 5820
    FraPrincipal.Height = 5775
        
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
    FraPrincipal.Height = 5295
    CmdImprimir.Top = 5340
    CmdCancelar.Top = 5340
    FrmRelAnexoFluxoRecebiveisItau.Height = 6240
End Sub

'Avança unidades na barra de progresso
Private Sub AvancarProgresso(valor As Integer)
    If (pgbImpressao.Value + valor <= pgbImpressao.Max) Then
        pgbImpressao.Value = pgbImpressao.Value + valor
    End If
End Sub

'Imprime a planilha de Fluxo Recebíveis - Banco Itaú
Private Sub CmdImprimir_Click()
    Dim rs                  As New ADODB.Recordset
    Dim XLO_COMANDO         As New ADODB.Command
    Dim XLI_INDEXADOR       As Integer  'Armazena o tipo de correção selecionado pelo usuário
    Dim XLT_MENSAIS         As String   'Armazena o código dos planos classificados como Mensais
    Dim XLT_INTERMEDIARIAS  As String   'Armazena o código dos planos classificados como Intermediárias
    Dim XLT_CHAVES          As String   'Armazena o código dos planos classificados como Chaves
    Dim XLT_REPASSE         As String   'Armazena o código dos planos classificados como Repasse
    
    Dim XLT_EMPREENDIMENTOS         As String
    Dim XLO_EMPREENDIMENTO()        As String
    
    Dim i               As Integer  'Auxiliar
    Dim j               As Integer  'Auxiliar
    Dim sheet           As Integer  'Quantidade de planilhas no workbook
         
    On Error GoTo TrataErro
                
    'Pega classificação dos planos
    TDBGrid1.MoveFirst
    While Not TDBGrid1.EOF
        If TDBGrid1.Columns("Mensais").Value = "-1" Then
            XLT_MENSAIS = XLT_MENSAIS + ", " + TDBGrid1.Columns("CodigoPlano")
        End If
        
        If TDBGrid1.Columns("Intermediárias").Value = "-1" Then
            XLT_INTERMEDIARIAS = XLT_INTERMEDIARIAS + ", " + TDBGrid1.Columns("CodigoPlano")
        End If
        
        If TDBGrid1.Columns("Chaves").Value = "-1" Then
            XLT_CHAVES = XLT_CHAVES + ", " + TDBGrid1.Columns("CodigoPlano")
        End If
        
        If TDBGrid1.Columns("Repasse").Value = "-1" Then
            XLT_REPASSE = XLT_REPASSE + ", " + TDBGrid1.Columns("CodigoPlano")
        End If
        
        TDBGrid1.MoveNext
    Wend
    
    'Remove a vírgula do início
    XLT_MENSAIS = Mid(XLT_MENSAIS, 2)
    XLT_INTERMEDIARIAS = Mid(XLT_INTERMEDIARIAS, 2)
    XLT_CHAVES = Mid(XLT_CHAVES, 2)
    XLT_REPASSE = Mid(XLT_REPASSE, 2)
                
                
    'Verifica preenchimento obrigatório
    If XLT_MENSAIS = "" Or XLT_INTERMEDIARIAS = "" Or XLT_CHAVES = "" Then
        MsgBox "É necessário informar pelo menos um plano para cada classificação (mensal, intermediárias e chaves).", vbInformation, "Atenção"
        FrmRelRelacaoClientes.MousePointer = vbDefault
        Exit Sub
    End If

        
    'Caixa de diálogo para selecionar local para armazenar a planilha
    dlgSalvar.CancelError = True
                              
    'Seta o Titulo da Tela de Dialogo salvar
    dlgSalvar.DialogTitle = "Salvar Arquivo"
    
    'Seta o filtro das extensões de arquivos que será usada
    dlgSalvar.Filter = "Documento do Excel(*.xls)|*.xls"
    
    'Indica qual sera o filtro padrão
    dlgSalvar.FilterIndex = 1
                
    dlgSalvar.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
            
    'Exibe a tela do dialogo para o usuário
    dlgSalvar.ShowSave


    'Verifica se o usuário selecionou o caminho do arquivo
    If (dlgSalvar.FileName <> "") Then
    
        'Indexador
        If optvlHistorico.Value = True Then
          XLI_INDEXADOR = 0
        ElseIf optIndexador1.Value = True Then
          XLI_INDEXADOR = 1
        ElseIf optIndexador2.Value = True Then
          XLI_INDEXADOR = 2
        End If
    
        'Altera o cursor do mouse
        FrmRelAnexoFluxoRecebiveisItau.MousePointer = vbHourglass
        
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
        
        'Inicializa a barra de progresso
        InicializarProgresso 0, j * 12
                
        'Impede que o usuário veja a planilha enquanto ela está sendo preenchida
        XFO_ExlObj.Visible = False
        
        i = 1
        sheet = 1
        XLT_EMPREENDIMENTOS = ""
        
        While i < TvwEmpreendimentos.Nodes.Count
            
            If (TvwEmpreendimentos.Nodes.Item(i).Tag = "Nó") And (TvwEmpreendimentos.Nodes.Item(i).Checked = True) Then
                
                XLO_EMPREENDIMENTO = Split(TvwEmpreendimentos.Nodes.Item(i).Text, "-")
                XLT_EMPREENDIMENTOS = Trim(XLO_EMPREENDIMENTO(0))
                              
                For j = i + 1 To i + XFI_QtdEmpreendimento - 1
                    If (TvwEmpreendimentos.Nodes.Item(j).Checked = True) Then
                        XLO_EMPREENDIMENTO = Split(TvwEmpreendimentos.Nodes.Item(j).Text, "-")
                        XLT_EMPREENDIMENTOS = XLT_EMPREENDIMENTOS & ", " & Trim(XLO_EMPREENDIMENTO(0))
                    End If
                Next
                
                AvancarProgresso 1
                
                '------------------
                'Imprime relatório
                '------------------
                If (XLT_EMPREENDIMENTOS <> "") Then
                
                    With XLO_COMANDO
                       .CommandTimeout = 300
                       .CommandText = "spCAPI_sel_fluxoComercializacaoITAU"
                       .CommandType = adCmdStoredProc
                       Set XLO_COMANDO.ActiveConnection = Conexao
                       'Cria automaticamente todos os parâmetros dentro do objeto Comando
                       .Parameters.Refresh
                    End With
                   
                    'Passa os parâmetros para a procedure
                    XLO_COMANDO.Parameters(1).Value = PCodEmpresa  'Empresa
                    XLO_COMANDO.Parameters(2).Value = DtBase.Value 'Data base para o cálculo
                    XLO_COMANDO.Parameters(3).Type = adBSTR
                    XLO_COMANDO.Parameters(3).Value = XLT_EMPREENDIMENTOS 'Empreendimentos
                    XLO_COMANDO.Parameters(4).Value = XLI_INDEXADOR      'Indexador
                    XLO_COMANDO.Parameters(5).Value = XLT_MENSAIS        'Planos classificados como mensais
                    XLO_COMANDO.Parameters(6).Value = XLT_INTERMEDIARIAS 'Planos classificados como intermediárias
                    XLO_COMANDO.Parameters(7).Value = XLT_CHAVES         'Planos classificados como chaves
                    XLO_COMANDO.Parameters(8).Value = XLT_REPASSE        'Planos classificados como repasse
                    
                    'Executa a procedure e guarda o resultado no objeto rs
                    Set rs = XLO_COMANDO.Execute
                    
                    AvancarProgresso 1
                
                    'Função para criar a planilha
                    ExportaExcel XLO_EMPREENDIMENTO(1), rs, sheet
                    
                    rs.Close
                                        
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
        
        FrmRelAnexoFluxoRecebiveisItau.MousePointer = vbDefault
        MsgBox "Relatório Exportado com sucesso!", vbInformation, "Informação"
        
        FinalizarProgresso
    End If
    
Exit Sub

TrataErro:
    
    FrmRelAnexoFluxoRecebiveisItau.MousePointer = vbDefault
    
    If Err.Number = cdlCancel Then
       Exit Sub
    End If
        
    MsgBox "Ocorreu um erro na exportação do relatório. Contate o suporte. Detalhe do Erro: " & Err.Description, vbCritical, "Erro"
    DeleteFile (dlgSalvar.FileName)
    
    FinalizarProgresso
End Sub

Private Sub ExportaExcel(nomeEmpreendimento As String, dados As ADODB.Recordset, sheet As Integer)
    Dim x       As Integer 'Auxiliar
    Dim i       As Integer 'Auxiliar
    Dim cont    As Integer 'Auxiliar
     
    
    'Seleciona a próxima planilha vazia para ser preenchida
    XFO_ExlObj.Sheets(sheet).Select
    
    'Configurando margem da página
    With XFO_ExlObj.ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.64)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.54)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .Orientation = xlLandscape
        .PrintTitleRows = "$1:$12" 'Repete o cabeçalho em todas as páginas
    End With
    
    AvancarProgresso 1
    
    'Preenchendo as células
    With XFO_ExlObj.ActiveSheet
        .Shapes.AddPicture "C:\QualiAdmFin\QualiCapi\Relatorios\Imagens\imgFluxoRecebiveisItau.png", True, True, 49, 1, 245, 55
        
        'Seta o nome da aba da planilha
        .Name = IIf(Len(Trim(nomeEmpreendimento)) > 31, Mid(Trim(nomeEmpreendimento), 1, 31), Trim(nomeEmpreendimento))
        
        .Cells(6, 6).Value = " UNIDADES COMERCIALIZADAS DO EMPREENDIMENTO":
        .Cells(6, 6).Font.Size = 9
        .Cells(7, 6).Value = "Edificio: "
        .Cells(7, 7).Value = nomeEmpreendimento
        .Cells(7, 7).Font.Size = 8
        .Cells(7, 9).Value = "    Nº contrato:"
        .Cells(7, 10).Value = txtNumContrato.Text
        .Cells(8, 1).Value = "Empresário: " & PEmpresa
        .Cells(8, 11).Value = "Data:  " & DtBase.Value
        .Cells(10, 1).Value = "N°"
        .Cells(10, 4).Value = "ENTRADA"
        .Range("E10:I10").Merge
        .Cells(10, 5).Value = "VALORES A RECEBER"
        .Cells(11, 1).Value = "DA UNIDADE"
        .Cells(11, 2).Value = "NOME DO COMPRADOR"
        .Cells(11, 3).Value = "VALOR DA VENDA"
        .Cells(11, 4).Value = "RECEBIDO ATE ESTA DATA (1)"
        .Range("E11:G11").Merge
        .Cells(11, 5).Value = "CONSTRUCAO (2)"
        .Range("H11:J11").Merge
        .Cells(11, 8).Value = "FINANCIAMENTO PROPRIO (3)"
        .Cells(11, 11).Value = "FINANCIAMENTO      <<            >>"
        .Cells(12, 5).Value = "MENSAIS"
        .Cells(12, 6).Value = "INTERMEDIÁRIAS"
        .Cells(12, 7).Value = "CHAVES"
        .Cells(12, 8).Value = "MENSAIS"
        .Range("I12:J12").Merge
        .Range("J10:K10").Merge
        .Cells(12, 9).Value = "INTERMEDIÁRIAS"
        .Cells(12, 11).Value = "-4"
        .Cells(12, 11).Font.Bold = True
    End With
    
    AvancarProgresso 1
    
    XFO_ExlObj.ActiveSheet.Range("A6:k11").Font.Bold = True
    XFO_ExlObj.ActiveSheet.Range("A10:k11").HorizontalAlignment = xlHAlignCenter
    
               
    With XFO_ExlObj.ActiveSheet.Range("A10:k12")
       .Interior.ColorIndex = 15
       .HorizontalAlignment = xlHAlignCenter
       .VerticalAlignment = xlVAlignTop
       .Font.Size = 7.5
    End With
    
    'Formata a largura da coluna
    With XFO_ExlObj.ActiveSheet
        .Range("A10").ColumnWidth = 8.29
        .Range("B10").ColumnWidth = 20.86
        .Range("C10").ColumnWidth = 11.57
        .Range("D10").ColumnWidth = 11.86
        .Range("E10").ColumnWidth = 10
        .Range("F10").ColumnWidth = 12
        .Range("G10").ColumnWidth = 10.14
        .Range("H10").ColumnWidth = 9.29
        .Range("I10").ColumnWidth = 14.86
        .Range("J10").ColumnWidth = 5.71
        .Range("K10").ColumnWidth = 16.29
    End With

    With XFO_ExlObj.ActiveSheet.Range("E12:k11")
        .Font.Size = 7.5
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    
    For x = 12 To 41 + 14 'Altera o n° 38 para quantidade de registro da consulta
         XFO_ExlObj.ActiveSheet.Range("I" & x & ":J" & x).Merge
    Next

    AvancarProgresso 1
    
    i = 13
    cont = 0
    
    'Insere os dados
    While Not dados.EOF
        With XFO_ExlObj.ActiveSheet
            .Cells(i, 1) = dados!Imovel
            .Cells(i, 2) = dados!Comprador
            .Cells(i, 3) = dados!ValorVenda
            .Cells(i, 4) = dados!ValorRecebido
            .Cells(i, 5) = dados!MensaisConstrucao
            .Cells(i, 6) = dados!IntermediariasConstrucao
            .Cells(i, 7) = dados!Chave
            .Cells(i, 8) = dados!MensaisFinanciamento
            .Cells(i, 9) = dados!IntermediariasFinanciamento
            .Cells(i, 11) = dados!Financiamento
        End With
        dados.MoveNext
        i = i + 1
        cont = cont + 1
    Wend
    
    
    AvancarProgresso 3
    
    cont = cont + 14
    
    'Formata o formato dos números
    XFO_ExlObj.ActiveSheet.Range("C13:K" & cont).NumberFormat = "#,##0.00"
    
    'Insere borda na planilha
    With XFO_ExlObj.ActiveSheet.Range("A10:K" & cont).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.ActiveSheet.Range("A10:K" & cont).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
        
    With XFO_ExlObj.ActiveSheet.Range("A10:K" & cont).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
        
    With XFO_ExlObj.ActiveSheet.Range("A" & cont & ":K" & cont).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
        
    With XFO_ExlObj.ActiveSheet.Range("A10:A" & cont).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
         
    With XFO_ExlObj.ActiveSheet.Range("K10:K" & cont).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
    
    AvancarProgresso 1
    
    'Formata a fonte e o tamanho da fonte
    With XFO_ExlObj.ActiveSheet.Range("A13:K" & cont)
        .Font.Size = 8
        .Font.Name = "Times New Roman"
    End With
    
    With XFO_ExlObj.ActiveSheet.Range("B13:B" & cont)
        .Font.Size = 7
        .Font.Name = "Times New Roman"
    End With
    
    'Formata o alinhamento do conteúdo
    With XFO_ExlObj.ActiveSheet.Range("A13:K" & cont)
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignTop
    End With
    
    'Formata a largura da linha:
    XFO_ExlObj.ActiveSheet.Range("A13:A" & cont).RowHeight = 22.5
    XFO_ExlObj.ActiveSheet.Range("A11").RowHeight = 24
    
    'Hachura os campos "SubTotal" e "Total Geral"
    XFO_ExlObj.ActiveSheet.Range("A" & cont - 1 & ":B" & cont).Select
    With XFO_ExlObj.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlGray16
        .PatternColorIndex = xlAutomatic
    End With
    
    AvancarProgresso 1
    
    With XFO_ExlObj.ActiveSheet
        'Fórmulas para calcular os totais
        .Range("A" & cont - 1 & ":B" & cont - 1).Merge
        .Range("A" & cont & ":B" & cont).Merge
        .Range("C" & cont - 1 & ":K" & cont).Font.Bold = True
        .Cells(cont - 1, 1).Value = "SUB TOTAL"
        .Cells(cont, 1).Value = "TOTAL GERAL"
        .Cells(cont - 1, 1).HorizontalAlignment = xlHAlignCenter
        .Cells(cont, 1).HorizontalAlignment = xlHAlignCenter
        .Cells(cont - 1, 4).Formula = "=SUM(D13:D" & cont - 2 & ")"
        .Cells(cont - 1, 5).Formula = "=SUM(E13:E" & cont - 2 & ")"
        .Cells(cont - 1, 6).Formula = "=SUM(F13:F" & cont - 2 & ")"
        .Cells(cont - 1, 7).Formula = "=SUM(G13:G" & cont - 2 & ")"
        .Cells(cont - 1, 8).Formula = "=SUM(H13:H" & cont - 2 & ")"
        .Cells(cont - 1, 9).Formula = "=SUM(I13:I" & cont - 2 & ")"
        .Cells(cont - 1, 11).Formula = "=SUM(K13:K" & cont - 2 & ")"
        .Range("C" & cont - 1 & ":K" & cont).Font.Name = "Arial"
        
        .Range("A10:k" & cont).WrapText = True
        
        cont = cont + 3
        .Range("A" & cont & ":K" & cont + 14).Font.Size = 8
        .Range("A" & cont & ":K" & cont + 14).Font.Name = "Arial"
        
        .Cells(cont, 1).Value = "INSTRUÇÕES PARA PREENCHIMENTO CONFORME ABAIXO:"
        .Cells(cont, 1).Font.Bold = True
        cont = cont + 1
        .Cells(cont, 1).Value = "- valores devem ser preenchidos em reais"
        cont = cont + 1
        .Cells(cont, 1).Value = "- nº e o tipo de cada unidade vendida, à vista, a prazo e permutada"
        cont = cont + 1
        .Cells(cont, 1).Value = "- nome completo do comprador para cada unidade vendida e permutada"
        .Cells(cont, 8).Value = "Até Chaves  (soma 1 + 2):"
        .Range("H" & cont).Font.Bold = True
        .Cells(cont, 11).Formula = "=SUM(D" & cont - 7 & ":G" & cont - 7 & ")"
        .Cells(cont, 11).Font.Size = 10
        .Range("K" & cont).Font.Bold = True
        cont = cont + 1
        .Cells(cont, 1).Value = "- valor de venda total de cada unidade vendida, permutada "
        .Cells(cont, 8).Value = "Pós Chave Soma (3 + 4)"
        .Range("H" & cont).Font.Bold = True
        .Cells(cont, 11).Formula = "=SUM(H" & cont - 8 & ":I" & cont - 8 & ")"
        .Cells(cont, 11).Font.Size = 10
        .Range("K" & cont).Font.Bold = True
        
        cont = cont + 1
        .Cells(cont, 1).Value = "- valor total a ser pago durante a fase de construção de cada unidade vendida "
        cont = cont + 1
        .Cells(cont, 1).Value = "- valor a ser pago nas chaves de cada unidade vendida "
        cont = cont + 1
        .Cells(cont, 1).Value = "- valor total do financiamento direto com a construtora"
        cont = cont + 1
        .Cells(cont, 1).Value = "- valor total a ser financiado pelo agente financeiro"
        
        cont = cont + 2
        .Range("H" & cont & ":K" & cont).Merge
        .Range("H" & cont & ":K" & cont).HorizontalAlignment = xlHAlignCenter
        .Cells(cont, 8).Font.Size = 10
        .Cells(cont, 8).Value = "Assinatura do responsável pela empresa"
        
        cont = cont + 3
        .Cells(cont, 1).Value = "IMPORTANTE: -  a)Anexar o(s) respectivo(s) Compromisso(s) de Venda e Compra, do empreendimento objeto de empréstimo."
        
        cont = cont + 1
        .Cells(cont, 1).Value = "b)    Na eventualidade de haver (em) outro(s) empreendimento (s),  financiado (s) pelo Banco Itaú S/A, relacionar neste impresso."
    End With
    
    AvancarProgresso 1
    
    'Insere borda no topo para assinatura
    With XFO_ExlObj.ActiveSheet.Range("H" & cont - 4 & ":K" & cont - 4).Borders(xlEdgeTop)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With

    'Coloca a palavra "IMPORTANTE" em negrito
    XFO_ExlObj.ActiveSheet.Range("A" & cont - 1).Select
    XFO_ExlObj.ActiveCell.Characters(1, 10).Font.Bold = True
      
    'Formata para não exibir as linhas de grade da planilha
    XFO_ExlObj.ActiveWindow.DisplayGridlines = False
    
    AvancarProgresso 1
End Sub



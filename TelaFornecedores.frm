VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaFornecedores 
   Caption         =   "Tabela de Fornecedores"
   ClientHeight    =   8535
   ClientLeft      =   1215
   ClientTop       =   1845
   ClientWidth     =   12330
   Icon            =   "TelaFornecedores.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   12330
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanMostraImóveis 
      Height          =   3525
      Left            =   960
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   11175
      _Version        =   65536
      _ExtentX        =   19711
      _ExtentY        =   6218
      _StockProps     =   15
      Caption         =   "Imóveis"
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
      Begin VB.CommandButton CmdRetornar2 
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
         Height          =   270
         Index           =   0
         Left            =   9720
         TabIndex        =   23
         Top             =   3120
         Width           =   975
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Bindings        =   "TelaFornecedores.frx":0442
         Height          =   2565
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   4524
         _LayoutType     =   4
         _RowHeight      =   25
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código do Empreendimento"
         Columns(0).DataField=   "empd_cd_empreendimento"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nome do Empreendimento"
         Columns(1).DataField=   "empd_tx_Nome"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Imóvel"
         Columns(2).DataField=   "imov_cd_imovel"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Contrato"
         Columns(3).DataField=   "cont_cd_contrato"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Data do Contrato"
         Columns(4).DataField=   "cont_dt_RegistroContrato"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Status do Contrato"
         Columns(5).DataField=   "Status"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).Locked=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColMove=   -1  'True
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3175"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3096"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8720"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(0)._MinWidth=64"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=4789"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=4710"
         Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=8720"
         Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(16)=   "Column(1)._MinWidth=-1"
         Splits(0)._ColumnProps(17)=   "Column(2).Width=2910"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2831"
         Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=8720"
         Splits(0)._ColumnProps(22)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(24)=   "Column(3).Width=2355"
         Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=2275"
         Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=8720"
         Splits(0)._ColumnProps(29)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(31)=   "Column(4).Width=2381"
         Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=2302"
         Splits(0)._ColumnProps(34)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=8720"
         Splits(0)._ColumnProps(36)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(37)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(38)=   "Column(4)._MinWidth=56550368"
         Splits(0)._ColumnProps(39)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(40)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(42)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(43)=   "Column(5)._ColStyle=532"
         Splits(0)._ColumnProps(44)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(45)=   "Column(5).Order=6"
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
         DirectionAfterEnter=   0
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.locked=-1,.bold=0"
         _StyleDefs(37)  =   ":id=28,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(38)  =   ":id=28,.fontname=MS Sans Serif"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15,.alignment=3"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0,.locked=-1"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0,.locked=-1"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14,.alignment=2"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0,.locked=-1"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14,.alignment=2"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=0,.locked=-1"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14,.alignment=2"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(62)  =   "Named:id=33:Normal"
         _StyleDefs(63)  =   ":id=33,.parent=0"
         _StyleDefs(64)  =   "Named:id=34:Heading"
         _StyleDefs(65)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H8000000F&"
         _StyleDefs(66)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1"
         _StyleDefs(67)  =   "Named:id=35:Footing"
         _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   "Named:id=36:Selected"
         _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(71)  =   "Named:id=37:Caption"
         _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(73)  =   "Named:id=38:HighlightRow"
         _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(75)  =   "Named:id=39:EvenRow"
         _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(77)  =   "Named:id=40:OddRow"
         _StyleDefs(78)  =   ":id=40,.parent=33"
         _StyleDefs(79)  =   "Named:id=41:RecordSelector"
         _StyleDefs(80)  =   ":id=41,.parent=34"
         _StyleDefs(81)  =   "Named:id=42:FilterBar"
         _StyleDefs(82)  =   ":id=42,.parent=33"
      End
   End
   Begin MSComDlg.CommonDialog DlgSalvar1 
      Left            =   7200
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSPanel PanAlterarCnpj 
      Height          =   1890
      Left            =   2880
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   3510
      _Version        =   65536
      _ExtentX        =   6191
      _ExtentY        =   3334
      _StockProps     =   15
      Caption         =   "Alteração de CNPJ"
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
         Height          =   270
         Left            =   840
         TabIndex        =   14
         Top             =   1440
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
         Height          =   270
         Index           =   0
         Left            =   2100
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin MSMask.MaskEdBox MskCnpjCpf 
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   960
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LblOriginal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "CNPJ Original:"
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
         Left            =   240
         TabIndex        =   18
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label LblCnpjCpf 
         BackColor       =   &H00E0E0E0&
         Caption         =   "999.999.999-99"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1620
         TabIndex        =   17
         Top             =   540
         Width           =   1755
      End
      Begin VB.Label LblNovo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "CNPJ Novo:"
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
         Left            =   420
         TabIndex        =   16
         Top             =   1020
         Width           =   1095
      End
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaFornecedores.frx":0457
      Height          =   315
      Left            =   3000
      TabIndex        =   8
      Top             =   3375
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "focl_tx_fantasia"
      BoundColumn     =   "focl_cd_forncli"
      Text            =   ""
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaFornecedores.frx":046C
      Height          =   3285
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   5794
      _LayoutType     =   4
      _RowHeight      =   25
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Fantasia"
      Columns(0).DataField=   "focl_tx_fantasia"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   80
      Columns(1)._MaxComboItems=   5
      Columns(1).ValueItems(0)._DefaultItem=   0
      Columns(1).ValueItems(0).Value=   "J"
      Columns(1).ValueItems(0).Value.vt=   8
      Columns(1).ValueItems(0).DisplayValue=   "S"
      Columns(1).ValueItems(0).DisplayValue.vt=   8
      Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(1)._DefaultItem=   0
      Columns(1).ValueItems(1).Value=   "3"
      Columns(1).ValueItems(1).Value.vt=   8
      Columns(1).ValueItems(1).DisplayValue=   "P"
      Columns(1).ValueItems(1).DisplayValue.vt=   8
      Columns(1).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(2)._DefaultItem=   0
      Columns(1).ValueItems(2).Value=   "4"
      Columns(1).ValueItems(2).Value.vt=   8
      Columns(1).ValueItems(2).DisplayValue=   "PS"
      Columns(1).ValueItems(2).DisplayValue.vt=   8
      Columns(1).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems.Count=   3
      Columns(1).Caption=   "Tipo"
      Columns(1).DataField=   "focl_tx_tipo"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Razão Social"
      Columns(2).DataField=   "focl_tx_razaosocial"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "CGC"
      Columns(3).DataField=   "focl_tx_cgccpf"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Telefone"
      Columns(4).DataField=   "focl_tx_fone"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Fax"
      Columns(5).DataField=   "focl_tx_fax"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Endereço"
      Columns(6).DataField=   "focl_tx_endereco"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Nº Porta"
      Columns(7).DataField=   "focl_nr_porta"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Bairro"
      Columns(8).DataField=   "focl_tx_bairro"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Cidade"
      Columns(9).DataField=   "focl_tx_cidade"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Estado"
      Columns(10).DataField=   "focl_tx_estado"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "CEP"
      Columns(11).DataField=   "focl_tx_cep"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "EMail"
      Columns(12).DataField=   "focl_tx_email"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Inscrição Estadual"
      Columns(13).DataField=   "focl_tx_iestadual"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Contato Comercial"
      Columns(14).DataField=   "focl_tx_contatocom"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Contato Financeiro"
      Columns(15).DataField=   "focl_tx_contatofin"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Conta Despesa"
      Columns(16).DataField=   "contarecdesp"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Conta Contábil"
      Columns(17).DataField=   "contacusto"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "Conta Adiantamento"
      Columns(18).DataField=   "contaprovadiant"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "Centro Custo"
      Columns(19).DataField=   "cere_tx_descricao"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   16
      Columns(20)._MaxComboItems=   5
      Columns(20).ValueItems(0)._DefaultItem=   0
      Columns(20).ValueItems(0).Value=   "0"
      Columns(20).ValueItems(0).Value.vt=   8
      Columns(20).ValueItems(0).DisplayValue=   "Cheque"
      Columns(20).ValueItems(0).DisplayValue.vt=   8
      Columns(20).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(20).ValueItems(1)._DefaultItem=   0
      Columns(20).ValueItems(1).Value=   "1"
      Columns(20).ValueItems(1).Value.vt=   8
      Columns(20).ValueItems(1).DisplayValue=   "Borderô 1"
      Columns(20).ValueItems(1).DisplayValue.vt=   8
      Columns(20).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(20).ValueItems(2)._DefaultItem=   0
      Columns(20).ValueItems(2).Value=   "2"
      Columns(20).ValueItems(2).Value.vt=   8
      Columns(20).ValueItems(2).DisplayValue=   "Borderô 2"
      Columns(20).ValueItems(2).DisplayValue.vt=   8
      Columns(20).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(20).ValueItems(3)._DefaultItem=   0
      Columns(20).ValueItems(3).Value=   "3"
      Columns(20).ValueItems(3).Value.vt=   8
      Columns(20).ValueItems(3).DisplayValue=   "Débito em Conta"
      Columns(20).ValueItems(3).DisplayValue.vt=   8
      Columns(20).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(20).ValueItems(4)._DefaultItem=   0
      Columns(20).ValueItems(4).Value=   "4"
      Columns(20).ValueItems(4).Value.vt=   8
      Columns(20).ValueItems(4).DisplayValue=   "Crédito em Conta"
      Columns(20).ValueItems(4).DisplayValue.vt=   8
      Columns(20).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
      Columns(20).ValueItems(5)._DefaultItem=   0
      Columns(20).ValueItems(5).Value=   "5"
      Columns(20).ValueItems(5).Value.vt=   8
      Columns(20).ValueItems(5).DisplayValue=   "Dinheiro"
      Columns(20).ValueItems(5).DisplayValue.vt=   8
      Columns(20).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
      Columns(20).ValueItems.Count=   6
      Columns(20).Caption=   "Tipo Pagamento"
      Columns(20).DataField=   "focl_tx_tipopag"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "Código"
      Columns(21).DataField=   "focl_cd_forncli"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "Qualificação"
      Columns(22).DataField=   "Qualificacoes"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "Contrato"
      Columns(23).DataField=   "CONTRATO"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "Código Atividade"
      Columns(24).DataField=   "focl_nr_codatividade"
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "Super Simples"
      Columns(25).DataField=   "focl_tx_SuperSimples"
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "Principal"
      Columns(26).DataField=   "focl_tx_fornecedorPrincipal"
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "Crítico"
      Columns(27).DataField=   "focl_tx_fornecedorcritico"
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "Última Qualificação"
      Columns(28).DataField=   "focl_dt_pesquisaUltimaQualificacao"
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   29
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).Locked=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=29"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3175"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3096"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8720"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(0)._MinWidth=64"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=1164"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=1085"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=2105873"
      Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(17)=   "Column(1)._MinWidth=-1"
      Splits(0)._ColumnProps(18)=   "Column(2).Width=4789"
      Splits(0)._ColumnProps(19)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._WidthInPix=4710"
      Splits(0)._ColumnProps(21)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(2)._ColStyle=2105872"
      Splits(0)._ColumnProps(23)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(25)=   "Column(2)._MinWidth=-1"
      Splits(0)._ColumnProps(26)=   "Column(3).Width=2910"
      Splits(0)._ColumnProps(27)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._WidthInPix=2831"
      Splits(0)._ColumnProps(29)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(3)._ColStyle=2105872"
      Splits(0)._ColumnProps(31)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(32)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(33)=   "Column(4).Width=2355"
      Splits(0)._ColumnProps(34)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(4)._WidthInPix=2275"
      Splits(0)._ColumnProps(36)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(37)=   "Column(4)._ColStyle=2105872"
      Splits(0)._ColumnProps(38)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(39)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(40)=   "Column(5).Width=2381"
      Splits(0)._ColumnProps(41)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(5)._WidthInPix=2302"
      Splits(0)._ColumnProps(43)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(44)=   "Column(5)._ColStyle=2105872"
      Splits(0)._ColumnProps(45)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(46)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(47)=   "Column(5)._MinWidth=56550368"
      Splits(0)._ColumnProps(48)=   "Column(6).Width=4551"
      Splits(0)._ColumnProps(49)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(6)._WidthInPix=4471"
      Splits(0)._ColumnProps(51)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(52)=   "Column(6)._ColStyle=2105872"
      Splits(0)._ColumnProps(53)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(54)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(55)=   "Column(6)._MinWidth=56549904"
      Splits(0)._ColumnProps(56)=   "Column(7).Width=1482"
      Splits(0)._ColumnProps(57)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(7)._WidthInPix=1402"
      Splits(0)._ColumnProps(59)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(60)=   "Column(7)._ColStyle=2105872"
      Splits(0)._ColumnProps(61)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(62)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(63)=   "Column(8).Width=2593"
      Splits(0)._ColumnProps(64)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(8)._WidthInPix=2514"
      Splits(0)._ColumnProps(66)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(67)=   "Column(8)._ColStyle=2097680"
      Splits(0)._ColumnProps(68)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(69)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(70)=   "Column(9).Width=2434"
      Splits(0)._ColumnProps(71)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(9)._WidthInPix=2355"
      Splits(0)._ColumnProps(73)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(74)=   "Column(9)._ColStyle=2097680"
      Splits(0)._ColumnProps(75)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(76)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(77)=   "Column(10).Width=1217"
      Splits(0)._ColumnProps(78)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(10)._WidthInPix=1138"
      Splits(0)._ColumnProps(80)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(81)=   "Column(10)._ColStyle=2097681"
      Splits(0)._ColumnProps(82)=   "Column(10).WrapText=1"
      Splits(0)._ColumnProps(83)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(84)=   "Column(11).Width=1720"
      Splits(0)._ColumnProps(85)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(11)._WidthInPix=1640"
      Splits(0)._ColumnProps(87)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(88)=   "Column(11)._ColStyle=2097680"
      Splits(0)._ColumnProps(89)=   "Column(11).WrapText=1"
      Splits(0)._ColumnProps(90)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(91)=   "Column(12).Width=4419"
      Splits(0)._ColumnProps(92)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(93)=   "Column(12)._WidthInPix=4339"
      Splits(0)._ColumnProps(94)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(95)=   "Column(12)._ColStyle=2097680"
      Splits(0)._ColumnProps(96)=   "Column(12).WrapText=1"
      Splits(0)._ColumnProps(97)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(98)=   "Column(13).Width=2355"
      Splits(0)._ColumnProps(99)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(100)=   "Column(13)._WidthInPix=2275"
      Splits(0)._ColumnProps(101)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(102)=   "Column(13)._ColStyle=2097680"
      Splits(0)._ColumnProps(103)=   "Column(13).WrapText=1"
      Splits(0)._ColumnProps(104)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(105)=   "Column(14).Width=3519"
      Splits(0)._ColumnProps(106)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(14)._WidthInPix=3440"
      Splits(0)._ColumnProps(108)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(109)=   "Column(14)._ColStyle=2097680"
      Splits(0)._ColumnProps(110)=   "Column(14).WrapText=1"
      Splits(0)._ColumnProps(111)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(112)=   "Column(15).Width=3493"
      Splits(0)._ColumnProps(113)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(114)=   "Column(15)._WidthInPix=3413"
      Splits(0)._ColumnProps(115)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(116)=   "Column(15)._ColStyle=2097680"
      Splits(0)._ColumnProps(117)=   "Column(15).WrapText=1"
      Splits(0)._ColumnProps(118)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(119)=   "Column(16).Width=4948"
      Splits(0)._ColumnProps(120)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(121)=   "Column(16)._WidthInPix=4868"
      Splits(0)._ColumnProps(122)=   "Column(16)._EditAlways=0"
      Splits(0)._ColumnProps(123)=   "Column(16)._ColStyle=2097680"
      Splits(0)._ColumnProps(124)=   "Column(16).WrapText=1"
      Splits(0)._ColumnProps(125)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(126)=   "Column(17).Width=4921"
      Splits(0)._ColumnProps(127)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(128)=   "Column(17)._WidthInPix=4842"
      Splits(0)._ColumnProps(129)=   "Column(17)._EditAlways=0"
      Splits(0)._ColumnProps(130)=   "Column(17)._ColStyle=2097680"
      Splits(0)._ColumnProps(131)=   "Column(17).WrapText=1"
      Splits(0)._ColumnProps(132)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(133)=   "Column(18).Width=4895"
      Splits(0)._ColumnProps(134)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(135)=   "Column(18)._WidthInPix=4815"
      Splits(0)._ColumnProps(136)=   "Column(18)._EditAlways=0"
      Splits(0)._ColumnProps(137)=   "Column(18)._ColStyle=2097680"
      Splits(0)._ColumnProps(138)=   "Column(18).WrapText=1"
      Splits(0)._ColumnProps(139)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(140)=   "Column(19).Width=6535"
      Splits(0)._ColumnProps(141)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(142)=   "Column(19)._WidthInPix=6456"
      Splits(0)._ColumnProps(143)=   "Column(19)._EditAlways=0"
      Splits(0)._ColumnProps(144)=   "Column(19)._ColStyle=2097680"
      Splits(0)._ColumnProps(145)=   "Column(19).WrapText=1"
      Splits(0)._ColumnProps(146)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(147)=   "Column(20).Width=2090"
      Splits(0)._ColumnProps(148)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(149)=   "Column(20)._WidthInPix=2011"
      Splits(0)._ColumnProps(150)=   "Column(20)._EditAlways=0"
      Splits(0)._ColumnProps(151)=   "Column(20)._ColStyle=2097680"
      Splits(0)._ColumnProps(152)=   "Column(20).WrapText=1"
      Splits(0)._ColumnProps(153)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(154)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(155)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(156)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(157)=   "Column(21)._EditAlways=0"
      Splits(0)._ColumnProps(158)=   "Column(21).AllowSizing=0"
      Splits(0)._ColumnProps(159)=   "Column(21)._ColStyle=2097680"
      Splits(0)._ColumnProps(160)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(161)=   "Column(21).WrapText=1"
      Splits(0)._ColumnProps(162)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(163)=   "Column(22).Width=2461"
      Splits(0)._ColumnProps(164)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(165)=   "Column(22)._WidthInPix=2381"
      Splits(0)._ColumnProps(166)=   "Column(22)._EditAlways=0"
      Splits(0)._ColumnProps(167)=   "Column(22)._ColStyle=2097681"
      Splits(0)._ColumnProps(168)=   "Column(22).WrapText=1"
      Splits(0)._ColumnProps(169)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(170)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(171)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(172)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(173)=   "Column(23)._EditAlways=0"
      Splits(0)._ColumnProps(174)=   "Column(23)._ColStyle=2097681"
      Splits(0)._ColumnProps(175)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(176)=   "Column(23).WrapText=1"
      Splits(0)._ColumnProps(177)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(178)=   "Column(24).Width=873"
      Splits(0)._ColumnProps(179)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(180)=   "Column(24)._WidthInPix=794"
      Splits(0)._ColumnProps(181)=   "Column(24)._EditAlways=0"
      Splits(0)._ColumnProps(182)=   "Column(24)._ColStyle=2097680"
      Splits(0)._ColumnProps(183)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(184)=   "Column(24).WrapText=1"
      Splits(0)._ColumnProps(185)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(186)=   "Column(25).Width=2646"
      Splits(0)._ColumnProps(187)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(188)=   "Column(25)._WidthInPix=2566"
      Splits(0)._ColumnProps(189)=   "Column(25)._EditAlways=0"
      Splits(0)._ColumnProps(190)=   "Column(25)._ColStyle=2097681"
      Splits(0)._ColumnProps(191)=   "Column(25).WrapText=1"
      Splits(0)._ColumnProps(192)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(193)=   "Column(26).Width=1508"
      Splits(0)._ColumnProps(194)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(195)=   "Column(26)._WidthInPix=1429"
      Splits(0)._ColumnProps(196)=   "Column(26)._EditAlways=0"
      Splits(0)._ColumnProps(197)=   "Column(26)._ColStyle=2097681"
      Splits(0)._ColumnProps(198)=   "Column(26).WrapText=1"
      Splits(0)._ColumnProps(199)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(200)=   "Column(27).Width=1773"
      Splits(0)._ColumnProps(201)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(202)=   "Column(27)._WidthInPix=1693"
      Splits(0)._ColumnProps(203)=   "Column(27)._EditAlways=0"
      Splits(0)._ColumnProps(204)=   "Column(27)._ColStyle=2105873"
      Splits(0)._ColumnProps(205)=   "Column(27).WrapText=1"
      Splits(0)._ColumnProps(206)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(207)=   "Column(28).Width=714"
      Splits(0)._ColumnProps(208)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(209)=   "Column(28)._WidthInPix=635"
      Splits(0)._ColumnProps(210)=   "Column(28)._EditAlways=0"
      Splits(0)._ColumnProps(211)=   "Column(28)._ColStyle=2097681"
      Splits(0)._ColumnProps(212)=   "Column(28).WrapText=1"
      Splits(0)._ColumnProps(213)=   "Column(28).Order=29"
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
      DirectionAfterEnter=   0
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.locked=-1,.bold=0"
      _StyleDefs(37)  =   ":id=28,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(38)  =   ":id=28,.fontname=MS Sans Serif"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15,.alignment=3"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=126,.parent=13,.alignment=2,.locked=-1,.bold=0"
      _StyleDefs(43)  =   ":id=126,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(44)  =   ":id=126,.fontname=MS Sans Serif"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=123,.parent=14,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=124,.parent=15,.bold=0,.fontsize=825"
      _StyleDefs(47)  =   ":id=124,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(48)  =   ":id=124,.fontname=MS Sans Serif"
      _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=125,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(59)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(63)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(64)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(67)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(68)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(7).Style:id=114,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(71)  =   "Splits(0).Columns(7).HeadingStyle:id=111,.parent=14,.alignment=2"
      _StyleDefs(72)  =   "Splits(0).Columns(7).FooterStyle:id=112,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(7).EditorStyle:id=113,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=0"
      _StyleDefs(75)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(76)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=0"
      _StyleDefs(79)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14,.alignment=2"
      _StyleDefs(80)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(10).Style:id=70,.parent=13,.alignment=2"
      _StyleDefs(83)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(84)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(11).Style:id=74,.parent=13,.alignment=0"
      _StyleDefs(87)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=14,.alignment=2"
      _StyleDefs(88)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(12).Style:id=78,.parent=13,.alignment=0"
      _StyleDefs(91)  =   "Splits(0).Columns(12).HeadingStyle:id=75,.parent=14,.alignment=2"
      _StyleDefs(92)  =   "Splits(0).Columns(12).FooterStyle:id=76,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(12).EditorStyle:id=77,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(13).Style:id=82,.parent=13,.alignment=0"
      _StyleDefs(95)  =   "Splits(0).Columns(13).HeadingStyle:id=79,.parent=14,.alignment=2"
      _StyleDefs(96)  =   "Splits(0).Columns(13).FooterStyle:id=80,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(13).EditorStyle:id=81,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(14).Style:id=86,.parent=13,.alignment=0"
      _StyleDefs(99)  =   "Splits(0).Columns(14).HeadingStyle:id=83,.parent=14,.alignment=2"
      _StyleDefs(100) =   "Splits(0).Columns(14).FooterStyle:id=84,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(14).EditorStyle:id=85,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(15).Style:id=90,.parent=13,.alignment=0"
      _StyleDefs(103) =   "Splits(0).Columns(15).HeadingStyle:id=87,.parent=14,.alignment=2"
      _StyleDefs(104) =   "Splits(0).Columns(15).FooterStyle:id=88,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(15).EditorStyle:id=89,.parent=17"
      _StyleDefs(106) =   "Splits(0).Columns(16).Style:id=94,.parent=13,.alignment=0"
      _StyleDefs(107) =   "Splits(0).Columns(16).HeadingStyle:id=91,.parent=14,.alignment=2"
      _StyleDefs(108) =   "Splits(0).Columns(16).FooterStyle:id=92,.parent=15"
      _StyleDefs(109) =   "Splits(0).Columns(16).EditorStyle:id=93,.parent=17"
      _StyleDefs(110) =   "Splits(0).Columns(17).Style:id=98,.parent=13,.alignment=0"
      _StyleDefs(111) =   "Splits(0).Columns(17).HeadingStyle:id=95,.parent=14,.alignment=2"
      _StyleDefs(112) =   "Splits(0).Columns(17).FooterStyle:id=96,.parent=15"
      _StyleDefs(113) =   "Splits(0).Columns(17).EditorStyle:id=97,.parent=17"
      _StyleDefs(114) =   "Splits(0).Columns(18).Style:id=102,.parent=13,.alignment=0"
      _StyleDefs(115) =   "Splits(0).Columns(18).HeadingStyle:id=99,.parent=14,.alignment=2"
      _StyleDefs(116) =   "Splits(0).Columns(18).FooterStyle:id=100,.parent=15"
      _StyleDefs(117) =   "Splits(0).Columns(18).EditorStyle:id=101,.parent=17"
      _StyleDefs(118) =   "Splits(0).Columns(19).Style:id=106,.parent=13,.alignment=0"
      _StyleDefs(119) =   "Splits(0).Columns(19).HeadingStyle:id=103,.parent=14,.alignment=2"
      _StyleDefs(120) =   "Splits(0).Columns(19).FooterStyle:id=104,.parent=15"
      _StyleDefs(121) =   "Splits(0).Columns(19).EditorStyle:id=105,.parent=17"
      _StyleDefs(122) =   "Splits(0).Columns(20).Style:id=110,.parent=13,.alignment=0"
      _StyleDefs(123) =   "Splits(0).Columns(20).HeadingStyle:id=107,.parent=14,.alignment=2"
      _StyleDefs(124) =   "Splits(0).Columns(20).FooterStyle:id=108,.parent=15"
      _StyleDefs(125) =   "Splits(0).Columns(20).EditorStyle:id=109,.parent=17"
      _StyleDefs(126) =   "Splits(0).Columns(21).Style:id=118,.parent=13,.alignment=0"
      _StyleDefs(127) =   "Splits(0).Columns(21).HeadingStyle:id=115,.parent=14,.alignment=2"
      _StyleDefs(128) =   "Splits(0).Columns(21).FooterStyle:id=116,.parent=15"
      _StyleDefs(129) =   "Splits(0).Columns(21).EditorStyle:id=117,.parent=17"
      _StyleDefs(130) =   "Splits(0).Columns(22).Style:id=130,.parent=13"
      _StyleDefs(131) =   "Splits(0).Columns(22).HeadingStyle:id=127,.parent=14"
      _StyleDefs(132) =   "Splits(0).Columns(22).FooterStyle:id=128,.parent=15"
      _StyleDefs(133) =   "Splits(0).Columns(22).EditorStyle:id=129,.parent=17"
      _StyleDefs(134) =   "Splits(0).Columns(23).Style:id=134,.parent=13"
      _StyleDefs(135) =   "Splits(0).Columns(23).HeadingStyle:id=131,.parent=14"
      _StyleDefs(136) =   "Splits(0).Columns(23).FooterStyle:id=132,.parent=15"
      _StyleDefs(137) =   "Splits(0).Columns(23).EditorStyle:id=133,.parent=17"
      _StyleDefs(138) =   "Splits(0).Columns(24).Style:id=122,.parent=13,.alignment=0"
      _StyleDefs(139) =   "Splits(0).Columns(24).HeadingStyle:id=119,.parent=14,.alignment=2"
      _StyleDefs(140) =   "Splits(0).Columns(24).FooterStyle:id=120,.parent=15"
      _StyleDefs(141) =   "Splits(0).Columns(24).EditorStyle:id=121,.parent=17"
      _StyleDefs(142) =   "Splits(0).Columns(25).Style:id=150,.parent=13,.alignment=2"
      _StyleDefs(143) =   "Splits(0).Columns(25).HeadingStyle:id=147,.parent=14"
      _StyleDefs(144) =   "Splits(0).Columns(25).FooterStyle:id=148,.parent=15"
      _StyleDefs(145) =   "Splits(0).Columns(25).EditorStyle:id=149,.parent=17"
      _StyleDefs(146) =   "Splits(0).Columns(26).Style:id=138,.parent=13,.alignment=2"
      _StyleDefs(147) =   "Splits(0).Columns(26).HeadingStyle:id=135,.parent=14"
      _StyleDefs(148) =   "Splits(0).Columns(26).FooterStyle:id=136,.parent=15"
      _StyleDefs(149) =   "Splits(0).Columns(26).EditorStyle:id=137,.parent=17"
      _StyleDefs(150) =   "Splits(0).Columns(27).Style:id=142,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(151) =   "Splits(0).Columns(27).HeadingStyle:id=139,.parent=14"
      _StyleDefs(152) =   "Splits(0).Columns(27).FooterStyle:id=140,.parent=15"
      _StyleDefs(153) =   "Splits(0).Columns(27).EditorStyle:id=141,.parent=17"
      _StyleDefs(154) =   "Splits(0).Columns(28).Style:id=146,.parent=13,.alignment=2"
      _StyleDefs(155) =   "Splits(0).Columns(28).HeadingStyle:id=143,.parent=14"
      _StyleDefs(156) =   "Splits(0).Columns(28).FooterStyle:id=144,.parent=15"
      _StyleDefs(157) =   "Splits(0).Columns(28).EditorStyle:id=145,.parent=17"
      _StyleDefs(158) =   "Named:id=33:Normal"
      _StyleDefs(159) =   ":id=33,.parent=0,.alignment=2"
      _StyleDefs(160) =   "Named:id=34:Heading"
      _StyleDefs(161) =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(162) =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1"
      _StyleDefs(163) =   "Named:id=35:Footing"
      _StyleDefs(164) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(165) =   "Named:id=36:Selected"
      _StyleDefs(166) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(167) =   "Named:id=37:Caption"
      _StyleDefs(168) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(169) =   "Named:id=38:HighlightRow"
      _StyleDefs(170) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(171) =   "Named:id=39:EvenRow"
      _StyleDefs(172) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(173) =   "Named:id=40:OddRow"
      _StyleDefs(174) =   ":id=40,.parent=33"
      _StyleDefs(175) =   "Named:id=41:RecordSelector"
      _StyleDefs(176) =   ":id=41,.parent=34"
      _StyleDefs(177) =   "Named:id=42:FilterBar"
      _StyleDefs(178) =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   795
      Left            =   480
      TabIndex        =   0
      Top             =   3780
      Width           =   10560
      _Version        =   65536
      _ExtentX        =   18627
      _ExtentY        =   1402
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
      Begin VB.CommandButton cmdExpDominio 
         Caption         =   "&Exp.Dominio"
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
         Left            =   8040
         TabIndex        =   26
         Top             =   15
         Width           =   1320
      End
      Begin VB.CheckBox ChkListarFornColaborador 
         Caption         =   "Ocultar Colaboradores"
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
         Left            =   8160
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton CmdRamosAtividades 
         Caption         =   "&Ramos Ativ."
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
         Left            =   5280
         TabIndex        =   24
         Top             =   15
         Width           =   1320
      End
      Begin VB.CommandButton cmdExport 
         Height          =   330
         Left            =   480
         Picture         =   "TelaFornecedores.frx":0481
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton CmdImoveis 
         Caption         =   "I&móveis"
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
         Left            =   6660
         TabIndex        =   19
         Top             =   15
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CommandButton CmdAlterarCnpj 
         Caption         =   "&Alterar CNPJ"
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
         Left            =   5280
         TabIndex        =   11
         Top             =   420
         Width           =   1320
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
         Left            =   6660
         TabIndex        =   7
         Top             =   420
         Width           =   1320
      End
      Begin VB.CommandButton Cmdimprimir 
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
         Left            =   3900
         TabIndex        =   6
         Top             =   420
         Width           =   1320
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
         Left            =   2520
         TabIndex        =   5
         Top             =   420
         Width           =   1320
      End
      Begin VB.CommandButton Botao_Fil 
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
         Left            =   1140
         TabIndex        =   4
         Top             =   420
         Width           =   1320
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
         Left            =   3900
         TabIndex        =   3
         Top             =   15
         Width           =   1320
      End
      Begin VB.CommandButton CmdAlterar 
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
         Left            =   2520
         TabIndex        =   2
         Top             =   15
         Width           =   1320
      End
      Begin VB.CommandButton CmdInserir 
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
         Left            =   1140
         TabIndex        =   1
         Top             =   15
         Width           =   1320
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5355
      Top             =   4800
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8880
      Top             =   3480
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
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
      Caption         =   "Adodc2"
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
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Fantasia:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3420
      Width           =   2670
   End
End
Attribute VB_Name = "TelaFornecedores"
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
Dim XLD_TOTALID As String         'Total IDs
Dim FRM_SQLAtivo As String        'Atualizar o SQL Ativo após modificações de filtro e ordem
Sub TelaValoresGlobais(GuardaOUPegaValores As String)
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
        xLocalCampo = Ordem
        xLocalDescricao = xLocalDescricao
        
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
Private Sub HabilitaBotoes()

    'Objetivo: Habilita os Botões da Tela
    CmdInserir.Enabled = True
    If Not FunTabelaVazia(Adodc1) Then
        CmdAlterar.Enabled = True
        CmdExcluir.Enabled = True
        CmdImprimir.Enabled = True
        CmdOrdem.Enabled = True
        cboLocalizar.Enabled = True
    Else
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdImprimir.Enabled = False
        CmdOrdem.Enabled = False
        cboLocalizar.Enabled = False
    End If
    
    CmdSair.Enabled = True
    'Verifica se tem Filtro Ativo
    If ContFil = 0 Then
       Botao_Fil.Enabled = False
    Else
       Botao_Fil.Enabled = True
    End If
End Sub
Private Sub DesabilitaBotoes()
    'Objetivo: Desabilita os Botões da Tela
    CmdInserir.Enabled = False
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    'Botao_Fil.Enabled = False
    CmdSair.Enabled = False
    CmdImprimir.Enabled = False
    CmdOrdem.Enabled = False
End Sub
Private Sub Botao_Fil_Click()
    Dim Fil As Integer
    DesabilitaBotoes
    Botao_Fil.Enabled = False
    Set Formulario = TelaFornecedores
    TelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    TelaValoresGlobais "G"

    If Filtrou = True Then
        Fil = ContFil
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        ContFil = Fil
    End If
    
    HabilitaBotoes
    
    TDBGrid1.SetFocus
End Sub

Private Sub cboLocalizar_Change()
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
End Sub

Private Sub ChkListarFornColaborador_Click()
    CARREGAGRID
End Sub

Private Sub CARREGAGRID()
    If ChkListarFornColaborador.Value = 1 Then
        subCarregaDadosNV Adodc1, "", "focl_tx_fantasia", "", "SELECT focl_cd_FornCli, focl_tx_Fantasia, focl_tx_RazaoSocial, focl_tx_CgcCpf, focl_tx_Fone, focl_tx_Fax, focl_tx_Endereco, focl_nr_Porta, focl_tx_Bairro, focl_tx_Cidade, focl_tx_Estado, focl_tx_Cep, focl_tx_EMail, contarecdesp, contaprovadiant, cere_tx_Descricao, focl_tx_TipoPag, Qualificacoes, focl_tx_SuperSimples, focl_tx_fornecedorcritico, focl_dt_pesquisaUltimaQualificacao FROM ConsGENFornCli LEFT JOIN vetorh.dbo.r034fun ON dbo.fcn_putzero(r034fun.numcpf,11) = dbo.OnlyNumbers(ConsGENFornCli.focl_tx_CgcCpf) WHERE focl_tx_tipo = '" & Pessoa & "' and focl_tx_classe='" & FornClie & "' and (empr_cd_empresa = " & PCodEmpresa & " or (empr_cd_empresa is null)) AND r034fun.numcpf IS NULL "
    Else
        subCarregaDadosNV Adodc1, "", "focl_tx_fantasia", "", "SELECT * FROM ConsGENFornCli WHERE focl_tx_tipo = '" & Pessoa & "' and focl_tx_classe='" & FornClie & "' and (empr_cd_empresa = " & PCodEmpresa & " or (empr_cd_empresa is null))"
    End If
End Sub

Private Sub CmdAlterar_Click()
    If FornClie = "F" Then
        If Not FunVerifica_Permissao("TelaFornecedores", "4") Then
            Exit Sub
        End If
    Else
        If Not FunVerifica_Permissao("TelaFornecedores", "7") Then
            Exit Sub
        End If
    End If
    
    Chave = Adodc1.Recordset.Fields("focl_cd_forncli")
'    ChaveBanco = Adodc1.Recordset.Fields("banc_cd_codigo")
    XInserir = "A"
    xChave = Adodc1.Recordset.bookmark
    DesabilitaBotoes
    FrmFornecedores.Show 1
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    cboLocalizar.BoundText = Chave
    If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    HabilitaBotoes
End Sub

Private Sub CmdAlterarCnpj_Click()
  
  PanBotoes.Enabled = False
  PanAlterarCnpj.Left = (TDBGrid1.Width - PanAlterarCnpj.Width) / 2
  PanAlterarCnpj.Top = (TDBGrid1.Height - PanAlterarCnpj.Height) / 2
  TDBGrid1.Enabled = False
  PanAlterarCnpj.Visible = True
  LblCnpjCpf.Caption = Adodc1.Recordset!focl_tx_CGCCPF
  MskCnpjCpf.PromptInclude = False
  MskCnpjCpf.Text = ""
  MskCnpjCpf.PromptInclude = True
  
  If Pessoa = "F" Then
    PanAlterarCnpj.Caption = "Alteração de CPF"
    LblOriginal.Caption = "CPF Original:"
    LblNovo.Caption = "CPF Novo:"
    MskCnpjCpf.Mask = "###.###.###-##"
    MskCnpjCpf.Text = "   .   .   -  "
  Else
    PanAlterarCnpj.Caption = "Alteração de CNPJ"
    LblOriginal.Caption = "CNPJ Original:"
    LblNovo.Caption = "CNPJ Novo:"
    MskCnpjCpf.Mask = "##.###.###/####-##"
    MskCnpjCpf.Text = "  .   .   /    -  "
  End If
  
  MskCnpjCpf.SetFocus
  
End Sub

Private Sub CmdConfirmar_Click()
  On Error GoTo TrataErro:
  
  Dim XLO_CODIGO As New ADODB.Recordset
  Dim XLT_SQL As String
   
  MskCnpjCpf.PromptInclude = False
  If Pessoa = "F" Then
    If Not FunConfereCPF(MskCnpjCpf.Text) Then
       MsgBox "CPF Invalido!", vbCritical, "Atenção!"
       MskCnpjCpf.SetFocus
       MskCnpjCpf.PromptInclude = True
       Exit Sub
     End If
  Else
    If Not FunConfereCGC(MskCnpjCpf.Text) Then
       MsgBox "CNPJ Invalido!", vbCritical, "Atenção!"
       MskCnpjCpf.SetFocus
       MskCnpjCpf.PromptInclude = True
       Exit Sub
     End If
  End If
  
  MskCnpjCpf.PromptInclude = True
  Conexao.BeginTrans
  
  XLT_SQL = "SELECT focl_cd_FornCli FROM FornClientes WHERE" & _
    " focl_tx_CgcCpf = '" & MskCnpjCpf.Text & "'" & _
    " AND focl_tx_Tipo='" & Pessoa & "'" & _
    " AND focl_tx_Classe='" & FornClie & "'"
    
  SubQOpenRecordset XLO_CODIGO, XLT_SQL, Estatico
  
  If Not XLO_CODIGO.EOF Then
    If Pessoa = "F" Then
      MsgBox "Este CPF já existe!", vbCritical, "ATENÇÃO"
    Else
      MsgBox "Este CNPJ já existe!", vbCritical, "ATENÇÃO"
    End If
    MskCnpjCpf.SetFocus
    Conexao.RollbackTrans
    Exit Sub
  Else
    XLO_CODIGO.Close
    Set XLO_CODIGO = Nothing
    
    Conexao.Execute "UPDATE FornClientes SET focl_tx_CgcCpf='" & MskCnpjCpf.Text & "'" & _
      " WHERE focl_tx_CgcCpf='" & Adodc1.Recordset!focl_tx_CGCCPF & "'"
  End If
  
  Conexao.CommitTrans
  Adodc1.Recordset.Requery adCmdText
  
  'Call CmdRetornar_Click
  
  Exit Sub
TrataErro:
    Conexao.RollbackTrans
    funTrataErros (ComMensagem)

End Sub

Private Sub CmdExcluir_Click()
    Dim opcao As Integer
    Dim Houve_Erro As Integer
    Dim ResReg As ADODB.Recordset
    Dim Posicao As Variant
    
    If FornClie = "F" Then
        If Not FunVerifica_Permissao("TelaFornecedores", "5") Then
            Exit Sub
        End If
    Else
        If Not FunVerifica_Permissao("TelaFornecedores", "8") Then
            Exit Sub
        End If
    End If
    
    Houve_Erro = False
    
    Posicao = Adodc1.Recordset.bookmark
        
    DesabilitaBotoes
   
    opcao = MsgBox("Confirma Remoção do Registro ?", 273, "CUIDADO")
   
    If opcao = 1 Then  'Confirma exclusão
        On Error GoTo RotuloErro
        'QDelete Data1
        Conexao.BeginTrans
      
        SubQOpenRecordset ResReg, "SELECT * FROM NotasFiscais WHERE focl_cd_forncli =" & Adodc1.Recordset.Fields("focl_cd_forncli") & " and empr_cd_empresa =" & PCodEmpresa, Estatico
        If ResReg.EOF Then
           ResReg.Close
           'Verificar se esse Fornecedor existe para outra empresa
           SubQOpenRecordset ResReg, "SELECT * FROM PlcCustoForncli WHERE focl_cd_forncli =" & Adodc1.Recordset.Fields("focl_cd_forncli") & " and empr_cd_empresa <>" & PCodEmpresa, Estatico
           If Not ResReg.EOF Then
                Dim XContaRecDesp As String, XContaCusto As String
                Dim XContaProvAdiant As String, XPcr As String
    
                'Se exitir apaga os registros da tabela de relacionamento PlCCustoFornCli, mas não apaga o registo da tabela FornClientes
                'Conexao.Execute "DELETE FROM PlCCustoFornCli WHERE focl_cd_forncli =" & Adodc1.Recordset.Fields("focl_cd_forncli") & " and empr_cd_empresa =" & PCodEmpresa

                XContaRecDesp = ""
                XContaCusto = ""
                XContaProvAdiant = ""
                XPcr = ""

                Sql = "UPDATE PlccustoForncli " & _
                      "SET plco_cd_ContaRecDesp = " & FunNuloBanco(XContaRecDesp) & ", " & _
                         " plco_cd_ContaCusto = " & FunNuloBanco(XContaCusto) & "," & _
                         " plco_cd_contaprovadiant = " & FunNuloBanco(XContaProvAdiant) & "," & _
                         " cere_cd_pcr = " & FunNuloBanco(XPcr) & "" & _
                         " WHERE focl_cd_forncli = " & Adodc1.Recordset.Fields("focl_cd_forncli") & " and empr_cd_empresa =" & PCodEmpresa
                Conexao.Execute Sql
   
                If FornClie = "F" Then
                    'MsgBox "Apenas as Contas Contábeis e o Centro de Custo foram excluídos, pois o Fornecedor está sendo utilizado por outra empresa.", vbCritical + vbOKOnly, "ATENÇÃO"
                    opcao = MsgBox("Apenas as Contas Contábeis e o Centro de Custo foram excluídos, pois o Fornecedor está sendo utilizado por outra empresa. Deseja excluir, permanentemente, o Fornecedor ?", 273, "CUIDADO")
                Else
                    'MsgBox "Apenas as Contas Contábeis e o Centro de Custo foram excluídos, pois o Cliente está sendo utilizado por outra empresa.", vbCritical + vbOKOnly, "ATENÇÃO"
                    opcao = MsgBox("Apenas as Contas Contábeis e o Centro de Custo foram excluídos, pois o Cliente está sendo utilizado por outra empresa. Deseja excluir, permanentemente, o Cliente ?", 273, "CUIDADO")
                End If
                If opcao = 1 Then  'Confirma exclusão
                    Conexao.Execute "DELETE FROM PlCCustoFornCli WHERE focl_cd_forncli =" & Adodc1.Recordset.Fields("focl_cd_forncli")
                    Conexao.Execute "DELETE FROM FornClientes Where focl_cd_forncli=" & Adodc1.Recordset.Fields("focl_cd_forncli")
                    Conexao.Execute "DELETE FROM FornCliente_RamosAtividades WHERE focl_cd_forncli = " & Adodc1.Recordset.Fields("focl_cd_forncli")
                End If
            Else
                'Se não existir apaga os registros da tabela de relacionamento PlCCustoFornCli e apaga o registo da tabela FornClientes
                Conexao.Execute "DELETE FROM PlCCustoFornCli WHERE focl_cd_forncli =" & Adodc1.Recordset.Fields("focl_cd_forncli")
                Conexao.Execute "DELETE FROM FornClientes Where focl_cd_forncli=" & Adodc1.Recordset.Fields("focl_cd_forncli")
                Conexao.Execute "DELETE FROM FornCliente_RamosAtividades WHERE focl_cd_forncli = " & Adodc1.Recordset.Fields("focl_cd_forncli")
            End If
            ResReg.Close
            Set ResReg = Nothing
      Else
            If FornClie = "F" Then
                MsgBox "O Fornecedor possui Notas Fiscais cadastradas. Não é possível excluí-lo.", vbCritical + vbOKOnly, "ATENÇÃO"
            Else
                MsgBox "O Cliente possui Notas Fiscais cadastradas. Não é possível excluí-lo.", vbCritical + vbOKOnly, "ATENÇÃO"
            End If
            ResReg.Close
            Set ResReg = Nothing
      End If

      Conexao.CommitTrans
      If Houve_Erro = False Then
         FRM_Quantid = funOperacaoQuantid("-")
         subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
      End If
    End If
    
    cboLocalizar.Text = ""
    TDBGrid1.SetFocus
    Adodc1.Refresh
    HabilitaBotoes
    TDBGrid1.Refresh
    TDBGrid1.Col = 0
        
    'Posicionando o cursor na linha anterior - Lucas Santiago 22.01.09
    If Posicao > 1 Then
        Adodc1.Recordset.bookmark = Posicao - 1
    End If
    
    Exit Sub
    
RotuloErro:
    Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next
End Sub

Private Sub cmdExpDominio_Click()
    'TelaExpFornecedoresDominio.Show
End Sub

Private Sub cmdExport_Click()
    Dim sLine As String
    Dim tmpFile As String
    Dim outFile As String
    Dim i As Integer
    
    ' Sets the Dialog Title to Save File
    CommonDialog1.DialogTitle = "Salvar arquivo"

    ' Sets the File List box to Text File and All Files
    CommonDialog1.Filter = "Arquivo CSV (*.csv)|*.csv"

    ' Set the default files type to Text File
    CommonDialog1.FilterIndex = 1

    ' Sets the flags - Hide Read only, prompt to overwrite, and path must exist
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist

    ' Set dialog box so an error occurs if the dialogbox is cancelled
    CommonDialog1.CancelError = True

    ' Enables error handling to catch cancel error
    On Error Resume Next
    ' display the dialog box
    CommonDialog1.ShowSave
    If Err Then
        ' This code runs if the dialog was cancelled
        Exit Sub
    End If
    
    'Call your export function
    outFile = CommonDialog1.FileName
    tmpFile = outFile & ".tmp"
    
    TDBGrid1.ExportToDelimitedFile tmpFile, , ";"
 
    Close tmpFile
    
    'Now use this to add the Headers (to a new file)
    Open tmpFile For Input As #1
    Open outFile For Output As #2
 
    'Write Headers
    For i = 0 To TDBGrid1.Columns.Count - 1
         If i = 0 And TDBGrid1.Columns(i).Visible Then
             sLine = Chr(34) & TDBGrid1.Columns(i).Caption & Chr(34)
         ElseIf TDBGrid1.Columns(i).Visible Then
             sLine = sLine & ";" & Chr(34) & TDBGrid1.Columns(i).Caption & Chr(34)
         End If
    Next
 
    Print #2, sLine

     'Write contents of original file
    Do Until EOF(1)
         Line Input #1, sLine
         Print #2, sLine
    Loop
 
    Close #1
    Close #2
 
    Kill tmpFile
          
    Screen.MousePointer = vbDefault
    
    MsgBox "Total exportado: " & Adodc1.Recordset.RecordCount & " registros", vbInformation, "Informação"

End Sub

Private Sub CmdImoveis_Click()
PanBotoes.Enabled = False
  PanMostraImóveis.Left = (TDBGrid1.Width - PanMostraImóveis.Width) / 2
  PanMostraImóveis.Top = (TDBGrid1.Height - PanMostraImóveis.Height) / 2
  TDBGrid1.Enabled = False
  PanMostraImóveis.Visible = True
  Chave = Adodc1.Recordset.Fields("focl_cd_FornCli")
  
subCarregaDadosNV Adodc2, "", "focl_tx_fantasia", "", "select CASE WHEN dbo.Contratos.cont_tx_Status  =  'A' THEN 'ATIVO' WHEN dbo.Contratos.cont_tx_Status  =  'I' THEN 'INATIVO' WHEN dbo.Contratos.cont_tx_Status  =  'R' THEN 'RESERVADO' WHEN dbo.Contratos.cont_tx_Status  =  'C' THEN 'CEDIDO' END AS Status ,* from FornClientes inner join Contratos on FornClientes.focl_cd_FornCli = Contratos .focl_cd_FornCli inner join Empreendimentos on Contratos.empd_cd_Empreendimento = Empreendimentos.empd_cd_Empreendimento where FornClientes.focl_cd_FornCli =" & Chave
        
    'xChave = Adodc1.Recordset.Bookmark
    'DesabilitaBotoes
    
    'TelaManImoveisPorCliente.Show 1
    'subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    'cboLocalizar.BoundText = Chave
    'If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    'HabilitaBotoes
End Sub

Private Sub CmdImprimir_Click()
    DesabilitaBotoes
    If FornClie = "F" Then
        If Pessoa = "F" Then
            Call subImprimeListagemGRID(2, TDBGrid1.PrintInfo, "Listagem de Fornecedores (Pessoa Física)")
        ElseIf Pessoa = "J" Then
            Call subImprimeListagemGRID(2, TDBGrid1.PrintInfo, "Listagem de Fornecedores de Serviços (Pessoa Jurídica)")
        ElseIf Pessoa = "3" Then
            Call subImprimeListagemGRID(2, TDBGrid1.PrintInfo, "Listagem de Fornecedores de Produtos (Pessoa Jurídica)")
        ElseIf Pessoa = "4" Then
            Call subImprimeListagemGRID(2, TDBGrid1.PrintInfo, "Listagem de Fornecedores de Produtos e Serviços (Pessoa Jurídica)")
        End If
    Else
        If Pessoa = "F" Then
            Call subImprimeListagemGRID(2, TDBGrid1.PrintInfo, "Listagem de Clientes (Pessoa Física)")
        Else
            Call subImprimeListagemGRID(2, TDBGrid1.PrintInfo, "Listagem de Clientes (Pessoa Jurídica)")
        End If
    End If
    HabilitaBotoes
End Sub
Private Sub CmdInserir_Click()
    If FornClie = "F" Then
        If Not FunVerifica_Permissao("TelaFornecedores", "3") Then
            Exit Sub
        End If
    Else
        If Not FunVerifica_Permissao("TelaFornecedores", "6") Then
            Exit Sub
        End If
    End If
    
    Chave = -1
    XInserir = "I"
    cboLocalizar.Text = ""
    DesabilitaBotoes
    FrmFornecedores.Show 1
    If Chave <> -1 Then
        FRM_Quantid = funOperacaoQuantid("+")
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        cboLocalizar.BoundText = Chave
        If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
    HabilitaBotoes
End Sub
Private Sub CmdOrdem_Click()
    
    DesabilitaBotoes
    Set Formulario = TelaFornecedores
    TelaValoresGlobais "P"
    TelaOrdem.Show 1
    TelaValoresGlobais "G"
    HabilitaBotoes
    
    lblLocalizar.Caption = "Localizar " & xLocalDescricao & ":"
    If InStr(1, xLocalCampo, ",") > 0 Then
        cboLocalizar.ListField = Left(xLocalCampo, InStr(1, xLocalCampo, ",") - 1)
    Else
        cboLocalizar.ListField = xLocalCampo
    End If
    cboLocalizar.Text = ""
    
    TDBGrid1.SetFocus
    
End Sub

Private Sub CmdRamosAtividades_Click()
    TelaFornecedores.MousePointer = vbHourglass
    TelaFornecedoresRamosAtividades.Show
    TelaFornecedores.MousePointer = vbDefault
End Sub

Private Sub CmdRetornar_Click(Index As Integer)
PanBotoes.Enabled = True
TDBGrid1.Enabled = True
PanAlterarCnpj.Visible = False
End Sub

Private Sub CmdRetornar2_Click(Index As Integer)
PanBotoes.Enabled = True
PanMostraImóveis.Visible = False
TDBGrid1.Enabled = True
End Sub

  
Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    MdiPrincipal.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub TDBGrid1_DblClick()
    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
'    If TDBGrid1.Columns("CONTRATO").CellText(Bookmark) <> "" Then
'        RowStyle.ForeColor = vbBlue
'    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    cboLocalizar.Text = ""
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaFornecedores
        TelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, x
        TelaValoresGlobais "G"
    End If
    
    HabilitaBotoes
    
End Sub
Private Sub Form_Activate()
    Adodc1.Refresh
    Set Formulario = TelaFornecedores
    XFormForn = "TelaFornecedores"
    If Not PrimeiraVez Then
        TelaValoresGlobais "P"
    End If
    PrimeiraVez = False
    
    If XGT_NomeAplic = "CAPI" Then
        'CmdInserir.Left = 90
        'CmdAlterar.Left = 1167
        'CmdExcluir.Left = 2244
        'CmdImoveis.Left = 3321
        'Botao_Fil.Left = 4398
        'CmdOrdem.Left = 5475
        'Cmdimprimir.Left = 6552
        'CmdSair.Left = 7635
        
       ' CmdOrdem.Left = 1545
       ' CmdImprimir.Left = 3030
       ' CmdImoveis.Left = 4515
        
        CmdImoveis.Visible = True
    End If
    MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
End Sub
Private Sub Form_Load()
    Dim XLT_SQL As String
    Dim ClienteImoveis As ADODB.Recordset
    Dim Negrito As Boolean
    Set Formulario = TelaFornecedores
    Call subAjustaTelaLoc
    
    subManutencaoJanelasAtivas "I", "TelaFornecedores"
    
    DoEvents
    
    
    xLocalDescricao = "Fantasia"
    
    'Mostrar Fornecedores de Propoduto e Produtos/Serviços
    If FornClie = "F" Then 'Fornecedor
        If Pessoa = "F" Then 'Pessoa Física
            TelaFornecedores.Caption = "Tabela Fornecedores (Serviços Pessoa Física)"
            CmdAlterarCnpj.Caption = "Alterar CPF"
            TDBGrid1.Columns(0).Caption = "Nome"
            TDBGrid1.Columns(2).Caption = "Nome Completo"
            TDBGrid1.Columns(3).Caption = "CPF"
            TDBGrid1.Columns(13).Visible = False
            TDBGrid1.Columns(14).Visible = False
            TDBGrid1.Columns(15).Visible = False
            TDBGrid1.Columns(17).Visible = False
            TDBGrid1.Columns(26).Visible = False
            ChkListarFornColaborador.Visible = True
            CARREGAGRID
            
            
              DoEvents
              
            '*************************************************************************************
            'XLT_SQL = "SELECT *, " & _
                    "(SELECT plco_cd_ContaRecDesp FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & ") AS plco_cd_ContaRecDesp, " & _
                    "(SELECT plco_cd_ContaCusto FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & ") AS plco_cd_ContaCusto, " & _
                    "(SELECT plco_cd_ContaProvAdiant FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & ") AS plco_cd_ContaProvAdiant, " & _
                    "(SELECT cere_cd_Pcr FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & ") AS cere_cd_Pcr, " & _
                    "(SELECT PlanoContas.plco_cd_CodReduzido FROM PlanoContas RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON PlanoContas.plco_cd_Conta = PlCCustoFornCli.plco_cd_ContaRecDesp WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS CodRedRecDesp, " & _
                    "(SELECT  PlanoContas.plco_tx_DescContabil FROM PlanoContas RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON PlanoContas.plco_cd_Conta = PlCCustoFornCli.plco_cd_ContaRecDesp WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS contarecdesp, " & _
                    "(SELECT  PlanoContas_1.plco_cd_CodReduzido FROM PlanoContas PlanoContas_1 RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON PlCCustoFornCli.plco_cd_ContaCusto = PlanoContas_1.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS CodRedCusto, " & _
                    "(SELECT  PlanoContas_1.plco_tx_DescContabil FROM PlanoContas PlanoContas_1 RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON PlCCustoFornCli.plco_cd_ContaCusto = PlanoContas_1.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS contacusto, " & _
                    "(SELECT  PlanoContas_2.plco_cd_CodReduzido FROM PlanoContas PlanoContas_2 RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON PlCCustoFornCli.plco_cd_ContaProvAdiant = PlanoContas_2.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS CodRedProvAdiant, " & _
                    "(SELECT  PlanoContas_2.plco_tx_DescContabil FROM PlanoContas PlanoContas_2 RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON PlCCustoFornCli.plco_cd_ContaProvAdiant = PlanoContas_2.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS contaprovadiant, " & _
                    "(SELECT  dbo.CentroResultados.cere_tx_Descricao FROM CentroResultados RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON CentroResultados.cere_cd_Pcr = PlCCustoFornCli.cere_cd_Pcr WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS cere_tx_Descricao, " & _
                    "(SELECT  dbo.CentroResultados.cere_cd_Estruturado FROM dbo.CentroResultados RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON CentroResultados.cere_cd_Pcr = PlCCustoFornCli.cere_cd_Pcr WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS cere_cd_Estruturado" & _
                    " FROM fornclientes WHERE focl_tx_tipo = '" & Pessoa & "' and focl_tx_classe='" & FornClie & "'"
            '*************************************************************************************
            'subCarregaDadosNV Adodc1, "", "focl_tx_fantasia", "", XLT_SQL

        ElseIf Pessoa = "J" Or Pessoa = "3" Or Pessoa = "4" Then 'Pessoa Jurídica
            subCarregaDadosNV Adodc1, "", "focl_tx_fantasia", "", "SELECT * FROM ConsGENFornCli WHERE focl_tx_tipo <> 'F' and focl_tx_classe='" & FornClie & "' and (empr_cd_empresa = " & PCodEmpresa & " or (empr_cd_empresa is null))"
            
            DoEvents
            
            '*************************************************************************************
            'XLT_SQL = "SELECT *, " & _
                    "(SELECT plco_cd_ContaRecDesp FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & ") AS plco_cd_ContaRecDesp, " & _
                    "(SELECT plco_cd_ContaCusto FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & ") AS plco_cd_ContaCusto, " & _
                    "(SELECT plco_cd_ContaProvAdiant FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & ") AS plco_cd_ContaProvAdiant, " & _
                    "(SELECT cere_cd_Pcr FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & ") AS cere_cd_Pcr, " & _
                    "(SELECT PlanoContas.plco_cd_CodReduzido FROM PlanoContas RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON PlanoContas.plco_cd_Conta = PlCCustoFornCli.plco_cd_ContaRecDesp WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS CodRedRecDesp, " & _
                    "(SELECT  PlanoContas.plco_tx_DescContabil FROM PlanoContas RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON PlanoContas.plco_cd_Conta = PlCCustoFornCli.plco_cd_ContaRecDesp WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS contarecdesp, " & _
                    "(SELECT  PlanoContas_1.plco_cd_CodReduzido FROM PlanoContas PlanoContas_1 RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON PlCCustoFornCli.plco_cd_ContaCusto = PlanoContas_1.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS CodRedCusto, " & _
                    "(SELECT  PlanoContas_1.plco_tx_DescContabil FROM PlanoContas PlanoContas_1 RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON PlCCustoFornCli.plco_cd_ContaCusto = PlanoContas_1.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS contacusto, " & _
                    "(SELECT  PlanoContas_2.plco_cd_CodReduzido FROM PlanoContas PlanoContas_2 RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON PlCCustoFornCli.plco_cd_ContaProvAdiant = PlanoContas_2.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS CodRedProvAdiant, " & _
                    "(SELECT  PlanoContas_2.plco_tx_DescContabil FROM PlanoContas PlanoContas_2 RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON PlCCustoFornCli.plco_cd_ContaProvAdiant = PlanoContas_2.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS contaprovadiant, " & _
                    "(SELECT  dbo.CentroResultados.cere_tx_Descricao FROM CentroResultados RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON CentroResultados.cere_cd_Pcr = PlCCustoFornCli.cere_cd_Pcr WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") AS cere_tx_Descricao, " & _
                    "(SELECT  dbo.CentroResultados.cere_cd_Estruturado FROM dbo.CentroResultados RIGHT OUTER JOIN " & _
                    "PlCCustoFornCli ON CentroResultados.cere_cd_Pcr = PlCCustoFornCli.cere_cd_Pcr WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ")  AS cere_cd_Estruturado" & _
                    " FROM fornclientes WHERE focl_tx_tipo <> 'F' and focl_tx_classe='" & FornClie & "'"
            '*************************************************************************************
            'subCarregaDadosNV Adodc1, "", "focl_tx_fantasia", "", XLT_SQL
            
            CmdAlterarCnpj.Caption = "Alterar CNPJ"
            TDBGrid1.Columns(0).Caption = "Fantasia"
            TDBGrid1.Columns(1).Visible = True
            TDBGrid1.Columns(2).Visible = True
            TDBGrid1.Columns(2).Caption = "Razão Social"
            TDBGrid1.Columns(3).Caption = "CGC"
            TDBGrid1.Columns(13).Visible = True
            TDBGrid1.Columns(14).Visible = True
            TDBGrid1.Columns(15).Visible = True
            TDBGrid1.Columns(17).Visible = True
            TDBGrid1.Columns(25).Visible = True         'Henrique Lima - 16/04/09
            TDBGrid1.Columns(26).Visible = True         'Henrique Lima - 16/04/09
            If Pessoa = "3" Then 'Produtos
                TelaFornecedores.Caption = "Tabela Fornecedores (Produtos)"
            ElseIf Pessoa = "4" Then 'Produtos e Serviços
                TelaFornecedores.Caption = "Tabela Fornecedores (Produtos e Serviços)"
            Else
                TelaFornecedores.Caption = "Tabela Fornecedores (Serviços Pessoa Jurídica)"
            End If
        End If
        If XGT_NomeAplic = "QUALIPAT" Then 'Está na aplicação do QUALIPAT
            TDBGrid1.Columns(16).Visible = False
            TDBGrid1.Columns(17).Visible = False
            TDBGrid1.Columns(18).Visible = False
            TDBGrid1.Columns(19).Visible = False
            TDBGrid1.Columns(20).Visible = False
            TDBGrid1.Columns(21).Visible = False
            TDBGrid1.Columns(22).Visible = False
        End If
        
    
    ElseIf FornClie = "C" Then 'Cliente
    
    
        subCarregaDadosNV Adodc1, "", "focl_tx_fantasia", "", "SELECT * FROM ConsGENFornCli WHERE focl_tx_tipo = '" & Pessoa & "' and focl_tx_classe='" & FornClie & "' and (empr_cd_empresa = " & PCodEmpresa & " or (empr_cd_empresa is null))"
        

        
        DoEvents
      
        
                    '*************************************************************************************
        'XLT_SQL = "SELECT *, " & _
                "(SELECT plco_cd_ContaRecDesp FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & "), " & _
                "(SELECT plco_cd_ContaCusto FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & "), " & _
                "(SELECT plco_cd_ContaProvAdiant FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & "), " & _
                "(SELECT cere_cd_Pcr FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & "), " & _
                "(SELECT cere_cd_Pcr FROM plccustoforncli WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and empr_cd_empresa = " & PCodEmpresa & "), " & _
                "(SELECT dbo.PlanoContas.plco_cd_CodReduzido AS CodRedRecDesp FROM   dbo.PlanoContas RIGHT OUTER JOIN " & _
                "dbo.PlCCustoFornCli ON dbo.PlanoContas.plco_cd_Conta = dbo.PlCCustoFornCli.plco_cd_ContaRecDesp WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & "), " & _
                "(SELECT  dbo.PlanoContas.plco_tx_DescContabil AS contarecdesp FROM  dbo.PlanoContas RIGHT OUTER JOIN " & _
                "dbo.PlCCustoFornCli ON dbo.PlanoContas.plco_cd_Conta = dbo.PlCCustoFornCli.plco_cd_ContaRecDesp WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & "), " & _
                "(SELECT  PlanoContas_1.plco_cd_CodReduzido AS CodRedCusto FROM  dbo.PlanoContas PlanoContas_1 RIGHT OUTER JOIN " & _
                "dbo.PlCCustoFornCli ON dbo.PlCCustoFornCli.plco_cd_ContaCusto = PlanoContas_1.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & "), " & _
                "(SELECT  PlanoContas_1.plco_tx_DescContabil AS contacusto FROM    dbo.PlanoContas PlanoContas_1 RIGHT OUTER JOIN " & _
                "dbo.PlCCustoFornCli ON dbo.PlCCustoFornCli.plco_cd_ContaCusto = PlanoContas_1.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & "), " & _
                "(SELECT  PlanoContas_2.plco_cd_CodReduzido AS CodRedProvAdiant FROM         dbo.PlanoContas PlanoContas_2 RIGHT OUTER JOIN " & _
                "dbo.PlCCustoFornCli ON dbo.PlCCustoFornCli.plco_cd_ContaProvAdiant = PlanoContas_2.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & "), " & _
                "(SELECT  PlanoContas_2.plco_tx_DescContabil AS contaprovadiant FROM         dbo.PlanoContas PlanoContas_2 RIGHT OUTER JOIN " & _
                "dbo.PlCCustoFornCli ON dbo.PlCCustoFornCli.plco_cd_ContaProvAdiant = PlanoContas_2.plco_cd_Conta WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & "), " & _
                "(SELECT  dbo.CentroResultados.cere_tx_Descricao AS cere_tx_Descricao FROM       dbo.CentroResultados RIGHT OUTER JOIN " & _
                "dbo.PlCCustoFornCli ON dbo.CentroResultados.cere_cd_Pcr = dbo.PlCCustoFornCli.cere_cd_Pcr WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & "), " & _
                "(SELECT  dbo.CentroResultados.cere_cd_Estruturado AS cere_cd_Estruturado FROM dbo.CentroResultados RIGHT OUTER JOIN " & _
                "dbo.PlCCustoFornCli ON dbo.CentroResultados.cere_cd_Pcr = dbo.PlCCustoFornCli.cere_cd_Pcr WHERE plccustoforncli.focl_cd_forncli = fornclientes.focl_cd_forncli and PlCCustoFornCli.empr_cd_empresa = " & PCodEmpresa & ") " & _
                "FROM fornclientes WHERE focl_tx_tipo = '" & Pessoa & "' and focl_tx_classe='" & FornClie & "'"
        '*************************************************************************************
        'subCarregaDadosNV Adodc1, "", "focl_tx_fantasia", "", XLT_SQL
        
        TDBGrid1.Columns(13).Visible = False
        TDBGrid1.Columns(14).Visible = False
        TDBGrid1.Columns(15).Visible = False
        TDBGrid1.Columns(16).Caption = "Conta Receita"
        TDBGrid1.Columns(18).Visible = False
        TDBGrid1.Columns(20).Visible = False
        TDBGrid1.Columns(21).Visible = False
        TDBGrid1.Columns(22).Visible = False
        If Pessoa = "F" Then 'Pessoa Física
            CmdAlterarCnpj.Caption = "Alterar CPF"
            TelaFornecedores.Caption = "Tabela Clientes (Pessoa Física)"
            TDBGrid1.Columns(0).Caption = "Nome"
            TDBGrid1.Columns(2).Caption = "Nome Completo"
            TDBGrid1.Columns(3).Caption = "CPF"
        ElseIf Pessoa = "J" Then 'Pessoa Jurídica
            CmdAlterarCnpj.Caption = "Alterar CNPJ"
            TelaFornecedores.Caption = "Tabela Clientes (Pessoa Jurídica)"
            TDBGrid1.Columns(0).Caption = "Fantasia"
            TDBGrid1.Columns(2).Visible = True
            TDBGrid1.Columns(2).Caption = "Razão Social"
            TDBGrid1.Columns(3).Caption = "CGC"
            'TDBGrid1.Columns(13).Visible = True
            'TDBGrid1.Columns(14).Visible = True
            'TDBGrid1.Columns(15).Visible = True
        End If
      
        If XGT_NomeAplic = "CAPI" Or XGT_NomeAplic = "QUALIPAT" Then 'Está na aplicação do CAPI
            TDBGrid1.Columns(16).Visible = False
            TDBGrid1.Columns(17).Visible = False
            TDBGrid1.Columns(19).Visible = False
        End If
    End If
    HabilitaBotoes
    
    filtra_especial = True
    
    TelaValoresGlobais "G"
    PrimeiraVez = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaFornecedores"
End Sub

Private Sub TDBGrid1_SelChange(Cancel As Integer)

'Rotina para agrupar od ID dos FornClientes conjunto de campos selecionados de uma coluna

Dim XLD_ID As String
    Dim i As Integer

    i = 0
            For i = 0 To TDBGrid1.SelBookmarks.Count - 1 Step 1
                If (Adodc1.Recordset(0)) <> "" Then
                    XLD_ID = XLD_ID & (Adodc1.Recordset(0)) & ","
                    Adodc1.Recordset.MoveNext
                End If
            Next
        XLD_TOTALID = XLD_ID
End Sub

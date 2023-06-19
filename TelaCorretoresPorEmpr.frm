VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form TelaCorretoresPorEmpr 
   Caption         =   "Tabela de Corretores por Empreendimento"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   1260
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   9840
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanCorretores 
      Height          =   2925
      Left            =   2520
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   8475
      _Version        =   65536
      _ExtentX        =   14949
      _ExtentY        =   5159
      _StockProps     =   15
      Caption         =   " Cadastro de Corretores"
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
         Left            =   6855
         TabIndex        =   12
         Top             =   1620
         Width           =   975
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
         Left            =   5790
         TabIndex        =   11
         Top             =   1620
         Width           =   975
      End
      Begin VB.TextBox TxtComissao 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1740
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1590
         Width           =   705
      End
      Begin MSAdodcLib.Adodc DatFornecedor 
         Height          =   330
         Left            =   5760
         Top             =   2280
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
         Caption         =   "DatFornecedor"
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
      Begin MSDataListLib.DataCombo CboCorretor 
         Bindings        =   "TelaCorretoresPorEmpr.frx":0000
         Height          =   315
         Left            =   1740
         TabIndex        =   13
         Top             =   1125
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "corr_tx_Nome"
         BoundColumn     =   "corr_cd_Corretor"
         Text            =   "CboCorretor"
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
      Begin MSAdodcLib.Adodc DatEmpreendimento 
         Height          =   330
         Left            =   3240
         Top             =   2280
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
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
      Begin MSMask.MaskEdBox MskEmpreendimento 
         Height          =   315
         Left            =   1740
         TabIndex        =   16
         Top             =   645
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AAAA"
         PromptChar      =   " "
      End
      Begin MSDataListLib.DataCombo CboEmpreendimento 
         Bindings        =   "TelaCorretoresPorEmpr.frx":001C
         Height          =   315
         Left            =   2460
         TabIndex        =   17
         Top             =   645
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "empd_tx_nome"
         BoundColumn     =   "empd_cd_empreendimento"
         Text            =   "CboEmpreendimento"
      End
      Begin VB.Label LblEmpreendimento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   18
         Top             =   705
         Width           =   1515
      End
      Begin VB.Label LblComissao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comissão (%):"
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
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1635
         Width           =   1515
      End
      Begin VB.Label LblCorretor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Corretor:"
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
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Top             =   1200
         Width           =   1515
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   3240
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Bindings        =   "TelaCorretoresPorEmpr.frx":003C
      Height          =   3690
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   14580
      _ExtentX        =   25718
      _ExtentY        =   6509
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "empd_cd_empreendimento"
      Columns(0).DataWidth=   4
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Empreendimento"
      Columns(1).DataField=   "empd_tx_nome"
      Columns(1).DataWidth=   50
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Corretor"
      Columns(2).DataField=   "corr_tx_Nome"
      Columns(2).DataWidth=   65
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Comissão (%)"
      Columns(3).DataField=   "coem_vl_PercentPadrao"
      Columns(3).DataWidth=   15
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=529"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=8229"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=8149"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=528"
      Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(17)=   "Column(1)._MinWidth=4"
      Splits(0)._ColumnProps(18)=   "Column(2).Width=8096"
      Splits(0)._ColumnProps(19)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._WidthInPix=8017"
      Splits(0)._ColumnProps(21)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(2)._ColStyle=528"
      Splits(0)._ColumnProps(23)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(24)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(25)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(26)=   "Column(3).Width=2275"
      Splits(0)._ColumnProps(27)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._WidthInPix=2196"
      Splits(0)._ColumnProps(29)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(3)._ColStyle=530"
      Splits(0)._ColumnProps(31)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(32)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(33)=   "Column(3).Order=4"
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
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15,.alignment=3"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17,.alignment=3"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(54)  =   "Named:id=33:Normal"
      _StyleDefs(55)  =   ":id=33,.parent=0"
      _StyleDefs(56)  =   "Named:id=34:Heading"
      _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(58)  =   ":id=34,.wraptext=-1"
      _StyleDefs(59)  =   "Named:id=35:Footing"
      _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(61)  =   "Named:id=36:Selected"
      _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=37:Caption"
      _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(65)  =   "Named:id=38:HighlightRow"
      _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(67)  =   "Named:id=39:EvenRow"
      _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(69)  =   "Named:id=40:OddRow"
      _StyleDefs(70)  =   ":id=40,.parent=33"
      _StyleDefs(71)  =   "Named:id=41:RecordSelector"
      _StyleDefs(72)  =   ":id=41,.parent=34"
      _StyleDefs(73)  =   "Named:id=42:FilterBar"
      _StyleDefs(74)  =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   8775
      _Version        =   65536
      _ExtentX        =   15478
      _ExtentY        =   1296
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
         Left            =   4500
         TabIndex        =   8
         Top             =   120
         Width           =   975
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
         Left            =   1275
         TabIndex        =   0
         Top             =   120
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
         Left            =   7440
         TabIndex        =   5
         Top             =   120
         Width           =   615
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
         Left            =   6300
         TabIndex        =   4
         Top             =   120
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
         Left            =   5580
         TabIndex        =   3
         Top             =   120
         Width           =   735
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
         Left            =   3420
         TabIndex        =   2
         Top             =   120
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
         Left            =   2340
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
   End
End
Attribute VB_Name = "TelaCorretoresPorEmpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************************
'Variaveis que guardarão valores específicos de cada formulário
Dim PrimeiraVez As Boolean

Dim FRM_ContFil As Integer          ' Contador que armazena o número de filtros ativos
Dim FRM_Ordem As String             ' Nome do Campo que tem a chave da Tabela ativa
Dim FRM_Campos_Obrig As String      ' Lista dos campos que são obrigatórios na tabela
Dim FRM_TipoVar As String           ' Tipo do campo a ser filtrado
Dim FRM_FiltroOrig As String        ' Indica os filtros determinados pelo próprio sistema
Dim FRM_FiltroAtual As String       ' Indica o filtro atual da tabela corrente
Dim FRM_filtra_especial As Integer  ' Indica se a tabela utilizada possui filtros pré-definidos pelo sistema em FiltroOrig
Dim FRM_Tem_Filtro As Boolean       ' Informa se já foi relizado um filtro na tabela pelo usuário
Dim FRM_ConTelaOrdem As Integer     ' Indica o número de chaves utilizados para ordenação
Dim FRM_Quantid As Long             ' Quantidade de registros da tabela
Dim FRM_Filtrou As Boolean          ' Informa se Algum campo foi filtrado
Dim FRM_Ordens() As Integer         ' Vetor que mantem a ordem dos campos
Dim FRM_Filtros() As String         ' Array para armazenar os Filtros da Tabela

Dim FRM_BarraFerramentasPressionada As Boolean
Dim FRM_VariasVezes As Boolean      ' Indica se vai executar a insercao na TelaPadrao varias ou uma vez
Dim FRM_XInserir As String          ' Flag para identificar se entrou na rotina para inserir ou alterar ('I'Inserir, 'A'Alterar, 'C'Consultar)
Dim FRM_XOrigem As String           ' Identifica a origem (que formulario) chamou o formulário ativo
Dim FRM_SQLAtivo As String          ' Atualizar o SQL Ativo após modificações de filtro e ordem

'
' Rotina para tratar informações de tela (get/set)
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

'
' Rotina para habilitar os botões da tela inicial e tratar filtro ativo
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

'
' Rotina para desabilitar os botões da tela inicial e filtro ativo
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

'
' Rotina para inicialização do registro de Log (desativada)
'Private Sub PreencheInformacoesOriginaisLog()
'
'    ReDim XGM_MATRIZLOG(18, 3) As Variant
'
'    XGM_MATRIZLOG(0, 1) = MskCodigo.Text
'    XGM_MATRIZLOG(1, 1) = TxtNome.Text
'    XGM_MATRIZLOG(2, 1) = FunNuloVal(TxtTotUnidades.Text)
'    XGM_MATRIZLOG(3, 1) = TxtEndereco.Text
'    XGM_MATRIZLOG(4, 1) = TxtBairro.Text
'    XGM_MATRIZLOG(5, 1) = TxtCidade.Text
'    XGM_MATRIZLOG(6, 1) = MskEstado.Text
'    XGM_MATRIZLOG(7, 1) = txtBloco.Text
'    XGM_MATRIZLOG(8, 1) = MskCep.Text
'    XGM_MATRIZLOG(9, 1) = Replace(txtPais.Text, "'", "''")
'    XGM_MATRIZLOG(10, 1) = TxtResponsavel.Text
'    XGM_MATRIZLOG(11, 1) = TxtFone.Text
'    XGM_MATRIZLOG(12, 1) = TxtFax.Text
'    XGM_MATRIZLOG(13, 1) = TxtCga.Text
'    XGM_MATRIZLOG(14, 1) = FunNulo(DatConta.Recordset.Fields("plco_tx_DescContabil")) & " - " & FunNulo(DatConta.Recordset.Fields("plco_cd_Estruturado"))
'    XGM_MATRIZLOG(15, 1) = FunNulo(DatPcr.Recordset.Fields("cere_tx_Descricao")) & " - " & FunNulo(DatPcr.Recordset.Fields("cere_cd_Estruturado"))
'    XGM_MATRIZLOG(16, 1) = FunNulo(CboCCorrente.BoundText)
'    XGM_MATRIZLOG(17, 1) = txtCodigoMunicipioIbge.Text
'
'End Sub

'
' Rotina para gravação do registro de Log (desativada)
'Private Sub GravaLog(XLT_OPERACAO As ENUM_OPERACOES, XLT_ORDEM As String)
'
'    Dim XLI_INDICE As Integer
'
'    If XLT_OPERACAO = EXCLUIR Then
'
'        ReDim XGM_MATRIZLOG(2, 2) As Variant
'
'        XGM_MATRIZLOG(0, 0) = "Código"
'        XGM_MATRIZLOG(1, 0) = "Nome"
'
'        XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("empd_cd_empreendimento")
'        XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("empd_tx_Nome")
'
'    Else
'
'        If XLT_OPERACAO <> ALTERAR Then
'            ReDim XGM_MATRIZLOG(18, 2) As Variant
'            XLI_INDICE = 1
'        Else
'            XLI_INDICE = 2
'        End If
'
'        'Campos
'        XGM_MATRIZLOG(0, 0) = "Código"
'        XGM_MATRIZLOG(1, 0) = "Nome"
'        XGM_MATRIZLOG(2, 0) = "Total Unidades"
'        XGM_MATRIZLOG(3, 0) = "Endereço"
'        XGM_MATRIZLOG(4, 0) = "Bairro"
'        XGM_MATRIZLOG(5, 0) = "Cidade"
'        XGM_MATRIZLOG(6, 0) = "Estado"
'        XGM_MATRIZLOG(7, 0) = "Bloco"
'        XGM_MATRIZLOG(8, 0) = "Cep"
'        XGM_MATRIZLOG(9, 0) = "Pais"
'        XGM_MATRIZLOG(10, 0) = "Responsável"
'        XGM_MATRIZLOG(11, 0) = "Fone"
'        XGM_MATRIZLOG(12, 0) = "Fax"
'        XGM_MATRIZLOG(13, 0) = "CGA"
'        XGM_MATRIZLOG(14, 0) = "Conta Contábil"
'        XGM_MATRIZLOG(15, 0) = "Centro de Custo"
'        XGM_MATRIZLOG(16, 0) = "Conta Corrente"
'        XGM_MATRIZLOG(17, 0) = "Cód. Município IBGE"
'
'        'Dados
'        XGM_MATRIZLOG(0, XLI_INDICE) = MskCodigo.Text
'        XGM_MATRIZLOG(1, XLI_INDICE) = TxtNome.Text
'        XGM_MATRIZLOG(2, XLI_INDICE) = FunNuloVal(TxtTotUnidades.Text)
'        XGM_MATRIZLOG(3, XLI_INDICE) = TxtEndereco.Text
'        XGM_MATRIZLOG(4, XLI_INDICE) = TxtBairro.Text
'        XGM_MATRIZLOG(5, XLI_INDICE) = TxtCidade.Text
'        XGM_MATRIZLOG(6, XLI_INDICE) = MskEstado.Text
'        XGM_MATRIZLOG(7, XLI_INDICE) = txtBloco.Text
'        XGM_MATRIZLOG(8, XLI_INDICE) = MskCep.Text
'        XGM_MATRIZLOG(9, XLI_INDICE) = Replace(txtPais.Text, "'", "''")
'        XGM_MATRIZLOG(10, XLI_INDICE) = TxtResponsavel.Text
'        XGM_MATRIZLOG(11, XLI_INDICE) = TxtFone.Text
'        XGM_MATRIZLOG(12, XLI_INDICE) = TxtFax.Text
'        XGM_MATRIZLOG(13, XLI_INDICE) = TxtCga.Text
'        XGM_MATRIZLOG(14, XLI_INDICE) = FunNulo(DatConta.Recordset.Fields("plco_tx_DescContabil")) & " - " & FunNulo(DatConta.Recordset.Fields("plco_cd_Estruturado"))
'        XGM_MATRIZLOG(15, XLI_INDICE) = FunNulo(DatPcr.Recordset.Fields("cere_tx_Descricao")) & " - " & FunNulo(DatPcr.Recordset.Fields("cere_cd_Estruturado"))
'        XGM_MATRIZLOG(16, XLI_INDICE) = FunNulo(CboCCorrente.BoundText)
'        XGM_MATRIZLOG(17, XLI_INDICE) = txtCodigoMunicipioIbge.Text
'
'    End If
'
'    Call subRegistraLog("TelaCorretoresPorEmpr", XLT_ORDEM, funCriaDescricaoLog(XGM_MATRIZLOG, XLT_OPERACAO, EMPREENDIMENTO_CAPI))
'
'End Sub

'
' Rotina para tratar combo de Corretores
Private Sub CboCorretor_Change()

    If CboCorretor.Text <> "" Then
        DatFornecedor.Recordset.bookmark = CboCorretor.SelectedItem
        
        ' Carrega o valor do percentual padrão
        TxtComissao.Text = Format(DatFornecedor.Recordset.Fields("corr_vl_PercentComissao"), "standard")
    End If
    
End Sub

'
' Rotina para tratar combo de Empreendimentos
Private Sub CboEmpreendimento_Change()

    If CboEmpreendimento.Text <> "" Then
        DatEmpreendimento.Recordset.bookmark = CboEmpreendimento.SelectedItem
        MskEmpreendimento.Text = CboEmpreendimento.BoundText
    End If
    
End Sub

'
' Rotina que trata o botão Confirmar, atualização dos registros do cadastro de corretores (implementar)
Private Sub CmdConfirmar_Click()

  Dim XLT_SQL       As String
  Dim XLI_HOUVEERRO As Integer
  Dim XLI_BOOKMARK  As Integer
  Dim XLT_EMPR_CD   As String       ' Código da Empresa (empr_cd_Empresa)
  Dim XLT_EMPD_CD   As String       ' Código do Empreendimento (empd_cd_Empreendimento)
  Dim XLT_CORR_CD   As String       ' Código do Corretor (corr_cd_Corretor)
  Dim XLT_COEM_VL   As String       ' Comissão padrão (coem_vl_PercentPadrao)
  
  XLI_HOUVEERRO = False
  
   If CboEmpreendimento.Text = "" Or CboCorretor.Text = "" Or TxtComissao.Text = "" Then
        MsgBox "O preenchimento dos campos Empreendimento, Corretor e Comissão são obrigatórios.", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
    
    XLT_EMPR_CD = DatEmpreendimento.Recordset.Fields("empr_cd_Empresa")
    XLT_EMPD_CD = DatEmpreendimento.Recordset.Fields("empd_cd_Empreendimento")
    XLT_CORR_CD = DatFornecedor.Recordset.Fields("corr_cd_Corretor")
    XLT_COEM_VL = TxtComissao.Text
        
    On Error GoTo RotuloErro
    
    If PanCorretores.Tag = "I" Then    ' Inclusão
    
        XLT_SQL = "INSERT INTO CorretoresEmpreendimento" & _
                  "            (empr_cd_Empresa, empd_cd_Empreendimento, corr_cd_Corretor, coem_vl_PercentPadrao)" & _
                  "     VALUES (" & XLT_EMPR_CD & ", '" & XLT_EMPD_CD & "', " & XLT_CORR_CD & ", " & FunTrataFloat(XLT_COEM_VL) & ") "
              
        Conexao.BeginTrans
        Conexao.Execute XLT_SQL
        Conexao.CommitTrans

'        GravaLog CADASTRAR, "2"

    Else                                ' Alteração do percentual
    
        XLT_SQL = "UPDATE CorretoresEmpreendimento SET coem_vl_PercentPadrao = " & FunTrataFloat(XLT_COEM_VL) & _
                  " WHERE empr_cd_Empresa = " & XLT_EMPR_CD & _
                  "   AND empd_cd_Empreendimento = '" & XLT_EMPD_CD & "'" & _
                  "   AND corr_cd_Corretor = " & XLT_CORR_CD
    
        Conexao.BeginTrans
        Conexao.Execute XLT_SQL
        Conexao.CommitTrans

'        GravaLog ALTERAR, "3"

    End If

    PanBotoes.Enabled = True
    PanCorretores.Visible = False
    PanCorretores.Enabled = False
    subHabilitaBotoes
    
    If PanCorretores.Tag <> "I" Then
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    Else
        FRM_Quantid = funOperacaoQuantid("+")
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    End If
    
        
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    
  Exit Sub
    
RotuloErro:
    XLI_HOUVEERRO = True
    funTrataErros (ComMensagem)
    Resume Next
    Exit Sub
   
End Sub

'
' Rotina para remover filtros de tela
Private Sub CmdRemoverFiltro_Click()

    Dim XLI_FIL As Integer
    
    subDesabilitaBotoes

    Set Formulario = TelaCorretoresPorEmpr
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

    subHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdOrdem.Enabled = False
       CmdImprimir.Enabled = False
    End If
    
    TDBGrid1.SetFocus

End Sub

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdRemoverFiltro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MdiPrincipal.BarraStatus.Panels(1).Text = "Mostra os filtros ativos da tabela"
    
End Sub

'
' Rotina que trata o botão Alterar, chamada da tela de cadastro de corretores (implementar)
Private Sub CmdAlterar_Click()
    
    PanBotoes.Enabled = False
    PanCorretores.Visible = True
    PanCorretores.Enabled = True
    PanCorretores.Left = (TDBGrid1.Width - PanCorretores.Width) / 2
    PanCorretores.Top = (TDBGrid1.Height - PanCorretores.Height) / 2
    PanCorretores.Tag = Adodc1.Recordset.Fields("corr_cd_corretor")
    PanCorretores.Caption = " Alterar a Comissão(%) do Corretor"
    
    '
    ' Atualiza campos/combos da tela com os dados lidos em Adodc1
    MskEmpreendimento.Text = Adodc1.Recordset.Fields("empd_cd_Empreendimento")
    CboEmpreendimento.Text = Adodc1.Recordset.Fields("empd_tx_Nome")
    CboEmpreendimento.BoundText = Adodc1.Recordset.Fields("empd_cd_Empreendimento")
    CboCorretor.Text = Adodc1.Recordset.Fields("corr_tx_Nome")
    CboCorretor.BoundText = Adodc1.Recordset.Fields("corr_cd_Corretor")
    TxtComissao.Text = Format(Adodc1.Recordset.Fields("coem_vl_PercentPadrao"), "standard")
    
    '
    ' Bloqueia combos para evitar alteração da associação (deve ser feita uma exclusão e posterior inclusão)
    MskEmpreendimento.Enabled = False
    CboEmpreendimento.Enabled = False
    CboCorretor.Enabled = False
    
    subDesabilitaBotoes

    TDBGrid1.Enabled = False
    
    TxtComissao.SetFocus

End Sub

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdAlterar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
  
End Sub

'
' Rotina para tratar o botão de retornar (saída da tela de manutenção do corretor)
Private Sub CmdRetornar_Click()

    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.CancelUpdate
    
    PanBotoes.Enabled = True
    PanCorretores.Visible = False
    PanCorretores.Enabled = False
    subHabilitaBotoes
        
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    
End Sub

'
' Rotina para tratar o botão de imprimir (impressão do grid)
Private Sub CmdImprimir_Click()

    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Corretores por Empreendimento")
    
End Sub

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MdiPrincipal.BarraStatus.Panels(1).Text = "Imprime os dados da tabela"
    
End Sub

'
' Rotina que trata o botão sair
Private Sub CmdSair_Click()

    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    Unload Me
    
End Sub

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
    
End Sub

'
' Rotina que trata o botão ordem (classificação das linhas no grid)
Private Sub CmdOrdem_Click()
    
    subDesabilitaBotoes
    Set Formulario = TelaCorretoresPorEmpr
    subTelaValoresGlobais "P"
    TelaOrdem.Show 1
    subTelaValoresGlobais "G"
    subHabilitaBotoes
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus

End Sub

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdOrdem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MdiPrincipal.BarraStatus.Panels(1).Text = "Ordena os registros atuais da tabela"
    
End Sub

'
' Rotina que trata o botão excluir
Private Sub CmdExcluir_Click()

    Dim XLT_SQL       As String
    Dim XLI_OPCAO As Integer
    Dim XLI_HOUVEERRO As Integer
    Dim XLI_BOOKMARK  As Integer
    Dim XLT_EMPR_CD   As String       ' Código da Empresa (empr_cd_Empresa)
    Dim XLT_EMPD_CD   As String       ' Código do Empreendimento (empd_cd_Empreendimento)
    Dim XLT_CORR_CD   As String       ' Código do Corretor (corr_cd_Corretor)
                
    On Error GoTo RotuloErro
    
    subDesabilitaBotoes
    
    XLI_OPCAO = MsgBox("Confirma Remoção do Registro ?", 273, "CUIDADO")
    
    If XLI_OPCAO = 1 Then  'Confirma exclusão
    
        XLT_EMPR_CD = Adodc1.Recordset.Fields("empr_cd_Empresa")
        XLT_EMPD_CD = Adodc1.Recordset.Fields("empd_cd_Empreendimento")
        XLT_CORR_CD = Adodc1.Recordset.Fields("corr_cd_Corretor")
    
        XLT_SQL = "DELETE FROM CorretoresEmpreendimento " & _
                  " WHERE empr_cd_Empresa = " & XLT_EMPR_CD & _
                  "   AND empd_cd_Empreendimento = '" & XLT_EMPD_CD & "'" & _
                  "   AND corr_cd_Corretor = " & XLT_CORR_CD
              
        Conexao.BeginTrans
        Conexao.Execute XLT_SQL
        Conexao.CommitTrans
      
'      GravaLog EXCLUIR, "4"
      
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
       
    TDBGrid1.SetFocus
    TDBGrid1.Col = 0
    
    Exit Sub
    
RotuloErro:
    
    Conexao.RollbackTrans
    If Err.Number = -2147467259 Or Err.Number = -2147217873 Then
      MsgBox "Esse corretor não pode ser excluído pois já possui contratos relacionados!", vbCritical, "ATENÇÃO"
    End If
    
    subHabilitaBotoes
    
End Sub

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
    
End Sub

'
' Rotina que trata o botão inserir
Private Sub CmdInserir_Click()
    
    PanBotoes.Enabled = False
    PanCorretores.Visible = True
    PanCorretores.Enabled = True
    PanCorretores.Left = (TDBGrid1.Width - PanCorretores.Width) / 2
    PanCorretores.Top = (TDBGrid1.Height - PanCorretores.Height) / 2
    PanCorretores.Tag = "I"
    PanCorretores.Caption = " Associar Corretor ao Empreendimento"
    
    '
    ' Desbloqueia combos e campos
    MskEmpreendimento.Enabled = True
    CboEmpreendimento.Enabled = True
    CboCorretor.Enabled = True
    TxtComissao.Enabled = True

    '
    ' Inicializa campos/combos da tela
    MskEmpreendimento.Text = ""
    CboEmpreendimento.Text = ""
    CboEmpreendimento.BoundText = ""
    CboCorretor.Text = ""
    CboCorretor.BoundText = ""
    TxtComissao.Text = ""
    
    subDesabilitaBotoes

    TDBGrid1.Enabled = False
    
    CboEmpreendimento.SetFocus
    
End Sub

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdInserir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MdiPrincipal.BarraStatus.Panels(1).Text = "Insere um novo registro na tabela"
    
End Sub

'
' Rotina para tratar erro no acesso ao banco de dados (Adodc)
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

Private Sub MskEmpreendimento_GotFocus()

    Call subSelecionaMSK(MskEmpreendimento)
    
End Sub

Private Sub MskEmpreendimento_LostFocus()

    If MskEmpreendimento.Text <> "" Then
        SubPreencheComZero MskEmpreendimento
        CboEmpreendimento.BoundText = MskEmpreendimento.Text
        
        If CboEmpreendimento.Text = "" Then
            MsgBox "Não existe Empreendimento com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskEmpreendimento.SetFocus
            Exit Sub
        End If
        
    End If
    
End Sub

Private Sub MskEmpreendimento_KeyPress(KeyAscii As Integer)

    If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
    End If
    
End Sub


'
' Rotina para tratar evento de do mouse (DblClick) sobre linha do grid
Private Sub TDBGrid1_DblClick()

    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
    
End Sub

'
' Rotina para tratar evento de do mouse (MouseUp) sobre linha do grid
Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaCorretoresPorEmpr
        subTelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, x
        subTelaValoresGlobais "G"
    End If
    
    subHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) = True Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If
    
End Sub

'
' Rotina para a ativação do formulário
Private Sub Form_Activate()

    Set Formulario = TelaCorretoresPorEmpr
    
    If Not PrimeiraVez Then
        subTelaValoresGlobais "P"
    End If
    
    PrimeiraVez = False

End Sub

'
' Rotina para carregar o formulário
Private Sub Form_Load()
    
    Dim XLT_SQL As String
    
    Set Formulario = TelaCorretoresPorEmpr
    
    Call SubAjustaTela
    
    subManutencaoJanelasAtivas "I", "TelaCorretoresPorEmpr"
         
    XLT_SQL = "SELECT A.*," & _
              "       B.empd_tx_Nome," & _
              "       C.corr_tx_Nome," & _
              "       C.corr_vl_PercentComissao" & _
              "  FROM CorretoresEmpreendimento A" & _
              "       JOIN" & _
              "       Empreendimentos          B ON  A.empd_cd_Empreendimento = B.empd_cd_Empreendimento" & _
              "                                  AND A.empr_cd_Empresa        = B.empr_cd_Empresa" & _
              "       JOIN" & _
              "       Corretores              C ON  A.corr_cd_Corretor        = C.corr_cd_Corretor" & _
              " WHERE A.empr_cd_Empresa       = " & PCodEmpresa & _
              " ORDER BY A.empd_cd_Empreendimento, C.corr_tx_Nome"
    
    '
    ' Carrega Adodc1 (dados para o Grid)
    subCarregaDadosNV Adodc1, "CorretoresEmpreendimento", "", "", XLT_SQL
    
    '
    ' Carrega DatFornecedor (combo de corretores)
    subConectarControleDadosNV DatFornecedor, "SELECT * FROM Corretores WHERE focl_cd_FornCli IS NOT NULL ORDER BY corr_tx_Nome", Estatico
    
    '
    ' Carrega DatEmpreendimento (combo de empreendimentos)
    subConectarControleDadosNV DatEmpreendimento, "SELECT * FROM Empreendimentos WHERE empr_cd_Empresa = " & PCodEmpresa & " ORDER BY empd_cd_Empreendimento", Estatico
    
    subHabilitaBotoes
    filtra_especial = True
    
    subTelaValoresGlobais "G"
    PrimeiraVez = True
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If
    
    If BarraFerramentasPressionada = True Then
      Call CmdInserir_Click
      BarraFerramentasPressionada = False
    End If

End Sub

'
' Rotina para descarregar o formulário
Private Sub Form_Unload(Cancel As Integer)

    subManutencaoJanelasAtivas "R", "TelaCorretoresPorEmpr"
    
End Sub


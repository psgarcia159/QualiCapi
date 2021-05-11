VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaVinculosCR 
   Caption         =   "Vinculos do Centro de Resultado"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   1260
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   10500
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaVinculosCR.frx":0000
      Height          =   3105
      Left            =   165
      TabIndex        =   0
      Top             =   1080
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   5477
      _LayoutType     =   4
      _RowHeight      =   25
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Centro de Resultado"
      Columns(0).DataField=   "cere_cd_Estruturado"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descrição"
      Columns(1).DataField=   "cere_tx_descricao"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1826"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1746"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=1701322862"
      Splits(0)._ColumnProps(10)=   "Column(1).Width=3281"
      Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=3201"
      Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=528"
      Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(16)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(17)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(18)=   "Column(1)._MinWidth=103613536"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=212,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(9)   =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(10)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton CmdRemover 
      BackColor       =   &H00000000&
      Caption         =   "<<"
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
      Left            =   4800
      TabIndex        =   3
      Top             =   2520
      Width           =   690
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2340
      Top             =   4140
      Visible         =   0   'False
      Width           =   2070
      _ExtentX        =   3651
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   15
      Top             =   4110
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin Threed.SSPanel PanCR 
      Height          =   510
      Left            =   195
      TabIndex        =   5
      Top             =   60
      Width           =   10185
      _Version        =   65536
      _ExtentX        =   17965
      _ExtentY        =   900
      _StockProps     =   15
      ForeColor       =   4210688
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
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
      Height          =   225
      Left            =   4440
      TabIndex        =   4
      Top             =   4275
      Width           =   1470
   End
   Begin VB.CommandButton CmdInserir 
      BackColor       =   &H00000000&
      Caption         =   ">>"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   2040
      Width           =   690
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
      Bindings        =   "TelaVinculosCR.frx":0015
      Height          =   3105
      Left            =   5745
      TabIndex        =   2
      Top             =   1080
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   5477
      _LayoutType     =   4
      _RowHeight      =   25
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Centro de Resultado"
      Columns(0).DataField=   "cere_cd_estruturado"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descrição"
      Columns(1).DataField=   "cere_tx_descricao"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=1701322862"
      Splits(0)._ColumnProps(10)=   "Column(1).Width=3466"
      Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=3387"
      Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=528"
      Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(16)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(17)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(18)=   "Column(1)._MinWidth=103613536"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=212,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(9)   =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(10)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label2 
      Caption         =   "Centro de Resultado (Não Vinculados):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   195
      TabIndex        =   7
      Top             =   705
      Width           =   4065
   End
   Begin VB.Label Label1 
      Caption         =   "Códigos Vinculados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   5715
      TabIndex        =   6
      Top             =   705
      Width           =   2880
   End
End
Attribute VB_Name = "TelaVinculosCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xPlanoContas As Long
Dim xItem As Long
Dim errLoop As ADODB.Error
Dim StrError As String

Private Sub CmdRemover_Click()
    Dim Sql As String
    
    'On Error GoTo MostraErro
    
    If IsNull(Adodc2.Recordset.Fields("cere_cd_pcr")) Then
        MsgBox "Selecione o item do Centro de Resultado (a direita) a ser removido.", vbInformation + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If

    Sql = "DELETE FROM FuncionarioCR WHERE " & _
               "cere_cd_pcr = " & Str(Adodc2.Recordset.Fields("cere_cd_pcr")) & " AND " & _
               "func_cd_operador = '" & Adodc2.Recordset.Fields("func_cd_operador") & "'"
    Conexao.Execute Sql
    
    subCarregaDadosNV Adodc1, "", "cere_cd_estruturado", "", "SELECT CentroResultados.*, FuncionarioCR.func_cd_operador FROM CentroResultados LEFT JOIN FuncionarioCR ON CentroResultados.cere_cd_pcr = FuncionarioCR.cere_cd_pcr WHERE (FuncionarioCR.func_cd_operador is null) and empr_cd_empresa=" & PCodEmpresa
    subCarregaDadosNV Adodc2, "", "cere_cd_estruturado", "", "SELECT CentroResultados.*, FuncionarioCR.func_cd_operador, FuncionarioCR.cere_cd_pcr FROM FuncionarioCR INNER JOIN CentroResultados ON CentroResultados.cere_cd_pcr = FuncionarioCR.cere_cd_pcr WHERE FuncionarioCR.func_cd_operador = " & Chave
    
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

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub CmdInserir_Click()
    Dim Sql As String
    
    On Error GoTo MostraErro
    
    If IsNull(Adodc1.Recordset.Fields("cere_cd_pcr")) Then
        MsgBox "Selecione o Centro de Resultado (a esquerda) a ser vinculado.", vbInformation + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
    'If Adodc1.Recordset.Fields("plco_cd_ultnivel") = 0 Then
    If Adodc1.Recordset.Fields("cere_tx_tipo") <> "A" Then
        MsgBox "Só é possível vincular Centro de Resultados do nível analítico.", vbInformation + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
    
    Sql = "INSERT INTO FuncionarioCR (cere_cd_Pcr, func_cd_operador) " & _
          "VALUES ('" & Adodc1.Recordset.Fields("cere_cd_pcr") & "'," & Chave & ")"
          Conexao.Execute Sql
    
    'subCarregaDadosNV Adodc1, "", "cere_cd_estruturado", "", "SELECT CentroResultados.*, RelacaoPlanos.itpd_cd_item FROM CentroResultados LEFT JOIN RelacaoPlanos ON CentroResultados.plco_cd_conta = RelacaoPlanos.plco_cd_conta WHERE (RelacaoPlanos.itpd_cd_item is null) and empr_cd_empresa=" & PCodEmpresa
    subCarregaDadosNV Adodc1, "", "cere_cd_estruturado", "", "SELECT CentroResultados.*, FuncionarioCR.func_cd_operador FROM CentroResultados LEFT JOIN FuncionarioCR ON CentroResultados.cere_cd_Pcr = FuncionarioCR.cere_cd_Pcr WHERE (FuncionarioCR.func_cd_operador is null) and empr_cd_empresa=" & PCodEmpresa
    subCarregaDadosNV Adodc2, "", "cere_cd_estruturado", "", "SELECT CentroResultados.*, FuncionarioCR.func_cd_operador, FuncionarioCR.cere_cd_pcr FROM FuncionarioCR INNER JOIN CentroResultados ON CentroResultados.cere_cd_Pcr = FuncionarioCR.cere_cd_pcr WHERE FuncionarioCR.func_cd_operador = " & Chave
    
    
    
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

Private Sub Form_Load()
    
    'Ajustar a Tela à resolução do monitor
    TelaVinculosCR.Top = Int(Resolucaoy - TelaVinculosCR.Height) / 2
    TelaVinculosCR.Left = Int(ResolucaoX - TelaVinculosCR.Width) / 2
    
    subCarregaDadosNV Adodc1, "", "cere_cd_Estruturado", "", _
        "SELECT * FROM CentroResultados WHERE cere_cd_Pcr NOT IN (SELECT cere_cd_pcr FROM FuncionarioCR WHERE func_cd_operador = " & Chave & ") and empr_cd_empresa=" & PCodEmpresa
  
    subCarregaDadosNV Adodc2, "", "cere_cd_estruturado", "", "SELECT CentroResultados.*, FuncionarioCR.cere_cd_pcr, FuncionarioCR.func_cd_operador FROM FuncionarioCR INNER JOIN CentroResultados ON CentroResultados.cere_cd_pcr = FuncionarioCR.cere_cd_pcr Where FuncionarioCR.func_cd_operador = " & Chave
  
    PanCR.Caption = "Vínculos de CR para: " & TelaOperador.Adodc1.Recordset.Fields("func_tx_nome")
  
End Sub

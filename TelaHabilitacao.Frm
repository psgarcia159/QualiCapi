VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaHabilitacao 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  "
   ClientHeight    =   5400
   ClientLeft      =   975
   ClientTop       =   2790
   ClientWidth     =   9225
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "TelaHabilitacao.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5400
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc dataModulo 
      Height          =   330
      Left            =   1860
      Top             =   660
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox PanAvancado 
      Height          =   3690
      Left            =   2280
      ScaleHeight     =   3630
      ScaleWidth      =   4500
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   4560
      Begin VB.CommandButton CmdDesabilitaModulo 
         Caption         =   "Desabilitar Rotinas do Módulo Selecionado"
         Height          =   330
         Left            =   420
         TabIndex        =   16
         Top             =   2670
         Width           =   3840
      End
      Begin VB.CommandButton CmdHabilitaModulo 
         Caption         =   "Habilitar Rotinas do Módulo Selecionado"
         Height          =   330
         Left            =   420
         TabIndex        =   15
         Top             =   2178
         Width           =   3840
      End
      Begin VB.CommandButton CmdDesiste 
         Caption         =   "Desiste"
         Height          =   330
         Left            =   420
         TabIndex        =   7
         Top             =   3165
         Width           =   3840
      End
      Begin VB.CommandButton CmdHabiltaGrupo 
         Caption         =   "Habilitar Rotinas do Grupo Selecionado"
         Height          =   330
         Left            =   420
         TabIndex        =   6
         Top             =   1194
         Width           =   3840
      End
      Begin VB.CommandButton CmdDesabiltaGrupo 
         Caption         =   "Desabilitar Rotinas do Grupo Selecionado"
         Height          =   330
         Left            =   420
         TabIndex        =   5
         Top             =   1686
         Width           =   3840
      End
      Begin VB.CommandButton CmdDesabiltaTodas 
         Caption         =   "Desabilitar Todas Rotinas"
         Height          =   330
         Left            =   420
         TabIndex        =   4
         Top             =   702
         Width           =   3840
      End
      Begin VB.CommandButton CmdHabiltaTodas 
         Caption         =   "Habilitar Todas Rotinas"
         Height          =   330
         Left            =   420
         TabIndex        =   3
         Top             =   210
         Width           =   3840
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
      Bindings        =   "TelaHabilitacao.frx":0442
      Height          =   3720
      Left            =   120
      TabIndex        =   9
      Top             =   1050
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   6562
      _LayoutType     =   4
      _RowHeight      =   21
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Descrição"
      Columns(0).DataField=   "roti_tx_descricao"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   112
      Columns(1)._MaxComboItems=   5
      Columns(1).ValueItems(0)._DefaultItem=   0
      Columns(1).ValueItems(0).Value=   "S"
      Columns(1).ValueItems(0).Value.vt=   8
      Columns(1).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(1).ValueItems(0).DisplayValue(0)=   "bHQAAAYDAABCTQYDAAAAAAAANgAAACgAAAAPAAAADwAAAAEAGAAAAAAA0AIAAAAAAAAAAAAAAAAA"
      Columns(1).ValueItems(0).DisplayValue(1)=   "AAAAAADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAADGx8bG"
      Columns(1).ValueItems(0).DisplayValue(2)=   "x8aEhoSEhoSEhoSEhoSEhoSEhoSEhoSEhoSEhoSEhoSEhoSEhoTGx8YAAADGx8b////Gx8bGx8bG"
      Columns(1).ValueItems(0).DisplayValue(3)=   "x8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8aEhoTGx8YAAADGx8b////Gx8YAAAAAAADGx8bGx8bG"
      Columns(1).ValueItems(0).DisplayValue(4)=   "x8bGx8bGx8bGx8bGx8bGx8aEhoTGx8YAAADGx8b////Gx8YAAAAAAAAAAADGx8bGx8bGx8bGx8bG"
      Columns(1).ValueItems(0).DisplayValue(5)=   "x8bGx8bGx8aEhoTGx8YAAADGx8b////Gx8YAAAAAAAAAAAAAAADGx8bGx8bGx8bGx8bGx8bGx8aE"
      Columns(1).ValueItems(0).DisplayValue(6)=   "hoTGx8YAAADGx8b////Gx8YAAAAAAADGx8YAAAAAAADGx8bGx8bGx8bGx8bGx8aEhoTGx8YAAADG"
      Columns(1).ValueItems(0).DisplayValue(7)=   "x8b////Gx8YAAAAAAADGx8bGx8YAAAAAAADGx8bGx8bGx8bGx8aEhoTGx8YAAADGx8b////Gx8bG"
      Columns(1).ValueItems(0).DisplayValue(8)=   "x8YAAADGx8bGx8bGx8YAAAAAAADGx8bGx8bGx8aEhoTGx8YAAADGx8b////Gx8bGx8bGx8bGx8bG"
      Columns(1).ValueItems(0).DisplayValue(9)=   "x8bGx8bGx8YAAAAAAADGx8bGx8aEhoTGx8YAAADGx8b////Gx8bGx8bGx8bGx8bGx8bGx8bGx8bG"
      Columns(1).ValueItems(0).DisplayValue(10)=   "x8YAAAAAAADGx8aEhoTGx8YAAADGx8b////Gx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bG"
      Columns(1).ValueItems(0).DisplayValue(11)=   "x8aEhoTGx8YAAADGx8b////Gx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8aEhoTGx8YA"
      Columns(1).ValueItems(0).DisplayValue(12)=   "AADGx8b////////////////////////////////////////////////Gx8bGx8YAAADGx8bGx8bG"
      Columns(1).ValueItems(0).DisplayValue(13)=   "x8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAAA="
      Columns(1).ValueItems(0).DisplayValue.vt=   9
      Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(1)._DefaultItem=   0
      Columns(1).ValueItems(1).Value=   "N"
      Columns(1).ValueItems(1).Value.vt=   8
      Columns(1).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(1).ValueItems(1).DisplayValue(0)=   "bHQAAAYDAABCTQYDAAAAAAAANgAAACgAAAAPAAAADwAAAAEAGAAAAAAA0AIAAAAAAAAAAAAAAAAA"
      Columns(1).ValueItems(1).DisplayValue(1)=   "AAAAAADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAADGx8bG"
      Columns(1).ValueItems(1).DisplayValue(2)=   "x8aEhoSEhoSEhoSEhoSEhoSEhoSEhoSEhoSEhoSEhoSEhoSEhoTGx8YAAADGx8b////Gx8bGx8bG"
      Columns(1).ValueItems(1).DisplayValue(3)=   "x8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8aEhoTGx8YAAADGx8b////Gx8bGx8bGx8bGx8bGx8bG"
      Columns(1).ValueItems(1).DisplayValue(4)=   "x8bGx8bGx8bGx8bGx8bGx8aEhoTGx8YAAADGx8b////Gx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bG"
      Columns(1).ValueItems(1).DisplayValue(5)=   "x8bGx8bGx8aEhoTGx8YAAADGx8b////Gx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8aE"
      Columns(1).ValueItems(1).DisplayValue(6)=   "hoTGx8YAAADGx8b////Gx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8aEhoTGx8YAAADG"
      Columns(1).ValueItems(1).DisplayValue(7)=   "x8b////Gx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8aEhoTGx8YAAADGx8b////Gx8bG"
      Columns(1).ValueItems(1).DisplayValue(8)=   "x8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8aEhoTGx8YAAADGx8b////Gx8bGx8bGx8bGx8bG"
      Columns(1).ValueItems(1).DisplayValue(9)=   "x8bGx8bGx8bGx8bGx8bGx8bGx8aEhoTGx8YAAADGx8b////Gx8bGx8bGx8bGx8bGx8bGx8bGx8bG"
      Columns(1).ValueItems(1).DisplayValue(10)=   "x8bGx8bGx8bGx8aEhoTGx8YAAADGx8b////Gx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bG"
      Columns(1).ValueItems(1).DisplayValue(11)=   "x8aEhoTGx8YAAADGx8b////Gx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8aEhoTGx8YA"
      Columns(1).ValueItems(1).DisplayValue(12)=   "AADGx8b////////////////////////////////////////////////Gx8bGx8YAAADGx8bGx8bG"
      Columns(1).ValueItems(1).DisplayValue(13)=   "x8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAAA="
      Columns(1).ValueItems(1).DisplayValue.vt=   9
      Columns(1).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems.Count=   2
      Columns(1).Caption=   "Habilitado"
      Columns(1).DataField=   "habi_bl_habilitado"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=13282"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=13176"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=609"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=503"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=131073"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=128,.bold=-1,.fontsize=825,.italic=0"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.alignment=2"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17,.locked=-1"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15,.alignment=2"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0"
      _StyleDefs(40)  =   "Named:id=34:Heading"
      _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=34,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=35:Footing"
      _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=36:Selected"
      _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=37:Caption"
      _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(49)  =   "Named:id=38:HighlightRow"
      _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=39:EvenRow"
      _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=40:OddRow"
      _StyleDefs(54)  =   ":id=40,.parent=33"
      _StyleDefs(55)  =   "Named:id=41:RecordSelector"
      _StyleDefs(56)  =   ":id=41,.parent=34"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc dataClass 
      Height          =   330
      Left            =   6915
      Top             =   630
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7050
      Top             =   4890
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   2
      CommandType     =   1
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSPanel PanUsuario 
      Height          =   420
      Left            =   150
      TabIndex        =   8
      Top             =   30
      Width           =   8940
      _Version        =   65536
      _ExtentX        =   15769
      _ExtentY        =   741
      _StockProps     =   15
      Caption         =   "Nome do Usuário"
      ForeColor       =   16711680
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
      FloodColor      =   12632256
   End
   Begin VB.CommandButton CmdAvancado 
      Caption         =   "Avançado"
      Height          =   330
      Left            =   2880
      TabIndex        =   1
      Top             =   4950
      Width           =   1620
   End
   Begin VB.CommandButton Botao_Ret 
      Caption         =   "Retornar"
      Height          =   330
      Left            =   4695
      TabIndex        =   0
      Top             =   4950
      Width           =   1620
   End
   Begin MSDataListLib.DataCombo Classificacao 
      Bindings        =   "TelaHabilitacao.frx":0457
      Height          =   315
      Left            =   5400
      TabIndex        =   10
      Top             =   645
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "roti_tx_classificacao"
      BoundColumn     =   "roti_tx_classificacao"
      Text            =   "Classificacao"
   End
   Begin MSDataListLib.DataCombo CboModulo 
      Bindings        =   "TelaHabilitacao.frx":046F
      Height          =   315
      Left            =   855
      TabIndex        =   12
      Top             =   645
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "roti_tx_modulo"
      BoundColumn     =   "roti_tx_modulo"
      Text            =   "CboModulo"
   End
   Begin Threed.SSCommand CmdLimparCategoria 
      Height          =   300
      Left            =   4125
      TabIndex        =   14
      Top             =   645
      Width           =   390
      _Version        =   65536
      _ExtentX        =   688
      _ExtentY        =   529
      _StockProps     =   78
      ForeColor       =   -2147483630
      Enabled         =   0   'False
      MouseIcon       =   "TelaHabilitacao.frx":0488
      Picture         =   "TelaHabilitacao.frx":08DA
   End
   Begin VB.Label LblModulo 
      Alignment       =   1  'Right Justify
      Caption         =   "Módulo:"
      Height          =   270
      Left            =   30
      TabIndex        =   13
      Top             =   675
      Width           =   810
   End
   Begin VB.Label LblGrupo 
      Alignment       =   1  'Right Justify
      Caption         =   "Grupo:"
      Height          =   270
      Left            =   4680
      TabIndex        =   11
      Top             =   675
      Width           =   645
   End
End
Attribute VB_Name = "TelaHabilitacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Houve_Erro As Integer
'Dim Primeiro As Boolean

Private Sub Form_Load()
  'Inicializa o Data Control.
'  Primeiro = True
'  subConectarControleDadosNV dataClass, "SELECT DISTINCT roti_tx_classificacao FROM ROTINAS", Estatico
  subConectarControleDadosNV dataModulo, _
    "SELECT roti_tx_modulo " & _
    "FROM ROTINAS " & _
    "GROUP BY roti_tx_modulo", Estatico
  CboModulo.Text = dataModulo.Recordset.Fields("roti_tx_modulo").Value
'  Classificacao.Text = dataClass.Recordset("roti_tx_classificacao").Value
  
  PanUsuario.Caption = TelaOperador.Adodc1.Recordset.Fields("func_tx_nome")
End Sub

Private Sub Botao_Ret_Click()
  MdiPrincipal.BarraStatus.Panels(1).Text = ""
  MdiPrincipal.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Botao_Ret_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna a Tela Principal"
End Sub

Private Sub CboModulo_Change()
  If CboModulo.Text <> "" Then
    subConectarControleDadosNV dataClass, _
      "SELECT roti_tx_classificacao " & _
      "FROM ROTINAS " & _
      "WHERE roti_tx_modulo = '" & CboModulo.Text & "' " & _
      "GROUP BY roti_tx_classificacao", Estatico
    Classificacao.BoundText = dataClass.Recordset("roti_tx_classificacao").Value
    subReCarregaGrid
  End If
End Sub

Private Sub Classificacao_Change()
  If Classificacao.Text <> "" Then
    subReCarregaGrid
  End If
End Sub

Private Sub CmdAvancado_Click()
  PanAvancado.Enabled = True
  PanAvancado.Visible = True
  CmdDesiste.SetFocus
End Sub

Private Sub CmdDesabilitaModulo_Click()
  Conexao.BeginTrans
  Conexao.Execute funTrataSql("UPDATE  Habilitacoes SET habi_bl_habilitado = 'N'  WHERE func_cd_operador = " & TelaOperador.Adodc1.Recordset.Fields("func_cd_operador") & " AND roti_tx_programa IN (Select roti_tx_programa from rotinas where roti_tx_modulo in ('" & CboModulo.Text & "'))")
  Conexao.CommitTrans
  TelaHabilitacao.Adodc1.Refresh
  PanAvancado.Visible = False
  Botao_Ret.SetFocus
End Sub

Private Sub CmdDesabiltaGrupo_Click()
  Conexao.BeginTrans
  Conexao.Execute funTrataSql("UPDATE  Habilitacoes SET habi_bl_habilitado = 'N'  WHERE func_cd_operador = " & TelaOperador.Adodc1.Recordset.Fields("func_cd_operador") & " AND roti_tx_programa IN (Select roti_tx_programa from rotinas where  roti_tx_classificacao = '" & Classificacao.Text & "' AND roti_tx_modulo in ('" & CboModulo.Text & "'))")
  
  Conexao.CommitTrans
  TelaHabilitacao.Adodc1.Refresh
  PanAvancado.Visible = False
  Botao_Ret.SetFocus
End Sub

Private Sub CmdDesabiltaTodas_Click()
  Conexao.BeginTrans
  Conexao.Execute funTrataSql("UPDATE Habilitacoes SET habi_bl_habilitado = 'N'  WHERE  func_cd_operador = " & TelaOperador.Adodc1.Recordset.Fields("func_cd_operador"))
  Conexao.CommitTrans
  TelaHabilitacao.Adodc1.Refresh
  PanAvancado.Visible = False
  Botao_Ret.SetFocus
End Sub

Private Sub CmdDesiste_Click()
  PanAvancado.Visible = False
  Botao_Ret.SetFocus
End Sub

Private Sub CmdHabilitaModulo_Click()
  Conexao.BeginTrans
  Conexao.Execute funTrataSql("UPDATE  Habilitacoes SET habi_bl_habilitado = 'S'  WHERE func_cd_operador = " & TelaOperador.Adodc1.Recordset.Fields("func_cd_operador") & " AND roti_tx_programa IN (Select roti_tx_programa from rotinas where roti_tx_modulo in ('" & CboModulo.Text & "'))")
  Conexao.CommitTrans
  TelaHabilitacao.Adodc1.Refresh
  PanAvancado.Visible = False
  Botao_Ret.SetFocus
End Sub

Private Sub CmdHabiltaGrupo_Click()
  Conexao.BeginTrans
  Conexao.Execute funTrataSql("UPDATE  Habilitacoes SET habi_bl_habilitado = 'S'  WHERE func_cd_operador = " & TelaOperador.Adodc1.Recordset.Fields("func_cd_operador") & " AND roti_tx_programa IN (Select roti_tx_programa from rotinas where  roti_tx_classificacao = '" & Classificacao.Text & "' AND roti_tx_modulo in ('" & CboModulo.Text & "'))")
  Conexao.CommitTrans
  TelaHabilitacao.Adodc1.Refresh
  PanAvancado.Visible = False
  Botao_Ret.SetFocus
End Sub

Private Sub CmdHabiltaTodas_Click()
  Conexao.BeginTrans
  Conexao.Execute funTrataSql("UPDATE  Habilitacoes  SET habi_bl_habilitado = 'S' WHERE func_cd_operador = " & TelaOperador.Adodc1.Recordset.Fields("func_cd_operador"))
  Conexao.CommitTrans
  TelaHabilitacao.Adodc1.Refresh
  PanAvancado.Visible = False
  Botao_Ret.SetFocus
End Sub

Private Sub CmdLimparCategoria_Click()
  CboModulo.BoundText = 0
  CboModulo.Text = ""
End Sub

Private Sub Table1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
End Sub

Private Sub TDBGrid2_DblClick()
  Dim Descricao As String
  Dim Linha As Integer
  
  Linha = TelaHabilitacao.TDBGrid2.Bookmark

  Descricao = Adodc1.Recordset("roti_tx_descricao").Value
  
  Conexao.BeginTrans
  If TelaHabilitacao.TDBGrid2.Columns.Item(1).Value = "S" Then
'    Conexao.Execute funTrataSql( _
'      "UPDATE Habilitacoes SET" & _
'      "  habi_bl_habilitado = 'N' " & _
'      "WHERE func_cd_operador = " & Cod_Operador & _
'      "  AND roti_tx_programa + roti_tx_ordem IN (Select roti_tx_programa+roti_tx_ordem from rotinas where  roti_tx_descricao = '" & Descricao & "')")
    Conexao.Execute funTrataSql( _
      "UPDATE Habilitacoes SET" & _
      "  habi_bl_habilitado = 'N' " & _
      "WHERE func_cd_operador = " & Adodc1.Recordset("func_cd_operador").Value & _
      "  AND roti_tx_programa = '" & Adodc1.Recordset("roti_tx_programa").Value & "' " & _
      "  AND roti_tx_ordem = " & Adodc1.Recordset("roti_tx_ordem").Value)
  Else
'    Conexao.Execute funTrataSql("UPDATE Habilitacoes SET habi_bl_habilitado = 'S'  WHERE func_cd_operador = " & Cod_Operador & " AND roti_tx_programa+roti_tx_ordem IN (Select roti_tx_programa+roti_tx_ordem from rotinas where roti_tx_descricao = '" & Descricao & "')")
    Conexao.Execute funTrataSql( _
      "UPDATE Habilitacoes SET" & _
      "  habi_bl_habilitado = 'S' " & _
      "WHERE func_cd_operador = " & Adodc1.Recordset("func_cd_operador").Value & _
      "  AND roti_tx_programa = '" & Adodc1.Recordset("roti_tx_programa").Value & "' " & _
      "  AND roti_tx_ordem = '" & Adodc1.Recordset("roti_tx_ordem").Value) & "'"
  End If
  Conexao.CommitTrans
  TelaHabilitacao.Adodc1.Refresh
  TelaHabilitacao.TDBGrid2.Refresh
  TelaHabilitacao.TDBGrid2.Bookmark = Linha
End Sub

Private Sub TDBGrid2_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub subReCarregaGrid()
  
'  Set ADODC1.Recordset = Conexao.Execute( _
'    "SELECT Habilitacoes.*, Rotinas.roti_tx_descricao " & _
'    "FROM Habilitacoes, Rotinas " & _
'    "WHERE Habilitacoes.roti_tx_programa = rotinas.roti_tx_programa " & _
'    "  AND Habilitacoes.roti_tx_ordem = rotinas.roti_tx_ordem " & _
'    "  AND Habilitacoes.func_cd_operador = " & Chave & _
'    "  AND rotinas.roti_tx_classificacao = '" & Classificacao.Text & "' " & _
'    "  AND rotinas.roti_tx_modulo = '" & CboModulo.Text & "' " & _
'    "ORDER BY rotinas.roti_tx_programa, rotinas.roti_tx_ordem")
  subCarregaDadosNV Adodc1, "", "", "", _
    "SELECT Habilitacoes.*, Rotinas.roti_tx_descricao " & _
    "FROM Habilitacoes, Rotinas " & _
    "WHERE Habilitacoes.roti_tx_programa = rotinas.roti_tx_programa " & _
    "  AND Habilitacoes.roti_tx_ordem = rotinas.roti_tx_ordem " & _
    "  AND Habilitacoes.func_cd_operador = " & Chave & _
    "  AND rotinas.roti_tx_classificacao = '" & Classificacao.Text & "' " & _
    "  AND rotinas.roti_tx_modulo = '" & CboModulo.Text & "' " & _
    "ORDER BY rotinas.roti_tx_programa, rotinas.roti_tx_ordem"
End Sub

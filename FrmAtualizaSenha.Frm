VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmAtualizaSenha 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atualização do Operador"
   ClientHeight    =   6675
   ClientLeft      =   975
   ClientTop       =   2385
   ClientWidth     =   11085
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
   HelpContextID   =   1090
   Icon            =   "FrmAtualizaSenha.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6675
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox SenhaEmail 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      HelpContextID   =   1090
      IMEMode         =   3  'DISABLE
      Left            =   5640
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   4
      Tag             =   "&Senha:"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox TxtEmail 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      HelpContextID   =   1090
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "&Email:"
      Top             =   1560
      Width           =   5835
   End
   Begin VB.CommandButton Botao_Des 
      Caption         =   "&Desiste"
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Botao_Con 
      Caption         =   "&Confirma"
      Height          =   375
      Left            =   8400
      TabIndex        =   7
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Nome 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      HelpContextID   =   1090
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "&Nome:"
      Top             =   420
      Width           =   8715
   End
   Begin VB.TextBox Senha 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      HelpContextID   =   1090
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   5
      Tag             =   "&Senha:"
      Top             =   1920
      Width           =   1605
   End
   Begin VB.TextBox Funcao 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      HelpContextID   =   1090
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "&Função:"
      Top             =   810
      Width           =   6915
   End
   Begin MSAdodcLib.Adodc DatPcr 
      Height          =   330
      Left            =   4740
      Negotiate       =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "DatPcr"
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
   Begin Threed.SSCommand CmdLimparPcr 
      Height          =   330
      Left            =   7455
      TabIndex        =   6
      Top             =   1185
      Width           =   360
      _Version        =   65536
      _ExtentX        =   635
      _ExtentY        =   582
      _StockProps     =   78
      ForeColor       =   -2147483630
      MouseIcon       =   "FrmAtualizaSenha.frx":0442
      Picture         =   "FrmAtualizaSenha.frx":0894
   End
   Begin MSDataListLib.DataCombo CboPcr 
      Bindings        =   "FrmAtualizaSenha.frx":09A6
      Height          =   330
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   582
      _Version        =   393216
      Style           =   2
      ListField       =   "desccombo"
      BoundColumn     =   "cere_cd_PCR"
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
   Begin Threed.SSPanel PanVinculo 
      Height          =   3735
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   10695
      _Version        =   65536
      _ExtentX        =   18865
      _ExtentY        =   6588
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton CmdRemoverTodas 
         BackColor       =   &H00000000&
         Caption         =   "<<"
         Height          =   330
         Left            =   4800
         TabIndex        =   25
         Top             =   2520
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton CmdInserirTodas 
         BackColor       =   &H00000000&
         Caption         =   ">>"
         Height          =   330
         Left            =   4800
         TabIndex        =   24
         Top             =   1080
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton CmdRemover 
         BackColor       =   &H00000000&
         Caption         =   "<"
         Height          =   330
         Left            =   4800
         TabIndex        =   23
         Top             =   2040
         Width           =   690
      End
      Begin VB.CommandButton CmdInserir 
         BackColor       =   &H00000000&
         Caption         =   ">"
         Height          =   330
         Left            =   4800
         TabIndex        =   22
         Top             =   1560
         Width           =   690
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Bindings        =   "FrmAtualizaSenha.frx":09BB
         Height          =   3105
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   5477
         _LayoutType     =   4
         _RowHeight      =   25
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "empr_cd_empresa"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Empresa"
         Columns(1).DataField=   "empr_tx_razaosocial"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1296"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1217"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=103613536"
         Splits(0)._ColumnProps(10)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=20"
         Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   3120
         Top             =   120
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   0
         Top             =   0
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Bindings        =   "FrmAtualizaSenha.frx":09D0
         Height          =   3105
         Left            =   5880
         TabIndex        =   21
         Top             =   480
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   5477
         _LayoutType     =   4
         _RowHeight      =   25
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "empr_cd_empresa"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Empresa"
         Columns(1).DataField=   "empr_tx_razaosocial"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1429"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1349"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=103613536"
         Splits(0)._ColumnProps(10)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=20"
         Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
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
      Begin VB.Label Label6 
         Caption         =   "Empresas (Não Vinculadas):"
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
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   4065
      End
      Begin VB.Label Label7 
         Caption         =   "Empresas Vinculadas:"
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
         Left            =   5880
         TabIndex        =   18
         Top             =   120
         Width           =   2880
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha de E-mail:"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label LblEmail 
      Caption         =   "E-mail:"
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha:"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Função:"
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   900
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   540
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Código:"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Cod_Operador 
      Caption         =   "Código do Usuário"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1995
      TabIndex        =   10
      Tag             =   "Código:"
      Top             =   120
      Width           =   1590
   End
   Begin VB.Label LblPCR2 
      Alignment       =   1  'Right Justify
      Caption         =   "Centro Custo:"
      Height          =   210
      Left            =   660
      TabIndex        =   9
      Top             =   1260
      Width           =   1275
   End
End
Attribute VB_Name = "FrmAtualizaSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tb_Operadores As ADODB.Recordset
Dim Codigo As String
Dim errLoop As ADODB.Error
Dim StrError As String

Private Sub Botao_Con_Click()
  'Verificações
  If Trim$(Nome.Text) = "" Then
     MsgBox "Nome do operador deve ser informado.", vbCritical, "ERRO"
     Nome.SetFocus
     Exit Sub
  ElseIf Trim$(Senha.Text) = "" Then
     MsgBox "Senha do usuário deve ser informado.", vbCritical, "ERRO"
     Senha.SetFocus
     Exit Sub
  End If

  On Error GoTo RotuloErro
  Conexao.BeginTrans
  If XInserir = "E" Then
    SubQOpenRecordset Tb_Operadores, _
      "select * from Funcionarios where func_cd_operador = " & Cod_Operador, Dinamico
  Else
    SubQOpenRecordset Tb_Operadores, "select * from Funcionarios", Dinamico
    Tb_Operadores.AddNew
    Tb_Operadores("func_cd_operador").Value = Val(Codigo)
  End If
  
  Tb_Operadores("func_tx_nome").Value = Trim$(Nome.Text)
  Tb_Operadores("func_tx_senha").Value = Crip_Senha(Trim$(Senha.Text))
  
'  If XGB_EMAIL Then
'    Tb_Operadores("func_tx_senhaEmail").Value = Crip_Senha(Trim$(SenhaEmail.Text))
'  End If
  
  Tb_Operadores("func_tx_funcao").Value = Trim$(Funcao.Text)
  'Tb_Operadores!func_cd_Empresa = funNull(CboEmpresa.BoundText)
  Tb_Operadores("func_tx_Email").Value = Trim$(TxtEmail.Text)
  If CboPcr.BoundText <> "" Then
    Tb_Operadores("cere_cd_PCR").Value = CboPcr.BoundText
     XGL_CODPCR = CboPcr.BoundText
  Else
    Tb_Operadores("cere_cd_PCR").Value = Null
  End If
  Tb_Operadores.Update
  If XInserir = "I" Then
    FrmAtualizaSenha.MousePointer = vbHourglass
    Guarda_Ultimo
    subAdicionarRotinas Codigo
  End If
  MdiPrincipal.MousePointer = vbDefault
  FrmAtualizaSenha.MousePointer = vbDefault
  Conexao.CommitTrans
  Unload Me
  Exit Sub

RotuloErro:
  Conexao.RollbackTrans
  'funTrataErros ComMensagem
End Sub

Private Sub Botao_Des_Click()
  MdiPrincipal.BarraStatus.Panels(1).Text = ""
  MdiPrincipal.MousePointer = vbDefault
  Unload Me
End Sub

Private Function Crip_Senha(ByVal Senha As String) As String
  Dim sen As String
  Dim a As Integer

  sen = ""
  For a = 1 To Len(Senha)
      sen = sen + UCase$(Chr$(Asc(Mid$(Senha, a, 1)) + 1))
  Next
  Crip_Senha = sen
End Function

Private Sub CmdInserir_Click()
    Dim Sql As String
    
    On Error GoTo MostraErro
    
    If IsNull(Adodc1.Recordset.Fields("empr_cd_empresa")) Then
        MsgBox "Selecione a empresa (a esquerda) a ser vinculada.", vbInformation + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
    If XInserir = "I" Then
        Chave = Me.Cod_Operador
    End If
                    
    Sql = "INSERT INTO FuncionarioEmpresa(empr_cd_empresa, func_cd_operador) " & _
          "VALUES ('" & Adodc1.Recordset.Fields("empr_cd_empresa") & "'," & Chave & ")"
          Conexao.Execute Sql
    
    subCarregaDadosNV Adodc1, "", "Empresas.empr_cd_empresa", "", "SELECT empr_cd_empresa, empr_tx_razaosocial FROM Empresas WHERE empr_cd_empresa NOT IN (SELECT empr_cd_empresa FROM FuncionarioEmpresa WHERE func_cd_operador = " & Chave & ")"
    subCarregaDadosNV Adodc2, "", "Empresas.empr_cd_empresa", "", "SELECT Empresas.*, FuncionarioEmpresa.func_cd_operador, FuncionarioEmpresa.empr_cd_empresa FROM FuncionarioEmpresa INNER JOIN Empresas ON FuncionarioEmpresa.empr_cd_empresa = Empresas.empr_cd_empresa WHERE FuncionarioEmpresa.func_cd_operador = " & Chave
    
    'subCarregaDadosNV Adodc1, "", "cere_cd_estruturado", "", "SELECT CentroResultados.*, FuncionarioCR.func_cd_operador FROM CentroResultados LEFT JOIN FuncionarioCR ON CentroResultados.cere_cd_Pcr = FuncionarioCR.cere_cd_Pcr WHERE (FuncionarioCR.func_cd_operador is null) and empr_cd_empresa=" & PCodEmpresa
    'subCarregaDadosNV Adodc2, "", "cere_cd_estruturado", "", "SELECT CentroResultados.*, FuncionarioCR.func_cd_operador, FuncionarioCR.cere_cd_pcr FROM FuncionarioCR INNER JOIN CentroResultados ON CentroResultados.cere_cd_Pcr = FuncionarioCR.cere_cd_pcr WHERE FuncionarioCR.func_cd_operador = " & Chave
    
    
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

Private Sub CmdLimparEmpresa_Click()

End Sub

Private Sub CmdLimparPcr_Click()
  CboPcr.BoundText = 0
  CboPcr.Text = ""
End Sub

Private Sub CmdRemover_Click()
    Dim Sql As String
    
    'On Error GoTo MostraErro
    
    If IsNull(Adodc2.Recordset.Fields("empr_cd_empresa")) Then
        MsgBox "Selecione a Empresa (a direita) a ser removida.", vbInformation + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If

    Sql = "DELETE FROM FuncionarioEmpresa WHERE " & _
               "empr_cd_empresa = " & Str(Adodc2.Recordset.Fields("empr_cd_empresa")) & " AND " & _
               "func_cd_operador = '" & Adodc2.Recordset.Fields("func_cd_operador") & "'"
    Conexao.Execute Sql
    
   subCarregaDadosNV Adodc1, "", "Empresas.empr_cd_empresa", "", "SELECT empr_cd_empresa, empr_tx_razaosocial FROM Empresas WHERE empr_cd_empresa NOT IN (SELECT empr_cd_empresa FROM FuncionarioEmpresa WHERE func_cd_operador = " & Chave & ")"
    subCarregaDadosNV Adodc2, "", "Empresas.empr_cd_empresa", "", "SELECT Empresas.*, FuncionarioEmpresa.func_cd_operador, FuncionarioEmpresa.empr_cd_empresa FROM FuncionarioEmpresa INNER JOIN Empresas ON FuncionarioEmpresa.empr_cd_empresa = Empresas.empr_cd_empresa WHERE FuncionarioEmpresa.func_cd_operador = " & Chave
    
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
  subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
  'subConectarControleDadosNV DatEmpresa, "SELECT empr_cd_Empresa,empr_tx_Fantasia FROM Empresas", Estatico
  
  If XInserir = "E" Then
    SubQOpenRecordset Tb_Operadores, "select * from Funcionarios where func_cd_operador = " & Chave, Dinamico
    Nome.Text = Trim$(Tb_Operadores("func_tx_nome").Value)
    Senha.Text = Decrip_Senha(Trim$(Tb_Operadores("func_tx_senha").Value))
    CboPcr.BoundText = FunNulo(Tb_Operadores("cere_cd_pcr").Value)
    Funcao.Text = Trim$(FunNulo(Tb_Operadores("func_tx_funcao").Value))
    Cod_Operador.Caption = Format$(Tb_Operadores("func_cd_operador").Value, "00")
'    CboEmpresa.BoundText = FunNulo(Tb_Operadores!func_cd_Empresa)
    TxtEmail.Text = (Trim$(FunNulo(Tb_Operadores("func_tx_Email").Value)))

'    'Se a empresa trabalha com envio de e-mails no Qualicomp
'    If XGB_EMAIL Then

'        SenhaEmail.Text = Decrip_Senha(Trim$(FunNulo(Tb_Operadores("func_tx_senhaemail").Value)))
'        SenhaEmail.Enabled = True
'    End If
  
  Else
    SubQOpenRecordset Tb_Operadores, "select * from Funcionarios", Dinamico
    Pega_Ultimo
    Cod_Operador.Caption = Codigo
  End If
  
  If pTipoCopia = "Multi" Then
    Me.PanVinculo.Visible = True
  Else
     Me.Left = Int((ResolucaoX - Me.ScaleWidth) / 2)
     Me.Top = Int((Resolucaoy - Me.ScaleHeight) / 2)
     Me.Height = 3540
     Me.Botao_Con.Top = 2500
     Me.Botao_Des.Top = 2500
  End If
  
  If pTipoCopia = "Multi" Then
     If XInserir <> "I" Then
         subCarregaDadosNV Adodc1, "", "empr_cd_empresa", "", _
             "SELECT * FROM Empresas WHERE empr_cd_empresa NOT IN (SELECT empr_cd_empresa FROM FuncionarioEmpresa WHERE func_cd_operador = " & Chave & ")"
         
         subCarregaDadosNV Adodc2, "", "Empresas.empr_cd_empresa", "", "SELECT Empresas.*, FuncionarioEmpresa.empr_cd_empresa, FuncionarioEmpresa.func_cd_operador FROM FuncionarioEmpresa INNER JOIN Empresas ON Empresas.empr_cd_empresa = FuncionarioEmpresa.empr_cd_empresa Where FuncionarioEmpresa.func_cd_operador = " & Chave
    Else
       subCarregaDadosNV Adodc1, "", "empr_cd_empresa", "", _
             "SELECT * FROM Empresas"
    End If
  End If
  
  'PanCR.Caption = "Vínculos de Empresa para: " & TelaOperador.Adodc1.Recordset.Fields("func_tx_nome")
  
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Tb_Operadores.Close
End Sub

'Private Sub Grava_Habilitacoes()
'    Dim Tb_Habilit As Object
'    Dim Tb_Rotinas As Object
'
'    SubQOpenRecordset Tb_Habilit, "select * from Habilitacoes", Dinamico
'    SubQOpenRecordset Tb_Rotinas, "select * from Rotinas", Dinamico
'
'    Do While Not Tb_Rotinas.EOF
'       Tb_Habilit.AddNew
'       Tb_Habilit!func_cd_operador = Val(Codigo)
'       Tb_Habilit!roti_tx_programa = Tb_Rotinas!roti_tx_programa
'       Tb_Habilit!roti_tx_ordem = Tb_Rotinas!roti_tx_ordem
'       Tb_Habilit!habi_bl_habilitado = "N"
'       Tb_Habilit.Update
'       Tb_Rotinas.MoveNext
'    Loop
'
'    Tb_Rotinas.Close
'    Tb_Habilit.Close
'End Sub
Private Sub Guarda_Ultimo()
    Dim Tabela As Object
    SubQOpenRecordset Tabela, "select * from Ultimos", Dinamico
    If Tabela.EOF Then
       Tabela.AddNew
       Tabela!Ulti_cd_operador = 1
    Else
       If IsNull(Tabela!Ulti_cd_operador) Then
          Tabela!Ult_Operador = 1
       Else
          Tabela!Ulti_cd_operador = Val(Codigo)
       End If
    End If
    Tabela.Update
    Tabela.Close
End Sub

Private Sub Pega_Ultimo()
    Dim Tabela As Object
    SubQOpenRecordset Tabela, "select * from Ultimos", Dinamico
    If Tabela.EOF Then
       Codigo = "01"
    Else
       If IsNull(Tabela!Ulti_cd_operador) Then
          Codigo = "01"
       Else
          Codigo = Format$(Tabela!Ulti_cd_operador + 1, "00")
       End If
    End If
    Tabela.Close
End Sub

Private Sub lblVinculoos_Click()

End Sub

Private Sub Senha_KeyPress(KeyAscii As Integer)
    'Aceita apenas caracteres entre 0 - 9; A - Z ; a - z
    If (KeyAscii = vbKeyBack) Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or (KeyAscii >= 97 And KeyAscii <= 122) Then ' 97 = a ; 122 = z
       KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    Else
       KeyAscii = 0
    End If
End Sub
Private Function Decrip_Senha(ByVal Senha As String) As String
    
    Dim sen As String
    Dim a As Integer

    sen = ""
    For a = 1 To Len(Senha)
        sen = sen + UCase$(Chr$(Asc(Mid$(Senha, a, 1)) - 1))
    Next
    Decrip_Senha = sen

End Function

Private Sub subAdicionarRotinas(XLI_CODOPERADOR As String)
  Dim XLT_
  Conexao.Execute _
    "INSERT INTO Habilitacoes " & _
    "SELECT Funcionarios.func_cd_Operador," & _
    "  Rotinas.roti_tx_Programa," & _
    "  Rotinas.roti_tx_Ordem," & _
    "  'N' AS habi_bl_habilitado " & _
    "FROM Rotinas, Funcionarios " & _
    "WHERE Funcionarios.func_cd_operador = " & XLI_CODOPERADOR
End Sub

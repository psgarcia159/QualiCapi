VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TelaExportacaoDimob 
   Caption         =   "Exportação DIMOB"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   10110
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanAno 
      Height          =   1275
      Left            =   3240
      TabIndex        =   19
      Top             =   1320
      Visible         =   0   'False
      Width           =   3600
      _Version        =   65536
      _ExtentX        =   6350
      _ExtentY        =   2249
      _StockProps     =   15
      Caption         =   "Informe o ano da Declaração"
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
      Begin VB.CommandButton cmdConfirmar 
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
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DtpAno 
         Height          =   315
         Left            =   1440
         TabIndex        =   22
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   83689475
         UpDown          =   -1  'True
         CurrentDate     =   38718
      End
      Begin VB.Label LblAno 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ano:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   840
         TabIndex        =   23
         Top             =   525
         Width           =   540
      End
   End
   Begin Threed.SSPanel PanInformacoes 
      Height          =   4575
      Left            =   2880
      TabIndex        =   1
      Top             =   10
      Visible         =   0   'False
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   8070
      _StockProps     =   15
      Caption         =   "Informações para Importação"
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
      Begin VB.CommandButton CmdRetornarExp 
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
         Left            =   3240
         TabIndex        =   21
         Top             =   4080
         Width           =   975
      End
      Begin VB.Frame FraSituacao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Situação Especial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   4095
         Begin VB.OptionButton OptFusao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "02 - Fusão"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton OptCisao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "04 - Cisão Total"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   2160
            Width           =   1935
         End
         Begin VB.OptionButton OptIncorporacao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "03 - Incorporação/Incorporada"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1800
            Width           =   3135
         End
         Begin VB.OptionButton OptExtincao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "01 - Extinção"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   1455
         End
         Begin VB.OptionButton OptNormal 
            BackColor       =   &H00E0E0E0&
            Caption         =   "00 - Normal"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox ChkSituacao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Situação Especial?"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DtpSituacao 
            Height          =   375
            Left            =   2520
            TabIndex        =   13
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   83689473
            CurrentDate     =   39448
         End
         Begin VB.Label LblData 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Data:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2040
            TabIndex        =   24
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.TextBox TxtRecibo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Frame FraTipoDeclaracao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo da Declaração:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   120
         TabIndex        =   7
         Top             =   440
         Width           =   4095
         Begin VB.OptionButton OptRetificadora 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Retificadora"
            Height          =   255
            Left            =   2400
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton OptOriginal 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Original"
            Height          =   255
            Left            =   720
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CommandButton CmdConfirmarExp 
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
         Left            =   2160
         TabIndex        =   2
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label LblRecibo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Número do Recibo:"
         Enabled         =   0   'False
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
         TabIndex        =   18
         Top             =   1080
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog DlgSalvar 
      Left            =   7800
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaExportacaoDimob.frx":0000
      Height          =   4950
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   8731
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "CNPJ/CPF"
      Columns(0).DataField=   "focl_tx_CgcCpf"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Cliente"
      Columns(1).DataField=   "focl_tx_RazaoSocial"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Contrato"
      Columns(2).DataField=   "Contrato"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Data do Contrato"
      Columns(3).DataField=   "cont_dt_RegistroContrato"
      Columns(3).NumberFormat=   "dd/mm/yy"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Total Operação"
      Columns(4).DataField=   "TotalOperacao"
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Valor Pago"
      Columns(5).DataField=   "ValorPago"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Endereço"
      Columns(6).DataField=   "empd_tx_endereco"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Bairro"
      Columns(7).DataField=   "empd_tx_Bairro"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Cidade"
      Columns(8).DataField=   "empd_tx_cidade"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Estado"
      Columns(9).DataField=   "empd_tx_estado"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "CEP"
      Columns(10).DataField=   "empd_nr_cep"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2831"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2752"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=5424"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=5345"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=528"
      Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(1)._MinWidth=4"
      Splits(0)._ColumnProps(16)=   "Column(2).Width=1984"
      Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=1905"
      Splits(0)._ColumnProps(19)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=528"
      Splits(0)._ColumnProps(21)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(22)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(24)=   "Column(3).Width=1640"
      Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=1561"
      Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=529"
      Splits(0)._ColumnProps(29)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(31)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=530"
      Splits(0)._ColumnProps(36)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(37)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(38)=   "Column(4)._MinWidth=89391764"
      Splits(0)._ColumnProps(39)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(40)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(42)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(43)=   "Column(5)._ColStyle=530"
      Splits(0)._ColumnProps(44)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(45)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(46)=   "Column(5)._MinWidth=89391764"
      Splits(0)._ColumnProps(47)=   "Column(6).Width=5689"
      Splits(0)._ColumnProps(48)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(6)._WidthInPix=5609"
      Splits(0)._ColumnProps(50)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(51)=   "Column(6)._ColStyle=528"
      Splits(0)._ColumnProps(52)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(53)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(54)=   "Column(6)._MinWidth=89391764"
      Splits(0)._ColumnProps(55)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(56)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(58)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(59)=   "Column(7)._ColStyle=20"
      Splits(0)._ColumnProps(60)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(61)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(62)=   "Column(8).Width=2884"
      Splits(0)._ColumnProps(63)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(8)._WidthInPix=2805"
      Splits(0)._ColumnProps(65)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(66)=   "Column(8)._ColStyle=528"
      Splits(0)._ColumnProps(67)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(68)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(69)=   "Column(9).Width=1244"
      Splits(0)._ColumnProps(70)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(9)._WidthInPix=1164"
      Splits(0)._ColumnProps(72)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(73)=   "Column(9)._ColStyle=528"
      Splits(0)._ColumnProps(74)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(75)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(76)=   "Column(10).Width=1773"
      Splits(0)._ColumnProps(77)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(10)._WidthInPix=1693"
      Splits(0)._ColumnProps(79)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(80)=   "Column(10)._ColStyle=528"
      Splits(0)._ColumnProps(81)=   "Column(10).WrapText=1"
      Splits(0)._ColumnProps(82)=   "Column(10).Order=11"
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
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=78,.parent=13,.alignment=0"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=0,.locked=0"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15,.alignment=3"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17,.alignment=3"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=82,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=79,.parent=14,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=80,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=81,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=98,.parent=13,.alignment=1"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=95,.parent=14,.alignment=2"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=96,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=97,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=90,.parent=13,.alignment=1"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=87,.parent=14,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=88,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=89,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=0"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=94,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=91,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=92,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=93,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=54,.parent=13,.alignment=0"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=58,.parent=13,.alignment=0"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=62,.parent=13,.alignment=0"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=17"
      _StyleDefs(82)  =   "Named:id=33:Normal"
      _StyleDefs(83)  =   ":id=33,.parent=0"
      _StyleDefs(84)  =   "Named:id=34:Heading"
      _StyleDefs(85)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(86)  =   ":id=34,.wraptext=-1"
      _StyleDefs(87)  =   "Named:id=35:Footing"
      _StyleDefs(88)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(89)  =   "Named:id=36:Selected"
      _StyleDefs(90)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=37:Caption"
      _StyleDefs(92)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(93)  =   "Named:id=38:HighlightRow"
      _StyleDefs(94)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=39:EvenRow"
      _StyleDefs(96)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(97)  =   "Named:id=40:OddRow"
      _StyleDefs(98)  =   ":id=40,.parent=33"
      _StyleDefs(99)  =   "Named:id=41:RecordSelector"
      _StyleDefs(100) =   ":id=41,.parent=34"
      _StyleDefs(101) =   "Named:id=42:FilterBar"
      _StyleDefs(102) =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   420
      Left            =   3000
      TabIndex        =   3
      Top             =   5160
      Width           =   3960
      _Version        =   65536
      _ExtentX        =   6985
      _ExtentY        =   741
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
      Begin VB.CommandButton CmdExportar 
         Caption         =   "&Exportar"
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
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   975
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
         Left            =   1380
         TabIndex        =   5
         Top             =   60
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
         Left            =   2700
         TabIndex        =   4
         Top             =   60
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   60
      Top             =   5100
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
End
Attribute VB_Name = "TelaExportacaoDimob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub subDesabHabBotoes(XHab)
  
  CmdExportar.Enabled = XHab
  CmdImprimir.Enabled = XHab
  CmdSair.Enabled = XHab
  
End Sub

Private Sub ChkSituacao_Click()
    
    If ChkSituacao.Value = 1 Then
        DtpSituacao.Enabled = True
        LblData.Enabled = True
        
        'Alterado em 09/02/09 - Larissa: a opção "Normal" foi desabilitada porque ela não existe mais na nova versão do Programa Dimob.
        OptNormal.Value = False
        OptNormal.Enabled = False
        OptExtincao.Value = True
        
        OptExtincao.Enabled = True
        OptFusao.Enabled = True
        OptIncorporacao.Enabled = True
        OptCisao.Enabled = True
    Else
        LblData.Enabled = False
        DtpSituacao.Enabled = False
        OptNormal.Enabled = False
        OptNormal.Value = False
        OptExtincao.Enabled = False
        OptExtincao.Value = False
        OptFusao.Enabled = False
        OptFusao.Value = False
        OptIncorporacao.Enabled = False
        OptIncorporacao.Value = False
        OptCisao.Enabled = False
        OptCisao.Value = False
    End If

End Sub

Private Sub CmdConfirmar_Click()

  subDesabHabBotoes (True)
  PanAno.Visible = False
  subCarregaDadosNV Adodc1, "ConsCAPExportacaoDimob", "Contrato", _
    "", "SELECT *, ISNULL((SELECT SUM(titu_vl_ValorPago)" & _
    " From Titulos WHERE  dbo.ConsCAPExportacaoDimob.empr_cd_Empresa = dbo.Titulos.empr_cd_Empresa AND" & _
    " dbo.ConsCAPExportacaoDimob.empd_cd_Empreendimento = dbo.Titulos.empd_cd_Empreendimento AND" & _
    " dbo.ConsCAPExportacaoDimob.imov_cd_Imovel = dbo.Titulos.imov_cd_Imovel AND" & _
    " dbo.ConsCAPExportacaoDimob.cont_cd_Contrato = dbo.Titulos.cont_cd_Contrato AND" & _
    " Year(dbo.Titulos.titu_dt_Pagamento)=" & Year(DtpAno.Value) & "), 0)" & _
    " AS ValorPago FROM ConsCAPExportacaoDimob WHERE " & _
    " Year(cont_dt_RegistroContrato) = " & Year(DtpAno.Value) & " AND " & _
    " empr_cd_Empresa = " & PCodEmpresa
    '" AND focl_tx_CGCCPF<>'000.614.175-73'"
    ' Para atender demanda JMJ (não confirmada) comentar a linha após o WHERE

End Sub

Private Sub CmdConfirmarExp_Click()
        
    Dim XLT_TEXTO As String
    Dim XLT_ENDERECO As String
    Dim XLT_CGCEMP As String
    Dim XLT_CGCCPF As String
    Dim XLI_SEQUENCIAL As Integer
    Dim XLI_SEQUENCIALHEXA As Integer
    Dim XLT_CPFRESPONSAVEL As String
    Dim XLT_TIPO As Integer 'Tipo de Declaração 0 - Original e 1 - Retificadora
    Dim XLT_RECIBO As String 'Número do Recibo no caso de declaração retificadora
    Dim XLT_SITUACAO As Integer 'Situação Especial? 1 - Sim 0 - Não
    Dim XLT_DATASITUACAO 'Data da Situação Especial
    Dim XLT_TIPOSITUACAO 'Tipo da situação Especial
    Dim XLI_MUNICIPIO As Integer 'Código do Município
    
    Adodc1.Recordset.MoveFirst
    
    If IsNull(Adodc1.Recordset!empr_tx_CpfResponsavel) Or Adodc1.Recordset!empr_tx_CpfResponsavel = "   .   .   -  " Then
        MsgBox ("É necessário cadastrar o CPF do responsável pela empresa, na tela Padrões de Empresa."), vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
        
    'O número do recibo é obrigatório para a declaração retificadora
    If OptRetificadora.Value = True Then
        If TxtRecibo.Text = "" Then
            MsgBox ("O número do recibo é obrigatório na declaração retificadora, favor inserir!"), vbCritical, "ATENÇÃO"
            Exit Sub
        End If
    End If
    
'    'A data e o tipo de situação é obrigatória no caso de situação especial
'    If ChkSituacao.Value = True Then
'        If OptNormal = OptExtincao = OptFusao = OptIncorporacao = OptCisao = True Then
'            MsgBox ("Favor indicar o tipo da Situação Especial!"), vbCritical, "ATENÇÃO"
'            Exit Sub
'        End If
'    End If
    
    'Verifica se o endereço está com os campos preenchidos
    While Not Adodc1.Recordset.EOF
        If Len(Adodc1.Recordset!focl_tx_CGCCPF) > 14 Then
          If FunConfereCGC(Mid(Adodc1.Recordset!focl_tx_CGCCPF, 1, 2) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 4, 3) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 8, 3) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 12, 4) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 17, 2)) = False Then
             MsgBox "O CNPJ do Cliente " & Adodc1.Recordset!focl_tx_RazaoSocial & _
                 vbCrLf & "está incorreto."
             Exit Sub
           End If
        Else
           If FunConfereCPF(Mid(Adodc1.Recordset!focl_tx_CGCCPF, 1, 3) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 5, 3) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 9, 3) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 13, 2)) = False Then
             MsgBox "O CPF do Cliente " & Adodc1.Recordset!focl_tx_RazaoSocial & _
                 vbCrLf & "está incorreto."
             
             Exit Sub
           End If
         End If
        
        If Adodc1.Recordset!empd_tx_Endereco = "" Or Adodc1.Recordset!empd_tx_Bairro = "" _
          Or Not IsNumeric(Left(Adodc1.Recordset!empd_nr_Cep, 2) + Mid(Adodc1.Recordset!empd_nr_Cep, 4, 3) + Right(Adodc1.Recordset!empd_nr_Cep, 3)) _
          Or Adodc1.Recordset!empd_tx_Cidade = "" Or Adodc1.Recordset!empd_tx_Estado = "" Then
          MsgBox ("O endereço do empreendimento " & Left(Adodc1.Recordset!contrato, 4) & " está preenchido incorretamente"), vbCritical, "ATENÇÃO"
          Exit Sub
        End If
        
        Adodc1.Recordset.MoveNext
    Wend
    
    DlgSalvar.DialogTitle = "Arquivo de Exportação para o DIMOB"
    DlgSalvar.DefaultExt = "txt"
    DlgSalvar.InitDir = App.Path
    DlgSalvar.CancelError = False
    DlgSalvar.ShowSave
    
    If DlgSalvar.FileName = "" Then
      Exit Sub
    End If
    
    If DlgSalvar.FileName <> "" Then
      XGT_LOCALARQ = DlgSalvar.FileName
      XGT_ARQUIVO = DlgSalvar.FileTitle
    End If
    
    
    Adodc1.Recordset.MoveFirst
    
    Open XGT_LOCALARQ For Output As #1
    
    'Declaração é original ou retificadora
    If OptRetificadora.Value = True Then
        XLT_TIPO = "1" 'Sim
        XLT_RECIBO = Format(Trim(TxtRecibo.Text), "0000000000")
    Else
        XLT_TIPO = "0" 'Não
        XLT_RECIBO = String(10, "0")
    End If
    
    'Situação Especial / Data da Situação Especial
    If ChkSituacao.Value = 1 Then
        XLT_SITUACAO = "1"
        XLT_DATASITUACAO = Format(DtpSituacao.Value, "ddmmyyyy")
        
        'Tipo da Situação Especial
        If OptNormal.Value = True Then
            XLT_TIPOSITUACAO = Left(OptNormal.Caption, 2)
        ElseIf OptExtincao.Value = True Then
             XLT_TIPOSITUACAO = Left(OptExtincao.Caption, 2)
        ElseIf OptFusao.Value = True Then
             XLT_TIPOSITUACAO = Left(OptFusao.Caption, 2)
        ElseIf OptIncorporacao.Value = True Then
             XLT_TIPOSITUACAO = Left(OptIncorporacao.Caption, 2)
        ElseIf OptCisao.Value = True Then
             XLT_TIPOSITUACAO = Left(OptCisao.Caption, 2)
        End If
    Else
        XLT_SITUACAO = "0"
        XLT_DATASITUACAO = "00000000"
        XLT_TIPOSITUACAO = Space(2)
    End If
      
    'Sequencial para gerar números em hexadecimal
    XLI_SEQUENCIALHEXA = 1
    XLI_SEQUENCIAL = 1
    'Format(Hex(XLI_SEQUENCIALHEXA), "00")
    XLT_TEXTO = "DIMOB" & _
             Space(369)
    Print #1, funTiraAcento(XLT_TEXTO, True)
    
    
    XLI_SEQUENCIALHEXA = XLI_SEQUENCIALHEXA + 1
    
    XLT_CGCEMP = Mid(XGT_CGC, 1, 2) + Mid(XGT_CGC, 4, 3) + Mid(XGT_CGC, 8, 3) + Mid(XGT_CGC, 12, 4) + Mid(XGT_CGC, 17, 2)
    XLT_CPFRESPONSAVEL = Mid(Adodc1.Recordset!empr_tx_CpfResponsavel, 1, 3) + Mid(Adodc1.Recordset!empr_tx_CpfResponsavel, 5, 3) + Mid(Adodc1.Recordset!empr_tx_CpfResponsavel, 9, 3) + Mid(Adodc1.Recordset!empr_tx_CpfResponsavel, 13, 2)
    
    XLT_TEXTO = "R01" & _
               XLT_CGCEMP & _
               Format(DtpAno.Value, "yyyy") & _
               XLT_TIPO & _
               XLT_RECIBO & _
               XLT_SITUACAO & _
               XLT_DATASITUACAO & _
               XLT_TIPOSITUACAO & _
               Left(PEmpresa, 60) & Space(60 - Len(Left(PEmpresa, 60))) & _
               XLT_CPFRESPONSAVEL & _
               Left(Replace(funTiraAcento(PEndereco + " " + XGT_NUMPORTA + " " + XGT_BAIRRO + " " + PCidade + " " + XGT_ESTADO + " " + XGT_CEP, True), ".", ""), 120) & Space(120 - Len(Left(Replace(funTiraAcento(PEndereco + " " + XGT_NUMPORTA + " " + XGT_BAIRRO + " " + PCidade + " " + XGT_ESTADO + " " + XGT_CEP, True), ".", ""), 120))) & _
               Left(XGT_ESTADO, 2) & _
               "3849" & _
               Space(30)
    Print #1, funTiraAcento(XLT_TEXTO, True)
    
    While Not Adodc1.Recordset.EOF
        
        XLI_SEQUENCIALHEXA = XLI_SEQUENCIALHEXA + 1
        
        If Adodc1.Recordset!focl_tx_Tipo = "F" Then
            XLT_CGCCPF = Mid(Adodc1.Recordset!focl_tx_CGCCPF, 1, 3) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 5, 3) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 9, 3) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 13, 2)
        Else
            XLT_CGCCPF = Mid(Adodc1.Recordset!focl_tx_CGCCPF, 1, 2) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 4, 3) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 8, 3) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 12, 4) + Mid(Adodc1.Recordset!focl_tx_CGCCPF, 17, 2)
        End If
        
        XLT_ENDERECO = Adodc1.Recordset!empd_tx_Endereco & "," & Adodc1.Recordset!empd_tx_Bairro
        
        If Adodc1.Recordset!empd_tx_Cidade = "SALVADOR" Then
            XLI_MUNICIPIO = "3849"
        Else 'Mata de São João
            XLI_MUNICIPIO = "3721"
        End If
        
        XLT_TEXTO = "R03" & _
                    XLT_CGCEMP & _
                    Format(DtpAno.Value, "yyyy") & _
                    Format(XLI_SEQUENCIAL, "00000") & _
                    Left(XLT_CGCCPF, 14) & Space(14 - Len(Left(XLT_CGCCPF, 14))) & _
                    Left(Adodc1.Recordset!focl_tx_RazaoSocial, 60) & Space(60 - Len(Left(Adodc1.Recordset!focl_tx_RazaoSocial, 60))) & _
                    Mid(Adodc1.Recordset!contrato, 6, 4) & Right(Adodc1.Recordset!contrato, 2) & _
                    Format(Adodc1.Recordset!cont_dt_RegistroContrato, "ddmmyyyy") & _
                    Format(Adodc1.Recordset!TotalOperacao * 100, "00000000000000") & _
                    Format(Adodc1.Recordset!VALORPAGO * 100, "00000000000000") & _
                    "U" & _
                    Left(XLT_ENDERECO, 60) & Space(60 - Len(Left(XLT_ENDERECO, 60))) & _
                    Left(Adodc1.Recordset!empd_nr_Cep, 2) + Mid(Adodc1.Recordset!empd_nr_Cep, 4, 3) + Right(Adodc1.Recordset!empd_nr_Cep, 3) & _
                    XLI_MUNICIPIO & _
                    Space(20) & _
                    Left(Adodc1.Recordset!empd_tx_Estado, 2) & Space(2 - Len(Left(Adodc1.Recordset!empd_tx_Cidade, 2))) & _
                    Space(10)
        Print #1, funTiraAcento(XLT_TEXTO, True)
        XLI_SEQUENCIAL = XLI_SEQUENCIAL + 1
        
        Adodc1.Recordset.MoveNext
    Wend
    
    XLT_TEXTO = "T9" & _
              Space(100)
    Print #1, funTiraAcento(XLT_TEXTO, True)
    
    Close #1
    MsgBox "Exportação realizada com sucesso", vbInformation
    PanInformacoes.Visible = False
    
End Sub

Private Sub CmdExportar_Click()
    
    PanInformacoes.Visible = True
    PanInformacoes.Left = (TDBGrid1.Width - PanInformacoes.Width) / 2
    PanInformacoes.Top = (TDBGrid1.Height - PanInformacoes.Height) / 2
    
End Sub

Private Sub cmdRetornarClasse_Click()
    Unload Me
End Sub

Private Sub CmdRetornar_Click()
    
    PanAno.Visible = False

End Sub

Private Sub CmdImprimir_Click()
    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Bancos")
End Sub

Private Sub CmdRetornarExp_Click()
    
    PanInformacoes.Visible = False
    
End Sub

Private Sub CmdSair_Click()
 Unload Me
End Sub

Private Sub DtpSituacao_LostFocus()
    If (Right(DtpSituacao.Value, 4) <> Format(DtpAno, "yyyy")) Then
        MsgBox ("O ano da data do evento não pode ser diferente do ano da declaração!"), vbInformation, "ATENÇÃO"
        DtpSituacao.SetFocus
    End If
End Sub

Private Sub Form_Load()

    Set Formulario = TelaExportacaoDimob
    
    Call SubAjustaTela
    
    subManutencaoJanelasAtivas "I", "TelaExportacaoDimob"
    
    subDesabHabBotoes (False)
    PanAno.Visible = True
    PanAno.Left = (TDBGrid1.Width - PanAno.Width) / 2
    PanAno.Top = (TDBGrid1.Height - PanAno.Height) / 2
    DtpAno.Year = Year(Date) - 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
   subManutencaoJanelasAtivas "R", "TelaExportacaoDimob"
  
End Sub

Private Sub OptOriginal_Click()
    TxtRecibo.Enabled = False
    TxtRecibo.Text = ""
    LblRecibo.Enabled = False
End Sub

Private Sub OptRetificadora_Click()

    TxtRecibo.Enabled = True
    LblRecibo.Enabled = True
    TxtRecibo.SetFocus

End Sub

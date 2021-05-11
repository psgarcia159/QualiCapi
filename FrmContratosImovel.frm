VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManContratosImovel 
   Caption         =   "Tabela de Contratos"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   9390
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   9390
   Begin Threed.SSPanel PanDistratar 
      Height          =   1470
      Left            =   540
      TabIndex        =   16
      Top             =   1980
      Visible         =   0   'False
      Width           =   3570
      _Version        =   65536
      _ExtentX        =   6297
      _ExtentY        =   2593
      _StockProps     =   15
      Caption         =   " Distrato"
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
         Height          =   330
         Left            =   1140
         TabIndex        =   18
         Top             =   1020
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
         Height          =   330
         Left            =   2280
         TabIndex        =   19
         Top             =   1020
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DtpDtDistrato 
         Height          =   315
         Left            =   1980
         TabIndex        =   17
         Top             =   420
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   49676291
         CurrentDate     =   37180
      End
      Begin VB.Label LblDtDistrato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data do Distrato:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   420
         TabIndex        =   20
         Top             =   480
         Width           =   1515
      End
   End
   Begin Threed.SSPanel PanCessao 
      Height          =   2565
      Left            =   1560
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   7365
      _Version        =   65536
      _ExtentX        =   12991
      _ExtentY        =   4524
      _StockProps     =   15
      Caption         =   "Cessão"
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
      Begin VB.CheckBox ChkTitulosPagos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Copiar os títulos pagos?"
         Height          =   195
         Left            =   1740
         TabIndex        =   24
         Top             =   1860
         Width           =   2115
      End
      Begin VB.CommandButton CmdRetornarCessao 
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
         Left            =   6120
         TabIndex        =   26
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton CmdConfirmarCessao 
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
         Left            =   4980
         TabIndex        =   25
         Top             =   2100
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DtpDtDistratoCessao 
         Height          =   315
         Left            =   6000
         TabIndex        =   23
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   49676291
         CurrentDate     =   37180
      End
      Begin MSDataListLib.DataCombo CboContratoInicial 
         Bindings        =   "FrmContratosImovel.frx":0000
         Height          =   315
         Left            =   1740
         TabIndex        =   30
         Top             =   960
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Parcela"
         BoundColumn     =   "titulo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo CboContratoFinal 
         Bindings        =   "FrmContratosImovel.frx":0021
         Height          =   315
         Left            =   1740
         TabIndex        =   31
         Top             =   1380
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Parcela"
         BoundColumn     =   "titulo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc DatContratoInicial 
         Height          =   330
         Left            =   2280
         Top             =   2160
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
         Caption         =   "DatTitulos"
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
      Begin MSAdodcLib.Adodc DatContratoFinal 
         Height          =   330
         Left            =   2340
         Top             =   2160
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
         Caption         =   "DatTitulos"
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
      Begin VB.Label LblDtContrato 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label1"
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
         Left            =   1740
         TabIndex        =   35
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label LblDataContrato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data do Contrato:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   34
         Top             =   660
         Width           =   1635
      End
      Begin VB.Label LblNomeCliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label1"
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
         Height          =   195
         Left            =   1740
         TabIndex        =   33
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label LblCliente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cliente:"
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
         Left            =   1020
         TabIndex        =   32
         Top             =   360
         Width           =   675
      End
      Begin VB.Label LblCodigoInicial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Título Inicial:"
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
         Left            =   480
         TabIndex        =   29
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label LlbCodigoFinal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Título Final:"
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
         Left            =   600
         TabIndex        =   28
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label LblDtDistratoCessao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data do Distrato:"
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
         Left            =   4500
         TabIndex        =   27
         Top             =   660
         Width           =   1515
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1980
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin Threed.SSPanel PanCodigo 
      Height          =   630
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   9255
      _Version        =   65536
      _ExtentX        =   16325
      _ExtentY        =   1111
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
      Begin VB.Label LblCodImovel 
         Caption         =   "Label1"
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
         Height          =   195
         Left            =   8220
         TabIndex        =   14
         Top             =   180
         Width           =   735
      End
      Begin VB.Label LblNomeEmpreendimento 
         Caption         =   "Label1"
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
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Top             =   180
         Width           =   3735
      End
      Begin VB.Label LblCodEmpreendimento 
         Caption         =   "Label1"
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
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   180
         Width           =   675
      End
      Begin VB.Label LblCodigo 
         Alignment       =   1  'Right Justify
         Caption         =   "Código do Imóvel:"
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
         Left            =   6540
         TabIndex        =   11
         Top             =   180
         Width           =   1605
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
         Left            =   240
         TabIndex        =   10
         Top             =   180
         Width           =   1485
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "FrmContratosImovel.frx":0040
      Height          =   3675
      Left            =   60
      TabIndex        =   15
      Top             =   780
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   6482
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Contrato"
      Columns(0).DataField=   "Contrato"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Status"
      Columns(1).DataField=   "Status"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Cliente"
      Columns(2).DataField=   "Cliente"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Data da Venda"
      Columns(3).DataField=   "cont_dt_Venda"
      Columns(3).NumberFormat=   "dd/mm/yy"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Data do Contrato"
      Columns(4).DataField=   "cont_dt_RegistroContrato"
      Columns(4).NumberFormat=   "dd/mm/yy"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Data Base"
      Columns(5).DataField=   "cont_dt_Base"
      Columns(5).NumberFormat=   "dd/mm/yy"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Data do Distrato"
      Columns(6).DataField=   "cont_dt_Distrato"
      Columns(6).NumberFormat=   "dd/mm/yy"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Data da Quitação"
      Columns(7).DataField=   "cont_dt_Quitacao"
      Columns(7).NumberFormat=   "dd/mm/yy"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8720"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2249"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2170"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=529"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=6509"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6429"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=8720"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1879"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1799"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=529"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=1879"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1799"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=529"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=1799"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1720"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=529"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1852"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1773"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=529"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=1905"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1826"
      Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=529"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bold=-1,.fontsize=825"
      _StyleDefs(21)  =   ":id=14,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(22)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3,.alignment=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.alignment=3"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.valignment=2"
      _StyleDefs(33)  =   ":id=28,.locked=-1"
      _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(63)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(65)  =   "Named:id=33:Normal"
      _StyleDefs(66)  =   ":id=33,.parent=0"
      _StyleDefs(67)  =   "Named:id=34:Heading"
      _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(69)  =   ":id=34,.wraptext=-1"
      _StyleDefs(70)  =   "Named:id=35:Footing"
      _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   "Named:id=36:Selected"
      _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(74)  =   "Named:id=37:Caption"
      _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(76)  =   "Named:id=38:HighlightRow"
      _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(78)  =   "Named:id=39:EvenRow"
      _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(80)  =   "Named:id=40:OddRow"
      _StyleDefs(81)  =   ":id=40,.parent=33"
      _StyleDefs(82)  =   "Named:id=41:RecordSelector"
      _StyleDefs(83)  =   ":id=41,.parent=34"
      _StyleDefs(84)  =   "Named:id=42:FilterBar"
      _StyleDefs(85)  =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   780
      Left            =   300
      TabIndex        =   21
      Top             =   4560
      Width           =   8760
      _Version        =   65536
      _ExtentX        =   15452
      _ExtentY        =   1376
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
      Begin VB.CommandButton CmdTitulo 
         Caption         =   "&Títulos"
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
         Left            =   360
         TabIndex        =   5
         Top             =   420
         Width           =   1755
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
         Left            =   1995
         TabIndex        =   1
         Top             =   0
         Width           =   1455
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
         Left            =   360
         TabIndex        =   0
         Top             =   0
         Width           =   1455
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
         Left            =   6600
         TabIndex        =   8
         Top             =   420
         Width           =   1755
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
         Left            =   3630
         TabIndex        =   2
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton CmdMapaFechamento 
         BackColor       =   &H00000000&
         Caption         =   "&Mapa Fechamento"
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
         Left            =   2440
         TabIndex        =   6
         Top             =   420
         Width           =   1755
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
         Left            =   4520
         TabIndex        =   7
         Top             =   420
         Width           =   1755
      End
      Begin VB.CommandButton CmdCessao 
         Caption         =   "&Cessão"
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
         Left            =   6900
         TabIndex        =   4
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton CmdDistratar 
         Caption         =   "&Distratar"
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
         Left            =   5265
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmManContratosImovel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XFT_CODIMOVEL As String 'Código do imóvel
Dim XFT_CODEMPREENDIMENTO As String 'Código do empreendimento
Dim XFT_CODCONTRATO As String 'Código do Contrato

Sub subAjustaTelaContrato()
    Formulario.Top = 0
    Formulario.Height = Resolucaoy - 1380
    Formulario.Width = ResolucaoX - 60
    Formulario.Left = 5
    Formulario.PanCodigo.Left = 75
    Formulario.PanCodigo.Width = Formulario.Width - 300
    Formulario.TDBGrid1.Top = PanCodigo.Top + PanCodigo.Height + 40
    Formulario.TDBGrid1.Left = 75
    Formulario.TDBGrid1.Width = Formulario.Width - 300
    Formulario.TDBGrid1.Height = Formulario.Height - Formulario.TDBGrid1.Top - Formulario.PanBotoes.Height - 100
    Formulario.PanBotoes.Top = Formulario.TDBGrid1.Top + Formulario.TDBGrid1.Height + 50
    Formulario.PanBotoes.Left = 75 + (Formulario.TDBGrid1.Width - Formulario.PanBotoes.Width) / 2
End Sub

Private Sub CmdAlterar_Click()
  
  Chave = Adodc1.Recordset.Fields("Contrato")
  XInserir = "A"
  FrmManContratos.Show 1
  subCarregaDadosNV Adodc1, "", "Contrato", _
        "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
    
End Sub

Private Sub CmdCessao_Click()
    
  Dim XLT_SQL As String
  
  PanBotoes.Enabled = False
  PanCessao.Left = (TDBGrid1.Width - PanCessao.Width) / 2
  PanCessao.Top = (TDBGrid1.Height - PanCessao.Height) / 2
  TDBGrid1.Enabled = False
  
  DtpDtDistratoCessao.Value = Date
  LblNomeCliente = Adodc1.Recordset.Fields("Cliente")
  LblDtContrato = Format(Adodc1.Recordset.Fields("cont_dt_RegistroContrato"), "dd/mm/yy")
  
  If NomeSgbd = "Access" Then
    XLT_SQL = "SELECT Titulo,IIF(Isnull(titu_dt_Pagamento)=true," & _
      "(Titulo + ' - ' + str(titu_dt_prorrogacao) + ' - ' + tipl_tx_descricao + ' - ' + natureza)," & _
      "(Titulo + ' - ' + str(titu_dt_prorrogacao) + ' - ' + str(titu_dt_Pagamento)" & _
      "+ ' - ' + tipl_tx_descricao + ' - ' + natureza)) AS Parcela" & _
      " FROM ConsCapTitulos" & _
      " WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "'" & _
      " AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "' AND cont_cd_Contrato='" & Right(Adodc1.Recordset.Fields("Contrato"), 2) & "'" & _
      " ORDER BY Titulo"
  Else
    XLT_SQL = "SELECT Titulo,CASE WHEN titu_dt_Pagamento IS NULL THEN " & _
      "(Titulo + ' - ' + CONVERT(varchar(8), titu_dt_prorrogacao, 3) + ' - ' + tipl_tx_descricao + ' - ' + natureza) " & _
      "WHEN titu_dt_Pagamento IS NOT NULL THEN " & _
      "(Titulo + ' - ' + CONVERT(varchar(8), titu_dt_prorrogacao, 3) + ' - ' + CONVERT(varchar(8), titu_dt_Pagamento, 3)" & _
      "+ ' - ' + tipl_tx_descricao + ' - ' + natureza) END AS Parcela" & _
      " FROM ConsCapTitulos" & _
      " WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "'" & _
      " AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "' AND cont_cd_Contrato='" & Right(Adodc1.Recordset.Fields("Contrato"), 2) & "'" & _
      " ORDER BY Titulo"
   End If
  
  subConectarControleDadosNV DatContratoInicial, XLT_SQL, Estatico
  
  If Not DatContratoInicial.Recordset.EOF = True Then
    DatContratoInicial.Recordset.MoveFirst
    CboContratoInicial.BoundText = DatContratoInicial.Recordset.Fields("Titulo")
  
    subConectarControleDadosNV DatContratoFinal, XLT_SQL, Estatico
    DatContratoFinal.Recordset.MoveLast
    CboContratoFinal.BoundText = DatContratoFinal.Recordset.Fields("Titulo")
  
    ChkTitulosPagos.Value = 0
    PanCessao.Visible = True
    PanCessao.Enabled = True
    PanCessao.Tag = "C"
    XFT_CODCONTRATO = Right$(Adodc1.Recordset.Fields("Contrato"), 2)
    DtpDtDistratoCessao.SetFocus
  Else
    MsgBox "Esse contrato não possui nenhum título", vbCritical, "ATENÇÃO !"
    PanBotoes.Enabled = True
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
  End If
  
End Sub

Private Sub CmdConfirmar_Click()
  
  XFT_CODCONTRATO = Right$(Adodc1.Recordset.Fields("Contrato"), 2)
  
  If DtpDtDistrato.Value < Adodc1.Recordset.Fields("cont_dt_RegistroContrato") Then
    MsgBox "A Data do Distrato deve se maior ou igual que a Data do Contrato.", vbCritical, "ATENÇÃO !"
    Exit Sub
  End If
  
  Conexao.BeginTrans
  Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas -1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
  Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='D' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
  Conexao.Execute "UPDATE Contratos SET cont_tx_Status='I',cont_dt_Distrato=" & FunNuloData(DtpDtDistrato.Value, NomeSgbd) & " WHERE empr_cd_Empresa=" & PCodEmpresa & " AND cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
  
  Conexao.CommitTrans
  subCarregaDadosNV Adodc1, "", "Contrato", _
       "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
  Call CmdRetornar_Click
    
End Sub

Private Sub CmdConfirmarCessao_Click()
    
  Dim XLO_CONTRATO As New ADODB.Recordset
  Dim XLO_TITULO As New ADODB.Recordset
  Dim XLT_NOVOCODCONTRATO As String
  Dim XLF_TITULOINICIAL As String
  Dim XLF_TITULOFINAL As String
  Dim XLT_SQL As String
  
  Dim XLT_CODIMOVEL As String, XLT_CODEMPREENDIMENTO As String
  Dim XLT_CODCONTRATO As String, XLT_CODPLANO As String
  Dim XLT_CODPARCELA As String, XLT_CODRESIDUO As String
    
  If DtpDtDistratoCessao.Value < Adodc1.Recordset.Fields("cont_dt_RegistroContrato") Then
    MsgBox "A Data do Distrato deve se maior ou igual que a Data do Contrato.", vbCritical, "ATENÇÃO !"
    Exit Sub
  End If
  
  'Inativa o contrato que será cedido para poder ser criado
   'um contrato ativo
  
  Conexao.BeginTrans
  
  Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas -1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
  Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='D' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
  Conexao.Execute "UPDATE Contratos SET cont_tx_Status='I',cont_dt_Distrato=" & FunNuloData(DtpDtDistratoCessao.Value, NomeSgbd) & "  WHERE empr_cd_Empresa=" & PCodEmpresa & " AND cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
  
  Chave = XFT_CODEMPREENDIMENTO + "." + XFT_CODIMOVEL
  XInserir = "C"
  XOrigem = "PanCessao"
  FrmManContratos.Show 1
  XOrigem = ""
  
  'Conexao.CommitTrans
  
  'Codigo do novo contrato a ser criado
  'XLT_NOVOCODCONTRATO = Format(CStr(CInt(XFT_CODCONTRATO) + 1), "00")
  XLT_NOVOCODCONTRATO = Chave
  
  'Verifica se o contrato foi realmente criado e com status ATIVO
  SubQOpenRecordset XLO_CONTRATO, "SELECT * FROM Contratos WHERE " & _
            "empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento= '" & XFT_CODEMPREENDIMENTO & "' " & _
            "AND imov_cd_Imovel='" & XFT_CODIMOVEL & "' AND cont_cd_Contrato ='" & XLT_NOVOCODCONTRATO & "' " & _
            "AND cont_tx_Status='A' ", "Estatico"
            
    If XLO_CONTRATO.EOF Then
    Conexao.RollbackTrans
    'Ativa o contrato novamente, pois não foi criado um contrato ativo
   ' Conexao.BeginTrans
   ' Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas +1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
   ' Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='V' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
   ' Conexao.Execute "UPDATE Contratos SET cont_tx_Status='A' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
   ' Conexao.CommitTrans
    MsgBox "Um contrato ATIVO não foi criado para herdar os títulos.", vbCritical, "ATENÇÃO !"
    XLO_CONTRATO.Close
    Set XLO_CONTRATO = Nothing
    Exit Sub
  Else
    
    Conexao.CommitTrans
    XLF_TITULOINICIAL = XFT_CODEMPREENDIMENTO + "." + XFT_CODIMOVEL + "." + XFT_CODCONTRATO + "." + CboContratoInicial.BoundText
    XLF_TITULOFINAL = XFT_CODEMPREENDIMENTO + "." + XFT_CODIMOVEL + "." + XFT_CODCONTRATO + "." + CboContratoFinal.BoundText
    
    
    If ChkTitulosPagos.Value = 0 Then
      SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCAPTitulos WHERE Titulo >= '" & XLF_TITULOINICIAL & "' AND Titulo <= '" & XLF_TITULOFINAL & "' AND titu_dt_Pagamento is Null ORDER BY Titulo", Estatico
    Else
      SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCAPTitulos WHERE Titulo >= '" & XLF_TITULOINICIAL & "' AND Titulo <= '" & XLF_TITULOFINAL & "' ORDER BY Titulo", Estatico
    End If
        
    While Not XLO_TITULO.EOF
      
      XLT_CODPLANO = Mid$(XLO_TITULO!Titulo, 14, 2)
      XLT_CODPARCELA = Mid$(XLO_TITULO!Titulo, 17, 3)
      XLT_CODRESIDUO = Right$(XLO_TITULO!Titulo, 2)
      
      Conexao.BeginTrans
      
      
      XLT_SQL = "INSERT INTO Titulos (empr_cd_Empresa, empd_cd_Empreendimento, imov_cd_Imovel,cont_cd_Contrato, " & _
        "titu_cd_Plano,titu_cd_Parcela,titu_cd_Residuo,obse_cd_Observacao, " & _
        "moed_cd_Moeda1, moed_cd_Moeda2, titu_vl_Parcela, " & _
        "titu_vl_ParcelaIndex1, titu_vl_ParcelaIndex2, titu_dt_Vencimento, tipl_cd_TipoPlano, " & _
        "titu_vl_PercJurosFin, titu_vl_PercMultaMora, titu_vl_PercJurosMora, " & _
        "cere_cd_Pcr,plco_cd_Conta,coco_cd_Codigo,titu_dt_Base," & _
        "titu_dt_Prorrogacao,titu_tx_NaturezaPlano,titu_vl_ValorPago, " & _
        "titu_dt_Pagamento,titu_vl_Multa,titu_vl_Juros,titu_vl_Seguro, " & _
        "titu_vl_Outros,titu_vl_Desconto,titu_vl_Desagio,titu_tx_Financia," & _
        "titu_tx_IndicePagamento,titu_dt_Deposito,titu_dt_Repasse, " & _
        "titu_tx_LocalPagto,titu_tx_DocPagto,titu_tx_TipoPag, " & _
        "titu_tx_CorrigeParcela, titu_tx_PreDatado, titu_dt_Criacao, " & _
        "titu_dt_BasePagto,moed_cd_Moeda3) "
   
   XLT_SQL = XLT_SQL + "VALUES (" & PCodEmpresa & ",'" & XFT_CODEMPREENDIMENTO & "','" & XFT_CODIMOVEL & "','" & XLT_NOVOCODCONTRATO & "'," & _
        "'" & XLT_CODPLANO & "','" & XLT_CODPARCELA & "','" & XLT_CODRESIDUO & "'," & FunNuloBancoVal(XLO_TITULO!obse_cd_Observacao) & ", " & _
        "" & FunNuloBancoVal(XLO_TITULO!moed_cd_Moeda1) & "," & FunNuloBancoVal(XLO_TITULO!moed_cd_Moeda2) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Parcela)) & "," & _
        "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_ParcelaIndex1)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_ParcelaIndex2)) & "," & FunNuloData(XLO_TITULO!titu_dt_Vencimento, NomeSgbd) & "," & XLO_TITULO!tipl_cd_TipoPlano & "," & _
        "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_PercJurosFin)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_PercMultaMora)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_PercJurosMora)) & ", " & _
        "" & FunNuloBanco(XLO_TITULO!cere_cd_Pcr) & "," & FunNuloBanco(XLO_TITULO!plco_cd_Conta) & "," & FunNuloBancoVal(XLO_TITULO!coco_cd_Codigo) & ", " & FunNuloData(XLO_TITULO!titu_dt_Base, NomeSgbd) & "," & _
        "" & FunNuloData(XLO_TITULO!titu_dt_Prorrogacao, NomeSgbd) & ",'" & XLO_TITULO!titu_tx_NaturezaPlano & "'," & _
        "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_ValorPago)) & "," & FunNuloData(XLO_TITULO!titu_dt_Pagamento, NomeSgbd) & "," & _
        "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Multa)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Juros)) & "," & _
        "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Seguro)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Outros)) & "," & _
        "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Desconto)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_Desagio)) & "," & _
        "'" & XLO_TITULO!titu_tx_Financia & "','" & XLO_TITULO!titu_tx_IndicePagamento & "'," & FunNuloData(XLO_TITULO!titu_dt_Deposito, NomeSgbd) & "," & _
        "" & FunNuloData(XLO_TITULO!titu_dt_Repasse, NomeSgbd) & ",'" & XLO_TITULO!titu_tx_LocalPagto & "','" & XLO_TITULO!titu_tx_DocPagto & "'," & _
        "'" & XLO_TITULO!titu_tx_TipoPag & "','" & XLO_TITULO!titu_tx_CorrigeParcela & "','" & XLO_TITULO!titu_tx_PreDatado & "'," & FunNuloData(XLO_TITULO!titu_dt_Criacao, NomeSgbd) & "," & _
        "" & FunNuloData(XLO_TITULO!titu_dt_BasePagto, NomeSgbd) & "," & FunNuloBancoVal(XLO_TITULO!moed_cd_Moeda3) & ") "
        
      Conexao.Execute (XLT_SQL)
      
      Conexao.CommitTrans
      XLO_TITULO.MoveNext
      
    Wend
    
    Conexao.Execute ("UPDATE Contratos SET cont_tx_Status='C' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'")
  
  End If
  
  XLO_TITULO.Close
  Set XLO_TITULO = Nothing
  
  subCarregaDadosNV Adodc1, "", "Contrato", _
       "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
  Call CmdRetornarCessao_Click
    
End Sub

Private Sub CmdRetornar_Click()
  
  PanDistratar.Visible = False
  PanDistratar.Enabled = False
  PanBotoes.Enabled = True
  'SubHabilitaBotoes
  TDBGrid1.Enabled = True
  TDBGrid1.SetFocus
  
End Sub

Private Sub CmdDistratar_Click()
  
  Dim XLO_CONTRATO As New ADODB.Recordset
  
  'Verifica se vai distratar ou ativar
  If CmdDistratar.Caption = "&Distratar" Then
    PanBotoes.Enabled = False
    PanDistratar.Left = (TDBGrid1.Width - PanDistratar.Width) / 2
    PanDistratar.Top = (TDBGrid1.Height - PanDistratar.Height) / 2
    TDBGrid1.Enabled = False
    DtpDtDistrato.Value = Date
    PanDistratar.Visible = True
    PanDistratar.Enabled = True
    PanDistratar.Tag = "D"
    DtpDtDistrato.SetFocus
  
  'Ativa um contrato
  Else
    If MsgBox("Confirma Ativação do Contrato ?", 273, "ATENÇÂO") = vbOK Then
    
      XFT_CODCONTRATO = Right$(Adodc1.Recordset.Fields("Contrato"), 2)
       
      SubQOpenRecordset XLO_CONTRATO, "SELECT cont_tx_status,cont_cd_Contrato FROM Contratos WHERE " & _
           "empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento= '" & XFT_CODEMPREENDIMENTO & "' " & _
           "AND imov_cd_Imovel='" & XFT_CODIMOVEL & "' AND (cont_tx_Status ='A' OR cont_tx_Status ='R')", "Estatico"
       
      If Not (XLO_CONTRATO.BOF And XLO_CONTRATO.EOF) Then
       
        'Verifica se já não existe um contrato ATIVO
        If XLO_CONTRATO!cont_tx_Status = "A" Then
          MsgBox "Esse imóvel já está vendido.", vbCritical, "ATENÇÃO !"
          Exit Sub
        End If
          
        'Verifica se já não existe um contrato RESERVADO
          'quando o  contrato que estiver sendo Ativado for INATIVO
        If Adodc1.Recordset.Fields("Status") = "INATIVO" Then
          If XLO_CONTRATO!cont_tx_Status = "R" Then
            MsgBox "Esse imóvel já está reservado.", vbCritical, "ATENÇÃO !"
            Exit Sub
          End If
          
        'Altera de RESERVADO para ATIVO, chama o formulário de contratos
        Else
          Chave = Adodc1.Recordset.Fields("Contrato")
          XInserir = "A"
          FrmManContratos.Show 1
          subCarregaDadosNV Adodc1, "", "Contrato", _
             "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
        End If
            
      'Altera de INATIVO para ATIVO
      Else
        Conexao.BeginTrans
        Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas +1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
        Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='V' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
        Conexao.Execute "UPDATE Contratos SET cont_tx_Status='A',cont_dt_Distrato=Null WHERE empr_cd_Empresa=" & PCodEmpresa & " AND cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
        Conexao.CommitTrans
  
        subCarregaDadosNV Adodc1, "", "Contrato", _
             "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
      End If
       
      XLO_CONTRATO.Close
      Set XLO_CONTRATO = Nothing
    End If
    
  End If

End Sub

Private Sub CmdExcluir_Click()
    
  Dim XLI_HOUVEERRO As Integer
  Dim XLT_STATUS As String 'Pega o Status do contrato
  Dim XLO_CONTRATO As New ADODB.Recordset
  
  XLI_HOUVEERRO = False
  
  'SubDesabilitaBotoes
     
  SubQOpenRecordset XLO_CONTRATO, "SELECT * FROM Titulos WHERE" & _
    " empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_Empreendimento= '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "'" & _
    " AND imov_cd_imovel ='" & Adodc1.Recordset.Fields("imov_cd_imovel") & "' AND cont_cd_Contrato ='" & Adodc1.Recordset.Fields("cont_cd_Contrato") & "'", Estatico
  
  If Not XLO_CONTRATO.EOF Then
    While Not XLO_CONTRATO.EOF
      If Not IsNull(XLO_CONTRATO!titu_dt_Pagamento) Then
        MsgBox "Esse contrato não poderá ser excluído pois contém Títulos pagos .", vbCritical, "ATENÇÃO"
        XLO_CONTRATO.Close
        Set XLO_CONTRATO = Nothing
        Exit Sub
      End If
      XLO_CONTRATO.MoveNext
    Wend
    'Else
      If MsgBox("Confirma Remoção do Contrato e seus Títulos?", 273, "CUIDADO") = vbOK Then
        Conexao.BeginTrans
        Conexao.Execute ("DELETE FROM Titulos WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_Empreendimento= '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "' AND imov_cd_imovel ='" & Adodc1.Recordset.Fields("imov_cd_imovel") & "' AND cont_cd_Contrato ='" & Adodc1.Recordset.Fields("cont_cd_Contrato") & "'")
        Conexao.Execute "DELETE FROM Contratos WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_Empreendimento= '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "' AND imov_cd_imovel ='" & Adodc1.Recordset.Fields("imov_cd_imovel") & "' AND cont_cd_Contrato ='" & Adodc1.Recordset.Fields("cont_cd_Contrato") & "'"
        
        XLT_STATUS = Adodc1.Recordset.Fields("cont_tx_Status")
        'Se o contrato deletado tiver o status ativo,
        'o status do imovel se torna disponível
        If XLT_STATUS = "A" Or XLT_STATUS = "R" Then
          Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas -1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
          Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='D' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
        End If
        Conexao.CommitTrans
      End If
    'End If
  Else
    XLT_STATUS = Adodc1.Recordset.Fields("cont_tx_Status")
     
    If MsgBox("Confirma Remoção do Contrato e seus Títulos?", 273, "CUIDADO") = vbOK Then
     
      Conexao.BeginTrans
      Conexao.Execute "DELETE FROM Contratos WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_Empreendimento= '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "' AND imov_cd_imovel ='" & Adodc1.Recordset.Fields("imov_cd_imovel") & "' AND cont_cd_Contrato ='" & Adodc1.Recordset.Fields("cont_cd_Contrato") & "'"
      
      'Se o contrato deletado tiver o status ativo,
      'o status do imovel se torna disponível
      If XLT_STATUS = "A" Or XLT_STATUS = "R" Then
        Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas -1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
        Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='D' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
      End If
      Conexao.CommitTrans
    End If
  End If
      
  If XLI_HOUVEERRO = False Then
     subCarregaDadosNV Adodc1, "", "Contrato", _
     "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
     If ContFil <> 0 Then
        subContaReg Adodc1
     End If
  End If
  
  'SubHabilitaBotoes
  
  If FunTabelaVazia(Adodc1) = True Then
     CmdAlterar.Enabled = False
     CmdExcluir.Enabled = False
     CmdDistratar.Enabled = False
     CmdImprimir.Enabled = False
     CmdTitulo.Enabled = False
     CmdMapaFechamento.Enabled = False
     CmdCessao.Enabled = False
  End If
  
  XLO_CONTRATO.Close
  Set XLO_CONTRATO = Nothing
  
  TDBGrid1.SetFocus
  'TDBGrid1.Col = 0
  
  Exit Sub
    
RotuloErro:
    XLI_HOUVEERRO = True
    funTrataErros (ComMensagem)
End Sub


Private Sub CmdImprimir_Click()
  Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Contratos")
End Sub

Private Sub CmdInserir_Click()
    Dim XLO_CONTRATO As New ADODB.Recordset

    SubQOpenRecordset XLO_CONTRATO, "SELECT cont_tx_Status FROM Contratos WHERE " & _
      "empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento= '" & XFT_CODEMPREENDIMENTO & "' " & _
      "AND imov_cd_Imovel='" & XFT_CODIMOVEL & "' AND cont_tx_Status = 'A'", "Estatico"
    If Not (XLO_CONTRATO.BOF And XLO_CONTRATO.EOF) Then
        MsgBox "Esse imóvel possui um contrato ATIVO.", vbCritical, "ATENÇÃO !"
    Else
      Chave = XFT_CODEMPREENDIMENTO + "." + XFT_CODIMOVEL
      XInserir = "I"
      FrmManContratos.Show 1
      subCarregaDadosNV Adodc1, "", "Contrato", _
          "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
    End If
    XLO_CONTRATO.Close
    If Not Adodc1.Recordset.EOF Then
      CmdAlterar.Enabled = True
      CmdExcluir.Enabled = True
      CmdImprimir.Enabled = True
    End If

End Sub

Private Sub CmdMapaFechamento_Click()
  
  Dim XLO_CONTRATO As New ADODB.Recordset
    
    Chave = Adodc1.Recordset.Fields("Contrato")
    XInserir = "I"
    FrmManMapaPlanos.Show 1
  
End Sub

Private Sub CmdRetornarCessao_Click()
  
  PanCessao.Visible = False
  PanCessao.Enabled = False
  PanBotoes.Enabled = True
  TDBGrid1.Enabled = True
  TDBGrid1.SetFocus
  
End Sub

Private Sub CmdSair_Click()
   
   Unload Me
   
End Sub

Private Sub CmdTitulo_Click()
  
  XOrigem = "FrmManContratosImovel"
  Chave = Adodc1.Recordset.Fields("Contrato")
  TelaFinTitulos.Show
  If FunLocalizaJanelasAtivas("TelaFinTitulos") <> -1 Then
      TelaFinTitulos.SetFocus
  End If
  
End Sub

Private Sub Form_Activate()
  
  Set Formulario = FrmManContratosImovel
  XOrigem = "FrmManContratosImovel"
  
End Sub

Private Sub Form_Load()

    Dim XLO_IMOVEL As New ADODB.Recordset
    Dim XLB_ACHOU As Boolean
    
    Set Formulario = FrmManContratosImovel
    XFT_CODEMPREENDIMENTO = Left$(Chave, 4)
    XFT_CODIMOVEL = Right$(Chave, 4)
 
    subCarregaDadosNV Adodc1, "", "Contrato", _
                            "", "SELECT * FROM ConsCAPContratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "'"
    SubQOpenRecordset XLO_IMOVEL, "SELECT empd_tx_Nome FROM Empreendimentos WHERE empd_cd_Empreendimento ='" & XFT_CODEMPREENDIMENTO & "' AND empr_cd_empresa = " & PCodEmpresa, Estatico
    If Not XLO_IMOVEL.EOF Then
      LblNomeEmpreendimento.Caption = XLO_IMOVEL!empd_tx_Nome
      LblCodEmpreendimento.Caption = XFT_CODEMPREENDIMENTO
      LblCodImovel.Caption = XFT_CODIMOVEL
    End If
    XLO_IMOVEL.Close
    Set XLO_IMOVEL = Nothing
    
    If Adodc1.Recordset.EOF Then
      CmdAlterar.Enabled = False
      CmdExcluir.Enabled = False
      CmdCessao.Enabled = False
      CmdDistratar.Enabled = False
      CmdTitulo.Enabled = False
      CmdMapaFechamento.Enabled = False
      CmdImprimir.Enabled = False
    Else
    
      XLB_ACHOU = False
      While (Not Adodc1.Recordset.EOF) And (XLB_ACHOU = False)
        If Adodc1.Recordset("cont_tx_Status") = "A" Then
          XLB_ACHOU = True
        Else
          Adodc1.Recordset.MoveNext
        End If
      Wend
      If XLB_ACHOU = False Then
        Adodc1.Recordset.MoveFirst
      End If
    
    End If
    
    Call subAjustaTelaContrato
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Chave = XFT_CODEMPREENDIMENTO + "." + XFT_CODIMOVEL
   Set FrmManContratosImovel = Nothing
   
End Sub

Private Sub TDBGrid1_DblClick()
  If Adodc1.Recordset.RecordCount <> 0 Then
    CmdAlterar_Click
  End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  
  If Not Adodc1.Recordset.EOF Then
  
    If Adodc1.Recordset.Fields("Status") = "ATIVO" Then
      CmdDistratar.Caption = "&Distratar"
      CmdDistratar.Enabled = True
      CmdCessao.Enabled = True
      CmdAlterar.Enabled = True
      CmdMapaFechamento.Enabled = True
      CmdTitulo.Enabled = True
    ElseIf Adodc1.Recordset.Fields("Status") = "INATIVO" Then
      CmdDistratar.Caption = "&Ativar"
      CmdDistratar.Enabled = True
      CmdCessao.Enabled = -False
      CmdAlterar.Enabled = True
      CmdMapaFechamento.Enabled = True
      CmdTitulo.Enabled = True
    ElseIf Adodc1.Recordset.Fields("Status") = "RESERVADO" Then
      CmdDistratar.Caption = "&Ativar"
      CmdDistratar.Enabled = True
      CmdCessao.Enabled = False
      CmdAlterar.Enabled = True
      CmdMapaFechamento.Enabled = False
      CmdTitulo.Enabled = False
    ElseIf Adodc1.Recordset.Fields("Status") = "CEDIDO" Then
      CmdDistratar.Enabled = False
      CmdCessao.Enabled = False
      CmdAlterar.Enabled = False
      CmdMapaFechamento.Enabled = False
      CmdTitulo.Enabled = False
    End If
  
  End If
  
End Sub

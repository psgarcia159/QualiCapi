VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmTransfBancaria 
   Caption         =   "Transferência Bancária"
   ClientHeight    =   6705
   ClientLeft      =   225
   ClientTop       =   2295
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkNaoSaiRelFinanceiro 
      Caption         =   "Não incluir nos Relatórios Financeiros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   7080
      TabIndex        =   39
      Top             =   6240
      Width           =   2220
   End
   Begin VB.CheckBox ChkComplemento 
      Caption         =   "Com Complemento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   240
      TabIndex        =   22
      Top             =   6240
      Width           =   2070
   End
   Begin TabDlg.SSTab TabTransfBancaria 
      Height          =   5820
      Left            =   105
      TabIndex        =   0
      Top             =   285
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   10266
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Transferência"
      TabPicture(0)   =   "FrmTransfBancaria.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraGeral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Complemento"
      TabPicture(1)   =   "FrmTransfBancaria.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PanDistribuicao"
      Tab(1).Control(1)=   "TDBGrid3"
      Tab(1).Control(2)=   "CmdInsereComp"
      Tab(1).Control(3)=   "CmdRemoverComp"
      Tab(1).Control(4)=   "CmdAlteraComp"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(6)=   "LblTotalComplemento"
      Tab(1).Control(7)=   "LblTotalComplemento2"
      Tab(1).ControlCount=   8
      Begin Threed.SSPanel PanDistribuicao 
         Height          =   2850
         Left            =   -74565
         TabIndex        =   42
         Top             =   1440
         Visible         =   0   'False
         Width           =   8580
         _Version        =   65536
         _ExtentX        =   15134
         _ExtentY        =   5027
         _StockProps     =   15
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
         Begin VB.ComboBox CboStatusCompl 
            Height          =   315
            ItemData        =   "FrmTransfBancaria.frx":0038
            Left            =   5370
            List            =   "FrmTransfBancaria.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   255
            Width           =   3090
         End
         Begin VB.ComboBox CboTipoPagCompl 
            Height          =   315
            ItemData        =   "FrmTransfBancaria.frx":0060
            Left            =   1575
            List            =   "FrmTransfBancaria.frx":0079
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   675
            Width           =   3090
         End
         Begin VB.TextBox TxtVlCompl 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   47
            Top             =   255
            Width           =   1335
         End
         Begin VB.TextBox TxtDocPagtoCompl 
            Height          =   315
            Left            =   6855
            MaxLength       =   50
            TabIndex        =   50
            Top             =   675
            Width           =   1590
         End
         Begin VB.CommandButton Botao_Conf 
            BackColor       =   &H00000000&
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
            Left            =   6150
            TabIndex        =   58
            Top             =   2400
            Width           =   1110
         End
         Begin VB.CommandButton CmdSairDistribuicao 
            BackColor       =   &H00000000&
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
            Left            =   7350
            TabIndex        =   60
            Top             =   2400
            Width           =   1110
         End
         Begin MSDataListLib.DataCombo CboHistCompl 
            Bindings        =   "FrmTransfBancaria.frx":00D3
            Height          =   315
            Left            =   1575
            TabIndex        =   56
            Top             =   1965
            Width           =   6480
            _ExtentX        =   11430
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "hist_tx_descricao"
            BoundColumn     =   "hist_cd_codigo"
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MskConta2 
            Height          =   330
            Left            =   1575
            TabIndex        =   53
            Top             =   1530
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo CboPlanoContas2 
            Bindings        =   "FrmTransfBancaria.frx":00EE
            Height          =   330
            Left            =   2505
            TabIndex        =   54
            Top             =   1530
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "plco_cd_codreduzido"
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
         Begin Threed.SSCommand CmdLimparHistorico2 
            Height          =   330
            Left            =   8085
            TabIndex        =   55
            Top             =   1950
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   582
            _StockProps     =   78
            Enabled         =   0   'False
            MouseIcon       =   "FrmTransfBancaria.frx":0105
            Picture         =   "FrmTransfBancaria.frx":0557
         End
         Begin MSMask.MaskEdBox MskPcr2 
            Height          =   330
            Left            =   1575
            TabIndex        =   51
            Top             =   1095
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboPcr2 
            Bindings        =   "FrmTransfBancaria.frx":0669
            Height          =   330
            Left            =   2520
            TabIndex        =   52
            Top             =   1095
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "cere_cd_estruturado"
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
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Status:"
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
            Left            =   4680
            TabIndex        =   65
            Top             =   285
            Width           =   645
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tipo Pagamento:"
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
            Left            =   45
            TabIndex        =   64
            Top             =   705
            Width           =   1485
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Centro Custo:"
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
            Left            =   60
            TabIndex        =   63
            Top             =   1140
            Width           =   1470
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Conta Contábil:"
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
            Left            =   45
            TabIndex        =   62
            Top             =   1575
            Width           =   1485
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Valor:"
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
            Left            =   690
            TabIndex        =   61
            Top             =   285
            Width           =   840
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Histórico:"
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
            Left            =   675
            TabIndex        =   59
            Top             =   1995
            Width           =   855
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Doc. Pagto.:"
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
            Left            =   5685
            TabIndex        =   57
            Top             =   705
            Width           =   1110
         End
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid3 
         Height          =   2910
         Left            =   -74880
         TabIndex        =   43
         Top             =   1410
         Width           =   9060
         _ExtentX        =   15981
         _ExtentY        =   5133
         _LayoutType     =   4
         _RowHeight      =   20
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "CodPcr"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Centro Custo"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "CodConta"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Conta Contábil"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Histórico"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   80
         Columns(5)._MaxComboItems=   5
         Columns(5).ValueItems(0)._DefaultItem=   0
         Columns(5).ValueItems(0).Value=   "C"
         Columns(5).ValueItems(0).Value.vt=   8
         Columns(5).ValueItems(0).DisplayValue=   "Compromissado"
         Columns(5).ValueItems(0).DisplayValue.vt=   8
         Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(1)._DefaultItem=   0
         Columns(5).ValueItems(1).Value=   "R"
         Columns(5).ValueItems(1).Value.vt=   8
         Columns(5).ValueItems(1).DisplayValue=   "Realizado"
         Columns(5).ValueItems(1).DisplayValue.vt=   8
         Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems.Count=   2
         Columns(5).Caption=   "Status"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   80
         Columns(6)._MaxComboItems=   5
         Columns(6).ValueItems(0)._DefaultItem=   0
         Columns(6).ValueItems(0).Value=   "1"
         Columns(6).ValueItems(0).Value.vt=   8
         Columns(6).ValueItems(0).DisplayValue=   "Cheque"
         Columns(6).ValueItems(0).DisplayValue.vt=   8
         Columns(6).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(6).ValueItems(1)._DefaultItem=   0
         Columns(6).ValueItems(1).Value=   "2"
         Columns(6).ValueItems(1).Value.vt=   8
         Columns(6).ValueItems(1).DisplayValue=   "Borderô 1"
         Columns(6).ValueItems(1).DisplayValue.vt=   8
         Columns(6).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(6).ValueItems(2)._DefaultItem=   0
         Columns(6).ValueItems(2).Value=   "3"
         Columns(6).ValueItems(2).Value.vt=   8
         Columns(6).ValueItems(2).DisplayValue=   "Borderô 2"
         Columns(6).ValueItems(2).DisplayValue.vt=   8
         Columns(6).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(6).ValueItems(3)._DefaultItem=   0
         Columns(6).ValueItems(3).Value=   "4"
         Columns(6).ValueItems(3).Value.vt=   8
         Columns(6).ValueItems(3).DisplayValue=   "Débito em Conta"
         Columns(6).ValueItems(3).DisplayValue.vt=   8
         Columns(6).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
         Columns(6).ValueItems(4)._DefaultItem=   0
         Columns(6).ValueItems(4).Value=   "5"
         Columns(6).ValueItems(4).Value.vt=   8
         Columns(6).ValueItems(4).DisplayValue=   "Crédito em Conta"
         Columns(6).ValueItems(4).DisplayValue.vt=   8
         Columns(6).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
         Columns(6).ValueItems(5)._DefaultItem=   0
         Columns(6).ValueItems(5).Value=   "6"
         Columns(6).ValueItems(5).Value.vt=   8
         Columns(6).ValueItems(5).DisplayValue=   "Dinheiro"
         Columns(6).ValueItems(5).DisplayValue.vt=   8
         Columns(6).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
         Columns(6).ValueItems(6)._DefaultItem=   0
         Columns(6).ValueItems(6).Value=   "7"
         Columns(6).ValueItems(6).Value.vt=   8
         Columns(6).ValueItems(6).DisplayValue=   "7"
         Columns(6).ValueItems(6).DisplayValue.vt=   8
         Columns(6).ValueItems(6)._PropDict=   "_DefaultItem,517,2"
         Columns(6).ValueItems.Count=   7
         Columns(6).Caption=   "Tipo Pagto"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Doc. Pagto"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Valor"
         Columns(8).DataField=   ""
         Columns(8).NumberFormat=   "Standard"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).Locked=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1048848"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(8)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(9)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(10)=   "Column(1).Width=2302"
         Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=2223"
         Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=1048848"
         Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(17)=   "Column(2).Width=2170"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2090"
         Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=1049104"
         Splits(0)._ColumnProps(22)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(23)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(24)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(25)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(26)=   "Column(3).Width=2328"
         Splits(0)._ColumnProps(27)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(3)._WidthInPix=2249"
         Splits(0)._ColumnProps(29)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(3)._ColStyle=1057040"
         Splits(0)._ColumnProps(31)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(32)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(33)=   "Column(4).Width=3281"
         Splits(0)._ColumnProps(34)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(4)._WidthInPix=3201"
         Splits(0)._ColumnProps(36)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(37)=   "Column(4)._ColStyle=1057296"
         Splits(0)._ColumnProps(38)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(39)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(40)=   "Column(5).Width=1852"
         Splits(0)._ColumnProps(41)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(5)._WidthInPix=1773"
         Splits(0)._ColumnProps(43)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(44)=   "Column(5)._ColStyle=1057296"
         Splits(0)._ColumnProps(45)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(46)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(47)=   "Column(5)._MinWidth=33"
         Splits(0)._ColumnProps(48)=   "Column(6).Width=1799"
         Splits(0)._ColumnProps(49)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(50)=   "Column(6)._WidthInPix=1720"
         Splits(0)._ColumnProps(51)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(52)=   "Column(6)._ColStyle=1057296"
         Splits(0)._ColumnProps(53)=   "Column(6).WrapText=1"
         Splits(0)._ColumnProps(54)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(55)=   "Column(6)._MinWidth=86574240"
         Splits(0)._ColumnProps(56)=   "Column(7).Width=1746"
         Splits(0)._ColumnProps(57)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(7)._WidthInPix=1667"
         Splits(0)._ColumnProps(59)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(60)=   "Column(7)._ColStyle=1057296"
         Splits(0)._ColumnProps(61)=   "Column(7).WrapText=1"
         Splits(0)._ColumnProps(62)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(63)=   "Column(7)._MinWidth=86574240"
         Splits(0)._ColumnProps(64)=   "Column(8).Width=1879"
         Splits(0)._ColumnProps(65)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(66)=   "Column(8)._WidthInPix=1799"
         Splits(0)._ColumnProps(67)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(68)=   "Column(8)._ColStyle=1057298"
         Splits(0)._ColumnProps(69)=   "Column(8).WrapText=1"
         Splits(0)._ColumnProps(70)=   "Column(8).Order=9"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1,5
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=48,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=43"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=44"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=43"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=44"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=45"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=47"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=43,.alignment=0"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=44,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=45"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=47"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=43,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=44"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=45"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=47"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=24,.parent=43,.alignment=0,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=21,.parent=44,.alignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=22,.parent=45"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=23,.parent=47"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(57)  =   ":id=28,.locked=-1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=45"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=47"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=20,.parent=43,.locked=-1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=44,.alignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=45"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=47"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=32,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(66)  =   ":id=32,.locked=-1"
         _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=45"
         _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=47"
         _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=16,.parent=43,.alignment=1,.locked=-1"
         _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=13,.parent=44,.alignment=2"
         _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=14,.parent=45"
         _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=15,.parent=47"
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
         _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(87)  =   "Named:id=39:EvenRow"
         _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(89)  =   "Named:id=40:OddRow"
         _StyleDefs(90)  =   ":id=40,.parent=33"
         _StyleDefs(91)  =   "Named:id=41:RecordSelector"
         _StyleDefs(92)  =   ":id=41,.parent=34"
         _StyleDefs(93)  =   "Named:id=42:FilterBar"
         _StyleDefs(94)  =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton CmdInsereComp 
         Caption         =   "Ins&erir"
         Height          =   315
         Left            =   -74910
         TabIndex        =   46
         Top             =   4425
         Width           =   1050
      End
      Begin VB.CommandButton CmdRemoverComp 
         Caption         =   "Excl&uir"
         Height          =   315
         Left            =   -72660
         TabIndex        =   45
         Top             =   4425
         Width           =   1050
      End
      Begin VB.CommandButton CmdAlteraComp 
         Caption         =   "&Al&terar"
         Height          =   315
         Left            =   -73785
         TabIndex        =   44
         Top             =   4425
         Width           =   1050
      End
      Begin VB.Frame Frame2 
         Height          =   690
         Left            =   -74895
         TabIndex        =   23
         Top             =   570
         Width           =   2355
         Begin VB.OptionButton OptCred 
            Caption         =   "Crédito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   165
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton OptDeb 
            Caption         =   "Débito"
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
            Left            =   1260
            TabIndex        =   25
            Top             =   255
            Width           =   960
         End
      End
      Begin VB.Frame FraGeral 
         Height          =   5205
         Left            =   120
         TabIndex        =   2
         Top             =   420
         Width           =   8985
         Begin VB.TextBox TxtNumCarta 
            Height          =   315
            Left            =   1845
            MaxLength       =   50
            TabIndex        =   71
            Top             =   5280
            Visible         =   0   'False
            Width           =   2835
         End
         Begin VB.TextBox txtReferencia 
            Height          =   735
            Left            =   1845
            MaxLength       =   400
            MultiLine       =   -1  'True
            TabIndex        =   69
            Top             =   4320
            Width           =   6855
         End
         Begin Threed.SSFrame FraBcOrigem 
            Height          =   1380
            Left            =   165
            TabIndex        =   3
            Top             =   975
            Width           =   8670
            _Version        =   65536
            _ExtentX        =   15293
            _ExtentY        =   2434
            _StockProps     =   14
            Caption         =   "Banco Origem"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShadowStyle     =   1
            Begin VB.TextBox TxtCCorrente 
               Height          =   330
               Left            =   1680
               MaxLength       =   4
               TabIndex        =   9
               Top             =   210
               Width           =   450
            End
            Begin MSDataListLib.DataCombo CboCCorrente1 
               Bindings        =   "FrmTransfBancaria.frx":067E
               Height          =   330
               Left            =   2205
               TabIndex        =   1
               Top             =   210
               Width           =   6360
               _ExtentX        =   11218
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "desccombo"
               BoundColumn     =   "coco_cd_codigo"
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
            Begin MSMask.MaskEdBox MskConta 
               Height          =   330
               Left            =   1680
               TabIndex        =   12
               Top             =   570
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               PromptChar      =   "_"
            End
            Begin MSDataListLib.DataCombo CboPlanoContas 
               Bindings        =   "FrmTransfBancaria.frx":069D
               Height          =   330
               Left            =   2610
               TabIndex        =   13
               Top             =   570
               Width           =   5955
               _ExtentX        =   10504
               _ExtentY        =   582
               _Version        =   393216
               Style           =   2
               ListField       =   "desccombo"
               BoundColumn     =   "plco_cd_codreduzido"
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
            Begin MSMask.MaskEdBox MskPcr 
               Height          =   330
               Left            =   1680
               TabIndex        =   10
               Top             =   930
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo CboPcr 
               Bindings        =   "FrmTransfBancaria.frx":06B4
               Height          =   330
               Left            =   2610
               TabIndex        =   11
               Top             =   930
               Width           =   5955
               _ExtentX        =   10504
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "desccombo"
               BoundColumn     =   "cere_cd_estruturado"
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
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "Centro Custo:"
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
               Left            =   75
               TabIndex        =   34
               Top             =   975
               Width           =   1560
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "Conta Contábil:"
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
               Left            =   150
               TabIndex        =   33
               Top             =   615
               Width           =   1485
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "C/C:"
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
               Left            =   300
               TabIndex        =   32
               Top             =   255
               Width           =   1335
            End
         End
         Begin VB.ComboBox CboStatus 
            Height          =   315
            ItemData        =   "FrmTransfBancaria.frx":06C9
            Left            =   4260
            List            =   "FrmTransfBancaria.frx":06D3
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   2385
         End
         Begin VB.TextBox TxtVlTransf 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7425
            MaxLength       =   50
            TabIndex        =   6
            Top             =   240
            Width           =   1395
         End
         Begin VB.ComboBox CboTipoPag 
            Height          =   315
            ItemData        =   "FrmTransfBancaria.frx":06F1
            Left            =   2055
            List            =   "FrmTransfBancaria.frx":070A
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   600
            Width           =   2955
         End
         Begin VB.TextBox TxtDocPagto 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7425
            MaxLength       =   7
            TabIndex        =   8
            Top             =   630
            Width           =   1395
         End
         Begin MSAdodcLib.Adodc DatHistorico 
            Height          =   330
            Left            =   2880
            Top             =   3855
            Visible         =   0   'False
            Width           =   2340
            _ExtentX        =   4128
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
            Caption         =   "DatHistorico"
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
         Begin MSDataListLib.DataCombo CboHistorico 
            Bindings        =   "FrmTransfBancaria.frx":0764
            Height          =   315
            Left            =   1845
            TabIndex        =   20
            Top             =   3840
            Width           =   6585
            _ExtentX        =   11615
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "hist_tx_descricao"
            BoundColumn     =   "hist_cd_historico"
            Text            =   ""
         End
         Begin Threed.SSCommand CmdLimparHistorico 
            Height          =   330
            Left            =   8460
            TabIndex        =   21
            Top             =   3825
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   582
            _StockProps     =   78
            Enabled         =   0   'False
            MouseIcon       =   "FrmTransfBancaria.frx":077F
            Picture         =   "FrmTransfBancaria.frx":0BD1
         End
         Begin Threed.SSFrame FraBcDestino 
            Height          =   1395
            Left            =   165
            TabIndex        =   35
            Top             =   2385
            Width           =   8670
            _Version        =   65536
            _ExtentX        =   15293
            _ExtentY        =   2461
            _StockProps     =   14
            Caption         =   "Banco Destino"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShadowStyle     =   1
            Begin VB.TextBox TxtCCorrente2 
               Height          =   330
               Left            =   1680
               MaxLength       =   4
               TabIndex        =   14
               Top             =   195
               Width           =   450
            End
            Begin MSMask.MaskEdBox MskConta3 
               Height          =   330
               Left            =   1680
               TabIndex        =   18
               Top             =   570
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               PromptChar      =   "_"
            End
            Begin MSDataListLib.DataCombo CboPlanoContas3 
               Bindings        =   "FrmTransfBancaria.frx":0CE3
               Height          =   330
               Left            =   2610
               TabIndex        =   19
               Top             =   570
               Width           =   5955
               _ExtentX        =   10504
               _ExtentY        =   582
               _Version        =   393216
               Style           =   2
               ListField       =   "desccombo"
               BoundColumn     =   "plco_cd_codreduzido"
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
            Begin MSMask.MaskEdBox MskPcr3 
               Height          =   330
               Left            =   1680
               TabIndex        =   16
               Top             =   930
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo CboPcr3 
               Bindings        =   "FrmTransfBancaria.frx":0CFA
               Height          =   330
               Left            =   2610
               TabIndex        =   17
               Top             =   930
               Width           =   5955
               _ExtentX        =   10504
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "desccombo"
               BoundColumn     =   "cere_cd_estruturado"
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
            Begin MSDataListLib.DataCombo CboCCorrente2 
               Bindings        =   "FrmTransfBancaria.frx":0D0F
               Height          =   330
               Left            =   2205
               TabIndex        =   15
               Top             =   195
               Width           =   6360
               _ExtentX        =   11218
               _ExtentY        =   582
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "desccombo"
               BoundColumn     =   "coco_cd_codigo"
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
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "C/C:"
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
               Left            =   165
               TabIndex        =   38
               Top             =   255
               Width           =   1470
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               Caption         =   "Conta Contábil:"
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
               Left            =   150
               TabIndex        =   37
               Top             =   615
               Width           =   1485
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               Caption         =   "Centro Custo:"
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
               Left            =   75
               TabIndex        =   36
               Top             =   975
               Width           =   1560
            End
         End
         Begin MSComCtl2.DTPicker DtpDtTransf 
            Height          =   315
            Left            =   2055
            TabIndex        =   4
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   81461249
            CurrentDate     =   37460
         End
         Begin VB.Label LblCarta 
            Alignment       =   1  'Right Justify
            Caption         =   "N. Carta:"
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
            Left            =   840
            TabIndex        =   70
            Top             =   5400
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Referente:"
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
            Left            =   840
            TabIndex        =   68
            Top             =   4320
            Width           =   1005
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Pagamento:"
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
            Left            =   480
            TabIndex        =   31
            Top             =   675
            Width           =   1560
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Data Transferência:"
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
            Left            =   270
            TabIndex        =   30
            Top             =   255
            Width           =   1770
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Status:"
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
            Left            =   3495
            TabIndex        =   29
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor:"
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
            Left            =   6720
            TabIndex        =   28
            Top             =   255
            Width           =   675
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Histórico:"
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
            Left            =   900
            TabIndex        =   27
            Top             =   3870
            Width           =   885
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Doc. Pagto.:"
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
            Left            =   6270
            TabIndex        =   26
            Top             =   675
            Width           =   1125
         End
      End
      Begin VB.Label LblTotalComplemento 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -67215
         TabIndex        =   67
         Top             =   4410
         Width           =   1380
      End
      Begin VB.Label LblTotalComplemento2 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   210
         Left            =   -68280
         TabIndex        =   66
         Top             =   4455
         Width           =   1170
      End
   End
   Begin MSAdodcLib.Adodc DatContaCorrente 
      Height          =   330
      Left            =   0
      Top             =   3915
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
      Caption         =   "DatContaCorrente"
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
   Begin MSAdodcLib.Adodc DatConta 
      Height          =   330
      Left            =   30
      Top             =   3570
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
      Caption         =   "DatConta"
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
   Begin MSAdodcLib.Adodc DatPCR 
      Height          =   330
      Left            =   0
      Top             =   4290
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
      Caption         =   "DatPCR"
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
   Begin MSAdodcLib.Adodc DatBanco 
      Height          =   330
      Left            =   0
      Top             =   4635
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
      Caption         =   "DatBanco"
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
   Begin Threed.SSCommand CmdGravar 
      Height          =   315
      Left            =   3960
      TabIndex        =   40
      Top             =   6240
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "&Gravar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   315
      Left            =   5520
      TabIndex        =   41
      Top             =   6240
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "&Sair"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   315
      Left            =   2400
      TabIndex        =   72
      Top             =   6240
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "&Novo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmTransfBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XContaOrigem As String 'Armazena o nº da conta origem
Dim XCCOrigem As String 'Armazena o código da conta contábil origem
Dim XPcrOrigem As String 'Armazena o código do Pcr do banco origem
Dim XContaDestino As String 'Armazena o nº da conta destino
Dim XCCDestino As String 'Armazena o código da conta contábil destino
Dim XPcrDestino As String 'Armazena o código do Pcr do banco destino
Dim XCCCompl As String 'Armazena o nº da conta contábil do complemento
Dim XPcrCompl As String 'Armazena o código do Pcr do complemento
Dim XFL_FATURA As Long 'Armazena nº da fatura correspondente à despesa
Dim XFT_INSERIR
Dim Vetor3 As New XArray

'Variáveis utilizadas na gravação do formulário
Dim Sql As String
Dim XTipoDoc As Integer
Dim XControle As Long
Dim XFL_CODDUPL As Long     'Armazena o código de pelo 1 das partes da duplicata para utilizar na exp. para contabilidade
Dim XLT_NAOSAIRELFINANCEIRO As String  'Indica se a transferencia sairá ou não nos relatórios financeiros

Sub subPreencheCampos()
    Dim Sql As String
    Dim ResTransf As Object
    Dim Status As Integer
    Dim xcompl As Boolean
    Dim X As Integer

    xcompl = False
    
    SubQOpenRecordset ResTransf, "SELECT * FROM ConsFINTransfBancaria WHERE nofi_tx_controle ='" & XControleTransf & "' and empr_cd_empresa = " & PCodEmpresa & "", "Estatico"
    ResTransf.MoveFirst
    Do While Not ResTransf.EOF
        If Mid(ResTransf!nofi_nr_documento, 1, 1) = "R" Then
            xcompl = True
            ChkComplemento.Value = 1
            TabTransfBancaria.TabEnabled(1) = True
        End If
        
        If Not xcompl Then
            If ResTransf!nofi_tx_tipomov = "S" Then
                'Dados de NotasFiscais
                CboHistorico.Text = FunNulo(ResTransf!nofi_tx_historico)
                txtReferencia.Text = FunNulo(ResTransf!nofi_tx_ReferenciaTransferencia) '29/09/08 -Patrícia
                
                TxtNumCarta.Text = FunNulo(ResTransf!nofi_tx_NumCartaTransferencia) '20/10/08 -Patrícia

                'Dados de Duplicatas
                DtpDtTransf.Value = Format(ResTransf!dupl_dt_vencimento, "dd/mm/yy")
                
                If ResTransf!dupl_tx_status = "C" Then
                    Status = 0
                ElseIf ResTransf!dupl_tx_status = "R" Then
                    Status = 1
                End If
                
                CboStatus.ListIndex = Status
                TxtVlTransf.Text = Format$(ResTransf!dupl_vl_valor, "standard")
                
                If Not IsNull(ResTransf!dupl_tx_tipopagrec) Then
                   For X = 0 To CboTipoPag.ListCount - 1
                       If CboTipoPag.ItemData(X) = CInt(ResTransf!dupl_tx_tipopagrec) Then
                           CboTipoPag.ListIndex = X
                       End If
                   Next
                End If
                
                TxtDocPagto.Text = FunNulo(ResTransf!dupl_tx_docpagrec)
                CboCCorrente1.BoundText = ResTransf!coco_cd_codigo
                
                'Dados de Distribuicao
                MskPcr.Text = ResTransf!cere_cd_estruturado
                CboPcr.BoundText = ResTransf!cere_cd_estruturado
                MskConta.Text = ResTransf!plco_cd_CodReduzido
                CboPlanoContas.BoundText = ResTransf!plco_cd_CodReduzido
            Else
                CboCCorrente2.BoundText = ResTransf!coco_cd_codigo
                'Dados de Distribuicao
                MskPcr3.Text = ResTransf!cere_cd_estruturado
                CboPcr3.BoundText = ResTransf!cere_cd_estruturado
                MskConta3.Text = ResTransf!plco_cd_CodReduzido
                CboPlanoContas3.BoundText = ResTransf!plco_cd_CodReduzido
            End If
            
            If ResTransf!dupl_tx_NaoRelFinanceiro = "X" Then
                ChkNaoSaiRelFinanceiro.Value = 1
            Else
                ChkNaoSaiRelFinanceiro.Value = 0
            End If
        Else
            If ResTransf!nofi_tx_tipomov = "S" Then
                OptDeb.Value = True
            Else
                OptCred.Value = True
            End If
            
            'Dados de Duplicatas
            TxtVlCompl.Text = Format$(ResTransf!dupl_vl_valor, "standard")
            
            If ResTransf!dupl_tx_status = "C" Then
                Status = 0
            ElseIf ResTransf!dupl_tx_status = "R" Then
                Status = 1
            End If
            
            CboStatusCompl.ListIndex = Status
            CboTipoPagCompl.ListIndex = ((ResTransf!dupl_tx_tipopagrec) - 1)
            TxtDocPagtoCompl.Text = FunNulo(ResTransf!dupl_tx_docpagrec)

            'Dados de NotasFiscais
            CboHistCompl.Text = FunNulo(ResTransf!nofi_tx_historico)
            
            If Len(Trim(FunNulo(ResTransf!nofi_tx_ReferenciaTransferencia))) > 0 Then
                txtReferencia.Text = FunNulo(ResTransf!nofi_tx_ReferenciaTransferencia)
            Else
                txtReferencia.Text = ""
            End If
                        
            If Len(Trim(FunNulo(ResTransf!nofi_tx_NumCartaTransferencia))) > 0 Then 'Inserido em 20/10/08 - patrícia
                TxtNumCarta.Text = FunNulo(ResTransf!nofi_tx_NumCartaTransferencia)
            Else
                TxtNumCarta.Text = ""
            End If

            'Dados de Distribuicao
            MskPcr2.Text = ResTransf!cere_cd_estruturado
            CboPcr2.BoundText = ResTransf!cere_cd_estruturado
            MskConta2.Text = ResTransf!plco_cd_CodReduzido
            CboPlanoContas2.BoundText = ResTransf!plco_cd_CodReduzido
        End If
        ResTransf.MoveNext
    Loop
    ResTransf.Close
    
    If Not xcompl Then
         TabTransfBancaria.TabEnabled(1) = False
    End If
End Sub

Sub subPreencheCamposConta()
    Dim Sql As String
    Dim ResTransf As Object
    Dim Status As Integer
    Dim X As Integer
    
    SubQOpenRecordset ResTransf, "SELECT * FROM ConsFINCompDuplicata WHERE dupl_nr_fatura = " & NFatura, "Estatico"
    ResTransf.MoveFirst
    If Not ResTransf.EOF Then
            
        'Dados de Distribuicao
        TxtCCorrente.Text = ResTransf!coco_cd_codigo
        TxtCCorrente_LostFocus
        
    End If
    ResTransf.Close
    
End Sub


Function funChecarCampos() As Boolean
    funChecarCampos = True
    
    If FunObrigatorioCBOX(CboStatus, "O Status da transferência é obrigatório!") Then
        CboStatus.SetFocus
        Exit Function
    End If
    
    If FunObrigatorioTXT(TxtVlTransf, "O valor da transferência é obrigatório!") Then
        TxtVlTransf.SetFocus
        Exit Function
    End If
    
    If FunObrigatorioCBOX(CboTipoPag, "O tipo de pagamento da transferência é obrigatório!") Then
        CboTipoPag.SetFocus
        Exit Function
    End If
    
    If CboStatus.ItemData(CboStatus.ListIndex) = "2" And CboTipoPag.ItemData(CboTipoPag.ListIndex) = "1" Then
       If FunObrigatorioTXT(TxtDocPagto, "O documento de pagamento da transferência é obrigatório!") Then
            TxtDocPagto.SetFocus
            Exit Function
        End If
    End If
    
    If FunObrigatorioCBO(CboCCorrente1, "O banco origem é obrigatório!") Then
        CboCCorrente1.SetFocus
        Exit Function
    End If
    
    If FunObrigatorioCBO(CboCCorrente2, "O banco destino é obrigatório!") Then
        CboCCorrente2.SetFocus
        Exit Function
    End If
    
    If CboPcr.Text = "" Or CboPcr3.Text = "" Then
        MsgBox "O centro de custo é obrigatório!", vbCritical, "CAMPO OBRIGATÓRIO"
        Exit Function
    End If
    
    funChecarCampos = False
End Function

Sub subGravaComplemento()
    Dim XStatus As String
    Dim xNumFat As Long
    Dim XNumDoc As String
    Dim XMovimentacao As String
    Dim XNumConta As Integer
    Dim ResFormulario As Object
    
    'Insere um crédito ou débito na tabela de notas fiscais/duplicatas/distribuicao
    TDBGrid3.MoveFirst
    While Not TDBGrid3.EOF
        If OptDeb.Value = True Then
            XMovimentacao = "S"
        Else
            XMovimentacao = "E"
        End If
        
        'Abre o banco de Dados
        SubQOpenRecordset ResFormulario, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal= -1", Dinamico
        ResFormulario.AddNew
                    
        'Gera novo Nº de documento e de fatura
        xNumFat = FunPegaGuardaUltimo("F")
        XNumDoc = Format(FunPegaGuardaUltimo("D"), "R000000000") 'R - Rendimento
                    
        ResFormulario!nofi_tx_serie = "U"
        ResFormulario!nofi_tx_status = TDBGrid3.Columns("Status").Value
        ResFormulario!tido_cd_TipoDoc = XTipoDoc
        ResFormulario!nofi_dt_entrada = Format$(DtpDtTransf.Value, "DD/MM/YY")
        ResFormulario!nofi_dt_emissao = Format$(DtpDtTransf.Value, "DD/MM/YY")
        ResFormulario!nofi_tx_historico = TDBGrid3.Columns("Histórico") 'FunNulo(CboHistCompl.Text)
        ResFormulario!nofi_vl_valor = TDBGrid3.Columns("Valor") 'FunNuloVal(TxtVlCompl.Text)
        ResFormulario!nofi_tx_tipomov = XMovimentacao
        ResFormulario!nofi_nr_fatura = xNumFat
        ResFormulario!nofi_nr_documento = XNumDoc
        ResFormulario!nofi_tx_distribuir = "N"
        ResFormulario!nofi_tx_tipofatura = "1"
        ResFormulario!nofi_tx_controle = XControle
        ResFormulario!empr_cd_empresa = PCodEmpresa
        
        ResFormulario.Update
    
        Chave = ResFormulario!nofi_cd_notafiscal
        
        ResFormulario.Close
                    
        If XMovimentacao = "E" Then
            XMovimentacao = "C"
        Else
            XMovimentacao = "D"
        End If
        
        'Distribuição
        Sql = "INSERT INTO Distribuicao (nofi_cd_notafiscal,plco_cd_conta,cere_cd_pcr,dist_vl_valor)" & _
                          " VALUES (" & Chave & ",'" & TDBGrid3.Columns("CodConta") & "','" & TDBGrid3.Columns("CodPcr") & "'," & FunTrataFloat(TDBGrid3.Columns("Valor")) & ")"
        Conexao.Execute Sql
                                
        XNumConta = CboCCorrente1.BoundText 'banco origem
        
        'Duplicata
        Sql = "INSERT INTO Duplicatas (dupl_nr_duplicata,dupl_tx_debcred,dupl_tx_status,dupl_dt_vencimento,dupl_vl_valor," & _
           "dupl_dt_pagrec, dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo, dupl_nr_fatura)" & _
           " VALUES ('U','" & XMovimentacao & "','" & TDBGrid3.Columns("Status").Value & "'," & FunNuloData(DtpDtTransf.Value, NomeSgbd) & "," & FunTrataFloat(TDBGrid3.Columns("Valor")) & "," & _
           "" & FunNuloData(DtpDtTransf.Value, NomeSgbd) & ",'" & TDBGrid3.Columns(6).Value & "','" & FunNulo(TDBGrid3.Columns(7)) & "'," & XNumConta & ", " & xNumFat & ")"
        Conexao.Execute Sql
        
        TDBGrid3.MoveNext
    Wend
End Sub

Sub subGravaTransf()
    Dim XStatus As String
    Dim xNumFat As Long
    Dim XNumDoc As String
    Dim XMovimentacao As String
    Dim Cont As Integer
    Dim XNumConta As Integer 'Armazena o código da conta origem ou destino
    Dim ResFormulario As Object
    Dim ResDoc As Object
    Dim XConta As String
    Dim XPcr As String
    Dim XTipoPag As String
    Dim Xdoc As String
    Dim xdatapag As String
    
    XStatus = funVerificaStatus(CboStatus.ListIndex)
    If XStatus = "C" Then
        xdatapag = ""
    Else
        xdatapag = Format(DtpDtTransf.Value, "dd/mm/yyyy")
    End If

    'Obter código de documento de pagto se não for definido um e se for realizado
    If TxtDocPagto.Text = "" And XStatus = "R" Then
        Xdoc = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
    Else
        Xdoc = TxtDocPagto.Text
    End If
    
    'Obter o código do tipo de documento p/ transf.
    SubQOpenRecordset ResDoc, "SELECT tido_cd_tipodoc FROM TipoDocumentos WHERE tido_tx_sigla='TB'", "Estatico"
    XTipoDoc = ResDoc!tido_cd_TipoDoc
    ResDoc.Close
    
    XControle = FunPegaGuardaUltimo("C") 'Gera um nº de controle que será o mesmo para todos os registros desta gravação
    Cont = 0
    XMovimentacao = "S"
    
    While (Cont <= 1)
        'Insere o crédito e o débito na tabela de notas fiscais/duplicatas/distribuicao
    
        'Abre o banco de Dados
        SubQOpenRecordset ResFormulario, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal= -1", Dinamico
        ResFormulario.AddNew
                    
        'Gera novo Nº de documento e de fatura
        xNumFat = FunPegaGuardaUltimo("F")
        XNumDoc = Format(FunPegaGuardaUltimo("D"), "A000000000")
                    
        ResFormulario!nofi_tx_serie = "U"
        ResFormulario!nofi_tx_status = XStatus
        ResFormulario!tido_cd_TipoDoc = XTipoDoc
        ResFormulario!nofi_dt_entrada = Format$(DtpDtTransf.Value, "DD/MM/YY")
        ResFormulario!nofi_dt_emissao = Format$(DtpDtTransf.Value, "DD/MM/YY")
        ResFormulario!nofi_tx_historico = FunNulo(CboHistorico.Text)
        ResFormulario!nofi_vl_valor = FunNuloVal(TxtVlTransf.Text)
        ResFormulario!nofi_tx_tipomov = XMovimentacao
        ResFormulario!nofi_nr_fatura = xNumFat
        ResFormulario!nofi_nr_documento = XNumDoc
        ResFormulario!nofi_tx_distribuir = "N"
        ResFormulario!nofi_tx_tipofatura = "1"
        ResFormulario!nofi_tx_controle = XControle
        ResFormulario!empr_cd_empresa = PCodEmpresa
        
        If (Len(Trim(TxtNumCarta.Text)) > 0) Then 'inserido em 20/10/08 - Patrícia
            ResFormulario!nofi_tx_NumCartaTransferencia = Mid(Trim(TxtNumCarta.Text), 1, 30)
        Else
            ResFormulario!nofi_tx_NumCartaTransferencia = Null
        End If
        
        If (Len(Trim(txtReferencia.Text)) > 0) Then
            ResFormulario!nofi_tx_ReferenciaTransferencia = Trim(txtReferencia.Text)
        Else
            ResFormulario!nofi_tx_ReferenciaTransferencia = Null
        End If
        
        ResFormulario.Update
    
        Chave = ResFormulario!nofi_cd_notafiscal
            
        ResFormulario.Close
                    
        XTipoPag = CboTipoPag.ItemData(CboTipoPag.ListIndex)
              
        If Cont = 0 Then
           XMovimentacao = "D"
           XNumConta = CboCCorrente1.BoundText 'Conta origem
           XConta = XCCOrigem
           XPcr = XPcrOrigem
           
           XFL_FATURA = xNumFat
        Else
           XMovimentacao = "C"
           XNumConta = CboCCorrente2.BoundText 'Conta Destino
           XConta = XCCDestino
           XPcr = XPcrDestino
           If XTipoPag = "4" Then 'Débito em conta
              XTipoPag = "5"
           End If
        End If
                    
        'Distribuição
        Sql = "INSERT INTO Distribuicao (nofi_cd_notafiscal,plco_cd_conta,cere_cd_pcr,dist_vl_valor)" & _
                          " VALUES (" & Chave & ",'" & XConta & "','" & XPcr & "'," & FunNuloVal(FunTrataFloat(TxtVlTransf.Text)) & ")"
        Conexao.Execute Sql
      
        'Duplicatas
        Sql = "INSERT INTO Duplicatas (dupl_nr_duplicata,dupl_tx_debcred,dupl_tx_status,dupl_dt_vencimento,dupl_vl_valor," & _
              "dupl_dt_pagrec,dupl_tx_tipopagrec,dupl_tx_docpagrec,coco_cd_codigo,dupl_nr_fatura, dupl_tx_naorelfinanceiro)" & _
              " VALUES ('U','" & XMovimentacao & "','" & FunNulo(XStatus) & "'," & FunNuloData(DtpDtTransf.Value, NomeSgbd) & "," & FunNuloVal(FunTrataFloat(TxtVlTransf.Text)) & "," & FunNuloData(xdatapag, NomeSgbd) & "," & _
              "'" & CStr(XTipoPag) & "','" & Xdoc & "'," & XNumConta & "," & xNumFat & ", '" & XLT_NAOSAIRELFINANCEIRO & "')"
        Conexao.Execute Sql

        Cont = Cont + 1
                
        '******************REGISTRA LOG DA OPERAÇÃO**************************
        ReDim XGM_MATRIZLOG(6, 2) As Variant
        
        XGM_MATRIZLOG(0, 0) = "Nº Doc" 'Número de documento da nota fiscal.
        XGM_MATRIZLOG(1, 0) = "PP" 'Número de fatura da nota fiscal.
        XGM_MATRIZLOG(2, 0) = "Status"  'Status da transferencia
        XGM_MATRIZLOG(3, 0) = "Movimentação"  'Informa se a nota é débito ou crédito
        XGM_MATRIZLOG(4, 0) = "Valor" 'Valor da nota fiscal.
        XGM_MATRIZLOG(5, 0) = "Historico" 'Histórico da nota fiscal
        XGM_MATRIZLOG(0, 1) = XNumDoc
        XGM_MATRIZLOG(1, 1) = xNumFat
        XGM_MATRIZLOG(2, 1) = XStatus
        XGM_MATRIZLOG(3, 1) = XMovimentacao
        XGM_MATRIZLOG(4, 1) = FunNuloVal(TxtVlTransf.Text)
        XGM_MATRIZLOG(5, 1) = FunNulo(CboHistorico.Text)
        
        Call subRegistraLog("FrmTransfBancaria", "1", funCriaDescricaoLog(XGM_MATRIZLOG, Transferencia, TRANSFERENCIA_BANCARIA))
        '*******************************************************************
        
        XMovimentacao = "E"
    Wend
    
    '***********CONTABILIDADE************
    If CboStatus.Text = "Realizado" Then
        SubQOpenRecordset ResFormulario, "SELECT dupl_cd_duplicata FROM Duplicatas WHERE dupl_nr_fatura = " & xNumFat & "", Estatico
        XFL_CODDUPL = ResFormulario!dupl_cd_Duplicata
        ResFormulario.Close
    End If
    '************************************
End Sub

Sub subPrepararInsercao()
    'Tab 0
    DtpDtTransf.Value = Format$(Now, "DD/MM/YY")
    CboStatus.ListIndex = 0
    TxtCCorrente.Text = ""
    TxtVlTransf.Text = ""
    CboTipoPag.ListIndex = 0
    CboCCorrente1.Text = ""
    CboCCorrente2.Text = ""
    CboHistorico.Text = ""
    TxtDocPagto.Text = ""
    MskPcr.Mask = ""
    MskPcr.Text = ""
    MskPcr.Mask = pPCRmascara
    MskPcr3.Mask = ""
    MskPcr3.Text = ""
    MskPcr3.Mask = pPCRmascara

    'Tab 1
    OptCred.Value = True
    OptDeb.Value = False
    CboStatusCompl.ListIndex = 0
    CboTipoPagCompl.ListIndex = 0
    MskPcr2.Mask = ""
    MskPcr2.Text = ""
    MskPcr2.Mask = pPCRmascara
    MskConta2.Text = ""
    CboPlanoContas2.Text = ""
    CboPcr2.Text = ""
    TxtVlCompl.Text = ""
    CboHistCompl.Text = ""
    TxtDocPagtoCompl.Text = ""
End Sub

Function funVerificaStatus(Status As Integer) As String
    If Status = 0 Then 'Compromissado
            funVerificaStatus = "C"
    ElseIf Status = 1 Then 'Realizado
            funVerificaStatus = "R"
    End If
End Function

Private Sub botao_conf_Click()
    Dim XStatus As String
    
    'Função para verificar preenchimento dos campos obrigatórios
    If FunObrigatorioCBO(CboPcr2, "O centro de custo é obrigatório!") Then
        CboPcr2.SetFocus
        Exit Sub
    End If
    
    If FunObrigatorioMSK(MskConta2, "A conta contábil é obrigatória!") Then
        MskConta2.SetFocus
        Exit Sub
    End If
    
    If FunObrigatorioTXT(TxtVlCompl, "O valor é obrigatório!") Then
        TxtVlCompl.SetFocus
        Exit Sub
    End If
      
    If CboStatusCompl.ListIndex = 0 Then 'Compromissado
        XStatus = "C"
    ElseIf CboStatusCompl.ListIndex = 1 Then 'Realizado
        XStatus = "R"
    End If
      
    If XFT_INSERIR = "I" Then
        subInsereRegistroVetor Array(DatPCR.Recordset.Fields("cere_cd_pcr"), MskPcr2.Text, DatConta.Recordset.Fields("plco_cd_conta"), MskConta2.Text, CboHistCompl.Text, XStatus, CboTipoPagCompl.ItemData(CboTipoPagCompl.ListIndex), TxtDocPagtoCompl.Text, TxtVlCompl.Text), Vetor3, TDBGrid3
    Else
        TDBGrid3.Columns(0).Text = DatPCR.Recordset.Fields("cere_cd_pcr")
        TDBGrid3.Columns(1).Text = MskPcr2.Text
        TDBGrid3.Columns(2).Text = DatConta.Recordset.Fields("plco_cd_conta")
        TDBGrid3.Columns(3).Text = MskConta2.Text
        TDBGrid3.Columns(4).Text = CboHistCompl.Text
        TDBGrid3.Columns(6).Text = XStatus
        TDBGrid3.Columns(7).Text = CboTipoPagCompl.ItemData(CboTipoPagCompl.ListIndex)
        TDBGrid3.Columns(8).Text = TxtDocPagtoCompl.Text
        TDBGrid3.Columns(9).Text = TxtVlCompl.Text
        TDBGrid3.Update
    End If
    
    PanDistribuicao.Enabled = False
    PanDistribuicao.Visible = False
    If Vetor3(0, 0) <> Empty Then
       'HabilitaBotoes
    Else
        CmdInsereComp.Enabled = True
    End If
    
    TDBGrid3.Enabled = True
End Sub

Private Sub CboCCorrente1_Change()
   If CboCCorrente1.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboCCorrente1.SelectedItem
        XContaOrigem = DatContaCorrente.Recordset.Fields("coco_tx_conta")
        DatConta.Recordset.MoveFirst
        DatConta.Recordset.Find "plco_cd_conta =" & DatContaCorrente.Recordset.Fields("plco_cd_conta")
        XCCOrigem = DatContaCorrente.Recordset.Fields("plco_cd_conta")
        MskConta.Text = DatConta.Recordset.Fields("plco_cd_codreduzido")
        CboPlanoContas.BoundText = MskConta.Text
        MskConta.Enabled = False
        CboPlanoContas.Enabled = False
        TxtCCorrente.Text = CboCCorrente1.BoundText
   End If
End Sub

Private Sub CboCCorrente2_Change()
   If CboCCorrente2.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboCCorrente2.SelectedItem
        XContaDestino = DatContaCorrente.Recordset.Fields("coco_tx_conta")
        DatConta.Recordset.MoveFirst
        DatConta.Recordset.Find "plco_cd_conta =" & DatContaCorrente.Recordset.Fields("plco_cd_conta")
        XCCDestino = DatConta.Recordset.Fields("plco_cd_conta")
        MskConta3.Text = DatConta.Recordset.Fields("plco_cd_codreduzido")
        CboPlanoContas3.BoundText = MskConta3.Text
        MskConta3.Enabled = False
        CboPlanoContas3.Enabled = False
        TxtCCorrente2.Text = CboCCorrente2.BoundText
   End If
End Sub

Private Sub CboHistorico_GotFocus()
    CboHistorico.Text = "Transf. Bancária da C/C " & XContaOrigem & " para C/C " & XContaDestino
End Sub

Private Sub CboHistorico_KeyPress(KeyAscii As Integer)
    If Len(CboHistorico) = 100 Then
        If KeyAscii <> 8 And CboHistorico.SelLength = 0 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub CboPcr_Change()
    If CboPcr.Text <> "" Then
        DatPCR.Recordset.Bookmark = CboPcr.SelectedItem
        If DatPCR.Recordset.Fields("cere_tx_tipo") <> "A" Then
            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPcr.BoundText = MskPcr.Text
            CboPcr.SetFocus
            Exit Sub
        End If
        MskPcr.Text = CboPcr.BoundText
        XPcrOrigem = DatPCR.Recordset.Fields("cere_cd_pcr")
    End If
End Sub

Private Sub CboPcr2_Change()
    If CboPcr2.Text <> "" Then
        DatPCR.Recordset.Bookmark = CboPcr2.SelectedItem
        If DatPCR.Recordset.Fields("cere_tx_tipo") <> "A" Then
            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPcr2.BoundText = MskPcr2.Text
            CboPcr2.SetFocus
            Exit Sub
        End If
        MskPcr2.Text = CboPcr2.BoundText
        XPcrCompl = DatPCR.Recordset.Fields("cere_cd_pcr")
    End If
End Sub

Private Sub CboPcr3_Change()
    If CboPcr3.Text <> "" Then
        DatPCR.Recordset.Bookmark = CboPcr3.SelectedItem
        If DatPCR.Recordset.Fields("cere_tx_tipo") <> "A" Then
            MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPcr3.BoundText = MskPcr3.Text
            CboPcr3.SetFocus
            Exit Sub
        End If
        MskPcr3.Text = CboPcr3.BoundText
        XPcrDestino = DatPCR.Recordset.Fields("cere_cd_pcr")
    End If
End Sub

Private Sub cboPlanoContas_Change()
'    If CboPlanoContas.Text <> "" Then
'        DatConta.Recordset.Bookmark = CboPlanoContas.SelectedItem
'        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
'            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
'            CboPlanoContas.BoundText = MskConta.Text  '? Não deveria permitir gravar uma conta que não é analítica
'            CboPlanoContas.SetFocus
'            Exit Sub
'        End If
'        MskConta.Text = CboPlanoContas.BoundText
'        XCCOrigem = DatConta.Recordset.Fields("plco_cd_conta")
'    End If
End Sub

Private Sub CboPlanoContas2_Change()
    If CboPlanoContas2.Text <> "" Then
        DatConta.Recordset.Bookmark = CboPlanoContas2.SelectedItem
        
        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboPlanoContas2.BoundText = MskConta2.Text
            CboPlanoContas2.SetFocus
            Exit Sub
        End If
        
        MskConta2.Text = CboPlanoContas2.BoundText
        XCCCompl = DatConta.Recordset.Fields("plco_cd_conta")
    End If
End Sub

Private Sub CboPlanoContas3_Change()
'    If CboPlanoContas3.Text <> "" Then
'        DatConta.Recordset.Bookmark = CboPlanoContas3.SelectedItem
'        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
'            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
'            CboPlanoContas3.BoundText = MskConta3.Text
'            CboPlanoContas3.SetFocus
'            Exit Sub
'        End If
'        MskConta3.Text = CboPlanoContas3.BoundText
'        XCCDestino = DatConta.Recordset.Fields("plco_cd_conta")
'    End If

End Sub

Private Sub CboStatus_LostFocus()
    If CboStatus.ListIndex = 1 And (CboTipoPag.ItemData(CboTipoPag.ListIndex)) <> "6" Then 'Status = Realizado e Tipo Pag <> Dinheiro
        TxtDocPagto.Enabled = True
    Else
        TxtDocPagto.Enabled = False
    End If
End Sub

Private Sub CboStatusCompl_LostFocus()
    If CboStatusCompl.ListIndex = 1 And (CboTipoPagCompl.ItemData(CboTipoPagCompl.ListIndex)) <> "6" Then 'Status = Realizado e Tipo Pag <> Dinheiro
        TxtDocPagtoCompl.Enabled = True
    Else
        TxtDocPagtoCompl.Enabled = False
    End If
End Sub

Private Sub CboTipoPag_LostFocus()
    If CboStatus.ListIndex = 1 And (CboTipoPag.ItemData(CboTipoPag.ListIndex)) <> "6" Then 'Status = Realizado e Tipo Pag <> Dinheiro
        TxtDocPagto.Enabled = True
    Else
        TxtDocPagto.Enabled = False
    End If
End Sub

Private Sub CboTipoPagCompl_LostFocus()
    If CboStatusCompl.ListIndex = 1 And (CboTipoPagCompl.ItemData(CboTipoPagCompl.ListIndex)) <> "6" Then 'Status = Realizado e Tipo Pag <> Dinheiro
        TxtDocPagtoCompl.Enabled = True
    Else
        TxtDocPagtoCompl.Enabled = False
    End If
End Sub

Private Sub ChkComplemento_Click()
    If ChkComplemento.Value = 1 Then
       TabTransfBancaria.TabEnabled(1) = True
       
       Sql = "SELECT * FROM ConsFINTransfBancaria WHERE nofi_tx_controle ='" & XControleTransf & "' AND left(nofi_nr_documento, 1) = 'R'  and empr_cd_empresa = " & PCodEmpresa & ""
       subCarregaVetor Sql, Array("cere_cd_pcr", "cere_cd_estruturado", "plco_cd_conta", "plco_cd_codreduzido", "nofi_tx_historico", "dupl_tx_status", "dupl_tx_TipoPagRec", "dupl_tx_DocPagRec", "dupl_vl_valor"), Vetor3, TDBGrid3
    Else
       TabTransfBancaria.TabEnabled(1) = False
    End If
End Sub

Private Sub CmdAlteraComp_Click()
    PanDistribuicao.Visible = True
    PanDistribuicao.Enabled = True
    XFT_INSERIR = "A"
    
    If TDBGrid3.Columns(1).Text <> "" Then
        MskPcr2.Text = TDBGrid3.Columns(1).Text
    End If
    CboPcr2.BoundText = MskPcr2.Text
    If TDBGrid3.Columns(3).Text <> "" Then
        MskConta2.Text = TDBGrid3.Columns(3).Text
    End If
    CboPlanoContas2.BoundText = MskConta2.Text
    
    TxtVlCompl.SetFocus
    Exit Sub
End Sub

Private Sub CmdDesistir_Click()
    Unload Me
End Sub

Private Sub cmdGravar_Click()
    Dim ResNF As Object
    Dim RESDUPL As Object

    If ChkNaoSaiRelFinanceiro.Value = 1 Then
        XLT_NAOSAIRELFINANCEIRO = "X"
    Else
        XLT_NAOSAIRELFINANCEIRO = ""
    End If
    
    If CboCCorrente1.BoundText = CboCCorrente2.BoundText Then
        MsgBox "Não é possível transferir para a mesma conta!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
    If funChecarCampos Then Exit Sub
    
    Conexao.BeginTrans
    
    '***************CONTABILIDADE*******************
    If XGB_PQUALICONT Then 'Se possui o módulo de contabilidade
        If funMesFechado(DtpDtTransf.Value) <> "A" Then
            MsgBox "Não é possível realizar a Transferência nessa data, pois o mês contábil já foi fechado. " _
                + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            Conexao.CommitTrans
            Exit Sub
        End If
    End If
    '***********************************************
    
    '///////APAGAR TRANSFERÊNCIA BANCÁRIA///////////
    If XFormulario = "TelaNotasFiscais" Or XFormulario = "TelaFatura" Or XFormulario = "TelaDuplicatas" Then
        
        SubQOpenRecordset ResNF, "SELECT * FROM NotasFiscais WHERE nofi_tx_controle = '" & XControleTransf & "' and empr_cd_empresa = " & PCodEmpresa & "", "Estatico"
        If Not ResNF.EOF Then
            ResNF.MoveFirst
        End If
        
        Do While Not ResNF.EOF
            '***************CONTABILIDADE*******************
            'Exclui a parte da contabilidade
            Conexao.Execute "DELETE FROM Lancamentos " & _
                            " WHERE lanc_cd_OrigemFin= " & ResNF!nofi_cd_notafiscal & "" & _
                            " AND lanc_tx_Origem = 'FN'  " & _
                            " AND empr_cd_empresa = " & PCodEmpresa & ""
    
            SubQOpenRecordset RESDUPL, "SELECT * FROM Duplicatas WHERE dupl_nr_Fatura = " & ResNF!nofi_nr_fatura & " and dupl_tx_Status='R'", Estatico
            If Not RESDUPL.EOF Then
                Conexao.Execute "DELETE FROM Lancamentos " & _
                            " WHERE lanc_cd_OrigemFin= " & RESDUPL!dupl_cd_Duplicata & "" & _
                            " AND lanc_tx_Origem = 'FD'  " & _
                            " AND empr_cd_empresa = " & PCodEmpresa & ""
            End If
            RESDUPL.Close
            '***********************************************
            
            Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura = " & ResNF!nofi_nr_fatura
            Conexao.Execute "DELETE FROM DescontosNF WHERE deen_nr_fatura = " & ResNF!nofi_nr_fatura
            Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_notafiscal=" & ResNF!nofi_cd_notafiscal
                                    
            ResNF.MoveNext
        Loop
        
        Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_tx_controle='" & XControleTransf & "'  and empr_cd_empresa = " & PCodEmpresa & ""
        ResNF.Close
    End If
    '//////////////////////////////////////////////////////
    
    Call subGravaTransf
    
    If ChkComplemento = 1 Then
        Call subGravaComplemento
    End If
    
    '***************CONTABILIDADE*******************
    If CboStatus.Text = "Realizado" Then 'contabiliza só se for realizado, pois não existe provisão para TB
'        SubContabilizarParcela XFL_CODDUPL, DtpDtTransf, False, False
    End If
    '***********************************************
    
    Conexao.CommitTrans
    
    'Imprime PP
    'If XGT_IMPRIMEPP = "S" Then
    '    If MsgBox("Deseja Imprimir Processo de Pagamento?", vbQuestion + vbYesNo, "ATENÇÃO") = vbYes Then
    '        subImprimePP XFL_FATURA, False
    '    End If
    'End If
    
    TabTransfBancaria.Tab = 0
    TabTransfBancaria.TabEnabled(1) = False
    
    'Cássio Medeiros 16/12/2013
    'Formulário mantem ativo para permitir novos registros após a gravação ao clicar em novo
    MsgBox "Operação realizada com sucesso.", vbInformation + vbOKOnly, "Informação"
    CmdGravar.Enabled = False
    'Unload Me
    
'Exit Sub
'ResumoDoErro:
'    Unload Me
'    Exit Sub

TrataErro:
'    funTrataErros (ComMensagem)
'    Conexao.RollbackTrans
'    Resume ResumoDoErro
End Sub

Private Sub CmdInsereComp_Click()
    XFT_INSERIR = "I"
        
    TxtVlCompl.Text = ""
    CboStatusCompl.ListIndex = 1
    CboTipoPagCompl.ListIndex = 1
    TxtDocPagtoCompl.Text = ""
    MskConta2.Text = ""
    CboPlanoContas2.Text = ""
    MskPcr2.PromptInclude = False
    MskPcr2.Text = ""
    MskPcr2.PromptInclude = True
    CboPcr2.Text = ""
    CboHistCompl.Text = ""
    
    PanDistribuicao.Visible = True
    PanDistribuicao.Enabled = True
    TxtVlCompl.SetFocus
End Sub

Private Sub CmdLimparHistorico_Click()
    CboHistorico.BoundText = 0
    CboHistorico.Text = ""
End Sub

Private Sub CmdLimparHistorico2_Click()
    CboHistCompl.BoundText = 0
    CboHistCompl.Text = ""
End Sub


Private Sub cmdNovo_Click()
    LimpaCampos
End Sub

Private Sub CmdRemoverComp_Click()
    TDBGrid3.Enabled = False
   
    If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
        subRemoveRegistroVetor Vetor3, TDBGrid3
    End If

    If Vetor3(0, 0) <> Empty Then
        'HabilitaBotoes
    Else
        CmdInsereComp.Enabled = True
    End If
   
    TDBGrid3.Enabled = True
 
    Exit Sub
End Sub

Private Sub CmdSairDistribuicao_Click()
    PanDistribuicao.Visible = False
    PanDistribuicao.Enabled = False
       
    XFT_INSERIR = ""
    TDBGrid3.Enabled = True
    
    If Vetor3(0, 0) <> Empty Then
       'HabilitaBotoes
    Else
       CmdInsereComp.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    Set Formulario = FrmTransfBancaria
    
End Sub

Private Sub Form_Load()

    FrmTransfBancaria.KeyPreview = True
 
    Dim ResTransfBanc As Object
    Dim XLO_RSEMPRESA As ADODB.Recordset ' Declaração em 20/10/08 - Patrícia
    
    subManutencaoJanelasAtivas "I", "FrmTransfBancaria"
    'Pesquisa por CGC Raiz - Samuel 26/09/18
    subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos ORDER BY hist_tx_descricao", Estatico
    subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
    If Mid(PEmpresa, 1, 10) = "QUALIDADOS" Or Mid(PEmpresa, 1, 8) = "QUALINFO" Then
        subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' AND coco_tx_status = 'A' ORDER BY coco_tx_Descricao", Estatico
    Else
        subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " AND coco_tx_status = 'A' ORDER BY coco_tx_Descricao", Estatico
    End If
    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_Tipo = 'A' ORDER BY plco_tx_desccontabil", Estatico
    subConectarControleDadosNV DatBanco, "SELECT * FROM Bancos", Estatico
    
    ChkComplemento.Value = 0
    CmdGravar.Caption = "&Gravar"

    'Combo do tipo de pagto/rec
    'Tab 0
    CboTipoPag.RemoveItem (1)
    CboTipoPag.RemoveItem (1)
    CboTipoPag.RemoveItem (1)
    CboTipoPag.RemoveItem (2)

    'Tab 1
    CboTipoPagCompl.ListIndex = 0
    CboTipoPagCompl.RemoveItem (1)
    CboTipoPagCompl.RemoveItem (1)
    CboTipoPagCompl.RemoveItem (1)
    
    If XFormulario = "FrmTransfBancaria(TelaNotasFiscais)" Then
        subPrepararInsercao 'Prepara os campo para a inserção
        subPreencheCamposConta 'Insere a conta corrente do registro selecionado na lista
    ElseIf XFormulario = "TelaNotasFiscais" Or XFormulario = "TelaFatura" Or XFormulario = "TelaDuplicatas" Then
        subPreencheCampos
    ElseIf XFormulario = "FrmTransfBancaria" Then
        TabTransfBancaria.TabEnabled(1) = False
        subPrepararInsercao 'Prepara os campo para a inserção
    End If
    
    If CboStatus.Text = "Realizado" Then

        SubQOpenRecordset ResTransfBanc, "SELECT * FROM ConsFINDuplicatas WHERE dupl_tx_DocPagRec = '" & TxtDocPagto.Text & _
                                            "' AND  dupl_dt_PagRec = " & FunNuloData(DtpDtTransf.Value, NomeSgbd) & _
                                            " AND coco_cd_codigo = " & TxtCCorrente.Text, "Estatico"
        
        If ResTransfBanc.RecordCount > 1 Then
        
            MsgBox "Esta transferência pertence a um cheque consolidado e não poderá ser alterada! Cancele o cheque para efetuar a operação.", vbCritical + vbOKOnly, "ATENÇÃO"
            CmdGravar.Enabled = False
            
        End If
        ResTransfBanc.Close
        
    End If
   
   
    SubQOpenRecordset XLO_RSEMPRESA, "SELECT empr_tx_cliente FROM Empresas", "Estatico"
    
    If Not (XLO_RSEMPRESA.EOF) Then ' If inserido em 20/10/08 - Patrícia. Inclusão do campo Carta solicitado pelo cliente Instituto Aliança
         If Trim(XLO_RSEMPRESA!empr_tx_cliente) = "ALIANCA" Then
            LblCarta.Visible = True
            TxtNumCarta.Visible = True
            LblCarta.Top = 5280
            TxtNumCarta.Top = 5160

            FraGeral.Height = 5805
            TabTransfBancaria.Height = 6420
            ChkComplemento.Top = 6840
            CmdGravar.Top = 6840
            CmdDesistir.Top = 6840
            cmdNovo.Top = 6840
            ChkNaoSaiRelFinanceiro.Top = 6840
            FrmTransfBancaria.Height = 7830

         Else
            LblCarta.Visible = False
            TxtNumCarta.Visible = False
            LblCarta.Top = 5400
            TxtNumCarta.Top = 5280

            FraGeral.Height = 5205
            TabTransfBancaria.Height = 5820
            ChkComplemento.Top = 6240
            CmdGravar.Top = 6240
            CmdDesistir.Top = 6240
            ChkNaoSaiRelFinanceiro.Top = 6240
            FrmTransfBancaria.Height = 7215
         End If
    End If
    XLO_RSEMPRESA.Close
    CmdGravar.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "FrmTransfBancaria"
End Sub

Private Sub MskConta_GotFocus()
'    Call subSelecionaMSK(MskConta)
End Sub

Private Sub mskConta_LostFocus()
'    If MskConta.Text <> "" Then
'        CboPlanoContas.BoundText = MskConta.Text
'        If CboPlanoContas.Text = "" Then
'            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
'            MskConta.SetFocus
'        End If
'    End If
End Sub


Private Sub MskConta2_GotFocus()
    Call subSelecionaMSK(MskConta2)
End Sub

Private Sub MskConta2_LostFocus()
    If MskConta2.Text <> "" Then
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            MskConta2.Text = Format(MskConta2.Text, XGT_MSKCODREDUZIDO)
        End If
        CboPlanoContas2.BoundText = MskConta2.Text
        If CboPlanoContas2.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskConta2.SetFocus
        End If
    End If

End Sub

Private Sub MskConta3_GotFocus()
    'Call subSelecionaMSK(MskConta3)
End Sub

Private Sub MskConta3_LostFocus()
'    If MskConta3.Text <> "" Then
'        CboPlanoContas3.BoundText = MskConta3.Text
'        If CboPlanoContas3.Text = "" Then
'            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
'            MskConta3.SetFocus
'        End If
'    End If
End Sub

Private Sub MskPcr_GotFocus()
    Call subSelecionaMSK(MskPcr)
End Sub

Private Sub MskPcr_LostFocus()
    MskPcr.PromptInclude = False
    If MskPcr.Text <> "" Then
        If Len(MskPcr.Text) <> Len(pPCRmascara) And pPCRnivel = 1 Then
            MskPcr.Text = funFormataCCusto(MskPcr.Text)
        End If
        MskPcr.Mask = pPCRmascara
        MskPcr.PromptInclude = True
        CboPcr.BoundText = MskPcr.Text
        If CboPcr.Text = "" Then
            MsgBox "Não existe item do Centro de Custo com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            If MskPcr.Visible = True Then MskPcr.SetFocus
        End If
    End If
    MskPcr.PromptInclude = True

End Sub

Private Sub MskPcr2_GotFocus()
    Call subSelecionaMSK(MskPcr2)
End Sub

Private Sub MskPcr2_LostFocus()
    MskPcr2.PromptInclude = False
    If MskPcr2.Text <> "" Then
        If Len(MskPcr2.Text) <> Len(pPCRmascara) And pPCRnivel = 1 Then
            MskPcr2.Text = funFormataCCusto(MskPcr2.Text)
        End If
        MskPcr2.Mask = pPCRmascara
        MskPcr2.PromptInclude = True
        CboPcr2.BoundText = MskPcr2.Text
        If CboPcr2.Text = "" Then
            MsgBox "Não existe item do Centro de Custo com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            If MskPcr2.Visible = True Then MskPcr2.SetFocus
        End If
    End If
    MskPcr2.PromptInclude = True
End Sub

Private Sub MskPcr3_GotFocus()
    Call subSelecionaMSK(MskPcr3)
End Sub

Private Sub MskPcr3_LostFocus()
    MskPcr3.PromptInclude = False
    If MskPcr3.Text <> "" Then
        If Len(MskPcr3.Text) <> Len(pPCRmascara) And pPCRnivel = 1 Then
            MskPcr3.Text = funFormataCCusto(MskPcr3.Text)
        End If
        MskPcr3.Mask = pPCRmascara
        MskPcr3.PromptInclude = True
        CboPcr3.BoundText = MskPcr3.Text
        If CboPcr3.Text = "" Then
            MsgBox "Não existe item do Centro de Custo com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            If MskPcr3.Visible = True Then MskPcr3.SetFocus
        End If
    End If
    MskPcr3.PromptInclude = True
End Sub

Private Sub TxtCCorrente_GotFocus()
    Call subSelecionaTXT(TxtCCorrente)
End Sub

Private Sub TxtCCorrente_LostFocus()
    If TxtCCorrente.Text <> "" Then
        CboCCorrente1.BoundText = TxtCCorrente.Text
        If CboCCorrente1.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtCCorrente.SetFocus
        End If
    End If
End Sub

Private Sub TxtCCorrente2_GotFocus()
    Call subSelecionaTXT(TxtCCorrente2)
End Sub

Private Sub TxtCCorrente2_LostFocus()
    If TxtCCorrente2.Text <> "" Then
        CboCCorrente2.BoundText = TxtCCorrente2.Text
        If CboCCorrente2.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtCCorrente2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDocPagto_GotFocus()
    If CboTipoPag.ListIndex = 3 Or CboTipoPag.ListIndex = 4 Or CboTipoPag.ListIndex = 5 Then
        TxtDocPagto.Enabled = False
    Else
        TxtDocPagto.Enabled = True
    End If
End Sub

Private Sub TxtDocPagto_LostFocus()
    If CboTipoPag.ListIndex = 0 Then
        TxtDocPagto.Text = Format(TxtDocPagto, "0000000")
    End If
End Sub

Private Sub TxtDocPagtoCompl_GotFocus()
    If CboTipoPagCompl.ListIndex = 3 Or CboTipoPagCompl.ListIndex = 4 Or CboTipoPagCompl.ListIndex = 5 Then
        TxtDocPagtoCompl.Enabled = False
    Else
        TxtDocPagtoCompl.Enabled = True
    End If
End Sub

Private Sub TxtVlCompl_KeyPress(KeyAscii As Integer)
  
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
  
End Sub

Private Sub TxtVlCompl_LostFocus()
    
    If IsNumeric(TxtVlCompl.Text) Then
        TxtVlCompl.Text = Format$(TxtVlCompl.Text, "standard")
    Else
        TxtVlCompl.Text = ""
    End If
    
End Sub

Private Sub TxtVlTransf_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtVlTransf_LostFocus()
    If IsNumeric(TxtVlTransf.Text) Then
        TxtVlTransf.Text = Format$(TxtVlTransf.Text, "##,##0.00")
    Else
        TxtVlTransf.Text = ""
    End If
End Sub

Private Sub LimpaCampos()
    DtpDtTransf.Value = Format(Now, "dd/MM/yyyy")
    MskConta.Text = ""
    TxtCCorrente.Text = ""
    TxtDocPagto.Text = ""
    TxtVlCompl.Text = ""
    CboCCorrente1.Text = ""
    CboCCorrente2.Text = ""
    CboPlanoContas3.Text = ""
    CboHistCompl.Text = ""
    txtReferencia.Text = ""
    CboPlanoContas3.Text = ""
    CboPlanoContas.Text = ""
    MskPcr3.Mask = ""
    MskPcr3.Text = ""
    TxtCCorrente2.Text = ""
    TxtNumCarta.Text = ""
    TDBGrid3.EmptyRows = True
    CboPcr.Text = ""
    CboPcr3.Text = ""
    TxtVlTransf.Text = ""
    MskConta3.Mask = ""
    MskConta3.Text = ""
    MskPcr.Mask = ""
    MskPcr.Text = ""
    CboHistorico.Text = ""
    ChkComplemento.Value = Unchecked
    CboStatus.ListIndex = 0
    CboTipoPag.ListIndex = 0
    MskPcr.Mask = pPCRmascara
    MskPcr3.Mask = pPCRmascara
    OptCred.Value = True
    OptDeb.Value = False
    CboStatusCompl.ListIndex = 0
    CboTipoPagCompl.ListIndex = 0
    MskPcr2.Mask = pPCRmascara
    MskConta2.Text = ""
    CboPlanoContas2.Text = ""
    CboPcr2.Text = ""
    TxtDocPagtoCompl.Text = ""
    ChkNaoSaiRelFinanceiro.Value = Unchecked
    XControleTransf = -1
    CmdGravar.Enabled = True
End Sub

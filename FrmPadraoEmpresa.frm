VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{12FEA440-3264-4401-B639-B45668B09A49}#9.0#0"; "DBImage.ocx"
Begin VB.Form FrmPadraoEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Padrões Utilizados"
   ClientHeight    =   8205
   ClientLeft      =   2205
   ClientTop       =   1995
   ClientWidth     =   8505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPadraoEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   8505
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc DatCNAEDisponivel 
      Height          =   330
      Left            =   3720
      Top             =   9240
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "DatCNAEDisponivel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatCNAEAssociado 
      Height          =   330
      Left            =   3720
      Top             =   8880
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "DatCNAEAssociado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatCNAEClasse 
      Height          =   330
      Left            =   3720
      Top             =   8520
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "DatCNAEClasse"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatCNAEGrupo 
      Height          =   330
      Left            =   1920
      Top             =   9240
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "DatCNAEGrupo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatCNAEDivisao 
      Height          =   330
      Left            =   1920
      Top             =   8880
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "DatCNAEDivisao"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatCNAESecao 
      Height          =   330
      Left            =   1920
      Top             =   8520
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "DatCNAESecao"
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
   Begin MSAdodcLib.Adodc DatEmpresaPrincipal 
      Height          =   375
      Left            =   60
      Top             =   8160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CheckBox ChkAtivar 
      Caption         =   "Empresa Ativa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   93
      Top             =   7800
      Width           =   2070
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "&Sair"
      Height          =   360
      Left            =   6960
      TabIndex        =   95
      Top             =   7800
      Width           =   1275
   End
   Begin VB.CommandButton CmdGravar 
      Caption         =   "&Gravar"
      Height          =   360
      Left            =   5520
      TabIndex        =   94
      Top             =   7800
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog DlgAbrirFigura 
      Left            =   3240
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatEmpresas 
      Height          =   330
      Left            =   60
      Top             =   8520
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSAdodcLib.Adodc datpadroes 
      Height          =   330
      Left            =   60
      Top             =   9240
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "DatPadrao"
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
   Begin MSAdodcLib.Adodc DatContaCorrente 
      Height          =   330
      Left            =   60
      Top             =   8880
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Dados Gerais"
      TabPicture(0)   =   "FrmPadraoEmpresa.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePadrao(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frRet"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Endereço"
      TabPicture(1)   =   "FrmPadraoEmpresa.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FramePadrao(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Padrões"
      TabPicture(2)   =   "FrmPadraoEmpresa.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "FraNotasEstoque"
      Tab(2).Control(3)=   "FraVinculaCRCC"
      Tab(2).Control(4)=   "FraFiltra"
      Tab(2).Control(5)=   "FraOrdemComboDesc"
      Tab(2).Control(6)=   "FraModeloRecibo"
      Tab(2).Control(7)=   "FraRelatorioQualiCapi"
      Tab(2).Control(8)=   "FraPP"
      Tab(2).Control(9)=   "FraQualifats"
      Tab(2).Control(10)=   "FraCapiQualifin"
      Tab(2).Control(11)=   "FraImportacao"
      Tab(2).Control(12)=   "FraExpContabilidade"
      Tab(2).Control(13)=   "FraMascara"
      Tab(2).Control(14)=   "FraCategoriaNvPlContas"
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "&Logomarca"
      TabPicture(3)   =   "FrmPadraoEmpresa.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FraLogoMarca"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&Responsáveis"
      TabPicture(4)   =   "FrmPadraoEmpresa.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FramePadrao(2)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Impostos Federais"
      TabPicture(5)   =   "FrmPadraoEmpresa.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "FramePadrao(4)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&CNAE"
      TabPicture(6)   =   "FrmPadraoEmpresa.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "TDBGridCNAEAssociado"
      Tab(6).Control(1)=   "TDBGridCNAEDisponivel"
      Tab(6).Control(2)=   "FramePadrao(5)"
      Tab(6).Control(3)=   "CmdInserirCNAE"
      Tab(6).Control(4)=   "CmdRemoverCNAE"
      Tab(6).ControlCount=   5
      Begin VB.CommandButton CmdRemoverCNAE 
         Caption         =   "<<"
         Height          =   495
         Left            =   -71160
         TabIndex        =   180
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton CmdInserirCNAE 
         Caption         =   ">>"
         Height          =   495
         Left            =   -71160
         TabIndex        =   179
         Top             =   4440
         Width           =   495
      End
      Begin VB.Frame Frame4 
         Caption         =   "Contas a Pagar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   -70200
         TabIndex        =   168
         Top             =   6720
         Width           =   3375
         Begin VB.CheckBox chkAlertarSaldoAdiantamento 
            Caption         =   "Alertar saldo de Adiantamento"
            Height          =   255
            Left            =   180
            TabIndex        =   169
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo Plano de Contas Referencial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   -70200
         TabIndex        =   165
         Top             =   5970
         Width           =   3375
         Begin VB.ComboBox CboTipoPlanoContasReferencial 
            Height          =   315
            ItemData        =   "FrmPadraoEmpresa.frx":0506
            Left            =   120
            List            =   "FrmPadraoEmpresa.frx":0528
            Style           =   2  'Dropdown List
            TabIndex        =   166
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame frRet 
         Caption         =   "RET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   160
         Top             =   5040
         Width           =   8055
         Begin MSDataListLib.DataCombo CboEmpresaVinculada 
            Bindings        =   "FrmPadraoEmpresa.frx":0613
            Height          =   330
            Left            =   1515
            TabIndex        =   161
            Top             =   240
            Width           =   5820
            _ExtentX        =   10266
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "empr_tx_fantasia"
            BoundColumn     =   "empr_cd_empresa"
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
         Begin Threed.SSCommand cmdLimparEmpresa 
            Height          =   315
            Left            =   7455
            TabIndex        =   163
            Top             =   240
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmPadraoEmpresa.frx":0635
            Picture         =   "FrmPadraoEmpresa.frx":0A87
         End
         Begin VB.Label lblMsg 
            Caption         =   "Obs.: Os lançamentos contábeis serão copiados automaticamente para a empresa vinculada no fechamento  contábil de cada mês."
            Height          =   375
            Left            =   120
            TabIndex        =   164
            Top             =   600
            Width           =   7815
         End
         Begin VB.Label Label35 
            Caption         =   "Empresa:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   735
            TabIndex        =   162
            Top             =   315
            Width           =   705
         End
      End
      Begin VB.Frame FramePadrao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Index           =   4
         Left            =   -74880
         TabIndex        =   147
         Top             =   720
         Width           =   8055
         Begin VB.CheckBox ChkSimplesNacional 
            Caption         =   "Simples Nacional"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1515
            TabIndex        =   157
            Top             =   2880
            Width           =   1575
         End
         Begin MSMask.MaskEdBox NrIRPJ 
            Height          =   315
            Left            =   1515
            TabIndex        =   148
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NrPIS 
            Height          =   315
            Left            =   1515
            TabIndex        =   150
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NrCofins 
            Height          =   315
            Left            =   1515
            TabIndex        =   152
            Top             =   1800
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NrCsll 
            Height          =   315
            Left            =   1515
            TabIndex        =   154
            Top             =   2400
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NrCPP 
            Height          =   315
            Left            =   1515
            TabIndex        =   159
            Top             =   3360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin VB.Label Label34 
            Caption         =   "Alíquota CPP:"
            Height          =   375
            Left            =   480
            TabIndex        =   158
            Top             =   3420
            Width           =   1095
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Alíquota CSLL:"
            Height          =   255
            Left            =   150
            TabIndex        =   155
            Top             =   2460
            Width           =   1290
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Alíquota COFINS:"
            Height          =   255
            Left            =   150
            TabIndex        =   153
            Top             =   1860
            Width           =   1290
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Alíquota PIS:"
            Height          =   255
            Left            =   150
            TabIndex        =   151
            Top             =   1260
            Width           =   1290
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Alíquota IRPJ:"
            Height          =   255
            Left            =   150
            TabIndex        =   149
            Top             =   660
            Width           =   1290
         End
      End
      Begin VB.Frame FraNotasEstoque 
         Caption         =   "Notas do Estoque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   -74880
         TabIndex        =   90
         Top             =   6720
         Width           =   4515
         Begin VB.CheckBox ChkAlteraNotasEstoque 
            Caption         =   "Permitir alterar notas do Estoque no Financeiro"
            Enabled         =   0   'False
            Height          =   255
            Left            =   180
            TabIndex        =   146
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame FraVinculaCRCC 
         Caption         =   "Vincular CRxCC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   -70200
         TabIndex        =   143
         Top             =   5220
         Visible         =   0   'False
         Width           =   3370
         Begin VB.OptionButton OptVinculaCRCC 
            Caption         =   "Sim"
            Height          =   255
            Left            =   555
            TabIndex        =   145
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton OptNaoVinculaCRCC 
            Caption         =   "Não"
            Height          =   255
            Left            =   1680
            TabIndex        =   144
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.Frame FraFiltra 
         Caption         =   "Filtro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -74880
         TabIndex        =   99
         Top             =   3460
         Width           =   4515
         Begin VB.CheckBox ChkCancelados 
            Caption         =   "Incluir Cancelados"
            Height          =   270
            Left            =   2475
            TabIndex        =   116
            Top             =   315
            Width           =   1650
         End
         Begin VB.TextBox TxtQtdDiasFiltro 
            Height          =   300
            Left            =   1605
            MaxLength       =   4
            TabIndex        =   100
            Top             =   285
            Width           =   615
         End
         Begin VB.Label LblQtdDiasFiltra 
            Caption         =   "Qtd. de dias para filtro de Pag / Rec:"
            Height          =   420
            Left            =   195
            TabIndex        =   101
            Top             =   195
            Width           =   1500
         End
      End
      Begin VB.Frame FraOrdemComboDesc 
         Caption         =   "Ordenar Combo de Desconto por:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   -70200
         TabIndex        =   96
         Top             =   3060
         Width           =   3370
         Begin VB.OptionButton OptDescricao 
            Caption         =   "Descrição"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   225
            TabIndex        =   98
            Top             =   240
            Value           =   -1  'True
            Width           =   1230
         End
         Begin VB.OptionButton OptCodReduzido 
            Caption         =   "Código Reduzido"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1560
            TabIndex        =   97
            Top             =   240
            Width           =   1740
         End
      End
      Begin VB.Frame FramePadrao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Index           =   2
         Left            =   -74880
         TabIndex        =   92
         Top             =   720
         Width           =   8055
         Begin VB.Frame FramePadrao 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3975
            Index           =   3
            Left            =   0
            TabIndex        =   125
            Top             =   0
            Width           =   8055
            Begin VB.Frame FraCOAF 
               Caption         =   "Resolução COAF"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   780
               Left            =   120
               TabIndex        =   140
               Top             =   2280
               Width           =   7890
               Begin VB.TextBox TxtResponsavelCoaf 
                  DataSource      =   "DatFormulario"
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
                  Left            =   2040
                  MaxLength       =   30
                  MultiLine       =   -1  'True
                  TabIndex        =   142
                  Top             =   240
                  Width           =   5775
               End
               Begin VB.Label LblResponsavelCoaf 
                  Caption         =   "Responsável no COAF:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   141
                  Top             =   315
                  Width           =   1830
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Doação"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   780
               Left            =   90
               TabIndex        =   135
               Top             =   3120
               Width           =   7890
               Begin MSMask.MaskEdBox MaskEdBox1 
                  Height          =   315
                  Left            =   1545
                  TabIndex        =   136
                  Top             =   270
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   7
                  PromptChar      =   "_"
               End
               Begin MSDataListLib.DataCombo DataCombo1 
                  Bindings        =   "FrmPadraoEmpresa.frx":0B99
                  Height          =   330
                  Left            =   2280
                  TabIndex        =   137
                  Top             =   270
                  Width           =   5190
                  _ExtentX        =   9155
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
               Begin Threed.SSCommand SSCommand2 
                  Height          =   315
                  Left            =   7470
                  TabIndex        =   138
                  Top             =   270
                  Width           =   345
                  _Version        =   65536
                  _ExtentX        =   609
                  _ExtentY        =   556
                  _StockProps     =   78
                  Enabled         =   0   'False
                  MouseIcon       =   "FrmPadraoEmpresa.frx":0BB9
                  Picture         =   "FrmPadraoEmpresa.frx":100B
               End
               Begin MSAdodcLib.Adodc Adodc1 
                  Height          =   330
                  Left            =   3375
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   2415
                  _ExtentX        =   4260
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
                  Caption         =   "DatContaRecDoacao"
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
               Begin VB.Label Label28 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Receita de Doação:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   45
                  TabIndex        =   139
                  Top             =   315
                  Width           =   1470
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Setor de Compras"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1830
               Left            =   120
               TabIndex        =   126
               Top             =   360
               Width           =   7890
               Begin VB.TextBox Text4 
                  DataSource      =   "DatFormulario"
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
                  Left            =   1440
                  MaxLength       =   30
                  MultiLine       =   -1  'True
                  TabIndex        =   130
                  Top             =   240
                  Width           =   4470
               End
               Begin VB.TextBox Text3 
                  DataSource      =   "DatFormulario"
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
                  Left            =   1455
                  MaxLength       =   100
                  MultiLine       =   -1  'True
                  TabIndex        =   129
                  Top             =   1335
                  Width           =   4470
               End
               Begin VB.TextBox Text2 
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
                  Left            =   1455
                  MaxLength       =   14
                  TabIndex        =   128
                  Top             =   645
                  Width           =   1725
               End
               Begin VB.TextBox Text1 
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
                  Left            =   1455
                  MaxLength       =   14
                  TabIndex        =   127
                  Top             =   990
                  Width           =   1725
               End
               Begin VB.Label Label27 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Responsável:"
                  Height          =   255
                  Left            =   270
                  TabIndex        =   134
                  Top             =   360
                  Width           =   1110
               End
               Begin VB.Label Label26 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Email:"
                  Height          =   255
                  Left            =   870
                  TabIndex        =   133
                  Top             =   1395
                  Width           =   510
               End
               Begin VB.Label Label25 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Telefone:"
                  Height          =   195
                  Left            =   540
                  TabIndex        =   132
                  Top             =   705
                  Width           =   840
               End
               Begin VB.Label Label24 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Fax:"
                  Height          =   210
                  Left            =   870
                  TabIndex        =   131
                  Top             =   1050
                  Width           =   510
               End
            End
         End
         Begin VB.Frame FraDadosGerais 
            Caption         =   "Setor de Compras"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1830
            Left            =   90
            TabIndex        =   107
            Top             =   225
            Width           =   7890
            Begin VB.TextBox txtCOMFax 
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
               Left            =   1455
               MaxLength       =   14
               TabIndex        =   111
               Top             =   990
               Width           =   1725
            End
            Begin VB.TextBox txtCOMTelefone 
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
               Left            =   1455
               MaxLength       =   14
               TabIndex        =   110
               Top             =   645
               Width           =   1725
            End
            Begin VB.TextBox txtCOMEmail 
               DataSource      =   "DatFormulario"
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
               Left            =   1455
               MaxLength       =   100
               MultiLine       =   -1  'True
               TabIndex        =   109
               Top             =   1335
               Width           =   4470
            End
            Begin VB.TextBox txtCOMResponsavel 
               DataSource      =   "DatFormulario"
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
               Left            =   1455
               MaxLength       =   30
               MultiLine       =   -1  'True
               TabIndex        =   108
               Top             =   300
               Width           =   4470
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Fax:"
               Height          =   210
               Left            =   870
               TabIndex        =   115
               Top             =   1050
               Width           =   510
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               Caption         =   "Telefone:"
               Height          =   195
               Left            =   540
               TabIndex        =   114
               Top             =   705
               Width           =   840
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               Caption         =   "Email:"
               Height          =   255
               Left            =   870
               TabIndex        =   113
               Top             =   1395
               Width           =   510
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "Responsável:"
               Height          =   255
               Left            =   270
               TabIndex        =   112
               Top             =   360
               Width           =   1110
            End
         End
         Begin VB.Frame FraDoacao 
            Caption         =   "Doação"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   90
            TabIndex        =   102
            Top             =   3015
            Width           =   7890
            Begin MSMask.MaskEdBox MskReceitaDoacao 
               Height          =   315
               Left            =   1545
               TabIndex        =   103
               Top             =   270
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               PromptChar      =   "_"
            End
            Begin MSDataListLib.DataCombo CboReceitaDoacao 
               Bindings        =   "FrmPadraoEmpresa.frx":111D
               Height          =   330
               Left            =   2280
               TabIndex        =   104
               Top             =   270
               Width           =   5190
               _ExtentX        =   9155
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
            Begin Threed.SSCommand SSCommand1 
               Height          =   315
               Left            =   7470
               TabIndex        =   106
               Top             =   270
               Width           =   345
               _Version        =   65536
               _ExtentX        =   609
               _ExtentY        =   556
               _StockProps     =   78
               Enabled         =   0   'False
               MouseIcon       =   "FrmPadraoEmpresa.frx":113D
               Picture         =   "FrmPadraoEmpresa.frx":158F
            End
            Begin MSAdodcLib.Adodc DatContaRecDoacao 
               Height          =   330
               Left            =   3375
               Top             =   180
               Visible         =   0   'False
               Width           =   2415
               _ExtentX        =   4260
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
               Caption         =   "DatContaRecDoacao"
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
            Begin VB.Label LblReceitaDoacao 
               Alignment       =   1  'Right Justify
               Caption         =   "Receita de Doação:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   45
               TabIndex        =   105
               Top             =   315
               Width           =   1470
            End
         End
      End
      Begin VB.Frame FraModeloRecibo 
         Caption         =   "Modelo Recibo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   -74880
         TabIndex        =   91
         Top             =   5970
         Visible         =   0   'False
         Width           =   4515
         Begin VB.OptionButton OptModelo2 
            Caption         =   "Modelo 2"
            Height          =   255
            Left            =   2340
            TabIndex        =   36
            Top             =   240
            Width           =   1515
         End
         Begin VB.OptionButton OptModelo1 
            Caption         =   "Modelo 1"
            Height          =   255
            Left            =   480
            TabIndex        =   35
            Top             =   240
            Width           =   1515
         End
      End
      Begin VB.Frame FraRelatorioQualiCapi 
         Caption         =   "Relatórios QUALICAPI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74880
         TabIndex        =   88
         Top             =   5205
         Visible         =   0   'False
         Width           =   4515
         Begin VB.CheckBox ChkIncluiAcrescimos 
            Caption         =   "Incluir Multa, Juros ou Deságio aos valores indexados"
            Height          =   330
            Left            =   180
            TabIndex        =   34
            Top             =   240
            Value           =   1  'Checked
            Width           =   4215
         End
      End
      Begin VB.Frame FraLogoMarca 
         Caption         =   " Logomarca: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3315
         Left            =   -74820
         TabIndex        =   85
         Top             =   780
         Width           =   7990
         Begin VB.CommandButton CmdLimpar 
            Caption         =   "Limpar"
            Height          =   375
            Left            =   3360
            TabIndex        =   47
            Top             =   2400
            Width           =   1215
         End
         Begin VB.CommandButton CmdProcura 
            Caption         =   "Procurar..."
            Height          =   375
            Left            =   6480
            TabIndex        =   45
            Top             =   600
            Width           =   1215
         End
         Begin prjDBImage.DBImage DBImage1 
            Height          =   1335
            Left            =   180
            TabIndex        =   46
            Top             =   1440
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   2355
            NOIMG           =   "FrmPadraoEmpresa.frx":16A1
            RIGHTCLICKMENU  =   0   'False
         End
         Begin Threed.SSPanel TxtArquivo 
            DataField       =   "Logomarca"
            DataSource      =   "DataPadroes"
            Height          =   435
            Left            =   180
            TabIndex        =   44
            Top             =   600
            Width           =   6150
            _Version        =   65536
            _ExtentX        =   10848
            _ExtentY        =   767
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   1
            BevelInner      =   1
            Alignment       =   1
         End
         Begin VB.Label LblNomeArquivo 
            Caption         =   "Nome Arquivo:"
            Height          =   195
            Left            =   180
            TabIndex        =   87
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label Label17 
            Caption         =   "Imagem:"
            Height          =   195
            Left            =   180
            TabIndex        =   86
            Top             =   1200
            Width           =   1140
         End
      End
      Begin VB.Frame FraPP 
         Caption         =   "Utiliza Proc. Pagto do Qualifin: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -70200
         TabIndex        =   84
         Top             =   2340
         Visible         =   0   'False
         Width           =   3370
         Begin VB.OptionButton OptNao 
            Caption         =   "Não"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1680
            TabIndex        =   41
            Top             =   240
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton OptSim 
            Caption         =   "Sim"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   555
            TabIndex        =   40
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.Frame FraQualifats 
         Caption         =   "QUALIFATS "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -70200
         TabIndex        =   82
         Top             =   4020
         Width           =   3370
         Begin VB.TextBox TxtPercentISS 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   915
            MaxLength       =   6
            TabIndex        =   43
            Top             =   570
            Width           =   630
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2295
            MaxLength       =   14
            TabIndex        =   42
            Top             =   225
            Width           =   930
         End
         Begin VB.Label LblPercentISS 
            Caption         =   "ISS (%):"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   89
            Top             =   600
            Width           =   780
         End
         Begin VB.Label LblNumDoc 
            Caption         =   "Último Nº de Documento:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   83
            Top             =   270
            Width           =   2235
         End
      End
      Begin VB.Frame FraCapiQualifin 
         Caption         =   " Interligação QUALICAPI/QUALIFIN: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   -74880
         TabIndex        =   79
         Top             =   4380
         Visible         =   0   'False
         Width           =   4515
         Begin VB.CheckBox ChkExpCompromissos 
            Caption         =   "Exportar inclusive os compromissos"
            Height          =   270
            Left            =   225
            TabIndex        =   33
            Top             =   195
            Value           =   1  'Checked
            Width           =   2835
         End
      End
      Begin VB.Frame FraImportacao 
         Caption         =   " Importação: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   -70200
         TabIndex        =   72
         Top             =   540
         Width           =   3370
         Begin VB.CheckBox ChkCCustos 
            Caption         =   "Centro de Custos"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   180
            TabIndex        =   38
            Top             =   600
            Width           =   1605
         End
         Begin VB.CheckBox ChkPlContas 
            Caption         =   "Plano de Contas"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   180
            TabIndex        =   37
            Top             =   240
            Width           =   2220
         End
         Begin MSDataListLib.DataCombo CboEmpresas 
            Bindings        =   "FrmPadraoEmpresa.frx":16BD
            Height          =   330
            Left            =   180
            TabIndex        =   39
            Top             =   1200
            Width           =   3070
            _ExtentX        =   5424
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "empr_tx_fantasia"
            BoundColumn     =   "empr_cd_empresa"
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
         Begin VB.Label LblEmpresa 
            Caption         =   "Empresa:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   73
            Top             =   960
            Width           =   825
         End
      End
      Begin VB.Frame FraExpContabilidade 
         Caption         =   " Exportação para Contabilidade: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74880
         TabIndex        =   71
         Top             =   1650
         Width           =   4515
         Begin VB.CheckBox ChkExportarQualiPat 
            Caption         =   "Exportar movimentos do QualiPat"
            Height          =   255
            Left            =   240
            TabIndex        =   167
            Top             =   960
            Width           =   4095
         End
         Begin VB.CheckBox ChkApropriarCR 
            Caption         =   "Apropriar CR a lançamentos de Ativo/Passivo"
            Height          =   255
            Left            =   240
            TabIndex        =   117
            Top             =   600
            Width           =   4095
         End
         Begin VB.TextBox TxtQtdDiasProvisao 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2820
            MaxLength       =   2
            TabIndex        =   32
            Top             =   1245
            Width           =   420
         End
         Begin VB.CheckBox ChkCaixa 
            Caption         =   "Base Caixa"
            Height          =   270
            Left            =   240
            TabIndex        =   30
            Top             =   240
            Value           =   1  'Checked
            Width           =   1110
         End
         Begin VB.CheckBox ChkComp 
            Caption         =   "Base Provisão e Caixa"
            Height          =   270
            Left            =   2160
            TabIndex        =   31
            Top             =   240
            Width           =   1965
         End
         Begin VB.Label LblQtdDiasProvisao 
            Caption         =   "Qtd. de dias para provisionamento:"
            Height          =   195
            Left            =   210
            TabIndex        =   78
            Top             =   1290
            Width           =   2625
         End
      End
      Begin VB.Frame FraMascara 
         Caption         =   " Máscaras: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   -74880
         TabIndex        =   68
         Top             =   540
         Width           =   4515
         Begin VB.TextBox TxtMascaraCC 
            Height          =   315
            Left            =   1500
            MaxLength       =   15
            TabIndex        =   29
            Top             =   570
            Width           =   2850
         End
         Begin VB.TextBox TxtMascaraPC 
            Height          =   315
            Left            =   1485
            MaxLength       =   20
            TabIndex        =   28
            Top             =   240
            Width           =   2850
         End
         Begin VB.Label LblMascaraCC 
            Alignment       =   1  'Right Justify
            Caption         =   "Centro de Custos:"
            Height          =   240
            Left            =   60
            TabIndex        =   70
            Top             =   660
            Width           =   1410
         End
         Begin VB.Label LblMascaraPC 
            Alignment       =   1  'Right Justify
            Caption         =   "Plano de Contas:"
            Height          =   240
            Left            =   180
            TabIndex        =   69
            Top             =   300
            Width           =   1290
         End
      End
      Begin VB.Frame FramePadrao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   8055
         Begin VB.CheckBox ChkCnpjObrigatorio 
            Caption         =   "CNPJ dos fornecedores obrigatório"
            Height          =   270
            Left            =   3960
            TabIndex        =   9
            Top             =   3660
            Value           =   1  'Checked
            Width           =   2910
         End
         Begin MSMask.MaskEdBox MskCpfResponsavel 
            Height          =   315
            Left            =   1515
            TabIndex        =   8
            ToolTipText     =   "CPF do responsável perante a SRF"
            Top             =   3600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "###.###.###-##"
            PromptChar      =   " "
         End
         Begin VB.TextBox TxtHomePage 
            DataSource      =   "DatFormulario"
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
            Left            =   1500
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   2760
            Width           =   6330
         End
         Begin VB.TextBox Txtfantasia 
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
            Left            =   1515
            MaxLength       =   20
            TabIndex        =   1
            Top             =   660
            Width           =   6330
         End
         Begin VB.TextBox TxtNome 
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
            Left            =   1515
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1080
            Width           =   6330
         End
         Begin VB.TextBox TxtIE 
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
            Left            =   1515
            MaxLength       =   15
            TabIndex        =   3
            Top             =   1500
            Width           =   3000
         End
         Begin VB.TextBox TxtIM 
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
            Left            =   1515
            MaxLength       =   15
            TabIndex        =   4
            Top             =   1920
            Width           =   3000
         End
         Begin VB.TextBox TxtEmail 
            DataSource      =   "DatFormulario"
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
            Left            =   1515
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   2340
            Width           =   4470
         End
         Begin MSDataListLib.DataCombo CboCCorrente 
            Bindings        =   "FrmPadraoEmpresa.frx":16D7
            Height          =   330
            Left            =   1515
            TabIndex        =   7
            Top             =   3180
            Width           =   6330
            _ExtentX        =   11165
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
         Begin MSMask.MaskEdBox MskCgc 
            Height          =   315
            Left            =   1515
            TabIndex        =   0
            ToolTipText     =   "CPF do responsável perante a SRF"
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   18
            Mask            =   "##.###.###/####-##"
            PromptChar      =   " "
         End
         Begin VB.Label LblCpfResponsavel 
            Alignment       =   1  'Right Justify
            Caption         =   "CPF Responsável:"
            Height          =   255
            Left            =   90
            TabIndex        =   81
            Top             =   3660
            Width           =   1350
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Home Page:"
            Height          =   255
            Left            =   150
            TabIndex        =   62
            Top             =   2790
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "CNPJ:"
            Height          =   255
            Left            =   150
            TabIndex        =   57
            Top             =   300
            Width           =   1290
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Fantasia:"
            Height          =   255
            Left            =   150
            TabIndex        =   56
            Top             =   705
            Width           =   1290
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Razão Social:"
            Height          =   255
            Left            =   150
            TabIndex        =   55
            Top             =   1125
            Width           =   1290
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "I. Estadual:"
            Height          =   255
            Left            =   150
            TabIndex        =   54
            Top             =   1560
            Width           =   1290
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "I. Municipal:"
            Height          =   255
            Left            =   150
            TabIndex        =   53
            Top             =   1980
            Width           =   1290
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Email:"
            Height          =   255
            Left            =   150
            TabIndex        =   52
            Top             =   2400
            Width           =   1290
         End
         Begin VB.Label LblCCorrente 
            Alignment       =   1  'Right Justify
            Caption         =   "Conta Corrente:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   150
            TabIndex        =   51
            Top             =   3240
            Width           =   1290
         End
      End
      Begin VB.Frame FramePadrao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Index           =   1
         Left            =   -74880
         TabIndex        =   49
         Top             =   720
         Width           =   8055
         Begin VB.TextBox txtCodigoMunicipioIbge 
            Height          =   315
            Left            =   5520
            TabIndex        =   18
            Top             =   1980
            Width           =   2240
         End
         Begin VB.TextBox TxtCGA 
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
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   16
            Top             =   1560
            Width           =   2955
         End
         Begin VB.CheckBox ChkEmail 
            Caption         =   "Enviar e-mails?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6240
            TabIndex        =   21
            Top             =   3000
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Frame FraEmail 
            Caption         =   "Configurações de E-mail: "
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   120
            TabIndex        =   118
            Top             =   3240
            Visible         =   0   'False
            Width           =   7815
            Begin VB.TextBox Senha 
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
               Height          =   315
               HelpContextID   =   1090
               IMEMode         =   3  'DISABLE
               Left            =   5640
               MaxLength       =   10
               PasswordChar    =   "*"
               TabIndex        =   27
               Tag             =   "&Senha:"
               Top             =   960
               Width           =   1965
            End
            Begin VB.TextBox TxtEmailAutent 
               Enabled         =   0   'False
               Height          =   315
               Left            =   840
               TabIndex        =   26
               Top             =   960
               Width           =   4080
            End
            Begin VB.CheckBox ChkAutenticar 
               Caption         =   "Servidor de saída requer autenticação"
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   4440
               TabIndex        =   25
               Top             =   240
               Width           =   3135
            End
            Begin VB.TextBox TxtPorta 
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
               Left            =   2760
               MaxLength       =   50
               TabIndex        =   23
               Top             =   600
               Width           =   600
            End
            Begin VB.TextBox TxtBanco 
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
               Left            =   840
               MaxLength       =   50
               TabIndex        =   22
               Top             =   600
               Width           =   1200
            End
            Begin VB.TextBox TxtServidor 
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
               Left            =   4200
               MaxLength       =   50
               TabIndex        =   24
               Top             =   600
               Width           =   3400
            End
            Begin VB.Label LblSenha 
               Caption         =   "Senha:"
               Enabled         =   0   'False
               Height          =   255
               Left            =   5040
               TabIndex        =   124
               Top             =   960
               Width           =   735
            End
            Begin VB.Label LblEmailAutent 
               Caption         =   "Usuário:"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   123
               Top             =   960
               Width           =   735
            End
            Begin VB.Label Label23 
               Caption         =   "Banco:"
               Height          =   255
               Left            =   120
               TabIndex        =   121
               Top             =   600
               Width           =   615
            End
            Begin VB.Label LblPorta 
               Caption         =   "Porta:"
               Height          =   255
               Left            =   2160
               TabIndex        =   120
               Top             =   600
               Width           =   495
            End
            Begin VB.Label LblServidor 
               Caption         =   "Servidor:"
               Height          =   255
               Left            =   3480
               TabIndex        =   119
               Top             =   600
               Width           =   975
            End
         End
         Begin VB.TextBox TxtFax 
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
            Left            =   1140
            MaxLength       =   14
            TabIndex        =   20
            Top             =   2820
            Width           =   1725
         End
         Begin VB.TextBox TxtTelefone 
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
            Left            =   1140
            MaxLength       =   14
            TabIndex        =   19
            Top             =   2400
            Width           =   1725
         End
         Begin VB.TextBox TxtPais 
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
            Left            =   1140
            MaxLength       =   15
            TabIndex        =   17
            Top             =   1980
            Width           =   2625
         End
         Begin VB.TextBox TxtEstado 
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
            Left            =   1140
            MaxLength       =   20
            TabIndex        =   15
            Top             =   1560
            Width           =   2985
         End
         Begin VB.TextBox TxtCEP 
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
            Left            =   6360
            MaxLength       =   10
            TabIndex        =   14
            Top             =   1140
            Width           =   1395
         End
         Begin VB.TextBox TxtCidade 
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
            Left            =   1140
            MaxLength       =   20
            TabIndex        =   13
            Top             =   1140
            Width           =   4545
         End
         Begin VB.TextBox TxtNumPorta 
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
            Left            =   6360
            MaxLength       =   10
            TabIndex        =   12
            Top             =   720
            Width           =   1395
         End
         Begin VB.TextBox TxtBairro 
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
            Left            =   1140
            MaxLength       =   20
            TabIndex        =   11
            Top             =   720
            Width           =   4545
         End
         Begin VB.TextBox TxtEndereco 
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
            Left            =   1140
            MaxLength       =   60
            TabIndex        =   10
            Top             =   300
            Width           =   6600
         End
         Begin VB.Label Label33 
            Caption         =   "Cód. Município IBGE:"
            Height          =   255
            Left            =   3960
            TabIndex        =   156
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label LblCga 
            Caption         =   "CGA:"
            Height          =   255
            Left            =   4320
            TabIndex        =   122
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Fax:"
            Height          =   210
            Left            =   570
            TabIndex        =   67
            Top             =   2880
            Width           =   510
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Telefone:"
            Height          =   195
            Left            =   240
            TabIndex        =   66
            Top             =   2460
            Width           =   840
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "País:"
            Height          =   240
            Left            =   555
            TabIndex        =   65
            Top             =   2040
            Width           =   525
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Estado:"
            Height          =   225
            Left            =   300
            TabIndex        =   64
            Top             =   1620
            Width           =   780
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "CEP:"
            Height          =   180
            Left            =   5700
            TabIndex        =   63
            Top             =   1200
            Width           =   555
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Cidade:"
            Height          =   225
            Left            =   150
            TabIndex        =   61
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Nº:"
            Height          =   210
            Left            =   5745
            TabIndex        =   60
            Top             =   765
            Width           =   510
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Bairro:"
            Height          =   240
            Left            =   480
            TabIndex        =   59
            Top             =   780
            Width           =   600
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Endereço:"
            Height          =   225
            Left            =   180
            TabIndex        =   58
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame FraCategoriaNvPlContas 
         Height          =   1275
         Left            =   -70200
         TabIndex        =   74
         Top             =   3360
         Visible         =   0   'False
         Width           =   3315
         Begin VB.TextBox TxtNivelCatFin 
            Height          =   315
            Left            =   840
            MaxLength       =   3
            TabIndex        =   77
            Top             =   800
            Width           =   855
         End
         Begin VB.Label Label18 
            Caption         =   "SEM USO - pode ser excluida a variavel de todo o sistema"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   825
            Left            =   1905
            TabIndex        =   80
            Top             =   495
            Width           =   1515
         End
         Begin VB.Label LblNivel 
            Caption         =   "Nível:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   76
            Top             =   840
            Width           =   450
         End
         Begin VB.Label LblCatNvFinanceiro 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Nível de Categoria que será Contabilizado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   300
            TabIndex        =   75
            Top             =   240
            Width           =   3000
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame FramePadrao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Index           =   5
         Left            =   -74880
         TabIndex        =   170
         Top             =   480
         Width           =   8055
         Begin MSDataListLib.DataCombo CboCNAESecao 
            Bindings        =   "FrmPadraoEmpresa.frx":16F6
            DataSource      =   "DatCNAESecao"
            Height          =   330
            Left            =   1440
            TabIndex        =   171
            Top             =   240
            Width           =   6330
            _ExtentX        =   11165
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "COMBOSecao"
            BoundColumn     =   "cnse_cd_CnaeSecao"
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
         Begin MSDataListLib.DataCombo CboCNAEDivisao 
            Bindings        =   "FrmPadraoEmpresa.frx":1711
            DataSource      =   "DatCNAEDivisao"
            Height          =   330
            Left            =   1440
            TabIndex        =   173
            Top             =   720
            Width           =   6330
            _ExtentX        =   11165
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "COMBODivisao"
            BoundColumn     =   "cndi_cd_CnaeDivisao"
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
         Begin MSDataListLib.DataCombo CboCNAEGrupo 
            Bindings        =   "FrmPadraoEmpresa.frx":172E
            DataSource      =   "DatCNAEGrupo"
            Height          =   330
            Left            =   1440
            TabIndex        =   175
            Top             =   1200
            Width           =   6330
            _ExtentX        =   11165
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "COMBOGrupo"
            BoundColumn     =   "cngr_cd_CnaeGrupo"
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
         Begin MSDataListLib.DataCombo CboCNAEClasse 
            Bindings        =   "FrmPadraoEmpresa.frx":1749
            DataSource      =   "DatCNAEClasse"
            Height          =   330
            Left            =   1440
            TabIndex        =   177
            Top             =   1680
            Width           =   6330
            _ExtentX        =   11165
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "COMBOClasse"
            BoundColumn     =   "cncl_cd_CnaeClasse"
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
         Begin VB.Label LblCNAEClasse 
            Alignment       =   1  'Right Justify
            Caption         =   "CNAE Classe:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   178
            Top             =   1740
            Width           =   1290
         End
         Begin VB.Label LblCNAEGrupo 
            Alignment       =   1  'Right Justify
            Caption         =   "CNAE Grupo:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   176
            Top             =   1260
            Width           =   1290
         End
         Begin VB.Label LblCNAEDivisao 
            Alignment       =   1  'Right Justify
            Caption         =   "CNAE Divisão:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   174
            Top             =   780
            Width           =   1290
         End
         Begin VB.Label LblCNAESecao 
            Alignment       =   1  'Right Justify
            Caption         =   "CNAE Seção:"
            Height          =   255
            Left            =   120
            TabIndex        =   172
            Top             =   300
            Width           =   1290
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGridCNAEDisponivel 
         Bindings        =   "FrmPadraoEmpresa.frx":1765
         Height          =   4305
         Left            =   -74880
         TabIndex        =   181
         Top             =   3240
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   7594
         _LayoutType     =   4
         _RowHeight      =   25
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "cnsc_tx_Codigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descrição"
         Columns(1).DataField=   "cnsc_tx_Descricao"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGridCNAEAssociado 
         Bindings        =   "FrmPadraoEmpresa.frx":1785
         Height          =   4305
         Left            =   -70560
         TabIndex        =   182
         Top             =   3240
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   7594
         _LayoutType     =   4
         _RowHeight      =   25
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "cnsc_tx_Codigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descrição"
         Columns(1).DataField=   "cnsc_tx_Descricao"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
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
   End
End
Attribute VB_Name = "FrmPadraoEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ExisteEmpresa As Boolean
Dim ExistePlContas As Boolean
Dim ExisteCCustos As Boolean
Dim EmpresaAtiva As Boolean
Private Function Decrip_Senha(ByVal Senha As String) As String

    Dim sen As String
    Dim a As Integer

    sen = ""
    For a = 1 To Len(Senha)
        sen = sen + UCase$(Chr$(Asc(Mid$(Senha, a, 1)) - 1))
    Next
    Decrip_Senha = sen

End Function
Private Function Crip_Senha(ByVal Senha As String) As String
  Dim sen As String
  Dim a As Integer

  sen = ""
  For a = 1 To Len(Senha)
      sen = sen + UCase$(Chr$(Asc(Mid$(Senha, a, 1)) + 1))
  Next
  Crip_Senha = sen
End Function

Sub SubGravarImagemBanco(XLA_RECORDSET As ADODB.Recordset, XLS_LocalImagem As String)
  XLA_RECORDSET("empr_im_logo") = DBImage1.PhotoField
End Sub

Sub ImportaPContasCCustos(Chave As Integer)
    Dim Padrao As Database
    Dim XCodReduzido As String, XCodEstruturado As String, XTipo As String
    Dim XRecDesp As String, xdescricao As String
    Dim xUltNivel As Byte
    Dim XCodNovo As String, XResponsavel As String, XEndereco As String
    Dim XConta As String, XCodDesc As String, Sql As String
    Dim XCodForncli As Integer
    Dim ResPadrao As ADODB.Recordset
    Dim XLI_TAMNIVELPCR As Integer
    Dim XLI_TAMNIVELCONTA As Integer
    Dim XCodEstruturadoSM As String
    Dim Xnatureza As String
    Dim xTaxaAdm As String
    Dim XStatus As String
        
    'On Error GoTo TrataErro
   
    If ChkPlContas.Value = 1 Then 'Importando o Plano de contas
        'Copia Plano de Contas
        SubQOpenRecordset ResPadrao, "SELECT * FROM PlanoContas WHERE empr_cd_empresa=" & CboEmpresas.BoundText & "", Dinamico
        
        While Not ResPadrao.EOF
            XCodReduzido = ResPadrao("plco_cd_codreduzido").Value
            XCodEstruturado = ResPadrao("plco_cd_estruturado").Value
            XCodEstruturadoSM = ResPadrao("plco_cd_EstrutSM").Value
            Xnatureza = FunNuloBanco(ResPadrao("natu_cd_Natureza").Value)
            XTipo = ResPadrao("plco_tx_tipo").Value
            XRecDesp = FunNulo(ResPadrao("plco_tx_recdesp").Value)
            xdescricao = FunNulo(ResPadrao("plco_tx_desccontabil").Value)
            xUltNivel = ResPadrao("plco_cd_ultnivel").Value
            xTaxaAdm = FunNulo(ResPadrao("plco_tx_TaxaAdm").Value)
            XStatus = FunNulo(ResPadrao("plco_tx_status").Value)
            
            XLI_TAMNIVELCONTA = ResPadrao("plco_nr_TamNivel").Value
            
            If XCodNovo = "" Then
                XCodNovo = Trim(Str(FunPegaUltimo("ulti_cd_planoccr")))
            Else
                XCodNovo = XCodNovo + 1
            End If
            ResPadrao("plco_cd_ControleImp").Value = XCodNovo
            
            ResPadrao.Update
                
            Sql = "INSERT INTO PlanoContas (plco_cd_conta, plco_cd_CodReduzido,plco_cd_estruturado, plco_tx_desccontabil, plco_tx_recdesp, plco_cd_ultnivel, plco_tx_tipo,empr_cd_empresa, plco_nr_TamNivel, plco_cd_EstrutSM,natu_cd_Natureza, plco_tx_TaxaAdm, plco_tx_status) " & _
                  "VALUES ('" & XCodNovo & "', '" & XCodReduzido & "','" & XCodEstruturado & "','" & FunNulo(xdescricao) & "','" & XRecDesp & "'," & xUltNivel & ",'" & XTipo & "', " & Chave & ", " & XLI_TAMNIVELCONTA & ", '" & XCodEstruturadoSM & "' , " & Xnatureza & ", '" & xTaxaAdm & "', '" & XStatus & "')"
                  Conexao.Execute Sql
            ResPadrao.MoveNext
        Wend
              
        ResPadrao.Close
        Set ResPadrao = Nothing
            
        Call subGuardaUltimo("ulti_cd_planoccr", CLng(XCodNovo))
    End If
    
    If ChkCCustos.Value = 1 Then 'Imposrtando o Centro de Custos
        'Copia Centro de Custos
        XCodNovo = ""
        SubQOpenRecordset ResPadrao, "SELECT * FROM CentroResultados WHERE empr_cd_empresa=" & CboEmpresas.BoundText & "", Dinamico
                
        While Not ResPadrao.EOF
            xdescricao = ResPadrao("cere_tx_descricao").Value
            XResponsavel = FunNulo(ResPadrao("cere_tx_responsavel").Value)
            XEndereco = FunNulo(ResPadrao("cere_tx_endereco").Value)
            XCodEstruturado = ResPadrao("cere_cd_estruturado").Value
            XTipo = ResPadrao("cere_tx_tipo").Value
            xUltNivel = ResPadrao("cere_cd_ultnivel").Value
            XLI_TAMNIVELPCR = ResPadrao("cere_nr_TamNivel").Value
            XCodEstruturadoSM = Replace(RTrim(LTrim(XCodEstruturado)), ".", "")
            XStatus = FunNulo(ResPadrao("cere_tx_status").Value)
            
            If XCodNovo = "" Then
                XCodNovo = Trim(Str(FunPegaUltimo("ulti_cd_pcr")))
            Else
                XCodNovo = XCodNovo + 1
            End If
            ResPadrao("cere_cd_ControleImp").Value = XCodNovo
            
            ResPadrao.Update
                
            Sql = "INSERT INTO CentroResultados (cere_cd_pcr, cere_cd_estruturado, cere_tx_endereco,cere_tx_responsavel, cere_tx_descricao, cere_cd_ultnivel, cere_tx_tipo,empr_cd_empresa, cere_bl_AlmoxCentral, cere_bl_ContrAlmoxarifado, cere_nr_TamNivel, cere_cd_EstrutSM, cere_tx_status) " & _
                  "VALUES (" & CStr(XCodNovo) & ", '" & XCodEstruturado & "','" & XEndereco & "','" & XResponsavel & "','" & FunNulo(xdescricao) & "'," & xUltNivel & ",'" & XTipo & "', " & Chave & ", 'N','N'," & XLI_TAMNIVELPCR & ",'" & XCodEstruturadoSM & "','" & XStatus & "')"
                  Conexao.Execute Sql
            ResPadrao.MoveNext
        Wend
        ResPadrao.Close
        Set ResPadrao = Nothing
        Call subGuardaUltimo("ulti_cd_pcr", CLng(XCodNovo))
    End If
    
    If ChkPlContas.Value = 1 Then 'Importou o Plano de Contas, deve atualizar tabelas de relacionamento
        'Copia tabela de relacionamento PlCDescontosImpostos atualizando os códigos para a nova empresa
        SubQOpenRecordset ResPadrao, "SELECT PlCDescontosImpostos.*, PlanoContas.plco_cd_ControleImp AS Conta FROM PlCDescontosImpostos LEFT JOIN PlanoContas ON PlCDescontosImpostos.plco_cd_Conta = PlanoContas.plco_cd_conta WHERE PlCDescontosImpostos.empr_cd_empresa=" & CboEmpresas.BoundText & "", Dinamico
        
        While Not ResPadrao.EOF
              XConta = FunNulo(ResPadrao("Conta").Value)
              XCodDesc = ResPadrao("desc_cd_desconto").Value
              
              Sql = "INSERT INTO PlCDescontosImpostos (desc_cd_desconto, empr_cd_empresa,plco_cd_Conta) " & _
                    "VALUES (" & XCodDesc & "," & Chave & ",'" & XConta & "')"
                    Conexao.Execute Sql
              ResPadrao.MoveNext
        Wend
        ResPadrao.Close
        Set ResPadrao = Nothing
        'Copia tabela de relacionamento PlCDescontosAcrescimos atualizando os códigos para a nova empresa
        SubQOpenRecordset ResPadrao, "SELECT PlCDescontosAcrescimos.*, PlanoContas.plco_cd_ControleImp AS Conta FROM PlCDescontosAcrescimos LEFT JOIN PlanoContas ON PlCDescontosAcrescimos.plco_cd_Conta = PlanoContas.plco_cd_conta WHERE PlCDescontosAcrescimos.empr_cd_empresa=" & CboEmpresas.BoundText & "", Dinamico
        
        While Not ResPadrao.EOF
              XConta = FunNulo(ResPadrao("Conta").Value)
              XCodDesc = ResPadrao("deac_cd_descacresc").Value
              
              Sql = "INSERT INTO PlCDescontosAcrescimos (deac_cd_descacresc, empr_cd_empresa,plco_cd_Conta) " & _
                    "VALUES (" & XCodDesc & "," & Chave & ",'" & XConta & "')"
                    Conexao.Execute Sql
              ResPadrao.MoveNext
        Wend
        ResPadrao.Close
        Set ResPadrao = Nothing
    Else
        If XInserir = "I" Then
            If Not ExistePlContas Then 'Não importou o plano de contas, Não possui um Plano, mas deve gerar registros vazios na tabela de relacionamento
                'Gera para cada registro de Descontos Impostos um correspondente em PlCDescontosImpostos
                SubQOpenRecordset ResPadrao, "SELECT DescontosImpostos.desc_cd_desconto FROM DescontosImpostos", Estatico
                       
                While Not ResPadrao.EOF
                      XCodDesc = ResPadrao("desc_cd_desconto").Value
                      
                      Sql = "INSERT INTO PlCDescontosImpostos (desc_cd_desconto, empr_cd_empresa) " & _
                            "VALUES (" & XCodDesc & "," & Chave & ")"
                            Conexao.Execute Sql
                      ResPadrao.MoveNext
                Wend
                ResPadrao.Close
                Set ResPadrao = Nothing
                  
                'Copia tabela de relacionamento PlCDescontosAcrescimos atualizando os códigos para a nova empresa
                SubQOpenRecordset ResPadrao, "SELECT DescontosAcrescimos.deac_cd_descacresc FROM DescontosAcrescimos", Estatico
                
                While Not ResPadrao.EOF
                      XCodDesc = ResPadrao("deac_cd_descacresc").Value
                      
                      Sql = "INSERT INTO PlCDescontosAcrescimos (deac_cd_descacresc, empr_cd_empresa) " & _
                            "VALUES (" & XCodDesc & "," & Chave & ")"
                            Conexao.Execute Sql
                      ResPadrao.MoveNext
                Wend
                ResPadrao.Close
                Set ResPadrao = Nothing
            End If
            
            If Not ExisteCCustos Then 'Não importou o Centro de custos, Não possui um, mas deve gerar registros vazios na tabela de relacionamento
            'If ChkCCustos.Value = 0 Then
                'Copia tabela de relacionamento PlCCustoFornCli atualizando os códigos para a nova empresa
                SubQOpenRecordset ResPadrao, "SELECT * FROM FornClientes", Estatico
                
                While Not ResPadrao.EOF
                    XCodForncli = ResPadrao("focl_cd_forncli").Value
            
                    Sql = "INSERT INTO PlccustoForncli (focl_cd_forncli, empr_cd_empresa) " & _
                          "VALUES (" & XCodForncli & "," & Chave & ")"
            
                          Conexao.Execute Sql
                    ResPadrao.MoveNext
                Wend
                ResPadrao.Close
                Set ResPadrao = Nothing
                'Conexao.CommitTrans
                Exit Sub
            End If
        End If
    End If

    'Copia tabela de relacionamento PlCCustoFornCli atualizando os códigos para a nova empresa
    
    If ChkCCustos.Value = 1 Or ChkPlContas.Value = 1 Then
        SubQOpenRecordset ResPadrao, "SELECT * FROM ConsGENPlCCustoFornCli WHERE empr_cd_empresa=" & CboEmpresas.BoundText & "", Dinamico
    Dim XContaRecDesp As String, XContaCusto As String
    Dim XContaProvAdiant As String, XPcr As String
    
        While Not ResPadrao.EOF
            XContaRecDesp = FunNulo(ResPadrao("ContaRecDesp").Value)
            XContaCusto = FunNulo(ResPadrao("ContaCusto").Value)
            XContaProvAdiant = FunNulo(ResPadrao("ContaProvAdiant").Value)
            XPcr = FunNulo(ResPadrao("cere_cd_ControleImp").Value)
            XCodForncli = ResPadrao("focl_cd_forncli").Value
    
            'Sql = "INSERT INTO PlccustoForncli (focl_cd_forncli, empr_cd_empresa,plco_cd_ContaRecDesp,plco_cd_ContaCusto,plco_cd_contaprovadiant,cere_cd_pcr) " & _
                  '"VALUES (" & XCodForncli & "," & Chave & ",'" & funNull(XContaRecDesp) & "','" & FunNulo(XContaCusto) & "','" & FunNulo(XContaProvAdiant) & "','" & FunNulo(XPcr) & "')"
                  
                  If FunNulo(XContaRecDesp) = "" And FunNulo(XContaCusto) = "" And FunNulo(XContaProvAdiant) = "" And FunNulo(XPcr) = "" Then
                    Sql = "INSERT INTO PlccustoForncli (focl_cd_forncli, empr_cd_empresa) VALUES (" & XCodForncli & "," & Chave & ")"
                ElseIf Len(XCodForncli) > 0 And Len(Chave) > 0 And Len(XContaRecDesp) > 0 And FunNulo(XContaCusto) = "" And FunNulo(XContaProvAdiant) = "" And FunNulo(XPcr) = "" Then
                    Sql = "INSERT INTO PlccustoForncli (focl_cd_forncli, empr_cd_empresa,plco_cd_ContaRecDesp) " & "VALUES (" & XCodForncli & "," & Chave & ",'" & funNull(XContaRecDesp) & "')"
                Else
                    Sql = "INSERT INTO PlccustoForncli (focl_cd_forncli, empr_cd_empresa,plco_cd_ContaRecDesp,plco_cd_ContaCusto,plco_cd_contaprovadiant,cere_cd_pcr) " & "VALUES (" & XCodForncli & "," & Chave & ",'" & funNull(XContaRecDesp) & "','" & FunNulo(XContaCusto) & "','" & FunNulo(XContaProvAdiant) & "','" & FunNulo(XPcr) & "')"
                End If
             

                  Conexao.Execute Sql
            ResPadrao.MoveNext
        Wend
        ResPadrao.Close
        Set ResPadrao = Nothing
    End If
    
    
    Unload Me
    
Exit Sub
    
TrataErro:
    
    funTrataErros (ComMensagem)
End Sub


Sub PreencheCampos()
  Dim ResPadrao As ADODB.Recordset
  
  SubQOpenRecordset ResPadrao, "SELECT * FROM Empresas WHERE empr_cd_empresa=" & Chave, "Estatico"
  If Not (ResPadrao.BOF And ResPadrao.EOF) Then
    Chave = FunNuloVal(ResPadrao("empr_cd_empresa").Value)
    MskCgc.PromptInclude = False
    MskCgc.Text = FunNuloVal(ResPadrao("empr_tx_CGC").Value)
    MskCgc.PromptInclude = True
    Txtfantasia.Text = FunNuloVal(ResPadrao("empr_tx_fantasia").Value)
    TxtNome.Text = FunNuloVal(ResPadrao("empr_tx_razaosocial").Value)
    TxtEndereco.Text = FunNulo(ResPadrao("empr_tx_endereco").Value)
    TxtBairro.Text = FunNulo(ResPadrao("empr_tx_Bairro").Value)
    TxtCidade.Text = FunNulo(ResPadrao("empr_tx_cidade").Value)
    TxtEstado.Text = FunNulo(ResPadrao("empr_tx_Estado").Value)
    TxtPais.Text = FunNulo(ResPadrao("empr_tx_pais").Value)
    txtCodigoMunicipioIbge.Text = FunNulo(ResPadrao("empr_cd_MunicipioIbge").Value)
    TxtTelefone.Text = FunNulo(ResPadrao("empr_tx_telefone").Value)
    TxtFax.Text = FunNulo(ResPadrao("empr_tx_fax").Value)
    TxtCEP.Text = FunNulo(ResPadrao("empr_tx_Cep").Value)
    TxtNumPorta.Text = FunNulo(ResPadrao("empr_nr_Porta").Value)
    TxtQtdDiasFiltro.Text = FunNulo(ResPadrao("empr_nr_QtdDiasFiltro").Value)
    NrIRPJ.Text = FunNulo(ResPadrao("empr_vl_IRPJ").Value)
    NrPIS.Text = FunNulo(ResPadrao("empr_vl_PIS").Value)
    NrCofins.Text = FunNulo(ResPadrao("empr_vl_COFINS").Value)
    NrCsll.Text = FunNulo(ResPadrao("empr_vl_CSLL").Value)
    
    If XGT_NomeAplic = "Capi" Then
        If Not IsNull(ResPadrao("empr_tx_ResponsavelCoaf").Value) Then
            TxtResponsavelCoaf.Text = FunNulo(ResPadrao("empr_tx_ResponsavelCoaf").Value)
        End If
    End If
    
    If Not IsNull(ResPadrao("Empr_cd_codcontacorrente").Value) Then
      CboCCorrente.BoundText = FunNuloVal(ResPadrao("Empr_cd_codcontacorrente").Value)
    End If
    TxtIE.Text = FunNulo(ResPadrao("empr_tx_InscrEst").Value)
    TxtIM.Text = FunNulo(ResPadrao("empr_tx_InscrMun").Value)
    TxtEmail.Text = FunNulo(ResPadrao("empr_tx_email").Value)
    TxtHomePage.Text = FunNulo(ResPadrao("empr_tx_homepage").Value)
    SubCarregarImagemBanco ResPadrao
    TxtMascaraPC.Text = FunNulo(ResPadrao("empr_tx_planocmascara").Value)
    TxtMascaraCC.Text = FunNulo(ResPadrao("empr_tx_pcrmascara").Value)
    TxtBanco.Text = FunNulo(ResPadrao("empr_tx_Banco").Value)
    TxtPorta.Text = FunNulo(ResPadrao("empr_tx_PortaSmtp").Value)
    TxtServidor.Text = FunNulo(ResPadrao("empr_tx_ServidorSmtp").Value)
    TxtCGA.Text = FunNulo(ResPadrao("empr_nr_CGA").Value)
    
    If XGT_NomeAplic = UCase("Compras") Then
        Senha.Text = Decrip_Senha(Trim$(FunNulo(ResPadrao("empr_tx_SenhaAutenticacao"))))
        TxtEmailAutent = FunNulo(ResPadrao("empr_tx_EmailAutenticacao").Value)
            
        If ResPadrao("empr_tx_EnviarEmail").Value = "S" Then
           ChkEmail.Value = 1
        Else
           ChkEmail.Value = 0
        End If
        
        If ResPadrao("empr_tx_Autentica").Value = "S" Then
           ChkAutenticar.Value = 1
        Else
           ChkAutenticar.Value = 0
        End If
    End If
    
    If FunNuloVal(ResPadrao("empr_nr_CategoriaNvContabil").Value) = 0 Then
       TxtNivelCatFin.Text = 1
    Else
       TxtNivelCatFin.Text = ResPadrao("empr_nr_CategoriaNvContabil").Value
    End If
    
    If ResPadrao("empr_tx_ExpCompromissoCapi").Value = "S" Then
       ChkExpCompromissos.Value = 1
    Else
       ChkExpCompromissos.Value = 0
    End If
    
    
    
    TxtQtdDiasProvisao.Text = FunNulo(ResPadrao("Empr_nr_qtdDiasProvisao").Value)
    If ResPadrao("Empr_tx_ExpContabil").Value = "PC" Then 'Base Provisão
       ChkComp.Value = 1
    Else 'Base Caixa
       ChkCaixa.Value = 1
    End If
    
    If ResPadrao("empr_tx_atual").Value = "S" Then
       ChkAtivar.Value = 1
       EmpresaAtiva = True
       ChkAtivar.Enabled = False
    Else
       ChkAtivar.Value = 0
       EmpresaAtiva = False
       ChkAtivar.Enabled = True
    End If
    
    If IsNull(ResPadrao("empr_tx_CpfResponsavel").Value) Or ResPadrao("empr_tx_CpfResponsavel").Value = "" Then
     MskCpfResponsavel.Text = "   .   .   -  "
    Else
     MskCpfResponsavel.PromptInclude = False
     MskCpfResponsavel.Text = ResPadrao("empr_tx_CpfResponsavel").Value
     MskCpfResponsavel.PromptInclude = True
    End If
       
    If IsNull(ResPadrao("empr_nr_DocQualifats").Value) Then
       TxtNumDoc.Text = ""
    Else
       TxtNumDoc.Text = ResPadrao("empr_nr_DocQualifats").Value
    End If
    TxtPercentISS.Text = Format(FunNulo(ResPadrao("empr_vl_percentIss").Value), "standard")
    
    If ResPadrao("empr_tx_ImprimePP").Value = "S" Then
       OptSim.Value = True  'Indica se a empresa utiliza PP (se sim deve mandar sempre direto p/ impressora)
    Else
       OptNao.Value = True
    End If
    
    If ResPadrao("empr_tx_OrdemDescontos").Value = "C" Then
        'Indica se a empresa deseja ordernar o combo de descontos por código reduzido
        OptCodReduzido.Value = True
    Else
        'Indica se a empresa deseja ordernar o combo de descontos por descricao
        OptDescricao.Value = True
    End If
    
    If ResPadrao!empr_tx_IncluiAcrescimos = "S" Then
     ChkIncluiAcrescimos.Value = 1
    Else
     ChkIncluiAcrescimos.Value = 0
    End If
    
    If ResPadrao!empr_tx_IncluiCancelados = "S" Then
        ChkCancelados.Value = 1
    Else
        ChkCancelados.Value = 0
    End If
    
    If ResPadrao!empr_tx_CrAtivoPassivo = "S" Then
     ChkApropriarCR.Value = 1
    Else
     ChkApropriarCR.Value = 0
    End If
    
    If ResPadrao!empr_tx_ExpMovPatCont = "S" Then
     ChkExportarQualiPat.Value = 1
    Else
     ChkExportarQualiPat.Value = 0
    End If
    
    If ResPadrao!empr_tx_AlertarSaldoAdiantamento = "S" Then
     chkAlertarSaldoAdiantamento.Value = 1
    Else
     chkAlertarSaldoAdiantamento.Value = 0
    End If
    
    If ResPadrao!empr_tx_ModeloRecibo = "1" Then
     OptModelo1.Value = 1
    Else
     OptModelo2.Value = 1
    End If
    
    If ResPadrao!empr_tx_VinculaCrCc = "S" Then
     OptVinculaCRCC.Value = 1
    Else
     OptNaoVinculaCRCC.Value = 1
    End If
    
    ChkCnpjObrigatorio.Value = FunNuloVal(ResPadrao!empr_tx_CnpjObrigatorio)
    
    'Preenchendo os dados da orelha do Dpto. de Compras
    txtCOMResponsavel.Text = FunNulo(ResPadrao("empr_tx_ComResp").Value)
    txtCOMTelefone.Text = FunNulo(ResPadrao("empr_tx_ComTelefone").Value)
    txtCOMFax.Text = FunNulo(ResPadrao("empr_tx_ComFax").Value)
    txtCOMEmail.Text = FunNulo(ResPadrao("empr_tx_ComEmail").Value)
    
    'Preenchendo os dados do frame de Doação
    If Not IsNull(ResPadrao!plco_cd_ReceitaDoacao) Then
        DatContaRecDoacao.Recordset.MoveFirst
        DatContaRecDoacao.Recordset.Find "plco_cd_Conta =" & ResPadrao!plco_cd_ReceitaDoacao
        MskReceitaDoacao.Text = DatContaRecDoacao.Recordset.Fields("plco_cd_Codreduzido")
        CboReceitaDoacao.BoundText = MskReceitaDoacao.Text
    End If
    
    If ResPadrao!empr_tx_SimplesNacional = "S" Then
        ChkSimplesNacional.Value = 1
    Else
        ChkSimplesNacional.Value = 0
    End If
    
    NrCPP.Text = FunNulo(ResPadrao!empr_vl_CPP)
    
    CboEmpresaVinculada.BoundText = FunNulo(ResPadrao!empr_cd_empresaVinculada)
    CboTipoPlanoContasReferencial.ListIndex = IIf(Not IsNull(ResPadrao.Fields("empr_nr_TipoPlanoReferencial").Value), ResPadrao.Fields("empr_nr_TipoPlanoReferencial").Value, 0)
    
    ExisteEmpresa = True
  End If
  ResPadrao.Close
    
End Sub

Sub PreparaInsercao()
  TxtNivelCatFin.Text = 1
End Sub

Sub SubCarregarImagemBanco(XLA_RECORDSET As ADODB.Recordset)
  DBImage1.PhotoField = XLA_RECORDSET("empr_im_logo")
End Sub

Private Sub CboCNAESecao_Change()
    CboCNAEDivisao.Text = ""
    CboCNAEGrupo.Text = ""
    CboCNAEClasse.Text = ""
    
    If CboCNAESecao.Text <> "" Then
        subConectarControleDadosNV DatCNAEDivisao, "SELECT * FROM CONSCNAEDivisao where cnse_cd_CnaeSecao = " & CboCNAESecao.BoundText & "", Estatico
    End If
End Sub

Private Sub CboCNAEDivisao_Change()
    CboCNAEGrupo.Text = ""
    CboCNAEClasse.Text = ""
    
    If CboCNAEDivisao.Text <> "" Then
        subConectarControleDadosNV DatCNAEGrupo, "SELECT * FROM CONSCNAEGrupo where cndi_cd_CnaeDivisao = " & CboCNAEDivisao.BoundText & "", Estatico
    End If
End Sub

Private Sub CboCNAEGrupo_Change()
    CboCNAEClasse.Text = ""
    
    If CboCNAEGrupo.Text <> "" Then
        subConectarControleDadosNV DatCNAEClasse, "SELECT * FROM CONSCNAEClasse WHERE cngr_cd_CnaeGrupo = " & CboCNAEGrupo.BoundText & "", Estatico
    End If
End Sub

Private Sub CboCNAEClasse_Change()

    If CboCNAEClasse.Text <> "" Then
        subConectarControleDadosNV DatCNAEDisponivel, "SELECT CnaeSubClasse.* FROM CnaeSubClasse LEFT JOIN Empresa_CNAESubClasse ON Empresa_CNAESubClasse.cnsc_cd_CnaeSubClasse = CnaeSubClasse.cnsc_cd_CnaeSubClasse WHERE cncl_cd_CnaeClasse = " & CboCNAEClasse.BoundText & " AND CnaeSubClasse.cnsc_cd_CnaeSubClasse NOT IN (SELECT cnsc_cd_CnaeSubClasse FROM Empresa_CNAESubClasse WHERE empr_cd_empresa= " & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa").Value & " )", Estatico
    End If
End Sub

Private Sub CboReceitaDoacao_Change()
    If CboReceitaDoacao.Text <> "" Then
        DatContaRecDoacao.Recordset.Bookmark = CboReceitaDoacao.SelectedItem
        If DatContaRecDoacao.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
            CboReceitaDoacao.BoundText = MskReceitaDoacao.Text
            CboReceitaDoacao.SetFocus
            Exit Sub
        End If
        MskReceitaDoacao.Text = CboReceitaDoacao.BoundText
    End If
End Sub

Private Sub ChkAutenticar_Click()
    If ChkAutenticar.Value = 1 Then
        LblEmailAutent.Enabled = True
        LblSenha.Enabled = True
        TxtEmailAutent.Enabled = True
        Senha.Enabled = True
    Else
        TxtEmailAutent.Enabled = False
        Senha.Enabled = False
        LblEmailAutent.Enabled = False
        LblSenha.Enabled = False
    End If
    
End Sub

Private Sub ChkCaixa_Click()
    If ChkCaixa.Value = 1 Then
        ChkComp.Value = 0
        TxtQtdDiasProvisao.Enabled = False
    End If
End Sub

Private Sub ChkComp_Click()
    If ChkComp.Value = 1 Then
        ChkCaixa.Value = 0
        TxtQtdDiasProvisao.Enabled = True
    End If
End Sub

Private Sub ChkEmail_Click()

    If ChkEmail.Value = 1 Then
        FraEmail.Enabled = True
    Else
        FraEmail.Enabled = False
        TxtBanco.Text = ""
        TxtPorta.Text = ""
        TxtServidor = ""
    End If
    
End Sub
Private Sub CmdGravar_Click()
  Me.MousePointer = vbHourglass
  Dim ResPadrao As Object
  Dim XNivelPC As Integer, XNivelCC As Integer, x As Integer
  Dim XMascaraPC As String, XMascaraCC As String

  ChkAlteraNotasEstoque.Enabled = False

  If FunObrigatorioMSK(MskCgc, "O preenchimento do campo CNPJ é obrigatório") Then Exit Sub
  
  If TxtNome.Text = "" Then
    MsgBox "A Razão Social da empresa é obrigatório.", vbInformation, "ATENÇÃO"
    Exit Sub
  End If
  
  If TxtNumPorta <> "" Then
      If Not IsNumeric(TxtNumPorta.Text) Then
          MsgBox "O número da porta deve ter formato numérico!", vbInformation, "AVISO"
           Me.MousePointer = vbDefault
          Exit Sub
      End If
  End If
  
  If NrIRPJ <> "" Then
      If Not IsNumeric(NrIRPJ.Text) Then
          MsgBox "O número do IRPJ deve ter formato numérico!", vbInformation, "AVISO"
           Me.MousePointer = vbDefault
          Exit Sub
      End If
  End If
  
  If NrPIS <> "" Then
      If Not IsNumeric(NrPIS.Text) Then
          MsgBox "O número do PIS deve ter formato numérico!", vbInformation, "AVISO"
           Me.MousePointer = vbDefault
          Exit Sub
      End If
  End If
  
  If NrCofins <> "" Then
      If Not IsNumeric(NrCofins.Text) Then
          MsgBox "O número do COFINS deve ter formato numérico!", vbInformation, "AVISO"
           Me.MousePointer = vbDefault
          Exit Sub
      End If
  End If
  
  If NrCsll <> "" Then
      If Not IsNumeric(NrCsll.Text) Then
          MsgBox "O número do CSLL deve ter formato numérico!", vbInformation, "AVISO"
           Me.MousePointer = vbDefault
          Exit Sub
      End If
  End If
  
  If XGB_EMAIL And XInserir <> "I" Then
  
    If TxtPorta = "" Or TxtServidor = "" Or TxtBanco = "" Then
        MsgBox ("Favor preencher os campos referentes ao envio de E-mail."), vbInformation, "Preencher campos"
        Exit Sub
    End If
    
    If ChkAutenticar.Value = 1 Then
        If Senha.Text = "" Or TxtEmailAutent = "" Then
            MsgBox ("Favor preencher os campos referentes ao envio de E-mail."), vbInformation, "Preencher campos"
        Exit Sub
        End If
    End If
  End If
  
  If ChkPlContas.Value = 1 Or ChkCCustos.Value = 1 Then
      If FunNulo(CboEmpresas.Text) = "" Then
          MsgBox "É necessário seleciona a empresa de Onde serão Copiadas as informações", vbInformation, "ATENÇÃO"
          Exit Sub
      End If
    SubQOpenRecordset ResPadrao, "SELECT empr_tx_pcrmascara,empr_nr_pcrnivel,empr_tx_planocmascara,empr_nr_planocnivel FROM empresas WHERE empr_cd_empresa = " & CboEmpresas.BoundText, Estatico
    XNivelPC = ResPadrao("empr_nr_planocnivel").Value
    XMascaraPC = ResPadrao("empr_tx_planocmascara").Value
    XNivelCC = ResPadrao("empr_nr_pcrnivel").Value
    XMascaraCC = ResPadrao("empr_tx_pcrmascara").Value
    ResPadrao.Close
  End If

  
  If ChkCCustos.Value = 0 Then
    If FunObrigatorioTXT(TxtMascaraCC, "O preenchimento do campo 'Máscara Centro de Custos' é obrigatório") Then Exit Sub
    x = 1
    XNivelCC = 1 'Se não existir ao menos um ponto, o nível será 1
    While x <= (Len(TxtMascaraCC.Text) + 1)
      If Mid(TxtMascaraCC.Text, x, 1) = "." Then
        XNivelCC = XNivelCC + 1
      End If
      x = x + 1
    Wend
    XMascaraCC = TxtMascaraCC.Text
  End If
  
  If ChkPlContas.Value = 0 Then
    If FunObrigatorioTXT(TxtMascaraPC, "O preenchimento do campo 'Máscara Plano de Contas' é obrigatório") Then Exit Sub
    x = 1
    XNivelPC = 1
    While x <= Len(TxtMascaraPC.Text)
      If Mid(TxtMascaraPC.Text, x, 1) = "." Then
        XNivelPC = XNivelPC + 1
      End If
      x = x + 1
    Wend
    XMascaraPC = TxtMascaraPC.Text
  End If
  
  'VERIFICA SE NIVEL DA CATEGORIA PARA FINANCEIRO ESTA DENTRO DO LIMITE
  If Not IsNumeric(TxtNivelCatFin.Text) Then
      MsgBox "O Nível da Categoria tem de ser um valor númerico"
      Exit Sub
  End If
  If FunNuloVal(TxtNivelCatFin.Text) > FunDefineUltimoNivel(pCategoriaMascara) Then
      MsgBox "O Nível Informado é superior ao Ultimo nível de categoria Cadastrado"
      Exit Sub
  End If
  
  If MskCpfResponsavel.Text <> "   .   .   -  " Then
    MskCpfResponsavel.PromptInclude = False
    If Not FunConfereCPF(MskCpfResponsavel.Text) Then
      MsgBox "O Cpf do Responsável perante a SRF está incorreto.", vbExclamation, "ATENÇÃO"
      MskCpfResponsavel.PromptInclude = True
      Exit Sub
    End If
  MskCpfResponsavel.PromptInclude = True
  End If
  
  If ChkSimplesNacional.Value = 1 And NrCPP.Text = "" Then
    MsgBox "É obrigatório informar o CPP quando a empresa for Super Simples.", vbInformation, "ATENÇÃO"
    Exit Sub
  End If
  
  'On Error GoTo TrataErro
  
  Conexao.BeginTrans
  If Not ExisteEmpresa Then
    'Abre o banco de Dados
    SubQOpenRecordset ResPadrao, "SELECT * FROM empresas WHERE empr_cd_empresa= -1", Dinamico
    ResPadrao.AddNew
  Else    'Alteração
    SubQOpenRecordset ResPadrao, "SELECT * FROM empresas WHERE empr_cd_empresa = " & Chave, Dinamico
  End If
  
  ResPadrao("empr_tx_CGC").Value = FunNulo(MskCgc.Text)
  ResPadrao("empr_tx_razaosocial").Value = FunNulo(TxtNome.Text)
  ResPadrao("empr_tx_fantasia").Value = FunNulo(Txtfantasia.Text)
  ResPadrao("empr_tx_endereco").Value = FunNulo(TxtEndereco.Text)
  ResPadrao("empr_nr_Porta").Value = FunNulo(TxtNumPorta.Text)
  ResPadrao("empr_tx_Bairro").Value = FunNulo(TxtBairro.Text)
  ResPadrao("empr_tx_cidade").Value = FunNulo(TxtCidade.Text)
  ResPadrao("empr_tx_Estado").Value = FunNulo(TxtEstado.Text)
  ResPadrao("empr_tx_pais").Value = FunNulo(TxtPais.Text)
  ResPadrao("empr_cd_MunicipioIbge").Value = FunNulo(txtCodigoMunicipioIbge.Text)
  ResPadrao("empr_tx_Cep").Value = FunNulo(TxtCEP.Text)
  ResPadrao("empr_tx_telefone").Value = FunNulo(TxtTelefone.Text)
  If NrIRPJ.Text = "" Then
     ResPadrao("empr_vl_IRPJ").Value = Null
  Else
     ResPadrao("empr_vl_IRPJ").Value = funNull(NrIRPJ.Text)
  End If
  If NrPIS.Text = "" Then
     ResPadrao("empr_vl_PIS").Value = Null
  Else
     ResPadrao("empr_vl_PIS").Value = funNull(NrPIS.Text)
 End If
  If NrCofins.Text = "" Then
     ResPadrao("empr_vl_COFINS").Value = Null
     Else
        ResPadrao("empr_vl_COFINS").Value = FunNulo(NrCofins.Text)
 End If
  If NrCsll.Text = "" Then
     ResPadrao("empr_vl_CSLL").Value = Null
  Else
        ResPadrao("empr_vl_CSLL").Value = funNull(NrCsll.Text)
  End If
  ResPadrao("empr_tx_fax").Value = FunNulo(TxtFax.Text)
  If CboCCorrente.BoundText <> "" Then ResPadrao("Empr_cd_codcontacorrente").Value = CboCCorrente.BoundText
  ResPadrao("empr_tx_InscrEst").Value = FunNulo(TxtIE.Text)
  ResPadrao("empr_tx_InscrMun").Value = FunNulo(TxtIM.Text)
  ResPadrao("empr_tx_email").Value = FunNulo(TxtEmail.Text)
  ResPadrao("empr_tx_homepage").Value = FunNulo(TxtHomePage.Text)
  ResPadrao("empr_tx_planocmascara").Value = XMascaraPC
  ResPadrao("empr_nr_planocnivel").Value = XNivelPC
  ResPadrao("empr_tx_pcrmascara").Value = XMascaraCC
  ResPadrao("empr_nr_pcrnivel").Value = XNivelCC
  ResPadrao("empr_nr_CategoriaNvContabil").Value = FunNulo(TxtNivelCatFin.Text)
  ResPadrao("Empr_nr_qtdDiasProvisao").Value = FunNuloVal(TxtQtdDiasProvisao.Text)
  ResPadrao("empr_tx_CpfResponsavel").Value = MskCpfResponsavel
  ResPadrao("empr_nr_DocQualifats").Value = funNull(TxtNumDoc.Text)
  ResPadrao("empr_vl_PercentISS").Value = funNull(TxtPercentISS.Text)
  ResPadrao("empr_nr_QtdDiasFiltro").Value = funNull(TxtQtdDiasFiltro.Text)
  ResPadrao("empr_tx_Banco").Value = funNull(TxtBanco.Text)
  ResPadrao("empr_tx_ServidorSmtp").Value = funNull(TxtServidor.Text)
  ResPadrao("empr_tx_PortaSmtp").Value = funNull(TxtPorta.Text)
  ResPadrao("empr_nr_Cga").Value = funNull(TxtCGA.Text)
  ResPadrao("empr_cd_empresaVinculada").Value = IIf(CboEmpresaVinculada.Text <> "", CboEmpresaVinculada.BoundText, Null)
  If CboTipoPlanoContasReferencial.ListIndex > -1 Then
    ResPadrao("empr_nr_TipoPlanoReferencial").Value = CboTipoPlanoContasReferencial.ItemData(CboTipoPlanoContasReferencial.ListIndex)
  Else
    ResPadrao("empr_nr_TipoPlanoReferencial").Value = Null
  End If

  
  If XGT_NomeAplic = UCase("Compras") Then
      ResPadrao("empr_tx_EmailAutenticacao").Value = funNull(TxtEmailAutent.Text)
      ResPadrao("empr_tx_SenhaAutenticacao").Value = Crip_Senha(Trim$(FunNulo(Senha.Text)))
      ResPadrao("empr_tx_ResponsavelCoaf").Value = funNull(TxtResponsavelCoaf.Text)
      
      If ChkEmail.Value = 1 Then
        ResPadrao!empr_tx_EnviarEmail = "S"
      Else
        ResPadrao!empr_tx_EnviarEmail = "N"
      End If
    
      If ChkAutenticar.Value = 1 Then
        ResPadrao!empr_tx_Autentica = "S"
      Else
        ResPadrao!empr_tx_Autentica = "N"
      End If
  End If
  
  ' Indica que a empresa pode ou não alterar as notas de Estoque - Lucas Santiago 23.01.09
  If ChkAlteraNotasEstoque = 1 Then
    ResPadrao!empr_tx_AltNotasEstoque = "S"
  Else
    ResPadrao!empr_tx_AltNotasEstoque = "N"
  End If
  
  If ChkApropriarCR.Value = 1 Then
    ResPadrao!empr_tx_CrAtivoPassivo = "S"
  Else
    ResPadrao!empr_tx_CrAtivoPassivo = "N"
  End If
  
  If ChkExportarQualiPat.Value = 1 Then
    ResPadrao!empr_tx_ExpMovPatCont = "S"
  Else
    ResPadrao!empr_tx_ExpMovPatCont = "N"
  End If
  
  If chkAlertarSaldoAdiantamento.Value = 1 Then
    ResPadrao!empr_tx_AlertarSaldoAdiantamento = "S"
  Else
    ResPadrao!empr_tx_AlertarSaldoAdiantamento = "N"
  End If
  
  If ChkCancelados.Value = 1 Then
    ResPadrao!empr_tx_IncluiCancelados = "S"
    XGB_INCLUICANCELADOS = True
  Else
    ResPadrao!empr_tx_IncluiCancelados = "N"
    XGB_INCLUICANCELADOS = False
  End If
  
  If TxtQtdDiasFiltro.Text = "" Then
    XGI_QTDDIASFILTRO = 0
  Else
    XGI_QTDDIASFILTRO = TxtQtdDiasFiltro.Text
  End If

  'Indica se a empresa utiliza PP (se sim deve mandar sempre direto p/ impressora)
  If OptSim.Value = True Then
      ResPadrao("empr_tx_ImprimePP").Value = "S"
  Else
      ResPadrao("empr_tx_ImprimePP").Value = "N"
  End If
  
  'Indica se a empresa deseja ordenar o combo de descontos por descrição ou por código reduzido
  If OptDescricao.Value = True Then
      ResPadrao("empr_tx_OrdemDescontos").Value = "D" 'Descrição
  Else
      ResPadrao("empr_tx_OrdemDescontos").Value = "C" 'Código Reduzido
  End If
  
  If ChkIncluiAcrescimos.Value = 1 Then
    ResPadrao!empr_tx_IncluiAcrescimos = "S"
  Else
    ResPadrao!empr_tx_IncluiAcrescimos = "N"
  End If
  
  If OptModelo1 = True Then
    ResPadrao!empr_tx_ModeloRecibo = "1"
  Else
    ResPadrao!empr_tx_ModeloRecibo = "2"
  End If
  
  ResPadrao!empr_tx_CnpjObrigatorio = ChkCnpjObrigatorio.Value
  
  If ChkExpCompromissos Then
      ResPadrao("empr_tx_ExpCompromissoCapi").Value = "S"
  Else
      ResPadrao("empr_tx_ExpCompromissoCapi").Value = "N"
  End If
  
  If ChkAtivar.Value = 1 Then
    ResPadrao("empr_tx_atual").Value = "S"
  Else
    ResPadrao("empr_tx_atual").Value = "N"
  End If
  
  If ChkComp.Value = 1 Then  'Base Provisão
      ResPadrao("Empr_tx_ExpContabil").Value = "PC"
  Else
      ResPadrao("Empr_tx_ExpContabil").Value = "CX" 'Base Caixa
  End If
  
  'Grava logotipo
  SubGravarImagemBanco ResPadrao, TxtArquivo.Caption
  
  'Grava dados da orelha do Dpto. de Compras
  ResPadrao("empr_tx_ComResp").Value = txtCOMResponsavel.Text
  ResPadrao("empr_tx_ComTelefone").Value = txtCOMTelefone.Text
  ResPadrao("empr_tx_ComFax").Value = txtCOMFax.Text
  ResPadrao("empr_tx_ComEmail").Value = txtCOMEmail.Text
  
  'Gravando os dados
  If CboReceitaDoacao.Text <> "" Then
      DatContaRecDoacao.Recordset.Bookmark = CboReceitaDoacao.SelectedItem
      ResPadrao!plco_cd_ReceitaDoacao = DatContaRecDoacao.Recordset.Fields("plco_cd_conta")
  End If
  
  'Inserido em 31/07/08 - Larissa
  If OptVinculaCRCC = True Then
    ResPadrao("empr_tx_VinculaCrCc").Value = "S"
  Else
    ResPadrao("empr_tx_VinculaCrCc").Value = "N"
  End If
  
  If ChkSimplesNacional.Value = 1 Then
        ResPadrao("empr_tx_SimplesNacional").Value = "S"
  Else
        ResPadrao("empr_tx_SimplesNacional").Value = "N"
  End If
  
  ResPadrao("empr_vl_CPP").Value = funNull(NrCPP.Text)
  
  ResPadrao.Update
  
  If XInserir = "I" Then Chave = ResPadrao("empr_cd_empresa").Value
    ResPadrao.Close
  'End If
  
    If Not EmpresaAtiva And ChkAtivar.Value = 1 Then
      'Atualizar a tabela de empresas
      SubQOpenRecordset ResPadrao, "SELECT * FROM empresas WHERE empr_cd_empresa <> " & _
      Chave & " And empr_tx_atual = 'S'", Dinamico
      ResPadrao("empr_tx_atual").Value = "N"
      ResPadrao.Update
      ResPadrao.Close
    End If
    
  If ChkAtivar.Value = 1 Then
      PCodEmpresa = Chave
      PFantasia = Txtfantasia.Text
      PEmpresa = TxtNome.Text
      PEndereco = TxtEndereco.Text
      pPCRmascara = XMascaraCC
      pPCRnivel = XNivelCC
      pPlanoCmascara = XMascaraPC
      pPlanoCNivel = XNivelPC
      'PLogo = FunNulo(Adodc1.Recordset.Fields("empr_im_caminhoLogo"))
      XGI_PCATEGORIANVFINANCEIRO = TxtNivelCatFin.Text
      If OptModelo1 = True Then
        XGT_MODELORECIBO = "1"
      Else
        XGT_MODELORECIBO = "2"
      End If
      
      XGB_CNPJOBRIGATORIO = ChkCnpjObrigatorio.Value
  
      MdiPrincipal.BarraStatus.Panels(2).Text = PFantasia
  End If
  
  If XInserir = "I" And XGB_PCAPI Then
      subImportaFormularios (Chave)
  End If
  
  ImportaPContasCCustos (Chave)
  subImportaCategorias (Chave)
  Conexao.CommitTrans
  XGB_RESPOSTA = True
  Me.MousePointer = vbDefault
  Unload Me
  Exit Sub
  
TrataErro:
  funTrataErros (ComMensagem)
End Sub

Private Sub CmdLimpar_Click()
  
  DBImage1.LoadImage ("")
  TxtArquivo.Caption = ""
  
End Sub

Private Sub cmdLimparEmpresa_Click()
    CboEmpresaVinculada.BoundText = 0
    CboEmpresaVinculada.Text = ""
End Sub

Private Sub CmdProcura_Click()
  DlgAbrirFigura.ShowOpen
  If DlgAbrirFigura.FileName <> "" Then
    DBImage1.LoadImage DlgAbrirFigura.FileName
    TxtArquivo.Caption = DlgAbrirFigura.FileName
  End If
  If Err Then
      MsgBox "Arquivo inválido !!", vbCritical, "ATENÇÃO"
  End If
End Sub

Private Sub CmdInserirCNAE_Click()
    Dim Sql As String
    Dim errLoop As ADODB.Error
    Dim StrError As String
    
    On Error GoTo MostraErro
    
    If DatCNAEDisponivel.Recordset.RecordCount = 0 Then
        MsgBox "Selecione o CNAE (a esquerda) a ser vinculado.", vbInformation + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
   
    Sql = "INSERT INTO Empresa_CNAESubClasse (cnsc_cd_CnaeSubClasse, empr_cd_Empresa) " & _
          "VALUES ('" & DatCNAEDisponivel.Recordset.Fields("cnsc_cd_CnaeSubClasse") & "'," & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa").Value & ")"
          Conexao.Execute Sql
        
        subConectarControleDadosNV DatCNAEDisponivel, "SELECT CnaeSubClasse.* FROM CnaeSubClasse LEFT JOIN Empresa_CNAESubClasse ON Empresa_CNAESubClasse.cnsc_cd_CnaeSubClasse = CnaeSubClasse.cnsc_cd_CnaeSubClasse WHERE cncl_cd_CnaeClasse = " & CboCNAEClasse.BoundText & " AND CnaeSubClasse.cnsc_cd_CnaeSubClasse NOT IN (SELECT cnsc_cd_CnaeSubClasse FROM Empresa_CNAESubClasse WHERE empr_cd_Empresa = " & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa").Value & " )", Estatico
        subConectarControleDadosNV DatCNAEAssociado, "SELECT * FROM CnaeSubClasse INNER JOIN Empresa_CNAESubClasse ON Empresa_CNAESubClasse.cnsc_cd_CnaeSubClasse = CnaeSubClasse.cnsc_cd_CnaeSubClasse WHERE Empresa_CNAESubClasse.empr_cd_Empresa = " & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa").Value, Estatico
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

Private Sub CmdRemoverCNAE_Click()
    Dim Sql As String
    Dim errLoop As ADODB.Error
    Dim StrError As String
    
    'On Error GoTo MostraErro
    
    If DatCNAEAssociado.Recordset.RecordCount = 0 Then
        MsgBox "Selecione o item do CNAE (a direita) a ser removido.", vbInformation + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If

    Sql = "DELETE FROM Empresa_CNAESubClasse WHERE " & _
               "cnsc_cd_CnaeSubClasse = " & Str(DatCNAEAssociado.Recordset.Fields("cnsc_cd_CnaeSubClasse")) & " AND " & _
               "empr_cd_Empresa = '" & DatCNAEAssociado.Recordset.Fields("empr_cd_Empresa") & "'"
    Conexao.Execute Sql
    
    If CboCNAEClasse = "" Then
        subConectarControleDadosNV DatCNAEAssociado, "SELECT * FROM CnaeSubClasse INNER JOIN Empresa_CNAESubClasse ON Empresa_CNAESubClasse.cnsc_cd_CnaeSubClasse = CnaeSubClasse.cnsc_cd_CnaeSubClasse WHERE Empresa_CNAESubClasse.empr_cd_Empresa = " & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa").Value, Estatico
        MsgBox "CNAE associado removido, caso queira associar outro CNAE, favor preencher os campos.", vbInformation, "ATENÇÃO"
    Else
        subConectarControleDadosNV DatCNAEAssociado, "SELECT * FROM CnaeSubClasse INNER JOIN Empresa_CNAESubClasse ON Empresa_CNAESubClasse.cnsc_cd_CnaeSubClasse = CnaeSubClasse.cnsc_cd_CnaeSubClasse WHERE Empresa_CNAESubClasse.empr_cd_Empresa = " & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa").Value, Estatico
        subConectarControleDadosNV DatCNAEDisponivel, "SELECT CnaeSubClasse.* FROM CnaeSubClasse LEFT JOIN Empresa_CNAESubClasse ON Empresa_CNAESubClasse.cnsc_cd_CnaeSubClasse = CnaeSubClasse.cnsc_cd_CnaeSubClasse WHERE cncl_cd_CnaeClasse = " & CboCNAEClasse.BoundText & " AND CnaeSubClasse.cnsc_cd_CnaeSubClasse NOT IN (SELECT cnsc_cd_CnaeSubClasse FROM Empresa_CNAESubClasse WHERE empr_cd_empresa= " & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa").Value & " )", Estatico
    End If
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
  ChkAlteraNotasEstoque.Enabled = False ' Inserido - Lucas Santiago 23.01.09
  XGB_RESPOSTA = False
  Unload Me
End Sub

Private Sub Form_Activate()
  FrmPadraoEmpresa.Top = 1000
  FrmPadraoEmpresa.Left = 0
End Sub
Private Sub Form_Load()
    Dim ResReg As Object
    Dim ResEstoque As Object
    Dim XLO_EMPRESA As ADODB.Recordset
  
    SubQOpenRecordset XLO_EMPRESA, "SELECT TOP 1 empr_tx_cliente FROM empresas WHERE empr_tx_cliente IS NOT NULL", Estatico
  
    If Not (XLO_EMPRESA.BOF Or XLO_EMPRESA.EOF) Then
  
        'Conecta o controle de Dados
        ExisteEmpresa = False
        EmpresaAtiva = False
        ExistePlContas = False
        ExisteCCustos = False
        
        If pTipoCopia = "Multi" Then
          subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa").Value & " ORDER BY DESCCOMBO", Estatico
          subConectarControleDadosNV DatEmpresas, "SELECT * FROM Empresas", Estatico
        Else
          subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo", Estatico
          Chave = PCodEmpresa
        End If
        
        subConectarControleDadosNV DatEmpresaPrincipal, "SELECT * FROM Empresas WHERE empr_cd_empresa <> " & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa").Value, Estatico
        subConectarControleDadosNV DatCNAESecao, "SELECT * FROM CONSCNAESecao", Estatico
        subConectarControleDadosNV DatCNAEAssociado, "SELECT * FROM CnaeSubClasse INNER JOIN Empresa_CNAESubClasse ON Empresa_CNAESubClasse.cnsc_cd_CnaeSubClasse = CnaeSubClasse.cnsc_cd_CnaeSubClasse WHERE Empresa_CNAESubClasse.empr_cd_Empresa = " & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa").Value, Estatico
        
        'Marca o checkbox chkAltNotasFiscais - Lucas Santiago - 23.01.09
        SubQOpenRecordset ResEstoque, "SELECT empr_tx_AltNotasEstoque FROM Empresas WHERE empr_cd_empresa = " & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa").Value, Estatico
        If ResEstoque!empr_tx_AltNotasEstoque = "S" Then
          ChkAlteraNotasEstoque = 1
        End If
  
        'Lucas Santiago, 14.01.09 - Removido o comando abaixo da condição If XGB_PQUALIEST
        subConectarControleDadosNV DatContaRecDoacao, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and natu_cd_natureza = 3 ORDER BY plco_tx_desccontabil", Estatico
        
        If XGB_PQUALIEST Then 'Possui módulo de Estoque - Habilitar frame de doação e conectar os datas
          FraDoacao.Enabled = True
        Else
          FraDoacao.Enabled = False
        End If
        
        If XGB_PCAPI Then     'Qdo. interligado com o capi incluir o frame para definir exportacao p/ Qualifin
          FraCapiQualifin.Visible = True
          FraRelatorioQualiCapi.Visible = True
          FraModeloRecibo.Visible = True
        End If
  
        If XGB_PQUALIFIN Then ' Torna visível o frame para indicar se imprime PP
          FraPP.Visible = True
        Else
          FraPP.Visible = False
        End If
        
        If XGB_PQUALIFATS Then ' Torna visível o frame com os padrões do qualifats
          FraQualifats.Visible = True
        Else
          FraQualifats.Visible = False
        End If
        
        If XGB_PCOMPRAS Then ' Torna visível o frame com os padrões de E-mail
          ChkEmail.Visible = True
          FraEmail.Visible = True
        Else
          ChkEmail.Visible = False
          FraEmail.Visible = False
        End If
        
        If XInserir = "A" Then
          CboCCorrente.Enabled = True
          
          'verifica Plano de Contas
          If pTipoCopia = "Multi" Then
            SubQOpenRecordset ResReg, "SELECT * FROM PlanoContas WHERE empr_cd_empresa =" & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa"), Estatico
          Else
            SubQOpenRecordset ResReg, "SELECT * FROM PlanoContas WHERE empr_cd_empresa =" & PCodEmpresa, Estatico
          End If
          If Not ResReg.EOF Then
            'Desabilitar Máscaras e Orelha de Importação
            ChkPlContas.Value = 0
            ChkPlContas.Enabled = False
            TxtMascaraPC.Enabled = False
            LblMascaraPC.Enabled = False
            ExistePlContas = True
          End If
          ResReg.Close
          
          'verifica centro de custo
          If pTipoCopia = "Multi" Then
            SubQOpenRecordset ResReg, "SELECT * FROM CentroResultados WHERE empr_cd_empresa =" & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa"), Estatico
          Else
            SubQOpenRecordset ResReg, "SELECT * FROM CentroResultados WHERE empr_cd_empresa =" & PCodEmpresa, Estatico
          End If
          If Not ResReg.EOF Then
            'Desabilitar Máscaras e Orelha de Importação
            ChkCCustos.Value = 0
            ChkCCustos.Enabled = False
            TxtMascaraCC.Enabled = False
            LblMascaraCC.Enabled = False
            ExisteCCustos = True
          End If
          ResReg.Close
          If ExistePlContas And ExisteCCustos Then
            CboEmpresas.Enabled = False
          End If
          
          'verifica nivel de categoria
          If pTipoCopia = "Multi" Then
              XGT_SQL = "SELECT * FROM ConsGENCategorias WHERE empr_cd_empresa =" & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa")
          Else
              XGT_SQL = "SELECT * FROM ConsGENCategorias WHERE empr_cd_empresa =" & PCodEmpresa
          End If
          SubQOpenRecordset ResReg, XGT_SQL, Estatico
          If Not ResReg.EOF Then
              TxtNivelCatFin.Enabled = False
          End If
          
          ResReg.Close ' Inserido em 13/10/08 - Patrícia
          'Inserido em 31/07/08 - Larissa
          FraVinculaCRCC.Visible = True
          FraVinculaCRCC.Enabled = True
          
          PreencheCampos
          
          'verifica Lancamentos Contábeis
          If pTipoCopia = "Multi" Then
            SubQOpenRecordset ResReg, "SELECT * FROM Lancamentos WHERE empr_cd_empresa =" & TelaEmpresas.Adodc1.Recordset.Fields("empr_cd_empresa"), Estatico
          Else
            SubQOpenRecordset ResReg, "SELECT * FROM Lancamentos WHERE empr_cd_empresa =" & PCodEmpresa, Estatico
          End If
          If Not ResReg.EOF Then
            'Desabilitar características de contabilização
            FraExpContabilidade.Enabled = False
            ChkCaixa.Enabled = False
            ChkComp.Enabled = False
            TxtQtdDiasProvisao.Enabled = False
            LblQtdDiasProvisao.Enabled = False
            ChkApropriarCR.Enabled = False
            ChkExportarQualiPat.Enabled = False
          End If
          ResReg.Close
          
          'Por Larissa - 19/02/2008
          'Desabilita a opção de importar plano de contas e centros de custo:
          FraImportacao.Enabled = False
          ChkPlContas.Enabled = False
          ChkCCustos.Enabled = False
          CboEmpresas.Enabled = False
          LblEmpresa.Enabled = False
          
          subConectarControleDadosNV datpadroes, "select * from empresas where empr_cd_empresa=" & Chave, Dinamico
        Else
          CboCCorrente.Enabled = False
          
          'Habilita esse campo apenas na inserção - Lucas Santiago 23.01.09
          ChkAlteraNotasEstoque.Enabled = True
          PreparaInsercao
        End If
    End If
  
End Sub
Private Sub MskCGC_LostFocus()
  MskCgc.PromptInclude = False
  If MskCgc.Text <> "" Then
    If Not FunConfereCGC(MskCgc) Then
      MsgBox "CNPJ Invalido!", vbInformation + vbOKOnly, "Atenção!"
      MskCgc.PromptInclude = True
      MskCgc.SetFocus
      Exit Sub
    End If
  End If
  MskCgc.PromptInclude = True
  
End Sub
Private Sub MskReceitaDoacao_GotFocus()
    Call subSelecionaMSK(MskReceitaDoacao)
End Sub
Private Sub MskReceitaDoacao_LostFocus()
    If MskReceitaDoacao.Text <> "" Then
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            MskReceitaDoacao.Text = Format(MskReceitaDoacao.Text, XGT_MSKCODREDUZIDO)
        End If
        CboReceitaDoacao.BoundText = MskReceitaDoacao.Text
        If CboReceitaDoacao.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskReceitaDoacao.SetFocus
        End If
    
    End If
End Sub

Private Sub Txtfantasia_LostFocus()
    If FunNulo(TxtNome.Text) = "" Then
        TxtNome.Text = Txtfantasia.Text
    End If
End Sub
Public Sub subImportaCategorias(Chave As Integer)
    Dim XLO_CATEGORIAS As Object
    Dim XLO_CATEMPRESA As Object
    
    XGT_SQL = "SELECT * FROM CategoriaPlanoContas WHERE empr_cd_empresa = " & Chave
    SubQOpenRecordset XLO_CATEMPRESA, XGT_SQL, Estatico
    
    XGT_SQL = "SELECT * FROM CategoriaS"
    SubQOpenRecordset XLO_CATEGORIAS, XGT_SQL, Estatico

    If Not XLO_CATEGORIAS.EOF Then
        
        XLO_CATEGORIAS.MoveFirst
        Do While Not XLO_CATEGORIAS.EOF
            XLO_CATEMPRESA.Filter = "cate_cd_Categoria=" & XLO_CATEGORIAS.Fields("cate_cd_Categoria")
            If XLO_CATEMPRESA.EOF Then
                XGT_SQL = "INSERT CategoriaPlanoContas(empr_cd_Empresa, cate_cd_Categoria) " & _
                          "VALUES (" & Chave & ", " & XLO_CATEGORIAS.Fields("cate_cd_Categoria") & _
                          ")"
                Conexao.Execute XGT_SQL
                XLO_CATEMPRESA.Close
                XLO_CATEMPRESA.Open
                XLO_CATEMPRESA.MoveFirst
            End If
            XLO_CATEGORIAS.MoveNext
        Loop
        
    End If
    XLO_CATEGORIAS.Close ' Inserido em 13/10/08 - Patrícia
    
End Sub
Private Sub subImportaFormularios(Chave As Integer)
    
    Dim XLO_FORMULARIOS As ADODB.Recordset
    Dim Sql As String
    
   'XGT_SQL = "SELECT * FROM Formularios WHERE empr_cd_empresa = " & Chave
    
    'XGT_SQL = "SELECT * INTO #form FROM Formularios WHERE empr_cd_empresa = (SELECT MIN(empr_cd_empresa) FROM Empresas)"
    
   XGT_SQL = "SELECT * FROM Formularios WHERE empr_cd_empresa = (SELECT MIN(empr_cd_empresa) FROM Empresas)"
   SubQOpenRecordset XLO_FORMULARIOS, XGT_SQL, Estatico

    If Not XLO_FORMULARIOS.EOF Then
    
        XLO_FORMULARIOS.MoveFirst
        
        Do While Not XLO_FORMULARIOS.EOF
        
            Sql = "INSERT INTO Formularios VALUES('" & XLO_FORMULARIOS!form_tx_Descricao & "','" & XLO_FORMULARIOS!form_tx_nomecampo & " ','" & XLO_FORMULARIOS!form_nr_imprime & "','" & FunTrataFloat(XLO_FORMULARIOS!form_nr_altura) & "','" & FunTrataFloat(XLO_FORMULARIOS!form_nr_largura) & "','" & XLO_FORMULARIOS!form_nr_tamfonte & "','" & XLO_FORMULARIOS!form_nr_negrito & "','" & XLO_FORMULARIOS!Form_tx_TipoDocumento & "'," & Chave & ")"
            Conexao.Execute (Sql)
            XLO_FORMULARIOS.MoveNext
        Loop
    
    End If
    
    XLO_FORMULARIOS.Close
    
End Sub
Private Sub TxtIM_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtpercentiss_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtpercentiss_LostFocus()
    If IsNumeric(TxtPercentISS.Text) Then
        TxtPercentISS.Text = Format$(TxtPercentISS.Text, "##,##0.00")
    Else
        TxtPercentISS.Text = ""
    End If
End Sub

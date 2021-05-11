VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaManFiadores 
   Caption         =   "Tabela de Fiadores"
   ClientHeight    =   4545
   ClientLeft      =   1965
   ClientTop       =   2850
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   8895
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanFiador 
      Height          =   3480
      Left            =   345
      TabIndex        =   23
      Top             =   300
      Visible         =   0   'False
      Width           =   8310
      _Version        =   65536
      _ExtentX        =   14658
      _ExtentY        =   6138
      _StockProps     =   15
      Caption         =   " Inserção de Fiador"
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
         Left            =   6135
         TabIndex        =   20
         Top             =   3015
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
         Left            =   7155
         TabIndex        =   21
         Top             =   3015
         Width           =   975
      End
      Begin VB.Frame FraDados 
         BackColor       =   &H00E0E0E0&
         Height          =   2640
         Left            =   210
         TabIndex        =   27
         Top             =   300
         Width           =   7935
         Begin VB.TextBox TxtCidade 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5040
            MaxLength       =   20
            TabIndex        =   13
            Top             =   1395
            Width           =   2715
         End
         Begin VB.TextBox TxtOrgEmissor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6630
            MaxLength       =   8
            TabIndex        =   9
            Top             =   225
            Width           =   1125
         End
         Begin VB.TextBox TxtIdentidade 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3630
            MaxLength       =   20
            TabIndex        =   8
            Top             =   225
            Width           =   1530
         End
         Begin VB.TextBox TxtBairro 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1170
            MaxLength       =   20
            TabIndex        =   12
            Top             =   1395
            Width           =   2715
         End
         Begin VB.TextBox TxtEndereco 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1005
            Width           =   6585
         End
         Begin VB.TextBox TxtEMail 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4230
            MaxLength       =   30
            TabIndex        =   16
            Top             =   1785
            Width           =   3525
         End
         Begin VB.TextBox TxtNome 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   10
            Top             =   615
            Width           =   6585
         End
         Begin MSMask.MaskEdBox MskCpf 
            Height          =   315
            Left            =   1155
            TabIndex        =   7
            Top             =   225
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "###.###.###-##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MskCep 
            Height          =   315
            Left            =   2280
            TabIndex        =   15
            Top             =   1785
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##.###-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MskEstado 
            Height          =   315
            Left            =   1170
            TabIndex        =   14
            Top             =   1785
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCelular 
            Height          =   315
            Left            =   6090
            TabIndex        =   19
            Top             =   2175
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MskFax 
            Height          =   315
            Left            =   3510
            TabIndex        =   18
            Top             =   2175
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MskFone 
            Height          =   315
            Left            =   1170
            TabIndex        =   17
            Top             =   2175
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin VB.Label LblOrgEmissor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Org. Emissor:"
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
            Height          =   225
            Left            =   5385
            TabIndex        =   40
            Top             =   270
            Width           =   1200
         End
         Begin VB.Label LblIdentidade 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "RG:"
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
            Height          =   225
            Left            =   3075
            TabIndex        =   39
            Top             =   270
            Width           =   480
         End
         Begin VB.Label LblEMail 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "E-Mail:"
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
            Height          =   225
            Left            =   3495
            TabIndex        =   38
            Top             =   1830
            Width           =   690
         End
         Begin VB.Label LblCep 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "CEP:"
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
            Height          =   225
            Left            =   1770
            TabIndex        =   37
            Top             =   1830
            Width           =   465
         End
         Begin VB.Label LblCelular 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Celular:"
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
            Height          =   225
            Left            =   5235
            TabIndex        =   36
            Top             =   2205
            Width           =   825
         End
         Begin VB.Label LblEstado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Estado:"
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
            Height          =   225
            Left            =   480
            TabIndex        =   35
            Top             =   1830
            Width           =   660
         End
         Begin VB.Label LblCidade 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cidade:"
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
            Height          =   225
            Left            =   4275
            TabIndex        =   34
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label LblFax 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fax:"
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
            Height          =   225
            Left            =   3000
            TabIndex        =   33
            Top             =   2235
            Width           =   465
         End
         Begin VB.Label LblNome 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nome:"
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
            Height          =   210
            Left            =   405
            TabIndex        =   32
            Top             =   660
            Width           =   720
         End
         Begin VB.Label LblEndereco 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Endereço:"
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
            Height          =   195
            Left            =   255
            TabIndex        =   31
            Top             =   1035
            Width           =   885
         End
         Begin VB.Label LblBairro 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Bairro:"
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
            Height          =   225
            Left            =   435
            TabIndex        =   30
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label LblFone 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fone:"
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
            Height          =   225
            Left            =   525
            TabIndex        =   29
            Top             =   2205
            Width           =   600
         End
         Begin VB.Label LblCpf 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "CPF:"
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
            Left            =   360
            TabIndex        =   28
            Top             =   270
            Width           =   750
         End
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4590
      Top             =   3285
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Bindings        =   "TelaFiadores.frx":0000
      Height          =   3690
      Left            =   45
      TabIndex        =   24
      Top             =   15
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   6509
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nome"
      Columns(0).DataField=   "fiad_tx_nome"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "RG"
      Columns(1).DataField=   "fiad_nr_identidade"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Órgão Emissor"
      Columns(2).DataField=   "fiad_tx_orgemissor"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "CPF"
      Columns(3).DataField=   "fiad_nr_cpf"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Endereço"
      Columns(4).DataField=   "fiad_tx_endereco"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Bairro"
      Columns(5).DataField=   "fiad_tx_bairro"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Cidade"
      Columns(6).DataField=   "fiad_tx_cidade"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Estado"
      Columns(7).DataField=   "fiad_tx_estado"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Telefone"
      Columns(8).DataField=   "fiad_nr_fone"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Fax"
      Columns(9).DataField=   "fiad_nr_fax"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Celular"
      Columns(10).DataField=   "fiad_nr_celular"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Email"
      Columns(11).DataField=   "fiad_tx_email"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=5080"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5001"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=2196"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2117"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=528"
      Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(16)=   "Column(1)._MinWidth=4"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=1693"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=1614"
      Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=528"
      Splits(0)._ColumnProps(22)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(24)=   "Column(2)._MinWidth=89391764"
      Splits(0)._ColumnProps(25)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(3)._ColStyle=528"
      Splits(0)._ColumnProps(30)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(31)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(32)=   "Column(4).Width=5239"
      Splits(0)._ColumnProps(33)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(4)._WidthInPix=5159"
      Splits(0)._ColumnProps(35)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(4)._ColStyle=528"
      Splits(0)._ColumnProps(37)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(38)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(39)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(40)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(42)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(43)=   "Column(5)._ColStyle=528"
      Splits(0)._ColumnProps(44)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(45)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(46)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(47)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(49)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(50)=   "Column(6)._ColStyle=528"
      Splits(0)._ColumnProps(51)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(52)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(53)=   "Column(7).Width=1217"
      Splits(0)._ColumnProps(54)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(7)._WidthInPix=1138"
      Splits(0)._ColumnProps(56)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(57)=   "Column(7)._ColStyle=528"
      Splits(0)._ColumnProps(58)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(59)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(60)=   "Column(8).Width=2461"
      Splits(0)._ColumnProps(61)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(8)._WidthInPix=2381"
      Splits(0)._ColumnProps(63)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(64)=   "Column(8)._ColStyle=528"
      Splits(0)._ColumnProps(65)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(66)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(67)=   "Column(9).Width=2566"
      Splits(0)._ColumnProps(68)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(9)._WidthInPix=2487"
      Splits(0)._ColumnProps(70)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(71)=   "Column(9)._ColStyle=528"
      Splits(0)._ColumnProps(72)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(73)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(74)=   "Column(10).Width=2461"
      Splits(0)._ColumnProps(75)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(10)._WidthInPix=2381"
      Splits(0)._ColumnProps(77)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(78)=   "Column(10)._ColStyle=528"
      Splits(0)._ColumnProps(79)=   "Column(10).WrapText=1"
      Splits(0)._ColumnProps(80)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(81)=   "Column(11).Width=4524"
      Splits(0)._ColumnProps(82)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(11)._WidthInPix=4445"
      Splits(0)._ColumnProps(84)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(85)=   "Column(11)._ColStyle=528"
      Splits(0)._ColumnProps(86)=   "Column(11).WrapText=1"
      Splits(0)._ColumnProps(87)=   "Column(11).Order=12"
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
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(21)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(22)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7,.wraptext=-1"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.locked=0"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15,.alignment=3"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17,.alignment=3"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=0"
      _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=0"
      _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=0"
      _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=0"
      _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=0"
      _StyleDefs(65)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(66)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=0"
      _StyleDefs(69)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14,.alignment=2"
      _StyleDefs(70)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=0"
      _StyleDefs(73)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14,.alignment=2"
      _StyleDefs(74)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(11).Style:id=82,.parent=13,.alignment=0"
      _StyleDefs(77)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14,.alignment=2"
      _StyleDefs(78)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(80)  =   "Named:id=33:Normal"
      _StyleDefs(81)  =   ":id=33,.parent=0"
      _StyleDefs(82)  =   "Named:id=34:Heading"
      _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(84)  =   ":id=34,.wraptext=-1"
      _StyleDefs(85)  =   "Named:id=35:Footing"
      _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   "Named:id=36:Selected"
      _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=37:Caption"
      _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(91)  =   "Named:id=38:HighlightRow"
      _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(93)  =   "Named:id=39:EvenRow"
      _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(95)  =   "Named:id=40:OddRow"
      _StyleDefs(96)  =   ":id=40,.parent=33"
      _StyleDefs(97)  =   "Named:id=41:RecordSelector"
      _StyleDefs(98)  =   ":id=41,.parent=34"
      _StyleDefs(99)  =   "Named:id=42:FilterBar"
      _StyleDefs(100) =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   360
      Left            =   120
      TabIndex        =   22
      Top             =   4125
      Width           =   8760
      _Version        =   65536
      _ExtentX        =   15452
      _ExtentY        =   635
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
         Left            =   315
         TabIndex        =   0
         Top             =   0
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
         Left            =   7560
         TabIndex        =   6
         Top             =   0
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
         Left            =   6360
         TabIndex        =   5
         Top             =   0
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
         Left            =   5160
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
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
         Left            =   3960
         TabIndex        =   3
         Top             =   0
         Width           =   975
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
         Left            =   2760
         TabIndex        =   2
         Top             =   0
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
         Left            =   1560
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaFiadores.frx":0015
      Height          =   315
      Left            =   2760
      TabIndex        =   25
      Top             =   3750
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "fiad_tx_nome"
      BoundColumn     =   "fiad_cd_fiador"
      Text            =   ""
   End
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Nome:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   3810
      Width           =   2670
   End
End
Attribute VB_Name = "TelaManFiadores"
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
Dim FRM_SQLAtivo As String        'Atualizar o SQL Ativo após modificações de filtro e ordem

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

Private Sub cboLocalizar_Click(Area As Integer)
    If Area = 2 Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If

End Sub

Private Sub cboLocalizar_KeyUp(KeyCode As Integer, Shift As Integer)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If

End Sub


Private Sub cboLocalizar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
End Sub


Private Sub CmdConfirmar_Click()
    Dim XLT_SQL As String
    
    
    'Função para verificar preenchimento dos campos obrigatórios
    MskCpf.PromptInclude = False
    If TxtNome.Text = "" Or MskCpf.Text = "" Then
        MsgBox "O preenchimento dos campos Nome e CPF é obrigatório.", vbCritical + vbOKOnly, "ATENÇÃO"
        Exit Sub
    End If
    MskCpf.PromptInclude = True
    
    Conexao.BeginTrans
    If PanFiador.Tag = "I" Then
        XLT_SQL = "INSERT INTO Fiadores (fiad_tx_nome, fiad_nr_cpf,fiad_nr_identidade, fiad_tx_orgemissor, " & _
                                    "fiad_tx_endereco, fiad_tx_bairro, fiad_tx_cidade, fiad_tx_estado, fiad_nr_cep, " & _
                                    "fiad_tx_email, fiad_nr_fone, fiad_nr_fax, fiad_nr_celular) " & _
                                    "VALUES ('" & TxtNome.Text & "','" & MskCpf.Text & "','" & TxtIdentidade.Text & "','" & TxtOrgEmissor.Text & "'," & _
                                    "'" & TxtEndereco.Text & "','" & TxtBairro.Text & "','" & TxtCidade.Text & "','" & MskEstado.Text & "','" & MskCep.Text & "'," & _
                                    "'" & TxtEMail.Text & "','" & MskFone.Text & "','" & MskFax.Text & "','" & MskCelular.Text & "')"
        Conexao.Execute XLT_SQL
    Else    'Alteração
        XLT_SQL = "UPDATE Fiadores SET fiad_tx_nome='" & TxtNome.Text & "'," & _
                                    "fiad_nr_cpf='" & MskCpf.Text & "'," & _
                                    "fiad_nr_identidade='" & TxtIdentidade.Text & "'," & _
                                    "fiad_tx_orgemissor='" & TxtOrgEmissor.Text & "'," & _
                                    "fiad_tx_endereco='" & TxtEndereco.Text & "'," & _
                                    "fiad_tx_bairro='" & TxtBairro.Text & "'," & _
                                    "fiad_tx_cidade='" & TxtCidade.Text & "'," & _
                                    "fiad_tx_estado='" & MskEstado.Text & "'," & _
                                    "fiad_nr_cep='" & MskCep.Text & "'," & _
                                    "fiad_tx_email='" & TxtEMail.Text & "'," & _
                                    "fiad_nr_fone='" & MskFone.Text & "'," & _
                                    "fiad_nr_fax='" & MskFax.Text & "'," & _
                                    "fiad_nr_celular='" & MskCelular.Text & "'" & _
                        " WHERE fiad_cd_fiador = " & PanFiador.Tag
        Conexao.Execute XLT_SQL
    End If
    Conexao.CommitTrans
    
    PanFiador.Visible = False
    PanBotoes.Enabled = True
    subHabilitaBotoes
    
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    If PanFiador.Tag <> "I" Then
        cboLocalizar.BoundText = PanFiador.Tag
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
    
    cboLocalizar.Enabled = True
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus

End Sub

Private Sub CmdImprimir_Click()
    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Fiadores")
End Sub

Private Sub CmdRemoverFiltro_Click()
    Dim XLI_FIL As Integer
    
    subDesabilitaBotoes

    Set Formulario = TelaManFiadores
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

Private Sub CmdRemoverFiltro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Mostra os filtros ativos da tabela"
End Sub

Private Sub CmdAlterar_Click()
    PanBotoes.Enabled = False
    PanFiador.Visible = True
    PanFiador.Enabled = True
    subDesabilitaBotoes
    subPreencheCampos
    SubDesabHabCampos (True)
    MskCpf.Enabled = False
    PanFiador.Left = (TDBGrid1.Width - PanFiador.Width) / 2
    PanFiador.Top = (TDBGrid1.Height - PanFiador.Height) / 2
    PanFiador.Tag = Adodc1.Recordset.Fields("fiad_cd_fiador")
    PanFiador.Caption = " Alteração de Fiador"
    cboLocalizar.Enabled = False
    TDBGrid1.Enabled = False
    TxtNome.SetFocus
End Sub
Sub subPreencheCampos()
    TxtNome.Text = Adodc1.Recordset.Fields("fiad_tx_nome")
    MskCpf.Text = Adodc1.Recordset.Fields("fiad_nr_cpf")
    TxtIdentidade.Text = Adodc1.Recordset.Fields("fiad_nr_identidade")
    TxtOrgEmissor.Text = FunNulo(Adodc1.Recordset.Fields("fiad_tx_orgemissor"))
    TxtEndereco.Text = FunNulo(Adodc1.Recordset.Fields("fiad_tx_endereco"))
    TxtBairro.Text = FunNulo(Adodc1.Recordset.Fields("fiad_tx_bairro"))
    TxtCidade.Text = FunNulo(Adodc1.Recordset.Fields("fiad_tx_cidade"))
    MskEstado.Text = FunNulo(Adodc1.Recordset.Fields("fiad_tx_estado"))
    MskCep.Text = FunNulo(Adodc1.Recordset.Fields("fiad_nr_cep"))
    TxtEMail.Text = FunNulo(Adodc1.Recordset.Fields("fiad_tx_email"))
    MskFone.Text = FunNulo(Adodc1.Recordset.Fields("fiad_nr_fone"))
    MskFax.Text = FunNulo(Adodc1.Recordset.Fields("fiad_nr_fax"))
    MskCelular.Text = FunNulo(Adodc1.Recordset.Fields("fiad_nr_celular"))
End Sub

Sub SubDesabHabCampos(XHab As Boolean)
    TxtNome.Enabled = XHab
    TxtIdentidade.Enabled = XHab
    TxtOrgEmissor.Enabled = XHab
    TxtEndereco.Enabled = XHab
    TxtBairro.Enabled = XHab
    TxtCidade.Enabled = XHab
    MskEstado.Enabled = XHab
    MskCep.Enabled = XHab
    TxtEMail.Enabled = XHab
    MskFone.Enabled = XHab
    MskFax.Enabled = XHab
    MskCelular.Enabled = XHab
    
    LblNome.Enabled = XHab
    LblIdentidade.Enabled = XHab
    LblOrgEmissor.Enabled = XHab
    LblEndereco.Enabled = XHab
    LblBairro.Enabled = XHab
    LblCidade.Enabled = XHab
    LblEstado.Enabled = XHab
    LblCep.Enabled = XHab
    LblEMail.Enabled = XHab
    LblFone.Enabled = XHab
    LblFax.Enabled = XHab
    LblCelular.Enabled = XHab
End Sub

Private Sub CmdRetornar_Click()
    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.CancelUpdate
    PanBotoes.Enabled = True
    PanFiador.Visible = False
    PanFiador.Enabled = False
    subHabilitaBotoes
    cboLocalizar.Enabled = True
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub

Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Imprime os dados da tabela"
End Sub

Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    Unload Me
End Sub

Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub
Private Sub CmdOrdem_Click()
    
    subDesabilitaBotoes
    Set Formulario = TelaManFiadores
    subTelaValoresGlobais "P"
    TelaOrdem.Show 1
    subTelaValoresGlobais "G"
    subHabilitaBotoes
    
    lblLocalizar.Caption = "Localizar " & xLocalDescricao & ": "
    cboLocalizar.ListField = xLocalCampo
    cboLocalizar.Text = ""
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus

End Sub

Private Sub CmdOrdem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Ordena os registros atuais da tabela"
End Sub

Private Sub CmdExcluir_Click()
    Dim XLI_OPCAO As Integer
    Dim XLO_FIADOR As Object

    Dim XLI_HOUVEERRO As Integer
    
    XLI_HOUVEERRO = False
    
    subDesabilitaBotoes
    
    XLI_OPCAO = MsgBox("Confirma Remoção do Registro ?", 273, "CUIDADO")
    
    If XLI_OPCAO = 1 Then  'Confirma exclusão
       On Error GoTo RotuloErro

       Conexao.BeginTrans
       Conexao.Execute "DELETE FROM fiadores WHERE fiad_cd_fiador=" & Adodc1.Recordset.Fields("fiad_cd_fiador")
       Conexao.CommitTrans
      
       FRM_Quantid = funOperacaoQuantid("-")
       subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
       If ContFil <> 0 Then
          subContaReg Adodc1
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
       
    cboLocalizar.Text = ""
    TDBGrid1.SetFocus
    TDBGrid1.Col = 0
    
    Exit Sub
    
RotuloErro:
    Conexao.RollbackTrans
    If Err.Number = -2147467259 Then
      MsgBox "Esse fiador não pode ser excluído pois já possui algum contrato relacionado!", vbCritical, "ATENÇÃO"
    End If
    subHabilitaBotoes
End Sub
Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
End Sub

Private Sub CmdInserir_Click()
    PanBotoes.Enabled = False
    PanFiador.Visible = True
    PanFiador.Enabled = True
    subPrepararInsercao
    SubDesabHabCampos (False)
    PanFiador.Left = (TDBGrid1.Width - PanFiador.Width) / 2
    PanFiador.Top = (TDBGrid1.Height - PanFiador.Height) / 2
    PanFiador.Caption = " Inserção de Fiador"
    PanFiador.Tag = "I"
    subDesabilitaBotoes
    cboLocalizar.Text = ""
    cboLocalizar.Enabled = False
    TDBGrid1.Enabled = False
    LblCpf.Enabled = True
    MskCpf.Enabled = True
    MskCpf.SetFocus
End Sub
Sub subPrepararInsercao()
    
    TxtNome.Text = ""
    TxtIdentidade.Text = ""
    TxtOrgEmissor.Text = ""
    TxtEndereco.Text = ""
    TxtBairro.Text = ""
    TxtCidade.Text = ""
    TxtEMail.Text = ""
    MskCpf.Text = "   .   .   -  "
    MskEstado.Text = ""
    MskCep.Text = "  .   -   "
    MskFone.Mask = ""
    MskFone.Text = ""
    MskFax.Text = ""
    MskCelular.Text = ""
    
End Sub

Private Sub CmdInserir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Insere um novo registro na tabela"
End Sub
  
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

Private Sub MskCpf_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 13
    End If
End Sub


Private Sub MskCpf_LostFocus()
    Dim XLO_CPF As Object
    Dim XLT_SQL As String
    
    If PanFiador.Tag = "I" Then
        MskCpf.PromptInclude = False
        If MskCpf.Text <> "" Then
            'verifica CPF e CGC
            If Not FunConfereCPF(MskCpf) Then
                MsgBox "CPF Invalido!", vbInformation + vbOKOnly, "Atenção!"
                MskCpf.PromptInclude = True
                MskCpf.SetFocus
                MskCpf.SelStart = 0
                MskCpf.SelLength = 14
                Exit Sub
            End If
        Else
            CmdRetornar.SetFocus
            Exit Sub
        End If
        MskCpf.PromptInclude = True
        
        SubQOpenRecordset XLO_CPF, "SELECT * FROM Fiadores WHERE fiad_nr_cpf = '" & MskCpf.Text & "'", "Estatico"
    
        If Not (XLO_CPF.BOF And XLO_CPF.EOF) Then
            XLO_CPF.MoveFirst
            PanFiador.Tag = XLO_CPF!fiad_cd_fiador
            XLO_CPF.Close
            'alterar para setar o combo da Tela de chamada
            TelaManFiadores.cboLocalizar.BoundText = PanFiador.Tag
            TelaManFiadores.Adodc1.Recordset.bookmark = TelaManFiadores.cboLocalizar.SelectedItem
            subPreencheCampos
            MskCpf.Enabled = False
            LblCpf.Enabled = False
        End If
        SubDesabHabCampos (True)
        TxtIdentidade.SetFocus
    End If
End Sub


Private Sub TDBGrid1_DblClick()
    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    cboLocalizar.Text = ""
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaManFiadores
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


Private Sub Form_Activate()
    Set Formulario = TelaManFiadores
    If Not PrimeiraVez Then
        subTelaValoresGlobais "P"
    End If
    
    PrimeiraVez = False

End Sub

Private Sub Form_Load()
    
    Set Formulario = TelaManFiadores
    
    Call subAjustaTelaLoc
    
    subManutencaoJanelasAtivas "I", "TelaManFiadores"
         
    'Associa um banco de dados à tabela profissões
    subCarregaDadosNV Adodc1, "Fiadores", "fiad_tx_nome", _
                            "", "SELECT * FROM Fiadores"
                               
    subHabilitaBotoes
    
    filtra_especial = False
    
    subTelaValoresGlobais "G"
    PrimeiraVez = True
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaManFiadores"
End Sub









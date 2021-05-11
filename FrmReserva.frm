VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmManReserva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Reservas"
   ClientHeight    =   5355
   ClientLeft      =   30
   ClientTop       =   1335
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleMode       =   0  'User
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraPrincipal 
      Height          =   4770
      Left            =   0
      TabIndex        =   27
      Top             =   -60
      Width           =   9345
      Begin VB.Frame FraDadosGerais 
         Height          =   675
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   9135
         Begin VB.Label LblGaragem2 
            Caption         =   "Garagem"
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
            Left            =   7260
            TabIndex        =   43
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label LblGaragem 
            Alignment       =   1  'Right Justify
            Caption         =   "Garagem:"
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
            Left            =   6360
            TabIndex        =   42
            Top             =   240
            Width           =   885
         End
         Begin VB.Label LblAndar2 
            Caption         =   "Andar"
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
            Left            =   5820
            TabIndex        =   41
            Top             =   240
            Width           =   375
         End
         Begin VB.Label LblArea2 
            Caption         =   "Area"
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
            Left            =   3960
            TabIndex        =   40
            Top             =   240
            Width           =   435
         End
         Begin VB.Label LblTipoImovel2 
            Caption         =   "Tipo Imovel"
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
            Left            =   1500
            TabIndex        =   39
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Lblm2 
            Caption         =   "m²"
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
            Left            =   4500
            TabIndex        =   38
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LblArea 
            Alignment       =   1  'Right Justify
            Caption         =   "Área:"
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
            Left            =   3420
            TabIndex        =   37
            Top             =   240
            Width           =   495
         End
         Begin VB.Label LblTipoImovel 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo do Imóvel:"
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
            Left            =   60
            TabIndex        =   36
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label LblAndar 
            Alignment       =   1  'Right Justify
            Caption         =   "Andar:"
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
            Left            =   5100
            TabIndex        =   35
            Top             =   240
            Width           =   690
         End
      End
      Begin Threed.SSPanel PanCodigo 
         Height          =   870
         Left            =   120
         TabIndex        =   28
         Top             =   180
         Width           =   9135
         _Version        =   65536
         _ExtentX        =   16113
         _ExtentY        =   1535
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
         Begin VB.TextBox TxtImovel 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8040
            TabIndex        =   2
            Top             =   60
            Width           =   855
         End
         Begin VB.TextBox TxtEmpreendimento 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2580
            TabIndex        =   1
            Top             =   60
            Width           =   3615
         End
         Begin MSMask.MaskEdBox MskEmpreendimento 
            Height          =   315
            Left            =   1800
            TabIndex        =   0
            Top             =   60
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSAdodcLib.Adodc DatEmpreendimento 
            Height          =   330
            Left            =   6660
            Top             =   480
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
         Begin MSMask.MaskEdBox MskContrato 
            Height          =   315
            Left            =   1260
            TabIndex        =   3
            Top             =   480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin VB.Label LblStatus2 
            Caption         =   "Reservado"
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
            Left            =   2700
            TabIndex        =   44
            Top             =   540
            Width           =   1035
         End
         Begin VB.Label LblStatus 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1860
            TabIndex        =   33
            Top             =   540
            Width           =   750
         End
         Begin VB.Label LblContrato 
            Caption         =   "Contrato:"
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
            Left            =   300
            TabIndex        =   32
            Top             =   540
            Width           =   855
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
            Left            =   255
            TabIndex        =   30
            Top             =   120
            Width           =   1485
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
            Left            =   6405
            TabIndex        =   29
            Top             =   120
            Width           =   1605
         End
      End
      Begin TabDlg.SSTab TabAuxiliar 
         Height          =   2805
         Left            =   60
         TabIndex        =   31
         Top             =   1920
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   4948
         _Version        =   393216
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "&Compradores"
         TabPicture(0)   =   "FrmReserva.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LblCliente"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "LblCliTitular"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "CboComprador"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "TDBGrid1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "DatComprador"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "CmdTitular"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "CmdInsComprador"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "CmdExcComprador"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "&Dados Financeiros"
         TabPicture(1)   =   "FrmReserva.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FraValores"
         Tab(1).Control(1)=   "FraDatas"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "C&orretores"
         TabPicture(2)   =   "FrmReserva.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "CmdInsCorretor"
         Tab(2).Control(1)=   "CmdExcCorretor"
         Tab(2).Control(2)=   "CmdAltCorretor"
         Tab(2).Control(3)=   "PanCorretor"
         Tab(2).Control(4)=   "TDBGrid2"
         Tab(2).ControlCount=   5
         Begin VB.CommandButton CmdInsCorretor 
            Caption         =   "&Inserir"
            Height          =   255
            Left            =   -66960
            TabIndex        =   16
            Top             =   540
            Width           =   975
         End
         Begin VB.CommandButton CmdExcCorretor 
            Caption         =   "&Excluir"
            Height          =   255
            Left            =   -66960
            TabIndex        =   18
            Top             =   1260
            Width           =   975
         End
         Begin VB.CommandButton CmdAltCorretor 
            Caption         =   "&Alterar"
            Height          =   255
            Left            =   -66960
            TabIndex        =   17
            Top             =   900
            Width           =   975
         End
         Begin VB.CommandButton CmdExcComprador 
            Caption         =   "&Excluir"
            Height          =   255
            Left            =   8040
            TabIndex        =   6
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton CmdInsComprador 
            Caption         =   "&Inserir"
            Height          =   255
            Left            =   8040
            TabIndex        =   5
            Top             =   465
            Width           =   975
         End
         Begin VB.CommandButton CmdTitular 
            Caption         =   "&Titular"
            Height          =   255
            Left            =   8040
            TabIndex        =   7
            Top             =   1215
            Width           =   975
         End
         Begin VB.Frame FraDatas 
            Height          =   615
            Left            =   -74940
            TabIndex        =   54
            Top             =   1740
            Width           =   9045
            Begin MSComCtl2.DTPicker DtpDtLimiteReserva 
               Height          =   315
               Left            =   6420
               TabIndex        =   15
               Top             =   180
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MM/yy"
               Format          =   63832067
               CurrentDate     =   37180
            End
            Begin MSComCtl2.DTPicker DtpDtReserva 
               Height          =   315
               Left            =   1980
               TabIndex        =   14
               Top             =   180
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd/MM/yy"
               Format          =   63832067
               CurrentDate     =   37180
            End
            Begin VB.Label LblDtreserva 
               Caption         =   "Reserva:"
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
               Left            =   1200
               TabIndex        =   56
               Top             =   240
               Width           =   735
            End
            Begin VB.Label LblDtLimiteReserva 
               Caption         =   "Limite da Reserva:"
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
               Left            =   4680
               TabIndex        =   55
               Top             =   240
               Width           =   1635
            End
         End
         Begin VB.Frame FraValores 
            Height          =   930
            Left            =   -74940
            TabIndex        =   45
            Top             =   600
            Width           =   9045
            Begin VB.TextBox TxtOutros 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6000
               TabIndex        =   13
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox TxtSinal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   840
               MaxLength       =   20
               TabIndex        =   9
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox TxtPoupanca 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3720
               MaxLength       =   20
               TabIndex        =   10
               Top             =   180
               Width           =   1215
            End
            Begin VB.TextBox TxtTotVenda 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   7680
               MaxLength       =   20
               TabIndex        =   46
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox TxtFgts 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6000
               MaxLength       =   20
               TabIndex        =   12
               Top             =   180
               Width           =   1215
            End
            Begin VB.TextBox TxtFinanciamento 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3720
               MaxLength       =   20
               TabIndex        =   11
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox TxtAVista 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   840
               MaxLength       =   20
               TabIndex        =   8
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label LblOutros 
               Caption         =   "Outros:"
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
               Left            =   5340
               TabIndex        =   53
               Top             =   600
               Width           =   675
            End
            Begin VB.Label LblSinal 
               Alignment       =   1  'Right Justify
               Caption         =   "Sinal:"
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
               Left            =   240
               TabIndex        =   52
               Top             =   600
               Width           =   585
            End
            Begin VB.Label LblPoupanca 
               Alignment       =   1  'Right Justify
               Caption         =   "Poupança:"
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
               Left            =   2700
               TabIndex        =   51
               Top             =   240
               Width           =   960
            End
            Begin VB.Label LblTotVenda 
               Alignment       =   1  'Right Justify
               Caption         =   "Total a Prazo"
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
               Left            =   7680
               TabIndex        =   50
               Top             =   300
               Width           =   1185
            End
            Begin VB.Label LblFgts 
               Alignment       =   1  'Right Justify
               Caption         =   "FGTS:"
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
               Left            =   5280
               TabIndex        =   49
               Top             =   240
               Width           =   705
            End
            Begin VB.Label LblFinanciamento 
               Alignment       =   1  'Right Justify
               Caption         =   "Financiamento:"
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
               Left            =   2400
               TabIndex        =   48
               Top             =   600
               Width           =   1305
            End
            Begin VB.Label LblAVista 
               Alignment       =   1  'Right Justify
               Caption         =   "À Vista:"
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
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   705
            End
         End
         Begin MSAdodcLib.Adodc DatComprador 
            Height          =   330
            Left            =   3600
            Top             =   420
            Visible         =   0   'False
            Width           =   2100
            _ExtentX        =   3704
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
            Caption         =   "DatComprador"
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
         Begin TrueDBGrid70.TDBGrid TDBGrid1 
            Height          =   1650
            Left            =   240
            TabIndex        =   57
            Top             =   945
            Width           =   7650
            _ExtentX        =   13494
            _ExtentY        =   2910
            _LayoutType     =   4
            _RowHeight      =   12
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nome"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Titular"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=65808"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(9)=   "Column(0).FetchStyle=1"
            Splits(0)._ColumnProps(10)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(11)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(12)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(13)=   "Column(1).Width=7858"
            Splits(0)._ColumnProps(14)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(1)._WidthInPix=7779"
            Splits(0)._ColumnProps(16)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(1)._ColStyle=74256"
            Splits(0)._ColumnProps(18)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(19)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(20)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(21)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(22)=   "Column(1)._MinWidth=80"
            Splits(0)._ColumnProps(23)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(24)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(26)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(2).AllowSizing=0"
            Splits(0)._ColumnProps(28)=   "Column(2)._ColStyle=65808"
            Splits(0)._ColumnProps(29)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(30)=   "Column(2).WrapText=1"
            Splits(0)._ColumnProps(31)=   "Column(2).FetchStyle=1"
            Splits(0)._ColumnProps(32)=   "Column(2).AllowFocus=0"
            Splits(0)._ColumnProps(33)=   "Column(2).Order=3"
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
            HeadLines       =   1
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H8000000E&,.fgcolor=&H0&"
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
            _StyleDefs(18)  =   "Splits(0).Style:id=21,.parent=1,.valignment=2,.wraptext=-1"
            _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=80,.parent=4"
            _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=22,.parent=2"
            _StyleDefs(21)  =   "Splits(0).FooterStyle:id=23,.parent=3"
            _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=24,.parent=5"
            _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=76,.parent=6"
            _StyleDefs(24)  =   "Splits(0).EditorStyle:id=75,.parent=7"
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=21"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(35)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
            _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=16,.parent=21"
            _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=22"
            _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=23"
            _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=75"
            _StyleDefs(43)  =   "Named:id=33:Normal"
            _StyleDefs(44)  =   ":id=33,.parent=0"
            _StyleDefs(45)  =   "Named:id=34:Heading"
            _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   ":id=34,.wraptext=-1"
            _StyleDefs(48)  =   "Named:id=35:Footing"
            _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   "Named:id=36:Selected"
            _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(52)  =   "Named:id=37:Caption"
            _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(54)  =   "Named:id=38:HighlightRow"
            _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
            _StyleDefs(56)  =   "Named:id=39:EvenRow"
            _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(58)  =   "Named:id=40:OddRow"
            _StyleDefs(59)  =   ":id=40,.parent=33"
            _StyleDefs(60)  =   "Named:id=41:RecordSelector"
            _StyleDefs(61)  =   ":id=41,.parent=34"
            _StyleDefs(62)  =   "Named:id=42:FilterBar"
            _StyleDefs(63)  =   ":id=42,.parent=33"
         End
         Begin MSDataListLib.DataCombo CboComprador 
            Bindings        =   "FrmReserva.frx":0054
            Height          =   315
            Left            =   2220
            TabIndex        =   4
            Top             =   525
            Width           =   5625
            _ExtentX        =   9922
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "focl_tx_RazaoSocial"
            BoundColumn     =   "focl_cd_forncli"
            Text            =   "CboComprador"
         End
         Begin Threed.SSPanel PanCorretor 
            Height          =   1800
            Left            =   -73200
            TabIndex        =   60
            Top             =   720
            Visible         =   0   'False
            Width           =   5400
            _Version        =   65536
            _ExtentX        =   9525
            _ExtentY        =   3175
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
            Enabled         =   0   'False
            Begin VB.TextBox TxtPercComissao 
               Height          =   315
               Left            =   1500
               TabIndex        =   21
               Top             =   1020
               Width           =   855
            End
            Begin VB.TextBox TxtComissao 
               Height          =   315
               Left            =   4260
               MaxLength       =   20
               TabIndex        =   22
               Top             =   1020
               Width           =   930
            End
            Begin VB.CommandButton CmdConfirmaCorretor 
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
               Height          =   285
               Left            =   3195
               TabIndex        =   23
               Top             =   1440
               Width           =   975
            End
            Begin VB.CommandButton CmdSairCorretor 
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
               Height          =   285
               Left            =   4200
               TabIndex        =   24
               Top             =   1440
               Width           =   975
            End
            Begin MSDataListLib.DataCombo CboCorretor 
               Bindings        =   "FrmReserva.frx":006F
               Height          =   315
               Left            =   825
               TabIndex        =   19
               Top             =   165
               Width           =   4410
               _ExtentX        =   7779
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "corr_tx_nome"
               BoundColumn     =   "corr_cd_corretor"
               Text            =   "CboCorretor"
            End
            Begin MSComCtl2.DTPicker DtpPagamento 
               Height          =   315
               Left            =   3660
               TabIndex        =   20
               Top             =   600
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "dd/MM/yy"
               DateIsNull      =   -1  'True
               Format          =   63832067
               CurrentDate     =   37818.8262615741
            End
            Begin MSAdodcLib.Adodc DatCorretor 
               Height          =   330
               Left            =   2295
               Top             =   195
               Visible         =   0   'False
               Width           =   2130
               _ExtentX        =   3757
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
               Caption         =   "DatCorretor"
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
            Begin VB.Label LblValorDoImovel 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Valor"
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
               Left            =   1380
               TabIndex        =   66
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label LblPercComissao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Comissão (%): "
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
               Left            =   240
               TabIndex        =   65
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label LblValorImovel 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Valor Imóvel"
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
               TabIndex        =   64
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label LblNomeCorretor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Nome:"
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
               Left            =   180
               TabIndex        =   63
               Top             =   225
               Width           =   600
            End
            Begin VB.Label LblComissao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Comissão:"
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
               Left            =   3300
               TabIndex        =   62
               Top             =   1080
               Width           =   825
            End
            Begin VB.Label LblPagamento 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Pagamento:"
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
               Left            =   2640
               TabIndex        =   61
               Top             =   630
               Width           =   1035
            End
         End
         Begin TrueDBGrid70.TDBGrid TDBGrid2 
            Height          =   1935
            Left            =   -74760
            TabIndex        =   67
            Top             =   615
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   3413
            _LayoutType     =   4
            _RowHeight      =   12
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nome"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Comissão"
            Columns(2).DataField=   ""
            Columns(2).NumberFormat=   "Standard"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Pagamento"
            Columns(3).DataField=   ""
            Columns(3).NumberFormat=   "dd/mm/yy"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=65808"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(9)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(10)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(11)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(12)=   "Column(1).Width=7858"
            Splits(0)._ColumnProps(13)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._WidthInPix=7779"
            Splits(0)._ColumnProps(15)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(1)._ColStyle=74256"
            Splits(0)._ColumnProps(17)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(18)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(19)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(20)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(21)=   "Column(1)._MinWidth=80"
            Splits(0)._ColumnProps(22)=   "Column(2).Width=2143"
            Splits(0)._ColumnProps(23)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(2)._WidthInPix=2064"
            Splits(0)._ColumnProps(25)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(26)=   "Column(2)._ColStyle=74256"
            Splits(0)._ColumnProps(27)=   "Column(2).WrapText=1"
            Splits(0)._ColumnProps(28)=   "Column(2).FetchStyle=1"
            Splits(0)._ColumnProps(29)=   "Column(2).AllowFocus=0"
            Splits(0)._ColumnProps(30)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(31)=   "Column(2)._MinWidth=89769656"
            Splits(0)._ColumnProps(32)=   "Column(3).Width=2752"
            Splits(0)._ColumnProps(33)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(3)._WidthInPix=2672"
            Splits(0)._ColumnProps(35)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(3)._ColStyle=74256"
            Splits(0)._ColumnProps(37)=   "Column(3).WrapText=1"
            Splits(0)._ColumnProps(38)=   "Column(3).FetchStyle=1"
            Splits(0)._ColumnProps(39)=   "Column(3).AllowFocus=0"
            Splits(0)._ColumnProps(40)=   "Column(3).Order=4"
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
            HeadLines       =   1
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H8000000E&,.fgcolor=&H0&"
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
            _StyleDefs(18)  =   "Splits(0).Style:id=21,.parent=1,.valignment=2,.wraptext=-1"
            _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=80,.parent=4"
            _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=22,.parent=2"
            _StyleDefs(21)  =   "Splits(0).FooterStyle:id=23,.parent=3"
            _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=24,.parent=5"
            _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=76,.parent=6"
            _StyleDefs(24)  =   "Splits(0).EditorStyle:id=75,.parent=7"
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=21"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(35)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
            _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=90,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(40)  =   ":id=90,.wraptext=-1,.locked=-1"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=22,.alignment=2"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=23"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=75"
            _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=94,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(45)  =   ":id=94,.wraptext=-1,.locked=-1"
            _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=91,.parent=22,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=92,.parent=23"
            _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=93,.parent=75"
            _StyleDefs(49)  =   "Named:id=33:Normal"
            _StyleDefs(50)  =   ":id=33,.parent=0"
            _StyleDefs(51)  =   "Named:id=34:Heading"
            _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(53)  =   ":id=34,.wraptext=-1"
            _StyleDefs(54)  =   "Named:id=35:Footing"
            _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   "Named:id=36:Selected"
            _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(58)  =   "Named:id=37:Caption"
            _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(60)  =   "Named:id=38:HighlightRow"
            _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
            _StyleDefs(62)  =   "Named:id=39:EvenRow"
            _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(64)  =   "Named:id=40:OddRow"
            _StyleDefs(65)  =   ":id=40,.parent=33"
            _StyleDefs(66)  =   "Named:id=41:RecordSelector"
            _StyleDefs(67)  =   ":id=41,.parent=34"
            _StyleDefs(68)  =   "Named:id=42:FilterBar"
            _StyleDefs(69)  =   ":id=42,.parent=33"
         End
         Begin VB.Label LblCliTitular 
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente Titular"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   8025
            TabIndex        =   59
            Top             =   2295
            Width           =   1005
         End
         Begin VB.Label LblCliente 
            Alignment       =   1  'Right Justify
            Caption         =   "Selecione o Cliente:"
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
            Left            =   180
            TabIndex        =   58
            Top             =   585
            Width           =   1995
         End
      End
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   6000
      TabIndex        =   26
      Top             =   4860
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
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
   Begin Threed.SSCommand CmdGravar 
      Height          =   330
      Left            =   2220
      TabIndex        =   25
      Top             =   4860
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
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
End
Attribute VB_Name = "FrmManReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VFV_VETOR1 As New XArray 'Vetor de Compradores
Dim VFV_VETOR2 As New XArray 'Vetor de Corretores
Dim VFV_VETOR3 As New XArray 'Vetor de Fiadores
Dim XFB_ENTROUCOMP As Boolean 'Controla a entrada na orelha de Compradores
Dim XFB_ENTROUFIAD As Boolean 'Controla a entrada na orelha de Fiadores
Dim XFB_ENTROUCORR As Boolean 'Controla a entrada na orelha de Corretores
Dim XFI_CODCLIENTE As Integer, XFI_CODCORRETOR As Integer
Dim XFT_CODIMOVEL As String, XFT_INSCORRETOR As String
Dim XFT_CODEMPREENDIMENTO As String, XFT_CODCONTRATO As String

Dim errLoop As adodb.Error
Dim StrError As String
Sub subDesabilitaBotoes()
    If TabAuxiliar.Tab = 0 Then
        CmdInsComprador.Enabled = False
        CmdExcComprador.Enabled = False
    ElseIf TabAuxiliar.Tab = 2 Then
        CmdInsCorretor.Enabled = False
        CmdAltCorretor.Enabled = False
        CmdExcCorretor.Enabled = False
    End If
End Sub
Sub subHabilitaBotoes()
    If TabAuxiliar.Tab = 0 Then
        CmdInsComprador.Enabled = True
        CmdExcComprador.Enabled = True
    ElseIf TabAuxiliar.Tab = 2 Then
        CmdInsCorretor.Enabled = True
        CmdAltCorretor.Enabled = True
        CmdExcCorretor.Enabled = True
    End If
End Sub
Sub subPreencheCampos()
    Dim XLT_SQL As String
    Dim XLO_CONTRATO As New adodb.Recordset
    
    SubQOpenRecordset XLO_CONTRATO, "SELECT * FROM Contratos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento= '" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XFT_CODIMOVEL & "' AND cont_cd_Contrato='" & XFT_CODCONTRATO & "'", "Estatico"
    If Not (XLO_CONTRATO.BOF And XLO_CONTRATO.EOF) Then
      
      TabAuxiliar.Enabled = True
      
      'Dados da Venda
      TxtAVista.Text = Format(XLO_CONTRATO!cont_vl_Avista, "standard")
      
      If CDbl(TxtAVista.Text) = 0# Then
        TxtAVista.Text = ""
        TxtSinal.Text = Format(XLO_CONTRATO!cont_vl_Sinal, "standard")
        TxtPoupanca.Text = Format(XLO_CONTRATO!cont_vl_Poupanca, "standard")
        TxtFinanciamento.Text = Format(XLO_CONTRATO!cont_vl_Financiamento, "standard")
        TxtTotVenda.Text = Format(XLO_CONTRATO!cont_vl_Total, "standard")
        TxtFgts.Text = Format(XLO_CONTRATO!cont_vl_Fgts, "standard")
        TxtOutros.Text = Format(XLO_CONTRATO!cont_vl_Outros, "standard")
      End If
      
      XFI_CODCLIENTE = XLO_CONTRATO!focl_cd_FornCli
    End If
    XLO_CONTRATO.Close ' Inserido em 13/10/08 - Patrícia
      
End Sub
Sub subPrepararInsercao()
    
    MskEmpreendimento.Text = ""
    TxtEmpreendimento.Text = ""
    TxtImovel.Text = ""
    TxtAVista.Text = ""
    TxtSinal.Text = ""
    TxtPoupanca.Text = ""
    TxtFinanciamento = ""
    TxtTotVenda.Text = ""
    TxtFgts.Text = ""
    CboComprador.Text = ""
    
End Sub
Private Sub CboComprador_DblClick(Area As Integer)
    
    Dim XLV_CHAVETEMP As Variant

    XLV_CHAVETEMP = Chave

    FornClie = "C" 'Cliente
    Pessoa = "F" 'Pessoa Física

    XFormForn = "FrmManContratos"
    FrmFornecedores.Show 1
    Chave = XLV_CHAVETEMP
    Set Formulario = FrmManContratos
    FrmManContratos.DatComprador.Refresh
    FrmManContratos.CboComprador.BoundText = XGV_FornNotaFiscal
    
End Sub
Private Sub CboComprador_GotFocus()
    TabAuxiliar.Tab = 0
End Sub
Private Sub CboCorretor_Change()
    TxtPercComissao.Text = DatCorretor.Recordset.Fields("corr_vl_percentcomissao")
    If TxtAVista.Text <> "" Then
        TxtComissao.Text = Format(((CDbl(TxtAVista.Text) * DatCorretor.Recordset.Fields("corr_vl_percentcomissao")) / 100), "standard")
    Else
        TxtAVista.Text = 0#
        TxtComissao.Text = Format(((CDbl(TxtAVista.Text) * DatCorretor.Recordset.Fields("corr_vl_percentcomissao")) / 100), "standard")
    End If
End Sub
Private Sub CmdAltCorretor_Click()
    CboCorretor.BoundText = TDBGrid2.Columns("Codigo")
    XFI_CODCORRETOR = TDBGrid2.Columns("Codigo")
    'TxtPercComissao.Text = TDBGrid2.Columns("Comissão (%)")
    TxtComissao.Text = TDBGrid2.Columns("Comissão")
    If Not IsNull(DtpPagamento.Value) Then
      DtpPagamento.Value = TDBGrid2.Columns("Pagamento")
    Else
     DtpPagamento.Value = Null
    End If
    XFT_INSCORRETOR = "A"
    PanCorretor.Enabled = True
    PanCorretor.Visible = True
End Sub
Private Sub CmdConfirmaCorretor_Click()
    If CboCorretor.Text <> "" Then
        DatCorretor.Recordset.bookmark = CboCorretor.SelectedItem
        
        If XFT_INSCORRETOR = "A" And XFI_CODCORRETOR = CboCorretor.BoundText Then
            TDBGrid2.Columns(0).Text = CboCorretor.BoundText
            TDBGrid2.Columns(1).Text = CboCorretor.Text
            'TDBGrid2.Columns(2).Text = TxtPercComissao.Text
            TDBGrid2.Columns(2).Text = TxtComissao.Text
            TDBGrid2.Columns(3).Text = DtpPagamento.Value
            TDBGrid2.Update
        Else
            If funExisteChave(1, Array("0"), Array(DatCorretor.Recordset.Fields("corr_cd_corretor")), VFV_VETOR2, TDBGrid2) Then
                MsgBox "Esse Corretor já foi cadastrado", vbCritical, "ATENÇÃO !"
                Exit Sub
            Else
                If XFT_INSCORRETOR = "I" Then
                    subInsereRegistroVetor Array(CboCorretor.BoundText, CboCorretor.Text, TxtComissao.Text, DtpPagamento.Value), VFV_VETOR2, TDBGrid2
                ElseIf XFT_INSCORRETOR = "A" Then
                    TDBGrid2.Columns(0).Text = CboCorretor.BoundText
                    TDBGrid2.Columns(1).Text = CboCorretor.Text
                    'TDBGrid2.Columns(2).Text = PercComissao.Text
                    TDBGrid2.Columns(2).Text = TxtComissao.Text
                    TDBGrid2.Columns(3).Text = DtpPagamento.Value
                    TDBGrid2.Update
                End If
            End If
        End If
        
        If Not CmdExcCorretor.Enabled Then CmdExcCorretor.Enabled = True
        If Not CmdAltCorretor.Enabled Then CmdAltCorretor.Enabled = True
    End If
    
    PanCorretor.Enabled = False
    PanCorretor.Visible = False
End Sub
Private Sub CmdDesistir_Click()
    Unload Me
    Chave = Left$(Chave, 9)
End Sub
Private Sub CmdExcComprador_Click()
   If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then
        subRemoveRegistroVetor VFV_VETOR1, TDBGrid1
        If VFV_VETOR1.Count(1) = 1 Then
            If IsEmpty(VFV_VETOR1(0, 0)) Then
                CmdExcComprador.Enabled = False
            End If
        End If
    End If
End Sub
Private Sub CmdExcCorretor_Click()
   If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then
        subRemoveRegistroVetor VFV_VETOR2, TDBGrid2
        If VFV_VETOR2.Count(1) = 1 Then
            If IsEmpty(VFV_VETOR2(0, 0)) Then
                CmdExcCorretor.Enabled = False
            End If
        End If
    End If
End Sub
Private Sub cmdGravar_Click()
    
    Dim XLT_STATUS As String, XLT_SQL As String
    Dim XLT_JUROS As String, XLT_DESAGIO As String, XLT_SEGURO As String
    Dim XLT_PCR As String, XLT_CONTA As String
    Dim XLO_IMOVEL As Object 'Usado para recuperar o Status do imóvel
    'Dim XLO_CONTRATO As Object 'Usado para verificar se o status do contrato está ativo
    
    'On Error GoTo RotuloErro
    
        
        If XFI_CODCLIENTE = 0 Then
            MsgBox "É preciso definir um Cliente como Titular.", vbCritical, "ATENÇÃO !"
            TabAuxiliar.Tab = 0
            CmdTitular.SetFocus
            Exit Sub
        End If
        
        Conexao.BeginTrans
        If XInserir = "I" Then
            
            XLT_SQL = "INSERT INTO Contratos (empr_cd_Empresa,cont_cd_Contrato,imov_cd_imovel, empd_cd_empreendimento, " & _
              "cont_dt_Venda,cont_dt_LimiteReserva,cont_tx_Status," & _
              "focl_cd_FornCli,cont_vl_AVista, cont_vl_Sinal, " & _
              "cont_vl_Poupanca, cont_vl_Financiamento, cont_vl_Total, cont_vl_Fgts,cont_vl_Outros, " & _
              "cont_tx_TipoJurosMora, cont_tx_Desagio,cont_vl_TaxaSegVida, cont_nr_TipoSeguro, " & _
              "cont_tx_MultaProRata, cont_tx_CorrecProRata, " & _
              "cont_vl_PercJurosFin, cont_vl_PercJurosMulta, cont_vl_PercJurosMora) " & _
              "VALUES (" & PCodEmpresa & ",'" & MskContrato.Text & "','" & XFT_CODIMOVEL & "','" & XFT_CODEMPREENDIMENTO & "', " & _
              "" & FunNuloData(DtpDtReserva.Value, NomeSgbd) & "," & FunNuloData(DtpDtLimiteReserva.Value, NomeSgbd) & ",'R', " & _
              "" & XFI_CODCLIENTE & "," & FunNuloVal(FunTrataFloat(TxtAVista.Text)) & ", " & FunNuloVal(FunTrataFloat(TxtSinal.Text)) & ", " & _
              "" & FunNuloVal(FunTrataFloat(TxtPoupanca.Text)) & ", " & FunNuloVal(FunTrataFloat(TxtFinanciamento.Text)) & "," & FunNuloVal(FunTrataFloat(TxtTotVenda.Text)) & "," & FunNuloVal(FunTrataFloat(TxtFgts.Text)) & "," & FunNuloVal(FunTrataFloat(TxtOutros.Text)) & "," & _
              "'D','D',0,'""'," & _
              "0,0," & _
              "0,0,0)"
            
            Conexao.Execute XLT_SQL
            
            'Altera o ststus do imóvel para reservado
            Conexao.Execute "UPDATE Imoveis SET imov_tx_Status='R' WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
            
            'Quando é feito a reserva
            'soma 1 ao número de unidades vendidas do empreendimento
            Conexao.Execute "UPDATE Empreendimentos SET empd_nr_UnVendidas=empd_nr_UnVendidas +1 WHERE empr_cd_Empresa=" & PCodEmpresa & " AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
            
        
        Else    'Alteração
          
              XLT_SQL = "UPDATE Contratos SET " & _
              "cont_dt_Venda=" & FunNuloData(DtpDtReserva.Value, NomeSgbd) & "," & _
              "cont_dt_LimiteReserva= " & FunNuloData(DtpDtLimiteReserva.Value, NomeSgbd) & "," & _
              "focl_cd_FornCli = " & XFI_CODCLIENTE & "," & _
              "cont_vl_AVista=" & FunNuloVal(FunTrataFloat(TxtAVista.Text)) & "," & _
              "cont_vl_Sinal=" & FunNuloVal(FunTrataFloat(TxtSinal.Text)) & "," & _
              "cont_vl_Poupanca=" & FunNuloVal(FunTrataFloat(TxtPoupanca.Text)) & "," & _
              "cont_vl_Financiamento= " & FunNuloVal(FunTrataFloat(TxtFinanciamento.Text)) & "," & _
              "cont_vl_Total=" & FunNuloVal(FunTrataFloat(TxtTotVenda.Text)) & "," & _
              "cont_vl_Fgts=" & FunNuloVal(FunTrataFloat(TxtFgts.Text)) & "," & _
              "cont_vl_Outros=" & FunNuloVal(FunTrataFloat(TxtOutros.Text)) & "" & _
              " WHERE empr_cd_Empresa=" & PCodEmpresa & " AND cont_cd_Contrato='" & MskContrato.Text & "' AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
            
            Conexao.Execute XLT_SQL
            
        End If
             
        'Compradores
        If XFB_ENTROUCOMP Then
            If XInserir = "A" Then
               Conexao.Execute "DELETE from Compradores WHERE imov_cd_imovel = '" & XFT_CODIMOVEL & "'"
            End If
            If VFV_VETOR1(0, 0) <> Empty Then
                subGravaVetor "Compradores", "focl_cd_FornCli, imov_cd_imovel, empd_cd_empreendimento,empr_cd_Empresa,cont_cd_Contrato", Array("N", "T", "T", "T", "N"), 1, Array(0), "'" & XFT_CODIMOVEL & "', '" & XFT_CODEMPREENDIMENTO & "', " & PCodEmpresa & ",'" & MskContrato.Text & "'", VFV_VETOR1, TDBGrid1
            End If
        End If
 
        
        'Corretores
        If XFB_ENTROUCORR Then
            If XInserir = "A" Then
               Conexao.Execute "DELETE from Comissoes WHERE imov_cd_imovel = '" & XFT_CODIMOVEL & "'"
            End If
            If VFV_VETOR2(0, 0) <> Empty Then
                subGravaVetor "Comissoes", "corr_cd_Corretor, comi_vl_Comissao, comi_dt_Pagamento, imov_cd_imovel, empd_cd_empreendimento,empr_cd_Empresa,cont_cd_Contrato", Array("N", "N", "D", "T", "T", "T", "N"), 3, Array(0, 2, 3), "'" & XFT_CODIMOVEL & "', '" & XFT_CODEMPREENDIMENTO & "', " & PCodEmpresa & ",'" & MskContrato.Text & "'", VFV_VETOR2, TDBGrid2
            End If
        End If
        
        'Atualiza o último contrato na tabela de imóveis
        Conexao.Execute "UPDATE Imoveis SET imov_cd_UltimoContrato='" & MskContrato.Text & "' " & _
        "WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'"
                
        Conexao.CommitTrans
        
        'CmdGravar.Caption = "&Novo"
        'If XInserir = "A" Then
            Call CmdDesistir_Click
        'End If

    Exit Sub
    
RotuloErro:
   Select Case Err.Number
       'MsgBox "Não é possivel concluir a gravação." & Chr(10) & _
                       "A tabela já possui registros cadastrados", _
                       vbCritical + vbApplicationModal + vbOKOnly, "ATENÇÃO"
   End Select
    funTrataErros (ComMensagem)
   ' Enumera a coleção de erros e apresenta
   ' as propriedades de cada erro ocorrido
'   For Each errLoop In Conexao.Errors
'      StrError = "Error #" & errLoop.Number & vbCr & _
'         "   " & errLoop.Description & vbCr & _
'         "   (Source: " & errLoop.Source & ")" & vbCr & _
'         "   (SQL State: " & errLoop.SQLState & ")" & vbCr & _
'         "   (NativeError: " & errLoop.NativeError & ")" & vbCr
'
'   MsgBox StrError, vbCritical + vbOKOnly, "ATENÇÃO !!!"
'   Next

   Resume Next
End Sub
Private Sub CmdInsComprador_Click()
  
  If CboComprador.Text <> "" Then
    DatComprador.Recordset.bookmark = CboComprador.SelectedItem
    
    If funExisteChave(1, Array("0"), Array(DatComprador.Recordset.Fields("focl_cd_forncli")), VFV_VETOR1, TDBGrid1) Then
        MsgBox "Esse Cliente já foi cadastrado", vbCritical, "ATENÇÃO !"
        Exit Sub
    End If
    subInsereRegistroVetor Array(CboComprador.BoundText, CboComprador.Text, ""), VFV_VETOR1, TDBGrid1
    
    If VFV_VETOR1.UpperBound(1) = 0 Then
      Call CmdTitular_Click
    End If
    
    If Not CmdExcComprador.Enabled Then CmdExcComprador.Enabled = True
  End If
  
End Sub
Private Sub CmdInsCorretor_Click()
    
    CboCorretor.Text = ""
    TxtPercComissao.Text = ""
    TxtComissao.Text = ""
    DtpPagamento.Value = ""
    XFT_INSCORRETOR = "I"
    PanCorretor.Visible = True
    PanCorretor.Enabled = True
    
End Sub
Private Sub CmdInsCorretor_GotFocus()
    TabAuxiliar.Tab = 2
End Sub
Private Sub CmdSairCorretor_Click()
    PanCorretor.Enabled = False
    PanCorretor.Visible = False
End Sub
Private Sub CmdTitular_Click()
    Dim XLI_POS As Integer
    Dim XLI_POSPADRAO As Integer
    
    If VFV_VETOR1(0, 0) <> Empty Then
        XLI_POS = TDBGrid1.bookmark
        TDBGrid1.MoveFirst
        Do While Not TDBGrid1.EOF
            If TDBGrid1.Columns("Titular").Value = "S" Then
                XLI_POSPADRAO = TDBGrid1.bookmark
                If XLI_POSPADRAO = XLI_POS Then
                    MsgBox "Este Cliente já é o Titular!", vbCritical, "ATENÇÃO"
                    XFI_CODCLIENTE = TDBGrid1.Columns("Codigo")
                    Exit Sub
                Else
                    If MsgBox("Já existe um Cliente Titular. Deseja alterar?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma alteração
                        TDBGrid1.Columns("Titular").Text = "N"
                        TDBGrid1.Update
                        TDBGrid1.bookmark = XLI_POS
                        TDBGrid1.Columns("Titular").Text = "S"
                        TDBGrid1.Update
                        XFI_CODCLIENTE = TDBGrid1.Columns("Codigo")
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            End If
            TDBGrid1.MoveNext
        Loop
        TDBGrid1.bookmark = XLI_POS
        TDBGrid1.Columns("Titular").Text = "S"
        TDBGrid1.Update
        XFI_CODCLIENTE = TDBGrid1.Columns("Codigo")
    End If

End Sub
Private Sub DtpDtReserva_LostFocus()
  
  DtpDtLimiteReserva.Value = DateAdd("D", 15, DtpDtReserva.Value)
  
End Sub
Private Sub Form_Activate()
    Set Formulario = FrmManReserva
End Sub
Private Sub Form_Load()
    Dim XLT_SQL  As String
    Dim XLO_IMOVEL As New adodb.Recordset  'Usado para preencher os dados do imóvel
      
    Set Formulario = FrmManReserva
        
    If BarraFerramentasPressionada Then
        XInserir = "I"
    End If

   ' CliqueBarraFerramentas = False
    
    If Len(Chave) = 9 Then
      XFT_CODEMPREENDIMENTO = Left$(Chave, 4)
      XFT_CODIMOVEL = Right$(Chave, 4)
    Else
      XFT_CODEMPREENDIMENTO = Left$(Chave, 4)
      XFT_CODIMOVEL = Mid$(Chave, 6, 4)
      XFT_CODCONTRATO = Right$(Chave, 2)
    End If
    
    XFB_ENTROUCOMP = False
    XFB_ENTROUFIAD = False
    'XFB_ENTROUCORR = False
    
    subConectarControleDadosNV DatComprador, "SELECT focl_cd_forncli,focl_tx_classe,focl_tx_razaoSocial FROM FornClientes WHERE focl_tx_classe = 'C' ORDER BY focl_tx_razaoSocial", Estatico
    subConectarControleDadosNV DatCorretor, "SELECT * FROM Corretores ORDER BY corr_tx_nome", Estatico
                                
    'Preenche dados
    MskEmpreendimento.Text = XFT_CODEMPREENDIMENTO
    TxtImovel.Text = XFT_CODIMOVEL
    
    SubQOpenRecordset XLO_IMOVEL, "SELECT * FROM ConsCapImoveis WHERE empr_cd_Empresa=" & PCodEmpresa & " AND imov_cd_imovel='" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "'", "Estatico"
      If Not (XLO_IMOVEL.BOF Or XLO_IMOVEL.EOF) Then
        TxtEmpreendimento.Text = XLO_IMOVEL!empreendimento
        LblTipoImovel2.Caption = XLO_IMOVEL!Tipo
        LblGaragem2.Caption = XLO_IMOVEL!imov_tx_Garagem
        LblArea2.Caption = XLO_IMOVEL!imov_nr_Area
        LblAndar2.Caption = FunNulo(XLO_IMOVEL!imov_nr_Andar)
        LblValorDoImovel.Caption = Format(XLO_IMOVEL!imov_vl_Avista, "standard")
        
        If XInserir = "I" Then
          TxtAVista.Text = Format(XLO_IMOVEL!imov_vl_Avista, "standard")
          'Preenche o numero do contrato automaticamente
          If IsNull(XLO_IMOVEL!imov_cd_UltimoContrato) Then
            MskContrato.Text = "01"
          Else
            MskContrato.Text = Format(CInt(XLO_IMOVEL!imov_cd_UltimoContrato) + 1, "00")
          End If
          DtpDtReserva.Value = Date
          DtpDtLimiteReserva.Value = DateAdd("D", 15, DtpDtReserva.Value)
        ElseIf XInserir = "A" Then
          Me.Caption = "Alteração de Reservas"
          MskContrato.Text = XFT_CODCONTRATO
          subPreencheCampos
        End If
      End If
      
    'Compradores
    XLT_SQL = "SELECT focl_cd_FornCli, focl_tx_RazaoSocial, focl_tx_RazaoSocial FROM ConsCAPCompradores WHERE empr_cd_Empresa = " & PCodEmpresa & " AND imov_cd_imovel = '" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND cont_cd_Contrato='" & MskContrato.Text & "'"
    subCarregaVetorCompradores XLT_SQL, Array("focl_cd_FornCli", "focl_tx_RazaoSocial", "focl_tx_RazaoSocial"), VFV_VETOR1, TDBGrid1
    XFB_ENTROUCOMP = True
         
    XLO_IMOVEL.Close
    Set XLO_IMOVEL = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "FrmManReserva"
End Sub
Private Sub TabAuxiliar_Click(PreviousTab As Integer)
    Dim XLT_SQL As String
    
    If TabAuxiliar.Tab = 0 And Not XFB_ENTROUCOMP Then
        'Compradores
        XLT_SQL = "SELECT focl_cd_FornCli, focl_tx_RazaoSocial, focl_tx_RazaoSocial FROM ConsCAPCompradores WHERE empr_cd_Empresa = " & PCodEmpresa & " AND imov_cd_imovel = '" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND cont_cd_Contrato='" & MskContrato.Text & "'"
        subCarregaVetorCompradores XLT_SQL, Array("focl_cd_FornCli", "focl_tx_RazaoSocial", "focl_tx_RazaoSocial"), VFV_VETOR1, TDBGrid1
        XFB_ENTROUCOMP = True
    ElseIf TabAuxiliar.Tab = 2 And Not XFB_ENTROUCORR Then
        'Corretores
        XLT_SQL = "SELECT * FROM ConsCAPCorretores WHERE empr_cd_Empresa = " & PCodEmpresa & " AND imov_cd_imovel = '" & XFT_CODIMOVEL & "' AND empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND cont_cd_Contrato='" & MskContrato.Text & "'"
        subCarregaVetor XLT_SQL, Array("corr_cd_Corretor", "corr_tx_Nome", "comi_vl_Comissao", "comi_dt_Pagamento"), VFV_VETOR2, TDBGrid2
        XFB_ENTROUCORR = True
    End If
End Sub
Sub subCarregaVetorCompradores(Sql As String, VetorCampos As Variant, Vetor As XArray, Grid As TDBGrid)   'Carrega o vetores
    Dim XLO_COMPRADOR As adodb.Recordset
    Dim XLI_CONT As Integer
    Dim XLI_X As Integer
    
    'sql - string sql com a tabela ou consulta para carregar o vetor e o grid
    'VetorCampos - vetor simples com os campos na respectiva ordem do grid
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    
    SubQOpenRecordset XLO_COMPRADOR, Sql, Estatico
    Vetor.ReDim 0, 0, 0, Grid.Columns.Count - 1
    Vetor.Clear
    Grid.Array = Vetor
    If Not (XLO_COMPRADOR.EOF And XLO_COMPRADOR.BOF) Then
        XLO_COMPRADOR.MoveFirst
        Do While Not XLO_COMPRADOR.EOF
            If Vetor.UpperBound(1) <> 0 Or Vetor(0, 0) <> Empty Then
                Vetor.Insert 1, Vetor.UpperBound(1) + 1
            End If
            XLI_CONT = Vetor.UpperBound(1)
            For XLI_X = 0 To Grid.Columns.Count - 2
                If VetorCampos(XLI_X) <> "" Then
                    subSetaVal Vetor, XLI_CONT, XLI_X, FunNulo(XLO_COMPRADOR(VetorCampos(XLI_X)))
                End If
            Next
            'O próximo if preenche a coluna titular com "S" para o comprador Titular
            If XFI_CODCLIENTE = XLO_COMPRADOR!focl_cd_FornCli And Grid.Columns.Count = 3 Then
                subSetaVal Vetor, XLI_CONT, XLI_X, "S"
            End If
            XLO_COMPRADOR.MoveNext
        Loop
        
        With Grid
            .ReBind
        End With
        
    End If
    XLO_COMPRADOR.Close
    Set XLO_COMPRADOR = Nothing
End Sub
Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    If TDBGrid1.Columns("Titular").CellText(bookmark) = "S" Then
        CellStyle.ForeColor = vbRed
    End If
End Sub
Private Sub TxtAVista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtAVista_LostFocus()
    If IsNumeric(TxtAVista.Text) Then
        TxtAVista.Text = Format$(TxtAVista.Text, "##,##0.00")
        TxtSinal.Text = ""
        TxtPoupanca.Text = ""
        TxtFinanciamento.Text = ""
        TxtFgts.Text = ""
        TxtOutros.Text = ""
        TxtTotVenda.Text = ""
    Else
        TxtAVista.Text = ""
    End If
End Sub
Private Sub TxtComissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtComissao_LostFocus()
    If IsNumeric(TxtComissao.Text) Then
        TxtComissao.Text = Format$(TxtComissao.Text, "##,##0.00")
        TxtPercComissao.Text = Format(CDbl(TxtComissao.Text) * 100 / (CDbl(TxtAVista.Text)), "standard")
    Else
        TxtComissao.Text = ""
    End If
End Sub
Private Sub TxtFgts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtFgts_LostFocus()
     If IsNumeric(TxtFgts.Text) Then
        TxtFgts.Text = Format$(TxtFgts.Text, "##,##0.00")
        TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        TxtAVista.Text = ""
    Else
        If TxtFgts.Text = "" Then
          TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        Else
          TxtFgts.Text = ""
        End If
    End If
End Sub
Private Sub TxtFinanciamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtFinanciamento_LostFocus()
  If IsNumeric(TxtFinanciamento.Text) Then
        TxtFinanciamento.Text = Format$(TxtFinanciamento.Text, "##,##0.00")
        TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        TxtAVista.Text = ""
    Else
        If TxtFinanciamento.Text = "" Then
          TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        Else
          TxtFinanciamento.Text = ""
        End If
    End If
End Sub
Private Sub TxtOutros_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
  End If
End Sub
Private Sub TxtOutros_LostFocus()
   
   If IsNumeric(TxtOutros.Text) Then
        TxtOutros.Text = Format$(TxtOutros.Text, "##,##0.00")
        TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        TxtAVista.Text = ""
    Else
        If TxtOutros.Text = "" Then
          TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        Else
          TxtOutros.Text = ""
        End If
    End If

End Sub
Private Sub TxtPercComissao_LostFocus()
  TxtComissao.Text = Format(CDbl(TxtAVista.Text) * CDbl(TxtPercComissao.Text) / 100, "standard")
End Sub
Private Sub TxtPoupanca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtPoupanca_LostFocus()
    If IsNumeric(TxtPoupanca.Text) Then
        TxtPoupanca.Text = Format$(TxtPoupanca.Text, "##,##0.00")
        TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        TxtAVista.Text = ""
    Else
        If TxtPoupanca.Text = "" Then
          TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        Else
          TxtPoupanca.Text = ""
        End If
    End If
End Sub
Private Sub TxtSinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtSinal_LostFocus()
    If IsNumeric(TxtSinal.Text) Then
        TxtSinal.Text = Format$(TxtSinal.Text, "##,##0.00")
        TxtTotVenda.Text = Format(CDbl(TxtSinal.Text) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        TxtAVista.Text = ""
    Else
        If TxtSinal.Text = "" Then
          TxtTotVenda.Text = Format(CDbl(FunNuloVal(TxtSinal.Text)) + CDbl(FunNuloVal(TxtPoupanca.Text)) + CDbl(FunNuloVal(TxtFinanciamento.Text) + CDbl(FunNuloVal(TxtFgts.Text)) + CDbl(FunNuloVal(TxtOutros.Text))), "standard")
        Else
          TxtSinal.Text = ""
        End If
    End If
End Sub
Private Sub TxtTotVenda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtTotVenda_LostFocus()
    If IsNumeric(TxtTotVenda.Text) Then
        TxtTotVenda.Text = Format$(TxtTotVenda.Text, "##,##0.00")
    Else
        TxtTotVenda.Text = ""
    End If
End Sub



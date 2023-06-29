VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProvisaoImpostos 
   Caption         =   "Provisão Impostos"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12570
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   12570
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "Pesquisar"
      Height          =   495
      Left            =   120
      TabIndex        =   99
      Top             =   7560
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   7215
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   12726
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Contas Correntes"
      TabPicture(0)   =   "frmProvisaoImpostos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraNumeroProvisao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmINSS13"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmProvisao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraCPP"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "IRPJ/CSLL/COFINS/PIS"
      TabPicture(1)   =   "frmProvisaoImpostos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label(4)"
      Tab(1).Control(1)=   "adoIRPJCSLLCOFINSPIS"
      Tab(1).Control(2)=   "dtpCompetenciaPISCOFINS"
      Tab(1).Control(3)=   "tdbgIRPJCSLLCOFINSPIS"
      Tab(1).Control(4)=   "cmdExportIRPJCSLLCOFINSPIS"
      Tab(1).Control(5)=   "fraAliquotas"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "ISS"
      TabPicture(2)   =   "frmProvisaoImpostos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label(8)"
      Tab(2).Control(1)=   "adoISSPagamento"
      Tab(2).Control(2)=   "adoISS"
      Tab(2).Control(3)=   "tdbgISS"
      Tab(2).Control(4)=   "dtpCompetenciaISS"
      Tab(2).Control(5)=   "cmdExportarISS"
      Tab(2).ControlCount=   6
      Begin VB.Frame Frame7 
         Caption         =   "Regime"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   97
         Top             =   480
         Width           =   1575
         Begin VB.ComboBox cmbRegime 
            Height          =   315
            ItemData        =   "frmProvisaoImpostos.frx":0054
            Left            =   120
            List            =   "frmProvisaoImpostos.frx":005E
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame fraCPP 
         Caption         =   "CPP (INSS Simples Nacional)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   82
         Top             =   6240
         Width           =   12135
         Begin VB.TextBox txtCCProvisaoCPPOrigem 
            Height          =   315
            Left            =   1080
            MaxLength       =   7
            TabIndex        =   84
            Top             =   300
            Width           =   555
         End
         Begin VB.TextBox txtCCProvisaoCPPDestino 
            Height          =   315
            Left            =   7080
            MaxLength       =   7
            TabIndex        =   83
            Top             =   300
            Width           =   555
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoCPPOrigem 
            Bindings        =   "frmProvisaoImpostos.frx":0076
            Height          =   315
            Left            =   1680
            TabIndex        =   85
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoCPPDestino 
            Bindings        =   "frmProvisaoImpostos.frx":009B
            Height          =   315
            Left            =   7680
            TabIndex        =   86
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datCCProvisaoCPPOrigem 
            Height          =   330
            Left            =   11040
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin MSAdodcLib.Adodc datCCProvisaoCPPDestino 
            Height          =   330
            Left            =   11040
            Top             =   480
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin VB.Label lblCCCPPDestino 
            Caption         =   "C/C Destino:"
            Height          =   255
            Left            =   6000
            TabIndex        =   88
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblCCCPPOrigem 
            Caption         =   "C/C Origem:"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Baixa Impostos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         TabIndex        =   74
         Top             =   480
         Width           =   7095
         Begin VB.CheckBox chkBaixaImpostos 
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   255
         End
         Begin MSComCtl2.DTPicker dtpBaixaCOFINS 
            Height          =   315
            Left            =   1200
            TabIndex        =   75
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "ddd dd/MM/yyyy"
            DateIsNull      =   -1  'True
            Format          =   146669571
            CurrentDate     =   37460
         End
         Begin MSComCtl2.DTPicker dtpBaixaPIS 
            Height          =   315
            Left            =   3240
            TabIndex        =   76
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "ddd dd/MM/yyyy"
            DateIsNull      =   -1  'True
            Format          =   146669571
            CurrentDate     =   37460
         End
         Begin MSComCtl2.DTPicker dtpBaixaISS 
            Height          =   315
            Left            =   5280
            TabIndex        =   79
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "ddd dd/MM/yyyy"
            Format          =   146669571
            CurrentDate     =   37460
         End
         Begin VB.Label Label18 
            Caption         =   "ISS:"
            Height          =   255
            Left            =   4920
            TabIndex        =   80
            Top             =   285
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "COFINS:"
            Height          =   255
            Left            =   480
            TabIndex        =   78
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "PIS:"
            Height          =   255
            Left            =   2880
            TabIndex        =   77
            Top             =   285
            Width           =   375
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Mês Provisão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   72
         Top             =   480
         Width           =   1455
         Begin MSComCtl2.DTPicker dtpProvisao 
            Height          =   315
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MM/yyyy"
            DateIsNull      =   -1  'True
            Format          =   146669571
            CurrentDate     =   37460
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "IRPF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   65
         Top             =   4560
         Width           =   12135
         Begin VB.TextBox txtCCProvisaoIRPFOrigem 
            Height          =   315
            Left            =   1080
            MaxLength       =   7
            TabIndex        =   67
            Top             =   300
            Width           =   555
         End
         Begin VB.TextBox txtCCProvisaoIRPFDestino 
            Height          =   315
            Left            =   7080
            MaxLength       =   7
            TabIndex        =   66
            Top             =   300
            Width           =   555
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoIRPFDestino 
            Bindings        =   "frmProvisaoImpostos.frx":00C1
            Height          =   315
            Left            =   7680
            TabIndex        =   68
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoIRPFOrigem 
            Bindings        =   "frmProvisaoImpostos.frx":00E7
            Height          =   315
            Left            =   1680
            TabIndex        =   69
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datCCProvisaoIRPFOrigem 
            Height          =   330
            Left            =   11040
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin MSAdodcLib.Adodc datCCProvisaoIRPFDestino 
            Height          =   330
            Left            =   11040
            Top             =   480
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin VB.Label Label15 
            Caption         =   "C/C Destino:"
            Height          =   255
            Left            =   6000
            TabIndex        =   71
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "C/C Origem:"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdExportarISS 
         Height          =   330
         Left            =   -63480
         Picture         =   "frmProvisaoImpostos.frx":010D
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   1020
         Width           =   375
      End
      Begin VB.Frame fraAliquotas 
         Caption         =   "Alíquotas"
         Height          =   855
         Left            =   -74760
         TabIndex        =   53
         Top             =   480
         Width           =   10095
         Begin VB.TextBox txtAliquotaCPP 
            Height          =   285
            Left            =   8520
            TabIndex        =   91
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtAliquotaCOFINS 
            Height          =   285
            Left            =   4800
            TabIndex        =   57
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtAliquotaPIS 
            Height          =   285
            Left            =   6720
            TabIndex        =   56
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtAliquotaCSLL 
            Height          =   285
            Left            =   2640
            TabIndex        =   55
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtAliquotaIRPJ 
            Height          =   285
            Left            =   720
            TabIndex        =   54
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblCPP 
            Caption         =   "CPP"
            Height          =   255
            Left            =   8160
            TabIndex        =   96
            Top             =   390
            Width           =   375
         End
         Begin VB.Label lblPIS 
            Caption         =   "PIS"
            Height          =   255
            Left            =   6240
            TabIndex        =   95
            Top             =   390
            Width           =   375
         End
         Begin VB.Label lblCOFINS 
            Caption         =   "COFINS"
            Height          =   255
            Left            =   4080
            TabIndex        =   94
            Top             =   390
            Width           =   735
         End
         Begin VB.Label lblCSLL 
            Caption         =   "CSLL"
            Height          =   255
            Left            =   2040
            TabIndex        =   93
            Top             =   390
            Width           =   495
         End
         Begin VB.Label lblIRRF 
            Caption         =   "IRRF"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   390
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdExportIRPJCSLLCOFINSPIS 
         Height          =   330
         Left            =   -63240
         Picture         =   "frmProvisaoImpostos.frx":02BC
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   840
         Width           =   375
      End
      Begin VB.Frame frmProvisao 
         Caption         =   "PIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   45
         Top             =   3720
         Width           =   12135
         Begin VB.TextBox txtCCProvisaoPISDestino 
            Height          =   315
            Left            =   7080
            MaxLength       =   7
            TabIndex        =   47
            Top             =   300
            Width           =   555
         End
         Begin VB.TextBox txtCCProvisaoPISOrigem 
            Height          =   315
            Left            =   1080
            MaxLength       =   7
            TabIndex        =   46
            Top             =   300
            Width           =   555
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoPISDestino 
            Bindings        =   "frmProvisaoImpostos.frx":046B
            Height          =   315
            Left            =   7680
            TabIndex        =   48
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoPISOrigem 
            Bindings        =   "frmProvisaoImpostos.frx":0491
            Height          =   315
            Left            =   1680
            TabIndex        =   49
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datCCProvisaoPISOrigem 
            Height          =   330
            Left            =   11040
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin MSAdodcLib.Adodc datCCProvisaoPISDestino 
            Height          =   330
            Left            =   11040
            Top             =   480
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin VB.Label lblProvisao13Origem 
            Caption         =   "C/C Origem:"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblProvisao13Destino 
            Caption         =   "C/C Destino:"
            Height          =   255
            Left            =   6000
            TabIndex        =   50
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame frmINSS13 
         Caption         =   "COFINS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   12135
         Begin VB.TextBox txtCCProvisaoCOFINSOrigem 
            Height          =   315
            Left            =   1080
            MaxLength       =   7
            TabIndex        =   40
            Top             =   300
            Width           =   555
         End
         Begin VB.TextBox txtCCProvisaoCOFINSDestino 
            Height          =   315
            Left            =   7080
            MaxLength       =   7
            TabIndex        =   39
            Top             =   300
            Width           =   555
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoCOFINSOrigem 
            Bindings        =   "frmProvisaoImpostos.frx":04B6
            Height          =   315
            Left            =   1680
            TabIndex        =   41
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoCOFINSDestino 
            Bindings        =   "frmProvisaoImpostos.frx":04DE
            Height          =   315
            Left            =   7680
            TabIndex        =   42
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datCCProvisaoCOFINSOrigem 
            Height          =   330
            Left            =   11040
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin MSAdodcLib.Adodc datCCProvisaoCOFINSDestino 
            Height          =   330
            Left            =   11040
            Top             =   480
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin VB.Label Label5 
            Caption         =   "C/C Destino:"
            Height          =   255
            Left            =   6000
            TabIndex        =   44
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "C/C Origem:"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "IRPJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   12135
         Begin VB.TextBox txtCCProvisaoIRPJDestino 
            Height          =   315
            Left            =   7080
            MaxLength       =   7
            TabIndex        =   33
            Top             =   300
            Width           =   555
         End
         Begin VB.TextBox txtCCProvisaoIRPJOrigem 
            Height          =   315
            Left            =   1080
            MaxLength       =   7
            TabIndex        =   32
            Top             =   300
            Width           =   555
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoIRPJOrigem 
            Bindings        =   "frmProvisaoImpostos.frx":0507
            Height          =   315
            Left            =   1680
            TabIndex        =   34
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoIRPJDestino 
            Bindings        =   "frmProvisaoImpostos.frx":052D
            Height          =   315
            Left            =   7680
            TabIndex        =   35
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datCCProvisaoIRPJOrigem 
            Height          =   330
            Left            =   11040
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin MSAdodcLib.Adodc datCCProvisaoIRPJDestino 
            Height          =   330
            Left            =   11040
            Top             =   480
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin VB.Label Label7 
            Caption         =   "C/C Origem:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "C/C Destino:"
            Height          =   255
            Left            =   6000
            TabIndex        =   36
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "CSLL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   12135
         Begin VB.TextBox txtCCProvisaoCSLLOrigem 
            Height          =   315
            Left            =   1080
            MaxLength       =   7
            TabIndex        =   26
            Top             =   300
            Width           =   555
         End
         Begin VB.TextBox txtCCProvisaoCSLLDestino 
            Height          =   315
            Left            =   7080
            MaxLength       =   7
            TabIndex        =   25
            Top             =   300
            Width           =   555
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoCSLLOrigem 
            Bindings        =   "frmProvisaoImpostos.frx":0554
            Height          =   315
            Left            =   1680
            TabIndex        =   27
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoCSLLDestino 
            Bindings        =   "frmProvisaoImpostos.frx":057A
            Height          =   315
            Left            =   7680
            TabIndex        =   28
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datCCProvisaoCSLLOrigem 
            Height          =   330
            Left            =   11040
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin MSAdodcLib.Adodc datCCProvisaoCSLLDestino 
            Height          =   330
            Left            =   11040
            Top             =   480
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin VB.Label Label9 
            Caption         =   "C/C Destino:"
            Height          =   255
            Left            =   6000
            TabIndex        =   30
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "C/C Origem:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame fraNumeroProvisao 
         Caption         =   "Número da Provisão"
         Height          =   615
         Left            =   10560
         TabIndex        =   22
         Top             =   480
         Width           =   1695
         Begin VB.Label lblNumeroProvisao 
            Caption         =   "0"
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
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ISS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   15
         Top             =   5400
         Width           =   12135
         Begin VB.TextBox txtCCProvisaoISSDestino 
            Height          =   315
            Left            =   7080
            MaxLength       =   7
            TabIndex        =   17
            Top             =   300
            Width           =   555
         End
         Begin VB.TextBox txtCCProvisaoISSOrigem 
            Height          =   315
            Left            =   1080
            MaxLength       =   7
            TabIndex        =   16
            Top             =   300
            Width           =   555
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoISSOrigem 
            Bindings        =   "frmProvisaoImpostos.frx":05A1
            Height          =   315
            Left            =   1680
            TabIndex        =   18
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboCCProvisaoISSDestino 
            Bindings        =   "frmProvisaoImpostos.frx":05C6
            Height          =   315
            Left            =   7680
            TabIndex        =   19
            Top             =   300
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "coco_cd_codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datCCProvisaoISSOrigem 
            Height          =   330
            Left            =   11040
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin MSAdodcLib.Adodc datCCProvisaoISSDestino 
            Height          =   330
            Left            =   11040
            Top             =   480
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   ""
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
         Begin VB.Label Label4 
            Caption         =   "C/C Origem:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "C/C Destino:"
            Height          =   255
            Left            =   6000
            TabIndex        =   20
            Top             =   360
            Width           =   975
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdbgIRPJCSLLCOFINSPIS 
         Bindings        =   "frmProvisaoImpostos.frx":05EC
         Height          =   5055
         Left            =   -74760
         TabIndex        =   58
         Top             =   1440
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   8916
         _LayoutType     =   4
         _RowHeight      =   25
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Centro de Resultado"
         Columns(0).DataField=   "CRDescricao"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Estruturado"
         Columns(1).DataField=   "CRCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   16
         Columns(2)._MaxComboItems=   5
         Columns(2).ValueItems(0)._DefaultItem=   0
         Columns(2).ValueItems(0).Value=   "1"
         Columns(2).ValueItems(0).Value.vt=   8
         Columns(2).ValueItems(0).DisplayValue=   "Simples"
         Columns(2).ValueItems(0).DisplayValue.vt=   8
         Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(1)._DefaultItem=   0
         Columns(2).ValueItems(1).Value=   "2"
         Columns(2).ValueItems(1).Value.vt=   8
         Columns(2).ValueItems(1).DisplayValue=   "Simples"
         Columns(2).ValueItems(1).DisplayValue.vt=   8
         Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(2)._DefaultItem=   0
         Columns(2).ValueItems(2).Value=   "3"
         Columns(2).ValueItems(2).Value.vt=   8
         Columns(2).ValueItems(2).DisplayValue=   "Composta"
         Columns(2).ValueItems(2).DisplayValue.vt=   8
         Columns(2).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems.Count=   3
         Columns(2).Caption=   "Faturamento"
         Columns(2).DataField=   "ValorFaturado"
         Columns(2).NumberFormat=   "Standard"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "% Lucro"
         Columns(3).DataField=   "PercentualLucro"
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Devido IRPJ"
         Columns(4).DataField=   "ValorDevidoIRRF"
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Retido IRPJ"
         Columns(5).DataField=   "ValorRetidoIRRF"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "% IRPJ"
         Columns(6).DataField=   "RetencaoIRRF"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Provisão IRPJ"
         Columns(7).DataField=   "ValorProvisaoIRRF"
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Devido CSLL"
         Columns(8).DataField=   "ValorDevidoCSLL"
         Columns(8).NumberFormat=   "Standard"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Retido CSLL"
         Columns(9).DataField=   "ValorRetidoCSLL"
         Columns(9).NumberFormat=   "Standard"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "% CSLL"
         Columns(10).DataField=   "RetencaoCSLL"
         Columns(10).NumberFormat=   "Standard"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Provisão CSLL"
         Columns(11).DataField=   "ValorProvisaoCSLL"
         Columns(11).NumberFormat=   "Standard"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Devido COFINS"
         Columns(12).DataField=   "ValorDevidoCOFINS"
         Columns(12).NumberFormat=   "Standard"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Retido COFINS"
         Columns(13).DataField=   "ValorRetidoCOFINS"
         Columns(13).NumberFormat=   "Standard"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "% COFINS"
         Columns(14).DataField=   "RetencaoCOFINS"
         Columns(14).NumberFormat=   "Standard"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "IRPJ -> COFINS"
         Columns(15).DataField=   "TransfIRPJCOFINS"
         Columns(15).NumberFormat=   "Standard"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "CSLL -> COFINS"
         Columns(16).DataField=   "TransfCSLLCOFINS"
         Columns(16).NumberFormat=   "Standard"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "Provisão COFINS"
         Columns(17).DataField=   "ValorProvisaoCOFINS"
         Columns(17).NumberFormat=   "Standard"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "Devido PIS"
         Columns(18).DataField=   "ValorDevidoPIS"
         Columns(18).NumberFormat=   "Standard"
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "Retido PIS"
         Columns(19).DataField=   "ValorRetidoPIS"
         Columns(19).NumberFormat=   "Standard"
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   0
         Columns(20)._MaxComboItems=   5
         Columns(20).Caption=   "% PIS"
         Columns(20).DataField=   "RetencaoPIS"
         Columns(20).NumberFormat=   "Standard"
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(21)._VlistStyle=   0
         Columns(21)._MaxComboItems=   5
         Columns(21).Caption=   "COFINS -> PIS"
         Columns(21).DataField=   "TransfCOFINSPIS"
         Columns(21).NumberFormat=   "Standard"
         Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(22)._VlistStyle=   0
         Columns(22)._MaxComboItems=   5
         Columns(22).Caption=   "Provisão PIS"
         Columns(22).DataField=   "ValorProvisaoPIS"
         Columns(22).NumberFormat=   "Standard"
         Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(23)._VlistStyle=   0
         Columns(23)._MaxComboItems=   5
         Columns(23).Caption=   "PIS -> IRPF"
         Columns(23).DataField=   "TransPISIRPF"
         Columns(23).NumberFormat=   "Standard"
         Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(24)._VlistStyle=   0
         Columns(24)._MaxComboItems=   5
         Columns(24).Caption=   "% CPP"
         Columns(24).DataField=   "RetencaoCPP"
         Columns(24).NumberFormat=   "Standard"
         Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(25)._VlistStyle=   0
         Columns(25)._MaxComboItems=   5
         Columns(25).Caption=   "Provisão CPP"
         Columns(25).DataField=   "ValorProvisaoCPP"
         Columns(25).NumberFormat=   "Standard"
         Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   26
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).SizeMode=   2
         Splits(0).Size  =   4
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   0
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=26"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4419"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4339"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=113"
         Splits(0)._ColumnProps(10)=   "Column(1).Width=1588"
         Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=1508"
         Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=196628"
         Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(17)=   "Column(2).Width=2514"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2434"
         Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=197394"
         Splits(0)._ColumnProps(22)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(23)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(25)=   "Column(3).Width=1244"
         Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=1164"
         Splits(0)._ColumnProps(28)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(3)._ColStyle=786"
         Splits(0)._ColumnProps(30)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(31)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(32)=   "Column(3)._MinWidth=29"
         Splits(0)._ColumnProps(33)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(34)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(36)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(37)=   "Column(4)._ColStyle=786"
         Splits(0)._ColumnProps(38)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(39)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(40)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(41)=   "Column(4)._MinWidth=29"
         Splits(0)._ColumnProps(42)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(43)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(45)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(46)=   "Column(5)._ColStyle=20"
         Splits(0)._ColumnProps(47)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(48)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(49)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(50)=   "Column(5)._MinWidth=29"
         Splits(0)._ColumnProps(51)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(52)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(53)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(54)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(55)=   "Column(6)._ColStyle=20"
         Splits(0)._ColumnProps(56)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(57)=   "Column(6).WrapText=1"
         Splits(0)._ColumnProps(58)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(59)=   "Column(6)._MinWidth=29"
         Splits(0)._ColumnProps(60)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(61)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(62)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(63)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(64)=   "Column(7)._ColStyle=20"
         Splits(0)._ColumnProps(65)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(66)=   "Column(7).WrapText=1"
         Splits(0)._ColumnProps(67)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(68)=   "Column(7)._MinWidth=29"
         Splits(0)._ColumnProps(69)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(70)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(72)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(73)=   "Column(8)._ColStyle=20"
         Splits(0)._ColumnProps(74)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(75)=   "Column(8).WrapText=1"
         Splits(0)._ColumnProps(76)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(77)=   "Column(8)._MinWidth=29"
         Splits(0)._ColumnProps(78)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(79)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(80)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(81)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(82)=   "Column(9)._ColStyle=20"
         Splits(0)._ColumnProps(83)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(84)=   "Column(9).WrapText=1"
         Splits(0)._ColumnProps(85)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(86)=   "Column(9)._MinWidth=29"
         Splits(0)._ColumnProps(87)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(88)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(89)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(90)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(91)=   "Column(10)._ColStyle=20"
         Splits(0)._ColumnProps(92)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(93)=   "Column(10).WrapText=1"
         Splits(0)._ColumnProps(94)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(95)=   "Column(10)._MinWidth=29"
         Splits(0)._ColumnProps(96)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(97)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(98)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(99)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(100)=   "Column(11)._ColStyle=20"
         Splits(0)._ColumnProps(101)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(102)=   "Column(11).WrapText=1"
         Splits(0)._ColumnProps(103)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(104)=   "Column(11)._MinWidth=29"
         Splits(0)._ColumnProps(105)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(106)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(107)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(108)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(109)=   "Column(12)._ColStyle=20"
         Splits(0)._ColumnProps(110)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(111)=   "Column(12).WrapText=1"
         Splits(0)._ColumnProps(112)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(113)=   "Column(12)._MinWidth=29"
         Splits(0)._ColumnProps(114)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(115)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(116)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(117)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(118)=   "Column(13)._ColStyle=20"
         Splits(0)._ColumnProps(119)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(120)=   "Column(13).WrapText=1"
         Splits(0)._ColumnProps(121)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(122)=   "Column(13)._MinWidth=29"
         Splits(0)._ColumnProps(123)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(124)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(125)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(126)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(127)=   "Column(14)._ColStyle=20"
         Splits(0)._ColumnProps(128)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(129)=   "Column(14).WrapText=1"
         Splits(0)._ColumnProps(130)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(131)=   "Column(14)._MinWidth=29"
         Splits(0)._ColumnProps(132)=   "Column(15).Width=2725"
         Splits(0)._ColumnProps(133)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(134)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(135)=   "Column(15)._EditAlways=0"
         Splits(0)._ColumnProps(136)=   "Column(15)._ColStyle=20"
         Splits(0)._ColumnProps(137)=   "Column(15).Visible=0"
         Splits(0)._ColumnProps(138)=   "Column(15).WrapText=1"
         Splits(0)._ColumnProps(139)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(140)=   "Column(15)._MinWidth=29"
         Splits(0)._ColumnProps(141)=   "Column(16).Width=2725"
         Splits(0)._ColumnProps(142)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(143)=   "Column(16)._WidthInPix=2646"
         Splits(0)._ColumnProps(144)=   "Column(16)._EditAlways=0"
         Splits(0)._ColumnProps(145)=   "Column(16)._ColStyle=20"
         Splits(0)._ColumnProps(146)=   "Column(16).Visible=0"
         Splits(0)._ColumnProps(147)=   "Column(16).WrapText=1"
         Splits(0)._ColumnProps(148)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(149)=   "Column(16)._MinWidth=29"
         Splits(0)._ColumnProps(150)=   "Column(17).Width=2725"
         Splits(0)._ColumnProps(151)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(152)=   "Column(17)._WidthInPix=2646"
         Splits(0)._ColumnProps(153)=   "Column(17)._EditAlways=0"
         Splits(0)._ColumnProps(154)=   "Column(17)._ColStyle=20"
         Splits(0)._ColumnProps(155)=   "Column(17).Visible=0"
         Splits(0)._ColumnProps(156)=   "Column(17).WrapText=1"
         Splits(0)._ColumnProps(157)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(158)=   "Column(17)._MinWidth=29"
         Splits(0)._ColumnProps(159)=   "Column(18).Width=2725"
         Splits(0)._ColumnProps(160)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(161)=   "Column(18)._WidthInPix=2646"
         Splits(0)._ColumnProps(162)=   "Column(18)._EditAlways=0"
         Splits(0)._ColumnProps(163)=   "Column(18)._ColStyle=20"
         Splits(0)._ColumnProps(164)=   "Column(18).Visible=0"
         Splits(0)._ColumnProps(165)=   "Column(18).WrapText=1"
         Splits(0)._ColumnProps(166)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(167)=   "Column(18)._MinWidth=29"
         Splits(0)._ColumnProps(168)=   "Column(19).Width=2461"
         Splits(0)._ColumnProps(169)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(170)=   "Column(19)._WidthInPix=2381"
         Splits(0)._ColumnProps(171)=   "Column(19)._EditAlways=0"
         Splits(0)._ColumnProps(172)=   "Column(19)._ColStyle=532"
         Splits(0)._ColumnProps(173)=   "Column(19).Visible=0"
         Splits(0)._ColumnProps(174)=   "Column(19).WrapText=1"
         Splits(0)._ColumnProps(175)=   "Column(19).AllowFocus=0"
         Splits(0)._ColumnProps(176)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(177)=   "Column(19)._MinWidth=29"
         Splits(0)._ColumnProps(178)=   "Column(20).Width=2725"
         Splits(0)._ColumnProps(179)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(180)=   "Column(20)._WidthInPix=2646"
         Splits(0)._ColumnProps(181)=   "Column(20)._EditAlways=0"
         Splits(0)._ColumnProps(182)=   "Column(20)._ColStyle=20"
         Splits(0)._ColumnProps(183)=   "Column(20).Visible=0"
         Splits(0)._ColumnProps(184)=   "Column(20).WrapText=1"
         Splits(0)._ColumnProps(185)=   "Column(20).Order=21"
         Splits(0)._ColumnProps(186)=   "Column(20)._MinWidth=123792704"
         Splits(0)._ColumnProps(187)=   "Column(21).Width=2725"
         Splits(0)._ColumnProps(188)=   "Column(21).DividerColor=0"
         Splits(0)._ColumnProps(189)=   "Column(21)._WidthInPix=2646"
         Splits(0)._ColumnProps(190)=   "Column(21)._EditAlways=0"
         Splits(0)._ColumnProps(191)=   "Column(21)._ColStyle=20"
         Splits(0)._ColumnProps(192)=   "Column(21).Visible=0"
         Splits(0)._ColumnProps(193)=   "Column(21).WrapText=1"
         Splits(0)._ColumnProps(194)=   "Column(21).Order=22"
         Splits(0)._ColumnProps(195)=   "Column(21)._MinWidth=123792704"
         Splits(0)._ColumnProps(196)=   "Column(22).Width=2461"
         Splits(0)._ColumnProps(197)=   "Column(22).DividerColor=0"
         Splits(0)._ColumnProps(198)=   "Column(22)._WidthInPix=2381"
         Splits(0)._ColumnProps(199)=   "Column(22)._EditAlways=0"
         Splits(0)._ColumnProps(200)=   "Column(22)._ColStyle=20"
         Splits(0)._ColumnProps(201)=   "Column(22).Visible=0"
         Splits(0)._ColumnProps(202)=   "Column(22).WrapText=1"
         Splits(0)._ColumnProps(203)=   "Column(22).Order=23"
         Splits(0)._ColumnProps(204)=   "Column(22)._MinWidth=123792704"
         Splits(0)._ColumnProps(205)=   "Column(23).Width=2725"
         Splits(0)._ColumnProps(206)=   "Column(23).DividerColor=0"
         Splits(0)._ColumnProps(207)=   "Column(23)._WidthInPix=2646"
         Splits(0)._ColumnProps(208)=   "Column(23)._EditAlways=0"
         Splits(0)._ColumnProps(209)=   "Column(23)._ColStyle=20"
         Splits(0)._ColumnProps(210)=   "Column(23).Visible=0"
         Splits(0)._ColumnProps(211)=   "Column(23).WrapText=1"
         Splits(0)._ColumnProps(212)=   "Column(23).Order=24"
         Splits(0)._ColumnProps(213)=   "Column(24).Width=2725"
         Splits(0)._ColumnProps(214)=   "Column(24).DividerColor=0"
         Splits(0)._ColumnProps(215)=   "Column(24)._WidthInPix=2646"
         Splits(0)._ColumnProps(216)=   "Column(24)._EditAlways=0"
         Splits(0)._ColumnProps(217)=   "Column(24)._ColStyle=20"
         Splits(0)._ColumnProps(218)=   "Column(24).Visible=0"
         Splits(0)._ColumnProps(219)=   "Column(24).WrapText=1"
         Splits(0)._ColumnProps(220)=   "Column(24).Order=25"
         Splits(0)._ColumnProps(221)=   "Column(25).Width=2725"
         Splits(0)._ColumnProps(222)=   "Column(25).DividerColor=0"
         Splits(0)._ColumnProps(223)=   "Column(25)._WidthInPix=2646"
         Splits(0)._ColumnProps(224)=   "Column(25)._EditAlways=0"
         Splits(0)._ColumnProps(225)=   "Column(25)._ColStyle=20"
         Splits(0)._ColumnProps(226)=   "Column(25).Visible=0"
         Splits(0)._ColumnProps(227)=   "Column(25).WrapText=1"
         Splits(0)._ColumnProps(228)=   "Column(25).Order=26"
         Splits(1)._UserFlags=   0
         Splits(1).ExtendRightColumn=   -1  'True
         Splits(1).MarqueeStyle=   3
         Splits(1).Size  =   5
         Splits(1).Size.vt=   2
         Splits(1).RecordSelectors=   0   'False
         Splits(1).RecordSelectorWidth=   503
         Splits(1)._SavedRecordSelectors=   0   'False
         Splits(1).AllowColMove=   -1  'True
         Splits(1).DividerStyle=   2
         Splits(1).DividerColor=   12632256
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=26"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=4419"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=4339"
         Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=528"
         Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(7)=   "Column(0).WrapText=1"
         Splits(1)._ColumnProps(8)=   "Column(0).AllowFocus=0"
         Splits(1)._ColumnProps(9)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(10)=   "Column(1).Width=2725"
         Splits(1)._ColumnProps(11)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(12)=   "Column(1)._WidthInPix=2646"
         Splits(1)._ColumnProps(13)=   "Column(1)._EditAlways=0"
         Splits(1)._ColumnProps(14)=   "Column(1)._ColStyle=20"
         Splits(1)._ColumnProps(15)=   "Column(1).Visible=0"
         Splits(1)._ColumnProps(16)=   "Column(1).WrapText=1"
         Splits(1)._ColumnProps(17)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(18)=   "Column(1)._MinWidth=104849504"
         Splits(1)._ColumnProps(19)=   "Column(2).Width=2461"
         Splits(1)._ColumnProps(20)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(21)=   "Column(2)._WidthInPix=2381"
         Splits(1)._ColumnProps(22)=   "Column(2)._EditAlways=0"
         Splits(1)._ColumnProps(23)=   "Column(2)._ColStyle=197138"
         Splits(1)._ColumnProps(24)=   "Column(2).Visible=0"
         Splits(1)._ColumnProps(25)=   "Column(2).WrapText=1"
         Splits(1)._ColumnProps(26)=   "Column(2).AllowFocus=0"
         Splits(1)._ColumnProps(27)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(28)=   "Column(2)._MinWidth=104849504"
         Splits(1)._ColumnProps(29)=   "Column(3).Width=2725"
         Splits(1)._ColumnProps(30)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(31)=   "Column(3)._WidthInPix=2646"
         Splits(1)._ColumnProps(32)=   "Column(3)._EditAlways=0"
         Splits(1)._ColumnProps(33)=   "Column(3)._ColStyle=20"
         Splits(1)._ColumnProps(34)=   "Column(3).Visible=0"
         Splits(1)._ColumnProps(35)=   "Column(3).WrapText=1"
         Splits(1)._ColumnProps(36)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(37)=   "Column(4).Width=2725"
         Splits(1)._ColumnProps(38)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(39)=   "Column(4)._WidthInPix=2646"
         Splits(1)._ColumnProps(40)=   "Column(4)._EditAlways=0"
         Splits(1)._ColumnProps(41)=   "Column(4)._ColStyle=197394"
         Splits(1)._ColumnProps(42)=   "Column(4).WrapText=1"
         Splits(1)._ColumnProps(43)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(44)=   "Column(5).Width=2725"
         Splits(1)._ColumnProps(45)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(46)=   "Column(5)._WidthInPix=2646"
         Splits(1)._ColumnProps(47)=   "Column(5)._EditAlways=0"
         Splits(1)._ColumnProps(48)=   "Column(5)._ColStyle=197394"
         Splits(1)._ColumnProps(49)=   "Column(5).WrapText=1"
         Splits(1)._ColumnProps(50)=   "Column(5).Order=6"
         Splits(1)._ColumnProps(51)=   "Column(6).Width=1058"
         Splits(1)._ColumnProps(52)=   "Column(6).DividerColor=0"
         Splits(1)._ColumnProps(53)=   "Column(6)._WidthInPix=979"
         Splits(1)._ColumnProps(54)=   "Column(6)._EditAlways=0"
         Splits(1)._ColumnProps(55)=   "Column(6)._ColStyle=197394"
         Splits(1)._ColumnProps(56)=   "Column(6).WrapText=1"
         Splits(1)._ColumnProps(57)=   "Column(6).Order=7"
         Splits(1)._ColumnProps(58)=   "Column(7).Width=2725"
         Splits(1)._ColumnProps(59)=   "Column(7).DividerColor=0"
         Splits(1)._ColumnProps(60)=   "Column(7)._WidthInPix=2646"
         Splits(1)._ColumnProps(61)=   "Column(7)._EditAlways=0"
         Splits(1)._ColumnProps(62)=   "Column(7)._ColStyle=197394"
         Splits(1)._ColumnProps(63)=   "Column(7).WrapText=1"
         Splits(1)._ColumnProps(64)=   "Column(7).Order=8"
         Splits(1)._ColumnProps(65)=   "Column(8).Width=2725"
         Splits(1)._ColumnProps(66)=   "Column(8).DividerColor=0"
         Splits(1)._ColumnProps(67)=   "Column(8)._WidthInPix=2646"
         Splits(1)._ColumnProps(68)=   "Column(8)._EditAlways=0"
         Splits(1)._ColumnProps(69)=   "Column(8)._ColStyle=197394"
         Splits(1)._ColumnProps(70)=   "Column(8).WrapText=1"
         Splits(1)._ColumnProps(71)=   "Column(8).Order=9"
         Splits(1)._ColumnProps(72)=   "Column(9).Width=2725"
         Splits(1)._ColumnProps(73)=   "Column(9).DividerColor=0"
         Splits(1)._ColumnProps(74)=   "Column(9)._WidthInPix=2646"
         Splits(1)._ColumnProps(75)=   "Column(9)._EditAlways=0"
         Splits(1)._ColumnProps(76)=   "Column(9)._ColStyle=197394"
         Splits(1)._ColumnProps(77)=   "Column(9).WrapText=1"
         Splits(1)._ColumnProps(78)=   "Column(9).Order=10"
         Splits(1)._ColumnProps(79)=   "Column(10).Width=1244"
         Splits(1)._ColumnProps(80)=   "Column(10).DividerColor=0"
         Splits(1)._ColumnProps(81)=   "Column(10)._WidthInPix=1164"
         Splits(1)._ColumnProps(82)=   "Column(10)._EditAlways=0"
         Splits(1)._ColumnProps(83)=   "Column(10)._ColStyle=197394"
         Splits(1)._ColumnProps(84)=   "Column(10).WrapText=1"
         Splits(1)._ColumnProps(85)=   "Column(10).Order=11"
         Splits(1)._ColumnProps(86)=   "Column(11).Width=2725"
         Splits(1)._ColumnProps(87)=   "Column(11).DividerColor=0"
         Splits(1)._ColumnProps(88)=   "Column(11)._WidthInPix=2646"
         Splits(1)._ColumnProps(89)=   "Column(11)._EditAlways=0"
         Splits(1)._ColumnProps(90)=   "Column(11)._ColStyle=197394"
         Splits(1)._ColumnProps(91)=   "Column(11).WrapText=1"
         Splits(1)._ColumnProps(92)=   "Column(11).Order=12"
         Splits(1)._ColumnProps(93)=   "Column(12).Width=2725"
         Splits(1)._ColumnProps(94)=   "Column(12).DividerColor=0"
         Splits(1)._ColumnProps(95)=   "Column(12)._WidthInPix=2646"
         Splits(1)._ColumnProps(96)=   "Column(12)._EditAlways=0"
         Splits(1)._ColumnProps(97)=   "Column(12)._ColStyle=197394"
         Splits(1)._ColumnProps(98)=   "Column(12).WrapText=1"
         Splits(1)._ColumnProps(99)=   "Column(12).Order=13"
         Splits(1)._ColumnProps(100)=   "Column(13).Width=2725"
         Splits(1)._ColumnProps(101)=   "Column(13).DividerColor=0"
         Splits(1)._ColumnProps(102)=   "Column(13)._WidthInPix=2646"
         Splits(1)._ColumnProps(103)=   "Column(13)._EditAlways=0"
         Splits(1)._ColumnProps(104)=   "Column(13)._ColStyle=197394"
         Splits(1)._ColumnProps(105)=   "Column(13).WrapText=1"
         Splits(1)._ColumnProps(106)=   "Column(13).Order=14"
         Splits(1)._ColumnProps(107)=   "Column(14).Width=1588"
         Splits(1)._ColumnProps(108)=   "Column(14).DividerColor=0"
         Splits(1)._ColumnProps(109)=   "Column(14)._WidthInPix=1508"
         Splits(1)._ColumnProps(110)=   "Column(14)._EditAlways=0"
         Splits(1)._ColumnProps(111)=   "Column(14)._ColStyle=197394"
         Splits(1)._ColumnProps(112)=   "Column(14).WrapText=1"
         Splits(1)._ColumnProps(113)=   "Column(14).Order=15"
         Splits(1)._ColumnProps(114)=   "Column(15).Width=2725"
         Splits(1)._ColumnProps(115)=   "Column(15).DividerColor=0"
         Splits(1)._ColumnProps(116)=   "Column(15)._WidthInPix=2646"
         Splits(1)._ColumnProps(117)=   "Column(15)._EditAlways=0"
         Splits(1)._ColumnProps(118)=   "Column(15)._ColStyle=197394"
         Splits(1)._ColumnProps(119)=   "Column(15).WrapText=1"
         Splits(1)._ColumnProps(120)=   "Column(15).Order=16"
         Splits(1)._ColumnProps(121)=   "Column(16).Width=2725"
         Splits(1)._ColumnProps(122)=   "Column(16).DividerColor=0"
         Splits(1)._ColumnProps(123)=   "Column(16)._WidthInPix=2646"
         Splits(1)._ColumnProps(124)=   "Column(16)._EditAlways=0"
         Splits(1)._ColumnProps(125)=   "Column(16)._ColStyle=197394"
         Splits(1)._ColumnProps(126)=   "Column(16).WrapText=1"
         Splits(1)._ColumnProps(127)=   "Column(16).Order=17"
         Splits(1)._ColumnProps(128)=   "Column(17).Width=2725"
         Splits(1)._ColumnProps(129)=   "Column(17).DividerColor=0"
         Splits(1)._ColumnProps(130)=   "Column(17)._WidthInPix=2646"
         Splits(1)._ColumnProps(131)=   "Column(17)._EditAlways=0"
         Splits(1)._ColumnProps(132)=   "Column(17)._ColStyle=197394"
         Splits(1)._ColumnProps(133)=   "Column(17).WrapText=1"
         Splits(1)._ColumnProps(134)=   "Column(17).Order=18"
         Splits(1)._ColumnProps(135)=   "Column(18).Width=2725"
         Splits(1)._ColumnProps(136)=   "Column(18).DividerColor=0"
         Splits(1)._ColumnProps(137)=   "Column(18)._WidthInPix=2646"
         Splits(1)._ColumnProps(138)=   "Column(18)._EditAlways=0"
         Splits(1)._ColumnProps(139)=   "Column(18)._ColStyle=197394"
         Splits(1)._ColumnProps(140)=   "Column(18).WrapText=1"
         Splits(1)._ColumnProps(141)=   "Column(18).Order=19"
         Splits(1)._ColumnProps(142)=   "Column(19).Width=2461"
         Splits(1)._ColumnProps(143)=   "Column(19).DividerColor=0"
         Splits(1)._ColumnProps(144)=   "Column(19)._WidthInPix=2381"
         Splits(1)._ColumnProps(145)=   "Column(19)._EditAlways=0"
         Splits(1)._ColumnProps(146)=   "Column(19)._ColStyle=197394"
         Splits(1)._ColumnProps(147)=   "Column(19).WrapText=1"
         Splits(1)._ColumnProps(148)=   "Column(19).AllowFocus=0"
         Splits(1)._ColumnProps(149)=   "Column(19).Order=20"
         Splits(1)._ColumnProps(150)=   "Column(20).Width=1058"
         Splits(1)._ColumnProps(151)=   "Column(20).DividerColor=0"
         Splits(1)._ColumnProps(152)=   "Column(20)._WidthInPix=979"
         Splits(1)._ColumnProps(153)=   "Column(20)._EditAlways=0"
         Splits(1)._ColumnProps(154)=   "Column(20)._ColStyle=197394"
         Splits(1)._ColumnProps(155)=   "Column(20).WrapText=1"
         Splits(1)._ColumnProps(156)=   "Column(20).Order=21"
         Splits(1)._ColumnProps(157)=   "Column(21).Width=2725"
         Splits(1)._ColumnProps(158)=   "Column(21).DividerColor=0"
         Splits(1)._ColumnProps(159)=   "Column(21)._WidthInPix=2646"
         Splits(1)._ColumnProps(160)=   "Column(21)._EditAlways=0"
         Splits(1)._ColumnProps(161)=   "Column(21)._ColStyle=197394"
         Splits(1)._ColumnProps(162)=   "Column(21).WrapText=1"
         Splits(1)._ColumnProps(163)=   "Column(21).Order=22"
         Splits(1)._ColumnProps(164)=   "Column(22).Width=2461"
         Splits(1)._ColumnProps(165)=   "Column(22).DividerColor=0"
         Splits(1)._ColumnProps(166)=   "Column(22)._WidthInPix=2381"
         Splits(1)._ColumnProps(167)=   "Column(22)._EditAlways=0"
         Splits(1)._ColumnProps(168)=   "Column(22)._ColStyle=197394"
         Splits(1)._ColumnProps(169)=   "Column(22).WrapText=1"
         Splits(1)._ColumnProps(170)=   "Column(22).Order=23"
         Splits(1)._ColumnProps(171)=   "Column(23).Width=2725"
         Splits(1)._ColumnProps(172)=   "Column(23).DividerColor=0"
         Splits(1)._ColumnProps(173)=   "Column(23)._WidthInPix=2646"
         Splits(1)._ColumnProps(174)=   "Column(23)._EditAlways=0"
         Splits(1)._ColumnProps(175)=   "Column(23)._ColStyle=196626"
         Splits(1)._ColumnProps(176)=   "Column(23).WrapText=1"
         Splits(1)._ColumnProps(177)=   "Column(23).Order=24"
         Splits(1)._ColumnProps(178)=   "Column(24).Width=2725"
         Splits(1)._ColumnProps(179)=   "Column(24).DividerColor=0"
         Splits(1)._ColumnProps(180)=   "Column(24)._WidthInPix=2646"
         Splits(1)._ColumnProps(181)=   "Column(24)._EditAlways=0"
         Splits(1)._ColumnProps(182)=   "Column(24)._ColStyle=18"
         Splits(1)._ColumnProps(183)=   "Column(24).WrapText=1"
         Splits(1)._ColumnProps(184)=   "Column(24).Order=25"
         Splits(1)._ColumnProps(185)=   "Column(25).Width=2725"
         Splits(1)._ColumnProps(186)=   "Column(25).DividerColor=0"
         Splits(1)._ColumnProps(187)=   "Column(25)._WidthInPix=2646"
         Splits(1)._ColumnProps(188)=   "Column(25)._EditAlways=0"
         Splits(1)._ColumnProps(189)=   "Column(25)._ColStyle=18"
         Splits(1)._ColumnProps(190)=   "Column(25).WrapText=1"
         Splits(1)._ColumnProps(191)=   "Column(25).Order=26"
         Splits(1)._ColumnProps(192)=   "Column(25)._MinWidth=106503328"
         Splits.Count    =   2
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         ColumnFooters   =   -1  'True
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   2
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
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
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=42,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(24)  =   "Splits(0).Style:id=29,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=128,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=30,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=31,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=32,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=124,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=123,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=125,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=126,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=127,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=129,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=130,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=134,.parent=29,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=131,.parent=30,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=132,.parent=31"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=133,.parent=123"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=90,.parent=29"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=87,.parent=30"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=88,.parent=31,.namedParent=35"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=89,.parent=123"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=138,.parent=29,.alignment=1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=135,.parent=30,.alignment=1"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=136,.parent=31,.namedParent=35"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=137,.parent=123"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=226,.parent=29,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=223,.parent=30,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=224,.parent=31"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=225,.parent=123"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=234,.parent=29,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=231,.parent=30,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=232,.parent=31"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=233,.parent=123"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=29"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=30"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=31"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=123"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=266,.parent=29"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=263,.parent=30"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=264,.parent=31"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=265,.parent=123"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=106,.parent=29"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=103,.parent=30"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=104,.parent=31"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=105,.parent=123"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=242,.parent=29"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=239,.parent=30"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=240,.parent=31"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=241,.parent=123"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=114,.parent=29"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=111,.parent=30"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=112,.parent=31"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=113,.parent=123"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=274,.parent=29"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=271,.parent=30"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=272,.parent=31"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=273,.parent=123"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=170,.parent=29"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=167,.parent=30"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=168,.parent=31"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=169,.parent=123"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=250,.parent=29"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=247,.parent=30"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=248,.parent=31"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=249,.parent=123"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=178,.parent=29"
         _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=175,.parent=30"
         _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=176,.parent=31"
         _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=177,.parent=123"
         _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=282,.parent=29"
         _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=279,.parent=30"
         _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=280,.parent=31"
         _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=281,.parent=123"
         _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=202,.parent=29"
         _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=199,.parent=30"
         _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=200,.parent=31"
         _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=201,.parent=123"
         _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=210,.parent=29"
         _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=207,.parent=30"
         _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=208,.parent=31"
         _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=209,.parent=123"
         _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=194,.parent=29"
         _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=191,.parent=30"
         _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=192,.parent=31"
         _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=193,.parent=123"
         _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=258,.parent=29"
         _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=255,.parent=30"
         _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=256,.parent=31"
         _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=257,.parent=123"
         _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=142,.parent=29"
         _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=139,.parent=30,.alignment=2"
         _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=140,.parent=31"
         _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=141,.parent=123"
         _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=290,.parent=29"
         _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=287,.parent=30"
         _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=288,.parent=31"
         _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=289,.parent=123"
         _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=218,.parent=29"
         _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=215,.parent=30"
         _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=216,.parent=31"
         _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=217,.parent=123"
         _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=150,.parent=29"
         _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=147,.parent=30"
         _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=148,.parent=31"
         _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=149,.parent=123"
         _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=74,.parent=29"
         _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=71,.parent=30"
         _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=72,.parent=31"
         _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=73,.parent=123"
         _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=46,.parent=29"
         _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=43,.parent=30"
         _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=44,.parent=31"
         _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=45,.parent=123"
         _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=66,.parent=29"
         _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=63,.parent=30"
         _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=64,.parent=31"
         _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=65,.parent=123"
         _StyleDefs(140) =   "Splits(1).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(141) =   "Splits(1).CaptionStyle:id=22,.parent=4"
         _StyleDefs(142) =   "Splits(1).HeadingStyle:id=14,.parent=2"
         _StyleDefs(143) =   "Splits(1).FooterStyle:id=15,.parent=3,.valignment=2,.bold=-1,.fontsize=825"
         _StyleDefs(144) =   ":id=15,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(145) =   ":id=15,.fontname=MS Sans Serif"
         _StyleDefs(146) =   "Splits(1).InactiveStyle:id=16,.parent=5"
         _StyleDefs(147) =   "Splits(1).SelectedStyle:id=18,.parent=6"
         _StyleDefs(148) =   "Splits(1).EditorStyle:id=17,.parent=7"
         _StyleDefs(149) =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(150) =   "Splits(1).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(151) =   "Splits(1).OddRowStyle:id=21,.parent=10"
         _StyleDefs(152) =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(153) =   "Splits(1).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(154) =   "Splits(1).Columns(0).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(155) =   "Splits(1).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
         _StyleDefs(156) =   "Splits(1).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(157) =   "Splits(1).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(158) =   "Splits(1).Columns(1).Style:id=94,.parent=13"
         _StyleDefs(159) =   "Splits(1).Columns(1).HeadingStyle:id=91,.parent=14"
         _StyleDefs(160) =   "Splits(1).Columns(1).FooterStyle:id=92,.parent=15"
         _StyleDefs(161) =   "Splits(1).Columns(1).EditorStyle:id=93,.parent=17"
         _StyleDefs(162) =   "Splits(1).Columns(2).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(163) =   "Splits(1).Columns(2).HeadingStyle:id=51,.parent=14,.alignment=2"
         _StyleDefs(164) =   "Splits(1).Columns(2).FooterStyle:id=52,.parent=15,.namedParent=35,.bold=-1"
         _StyleDefs(165) =   ":id=52,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(166) =   ":id=52,.fontname=MS Sans Serif"
         _StyleDefs(167) =   "Splits(1).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(168) =   "Splits(1).Columns(3).Style:id=230,.parent=13"
         _StyleDefs(169) =   "Splits(1).Columns(3).HeadingStyle:id=227,.parent=14"
         _StyleDefs(170) =   "Splits(1).Columns(3).FooterStyle:id=228,.parent=15"
         _StyleDefs(171) =   "Splits(1).Columns(3).EditorStyle:id=229,.parent=17"
         _StyleDefs(172) =   "Splits(1).Columns(4).Style:id=238,.parent=13,.alignment=1,.bgcolor=&HFFD5AA&"
         _StyleDefs(173) =   "Splits(1).Columns(4).HeadingStyle:id=235,.parent=14,.alignment=1"
         _StyleDefs(174) =   "Splits(1).Columns(4).FooterStyle:id=236,.parent=15,.namedParent=35,.alignment=1"
         _StyleDefs(175) =   ":id=236,.bgcolor=&HFFD5AA&"
         _StyleDefs(176) =   "Splits(1).Columns(4).EditorStyle:id=237,.parent=17"
         _StyleDefs(177) =   "Splits(1).Columns(5).Style:id=86,.parent=13,.alignment=1,.bgcolor=&HFFD5AA&"
         _StyleDefs(178) =   "Splits(1).Columns(5).HeadingStyle:id=83,.parent=14,.alignment=1"
         _StyleDefs(179) =   "Splits(1).Columns(5).FooterStyle:id=84,.parent=15,.namedParent=35,.alignment=1"
         _StyleDefs(180) =   ":id=84,.bgcolor=&HFFD5AA&"
         _StyleDefs(181) =   "Splits(1).Columns(5).EditorStyle:id=85,.parent=17"
         _StyleDefs(182) =   "Splits(1).Columns(6).Style:id=270,.parent=13,.alignment=1,.bgcolor=&HFFD5AA&"
         _StyleDefs(183) =   "Splits(1).Columns(6).HeadingStyle:id=267,.parent=14,.alignment=1"
         _StyleDefs(184) =   "Splits(1).Columns(6).FooterStyle:id=268,.parent=15,.namedParent=35"
         _StyleDefs(185) =   ":id=268,.bgcolor=&HFFD5AA&"
         _StyleDefs(186) =   "Splits(1).Columns(6).EditorStyle:id=269,.parent=17"
         _StyleDefs(187) =   "Splits(1).Columns(7).Style:id=110,.parent=13,.alignment=1,.bgcolor=&HFFD5AA&"
         _StyleDefs(188) =   "Splits(1).Columns(7).HeadingStyle:id=107,.parent=14,.alignment=1"
         _StyleDefs(189) =   "Splits(1).Columns(7).FooterStyle:id=108,.parent=15,.namedParent=35"
         _StyleDefs(190) =   ":id=108,.bgcolor=&HFFD5AA&"
         _StyleDefs(191) =   "Splits(1).Columns(7).EditorStyle:id=109,.parent=17"
         _StyleDefs(192) =   "Splits(1).Columns(8).Style:id=246,.parent=13,.alignment=1,.bgcolor=&HA8C5FF&"
         _StyleDefs(193) =   "Splits(1).Columns(8).HeadingStyle:id=243,.parent=14,.alignment=1"
         _StyleDefs(194) =   "Splits(1).Columns(8).FooterStyle:id=244,.parent=15,.namedParent=35"
         _StyleDefs(195) =   ":id=244,.bgcolor=&HA8C5FF&"
         _StyleDefs(196) =   "Splits(1).Columns(8).EditorStyle:id=245,.parent=17"
         _StyleDefs(197) =   "Splits(1).Columns(9).Style:id=118,.parent=13,.alignment=1,.bgcolor=&HAEC9FF&"
         _StyleDefs(198) =   "Splits(1).Columns(9).HeadingStyle:id=115,.parent=14,.alignment=1"
         _StyleDefs(199) =   "Splits(1).Columns(9).FooterStyle:id=116,.parent=15,.namedParent=35"
         _StyleDefs(200) =   ":id=116,.bgcolor=&HA8C5FF&"
         _StyleDefs(201) =   "Splits(1).Columns(9).EditorStyle:id=117,.parent=17"
         _StyleDefs(202) =   "Splits(1).Columns(10).Style:id=278,.parent=13,.alignment=1,.bgcolor=&HA8C5FF&"
         _StyleDefs(203) =   "Splits(1).Columns(10).HeadingStyle:id=275,.parent=14,.alignment=1"
         _StyleDefs(204) =   "Splits(1).Columns(10).FooterStyle:id=276,.parent=15,.namedParent=35"
         _StyleDefs(205) =   ":id=276,.bgcolor=&HA8C5FF&"
         _StyleDefs(206) =   "Splits(1).Columns(10).EditorStyle:id=277,.parent=17"
         _StyleDefs(207) =   "Splits(1).Columns(11).Style:id=174,.parent=13,.alignment=1,.bgcolor=&HAEC9FF&"
         _StyleDefs(208) =   "Splits(1).Columns(11).HeadingStyle:id=171,.parent=14,.alignment=1"
         _StyleDefs(209) =   "Splits(1).Columns(11).FooterStyle:id=172,.parent=15,.namedParent=35"
         _StyleDefs(210) =   ":id=172,.bgcolor=&HA8C5FF&"
         _StyleDefs(211) =   "Splits(1).Columns(11).EditorStyle:id=173,.parent=17"
         _StyleDefs(212) =   "Splits(1).Columns(12).Style:id=254,.parent=13,.alignment=1,.bgcolor=&HD7FFD7&"
         _StyleDefs(213) =   "Splits(1).Columns(12).HeadingStyle:id=251,.parent=14,.alignment=1"
         _StyleDefs(214) =   "Splits(1).Columns(12).FooterStyle:id=252,.parent=15,.namedParent=35"
         _StyleDefs(215) =   ":id=252,.bgcolor=&HD7FFD7&"
         _StyleDefs(216) =   "Splits(1).Columns(12).EditorStyle:id=253,.parent=17"
         _StyleDefs(217) =   "Splits(1).Columns(13).Style:id=182,.parent=13,.alignment=1,.bgcolor=&HD7FFD7&"
         _StyleDefs(218) =   "Splits(1).Columns(13).HeadingStyle:id=179,.parent=14,.alignment=1"
         _StyleDefs(219) =   "Splits(1).Columns(13).FooterStyle:id=180,.parent=15,.namedParent=35"
         _StyleDefs(220) =   ":id=180,.bgcolor=&HD7FFD7&"
         _StyleDefs(221) =   "Splits(1).Columns(13).EditorStyle:id=181,.parent=17"
         _StyleDefs(222) =   "Splits(1).Columns(14).Style:id=286,.parent=13,.alignment=1,.bgcolor=&HD7FFD7&"
         _StyleDefs(223) =   "Splits(1).Columns(14).HeadingStyle:id=283,.parent=14,.alignment=1"
         _StyleDefs(224) =   "Splits(1).Columns(14).FooterStyle:id=284,.parent=15,.namedParent=35"
         _StyleDefs(225) =   ":id=284,.bgcolor=&HD7FFD7&"
         _StyleDefs(226) =   "Splits(1).Columns(14).EditorStyle:id=285,.parent=17"
         _StyleDefs(227) =   "Splits(1).Columns(15).Style:id=206,.parent=13,.alignment=1,.bgcolor=&HD7FFD7&"
         _StyleDefs(228) =   "Splits(1).Columns(15).HeadingStyle:id=203,.parent=14,.alignment=1"
         _StyleDefs(229) =   "Splits(1).Columns(15).FooterStyle:id=204,.parent=15,.namedParent=35"
         _StyleDefs(230) =   ":id=204,.bgcolor=&HD7FFD7&"
         _StyleDefs(231) =   "Splits(1).Columns(15).EditorStyle:id=205,.parent=17"
         _StyleDefs(232) =   "Splits(1).Columns(16).Style:id=214,.parent=13,.alignment=1,.bgcolor=&HD7FFD7&"
         _StyleDefs(233) =   "Splits(1).Columns(16).HeadingStyle:id=211,.parent=14,.alignment=1"
         _StyleDefs(234) =   "Splits(1).Columns(16).FooterStyle:id=212,.parent=15,.namedParent=35"
         _StyleDefs(235) =   ":id=212,.bgcolor=&HD7FFD7&"
         _StyleDefs(236) =   "Splits(1).Columns(16).EditorStyle:id=213,.parent=17"
         _StyleDefs(237) =   "Splits(1).Columns(17).Style:id=198,.parent=13,.alignment=1,.bgcolor=&HD7FFD7&"
         _StyleDefs(238) =   "Splits(1).Columns(17).HeadingStyle:id=195,.parent=14,.alignment=1"
         _StyleDefs(239) =   "Splits(1).Columns(17).FooterStyle:id=196,.parent=15,.namedParent=35"
         _StyleDefs(240) =   ":id=196,.bgcolor=&HD7FFD7&"
         _StyleDefs(241) =   "Splits(1).Columns(17).EditorStyle:id=197,.parent=17"
         _StyleDefs(242) =   "Splits(1).Columns(18).Style:id=262,.parent=13,.alignment=1,.bgcolor=&HC4FFFF&"
         _StyleDefs(243) =   "Splits(1).Columns(18).HeadingStyle:id=259,.parent=14,.alignment=1"
         _StyleDefs(244) =   "Splits(1).Columns(18).FooterStyle:id=260,.parent=15,.namedParent=35"
         _StyleDefs(245) =   ":id=260,.bgcolor=&HC4FFFF&"
         _StyleDefs(246) =   "Splits(1).Columns(18).EditorStyle:id=261,.parent=17"
         _StyleDefs(247) =   "Splits(1).Columns(19).Style:id=58,.parent=13,.alignment=1,.bgcolor=&HC4FFFF&"
         _StyleDefs(248) =   "Splits(1).Columns(19).HeadingStyle:id=55,.parent=14,.alignment=1"
         _StyleDefs(249) =   "Splits(1).Columns(19).FooterStyle:id=56,.parent=15,.namedParent=35"
         _StyleDefs(250) =   ":id=56,.bgcolor=&HC4FFFF&"
         _StyleDefs(251) =   "Splits(1).Columns(19).EditorStyle:id=57,.parent=17"
         _StyleDefs(252) =   "Splits(1).Columns(20).Style:id=294,.parent=13,.alignment=1,.bgcolor=&HC4FFFF&"
         _StyleDefs(253) =   "Splits(1).Columns(20).HeadingStyle:id=291,.parent=14,.alignment=1"
         _StyleDefs(254) =   "Splits(1).Columns(20).FooterStyle:id=292,.parent=15,.namedParent=35"
         _StyleDefs(255) =   ":id=292,.bgcolor=&HC4FFFF&"
         _StyleDefs(256) =   "Splits(1).Columns(20).EditorStyle:id=293,.parent=17"
         _StyleDefs(257) =   "Splits(1).Columns(21).Style:id=222,.parent=13,.alignment=1,.bgcolor=&HC4FFFF&"
         _StyleDefs(258) =   "Splits(1).Columns(21).HeadingStyle:id=219,.parent=14,.alignment=1"
         _StyleDefs(259) =   "Splits(1).Columns(21).FooterStyle:id=220,.parent=15,.namedParent=35"
         _StyleDefs(260) =   ":id=220,.bgcolor=&HC4FFFF&"
         _StyleDefs(261) =   "Splits(1).Columns(21).EditorStyle:id=221,.parent=17"
         _StyleDefs(262) =   "Splits(1).Columns(22).Style:id=62,.parent=13,.alignment=1,.bgcolor=&HC4FFFF&"
         _StyleDefs(263) =   "Splits(1).Columns(22).HeadingStyle:id=59,.parent=14,.alignment=1"
         _StyleDefs(264) =   "Splits(1).Columns(22).FooterStyle:id=60,.parent=15,.namedParent=35"
         _StyleDefs(265) =   ":id=60,.bgcolor=&HC4FFFF&"
         _StyleDefs(266) =   "Splits(1).Columns(22).EditorStyle:id=61,.parent=17"
         _StyleDefs(267) =   "Splits(1).Columns(23).Style:id=78,.parent=13,.alignment=1"
         _StyleDefs(268) =   "Splits(1).Columns(23).HeadingStyle:id=75,.parent=14"
         _StyleDefs(269) =   "Splits(1).Columns(23).FooterStyle:id=76,.parent=15,.alignment=1"
         _StyleDefs(270) =   "Splits(1).Columns(23).EditorStyle:id=77,.parent=17"
         _StyleDefs(271) =   "Splits(1).Columns(24).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(272) =   "Splits(1).Columns(24).HeadingStyle:id=47,.parent=14"
         _StyleDefs(273) =   "Splits(1).Columns(24).FooterStyle:id=48,.parent=15"
         _StyleDefs(274) =   "Splits(1).Columns(24).EditorStyle:id=49,.parent=17"
         _StyleDefs(275) =   "Splits(1).Columns(25).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(276) =   "Splits(1).Columns(25).HeadingStyle:id=67,.parent=14"
         _StyleDefs(277) =   "Splits(1).Columns(25).FooterStyle:id=68,.parent=15"
         _StyleDefs(278) =   "Splits(1).Columns(25).EditorStyle:id=69,.parent=17"
         _StyleDefs(279) =   "Named:id=33:Normal"
         _StyleDefs(280) =   ":id=33,.parent=0"
         _StyleDefs(281) =   "Named:id=34:Heading"
         _StyleDefs(282) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(283) =   ":id=34,.wraptext=-1"
         _StyleDefs(284) =   "Named:id=35:Footing"
         _StyleDefs(285) =   ":id=35,.parent=33,.alignment=1,.valignment=2,.bgcolor=&H80000005&"
         _StyleDefs(286) =   ":id=35,.fgcolor=&H80000012&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(287) =   ":id=35,.strikethrough=0,.charset=0"
         _StyleDefs(288) =   ":id=35,.fontname=MS Sans Serif"
         _StyleDefs(289) =   "Named:id=36:Selected"
         _StyleDefs(290) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(291) =   "Named:id=37:Caption"
         _StyleDefs(292) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(293) =   "Named:id=38:HighlightRow"
         _StyleDefs(294) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(295) =   "Named:id=39:EvenRow"
         _StyleDefs(296) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(297) =   "Named:id=40:OddRow"
         _StyleDefs(298) =   ":id=40,.parent=33"
         _StyleDefs(299) =   "Named:id=41:RecordSelector"
         _StyleDefs(300) =   ":id=41,.parent=34"
         _StyleDefs(301) =   "Named:id=42:FilterBar"
         _StyleDefs(302) =   ":id=42,.parent=33"
      End
      Begin MSComCtl2.DTPicker dtpCompetenciaPISCOFINS 
         Height          =   315
         Left            =   -64560
         TabIndex        =   59
         Top             =   840
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "MM/yyyy"
         DateIsNull      =   -1  'True
         Format          =   146669571
         CurrentDate     =   37460
      End
      Begin MSComCtl2.DTPicker dtpCompetenciaISS 
         Height          =   315
         Left            =   -74760
         TabIndex        =   62
         Top             =   840
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "MM/yyyy"
         DateIsNull      =   -1  'True
         Format          =   146669571
         CurrentDate     =   37460
      End
      Begin TrueOleDBGrid70.TDBGrid tdbgISS 
         Bindings        =   "frmProvisaoImpostos.frx":060F
         Height          =   5175
         Left            =   -74760
         TabIndex        =   63
         Top             =   1440
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   9128
         _LayoutType     =   4
         _RowHeight      =   25
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Centro de Resultado"
         Columns(0).DataField=   "CRDescricao"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Estruturado"
         Columns(1).DataField=   "CRCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   16
         Columns(2)._MaxComboItems=   5
         Columns(2).ValueItems(0)._DefaultItem=   0
         Columns(2).ValueItems(0).Value=   "1"
         Columns(2).ValueItems(0).Value.vt=   8
         Columns(2).ValueItems(0).DisplayValue=   "Simples"
         Columns(2).ValueItems(0).DisplayValue.vt=   8
         Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(1)._DefaultItem=   0
         Columns(2).ValueItems(1).Value=   "2"
         Columns(2).ValueItems(1).Value.vt=   8
         Columns(2).ValueItems(1).DisplayValue=   "Simples"
         Columns(2).ValueItems(1).DisplayValue.vt=   8
         Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(2)._DefaultItem=   0
         Columns(2).ValueItems(2).Value=   "3"
         Columns(2).ValueItems(2).Value.vt=   8
         Columns(2).ValueItems(2).DisplayValue=   "Composta"
         Columns(2).ValueItems(2).DisplayValue.vt=   8
         Columns(2).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems.Count=   3
         Columns(2).Caption=   "Faturamento"
         Columns(2).DataField=   "ValorFaturado"
         Columns(2).NumberFormat=   "Standard"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Valor Retido ISS"
         Columns(3).DataField=   "ValorRetidoISS"
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Retenção ISS"
         Columns(4).DataField=   "RetencaoISS"
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Provisão ISS"
         Columns(5).DataField=   "ValorProvisaoISS"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).Size  =   3
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   0
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=7938"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=7858"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=113"
         Splits(0)._ColumnProps(10)=   "Column(1).Width=1323"
         Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=1244"
         Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=20"
         Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(17)=   "Column(2).Width=2461"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2381"
         Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=528"
         Splits(0)._ColumnProps(22)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(23)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(24)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(25)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(26)=   "Column(3).Width=2461"
         Splits(0)._ColumnProps(27)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(3)._WidthInPix=2381"
         Splits(0)._ColumnProps(29)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(3)._ColStyle=532"
         Splits(0)._ColumnProps(31)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(32)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(33)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(34)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(35)=   "Column(3)._MinWidth=29"
         Splits(0)._ColumnProps(36)=   "Column(4).Width=2461"
         Splits(0)._ColumnProps(37)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(4)._WidthInPix=2381"
         Splits(0)._ColumnProps(39)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(4)._ColStyle=528"
         Splits(0)._ColumnProps(41)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(42)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(43)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(44)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(45)=   "Column(5).Width=2461"
         Splits(0)._ColumnProps(46)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(5)._WidthInPix=2381"
         Splits(0)._ColumnProps(48)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(49)=   "Column(5)._ColStyle=20"
         Splits(0)._ColumnProps(50)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(51)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(52)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(53)=   "Column(5)._MinWidth=123792704"
         Splits(1)._UserFlags=   0
         Splits(1).ExtendRightColumn=   -1  'True
         Splits(1).MarqueeStyle=   3
         Splits(1).Size  =   3
         Splits(1).Size.vt=   2
         Splits(1).RecordSelectors=   0   'False
         Splits(1).RecordSelectorWidth=   503
         Splits(1)._SavedRecordSelectors=   0   'False
         Splits(1).AllowColMove=   -1  'True
         Splits(1).DividerStyle=   2
         Splits(1).DividerColor=   12632256
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=6"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=4419"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=4339"
         Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=528"
         Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(7)=   "Column(0).WrapText=1"
         Splits(1)._ColumnProps(8)=   "Column(0).AllowFocus=0"
         Splits(1)._ColumnProps(9)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(10)=   "Column(1).Width=2725"
         Splits(1)._ColumnProps(11)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(12)=   "Column(1)._WidthInPix=2646"
         Splits(1)._ColumnProps(13)=   "Column(1)._EditAlways=0"
         Splits(1)._ColumnProps(14)=   "Column(1)._ColStyle=20"
         Splits(1)._ColumnProps(15)=   "Column(1).Visible=0"
         Splits(1)._ColumnProps(16)=   "Column(1).WrapText=1"
         Splits(1)._ColumnProps(17)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(18)=   "Column(1)._MinWidth=104849504"
         Splits(1)._ColumnProps(19)=   "Column(2).Width=2461"
         Splits(1)._ColumnProps(20)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(21)=   "Column(2)._WidthInPix=2381"
         Splits(1)._ColumnProps(22)=   "Column(2)._EditAlways=0"
         Splits(1)._ColumnProps(23)=   "Column(2)._ColStyle=530"
         Splits(1)._ColumnProps(24)=   "Column(2).WrapText=1"
         Splits(1)._ColumnProps(25)=   "Column(2).AllowFocus=0"
         Splits(1)._ColumnProps(26)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(27)=   "Column(2)._MinWidth=104849504"
         Splits(1)._ColumnProps(28)=   "Column(3).Width=2461"
         Splits(1)._ColumnProps(29)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(30)=   "Column(3)._WidthInPix=2381"
         Splits(1)._ColumnProps(31)=   "Column(3)._EditAlways=0"
         Splits(1)._ColumnProps(32)=   "Column(3)._ColStyle=530"
         Splits(1)._ColumnProps(33)=   "Column(3).WrapText=1"
         Splits(1)._ColumnProps(34)=   "Column(3).AllowFocus=0"
         Splits(1)._ColumnProps(35)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(36)=   "Column(4).Width=2461"
         Splits(1)._ColumnProps(37)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(38)=   "Column(4)._WidthInPix=2381"
         Splits(1)._ColumnProps(39)=   "Column(4)._EditAlways=0"
         Splits(1)._ColumnProps(40)=   "Column(4)._ColStyle=530"
         Splits(1)._ColumnProps(41)=   "Column(4).WrapText=1"
         Splits(1)._ColumnProps(42)=   "Column(4).AllowFocus=0"
         Splits(1)._ColumnProps(43)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(44)=   "Column(5).Width=2461"
         Splits(1)._ColumnProps(45)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(46)=   "Column(5)._WidthInPix=2381"
         Splits(1)._ColumnProps(47)=   "Column(5)._EditAlways=0"
         Splits(1)._ColumnProps(48)=   "Column(5)._ColStyle=18"
         Splits(1)._ColumnProps(49)=   "Column(5).WrapText=1"
         Splits(1)._ColumnProps(50)=   "Column(5).Order=6"
         Splits.Count    =   2
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         ColumnFooters   =   -1  'True
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   2
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
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
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=42,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(24)  =   "Splits(0).Style:id=29,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=128,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=30,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=31,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=32,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=124,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=123,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=125,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=126,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=127,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=129,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=130,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=134,.parent=29,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=131,.parent=30,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=132,.parent=31"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=133,.parent=123"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=90,.parent=29"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=87,.parent=30"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=88,.parent=31,.namedParent=35"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=89,.parent=123"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=138,.parent=29,.alignment=0"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=135,.parent=30,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=136,.parent=31"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=137,.parent=123"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=142,.parent=29"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=139,.parent=30,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=140,.parent=31"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=141,.parent=123"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=146,.parent=29,.alignment=0"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=143,.parent=30,.alignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=144,.parent=31"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=145,.parent=123"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=150,.parent=29"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=147,.parent=30"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=148,.parent=31"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=149,.parent=123"
         _StyleDefs(60)  =   "Splits(1).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(61)  =   "Splits(1).CaptionStyle:id=22,.parent=4"
         _StyleDefs(62)  =   "Splits(1).HeadingStyle:id=14,.parent=2"
         _StyleDefs(63)  =   "Splits(1).FooterStyle:id=15,.parent=3"
         _StyleDefs(64)  =   "Splits(1).InactiveStyle:id=16,.parent=5"
         _StyleDefs(65)  =   "Splits(1).SelectedStyle:id=18,.parent=6"
         _StyleDefs(66)  =   "Splits(1).EditorStyle:id=17,.parent=7"
         _StyleDefs(67)  =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(68)  =   "Splits(1).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(69)  =   "Splits(1).OddRowStyle:id=21,.parent=10"
         _StyleDefs(70)  =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(71)  =   "Splits(1).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(72)  =   "Splits(1).Columns(0).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(73)  =   "Splits(1).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
         _StyleDefs(74)  =   "Splits(1).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(75)  =   "Splits(1).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(76)  =   "Splits(1).Columns(1).Style:id=94,.parent=13"
         _StyleDefs(77)  =   "Splits(1).Columns(1).HeadingStyle:id=91,.parent=14"
         _StyleDefs(78)  =   "Splits(1).Columns(1).FooterStyle:id=92,.parent=15"
         _StyleDefs(79)  =   "Splits(1).Columns(1).EditorStyle:id=93,.parent=17"
         _StyleDefs(80)  =   "Splits(1).Columns(2).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(81)  =   "Splits(1).Columns(2).HeadingStyle:id=51,.parent=14,.alignment=2"
         _StyleDefs(82)  =   "Splits(1).Columns(2).FooterStyle:id=52,.parent=15,.namedParent=35"
         _StyleDefs(83)  =   "Splits(1).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(84)  =   "Splits(1).Columns(3).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(85)  =   "Splits(1).Columns(3).HeadingStyle:id=55,.parent=14,.alignment=2"
         _StyleDefs(86)  =   "Splits(1).Columns(3).FooterStyle:id=56,.parent=15,.namedParent=35"
         _StyleDefs(87)  =   "Splits(1).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(88)  =   "Splits(1).Columns(4).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(89)  =   "Splits(1).Columns(4).HeadingStyle:id=43,.parent=14,.alignment=2"
         _StyleDefs(90)  =   "Splits(1).Columns(4).FooterStyle:id=44,.parent=15,.namedParent=35"
         _StyleDefs(91)  =   "Splits(1).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(92)  =   "Splits(1).Columns(5).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(93)  =   "Splits(1).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(94)  =   "Splits(1).Columns(5).FooterStyle:id=60,.parent=15,.namedParent=35"
         _StyleDefs(95)  =   "Splits(1).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(96)  =   "Named:id=33:Normal"
         _StyleDefs(97)  =   ":id=33,.parent=0"
         _StyleDefs(98)  =   "Named:id=34:Heading"
         _StyleDefs(99)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(100) =   ":id=34,.wraptext=-1"
         _StyleDefs(101) =   "Named:id=35:Footing"
         _StyleDefs(102) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H80000005&,.fgcolor=&H80000012&"
         _StyleDefs(103) =   ":id=35,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(104) =   ":id=35,.fontname=MS Sans Serif"
         _StyleDefs(105) =   "Named:id=36:Selected"
         _StyleDefs(106) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(107) =   "Named:id=37:Caption"
         _StyleDefs(108) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(109) =   "Named:id=38:HighlightRow"
         _StyleDefs(110) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(111) =   "Named:id=39:EvenRow"
         _StyleDefs(112) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(113) =   "Named:id=40:OddRow"
         _StyleDefs(114) =   ":id=40,.parent=33"
         _StyleDefs(115) =   "Named:id=41:RecordSelector"
         _StyleDefs(116) =   ":id=41,.parent=34"
         _StyleDefs(117) =   "Named:id=42:FilterBar"
         _StyleDefs(118) =   ":id=42,.parent=33"
      End
      Begin MSAdodcLib.Adodc adoIRPJCSLLCOFINSPIS 
         Height          =   375
         Left            =   -75000
         Top             =   6240
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
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
         Caption         =   ""
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
      Begin MSAdodcLib.Adodc adoISS 
         Height          =   375
         Left            =   -75000
         Top             =   6240
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
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
         Caption         =   ""
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
      Begin MSAdodcLib.Adodc adoISSPagamento 
         Height          =   375
         Left            =   -75000
         Top             =   5760
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
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
         Caption         =   ""
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
      Begin VB.Label Label 
         Caption         =   "Competência:"
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   64
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Competência:"
         Height          =   255
         Index           =   4
         Left            =   -64560
         TabIndex        =   60
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   7560
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   9000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTotalProvisaoCPP 
      Alignment       =   1  'Right Justify
      Caption         =   "0,00"
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
      Left            =   10200
      TabIndex        =   90
      Top             =   9240
      Width           =   2055
   End
   Begin VB.Label lblTituloTotalProvisaoCPP 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Provisão CPP:"
      Height          =   255
      Left            =   8400
      TabIndex        =   89
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Line Line 
      X1              =   8160
      X2              =   12480
      Y1              =   9540
      Y2              =   9540
   End
   Begin VB.Label lblTotalProvisaoISS 
      Alignment       =   1  'Right Justify
      Caption         =   "0,00"
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
      Left            =   10200
      TabIndex        =   13
      Top             =   8880
      Width           =   2055
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Provisão ISS:"
      Height          =   255
      Left            =   8640
      TabIndex        =   12
      Top             =   8880
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Provisão CSLL:"
      Height          =   255
      Left            =   8640
      TabIndex        =   11
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label lblTituloTotalProvisaoIRRF 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Provisão IRRF:"
      Height          =   255
      Left            =   8640
      TabIndex        =   10
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label lblTotalProvisaoCSLL 
      Alignment       =   1  'Right Justify
      Caption         =   "0,00"
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
      Left            =   10200
      TabIndex        =   9
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lblTotalProvisaoIRRF 
      Alignment       =   1  'Right Justify
      Caption         =   "0,00"
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
      Left            =   10200
      TabIndex        =   8
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Provisão COFINS:"
      Height          =   255
      Left            =   8520
      TabIndex        =   7
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Provisão PIS:"
      Height          =   255
      Left            =   8640
      TabIndex        =   6
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label lblTotalProvisaoCOFINS 
      Alignment       =   1  'Right Justify
      Caption         =   "0,00"
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
      Left            =   10200
      TabIndex        =   5
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Label lblTotalProvisaoPIS 
      Alignment       =   1  'Right Justify
      Caption         =   "0,00"
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
      Left            =   10200
      TabIndex        =   4
      Top             =   8520
      Width           =   2055
   End
   Begin VB.Label lblTotalGeralProvisao 
      Alignment       =   1  'Right Justify
      Caption         =   "0,00"
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
      Left            =   10200
      TabIndex        =   3
      Top             =   9720
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Geral:"
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
      Left            =   8640
      TabIndex        =   2
      Top             =   9720
      Width           =   1575
   End
End
Attribute VB_Name = "frmProvisaoImpostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const CENTRORESULTADO As Integer = 0
Const ESTRUTURADO As Integer = 1
Const VALORFATURADO As Integer = 2
Const PERCENTUAL_LUCRO As Integer = 3

Const DEVIDO_IRPJ As Integer = 4
Const RETIDO_IRPJ As Integer = 5
Const PERCENTUAL_RETIDO_IRPJ As Integer = 6
Const PROVISAO_IRPJ As Integer = 7

Const DEVIDO_CSLL As Integer = 8
Const RETIDO_CSLL As Integer = 9
Const PERCENTUAL_RETIDO_CSLL As Integer = 10
Const PROVISAO_CSLL As Integer = 11

Const DEVIDO_COFINS As Integer = 12
Const RETIDO_COFINS As Integer = 13
Const PERCENTUAL_RETIDO_COFINS As Integer = 14
Const TRANSF_IRPJ_COFINS As Integer = 15
Const TRANSF_CSLL_COFINS As Integer = 16
Const PROVISAO_COFINS As Integer = 17

Const DEVIDO_PIS As Integer = 18
Const RETIDO_PIS As Integer = 19
Const PERCENTUAL_RETIDO_PIS As Integer = 20
Const TRANSF_COFINS_PIS As Integer = 21
Const PROVISAO_PIS As Integer = 22
Const TRANSF_PIS_IRPF As Integer = 23
Const PERCENTUAL_RETIDO_CPP As Integer = 24
Const PROVISAO_CPP As Integer = 25

Const GRIDISS_CENTRORESULTADO As Integer = 0
Const GRIDISS_ESTRUTURADO As Integer = 1
Const GRIDISS_FATURAMENTO As Integer = 2
Const GRIDISS_VALORRETIDOISS As Integer = 3
Const GRIDISS_RETENCAOISS As Integer = 4
Const GRIDISS_PROVISAOISS As Integer = 5

Dim XFB_SIMPLESNACIONAL As Boolean

Private Sub PreencherDadosISS()
    Dim competencia As String
    If (cmbRegime.Text = "Caixa") Then
        competencia = "0"
    Else
        competencia = "1"
    End If

    subCarregaDadosNV adoISS, "", "", "", "EXEC sp_ProvisaoISS @empresa = " & CStr(PCodEmpresa) & ", @emissao = " & FunNuloData("01/" & Month(dtpCompetenciaISS.Value) & "/" & Year(dtpCompetenciaISS.Value), "Sql Server") & ", @competencia = " & competencia
    subCarregaDadosNV adoISSPagamento, "", "", "", "EXEC sp_ProvisaoISS @empresa = " & CStr(PCodEmpresa) & ", @emissao = " & FunNuloData("01/" & Month(dtpCompetenciaISS.Value) & "/" & Year(dtpCompetenciaISS.Value), "Sql Server") & ", @competencia = " & competencia
    
    Dim XLD_FATURAMENTO As Double
    Dim XLD_VALORRETIDOISS As Double
    Dim XLD_RETENCAOISS As Double
    Dim XLD_PROVISAOISS As Double
    
    XLD_FATURAMENTO = 0
    
    XLD_VALORRETIDOISS = 0
    XLD_RETENCAOISS = 0
    XLD_PROVISAOISS = 0
    
    If (Not adoISS.Recordset.EOF And Not adoISS.Recordset.BOF) Then
        adoISS.Recordset.MoveFirst
    End If
    
    While Not adoISS.Recordset.EOF And Not adoISS.Recordset.BOF
        XLD_FATURAMENTO = XLD_FATURAMENTO + CDbl(adoISS.Recordset("ValorFaturado"))
        XLD_VALORRETIDOISS = XLD_VALORRETIDOISS + CDbl(adoISS.Recordset("ValorRetidoISS"))
        XLD_RETENCAOISS = XLD_RETENCAOISS + CDbl(adoISS.Recordset("RetencaoISS"))
        XLD_PROVISAOISS = XLD_PROVISAOISS + CDbl(adoISS.Recordset("ValorProvisaoISS"))
        
        adoISS.Recordset.MoveNext
    Wend
    
    tdbgISS.Columns(GRIDISS_ESTRUTURADO).FooterText = "TOTAIS"
    tdbgISS.Columns(GRIDISS_FATURAMENTO).FooterText = CStr(FormatNumber(XLD_FATURAMENTO, 2, True, True, True))
    tdbgISS.Columns(GRIDISS_VALORRETIDOISS).FooterText = CStr(FormatNumber(XLD_VALORRETIDOISS, 2, True, True, True))
    tdbgISS.Columns(GRIDISS_PROVISAOISS).FooterText = CStr(FormatNumber(XLD_PROVISAOISS, 2, True, True, True))
End Sub

Private Sub PreencherDadosIR()

    Dim competencia As String
    If (cmbRegime.Text = "Caixa") Then
        competencia = "0"
    Else
        competencia = "1"
    End If
 
    subCarregaDadosNV adoIRPJCSLLCOFINSPIS, "", "", "", "EXEC sp_ProvisaoIRPJ_CSLL_COFINS_PIS @empresa = " & CStr(PCodEmpresa) & ", @emissao = " & FunNuloData("01/" & Month(dtpCompetenciaPISCOFINS.Value) & "/" & Year(dtpCompetenciaPISCOFINS.Value), "Sql Server") & ", @aliquotaIrpj  = " & Replace(IIf(RTrim(LTrim(txtAliquotaIRPJ.Text)) = "", "0.00", txtAliquotaIRPJ.Text), ",", ".") & ", @aliquotaCSLL  = " & Replace(IIf(RTrim(LTrim(txtAliquotaCSLL.Text)) = "", "0.00", txtAliquotaCSLL.Text), ",", ".") & ", @aliquotaCPP  = " & Replace(IIf(RTrim(LTrim(txtAliquotaCPP.Text)) = "", "0.00", txtAliquotaCPP.Text), ",", ".") & ", @competencia = " & competencia

    Dim XLD_FATURAMENTO As Double
    
    Dim XLD_DEVIDO_IRPJ As Double
    Dim XLD_VALORRETIDO_IRPJ As Double
    Dim XLD_PROVISAO_IRPJ As Double
    
    Dim XLD_DEVIDO_CSLL As Double
    Dim XLD_VALORRETIDO_CSLL As Double
    Dim XLD_PROVISAO_CSLL As Double
    
    Dim XLD_DEVIDO_COFINS As Double
    Dim XLD_TRANSF_IRPJ_COFINS As Double
    Dim XLD_TRANSF_CSLL_COFINS As Double
    Dim XLD_VALORRETIDO_COFINS As Double
    Dim XLD_PROVISAO_COFINS As Double
    
    Dim XLD_DEVIDO_PIS As Double
    Dim XLD_VALORRETIDO_PIS As Double
    Dim XLD_TRANSF_COFINS_PIS As Double
    Dim XLD_PROVISAO_PIS As Double
    Dim XLD_TRANSF_PIS_IRPF As Double
    Dim XLD_PROVISAO_CPP As Double
    
    Dim XLD_PROVISAOISS As Double
            
    XLD_FATURAMENTO = 0

    XLD_DEVIDO_IRPJ = 0
    XLD_VALORRETIDO_IRPJ = 0
    XLD_PROVISAO_IRPJ = 0
    
    XLD_DEVIDO_CSLL = 0
    XLD_VALORRETIDO_CSLL = 0
    XLD_PROVISAO_CSLL = 0
    
    XLD_DEVIDO_COFINS = 0
    XLD_TRANSF_IRPJ_COFINS = 0
    XLD_TRANSF_CSLL_COFINS = 0
    XLD_VALORRETIDO_COFINS = 0
    XLD_PROVISAO_COFINS = 0
    
    XLD_DEVIDO_PIS = 0
    XLD_VALORRETIDO_PIS = 0
    XLD_TRANSF_COFINS_PIS = 0
    XLD_PROVISAO_PIS = 0
    XLD_TRANSF_PIS_IRPF = 0
    XLD_PROVISAO_CPP = 0
    
    XLD_PROVISAOISS = 0
    
    If Not adoIRPJCSLLCOFINSPIS.Recordset.EOF And Not adoIRPJCSLLCOFINSPIS.Recordset.BOF Then
        adoIRPJCSLLCOFINSPIS.Recordset.MoveFirst
    End If
    While Not adoIRPJCSLLCOFINSPIS.Recordset.EOF And Not adoIRPJCSLLCOFINSPIS.Recordset.BOF
        XLD_FATURAMENTO = XLD_FATURAMENTO + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorFaturado")))

        XLD_DEVIDO_IRPJ = XLD_DEVIDO_IRPJ + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorDevidoIRRF")))
        XLD_VALORRETIDO_IRPJ = XLD_VALORRETIDO_IRPJ + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorRetidoIRRF")))
        XLD_PROVISAO_IRPJ = XLD_PROVISAO_IRPJ + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorProvisaoIRRF")))
        
        XLD_DEVIDO_CSLL = XLD_DEVIDO_CSLL + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorDevidoCSLL")))
        XLD_VALORRETIDO_CSLL = XLD_VALORRETIDO_CSLL + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorRetidoCSLL")))
        XLD_PROVISAO_CSLL = XLD_PROVISAO_CSLL + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorProvisaoCSLL")))
        
        XLD_DEVIDO_COFINS = XLD_DEVIDO_COFINS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorDevidoCOFINS")))
        XLD_TRANSF_IRPJ_COFINS = XLD_TRANSF_IRPJ_COFINS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("TransfIRPJCOFINS")))
        XLD_TRANSF_CSLL_COFINS = XLD_TRANSF_CSLL_COFINS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("TransfCSLLCOFINS")))
        XLD_VALORRETIDO_COFINS = XLD_VALORRETIDO_COFINS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorRetidoCOFINS")))
        XLD_PROVISAO_COFINS = XLD_PROVISAO_COFINS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorProvisaoCOFINS")))
        
        XLD_DEVIDO_PIS = XLD_DEVIDO_PIS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorDevidoPIS")))
        XLD_TRANSF_COFINS_PIS = XLD_TRANSF_COFINS_PIS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("TransfCOFINSPIS")))
        XLD_VALORRETIDO_PIS = XLD_VALORRETIDO_PIS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorRetidoPIS")))
        XLD_PROVISAO_PIS = XLD_PROVISAO_PIS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorProvisaoPIS")))
        XLD_TRANSF_PIS_IRPF = XLD_TRANSF_PIS_IRPF + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("TransPISIRPF")))
        XLD_PROVISAO_CPP = XLD_PROVISAO_CPP + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorProvisaoCPP")))
        
        adoIRPJCSLLCOFINSPIS.Recordset.MoveNext
    Wend

    tdbgIRPJCSLLCOFINSPIS.Columns(ESTRUTURADO).FooterText = "TOTAIS"
    tdbgIRPJCSLLCOFINSPIS.Columns(VALORFATURADO).FooterText = CStr(FormatNumber(XLD_FATURAMENTO, 2, True, True, True))
    
    tdbgIRPJCSLLCOFINSPIS.Columns(DEVIDO_IRPJ).FooterText = CStr(FormatNumber(XLD_DEVIDO_IRPJ, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(RETIDO_IRPJ).FooterText = CStr(FormatNumber(XLD_VALORRETIDO_IRPJ, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(PROVISAO_IRPJ).FooterText = CStr(FormatNumber(XLD_PROVISAO_IRPJ, 2, True, True, True))
    
    tdbgIRPJCSLLCOFINSPIS.Columns(DEVIDO_CSLL).FooterText = CStr(FormatNumber(XLD_DEVIDO_CSLL, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(RETIDO_CSLL).FooterText = CStr(FormatNumber(XLD_VALORRETIDO_CSLL, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(PROVISAO_CSLL).FooterText = CStr(FormatNumber(XLD_PROVISAO_CSLL, 2, True, True, True))
    
    tdbgIRPJCSLLCOFINSPIS.Columns(DEVIDO_COFINS).FooterText = CStr(FormatNumber(XLD_DEVIDO_COFINS, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(RETIDO_COFINS).FooterText = CStr(FormatNumber(XLD_VALORRETIDO_COFINS, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(TRANSF_IRPJ_COFINS).FooterText = CStr(FormatNumber(XLD_TRANSF_IRPJ_COFINS, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(TRANSF_CSLL_COFINS).FooterText = CStr(FormatNumber(XLD_TRANSF_CSLL_COFINS, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(PROVISAO_COFINS).FooterText = CStr(FormatNumber(XLD_PROVISAO_COFINS, 2, True, True, True))
    
    tdbgIRPJCSLLCOFINSPIS.Columns(DEVIDO_PIS).FooterText = CStr(FormatNumber(XLD_DEVIDO_PIS, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(RETIDO_PIS).FooterText = CStr(FormatNumber(XLD_VALORRETIDO_PIS, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(TRANSF_COFINS_PIS).FooterText = CStr(FormatNumber(XLD_TRANSF_COFINS_PIS, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(PROVISAO_PIS).FooterText = CStr(FormatNumber(XLD_PROVISAO_PIS, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(TRANSF_PIS_IRPF).FooterText = CStr(FormatNumber(XLD_TRANSF_PIS_IRPF, 2, True, True, True))
    tdbgIRPJCSLLCOFINSPIS.Columns(PROVISAO_CPP).FooterText = CStr(FormatNumber(XLD_PROVISAO_CPP, 2, True, True, True))
    
    lblTotalProvisaoIRRF.Caption = CStr(FormatNumber(XLD_PROVISAO_IRPJ, 2, True, True, True))
    lblTotalProvisaoCSLL.Caption = CStr(FormatNumber(XLD_PROVISAO_CSLL, 2, True, True, True))
    lblTotalProvisaoCOFINS.Caption = CStr(FormatNumber(XLD_PROVISAO_COFINS, 2, True, True, True))
    lblTotalProvisaoPIS.Caption = CStr(FormatNumber(XLD_PROVISAO_PIS, 2, True, True, True))
    lblTotalProvisaoCPP.Caption = CStr(FormatNumber(XLD_PROVISAO_CPP, 2, True, True, True))
    
    If (XLD_PROVISAO_IRPJ >= 0) Then
        lblTotalProvisaoIRRF.ForeColor = &HFF0000
    Else
        lblTotalProvisaoIRRF.ForeColor = &HFF&
    End If
        
    If (XLD_PROVISAO_CSLL >= 0) Then
        lblTotalProvisaoCSLL.ForeColor = &HFF0000
    Else
        lblTotalProvisaoCSLL.ForeColor = &HFF&
    End If
        
    If (XLD_PROVISAO_COFINS >= 0) Then
        lblTotalProvisaoCOFINS.ForeColor = &HFF0000
    Else
        lblTotalProvisaoCOFINS.ForeColor = &HFF&
    End If
    
    If (XLD_PROVISAO_PIS >= 0) Then
        lblTotalProvisaoPIS.ForeColor = &HFF0000
    Else
        lblTotalProvisaoPIS.ForeColor = &HFF&
    End If
    
    If (XLD_PROVISAO_CPP >= 0) Then
        lblTotalProvisaoCPP.ForeColor = &HFF0000
    Else
        lblTotalProvisaoCPP.ForeColor = &HFF&
    End If
    
    If (tdbgISS.Columns(GRIDISS_PROVISAOISS).FooterText <> "") Then
        XLD_PROVISAOISS = CDbl(tdbgISS.Columns(GRIDISS_PROVISAOISS).FooterText)
    Else
        XLD_PROVISAOISS = 0
    End If
    
    lblTotalProvisaoISS.Caption = CStr(FormatNumber(XLD_PROVISAOISS, 2, True, True, True))
    If (XLD_PROVISAOISS >= 0) Then
        lblTotalProvisaoISS.ForeColor = &HFF0000
    Else
        lblTotalProvisaoISS.ForeColor = &HFF&
    End If
    
    lblTotalGeralProvisao.Caption = CStr(FormatNumber(XLD_PROVISAO_IRPJ + XLD_PROVISAO_CSLL + XLD_PROVISAO_COFINS + XLD_PROVISAO_PIS + XLD_PROVISAOISS, 2, True, True, True))
    
    If (XLD_PROVISAO_IRPJ + XLD_PROVISAO_CSLL + XLD_PROVISAO_COFINS + XLD_PROVISAO_PIS + XLD_PROVISAOISS >= 0) Then
        lblTotalGeralProvisao.ForeColor = &HFF0000
    Else
        lblTotalGeralProvisao.ForeColor = &HFF&
    End If
End Sub

Private Sub chkBaixaImpostos_Click()
    If chkBaixaImpostos.Value = 1 Then
        dtpBaixaCOFINS.Enabled = True
        dtpBaixaPIS.Enabled = True
        dtpBaixaISS.Enabled = True
        
    Else
        dtpBaixaCOFINS.Enabled = False
        dtpBaixaPIS.Enabled = False
        dtpBaixaISS.Enabled = False
    End If
End Sub

Private Sub cmdExportarISS_Click()
 Dim sLine As String
    Dim tmpFile As String
    Dim outFile As String
    Dim i As Integer
    
    ' Sets the Dialog Title to Save File
    CommonDialog1.DialogTitle = "Salvar arquivo"

    ' Sets the File List box to Text File and All Files
    CommonDialog1.Filter = "Arquivo CSV (*.csv)|*.csv"

    ' Set the default files type to Text File
    CommonDialog1.FilterIndex = 1

    ' Sets the flags - Hide Read only, prompt to overwrite, and path must exist
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist

    ' Set dialog box so an error occurs if the dialogbox is cancelled
    CommonDialog1.CancelError = True

    ' Enables error handling to catch cancel error
    On Error Resume Next
    ' display the dialog box
    CommonDialog1.ShowSave
    If Err Then
        ' This code runs if the dialog was cancelled
        Exit Sub
    End If
    
    'Call your export function
    outFile = CommonDialog1.FileName
    tmpFile = outFile & ".tmp"
    
    tdbgISS.ExportToDelimitedFile tmpFile, , ";"
 
    Close tmpFile
    
    'Now use this to add the Headers (to a new file)
    Open tmpFile For Input As #1
    Open outFile For Output As #2
 
    'Write Headers
    For i = 0 To tdbgISS.Columns.Count - 1
         If i = 0 Then
             sLine = Chr(34) & tdbgISS.Columns(i).Caption & Chr(34)
         Else
             sLine = sLine & ";" & Chr(34) & tdbgISS.Columns(i).Caption & Chr(34)
         End If
    Next
 
    Print #2, sLine

     'Write contents of original file
    Do Until EOF(1)
         Line Input #1, sLine
         Print #2, sLine
    Loop
 
    Close #1
    Close #2
 
    Kill tmpFile
          
    Screen.MousePointer = vbDefault
    
    MsgBox "Total exportado: " & adoISS.Recordset.RecordCount & " registros", vbInformation, "Informação"
    
End Sub

Private Sub cmdExportIRPJCSLLCOFINSPIS_Click()
    Dim sLine As String
    Dim tmpFile As String
    Dim outFile As String
    Dim i As Integer
    
    ' Sets the Dialog Title to Save File
    CommonDialog1.DialogTitle = "Salvar arquivo"

    ' Sets the File List box to Text File and All Files
    CommonDialog1.Filter = "Arquivo CSV (*.csv)|*.csv"

    ' Set the default files type to Text File
    CommonDialog1.FilterIndex = 1

    ' Sets the flags - Hide Read only, prompt to overwrite, and path must exist
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist

    ' Set dialog box so an error occurs if the dialogbox is cancelled
    CommonDialog1.CancelError = True

    ' Enables error handling to catch cancel error
    On Error Resume Next
    ' display the dialog box
    CommonDialog1.ShowSave
    If Err Then
        ' This code runs if the dialog was cancelled
        Exit Sub
    End If
    
    'Call your export function
    outFile = CommonDialog1.FileName
    tmpFile = outFile & ".tmp"
    
    tdbgIRPJCSLLCOFINSPIS.ExportToDelimitedFile tmpFile, , ";"
 
    Close tmpFile
    
    'Now use this to add the Headers (to a new file)
    Open tmpFile For Input As #1
    Open outFile For Output As #2
 
    'Write Headers
    
    For i = 0 To tdbgIRPJCSLLCOFINSPIS.Columns.Count - 1
         If i = 0 Then
             sLine = Chr(34) & tdbgIRPJCSLLCOFINSPIS.Columns(i).Caption & Chr(34)
         Else
             sLine = sLine & ";" & Chr(34) & tdbgIRPJCSLLCOFINSPIS.Columns(i).Caption & Chr(34)
         End If
    Next
 
    Print #2, sLine

     'Write contents of original file
    Do Until EOF(1)
         Line Input #1, sLine
         Print #2, sLine
    Loop
 
    Close #1
    Close #2
 
    Kill tmpFile
          
    Screen.MousePointer = vbDefault
    
    MsgBox "Total exportado: " & adoIRPJCSLLCOFINSPIS.Recordset.RecordCount & " registros", vbInformation, "Informação"
    
End Sub
Private Function validaCampos() As Boolean

    Dim result As Boolean
    Dim message As String
    
    result = False
    message = ""

    If cboCCProvisaoIRPJOrigem.Text = "" Then
        message = message & "Conta corrente origem para IRPJ é obrigatório" & vbCrLf
        result = True
    End If
    
    If cboCCProvisaoIRPJDestino.Text = "" Then
        message = message & "Conta corrente destino para IRPJ é obrigatório" & vbCrLf
        result = True
    End If
    
    If cboCCProvisaoCSLLOrigem.Text = "" Then
        message = message & "Conta corrente origem para CSLL é obrigatório" & vbCrLf
        result = True
    End If
    
    If cboCCProvisaoCSLLDestino.Text = "" Then
        message = message & "Conta corrente destino para CSLL é obrigatório" & vbCrLf
        result = True
    End If
    
    If cboCCProvisaoCOFINSOrigem.Text = "" Then
        message = message & "Conta corrente origem para CONFINS é obrigatório" & vbCrLf
        result = True
    End If
    
    If cboCCProvisaoCOFINSDestino.Text = "" Then
        message = message & "Conta corrente destino para CONFINS é obrigatório" & vbCrLf
        result = True
    End If
    
    If cboCCProvisaoPISOrigem.Text = "" Then
        message = message & "Conta corrente origem para PIS é obrigatório" & vbCrLf
        result = True
    End If
    
    If cboCCProvisaoPISDestino.Text = "" Then
        message = message & "Conta corrente destino para PIS é obrigatório" & vbCrLf
        result = True
    End If
    
    If cboCCProvisaoIRPFOrigem.Text = "" Then
        message = message & "Conta corrente origem para IRPF é obrigatório" & vbCrLf
        result = True
    End If
    
    If cboCCProvisaoIRPFDestino.Text = "" Then
        message = message & "Conta corrente destino para IRPF é obrigatório" & vbCrLf
        result = True
    End If
    
    If cboCCProvisaoISSOrigem.Text = "" Then
        message = message & "Conta corrente origem para ISS é obrigatório" & vbCrLf
        result = True
    End If
    
    If cboCCProvisaoISSDestino.Text = "" Then
        message = message & "Conta corrente destino para ISS é obrigatório" & vbCrLf
        result = True
    End If
    
    If result = True Then
         MsgBox message, vbInformation, "Informação"
    End If
    
    validaCampos = result
End Function

Private Sub cmdGravar_Click()
    Dim XLI_TIPODOCPAGAR As Integer
    Dim XLO_ProvisaoImpostos As ADODB.Recordset
    Dim XLI_TIPODOCRECEBER As Integer
    Dim XLI_PROVISAO As Integer
    Dim XLI_LOTE As Integer
    Dim XLI_SQL As String
    Dim XLI_CONTAIRRF As Integer
    Dim XLI_CONTACOFINS As Integer
    Dim XLI_REGIME As Integer
    
    Me.MousePointer = vbHourglass
    
    If validaCampos = True Then
        Exit Sub
    End If
    
    SubQOpenRecordset XLO_ProvisaoImpostos, "SELECT * FROM ContasProvisaoImpostos WHERE empr_cd_empresa =" & PCodEmpresa, "Estatico"
    
    If (XLO_ProvisaoImpostos.BOF And XLO_ProvisaoImpostos.EOF) Then
        
        XLI_SQL = "INSERT INTO ContasProvisaoImpostos (" & _
                  "empr_cd_empresa, coco_cd_origemIRPJ, coco_cd_destinoIRPJ, coco_cd_origemCSLL, coco_cd_destinoCSLL," & _
                  "coco_cd_origemCOFINS, coco_cd_destinoCOFINS, coco_cd_origemPIS, coco_cd_destinoPIS," & _
                  "coco_cd_origemIRPF, coco_cd_destinoIRPF, coco_cd_origemISS, coco_cd_destinoISS ) VALUES ( " & _
                  " " & PCodEmpresa & ", " & CInt(txtCCProvisaoIRPJOrigem.Text) & ", " & CInt(txtCCProvisaoIRPJDestino.Text) & ", " & CInt(txtCCProvisaoCSLLOrigem.Text) & "," & CInt(txtCCProvisaoCSLLDestino.Text) & "," & _
                  " " & CInt(txtCCProvisaoCOFINSOrigem.Text) & ", " & CInt(txtCCProvisaoCOFINSDestino.Text) & ", " & CInt(txtCCProvisaoPISOrigem.Text) & ", " & CInt(txtCCProvisaoPISDestino.Text) & "," & _
                  " " & CInt(txtCCProvisaoIRPFOrigem.Text) & ", " & CInt(txtCCProvisaoIRPFDestino.Text) & ", " & CInt(txtCCProvisaoISSOrigem.Text) & ", " & CInt(txtCCProvisaoISSDestino.Text) & ")"
    Else
        XLI_SQL = "UPDATE ContasProvisaoImpostos SET " & _
                  "coco_cd_origemIRPJ = " & CInt(txtCCProvisaoIRPJOrigem.Text) & ", coco_cd_destinoIRPJ = " & CInt(txtCCProvisaoIRPJDestino.Text) & ", coco_cd_origemCSLL = " & CInt(txtCCProvisaoCSLLOrigem.Text) & ", coco_cd_destinoCSLL = " & CInt(txtCCProvisaoCSLLDestino.Text) & "," & _
                  "coco_cd_origemCOFINS = " & CInt(txtCCProvisaoCOFINSOrigem.Text) & ", coco_cd_destinoCOFINS = " & CInt(txtCCProvisaoCOFINSDestino.Text) & ", coco_cd_origemPIS = " & CInt(txtCCProvisaoPISOrigem.Text) & ", coco_cd_destinoPIS = " & CInt(txtCCProvisaoPISDestino.Text) & "," & _
                  "coco_cd_origemIRPF = " & CInt(txtCCProvisaoIRPFOrigem.Text) & ", coco_cd_destinoIRPF = " & CInt(txtCCProvisaoIRPFDestino.Text) & ", coco_cd_origemISS = " & CInt(txtCCProvisaoISSOrigem.Text) & ",coco_cd_destinoISS = " & CInt(txtCCProvisaoISSDestino.Text) & " WHERE empr_cd_empresa = " & PCodEmpresa & " "
    End If
    
    Conexao.Execute XLI_SQL
    
    'PEGA OS CODIGOS DOS TIPOS DE DOCUMENTOS PARA PAGAR E RECEBER
    Dim XLO_RS As ADODB.Recordset
    SubQOpenRecordset XLO_RS, "SELECT tido_cd_tipodoc, tido_tx_sigla FROM TipoDocumentos WHERE tido_tx_sigla IN ('TBCR', 'TBCP')", Estatico
    While Not XLO_RS.EOF
        If (XLO_RS!tido_tx_sigla = "TBCP") Then
            XLI_TIPODOCPAGAR = FunNulo(XLO_RS!tido_cd_TipoDoc)
        ElseIf (XLO_RS!tido_tx_sigla = "TBCR") Then
            XLI_TIPODOCRECEBER = FunNulo(XLO_RS!tido_cd_TipoDoc)
        End If
        
        XLO_RS.MoveNext
    Wend
    XLO_RS.Close
    
    XLI_PROVISAO = FunPegaGuardaUltimo("P")
    
    SubQOpenRecordset XLO_RS, "SELECT * FROM Lancamentos WHERE lanc_dt_data = " & FunNuloData(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year))), NomeSgbd) & " AND lanc_tx_origem LIKE 'F%' AND empr_cd_empresa = " & PCodEmpresa & " ORDER BY lanc_nr_lote DESC, lanc_cd_lanc DESC", Estatico
        
    If XLO_RS.EOF Then
       XLI_LOTE = (FunDefinirUltimoLote(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year)))))
    Else
       XLI_LOTE = XLO_RS!lanc_nr_Lote
    End If
    XLO_RS.Close

    InsereProvisao "IRRF", XLI_PROVISAO, XLI_LOTE, _
        cboCCProvisaoIRPJOrigem.BoundText, _
        cboCCProvisaoIRPJDestino.BoundText, _
        cboCCProvisaoCOFINSDestino.BoundText, _
        XLI_TIPODOCPAGAR, XLI_TIPODOCRECEBER, adoIRPJCSLLCOFINSPIS
    
    InsereProvisao "CSLL", XLI_PROVISAO, XLI_LOTE, _
        cboCCProvisaoCSLLOrigem.BoundText, _
        cboCCProvisaoCSLLDestino.BoundText, _
        cboCCProvisaoCOFINSDestino.BoundText, _
        XLI_TIPODOCPAGAR, XLI_TIPODOCRECEBER, adoIRPJCSLLCOFINSPIS
    
    InsereProvisao "COFINS", XLI_PROVISAO, XLI_LOTE, _
        cboCCProvisaoCOFINSOrigem.BoundText, _
        cboCCProvisaoCOFINSDestino.BoundText, _
        cboCCProvisaoPISDestino.BoundText, _
        XLI_TIPODOCPAGAR, XLI_TIPODOCRECEBER, adoIRPJCSLLCOFINSPIS
    
    InsereProvisao "PIS", XLI_PROVISAO, XLI_LOTE, _
        cboCCProvisaoPISOrigem.BoundText, _
        cboCCProvisaoPISDestino.BoundText, _
        cboCCProvisaoIRPFDestino.BoundText, _
        XLI_TIPODOCPAGAR, XLI_TIPODOCRECEBER, adoIRPJCSLLCOFINSPIS
        
    If XFB_SIMPLESNACIONAL Then
        'INSS - CPP
        InsereProvisao "CPP", XLI_PROVISAO, XLI_LOTE, _
            cboCCProvisaoCPPOrigem.BoundText, _
            cboCCProvisaoCPPDestino.BoundText, _
            0, _
            XLI_TIPODOCPAGAR, XLI_TIPODOCRECEBER, adoIRPJCSLLCOFINSPIS
    End If
    
    
    If (adoISS.Recordset.RecordCount > 0) Then
        InsereProvisao "ISS", XLI_PROVISAO, XLI_LOTE, cboCCProvisaoISSOrigem.BoundText, cboCCProvisaoISSDestino.BoundText, 0, XLI_TIPODOCPAGAR, XLI_TIPODOCRECEBER, adoISS
    End If

    If (chkBaixaImpostos.Value = 1) Then
        BaixaImpostos "COFINS", XLI_PROVISAO, cboCCProvisaoCOFINSOrigem.BoundText, cboCCProvisaoCOFINSOrigem.Text, cboCCProvisaoCOFINSDestino.BoundText, cboCCProvisaoCOFINSDestino.Text, 137, 137, dtpBaixaCOFINS, adoIRPJCSLLCOFINSPIS
        BaixaImpostos "PIS", XLI_PROVISAO, cboCCProvisaoPISOrigem.BoundText, cboCCProvisaoPISOrigem.Text, cboCCProvisaoPISDestino.BoundText, cboCCProvisaoPISDestino.Text, 137, 137, dtpBaixaPIS, adoIRPJCSLLCOFINSPIS
        ''BaixaImpostos "ISS", XLI_PROVISAO, cboCCProvisaoISSDestino.BoundText, cboCCProvisaoISSOrigem.BoundText, 145, 145, dtpBaixaISS, adoISSPagamento
    End If
    
    lblNumeroProvisao.Caption = CStr(Format(XLI_PROVISAO, "0"))
    
    If cmbRegime.Text = "Competencia" Then
        XLI_REGIME = 1
    Else
        XLI_REGIME = 0
    End If
    
    Conexao.Execute "EXEC sp_ProvisaoIRPJ_CSLL_COFINS_PIS @empresa = " & CStr(PCodEmpresa) & ", @emissao = " & FunNuloData("01/" & Month(dtpCompetenciaPISCOFINS.Value) & "/" & Year(dtpCompetenciaPISCOFINS.Value), "Sql Server") & ", @aliquotaIrpj  = " & Replace(IIf(RTrim(LTrim(txtAliquotaIRPJ.Text)) = "", "0.00", txtAliquotaIRPJ.Text), ",", ".") & ", @aliquotaCSLL  = " & Replace(IIf(RTrim(LTrim(txtAliquotaCSLL.Text)) = "", "0.00", txtAliquotaCSLL.Text), ",", ".") & ", @aliquotaCPP  = " & Replace(IIf(RTrim(LTrim(txtAliquotaCPP.Text)) = "", "0.00", txtAliquotaCPP.Text), ",", ".") & ", @NumeroProvisao = " & CStr(Format(XLI_PROVISAO, "0")) & ", @competencia = " & XLI_REGIME
    Conexao.Execute "EXEC sp_ProvisaoISS @empresa = " & CStr(PCodEmpresa) & ", @emissao = " & FunNuloData("01/" & Month(dtpCompetenciaISS.Value) & "/" & Year(dtpCompetenciaISS.Value), "Sql Server") & ", @NumeroProvisao = " & CStr(Format(XLI_PROVISAO, "0")) & ", @competencia = " & XLI_REGIME
    
    cmdGravar.Enabled = False
    cmdPesquisar.Enabled = False
    
    Me.MousePointer = vbDefault
    
    MsgBox "Operação efetuada com sucesso", vbInformation, "Informação"
End Sub

Private Sub CmdPesquisar_Click()
    Dim XLD_DATAPROVISAO As Date
    
    If (cmbRegime.Text = "") Then
        MsgBox "Regime é campo obrigatório", vbOKOnly, "Aviso"
        Exit Sub
    End If

    XLD_DATAPROVISAO = CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year)

    dtpCompetenciaPISCOFINS.Value = XLD_DATAPROVISAO
    dtpCompetenciaISS.Value = XLD_DATAPROVISAO
    
    XLD_DATAPROVISAO = DateAdd("m", 1, XLD_DATAPROVISAO)
    
    Dim Dia As Integer
    Dim diaDaSemana As Integer
    
    diaDaSemana = Format(CDate("20/" & Month(XLD_DATAPROVISAO) & "/" & Year(XLD_DATAPROVISAO)), "w")
    
    'domingo = 1 e sábado = 7
    If (diaDaSemana = 1) Then
        Dia = 2
    ElseIf (diaDaSemana = 7) Then
        Dia = 1
    Else
        Dia = 0
    End If
    
    dtpBaixaCOFINS.Value = CDate(CStr(20 - Dia) & "/" & Month(XLD_DATAPROVISAO) & "/" & Year(XLD_DATAPROVISAO))
    dtpBaixaPIS.Value = CDate(CStr(20 - Dia) & "/" & Month(XLD_DATAPROVISAO) & "/" & Year(XLD_DATAPROVISAO))
    dtpBaixaISS.Value = CDate(CStr(5 - Dia) & "/" & Month(XLD_DATAPROVISAO) & "/" & Year(XLD_DATAPROVISAO))
    
    PreencherDadosISS
    PreencherDadosIR
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub InsereContasReceberProvisao( _
    XLT_IMPOSTO As String, _
    NumeroProvisao As Integer, _
    Lote As Integer, _
    ContaCorrente As Integer, _
    ContaContabilDestino As Integer, _
    ValorProvisao As Double, _
    TipoDocumentoReceber As Integer, _
    dados As Adodc)
       
    Dim XLI_I As Integer
    Dim XLI_FATURA As Long
    Dim XLI_IDNOTAFISCAL As Long
    Dim XLI_IDDUPLICATA As Long
    Dim XLO_RS As ADODB.Recordset
    
    SubQOpenRecordset XLO_RS, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal= -1", Dinamico
    XLO_RS.AddNew

    XLI_FATURA = FunPegaGuardaUltimo("F")
    XLO_RS!nofi_dt_emissao = Format$(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year))), "DD/MM/YY")
    XLO_RS!nofi_tx_tipomov = "E"
    XLO_RS!nofi_tx_status = "R"
    XLO_RS!nofi_nr_documento = Format(FunPegaGuardaUltimo("D"), "A000000000")
    XLO_RS!nofi_nr_provisao = NumeroProvisao
    XLO_RS!focl_cd_forncli = Null
    XLO_RS!nofi_tx_serie = "U"
    XLO_RS!tido_cd_TipoDoc = TipoDocumentoReceber
    XLO_RS!nofi_dt_entrada = Format$(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year))), "DD/MM/YY")
    XLO_RS!nofi_tx_historico = "CRÉDITO RETENÇÕES " & XLT_IMPOSTO & " " & Format$(dtpProvisao.Value, "MM/YYYY")
    XLO_RS!nofi_vl_valor = ValorProvisao
    
    If (dados.Recordset.RecordCount > 1) Then
        XLO_RS!nofi_tx_distribuir = "S"
    Else
        XLO_RS!nofi_tx_distribuir = "N"
    End If
    
    XLO_RS!nofi_nr_fatura = XLI_FATURA
    XLO_RS!nofi_tx_tipofatura = "1"
    XLO_RS!empr_cd_empresa = PCodEmpresa
    
    XLO_RS.Update
    XLI_IDNOTAFISCAL = XLO_RS!nofi_cd_notafiscal
    XLO_RS.Close
   
    dados.Recordset.MoveFirst
    While Not dados.Recordset.BOF And Not dados.Recordset.EOF
        
        If (CDbl(dados.Recordset("ValorRetido" & XLT_IMPOSTO).Value) > 0) Then
            
            Conexao.Execute "INSERT INTO DISTRIBUICAO " & _
                    "(" & _
                    "    plco_cd_Conta, " & _
                    "    cere_cd_Pcr, " & _
                    "    nofi_cd_notafiscal, " & _
                    "    dist_vl_valor, " & _
                    "    dist_vl_GlosaDevolucao, " & _
                    "    iden_cd_Projeto " & _
                    ") " & _
                    "VALUES " & _
                    "( " & _
                    "    " & ContaContabilDestino & ", " & _
                    "    " & dados.Recordset("CRId").Value & ", " & _
                    "    " & XLI_IDNOTAFISCAL & ", " & _
                    "    " & Replace(dados.Recordset("ValorRetido" & XLT_IMPOSTO).Value, ",", ".") & ", " & _
                    "    0, " & _
                    "    NULL " & _
                    ")"
        End If
        
        dados.Recordset.MoveNext
    Wend
    
    SubQOpenRecordset XLO_RS, "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata = -1", Dinamico
    XLO_RS.AddNew
    
    XLO_RS!dupl_nr_duplicata = "U"
    XLO_RS!dupl_dt_vencimento = Format$(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year))), "DD/MM/YY")
    XLO_RS!dupl_vl_GlosaDevolucao = 0
    XLO_RS!dupl_vl_valor = ValorProvisao
    XLO_RS!dupl_tx_status = "R"
    XLO_RS!dupl_tx_tipopagrec = "4"
    XLO_RS!dupl_tx_docpagrec = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
    XLO_RS!coco_cd_codigo = ContaCorrente
    XLO_RS!bafo_cd_codigo = Null
    XLO_RS!dupl_tx_debcred = "C"
    XLO_RS!dupl_nr_fatura = XLI_FATURA
    XLO_RS!dupl_dt_pagrec = Format$(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year))), "DD/MM/YY")
    
    XLO_RS.Update
    XLI_IDDUPLICATA = XLO_RS!dupl_cd_Duplicata
    XLO_RS.Close
End Sub

'************************************************************************************************************************************************************************************************
'* Considerando que a provisão que está sendo executada é a de Set/AAAA então o sistema vai lançar um ou mais compromissados no mês Out/AAAA
'* para pagar o imposto devido pela empresa.
'*
'* Essa conta é feita da seguinte forma:
'*
'*  Se TODAS as notas têm retenção do IMPOSTO: (TIPO 1)
'*                        - Um compromissado a PAGAR
'*                                - Competencia e Vencimento: em Out/AAAA
'*                                - Valor: [Faturamento Total] * [% devido]
'*    (1)                         - Conta Corrente: Conta Provisao
'*                                - Distribuido por todos os contratos
'*                                - Características do Tipo de Documento: Não sai em custo e não contabiliza
'*                                - Histórico: BX RETENÇÃO [IMPOSTO] MM/AAAA
'*
'*
'*
'*
'*
'* Se NÃO HÁ nenhuma nota com retenção do imposto, então lança:  (TIPO 2)
'*                        - Um compromissado a PAGAR
'*                                - Competencia e Vencimento: em Out/AAAA
'*                                - Valor: [Faturamento Total] * [% devido]
'*    (1)                         - Conta Corrente: Conta Provisao Imposto
'*                                - Distribuido por todos os contratos
'*                                - Características do Tipo de Documento: Não sai em custo e não contabiliza
'*                                - Histórico: TRANFS p/ [CONTA MOVIMENTO] MM/AAAA
'*
'*                        - Um compromissado a RECEBER
'*                                - Competencia e Vencimento: em Out/AAAA
'*                                - Valor: [Faturamento Total] * [% devido]
'*    (2)                         - Conta Corrente: Conta Movimento
'*                                - Distribuido por todos os contratos
'*                                - Características do Tipo de Documento: Não sai em custo e não contabiliza
'*                                - Histórico: TRANFS [CONTA IMPOSTO] p/ [CONTA MOVIMENTO] MM/AAAA
'*
'*                        - Um compromissado a PAGAR
'*                                - Competencia e Vencimento: em Out/AAAA
'*                                - Valor: [Faturamento Total] * [% devido]
'*    (4)                         - Conta Corrente: Conta Movimento
'*                                - Distribuido por todos os contratos
'*                                - Características do Tipo de Documento: Não sai em custo e não contabiliza
'*                                - Histórico: PG [IMPOSTO] MM/AAAA
'*
'*
'*
'*
'*
'*  Se ALGUMAS NOTAS possuem retenção e outras não:  (TIPO 4)
'*                        - Um compromissado a PAGAR
'*                                - Competencia e Vencimento: em Out/AAAA
'*                                - Valor: SOMA([Valores retidos de cada contrato]
'*    (1)                         - Conta Corrente: Conta Provisao
'*                                - Distribuido por todos os contratos
'*                                - Características do Tipo de Documento: Não sai em custo e não contabiliza
'*                                - Histórico: PG COMPLEMENTO [IMPOSTO] MM/AAAA
'*
'*                        - Um compromissado a RECEBER
'*                                - Competencia e Vencimento: em Out/AAAA
'*                                - Valor: ([Faturamento Total] * [% devido]) - SOMA([Valores retidos de cada contrato])
'*    (2)                         - Conta Corrente: Conta Movimento
'*                                - Distribuido por todos os contratos
'*                                - Características do Tipo de Documento: Não sai em custo e não contabiliza
'*                                - Histórico: PG [IMPOSTO] MM/AAAA
'*
'*                        - Um compromissado a PAGAR
'*                                - Competencia e Vencimento: em Out/AAAA
'*                                - Valor: ([Faturamento Total] * [% devido]) - SOMA([Valores retidos de cada contrato])
'*    (4)                         - Conta Corrente: Conta Movimento
'*                                - Distribuido por todos os contratos
'*                                - Características do Tipo de Documento: Não sai em custo e não contabiliza
'*                                - Histórico: PG [IMPOSTO] MM/AAAA
'*
'*                        - Um compromissado a PAGAR
'*                                - Competencia e Vencimento: em Out/AAAA
'*                                - Valor: ([Faturamento Total] * [% devido]) - SOMA([Valores retidos de cada contrato])
'*    (8)                         - Conta Corrente: Conta Provisao
'*                                - Distribuido por todos os contratos
'*                                - Características do Tipo de Documento: Não sai em custo e não contabiliza
'*                                - Histórico: PG [IMPOSTO] MM/AAAA
'***********************************************************************************************************************************************************************************************

Private Sub BaixaImpostos(Imposto As String, NumeroProvisao As Integer, ContaCorrenteOrigem As Integer, DescricaoContaCorrenteOrigem As String, ContaCorrenteDestino As Integer, DescricaoContaCorrenteDestino As String, TipoDocumentoPagar As Integer, TipoDocumentoReceber As Integer, Vencimento As Date, dados As Adodc)
    Dim XLI_I As Integer
    Dim XLI_FATURA As Long
    Dim XLI_IDNOTAFISCAL As Long
    Dim XLI_IDDUPLICATA As Long
    Dim XLD_TOTALRETIDO As Double
    Dim XLD_FATURAMENTO As Double
    Dim XLD_PERCENTUALDEVIDO As Double
    Dim XLO_RS As ADODB.Recordset
    Dim XLI_TIPO As Integer
    Dim XLD_VALOR As Double
    Dim XLI_CONTACONTABILORIGEM As Integer
    Dim XLI_CONTACONTABILDESTINO As Integer
    
    Dim XLB_TODASRETENCAO As Boolean
    Dim XLB_ALGUMASRETENCAO As Boolean
    Dim XLB_SEMRETENCAO As Boolean
    
    XLB_TODASRETENCAO = True
    XLB_SEMRETENCAO = True
    
    If (Imposto = "COFINS") Then
        XLD_PERCENTUALDEVIDO = 3 / 100
    ElseIf (Imposto = "PIS") Then
        XLD_PERCENTUALDEVIDO = 0.65 / 100
    End If
    
     'TOTAL RETIDO
    dados.Recordset.MoveFirst
    While Not dados.Recordset.BOF And Not dados.Recordset.EOF
        XLD_FATURAMENTO = XLD_FATURAMENTO + CDbl(dados.Recordset("ValorFaturado").Value)
        
        If CDbl(dados.Recordset("ValorRetido" & Imposto).Value) > 0 Then
            XLB_SEMRETENCAO = False
        ElseIf CDbl(dados.Recordset("ValorRetido" & Imposto).Value) = 0 Then
            XLB_TODASRETENCAO = False
        End If
        dados.Recordset.MoveNext
    Wend
            
    If XLB_TODASRETENCAO And Not XLB_SEMRETENCAO Then
        XLI_TIPO = 1
    ElseIf XLB_SEMRETENCAO And Not XLB_TODASRETENCAO Then
        XLI_TIPO = 7
    Else
        XLI_TIPO = 15
    End If

    'TOTAL RETIDO
    dados.Recordset.MoveFirst
    While Not dados.Recordset.BOF And Not dados.Recordset.EOF
        If (Imposto <> "ISS") Then
            XLD_TOTALRETIDO = XLD_TOTALRETIDO + CDbl(dados.Recordset("ValorRetido" & Imposto).Value)
        Else
            XLD_TOTALRETIDO = XLD_TOTALRETIDO + CDbl(dados.Recordset("ValorProvisao" & Imposto).Value)
        End If
        dados.Recordset.MoveNext
    Wend
    
    'Alteração filtro CC por CGC Raiz - Samuel 27/09/18
    If Mid(PEmpresa, 1, 10) = "QUALIDADOS" Or Mid(PEmpresa, 1, 8) = "QUALINFO" Then
        SubQOpenRecordset XLO_RS, "SELECT coco_cd_codigo, plco_cd_Conta FROM ContasCorrente WHERE coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' AND coco_cd_codigo IN (" & ContaCorrenteOrigem & ", " & ContaCorrenteDestino & ")", Estatico
    Else
        SubQOpenRecordset XLO_RS, "SELECT coco_cd_codigo, plco_cd_Conta FROM ContasCorrente WHERE empr_cd_empresa = " & CStr(PCodEmpresa) & " AND coco_cd_codigo IN (" & ContaCorrenteOrigem & ", " & ContaCorrenteDestino & ")", Estatico
    End If
    While Not XLO_RS.EOF
        If XLO_RS!coco_cd_codigo = ContaCorrenteOrigem Then
            XLI_CONTACONTABILORIGEM = XLO_RS!plco_cd_conta
        ElseIf XLO_RS!coco_cd_codigo = ContaCorrenteDestino Then
            XLI_CONTACONTABILDESTINO = XLO_RS!plco_cd_conta
        End If
        XLO_RS.MoveNext
    Wend
    XLO_RS.Close
    
    For XLI_I = 0 To 3
        SubQOpenRecordset XLO_RS, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal= -1", Dinamico
        XLO_RS.AddNew

        XLI_FATURA = FunPegaGuardaUltimo("F")
        XLO_RS!focl_cd_forncli = 376
        XLO_RS!nofi_dt_emissao = Format$(Vencimento, "DD/MM/YYYY")
        
        If (XLI_I <> 1) Then
            XLO_RS!nofi_tx_tipomov = "S"
        Else
            XLO_RS!nofi_tx_tipomov = "E"
        End If
        
        XLO_RS!nofi_tx_status = "C"
        XLO_RS!nofi_nr_documento = Format(FunPegaGuardaUltimo("D"), "A000000000")
        XLO_RS!nofi_nr_provisao = NumeroProvisao
        XLO_RS!focl_cd_forncli = 376
        XLO_RS!nofi_tx_serie = "U"
        
        If (XLI_I <> 1) Then
            XLO_RS!tido_cd_TipoDoc = TipoDocumentoPagar
        Else
            XLO_RS!tido_cd_TipoDoc = TipoDocumentoReceber
        End If
        
        XLO_RS!nofi_dt_entrada = Format$(Vencimento, "DD/MM/YYYY")
        
        If (XLI_I = 0) Then
            
            '1 - Histórico: BX RETENÇÃO [IMPOSTO] MM/AAAA
            '1 - Histórico: TRANFS p/ [CONTA MOVIMENTO] MM/AAAA
            '1 - Histórico: PG COMPLEMENTO [IMPOSTO] MM/AAAA
                        
            '1 - Valor: [Faturamento Total] * [% devido]
            '1 - Valor: [Faturamento Total] * [% devido]
            '1 - Valor: SOMA([Valores retidos de cada contrato]
            If (XLI_TIPO = 1) Then
                XLO_RS!nofi_tx_historico = "BX RETENÇÃO " & Imposto & " " & Format$(dtpProvisao.Value, "MM/YYYY")
                XLO_RS!nofi_vl_valor = XLD_FATURAMENTO * XLD_PERCENTUALDEVIDO
            ElseIf (XLI_TIPO = 7) Then
                XLO_RS!nofi_tx_historico = "TRANSF p/ " & DescricaoContaCorrenteDestino & " p/ BX RETENÇÃO " & Format$(dtpProvisao.Value, "MM/YYYY")
                XLO_RS!nofi_vl_valor = XLD_FATURAMENTO * XLD_PERCENTUALDEVIDO
            ElseIf (XLI_TIPO = 15) Then
                XLO_RS!nofi_tx_historico = "PG COMPLEMENTO " & Imposto & " " & Format$(dtpProvisao.Value, "MM/YYYY")
                XLO_RS!nofi_vl_valor = XLD_TOTALRETIDO
            End If
            
        ElseIf (XLI_I = 1) Then
        
            '2 - Histórico: TRANFS [CONTA IMPOSTO] p/ [CONTA MOVIMENTO] MM/AAAA
            '2 - Histórico: PG [IMPOSTO] MM/AAAA
            
            '2 - Valor: [Faturamento Total] * [% devido]
            '2 - Valor: ([Faturamento Total] * [% devido]) - SOMA([Valores retidos de cada contrato])
            If (XLI_TIPO = 7 And ((XLI_TIPO And 2) = 2)) Then
                XLO_RS!nofi_tx_historico = "TRANFS " & Imposto & " p/ " & DescricaoContaCorrenteDestino & " " & Format$(dtpProvisao.Value, "MM/YYYY")
                XLO_RS!nofi_vl_valor = XLD_FATURAMENTO * XLD_PERCENTUALDEVIDO
            ElseIf (XLI_TIPO = 15 And ((XLI_TIPO And 2) = 2)) Then
                XLO_RS!nofi_tx_historico = "PG IMPOSTO " & Imposto & " " & Format$(dtpProvisao.Value, "MM/YYYY")
                XLO_RS!nofi_vl_valor = Round((XLD_FATURAMENTO * XLD_PERCENTUALDEVIDO) - XLD_TOTALRETIDO, 2)
            End If
            
        ElseIf (XLI_I = 2) Then
        
            '4 - Histórico: PG [IMPOSTO] MM/AAAA
            '4 - Histórico: PG [IMPOSTO] MM/AAAA
            
            '4 - Valor: [Faturamento Total] * [% devido]
            '4 - Valor: ([Faturamento Total] * [% devido]) - SOMA([Valores retidos de cada contrato])
            If (XLI_TIPO = 7 And ((XLI_TIPO And 4) = 4) And XLI_I = 2) Then
                XLO_RS!nofi_tx_historico = "PG IMPOSTO " & Imposto & " " & Format$(dtpProvisao.Value, "MM/YYYY")
                XLO_RS!nofi_vl_valor = (XLD_FATURAMENTO * XLD_PERCENTUALDEVIDO)
            ElseIf (XLI_TIPO = 15 And ((XLI_TIPO And 4) = 4) And XLI_I = 2) Then
                XLO_RS!nofi_tx_historico = "PG IMPOSTO " & Imposto & " " & Format$(dtpProvisao.Value, "MM/YYYY")
                XLO_RS!nofi_vl_valor = (XLD_FATURAMENTO * XLD_PERCENTUALDEVIDO) - XLD_TOTALRETIDO
            End If
            
            
        ElseIf (XLI_I = 3) Then
        
            '8 - Histórico: PG [IMPOSTO] MM/AAAA
            '8 - Valor: ([Faturamento Total] * [% devido]) - SOMA([Valores retidos de cada contrato])
            If (XLI_TIPO = 15 And ((XLI_TIPO And 8) = 8) And XLI_I = 3) Then
                XLO_RS!nofi_tx_historico = "PG IMPOSTO " & Imposto & " " & Format$(dtpProvisao.Value, "MM/YYYY")
                XLO_RS!nofi_vl_valor = (XLD_FATURAMENTO * XLD_PERCENTUALDEVIDO) - XLD_TOTALRETIDO
            End If
        End If
        
        XLD_VALOR = XLO_RS!nofi_vl_valor
                             
        If (dados.Recordset.RecordCount > 1) Then
            XLO_RS!nofi_tx_distribuir = "S"
        Else
            XLO_RS!nofi_tx_distribuir = "N"
        End If

        XLO_RS!nofi_nr_fatura = XLI_FATURA
        XLO_RS!nofi_tx_tipofatura = "1"
        XLO_RS!empr_cd_empresa = PCodEmpresa

        XLO_RS.Update
        XLI_IDNOTAFISCAL = XLO_RS!nofi_cd_notafiscal
        XLO_RS.Close
         
        dados.Recordset.MoveFirst
        While Not dados.Recordset.BOF And Not dados.Recordset.EOF
            If (Imposto <> "ISS") Then
                If CDbl(dados.Recordset("ValorRetido" & Imposto).Value) > 0 And XLI_I = 0 And XLI_TIPO = 15 Then
                    
                    Conexao.Execute "INSERT INTO DISTRIBUICAO " & _
                            "(" & _
                            "    plco_cd_Conta, " & _
                            "    cere_cd_Pcr, " & _
                            "    nofi_cd_notafiscal, " & _
                            "    dist_vl_valor, " & _
                            "    dist_vl_GlosaDevolucao, " & _
                            "    iden_cd_Projeto " & _
                            ") " & _
                            "VALUES " & _
                            "( " & _
                            "    " & IIf(XLI_I = 0, XLI_CONTACONTABILORIGEM, XLI_CONTACONTABILDESTINO) & ", " & _
                            "    " & dados.Recordset("CRId").Value & ", " & _
                            "    " & XLI_IDNOTAFISCAL & ", " & _
                            "    " & Replace(dados.Recordset("ValorRetido" & Imposto).Value, ",", ".") & ", " & _
                            "    0, " & _
                            "    NULL " & _
                            ")"
                ElseIf CDbl(dados.Recordset("ValorRetido" & Imposto).Value) = 0 And XLI_TIPO = 15 Then
                    Conexao.Execute "INSERT INTO DISTRIBUICAO (plco_cd_Conta, cere_cd_Pcr, nofi_cd_notafiscal, dist_vl_valor, dist_vl_GlosaDevolucao, iden_cd_Projeto) VALUES (" & IIf(XLI_I = 0, XLI_CONTACONTABILORIGEM, XLI_CONTACONTABILDESTINO) & ", " & " " & dados.Recordset("CRId").Value & ", " & "    " & XLI_IDNOTAFISCAL & ", " & "    " & Replace(dados.Recordset("ValorFaturado").Value * XLD_PERCENTUALDEVIDO, ",", ".") & ", " & "    0, " & "    NULL " & ")"
                End If
            Else
                If (CDbl(dados.Recordset("ValorProvisao" & Imposto).Value) > 0) Then
                    
                    Conexao.Execute "INSERT INTO DISTRIBUICAO " & _
                            "(" & _
                            "    plco_cd_Conta, " & _
                            "    cere_cd_Pcr, " & _
                            "    nofi_cd_notafiscal, " & _
                            "    dist_vl_valor, " & _
                            "    dist_vl_GlosaDevolucao, " & _
                            "    iden_cd_Projeto " & _
                            ") " & _
                            "VALUES " & _
                            "( " & _
                            "    " & IIf(XLI_I = 0, XLI_CONTACONTABILORIGEM, XLI_CONTACONTABILDESTINO) & ", " & _
                            "    " & dados.Recordset("CRId").Value & ", " & _
                            "    " & XLI_IDNOTAFISCAL & ", " & _
                            "    " & Replace(dados.Recordset("ValorProvisao" & Imposto).Value, ",", ".") & ", " & _
                            "    0, " & _
                            "    NULL " & _
                            ")"
                End If
            End If
            dados.Recordset.MoveNext
        Wend

        SubQOpenRecordset XLO_RS, "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata = -1", Dinamico
        XLO_RS.AddNew

        XLO_RS!dupl_nr_duplicata = "U"
        XLO_RS!dupl_dt_vencimento = Format$(Vencimento, "DD/MM/YYYY")
        XLO_RS!dupl_vl_GlosaDevolucao = 0
        
        XLO_RS!dupl_vl_valor = XLD_VALOR
        
        XLO_RS!dupl_tx_status = "C"
        XLO_RS!dupl_tx_tipopagrec = "1"
        
        If (XLI_I = 0) Then
            XLO_RS!coco_cd_codigo = ContaCorrenteOrigem
        ElseIf (XLI_I = 1) Then
            XLO_RS!coco_cd_codigo = ContaCorrenteDestino
        ElseIf (XLI_I = 2) Then
            XLO_RS!coco_cd_codigo = ContaCorrenteDestino
        End If
        
        XLO_RS!bafo_cd_codigo = Null
        If (XLI_I <> 1) Then
            XLO_RS!dupl_tx_debcred = "D"
        Else
            XLO_RS!dupl_tx_debcred = "C"
        End If
        
        XLO_RS!dupl_nr_fatura = XLI_FATURA
        XLO_RS!dupl_dt_pagrec = Null

        XLO_RS.Update
        XLI_IDDUPLICATA = XLO_RS!dupl_cd_Duplicata
        XLO_RS.Close
    Next XLI_I
End Sub

Private Sub InsereProvisao(XLT_IMPOSTO As String, NumeroProvisao As Integer, Lote As Integer, _
    ContaCorrenteOrigem As Integer, _
    ContaCorrenteDestino As Integer, _
    ContaCorrenteDestinoTransferencia As Integer, _
    TipoDocumentoCusto As Integer, _
    TipoDocumentoNaoCusto As Integer, _
    dados As Adodc)
    
    Dim XLI_PAGAR As Integer
    Dim XLI_RECEBER As Integer
    Dim XLI_I As Integer
    Dim XLI_FATURA As Long
    Dim XLI_IDNOTAFISCAL As Long
    Dim XLI_IDDUPLICATA As Long
    Dim XLI_CONTACONTABILORIGEM As Integer
    Dim XLI_CONTACONTABILDESTINO As Integer
    Dim XLI_CONTACONTABILDESTINOTRANSF As Integer
    Dim XLO_RS As ADODB.Recordset
    Dim XLT_HISTORICO As String
    
    Dim XLD_TOTALRETIDO As Double
    Dim XLD_TOTALPOSITIVOS As Double
    Dim XLD_TOTALNEGATIVOS As Double
    
    SubQOpenRecordset XLO_RS, "SELECT coco_cd_codigo, plco_cd_Conta FROM ContasCorrente WHERE empr_cd_empresa = " & CStr(PCodEmpresa) & " AND coco_cd_codigo IN (" & ContaCorrenteOrigem & ", " & ContaCorrenteDestino & ", " & ContaCorrenteDestinoTransferencia & ")", Estatico
    While Not XLO_RS.EOF
        If XLO_RS!coco_cd_codigo = ContaCorrenteOrigem Then
            XLI_CONTACONTABILORIGEM = XLO_RS!plco_cd_conta
        ElseIf XLO_RS!coco_cd_codigo = ContaCorrenteDestino Then
            XLI_CONTACONTABILDESTINO = XLO_RS!plco_cd_conta
        ElseIf XLO_RS!coco_cd_codigo = ContaCorrenteDestinoTransferencia Then
            XLI_CONTACONTABILDESTINOTRANSF = XLO_RS!plco_cd_conta
        End If
        XLO_RS.MoveNext
    Wend
    XLO_RS.Close
        
    'TOTAL RETIDO
    dados.Recordset.MoveFirst
    While Not dados.Recordset.BOF And Not dados.Recordset.EOF
        If XLT_IMPOSTO = "CPP" Then
            XLD_TOTALRETIDO = XLD_TOTALRETIDO + CDbl(dados.Recordset("ValorProvisaoCPP").Value)
        Else
            XLD_TOTALRETIDO = XLD_TOTALRETIDO + CDbl(dados.Recordset("ValorRetido" & XLT_IMPOSTO).Value)
        End If
        dados.Recordset.MoveNext
    Wend
        
    If XLT_IMPOSTO <> "CPP" Then
        'INSERE CREDITO NA CONTA DE PROVISAO DO IMPOSTO (PASSO 1)
        InsereContasReceberProvisao XLT_IMPOSTO, NumeroProvisao, Lote, ContaCorrenteDestino, XLI_CONTACONTABILDESTINO, XLD_TOTALRETIDO, TipoDocumentoNaoCusto, dados
    End If
                
    'VALORES POSITIVOS E NEGATIVOS
    dados.Recordset.MoveFirst
    While Not dados.Recordset.BOF And Not dados.Recordset.EOF
        If (CDbl(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value) > 0) Then
            XLD_TOTALPOSITIVOS = XLD_TOTALPOSITIVOS + CDbl(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value)
        Else
            XLD_TOTALNEGATIVOS = XLD_TOTALNEGATIVOS + Abs(CDbl(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value))
        End If
        dados.Recordset.MoveNext
    Wend
        
    For XLI_I = 2 To IIf(ContaCorrenteDestinoTransferencia > 0, 5, 3) Step 1
        SubQOpenRecordset XLO_RS, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal= -1", Dinamico
        XLO_RS.AddNew
    
        XLI_FATURA = FunPegaGuardaUltimo("F")
        XLO_RS!nofi_dt_emissao = Format$(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year))), "DD/MM/YY")
        
        'If XLI_I Mod 2 = 0 Then
        If XLI_I = 2 Or XLI_I = 4 Then
            XLO_RS!nofi_tx_tipomov = "S"
        ElseIf XLI_I = 3 Or XLI_I = 5 Then
            XLO_RS!nofi_tx_tipomov = "E"
        End If
        
        XLO_RS!nofi_tx_status = "R"
        XLO_RS!nofi_nr_documento = Format(FunPegaGuardaUltimo("D"), "A000000000")
        XLO_RS!nofi_nr_provisao = NumeroProvisao
        XLO_RS!focl_cd_forncli = Null
        XLO_RS!nofi_tx_serie = "U"
        
        If XLI_I = 2 Then
            XLO_RS!tido_cd_TipoDoc = TipoDocumentoCusto
        ElseIf XLI_I = 3 Then
            XLO_RS!tido_cd_TipoDoc = TipoDocumentoNaoCusto
        ElseIf XLI_I = 4 Then
           XLO_RS!tido_cd_TipoDoc = TipoDocumentoCusto
        ElseIf XLI_I = 5 Then
            XLO_RS!tido_cd_TipoDoc = TipoDocumentoCusto
        End If
        
        XLO_RS!nofi_dt_entrada = Format$(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year))), "DD/MM/YY")
        
        If XLI_I = 2 Or XLI_I = 3 Then
            XLT_HISTORICO = "TRANSF PROVISÃO BANCO PARA " & XLT_IMPOSTO & " " & Format$(dtpProvisao.Value, "MM/YYYY")
            
        ElseIf XLI_I = 4 Or XLI_I = 5 Then
            If (XLT_IMPOSTO = "IRRF") Then
                XLT_HISTORICO = "TRANSF PROVISÃO " & XLT_IMPOSTO & " PARA COFINS " & Format$(dtpProvisao.Value, "MM/YYYY")
            ElseIf (XLT_IMPOSTO = "CSLL") Then
                XLT_HISTORICO = "TRANSF PROVISÃO " & XLT_IMPOSTO & " PARA COFINS " & Format$(dtpProvisao.Value, "MM/YYYY")
            ElseIf (XLT_IMPOSTO = "COFINS") Then
                XLT_HISTORICO = "TRANSF PROVISÃO " & XLT_IMPOSTO & " PARA PIS " & Format$(dtpProvisao.Value, "MM/YYYY")
            ElseIf (XLT_IMPOSTO = "PIS") Then
                XLT_HISTORICO = "TRANSF PROVISÃO " & XLT_IMPOSTO & " PARA IRPF " & Format$(dtpProvisao.Value, "MM/YYYY")
            End If
        End If
        
        XLO_RS!nofi_tx_historico = XLT_HISTORICO
        
        If XLI_I <= 3 Then
            XLO_RS!nofi_vl_valor = XLD_TOTALPOSITIVOS
        Else
            XLO_RS!nofi_vl_valor = XLD_TOTALNEGATIVOS
        End If
        
        If (dados.Recordset.RecordCount > 1) Then
            XLO_RS!nofi_tx_distribuir = "S"
        Else
            XLO_RS!nofi_tx_distribuir = "N"
        End If
        
        XLO_RS!nofi_nr_fatura = XLI_FATURA
        XLO_RS!nofi_tx_tipofatura = "1"
        XLO_RS!empr_cd_empresa = PCodEmpresa
        
        XLO_RS.Update
        XLI_IDNOTAFISCAL = XLO_RS!nofi_cd_notafiscal
        XLO_RS.Close
       
        dados.Recordset.MoveFirst
        While Not dados.Recordset.BOF And Not dados.Recordset.EOF
             If (XLI_I <= 3) Then
                 If (CDbl(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value) > 0) Then
                        
                        Conexao.Execute "INSERT INTO DISTRIBUICAO " & _
                                "(" & _
                                "    plco_cd_Conta, " & _
                                "    cere_cd_Pcr, " & _
                                "    nofi_cd_notafiscal, " & _
                                "    dist_vl_valor, " & _
                                "    dist_vl_GlosaDevolucao, " & _
                                "    iden_cd_Projeto " & _
                                ") " & _
                                "VALUES " & _
                                "( " & _
                                "    " & XLI_CONTACONTABILDESTINO & ", " & _
                                "    " & dados.Recordset("CRId").Value & ", " & _
                                "    " & XLI_IDNOTAFISCAL & ", " & _
                                "    " & Replace(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value, ",", ".") & ", " & _
                                "    0, " & _
                                "    NULL " & _
                                ")"
                End If
            Else
                If (CDbl(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value) < 0) Then
                        
                        Conexao.Execute "INSERT INTO DISTRIBUICAO " & _
                                "(" & _
                                "    plco_cd_Conta, " & _
                                "    cere_cd_Pcr, " & _
                                "    nofi_cd_notafiscal, " & _
                                "    dist_vl_valor, " & _
                                "    dist_vl_GlosaDevolucao, " & _
                                "    iden_cd_Projeto " & _
                                ") " & _
                                "VALUES " & _
                                "( " & _
                                "    " & IIf(XLI_I = 4, XLI_CONTACONTABILDESTINOTRANSF, XLI_CONTACONTABILDESTINO) & ", " & _
                                "    " & dados.Recordset("CRId").Value & ", " & _
                                "    " & XLI_IDNOTAFISCAL & ", " & _
                                "    " & Replace(Abs(CDbl(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value)), ",", ".") & ", " & _
                                "    0, " & _
                                "    NULL " & _
                                ")"
                End If
            End If
            dados.Recordset.MoveNext
        Wend
        
        SubQOpenRecordset XLO_RS, "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata = -1", Dinamico
        XLO_RS.AddNew
        
        XLO_RS!dupl_nr_duplicata = "U"
        XLO_RS!dupl_dt_vencimento = Format$(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year))), "DD/MM/YY")
        XLO_RS!dupl_vl_GlosaDevolucao = 0
        If XLI_I <= 3 Then
            XLO_RS!dupl_vl_valor = XLD_TOTALPOSITIVOS
        Else
            XLO_RS!dupl_vl_valor = XLD_TOTALNEGATIVOS
        End If
        
        XLO_RS!dupl_tx_status = "R"
        XLO_RS!dupl_tx_tipopagrec = "4"
        XLO_RS!dupl_tx_docpagrec = Format(CStr(FunPegaGuardaUltimo("PR")), "A00000")
        
        If (XLI_I = 2) Then
            XLO_RS!coco_cd_codigo = ContaCorrenteOrigem
        ElseIf (XLI_I = 3 Or XLI_I = 4) Then
            XLO_RS!coco_cd_codigo = ContaCorrenteDestino
        ElseIf (XLI_I = 5) Then
            XLO_RS!coco_cd_codigo = ContaCorrenteDestinoTransferencia
        End If
        
        XLO_RS!bafo_cd_codigo = Null
        
        If XLI_I = 2 Or XLI_I = 4 Then
            XLO_RS!dupl_tx_debcred = "D"
        ElseIf XLI_I = 3 Or XLI_I = 5 Then
            XLO_RS!dupl_tx_debcred = "C"
        End If
        
        XLO_RS!dupl_nr_fatura = XLI_FATURA
        XLO_RS!dupl_dt_pagrec = Format$(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year))), "DD/MM/YY")
        
        XLO_RS.Update
        XLI_IDDUPLICATA = XLO_RS!dupl_cd_Duplicata
        XLO_RS.Close
       
        'VALORES POSITIVOS
        If (XLI_I = 2) Then
            dados.Recordset.MoveFirst
            While Not dados.Recordset.BOF And Not dados.Recordset.EOF
                
                If CDbl(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value) > 0 Then
                    InsereLancamentoContabil _
                        Lote, _
                        CDbl(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value), _
                        XLI_CONTACONTABILDESTINO, _
                        dados.Recordset("CRId").Value, _
                        0, _
                        0, _
                        XLT_HISTORICO, _
                        XLI_IDNOTAFISCAL, _
                        "FN"
                End If
                dados.Recordset.MoveNext
            Wend
            
            InsereLancamentoContabil _
                   Lote, _
                   XLD_TOTALPOSITIVOS, _
                   0, _
                   0, _
                   XLI_CONTACONTABILORIGEM, _
                   0, _
                   XLT_HISTORICO, _
                   XLI_IDNOTAFISCAL, _
                   "FN"
        ElseIf (XLI_I = 4) Then
            dados.Recordset.MoveFirst
            While Not dados.Recordset.BOF And Not dados.Recordset.EOF
                If CDbl(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value) < 0 Then
                    InsereLancamentoContabil _
                        Lote, _
                        Abs(CDbl(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value)), _
                        0, _
                        0, _
                        IIf(XLI_I = 4, XLI_CONTACONTABILDESTINO, XLI_CONTACONTABILDESTINOTRANSF), _
                        dados.Recordset("CRId").Value, _
                        XLT_HISTORICO, _
                        XLI_IDNOTAFISCAL, _
                        "FN"
                End If
                dados.Recordset.MoveNext
            Wend
            
            dados.Recordset.MoveFirst
            While Not dados.Recordset.BOF And Not dados.Recordset.EOF
                
                If CDbl(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value) < 0 Then
                    InsereLancamentoContabil _
                        Lote, _
                        Abs(CDbl(dados.Recordset("ValorProvisao" & XLT_IMPOSTO).Value)), _
                        IIf(XLI_I = 4, XLI_CONTACONTABILDESTINOTRANSF, XLI_CONTACONTABILDESTINO), _
                        dados.Recordset("CRId").Value, _
                        0, _
                        0, _
                        XLT_HISTORICO, _
                        XLI_IDNOTAFISCAL, _
                        "FN"
                End If
                dados.Recordset.MoveNext
            Wend
        End If
    Next XLI_I
End Sub

Private Sub InsereLancamentoContabil(Lote As Integer, VALOR As Double, ContaDebito As Integer, CRDebito As Integer, ContaCredito As Integer, CRCredito As Integer, Historico As String, Origem As Long, OrigemFin As String)
    XGT_SQL = ""
    XGT_SQL = XGT_SQL & "INSERT INTO LANCAMENTOS "
    XGT_SQL = XGT_SQL & "("
    XGT_SQL = XGT_SQL & "   lanc_dt_data, "
    XGT_SQL = XGT_SQL & "   lanc_nr_Lote, "
    XGT_SQL = XGT_SQL & "   lanc_cd_Lanc, "
    XGT_SQL = XGT_SQL & "   empr_cd_empresa, "
    XGT_SQL = XGT_SQL & "   plco_cd_Debito, "
    XGT_SQL = XGT_SQL & "   plco_cd_Credito, "
    XGT_SQL = XGT_SQL & "   cere_cd_PcrDebitar, "
    XGT_SQL = XGT_SQL & "   cere_cd_PcrCreditar, "
    XGT_SQL = XGT_SQL & "   lanc_vl_valor, "
    XGT_SQL = XGT_SQL & "   lanc_tx_Historico, "
    XGT_SQL = XGT_SQL & "   lanc_tx_DescCustos, "
    XGT_SQL = XGT_SQL & "   iden_cd_Projeto, "
    XGT_SQL = XGT_SQL & "   lanc_cd_OrigemFin, "
    XGT_SQL = XGT_SQL & "   lanc_tx_Origem, "
    XGT_SQL = XGT_SQL & "   lanc_cd_Estoque, "
    XGT_SQL = XGT_SQL & "   lanc_cd_Patrimonio "
    XGT_SQL = XGT_SQL & ") "
    XGT_SQL = XGT_SQL & "VALUES "
    XGT_SQL = XGT_SQL & "( "
    XGT_SQL = XGT_SQL & "   " & FunNuloData(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year))), NomeSgbd) & ", "
    XGT_SQL = XGT_SQL & "   " & Lote & ", "
    XGT_SQL = XGT_SQL & "   " & funDefinirCodigoLanc(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & dtpProvisao.Month & "/" & dtpProvisao.Year))), Lote) & ", "
    XGT_SQL = XGT_SQL & "   " & CStr(PCodEmpresa) & ", "
    If (ContaDebito > 0) Then
    XGT_SQL = XGT_SQL & "   " & ContaDebito & ", "
    XGT_SQL = XGT_SQL & "   Null, "
    Else
    XGT_SQL = XGT_SQL & "   Null, "
    XGT_SQL = XGT_SQL & "   " & ContaCredito & ", "
    End If
    
    If (CRDebito > 0) Then
    XGT_SQL = XGT_SQL & "   " & CRDebito & ", "
    XGT_SQL = XGT_SQL & "   Null, "
    ElseIf (CRCredito > 0) Then
    XGT_SQL = XGT_SQL & "   Null, "
    XGT_SQL = XGT_SQL & "   " & CRCredito & ", "
    Else
    XGT_SQL = XGT_SQL & "   Null, "
    XGT_SQL = XGT_SQL & "   Null, "
    End If
    
    XGT_SQL = XGT_SQL & "   " & Replace(CStr(VALOR), ",", ".") & ", "
    XGT_SQL = XGT_SQL & "   '" & Historico & "' , "
    XGT_SQL = XGT_SQL & "   '" & Historico & "' , "
    XGT_SQL = XGT_SQL & "   null, "
    XGT_SQL = XGT_SQL & "   " & Origem & ", "
    XGT_SQL = XGT_SQL & "   '" & OrigemFin & "', "
    XGT_SQL = XGT_SQL & "   null, "
    XGT_SQL = XGT_SQL & "   null "
    XGT_SQL = XGT_SQL & ")"
                    
    Conexao.Execute XGT_SQL
End Sub

Private Sub Form_Load()
    Dim XLI_I As Integer
    Dim XLO_RS As ADODB.Recordset
    Dim XLO_ProvisaoImpostos As ADODB.Recordset

    lblNumeroProvisao.Caption = ""
    
    'Alteração filtro CC por CGC Raiz - Samuel 27/09/18
    
    If Mid(PEmpresa, 1, 10) = "QUALIDADOS" Or Mid(PEmpresa, 1, 8) = "QUALINFO" Then
        subConectarControleDadosNV datCCProvisaoPISOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'B' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoPISDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
    
        subConectarControleDadosNV datCCProvisaoCOFINSOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'B' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoCOFINSDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
    
        subConectarControleDadosNV datCCProvisaoIRPJOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'B' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoIRPJDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
    
        subConectarControleDadosNV datCCProvisaoIRPFOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoIRPFDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
    
        subConectarControleDadosNV datCCProvisaoCSLLOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'B' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoCSLLDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
    
        subConectarControleDadosNV datCCProvisaoISSOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'B' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoISSDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
    
        subConectarControleDadosNV datCCProvisaoCPPOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'B' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoCPPDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND coco_tx_CGCRaiz = '" & Mid(XGT_CGC, 1, 10) & "' ORDER BY coco_tx_Descricao", Estatico
    Else
        subConectarControleDadosNV datCCProvisaoPISOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'B' AND empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoPISDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    
        subConectarControleDadosNV datCCProvisaoCOFINSOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'B' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoCOFINSDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    
        subConectarControleDadosNV datCCProvisaoIRPJOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'B' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoIRPJDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    
        subConectarControleDadosNV datCCProvisaoIRPFOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoIRPFDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    
        subConectarControleDadosNV datCCProvisaoCSLLOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'B' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoCSLLDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    
        subConectarControleDadosNV datCCProvisaoISSOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'B' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoISSDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    
        subConectarControleDadosNV datCCProvisaoCPPOrigem, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'B' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
        subConectarControleDadosNV datCCProvisaoCPPDestino, "SELECT * FROM ConsGENCCcombo WHERE coco_tx_provisao = 'P' AND  empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    End If

    SubQOpenRecordset XLO_RS, "SELECT desc_tx_descricao, desc_vl_taxa FROM DescontosImpostos WHERE desc_tx_descricao in ('PIS', 'COFINS', 'IRRF', 'CONTRIBUIÇÃO SOCIAL')", Estatico
    While Not XLO_RS.EOF
        
        If (XLO_RS!desc_tx_descricao = "PIS") Then
            txtAliquotaPIS.Text = FunNulo(XLO_RS!desc_vl_taxa)
        ElseIf (XLO_RS!desc_tx_descricao = "COFINS") Then
            txtAliquotaCOFINS.Text = FunNulo(XLO_RS!desc_vl_taxa)
        ElseIf (XLO_RS!desc_tx_descricao = "IRRF") Then
            txtAliquotaIRPJ.Text = FunNulo(XLO_RS!desc_vl_taxa)
        ElseIf (XLO_RS!desc_tx_descricao = "CONTRIBUIÇÃO SOCIAL") Then
            txtAliquotaCSLL.Text = FunNulo(XLO_RS!desc_vl_taxa)
        End If
        
        XLO_RS.MoveNext
    Wend
    
    XLO_RS.Close
    
    SubQOpenRecordset XLO_RS, "SELECT empr_vl_IRPJ, empr_vl_PIS, empr_vl_COFINS, empr_vl_CSLL, empr_vl_CPP, empr_tx_simplesnacional FROM Empresas WHERE empr_cd_empresa = " & CStr(PCodEmpresa), Estatico
    While Not XLO_RS.EOF
        txtAliquotaIRPJ.Text = FunNulo(XLO_RS!empr_vl_IRPJ)
        txtAliquotaPIS.Text = FunNulo(XLO_RS!empr_vl_PIS)
        txtAliquotaCOFINS.Text = FunNulo(XLO_RS!empr_vl_COFINS)
        txtAliquotaCSLL.Text = FunNulo(XLO_RS!empr_vl_CSLL)
        txtAliquotaCPP.Text = FunNulo(XLO_RS!empr_vl_CPP)
        XFB_SIMPLESNACIONAL = (FunNulo(XLO_RS!empr_tx_simplesnacional) = "S")
        
        XLO_RS.MoveNext
    Wend
    XLO_RS.Close
    Set XLO_RS = Nothing
    
    If XFB_SIMPLESNACIONAL Then
        SSTab.TabCaption(1) = SSTab.TabCaption(1) & "/CPP"
        tdbgIRPJCSLLCOFINSPIS.Splits(0).Columns(PERCENTUAL_LUCRO).Visible = False
    Else
        txtCCProvisaoCPPOrigem.Enabled = False
        txtCCProvisaoCPPDestino.Enabled = False
        cboCCProvisaoCPPOrigem.Enabled = False
        cboCCProvisaoCPPDestino.Enabled = False
        fraCPP.ForeColor = &H80000011
        lblCCCPPOrigem.ForeColor = &H80000011
        lblCCCPPDestino.ForeColor = &H80000011
        lblTituloTotalProvisaoCPP.ForeColor = &H80000011
        lblTotalProvisaoCPP.ForeColor = &H80000011
        lblCPP.ForeColor = &H80000011
        txtAliquotaCPP.Enabled = False
        txtAliquotaCPP.Text = "0,00"
        tdbgIRPJCSLLCOFINSPIS.Splits(1).Columns(PERCENTUAL_RETIDO_CPP).Visible = False
        tdbgIRPJCSLLCOFINSPIS.Splits(1).Columns(PROVISAO_CPP).Visible = False
    End If
    
    If (XGT_MESTRABALHO = "00:00:00") Then
        dtpProvisao.Value = Now
        dtpBaixaCOFINS.Value = CDate("20/" & Month(Now) & "/" & Year(Now))
        dtpBaixaPIS.Value = CDate("20/" & Month(Now) & "/" & Year(Now))
        dtpBaixaISS.Value = CDate("20/" & Month(Now) & "/" & Year(Now))
    Else
        dtpProvisao.Value = DateAdd("d", -1, DateAdd("m", 1, XGT_MESTRABALHO))
        dtpBaixaCOFINS.Value = CDate("20/" & Month(DateAdd("d", -1, DateAdd("m", 1, XGT_MESTRABALHO))) & "/" & Year(DateAdd("d", -1, DateAdd("m", 1, XGT_MESTRABALHO))))
        dtpBaixaPIS.Value = CDate("20/" & Month(DateAdd("d", -1, DateAdd("m", 1, XGT_MESTRABALHO))) & "/" & Year(DateAdd("d", -1, DateAdd("m", 1, XGT_MESTRABALHO))))
        dtpBaixaISS.Value = CDate("20/" & Month(DateAdd("d", -1, DateAdd("m", 1, XGT_MESTRABALHO))) & "/" & Year(DateAdd("d", -1, DateAdd("m", 1, XGT_MESTRABALHO))))
    End If
    
    dtpCompetenciaPISCOFINS.Value = CDate("01/" & Month(Now) & "/" & Year(Now))
    
    SubQOpenRecordset XLO_ProvisaoImpostos, "SELECT * FROM ContasProvisaoImpostos WHERE empr_cd_empresa =" & PCodEmpresa, "Estatico"
    
    If Not (XLO_ProvisaoImpostos.BOF And XLO_ProvisaoImpostos.EOF) Then
        txtCCProvisaoPISOrigem.Text = XLO_ProvisaoImpostos!coco_cd_origemPIS
        txtCCProvisaoPISOrigem_LostFocus
        txtCCProvisaoCOFINSOrigem.Text = XLO_ProvisaoImpostos!coco_cd_origemCOFINS
        txtCCProvisaoCOFINSOrigem_LostFocus
        txtCCProvisaoIRPJOrigem.Text = XLO_ProvisaoImpostos!coco_cd_origemIRPJ
        txtCCProvisaoIRPJOrigem_LostFocus
        txtCCProvisaoIRPFOrigem.Text = XLO_ProvisaoImpostos!coco_cd_origemIRPF
        txtCCProvisaoIRPFOrigem_LostFocus
        txtCCProvisaoCSLLOrigem.Text = XLO_ProvisaoImpostos!coco_cd_origemCSLL
        txtCCProvisaoCSLLOrigem_LostFocus
        txtCCProvisaoISSOrigem.Text = XLO_ProvisaoImpostos!coco_cd_origemISS
        txtCCProvisaoISSOrigem_LostFocus
        
        txtCCProvisaoPISDestino.Text = XLO_ProvisaoImpostos!coco_cd_destinoPIS
        txtCCProvisaoPISDestino_LostFocus
        txtCCProvisaoCOFINSDestino.Text = XLO_ProvisaoImpostos!coco_cd_destinoCOFINS
        txtCCProvisaoCOFINSDestino_LostFocus
        txtCCProvisaoIRPJDestino.Text = XLO_ProvisaoImpostos!coco_cd_destinoIRPJ
        txtCCProvisaoIRPJDestino_LostFocus
        txtCCProvisaoIRPFDestino.Text = XLO_ProvisaoImpostos!coco_cd_destinoIRPF
        txtCCProvisaoIRPFDestino_LostFocus
        txtCCProvisaoCSLLDestino.Text = XLO_ProvisaoImpostos!coco_cd_destinoCSLL
        txtCCProvisaoCSLLDestino_LostFocus
        txtCCProvisaoISSDestino.Text = XLO_ProvisaoImpostos!coco_cd_destinoISS
        txtCCProvisaoISSDestino_LostFocus
    
        If XFB_SIMPLESNACIONAL Then
            txtCCProvisaoCPPOrigem.Text = XLO_ProvisaoImpostos!coco_cd_origemCPP
            txtCCProvisaoCPPOrigem_LostFocus
        
            txtCCProvisaoCPPDestino.Text = XLO_ProvisaoImpostos!coco_cd_destinoCPP
            txtCCProvisaoCPPDestino_LostFocus
        End If
    End If
    
'    If (PEmpresa <> "QUALIDADOS CONSULTORIA E SISTEMAS LTDA") Then
'        cmdGravar.Enabled = False
'        lblNumeroProvisao.Caption = CStr(xChave)
'    End If
    
'    If (PEmpresa = "QUALIDADOS CONSULTORIA E SISTEMAS LTDA" And xChave <= 0) Then
'        txtCCProvisaoPISOrigem.Text = 181
'        txtCCProvisaoPISOrigem_LostFocus
'        txtCCProvisaoCOFINSOrigem.Text = 181
'        txtCCProvisaoCOFINSOrigem_LostFocus
'        txtCCProvisaoIRPJOrigem.Text = 181
'        txtCCProvisaoIRPJOrigem_LostFocus
'        txtCCProvisaoIRPFOrigem.Text = 215
'        txtCCProvisaoIRPFOrigem_LostFocus
'        txtCCProvisaoCSLLOrigem.Text = 181
'        txtCCProvisaoCSLLOrigem_LostFocus
'        txtCCProvisaoISSOrigem.Text = 181
'        txtCCProvisaoISSOrigem_LostFocus
'
'
'        txtCCProvisaoPISDestino.Text = 215
'        txtCCProvisaoPISDestino_LostFocus
'        txtCCProvisaoCOFINSDestino.Text = 216
'        txtCCProvisaoCOFINSDestino_LostFocus
'        txtCCProvisaoIRPJDestino.Text = 213
'        txtCCProvisaoIRPJDestino_LostFocus
'        txtCCProvisaoIRPFDestino.Text = 242
'        txtCCProvisaoIRPFDestino_LostFocus
'        txtCCProvisaoCSLLDestino.Text = 214
'        txtCCProvisaoCSLLDestino_LostFocus
'        txtCCProvisaoISSDestino.Text = 202
'        txtCCProvisaoISSDestino_LostFocus
'    Else
        'cmdGravar.Enabled = False
        
        
        'lblNumeroProvisao.Caption = CStr(xChave)
        
'        XGT_SQL = _
'            "SELECT DISTINCT DUPLICATAS.dupl_cd_duplicata, DUPLICATAS.coco_cd_codigo, Duplicatas.dupl_dt_PagRec FROM NotasFiscais " & _
'            "INNER JOIN Distribuicao ON NotasFiscais.nofi_cd_notafiscal = Distribuicao.nofi_cd_notafiscal " & _
'            "INNER JOIN Duplicatas ON Duplicatas.dupl_nr_fatura  = NotasFiscais.nofi_nr_fatura " & _
'            "INNER JOIN PlanoContas ON PlanoContas.plco_cd_Conta = Distribuicao.plco_cd_Conta " & _
'            "WHERE NotasFiscais.nofi_nr_provisao = " & CStr(xChave) & "  AND NotasFiscais.nofi_tx_tipomov = 'S' " & _
'            "UNION ALL " & _
'            "SELECT DISTINCT DUPLICATAS.dupl_cd_duplicata, DUPLICATAS.coco_cd_codigo, Duplicatas.dupl_dt_PagRec FROM NotasFiscais " & _
'            "INNER JOIN Distribuicao ON NotasFiscais.nofi_cd_notafiscal = Distribuicao.nofi_cd_notafiscal " & _
'            "INNER JOIN Duplicatas ON Duplicatas.dupl_nr_fatura  = NotasFiscais.nofi_nr_fatura " & _
'            "INNER JOIN PlanoContas ON PlanoContas.plco_cd_Conta = Distribuicao.plco_cd_Conta " & _
'            "WHERE NotasFiscais.nofi_nr_provisao = " & CStr(xChave) & "  AND NotasFiscais.nofi_tx_tipomov = 'E' " & _
'            "ORDER BY 1 "
'
'        SubQOpenRecordset XLO_RS, XGT_SQL, Estatico
'
'        XLI_I = 0
'        While Not XLO_RS.EOF
'
'            If XLI_I = 0 Then
'                txtCCProvisaoPISOrigem.Text = XLO_RS!coco_cd_codigo
'                txtCCProvisaoPISOrigem_LostFocus
'            ElseIf XLI_I = 1 Then
'                txtCCProvisaoPISDestino.Text = XLO_RS!coco_cd_codigo
'                txtCCProvisaoPISDestino_LostFocus
'            ElseIf XLI_I = 2 Then
'                txtCCProvisaoCOFINSOrigem.Text = XLO_RS!coco_cd_codigo
'                txtCCProvisaoCOFINSOrigem_LostFocus
'            ElseIf XLI_I = 3 Then
'                txtCCProvisaoCOFINSDestino.Text = XLO_RS!coco_cd_codigo
'                txtCCProvisaoCOFINSDestino_LostFocus
'            ElseIf XLI_I = 4 Then
'                txtCCProvisaoIRPJOrigem.Text = XLO_RS!coco_cd_codigo
'                txtCCProvisaoIRPJOrigem_LostFocus
'            ElseIf XLI_I = 5 Then
'                txtCCProvisaoIRPJDestino.Text = XLO_RS!coco_cd_codigo
'                txtCCProvisaoIRPJDestino_LostFocus
'            ElseIf XLI_I = 6 Then
'                txtCCProvisaoCSLLOrigem.Text = XLO_RS!coco_cd_codigo
'                txtCCProvisaoCSLLOrigem_LostFocus
'            ElseIf XLI_I = 7 Then
'                txtCCProvisaoISSDestino.Text = XLO_RS!coco_cd_codigo
'                txtCCProvisaoISSDestino_LostFocus
'            ElseIf XLI_I = 8 Then
'                txtCCProvisaoISSOrigem.Text = XLO_RS!coco_cd_codigo
'                txtCCProvisaoISSOrigem_LostFocus
'            ElseIf XLI_I = 9 Then
'                txtCCProvisaoCSLLDestino.Text = XLO_RS!coco_cd_codigo
'                txtCCProvisaoCSLLDestino_LostFocus
'
'                dtpProvisao.Value = XLO_RS!dupl_dt_PagRec
'
'                dtpCompetenciaPISCOFINS.Value = CDate("01/" & Month(XLO_RS!dupl_dt_PagRec) & "/" & Year(XLO_RS!dupl_dt_PagRec))
'                dtpCompetenciaISS.Value = CDate("01/" & Month(XLO_RS!dupl_dt_PagRec) & "/" & Year(XLO_RS!dupl_dt_PagRec))
'            End If
'
'            XLI_I = XLI_I + 1
'            XLO_RS.MoveNext
'        Wend
'        XLO_RS.Close


'
'        If (PEmpresa = "QUALIDADOS CONSULTORIA E SISTEMAS LTDA") Then
'            txtCCProvisaoPISOrigem.Text = 181
'            txtCCProvisaoPISOrigem_LostFocus
'            txtCCProvisaoCOFINSOrigem.Text = 181
'            txtCCProvisaoCOFINSOrigem_LostFocus
'            txtCCProvisaoIRPJOrigem.Text = 181
'            txtCCProvisaoIRPJOrigem_LostFocus
'            txtCCProvisaoIRPFOrigem.Text = 215
'            txtCCProvisaoIRPFOrigem_LostFocus
'            txtCCProvisaoCSLLOrigem.Text = 181
'            txtCCProvisaoCSLLOrigem_LostFocus
'            txtCCProvisaoISSOrigem.Text = 181
'            txtCCProvisaoISSOrigem_LostFocus
'        End If
'
'        txtCCProvisaoPISDestino.Text = 215
'        txtCCProvisaoPISDestino_LostFocus
'        txtCCProvisaoCOFINSDestino.Text = 216
'        txtCCProvisaoCOFINSDestino_LostFocus
'        txtCCProvisaoIRPJDestino.Text = 213
'        txtCCProvisaoIRPJDestino_LostFocus
'        txtCCProvisaoIRPFDestino.Text = 242
'        txtCCProvisaoIRPFDestino_LostFocus
'        txtCCProvisaoCSLLDestino.Text = 214
'        txtCCProvisaoCSLLDestino_LostFocus
'        txtCCProvisaoISSDestino.Text = 202
'        txtCCProvisaoISSDestino_LostFocus
'--------------------------------------------
'        subCarregaDadosNV adoIRPJCSLLCOFINSPIS, "", "", "", "select * from HistoricoProvisaoImpostosIR where NumeroProvisao = " & CStr(xChave)
'        subCarregaDadosNV adoISS, "", "", "", "select * from HistoricoProvisaoImpostosISS where NumeroProvisao = " & CStr(xChave)
'
'        Dim XLD_FATURAMENTO As Double
'        Dim XLD_VALORRETIDOISS As Double
'        Dim XLD_RETENCAOISS As Double
'        Dim XLD_PROVISAOISS As Double
'
'
'        Dim XLD_DEVIDO_IRPJ As Double
'        Dim XLD_VALORRETIDO_IRPJ As Double
'        Dim XLD_PROVISAO_IRPJ As Double
'
'        Dim XLD_DEVIDO_CSLL As Double
'        Dim XLD_VALORRETIDO_CSLL As Double
'        Dim XLD_PROVISAO_CSLL As Double
'
'        Dim XLD_DEVIDO_COFINS As Double
'        Dim XLD_TRANSF_IRPJ_COFINS As Double
'        Dim XLD_TRANSF_CSLL_COFINS As Double
'        Dim XLD_VALORRETIDO_COFINS As Double
'        Dim XLD_PROVISAO_COFINS As Double
'
'        Dim XLD_DEVIDO_PIS As Double
'        Dim XLD_VALORRETIDO_PIS As Double
'        Dim XLD_TRANSF_COFINS_PIS As Double
'        Dim XLD_PROVISAO_PIS As Double
'
'
'
'        XLD_FATURAMENTO = 0
'
'        XLD_DEVIDO_IRPJ = 0
'        XLD_VALORRETIDO_IRPJ = 0
'        XLD_PROVISAO_IRPJ = 0
'
'        XLD_DEVIDO_CSLL = 0
'        XLD_VALORRETIDO_CSLL = 0
'        XLD_PROVISAO_CSLL = 0
'
'        XLD_DEVIDO_COFINS = 0
'        XLD_TRANSF_IRPJ_COFINS = 0
'        XLD_TRANSF_CSLL_COFINS = 0
'        XLD_VALORRETIDO_COFINS = 0
'        XLD_PROVISAO_COFINS = 0
'
'        XLD_DEVIDO_PIS = 0
'        XLD_VALORRETIDO_PIS = 0
'        XLD_TRANSF_COFINS_PIS = 0
'        XLD_PROVISAO_PIS = 0
'
'        XLD_VALORRETIDOISS = 0
'        XLD_RETENCAOISS = 0
'        XLD_PROVISAOISS = 0
'
'
'
'        If (Not adoISS.Recordset.EOF And Not adoISS.Recordset.BOF) Then
'            adoISS.Recordset.MoveFirst
'        End If
'
'        While Not adoISS.Recordset.EOF And Not adoISS.Recordset.BOF
'
'                XLD_FATURAMENTO = XLD_FATURAMENTO + CDbl(adoISS.Recordset("ValorFaturado"))
'
'                XLD_VALORRETIDOISS = XLD_VALORRETIDOISS + CDbl(adoISS.Recordset("ValorRetidoISS"))
'                XLD_RETENCAOISS = XLD_RETENCAOISS + CDbl(adoISS.Recordset("RetencaoISS"))
'                XLD_PROVISAOISS = XLD_PROVISAOISS + CDbl(adoISS.Recordset("ValorProvisaoISS"))
'
'                adoISS.Recordset.MoveNext
'        Wend
'
'
'        tdbgISS.Columns(GRIDISS_ESTRUTURADO).FooterText = "TOTAIS"
'        tdbgISS.Columns(GRIDISS_FATURAMENTO).FooterText = CStr(FormatNumber(XLD_FATURAMENTO, 2, True, True, True))
'        tdbgISS.Columns(GRIDISS_VALORRETIDOISS).FooterText = CStr(FormatNumber(XLD_VALORRETIDOISS, 2, True, True, True))
'
'        tdbgISS.Columns(GRIDISS_PROVISAOISS).FooterText = CStr(FormatNumber(XLD_PROVISAOISS, 2, True, True, True))
'
'
'        If Not adoIRPJCSLLCOFINSPIS.Recordset.EOF And Not adoIRPJCSLLCOFINSPIS.Recordset.BOF Then
'            adoIRPJCSLLCOFINSPIS.Recordset.MoveFirst
'        End If
'        While Not adoIRPJCSLLCOFINSPIS.Recordset.EOF And Not adoIRPJCSLLCOFINSPIS.Recordset.BOF
'
'                XLD_FATURAMENTO = XLD_FATURAMENTO + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorFaturado")))
'
'                XLD_DEVIDO_IRPJ = XLD_DEVIDO_IRPJ + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorDevidoIRRF")))
'                XLD_VALORRETIDO_IRPJ = XLD_VALORRETIDO_IRPJ + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorRetidoIRRF")))
'                XLD_PROVISAO_IRPJ = XLD_PROVISAO_IRPJ + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorProvisaoIRRF")))
'
'                XLD_DEVIDO_CSLL = XLD_DEVIDO_CSLL + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorDevidoCSLL")))
'                XLD_VALORRETIDO_CSLL = XLD_VALORRETIDO_CSLL + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorRetidoCSLL")))
'                XLD_PROVISAO_CSLL = XLD_PROVISAO_CSLL + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorProvisaoCSLL")))
'
'                XLD_DEVIDO_COFINS = XLD_DEVIDO_COFINS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorDevidoCOFINS")))
'                XLD_TRANSF_IRPJ_COFINS = XLD_TRANSF_IRPJ_COFINS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("TransfIRPJCOFINS")))
'                XLD_TRANSF_CSLL_COFINS = XLD_TRANSF_CSLL_COFINS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("TransfCSLLCOFINS")))
'                XLD_VALORRETIDO_COFINS = XLD_VALORRETIDO_COFINS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorRetidoCOFINS")))
'                XLD_PROVISAO_COFINS = XLD_PROVISAO_COFINS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorProvisaoCOFINS")))
'
'
'                XLD_DEVIDO_PIS = XLD_DEVIDO_PIS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorDevidoPIS")))
'                XLD_TRANSF_COFINS_PIS = XLD_TRANSF_COFINS_PIS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("TransfCOFINSPIS")))
'                XLD_VALORRETIDO_PIS = XLD_VALORRETIDO_PIS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorRetidoPIS")))
'                XLD_PROVISAO_PIS = XLD_PROVISAO_PIS + CDbl(FunNuloVal(adoIRPJCSLLCOFINSPIS.Recordset("ValorProvisaoPIS")))
'
'
'                adoIRPJCSLLCOFINSPIS.Recordset.MoveNext
'        Wend
'
'
'        tdbgIRPJCSLLCOFINSPIS.Columns(ESTRUTURADO).FooterText = "TOTAIS"
'        tdbgIRPJCSLLCOFINSPIS.Columns(VALORFATURADO).FooterText = CStr(FormatNumber(XLD_FATURAMENTO, 2, True, True, True))
'
'        tdbgIRPJCSLLCOFINSPIS.Columns(DEVIDO_IRPJ).FooterText = CStr(FormatNumber(XLD_DEVIDO_IRPJ, 2, True, True, True))
'        tdbgIRPJCSLLCOFINSPIS.Columns(RETIDO_IRPJ).FooterText = CStr(FormatNumber(XLD_VALORRETIDO_IRPJ, 2, True, True, True))
'        tdbgIRPJCSLLCOFINSPIS.Columns(PROVISAO_IRPJ).FooterText = CStr(FormatNumber(XLD_PROVISAO_IRPJ, 2, True, True, True))
'
'
'        tdbgIRPJCSLLCOFINSPIS.Columns(DEVIDO_CSLL).FooterText = CStr(FormatNumber(XLD_DEVIDO_CSLL, 2, True, True, True))
'        tdbgIRPJCSLLCOFINSPIS.Columns(RETIDO_CSLL).FooterText = CStr(FormatNumber(XLD_VALORRETIDO_CSLL, 2, True, True, True))
'        tdbgIRPJCSLLCOFINSPIS.Columns(PROVISAO_CSLL).FooterText = CStr(FormatNumber(XLD_PROVISAO_CSLL, 2, True, True, True))
'
'
'        tdbgIRPJCSLLCOFINSPIS.Columns(DEVIDO_COFINS).FooterText = CStr(FormatNumber(XLD_DEVIDO_COFINS, 2, True, True, True))
'        tdbgIRPJCSLLCOFINSPIS.Columns(RETIDO_COFINS).FooterText = CStr(FormatNumber(XLD_VALORRETIDO_COFINS, 2, True, True, True))
'        tdbgIRPJCSLLCOFINSPIS.Columns(TRANSF_IRPJ_COFINS).FooterText = CStr(FormatNumber(XLD_TRANSF_IRPJ_COFINS, 2, True, True, True))
'        tdbgIRPJCSLLCOFINSPIS.Columns(TRANSF_CSLL_COFINS).FooterText = CStr(FormatNumber(XLD_TRANSF_CSLL_COFINS, 2, True, True, True))
'        tdbgIRPJCSLLCOFINSPIS.Columns(PROVISAO_COFINS).FooterText = CStr(FormatNumber(XLD_PROVISAO_COFINS, 2, True, True, True))
'
'
'        tdbgIRPJCSLLCOFINSPIS.Columns(DEVIDO_PIS).FooterText = CStr(FormatNumber(XLD_DEVIDO_PIS, 2, True, True, True))
'        tdbgIRPJCSLLCOFINSPIS.Columns(RETIDO_PIS).FooterText = CStr(FormatNumber(XLD_VALORRETIDO_PIS, 2, True, True, True))
'        tdbgIRPJCSLLCOFINSPIS.Columns(TRANSF_COFINS_PIS).FooterText = CStr(FormatNumber(XLD_TRANSF_COFINS_PIS, 2, True, True, True))
'        tdbgIRPJCSLLCOFINSPIS.Columns(PROVISAO_PIS).FooterText = CStr(FormatNumber(XLD_PROVISAO_PIS, 2, True, True, True))
'
'        lblTotalProvisaoIRRF.Caption = CStr(FormatNumber(XLD_PROVISAO_IRPJ, 2, True, True, True))
'        lblTotalProvisaoCSLL.Caption = CStr(FormatNumber(XLD_PROVISAO_CSLL, 2, True, True, True))
'        lblTotalProvisaoCOFINS.Caption = CStr(FormatNumber(XLD_PROVISAO_COFINS, 2, True, True, True))
'        lblTotalProvisaoPIS.Caption = CStr(FormatNumber(XLD_PROVISAO_PIS, 2, True, True, True))
'
'
'        If (XLD_PROVISAO_IRPJ >= 0) Then
'            lblTotalProvisaoIRRF.ForeColor = &HFF0000
'        Else
'            lblTotalProvisaoIRRF.ForeColor = &HFF&
'        End If
'
'        If (XLD_PROVISAO_CSLL >= 0) Then
'            lblTotalProvisaoCSLL.ForeColor = &HFF0000
'        Else
'            lblTotalProvisaoCSLL.ForeColor = &HFF&
'        End If
'
'        If (XLD_PROVISAO_COFINS >= 0) Then
'            lblTotalProvisaoCOFINS.ForeColor = &HFF0000
'        Else
'            lblTotalProvisaoCOFINS.ForeColor = &HFF&
'        End If
'
'        If (XLD_PROVISAO_PIS >= 0) Then
'            lblTotalProvisaoPIS.ForeColor = &HFF0000
'        Else
'            lblTotalProvisaoPIS.ForeColor = &HFF&
'        End If
'
'        If (tdbgISS.Columns(GRIDISS_PROVISAOISS).FooterText <> "") Then
'            XLD_PROVISAOISS = CDbl(tdbgISS.Columns(GRIDISS_PROVISAOISS).FooterText)
'        Else
'            XLD_PROVISAOISS = 0
'        End If
'
'        lblTotalProvisaoISS.Caption = CStr(FormatNumber(XLD_PROVISAOISS, 2, True, True, True))
'        If (XLD_PROVISAOISS >= 0) Then
'            lblTotalProvisaoISS.ForeColor = &HFF0000
'        Else
'            lblTotalProvisaoISS.ForeColor = &HFF&
'        End If
'
'        lblTotalGeralProvisao.Caption = CStr(FormatNumber(XLD_PROVISAO_IRPJ + XLD_PROVISAO_CSLL + XLD_PROVISAO_COFINS + XLD_PROVISAO_PIS + XLD_PROVISAOISS, 2, True, True, True))
'
'        If (XLD_PROVISAO_IRPJ + XLD_PROVISAO_CSLL + XLD_PROVISAO_COFINS + XLD_PROVISAO_PIS + XLD_PROVISAOISS >= 0) Then
'            lblTotalGeralProvisao.ForeColor = &HFF0000
'        Else
'            lblTotalGeralProvisao.ForeColor = &HFF&
'        End If
  'End If
    
End Sub

Private Sub txtCCProvisaoPISOrigem_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoPISOrigem)
End Sub

Private Sub txtCCProvisaoPISOrigem_LostFocus()
   If txtCCProvisaoPISOrigem.Text <> "" Then
        cboCCProvisaoPISOrigem.BoundText = txtCCProvisaoPISOrigem.Text
        If cboCCProvisaoPISOrigem.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoPISOrigem.SetFocus
        End If
    End If
End Sub

Private Sub txtCCProvisaoPISDestino_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoPISDestino)
End Sub

Private Sub txtCCProvisaoPISDestino_LostFocus()
   If txtCCProvisaoPISDestino.Text <> "" Then
        cboCCProvisaoPISDestino.BoundText = txtCCProvisaoPISDestino.Text
        If cboCCProvisaoPISDestino.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoPISDestino.SetFocus
        End If
    End If
End Sub

Private Sub cboCCProvisaoPISOrigem_Change()
    If cboCCProvisaoPISOrigem.Text <> "" Then
        datCCProvisaoPISOrigem.Recordset.Bookmark = cboCCProvisaoPISOrigem.SelectedItem
        txtCCProvisaoPISOrigem.Text = cboCCProvisaoPISOrigem.BoundText
    End If
End Sub
    
Private Sub cboCCProvisaoPISDestino_Change()
    If cboCCProvisaoPISDestino.Text <> "" Then
        datCCProvisaoPISDestino.Recordset.Bookmark = cboCCProvisaoPISDestino.SelectedItem
        txtCCProvisaoPISDestino.Text = cboCCProvisaoPISDestino.BoundText
    End If
End Sub
    
Private Sub txtCCProvisaoCOFINSOrigem_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoCOFINSOrigem)
End Sub

Private Sub txtCCProvisaoCOFINSOrigem_LostFocus()
   If txtCCProvisaoCOFINSOrigem.Text <> "" Then
        cboCCProvisaoCOFINSOrigem.BoundText = txtCCProvisaoCOFINSOrigem.Text
        If cboCCProvisaoCOFINSOrigem.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoCOFINSOrigem.SetFocus
        End If
    End If
End Sub

Private Sub txtCCProvisaoCOFINSDestino_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoCOFINSDestino)
End Sub

Private Sub txtCCProvisaoCOFINSDestino_LostFocus()
   If txtCCProvisaoCOFINSDestino.Text <> "" Then
        cboCCProvisaoCOFINSDestino.BoundText = txtCCProvisaoCOFINSDestino.Text
        If cboCCProvisaoCOFINSDestino.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoCOFINSDestino.SetFocus
        End If
    End If
End Sub

Private Sub cboCCProvisaoCOFINSOrigem_Change()
    If cboCCProvisaoCOFINSOrigem.Text <> "" Then
        datCCProvisaoCOFINSOrigem.Recordset.Bookmark = cboCCProvisaoCOFINSOrigem.SelectedItem
        txtCCProvisaoCOFINSOrigem.Text = cboCCProvisaoCOFINSOrigem.BoundText
    End If
End Sub
    
Private Sub cboCCProvisaoCOFINSDestino_Change()
    If cboCCProvisaoCOFINSDestino.Text <> "" Then
        datCCProvisaoCOFINSDestino.Recordset.Bookmark = cboCCProvisaoCOFINSDestino.SelectedItem
        txtCCProvisaoCOFINSDestino.Text = cboCCProvisaoCOFINSDestino.BoundText
    End If
End Sub

Private Sub txtCCProvisaoIRPFOrigem_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoIRPFOrigem)
End Sub

Private Sub txtCCProvisaoIRPFOrigem_LostFocus()
   If txtCCProvisaoIRPFOrigem.Text <> "" Then
        cboCCProvisaoIRPFOrigem.BoundText = txtCCProvisaoIRPFOrigem.Text
        If cboCCProvisaoIRPFOrigem.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoIRPFOrigem.SetFocus
        End If
    End If
End Sub

Private Sub txtCCProvisaoIRPFDestino_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoIRPFDestino)
End Sub

Private Sub txtCCProvisaoIRPFDestino_LostFocus()
   If txtCCProvisaoIRPFDestino.Text <> "" Then
        cboCCProvisaoIRPFDestino.BoundText = txtCCProvisaoIRPFDestino.Text
        If cboCCProvisaoIRPFDestino.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoIRPFDestino.SetFocus
        End If
    End If
End Sub

Private Sub cboCCProvisaoIRPFOrigem_Change()
    If cboCCProvisaoIRPFOrigem.Text <> "" Then
        datCCProvisaoIRPFOrigem.Recordset.Bookmark = cboCCProvisaoIRPFOrigem.SelectedItem
        txtCCProvisaoIRPFOrigem.Text = cboCCProvisaoIRPFOrigem.BoundText
    End If
End Sub
    
Private Sub cboCCProvisaoIRPFDestino_Change()
    If cboCCProvisaoIRPFDestino.Text <> "" Then
        datCCProvisaoIRPFDestino.Recordset.Bookmark = cboCCProvisaoIRPFDestino.SelectedItem
        txtCCProvisaoIRPFDestino.Text = cboCCProvisaoIRPFDestino.BoundText
    End If
End Sub

Private Sub txtCCProvisaoIRPJOrigem_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoIRPJOrigem)
End Sub

Private Sub txtCCProvisaoIRPJOrigem_LostFocus()
   If txtCCProvisaoIRPJOrigem.Text <> "" Then
        cboCCProvisaoIRPJOrigem.BoundText = txtCCProvisaoIRPJOrigem.Text
        If cboCCProvisaoIRPJOrigem.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoIRPJOrigem.SetFocus
        End If
    End If
End Sub

Private Sub txtCCProvisaoIRPJDestino_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoIRPJDestino)
End Sub

Private Sub txtCCProvisaoIRPJDestino_LostFocus()
   If txtCCProvisaoIRPJDestino.Text <> "" Then
        cboCCProvisaoIRPJDestino.BoundText = txtCCProvisaoIRPJDestino.Text
        If cboCCProvisaoIRPJDestino.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoIRPJDestino.SetFocus
        End If
    End If
End Sub

Private Sub cboCCProvisaoIRPJOrigem_Change()
    If cboCCProvisaoIRPJOrigem.Text <> "" Then
        datCCProvisaoIRPJOrigem.Recordset.Bookmark = cboCCProvisaoIRPJOrigem.SelectedItem
        txtCCProvisaoIRPJOrigem.Text = cboCCProvisaoIRPJOrigem.BoundText
    End If
End Sub
    
Private Sub cboCCProvisaoIRPJDestino_Change()
    If cboCCProvisaoIRPJDestino.Text <> "" Then
        datCCProvisaoIRPJDestino.Recordset.Bookmark = cboCCProvisaoIRPJDestino.SelectedItem
        txtCCProvisaoIRPJDestino.Text = cboCCProvisaoIRPJDestino.BoundText
    End If
End Sub

Private Sub txtCCProvisaoCSLLOrigem_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoCSLLOrigem)
End Sub

Private Sub txtCCProvisaoCSLLOrigem_LostFocus()
   If txtCCProvisaoCSLLOrigem.Text <> "" Then
        cboCCProvisaoCSLLOrigem.BoundText = txtCCProvisaoCSLLOrigem.Text
        If cboCCProvisaoCSLLOrigem.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoCSLLOrigem.SetFocus
        End If
    End If
End Sub

Private Sub txtCCProvisaoCSLLDestino_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoCSLLDestino)
End Sub

Private Sub txtCCProvisaoCSLLDestino_LostFocus()
   If txtCCProvisaoCSLLDestino.Text <> "" Then
        cboCCProvisaoCSLLDestino.BoundText = txtCCProvisaoCSLLDestino.Text
        If cboCCProvisaoCSLLDestino.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoCSLLDestino.SetFocus
        End If
    End If
End Sub

Private Sub cboCCProvisaoCSLLOrigem_Change()
    If cboCCProvisaoCSLLOrigem.Text <> "" Then
        datCCProvisaoCSLLOrigem.Recordset.Bookmark = cboCCProvisaoCSLLOrigem.SelectedItem
        txtCCProvisaoCSLLOrigem.Text = cboCCProvisaoCSLLOrigem.BoundText
    End If
End Sub
    
Private Sub cboCCProvisaoCSLLDestino_Change()
    If cboCCProvisaoCSLLDestino.Text <> "" Then
        datCCProvisaoCSLLDestino.Recordset.Bookmark = cboCCProvisaoCSLLDestino.SelectedItem
        txtCCProvisaoCSLLDestino.Text = cboCCProvisaoCSLLDestino.BoundText
    End If
End Sub

Private Sub txtCCProvisaoISSOrigem_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoISSOrigem)
End Sub

Private Sub txtCCProvisaoISSOrigem_LostFocus()
   If txtCCProvisaoISSOrigem.Text <> "" Then
        cboCCProvisaoISSOrigem.BoundText = txtCCProvisaoISSOrigem.Text
        If cboCCProvisaoISSOrigem.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoISSOrigem.SetFocus
        End If
    End If
End Sub

Private Sub txtCCProvisaoISSDestino_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoISSDestino)
End Sub

Private Sub txtCCProvisaoISSDestino_LostFocus()
   If txtCCProvisaoISSDestino.Text <> "" Then
        cboCCProvisaoISSDestino.BoundText = txtCCProvisaoISSDestino.Text
        If cboCCProvisaoISSDestino.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoISSDestino.SetFocus
        End If
    End If
End Sub

Private Sub cboCCProvisaoISSOrigem_Change()
    If cboCCProvisaoISSOrigem.Text <> "" Then
        datCCProvisaoISSOrigem.Recordset.Bookmark = cboCCProvisaoISSOrigem.SelectedItem
        txtCCProvisaoISSOrigem.Text = cboCCProvisaoISSOrigem.BoundText
    End If
End Sub
    
Private Sub cboCCProvisaoISSDestino_Change()
    If cboCCProvisaoISSDestino.Text <> "" Then
        datCCProvisaoISSDestino.Recordset.Bookmark = cboCCProvisaoISSDestino.SelectedItem
        txtCCProvisaoISSDestino.Text = cboCCProvisaoISSDestino.BoundText
    End If
End Sub
Private Sub txtCCProvisaoCPPOrigem_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoCPPOrigem)
End Sub

Private Sub txtCCProvisaoCPPOrigem_LostFocus()
   If txtCCProvisaoCPPOrigem.Text <> "" Then
        cboCCProvisaoCPPOrigem.BoundText = txtCCProvisaoCPPOrigem.Text
        If cboCCProvisaoCPPOrigem.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoCPPOrigem.SetFocus
        End If
    End If
End Sub

Private Sub txtCCProvisaoCPPDestino_GotFocus()
    Call subSelecionaTXT(txtCCProvisaoCPPDestino)
End Sub

Private Sub txtCCProvisaoCPPDestino_LostFocus()
   If txtCCProvisaoCPPDestino.Text <> "" Then
        cboCCProvisaoCPPDestino.BoundText = txtCCProvisaoCPPDestino.Text
        If cboCCProvisaoCPPDestino.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            txtCCProvisaoCPPDestino.SetFocus
        End If
    End If
End Sub

Private Sub cboCCProvisaoCPPOrigem_Change()
    If cboCCProvisaoCPPOrigem.Text <> "" Then
        datCCProvisaoCPPOrigem.Recordset.Bookmark = cboCCProvisaoCPPOrigem.SelectedItem
        txtCCProvisaoCPPOrigem.Text = cboCCProvisaoCPPOrigem.BoundText
    End If
End Sub
    
Private Sub cboCCProvisaoCPPDestino_Change()
    If cboCCProvisaoCPPDestino.Text <> "" Then
        datCCProvisaoCPPDestino.Recordset.Bookmark = cboCCProvisaoCPPDestino.SelectedItem
        txtCCProvisaoCPPDestino.Text = cboCCProvisaoCPPDestino.BoundText
    End If
End Sub

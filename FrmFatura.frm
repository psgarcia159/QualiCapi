VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmFatura 
   Caption         =   "Composição de Fatura"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   1380
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   9420
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab TabFatura 
      Height          =   4755
      Left            =   0
      TabIndex        =   31
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8387
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Dados Gerais"
      TabPicture(0)   =   "FrmFatura.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "PanConjunto"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Composição de Duplicata"
      TabPicture(1)   =   "FrmFatura.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblTotalDupl"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LblValorTotal"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TDBGrid2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "CmdExcluiduplicata"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "CmdAlteraduplicata"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "CmdInsereduplicata"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "PnlCompDuplicata"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Descontos"
      TabPicture(2)   =   "FrmFatura.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TDBGrid3"
      Tab(2).Control(1)=   "pandescontos"
      Tab(2).Control(2)=   "CmdExcluiDesconto"
      Tab(2).Control(3)=   "CmdInsereDesconto"
      Tab(2).Control(4)=   "CmdAlteraDesconto"
      Tab(2).ControlCount=   5
      Begin Threed.SSPanel PnlCompDuplicata 
         Height          =   2775
         Left            =   -74460
         TabIndex        =   42
         Top             =   1020
         Visible         =   0   'False
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   4895
         _StockProps     =   15
         ForeColor       =   16711680
         BackColor       =   14737632
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
         Alignment       =   0
         Enabled         =   0   'False
         Begin MSAdodcLib.Adodc DatBancoFornecedor 
            Height          =   330
            Left            =   2175
            Top             =   1680
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
            Caption         =   "DatBancoFornecedor"
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
            Left            =   2565
            Top             =   1185
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
         Begin VB.CommandButton CmdConfirmaduplicata 
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
            Left            =   5700
            TabIndex        =   15
            Top             =   2310
            Width           =   1140
         End
         Begin VB.CommandButton CmdSairDuplicata 
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
            Left            =   7005
            TabIndex        =   16
            Top             =   2310
            Width           =   1140
         End
         Begin Threed.SSPanel PanPagto2 
            Height          =   1545
            Left            =   45
            TabIndex        =   43
            Top             =   135
            Width           =   8145
            _Version        =   65536
            _ExtentX        =   14367
            _ExtentY        =   2725
            _StockProps     =   15
            BackColor       =   14737632
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
            Begin VB.TextBox TxtCCorrente2 
               Height          =   330
               Left            =   1860
               MaxLength       =   4
               TabIndex        =   10
               Top             =   1050
               Width           =   450
            End
            Begin VB.ComboBox CboTipoPagto1 
               Height          =   315
               ItemData        =   "FrmFatura.frx":0054
               Left            =   1860
               List            =   "FrmFatura.frx":006D
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   570
               Width           =   3000
            End
            Begin VB.ComboBox CboStatusduplicata 
               Height          =   315
               ItemData        =   "FrmFatura.frx":00C7
               Left            =   6180
               List            =   "FrmFatura.frx":00D1
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   105
               Width           =   1935
            End
            Begin VB.ComboBox CboTipoPagto2 
               Height          =   315
               ItemData        =   "FrmFatura.frx":00EF
               Left            =   1875
               List            =   "FrmFatura.frx":0108
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   570
               Width           =   3000
            End
            Begin VB.TextBox TxtValorDuplicata 
               Height          =   315
               Left            =   3465
               TabIndex        =   6
               Top             =   105
               Width           =   1605
            End
            Begin VB.TextBox TxtDuplicata 
               Height          =   315
               Left            =   1860
               MaxLength       =   1
               TabIndex        =   5
               Top             =   105
               Width           =   630
            End
            Begin VB.TextBox TxtDocPagto2 
               Height          =   315
               Left            =   6600
               MaxLength       =   7
               TabIndex        =   9
               Top             =   570
               Width           =   1515
            End
            Begin MSDataListLib.DataCombo CboContaCorrente 
               Bindings        =   "FrmFatura.frx":0162
               Height          =   330
               Left            =   2340
               TabIndex        =   11
               Top             =   1050
               Width           =   5805
               _ExtentX        =   10239
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
               Height          =   315
               Left            =   5355
               TabIndex        =   49
               Top             =   120
               Width           =   765
            End
            Begin VB.Label LblCCorrente 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Conta Corrente:"
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
               Height          =   315
               Left            =   450
               TabIndex        =   48
               Top             =   1080
               Width           =   1365
            End
            Begin VB.Label LblTipoPagamento 
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
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   90
               TabIndex        =   47
               Top             =   600
               Width           =   1725
            End
            Begin VB.Label LblValorDupl 
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   2865
               TabIndex        =   46
               Top             =   120
               Width           =   615
            End
            Begin VB.Label LblCodDupl 
               BackStyle       =   0  'Transparent
               Caption         =   "Código:"
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
               Height          =   255
               Left            =   1140
               TabIndex        =   45
               Top             =   135
               Width           =   720
            End
            Begin VB.Label LblDocPagRec 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Doc. Pagamento:"
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
               Height          =   255
               Left            =   4740
               TabIndex        =   44
               Top             =   585
               Width           =   1815
            End
         End
         Begin Threed.SSPanel PanDatas 
            Height          =   495
            Left            =   120
            TabIndex        =   50
            Top             =   2160
            Width           =   5400
            _Version        =   65536
            _ExtentX        =   9525
            _ExtentY        =   873
            _StockProps     =   15
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin MSComCtl2.DTPicker DtpVencimento 
               Height          =   315
               Left            =   1215
               TabIndex        =   13
               Top             =   90
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               DateIsNull      =   -1  'True
               Format          =   81788929
               CurrentDate     =   37460
            End
            Begin MSComCtl2.DTPicker DtpPagamento 
               Height          =   315
               Left            =   3795
               TabIndex        =   14
               Top             =   105
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   81788929
               CurrentDate     =   37460
            End
            Begin VB.Label LblPagamento 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   2520
               TabIndex        =   52
               Top             =   120
               Width           =   1245
            End
            Begin VB.Label LblVencimento 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Vencimento:"
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
               Height          =   315
               Left            =   60
               TabIndex        =   51
               Top             =   120
               Width           =   1125
            End
         End
         Begin MSDataListLib.DataCombo CboBancoForn2 
            Bindings        =   "FrmFatura.frx":0181
            Height          =   330
            Left            =   1905
            TabIndex        =   12
            Top             =   1680
            Width           =   6285
            _ExtentX        =   11086
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "desccombo"
            BoundColumn     =   "bafo_cd_codigo"
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
         Begin VB.Label LblBcFornecedor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Banco Fornecedor:"
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
            Height          =   315
            Left            =   135
            TabIndex        =   53
            Top             =   1695
            Width           =   1725
         End
      End
      Begin VB.CommandButton CmdAlteraDesconto 
         Caption         =   "A&lterar"
         Height          =   330
         Left            =   -67080
         TabIndex        =   27
         Top             =   1350
         Width           =   1080
      End
      Begin VB.CommandButton CmdInsereDesconto 
         Caption         =   "Ins&erir"
         Height          =   330
         Left            =   -67080
         TabIndex        =   26
         Top             =   960
         Width           =   1080
      End
      Begin VB.CommandButton CmdExcluiDesconto 
         Caption         =   "E&xcluir"
         Height          =   330
         Left            =   -67080
         TabIndex        =   28
         Top             =   1740
         Width           =   1080
      End
      Begin VB.CommandButton CmdInsereduplicata 
         Caption         =   "I&nserir"
         Height          =   330
         Left            =   -74805
         TabIndex        =   18
         Top             =   4230
         Width           =   1080
      End
      Begin VB.CommandButton CmdAlteraduplicata 
         Caption         =   "&Alterar"
         Height          =   330
         Left            =   -73560
         TabIndex        =   19
         Top             =   4230
         Width           =   1080
      End
      Begin VB.CommandButton CmdExcluiduplicata 
         Caption         =   "&Excluir"
         Height          =   330
         Left            =   -72285
         TabIndex        =   20
         Top             =   4230
         Width           =   1080
      End
      Begin Threed.SSPanel PanConjunto 
         Height          =   1065
         Left            =   240
         TabIndex        =   34
         Top             =   480
         Width           =   8895
         _Version        =   65536
         _ExtentX        =   15690
         _ExtentY        =   1879
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Begin VB.Frame FraTipoMovimentacao 
            Caption         =   "Tipo de Movimentação:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   120
            TabIndex        =   0
            Top             =   120
            Width           =   2295
            Begin VB.OptionButton OptSaida 
               Caption         =   "A Pagar"
               Height          =   200
               Left            =   1320
               TabIndex        =   2
               Top             =   420
               Width           =   870
            End
            Begin VB.OptionButton OptEntrada 
               Caption         =   "A Receber"
               Height          =   200
               Left            =   120
               TabIndex        =   1
               Top             =   400
               Width           =   1215
            End
         End
         Begin MSAdodcLib.Adodc DatFornecedor 
            Height          =   330
            Left            =   4920
            Top             =   240
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
            Caption         =   "DatFornecedor"
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
         Begin MSDataListLib.DataCombo CboFornecedor 
            Bindings        =   "FrmFatura.frx":01A2
            Height          =   315
            Left            =   3960
            TabIndex        =   3
            Top             =   240
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "focl_tx_fantasia"
            BoundColumn     =   "focl_cd_forncli"
            Text            =   ""
         End
         Begin VB.Label LblValor 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3960
            TabIndex        =   54
            Top             =   630
            Width           =   1710
         End
         Begin VB.Label LblTotalFatura 
            Caption         =   "Total da Fatura:"
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
            Left            =   2520
            TabIndex        =   36
            Top             =   675
            Width           =   1395
         End
         Begin VB.Label Lblfornecedor 
            Alignment       =   1  'Right Justify
            Caption         =   "Fornecedor:"
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
            Left            =   2835
            TabIndex        =   35
            Top             =   240
            Width           =   1080
         End
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Height          =   2895
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   5106
         _LayoutType     =   4
         _RowHeight      =   20
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "CodigoNF"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nº Documento"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Série"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Emissão"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "dd/mm/yy"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Valor"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   80
         Columns(5)._MaxComboItems=   5
         Columns(5).ValueItems(0)._DefaultItem=   0
         Columns(5).ValueItems(0).Value=   "N"
         Columns(5).ValueItems(0).Value.vt=   8
         Columns(5).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(5).ValueItems(0).DisplayValue(0)=   "bHQAACYEAABCTSYEAAAAAAAANgAAACgAAAASAAAAEgAAAAEAGAAAAAAA8AMAAAAAAAAAAAAAAAAA"
         Columns(5).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
         Columns(5).ValueItems(0).DisplayValue(2)=   "//8AAP//////////////////////////////////////////////////////////////////////"
         Columns(5).ValueItems(0).DisplayValue(3)=   "/wAA////////////////////////////////////////////////////////////////////////"
         Columns(5).ValueItems(0).DisplayValue(4)=   "AAD///////////////////////////////////////////////////////////////////////8A"
         Columns(5).ValueItems(0).DisplayValue(5)=   "AP///////////4SGhISGhISGhISGhISGhISGhISGhISGhISGhISGhISGhP///////////////wAA"
         Columns(5).ValueItems(0).DisplayValue(6)=   "////////////hIaE////////////////////////////////////hIaE////////////////AAD/"
         Columns(5).ValueItems(0).DisplayValue(7)=   "//////////+EhoT///////////////////////////////////+EhoT///////////////8AAP//"
         Columns(5).ValueItems(0).DisplayValue(8)=   "/////////4SGhP///////////////////////////////////4SGhP///////////////wAA////"
         Columns(5).ValueItems(0).DisplayValue(9)=   "////////hIaE////////////////////////////////////hIaE////////////////AAD/////"
         Columns(5).ValueItems(0).DisplayValue(10)=   "//////+EhoT///////////////////////////////////+EhoT///////////////8AAP//////"
         Columns(5).ValueItems(0).DisplayValue(11)=   "/////4SGhP///////////////////////////////////4SGhP///////////////wAA////////"
         Columns(5).ValueItems(0).DisplayValue(12)=   "////hIaE////////////////////////////////////hIaE////////////////AAD/////////"
         Columns(5).ValueItems(0).DisplayValue(13)=   "//+EhoT///////////////////////////////////+EhoT///////////////8AAP//////////"
         Columns(5).ValueItems(0).DisplayValue(14)=   "/4SGhP///////////////////////////////////4SGhP///////////////wAA////////////"
         Columns(5).ValueItems(0).DisplayValue(15)=   "hIaEhIaEhIaEhIaEhIaEhIaEhIaEhIaEhIaEhIaEhIaE////////////////AAD/////////////"
         Columns(5).ValueItems(0).DisplayValue(16)=   "//////////////////////////////////////////////////////////8AAP//////////////"
         Columns(5).ValueItems(0).DisplayValue(17)=   "/////////////////////////////////////////////////////////wAA////////////////"
         Columns(5).ValueItems(0).DisplayValue(18)=   "////////////////////////////////////////////////////////AAA="
         Columns(5).ValueItems(0).DisplayValue.vt=   9
         Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(1)._DefaultItem=   0
         Columns(5).ValueItems(1).Value=   "S"
         Columns(5).ValueItems(1).Value.vt=   8
         Columns(5).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(5).ValueItems(1).DisplayValue(0)=   "bHQAABIEAABCTRIEAAAAAAAANgAAACgAAAARAAAAEwAAAAEAGAAAAAAA3AMAAAAAAAAAAAAAAAAA"
         Columns(5).ValueItems(1).DisplayValue(1)=   "AAAAAAD///////////////////////////////////////////////////////////////////8A"
         Columns(5).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////AP//////"
         Columns(5).ValueItems(1).DisplayValue(3)=   "/////////////////////////////////////////////////////////////wD/////////////"
         Columns(5).ValueItems(1).DisplayValue(4)=   "//////////////////////////////////////////////////////8A////////////hIaEhIaE"
         Columns(5).ValueItems(1).DisplayValue(5)=   "hIaEhIaEhIaEhIaEhIaEhIaEhIaEhIaEhIaE////////////AP///////////4SGhP//////////"
         Columns(5).ValueItems(1).DisplayValue(6)=   "/////////////////////////4SGhP///////////wD///////////+EhoT///////////8AAAD/"
         Columns(5).ValueItems(1).DisplayValue(7)=   "//////////////////+EhoT///////////8A////////////hIaE////////AAAAAAAAAAAA////"
         Columns(5).ValueItems(1).DisplayValue(8)=   "////////////hIaE////////////AP///////////4SGhP///wAAAAAAAAAAAAAAAAAAAP//////"
         Columns(5).ValueItems(1).DisplayValue(9)=   "/////4SGhP///////////wD///////////+EhoT///8AAAAAAAD///8AAAAAAAAAAAD///////+E"
         Columns(5).ValueItems(1).DisplayValue(10)=   "hoT///////////8A////////////hIaE////AAAA////////////AAAAAAAAAAAA////hIaE////"
         Columns(5).ValueItems(1).DisplayValue(11)=   "////////AP///////////4SGhP///////////////////////wAAAAAAAP///4SGhP//////////"
         Columns(5).ValueItems(1).DisplayValue(12)=   "/wD///////////+EhoT///////////////////////////8AAAD///+EhoT///////////8A////"
         Columns(5).ValueItems(1).DisplayValue(13)=   "////////hIaE////////////////////////////////////hIaE////////////AP//////////"
         Columns(5).ValueItems(1).DisplayValue(14)=   "/4SGhISGhISGhISGhISGhISGhISGhISGhISGhISGhISGhP///////////wD/////////////////"
         Columns(5).ValueItems(1).DisplayValue(15)=   "//////////////////////////////////////////////////8A////////////////////////"
         Columns(5).ValueItems(1).DisplayValue(16)=   "////////////////////////////////////////////AP//////////////////////////////"
         Columns(5).ValueItems(1).DisplayValue(17)=   "/////////////////////////////////////wD/////////////////////////////////////"
         Columns(5).ValueItems(1).DisplayValue(18)=   "//////////////////////////////8A"
         Columns(5).ValueItems(1).DisplayValue.vt=   9
         Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems.Count=   2
         Columns(5).Caption=   "Selecionar"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Status"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Fatura"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "tido_cd_tipodoc"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
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
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=1920300131"
         Splits(0)._ColumnProps(10)=   "Column(1).Width=2672"
         Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=2593"
         Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=1049104"
         Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(16)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(17)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(18)=   "Column(1)._MinWidth=-1"
         Splits(0)._ColumnProps(19)=   "Column(2).Width=2831"
         Splits(0)._ColumnProps(20)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._WidthInPix=2752"
         Splits(0)._ColumnProps(22)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(2)._ColStyle=1049104"
         Splits(0)._ColumnProps(24)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(25)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(26)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(27)=   "Column(2)._MinWidth=81614592"
         Splits(0)._ColumnProps(28)=   "Column(3).Width=3096"
         Splits(0)._ColumnProps(29)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(3)._WidthInPix=3016"
         Splits(0)._ColumnProps(31)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(32)=   "Column(3)._ColStyle=1049105"
         Splits(0)._ColumnProps(33)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(34)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(35)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(36)=   "Column(3)._MinWidth=1851880448"
         Splits(0)._ColumnProps(37)=   "Column(4).Width=3254"
         Splits(0)._ColumnProps(38)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(40)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(4)._ColStyle=1049362"
         Splits(0)._ColumnProps(42)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(43)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(44)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(45)=   "Column(5).Width=2355"
         Splits(0)._ColumnProps(46)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(5)._WidthInPix=2275"
         Splits(0)._ColumnProps(48)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(49)=   "Column(5)._ColStyle=1049105"
         Splits(0)._ColumnProps(50)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(51)=   "Column(5).AllowFocus=0"
         Splits(0)._ColumnProps(52)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(53)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(54)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(56)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(57)=   "Column(6)._ColStyle=1048848"
         Splits(0)._ColumnProps(58)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(59)=   "Column(6).WrapText=1"
         Splits(0)._ColumnProps(60)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(61)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(62)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(64)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(65)=   "Column(7)._ColStyle=1048848"
         Splits(0)._ColumnProps(66)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(67)=   "Column(7).WrapText=1"
         Splits(0)._ColumnProps(68)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(69)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(70)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(72)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(73)=   "Column(8)._ColStyle=1048848"
         Splits(0)._ColumnProps(74)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(75)=   "Column(8).WrapText=1"
         Splits(0)._ColumnProps(76)=   "Column(8).Order=9"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
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
         _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.valignment=2,.wraptext=-1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=44,.parent=2,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(27)  =   ":id=44,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(28)  =   ":id=44,.fontname=MS Sans Serif"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=2"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=2,.valignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=2"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=2,.valignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44,.alignment=2"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
         _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=16,.parent=43"
         _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=13,.parent=44"
         _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=14,.parent=45"
         _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=15,.parent=47"
         _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=20,.parent=43"
         _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=17,.parent=44"
         _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=18,.parent=45"
         _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=19,.parent=47"
         _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=24,.parent=43"
         _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=21,.parent=44"
         _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=22,.parent=45"
         _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=23,.parent=47"
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
      Begin Threed.SSPanel pandescontos 
         Height          =   1695
         Left            =   -74280
         TabIndex        =   37
         Top             =   1560
         Visible         =   0   'False
         Width           =   6855
         _Version        =   65536
         _ExtentX        =   12091
         _ExtentY        =   2990
         _StockProps     =   15
         ForeColor       =   16711680
         BackColor       =   14737632
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
         Alignment       =   0
         Enabled         =   0   'False
         Begin VB.CommandButton CmdConfirmaDesconto 
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
            Left            =   4080
            TabIndex        =   24
            Top             =   1200
            Width           =   1140
         End
         Begin VB.CommandButton CmdSairDesconto 
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
            Left            =   5400
            TabIndex        =   25
            Top             =   1200
            Width           =   1140
         End
         Begin VB.TextBox TxtValorDesconto 
            Height          =   315
            Left            =   4800
            TabIndex        =   23
            Top             =   720
            Width           =   1755
         End
         Begin MSAdodcLib.Adodc DatDescontosNF 
            Height          =   330
            Left            =   2400
            Top             =   240
            Visible         =   0   'False
            Width           =   2535
            _ExtentX        =   4471
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
            Caption         =   "DatDescontosNF"
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
         Begin MSDataListLib.DataCombo CboDescontos 
            Bindings        =   "FrmFatura.frx":01BE
            Height          =   315
            Left            =   1560
            TabIndex        =   21
            Top             =   240
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "DescricaoCombo"
            BoundColumn     =   "desc_cd_desconto"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker DtpVencDesconto 
            Height          =   315
            Left            =   1560
            TabIndex        =   22
            Top             =   720
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   81788929
            CurrentDate     =   37460
         End
         Begin VB.Label LblVlDesconto 
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4200
            TabIndex        =   40
            Top             =   735
            Width           =   975
         End
         Begin VB.Label LblDescontos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Descontos:"
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
            Height          =   315
            Left            =   360
            TabIndex        =   39
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label LblVencDesconto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimento:"
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
            Height          =   315
            Left            =   360
            TabIndex        =   38
            Top             =   735
            Width           =   1125
         End
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid3 
         Height          =   3135
         Left            =   -74280
         TabIndex        =   41
         Top             =   960
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5530
         _LayoutType     =   4
         _RowHeight      =   12
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).ValueItems(0)._DefaultItem=   0
         Columns(0).ValueItems(0).Value=   ""
         Columns(0).ValueItems(0).Value.vt=   8
         Columns(0).ValueItems(0).DisplayValue=   ""
         Columns(0).ValueItems(0).DisplayValue.vt=   8
         Columns(0).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(0).ValueItems.Count=   1
         Columns(0).Caption=   "Descrição"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Vencimento"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "dd/mm/yy"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Valor"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "Currency"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "codigo"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Fatura"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Recolhimento"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "NumDarf"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=6112"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6033"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1057296"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=-2147483630"
         Splits(0)._ColumnProps(10)=   "Column(1).Width=2619"
         Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=2540"
         Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=1049105"
         Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(16)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(19)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=1049362"
         Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=2064"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=1984"
         Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=1048848"
         Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(29)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=1048848"
         Splits(0)._ColumnProps(34)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(36)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=1048832"
         Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(42)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(45)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=1048832"
         Splits(0)._ColumnProps(47)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(48)=   "Column(6).Order=7"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=52,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(18)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(31)  =   ":id=28,.wraptext=-1,.locked=-1"
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=2,.valignment=2"
         _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=1"
         _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(43)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(44)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44"
         _StyleDefs(45)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(46)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
         _StyleDefs(47)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(48)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
         _StyleDefs(49)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
         _StyleDefs(50)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
         _StyleDefs(51)  =   "Splits(0).Columns(5).Style:id=70,.parent=43"
         _StyleDefs(52)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44"
         _StyleDefs(53)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
         _StyleDefs(54)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
         _StyleDefs(55)  =   "Splits(0).Columns(6).Style:id=16,.parent=43"
         _StyleDefs(56)  =   "Splits(0).Columns(6).HeadingStyle:id=13,.parent=44"
         _StyleDefs(57)  =   "Splits(0).Columns(6).FooterStyle:id=14,.parent=45"
         _StyleDefs(58)  =   "Splits(0).Columns(6).EditorStyle:id=15,.parent=47"
         _StyleDefs(59)  =   "Named:id=33:Normal"
         _StyleDefs(60)  =   ":id=33,.parent=0"
         _StyleDefs(61)  =   "Named:id=34:Heading"
         _StyleDefs(62)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   ":id=34,.wraptext=-1"
         _StyleDefs(64)  =   "Named:id=35:Footing"
         _StyleDefs(65)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(66)  =   "Named:id=36:Selected"
         _StyleDefs(67)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(68)  =   "Named:id=37:Caption"
         _StyleDefs(69)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(70)  =   "Named:id=38:HighlightRow"
         _StyleDefs(71)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(72)  =   "Named:id=39:EvenRow"
         _StyleDefs(73)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(74)  =   "Named:id=40:OddRow"
         _StyleDefs(75)  =   ":id=40,.parent=33"
         _StyleDefs(76)  =   "Named:id=41:RecordSelector"
         _StyleDefs(77)  =   ":id=41,.parent=34"
         _StyleDefs(78)  =   "Named:id=42:FilterBar"
         _StyleDefs(79)  =   ":id=42,.parent=33"
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid2 
         Height          =   3255
         Left            =   -74805
         TabIndex        =   55
         Top             =   750
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   5741
         _LayoutType     =   4
         _RowHeight      =   27
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Vencimento"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "dd/mm/yy"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Valor"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "Standard"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "CodStatus"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Status"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "CodTipoPagto"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Tipo Pag."
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Doc. Pag."
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Data Pag."
         Columns(8).DataField=   ""
         Columns(8).NumberFormat=   "dd/mm/yy"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "CodConta"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Conta Corrente (Empresa)"
         Columns(10).DataField=   ""
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "CodBancoforn"
         Columns(11).DataField=   ""
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Conta Corrente (Fornecedor)"
         Columns(12).DataField=   ""
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Fatura"
         Columns(13).DataField=   ""
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Cód. Duplicata"
         Columns(14).DataField=   ""
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   15
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColMove=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=15"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1058"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=979"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1057296"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=1693"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=1614"
         Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=1057297"
         Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(17)=   "Column(2).Width=2275"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2196"
         Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=1057554"
         Splits(0)._ColumnProps(22)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(23)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(25)=   "Column(2)._MinWidth=1953261894"
         Splits(0)._ColumnProps(26)=   "Column(3).Width=2064"
         Splits(0)._ColumnProps(27)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(3)._WidthInPix=1984"
         Splits(0)._ColumnProps(29)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(3)._ColStyle=1049362"
         Splits(0)._ColumnProps(31)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(32)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(33)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(34)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(35)=   "Column(4).Width=2143"
         Splits(0)._ColumnProps(36)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(4)._WidthInPix=2064"
         Splits(0)._ColumnProps(38)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(39)=   "Column(4)._ColStyle=1057296"
         Splits(0)._ColumnProps(40)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(41)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(42)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(43)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(44)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(46)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(5)._ColStyle=1048848"
         Splits(0)._ColumnProps(48)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(49)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(50)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(51)=   "Column(6).Width=1984"
         Splits(0)._ColumnProps(52)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(53)=   "Column(6)._WidthInPix=1905"
         Splits(0)._ColumnProps(54)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(55)=   "Column(6)._ColStyle=1057296"
         Splits(0)._ColumnProps(56)=   "Column(6).WrapText=1"
         Splits(0)._ColumnProps(57)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(58)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(59)=   "Column(7).Width=1720"
         Splits(0)._ColumnProps(60)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(61)=   "Column(7)._WidthInPix=1640"
         Splits(0)._ColumnProps(62)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(63)=   "Column(7)._ColStyle=1057296"
         Splits(0)._ColumnProps(64)=   "Column(7).WrapText=1"
         Splits(0)._ColumnProps(65)=   "Column(7).AllowFocus=0"
         Splits(0)._ColumnProps(66)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(67)=   "Column(8).Width=1535"
         Splits(0)._ColumnProps(68)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(8)._WidthInPix=1455"
         Splits(0)._ColumnProps(70)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(71)=   "Column(8)._ColStyle=1057297"
         Splits(0)._ColumnProps(72)=   "Column(8).WrapText=1"
         Splits(0)._ColumnProps(73)=   "Column(8).AllowFocus=0"
         Splits(0)._ColumnProps(74)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(75)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(76)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(77)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(78)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(79)=   "Column(9)._ColStyle=1048848"
         Splits(0)._ColumnProps(80)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(81)=   "Column(9).WrapText=1"
         Splits(0)._ColumnProps(82)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(83)=   "Column(10).Width=3863"
         Splits(0)._ColumnProps(84)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(85)=   "Column(10)._WidthInPix=3784"
         Splits(0)._ColumnProps(86)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(87)=   "Column(10)._ColStyle=1057296"
         Splits(0)._ColumnProps(88)=   "Column(10).WrapText=1"
         Splits(0)._ColumnProps(89)=   "Column(10).AllowFocus=0"
         Splits(0)._ColumnProps(90)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(91)=   "Column(11).Width=1270"
         Splits(0)._ColumnProps(92)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(93)=   "Column(11)._WidthInPix=1191"
         Splits(0)._ColumnProps(94)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(95)=   "Column(11)._ColStyle=1048848"
         Splits(0)._ColumnProps(96)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(97)=   "Column(11).WrapText=1"
         Splits(0)._ColumnProps(98)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(99)=   "Column(12).Width=3916"
         Splits(0)._ColumnProps(100)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(101)=   "Column(12)._WidthInPix=3836"
         Splits(0)._ColumnProps(102)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(103)=   "Column(12)._ColStyle=1057296"
         Splits(0)._ColumnProps(104)=   "Column(12).WrapText=1"
         Splits(0)._ColumnProps(105)=   "Column(12).AllowFocus=0"
         Splits(0)._ColumnProps(106)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(107)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(108)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(109)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(110)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(111)=   "Column(13)._ColStyle=1048848"
         Splits(0)._ColumnProps(112)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(113)=   "Column(13).WrapText=1"
         Splits(0)._ColumnProps(114)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(115)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(116)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(117)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(118)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(119)=   "Column(14)._ColStyle=1048848"
         Splits(0)._ColumnProps(120)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(121)=   "Column(14).WrapText=1"
         Splits(0)._ColumnProps(122)=   "Column(14).Order=15"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=196,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(37)  =   ":id=28,.locked=-1"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=2,.valignment=2"
         _StyleDefs(42)  =   ":id=32,.locked=-1"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(47)  =   ":id=58,.locked=-1"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=1,.valignment=2"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=1,.locked=0"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47,.locked=0"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(56)  =   ":id=66,.locked=-1"
         _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=2"
         _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
         _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
         _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=70,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=44"
         _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=45"
         _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=47"
         _StyleDefs(64)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(65)  =   ":id=74,.locked=-1"
         _StyleDefs(66)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44,.alignment=2"
         _StyleDefs(67)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
         _StyleDefs(68)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
         _StyleDefs(69)  =   "Splits(0).Columns(7).Style:id=78,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(70)  =   ":id=78,.locked=-1"
         _StyleDefs(71)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=44,.alignment=2"
         _StyleDefs(72)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=45"
         _StyleDefs(73)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=47"
         _StyleDefs(74)  =   "Splits(0).Columns(8).Style:id=82,.parent=43,.alignment=2,.valignment=2"
         _StyleDefs(75)  =   ":id=82,.locked=-1"
         _StyleDefs(76)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=44,.alignment=2"
         _StyleDefs(77)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=45"
         _StyleDefs(78)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=47"
         _StyleDefs(79)  =   "Splits(0).Columns(9).Style:id=86,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(80)  =   "Splits(0).Columns(9).HeadingStyle:id=83,.parent=44"
         _StyleDefs(81)  =   "Splits(0).Columns(9).FooterStyle:id=84,.parent=45"
         _StyleDefs(82)  =   "Splits(0).Columns(9).EditorStyle:id=85,.parent=47"
         _StyleDefs(83)  =   "Splits(0).Columns(10).Style:id=90,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(84)  =   ":id=90,.locked=-1"
         _StyleDefs(85)  =   "Splits(0).Columns(10).HeadingStyle:id=87,.parent=44,.alignment=2"
         _StyleDefs(86)  =   "Splits(0).Columns(10).FooterStyle:id=88,.parent=45"
         _StyleDefs(87)  =   "Splits(0).Columns(10).EditorStyle:id=89,.parent=47"
         _StyleDefs(88)  =   "Splits(0).Columns(11).Style:id=94,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(89)  =   "Splits(0).Columns(11).HeadingStyle:id=91,.parent=44"
         _StyleDefs(90)  =   "Splits(0).Columns(11).FooterStyle:id=92,.parent=45"
         _StyleDefs(91)  =   "Splits(0).Columns(11).EditorStyle:id=93,.parent=47"
         _StyleDefs(92)  =   "Splits(0).Columns(12).Style:id=98,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(93)  =   ":id=98,.locked=-1"
         _StyleDefs(94)  =   "Splits(0).Columns(12).HeadingStyle:id=95,.parent=44,.alignment=2"
         _StyleDefs(95)  =   "Splits(0).Columns(12).FooterStyle:id=96,.parent=45"
         _StyleDefs(96)  =   "Splits(0).Columns(12).EditorStyle:id=97,.parent=47"
         _StyleDefs(97)  =   "Splits(0).Columns(13).Style:id=16,.parent=43"
         _StyleDefs(98)  =   "Splits(0).Columns(13).HeadingStyle:id=13,.parent=44"
         _StyleDefs(99)  =   "Splits(0).Columns(13).FooterStyle:id=14,.parent=45"
         _StyleDefs(100) =   "Splits(0).Columns(13).EditorStyle:id=15,.parent=47"
         _StyleDefs(101) =   "Splits(0).Columns(14).Style:id=102,.parent=43,.alignment=0,.valignment=2"
         _StyleDefs(102) =   "Splits(0).Columns(14).HeadingStyle:id=99,.parent=44"
         _StyleDefs(103) =   "Splits(0).Columns(14).FooterStyle:id=100,.parent=45"
         _StyleDefs(104) =   "Splits(0).Columns(14).EditorStyle:id=101,.parent=47"
         _StyleDefs(105) =   "Named:id=33:Normal"
         _StyleDefs(106) =   ":id=33,.parent=0"
         _StyleDefs(107) =   "Named:id=34:Heading"
         _StyleDefs(108) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(109) =   ":id=34,.wraptext=-1"
         _StyleDefs(110) =   "Named:id=35:Footing"
         _StyleDefs(111) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(112) =   "Named:id=36:Selected"
         _StyleDefs(113) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(114) =   "Named:id=37:Caption"
         _StyleDefs(115) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(116) =   "Named:id=38:HighlightRow"
         _StyleDefs(117) =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(118) =   "Named:id=39:EvenRow"
         _StyleDefs(119) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(120) =   "Named:id=40:OddRow"
         _StyleDefs(121) =   ":id=40,.parent=33"
         _StyleDefs(122) =   "Named:id=41:RecordSelector"
         _StyleDefs(123) =   ":id=41,.parent=34"
         _StyleDefs(124) =   "Named:id=42:FilterBar"
         _StyleDefs(125) =   ":id=42,.parent=33"
      End
      Begin VB.Label LblValorTotal 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -67695
         TabIndex        =   33
         Top             =   4230
         Width           =   1485
      End
      Begin VB.Label LblTotalDupl 
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
         Left            =   -68775
         TabIndex        =   32
         Top             =   4275
         Width           =   1170
      End
   End
   Begin Threed.SSCommand CmdGravar 
      Height          =   330
      Left            =   2445
      TabIndex        =   29
      Top             =   4950
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
      Enabled         =   0   'False
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   5340
      TabIndex        =   30
      Top             =   4950
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
End
Attribute VB_Name = "FrmFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Reload As Boolean

Dim ValorTotalDuplicata As Double 'Acumula o total de duplicatas até compor o valor total da Nota Fiscal
Dim Xdebcred As String
Dim Fatura As Long 'gera o novo nº da fatura
Dim vazio As Boolean 'indica se o vetor com a relação das notas do fornecedor está vazio
Dim xcod As String 'Armazena o desconto escolhido na alteracao do desconto
Dim xcont As Integer

Dim PosCursor As Integer 'Guarda o Bookmark do grid
Dim XTipoFatura As String
Dim StatusNF As String  'Guarda o status das NFs
Dim XVlDuplicata As Double

'Variáveis para viabilizar utilização dos grids unbound
Dim Vetor1 As New XArray
Dim Vetor2 As New XArray
Dim Vetor3 As New XArray
Dim EntrouDesc As Boolean 'Controla a entrada na orelha de desconto
Dim EntrouDupl As Boolean 'Controla a entrada na orelha de duplicata
Dim XDuplicata As String 'Armazena o tipo de operação (A:alteração, I: inserção)do grid que será executada
Dim XDesconto As String 'Armazena o tipo de operação (A:alteração, I: inserção)do grid que será executada


Public Property Get CARREGAGRID() As Boolean
    CARREGAGRID = Reload
End Property

Function FunVerificaTipoDocumento() As Boolean
    
    Dim XLI_DOCUMENTO As Integer
    Dim XLT_SERIE  As String
    
    'XLI_DOCUMENTO = ""
    XLT_SERIE = ""
    
    TDBGrid1.MoveFirst
    
    'Pegando tipo de documento e série da primeira nota selecionada
    Do While Not Me.TDBGrid1.EOF
        If Me.TDBGrid1.Columns("Selecionar").Value = "S" Then
            XLI_DOCUMENTO = Me.TDBGrid1.Columns("tido_cd_tipodoc")
            XLT_SERIE = Me.TDBGrid1.Columns("Série")
        End If
        
       Me.TDBGrid1.MoveNext
    
    Loop
    
    TDBGrid1.MoveFirst
    
    'Verificando se todos os tipos de documento selecionados são iguais ao primeiro
    Do While Not Me.TDBGrid1.EOF
        If Me.TDBGrid1.Columns("Selecionar").Value = "S" Then
            If XLI_DOCUMENTO <> Me.TDBGrid1.Columns("tido_cd_tipodoc") Or XLT_SERIE <> Me.TDBGrid1.Columns("Série") Then
                FunVerificaTipoDocumento = False
                Exit Function
            End If
        End If
        Me.TDBGrid1.MoveNext
    Loop
    
    FunVerificaTipoDocumento = True
    
End Function
Sub subCarregaVetorGrid(Sql As String, VetorCampos As Variant, Vetor As XArray, Grid As TDBGrid)   'Carrega o vetores
    Dim ResAvaliacao As Object
    Dim Cont As Integer
    Dim X As Integer
    
    'Carrega os valores no vetor
    
    'sql - string sql com a tabela ou consulta para carregar o vetor e o grid
    'VetorCampos - vetor simples com os campos na respectiva ordem do grid
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    
    SubQOpenRecordset ResAvaliacao, Sql, Estatico
    Vetor.ReDim 0, 0, 0, Grid.Columns.Count - 1
    Vetor.Clear
    Grid.Array = Vetor
    If Not (ResAvaliacao.EOF And ResAvaliacao.BOF) Then
        ResAvaliacao.MoveFirst
        Do While Not ResAvaliacao.EOF
            If Vetor.UpperBound(1) <> 0 Or Vetor(0, 0) <> Empty Then
                Vetor.Insert 1, Vetor.UpperBound(1) + 1
            End If
            Cont = Vetor.UpperBound(1)
            For X = 0 To Grid.Columns.Count - 1
                subSetaVal Vetor, Cont, X, FunNulo(ResAvaliacao(VetorCampos(X)))
            Next
            ResAvaliacao.MoveNext
        Loop
        
        With Grid
            .ReBind
        End With
        
    End If
    ResAvaliacao.Close
End Sub
Sub subHabilitaCampos()
    If OptEntrada.Value = True Then
        Lblfornecedor.Caption = "Cliente:"
        CboFornecedor.Text = ""
        LblTipoPagamento.Caption = "Tipo Recebimento:"
        LblDocPagRec.Caption = "Doc. Recebimento:"
        LblBcFornecedor.Visible = False
        CboBancoForn2.Visible = False
        LblPagamento.Caption = "Recebimento:"
        PanPagto2.Top = 375
        TDBGrid2.Columns("Conta Corrente (Fornecedor)").Visible = False
        TDBGrid3.Columns("Vencimento").Visible = False
        LblVlDesconto.Left = LblVencDesconto.Left
        TxtValorDesconto.Left = DtpVencDesconto.Left
        LblVencDesconto.Visible = False
        DtpVencDesconto.Visible = False
    Else
        Lblfornecedor.Caption = "Fornecedor:"
        CboFornecedor.Text = ""
        LblTipoPagamento.Caption = "Tipo Pagamento:"
        LblDocPagRec.Caption = "Doc. Pagamento:"
        LblBcFornecedor.Visible = True
        CboBancoForn2.Visible = True
        LblPagamento.Caption = "Pagamento:"
        PanPagto2.Top = 135
        TDBGrid3.Columns("Vencimento").Visible = True
        LblVlDesconto.Left = 4200
        TxtValorDesconto.Left = 4800
        LblVencDesconto.Visible = True
        DtpVencDesconto.Visible = True
    End If
End Sub
Function funVerificaValores() As Boolean
    Dim SomaDesc As Double
    Dim Somadupl As Double
    Dim Sql As String
    
    SomaDesc = funSomaColuna(Vetor3, 2)
    Somadupl = funSomaColuna(Vetor2, 2)
    If Format(LblValor.Caption, "standard") <> Format((Somadupl + SomaDesc), "standard") Then
        MsgBox "A soma das Parcelas com os Descontos é diferente do valor bruto da Fatura!", vbCritical, "ATENÇÃO"
        TabFatura.Tab = 1
        If Format(LblValor.Caption, "standard") > Format((Somadupl + SomaDesc), "standard") Then
            CmdInsereduplicata_Click
            TxtValorDuplicata.Text = (CDbl(LblValor.Caption) - (Somadupl + SomaDesc))
        End If
        funVerificaValores = True
        Exit Function
    End If
End Function
Function funConsistenciaDatas() As Boolean
Dim DtEmissao As Date
'Verifica se as datas seguem a ordem de preenchimento
'Se ConsistenciaDatas= TRUE    >> Datas consistentes
'Se ConsistenciaDatas= FALSE  >> Datas não consistentes
    funConsistenciaDatas = False
    
    'obter a maior data de emissão
    If Vetor1(0, 0) <> Empty Then
       TDBGrid1.MoveFirst
       DtEmissao = TDBGrid1.Columns("Emissão").Value
       TDBGrid1.MoveNext
       Do While Not TDBGrid1.EOF
           If TDBGrid1.Columns("Selecionar").Value = "S" Then
              If DtEmissao > (TDBGrid1.Columns("Emissão").Value) Then
                 DtEmissao = TDBGrid1.Columns("Emissão").Value
              End If
           End If
           TDBGrid1.MoveNext
        Loop
    End If
    
    If TabFatura.Tab = 1 Then 'Orelha de Duplicata
        If Not IsNull(DtpVencimento.Value) Then
           If DateDiff("d", DtEmissao, DtpVencimento.Value) < 0 Then
                MsgBox "A data de vencimento é anterior a data de emissão!", vbCritical, "ATENÇÃO"
                DtpVencimento.SetFocus
                Exit Function
           End If
        End If
        If CboStatusduplicata.ListIndex = 1 Then   'Realizado
           If Not IsNull(DtpPagamento.Value) Then
                If DateDiff("d", DtEmissao, DtpPagamento.Value) < 0 Then
                    MsgBox "A data de pagamento é anterior a data de emissão!", vbCritical, "ATENÇÃO"
                    DtpPagamento.SetFocus
                    Exit Function
                End If
            End If
        End If
    ElseIf TabFatura.Tab = 2 Then 'Orelha de Descontos
        If Not IsNull(DtpVencDesconto.Value) Then
           If DateDiff("d", DtEmissao, DtpVencDesconto.Value) < 0 Then
                MsgBox "A data de vencimento é anterior a data de emissão!", vbCritical, "ATENÇÃO"
                DtpVencDesconto.SetFocus
                Exit Function
           End If
        End If
    End If
    
    funConsistenciaDatas = True
End Function


Sub subDesabilitaBotoes()
    If TabFatura.Tab = 2 Then
        CmdInsereDesconto.Enabled = False
        CmdAlteraDesconto.Enabled = False
        CmdExcluiDesconto.Enabled = False
    ElseIf TabFatura.Tab = 1 Then
        CmdInsereduplicata.Enabled = False
        CmdAlteraduplicata.Enabled = False
        CmdExcluiduplicata.Enabled = False
    End If
End Sub
                                                        
Sub subHabilitaBotoes()
    If TabFatura.Tab = 2 Then
        CmdInsereDesconto.Enabled = True
        CmdAlteraDesconto.Enabled = True
        CmdExcluiDesconto.Enabled = True
    ElseIf TabFatura.Tab = 1 Then
        CmdInsereduplicata.Enabled = True
        CmdAlteraduplicata.Enabled = True
        CmdExcluiduplicata.Enabled = True
    End If
End Sub
Sub subInsereRegistroVetorDupl(VetorCampos As Variant, Vetor As XArray, Grid As TDBGrid)  'Carrega o vetores
    Dim Cont As Integer
    Dim X As Integer
    
    'Insere um regitro no vetor
    
    'VetorCampos - vetor simples com os valores a inserir na respectiva ordem do grid
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    
    If Vetor.UpperBound(1) <> 0 Or Vetor(0, 0) <> Empty Then
        Vetor.Insert 1, Vetor.UpperBound(1) + 1
    End If
    Cont = Vetor.UpperBound(1)
    For X = 0 To Grid.Columns.Count - 2
        subSetaVal Vetor, Cont, X, FunNulo(VetorCampos(X))
    Next
    With Grid
        .ReBind
    End With
End Sub


Function funChecarCampos() As Boolean
    funChecarCampos = True
    
    If TabFatura.Tab = 1 Then
        'If funObrigatorioTXT(TxtValor, "O preenchimento do campo Valor é obrigatório.") Then Exit Function
        If OptEntrada.Value = True Then
            If FunObrigatorioCBOX(CboTipoPagto1, "O preenchimento do campo Tipo Recebimento é obrigatório.") Then Exit Function
        ElseIf FunObrigatorioCBOX(CboTipoPagto2, "O preenchimento do campo Tipo Pagamento é obrigatório.") Then Exit Function
        End If
        If FunObrigatorioCBO(CboContaCorrente, "O preenchimento do campo Conta Corrente é obrigatório.") Then Exit Function
        If FunObrigatorioCBOX(CboStatusduplicata, "O preenchimento do campo Status é obrigatório.") Then Exit Function
        If CboStatusduplicata.ListIndex = 2 Then
            If IsNull(DtpPagamento.Value) Then
                MsgBox "Data do Pagamento é Obrigatória", vbCritical + vbOKOnly, "ATENÇÃO'"
                DtpPagamento.SetFocus
                Exit Function
            End If
        End If
        If OptSaida.Value = True Then
            'If CboTipoPagto2.ListIndex = 2 Or CboTipoPagto2.ListIndex = 3 Or CboTipoPagto2.ListIndex = 4 Then
            If CboTipoPagto2.ListIndex = 2 Then
                If FunObrigatorioCBO(CboBancoForn2, "O preenchimento do Campo Banco do Fornecedor é obrigatório.") Then Exit Function
            End If
        End If
    End If
    funChecarCampos = False
End Function

Sub subPrepararInsercao()
        
   If TabFatura.Tab = 1 Then
        If Vetor2.UpperBound(1) > 0 Then
            TxtDuplicata.Text = Chr(66 + Vetor2.UpperBound(1))
        Else
            If Vetor2(0, 0) <> Empty Then
                TxtDuplicata.Text = "B"
            Else
                TxtDuplicata.Text = "A"
            End If
        End If
        TxtDuplicata.Enabled = False
        'TxtDuplicata.Text = ""
        If LblValor.Caption <> "" And LblValorTotal.Caption <> "" Then
            TxtValorDuplicata.Text = Format(((LblValor.Caption - funSomaColuna(Vetor3, 2)) - LblValorTotal.Caption), "standard")
        Else
            TxtValorDuplicata.Text = ""
        End If
        
        If Not IsNull(DatFornecedor.Recordset.Fields("focl_tx_TipoPag")) Then
            CboTipoPagto2.ListIndex = DatFornecedor.Recordset.Fields("focl_tx_TipoPag")
        Else
            CboTipoPagto2.ListIndex = 0
        End If
        
        TxtDocPagto2.Text = ""
        CboContaCorrente.Text = ""
        CboBancoForn2.Text = ""
        DtpVencimento.Value = Format(Now, "dd/MM/yy")
        DtpPagamento.Value = Null
        CboStatusduplicata.ListIndex = 0
   ElseIf TabFatura.Tab = 2 Then
        If Vetor2(0, 1) <> Empty Then
            DtpVencDesconto.Value = Vetor2(0, 1)
        End If
        TxtValorDesconto.Text = ""
        CboDescontos.Text = ""
   End If
End Sub

Private Sub CboBancoForn2_Change()
    If CboBancoForn2.Text <> "" Then
        DatBancoFornecedor.Recordset.Bookmark = CboBancoForn2.SelectedItem
    End If
End Sub

Private Sub CboContaCorrente_Change()
    If CboContaCorrente.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboContaCorrente.SelectedItem
        TxtCCorrente2.Text = CboContaCorrente.BoundText
    End If
End Sub

Private Sub CboDescontos_Change()
    If CboDescontos.Text <> "" And OptSaida.Value Then
        DatDescontosNF.Recordset.Bookmark = CboDescontos.SelectedItem
        If DatDescontosNF.Recordset.Fields("desc_tx_tipobase") = "2" Then
            If Vetor2(0, 1) <> Empty Then
                DtpVencDesconto.Value = Format$(Vetor2(0, 1), "DD/MM/YY")
            End If
        ElseIf Not IsNull(DatDescontosNF.Recordset.Fields("desc_tx_diarecolhimento")) Then
                Dim Ano As Integer
                Dim Mes As Integer

                Ano = Year(CDate(Vetor2(0, 1)))
                Mes = Month(CDate(Vetor2(0, 1)))
                If Mes = 12 Then
                   Mes = 1
                   Ano = Ano + 1
                Else
                   Mes = Mes + 1
                End If
                DtpVencDesconto.Value = Format$(DateSerial(Ano, Mes, DatDescontosNF.Recordset.Fields("desc_tx_diarecolhimento")), "dd/mm/yy")
            'Else
               ' DtpVencDesconto.Value = Null
        End If
    End If
End Sub


Private Sub CboFornecedor_DblClick(Area As Integer)
'    FrmFatura.MousePointer = vbHourglass
'    FrmFornecedores.Show 1
'    MdiPrincipal.MousePointer = vbDefault
End Sub

Private Sub CboFornecedor_LostFocus()
    Dim Sql As String
    Dim XLO_CHECAR As ADODB.Recordset
    
    StatusNF = "R"
    If CboFornecedor.Text <> "" And XInserir = "I" Then
        If OptSaida.Value = True Then
            If FunObrigatorioCBO(CboFornecedor, "Informe o Fornecedor!") Then Exit Sub
        Else
            If FunObrigatorioCBO(CboFornecedor, "Informe o Cliente!") Then Exit Sub
        End If
        
        subConectarControleDadosNV DatBancoFornecedor, "SELECT * FROM ConsFINBancoForn where focl_cd_forncli = " & CboFornecedor.BoundText, "Estatico"
        CboBancoForn2.Text = ""
        LblValor.Caption = ""
        
        If vazio = False Then
            TDBGrid1.MoveFirst
            subLimpaVetor Vetor1, TDBGrid1
            'cont = cont + 1
        End If
    
        'CHECAR SE EXISTEM NOTAS COM PENDENCIAS NO ESTOQUE
        Sql = "SELECT * FROM ConsFINFatura WHERE empr_cd_empresa = " & PCodEmpresa & " and nofi_nr_fatura = 0 and nofi_tx_PendEstoque = 'P' and focl_cd_forncli =" & CboFornecedor.BoundText
        SubQOpenRecordset XLO_CHECAR, Sql, Estatico
        If Not XLO_CHECAR.EOF Then
            MsgBox "Existem Notas com pendências no QualiMat.", vbCritical, "ATENÇÃO!"
        End If
        XLO_CHECAR.Close
    
        Sql = "SELECT * FROM ConsFINFatura WHERE empr_cd_empresa = " & PCodEmpresa & " and nofi_nr_fatura = 0 and focl_cd_forncli =" & CboFornecedor.BoundText
        subCarregaVetorGrid Sql, Array("nofi_cd_notafiscal", "nofi_nr_documento", "nofi_tx_serie", "nofi_dt_emissao", "nofi_vl_valor", "Selecionar", "nofi_tx_status", "nofi_nr_fatura", "tido_cd_tipodoc"), Vetor1, TDBGrid1
        If Vetor1(0, 0) = Empty Then
            vazio = True
            subDesabilitaBotoes
            CmdDesistir.Enabled = True
            MsgBox "Não exite Notas para Composição da Fatura", vbCritical, "ATENÇÃO!"
            TDBGrid1.Enabled = False
            TabFatura.TabEnabled(1) = False
        Else
            TDBGrid1.Enabled = True
            vazio = False
            TabFatura.TabEnabled(1) = True
            TabFatura.TabEnabled(2) = True
            CmdGravar.Enabled = True
    
            TDBGrid1.MoveFirst
            Do While Not TDBGrid1.EOF
                If StatusNF <> TDBGrid1.Columns("Status").Value Then
                    StatusNF = TDBGrid1.Columns("Status").Value
                End If
                TDBGrid1.Columns("Fatura") = Fatura
                TDBGrid1.Columns("Selecionar").Value = "N"
                TDBGrid1.MoveNext
            Loop
            TDBGrid1.MoveFirst
            TDBGrid1.SetFocus
        End If
        'TxtValor.Enabled = True
        subHabilitaBotoes
    End If

End Sub

Private Sub CboStatusduplicata_LostFocus()
    If CboStatusduplicata.ListIndex = 1 Then
        TxtDocPagto2.Enabled = True
        DtpPagamento.Enabled = True
    Else
        TxtDocPagto2.Enabled = False
        DtpPagamento.Enabled = False
    End If
End Sub

Private Sub CboTipoPagto2_LostFocus()
    If OptSaida.Value = True Then
        'If CboTipoPagto2.ListIndex = 2 Or CboTipoPagto2.ListIndex = 3 Or CboTipoPagto2.ListIndex = 4 Then
        If CboTipoPagto2.ListIndex = 2 Then
            CboBancoForn2.Enabled = True
        Else
            CboBancoForn2.Text = ""
            CboBancoForn2.Enabled = False
        End If
    End If
End Sub

Private Sub CmdAlteraDesconto_Click()
    Dim Xvalor As Double
           
    If TDBGrid3.Columns("NumDarf").Text <> "" Then
        MsgBox ("Este imposto já foi recolhido portanto não poderá ser alterado!"), vbExclamation, " Informação"
        Exit Sub
    End If
    
    FrmFatura.MousePointer = vbHourglass
    'If ExisteChave(1, Array("4"), Array(Fatura), Vetor1, TDBGrid1) Then Exit Sub
    ChaveD = -1
    XDesconto = "A"
    
    subPrepararInsercao
    CboDescontos.BoundText = TDBGrid3.Columns("codigo").Text
    xcod = CboDescontos.BoundText
    If IsDate(TDBGrid3.Columns("Vencimento")) Then
        DtpVencDesconto.Value = Format(TDBGrid3.Columns("Vencimento").Text, "dd/mm/yy")
    End If
    TxtValorDesconto.Text = Format(TDBGrid3.Columns("Valor").Text, "standard")
    
    pandescontos.Visible = True
    pandescontos.Enabled = True
    CboDescontos.SetFocus
    subDesabilitaBotoes
    FrmFatura.MousePointer = vbDefault

End Sub

Private Sub CmdAlteraduplicata_Click()
    Dim Status As String
    
    subDesabilitaBotoes
    TDBGrid1.Enabled = False
    
    PnlCompDuplicata.Visible = True
    PnlCompDuplicata.Enabled = True
    
    'Faturas só podem ser lançadas compromissadas
    CboStatusduplicata.Enabled = False
    
    XDuplicata = "A"
    Me.CboTipoPagto1.Text = TDBGrid2.Columns("Tipo Pag.").Text
    TxtDuplicata.Text = TDBGrid2.Columns("Código").Text
    DtpVencimento.Value = Format$(TDBGrid2.Columns("Vencimento").Text, "DD/MM/YY")
    TxtValorDuplicata.Text = Format$(TDBGrid2.Columns("Valor"), "standard")
    
    If TDBGrid2.Columns("CodStatus") = "C" Then 'Compromissado
        Status = "0"
    ElseIf TDBGrid2.Columns("CodStatus") = "R" Then 'Realizado
        Status = "1"
    End If
    CboStatusduplicata.ListIndex = Status
    
    If OptSaida.Value = True Then
        CboTipoPagto2.ListIndex = (TDBGrid2.Columns("CodTipoPagto") - 1)
        TxtDocPagto2.Text = FunNulo(TDBGrid2.Columns("Doc. Pag."))
        If TDBGrid2.Columns("Data Pag.") <> Null Or TDBGrid2.Columns("Data Pag.") <> "" Then
            DtpPagamento.Enabled = True
            DtpPagamento.Value = Format$(TDBGrid2.Columns("Data Pag."), "DD/MM/YY")
        End If
    Else
        CboTipoPagto1.ListIndex = (TDBGrid2.Columns("CodTipoPagto") - 1)
        TxtDocPagto2.Text = FunNulo(TDBGrid2.Columns("Doc. Rec."))
        If TDBGrid2.Columns("Data Rec.") <> Null Or TDBGrid2.Columns("Data Rec.") <> "" Then
            DtpPagamento.Enabled = True
            DtpPagamento.Value = Format$(TDBGrid2.Columns("Data Rec."), "DD/MM/YY")
        End If
    End If

    XVlDuplicata = CDbl(TxtValorDuplicata.Text)

    CboContaCorrente.BoundText = TDBGrid2.Columns("CodConta")
    CboBancoForn2.BoundText = TDBGrid2.Columns("CodBancoForn")
    
End Sub

Private Sub CmdConfirmaDesconto_Click()
       
   If FunObrigatorioCBO(CboDescontos, "É necessário selecionar um desconto!") Then
        CboDescontos.SetFocus
        Exit Sub
    End If
    
    If FunObrigatorioTXT(TxtValorDesconto, "É necessário preencher o campo destinado ao valor do desconto!") Then
        TxtValorDesconto.SetFocus
        Exit Sub
    End If
    
    If (xcod = CboDescontos.BoundText) Then
        If XDesconto = "I" Then
           subInsereRegistroVetor Array(CboDescontos.Text, Format(DtpVencDesconto.Value, "dd/mm/yyyy"), TxtValorDesconto.Text, CboDescontos.BoundText, "", "", ""), Vetor3, TDBGrid3
        ElseIf XDesconto = "A" Then
            TDBGrid3.Columns("Descrição").Text = CboDescontos.Text
            TDBGrid3.Columns("Vencimento").Text = Format(DtpVencDesconto.Value, "dd/mm/yyyy")
            TDBGrid3.Columns("Valor").Text = TxtValorDesconto.Text
            TDBGrid3.Columns("codigo").Text = CboDescontos.BoundText
            TDBGrid3.Update
        End If
    Else
        DatDescontosNF.Recordset.Bookmark = CboDescontos.SelectedItem
        If funExisteChave(1, Array(3), Array(CboDescontos.BoundText), Vetor3, TDBGrid3) Then
            MsgBox "Esse imposto já foi cadastrado", vbCritical, "ATENÇÃO!"
            If XDesconto = "I" Then
               CboDescontos.SetFocus
               CmdInsereDesconto_Click
            Else
                CboDescontos.SetFocus
                Exit Sub
            End If
        Else
            If XDesconto = "I" Then
                subInsereRegistroVetor Array(CboDescontos.Text, Format(DtpVencDesconto.Value, "dd/mm/yyyy"), TxtValorDesconto.Text, CboDescontos.BoundText, "", "", ""), Vetor3, TDBGrid3
            Else
                TDBGrid3.Columns("Descrição").Text = CboDescontos.Text
                TDBGrid3.Columns("Vencimento").Text = Format(DtpVencDesconto.Value, "dd/mm/yyyy")
                TDBGrid3.Columns("Valor").Text = TxtValorDesconto.Text
                TDBGrid3.Columns("codigo").Text = CboDescontos.BoundText
                TDBGrid3.Update
            End If
        End If
    End If

    pandescontos.Visible = False
    pandescontos.Enabled = False
    
    If Vetor3(0, 0) <> Empty Then
       subHabilitaBotoes
    Else
        CmdInsereDesconto.Enabled = True
    End If
    
    TDBGrid3.Enabled = True

End Sub


Private Sub CmdConfirmaduplicata_Click()
    Dim XStatus As String
    Dim xcod As String
    Dim Xtipopagrec As Integer
    Dim XPagRec As String
    Dim XLF_AUXILIAR1 As Double
    Dim XLF_AUXILIAR2 As Double

    On Error GoTo TrataErro
    
    If Me.CboStatusduplicata.Text = "Realizado" Then
        If Me.TxtDocPagto2.Text = "" Or IsNull(Me.DtpPagamento.Value) Then
            MsgBox "Campos Obrigatórios para lançamentos realizados: Doc. Pag/Rec e Data Pag/Rec!", vbCritical, "Financeiro"
            Exit Sub
        End If
    End If
    'Função para verificar preenchimento dos campos obrigatórios
    If funChecarCampos Then Exit Sub
        
    If XDuplicata = "A" Then
       ValorTotalDuplicata = ValorTotalDuplicata - XVlDuplicata
    End If

'******
    'AS VARIÁVEIS FORAM INTRODUZIDA PARA SOLUCIONAR UMA EXCEÇÃO REF. A APROX.,
    'ONDE OS 2 IFs ABAIXO NÃO RESOLVERAM
    XLF_AUXILIAR1 = Format(CDbl(LblValor.Caption), "#.0000000000")
    XLF_AUXILIAR2 = Format(ValorTotalDuplicata + CDbl(TxtValorDuplicata.Text), "#.0000000000")
    'A formatação abaixo foi implementada para tentar solucionar exceções na aproximação.
    'SÓ NÃO ATENDEU A EXCEÇÃO ACIMA->If (CDbl(LblValorLiquido1.Caption) + CDbl(TxtDevolucao.Text)) < (ValorTotalDuplicata + CDbl(TxtValorDuplicata.Text)) Then
    'TESTE->If (CDbl(Format(LblValorLiquido1.Caption, "standard")) + CDbl(Format(TxtDevolucao.Text, "standard"))) < (CDbl(Format(ValorTotalDuplicata, "standard")) + CDbl(Format(TxtValorDuplicata.Text, "standard"))) Then
    If XLF_AUXILIAR1 < XLF_AUXILIAR2 Then
'*****
    'If CDbl(LblValor.Caption) < (ValorTotalDuplicata + CDbl(TxtValorDuplicata.Text)) Then
        MsgBox "A soma das duplicatas excedeu o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
        TxtValorDuplicata.SetFocus
        CmdInsereduplicata_Click
    Else
        If OptEntrada.Value = True Then
            Xdebcred = "C"
            Xtipopagrec = CboTipoPagto1.ItemData(CboTipoPagto1.ListIndex)
            XPagRec = CboTipoPagto1.Text
        ElseIf OptSaida.Value = True Then
            Xtipopagrec = CboTipoPagto2.ItemData(CboTipoPagto2.ListIndex)
            Xdebcred = "D"
            XPagRec = CboTipoPagto2.Text
        End If
        
        If TxtDuplicata.Text = "" Then
            xcod = "U"
        Else
            xcod = TxtDuplicata.Text
        End If
    
        If CboStatusduplicata.ListIndex = 0 Then 'Compromissado
           XStatus = "C"
        ElseIf CboStatusduplicata.ListIndex = 1 Then 'Realizado
           XStatus = "R"
        End If

        If XDuplicata = "I" Then
            subInsereRegistroVetorDupl Array(xcod, Format(DtpVencimento.Value, "dd/mm/yyyy"), TxtValorDuplicata.Text, XStatus, CboStatusduplicata.Text, _
            Xtipopagrec, XPagRec, TxtDocPagto2.Text, Format(DtpPagamento.Value, "dd/mm/yyyy"), CboContaCorrente.BoundText, CboContaCorrente.Text, FunNuloValor(CboBancoForn2), CboBancoForn2.Text, ""), Vetor2, TDBGrid2
        ElseIf XDuplicata = "A" Then
            TDBGrid2.Columns("Código").Text = TxtDuplicata.Text
            TDBGrid2.Columns("Vencimento").Text = Format(DtpVencimento.Value, "dd/mm/yyyy")
            TDBGrid2.Columns("Valor").Text = TxtValorDuplicata.Text
            TDBGrid2.Columns("CodStatus").Text = XStatus
            TDBGrid2.Columns("Status").Text = CboStatusduplicata.Text
            TDBGrid2.Columns("CodTipoPagto").Text = Xtipopagrec
            If OptSaida.Value = True Then
                TDBGrid2.Columns("Tipo Pag.").Text = XPagRec
                TDBGrid2.Columns("Doc. Pag.").Text = TxtDocPagto2.Text
                TDBGrid2.Columns("Data Pag.").Text = Format(DtpPagamento.Value, "dd/mm/yyyy")
                TDBGrid2.Columns("CodBancoForn").Text = FunNuloValor(CboBancoForn2)
                TDBGrid2.Columns("Conta Corrente (Fornecedor)").Text = CboBancoForn2.Text
            Else
                TDBGrid2.Columns("Tipo Rec.").Text = XPagRec
                TDBGrid2.Columns("Doc. Rec.").Text = TxtDocPagto2.Text
                TDBGrid2.Columns("Data Rec.").Text = Format(DtpPagamento.Value, "dd/mm/yyyy")
            End If
            
            TDBGrid2.Columns("CodConta").Text = CboContaCorrente.BoundText
            TDBGrid2.Columns("Conta Corrente (Empresa)").Text = CboContaCorrente.Text
            TDBGrid2.Update
        End If

        PnlCompDuplicata.Visible = False
        PnlCompDuplicata.Enabled = False
            
        If Vetor2(0, 0) <> Empty Then
            subHabilitaBotoes
            LblValorTotal.Caption = funSomaColuna(Vetor2, 2)
            ValorTotalDuplicata = CDbl(LblValorTotal.Caption)
        Else
            CmdInsereduplicata.Enabled = True
        End If
        TDBGrid2.Enabled = True
    End If
    
    
ResumoDoErro:
    Exit Sub

TrataErro:
    funTrataErros (ComMensagem)
    'Conexao.Rollback
    Resume ResumoDoErro
End Sub


Private Sub CmdDesistir_Click()
    Unload Me
End Sub



Private Sub CmdExcluiDesconto_Click()
   Dim Houve_Erro As Integer

   Houve_Erro = False
           
    If TDBGrid3.Columns("NumDarf").Text <> "" Then
        MsgBox ("Este imposto já foi recolhido portanto não poderá ser excluído!"), vbExclamation, " Informação"
        Exit Sub
    End If
    
   subDesabilitaBotoes
   TDBGrid1.Enabled = False
   If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
      subRemoveRegistroVetor Vetor3, TDBGrid3
      If Vetor3(0, 0) <> Empty Then
        subHabilitaBotoes
      Else
        CmdInsereDesconto.Enabled = True
      End If
      TDBGrid1.Enabled = True
   Else
        subHabilitaBotoes
   End If
   Exit Sub
    
RotuloErro:
    Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next
End Sub

Private Sub CmdExcluiduplicata_Click()
   Dim Houve_Erro As Integer
   Dim ValorExcluido As Double

   Houve_Erro = False
    
   subDesabilitaBotoes
   
   TDBGrid2.Enabled = False
   
   If TDBGrid2.Columns("Status") = "Realizado" Then
        MsgBox "A duplicata não pode ser apagada porque já foi realizada.", vbCritical + vbOKOnly, "ATENÇÃO"
        subHabilitaBotoes
        Exit Sub
   End If
   
   If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical, "CUIDADO") = vbYes Then 'Confirma exclusão
        ValorExcluido = TDBGrid2.Columns("Valor").Value
        subRemoveRegistroVetor Vetor2, TDBGrid2
   End If
   
   If Vetor2(0, 0) <> Empty Then
       subHabilitaBotoes
   Else
       CmdInsereduplicata.Enabled = True
   End If
   LblValorTotal.Caption = Format(ValorTotalDuplicata - ValorExcluido, "standard")
   ValorTotalDuplicata = LblValorTotal.Caption
   TDBGrid2.Enabled = True

   Exit Sub

RotuloErro:
    Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next
End Sub

Private Sub cmdGravar_Click()
    Dim Sql As String
    Dim VetorFat As Double
    Dim Cont As Integer

    Reload = True

    On Error GoTo TrataErro
          
    CmdGravar.Enabled = False
    Me.MousePointer = vbHourglass
   
    
    'If Not EntrouDupl Then
    If Vetor2(0, 0) = Empty Then
       MsgBox "É necessário compor Duplicata!", vbCritical, "ATENÇÃO"
       Me.MousePointer = vbDefault
       CmdGravar.Enabled = True
       Exit Sub
    End If
    
    If FunVerificaTipoDocumento = False Then
        MsgBox "Não é possível gerar faturas de notas com tipos de documento ou séries diferentes. " & Chr(10) & " Favor alterar os tipos de documento ou a série das notas para realizar a operação.", vbInformation, "FATURAS"
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    If Vetor1(0, 0) <> Empty Then
        TDBGrid1.MoveFirst
        Do While Not TDBGrid1.EOF
            If TDBGrid1.Columns("Selecionar").Value = "S" Then
                VetorFat = VetorFat + CDbl(TDBGrid1.Columns("Valor").Value)
            End If
            TDBGrid1.MoveNext
        Loop
     
        If VetorFat > 0 And (((VetorFat * 100) \ 1) / 100) <> (((CDbl(LblValor.Caption) * 100) \ 1) / 100) Then
            MsgBox "O valor da fatura é diferente da soma das notas selecionadas!", vbCritical, "ATENÇÃO"
            Me.MousePointer = vbDefault
            CmdGravar.Enabled = True
            Exit Sub
        ElseIf VetorFat = 0 Then
            MsgBox "Nenhuma nota foi selecionada!", vbCritical, "ATENÇÃO"
            Me.MousePointer = vbDefault
            CmdGravar.Enabled = True
            Exit Sub
        End If
    End If
    
    'Verifica se o total da nota é igual a soma das duplicatas com os descontos ou se alterou o valor da nota
    If funVerificaValores Then
        Me.MousePointer = vbDefault
        CmdGravar.Enabled = True
        Exit Sub
    End If
    
    If XInserir = "I" Then
        Fatura = CStr(FunPegaGuardaUltimo("F"))
    ElseIf XInserir = "A" Then
        Fatura = NFatura
    End If
    
    If OptEntrada.Value = True Then
        Xdebcred = "C"
    ElseIf OptSaida.Value = True Then
        Xdebcred = "D"
    End If
            
    'Notas Fiscais
    
    'Definir Xtipofatura
    If Vetor1(0, 0) <> Empty Then
       TDBGrid1.MoveFirst
       Do While Not TDBGrid1.EOF
            If TDBGrid1.Columns("Selecionar").Value = "S" Then
               Cont = Cont + 1
            End If
            TDBGrid1.MoveNext
        Loop
    End If
    If Cont > 1 Then
        XTipoFatura = 3
    ElseIf Cont = 1 Then
        If Vetor2.UpperBound(1) > 0 Then
            XTipoFatura = 2
        Else
            XTipoFatura = 1
        End If
    End If
    
    Conexao.BeginTrans
    If Vetor1(0, 0) <> Empty And XFormulario <> "FrmNotasFiscais" And XFormulario <> "TelaDuplicatas" Then
        TDBGrid1.MoveFirst
        Do While Not TDBGrid1.EOF
            If TDBGrid1.Columns("Selecionar").Value = "S" Then
                Sql = "UPDATE NotasFiscais " & _
                      "SET nofi_nr_fatura = " & Fatura & "," & _
                      "nofi_tx_tipoFatura = '" & XTipoFatura & "'" & _
                      "WHERE nofi_cd_notafiscal = " & CLng(TDBGrid1.Columns("CodigoNF"))
                Conexao.Execute Sql
            End If
            TDBGrid1.MoveNext
        Loop
    End If
              
    'Descontos
    If EntrouDesc Then
        If XInserir = "A" Then
            Conexao.Execute "DELETE from DescontosNF WHERE deen_nr_fatura = " & Fatura
        End If
        If Vetor3(0, 0) <> Empty Then
            subGravaVetor "DescontosNF", "deen_dt_prevrecolhimento,deen_vl_valor,desc_cd_desconto, deen_dt_recolhimento, deen_nr_DarfDam, deen_nr_fatura", Array("D", "N", "N", "D", "T", "N"), 5, Array(1, 2, 3, 5, 6), "" & Fatura & "", Vetor3, TDBGrid1
        End If
    End If
    
    
    'Duplicatas
    
    'Definir o código da duplicata como "U" para caso de unica e
    '"A" para caso de ser a primeira de uma série
    If Vetor2.UpperBound(1) > 0 Then
        If Vetor2(0, 0) <> Empty Then
            Vetor2(0, 0) = "A"
        End If
    Else
        If Vetor2(0, 0) <> Empty Then
            Vetor2(0, 0) = "U"
        End If
    End If

    If EntrouDupl Then
        If XInserir = "A" Then
           Conexao.Execute "DELETE from Duplicatas WHERE dupl_nr_fatura = " & Fatura
        End If
        If Vetor2(0, 0) <> Empty Then
            subGravaVetor "Duplicatas", "dupl_nr_duplicata,dupl_dt_vencimento,dupl_vl_valor,dupl_tx_status,dupl_tx_tipopagrec,dupl_tx_docpagrec,dupl_dt_pagrec,coco_cd_codigo,bafo_cd_codigo,dupl_nr_fatura,dupl_tx_debcred", Array("T", "D", "N", "T", "T", "T", "D", "N", "N", "N", "T"), 9, Array(0, 1, 2, 3, 5, 7, 8, 9, 11), "" & Fatura & ",'" & Xdebcred & "'", Vetor2, TDBGrid2
        End If
    End If
    Conexao.CommitTrans
        
    '////Imprime processo de pagamento/////
    If XInserir = "I" Then
        If XGT_IMPRIMEPP = "S" Then
            If MsgBox("Deseja Imprimir Processo de Pagamento?", vbQuestion + vbYesNo, "ATENÇÃO") = vbYes Then
'                subImprimePP Fatura, True
            End If
        End If
    End If
    '/////////////////////////////////////
        
    Me.MousePointer = vbDefault
    CmdGravar.Enabled = True
            
    Unload Me
    Exit Sub
TrataErro:
    funTrataErros (ComMensagem)
    
    
    Conexao.RollbackTrans
End Sub


Private Sub CmdInsereDesconto_Click()
    FrmFatura.MousePointer = vbHourglass
    If XInserir = "I" Then
        If funExisteChave(1, Array("4"), Array(Fatura), Vetor3, TDBGrid3) Then Exit Sub
    End If
    ChaveD = -1
    XDesconto = "I"
    subPrepararInsercao
    pandescontos.Visible = True
    pandescontos.Enabled = True
    CboDescontos.SetFocus
    subDesabilitaBotoes
    FrmFatura.MousePointer = vbDefault
End Sub



Private Sub CmdInsereduplicata_Click()
    FrmFatura.MousePointer = vbHourglass
    XDuplicata = "I"
    ChaveDp = -1
    If LblValor.Caption = LblValorTotal.Caption Then
       MsgBox "As duplicatas já atingiram o valor total da Nota Fiscal", vbCritical, "ATENÇÃO"
    Else
        PnlCompDuplicata.Visible = True
        PnlCompDuplicata.Enabled = True
        TxtDocPagto2.Enabled = False
        DtpPagamento.Enabled = False
        
        subPrepararInsercao
        
        'subDesabilitaBotoes
        
        TxtValorDuplicata.SetFocus
    End If
    
    'Faturas só podem ser lançadas compromissadas
    If StatusNF = "R" Then
        CboStatusduplicata.ListIndex = 1
    End If
        CboStatusduplicata.Enabled = False
        
    FrmFatura.MousePointer = vbDefault
End Sub

Private Sub CmdSairDesconto_Click()
    pandescontos.Visible = False
    pandescontos.Enabled = False
    
    If Vetor3(0, 0) <> Empty Then
        subHabilitaBotoes
    Else
        subDesabilitaBotoes
        CmdInsereDesconto.Enabled = True
    End If
    XDesconto = ""
    TDBGrid1.Enabled = True
    
End Sub

Private Sub CmdSairDuplicata_Click()
    
    PnlCompDuplicata.Visible = False
    PnlCompDuplicata.Enabled = False
    
    XDuplicata = ""
    TDBGrid2.Enabled = True
 
    If Vetor2(0, 0) <> Empty Then
       subHabilitaBotoes
    Else
       CmdInsereduplicata.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    'Set Formulario = FrmFatura
    Reload = False
    
End Sub

Private Sub Form_Load()
    'Set Formulario = FrmFatura


    Dim Sql As String
    
    Reload = False
    'On Error GoTo TrataErro
    
    subManutencaoJanelasAtivas "I", "Frmfatura"
    
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
 
    EntrouDesc = False
    EntrouDupl = False

    If XInserir = "A" Then
        FrmFatura.Caption = "Alteração da Fatura Nº " + CStr(NFatura)
        If EntSaida = "E" Then
           subConectarControleDadosNV DatFornecedor, "SELECT * FROM FornClientes WHERE focl_tx_classe = 'C' ORDER BY focl_tx_fantasia ", Estatico
           OptEntrada.Value = True
           CboTipoPagto1.RemoveItem (1)
           CboTipoPagto1.RemoveItem (1)
           CboTipoPagto1.RemoveItem (1)
           CboTipoPagto1.RemoveItem (1)
        Else
           subConectarControleDadosNV DatFornecedor, "SELECT * FROM FornClientes WHERE focl_tx_classe = 'F' ORDER BY focl_tx_fantasia ", Estatico
           OptSaida.Value = True
        End If
        If XFormulario = "TelaFatura" Then
            CboFornecedor.Text = TelaFatura.Adodc1.Recordset.Fields("focl_tx_fantasia")
            If Not IsNull(TelaFatura.Adodc1.Recordset.Fields("nofi_tx_tipofatura")) Then
                XTipoFatura = TelaFatura.Adodc1.Recordset.Fields("nofi_tx_tipofatura")
            End If
            If TelaFatura.Adodc1.Recordset.Fields("tipo") = "A Pagar" Then
                OptSaida.Value = True
            Else
               OptEntrada.Value = True
            End If
        'ElseIf XFormulario = "FrmNotasFiscais" Then
        ElseIf XFormulario = "TelaNotasFiscais" Then
            CboFornecedor.Text = TelaNotasFiscais.Adodc1.Recordset.Fields("focl_tx_fantasia")
            XTipoFatura = "3"
            If EntSaida = "S" Then
                OptSaida.Value = True
            Else
               OptEntrada.Value = True
            End If
            FrmFatura.TabFatura.Tab = 1
        ElseIf XFormulario = "TelaDuplicatas" Then
            CboFornecedor.Text = TelaDuplicatas.Adodc1.Recordset.Fields("focl_tx_fantasia")
            XTipoFatura = "3"
            If TelaDuplicatas.Adodc1.Recordset.Fields("Saida") <> 0 Then
                OptSaida.Value = True
            ElseIf TelaDuplicatas.Adodc1.Recordset.Fields("A Receber") <> 0 Then
                OptEntrada.Value = True
            End If
            FrmFatura.TabFatura.Tab = 1
        End If
        'NotasFiscais
        Sql = "SELECT NotasFiscais.nofi_cd_notafiscal, NotasFiscais.nofi_nr_documento, NotasFiscais.nofi_tx_serie, NotasFiscais.nofi_dt_emissao, NotasFiscais.nofi_vl_Valor, FornClientes.focl_tx_fantasia, NotasFiscais.nofi_nr_fatura, NotasFiscais.nofi_tx_status, 'S' AS Selecionar,  NotasFiscais.tido_cd_tipodoc " & _
        "FROM NotasFiscais LEFT JOIN Fornclientes ON NotasFiscais.focl_cd_forncli = Fornclientes.focl_cd_forncli Where NotasFiscais.nofi_nr_fatura =" & NFatura
        subCarregaVetorGrid Sql, Array("nofi_cd_notafiscal", "nofi_nr_documento", "nofi_tx_serie", "nofi_dt_emissao", "nofi_vl_valor", "Selecionar", "nofi_tx_status", "nofi_nr_fatura", "tido_cd_tipodoc"), Vetor1, TDBGrid1
                                    
        'TxtValor.Text = Format$(funSomaColuna(Vetor1, 4), "standard")
        LblValor.Caption = Format$(funSomaColuna(Vetor1, 4), "standard")
        PanConjunto.Enabled = False
        TDBGrid1.Enabled = False
        TabFatura.TabEnabled(1) = True
        TabFatura.TabEnabled(2) = True
        CmdGravar.Enabled = True
        
        If EntSaida = "S" Then
            subConectarControleDadosNV DatBancoFornecedor, "SELECT * FROM ConsFINBancoForn where focl_cd_forncli = " & CboFornecedor.BoundText, "Estatico"
        End If
        StatusNF = TDBGrid1.Columns("Status").Value
    Else
        TabFatura.TabEnabled(1) = False
        TabFatura.TabEnabled(2) = False
        vazio = True
        OptSaida.Value = True
    End If
    
    'Descontos
    Sql = "SELECT DescontosNF.*,DescontosImpostos.desc_tx_descricao FROM DescontosNF LEFT JOIN DescontosImpostos ON DescontosNF.desc_cd_desconto = DescontosImpostos.desc_cd_desconto WHERE DescontosNF.deen_nr_fatura = " & NFatura
    subCarregaVetor Sql, Array("desc_tx_descricao", "deen_dt_prevrecolhimento", "deen_vl_valor", "desc_cd_desconto", "deen_nr_fatura", "deen_dt_recolhimento", "deen_nr_DarfDam"), Vetor3, TDBGrid3

    'Duplicatas
    Sql = "SELECT * FROM ConsFINCompDuplicata WHERE dupl_nr_fatura = " & NFatura
    'subCarregaVetor sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "dupl_vl_valor", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BaEmpresa", "bafo_cd_codigo", "BaFornecedor", "dupl_nr_fatura", "dupl_cd_duplicata"), Vetor2, TDBGrid2
    subCarregaVetor Sql, Array("dupl_nr_duplicata", "dupl_dt_vencimento", "valor", "dupl_tx_status", "status", "dupl_tx_tipopagrec", "tipoPagto", "dupl_tx_docpagrec", "dupl_dt_pagrec", "coco_cd_codigo", "BaEmpresa", "bafo_cd_codigo", "BaFornecedor", "dupl_nr_fatura", "dupl_cd_duplicata"), Vetor2, TDBGrid2
    
    TDBGrid2.ReBind
    
Exit Sub

TrataErro:
    If Err.Number = NumberTravaRegistro And Err.Description = DescriptionTravaRegistro Then
        If MsgBox("O exemplo está sendo alterado por outro usuário. Deseja tentar novamente?", vbCritical + vbQuestion + vbYesNo, "ATENÇÃO") = vbYes Then
            Resume
        Else
            Unload Me
            Exit Sub
        End If
    ElseIf Err.Number = NumberRegistroExcluido And Err.Description = DescriptionRegistroExcluido Then
        MsgBox "O exemplo foi excluído por outro usuário!", vbCritical, "ATENÇÃO"
        Unload Me
        Exit Sub
    Else
        funTrataErros (ComMensagem)
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "Frmfatura"
End Sub

Private Sub dtpPagamento_LostFocus()
    If Not funConsistenciaDatas Then Exit Sub
End Sub

Private Sub dtpVencDesconto_LostFocus()
    If Not funConsistenciaDatas Then Exit Sub
End Sub

Private Sub dtpVencimento_LostFocus()
    If Not funConsistenciaDatas Then Exit Sub
End Sub

Private Sub OptEntrada_Click()
    subConectarControleDadosNV DatFornecedor, "SELECT * FROM FornClientes WHERE focl_tx_classe = 'C' ORDER BY focl_tx_fantasia ", Estatico
End Sub


Private Sub OptEntrada_LostFocus()
    subHabilitaCampos
End Sub


Private Sub OptSaida_Click()
    subConectarControleDadosNV DatFornecedor, "SELECT * FROM FornClientes WHERE focl_tx_classe = 'F' ORDER BY focl_tx_fantasia ", Estatico
End Sub

Private Sub OptSaida_LostFocus()
    subHabilitaCampos
End Sub


Private Sub TabFatura_Click(PreviousTab As Integer)
    Dim Sql As String
    Dim ResForn As Object
    Dim XVlDesconto As Double
    Dim XData As Date
    Dim Ano As Integer
    Dim Mes As Integer
    Dim X As Byte
    
    If TabFatura.Tab <> 0 Then
        If LblValor.Caption = "" Then
            MsgBox "Não existem notas selecionadas!", vbCritical, "ATENÇÃO"
            TabFatura.Tab = 0
            Exit Sub
        End If
    End If

    If TabFatura.Tab = 1 And Not EntrouDupl Then
        If OptEntrada.Value = True Then 'Prepara o painel para Contas a Receber
            LblTipoPagamento.Caption = "Tipo Recebimento:"
            LblDocPagRec.Caption = "Doc. Recebimento:"
            LblBcFornecedor.Visible = False
            CboBancoForn2.Visible = False
            LblPagamento.Caption = "Recebimento:"
            PanPagto2.Top = 375
            CboTipoPagto2.Visible = False
            CboTipoPagto1.Visible = True
            TDBGrid2.Columns("Tipo Pag.").Caption = "Tipo Rec."
            TDBGrid2.Columns("Doc. Pag.").Caption = "Doc. Rec."
            TDBGrid2.Columns("Data Pag.").Caption = "Data Rec."
            TDBGrid2.Columns("Conta Corrente (Fornecedor)").Visible = False
        Else
            CboTipoPagto2.Visible = True
            CboTipoPagto1.Visible = False
        End If

        EntrouDupl = True
        If Vetor2(0, 0) = Empty Then
            subDesabilitaBotoes
            CmdInsereduplicata.Enabled = True
        Else
             If XFormulario = "TelaDuplicatas" Then
                X = 0
                While X <= Vetor2.UpperBound(1) 'Utilizado para a alteração da fatura, posicionando o cursor no registro desejado
                    If Vetor2(X, 14) = CStr(ChaveDp) Then
                        TDBGrid2.Bookmark = X
                        PosCursor = X
                        X = Vetor2.UpperBound(1)
                    End If
                    X = X + 1
                Wend
                CmdAlteraduplicata_Click
                TDBGrid2.Bookmark = PosCursor
            End If
        End If
        LblValorTotal = funSomaColuna(Vetor2, 2)
        ValorTotalDuplicata = CDbl(LblValorTotal.Caption)

    ElseIf TabFatura.Tab = 2 Then
        If Not EntrouDesc Then
            If OptSaida.Value = True Then  'É um fornecedor
                If XGT_ORDEMDESCONTO = "C" Then
                    'Ordernar por código reduzido
                    Sql = "SELECT (CASE WHEN plco_cd_CodReduzido IS NULL THEN desc_tx_descricao " & _
                            " ELSE plco_cd_CodReduzido + ' - ' + desc_tx_descricao END) " & _
                            " AS DescricaoCombo, * FROM ConsFINPlCDescImp " & _
                            " WHERE desc_tx_classe = 'D' and (empr_cd_empresa= " & PCodEmpresa & " or (empr_cd_empresa is null)) AND NOT plco_cd_CodReduzido IS NULL ORDER BY plco_cd_CodReduzido"
                Else
                    'Ordenar por Descrição
                    Sql = "SELECT (CASE WHEN plco_cd_CodReduzido IS NULL THEN desc_tx_descricao " & _
                            " ELSE desc_tx_descricao + ' (' + plco_cd_CodReduzido + ')' END) " & _
                            " AS DescricaoCombo, * FROM ConsFINPlCDescImp " & _
                            " WHERE desc_tx_classe = 'D' and (empr_cd_empresa= " & PCodEmpresa & " or (empr_cd_empresa is null)) ORDER BY desc_tx_descricao"
                End If
                subConectarControleDadosNV DatDescontosNF, Sql, "Estatico"
            Else  'É um Cliente
                If XGT_ORDEMDESCONTO = "C" Then
                    'Ordernar por código reduzido
                    Sql = "SELECT (CASE WHEN plco_cd_CodReduzido IS NULL THEN desc_tx_descricao " & _
                            " ELSE plco_cd_CodReduzido + ' - ' + desc_tx_descricao END) " & _
                            " AS DescricaoCombo, * FROM ConsFINPlCDescImp " & _
                            " WHERE desc_tx_classe = 'I' and and desc_tx_fonte = 'S' and (empr_cd_empresa= " & PCodEmpresa & " or (empr_cd_empresa is null)) AND NOT plco_cd_CodReduzido IS NULL ORDER BY plco_cd_CodReduzido"
                Else
                    'Ordenar por Descrição
                    Sql = "SELECT (CASE WHEN plco_cd_CodReduzido IS NULL THEN desc_tx_descricao " & _
                            " ELSE desc_tx_descricao + ' (' + plco_cd_CodReduzido + ')' END) " & _
                            " AS DescricaoCombo, * FROM ConsFINPlCDescImp " & _
                            " WHERE desc_tx_classe = 'I' and and desc_tx_fonte = 'S' and (empr_cd_empresa= " & PCodEmpresa & " or (empr_cd_empresa is null)) ORDER BY desc_tx_descricao"
                End If
                subConectarControleDadosNV DatDescontosNF, Sql, "Estatico"
            End If
            
           If Vetor3(0, 0) = Empty Then
                subDesabilitaBotoes
                CmdInsereDesconto.Enabled = True
            End If
        
            'Verifica se o fornecedor/cliente possui um imposto cadastrado
            If OptSaida.Value = True And XInserir = "I" Then 'É um fornecedor
                SubQOpenRecordset ResForn, "SELECT DesImpForn.*,DescontosImpostos.desc_tx_descricao,DescontosImpostos.desc_tx_diarecolhimento,DescontosImpostos.desc_tx_tipobase FROM DesImpForn LEFT JOIN DescontosImpostos ON DesImpForn.desc_cd_desconto = DescontosImpostos.desc_cd_desconto WHERE DesImpForn.focl_cd_forncli=" & CboFornecedor.BoundText, "Estatico"
                
                If Not (ResForn.BOF And ResForn.EOF) Then
                    ResForn.MoveFirst
                    Do While Not ResForn.EOF
                        If ResForn!desc_tx_tipobase = "1" Then
                            XVlDesconto = Format$(((CDbl(LblValor.Caption) * (ResForn!desi_Vl_percentual)) / 100), "##,##0.00")
                            Ano = Year(DtpVencimento.Value)
                            Mes = Month(DtpVencimento.Value)
                            If Mes = 12 Then
                                Mes = 1
                                Ano = Ano + 1
                            Else
                                Mes = Mes + 1
                            End If
                            XData = DateSerial(Ano, Mes, ResForn!desc_tx_DiaRecolhimento)
                            subInsereRegistroVetor Array(ResForn!desc_tx_descricao, XData, XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor3, TDBGrid3
                        ElseIf ResForn!desc_tx_tipobase = "2" Then
                            XVlDesconto = Format$(((CDbl(LblValor.Caption) * (ResForn!desi_Vl_percentual)) / 100), "##,##0.00")
                            subInsereRegistroVetor Array(ResForn!desc_tx_descricao, Format(DtpVencimento.Value, "dd/mm/yyyy"), XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor3, TDBGrid3
                        End If
                        
                        ResForn.MoveNext
                        subHabilitaBotoes
                    Loop
                End If
                
                ResForn.Close
                LblValor.Caption = Format$((CDbl(LblValor.Caption) - funSomaColuna(Vetor1, 2)), "##,##0.00")
            
            ElseIf OptEntrada.Value = True And XInserir = "I" Then 'É um Cliente
                TDBGrid3.Columns("Vencimento").Visible = False
                
                SubQOpenRecordset ResForn, "SELECT DescontosImpostos.desc_cd_desconto, DescontosImpostos.desc_tx_descricao,DescontosImpostos.desc_tx_diarecolhimento,DescontosImpostos.desc_tx_tipobase, DescontosImpostos.desc_tx_fonte, DescontosImpostos.desc_vl_taxa FROM DescontosImpostos where descontosimpostos.desc_tx_fonte = 'S'", "Estatico"
                
                If Not (ResForn.BOF And ResForn.EOF) Then
                    ResForn.MoveFirst
                    Do While Not ResForn.EOF
                       subInsereRegistroVetor Array(ResForn!desc_tx_descricao, "  /  /  ", XVlDesconto, ResForn!desc_cd_desconto, "", "", ""), Vetor3, TDBGrid3
                       ResForn.MoveNext
                       subHabilitaBotoes
                    Loop
                End If
                
                ResForn.Close
            End If
            EntrouDesc = True
        End If
    End If
End Sub


Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 5 Then
        If TDBGrid1.Columns("Selecionar").Value = "S" Then
            TDBGrid1.Columns("Selecionar").Value = "N"
        Else
            TDBGrid1.Columns("Selecionar").Value = "S"
        End If
    End If
End Sub

Private Sub TDBGrid1_DblClick()
    If Not TDBGrid1.EOF And Not TDBGrid1.BOF Then
        If TDBGrid1.Columns("Selecionar").Value = "N" Then
            TDBGrid1.Columns("Selecionar").Value = "S"
            If LblValor.Caption = "" Then
               LblValor.Caption = Format$(TDBGrid1.Columns("Valor"), "standard")
            Else
               LblValor.Caption = Format$((CDbl(LblValor.Caption) + CDbl(TDBGrid1.Columns("Valor"))), "standard")
            End If
        Else
            TDBGrid1.Columns("Selecionar").Value = "N"
            LblValor.Caption = Format$((CDbl(LblValor.Caption) - CDbl(TDBGrid1.Columns("Valor"))), "standard")
            If LblValor.Caption = "0,00" Then LblValor.Caption = ""
       End If
       TDBGrid1.Refresh
   End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then   'Verifica se o botão da direita foi pressionado
        MsgBox "Esta tela não pode ser filtrada."
    End If
    
    subHabilitaBotoes

End Sub

Private Sub TxtCCorrente2_GotFocus()
    Call subSelecionaTXT(TxtCCorrente2)
End Sub

Private Sub TxtCCorrente2_LostFocus()
    If TxtCCorrente2.Text <> "" Then
        CboContaCorrente.BoundText = TxtCCorrente2.Text
        If CboContaCorrente.Text = "" Then
            MsgBox "Não existe Conta Corrente com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            TxtCCorrente2.SetFocus
        End If
    End If
End Sub
Private Sub TxtValorDuplicata_LostFocus()
    If IsNumeric(TxtValorDuplicata.Text) Then
        TxtValorDuplicata.Text = Format$(TxtValorDuplicata.Text, "##,##0.00")
    Else
        TxtValorDuplicata.Text = ""
    End If
End Sub


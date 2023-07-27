VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form TelaPagamentoComissoes 
   Caption         =   "Pagamento de Comissões"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   1260
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   9840
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanPagamento 
      Height          =   7335
      Left            =   5760
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   9945
      _Version        =   65536
      _ExtentX        =   17542
      _ExtentY        =   12938
      _StockProps     =   15
      Caption         =   "Programação do Pagamento"
      ForeColor       =   16711680
      BackColor       =   15790320
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
      Begin VB.Frame FraPrincipal 
         Height          =   6495
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   9400
         Begin Threed.SSPanel PanCodigo 
            Height          =   1335
            Left            =   120
            TabIndex        =   14
            Top             =   180
            Width           =   9120
            _Version        =   65536
            _ExtentX        =   16087
            _ExtentY        =   2355
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
            Begin VB.TextBox TxtCpfCnpj 
               Alignment       =   2  'Center
               DataField       =   "focl_tx_CgcCpf"
               DataSource      =   "Adodc1"
               Enabled         =   0   'False
               Height          =   330
               Left            =   6910
               TabIndex        =   11
               Text            =   "99.999.999/9999-99"
               Top             =   840
               Width           =   1700
            End
            Begin VB.TextBox TxtForncecedor 
               DataField       =   "corr_tx_Nome"
               DataSource      =   "Adodc1"
               Enabled         =   0   'False
               Height          =   330
               Left            =   1750
               TabIndex        =   24
               Text            =   "Nome do Corretor (Fornecedor)"
               Top             =   840
               Width           =   4120
            End
            Begin VB.TextBox TxtCodEmpd 
               Alignment       =   2  'Center
               DataField       =   "empd_cd_Empreendimento"
               DataSource      =   "Adodc1"
               Enabled         =   0   'False
               Height          =   330
               Left            =   1750
               TabIndex        =   22
               Text            =   "AAAA"
               Top             =   120
               Width           =   615
            End
            Begin VB.TextBox TxtCliente 
               DataField       =   "Cliente"
               DataSource      =   "Adodc1"
               Enabled         =   0   'False
               Height          =   330
               Left            =   2565
               TabIndex        =   21
               Text            =   "Nome do Cliente"
               Top             =   480
               Width           =   4170
            End
            Begin VB.TextBox TxtContrato 
               Alignment       =   2  'Center
               DataField       =   "cont_cd_Contrato"
               DataSource      =   "Adodc1"
               Enabled         =   0   'False
               Height          =   330
               Left            =   7990
               TabIndex        =   20
               Text            =   "AA"
               Top             =   480
               Width           =   620
            End
            Begin VB.TextBox TxtImovel 
               Alignment       =   2  'Center
               DataField       =   "imov_cd_Imovel"
               DataSource      =   "Adodc1"
               Enabled         =   0   'False
               Height          =   330
               Left            =   1750
               TabIndex        =   19
               Text            =   "AAAA"
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox TxNomeEmpd 
               DataField       =   "empd_tx_Nome"
               DataSource      =   "Adodc1"
               Enabled         =   0   'False
               Height          =   330
               Left            =   2565
               TabIndex        =   18
               Text            =   "Nome Empreendimento"
               Top             =   120
               Width           =   6045
            End
            Begin VB.Label LblCpfCnpj 
               Alignment       =   1  'Right Justify
               Caption         =   "CPF/CNPJ:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Left            =   5880
               TabIndex        =   25
               Top             =   905
               Width           =   1020
            End
            Begin VB.Label LblCorretor 
               Alignment       =   1  'Right Justify
               Caption         =   "Corretor:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Left            =   120
               TabIndex        =   23
               Top             =   905
               Width           =   1530
            End
            Begin VB.Line Line1 
               BorderWidth     =   2
               X1              =   2400
               X2              =   2500
               Y1              =   270
               Y2              =   270
            End
            Begin VB.Line Line2 
               BorderWidth     =   2
               X1              =   2400
               X2              =   2500
               Y1              =   630
               Y2              =   630
            End
            Begin VB.Label LblImovel 
               Alignment       =   1  'Right Justify
               Caption         =   "Imóvel:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Left            =   120
               TabIndex        =   17
               Top             =   545
               Width           =   1530
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
               Height          =   250
               Left            =   120
               TabIndex        =   16
               Top             =   165
               Width           =   1530
            End
            Begin VB.Label LblContrato 
               Alignment       =   1  'Right Justify
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
               Left            =   6990
               TabIndex        =   15
               Top             =   545
               Width           =   915
            End
         End
         Begin Threed.SSPanel PanNotasFiscais 
            Height          =   3825
            Left            =   120
            TabIndex        =   26
            Top             =   1560
            Width           =   9120
            _Version        =   65536
            _ExtentX        =   16087
            _ExtentY        =   6747
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
            Begin VB.TextBox TxtStatus 
               Enabled         =   0   'False
               Height          =   330
               Left            =   1500
               TabIndex        =   63
               Text            =   "Compromissado"
               Top             =   120
               Width           =   1815
            End
            Begin VB.TextBox TxtCodTipoDoc 
               Height          =   330
               Left            =   4830
               MaxLength       =   4
               TabIndex        =   59
               Top             =   120
               Width           =   495
            End
            Begin VB.TextBox TxtHistorico 
               Height          =   330
               Left            =   1500
               MaxLength       =   4
               TabIndex        =   41
               Top             =   3045
               Width           =   450
            End
            Begin VB.TextBox TxtSerie 
               Alignment       =   2  'Center
               Height          =   330
               Left            =   3840
               MaxLength       =   3
               TabIndex        =   28
               Text            =   "U"
               Top             =   525
               Width           =   495
            End
            Begin VB.TextBox TxtValor 
               Alignment       =   1  'Right Justify
               DataField       =   "comi_vl_Comissao"
               DataSource      =   "Adodc1"
               Enabled         =   0   'False
               Height          =   315
               Left            =   7200
               TabIndex        =   27
               Text            =   "0,00"
               Top             =   525
               Width           =   1800
            End
            Begin MSComCtl2.DTPicker DtpEmissao 
               Height          =   330
               Left            =   4300
               TabIndex        =   29
               Top             =   935
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   582
               _Version        =   393216
               Format          =   93257729
               CurrentDate     =   36892
            End
            Begin MSMask.MaskEdBox MskNumdocumento 
               Height          =   330
               Left            =   1440
               TabIndex        =   30
               Top             =   525
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   582
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               PromptChar      =   " "
            End
            Begin MSComCtl2.DTPicker DtpEntrada 
               Height          =   330
               Left            =   1500
               TabIndex        =   31
               Top             =   935
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   582
               _Version        =   393216
               Format          =   93257729
               CurrentDate     =   36892
            End
            Begin MSComCtl2.DTPicker DtpVencimento 
               Height          =   330
               Left            =   7550
               TabIndex        =   37
               Top             =   935
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   582
               _Version        =   393216
               Format          =   93257731
               CurrentDate     =   36892
            End
            Begin MSMask.MaskEdBox Mskconta 
               Height          =   330
               Left            =   1500
               TabIndex        =   42
               Top             =   1440
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               PromptChar      =   "_"
            End
            Begin MSDataListLib.DataCombo CboPlanoContas 
               Bindings        =   "TelaPagamentoComissoes.frx":0000
               Height          =   330
               Left            =   2400
               TabIndex        =   43
               Top             =   1440
               Width           =   6230
               _ExtentX        =   10980
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
            Begin MSMask.MaskEdBox MskPcr 
               Height          =   330
               Left            =   1500
               TabIndex        =   44
               Top             =   1845
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo CboPcr 
               Bindings        =   "TelaPagamentoComissoes.frx":0017
               Height          =   330
               Left            =   2400
               TabIndex        =   45
               Top             =   1845
               Width           =   6230
               _ExtentX        =   10980
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
            Begin MSDataListLib.DataCombo CboCCorrente 
               Bindings        =   "TelaPagamentoComissoes.frx":002C
               Height          =   330
               Left            =   1500
               TabIndex        =   46
               Top             =   2235
               Width           =   7115
               _ExtentX        =   12541
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
            Begin Threed.SSCommand CmdLimparCCorrente 
               Height          =   330
               Left            =   8650
               TabIndex        =   47
               Top             =   2235
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   582
               _StockProps     =   78
               MouseIcon       =   "TelaPagamentoComissoes.frx":004B
               Picture         =   "TelaPagamentoComissoes.frx":049D
            End
            Begin Threed.SSCommand CmdLimparPcr 
               Height          =   330
               Left            =   8650
               TabIndex        =   48
               Top             =   1845
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   582
               _StockProps     =   78
               MouseIcon       =   "TelaPagamentoComissoes.frx":05AF
               Picture         =   "TelaPagamentoComissoes.frx":0A01
            End
            Begin Threed.SSCommand CmdLimparplanoContas 
               Height          =   330
               Left            =   8650
               TabIndex        =   49
               Top             =   1440
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   582
               _StockProps     =   78
               MouseIcon       =   "TelaPagamentoComissoes.frx":0B13
               Picture         =   "TelaPagamentoComissoes.frx":0F65
            End
            Begin MSDataListLib.DataCombo CboIdentificador 
               Bindings        =   "TelaPagamentoComissoes.frx":1077
               Height          =   315
               Left            =   1500
               TabIndex        =   50
               Top             =   2640
               Width           =   7115
               _ExtentX        =   12541
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "iden_tx_descricao"
               BoundColumn     =   "iden_cd_Projeto"
               Text            =   ""
            End
            Begin Threed.SSCommand CmdLimparIdentif 
               Height          =   330
               Left            =   8650
               TabIndex        =   51
               Top             =   2640
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   582
               _StockProps     =   78
               MouseIcon       =   "TelaPagamentoComissoes.frx":1096
               Picture         =   "TelaPagamentoComissoes.frx":14E8
            End
            Begin MSDataListLib.DataCombo CboHistorico 
               Bindings        =   "TelaPagamentoComissoes.frx":15FA
               Height          =   315
               Left            =   1980
               TabIndex        =   52
               Top             =   3045
               Width           =   6650
               _ExtentX        =   11721
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               ListField       =   "hist_tx_descricao"
               BoundColumn     =   "hist_cd_codigo"
               Text            =   ""
            End
            Begin Threed.SSCommand CmdLimparHistorico 
               Height          =   330
               Left            =   8650
               TabIndex        =   53
               Top             =   3045
               Width           =   345
               _Version        =   65536
               _ExtentX        =   609
               _ExtentY        =   582
               _StockProps     =   78
               MouseIcon       =   "TelaPagamentoComissoes.frx":1615
               Picture         =   "TelaPagamentoComissoes.frx":1A67
            End
            Begin MSDataListLib.DataCombo CboTipoDoc 
               Bindings        =   "TelaPagamentoComissoes.frx":1B79
               Height          =   315
               Left            =   5355
               TabIndex        =   60
               Top             =   120
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               ListField       =   "tido_tx_descricao"
               BoundColumn     =   "tido_cd_tipodoc"
               Text            =   "CboTipoDoc"
            End
            Begin VB.Label LblStatus1 
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
               Height          =   250
               Left            =   840
               TabIndex        =   62
               Top             =   185
               Width           =   645
            End
            Begin VB.Label LblTipoDoc 
               Caption         =   "Tipo Documento:"
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
               Left            =   3330
               TabIndex        =   61
               Top             =   180
               Width           =   1515
            End
            Begin VB.Label LblIdentificador2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Identificador:"
               DataField       =   "c"
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
               Left            =   330
               TabIndex        =   58
               Top             =   2700
               Width           =   1155
            End
            Begin VB.Label LblConta 
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
               Height          =   255
               Left            =   165
               TabIndex        =   57
               Top             =   1500
               Width           =   1320
            End
            Begin VB.Label LblPcr 
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
               Height          =   255
               Left            =   225
               TabIndex        =   56
               Top             =   1905
               Width           =   1260
            End
            Begin VB.Label LblCCorrente 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   2310
               Width           =   1365
            End
            Begin VB.Label LblHistorico 
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
               Height          =   255
               Left            =   300
               TabIndex        =   54
               Top             =   3105
               Width           =   1125
            End
            Begin VB.Label LblVencimnto 
               Alignment       =   1  'Right Justify
               Caption         =   "Dt. Vencimento:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Left            =   5950
               TabIndex        =   38
               Top             =   1000
               Width           =   1560
            End
            Begin VB.Label LblNumDoc 
               Alignment       =   1  'Right Justify
               Caption         =   "Nº Documento:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Left            =   90
               TabIndex        =   36
               Top             =   590
               Width           =   1365
            End
            Begin VB.Label LblSerie 
               Alignment       =   1  'Right Justify
               Caption         =   "Série:"
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
               Left            =   3120
               TabIndex        =   35
               Top             =   585
               Width           =   675
            End
            Begin VB.Label LblEmissao 
               Alignment       =   1  'Right Justify
               Caption         =   "Dt. Emissão:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Left            =   3160
               TabIndex        =   34
               Top             =   1000
               Width           =   1100
            End
            Begin VB.Label LblEntrada 
               Alignment       =   1  'Right Justify
               Caption         =   "Dt. Entrada:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Left            =   270
               TabIndex        =   33
               Top             =   1000
               Width           =   1185
            End
            Begin VB.Label LblValor1 
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
               Height          =   285
               Left            =   6615
               TabIndex        =   32
               Top             =   585
               Width           =   555
            End
         End
         Begin Threed.SSCommand CmdRetornar 
            Height          =   330
            Left            =   5280
            TabIndex        =   39
            Top             =   5760
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   582
            _StockProps     =   78
            Caption         =   "&Retornar"
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
            Left            =   2040
            TabIndex        =   40
            Top             =   5760
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
      Begin MSAdodcLib.Adodc DatPcr 
         Height          =   330
         Left            =   3600
         Top             =   7800
         Visible         =   0   'False
         Width           =   2115
         _ExtentX        =   3731
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
      Begin MSAdodcLib.Adodc DatConta 
         Height          =   330
         Left            =   6600
         Top             =   7800
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
      Begin MSAdodcLib.Adodc DatContaCorrente 
         Height          =   330
         Left            =   1920
         Top             =   8040
         Visible         =   0   'False
         Width           =   2235
         _ExtentX        =   3942
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
      Begin MSAdodcLib.Adodc DatIdentificador 
         Height          =   330
         Left            =   720
         Top             =   7800
         Visible         =   0   'False
         Width           =   2610
         _ExtentX        =   4604
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
         Caption         =   "DatIdentificador"
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
      Begin MSAdodcLib.Adodc DatHistorico 
         Height          =   330
         Left            =   5280
         Top             =   8040
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
      Begin MSAdodcLib.Adodc DatTipoDoc 
         Height          =   330
         Left            =   8040
         Top             =   7920
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
         Caption         =   "DatTipoDoc"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   3240
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaPagamentoComissoes.frx":1B92
      Height          =   3690
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   19500
      _ExtentX        =   34396
      _ExtentY        =   6509
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "Codigo"
      Columns(0).DataWidth=   4
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Status"
      Columns(1).DataField=   "DocumentoGerado"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Empreendimento"
      Columns(2).DataField=   "empd_tx_nome"
      Columns(2).DataWidth=   50
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Imóvel"
      Columns(3).DataField=   "imov_cd_Imovel"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Contrato"
      Columns(4).DataField=   "cont_cd_Contrato"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Cliente"
      Columns(5).DataField=   "Cliente"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Corretor"
      Columns(6).DataField=   "corr_tx_Nome"
      Columns(6).DataWidth=   65
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Comissão"
      Columns(7).DataField=   "comi_vl_Comissao"
      Columns(7).DataWidth=   15
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Data Pagamento Prevista"
      Columns(8).DataField=   "comi_dt_Pagamento"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Codigo Comissoes"
      Columns(9).DataField=   "comi_cd_Comissoes"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2937"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2858"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=529"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=1958"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=1879"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=529"
      Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(17)=   "Column(1)._MinWidth=4"
      Splits(0)._ColumnProps(18)=   "Column(2).Width=7805"
      Splits(0)._ColumnProps(19)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._WidthInPix=7726"
      Splits(0)._ColumnProps(21)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(2)._ColStyle=528"
      Splits(0)._ColumnProps(23)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(24)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(25)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(26)=   "Column(2)._MinWidth=4"
      Splits(0)._ColumnProps(27)=   "Column(3).Width=2302"
      Splits(0)._ColumnProps(28)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(3)._WidthInPix=2223"
      Splits(0)._ColumnProps(30)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(31)=   "Column(3)._ColStyle=529"
      Splits(0)._ColumnProps(32)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(33)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(34)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(35)=   "Column(4).Width=2381"
      Splits(0)._ColumnProps(36)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(4)._WidthInPix=2302"
      Splits(0)._ColumnProps(38)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(4)._ColStyle=529"
      Splits(0)._ColumnProps(40)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(41)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(42)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(43)=   "Column(5).Width=7276"
      Splits(0)._ColumnProps(44)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(5)._WidthInPix=7197"
      Splits(0)._ColumnProps(46)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(5)._ColStyle=528"
      Splits(0)._ColumnProps(48)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(49)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(50)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(51)=   "Column(6).Width=7250"
      Splits(0)._ColumnProps(52)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(6)._WidthInPix=7170"
      Splits(0)._ColumnProps(54)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(55)=   "Column(6)._ColStyle=528"
      Splits(0)._ColumnProps(56)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(57)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(58)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(59)=   "Column(7).Width=2831"
      Splits(0)._ColumnProps(60)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(7)._WidthInPix=2752"
      Splits(0)._ColumnProps(62)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(63)=   "Column(7)._ColStyle=530"
      Splits(0)._ColumnProps(64)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(65)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(66)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(67)=   "Column(8).Width=4974"
      Splits(0)._ColumnProps(68)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(8)._WidthInPix=4895"
      Splits(0)._ColumnProps(70)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(71)=   "Column(8)._ColStyle=529"
      Splits(0)._ColumnProps(72)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(73)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(74)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(75)=   "Column(8)._MinWidth=107251536"
      Splits(0)._ColumnProps(76)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(77)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(79)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(80)=   "Column(9)._ColStyle=20"
      Splits(0)._ColumnProps(81)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(82)=   "Column(9).Order=10"
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
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15,.alignment=3"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17,.alignment=3"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.alignment=0"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=0"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=66,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(78)  =   "Named:id=33:Normal"
      _StyleDefs(79)  =   ":id=33,.parent=0"
      _StyleDefs(80)  =   "Named:id=34:Heading"
      _StyleDefs(81)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(82)  =   ":id=34,.wraptext=-1"
      _StyleDefs(83)  =   "Named:id=35:Footing"
      _StyleDefs(84)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(85)  =   "Named:id=36:Selected"
      _StyleDefs(86)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(87)  =   "Named:id=37:Caption"
      _StyleDefs(88)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(89)  =   "Named:id=38:HighlightRow"
      _StyleDefs(90)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=39:EvenRow"
      _StyleDefs(92)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(93)  =   "Named:id=40:OddRow"
      _StyleDefs(94)  =   ":id=40,.parent=33"
      _StyleDefs(95)  =   "Named:id=41:RecordSelector"
      _StyleDefs(96)  =   ":id=41,.parent=34"
      _StyleDefs(97)  =   "Named:id=42:FilterBar"
      _StyleDefs(98)  =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   8775
      _Version        =   65536
      _ExtentX        =   15478
      _ExtentY        =   1720
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
         Left            =   4370
         TabIndex        =   8
         Top             =   480
         Width           =   980
      End
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
         Left            =   1280
         TabIndex        =   0
         Top             =   480
         Width           =   980
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
         Left            =   7460
         TabIndex        =   5
         Top             =   480
         Width           =   980
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
         Left            =   6430
         TabIndex        =   4
         Top             =   480
         Width           =   980
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
         Left            =   5400
         TabIndex        =   3
         Top             =   480
         Width           =   980
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
         Left            =   3340
         TabIndex        =   2
         Top             =   480
         Width           =   980
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
         Left            =   2310
         TabIndex        =   1
         Top             =   480
         Width           =   980
      End
      Begin MSDataListLib.DataCombo cboLocalizar 
         Bindings        =   "TelaPagamentoComissoes.frx":1BA7
         Height          =   315
         Left            =   2700
         TabIndex        =   9
         Top             =   0
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Codigo"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.Label lblLocalizar 
         Alignment       =   1  'Right Justify
         Caption         =   "Localizar Código:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   60
         Width           =   2670
      End
   End
End
Attribute VB_Name = "TelaPagamentoComissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************************
'Variaveis que guardarão valores específicos de cada formulário
Dim PrimeiraVez As Boolean

Dim FRM_ContFil As Integer          ' Contador que armazena o número de filtros ativos
Dim FRM_Ordem As String             ' Nome do Campo que tem a chave da Tabela ativa
Dim FRM_Campos_Obrig As String      ' Lista dos campos que são obrigatórios na tabela
Dim FRM_TipoVar As String           ' Tipo do campo a ser filtrado
Dim FRM_FiltroOrig As String        ' Indica os filtros determinados pelo próprio sistema
Dim FRM_FiltroAtual As String       ' Indica o filtro atual da tabela corrente
Dim FRM_filtra_especial As Integer  ' Indica se a tabela utilizada possui filtros pré-definidos pelo sistema em FiltroOrig
Dim FRM_Tem_Filtro As Boolean       ' Informa se já foi relizado um filtro na tabela pelo usuário
Dim FRM_ConTelaOrdem As Integer     ' Indica o número de chaves utilizados para ordenação
Dim FRM_Quantid As Long             ' Quantidade de registros da tabela
Dim FRM_Filtrou As Boolean          ' Informa se Algum campo foi filtrado
Dim FRM_Ordens() As Integer         ' Vetor que mantem a ordem dos campos
Dim FRM_Filtros() As String         ' Array para armazenar os Filtros da Tabela

Dim FRM_BarraFerramentasPressionada As Boolean
Dim FRM_VariasVezes As Boolean      ' Indica se vai executar a insercao na TelaPadrao varias ou uma vez
Dim FRM_XInserir As String          ' Flag para identificar se entrou na rotina para inserir ou alterar ('I'Inserir, 'A'Alterar, 'C'Consultar)
Dim FRM_XOrigem As String           ' Identifica a origem (que formulario) chamou o formulário ativo
Dim FRM_SQLAtivo As String          ' Atualizar o SQL Ativo após modificações de filtro e ordem

'
' Rotina para tratar informações de tela (get/set)
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

'
' Rotina para habilitar os botões da tela inicial e tratar filtro ativo
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

'
' Rotina para desabilitar os botões da tela inicial e filtro ativo
Private Sub subDesabilitaBotoes()

    'Objetivo: Desabilita os Botões da Tela
    CmdInserir.Enabled = False
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    CmdRemoverFiltro.Enabled = False
'    CmdSair.Enabled = False
    CmdImprimir.Enabled = False
    CmdOrdem.Enabled = False
    
End Sub

'
' Rotina para inicialização do registro de Log (desativada)
'Private Sub PreencheInformacoesOriginaisLog()
'
'    ReDim XGM_MATRIZLOG(18, 3) As Variant
'
'    XGM_MATRIZLOG(0, 1) = MskCodigo.Text
'    XGM_MATRIZLOG(1, 1) = TxtNome.Text
'    XGM_MATRIZLOG(2, 1) = FunNuloVal(TxtTotUnidades.Text)
'    XGM_MATRIZLOG(3, 1) = TxtEndereco.Text
'    XGM_MATRIZLOG(4, 1) = TxtBairro.Text
'    XGM_MATRIZLOG(5, 1) = TxtCidade.Text
'    XGM_MATRIZLOG(6, 1) = MskEstado.Text
'    XGM_MATRIZLOG(7, 1) = txtBloco.Text
'    XGM_MATRIZLOG(8, 1) = MskCep.Text
'    XGM_MATRIZLOG(9, 1) = Replace(txtPais.Text, "'", "''")
'    XGM_MATRIZLOG(10, 1) = TxtResponsavel.Text
'    XGM_MATRIZLOG(11, 1) = TxtFone.Text
'    XGM_MATRIZLOG(12, 1) = TxtFax.Text
'    XGM_MATRIZLOG(13, 1) = TxtCga.Text
'    XGM_MATRIZLOG(14, 1) = FunNulo(DatConta.Recordset.Fields("plco_tx_DescContabil")) & " - " & FunNulo(DatConta.Recordset.Fields("plco_cd_Estruturado"))
'    XGM_MATRIZLOG(15, 1) = FunNulo(DatPcr.Recordset.Fields("cere_tx_Descricao")) & " - " & FunNulo(DatPcr.Recordset.Fields("cere_cd_Estruturado"))
'    XGM_MATRIZLOG(16, 1) = FunNulo(CboCCorrente.BoundText)
'    XGM_MATRIZLOG(17, 1) = txtCodigoMunicipioIbge.Text
'
'End Sub

'
' Rotina para gravação do registro de Log (desativada)
'Private Sub GravaLog(XLT_OPERACAO As ENUM_OPERACOES, XLT_ORDEM As String)
'
'    Dim XLI_INDICE As Integer
'
'    If XLT_OPERACAO = EXCLUIR Then
'
'        ReDim XGM_MATRIZLOG(2, 2) As Variant
'
'        XGM_MATRIZLOG(0, 0) = "Código"
'        XGM_MATRIZLOG(1, 0) = "Nome"
'
'        XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("empd_cd_empreendimento")
'        XGM_MATRIZLOG(1, 1) = Adodc1.Recordset.Fields("empd_tx_Nome")
'
'    Else
'
'        If XLT_OPERACAO <> ALTERAR Then
'            ReDim XGM_MATRIZLOG(18, 2) As Variant
'            XLI_INDICE = 1
'        Else
'            XLI_INDICE = 2
'        End If
'
'        'Campos
'        XGM_MATRIZLOG(0, 0) = "Código"
'        XGM_MATRIZLOG(1, 0) = "Nome"
'        XGM_MATRIZLOG(2, 0) = "Total Unidades"
'        XGM_MATRIZLOG(3, 0) = "Endereço"
'        XGM_MATRIZLOG(4, 0) = "Bairro"
'        XGM_MATRIZLOG(5, 0) = "Cidade"
'        XGM_MATRIZLOG(6, 0) = "Estado"
'        XGM_MATRIZLOG(7, 0) = "Bloco"
'        XGM_MATRIZLOG(8, 0) = "Cep"
'        XGM_MATRIZLOG(9, 0) = "Pais"
'        XGM_MATRIZLOG(10, 0) = "Responsável"
'        XGM_MATRIZLOG(11, 0) = "Fone"
'        XGM_MATRIZLOG(12, 0) = "Fax"
'        XGM_MATRIZLOG(13, 0) = "CGA"
'        XGM_MATRIZLOG(14, 0) = "Conta Contábil"
'        XGM_MATRIZLOG(15, 0) = "Centro de Custo"
'        XGM_MATRIZLOG(16, 0) = "Conta Corrente"
'        XGM_MATRIZLOG(17, 0) = "Cód. Município IBGE"
'
'        'Dados
'        XGM_MATRIZLOG(0, XLI_INDICE) = MskCodigo.Text
'        XGM_MATRIZLOG(1, XLI_INDICE) = TxtNome.Text
'        XGM_MATRIZLOG(2, XLI_INDICE) = FunNuloVal(TxtTotUnidades.Text)
'        XGM_MATRIZLOG(3, XLI_INDICE) = TxtEndereco.Text
'        XGM_MATRIZLOG(4, XLI_INDICE) = TxtBairro.Text
'        XGM_MATRIZLOG(5, XLI_INDICE) = TxtCidade.Text
'        XGM_MATRIZLOG(6, XLI_INDICE) = MskEstado.Text
'        XGM_MATRIZLOG(7, XLI_INDICE) = txtBloco.Text
'        XGM_MATRIZLOG(8, XLI_INDICE) = MskCep.Text
'        XGM_MATRIZLOG(9, XLI_INDICE) = Replace(txtPais.Text, "'", "''")
'        XGM_MATRIZLOG(10, XLI_INDICE) = TxtResponsavel.Text
'        XGM_MATRIZLOG(11, XLI_INDICE) = TxtFone.Text
'        XGM_MATRIZLOG(12, XLI_INDICE) = TxtFax.Text
'        XGM_MATRIZLOG(13, XLI_INDICE) = TxtCga.Text
'        XGM_MATRIZLOG(14, XLI_INDICE) = FunNulo(DatConta.Recordset.Fields("plco_tx_DescContabil")) & " - " & FunNulo(DatConta.Recordset.Fields("plco_cd_Estruturado"))
'        XGM_MATRIZLOG(15, XLI_INDICE) = FunNulo(DatPcr.Recordset.Fields("cere_tx_Descricao")) & " - " & FunNulo(DatPcr.Recordset.Fields("cere_cd_Estruturado"))
'        XGM_MATRIZLOG(16, XLI_INDICE) = FunNulo(CboCCorrente.BoundText)
'        XGM_MATRIZLOG(17, XLI_INDICE) = txtCodigoMunicipioIbge.Text
'
'    End If
'
'    Call subRegistraLog("TelaPagamentoComissoes", XLT_ORDEM, funCriaDescricaoLog(XGM_MATRIZLOG, XLT_OPERACAO, EMPREENDIMENTO_CAPI))
'
'End Sub

Private Sub CboCCorrente_Change()

    If CboCCorrente.Text <> "" Then
        DatContaCorrente.Recordset.Bookmark = CboCCorrente.SelectedItem
    End If
    
End Sub

Private Sub CboHistorico_Change()

    If CboHistorico.Text <> "" And Not IsNull(CboHistorico.SelectedItem) Then
    
        DatHistorico.Recordset.Bookmark = CboHistorico.SelectedItem
        TxtHistorico.Text = CboHistorico.BoundText
        
    End If
    
End Sub

Private Sub CboHistorico_KeyPress(KeyAscii As Integer)
    If Len(CboHistorico) = 100 Then
        If KeyAscii <> 8 And CboHistorico.SelLength = 0 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub


Private Sub CboIdentificador_Change()

    If CboIdentificador.Text <> "" Then
        DatIdentificador.Recordset.Bookmark = CboIdentificador.SelectedItem
    End If

End Sub

Private Sub cboLocalizar_Click(Area As Integer)

    If Area = 2 Then
      Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If
    
End Sub

Private Sub cboLocalizar_KeyUp(KeyCode As Integer, Shift As Integer)

    If cboLocalizar.BoundText <> "" And Not IsNull(cboLocalizar.SelectedItem) Then
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If
    
End Sub

Private Sub cboLocalizar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If
    
End Sub


Private Sub CboPcr_Change()

    If CboPcr.Text <> "" Then
    
        DatPCR.Recordset.Bookmark = CboPcr.SelectedItem
        
        If DatPCR.Recordset.Fields("cere_tx_tipo") <> "A" Or DatPCR.Recordset.Fields("cere_tx_status") = "I" Then
        
            MsgBox "Este Centro de Custo não é analítico, ou se encontra inativo, logo não pode ser atribuído a uma Nota.", vbExclamation + vbOKOnly, "ATENÇÃO"
            CboPcr.BoundText = MskPcr.Text
            CboPcr.SetFocus
            Exit Sub
            
        End If
            
        MskPcr.Text = CboPcr.BoundText
     
    End If

End Sub


Private Sub cboPlanoContas_Change()

    If CboPlanoContas.Text <> "" Then
    
        DatConta.Recordset.Bookmark = CboPlanoContas.SelectedItem
        
        If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
            MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbExclamation + vbOKOnly, "ATENÇÃO"
            CboPlanoContas.BoundText = MskConta.Text
            CboPlanoContas.SetFocus
            Exit Sub
        End If
        
        MskConta.Text = CboPlanoContas.BoundText
        
    End If
    
End Sub


Private Sub CboTipoDoc_Change()

    If CboTipoDoc.Text <> "" Then
        
        DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
        TxtCodTipoDoc.Text = CboTipoDoc.BoundText
        
        If Not IsNull(CboTipoDoc.SelectedItem) Then
            DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
        End If
    
    End If
    
End Sub

Private Sub CboTipoDoc_LostFocus()

    Dim ResDoc As Object
    
    If Not IsNull(CboTipoDoc.SelectedItem) Then
        DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
        
        'Desabilita combo de contas correntes - Larissa 31/07/08
        If pVinculaCrCc = True Then
            CboPlanoContas.Enabled = False
        End If
        
    End If
    
'    MskNumdocumento.SetFocus
    
End Sub

Private Sub cmdGravar_Click()

    Dim XLI_HOUVEERRO As Integer
    Dim XLI_LOG       As String
    Dim XLT_SQL       As String
    
    Dim XLT_EMPR_CD   As String       ' Código da Empresa           (empr_cd_Empresa)
    Dim XLT_EMPD_CD   As String       ' Código do Empreendimento    (empd_cd_Empreendimento)
    Dim XLT_IMOV_CD   As String       ' Código do Imóvel            (imov_cd_Imovel)
    Dim XLT_CONT_CD   As String       ' Código do Contrato          (cont_cd_Contrato)
    Dim XLT_CORR_CD   As String       ' Código do Corretor          (corr_cd_Corretor)
    Dim XLT_FOCL_CD   As String       ' Código do Fornecedor        (focl_cd_FornCli)
    Dim XLT_TIDO_CD   As String       ' Código do Tipo Documento    (tido_cd_TipoDoc)
    Dim XLT_CERE_CD   As String       ' Código do Centro de Custo   (cere_cd_Pcr)
    Dim XLT_COCO_CD   As String       ' Código da Conta Corrente    (coco_cd_Codigo)
    Dim XLT_PLCO_CD   As String       ' Código da Conta Contábil    (plco_cd_Conta)
    Dim XLT_IDEN_CD   As String       ' Código do Identificador     (iden_cd_Projeto)
    Dim XLT_COPA_CD   As String       ' Código do Pagamento         (copa_tx_codigo)
    Dim XLT_COMI_CD   As String       ' Código da Comissão          (comi_cd_Comissoes)
    Dim XLT_COPA_NUM  As String       ' Número do Documento         (copa_nr_documento)
    Dim XLT_COPA_SER  As String       ' Série do Documento          (copa_tx_serie)
    Dim XLT_COPA_STD  As String       ' Status do Documento         (copa_tx_statusDocumento)
    Dim XLT_COPA_DTR  As String       ' Data de Registro/Entrada    (copa_dt_entrada)
    Dim XLT_COPA_DTE  As String       ' Data de Emissão             (copa_dt_emissao)
    Dim XLT_COPA_DTV  As String       ' Data de Vencimento          (copa_dt_vencimento)
    Dim XLT_COPA_HIS  As String       ' Histórico                   (copa_tx_historico)
    Dim XLT_COPA_VLR  As String       ' Valor do documento          (copa_vl_valor)
    Dim XLT_COPA_CAP  As String       ' N. titulo no QualiCapi      (copa_nr_titulocapi)
    Dim XLT_COPA_STP  As String       ' Status do Pagamento         (copa_tx_statusPagamento)
    
    XLI_HOUVEERRO = False
    
    '
    ' Valida campos de entrada (Se falso, sai da subrotina)
    If Not ConsisteCampos Then Exit Sub
        
    '
    ' Preenche as variáveis locais com os dados da tela de entrada
    XLT_EMPR_CD = Adodc1.Recordset.Fields("empr_cd_Empresa")
    XLT_EMPD_CD = Adodc1.Recordset.Fields("empd_cd_Empreendimento")
    XLT_IMOV_CD = Adodc1.Recordset.Fields("imov_cd_Imovel")
    XLT_CONT_CD = Adodc1.Recordset.Fields("cont_cd_Contrato")
    XLT_CORR_CD = Adodc1.Recordset.Fields("corr_cd_Corretor")
    XLT_FOCL_CD = Adodc1.Recordset.Fields("focl_cd_FornCli")
    XLT_COMI_CD = Adodc1.Recordset.Fields("comi_cd_Comissoes")
    XLT_TIDO_CD = FunNuloBanco(CboTipoDoc.BoundText)
    XLT_CERE_CD = DatPCR.Recordset.Fields("cere_cd_Pcr")
    XLT_COCO_CD = FunNuloBanco(CboCCorrente.BoundText)
    XLT_PLCO_CD = DatConta.Recordset.Fields("plco_cd_Conta")
    XLT_IDEN_CD = FunNuloBanco(CboIdentificador.BoundText)
    XLT_COPA_CD = Adodc1.Recordset.Fields("Codigo")
    XLT_COPA_NUM = MskNumdocumento.Text
    XLT_COPA_SER = TxtSerie.Text
    XLT_COPA_STD = "C"                                              ' Compromissado
    XLT_COPA_DTR = FunNuloData(DtpEntrada.Value, NomeSgbd)          ' Formato: '06/25/2023'
    XLT_COPA_DTE = FunNuloData(DtpEmissao.Value, NomeSgbd)
    XLT_COPA_DTV = FunNuloData(DtpVencimento.Value, NomeSgbd)
    XLT_COPA_HIS = CboHistorico.Text
    XLT_COPA_VLR = FunTrataFloat(Adodc1.Recordset.Fields("comi_vl_Comissao"))
    XLT_COPA_CAP = Adodc1.Recordset.Fields("Codigo")
    XLT_COPA_STP = "P"                                              ' Pagamento processado

    On Error GoTo RotuloErro
    
    '
    ' Monta o SQL de acordo com a opção (Inserir/Alterar)
    If PanPagamento.Tag = "I" Then     ' Inserção
        XLT_SQL = "INSERT INTO CorretoresPagamento" & _
                  " (empr_cd_Empresa, empd_cd_Empreendimento, imov_cd_Imovel, cont_cd_Contrato, corr_cd_Corretor, focl_cd_FornCli, tido_cd_TipoDoc, cere_cd_Pcr, coco_cd_Codigo, plco_cd_Conta, iden_cd_Projeto, copa_tx_codigo," & _
                  " copa_nr_documento, copa_tx_serie, copa_tx_statusDocumento, copa_dt_entrada, copa_dt_emissao, copa_dt_vencimento, copa_tx_historico, copa_vl_valor, copa_nr_titulocapi, copa_tx_statusPagamento, comi_cd_Comissoes)" & _
                  " VALUES (" & XLT_EMPR_CD & ", '" & XLT_EMPD_CD & "', '" & XLT_IMOV_CD & "', '" & XLT_CONT_CD & "', " & XLT_CORR_CD & ", " & XLT_FOCL_CD & ", " & XLT_TIDO_CD & ", '" & XLT_CERE_CD & "', " & XLT_COCO_CD & ", '" & _
                  XLT_PLCO_CD & "', " & XLT_IDEN_CD & ", '" & XLT_COPA_CD & "', '" & XLT_COPA_NUM & "', '" & XLT_COPA_SER & "', '" & XLT_COPA_STD & "', " & XLT_COPA_DTR & ", " & XLT_COPA_DTE & ", " & XLT_COPA_DTV & ", '" & _
                  XLT_COPA_HIS & "', " & XLT_COPA_VLR & ", '" & XLT_COPA_CAP & "', '" & XLT_COPA_STP & "', " & XLT_COMI_CD & ")"

'        XLI_LOG = "2"
    Else                                ' Alteração
        XLT_SQL = "UPDATE CorretoresPagamento SET" & _
                  " tido_cd_TipoDoc = " & XLT_TIDO_CD & ", cere_cd_Pcr = '" & XLT_CERE_CD & "', coco_cd_Codigo = " & XLT_COCO_CD & ", plco_cd_Conta = '" & XLT_PLCO_CD & "', iden_cd_Projeto = " & XLT_IDEN_CD & ", " & _
                  " copa_tx_codigo = '" & XLT_COPA_CD & "', copa_nr_documento = '" & XLT_COPA_NUM & "', copa_tx_serie = '" & XLT_COPA_SER & "', copa_dt_entrada = " & XLT_COPA_DTR & ", copa_dt_emissao = " & XLT_COPA_DTE & ", " & _
                  " copa_dt_vencimento = " & XLT_COPA_DTV & ", copa_tx_historico = '" & XLT_COPA_HIS & "', copa_vl_valor = " & XLT_COPA_VLR & ", copa_nr_titulocapi = '" & XLT_COPA_CAP & "' " & _
                  "WHERE empr_cd_Empresa = " & XLT_EMPR_CD & _
                  "  AND empd_cd_Empreendimento = '" & XLT_EMPD_CD & "'" & _
                  "  AND imov_cd_Imovel = '" & XLT_IMOV_CD & "'" & _
                  "  AND cont_cd_Contrato = '" & XLT_CONT_CD & "'" & _
                  "  AND corr_cd_Corretor = " & XLT_CORR_CD & _
                  "  AND comi_cd_Comissoes = " & XLT_COMI_CD
                      
'        XLI_LOG = "3"
    End If
    
    Conexao.BeginTrans
    Conexao.Execute XLT_SQL
    Conexao.CommitTrans
    
'    If XLI_LOG = "2" Then
'        GravaLog CADASTRAR, XLI_LOG
'    Else
'        GravaLog ALTERAR, XLI_LOG
'    End If
    
    '
    ' - Bloco com comandos para retorno à 1a. tela (Grid)
    ' ---------------------------------------------------
    PanBotoes.Enabled = True
    
    PanPagamento.Visible = False
    PanPagamento.Enabled = False
    
    subHabilitaBotoes
        
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If

    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    cboLocalizar.BoundText = Chave
    If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    
    Exit Sub
    
RotuloErro:

    XLI_HOUVEERRO = True
    funTrataErros (ComMensagem)
    Conexao.RollbackTrans
    
End Sub

Private Sub CmdLimparCCorrente_Click()

    CboCCorrente.Text = ""
    CboCCorrente.BoundText = ""
    CboCCorrente.SetFocus

End Sub

Private Sub CmdLimparHistorico_Click()
    
    TxtHistorico.Text = ""
    CboHistorico.Text = ""
    CboHistorico.BoundText = ""
    TxtHistorico.SetFocus
    
End Sub

Private Sub CmdLimparIdentif_Click()

    CboIdentificador.Text = ""
    CboIdentificador.BoundText = ""
    CboIdentificador.SetFocus
    
End Sub

Private Sub CmdLimparPcr_Click()

    MskPcr.Mask = ""
    MskPcr.Text = ""
    CboPcr.Text = ""
    CboPcr.BoundText = ""
    MskPcr.SetFocus
    
End Sub

Private Sub CmdLimparplanoContas_Click()

    MskConta.Mask = ""
    MskConta.Text = ""
    CboPlanoContas.Text = ""
    CboPlanoContas.BoundText = ""
    MskConta.SetFocus

End Sub

'
' Rotina para remover filtros de tela
Private Sub CmdRemoverFiltro_Click()

    Dim XLI_FIL As Integer
    
    subDesabilitaBotoes

    Set Formulario = TelaPagamentoComissoes
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

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdRemoverFiltro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MdiPrincipal.BarraStatus.Panels(1).Text = "Mostra os filtros ativos da tabela"
    
End Sub

'
' Rotina que trata o botão Alterar, chamada da tela de cadastro de corretores (implementar)
Private Sub CmdAlterar_Click()
    
    Dim XLT_SQL  As String
    Dim XLO_COPA As ADODB.Recordset

    
    If Adodc1.Recordset.Fields("DocumentoGerado") = "N" Then
        MsgBox "A programação do pagamento de comissão para este imóvel/corretor não foi realizada. Favor verificar!", vbExclamation, "AVISO"
        Exit Sub
    End If
        
    Chave = Adodc1.Recordset.Fields("Codigo")
    
    PanBotoes.Enabled = False
    
    PanPagamento.Visible = True
    PanPagamento.Enabled = True
    PanPagamento.Left = (TDBGrid1.Width - PanPagamento.Width) / 2
    PanPagamento.Top = (TDBGrid1.Height - PanPagamento.Height) / 2
    PanPagamento.Tag = "A"
    
    PanBotoes.Enabled = False
    
    '
    ' Atualiza campos/combos da tela com os dados lidos em CorretoresPagamento
    ' Lê conteúdo do documento gerado anteriormente e carrega os dados
    XLT_SQL = "SELECT A.*, B.plco_cd_codreduzido, B.plco_tx_desccontabil, C.cere_cd_estruturado, C.cere_tx_descricao" & _
              "  FROM CorretoresPagamento   A" & _
              "  LEFT JOIN PlanoContas      B ON  A.plco_cd_Conta = B.plco_cd_Conta" & _
              "  LEFT JOIN CentroResultados C ON  A.cere_cd_Pcr = C.cere_cd_Pcr" & _
              " WHERE A.empr_cd_Empresa = " & Adodc1.Recordset.Fields("empr_cd_Empresa") & _
              "   AND A.empd_cd_Empreendimento = '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "'" & _
              "   AND A.imov_cd_Imovel = '" & Adodc1.Recordset.Fields("imov_cd_Imovel") & "'" & _
              "   AND A.cont_cd_Contrato = '" & Adodc1.Recordset.Fields("cont_cd_Contrato") & "'" & _
              "   AND A.corr_cd_Corretor = " & Adodc1.Recordset.Fields("corr_cd_Corretor") & _
              "   AND A.comi_cd_Comissoes = " & Adodc1.Recordset.Fields("comi_cd_Comissoes")

    SubQOpenRecordset XLO_COPA, XLT_SQL, Estatico
    
    If (XLO_COPA.BOF Or XLO_COPA.EOF) Then
        MsgBox "Erro na leitura do registro de Pagamento dos Corretores. Entre em contato com o Suporte!", vbCritical, "ERRO"
        Exit Sub
    End If
    
    '
    ' Carrega os dados dos Combos e posiciona os dados dos objetos AdoDB em função das informações em XLO_COPA
    If Not IsNull(XLO_COPA!tido_cd_TipoDoc) Then
        CboTipoDoc.BoundText = FunNulo(XLO_COPA!tido_cd_TipoDoc)
        DatTipoDoc.Recordset.Bookmark = CboTipoDoc.SelectedItem
        TxtCodTipoDoc.Text = CboTipoDoc.BoundText
    End If
    
    If Not IsNull(XLO_COPA!plco_cd_CodReduzido) Then
        CboPlanoContas.BoundText = FunNulo(XLO_COPA!plco_cd_CodReduzido)
        DatConta.Recordset.Bookmark = CboPlanoContas.SelectedItem
        MskConta.Text = CboPlanoContas.BoundText
    End If
    
    If Not IsNull(XLO_COPA!cere_cd_estruturado) Then
        CboPcr.BoundText = FunNulo(XLO_COPA!cere_cd_estruturado)
        DatPCR.Recordset.Bookmark = CboPcr.SelectedItem
        MskPcr.Text = CboPcr.BoundText
    End If

    If Not IsNull(XLO_COPA!coco_cd_codigo) Then
        CboCCorrente.BoundText = FunNulo(XLO_COPA!coco_cd_codigo)
        DatContaCorrente.Recordset.Bookmark = CboCCorrente.SelectedItem
    End If
        
    If Not IsNull(XLO_COPA!iden_cd_Projeto) Then
        CboIdentificador.BoundText = FunNulo(XLO_COPA!iden_cd_Projeto)
        DatIdentificador.Recordset.Bookmark = CboIdentificador.SelectedItem
    End If
    
    CboHistorico.Text = FunNulo(XLO_COPA!copa_tx_historico)
    
    '
    ' Carrega demais campos
    MskNumdocumento.Text = XLO_COPA!copa_nr_documento
    TxtSerie.Text = XLO_COPA!copa_tx_serie
    TxtValor.Text = Format$(XLO_COPA!copa_vl_valor, "##,##0.00")
    DtpEntrada.Value = XLO_COPA!copa_dt_entrada
    DtpEmissao.Value = XLO_COPA!copa_dt_emissao
    DtpVencimento.Value = XLO_COPA!copa_dt_vencimento
               
    XLO_COPA.Close
    
    subDesabilitaBotoes

    TDBGrid1.Enabled = False
    
    TxtCodTipoDoc.SetFocus

End Sub

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdAlterar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
  
End Sub

'
' Rotina para tratar o botão de retornar (saída da tela PanPagamento)
Private Sub CmdRetornar_Click()

    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.CancelUpdate
    
    '
    ' - Bloco com comandos para retorno à 1a. tela (Grid)
    ' ---------------------------------------------------
    PanBotoes.Enabled = True
    
    PanPagamento.Visible = False
    PanPagamento.Enabled = False
    
    subHabilitaBotoes
        
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
    End If

    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    cboLocalizar.BoundText = Chave
    If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
            
End Sub

'
' Rotina para tratar o botão de imprimir (impressão do grid)
Private Sub CmdImprimir_Click()

    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Corretores por Empreendimento")
    
End Sub

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MdiPrincipal.BarraStatus.Panels(1).Text = "Imprime os dados da tabela"
    
End Sub

'
' Rotina que trata o botão sair
Private Sub CmdSair_Click()

    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    Unload Me
    
End Sub

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
    
End Sub

'
' Rotina que trata o botão ordem (classificação das linhas no grid)
Private Sub CmdOrdem_Click()
    
    subDesabilitaBotoes
    Set Formulario = TelaPagamentoComissoes
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

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdOrdem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MdiPrincipal.BarraStatus.Panels(1).Text = "Ordena os registros atuais da tabela"
    
End Sub

'
' Rotina que trata o botão excluir
Private Sub CmdExcluir_Click()

    Dim XLT_SQL       As String
    Dim XLI_OPCAO     As Integer
    Dim XLI_HOUVEERRO As Integer
                
    XLI_HOUVEERRO = False
    
    On Error GoTo RotuloErro
    
    If Adodc1.Recordset.Fields("DocumentoGerado") = "N" Then
        MsgBox "A programação do pagamento de comissão para este imóvel/corretor não foi realizada, não pode ser excluida. Favor verificar!", vbExclamation, "AVISO"
        Exit Sub
    End If
        
    Chave = Adodc1.Recordset.Fields("Codigo")
        
    subDesabilitaBotoes
    
    XLI_OPCAO = MsgBox("Confirma Remoção do Registro ?", vbExclamation, "CUIDADO")
    
    If XLI_OPCAO = 1 Then  'Confirma exclusão
    
        '
        ' Exclui o registro em CorretoresPagamento
'        XLT_SQL = "DELETE" & _
'                  "  FROM CorretoresPagamento" & _
'                  " WHERE empr_cd_Empresa = " & Adodc1.Recordset.Fields("empr_cd_Empresa") & _
'                  "   AND empd_cd_Empreendimento = '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "'" & _
'                  "   AND imov_cd_Imovel = '" & Adodc1.Recordset.Fields("imov_cd_Imovel") & "'" & _
'                  "   AND cont_cd_Contrato = '" & Adodc1.Recordset.Fields("cont_cd_Contrato") & "'" & _
'                  "   AND corr_cd_Corretor = " & Adodc1.Recordset.Fields("corr_cd_Corretor")
        
        XLT_SQL = "DELETE" & _
                  "  FROM CorretoresPagamento" & _
                  " WHERE empr_cd_Empresa = " & Adodc1.Recordset.Fields("empr_cd_Empresa") & _
                  "   AND empd_cd_Empreendimento = '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "'" & _
                  "   AND imov_cd_Imovel = '" & Adodc1.Recordset.Fields("imov_cd_Imovel") & "'" & _
                  "   AND cont_cd_Contrato = '" & Adodc1.Recordset.Fields("cont_cd_Contrato") & "'" & _
                  "   AND corr_cd_Corretor = " & Adodc1.Recordset.Fields("corr_cd_Corretor") & _
                  "   AND comi_cd_Comissoes = " & Adodc1.Recordset.Fields("comi_cd_Comissoes")
        
        Conexao.BeginTrans
        Conexao.Execute XLT_SQL
        Conexao.CommitTrans

'        GravaLog EXCLUIR, "4"
       
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
       
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    cboLocalizar.BoundText = Chave
    If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    
    Exit Sub
    
RotuloErro:

    XLI_HOUVEERRO = True
    funTrataErros (ComMensagem)
    Conexao.RollbackTrans
    
    If Err.Number = -2147467259 Or Err.Number = -2147217873 Then
      MsgBox "Esse pagamento a corretor não pode ser excluído!", vbExclamation, "ATENÇÃO"
    End If
    
    subHabilitaBotoes
    
End Sub

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
    
End Sub

'
' Rotina que trata o botão inserir
Private Sub CmdInserir_Click()

    Dim XLT_SQL  As String
    Dim XLO_RES  As ADODB.Recordset
    
    Dim ResEmpd As Object
    Dim ResCere As Object
    Dim ResCoco As Object
    Dim ResPlco As Object

    If Adodc1.Recordset.Fields("DocumentoGerado") <> "N" Then
        MsgBox "A programação do pagamento de comissão para este imóvel/corretor já foi realizada. Favor verificar!", vbExclamation, "AVISO"
        Exit Sub
    End If
    
    Chave = Adodc1.Recordset.Fields("Codigo")
        
    PanBotoes.Enabled = False
    
    PanPagamento.Visible = True
    PanPagamento.Enabled = True
    PanPagamento.Left = (TDBGrid1.Width - PanPagamento.Width) / 2
    PanPagamento.Top = (TDBGrid1.Height - PanPagamento.Height) / 2
    PanPagamento.Tag = "I"
        
    '
    ' Inicializa campos/combos da tela
    MskNumdocumento.Text = ""
    TxtSerie.Text = ""
    TxtValor.Text = Format$(TxtValor.Text, "##,##0.00")
    DtpEntrada.Value = Format(Now, "dd/MM/yyyy")
    DtpEmissao.Value = Format(Now, "dd/MM/yyyy")
    DtpVencimento.Value = Adodc1.Recordset.Fields("comi_dt_Pagamento")

    TxtCodTipoDoc.Text = ""
    CboTipoDoc.Text = ""
    CboTipoDoc.BoundText = ""

    MskConta.Text = ""
    CboPlanoContas.Text = ""
    CboPlanoContas.BoundText = ""

    MskPcr.Text = ""
    CboPcr.Text = ""
    CboPcr.BoundText = ""
    
    CboCCorrente.Text = ""
    CboCCorrente.BoundText = ""

    CboIdentificador.Text = ""
    CboIdentificador.BoundText = ""

    TxtHistorico.Text = ""
    CboHistorico.Text = ""
    CboHistorico.BoundText = ""
    
    '
    ' Carrega combos com valores default a partir de dados do Empreendimento
    XLT_SQL = "SELECT B.plco_cd_codreduzido, C.cere_cd_estruturado, D.coco_cd_Codigo" & _
              "  FROM Empreendimentos       A" & _
              "  LEFT JOIN PlanoContas      B ON  A.plco_cd_Conta = B.plco_cd_Conta" & _
              "  LEFT JOIN CentroResultados C ON  A.cere_cd_Pcr = C.cere_cd_Pcr" & _
              "  LEFT JOIN ContasCorrente   D ON  A.coco_cd_Codigo = D.coco_cd_Codigo" & _
              " WHERE A.empr_cd_Empresa = " & Adodc1.Recordset.Fields("empr_cd_Empresa") & _
              "   AND A.empd_cd_Empreendimento = '" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "'"
               
    SubQOpenRecordset XLO_RES, XLT_SQL, Estatico
    
    If Not (XLO_RES.BOF And XLO_RES.EOF) Then
    
        XLO_RES.MoveFirst
        
'        If Not IsNull(XLO_RES!plco_cd_CodReduzido) Then
'            CboPlanoContas.BoundText = FunNulo(XLO_RES!plco_cd_CodReduzido)
'            DatConta.Recordset.Bookmark = CboPlanoContas.SelectedItem
'            Mskconta.Text = CboPlanoContas.BoundText
'        End If
        
        If Not IsNull(XLO_RES!cere_cd_estruturado) Then
            CboPcr.BoundText = FunNulo(XLO_RES!cere_cd_estruturado)
            DatPCR.Recordset.Bookmark = CboPcr.SelectedItem
            MskPcr.Text = CboPcr.BoundText
        End If
    
        If Not IsNull(XLO_RES!coco_cd_codigo) Then
            CboCCorrente.BoundText = FunNulo(XLO_RES!coco_cd_codigo)
            DatContaCorrente.Recordset.Bookmark = CboCCorrente.SelectedItem
        End If
        
    End If
    
    XLO_RES.Close
    Set XLO_RES = Nothing
    
    subDesabilitaBotoes

    TDBGrid1.Enabled = False
    
    'TxtCodTipoDoc.SetFocus
    
End Sub

'
' Rotina para emitir mensagem na tela para evento de movimento do mouse
Private Sub CmdInserir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MdiPrincipal.BarraStatus.Panels(1).Text = "Insere um novo registro na tabela"
    
End Sub

'
' Rotina para tratar erro no acesso ao banco de dados (Adodc)
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

Private Sub dtpEmissao_LostFocus()

    If Not ConsisteDatas Then Exit Sub
    
End Sub

Private Sub dtpVencimento_LostFocus()

    If Not ConsisteDatas Then Exit Sub
    
End Sub

Private Sub MskConta_GotFocus()

    Call subSelecionaMSK(MskConta)
    
End Sub

Private Sub mskConta_LostFocus()

    MskConta.PromptInclude = False
    
    If MskConta.Text <> "" Then
    
        If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
            MskConta.Text = Format(MskConta.Text, XGT_MSKCODREDUZIDO)
        End If
        
        CboPlanoContas.BoundText = MskConta.Text
        
        If CboPlanoContas.Text = "" Then
            MsgBox "Não existe item do Plano de Contas com este código reduzido, ou não está vinculado ao Centro de Custo.", vbExclamation + vbOKOnly, "ATENÇÃO"
            MskConta.SetFocus
        End If
    
    End If
    
    MskConta.PromptInclude = True
        
End Sub


'
'Rotina para tratar o campo Número (da NF)
Private Sub MskNumdocumento_LostFocus()
    Dim ResFormulario As Object
    
    MskNumdocumento.PromptInclude = False
        
    'Checar se o documento já foi cadastrado
    'If MskNumdocumento.Text <> "" And XInserir = "I" Then
    If MskNumdocumento.Text <> "" Then
        SubQOpenRecordset ResFormulario, "SELECT * FROM NotasFiscais WHERE focl_cd_forncli = " & Adodc1.Recordset.Fields("focl_cd_forncli") & " and nofi_nr_documento = '" & MskNumdocumento.Text & "'", Estatico
        If Not ResFormulario.EOF Then
           MsgBox "Documento já cadastrado.", vbExclamation + vbOKOnly, "ATENÇÃO"
           MskNumdocumento.SetFocus
           ResFormulario.Close
           Exit Sub
        End If
        ResFormulario.Close
    End If
    
    MskNumdocumento.PromptInclude = True
    
'    TxtSerie.SetFocus
    
End Sub

Private Sub MskNumdocumento_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8 And KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 And KeyAscii <> 65 And KeyAscii <> 97) Then
        KeyAscii = 0
    End If
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
        CboPcr.BoundText = MskPcr.Text
        
        If CboPcr.Text = "" Then
            MsgBox "Não existe item do Centro de Custo com este código.", vbExclamation + vbOKOnly, "ATENÇÃO"
            If MskPcr.Visible = True Then MskPcr.SetFocus
        End If
        
    End If
    
    MskPcr.PromptInclude = True
    
End Sub


'
' Rotina para tratar evento de do mouse (DblClick) sobre linha do grid
Private Sub TDBGrid1_DblClick()

    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
    
End Sub

'
' Rotina para tratar evento de do mouse (MouseUp) sobre linha do grid
Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    cboLocalizar.Text = ""
    
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaPagamentoComissoes
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

'
' Rotina para a ativação do formulário chamadas posteriores (retorno à tela)
Private Sub Form_Activate()

    Set Formulario = TelaPagamentoComissoes
    
    If Not PrimeiraVez Then
        subTelaValoresGlobais "P"
    End If
    
    PrimeiraVez = False

End Sub

'
' Rotina para carregar o formulário
Private Sub Form_Load()
    
    Dim XLT_SQL As String
    
    Set Formulario = TelaPagamentoComissoes
    
    Call SubAjustaTela
    
    subManutencaoJanelasAtivas "I", "TelaPagamentoComissoes"
         
    '
    ' Select para carregar o Adodc1
    XLT_SQL = "SELECT A.*" & _
              "      ,'C.' + A.empd_cd_Empreendimento + '.' + A.imov_cd_Imovel + '.' + A.cont_cd_Contrato + '.' + RIGHT('00000' + CONVERT(VARCHAR, A.corr_cd_Corretor), 4) AS Codigo" & _
              "      ,B.corr_tx_Nome" & _
              "      ,C.*" & _
              "      ,D.empd_tx_Nome" & _
              "      ,G.focl_tx_RazaoSocial AS Cliente" & _
              "      ,ISNULL(H.copa_tx_statusPagamento, 'N') AS DocumentoGerado" & _
              "  FROM Comissoes                      A" & _
              "  JOIN Corretores                     B ON A.corr_cd_Corretor       = B.corr_cd_Corretor" & _
              "  JOIN FornClientes                   C ON B.focl_cd_FornCli        = C.focl_cd_FornCli" & _
              "  JOIN Empreendimentos                D ON A.empd_cd_Empreendimento = D.empd_cd_Empreendimento AND A.empr_cd_Empresa = D.empr_cd_Empresa" & _
              "  JOIN CorretoresEmpreendimento       E ON A.empr_cd_Empresa        = E.empr_cd_Empresa AND A.empd_cd_Empreendimento = E.empd_cd_Empreendimento AND A.corr_cd_Corretor = E.corr_cd_Corretor" & _
              "  JOIN Contratos                      F ON A.empr_cd_Empresa        = F.empr_cd_Empresa AND A.empd_cd_Empreendimento = F.empd_cd_Empreendimento AND A.imov_cd_Imovel = F.imov_cd_Imovel AND A.cont_cd_Contrato = F.cont_cd_Contrato" & _
              "  LEFT OUTER JOIN FornClientes        G ON F.focl_cd_FornCli        = G.focl_cd_FornCli" & _
              "  LEFT OUTER JOIN CorretoresPagamento H ON A.comi_cd_Comissoes      = H.comi_cd_Comissoes" & _
              " WHERE A.empr_cd_Empresa = " & PCodEmpresa & _
              "   AND A.comi_dt_Pagamento IS NOT NULL" & _
              "   AND (H.copa_tx_statusPagamento IS NULL OR H.copa_tx_statusPagamento = 'P')" & _
              " ORDER BY A.empd_cd_Empreendimento, A.imov_cd_Imovel, A.cont_cd_Contrato, B.corr_tx_Nome"
                            
    '
    ' Carrega Adodc's (dados para o Grid)
    subCarregaDadosNV Adodc1, "", "", "", XLT_SQL
    
    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and plco_tx_tipo = 'A' and (plco_tx_recdesp= 'D' or plco_tx_recdesp= 'G') AND plco_tx_status = 'A' ORDER BY plco_tx_desccontabil", Estatico
    
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " AND coco_tx_status = 'A' ORDER BY coco_tx_Descricao", Estatico
 
    subConectarControleDadosNV DatHistorico, "SELECT * FROM Historicos WHERE hist_tx_Tipo = 'D' ORDER BY hist_tx_descricao", Estatico
    
    subConectarControleDadosNV DatIdentificador, "SELECT * FROM Identificadores ORDER BY iden_tx_descricao", Estatico

    subConectarControleDadosNV DatPCR, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " AND cere_tx_status = 'A' ORDER BY cere_cd_estruturado", Estatico
 
    subConectarControleDadosNV DatTipoDoc, "SELECT * FROM Tipodocumentos WHERE (tido_tx_sigla IS NULL or tido_tx_sigla='NF' or tido_tx_sigla='NFF' or tido_tx_sigla='NFS'  or tido_tx_sigla='NFFS') and tido_bl_ativo = 1  ORDER BY tido_tx_descricao", Estatico
            
    subHabilitaBotoes
    filtra_especial = True
    
    subTelaValoresGlobais "G"
    PrimeiraVez = True
    
    If FunTabelaVazia(Adodc1) Then
    
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdImprimir.Enabled = False
        CmdInserir.Enabled = False
        CmdOrdem.Enabled = False
        
        MsgBox "Não foram encontrados registros para processar. É necessário a atualização das comissões dos corretores nos contratos dos imóveis, antes de entrar com os dados da programação de pagamento!", vbCritical, "ATENÇÃO !"
        
    End If
        
    If BarraFerramentasPressionada = True Then
      Call CmdInserir_Click
      BarraFerramentasPressionada = False
    End If

End Sub

'
' Rotina para descarregar o formulário
Private Sub Form_Unload(Cancel As Integer)

    subManutencaoJanelasAtivas "R", "TelaPagamentoComissoes"
    
End Sub

Private Sub TxtCodTipoDoc_GotFocus()
    Call subSelecionaTXT(TxtCodTipoDoc)
End Sub

Private Sub TxtCodTipoDoc_LostFocus()
    If TxtCodTipoDoc.Text <> "" Then
        CboTipoDoc.BoundText = TxtCodTipoDoc.Text
        If CboTipoDoc.Text = "" Then
            MsgBox "Não existe Tipo de Documento com este código.", vbExclamation + vbOKOnly, "ATENÇÃO"
            TxtCodTipoDoc.SetFocus
        End If
    End If
End Sub

Private Sub TxtHistorico_GotFocus()

    Call subSelecionaTXT(TxtHistorico)
    
End Sub

Private Sub TxtHistorico_LostFocus()

    If TxtHistorico.Text <> "" Then
    
        CboHistorico.BoundText = TxtHistorico.Text
        
        If CboHistorico.Text = "" Then
            MsgBox "Não existe Histórico com este código.", vbExclamation + vbOKOnly, "ATENÇÃO"
            TxtHistorico.SetFocus
        End If
        
    End If
    
End Sub

Function ConsisteDatas() As Boolean
    'Verifica se as datas seguem a ordem de preenchimento
    'Se ConsisteDatas= TRUE    >> Datas consistentes
    'Se ConsisteDatas= FALSE   >> Datas não consistentes
    ConsisteDatas = False

    If DateDiff("d", DtpEmissao.Value, DtpEntrada.Value) < 0 Then
         MsgBox "A Data de Entrada é anterior à Data de Emissão!", vbExclamation, "ATENÇÃO"
         DtpEmissao.SetFocus
         Exit Function
     End If
    
    If DateDiff("d", DtpEmissao.Value, DtpVencimento.Value) < 0 Then
         MsgBox "A Data de Vencimento é anterior à Data de Emissão!", vbExclamation, "ATENÇÃO"
         DtpVencimento.SetFocus
         Exit Function
     End If

    If DateDiff("d", DtpEntrada.Value, DtpVencimento.Value) < 0 Then
         MsgBox "A Data de Vencimento é anterior à Data de Entrada!", vbExclamation, "ATENÇÃO"
         DtpVencimento.SetFocus
         Exit Function
     End If

   ConsisteDatas = True
   
End Function

'
' Função para validar os campss da tela (se foram imputados)
Function ConsisteCampos() As Boolean

    ConsisteCampos = False
    
    If CboTipoDoc.Text = "" Then
         MsgBox "O Tipo do Documento deve ser informado!", vbExclamation, "ATENÇÃO"
         CboTipoDoc.SetFocus
         Exit Function
     End If
     
    If MskNumdocumento.Text = "" Then
         MsgBox "O Número do Documento deve ser informado!", vbExclamation, "ATENÇÃO"
         MskNumdocumento.SetFocus
         Exit Function
     End If
     
    If TxtSerie.Text = "" Then
         MsgBox "A Série do Documento deve ser informada!", vbExclamation, "ATENÇÃO"
         TxtSerie.SetFocus
         Exit Function
     End If
     
    If Not ConsisteDatas Then Exit Function
     
    If CboPlanoContas.Text = "" Then
         MsgBox "A Conta Contábil deve ser informada!", vbExclamation, "ATENÇÃO"
         CboPlanoContas.SetFocus
         Exit Function
     End If

    If CboPcr.Text = "" Then
         MsgBox "O Centro de Custo deve ser informado!", vbExclamation, "ATENÇÃO"
         CboPcr.SetFocus
         Exit Function
     End If

    If CboCCorrente.Text = "" Then
         MsgBox "A Conta Corrente deve ser informada!", vbExclamation, "ATENÇÃO"
         CboCCorrente.SetFocus
         Exit Function
     End If

    If CboHistorico.Text = "" Then
         MsgBox "O Histórico deve ser informado!", vbExclamation, "ATENÇÃO"
         CboHistorico.SetFocus
         Exit Function
     End If

    ConsisteCampos = True

End Function


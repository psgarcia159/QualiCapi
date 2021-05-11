VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFinAlterarTitulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração Múltipla de Títulos"
   ClientHeight    =   4425
   ClientLeft      =   30
   ClientTop       =   1335
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraPrincipal 
      Height          =   3990
      Left            =   60
      TabIndex        =   23
      Top             =   -60
      Width           =   9225
      Begin TabDlg.SSTab TabAuxiliar 
         Height          =   3255
         Left            =   120
         TabIndex        =   24
         Top             =   660
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5741
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Dados Financeiros"
         TabPicture(0)   =   "FrmAlterarTitulos.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LblObs"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "CmdLimparObs"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "DatObs"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "CboObs"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "FraData"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "FraPlano2"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "FraVlIndexador"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Dados Contábeis"
         TabPicture(1)   =   "FrmAlterarTitulos.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FraConta"
         Tab(1).Control(1)=   "DatPcr"
         Tab(1).Control(2)=   "DatConta"
         Tab(1).Control(3)=   "DatContaCorrente"
         Tab(1).Control(4)=   "DatIdentificador"
         Tab(1).ControlCount=   5
         Begin VB.Frame FraConta 
            Height          =   1815
            Left            =   -74880
            TabIndex        =   40
            Top             =   540
            Width           =   8775
            Begin VB.TextBox txtContratoBanco 
               Height          =   315
               Left            =   7440
               MaxLength       =   3
               TabIndex        =   50
               Top             =   1380
               Width           =   735
            End
            Begin MSMask.MaskEdBox Mskconta 
               Height          =   330
               Left            =   1440
               TabIndex        =   13
               Top             =   180
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               PromptChar      =   "_"
            End
            Begin MSDataListLib.DataCombo CboPlanoContas 
               Bindings        =   "FrmAlterarTitulos.frx":0038
               Height          =   330
               Left            =   2325
               TabIndex        =   14
               Top             =   180
               Width           =   5940
               _ExtentX        =   10478
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
               Left            =   1440
               TabIndex        =   16
               Top             =   570
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo CboPcr 
               Bindings        =   "FrmAlterarTitulos.frx":004F
               Height          =   330
               Left            =   2340
               TabIndex        =   17
               Top             =   570
               Width           =   5925
               _ExtentX        =   10451
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
               Bindings        =   "FrmAlterarTitulos.frx":0064
               Height          =   330
               Left            =   1440
               TabIndex        =   19
               Top             =   960
               Width           =   6825
               _ExtentX        =   12039
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
               Height          =   315
               Left            =   8340
               TabIndex        =   20
               Top             =   960
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmAlterarTitulos.frx":0083
               Picture         =   "FrmAlterarTitulos.frx":04D5
            End
            Begin Threed.SSCommand CmdLimparPcr 
               Height          =   315
               Left            =   8340
               TabIndex        =   18
               Top             =   600
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmAlterarTitulos.frx":05E7
               Picture         =   "FrmAlterarTitulos.frx":0A39
            End
            Begin Threed.SSCommand CmdLimparPlanoContas 
               Height          =   315
               Left            =   8340
               TabIndex        =   15
               Top             =   180
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmAlterarTitulos.frx":0B4B
               Picture         =   "FrmAlterarTitulos.frx":0F9D
            End
            Begin MSDataListLib.DataCombo CboIdentificador 
               Bindings        =   "FrmAlterarTitulos.frx":10AF
               Height          =   315
               Left            =   1440
               TabIndex        =   47
               Top             =   1380
               Width           =   3210
               _ExtentX        =   5662
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "iden_tx_descricao"
               BoundColumn     =   "iden_cd_Projeto"
               Text            =   ""
            End
            Begin Threed.SSCommand CmdLimparIdentif 
               Height          =   315
               Left            =   4680
               TabIndex        =   48
               Top             =   1380
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmAlterarTitulos.frx":10CE
               Picture         =   "FrmAlterarTitulos.frx":1520
            End
            Begin VB.Label lblContratoBanco 
               Caption         =   "Nº do Contrato no Banco:"
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
               Left            =   5160
               TabIndex        =   51
               Top             =   1440
               Width           =   2295
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
               Left            =   240
               TabIndex        =   49
               Top             =   1485
               Width           =   1155
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
               Left            =   60
               TabIndex        =   43
               Top             =   1080
               Width           =   1365
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
               Height          =   195
               Left            =   120
               TabIndex        =   42
               Top             =   660
               Width           =   1260
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
               Height          =   195
               Left            =   120
               TabIndex        =   41
               Top             =   240
               Width           =   1320
            End
         End
         Begin VB.Frame FraVlIndexador 
            Caption         =   "Indexadores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   240
            TabIndex        =   35
            Top             =   1980
            Width           =   8535
            Begin VB.CheckBox ChkReal2 
               Caption         =   "Real"
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
               Left            =   7260
               TabIndex        =   46
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox ChkReal1 
               Caption         =   "Real"
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
               Left            =   3000
               TabIndex        =   45
               Top             =   240
               Width           =   735
            End
            Begin MSDataListLib.DataCombo CboMoeda2 
               Bindings        =   "FrmAlterarTitulos.frx":1632
               Height          =   315
               Left            =   5520
               TabIndex        =   10
               Top             =   240
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "moed_tx_descricao"
               BoundColumn     =   "moed_cd_codmoeda"
               Text            =   "CboMoeda2"
            End
            Begin MSDataListLib.DataCombo CboMoeda1 
               Bindings        =   "FrmAlterarTitulos.frx":164A
               Height          =   315
               Left            =   1200
               TabIndex        =   8
               Top             =   240
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   "moed_tx_descricao"
               BoundColumn     =   "moed_cd_codmoeda"
               Text            =   "CboMoeda1"
            End
            Begin MSAdodcLib.Adodc DatMoeda1 
               Height          =   330
               Left            =   5520
               Top             =   240
               Visible         =   0   'False
               Width           =   1575
               _ExtentX        =   2778
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
               Caption         =   "DatMoeda1"
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
            Begin MSAdodcLib.Adodc DatMoeda2 
               Height          =   330
               Left            =   1200
               Top             =   240
               Visible         =   0   'False
               Width           =   1575
               _ExtentX        =   2778
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
               Caption         =   "DatMoeda2"
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
            Begin Threed.SSCommand CmdLimparIndex1 
               Height          =   315
               Left            =   3780
               TabIndex        =   9
               Top             =   240
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmAlterarTitulos.frx":1662
               Picture         =   "FrmAlterarTitulos.frx":1AB4
            End
            Begin Threed.SSCommand CmdLimparIndex2 
               Height          =   315
               Left            =   8040
               TabIndex        =   11
               Top             =   240
               Width           =   360
               _Version        =   65536
               _ExtentX        =   635
               _ExtentY        =   556
               _StockProps     =   78
               MouseIcon       =   "FrmAlterarTitulos.frx":1BC6
               Picture         =   "FrmAlterarTitulos.frx":2018
            End
            Begin VB.Label LblMoeda2 
               Alignment       =   1  'Right Justify
               Caption         =   "Indexador 2:"
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
               Left            =   4440
               TabIndex        =   37
               Top             =   300
               Width           =   1095
            End
            Begin VB.Label LblMoeda1 
               Alignment       =   1  'Right Justify
               Caption         =   "Indexador 1:"
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
               Left            =   120
               TabIndex        =   36
               Top             =   300
               Width           =   1095
            End
         End
         Begin VB.Frame FraPlano2 
            Height          =   915
            Left            =   240
            TabIndex        =   29
            Top             =   1020
            Width           =   8535
            Begin VB.TextBox TxtJuros 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2160
               TabIndex        =   5
               Top             =   540
               Width           =   975
            End
            Begin VB.TextBox TxtMulta 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6060
               TabIndex        =   6
               Top             =   180
               Width           =   960
            End
            Begin VB.TextBox TxtJurosFin 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2160
               TabIndex        =   4
               Top             =   180
               Width           =   960
            End
            Begin VB.TextBox TxtValorReal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6060
               TabIndex        =   7
               Top             =   540
               Width           =   1515
            End
            Begin VB.Label LblJuros 
               Alignment       =   1  'Right Justify
               Caption         =   "% Juros de Mora:"
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
               Left            =   480
               TabIndex        =   34
               Top             =   600
               Width           =   1620
            End
            Begin VB.Label LblMulta 
               Alignment       =   1  'Right Justify
               Caption         =   "% Multa:"
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
               Left            =   5220
               TabIndex        =   33
               Top             =   240
               Width           =   795
            End
            Begin VB.Label LblJurosFin 
               Alignment       =   1  'Right Justify
               Caption         =   "% Deságio:"
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
               Left            =   60
               TabIndex        =   32
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label LblValorReal 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor Histórico:"
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
               TabIndex        =   31
               Top             =   600
               Width           =   1485
            End
            Begin VB.Label LblReal 
               Caption         =   "R$"
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
               TabIndex        =   30
               Top             =   600
               Width           =   315
            End
         End
         Begin VB.Frame FraData 
            Height          =   555
            Left            =   240
            TabIndex        =   26
            Top             =   420
            Width           =   8535
            Begin MSMask.MaskEdBox MskDiaVencimento 
               Height          =   315
               Left            =   6600
               TabIndex        =   3
               Top             =   180
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   2
               Mask            =   "##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.DTPicker DtpDtBase 
               Height          =   315
               Left            =   1500
               TabIndex        =   2
               Top             =   180
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   69533697
               CurrentDate     =   37180
            End
            Begin VB.Label LblDiaVencimnto 
               Alignment       =   1  'Right Justify
               Caption         =   "Dia do Vencimento:"
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
               Left            =   4800
               TabIndex        =   28
               Top             =   240
               Width           =   1740
            End
            Begin VB.Label LblDtBase 
               Alignment       =   1  'Right Justify
               Caption         =   "Data Base:"
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
               Left            =   420
               TabIndex        =   27
               Top             =   240
               Width           =   1050
            End
         End
         Begin MSAdodcLib.Adodc DatPcr 
            Height          =   330
            Left            =   -70500
            Top             =   2640
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
            Left            =   -68460
            Top             =   2640
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
            Left            =   -74460
            Top             =   2640
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
         Begin MSDataListLib.DataCombo CboObs 
            Bindings        =   "FrmAlterarTitulos.frx":212A
            Height          =   315
            Left            =   1440
            TabIndex        =   12
            Top             =   2820
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "obse_tx_observacao"
            BoundColumn     =   "obse_cd_observacao"
            Text            =   "CboObs"
         End
         Begin MSAdodcLib.Adodc DatObs 
            Height          =   330
            Left            =   6960
            Top             =   2820
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
            Caption         =   "DatObs"
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
         Begin Threed.SSCommand CmdLimparObs 
            Height          =   315
            Left            =   6300
            TabIndex        =   44
            Top             =   2820
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmAlterarTitulos.frx":213F
            Picture         =   "FrmAlterarTitulos.frx":2591
         End
         Begin MSAdodcLib.Adodc DatIdentificador 
            Height          =   330
            Left            =   -72900
            Top             =   2640
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
         Begin VB.Label LblObs 
            Alignment       =   1  'Right Justify
            Caption         =   "Observação:"
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
            Left            =   240
            TabIndex        =   25
            Top             =   2880
            Width           =   1140
         End
      End
      Begin MSMask.MaskEdBox MskCodigoInicial 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   22
         Mask            =   "AAAA.AAAA.##.##.###.##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MskCodigoFinal 
         Height          =   315
         Left            =   6780
         TabIndex        =   1
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   22
         Mask            =   "AAAA.AAAA.##.##.###.##"
         PromptChar      =   " "
      End
      Begin VB.Label LblCodigoInicial 
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
         Left            =   240
         TabIndex        =   39
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label LlbCodigoFinal 
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
         Left            =   5640
         TabIndex        =   38
         Top             =   300
         Width           =   1095
      End
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   5700
      TabIndex        =   22
      Top             =   4020
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
      Left            =   2400
      TabIndex        =   21
      Top             =   4020
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
Attribute VB_Name = "FrmFinAlterarTitulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim errLoop As ADODB.Error
Dim StrError As String
Dim XLT_CODIMOVEL As String
Dim XLT_CODEMPREENDIMENTO As String
Dim XLT_CODCONTRATO As String
Dim XLT_CODPLANO As String
Dim XLT_CODPARCELA As String
Dim XLT_CODRESIDUO As String
Dim XLT_SQL As String
Dim XLD_VENCIMENTO As Date 'Armazena a data do vencimento do título
Dim XLT_PCR As String
Dim XLT_CONTA As String

Sub subImportaFinanceiro()
    
    Dim Titulo As String
     
    'Não atualiza a contabilidade porque compromissos do capi não são provisionados
    'só são contabilizados quando pagos.
   
    
    Titulo = XLT_CODEMPREENDIMENTO & "." & XLT_CODIMOVEL & "." & _
             XLT_CODCONTRATO & "." & XLT_CODPLANO & "." & XLT_CODPARCELA & "." & XLT_CODRESIDUO
                 
    
    'Altera Nota fiscal - Valor
    If Me.TxtValorReal <> "" Then
    
        XLT_SQL = "UPDATE NotasFiscais SET " & _
            " nofi_vl_valor = " & FunTrataFloat(Me.TxtValorReal) & " WHERE nofi_nr_titulocapi = '" & Titulo & "' AND empr_cd_empresa = " & PCodEmpresa
        
        Conexao.Execute XLT_SQL
    
    End If
                              
    'Altera Duplicata - Conta Corrente, dia de Vencimento e Valor
    If Me.CboCCorrente.Text <> "" Then
        
        XLT_SQL = "UPDATE Duplicatas SET " & _
            " coco_cd_codigo = " & FunNuloBancoVal(CboCCorrente.BoundText) & "" _
            & " WHERE dupl_nr_fatura = (SELECT nofi_nr_fatura FROM NotasFiscais  WHERE " & _
            " nofi_nr_titulocapi = '" & Titulo & "' AND empr_cd_empresa = " & PCodEmpresa & ")"
            
        Conexao.Execute XLT_SQL
    
    End If
    
    If Trim(Me.MskDiaVencimento) <> "" Then
        
        XLT_SQL = "UPDATE Duplicatas SET " & _
            " dupl_dt_vencimento = " & FunNuloData(XLD_VENCIMENTO, NomeSgbd) & "" _
            & " WHERE dupl_nr_fatura = (SELECT nofi_nr_fatura FROM NotasFiscais WHERE " & _
            " nofi_nr_titulocapi = '" & Titulo & "' AND empr_cd_empresa = " & PCodEmpresa & ")"
    
        Conexao.Execute XLT_SQL
    
    End If
    
    If Me.TxtValorReal <> "" Then
        
        XLT_SQL = "UPDATE Duplicatas SET " & _
            " dupl_vl_valor = " & FunTrataFloat(TxtValorReal.Text) & "" _
            & " WHERE dupl_nr_fatura = (SELECT nofi_nr_fatura FROM NotasFiscais WHERE " & _
            " nofi_nr_titulocapi = '" & Titulo & "' AND empr_cd_empresa = " & PCodEmpresa & ")"
        
        Conexao.Execute XLT_SQL
    
    End If
    
    'Altera Distribuicao - Plano de Contas, Centro de Custos e Identificador
    If Me.CboPlanoContas.Text <> "" Then
            
        XLT_SQL = "UPDATE Distribuicao SET " & _
            " plco_cd_Conta= " & FunNuloBanco(XLT_CONTA) & _
            " WHERE nofi_cd_notafiscal = (SELECT nofi_cd_notafiscal FROM NotasFiscais WHERE " & _
            " nofi_nr_titulocapi = '" & Titulo & "' AND empr_cd_empresa = " & PCodEmpresa & ")"
            
        Conexao.Execute XLT_SQL
        
    End If
    
    If Me.CboPcr.Text <> "" Then
            
        XLT_SQL = "UPDATE Distribuicao SET " & _
            " cere_cd_pcr = " & FunNuloBanco(XLT_PCR) & _
            " WHERE nofi_cd_notafiscal = (SELECT nofi_cd_notafiscal FROM NotasFiscais WHERE " & _
            " nofi_nr_titulocapi = '" & Titulo & "' AND empr_cd_empresa = " & PCodEmpresa & ")"
            
        Conexao.Execute XLT_SQL
        
    End If
    
    If Me.CboIdentificador <> "" Then
        
        XLT_SQL = "UPDATE Distribuicao SET " & _
            " iden_cd_Projeto= " & FunNuloBancoVal(CboIdentificador.BoundText) & _
            " WHERE nofi_cd_notafiscal = (SELECT nofi_cd_notafiscal FROM NotasFiscais WHERE " & _
            " nofi_nr_titulocapi = '" & Titulo & "' AND empr_cd_empresa = " & PCodEmpresa & ")"
            
        Conexao.Execute XLT_SQL
    
    End If
       
End Sub
Private Sub CboIdentificador_Change()
  
  If CboIdentificador.BoundText <> "" Then
    DatIdentificador.Recordset.bookmark = CboIdentificador.SelectedItem
  End If
  
End Sub

Private Sub CboMoeda1_Change()
   
   If CboMoeda1.Text <> "" Then
     DatMoeda1.Recordset.bookmark = CboMoeda1.SelectedItem
     ChkReal1.Value = 0
   End If
   
End Sub

Private Sub CboMoeda2_Change()
  
  If CboMoeda2.Text <> "" Then
    DatMoeda2.Recordset.bookmark = CboMoeda2.SelectedItem
    ChkReal2.Value = 0
  End If
  
End Sub

Private Sub ChkReal1_Click()
  
  If ChkReal1.Value = 1 Then
    CboMoeda1.BoundText = ""
  End If
  
End Sub

Private Sub ChkReal2_Click()
  
  If ChkReal2.Value = 1 Then
    CboMoeda2.BoundText = ""
  End If
  
End Sub

Private Sub CmdLimparCCorrente_Click()
  CboCCorrente.BoundText = ""
End Sub

Private Sub CmdLimparIdentif_Click()
  
  CboIdentificador.BoundText = ""
  
End Sub

Private Sub CmdLimparIndex1_Click()
  CboMoeda1.BoundText = ""
  ChkReal1.Value = 0
End Sub

Private Sub CmdLimparIndex2_Click()
  CboMoeda2.BoundText = ""
  ChkReal2.Value = 0
End Sub

Private Sub CmdLimparObs_Click()
  CboObs.BoundText = ""
End Sub

Private Sub CmdLimparPcr_Click()
  CboPcr.BoundText = ""
  MskPcr.PromptInclude = False
  MskPcr.Text = ""
  MskPcr.PromptInclude = True
End Sub

Private Sub CmdLimparplanoContas_Click()
  CboPlanoContas.BoundText = ""
  Mskconta.Text = ""
End Sub

Private Sub MskDiaVencimento_LostFocus()
  If MskDiaVencimento <> "" Then
    MskDiaVencimento = Format(MskDiaVencimento, "00")
  End If
End Sub

Private Sub MskDiaVencimento_GotFocus()
     Call subSelecionaMSK(MskDiaVencimento)
End Sub

Private Sub MskCodigoFinal_LostFocus()
  
  Dim XLO_TITULO As New ADODB.Recordset
  
  If MskCodigoFinal.Text <> "    .    .  .  .   .  " Then
    MskCodigoFinal = UCase(MskCodigoFinal)
    SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCAPTitulos WHERE Titulo = '" & MskCodigoFinal.Text & "'", Estatico
      If XLO_TITULO.EOF Then
        MsgBox "O título Final não existe!", vbCritical, "ATENÇÃO !"
        MskCodigoFinal.SetFocus
      End If
    XLO_TITULO.Close
    Set XLO_TITULO = Nothing
  End If
   
End Sub

Private Sub MskCodigoInicial_LostFocus()
  
  Dim XLO_TITULO As New ADODB.Recordset
  
  If MskCodigoInicial.Text <> "    .    .  .  .   .  " Then
    MskCodigoInicial = UCase(MskCodigoInicial)
    SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCAPTitulos WHERE Titulo = '" & MskCodigoInicial.Text & "'", Estatico
      If XLO_TITULO.EOF Then
        MsgBox "O título Inicial não existe!", vbCritical, "ATENÇÃO !"
        MskCodigoInicial.SetFocus
      Else
        If MskCodigoFinal.Text = "    .    .  .  .   .  " Then
          MskCodigoFinal.Text = MskCodigoInicial.Text
        End If
      End If
    XLO_TITULO.Close
    Set XLO_TITULO = Nothing
  End If
  
End Sub

Private Sub Mskconta_LostFocus()
  If Mskconta.Text <> "" Then
    CboPlanoContas.BoundText = Mskconta.Text
    If CboPlanoContas.Text = "" Then
      MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
      Mskconta.SetFocus
    End If
  End If
End Sub

Private Sub CboPcr_Change()
  If CboPcr.Text <> "" Then
    DatPcr.Recordset.bookmark = CboPcr.SelectedItem
    If DatPcr.Recordset.Fields("cere_tx_tipo") <> "A" Then
      MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
      CboPcr.BoundText = MskPcr.Text
      CboPcr.SetFocus
      Exit Sub
    End If
    MskPcr.Text = CboPcr.BoundText
  End If
End Sub

Private Sub CboPlanoContas_Change()
  If CboPlanoContas.Text <> "" Then
    DatConta.Recordset.bookmark = CboPlanoContas.SelectedItem
    If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
      MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
      CboPlanoContas.BoundText = Mskconta.Text
      CboPlanoContas.SetFocus
      Exit Sub
    End If
    Mskconta.Text = CboPlanoContas.BoundText
  End If
End Sub

Private Sub MskPcr_LostFocus()
    MskPcr.PromptInclude = False
    If MskPcr.Text <> "" Then
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

Private Sub CmdDesistir_Click()
    Unload Me
End Sub

Function FunAlteraDia(XLD_VENCIMENTO As Date, XLI_DIAS As Long) As Date
  
  'Altera o dia, mas se ultrapassar o limite do mês vai para último dia do mês
  
  Dim XLD_NOVOVENCIMENTO As Date
  
  'Altera o dia do vencimento
  XLD_NOVOVENCIMENTO = Format(XLD_VENCIMENTO, "mm/yyyy")
  XLD_NOVOVENCIMENTO = DateAdd("d", XLI_DIAS - 1, XLD_NOVOVENCIMENTO)
  
  'Verifica se o dia do vencimento existe naquele mês
  If Month(XLD_VENCIMENTO) <> Month(XLD_NOVOVENCIMENTO) Then
    XLD_NOVOVENCIMENTO = Format(XLD_VENCIMENTO, "mm/yyyy")
    If Month(XLD_VENCIMENTO) = 2 And Year(XLD_VENCIMENTO) Mod 4 = 0 Then
      XLD_NOVOVENCIMENTO = DateAdd("d", 28, XLD_NOVOVENCIMENTO)
    ElseIf Month(XLD_VENCIMENTO) = 2 And Year(XLD_VENCIMENTO) Mod 4 <> 0 Then
       XLD_NOVOVENCIMENTO = DateAdd("d", 27, XLD_NOVOVENCIMENTO)
    Else
       XLD_NOVOVENCIMENTO = DateAdd("d", 29, XLD_NOVOVENCIMENTO)
    End If
  End If
  FunAlteraDia = XLD_NOVOVENCIMENTO

End Function

Private Sub cmdGravar_Click()
        On Error GoTo RotuloErro
    
      Dim XLO_TITULO As New ADODB.Recordset
      Dim XLO_INDEXACAO As New ADODB.Recordset
      Dim XLI_VAZIO As Integer 'Usada para determinar se todos os campos estão vazio
      Dim XLF_VALORINDEXADO1 As Double 'Armazena o valor indexado 1
      Dim XLF_VALORINDEXADO2 As Double 'Armazena o valor indexado 2
      Dim XLT_PRORATA As String 'Armazena se a correção é ProRata
      
      If FunObrigatorioMSK(MskCodigoInicial, "O Código Inicial é Obrigatório") Then Exit Sub
      If FunObrigatorioMSK(MskCodigoFinal, "O Código Final é Obrigatório") Then Exit Sub
           
      If CboPcr.Text = "" Then
          XLT_PCR = ""
      Else
          XLT_PCR = DatPcr.Recordset.Fields("cere_cd_pcr")
      End If
      
      If CboPlanoContas.Text = "" Then
          XLT_CONTA = ""
      Else
          XLT_CONTA = DatConta.Recordset.Fields("plco_cd_conta")
      End If
      
      'Verifica se todos os campos são vazios
      'se forem não executa o update
      XLI_VAZIO = 0
      If CboObs.Text <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If TxtValorReal.Text <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If CboMoeda1.Text <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If ChkReal1.Value = 1 Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If ChkReal2.Value = 1 Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If CboMoeda2.Text <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If TxtJurosFin.Text <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If TxtMulta.Text <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If TxtJuros.Text <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If XLT_PCR <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If XLT_CONTA <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If CboCCorrente.Text <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If CboIdentificador.Text <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If Not IsNull(DtpDtBase.Value) Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
      If MskDiaVencimento.Text <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
        
      If Me.txtContratoBanco.Text <> "" Then
        XLI_VAZIO = XLI_VAZIO + 1
      End If
    
      If XLI_VAZIO = 0 Then
        MsgBox "Todos os campos estão vazios!", vbCritical, "ATENÇÃO "
        Exit Sub
      Else
    
      MsgBox "Os Títulos Pagos não serão alterados.", vbCritical, "ATENÇÃO!"
  
    SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCAPRecalcular WHERE Titulo >= '" & MskCodigoInicial.Text & "' AND Titulo <= '" & MskCodigoFinal.Text & "' AND empr_cd_Empresa = " & PCodEmpresa & " ORDER BY Titulo", Estatico
       
    Conexao.BeginTrans
      
    While Not XLO_TITULO.EOF
      
      XLT_CODEMPREENDIMENTO = Left$(XLO_TITULO!Titulo, 4)
      XLT_CODIMOVEL = Mid$(XLO_TITULO!Titulo, 6, 4)
      XLT_CODCONTRATO = Mid$(XLO_TITULO!Titulo, 11, 2)
      XLT_CODPLANO = Mid$(XLO_TITULO!Titulo, 14, 2)
      XLT_CODPARCELA = Mid$(XLO_TITULO!Titulo, 17, 3)
      XLT_CODRESIDUO = Right$(XLO_TITULO!Titulo, 2)
      
      If MskDiaVencimento.Text <> "" Then
        If MskDiaVencimento.Text > "31" Or MskDiaVencimento.Text = "00" Then
          MsgBox "O dia é inválido", vbCritical, "ATENÇÃO !"
          Exit Sub
        End If
        XLD_VENCIMENTO = FunAlteraDia(XLO_TITULO!titu_dt_Vencimento, CInt(MskDiaVencimento.Text))
      End If
      
      'Cáculo do valor indexado quando há mudança da Data Base ou
       'do valor hitórico ou de algum indexador
      
'      If XLO_TITULO!cont_tx_CorrecProRata = 1 Then
'        XLT_PRORATA = True
'      Else
'        XLT_PRORATA = False
'      End If
       XLT_PRORATA = FunNuloVal(XLO_TITULO!cont_tx_CorrecProRata)
      
      'Se a Data Base não for nula, faz a cálculo com base nessa data
      If Not IsNull(DtpDtBase.Value) Then
         
        'se o valor histórico for vazio, pegar o valor do título
        If TxtValorReal.Text = "" Then
        
          'se o indexador 1 for vazio, pegar o indexador 1 do título
          If CboMoeda1.BoundText = "" Then
            If ChkReal1.Value = 0 Then
              If Not IsNull(XLO_TITULO!moed_cd_Moeda1) Then
                XLF_VALORINDEXADO1 = FunCalcularValorIndexado(XLO_TITULO!titu_vl_Parcela, DtpDtBase.Value, XLO_TITULO!moed_cd_Moeda1, XLO_TITULO!Indexacao1, XLT_PRORATA, TelaCotacoes)
              End If
            Else
              XLF_VALORINDEXADO1 = 0
            End If
          Else
            XLF_VALORINDEXADO1 = FunCalcularValorIndexado(XLO_TITULO!titu_vl_Parcela, DtpDtBase.Value, CboMoeda1.BoundText, DatMoeda1.Recordset.Fields("moed_tx_Indexacao"), XLT_PRORATA, TelaCotacoes)
          End If
          
          'se o indexador 2 for vazio, pegar o indexador 2 do título
          If CboMoeda2.BoundText = "" Then
            If ChkReal2.Value = 0 Then
              If Not IsNull(XLO_TITULO!moed_cd_Moeda2) Then
                XLF_VALORINDEXADO2 = FunCalcularValorIndexado(XLO_TITULO!titu_vl_Parcela, DtpDtBase.Value, XLO_TITULO!moed_cd_Moeda2, XLO_TITULO!Indexacao2, XLT_PRORATA, TelaCotacoes)
              End If
            Else
              XLF_VALORINDEXADO2 = 0
            End If
          Else
            XLF_VALORINDEXADO2 = FunCalcularValorIndexado(XLO_TITULO!titu_vl_Parcela, DtpDtBase.Value, CboMoeda2.BoundText, DatMoeda2.Recordset.Fields("moed_tx_Indexacao"), XLT_PRORATA, TelaCotacoes)
          End If
        
        'se o valor histórico não for vazio
        Else
          
          'se o indexador 1 for vazio, pegar o indexador 1 do título
          If CboMoeda1.BoundText = "" Then
            If ChkReal1.Value = 0 Then
              If Not IsNull(XLO_TITULO!moed_cd_Moeda1) Then
                XLF_VALORINDEXADO1 = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), DtpDtBase.Value, XLO_TITULO!moed_cd_Moeda1, XLO_TITULO!Indexacao1, XLT_PRORATA, TelaCotacoes)
              End If
            Else
              XLF_VALORINDEXADO1 = CDbl(TxtValorReal.Text)
            End If
          Else
            XLF_VALORINDEXADO1 = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), DtpDtBase.Value, CboMoeda1.BoundText, DatMoeda1.Recordset.Fields("moed_tx_Indexacao"), XLT_PRORATA, TelaCotacoes)
          End If
          
          'se o indexador 2 for vazio, pegar o indexador 2 do título
          If CboMoeda2.BoundText = "" Then
            If ChkReal2.Value = 0 Then
              If Not IsNull(XLO_TITULO!moed_cd_Moeda2) Then
                XLF_VALORINDEXADO2 = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), DtpDtBase.Value, XLO_TITULO!moed_cd_Moeda2, XLO_TITULO!Indexacao2, XLT_PRORATA, TelaCotacoes)
              End If
            Else
              XLF_VALORINDEXADO2 = CDbl(TxtValorReal.Text)
            End If
          Else
            XLF_VALORINDEXADO2 = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), DtpDtBase.Value, CboMoeda2.BoundText, DatMoeda2.Recordset.Fields("moed_tx_Indexacao"), XLT_PRORATA, TelaCotacoes)
          End If
          
        End If
         
      'Se a Data Base for nula, pegar a Data Base do título
      Else
        
        'se o valor histórico for vazio, pegar o valor do título
        If TxtValorReal.Text = "" Then
          
          'se o indexador 1 não for vazio
          If ChkReal1.Value = 0 Then
            If CboMoeda1.BoundText <> "" Then
              XLF_VALORINDEXADO1 = FunCalcularValorIndexado(XLO_TITULO!titu_vl_Parcela, XLO_TITULO!titu_dt_Base, CboMoeda1.BoundText, DatMoeda1.Recordset.Fields("moed_tx_Indexacao"), XLT_PRORATA, TelaCotacoes)
            End If
          Else
            XLF_VALORINDEXADO1 = 0
          End If
          
          'se o indexador 2 não for vazio
          If ChkReal2.Value = 0 Then
            If CboMoeda2.BoundText <> "" Then
              XLF_VALORINDEXADO2 = FunCalcularValorIndexado(XLO_TITULO!titu_vl_Parcela, XLO_TITULO!titu_dt_Base, CboMoeda2.BoundText, DatMoeda2.Recordset.Fields("moed_tx_Indexacao"), XLT_PRORATA, TelaCotacoes)
            End If
          Else
            XLF_VALORINDEXADO2 = 0
          End If
        
        'se o valor histórico não for vazio
        Else
          
          'se o indexador 1 for vazio, pegar o indexador 1 do título
          If CboMoeda1.BoundText = "" Then
            If ChkReal1.Value = 0 Then
              If Not IsNull(XLO_TITULO!moed_cd_Moeda1) Then
                XLF_VALORINDEXADO1 = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), XLO_TITULO!titu_dt_Base, XLO_TITULO!moed_cd_Moeda1, XLO_TITULO!Indexacao1, XLT_PRORATA, TelaCotacoes)
              End If
            Else
              XLF_VALORINDEXADO1 = CDbl(TxtValorReal.Text)
            End If
          Else
            XLF_VALORINDEXADO1 = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), XLO_TITULO!titu_dt_Base, CboMoeda1.BoundText, DatMoeda1.Recordset.Fields("moed_tx_Indexacao"), XLT_PRORATA, TelaCotacoes)
          End If
          
          'se o indexador 2 for vazio, pegar o indexador 2 do título
          If CboMoeda2.BoundText = "" Then
            If ChkReal2.Value = 0 Then
              If Not IsNull(XLO_TITULO!moed_cd_Moeda2) Then
                XLF_VALORINDEXADO2 = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), XLO_TITULO!titu_dt_Base, XLO_TITULO!moed_cd_Moeda2, XLO_TITULO!Indexacao2, XLT_PRORATA, TelaCotacoes)
              End If
            Else
              XLF_VALORINDEXADO2 = CDbl(TxtValorReal.Text)
            End If
          Else
            XLF_VALORINDEXADO2 = FunCalcularValorIndexado(CDbl(TxtValorReal.Text), XLO_TITULO!titu_dt_Base, CboMoeda2.BoundText, DatMoeda2.Recordset.Fields("moed_tx_Indexacao"), XLT_PRORATA, TelaCotacoes)
          End If
        
        End If
        
      End If
      
    
      
      XLT_SQL = "UPDATE Titulos SET "
      If CboObs.Text <> "" Then
        XLT_SQL = XLT_SQL + "obse_cd_Observacao=" & FunNuloBancoVal(CboObs.BoundText) & ","
      End If
      If TxtValorReal.Text <> "" Then
        XLT_SQL = XLT_SQL + "titu_vl_Parcela=" & FunNuloVal(FunTrataFloat(TxtValorReal.Text)) & ","
      End If
      If CboMoeda1.Text <> "" Then
        XLT_SQL = XLT_SQL + "moed_cd_Moeda1=" & FunNuloBancoVal(CboMoeda1.BoundText) & ","
      End If
      If CboMoeda2.Text <> "" Then
        XLT_SQL = XLT_SQL + "moed_cd_Moeda2=" & FunNuloBancoVal(CboMoeda2.BoundText) & ","
      End If
      If ChkReal1.Value = 1 Then
        XLT_SQL = XLT_SQL + "moed_cd_Moeda1=Null,"
      End If
      If ChkReal2.Value = 1 Then
        XLT_SQL = XLT_SQL + "moed_cd_Moeda2=Null,"
      End If
      If TxtJurosFin.Text <> "" Then
        XLT_SQL = XLT_SQL + "titu_vl_PercJurosFin=" & FunNuloVal(FunTrataFloat(TxtJurosFin.Text)) & ","
      End If
      If TxtMulta.Text <> "" Then
        XLT_SQL = XLT_SQL + "titu_vl_PercMultaMora=" & FunNuloVal(FunTrataFloat(TxtMulta.Text)) & ","
      End If
      If TxtJuros.Text <> "" Then
        XLT_SQL = XLT_SQL + "titu_vl_PercJurosMora=" & FunNuloVal(FunTrataFloat(TxtJuros.Text)) & ","
      End If
      If Me.txtContratoBanco.Text <> "" Then
        XLT_SQL = XLT_SQL + "titu_nr_contratobanco=" & FunNulo((txtContratoBanco.Text)) & ","
      End If
      If XLT_PCR <> "" Then
        XLT_SQL = XLT_SQL + "cere_cd_Pcr=" & FunNuloBanco(XLT_PCR) & ","
      End If
      If XLT_CONTA <> "" Then
        XLT_SQL = XLT_SQL + "plco_cd_Conta= " & FunNuloBanco(XLT_CONTA) & ","
      End If
      If CboCCorrente.Text <> "" Then
        XLT_SQL = XLT_SQL + "coco_cd_Codigo= " & FunNuloBancoVal(CboCCorrente.BoundText) & ","
      End If
      If CboIdentificador.Text <> "" Then
        XLT_SQL = XLT_SQL + "iden_cd_Projeto= " & FunNuloBancoVal(CboIdentificador.BoundText) & ","
      End If
      If Not IsNull(DtpDtBase.Value) Then
        XLT_SQL = XLT_SQL + "titu_dt_Base=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ","
      End If
      If MskDiaVencimento.Text <> "" Then
        XLT_SQL = XLT_SQL + "titu_dt_Vencimento=" & FunNuloData(XLD_VENCIMENTO, NomeSgbd) & ","
        XLT_SQL = XLT_SQL + "titu_dt_Prorrogacao=" & FunNuloData(XLD_VENCIMENTO, NomeSgbd) & ","
      End If
      If Not IsNull(DtpDtBase) Or TxtValorReal.Text <> "" Or CboMoeda1.BoundText <> "" Or ChkReal1.Value = 1 Then
        XLT_SQL = XLT_SQL + "titu_vl_ParcelaIndex1=" & FunNuloVal(FunTrataFloat(XLF_VALORINDEXADO1)) & ","
      End If
      If Not IsNull(DtpDtBase) Or TxtValorReal.Text <> "" Or CboMoeda2.BoundText <> "" Or ChkReal2.Value = 1 Then
        XLT_SQL = XLT_SQL + "titu_vl_ParcelaIndex2=" & FunNuloVal(FunTrataFloat(XLF_VALORINDEXADO2)) & ","
      End If
        
        'Tira a última vírgula
        XLT_SQL = Left$(XLT_SQL, Len(XLT_SQL) - 1)
        
        XLT_SQL = XLT_SQL + " WHERE empr_cd_Empresa=" & PCodEmpresa & " AND " & _
                "empd_cd_empreendimento='" & XLT_CODEMPREENDIMENTO & "' AND " & _
                "imov_cd_imovel='" & XLT_CODIMOVEL & "' AND " & _
                "cont_cd_Contrato='" & XLT_CODCONTRATO & "' AND " & _
                "titu_cd_Plano= '" & XLT_CODPLANO & "' AND " & _
                "titu_cd_Parcela= '" & XLT_CODPARCELA & "' AND " & _
                "titu_cd_Residuo= '" & XLT_CODRESIDUO & "'"
      
      Conexao.Execute (XLT_SQL)
      
      'Verifica se o cliente tem o sistema financeiro
      'Verifica se utiliza exportação de compromissos para o financeiro
      '*****************Alterando o título no financeiro************
      If XGB_PQUALIFIN And XGT_EXPCOMPROMCAPI = "S" Then
        Call subImportaFinanceiro
      End If
      
      XLO_TITULO.MoveNext
      
    Wend
    
    Conexao.CommitTrans
          
    XLO_TITULO.Close
    Set XLO_TITULO = Nothing
    
  End If
  
  CmdDesistir_Click
  
  Exit Sub
  
RotuloErro:

  Conexao.RollbackTrans
  funTrataErros ComMensagem
     
End Sub

Private Sub Form_Activate()
    Set Formulario = FrmFinAlterarTitulos
End Sub

Private Sub Form_Load()
    Dim XLT_SQL  As String
   
    Set Formulario = FrmFinAlterarTitulos
   
    subConectarControleDadosNV DatMoeda1, "SELECT * FROM Moedas ORDER BY moed_tx_descricao", Estatico
    subConectarControleDadosNV DatMoeda2, "SELECT * FROM Moedas ORDER BY moed_tx_descricao", Estatico
    subConectarControleDadosNV DatObs, "SELECT * FROM Observacoes ORDER BY obse_tx_observacao", Estatico
    subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_cd_estruturado", Estatico
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico
    subConectarControleDadosNV DatIdentificador, "SELECT * FROM Identificadores ORDER BY iden_tx_descricao", Estatico
    
    MskCodigoInicial.Text = Chave

End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "FrmFinAlterarTitulos"
End Sub

Private Sub TxtJuros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtJuros_LostFocus()
    If IsNumeric(TxtJuros.Text) Then
        TxtJuros.Text = Format$(TxtJuros.Text, "##0.0000")
    Else
        TxtJuros.Text = ""
    End If
End Sub

Private Sub TxtJurosFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtJurosFin_LostFocus()
    If IsNumeric(TxtJurosFin.Text) Then
        TxtJurosFin.Text = Format$(TxtJurosFin.Text, "##0.0000")
    Else
        TxtJurosFin.Text = ""
    End If
End Sub

Private Sub TxtMulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtMulta_LostFocus()
    If IsNumeric(TxtMulta.Text) Then
        TxtMulta.Text = Format$(TxtMulta.Text, "##0.0000")
    Else
        TxtMulta.Text = ""
    End If
End Sub

Private Sub TxtValorReal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtValorReal_LostFocus()
    If IsNumeric(TxtValorReal.Text) Then
        TxtValorReal.Text = Format$(TxtValorReal.Text, "##,##0.00")
    Else
        TxtValorReal.Text = ""
    End If
End Sub

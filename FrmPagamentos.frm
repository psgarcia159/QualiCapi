VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFinPagamentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagamento de Parcela"
   ClientHeight    =   6375
   ClientLeft      =   1800
   ClientTop       =   2190
   ClientWidth     =   9375
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9375
   Begin VB.Frame FraDadosGerais 
      Height          =   1200
      Left            =   45
      TabIndex        =   29
      Top             =   -90
      Width           =   9300
      Begin MSDataListLib.DataCombo CboImovel 
         Bindings        =   "FrmPagamentos.frx":0000
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   540
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Imovel"
         BoundColumn     =   "imov_cd_imovel"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox MskEmpreendimento 
         Height          =   315
         Left            =   1650
         TabIndex        =   0
         Top             =   165
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "AAAA"
         PromptChar      =   " "
      End
      Begin MSDataListLib.DataCombo CboEmpreendimento 
         Bindings        =   "FrmPagamentos.frx":0018
         Height          =   315
         Left            =   2370
         TabIndex        =   1
         Top             =   165
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "empd_tx_nome"
         BoundColumn     =   "empd_cd_empreendimento"
         Text            =   "CboEmpreendimento"
      End
      Begin MSAdodcLib.Adodc DatEmpreendimento 
         Height          =   330
         Left            =   2580
         Top             =   225
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
      Begin MSAdodcLib.Adodc DatImovel 
         Height          =   330
         Left            =   2880
         Top             =   480
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
         Caption         =   "DatImovel"
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
         Left            =   7965
         TabIndex        =   40
         Top             =   630
         Width           =   795
      End
      Begin VB.Label LblContrato2 
         Caption         =   "01"
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
         Left            =   8805
         TabIndex        =   39
         Top             =   630
         Width           =   255
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
         Height          =   210
         Left            =   780
         TabIndex        =   38
         Top             =   600
         Width           =   810
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
         Height          =   240
         Left            =   60
         TabIndex        =   37
         Top             =   210
         Width           =   1545
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
         Height          =   240
         Left            =   90
         TabIndex        =   35
         Top             =   915
         Width           =   1365
      End
      Begin VB.Label LblDtVenda 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Venda:"
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
         Left            =   3885
         TabIndex        =   34
         Top             =   915
         Width           =   1095
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
         Height          =   195
         Left            =   1530
         TabIndex        =   33
         Top             =   915
         Width           =   2190
      End
      Begin VB.Label LblDtVenda2 
         Caption         =   "01/11/2002"
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
         Height          =   240
         Left            =   5040
         TabIndex        =   32
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label LblDtContrato 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Base Contrato:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6255
         TabIndex        =   31
         Top             =   915
         Width           =   1755
      End
      Begin VB.Label LblDtContrato2 
         Caption         =   "01/11/2002"
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
         Left            =   8070
         TabIndex        =   30
         Top             =   930
         Width           =   1140
      End
   End
   Begin VB.Frame FraPagamento 
      Height          =   4770
      Left            =   45
      TabIndex        =   23
      Top             =   1065
      Width           =   9315
      Begin VB.TextBox TxtTitulo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   260
         Left            =   1830
         TabIndex        =   4
         Top             =   180
         Width           =   6960
      End
      Begin MSAdodcLib.Adodc DatTitulos 
         Height          =   330
         Left            =   3735
         Top             =   165
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
      Begin MSDataListLib.DataCombo CboTitulos 
         Bindings        =   "FrmPagamentos.frx":0038
         Height          =   330
         Left            =   1785
         TabIndex        =   3
         Top             =   150
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Parcela"
         BoundColumn     =   "titulo"
         Text            =   "CboTitulos"
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
      Begin VB.Frame FraDatas 
         Enabled         =   0   'False
         Height          =   555
         Left            =   75
         TabIndex        =   24
         Top             =   705
         Width           =   9165
         Begin MSComCtl2.DTPicker DtpPagamento 
            Height          =   315
            Left            =   4440
            TabIndex        =   5
            Top             =   165
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   84344835
            CurrentDate     =   37180
         End
         Begin MSComCtl2.DTPicker DtpBasePagto 
            Height          =   315
            Left            =   7635
            TabIndex        =   6
            Top             =   165
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   84344835
            CurrentDate     =   37180
         End
         Begin MSComCtl2.DTPicker DtpVencimento 
            Height          =   315
            Left            =   1710
            TabIndex        =   25
            Top             =   165
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yy"
            Format          =   84344835
            CurrentDate     =   37180
         End
         Begin MSComCtl2.DTPicker DtpProrrogacao 
            Height          =   315
            Left            =   3360
            TabIndex        =   78
            Top             =   165
            Visible         =   0   'False
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yy"
            Format          =   84344835
            CurrentDate     =   37180
         End
         Begin VB.Label LblProrrogacao 
            Alignment       =   1  'Right Justify
            Caption         =   "Prorrogação:"
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
            Left            =   2250
            TabIndex        =   79
            Top             =   210
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label LblDtPagamento 
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
            Left            =   3420
            TabIndex        =   28
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label LblDtBasePagto 
            Caption         =   "Data Base Pagto:"
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
            Left            =   6120
            TabIndex        =   27
            Top             =   210
            Width           =   1575
         End
         Begin VB.Label LblVencimento 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   630
            TabIndex        =   26
            Top             =   210
            Width           =   1080
         End
      End
      Begin TabDlg.SSTab TabFinanceiro 
         Height          =   3375
         Left            =   75
         TabIndex        =   41
         Top             =   1320
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   5953
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         Enabled         =   0   'False
         TabCaption(0)   =   "Valores"
         TabPicture(0)   =   "FrmPagamentos.frx":0051
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LblValor2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "LblValor"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "PanPagto"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "FraValores1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "FraValores2"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Dados Bancários"
         TabPicture(1)   =   "FrmPagamentos.frx":006D
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FraDadosBancarios"
         Tab(1).Control(1)=   "CboCCorrente"
         Tab(1).Control(2)=   "DatContaCorrente"
         Tab(1).Control(3)=   "LblCCorrente"
         Tab(1).ControlCount=   4
         Begin VB.Frame FraValores2 
            Height          =   1875
            Left            =   5460
            TabIndex        =   61
            Top             =   345
            Width           =   3645
            Begin VB.TextBox TxtMulta 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   765
               TabIndex        =   10
               Top             =   1305
               Width           =   945
            End
            Begin VB.TextBox TxtJuros 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   765
               TabIndex        =   9
               Top             =   765
               Width           =   945
            End
            Begin VB.TextBox TxtSeguro 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   765
               TabIndex        =   8
               Top             =   225
               Width           =   945
            End
            Begin VB.TextBox TxtDesagio 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   2640
               TabIndex        =   62
               Top             =   225
               Width           =   945
            End
            Begin VB.TextBox TxtDesconto 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2640
               TabIndex        =   11
               Top             =   765
               Width           =   945
            End
            Begin VB.TextBox TxtOutros 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2640
               TabIndex        =   12
               Top             =   1305
               Width           =   945
            End
            Begin VB.Label LblMulta 
               Caption         =   "Multa:"
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
               Left            =   210
               TabIndex        =   68
               Top             =   1365
               Width           =   555
            End
            Begin VB.Label LblJuros 
               Caption         =   "Juros:"
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
               TabIndex        =   67
               Top             =   810
               Width           =   555
            End
            Begin VB.Label LblSeguro 
               Caption         =   "Seguro:"
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
               Left            =   75
               TabIndex        =   66
               Top             =   285
               Width           =   675
            End
            Begin VB.Label LblDesagio 
               Alignment       =   1  'Right Justify
               Caption         =   "Deságio:"
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
               Left            =   1830
               TabIndex        =   65
               Top             =   285
               Width           =   795
            End
            Begin VB.Label LblDesconto 
               Alignment       =   1  'Right Justify
               Caption         =   "Desconto:"
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
               Left            =   1695
               TabIndex        =   64
               Top             =   810
               Width           =   930
            End
            Begin VB.Label LblOutros 
               Alignment       =   1  'Right Justify
               Caption         =   "Outros(+):"
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
               Left            =   1785
               TabIndex        =   63
               Top             =   1365
               Width           =   840
            End
         End
         Begin VB.Frame FraValores1 
            Height          =   1875
            Left            =   60
            TabIndex        =   48
            Top             =   345
            Width           =   5370
            Begin VB.OptionButton OptValor1 
               Height          =   195
               Left            =   5070
               TabIndex        =   60
               Top             =   840
               Value           =   -1  'True
               Width           =   225
            End
            Begin VB.OptionButton OptValor2 
               Height          =   255
               Left            =   5070
               TabIndex        =   59
               Top             =   1350
               Width           =   195
            End
            Begin VB.TextBox TxtValorReal 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   1575
               TabIndex        =   53
               Top             =   225
               Width           =   1080
            End
            Begin VB.CheckBox ChkCorrecao 
               Caption         =   "Pagto. c/ Parcela Corrigida"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   315
               Left            =   2685
               TabIndex        =   7
               Top             =   210
               Value           =   1  'Checked
               Width           =   2655
            End
            Begin VB.TextBox TxtValorMoeda2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   1575
               TabIndex        =   52
               Top             =   1305
               Width           =   1080
            End
            Begin VB.TextBox TxtValorMoeda1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   1575
               TabIndex        =   51
               Top             =   765
               Width           =   1080
            End
            Begin VB.TextBox TxtVlCorrigido2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   3960
               TabIndex        =   50
               Top             =   1305
               Width           =   1080
            End
            Begin VB.TextBox TxtVlCorrigido1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   3960
               TabIndex        =   49
               Top             =   765
               Width           =   1080
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
               Left            =   210
               TabIndex        =   58
               Top             =   270
               Width           =   1365
            End
            Begin VB.Label LblValorMoeda2 
               Alignment       =   1  'Right Justify
               Caption         =   "Correção:"
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
               Left            =   45
               TabIndex        =   57
               Top             =   1365
               Width           =   1530
            End
            Begin VB.Label LblValorMoeda1 
               Alignment       =   1  'Right Justify
               Caption         =   "Correção:"
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
               Left            =   45
               TabIndex        =   56
               Top             =   810
               Width           =   1530
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor Atual:"
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
               Left            =   2775
               TabIndex        =   55
               Top             =   1365
               Width           =   1155
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor Atual:"
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
               Left            =   2895
               TabIndex        =   54
               Top             =   810
               Width           =   1020
            End
         End
         Begin VB.Frame FraDadosBancarios 
            Caption         =   "Dados do Pagamento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2265
            Left            =   -74805
            TabIndex        =   44
            Top             =   450
            Width           =   8745
            Begin VB.TextBox TxtDocPagto 
               Height          =   1275
               Left            =   2280
               MaxLength       =   250
               MultiLine       =   -1  'True
               TabIndex        =   17
               Top             =   735
               Width           =   6300
            End
            Begin VB.ComboBox CboLocalPagto 
               Height          =   315
               ItemData        =   "FrmPagamentos.frx":0089
               Left            =   4800
               List            =   "FrmPagamentos.frx":0099
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   285
               Width           =   2040
            End
            Begin VB.CheckBox ChkPreDatado 
               Caption         =   "Pré-Datado"
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
               Left            =   7260
               TabIndex        =   15
               Top             =   300
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DtpDeposito 
               Height          =   300
               Left            =   1020
               TabIndex        =   14
               Top             =   315
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "dd/MM/yy"
               Format          =   84344835
               CurrentDate     =   37180
            End
            Begin VB.Label LblDocPagto 
               Caption         =   "Documento Pagamento:"
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
               Left            =   195
               TabIndex        =   47
               Top             =   795
               Width           =   2115
            End
            Begin VB.Label LblLocalPgto 
               Caption         =   "Forma de Pagamento:"
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
               Left            =   2895
               TabIndex        =   46
               Top             =   330
               Width           =   1860
            End
            Begin VB.Label LblDtDeposito 
               Caption         =   "Depósito:"
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
               TabIndex        =   45
               Top             =   345
               Width           =   915
            End
         End
         Begin Threed.SSPanel PanPagto 
            Height          =   450
            Left            =   4620
            TabIndex        =   42
            Top             =   2520
            Width           =   4470
            _Version        =   65536
            _ExtentX        =   7885
            _ExtentY        =   794
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
            Begin VB.TextBox TxtValorPago 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2205
               TabIndex        =   13
               Top             =   75
               Width           =   1770
            End
            Begin VB.Label LblIndicePagamento 
               Caption         =   "Valor do Pagamento:"
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
               Height          =   255
               Left            =   375
               TabIndex        =   43
               Top             =   105
               Width           =   1890
            End
         End
         Begin MSDataListLib.DataCombo CboCCorrente 
            Bindings        =   "FrmPagamentos.frx":00CB
            Height          =   330
            Left            =   -73320
            TabIndex        =   18
            Top             =   2880
            Width           =   7035
            _ExtentX        =   12409
            _ExtentY        =   582
            _Version        =   393216
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
         Begin MSAdodcLib.Adodc DatContaCorrente 
            Height          =   330
            Left            =   -73320
            Top             =   2880
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
         Begin VB.Label LblValor 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor do Resíduo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   255
            Left            =   90
            TabIndex        =   77
            Top             =   2625
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.Label LblValor2 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1740
            TabIndex        =   76
            Top             =   2595
            Visible         =   0   'False
            Width           =   1245
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
            Height          =   270
            Left            =   -74760
            TabIndex        =   69
            Top             =   2895
            Width           =   1365
         End
      End
      Begin VB.Label LblDtBaseTitulo2 
         Caption         =   "01/11/2002"
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
         Left            =   8085
         TabIndex        =   74
         Top             =   525
         Width           =   1035
      End
      Begin VB.Label LblDtBaseTitulo 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Base Título:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6525
         TabIndex        =   73
         Top             =   510
         Width           =   1545
      End
      Begin VB.Label LblObs2 
         Caption         =   "Título em carteira"
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
         Height          =   225
         Left            =   1830
         TabIndex        =   71
         Top             =   525
         Width           =   4620
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
         Left            =   645
         TabIndex        =   70
         Top             =   510
         Width           =   1125
      End
      Begin VB.Label LblTitulosPend 
         Alignment       =   1  'Right Justify
         Caption         =   "Parcela a Pagar:"
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
         TabIndex        =   36
         Top             =   195
         Width           =   1710
      End
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   7875
      TabIndex        =   22
      Top             =   5880
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
      Left            =   1800
      TabIndex        =   20
      Top             =   5880
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
   Begin Threed.SSCommand CmdNovo 
      Height          =   330
      Left            =   180
      TabIndex        =   19
      Top             =   5880
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
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
      Enabled         =   0   'False
   End
   Begin Threed.SSCommand CmdRecibo 
      Height          =   330
      Left            =   3255
      TabIndex        =   21
      Top             =   5880
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Recibo"
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
   Begin Threed.SSCommand CmdCancelar 
      Height          =   330
      Left            =   4800
      TabIndex        =   72
      Top             =   5880
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Contrato"
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
   Begin Threed.SSCommand CmdTitulo 
      Height          =   330
      Left            =   6330
      TabIndex        =   75
      Top             =   5880
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Título"
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
Attribute VB_Name = "FrmFinPagamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XFB_CALCULA As Boolean 'Variável de controle usada para indicar se vai calcular ou não
  'devido ao cálculo dos campos ser efetuado no change ou no lost focus dos controles
Dim XLO_PERMISSAODESCONTO As ADODB.Recordset
Option Explicit
Function funChecarCampos() As Boolean
    funChecarCampos = True
    
    If XGB_PQUALIFIN Then 'Integrado com o Financeiro
        If FunObrigatorioCBO(CboCCorrente, "A Conta Corrente é Obrigatória") Then Exit Function
    End If

    funChecarCampos = False
End Function

Sub subCalculaCampos(XLB_CALCULASEGURO As Boolean)
    Dim XLT_SQL As String           'Armazena SQL's
    Dim XLF_VALOR As Double         'Armazena o valor sobre o qual deverá incidir juros
    Dim XLO_TITULO As ADODB.Recordset 'Tabela com os dados do contrato e dos títulos correspondentes
    Dim XLF_VLCORRIGIDO As Double   'Valor corrigido para ser utilizado no cálculo do seguro com base no saldo devedor
    Dim XLF_VLTOTCORRIGIDO As Double 'Valor total corrigido e desagiado para ser utilizado no cálculo do seguro com base no saldo devedor
    Dim XLF_TAXASEGURO As Double    'Taxa de seguro
    Dim XLF_VLSEGURO As Double      'Valor do Seguro calculado com base no Saldo Devedor
    Dim XLT_TIPOSEGURO As String    'Tipo de seguro cadastrado no contrato 1-Pelo Saldo Devedor 2-Pela Parcela ""Não possui seguro
    Dim XLB_CALCSEGUROS As String   'Informa se deseja ou nao calcular o seguro
    Dim XLF_JUROS As Double         'Valor calculado na funcao jurosmulta
    Dim XLF_MULTA As Double         'Valor calculado na funcao jurosmulta
    
    Dim XLO_PAGAMENTO As New ADODB.Recordset
    Dim XLT_SQLBASE As String
    Dim XLF_VALORCORRIGIDO As Double
    
    XLT_SQLBASE = FunCriaConsultaBase(DtpBasePagto, NomeSgbd, "ConsCAPBasePagamento", 1)
    XLT_SQLBASE = XLT_SQLBASE & " empr_cd_empresa=" & PCodEmpresa
    XLT_SQL = XLT_SQLBASE & " AND Titulo = '" & DatTitulos.Recordset!Titulo & "'"
    
    SubQOpenRecordset XLO_PAGAMENTO, XLT_SQL, Estatico
    
    While Not XLO_PAGAMENTO.EOF
      If Not funVerirficaCotacao(XLO_PAGAMENTO, 1, DtpBasePagto.Value, True, TelaCotacoes) Then
        XLO_PAGAMENTO.Requery adCmdText
        If Not funVerirficaCotacao(XLO_PAGAMENTO, 1, DtpBasePagto.Value, False, TelaCotacoes) Then
          Exit Sub
        End If
      End If
      
      If Not funVerirficaCotacao(XLO_PAGAMENTO, 2, DtpBasePagto.Value, True, TelaCotacoes) Then
        XLO_PAGAMENTO.Requery adCmdText
        If Not funVerirficaCotacao(XLO_PAGAMENTO, 2, DtpBasePagto.Value, False, TelaCotacoes) Then
          Exit Sub
        End If
      End If
      XLO_PAGAMENTO.MoveNext
    Wend
    
    XLO_PAGAMENTO.MoveFirst
    
    XLB_CALCSEGUROS = True
    'CÁLCULO DA CORREÇÃO MONETÁRIA
    DatTitulos.Recordset.bookmark = CboTitulos.SelectedItem
    
    If Not IsNull(DatTitulos.Recordset.Fields!moeda1) Then
        'Correção Monetária
        XLF_VALORCORRIGIDO = Format(funCalculaCorrecaoMonetaria(XLO_PAGAMENTO, 1, DtpBasePagto.Value), "standard")
        
        'Correção
        TxtValorMoeda1.Text = Format(FunNuloVal(XLF_VALORCORRIGIDO) - FunNuloVal(TxtValorReal.Text), "standard")
        LblValorMoeda1.Caption = "Correção " + DatTitulos.Recordset.Fields("moeda1") + ":"
        
        'Juros CEF
        TxtVlCorrigido1.Text = Format(XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XLO_PAGAMENTO, DtpBasePagto.Value, XLF_VALORCORRIGIDO), "standard"), "standard")
    Else
        TxtValorMoeda1.Text = "0,00"
        XLF_VALORCORRIGIDO = Format(TxtValorReal.Text, "standard")
        TxtVlCorrigido1.Text = Format(XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XLO_PAGAMENTO, DtpBasePagto.Value, XLF_VALORCORRIGIDO), "standard"), "standard")
        LblValorMoeda1.Caption = "Correção R$:"
    End If
    
    If Not IsNull(DatTitulos.Recordset.Fields!moeda2) Then
        'Correção Monetária
        XLF_VALORCORRIGIDO = Format(funCalculaCorrecaoMonetaria(XLO_PAGAMENTO, 2, DtpBasePagto.Value), "standard")
        
        'Correção
        TxtValorMoeda2.Text = Format(FunNuloVal(XLF_VALORCORRIGIDO) - FunNuloVal(TxtValorReal.Text), "standard")
        LblValorMoeda2.Caption = "Correção " + DatTitulos.Recordset.Fields("moeda2") + ":"
        
        'Juros CEF
        TxtVlCorrigido2.Text = Format(XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XLO_PAGAMENTO, DtpBasePagto.Value, XLF_VALORCORRIGIDO), "standard"), "standard")
    Else
        TxtValorMoeda2.Text = "0,00"
        XLF_VALORCORRIGIDO = Format(TxtValorReal.Text, "standard")
        TxtVlCorrigido2.Text = Format(XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XLO_PAGAMENTO, DtpBasePagto.Value, XLF_VALORCORRIGIDO), "standard"), "standard")
        LblValorMoeda2.Caption = "Correção R$:"
    End If
    
    If ChkCorrecao.Value = 0 Then
        XLF_VALOR = TxtValorReal.Text
    Else
        If OptValor1.Value = True Then 'usar o valor indexado 1
            XLF_VALOR = FunNuloVal(TxtVlCorrigido1.Text)
        Else
            XLF_VALOR = FunNuloVal(TxtVlCorrigido2.Text)
        End If
    End If
    
    If XLB_CALCULASEGURO = True Then
      'CÁLCULO DO SEGURO
      'Calcular seguro (qdo aplicável) sempre, exceto qdo já foi pago no mês ou se a natureza for financiamento
      XLF_TAXASEGURO = DatImovel.Recordset.Fields("cont_vl_TaxaSegVida")
      XLT_TIPOSEGURO = DatImovel.Recordset.Fields("cont_nr_TipoSeguro")
      
      If XLF_TAXASEGURO <> 0 And DatTitulos.Recordset.Fields!NATUREZA <> "FINANCIAMENTO" Then
          'If DatTitulos.Recordset.Fields!tipl_tx_Descricao = "MENSAL" Then 'Não pode ser mais dessa forma
          XLT_SQL = "SELECT titu_vl_seguro FROM Titulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
                      " AND empd_cd_Empreendimento= '" & DatTitulos.Recordset.Fields("empd_cd_Empreendimento") & "'" & _
                      " AND imov_cd_Imovel='" & DatTitulos.Recordset.Fields("imov_cd_imovel") & "'" & _
                      " AND cont_cd_Contrato='" & DatTitulos.Recordset.Fields("cont_cd_Contrato") & "'" & _
                      " AND month(titu_dt_pagamento)=" & Month(DtpPagamento.Value) & "" & _
                      " AND titu_vl_Seguro<>0 "
          SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
          If XLO_TITULO.EOF Then
              If (DtpPagamento.Value - DtpVencimento.Value) > 14 _
                  And IsNull(DatTitulos.Recordset.Fields!titu_dt_Pagamento) Then
                  If MsgBox("Prazo de Pagto. do seguro superior a 14 dias." + Chr(13) + "Deseja Calcular?", vbCritical + vbYesNo, "ATENÇÃO") = vbNo Then
                      XLB_CALCSEGUROS = False
                  End If
              End If
          Else
              XLB_CALCSEGUROS = False
          End If
          XLO_TITULO.Close
          Set XLO_TITULO = Nothing
  
          If XLB_CALCSEGUROS = True Then
              TxtSeguro.Enabled = True
              LblSeguro.Enabled = True
              If XLT_TIPOSEGURO = "2" Then 'Cálculo do seguro pela parcela
                  TxtSeguro.Text = Format(XLF_VALOR * XLF_TAXASEGURO, "standard")
              ElseIf XLT_TIPOSEGURO = "1" Then 'Cálculo do seguro pela saldo devedor
                  'Selecionar todos os títulos em aberto
                  XLT_SQL = XLT_SQLBASE & _
                    " AND left(Titulo,12)= '" & Left(XLO_PAGAMENTO!Titulo, 12) & "' AND titu_dt_Pagamento is null"
                                 
                  SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
                  While Not XLO_TITULO.EOF
                      If XLO_TITULO!Titulo <> DatTitulos.Recordset.Fields("Titulo") Then 'Excluir o título que está sendo pago
                          'O cálculo da correção monetária para cálculo do seguro deve ser com cotação cheia, ou seja, não usa pro-rata
                          'Deve ser considerado para cálculo do seguro apenas valores indexados,
                          'utilizar valor histórico apenas qdo. não existir valor indexado
                          If OptValor1.Value = True Then 'usar o valor indexado 1
                              If Not IsNull(XLO_TITULO!moed_cd_Moeda1) Then
                                  'XLF_VLCORRIGIDO = funCalculaCorrecaoMonetaria(XLO_TITULO, 1, "01/" + Trim(Str(Month(DtpBasePagto.Value))) + "/" + Trim(Str(Year(DtpBasePagto.Value))))
                                  XLF_VLCORRIGIDO = funCalculaCorrecaoMonetaria(XLO_TITULO, 1, Format(DtpBasePagto.Value, "\01/mm/yy"))
                                  XLF_VLCORRIGIDO = XLF_VLCORRIGIDO + Format(FunCalculaJurosCEF(XLO_PAGAMENTO, Format(DtpBasePagto.Value, "\01/mm/yy"), XLF_VLCORRIGIDO), "standard")
                              Else
                                  XLF_VLCORRIGIDO = XLO_TITULO!titu_vl_Parcela
                                  XLF_VLCORRIGIDO = XLF_VLCORRIGIDO + Format(FunCalculaJurosCEF(XLO_PAGAMENTO, Format(DtpBasePagto.Value, "\01/mm/yy"), XLF_VLCORRIGIDO), "standard")
                              End If
                          Else 'usar o valor indexado 2
                              If Not IsNull(XLO_TITULO!moed_cd_Moeda2) Then
                                  XLF_VLCORRIGIDO = funCalculaCorrecaoMonetaria(XLO_PAGAMENTO, 2, Format(DtpBasePagto.Value, "\01/mm/yy"))
                                  XLF_VLCORRIGIDO = XLF_VLCORRIGIDO + Format(FunCalculaJurosCEF(XLO_PAGAMENTO, Format(DtpBasePagto.Value, "\01/mm/yy"), XLF_VLCORRIGIDO), "standard")
                              Else
                                  XLF_VLCORRIGIDO = XLO_TITULO!titu_vl_Parcela
                                  XLF_VLCORRIGIDO = XLF_VLCORRIGIDO + Format(FunCalculaJurosCEF(XLO_PAGAMENTO, Format(DtpBasePagto.Value, "\01/mm/yy"), XLF_VLCORRIGIDO), "standard")
                              End If
                          End If
                          
                          'If DtpBasePagto.Value < XLO_TITULO!titu_dt_Prorrogacao Then
                          If DtpBasePagto.Value < XLO_TITULO!titu_dt_Vencimento Then
                              XLF_VLCORRIGIDO = XLF_VLCORRIGIDO - (funCalculaDesagio(XLO_TITULO, DtpBasePagto.Value, XLF_VLCORRIGIDO))
                          End If
                          XLF_VLTOTCORRIGIDO = XLF_VLTOTCORRIGIDO + XLF_VLCORRIGIDO
                      End If
                      XLO_TITULO.MoveNext
                  Wend
                  XLO_TITULO.Close
                  Set XLO_TITULO = Nothing
                  TxtSeguro.Text = Format(XLF_VLTOTCORRIGIDO * XLF_TAXASEGURO, "standard")
              End If
          Else
              TxtSeguro.Text = ""
              TxtSeguro.Enabled = False
              LblSeguro.Enabled = False
          End If
      Else
          TxtSeguro.Text = ""
          TxtSeguro.Enabled = False
          LblSeguro.Enabled = False
      End If
    End If
    
    If ChkCorrecao.Value = 0 Then
        XLF_VALOR = CDbl(TxtValorReal.Text)
    Else
        If OptValor1.Value = True Then 'usar o valor indexado 1
            XLF_VALOR = CDbl(FunNuloVal(TxtVlCorrigido1.Text))
        Else
            XLF_VALOR = CDbl(FunNuloVal(TxtVlCorrigido2.Text))
        End If
    End If
    
    XLF_VALOR = XLF_VALOR + CDbl(FunNuloVal(TxtSeguro.Text))
    
    If DtpBasePagto.Value > DtpVencimento.Value Then 'Pago em atraso
       
      'CÁLCULO DO JUROS e MULTA
      XLF_JUROS = FunCalculaJuros(XLO_PAGAMENTO, DtpBasePagto.Value, XLF_VALOR)
      XLF_MULTA = FunCalculaMulta(XLO_PAGAMENTO, DtpBasePagto.Value, XLF_VALOR)
      
      TxtDesagio.Text = ""
      TxtJuros.Text = FunTrataZero(Format(XLF_JUROS, "standard"))
      TxtMulta.Text = FunTrataZero(Format(XLF_MULTA, "standard"))
      TxtJuros.Enabled = True
      TxtMulta.Enabled = True
      
      'Se o usuário não tiver permissão para dar descontos ele não pode tirar multas nem juros
      'Solicitado pela Ebisa em 28/08/2006
      If XLO_PERMISSAODESCONTO.EOF Then
        TxtJuros.Enabled = False
        TxtMulta.Enabled = False
      End If
    
    ElseIf DtpBasePagto.Value < DtpVencimento.Value Then 'PAGO ANTECIPADAMENTE
      
      'CÁLCULO DO DESÁGIO
      TxtDesagio.Text = Format(funCalculaDesagio(XLO_PAGAMENTO, DtpBasePagto.Value, XLF_VALOR - CDbl(FunNuloVal(TxtSeguro.Text))), "standard")
      TxtJuros.Text = ""
      TxtMulta.Text = ""
      TxtJuros.Enabled = False
      TxtMulta.Enabled = False
      
    Else 'PAGO NO VENCIMENTO
      
      TxtJuros.Text = ""
      TxtJuros.Enabled = False
      TxtMulta.Text = ""
      TxtMulta.Enabled = False
      TxtDesagio.Text = ""
    End If
    
    TxtValorPago.Text = Format(XLF_VALOR + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
    XLO_PAGAMENTO.Close
    Set XLO_PAGAMENTO = Nothing
    
End Sub

'Retorna o valor calculado do residuo
Function funCalculaResiduo(XLF_TOTALTITULO As Double, XLF_VLCORRIGIDO As Double) As Double
    Dim XLF_PERCENT As Double    'Percentual encontrado da divisão entre valor pago e valor a pagar
    Dim XLF_MULTA
    Dim XLF_JUROS
    Dim XLF_DESAGIO

    If CDbl(TxtValorPago.Text) < Format(XLF_TOTALTITULO, "Standard") Then
        XLF_PERCENT = CDbl(TxtValorPago.Text) / XLF_TOTALTITULO
        XLF_MULTA = CDbl(FunNuloVal(TxtMulta.Text)) * XLF_PERCENT
        XLF_JUROS = CDbl(FunNuloVal(TxtJuros.Text)) * XLF_PERCENT
        XLF_DESAGIO = CDbl(FunNuloVal(TxtDesagio.Text)) * XLF_PERCENT
        
        funCalculaResiduo = XLF_VLCORRIGIDO - XLF_DESAGIO - CDbl(FunNuloVal(TxtDesconto.Text)) - CDbl(TxtValorPago.Text) + XLF_MULTA + XLF_JUROS + CDbl(FunNuloVal(TxtOutros.Text))
    ElseIf ChkCorrecao.Value = 0 Then
        funCalculaResiduo = XLF_VLCORRIGIDO - CDbl(TxtValorReal.Text)
    End If
End Function
Sub subPreencheCampos()
    Dim XLT_SQL As String
    Dim XLO_TITULO As Object
    Dim XLT_CHAVERESIDUO As String 'Armazena o código de um possível resíduo para verificar se o mesmo existe
    
    XLT_CHAVERESIDUO = Left(Chave, 20) & Format(Str(Int(Right(Chave, 2)) + 1), "00")
    XLT_SQL = "SELECT * FROM ConsCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
              " AND Titulo = '" & XLT_CHAVERESIDUO & "'"
    SubQOpenRecordset XLO_TITULO, XLT_SQL, "Estatico"
    If Not XLO_TITULO.EOF Then
        LblValor.Visible = True
        LblValor2.Visible = True
        LblValor2.Caption = Format(XLO_TITULO!titu_vl_Parcela, "standard")
    End If
    XLO_TITULO.Close
    
    XLT_SQL = "SELECT * FROM ConsCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
              " AND Titulo = '" & Chave & "'"
    SubQOpenRecordset XLO_TITULO, XLT_SQL, "Estatico"
    If Not (XLO_TITULO.BOF And XLO_TITULO.EOF) Then
        MskEmpreendimento.Text = XLO_TITULO!empd_cd_Empreendimento
        CboEmpreendimento.BoundText = XLO_TITULO!empd_cd_Empreendimento
        
        'Se o status do contrato for inativo, só deve permitir visualizar
        If XLO_TITULO!Status = "INATIVO" Then
            subConectarControleDadosNV DatImovel, "SELECT * FROM ConsCapContratos WHERE empr_cd_Empresa= " & PCodEmpresa & " AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' and cont_tx_Status='I' ORDER BY imov_cd_Imovel", Estatico
            CmdCancelar.Enabled = False
        End If
        
        CboImovel.BoundText = XLO_TITULO!imov_cd_Imovel
        CboTitulos.BoundText = Chave
        LblObs2.Caption = FunNulo(XLO_TITULO!obse_tx_Observacao)
        DtpVencimento.Value = Format(XLO_TITULO!titu_dt_Vencimento, "dd/MM/yy")
        DtpProrrogacao.Value = Format(XLO_TITULO!titu_dt_Prorrogacao, "dd/MM/yy")
        
        If DtpVencimento.Value = DtpProrrogacao.Value Then
            DtpProrrogacao.Visible = False
            LblProrrogacao.Visible = False
            LblVencimento.Left = 630
            DtpVencimento.Left = 1710
            LblDtPagamento.Left = 3420
            DtpPagamento.Left = 4440
            LblDtBasePagto.Left = 6120
            DtpBasePagto.Left = 7635
        Else
            DtpProrrogacao.Visible = True
            LblProrrogacao.Visible = True
            LblVencimento.Left = 45
            DtpVencimento.Left = 1125
            LblProrrogacao.Left = 2250
            DtpProrrogacao.Left = 3360
            LblDtPagamento.Left = 4455
            DtpPagamento.Left = 5475
            LblDtBasePagto.Left = 6550
            DtpBasePagto.Left = 8085
        End If
        
        If Not IsNull(XLO_TITULO!titu_dt_Pagamento) Then
            DtpPagamento.Value = Format(XLO_TITULO!titu_dt_Pagamento, "dd/MM/yy")
        End If
        If Not IsNull(XLO_TITULO!titu_dt_BasePagto) Then
            DtpBasePagto.Value = Format(XLO_TITULO!titu_dt_BasePagto, "dd/MM/yy")
        End If
        
        TxtValorReal.Text = Format(XLO_TITULO!titu_vl_Parcela, "standard")
        TxtVlCorrigido1.Text = Format(XLO_TITULO!titu_vl_Corrigido1, "standard")
        TxtValorMoeda1.Text = Format(XLO_TITULO!titu_vl_Corrigido1 - XLO_TITULO!titu_vl_Parcela, "standard")
        TxtVlCorrigido2.Text = Format(XLO_TITULO!titu_vl_Corrigido2, "standard")
        TxtValorMoeda2.Text = Format(XLO_TITULO!titu_vl_Corrigido2 - XLO_TITULO!titu_vl_Parcela, "standard")
               
        DatTitulos.Recordset.bookmark = CboTitulos.SelectedItem
        
        XFB_CALCULA = False 'Trava o calculo dos valores
        
        If XLO_TITULO!titu_tx_CorrigeParcela = "S" Then
            ChkCorrecao.Value = 1
        Else
            ChkCorrecao.Value = 0
        End If
        
        XFB_CALCULA = True 'Libera o calculo dos valores
        
        If XLO_TITULO!titu_tx_IndicePagamento = 1 Then
            OptValor1.Value = True
        ElseIf XLO_TITULO!titu_tx_IndicePagamento = 2 Then
            OptValor2.Value = True
        End If
        
        TxtSeguro.Text = Format(XLO_TITULO!titu_vl_Seguro, "standard")
        TxtJuros.Text = Format(XLO_TITULO!titu_vl_Juros, "standard")
        TxtMulta.Text = Format(XLO_TITULO!titu_vl_Multa, "standard")
        TxtDesagio.Text = Format(XLO_TITULO!titu_vl_Desagio, "standard")
        TxtOutros.Text = Format(XLO_TITULO!titu_vl_Outros, "standard")
        TxtDesconto.Text = Format(XLO_TITULO!titu_vl_Desconto, "standard")
        TxtValorPago.Text = Format(XLO_TITULO!titu_vl_ValorPago, "standard")
        
        If Not IsNull(XLO_TITULO!titu_dt_Deposito) Then
            DtpDeposito.Value = Format(XLO_TITULO!titu_dt_Deposito, "dd/MM/yy")
        End If
        
        If XLO_TITULO!titu_tx_PreDatado = "S" Then
            ChkPreDatado.Value = 1
        Else
            ChkPreDatado.Value = 0
        End If

        If Not IsNull(XLO_TITULO!titu_tx_LocalPagto) And XLO_TITULO!titu_tx_LocalPagto <> "" Then
            CboLocalPagto.ListIndex = XLO_TITULO!titu_tx_LocalPagto
        Else
            CboLocalPagto.ItemData(CboLocalPagto.ListIndex) = 1
        End If
        TxtDocPagto.Text = FunNulo(XLO_TITULO!titu_tx_DocPagto)
        
        If Not DatContaCorrente.Recordset.EOF Then
          DatContaCorrente.Recordset.MoveFirst
          If Not IsNull(XLO_TITULO!coco_cd_codigo) Then
            DatContaCorrente.Recordset.Find "coco_cd_Codigo =" & XLO_TITULO!coco_cd_codigo
            CboCCorrente.BoundText = FunNulo(XLO_TITULO!coco_cd_codigo)
          End If
        End If
        
        If Not FunVerifica_Permissao_SemMsg(Me.Name, "4") Then 'Realizar Pagamentos Conta Diversa Contrato
            CboCCorrente.Enabled = False
        Else
            CboCCorrente.Enabled = True
        End If
        
        If Not IsNull(DatTitulos.Recordset.Fields!moeda1) Then
            LblValorMoeda1.Caption = "Correção " + DatTitulos.Recordset.Fields("moeda1") + ":"
        Else
            LblValorMoeda1.Caption = "Correção R$:"
        End If
        If Not IsNull(DatTitulos.Recordset.Fields!moeda2) Then
            LblValorMoeda2.Caption = "Correção " + DatTitulos.Recordset.Fields("moeda2") + ":"
        Else
            LblValorMoeda2.Caption = "Correção R$:"
        End If
        LblDtBaseTitulo2.Caption = Format(DatTitulos.Recordset.Fields("titu_dt_Base"), "dd/MM/yy")
        TxtTitulo.Text = DatTitulos.Recordset.Fields!titu_cd_Plano + "." + DatTitulos.Recordset.Fields!titu_cd_Parcela + "." + DatTitulos.Recordset.Fields!titu_cd_residuo + _
                        "  -  " + DatTitulos.Recordset.Fields!tipl_tx_Descricao + _
                        "  -  " + DatTitulos.Recordset.Fields!NATUREZA
    End If
End Sub

Private Sub CboCCorrente_Click(Area As Integer)
   If CboCCorrente.Text <> "" Then
     DatContaCorrente.Recordset.bookmark = CboCCorrente.SelectedItem
   End If
End Sub

Private Sub CboEmpreendimento_Change()
    If CboEmpreendimento.Text <> "" Then
        DatEmpreendimento.Recordset.bookmark = CboEmpreendimento.SelectedItem
        MskEmpreendimento.Text = CboEmpreendimento.BoundText
        CboImovel.Enabled = True
        CboImovel.Text = ""
        subConectarControleDadosNV DatImovel, "SELECT * FROM ConsCapContratos WHERE empr_cd_Empresa= " & PCodEmpresa & " AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' and cont_tx_Status='A' ORDER BY imov_cd_Imovel", Estatico
        
        TabFinanceiro.Enabled = False
        FraDatas.Enabled = False
        TxtTitulo.Text = ""
        CboTitulos.Text = ""
    Else
      Call CmdNovo_Click
    End If
End Sub

Private Sub CboEmpreendimento_DblClick(Area As Integer)
    TelaManEmpreendimentos.Show
    If FunLocalizaJanelasAtivas("TelaManEmpreendimentos") <> -1 Then
        TelaManEmpreendimentos.SetFocus
    End If
End Sub

Private Sub CboImovel_Change()
    Dim XLT_SQL As String
    If CboImovel.Text <> "" Then
        DatImovel.Recordset.bookmark = CboImovel.SelectedItem
        If Not IsNull(DatImovel.Recordset.Fields("cont_dt_base")) Then
            LblDtContrato2.Caption = Format(DatImovel.Recordset.Fields("cont_dt_base"), "dd/MM/yy")
        End If
        LblContrato2.Caption = DatImovel.Recordset.Fields!cont_cd_Contrato
        LblDtVenda2.Caption = Format(DatImovel.Recordset.Fields!cont_dt_Venda, "dd/MM/yy")
        LblTipoImovel2.Caption = DatImovel.Recordset.Fields!tiim_tx_Descricao
        
        If XFormulario = "TelaFinTitulos" Then
          'Trazer todos os títulos já pagos
          XLT_SQL = "SELECT (CONVERT(varchar(8), titu_dt_vencimento, 3) " & _
             "+ '  -  ' + titu_cd_Plano + '.' + titu_cd_Parcela + '.' + titu_cd_Residuo " & _
             "+ '  -  ' + left(tipl_tx_descricao,17) + space (15 -len(left(tipl_tx_descricao,17))) + ' - ' + natureza+space(18 - len(natureza))) as Parcela, * " & _
             "FROM ConsCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & " " & _
             "AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' " & _
             "AND not titu_dt_Pagamento is null " & _
             "AND imov_cd_Imovel = '" & CboImovel.BoundText & "' ORDER BY titu_dt_vencimento, Titulo"
        Else
            'Trazer todos os títulos desse contrato que estão a vencer
            XLT_SQL = "SELECT (CONVERT(varchar(8), titu_dt_vencimento, 3) " & _
                "+ '  -  ' + titu_cd_Plano + '.' + titu_cd_Parcela + '.' + titu_cd_Residuo " & _
                "+ '  -  ' + left(tipl_tx_descricao,17) + space (17 -len(left(tipl_tx_descricao,17))) + ' - ' + natureza+space(18 - len(natureza))) as Parcela, * " & _
                "FROM ConsCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & " " & _
                "AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' " & _
                "AND titu_dt_Pagamento is null " & _
                "AND imov_cd_Imovel = '" & CboImovel.BoundText & "' " & _
                "AND status='ATIVO' ORDER BY titu_dt_vencimento, Titulo"
        End If
        subConectarControleDados DatTitulos, XLT_SQL, Estatico
        TabFinanceiro.Enabled = False
        FraDatas.Enabled = False
        CboTitulos.Text = ""
        CboTitulos.Enabled = True
        TxtTitulo.Text = ""
    End If
End Sub

Private Sub CboLocalPagto_Click()
   If CboLocalPagto.Text = "Cheque" Then
     ChkPreDatado.Visible = True
   Else
     ChkPreDatado.Visible = False
   End If
End Sub

Private Sub CboTitulos_Change()

    On Error GoTo RotuloErro

    Dim XLF_VALOR As Double  'Armazena o valor sobre o qual deverá incidir juros
    
    If XFormulario <> "TelaFinTitulos" And CboTitulos.Text <> "" Then
        DatTitulos.Recordset.bookmark = CboTitulos.SelectedItem
        DtpVencimento.Value = Format(DatTitulos.Recordset.Fields("titu_dt_vencimento"), "dd/MM/yy")
        DtpProrrogacao.Value = Format(DatTitulos.Recordset.Fields("titu_dt_Prorrogacao"), "dd/MM/yy")
        If DtpVencimento.Value = DtpProrrogacao.Value Then
            DtpProrrogacao.Visible = False
            LblProrrogacao.Visible = False
            LblVencimento.Left = 630
            DtpVencimento.Left = 1710
            LblDtPagamento.Left = 3420
            DtpPagamento.Left = 4440
            LblDtBasePagto.Left = 6120
            DtpBasePagto.Left = 7635
        Else
            DtpProrrogacao.Visible = True
            LblProrrogacao.Visible = True
            LblVencimento.Left = 45
            DtpVencimento.Left = 1125
            LblProrrogacao.Left = 2250
            DtpProrrogacao.Left = 3360
            LblDtPagamento.Left = 4455
            DtpPagamento.Left = 5475
            LblDtBasePagto.Left = 6550
            DtpBasePagto.Left = 8085
        End If
        LblDtBaseTitulo2.Caption = Format(DatTitulos.Recordset.Fields("titu_dt_Base"), "dd/MM/yy")
        TxtValorReal.Text = Format(DatTitulos.Recordset.Fields("titu_vl_Parcela"), "standard")
        If Not IsNull(DatTitulos.Recordset.Fields("moeda1")) Then
            LblValorMoeda1.Caption = "Correção " + DatTitulos.Recordset.Fields("moeda1") + ":"
        Else
            LblValorMoeda1.Caption = "Correção R$:"
        End If
        If Not IsNull(DatTitulos.Recordset.Fields("moeda2")) Then
            LblValorMoeda2.Caption = "Correção " + DatTitulos.Recordset.Fields("moeda2") + ":"
        Else
            LblValorMoeda2.Caption = "Correção R$:"
        End If
        LblObs2.Caption = FunNulo(DatTitulos.Recordset.Fields("obse_tx_Observacao"))
        
        If ChkCorrecao.Value = 1 Then
            XLF_VALOR = TxtValorReal.Text
        ElseIf OptValor1.Value = True Then
            XLF_VALOR = TxtVlCorrigido1.Text
        Else
            XLF_VALOR = TxtVlCorrigido2.Text
        End If
        
        If Not DatContaCorrente.Recordset.EOF Then
          DatContaCorrente.Recordset.MoveFirst
          If Not IsNull(DatTitulos.Recordset.Fields!coco_cd_codigo) Then
              DatContaCorrente.Recordset.Find "coco_cd_Codigo =" & DatTitulos.Recordset.Fields!coco_cd_codigo
              CboCCorrente.BoundText = FunNulo(DatTitulos.Recordset.Fields("coco_cd_Codigo"))
          Else
            CboCCorrente.BoundText = ""
          End If
        End If
        
        TabFinanceiro.Enabled = True
        FraDatas.Enabled = True
        CmdCancelar.Enabled = True 'Servindo como botão para acessar contrato
        CmdGravar.Enabled = True
        CmdTitulo.Enabled = True
        subCalculaCampos (True)
        TxtTitulo.Text = DatTitulos.Recordset.Fields!titu_cd_Plano + "." + DatTitulos.Recordset.Fields!titu_cd_Parcela + "." + DatTitulos.Recordset.Fields!titu_cd_residuo + _
                        "  -  " + DatTitulos.Recordset.Fields!tipl_tx_Descricao + _
                        "  -  " + DatTitulos.Recordset.Fields!NATUREZA
    End If
    
    Exit Sub
RotuloErro:
    funTrataErros (ComMensagem)
    Resume Next
End Sub

Private Sub ChkCorrecao_Click()
    'Define se o pagamento será com base no valor histórico ou no valor corrigido
    
    If XFB_CALCULA = True Then
      subCalculaCampos (True)
    End If
    
    If ChkCorrecao.Value = 1 Then 'Corrigir parcela
        If OptValor1.Value = True Then 'usar o valor indexado 1
            TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtVlCorrigido1.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
        Else
            TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtVlCorrigido2.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
        End If
    Else
        TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtValorReal.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
    End If
End Sub

Private Sub CmdCancelar_Click()
     
     Dim XLT_SQL As String
     Dim XLO_TITULO As Object
     Dim XLO_QUALIFIN As Object
     Dim XLT_CODEMPREENDIMENTO As String
     Dim XLT_CODIMOVEL As String
     Dim XLT_CODCONTRATO As String
     Dim XLT_CODPLANO As String
     Dim XLT_CODPARCELA As String
     Dim XLT_CODRESIDUO As String
     Dim XLT_TITULOQUALIFIN As String    'Código do título no financeiro
     Dim XLL_FATURA As Long              'Nº da fatura
     Dim XLL_CODDUPL As Long             'Código da Duplicata
     Dim XLD_DTBASE As Date
     
     Dim XLO_QUITACAO As New ADODB.Recordset
        
     
     On Error GoTo RotuloErro
     
     '***************CONTABILIDADE*******************
     'possui o módulo de contabilidade
     If XGB_PQUALICONT Then
        If funMesFechado(DtpPagamento) <> "A" Then
            MsgBox "Não é possível cancelar o Pagamento, pois o mês contábil já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            Exit Sub
        End If
     End If
     '***********************************************
     
     If XFormulario = "TelaFinTitulos" Then
     
        If Not FunVerifica_Permissao("FrmFinPagamentos", "2") Then
          Exit Sub
        End If
        
        DatTitulos.Recordset.bookmark = CboTitulos.SelectedItem
        XLT_CODEMPREENDIMENTO = DatTitulos.Recordset.Fields("empd_cd_Empreendimento")
        XLT_CODIMOVEL = DatTitulos.Recordset.Fields("imov_cd_Imovel")
        XLT_CODCONTRATO = DatTitulos.Recordset.Fields("cont_cd_Contrato")
        XLT_CODPLANO = DatTitulos.Recordset.Fields("titu_cd_Plano")
        XLT_CODPARCELA = DatTitulos.Recordset.Fields("titu_cd_Parcela")
        XLT_CODRESIDUO = DatTitulos.Recordset.Fields("titu_cd_Residuo")
            
        '**********Verificação do resíduo*****************
        'Verificar se existe resíduo para este título e se ele está pago
        'Se estiver, o resíduo deve ser cancelado antes
        'Senão, o seríduo pode ser excluído
        XLT_SQL = "SELECT * FROM Titulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
                        " AND empd_cd_Empreendimento= '" & XLT_CODEMPREENDIMENTO & "'" & _
                        " AND imov_cd_Imovel='" & XLT_CODIMOVEL & "'" & _
                        " AND cont_cd_Contrato='" & XLT_CODCONTRATO & "'" & _
                        " AND titu_cd_Plano='" & XLT_CODPLANO & "'" & _
                        " AND titu_cd_Parcela='" & XLT_CODPARCELA & "'" & _
                        " AND titu_cd_Residuo='" & Format(Str(Int(XLT_CODRESIDUO) + 1), "00") & "'"
        SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
        If Not XLO_TITULO.EOF Then
            If Not IsNull(XLO_TITULO!titu_dt_Pagamento) Then
                MsgBox "Não é possível cancelar esse pagamento, pois existe resíduo pago para esse título.", vbCritical + vbOKOnly, "ATENÇÃO"
                XLO_TITULO.Close
                Exit Sub
            End If
        End If
        XLO_TITULO.Close
        '******************************************************************
         
        Conexao.BeginTrans
        XLT_TITULOQUALIFIN = XLT_CODEMPREENDIMENTO + "." + XLT_CODIMOVEL + "." + XLT_CODCONTRATO + "." + XLT_CODPLANO + "." + XLT_CODPARCELA + "." + XLT_CODRESIDUO
        
        XLT_SQL = "SELECT * FROM Titulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
            " AND empd_cd_Empreendimento= '" & XLT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XLT_CODIMOVEL & "'" & _
            " AND cont_cd_Contrato='" & XLT_CODCONTRATO & "' AND titu_cd_Plano='" & XLT_CODPLANO & "'" & _
            " AND titu_cd_Parcela='" & XLT_CODPARCELA & "' AND titu_cd_Residuo='" & XLT_CODRESIDUO & "'"
        SubQOpenRecordset XLO_TITULO, XLT_SQL, Dinamico
            
        XLO_TITULO!titu_vl_Desagio = 0
        XLO_TITULO!titu_vl_Multa = 0
        XLO_TITULO!titu_vl_Juros = 0
        XLO_TITULO!titu_vl_Desconto = 0
        XLO_TITULO!titu_vl_Outros = 0
        XLO_TITULO!titu_vl_Seguro = 0
        XLO_TITULO!titu_vl_ValorPago = 0
        XLO_TITULO!titu_vl_Corrigido1 = Null
        XLO_TITULO!titu_vl_Corrigido2 = Null
        XLO_TITULO!titu_dt_Pagamento = Null
        XLO_TITULO!titu_dt_BasePagto = Null
        XLO_TITULO!titu_dt_Deposito = Null
        XLO_TITULO!titu_tx_LocalPagto = Null
        XLO_TITULO!titu_tx_DocPagto = Null
        XLO_TITULO!titu_tx_IndicePagamento = Null
        XLO_TITULO!titu_tx_TipoPag = Null
        'XLO_TITULO!coco_cd_Codigo = Null
        XLO_TITULO!moed_cd_Moeda3 = Null
        XLO_TITULO!titu_tx_ExpBanco = "N"
       
        If ChkCorrecao.Value = 1 Then
            XLO_TITULO!titu_tx_CorrigeParcela = "N"
        End If
        If ChkPreDatado.Value = 1 Then
            XLO_TITULO!titu_tx_PreDatado = "N"
        End If
            
        XLD_DTBASE = XLO_TITULO!titu_dt_Base
        
        XLO_TITULO.Update
        
        XLO_TITULO.Close

        'ATUALIZAR NO FINANCEIRO A PARCELA PRINCIPAL
        If XGB_PQUALIFIN Then
            If XGT_EXPCOMPROMCAPI = "S" Then
                SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM NotasFiscais " & _
                    "WHERE nofi_nr_TituloCapi LIKE '" & XLT_TITULOQUALIFIN & "' AND empr_cd_empresa = " & PCodEmpresa & "", Dinamico
                          
                If Not XLO_QUALIFIN.EOF Then
                    XLO_QUALIFIN!nofi_tx_status = "C"
                    XLO_QUALIFIN!nofi_dt_entrada = XLD_DTBASE
                    XLL_FATURA = XLO_QUALIFIN!nofi_nr_fatura
                    XLO_QUALIFIN.Update
                    XLO_QUALIFIN.Close
                
                    SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Duplicatas WHERE dupl_nr_fatura = " & XLL_FATURA, Dinamico
                    XLO_QUALIFIN!dupl_tx_status = "C"
                    XLO_QUALIFIN!dupl_tx_docpagrec = Null
                    XLO_QUALIFIN!dupl_dt_pagrec = Null
                    XLO_QUALIFIN!dupl_dt_concilia = Null
                                
                    'Deletar os descontos acrescimos
                    XLL_CODDUPL = XLO_QUALIFIN!dupl_cd_Duplicata
                    XLO_QUALIFIN.Update
                    XLO_QUALIFIN.Close
                    Conexao.Execute "DELETE FROM DuplDescAcresc WHERE dupl_cd_duplicata= " & XLL_CODDUPL
                End If
            Else 'Se não exporta os compromissos, o registro deve ser excluido do Qualifin
'                SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM NotasFiscais " & _
'                    "WHERE nofi_nr_TituloCapi LIKE '" & XLT_TITULOQUALIFIN & "' AND empr_cd_empresa = " & PCodEmpresa & "", Dinamico
                
                'Cássio Medeiros 16/12/2016
                'Incluído o código da duplicata para permitir a exclusão do lançamento contábil
                SubQOpenRecordset XLO_QUALIFIN, "SELECT nofi_nr_fatura, dupl_cd_Duplicata FROM NotasFiscais " & _
                    "LEFT JOIN Duplicatas ON Duplicatas.dupl_nr_fatura = NotasFiscais.nofi_nr_fatura " & _
                    "WHERE nofi_nr_TituloCapi LIKE '" & XLT_TITULOQUALIFIN & "' AND empr_cd_empresa = " & PCodEmpresa & "", Dinamico
                
                While Not XLO_QUALIFIN.EOF
                    
                    XLL_FATURA = XLO_QUALIFIN!nofi_nr_fatura
                    XLL_CODDUPL = XLO_QUALIFIN!dupl_cd_Duplicata
                    
                    Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_NotaFiscal in  (select nofi_cd_notafiscal from notasfiscais where nofi_nr_fatura = " & XLL_FATURA & ")"
                    Conexao.Execute "delete from dupldescacresc where dupl_cd_duplicata in (select dupl_cd_duplicata from duplicatas where dupl_nr_fatura = " & XLL_FATURA & ")"
                    Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura= " & XLL_FATURA
                    Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_nr_fatura= " & XLL_FATURA
                    
                    XLO_QUALIFIN.MoveNext
                Wend
            End If
        End If
        
        SubQOpenRecordset XLO_TITULO, "SELECT * FROM Titulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
            " AND empd_cd_Empreendimento= '" & XLT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XLT_CODIMOVEL & "'" & _
            " AND cont_cd_Contrato='" & XLT_CODCONTRATO & "' AND titu_cd_Plano='" & XLT_CODPLANO & "'" & _
            " AND titu_cd_Parcela='" & XLT_CODPARCELA & "' AND titu_cd_Residuo='" & Format(Str(Int(XLT_CODRESIDUO) + 1), "00") & "'", Estatico
        
        'só para os registros compromissados
        If IsNull(XLO_TITULO!titu_dt_Pagamento) Then
            Conexao.Execute "DELETE FROM Titulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
                    " AND empd_cd_Empreendimento= '" & XLT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XLT_CODIMOVEL & "'" & _
                    " AND cont_cd_Contrato='" & XLT_CODCONTRATO & "' AND titu_cd_Plano='" & XLT_CODPLANO & "'" & _
                    " AND titu_cd_Parcela='" & XLT_CODPARCELA & "' AND titu_cd_Residuo='" & XLO_TITULO!titu_cd_residuo & "'"
                    
            
            'DELETAR NO FINANCEIRO
            If XGB_PQUALIFIN And XGT_EXPCOMPROMCAPI = "S" Then 'Possui Qualifin e exporta compromissos
                XLT_TITULOQUALIFIN = XLT_CODEMPREENDIMENTO + "." + XLT_CODIMOVEL + "." + XLT_CODCONTRATO + "." + XLT_CODPLANO + "." + XLT_CODPARCELA + "." + XLO_TITULO!titu_cd_residuo
                SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM NotasFiscais " & _
                    "WHERE nofi_nr_TituloCapi LIKE '" & XLT_TITULOQUALIFIN & "' AND empr_cd_empresa = " & PCodEmpresa & "", Estatico
                            
                        While Not XLO_QUALIFIN.EOF
                            Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_Fatura = " & XLO_QUALIFIN!nofi_nr_fatura & ""
                            Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_NotaFiscal = " & XLO_QUALIFIN!nofi_cd_NotaFiscal & ""
                            XLO_QUALIFIN.MoveNext
                        Wend
                        XLO_QUALIFIN.Close
                
                Conexao.Execute "DELETE FROM NotasFiscais WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
                            " AND nofi_nr_TituloCapi LIKE '" & XLT_TITULOQUALIFIN & "'"
            End If
        End If

        '***************CONTABILIDADE*******************
        'Exclui a parte da contabilidade
        Conexao.Execute "DELETE FROM Lancamentos " & _
                        " WHERE lanc_cd_OrigemFin= " & XLL_CODDUPL & "" & _
                        " AND lanc_tx_Origem = 'C'" & _
                        " AND empr_cd_empresa = " & PCodEmpresa & "" & _
                        " AND (plco_cd_debito is null or (not plco_cd_debito is null and not cere_cd_PcrDebitar is null))"
                        
        'Diminuir do valor total da parte do crédito
        SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Lancamentos " & _
                    " WHERE lanc_dt_Data= " & FunNuloData(DtpDeposito, NomeSgbd) & "" & _
                        " AND lanc_tx_Origem = 'C'" & _
                        " AND empr_cd_empresa = " & PCodEmpresa & "" & _
                        " AND plco_cd_credito is null AND plco_cd_debito = '" & DatContaCorrente.Recordset.Fields("plco_cd_conta") & "'", Dinamico

        If Not XLO_QUALIFIN.EOF Then
            If XLO_QUALIFIN!lanc_vl_valor = CDbl(TxtValorPago.Text) Then
                Conexao.Execute "DELETE FROM Lancamentos " & _
                        " WHERE lanc_dt_Data= " & FunNuloData(DtpDeposito, NomeSgbd) & "" & _
                        " AND lanc_tx_Origem = 'C'" & _
                        " AND empr_cd_empresa = " & PCodEmpresa & "" & _
                        " AND plco_cd_credito is null AND plco_cd_debito = '" & DatContaCorrente.Recordset.Fields("plco_cd_conta") & "'"
            Else
                XLO_QUALIFIN!lanc_vl_valor = XLO_QUALIFIN!lanc_vl_valor - CDbl(TxtValorPago.Text)
                XLO_QUALIFIN.Update
            End If
        End If
        XLO_QUALIFIN.Close
        '***********************************************
        XLO_TITULO.Close
                
        SubQOpenRecordset XLO_QUITACAO, "SELECT titu_dt_Vencimento FROM Titulos " & _
        " WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & _
        " empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND " & _
        " imov_cd_Imovel='" & CboImovel.BoundText & "'" & " AND " & _
        " cont_cd_Contrato='" & LblContrato2 & "' AND " & _
        " titu_dt_Pagamento is null", Estatico

        If Not XLO_QUITACAO.EOF Then
          Conexao.Execute "UPDATE Contratos SET cont_dt_Quitacao = NULL" & _
          " WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & _
          " empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND " & _
          " imov_cd_Imovel='" & CboImovel.BoundText & "'" & " AND " & _
          " cont_cd_Contrato='" & LblContrato2 & "'"
        End If
        
        XLO_QUITACAO.Close
        Set XLO_QUITACAO = Nothing
                                
        Conexao.CommitTrans
        
        'Log
        GravaLog CANCELAR_PAGAMENTO_TITULO, "2"
        
        Set XLO_TITULO = Nothing
        
        XOrigem = "CANCELOU"
        
        Unload Me
     Else
        If Not FunVerifica_Permissao("TelaManContratos", "1") Then
          Exit Sub
        End If
        
        XInserir = "C"
        BarraFerramentasPressionada = False
        Chave = (MskEmpreendimento.Text + "." + CboImovel.Text + "." + DatTitulos.Recordset.Fields("cont_cd_Contrato"))
        FrmManContratos.Show 1
     End If
     
     TelaFinTitulos.Adodc1.Refresh
   
   Exit Sub
   
RotuloErro:
    Conexao.RollbackTrans
    
    funTrataErros (ComMensagem)
    Resume Next
End Sub

Private Sub CmdConfirmaDtDeposito_Click()
        
   
    
    'Grava Log
    GravaLog PAGTO_TITULO_DTDEPOSITO, "1"
    
    
    
    CmdGravar.Enabled = True
    TabFinanceiro.Enabled = True
End Sub



Private Sub GravaLog(XLT_OPERACAO As ENUM_OPERACOES, XLT_ORDEM As String)
   
    If (XLT_OPERACAO = PAGAMENTO_TITULO) Then
        ReDim XGM_MATRIZLOG(15, 2) As Variant
                
        'Campos
        XGM_MATRIZLOG(0, 0) = "Título"
        XGM_MATRIZLOG(1, 0) = "Data Pagamento"
        XGM_MATRIZLOG(2, 0) = "Data Base de Pagamento"
        XGM_MATRIZLOG(3, 0) = "Data Depósito"
        XGM_MATRIZLOG(4, 0) = "Valor Corrigido 1"
        XGM_MATRIZLOG(5, 0) = "Valor Corrigido 2"
        XGM_MATRIZLOG(6, 0) = "Seguro"
        XGM_MATRIZLOG(7, 0) = "Deságio"
        XGM_MATRIZLOG(8, 0) = "Juros"
        XGM_MATRIZLOG(9, 0) = "Desconto"
        XGM_MATRIZLOG(10, 0) = "Multa"
        XGM_MATRIZLOG(11, 0) = "Outros"
        XGM_MATRIZLOG(12, 0) = "Valor do Pagamento"
        XGM_MATRIZLOG(13, 0) = "Tipo de Pagamento"
        XGM_MATRIZLOG(14, 0) = "Doc. de Pagamento"
                    
        'Dados informados
        XGM_MATRIZLOG(0, 1) = MskEmpreendimento.Text & "." & CboImovel.BoundText & "." & _
                            LblContrato2 & "." & Mid$(CboTitulos.Text, 14, 2) & "." & _
                            Mid$(CboTitulos.Text, 17, 3) & "." & Mid$(CboTitulos.Text, 21, 2)

        XGM_MATRIZLOG(1, 1) = DtpPagamento.Value
        XGM_MATRIZLOG(2, 1) = DtpBasePagto.Value
        XGM_MATRIZLOG(3, 1) = DtpDeposito.Value
        XGM_MATRIZLOG(4, 1) = FunNuloVal(TxtVlCorrigido1.Text)
        XGM_MATRIZLOG(5, 1) = FunNuloVal(TxtVlCorrigido2.Text)
        XGM_MATRIZLOG(6, 1) = FunNuloVal(TxtSeguro.Text)
        XGM_MATRIZLOG(7, 1) = FunNuloVal(TxtDesagio.Text)
        XGM_MATRIZLOG(8, 1) = FunNuloVal(TxtJuros.Text)
        XGM_MATRIZLOG(9, 1) = FunNuloVal(TxtDesconto.Text)
        XGM_MATRIZLOG(10, 1) = FunNuloVal(TxtMulta.Text)
        XGM_MATRIZLOG(11, 1) = FunNuloVal(TxtOutros.Text)
        XGM_MATRIZLOG(12, 1) = TxtValorPago.Text
        XGM_MATRIZLOG(13, 1) = CboLocalPagto.ListIndex
        XGM_MATRIZLOG(14, 1) = TxtDocPagto.Text
        
    ElseIf (XLT_OPERACAO = CANCELAR_PAGAMENTO_TITULO) Then
        ReDim XGM_MATRIZLOG(1, 2) As Variant
        
        'Campos
        XGM_MATRIZLOG(0, 0) = "Título"
        
        'Dados informados
        XGM_MATRIZLOG(0, 1) = DatTitulos.Recordset.Fields("empd_cd_Empreendimento") & "." & _
            DatTitulos.Recordset.Fields("imov_cd_Imovel") & "." & _
            DatTitulos.Recordset.Fields("cont_cd_Contrato") & "." & _
            DatTitulos.Recordset.Fields("titu_cd_Plano") & "." & _
            DatTitulos.Recordset.Fields("titu_cd_Parcela") & "." & _
            DatTitulos.Recordset.Fields("titu_cd_Residuo")
            
    ElseIf (XLT_OPERACAO = PAGTO_TITULO_DTDEPOSITO) Then
        ReDim XGM_MATRIZLOG(4, 2) As Variant
        
        XGM_MATRIZLOG(0, 0) = "Título"
        XGM_MATRIZLOG(1, 0) = "Data Pagamento"
        XGM_MATRIZLOG(2, 0) = "Data Depósito"
        XGM_MATRIZLOG(3, 0) = "Chave"
                        
        XGM_MATRIZLOG(0, 1) = MskEmpreendimento.Text & "." & CboImovel.BoundText & "." & _
                            LblContrato2 & "." & Mid$(CboTitulos.Text, 14, 2) & "." & _
                            Mid$(CboTitulos.Text, 17, 3) & "." & Mid$(CboTitulos.Text, 21, 2)

        XGM_MATRIZLOG(1, 1) = DtpPagamento.Value
        XGM_MATRIZLOG(2, 1) = DtpDeposito.Value
        XGM_MATRIZLOG(3, 1) = ""
        
    End If
        
    Call subRegistraLog("FrmFinPagamentos", XLT_ORDEM, funCriaDescricaoLog(XGM_MATRIZLOG, XLT_OPERACAO, CONTRATO_CAPI))
End Sub

Private Sub CmdGravar_Click()
     Dim XLT_SQL As String          'Prepara select
     Dim XLO_TITULO As ADODB.Recordset        'Recordset para a tabela de titulos
     Dim XLO_MOEDA As ADODB.Recordset        'Recordset para a tabela de moedas
     Dim XLT_CODPLANO As String     'Código do TITULO
     Dim XLT_CODPARCELA As String   'Código da Parcela
     Dim XLT_CODRESIDUO As String   'Código do resíduo
     Dim XLF_VLRESIDUO As Double    'Valor do Resíduo
     Dim XLT_INDEXACAO1 As String   'tipo de indexação da moeda1
     Dim XLF_VLINDEX1 As Double     'Valor do Indexador 1
     Dim XLT_INDEXACAO2 As String   'tipo de indexação da moeda2
     Dim XLF_VLINDEX2 As Double     'Valor do Indexador 2
     Dim XLF_TOTALTITULO As Double  'Valor total a pagar do título
     Dim XLF_VLCORRIGIDO As Double  'Valor corrigido
     Dim XLT_PRORATA As String     'Indica se a correcao é prorata
     Dim XLT_NATUREZAPLANO As String 'ARMAZENA A NATUREZA PLANO
     Dim XLO_TIPOPLANO As ADODB.Recordset    'Recordset para a tabela de Tipo Planos
     Dim XLI_TIPOPLANO As Integer   'Codigo do tipo do plano
     Dim XLT_VERIFICADATA As String
     Dim XLT_VERIFICAMESATUAL As String
     Dim XLT_VERIFICADATAATUAL As String
     Dim XLO_QUITACAO As New ADODB.Recordset
     Dim XLO_PERMISSAORESIDUO As New ADODB.Recordset
     Dim Titulo As String
     
     
     On Error GoTo RotuloErro
    
     XLT_CODPLANO = Mid$(CboTitulos.Text, 14, 2)
     XLT_CODPARCELA = Mid$(CboTitulos.Text, 17, 3)
     XLT_CODRESIDUO = Mid$(CboTitulos.Text, 21, 2)

     If funChecarCampos Then Exit Sub
     
     'Verifica se o usuário tem permissão para realizar pagamentos fora do mês corrente
     XLT_VERIFICADATAATUAL = Format(Now, "dd/MM/yy")
     XLT_VERIFICAMESATUAL = Mid$(XLT_VERIFICADATAATUAL, 4, 2)
     XLT_VERIFICADATA = Format(DtpPagamento.Value, "dd/MM/yy")
     XLT_VERIFICADATA = Mid$(XLT_VERIFICADATA, 4, 2)
        If XLT_VERIFICADATA <> XLT_VERIFICAMESATUAL Then
          If Not FunVerifica_Permissao_SemMsg("FrmFinPagamentos", "5") Then
            MsgBox "Você não possui permissão para realizar pagamentos fora do mês corrente.", 16, "AVISO"
            DtpPagamento.SetFocus
            Exit Sub
          End If
        End If
    'Verifica a data de deposito e emite uma mensagem caso o usuário não tenha permissão para realizar o depósito na data
        XLT_VERIFICADATA = Format(DtpDeposito.Value, "dd/MM/yy")
        XLT_VERIFICADATA = Mid$(XLT_VERIFICADATA, 4, 2)
        If XLT_VERIFICADATA <> XLT_VERIFICAMESATUAL Then
            If Not FunVerifica_Permissao_SemMsg("FrmFinPagamentos", "5") Then
                MsgBox "Você não possui permissão para realizar pagamentos na data de depósito informada.", 16, "AVISO"
                DtpDeposito.SetFocus
                Exit Sub
            End If
        End If
        
            
     
     
              
      
     '***************CONTABILIDADE*******************
     'possui o módulo de contabilidade
     If XGB_PQUALICONT Then
        If funMesFechado(DtpPagamento) <> "A" Then
            MsgBox "Não é possível efetivar o Pagamento, pois o mês contábil já foi fechado. " _
                    + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
            Exit Sub
        End If
     End If
     '***********************************************
     
     '******************FINANCEIRO*******************
     'possui o módulo do financeiro
     If XGB_PQUALIFIN Then
        If IsNull(DatTitulos.Recordset!plco_cd_Conta) Then
          MsgBox "O título deve possuir uma Conta Contábil.", vbCritical, "ATENÇÃO"
          Exit Sub
        End If
        
        If IsNull(DatTitulos.Recordset!cere_cd_Pcr) Then
          MsgBox "O título deve possuir um Centro de Custo.", vbCritical, "ATENÇÃO"
          Exit Sub
        End If
     End If
     '***********************************************
      
    '************ verifica diferenca de dt de pagamento x dt deposito - inicio
    Dim XLO_RS As ADODB.Recordset
    Dim XLD_HOJE As Date
    SubQOpenRecordset XLO_RS, "SELECT CAST(FLOOR(CAST( getdate() AS float)) AS datetime) dataAtual", Estatico
    XLD_HOJE = CDate(XLO_RS!dataAtual)
    XLO_RS.Close
    
    If DtpPagamento.Value > XLD_HOJE Then
       MsgBox "A data de pagamento do título não pode ser maior que a data atual!", vbCritical, "ATENÇÃO"
       Exit Sub
    End If
    
    
    If DtpDeposito.Value < DtpPagamento.Value Then
       MsgBox "A data de depósito do título não pode ser menor que a data de pagamento!", vbCritical, "ATENÇÃO"
       Exit Sub
    End If
    
    
    
    Dim XLO_FERIADO As ADODB.Recordset
    Dim XLB_INVALIDO As Boolean
    Dim XLB_FERIADO As Boolean
    Dim XLI_DIAS As Integer
    Dim i As Integer
    
    
    SubQOpenRecordset XLO_FERIADO, "SELECT * FROM Feriados WHERE feri_dt_data = " & FunNuloData(DtpDeposito.Value, NomeSgbd) & "", Estatico
    If XLO_FERIADO.EOF And XLO_FERIADO.BOF Then
        XLO_FERIADO.Close
        Set XLO_FERIADO = Nothing
        XLB_FERIADO = False
    Else
        XLO_FERIADO.Close
        Set XLO_FERIADO = Nothing
        XLB_FERIADO = True
    End If
    
    If LCase(Format(DtpDeposito.Value, "ddd")) = "sáb" Or LCase(Format(DtpDeposito.Value, "ddd")) = "dom" Or LCase(Format(DtpDeposito.Value, "ddd")) = "sat" Or LCase(Format(DtpDeposito.Value, "ddd")) = "sun" Or XLB_FERIADO Then
        MsgBox "A data de depósito do título não pode ser um sábado, domingo ou feriado!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
    
    XLI_DIAS = DateDiff("d", DtpPagamento.Value, DtpDeposito.Value, vbSunday)
    
    i = 1
    XLB_INVALIDO = False
    
    While i < XLI_DIAS And XLB_INVALIDO = False
    
        
        SubQOpenRecordset XLO_FERIADO, "SELECT * FROM Feriados WHERE feri_dt_data = " & FunNuloData(DateAdd("d", i, DtpPagamento.Value), NomeSgbd) & "", Estatico
        If XLO_FERIADO.EOF And XLO_FERIADO.BOF Then
            XLO_FERIADO.Close
            Set XLO_FERIADO = Nothing
            XLB_FERIADO = False
        Else
            XLO_FERIADO.Close
            Set XLO_FERIADO = Nothing
            XLB_FERIADO = True
        End If
        
        If (LCase(Format(DateAdd("d", i, DtpPagamento.Value), "ddd")) <> "sáb" And LCase(Format(DateAdd("d", i, DtpPagamento.Value), "ddd")) <> "dom" And LCase(Format(DateAdd("d", i, DtpPagamento.Value), "ddd")) <> "sat" And LCase(Format(DateAdd("d", i, DtpPagamento.Value), "ddd")) <> "sun") _
            And _
            Not XLB_FERIADO Then
            XLB_INVALIDO = True
        Else
            i = i + 1
        End If
        
        
    Wend
    
    If (XLB_INVALIDO) _
        Then
        If MsgBox("A data de depósito difere da data de pagamento do título em mais de 1 dia. Deseja efetuar o pagamento assim mesmo?", vbExclamation + vbYesNo, "ATENÇÃO!") = vbYes Then
            If Not FunVerifica_Permissao(Me.Name, "3") Then 'Realizar Pagamentos  Data Dep Futuro
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    '************ verifica diferenca de dt de pagamento x dt deposito - fim
    
      
     Conexao.BeginTrans
     
     XLT_NATUREZAPLANO = DatTitulos.Recordset.Fields!NATUREZA
     
     XLT_SQL = "SELECT * FROM Titulos " & _
                 "WHERE empr_cd_empresa = " & PCodEmpresa & "" & " AND " & _
                       "empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND " & _
                       "imov_cd_Imovel='" & CboImovel.BoundText & "'" & " AND " & _
                       "cont_cd_Contrato='" & LblContrato2 & "' AND " & _
                       "titu_cd_Plano='" & XLT_CODPLANO & "'" & " AND " & _
                       "titu_cd_Parcela='" & XLT_CODPARCELA & "' AND " & _
                       "titu_cd_Residuo='" & XLT_CODRESIDUO & "'"
     SubQOpenRecordset XLO_TITULO, XLT_SQL, Dinamico
         
     XLO_TITULO!titu_vl_Desagio = FunNuloVal(TxtDesagio.Text)
     XLO_TITULO!titu_vl_Multa = FunNuloVal(TxtMulta.Text)
     XLO_TITULO!titu_vl_Juros = FunNuloVal(TxtJuros.Text)
     XLO_TITULO!titu_vl_Desconto = FunNuloVal(TxtDesconto.Text)
     XLO_TITULO!titu_vl_Outros = FunNuloVal(TxtOutros.Text)
     XLO_TITULO!titu_vl_Seguro = FunNuloVal(TxtSeguro.Text)
     XLO_TITULO!titu_vl_ValorPago = TxtValorPago.Text
     XLO_TITULO!titu_vl_Corrigido1 = FunNuloVal(TxtVlCorrigido1.Text)
     XLO_TITULO!titu_vl_Corrigido2 = FunNuloVal(TxtVlCorrigido2.Text)
     XLO_TITULO!titu_dt_Pagamento = DtpPagamento.Value
     XLO_TITULO!titu_dt_BasePagto = DtpBasePagto.Value
     XLO_TITULO!titu_dt_Deposito = DtpDeposito.Value
     XLO_TITULO!titu_tx_LocalPagto = CboLocalPagto.ListIndex
     XLO_TITULO!titu_tx_DocPagto = TxtDocPagto.Text
     XLO_TITULO!coco_cd_codigo = funNull(CboCCorrente.BoundText)
     XLO_TITULO!titu_tx_ExpBanco = "N"
     
     If ChkPreDatado.Value = 1 Then
         XLO_TITULO!titu_tx_PreDatado = "S"
     End If
     
     If OptValor1.Value = True Then 'usar o valor indexado 1
        XLO_TITULO!titu_tx_IndicePagamento = "1"
        XLO_TITULO!moed_cd_Moeda3 = funNull(DatTitulos.Recordset.Fields("moed_cd_Moeda1"))
        XLF_TOTALTITULO = (CDbl(FunNuloVal(TxtVlCorrigido1.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))))
        XLF_VLCORRIGIDO = CDbl(FunNuloVal(TxtVlCorrigido1.Text)) + CDbl(FunNuloVal(TxtSeguro.Text))
     Else 'usar o valor indexado 2
        XLO_TITULO!titu_tx_IndicePagamento = "2"
        XLO_TITULO!moed_cd_Moeda3 = funNull(DatTitulos.Recordset.Fields("moed_cd_Moeda2"))
        XLF_TOTALTITULO = (CDbl(FunNuloVal(TxtVlCorrigido2.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))))
        XLF_VLCORRIGIDO = CDbl(FunNuloVal(TxtVlCorrigido2.Text)) + CDbl(FunNuloVal(TxtSeguro.Text))
     End If
     
     If ChkCorrecao.Value = 1 Then
         XLO_TITULO!titu_tx_CorrigeParcela = "S"
     Else
         XLF_TOTALTITULO = (CDbl(FunNuloVal(TxtValorReal.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))))
     End If
        
     XLO_TITULO!titu_tx_TipoPag = funGeraTipoPag(Trim(Str(ChkCorrecao.Value)), CDbl(TxtValorPago.Text), XLF_TOTALTITULO, XLF_VLCORRIGIDO, DtpVencimento.Value, DtpPagamento.Value)
     
     '////////VERIFICAR SE HÁ RESÍDUO////////////
     XLF_VLRESIDUO = Format(funCalculaResiduo(XLF_TOTALTITULO, XLF_VLCORRIGIDO), "standard")
     If XLF_VLRESIDUO <> 0 Then
        'Atualizar valores de multa, juros e deságio da parte paga
        XLO_TITULO!titu_vl_Desagio = Format(((CDbl(FunNuloVal(TxtValorPago.Text)) * CDbl(FunNuloVal(TxtDesagio.Text))) / XLF_TOTALTITULO), "standard")
        XLO_TITULO!titu_vl_Multa = Format(((CDbl(FunNuloVal(TxtValorPago.Text)) * CDbl(FunNuloVal(TxtMulta.Text))) / XLF_TOTALTITULO), "standard")
        XLO_TITULO!titu_vl_Juros = Format(((CDbl(FunNuloVal(TxtValorPago.Text)) * CDbl(FunNuloVal(TxtJuros.Text))) / XLF_TOTALTITULO), "standard")
        
        FrmFINVencResiduo.DtpVencResiduo.Value = DtpVencimento.Value
        FrmFINVencResiduo.LblValor2.Caption = Format(XLF_VLRESIDUO, "standard")
               
        If ChkCorrecao.Value = 1 Then
           SubQOpenRecordset XLO_PERMISSAORESIDUO, "SELECT * FROM Habilitacoes WHERE " & _
             " func_cd_operador = " & func_cd_operador & " AND " & _
             " roti_tx_programa  = 'FrmFINVencResiduo' AND  roti_tx_ordem = '1' AND habi_bl_Habilitado='S' ", Dinamico
             
           If XLO_PERMISSAORESIDUO.EOF Then
             XGB_GERARESIDUO = True
           Else
             FrmFINVencResiduo.Show vbModal
           End If
           
           XLO_PERMISSAORESIDUO.Close
        Else
            XGB_GERARESIDUO = True
        End If
        
        Set XLO_PERMISSAORESIDUO = Nothing
        
        If XGB_GERARESIDUO Then
        
            XLT_PRORATA = DatImovel.Recordset.Fields("cont_tx_correcProRata")
            
            'Verificar o tipo de indexação
            If Not IsNull(DatTitulos.Recordset.Fields("moed_cd_Moeda1")) Then
                SubQOpenRecordset XLO_MOEDA, "SELECT * FROM Moedas WHERE Moedas.moed_cd_CodMoeda = " & DatTitulos.Recordset.Fields("moed_cd_Moeda1") & "", Estatico
                XLT_INDEXACAO1 = XLO_MOEDA!moed_tx_Indexacao
                XLO_MOEDA.Close
                
                XLF_VLINDEX1 = FunCalcularValorIndexado(Format(XLF_VLRESIDUO, "standard"), DtpBasePagto.Value, DatTitulos.Recordset.Fields("moed_cd_Moeda1"), XLT_INDEXACAO1, XLT_PRORATA, TelaCotacoes)
            Else
                XLF_VLINDEX1 = 0
            End If
            
            If Not IsNull(DatTitulos.Recordset.Fields("moed_cd_Moeda2")) Then
                SubQOpenRecordset XLO_MOEDA, "SELECT * FROM Moedas WHERE Moedas.moed_cd_CodMoeda = " & DatTitulos.Recordset.Fields("moed_cd_Moeda2") & "", Estatico
                XLT_INDEXACAO2 = XLO_MOEDA!moed_tx_Indexacao
                XLO_MOEDA.Close
            
                XLF_VLINDEX2 = FunCalcularValorIndexado(Format(XLF_VLRESIDUO, "standard"), DtpBasePagto.Value, DatTitulos.Recordset.Fields("moed_cd_Moeda2"), XLT_INDEXACAO2, XLT_PRORATA, TelaCotacoes)
            Else
                XLF_VLINDEX2 = 0
            End If
            
            'GRAVAR RESÍDUO
            XLT_CODRESIDUO = Format(Int(XLO_TITULO!titu_cd_residuo) + 1, "00")
            
            XLT_SQL = " SELECT tipl_cd_TipoPlano FROM TiposPlanos WHERE tipl_tx_sigla = 'R'"
            SubQOpenRecordset XLO_TIPOPLANO, XLT_SQL, Estatico
            XLI_TIPOPLANO = XLO_TIPOPLANO!tipl_cd_TipoPlano
            XLO_TIPOPLANO.Close
            
            XLT_SQL = "INSERT INTO Titulos (empr_cd_Empresa, empd_cd_Empreendimento, imov_cd_Imovel,cont_cd_Contrato, " & _
                            "titu_cd_Plano,titu_cd_Parcela,titu_cd_Residuo,obse_cd_Observacao, " & _
                            "moed_cd_Moeda1, moed_cd_Moeda2, titu_vl_Parcela, " & _
                            "titu_vl_ParcelaIndex1, titu_vl_ParcelaIndex2, titu_dt_Vencimento, tipl_cd_TipoPlano, " & _
                            "titu_vl_PercJurosFin, titu_vl_PercMultaMora, titu_vl_PercJurosMora, " & _
                            "cere_cd_Pcr,plco_cd_Conta,coco_cd_Codigo,titu_dt_Base," & _
                            "titu_dt_Prorrogacao,titu_tx_NaturezaPlano,titu_dt_Criacao,iden_cd_Projeto)" & _
                    " SELECT " & XLO_TITULO!empr_cd_Empresa & ",'" & XLO_TITULO!empd_cd_Empreendimento & "', " & _
                            "'" & XLO_TITULO!imov_cd_Imovel & "','" & XLO_TITULO!cont_cd_Contrato & "'," & _
                            "'" & XLO_TITULO!titu_cd_Plano & "','" & XLO_TITULO!titu_cd_Parcela & "'," & _
                            "'" & XLT_CODRESIDUO & "'," & FunNuloBancoVal(XLO_TITULO!obse_cd_Observacao) & ", " & _
                            "" & FunNuloBancoVal(XLO_TITULO!moed_cd_Moeda1) & "," & FunNuloBancoVal(XLO_TITULO!moed_cd_Moeda2) & "," & _
                            "" & FunNuloVal(FunTrataFloat(XLF_VLRESIDUO)) & "," & FunNuloVal(FunTrataFloat(XLF_VLINDEX1)) & "," & _
                            "" & FunNuloVal(FunTrataFloat(XLF_VLINDEX2)) & "," & FunNuloData(FrmFINVencResiduo.DtpVencResiduo.Value, NomeSgbd) & "," & _
                            "" & XLI_TIPOPLANO & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_PercJurosFin)) & "," & _
                            "" & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_PercMultaMora)) & "," & FunNuloVal(FunTrataFloat(XLO_TITULO!titu_vl_PercJurosMora)) & ", " & _
                            "" & FunNuloBanco(XLO_TITULO!cere_cd_Pcr) & "," & FunNuloBanco(XLO_TITULO!plco_cd_Conta) & "," & FunNuloBancoVal(CboCCorrente.BoundText) & "," & _
                            "" & FunNuloData(DtpBasePagto.Value, NomeSgbd) & "," & FunNuloData(FrmFINVencResiduo.DtpVencResiduo.Value, NomeSgbd) & "," & _
                            "'" & XLO_TITULO!titu_tx_NaturezaPlano & "'," & FunNuloData(DtpPagamento.Value, NomeSgbd) & ", " & FunNuloBancoVal(XLO_TITULO!iden_cd_Projeto) & ""
            
            Conexao.Execute XLT_SQL
            
            Unload FrmFINVencResiduo
        End If
     End If
     '//////////////////////////////////////////
     
     XLO_TITULO.Update
     
     If XGB_PQUALIFIN = True Then
       subGravaFinanceiro XLO_TITULO, "PAGAR", DatImovel.Recordset.Fields("focl_cd_FornCli"), DatTitulos.Recordset.Fields!tipl_tx_Descricao, XLT_NATUREZAPLANO
       
       'Se o desconto/acréscimo associado ao título não possuir uma conta contábil associada, o pagamento é cancelado.
       If XGB_ContinaPagamento = False Then
            Conexao.RollbackTrans
            CmdNovo.Enabled = True
            CmdRecibo.Enabled = True
            CboEmpreendimento.Enabled = False
            MskEmpreendimento.Enabled = False
            CboImovel.Enabled = False
            CboTitulos.Enabled = False
            DtpBasePagto.Enabled = False
            DtpPagamento.Enabled = False
            CmdGravar.Enabled = False
            TxtValorPago.Enabled = False
            TxtDesagio.Enabled = False
            TxtDesconto.Enabled = False
            TxtJuros.Enabled = False
            TxtMulta.Enabled = False
            TxtOutros.Enabled = False
            TxtSeguro.Enabled = False
            Unload Me
            Exit Sub
       End If
       
     End If
     XLO_TITULO.Close
     Set XLO_TITULO = Nothing
     
     If XLF_VLRESIDUO <> 0 And XGB_GERARESIDUO Then 'Gravar um compromisso no financeiro referente ao resíduo
         XLT_SQL = "SELECT * FROM Titulos " & _
                    "WHERE empr_cd_empresa = " & PCodEmpresa & "" & " AND " & _
                           "empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND " & _
                           "imov_cd_Imovel='" & CboImovel.BoundText & "'" & " AND " & _
                           "cont_cd_Contrato='" & LblContrato2 & "' AND " & _
                           "titu_cd_Plano='" & XLT_CODPLANO & "'" & " AND " & _
                           "titu_cd_Parcela='" & XLT_CODPARCELA & "' AND " & _
                           "titu_cd_Residuo='" & XLT_CODRESIDUO & "'"
         SubQOpenRecordset XLO_TITULO, XLT_SQL, Dinamico
         
         If XGB_PQUALIFIN = True And XGT_EXPCOMPROMCAPI = "S" Then
           subGravaFinanceiro XLO_TITULO, "COMPROMISSAR", DatImovel.Recordset.Fields("focl_cd_FornCli"), "RESÍDUO", XLT_NATUREZAPLANO
         End If
         
         XLO_TITULO.Close
         
         Set XLO_TITULO = Nothing
         
     End If
        
     SubQOpenRecordset XLO_QUITACAO, "SELECT titu_dt_Vencimento FROM Titulos " & _
        " WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & _
        " empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND " & _
        " imov_cd_Imovel='" & CboImovel.BoundText & "'" & " AND " & _
        " cont_cd_Contrato='" & LblContrato2 & "' AND " & _
        " titu_dt_Pagamento is null", Estatico

      If XLO_QUITACAO.EOF Then
        Conexao.Execute "UPDATE Contratos SET cont_dt_Quitacao=" & FunNuloData(DtpPagamento.Value, NomeSgbd) & _
        " WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & _
        " empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND " & _
        " imov_cd_Imovel='" & CboImovel.BoundText & "'" & " AND " & _
        " cont_cd_Contrato='" & LblContrato2 & "'"
        MsgBox "O Contrato foi Quitado com sucesso!", vbInformation, "ATENÇÃO"
      End If
      XLO_QUITACAO.Close ' Alteração em 10/10/08 - Patrícia
      Set XLO_QUITACAO = Nothing
        
     Conexao.CommitTrans
     
     'Log
     GravaLog PAGAMENTO_TITULO, "1"
     
          
     CmdNovo.Enabled = True
     CmdRecibo.Enabled = True
     ' */*/*/*/*/*/*/*/*/*
     CboEmpreendimento.Enabled = False
     MskEmpreendimento.Enabled = False
     CboImovel.Enabled = False
     CboTitulos.Enabled = False
     DtpBasePagto.Enabled = False
     DtpPagamento.Enabled = False
     CmdGravar.Enabled = False
     TxtValorPago.Enabled = False
     TxtDesagio.Enabled = False
     TxtDesconto.Enabled = False
     TxtJuros.Enabled = False
     TxtMulta.Enabled = False
     TxtOutros.Enabled = False
     TxtSeguro.Enabled = False
     
     XOrigem = "PAGOU"
     Titulo = TelaFinTitulos.Adodc1.Recordset.Fields("Titulo")
     
    TelaFinTitulos.Adodc1.Refresh

    TelaFinTitulos.cboLocalizar.BoundText = Titulo
    If Not IsNull(TelaFinTitulos.cboLocalizar.SelectedItem) Then TelaFinTitulos.Adodc1.Recordset.bookmark = TelaFinTitulos.cboLocalizar.SelectedItem
    'TelaFinTitulos.AtualizaGrid

     Exit Sub
     
RotuloErro:
    
    Conexao.RollbackTrans
    funTrataErros (ComMensagem)
    

End Sub

Private Sub CmdNovo_Click()
    
    MskEmpreendimento.Text = ""
    MskEmpreendimento.Enabled = True
    CboEmpreendimento.Text = ""
    CboEmpreendimento.Enabled = True
    CboImovel.Text = ""
    CboImovel.Enabled = False
    LblContrato2.Caption = ""
    LblDtContrato2.Caption = ""
    LblTipoImovel2.Caption = ""
    LblDtVenda2.Caption = ""
    LblContrato2.Caption = ""
    LblDtBaseTitulo2.Caption = ""
    TxtTitulo.Text = ""
    CboTitulos.Text = ""
    CboTitulos.Enabled = False
    LblObs2.Caption = ""
    DtpVencimento.Value = Format(Now, "dd/MM/yy") ' não aceita Null qdo não possui check
    DtpProrrogacao.Value = Format(Now, "dd/MM/yy") ' não aceita Null qdo não possui check
    DtpPagamento.Value = Format(Now, "dd/MM/yy")
    DtpPagamento.Enabled = True
    DtpBasePagto.Value = Format(Now, "dd/MM/yy")
    DtpBasePagto.Enabled = True
    ChkCorrecao.Value = 1
    TxtValorReal.Text = ""
    TxtValorMoeda1.Text = ""
    TxtVlCorrigido1.Text = ""
    OptValor1.Value = True
    TxtValorMoeda2.Text = ""
    TxtVlCorrigido2.Text = ""
    TxtSeguro.Text = ""
    TxtJuros.Text = ""
    TxtMulta.Text = ""
    TxtDesagio.Text = ""
    TxtOutros.Text = ""
    TxtDesconto.Text = ""
    TxtValorPago.Text = ""
    DtpDeposito.Value = Format(Now, "dd/MM/yy")
    ChkPreDatado.Value = 0
    CboLocalPagto.ListIndex = 0
    TxtDocPagto.Text = ""
    CboCCorrente.Text = ""
    CboCCorrente.Enabled = True
    CmdNovo.Enabled = False
    CmdGravar.Enabled = True
    TxtValorPago.Enabled = True
    TxtOutros.Enabled = True
    
End Sub

Private Sub CmdRecibo_Click()
  
  Dim XLT_CONSULTA As String
  Dim XLT_SQL As String
  Dim XLT_CONJUNCAO As String
  Dim XLT_FORMAPGTO As String
  Dim XLT_DOCPAGTO As String
  Dim XLT_DEPOSITO As String
  
  Me.MousePointer = vbHourglass
  
  XLT_SQL = " SELECT consCAPRecibo.*," & _
   "(SELECT COUNT(*) FROM Titulos AS tit WHERE tit.empr_cd_empresa=consCAPRecibo.empr_cd_empresa  " & _
   " AND tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento " & _
   " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND  tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato " & _
   " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_residuo = '00') AS TotalParcelas, " & _
   " IsNull((SELECT titu_vl_Parcela FROM ConscapTitulos as tit WHERE tit.empr_cd_empresa=consCAPRecibo.empr_cd_empresa  " & _
   " AND tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento" & _
   " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato" & _
   " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_Parcela = consCAPRecibo.titu_cd_Parcela" & _
   " AND (tit.titu_cd_Residuo = (consCAPRecibo.titu_cd_residuo + 1))), 0) AS ValorResiduo" & _
   " FROM consCAPRecibo WHERE " & _
   " Titulo = '" & DatTitulos.Recordset!Titulo & "' AND" & _
   " empr_cd_empresa = " & PCodEmpresa
  
  XLT_DOCPAGTO = TxtDocPagto.Text
  
  If CboCCorrente.BoundText <> "" And CboLocalPagto.Text = "Cheque" Then
    XLT_DEPOSITO = DatContaCorrente.Recordset!banc_tx_Descricao & " - AG: " & DatContaCorrente.Recordset!coco_cd_Agencia & ""
    If FunNulo(DatContaCorrente.Recordset!coco_nr_Dag) <> "" Then
      XLT_DEPOSITO = XLT_DEPOSITO & "-" & DatContaCorrente.Recordset!coco_nr_Dag
    End If
    XLT_DEPOSITO = XLT_DEPOSITO & " - C/C: " & DatContaCorrente.Recordset!coco_tx_Conta
    If FunNulo(DatContaCorrente.Recordset!coco_nr_Dac) <> "" Then
      XLT_DEPOSITO = XLT_DEPOSITO & "-" & DatContaCorrente.Recordset!coco_nr_Dac
    End If
  Else
    XLT_DEPOSITO = ""
  End If
  
  'Abre os recordsets e verifica se eles possuem dados
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL)
  If XGO_RSRELATORIO.EOF Then
    MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Erro na consulta"
    Me.MousePointer = vbDefault
    Exit Sub
  End If
  Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)

  'Carrega os relatórios
  FrmVisRelatorios.Caption = "Impressão de Recibo"
  
  If XGT_MODELORECIBO = "1" Then
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelRecibo.rpt")
  Else
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelRecibo2.rpt")
  End If

  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
 
'  Unload Me

  'Ativa os Relatórios
  With XGR_RELATORIO.FormulaFields
    .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    .GetItemByName("DocPgto").Text = Chr(34) & Replace(XLT_DOCPAGTO, vbCrLf, " - ") & Chr(34)
    .GetItemByName("DepositoBanco").Text = Chr(34) & XLT_DEPOSITO & Chr(34)
  End With
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Rodape"
  FrmVisRelatorios.SubVerRelatorioMoldal XGR_RELATORIO, Me
  Me.MousePointer = vbDefault

End Sub

Private Sub CmdSairDtDeposito_Click()
    DtpDeposito.Value = DtpPagamento.Value
    CmdGravar.Enabled = True
    TabFinanceiro.Enabled = True
End Sub

Private Sub CmdTitulo_Click()
    Chave = DatTitulos.Recordset.Fields("Titulo")
    XInserir = "C"
    FrmFinTitulos.Show 1
End Sub

Private Sub DtpBasePagto_LostFocus()
    subCalculaCampos (True)
End Sub


Private Sub DtpPagamento_LostFocus()
    DtpBasePagto.Value = DtpPagamento.Value
    DtpDeposito.Value = DtpPagamento.Value
    DtpBasePagto.SetFocus
End Sub

Private Sub Form_Activate()
    Set Formulario = FrmFinPagamentos
    FrmFinPagamentos.Top = 1005
    FrmFinPagamentos.Left = 0
End Sub

Private Sub Form_Load()
   
    Set Formulario = FrmFinPagamentos

    subConectarControleDados DatEmpreendimento, "SELECT empd_cd_Empreendimento,empd_tx_Nome FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_nome", Estatico
    subConectarControleDados DatContaCorrente, "SELECT * FROM ConsGENCCCombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY descCombo", Estatico
    DtpPagamento.Value = Format(Now, "dd/MM/yy")
    DtpBasePagto.Value = Format(Now, "dd/MM/yy")
    DtpDeposito.Value = Format(Now, "dd/MM/yy")
    LblTipoImovel2.Caption = ""
    LblContrato2.Caption = ""
    LblDtVenda2.Caption = ""
    LblDtContrato2.Caption = ""
    LblDtBaseTitulo2.Caption = ""
    LblObs2.Caption = ""
    CboLocalPagto.ListIndex = 0
    
    XFB_CALCULA = True 'Sempre calcula
    
    SubQOpenRecordset XLO_PERMISSAODESCONTO, "SELECT * FROM Habilitacoes WHERE " & _
        " func_cd_operador = " & func_cd_operador & " AND " & _
        " roti_tx_programa  = 'FrmFINDesconto' AND  roti_tx_ordem = '1' AND habi_bl_Habilitado='S' ", Dinamico
        
    If XLO_PERMISSAODESCONTO.EOF Then
       TxtDesconto.Enabled = False
       TxtJuros.Enabled = False
       TxtMulta.Enabled = False
    Else
       TxtDesconto.Enabled = True
       TxtJuros.Enabled = True
       TxtMulta.Enabled = True
    End If
        
    If Not XGB_PQUALIFIN Then 'Integrado com o Financeiro
        CboCCorrente.Enabled = False
        LblCCorrente.Enabled = False
    End If
    
    If XFormulario = "TelaFinTitulos" Then
        CmdNovo.Visible = False
        CmdGravar.Visible = False
        CmdTitulo.Visible = False
        CmdRecibo.Left = 1732
        CmdRecibo.Enabled = True
        CmdCancelar.Caption = "&Cancelar"
        CmdCancelar.Enabled = True
        CmdCancelar.Left = 4004
        CmdDesistir.Left = 6277
        subPreencheCampos
        
            
        'Desabilitar campos
        CboTitulos.Enabled = False
        MskEmpreendimento.Enabled = False
        CboEmpreendimento.Enabled = False
        CboImovel.Enabled = False
        FraValores1.Enabled = False
        FraValores2.Enabled = False
        PanPagto.Enabled = False
        FraDadosBancarios.Enabled = False
        CboCCorrente.Enabled = False
        TabFinanceiro.Enabled = True
    Else
        CmdNovo.Left = 180 '592
        CmdGravar.Left = 1719 '2298
        CmdGravar.Enabled = False
        CmdRecibo.Left = 3258 '4004
        CmdRecibo.Enabled = False
        CmdTitulo.Enabled = False
        CmdCancelar.Left = 4797 '5710
        CmdCancelar.Caption = "&Contrato"
        CmdTitulo.Left = 6336
        CmdDesistir.Left = 7875 '7417
        If XFormulario = "TelaFinTitulosP" Then
            subPreencheCampos
            ChkCorrecao.Value = 1
        End If
    End If
    
    'Flag utilizado para cancelar o pagamento de um título se existir desconto para o título e este desconto
    'não possuir conta contábil cadastrada. - Larissa 14/08/08
    XGB_ContinaPagamento = True
    
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "FrmFinPagamentos"
End Sub

Private Sub MskEmpreendimento_GotFocus()
    Call subSelecionaMSK(MskEmpreendimento)
End Sub

Private Sub MskEmpreendimento_LostFocus()
    If MskEmpreendimento.Text <> "" Then
        SubPreencheComZero MskEmpreendimento
        CboEmpreendimento.BoundText = MskEmpreendimento.Text
        If CboEmpreendimento.Text = "" Then
            MsgBox "Não existe Empreendimento com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskEmpreendimento.SetFocus
            Exit Sub
        End If
        'subConectarControleDadosNV DatImovel, "SELECT * FROM ConsCapImoveis WHERE empd_cd_empreendimento='" & MskEmpreendimento.Text & "' and imov_tx_status='A' ORDER BY imov_cd_imovel", Estatico
    Else
      CboEmpreendimento.BoundText = ""
    End If
End Sub

Private Sub MskEmpreendimento_KeyPress(KeyAscii As Integer)
If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
    End If
End Sub
Private Sub CmdDesistir_Click()
    Unload Me
End Sub

Private Sub OptValor1_Click()
    
    If XFormulario <> "TelaFinTitulos" And CboTitulos.Text <> "" Then
      subCalculaCampos (True)
      'Define qual o valor corrigido que será utilizado
      If OptValor1.Value = True And ChkCorrecao.Value = 1 Then 'usar o valor indexado 1
          TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtVlCorrigido1.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
      End If
    End If
      
End Sub

Private Sub OptValor2_Click()
    
    If XFormulario <> "TelaFinTitulos" And CboTitulos.Text <> "" Then
      subCalculaCampos (True)
      'Define qual o valor corrigido que será utilizado
      If OptValor2.Value = True And ChkCorrecao.Value = 1 Then 'usar o valor indexado 2
          TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtVlCorrigido2.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
      End If
    End If
    
End Sub


Private Sub TxtDesagio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDesconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDesconto_LostFocus()
    
    If Not IsNumeric(TxtDesconto.Text) And TxtDesconto.Text <> "" Then
      TxtDesconto.Text = "0,00"
    Else
      TxtDesconto.Text = Format(TxtDesconto.Text, "standard")
    End If
    
    If ChkCorrecao.Value = 1 Then 'Corrigir parcela
        If OptValor1.Value = True Then 'usar o valor indexado 1
            TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtVlCorrigido1.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
        Else
            TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtVlCorrigido2.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
        End If
    Else
        TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtValorReal.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
    End If
    TxtDesconto = Format(TxtDesconto, "standard")
End Sub


Private Sub TxtJuros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtJuros_LostFocus()
    
    If Not IsNumeric(TxtJuros.Text) And TxtJuros.Text <> "" Then
      TxtJuros.Text = "0,00"
    Else
      TxtJuros.Text = Format(TxtJuros.Text, "standard")
    End If
    
    If ChkCorrecao.Value = 1 Then 'Corrigir parcela
        If OptValor1.Value = True Then 'usar o valor indexado 1
            TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtVlCorrigido1.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
        Else
            TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtVlCorrigido2.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
        End If
    Else
        TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtValorReal.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
    End If
End Sub

Private Sub TxtMulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtMulta_LostFocus()
    
    If Not IsNumeric(TxtMulta.Text) And TxtMulta.Text <> "" Then
      TxtMulta.Text = "0,00"
    Else
      TxtMulta.Text = Format(TxtMulta.Text, "standard")
    End If
    
    If ChkCorrecao.Value = 1 Then 'Corrigir parcela
        If OptValor1.Value = True Then 'usar o valor indexado 1
            TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtVlCorrigido1.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
        Else
            TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtVlCorrigido2.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
        End If
    Else
        TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtValorReal.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
    End If
End Sub

Private Sub TxtOutros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtOutros_LostFocus()
    
    If Not IsNumeric(TxtOutros.Text) And TxtOutros.Text <> "" Then
      TxtOutros.Text = "0,00"
    Else
      TxtOutros.Text = Format(TxtOutros.Text, "standard")
    End If
    
    If ChkCorrecao.Value = 1 Then 'Corrigir parcela
        If OptValor1.Value = True Then 'usar o valor indexado 1
            TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtVlCorrigido1.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
        Else
            TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtVlCorrigido2.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
        End If
    Else
        TxtValorPago.Text = Format(CDbl(FunNuloVal(TxtValorReal.Text)) + CDbl(FunNuloVal(TxtSeguro.Text)) + CDbl(FunNuloVal(TxtJuros.Text) + CDbl(FunNuloVal(TxtMulta.Text)) + CDbl(FunNuloVal(TxtOutros.Text)) - CDbl(FunNuloVal(TxtDesagio.Text)) - CDbl(FunNuloVal(TxtDesconto.Text))), "standard")
    End If
    
End Sub

Private Sub TxtSeguro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtSeguro_LostFocus()

    subCalculaCampos (False)
    
End Sub


Private Sub TxtValorPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtValorPago_LostFocus()
    
    TxtValorPago = Format(TxtValorPago, "standard")
    
End Sub

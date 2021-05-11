VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TelaFinTitulos 
   Caption         =   "Tabela de Títulos"
   ClientHeight    =   4545
   ClientLeft      =   540
   ClientTop       =   1845
   ClientWidth     =   10935
   Icon            =   "TelaTitulos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   10935
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanPromissorias 
      Height          =   2865
      Left            =   960
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   7680
      _Version        =   65536
      _ExtentX        =   13547
      _ExtentY        =   5054
      _StockProps     =   15
      Caption         =   " Emissão de Nota Promissória"
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
      Begin VB.CheckBox chkImprimedataBase 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprime Data Base "
         Height          =   255
         Left            =   3120
         TabIndex        =   58
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Frame FraEmissao 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4440
         TabIndex        =   44
         Top             =   360
         Width           =   3135
         Begin MSComCtl2.DTPicker DTPData 
            Height          =   315
            Left            =   1785
            TabIndex        =   28
            Top             =   180
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   87687171
            CurrentDate     =   37637
            MinDate         =   29221
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Data da Emissão:"
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
            Left            =   270
            TabIndex        =   45
            Top             =   270
            Width           =   1635
         End
      End
      Begin VB.CheckBox chkCondomino 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Incluir Condômino"
         Height          =   255
         Left            =   1440
         TabIndex        =   57
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Frame fraCondomino 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados do Condômino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   52
         Top             =   2880
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox txtCpfCnpj 
            Height          =   285
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   56
            Top             =   720
            Width           =   1995
         End
         Begin VB.TextBox txtCondomino 
            Height          =   285
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   54
            Top             =   360
            Width           =   5115
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "CPF/CNPJ:"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Condômino:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkLaser 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Jato de Tinta"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command50 
         Caption         =   "Command1"
         Height          =   255
         Left            =   5280
         TabIndex        =   50
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame FraClausula 
         BackColor       =   &H00E0E0E0&
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
         Left            =   4425
         TabIndex        =   47
         Top             =   1065
         Width           =   3135
         Begin VB.CheckBox chkClausula 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cláusula de Reajuste"
            Height          =   195
            Left            =   660
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   30
            Top             =   285
            Width           =   1875
         End
      End
      Begin VB.TextBox TxtObs 
         Height          =   285
         Left            =   165
         MaxLength       =   100
         TabIndex        =   31
         Top             =   2025
         Width           =   7395
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir promissórias dos títulos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   150
         TabIndex        =   34
         Top             =   360
         Width           =   4215
         Begin VB.OptionButton OptSel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Selecionado"
            Height          =   195
            Left            =   2280
            TabIndex        =   27
            Top             =   300
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptFiltrado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Filtrados"
            Height          =   195
            Left            =   720
            TabIndex        =   26
            Top             =   300
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdImprimeNotas 
         Caption         =   "&Imprimir"
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
         Left            =   5520
         TabIndex        =   32
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton CmdRetornanotas 
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
         Left            =   6600
         TabIndex        =   33
         Top             =   2400
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5910
         Top             =   -30
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame FraFonte 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fonte:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   135
         TabIndex        =   46
         Top             =   1050
         Width           =   4200
         Begin VB.ComboBox CboFonte 
            Height          =   315
            Left            =   150
            Sorted          =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   3870
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Left            =   150
         TabIndex        =   35
         Top             =   1785
         Width           =   1095
      End
   End
   Begin Threed.SSPanel PanRepasse 
      Height          =   1650
      Left            =   180
      TabIndex        =   36
      Top             =   780
      Visible         =   0   'False
      Width           =   5130
      _Version        =   65536
      _ExtentX        =   9049
      _ExtentY        =   2910
      _StockProps     =   15
      Caption         =   " Data de Repasse"
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
      Begin VB.CommandButton CmdRepDesmarcar 
         Caption         =   "&Desfazer"
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
         Left            =   1740
         TabIndex        =   43
         Top             =   1140
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Repassar Títulos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1800
         TabIndex        =   39
         Top             =   360
         Width           =   3135
         Begin VB.OptionButton OptFiltrados 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Filtrados"
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   300
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptSelecionados 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Selecionado"
            Height          =   195
            Left            =   1320
            TabIndex        =   41
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdRepConfirma 
         Caption         =   "&Aplicar"
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
         Left            =   2820
         TabIndex        =   40
         Top             =   1155
         Width           =   975
      End
      Begin VB.CommandButton cmdRepExit 
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
         Left            =   3900
         TabIndex        =   42
         Top             =   1140
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpRepasse 
         Height          =   315
         Left            =   180
         TabIndex        =   37
         Top             =   540
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         Format          =   87687171
         CurrentDate     =   37459
      End
   End
   Begin Threed.SSPanel PanRecalcular 
      Height          =   2010
      Left            =   1920
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   4110
      _Version        =   65536
      _ExtentX        =   7250
      _ExtentY        =   3545
      _StockProps     =   15
      Caption         =   " Recalcular Títulos"
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
      Begin MSMask.MaskEdBox MskCodigoInicial 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   22
         Mask            =   "AAAA.AAAA.##.##.###.##"
         PromptChar      =   " "
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
         Left            =   2880
         TabIndex        =   20
         Top             =   1500
         Width           =   975
      End
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
         Left            =   1800
         TabIndex        =   19
         Top             =   1500
         Width           =   975
      End
      Begin MSMask.MaskEdBox MskCodigoFinal 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   22
         Mask            =   "AAAA.AAAA.##.##.###.##"
         PromptChar      =   " "
      End
      Begin VB.Label LlbCodigoFinal 
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
         TabIndex        =   22
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label LblCodigoInicial 
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
         TabIndex        =   21
         Top             =   540
         Width           =   1155
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6720
      Top             =   2940
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
      Bindings        =   "TelaTitulos.frx":030A
      Height          =   3270
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   5768
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "Titulo"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Tipo de Plano"
      Columns(1).DataField=   "tipl_tx_Descricao"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Natureza do Plano"
      Columns(2).DataField=   "Natureza"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Status do Contrato"
      Columns(3).DataField=   "Status"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Vencimento"
      Columns(4).DataField=   "titu_dt_Vencimento"
      Columns(4).NumberFormat=   "dd/mm/yy"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Valor Histórico em R$"
      Columns(5).DataField=   "titu_vl_parcela"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Indexador 1"
      Columns(6).DataField=   "Moeda1"
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Valor Indexado 1"
      Columns(7).DataField=   "titu_vl_ParcelaIndex1"
      Columns(7).NumberFormat=   "#,##0.0000"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   80
      Columns(8)._MaxComboItems=   5
      Columns(8).ValueItems(0)._DefaultItem=   0
      Columns(8).ValueItems(0).Value=   "D"
      Columns(8).ValueItems(0).Value.vt=   8
      Columns(8).ValueItems(0).DisplayValue=   "Disponível"
      Columns(8).ValueItems(0).DisplayValue.vt=   8
      Columns(8).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(8).ValueItems(1)._DefaultItem=   0
      Columns(8).ValueItems(1).Value=   "A"
      Columns(8).ValueItems(1).Value.vt=   8
      Columns(8).ValueItems(1).DisplayValue=   "Ativo"
      Columns(8).ValueItems(1).DisplayValue.vt=   8
      Columns(8).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(8).ValueItems(2)._DefaultItem=   0
      Columns(8).ValueItems(2).Value=   "I"
      Columns(8).ValueItems(2).Value.vt=   8
      Columns(8).ValueItems(2).DisplayValue=   "Inativo"
      Columns(8).ValueItems(2).DisplayValue.vt=   8
      Columns(8).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(8).ValueItems(3)._DefaultItem=   0
      Columns(8).ValueItems(3).Value=   "C"
      Columns(8).ValueItems(3).Value.vt=   8
      Columns(8).ValueItems(3).DisplayValue=   "Cedido"
      Columns(8).ValueItems(3).DisplayValue.vt=   8
      Columns(8).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(8).ValueItems.Count=   4
      Columns(8).Caption=   "Indexador 2"
      Columns(8).DataField=   "Moeda2"
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Valor Indexado 2"
      Columns(9).DataField=   "titu_vl_ParcelaIndex2"
      Columns(9).NumberFormat=   "#,##0.0000"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Pagamento"
      Columns(10).DataField=   "titu_dt_Pagamento"
      Columns(10).NumberFormat=   "dd/mm/yy"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Valor do Pagamento"
      Columns(11).DataField=   "titu_vl_ValorPago"
      Columns(11).NumberFormat=   "Standard"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Criação"
      Columns(12).DataField=   "titu_dt_Criacao"
      Columns(12).NumberFormat=   "dd/mm/yy"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Observação"
      Columns(13).DataField=   "obse_tx_Observacao"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Data Repasse"
      Columns(14).DataField=   "titu_dt_repasse"
      Columns(14).NumberFormat=   "dd/mm/yy"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Exp. Banco"
      Columns(15).DataField=   "Exportou"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Data Depósito"
      Columns(16).DataField=   "titu_dt_deposito"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Conta Corrente"
      Columns(17).DataField=   "coco_tx_descricao"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "Doc. Pagto"
      Columns(18).DataField=   "titu_tx_DocPagto"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   19
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=19"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3387"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3307"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=3440"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=3360"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=528"
      Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=2646"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2566"
      Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=528"
      Splits(0)._ColumnProps(22)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(23)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(25)=   "Column(3).Width=1905"
      Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=1826"
      Splits(0)._ColumnProps(28)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(3)._ColStyle=532"
      Splits(0)._ColumnProps(30)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(31)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(32)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(33)=   "Column(4).Width=1905"
      Splits(0)._ColumnProps(34)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(4)._WidthInPix=1826"
      Splits(0)._ColumnProps(36)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(37)=   "Column(4)._ColStyle=17"
      Splits(0)._ColumnProps(38)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(39)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(40)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(41)=   "Column(5).Width=2355"
      Splits(0)._ColumnProps(42)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(5)._WidthInPix=2275"
      Splits(0)._ColumnProps(44)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(45)=   "Column(5)._ColStyle=530"
      Splits(0)._ColumnProps(46)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(47)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(48)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(49)=   "Column(6).Width=1640"
      Splits(0)._ColumnProps(50)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(6)._WidthInPix=1561"
      Splits(0)._ColumnProps(52)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(53)=   "Column(6)._ColStyle=529"
      Splits(0)._ColumnProps(54)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(55)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(56)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(57)=   "Column(7).Width=1826"
      Splits(0)._ColumnProps(58)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(7)._WidthInPix=1746"
      Splits(0)._ColumnProps(60)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(61)=   "Column(7)._ColStyle=530"
      Splits(0)._ColumnProps(62)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(63)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(64)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(65)=   "Column(8).Width=1667"
      Splits(0)._ColumnProps(66)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(8)._WidthInPix=1588"
      Splits(0)._ColumnProps(68)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(69)=   "Column(8)._ColStyle=529"
      Splits(0)._ColumnProps(70)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(71)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(72)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(73)=   "Column(9).Width=1958"
      Splits(0)._ColumnProps(74)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(9)._WidthInPix=1879"
      Splits(0)._ColumnProps(76)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(77)=   "Column(9)._ColStyle=530"
      Splits(0)._ColumnProps(78)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(79)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(80)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(81)=   "Column(10).Width=1879"
      Splits(0)._ColumnProps(82)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(10)._WidthInPix=1799"
      Splits(0)._ColumnProps(84)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(85)=   "Column(10)._ColStyle=17"
      Splits(0)._ColumnProps(86)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(87)=   "Column(10).WrapText=1"
      Splits(0)._ColumnProps(88)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(89)=   "Column(11).Width=2249"
      Splits(0)._ColumnProps(90)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(11)._WidthInPix=2170"
      Splits(0)._ColumnProps(92)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(93)=   "Column(11)._ColStyle=530"
      Splits(0)._ColumnProps(94)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(95)=   "Column(11).WrapText=1"
      Splits(0)._ColumnProps(96)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(97)=   "Column(12).Width=2143"
      Splits(0)._ColumnProps(98)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(12)._WidthInPix=2064"
      Splits(0)._ColumnProps(100)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(101)=   "Column(12)._ColStyle=529"
      Splits(0)._ColumnProps(102)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(103)=   "Column(12).WrapText=1"
      Splits(0)._ColumnProps(104)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(105)=   "Column(13).Width=4180"
      Splits(0)._ColumnProps(106)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(13)._WidthInPix=4101"
      Splits(0)._ColumnProps(108)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(109)=   "Column(13)._ColStyle=528"
      Splits(0)._ColumnProps(110)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(111)=   "Column(13).WrapText=1"
      Splits(0)._ColumnProps(112)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(113)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(114)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(116)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(117)=   "Column(14)._ColStyle=20"
      Splits(0)._ColumnProps(118)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(119)=   "Column(14).WrapText=1"
      Splits(0)._ColumnProps(120)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(121)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(122)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(123)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(124)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(125)=   "Column(15)._ColStyle=20"
      Splits(0)._ColumnProps(126)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(127)=   "Column(15).WrapText=1"
      Splits(0)._ColumnProps(128)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(129)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(130)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(131)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(132)=   "Column(16)._EditAlways=0"
      Splits(0)._ColumnProps(133)=   "Column(16)._ColStyle=20"
      Splits(0)._ColumnProps(134)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(135)=   "Column(16).WrapText=1"
      Splits(0)._ColumnProps(136)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(137)=   "Column(16)._MinWidth=136065472"
      Splits(0)._ColumnProps(138)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(139)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(140)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(141)=   "Column(17)._EditAlways=0"
      Splits(0)._ColumnProps(142)=   "Column(17)._ColStyle=20"
      Splits(0)._ColumnProps(143)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(144)=   "Column(17).WrapText=1"
      Splits(0)._ColumnProps(145)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(146)=   "Column(17)._MinWidth=161851696"
      Splits(0)._ColumnProps(147)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(148)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(149)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(150)=   "Column(18)._EditAlways=0"
      Splits(0)._ColumnProps(151)=   "Column(18)._ColStyle=20"
      Splits(0)._ColumnProps(152)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(153)=   "Column(18).WrapText=1"
      Splits(0)._ColumnProps(154)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(155)=   "Column(18)._MinWidth=134436432"
      Splits(1)._UserFlags=   0
      Splits(1).ExtendRightColumn=   -1  'True
      Splits(1).MarqueeStyle=   3
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).AllowColMove=   -1  'True
      Splits(1).FetchRowStyle=   -1  'True
      Splits(1).DividerStyle=   4
      Splits(1).DividerColor=   12632256
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=19"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=3387"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=3307"
      Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=528"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).WrapText=1"
      Splits(1)._ColumnProps(8)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(9)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(10)=   "Column(1).Width=3440"
      Splits(1)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(12)=   "Column(1)._WidthInPix=3360"
      Splits(1)._ColumnProps(13)=   "Column(1)._EditAlways=0"
      Splits(1)._ColumnProps(14)=   "Column(1)._ColStyle=528"
      Splits(1)._ColumnProps(15)=   "Column(1).WrapText=1"
      Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(17)=   "Column(2).Width=2831"
      Splits(1)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(19)=   "Column(2)._WidthInPix=2752"
      Splits(1)._ColumnProps(20)=   "Column(2)._EditAlways=0"
      Splits(1)._ColumnProps(21)=   "Column(2)._ColStyle=528"
      Splits(1)._ColumnProps(22)=   "Column(2).WrapText=1"
      Splits(1)._ColumnProps(23)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(24)=   "Column(3).Width=1905"
      Splits(1)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(26)=   "Column(3)._WidthInPix=1826"
      Splits(1)._ColumnProps(27)=   "Column(3)._EditAlways=0"
      Splits(1)._ColumnProps(28)=   "Column(3)._ColStyle=532"
      Splits(1)._ColumnProps(29)=   "Column(3).WrapText=1"
      Splits(1)._ColumnProps(30)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(31)=   "Column(4).Width=1905"
      Splits(1)._ColumnProps(32)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(33)=   "Column(4)._WidthInPix=1826"
      Splits(1)._ColumnProps(34)=   "Column(4)._EditAlways=0"
      Splits(1)._ColumnProps(35)=   "Column(4)._ColStyle=17"
      Splits(1)._ColumnProps(36)=   "Column(4).WrapText=1"
      Splits(1)._ColumnProps(37)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(38)=   "Column(5).Width=2408"
      Splits(1)._ColumnProps(39)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(40)=   "Column(5)._WidthInPix=2328"
      Splits(1)._ColumnProps(41)=   "Column(5)._EditAlways=0"
      Splits(1)._ColumnProps(42)=   "Column(5)._ColStyle=530"
      Splits(1)._ColumnProps(43)=   "Column(5).WrapText=1"
      Splits(1)._ColumnProps(44)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(45)=   "Column(6).Width=1640"
      Splits(1)._ColumnProps(46)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(47)=   "Column(6)._WidthInPix=1561"
      Splits(1)._ColumnProps(48)=   "Column(6)._EditAlways=0"
      Splits(1)._ColumnProps(49)=   "Column(6)._ColStyle=529"
      Splits(1)._ColumnProps(50)=   "Column(6).WrapText=1"
      Splits(1)._ColumnProps(51)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(52)=   "Column(7).Width=1984"
      Splits(1)._ColumnProps(53)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(54)=   "Column(7)._WidthInPix=1905"
      Splits(1)._ColumnProps(55)=   "Column(7)._EditAlways=0"
      Splits(1)._ColumnProps(56)=   "Column(7)._ColStyle=530"
      Splits(1)._ColumnProps(57)=   "Column(7).WrapText=1"
      Splits(1)._ColumnProps(58)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(59)=   "Column(7)._MinWidth=103679444"
      Splits(1)._ColumnProps(60)=   "Column(8).Width=1667"
      Splits(1)._ColumnProps(61)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(62)=   "Column(8)._WidthInPix=1588"
      Splits(1)._ColumnProps(63)=   "Column(8)._EditAlways=0"
      Splits(1)._ColumnProps(64)=   "Column(8)._ColStyle=529"
      Splits(1)._ColumnProps(65)=   "Column(8).WrapText=1"
      Splits(1)._ColumnProps(66)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(67)=   "Column(9).Width=1958"
      Splits(1)._ColumnProps(68)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(69)=   "Column(9)._WidthInPix=1879"
      Splits(1)._ColumnProps(70)=   "Column(9)._EditAlways=0"
      Splits(1)._ColumnProps(71)=   "Column(9)._ColStyle=530"
      Splits(1)._ColumnProps(72)=   "Column(9).WrapText=1"
      Splits(1)._ColumnProps(73)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(74)=   "Column(10).Width=1879"
      Splits(1)._ColumnProps(75)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(76)=   "Column(10)._WidthInPix=1799"
      Splits(1)._ColumnProps(77)=   "Column(10)._EditAlways=0"
      Splits(1)._ColumnProps(78)=   "Column(10)._ColStyle=17"
      Splits(1)._ColumnProps(79)=   "Column(10).WrapText=1"
      Splits(1)._ColumnProps(80)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(81)=   "Column(11).Width=2249"
      Splits(1)._ColumnProps(82)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(83)=   "Column(11)._WidthInPix=2170"
      Splits(1)._ColumnProps(84)=   "Column(11)._EditAlways=0"
      Splits(1)._ColumnProps(85)=   "Column(11)._ColStyle=530"
      Splits(1)._ColumnProps(86)=   "Column(11).WrapText=1"
      Splits(1)._ColumnProps(87)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(88)=   "Column(12).Width=2143"
      Splits(1)._ColumnProps(89)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(90)=   "Column(12)._WidthInPix=2064"
      Splits(1)._ColumnProps(91)=   "Column(12)._EditAlways=0"
      Splits(1)._ColumnProps(92)=   "Column(12)._ColStyle=529"
      Splits(1)._ColumnProps(93)=   "Column(12).WrapText=1"
      Splits(1)._ColumnProps(94)=   "Column(12).Order=14"
      Splits(1)._ColumnProps(95)=   "Column(13).Width=4180"
      Splits(1)._ColumnProps(96)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(97)=   "Column(13)._WidthInPix=4101"
      Splits(1)._ColumnProps(98)=   "Column(13)._EditAlways=0"
      Splits(1)._ColumnProps(99)=   "Column(13)._ColStyle=528"
      Splits(1)._ColumnProps(100)=   "Column(13).WrapText=1"
      Splits(1)._ColumnProps(101)=   "Column(13).Order=15"
      Splits(1)._ColumnProps(102)=   "Column(13)._MinWidth=64"
      Splits(1)._ColumnProps(103)=   "Column(14).Width=2725"
      Splits(1)._ColumnProps(104)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(105)=   "Column(14)._WidthInPix=2646"
      Splits(1)._ColumnProps(106)=   "Column(14)._EditAlways=0"
      Splits(1)._ColumnProps(107)=   "Column(14)._ColStyle=529"
      Splits(1)._ColumnProps(108)=   "Column(14).WrapText=1"
      Splits(1)._ColumnProps(109)=   "Column(14).Order=13"
      Splits(1)._ColumnProps(110)=   "Column(15).Width=2725"
      Splits(1)._ColumnProps(111)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(112)=   "Column(15)._WidthInPix=2646"
      Splits(1)._ColumnProps(113)=   "Column(15)._EditAlways=0"
      Splits(1)._ColumnProps(114)=   "Column(15)._ColStyle=529"
      Splits(1)._ColumnProps(115)=   "Column(15).WrapText=1"
      Splits(1)._ColumnProps(116)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(117)=   "Column(15)._MinWidth=73194978"
      Splits(1)._ColumnProps(118)=   "Column(16).Width=2725"
      Splits(1)._ColumnProps(119)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(120)=   "Column(16)._WidthInPix=2646"
      Splits(1)._ColumnProps(121)=   "Column(16)._EditAlways=0"
      Splits(1)._ColumnProps(122)=   "Column(16)._ColStyle=20"
      Splits(1)._ColumnProps(123)=   "Column(16).WrapText=1"
      Splits(1)._ColumnProps(124)=   "Column(16).Order=17"
      Splits(1)._ColumnProps(125)=   "Column(16)._MinWidth=135981568"
      Splits(1)._ColumnProps(126)=   "Column(17).Width=2725"
      Splits(1)._ColumnProps(127)=   "Column(17).DividerColor=0"
      Splits(1)._ColumnProps(128)=   "Column(17)._WidthInPix=2646"
      Splits(1)._ColumnProps(129)=   "Column(17)._EditAlways=0"
      Splits(1)._ColumnProps(130)=   "Column(17)._ColStyle=20"
      Splits(1)._ColumnProps(131)=   "Column(17).WrapText=1"
      Splits(1)._ColumnProps(132)=   "Column(17).Order=18"
      Splits(1)._ColumnProps(133)=   "Column(18).Width=2725"
      Splits(1)._ColumnProps(134)=   "Column(18).DividerColor=0"
      Splits(1)._ColumnProps(135)=   "Column(18)._WidthInPix=2646"
      Splits(1)._ColumnProps(136)=   "Column(18)._EditAlways=0"
      Splits(1)._ColumnProps(137)=   "Column(18)._ColStyle=20"
      Splits(1)._ColumnProps(138)=   "Column(18).WrapText=1"
      Splits(1)._ColumnProps(139)=   "Column(18).Order=19"
      Splits(1)._ColumnProps(140)=   "Column(18)._MinWidth=134358256"
      Splits.Count    =   2
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
      _StyleDefs(24)  =   "Splits(0).Style:id=83,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=92,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=84,.parent=2,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(27)  =   ":id=84,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(28)  =   ":id=84,.fontname=MS Sans Serif"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=85,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=86,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=88,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=87,.parent=7,.wraptext=-1"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=89,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=90,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=91,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=93,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=94,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=98,.parent=83,.alignment=0,.locked=0"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=95,.parent=84,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=96,.parent=85,.alignment=3"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=97,.parent=87,.alignment=3"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=110,.parent=83,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=84,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=85"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=87"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=114,.parent=83,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=111,.parent=84,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=112,.parent=85"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=113,.parent=87"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=118,.parent=83"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=115,.parent=84,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=116,.parent=85"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=117,.parent=87"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=122,.parent=83,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=119,.parent=84"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=120,.parent=85"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=121,.parent=87"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=126,.parent=83,.alignment=1"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=123,.parent=84,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=124,.parent=85"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=125,.parent=87"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=130,.parent=83,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=127,.parent=84,.alignment=2"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=128,.parent=85"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=129,.parent=87"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=134,.parent=83,.alignment=1"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=131,.parent=84,.alignment=2"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=132,.parent=85"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=133,.parent=87"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=138,.parent=83,.alignment=2"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=135,.parent=84,.alignment=2"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=136,.parent=85"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=137,.parent=87"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=142,.parent=83,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=139,.parent=84,.alignment=2"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=140,.parent=85"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=141,.parent=87"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=146,.parent=83,.alignment=2"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=143,.parent=84"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=144,.parent=85"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=145,.parent=87"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=150,.parent=83,.alignment=1"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=147,.parent=84,.alignment=2"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=148,.parent=85"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=149,.parent=87"
      _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=154,.parent=83,.alignment=2"
      _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=151,.parent=84,.alignment=2"
      _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=152,.parent=85"
      _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=153,.parent=87"
      _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=158,.parent=83,.alignment=0"
      _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=155,.parent=84,.alignment=2"
      _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=156,.parent=85"
      _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=157,.parent=87"
      _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=162,.parent=83"
      _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=159,.parent=84"
      _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=160,.parent=85"
      _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=161,.parent=87"
      _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=170,.parent=83"
      _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=167,.parent=84"
      _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=168,.parent=85"
      _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=169,.parent=87"
      _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=178,.parent=83"
      _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=175,.parent=84"
      _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=176,.parent=85"
      _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=177,.parent=87"
      _StyleDefs(106) =   "Splits(0).Columns(17).Style:id=186,.parent=83"
      _StyleDefs(107) =   "Splits(0).Columns(17).HeadingStyle:id=183,.parent=84"
      _StyleDefs(108) =   "Splits(0).Columns(17).FooterStyle:id=184,.parent=85"
      _StyleDefs(109) =   "Splits(0).Columns(17).EditorStyle:id=185,.parent=87"
      _StyleDefs(110) =   "Splits(0).Columns(18).Style:id=194,.parent=83"
      _StyleDefs(111) =   "Splits(0).Columns(18).HeadingStyle:id=191,.parent=84"
      _StyleDefs(112) =   "Splits(0).Columns(18).FooterStyle:id=192,.parent=85"
      _StyleDefs(113) =   "Splits(0).Columns(18).EditorStyle:id=193,.parent=87"
      _StyleDefs(114) =   "Splits(1).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(115) =   "Splits(1).CaptionStyle:id=22,.parent=4"
      _StyleDefs(116) =   "Splits(1).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(117) =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(118) =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(119) =   "Splits(1).FooterStyle:id=15,.parent=3"
      _StyleDefs(120) =   "Splits(1).InactiveStyle:id=16,.parent=5"
      _StyleDefs(121) =   "Splits(1).SelectedStyle:id=18,.parent=6"
      _StyleDefs(122) =   "Splits(1).EditorStyle:id=17,.parent=7,.wraptext=-1"
      _StyleDefs(123) =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(124) =   "Splits(1).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(125) =   "Splits(1).OddRowStyle:id=21,.parent=10"
      _StyleDefs(126) =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(127) =   "Splits(1).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(128) =   "Splits(1).Columns(0).Style:id=28,.parent=13,.alignment=0,.locked=0"
      _StyleDefs(129) =   "Splits(1).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(130) =   "Splits(1).Columns(0).FooterStyle:id=26,.parent=15,.alignment=3"
      _StyleDefs(131) =   "Splits(1).Columns(0).EditorStyle:id=27,.parent=17,.alignment=3"
      _StyleDefs(132) =   "Splits(1).Columns(1).Style:id=102,.parent=13,.alignment=0"
      _StyleDefs(133) =   "Splits(1).Columns(1).HeadingStyle:id=99,.parent=14,.alignment=2"
      _StyleDefs(134) =   "Splits(1).Columns(1).FooterStyle:id=100,.parent=15"
      _StyleDefs(135) =   "Splits(1).Columns(1).EditorStyle:id=101,.parent=17"
      _StyleDefs(136) =   "Splits(1).Columns(2).Style:id=78,.parent=13,.alignment=0"
      _StyleDefs(137) =   "Splits(1).Columns(2).HeadingStyle:id=75,.parent=14,.alignment=2"
      _StyleDefs(138) =   "Splits(1).Columns(2).FooterStyle:id=76,.parent=15"
      _StyleDefs(139) =   "Splits(1).Columns(2).EditorStyle:id=77,.parent=17"
      _StyleDefs(140) =   "Splits(1).Columns(3).Style:id=82,.parent=13"
      _StyleDefs(141) =   "Splits(1).Columns(3).HeadingStyle:id=79,.parent=14,.alignment=2"
      _StyleDefs(142) =   "Splits(1).Columns(3).FooterStyle:id=80,.parent=15"
      _StyleDefs(143) =   "Splits(1).Columns(3).EditorStyle:id=81,.parent=17"
      _StyleDefs(144) =   "Splits(1).Columns(4).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(145) =   "Splits(1).Columns(4).HeadingStyle:id=63,.parent=14"
      _StyleDefs(146) =   "Splits(1).Columns(4).FooterStyle:id=64,.parent=15"
      _StyleDefs(147) =   "Splits(1).Columns(4).EditorStyle:id=65,.parent=17"
      _StyleDefs(148) =   "Splits(1).Columns(5).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(149) =   "Splits(1).Columns(5).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(150) =   "Splits(1).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(151) =   "Splits(1).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(152) =   "Splits(1).Columns(6).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(153) =   "Splits(1).Columns(6).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(154) =   "Splits(1).Columns(6).FooterStyle:id=30,.parent=15"
      _StyleDefs(155) =   "Splits(1).Columns(6).EditorStyle:id=31,.parent=17"
      _StyleDefs(156) =   "Splits(1).Columns(7).Style:id=46,.parent=13,.alignment=1"
      _StyleDefs(157) =   "Splits(1).Columns(7).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(158) =   "Splits(1).Columns(7).FooterStyle:id=44,.parent=15"
      _StyleDefs(159) =   "Splits(1).Columns(7).EditorStyle:id=45,.parent=17"
      _StyleDefs(160) =   "Splits(1).Columns(8).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(161) =   "Splits(1).Columns(8).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(162) =   "Splits(1).Columns(8).FooterStyle:id=48,.parent=15"
      _StyleDefs(163) =   "Splits(1).Columns(8).EditorStyle:id=49,.parent=17"
      _StyleDefs(164) =   "Splits(1).Columns(9).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(165) =   "Splits(1).Columns(9).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(166) =   "Splits(1).Columns(9).FooterStyle:id=56,.parent=15"
      _StyleDefs(167) =   "Splits(1).Columns(9).EditorStyle:id=57,.parent=17"
      _StyleDefs(168) =   "Splits(1).Columns(10).Style:id=106,.parent=13,.alignment=2"
      _StyleDefs(169) =   "Splits(1).Columns(10).HeadingStyle:id=103,.parent=14"
      _StyleDefs(170) =   "Splits(1).Columns(10).FooterStyle:id=104,.parent=15"
      _StyleDefs(171) =   "Splits(1).Columns(10).EditorStyle:id=105,.parent=17"
      _StyleDefs(172) =   "Splits(1).Columns(11).Style:id=74,.parent=13,.alignment=1"
      _StyleDefs(173) =   "Splits(1).Columns(11).HeadingStyle:id=71,.parent=14,.alignment=2"
      _StyleDefs(174) =   "Splits(1).Columns(11).FooterStyle:id=72,.parent=15"
      _StyleDefs(175) =   "Splits(1).Columns(11).EditorStyle:id=73,.parent=17"
      _StyleDefs(176) =   "Splits(1).Columns(12).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(177) =   "Splits(1).Columns(12).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(178) =   "Splits(1).Columns(12).FooterStyle:id=60,.parent=15"
      _StyleDefs(179) =   "Splits(1).Columns(12).EditorStyle:id=61,.parent=17"
      _StyleDefs(180) =   "Splits(1).Columns(13).Style:id=70,.parent=13,.alignment=0"
      _StyleDefs(181) =   "Splits(1).Columns(13).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(182) =   "Splits(1).Columns(13).FooterStyle:id=68,.parent=15"
      _StyleDefs(183) =   "Splits(1).Columns(13).EditorStyle:id=69,.parent=17"
      _StyleDefs(184) =   "Splits(1).Columns(14).Style:id=166,.parent=13,.alignment=2"
      _StyleDefs(185) =   "Splits(1).Columns(14).HeadingStyle:id=163,.parent=14,.alignment=2"
      _StyleDefs(186) =   "Splits(1).Columns(14).FooterStyle:id=164,.parent=15"
      _StyleDefs(187) =   "Splits(1).Columns(14).EditorStyle:id=165,.parent=17"
      _StyleDefs(188) =   "Splits(1).Columns(15).Style:id=174,.parent=13,.alignment=2"
      _StyleDefs(189) =   "Splits(1).Columns(15).HeadingStyle:id=171,.parent=14,.alignment=2"
      _StyleDefs(190) =   "Splits(1).Columns(15).FooterStyle:id=172,.parent=15"
      _StyleDefs(191) =   "Splits(1).Columns(15).EditorStyle:id=173,.parent=17"
      _StyleDefs(192) =   "Splits(1).Columns(16).Style:id=182,.parent=13"
      _StyleDefs(193) =   "Splits(1).Columns(16).HeadingStyle:id=179,.parent=14"
      _StyleDefs(194) =   "Splits(1).Columns(16).FooterStyle:id=180,.parent=15"
      _StyleDefs(195) =   "Splits(1).Columns(16).EditorStyle:id=181,.parent=17"
      _StyleDefs(196) =   "Splits(1).Columns(17).Style:id=190,.parent=13"
      _StyleDefs(197) =   "Splits(1).Columns(17).HeadingStyle:id=187,.parent=14"
      _StyleDefs(198) =   "Splits(1).Columns(17).FooterStyle:id=188,.parent=15"
      _StyleDefs(199) =   "Splits(1).Columns(17).EditorStyle:id=189,.parent=17"
      _StyleDefs(200) =   "Splits(1).Columns(18).Style:id=198,.parent=13"
      _StyleDefs(201) =   "Splits(1).Columns(18).HeadingStyle:id=195,.parent=14"
      _StyleDefs(202) =   "Splits(1).Columns(18).FooterStyle:id=196,.parent=15"
      _StyleDefs(203) =   "Splits(1).Columns(18).EditorStyle:id=197,.parent=17"
      _StyleDefs(204) =   "Named:id=33:Normal"
      _StyleDefs(205) =   ":id=33,.parent=0"
      _StyleDefs(206) =   "Named:id=34:Heading"
      _StyleDefs(207) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(208) =   ":id=34,.wraptext=-1"
      _StyleDefs(209) =   "Named:id=35:Footing"
      _StyleDefs(210) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(211) =   "Named:id=36:Selected"
      _StyleDefs(212) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(213) =   "Named:id=37:Caption"
      _StyleDefs(214) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(215) =   "Named:id=38:HighlightRow"
      _StyleDefs(216) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(217) =   "Named:id=39:EvenRow"
      _StyleDefs(218) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(219) =   "Named:id=40:OddRow"
      _StyleDefs(220) =   ":id=40,.parent=33"
      _StyleDefs(221) =   "Named:id=41:RecordSelector"
      _StyleDefs(222) =   ":id=41,.parent=34"
      _StyleDefs(223) =   "Named:id=42:FilterBar"
      _StyleDefs(224) =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   1140
      Left            =   60
      TabIndex        =   16
      Top             =   3420
      Width           =   10860
      _Version        =   65536
      _ExtentX        =   19156
      _ExtentY        =   2011
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
      Begin VB.CommandButton cmdExport 
         Height          =   330
         Left            =   8880
         Picture         =   "TelaTitulos.frx":031F
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdRepasse 
         Caption         =   "&Repasse"
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
         Left            =   3780
         TabIndex        =   9
         Top             =   780
         Width           =   1035
      End
      Begin VB.CommandButton cmdPromissoria 
         BackColor       =   &H00000000&
         Caption         =   "Pro&missórias"
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
         Left            =   2400
         TabIndex        =   8
         Top             =   780
         Width           =   1275
      End
      Begin VB.CommandButton CmdPlanos 
         BackColor       =   &H00000000&
         Caption         =   "P&lanos"
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
         Left            =   0
         TabIndex        =   6
         Top             =   795
         Width           =   1035
      End
      Begin VB.CommandButton CmdPagamento 
         Caption         =   "&Pagamento"
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
         TabIndex        =   7
         Top             =   780
         Width           =   1155
      End
      Begin VB.CommandButton CmdAlteracaoMultipla 
         BackColor       =   &H00000000&
         Caption         =   "Alteração &Múltipla"
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
         Left            =   3780
         TabIndex        =   3
         Top             =   360
         Width           =   1875
      End
      Begin VB.CommandButton CmdExclusaoMultipla 
         BackColor       =   &H00000000&
         Caption         =   "E&xclusão Múltipla"
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
         Left            =   5796
         TabIndex        =   4
         Top             =   360
         Width           =   1755
      End
      Begin VB.CommandButton CmdRecalcular 
         BackColor       =   &H00000000&
         Caption         =   "Re&calcular"
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
         Left            =   7680
         TabIndex        =   5
         Top             =   360
         Width           =   1095
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
         Left            =   0
         TabIndex        =   0
         Top             =   360
         Width           =   1095
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
         Left            =   7920
         TabIndex        =   13
         Top             =   780
         Width           =   855
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
         Left            =   6900
         TabIndex        =   12
         Top             =   780
         Width           =   915
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
         Left            =   6000
         TabIndex        =   11
         Top             =   780
         Width           =   795
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
         Left            =   4920
         TabIndex        =   10
         Top             =   780
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
         Left            =   2508
         TabIndex        =   2
         Top             =   360
         Width           =   1155
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
         Left            =   1224
         TabIndex        =   1
         Top             =   360
         Width           =   1155
      End
      Begin MSMask.MaskEdBox MskTitulo 
         Height          =   315
         Left            =   6660
         TabIndex        =   15
         ToolTipText     =   "Digite o título e aperte ENTER"
         Top             =   0
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   22
         Mask            =   "AAAA.AAAA.##.##.###.##"
         PromptChar      =   " "
      End
      Begin MSDataListLib.DataCombo cboLocalizar 
         Bindings        =   "TelaTitulos.frx":04CE
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Top             =   0
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "titulo"
         BoundColumn     =   "titulo"
         Text            =   ""
      End
      Begin VB.Label lblLocalizar 
         Alignment       =   1  'Right Justify
         Caption         =   "Localizar Código:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -1020
         TabIndex        =   49
         Top             =   60
         Width           =   2670
      End
      Begin VB.Label LblTitulo 
         Alignment       =   1  'Right Justify
         Caption         =   "Título:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6180
         TabIndex        =   48
         Top             =   60
         Width           =   450
      End
   End
End
Attribute VB_Name = "TelaFinTitulos"
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
Dim XFB_CARREGOUFONTE As Boolean

Function VerificaCpfCnpj() As Boolean
    
    If Me.txtCpfCnpj.Text = "" Or Me.txtCondomino = "" Then
        VerificaCpfCnpj = False
        Exit Function
    End If
    
    VerificaCpfCnpj = True
End Function


'Imprrime Promissórias na Laser - Notas da Costa Andrade
Private Sub subImprimePromissoriaLaser()
    
    Dim XLO_IMPRESSORA As Printer
    Dim XLT_IMPRESSORA As String
    Dim XLT_SQL As String
    
    
    'Abre os recordsets e verificar se eles possuem dados
    If OptSel.Value = True Then
        Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM consCAPFormularios WHERE titulo = '" & Adodc1.Recordset.Fields("Titulo") & "' AND empr_cd_empresa = " & PCodEmpresa & "")
    Else
        Adodc1.Recordset.MoveFirst
        
        While Not (Adodc1.Recordset.EOF)
            XLT_SQL = "'" & Adodc1.Recordset.Fields("Titulo") & "' , " & XLT_SQL
            Adodc1.Recordset.MoveNext
        Wend
        
        'Removendo a última vírgula
        XLT_SQL = Mid(XLT_SQL, 1, Len(XLT_SQL) - 2)
        
        Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM consCAPFormularios WHERE titulo IN (" & XLT_SQL & ") AND empr_cd_empresa = " & PCodEmpresa & " ORDER BY Titulo ")
    
    End If
    
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        TelaFinTitulos.MousePointer = vbDefault
        Exit Sub
    End If

    'Abre o relatório
    FrmVisRelatorios.Caption = "Promissórias"
    
    If UCase(PCliente) = "FRANISA" Then
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\NotaPromissoriaFranisa.rpt")
    ElseIf UCase(PCliente) = "COSTA ANDRADE" Then
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\NotaPromissoriaCA.rpt")
    ElseIf UCase(PCliente) = "ACTITUR" Then
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\NotaPromissoriaActitur.rpt")
    ElseIf UCase(PCliente) = "RJ" Then
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\NotaPromissoriaRJ.rpt")
    Else
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\NotaPromissoria.rpt")
    End If
    
    
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO, 3, 1


    If UCase(PCliente) <> "FRANISA" Then
    
        '********** Ativa os Relatórios
        With XGR_RELATORIO.FormulaFields
             
             If Me.chkCondomino.Value = 1 Then
               .GetItemByName("temcondomino").Text = Chr(34) + "sim" + Chr(34)
             Else
               .GetItemByName("temcondomino").Text = Chr(34) + "nao" + Chr(34)
             End If
             
             .GetItemByName("EmpresaCondomino").Text = Chr(34) + Me.txtCondomino.Text + Chr(34)
             
             .GetItemByName("CpfCnpjCondomino").Text = Chr(34) + Me.txtCpfCnpj.Text + Chr(34)
             
             If Me.chkImprimedataBase.Value = 1 Then
                .GetItemByName("imprimedatabase").Text = Chr(34) + "sim" + Chr(34)
             Else
                .GetItemByName("imprimedatabase").Text = Chr(34) + "nao" + Chr(34)
             End If
             
        End With
    End If
    
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""

    Exit Sub

End Sub

Private Sub SubImprimePromissoriaFranisa()

    Dim RecebeCampo As String
    Dim XLI_QTDPROMISSORIA As Integer 'Quantidade de promissorias a serem impressos
    Dim XLI_QTDIMPRESSOS As Integer 'Quantidade de promissorias impressas
    Dim XLS_SQL As String
    Dim XLO_RSTITU As ADODB.Recordset
    
    Open "lpt1" For Output As #1
    
    XLI_QTDPROMISSORIA = 0
            
    XLI_QTDIMPRESSOS = 0

    XLI_QTDIMPRESSOS = XLI_QTDIMPRESSOS + 1
    
    '***** Impressão de Promissórias *****
    'Código para impressão direto na porta

    XLS_SQL = "SELECT * FROM consCAPFormularios WHERE titulo = '" & Adodc1.Recordset.Fields("Titulo") & "' AND empr_cd_empresa = " & PCodEmpresa
    SubQOpenRecordset XLO_RSTITU, XLS_SQL, Estatico
    
    '****************CONFIGURANDO PROMISSORIAS - IMPRIMINDO PELA PORTA DA IMPRESSORA*******************
    'limpando todas as configuracoes da impressora
    Print #1, Chr(27) + "@"
    'setando fonte para roman
    Print #1, Chr(27) + "x1"; Chr(27) + "k0"
     
    'Codigo do Imovel
    RecebeCampo = XLO_RSTITU!CodImovel
    Print #1, Space(29); Chr(27) + "M"; funTiraAcento(RecebeCampo, True) + Space(8) + XLO_RSTITU!CodTitulo + Space(15) + FunNulo(Format$(XLO_RSTITU!ValorTitulo, "standard"))
    
    'Pula Linha
    Print #1, ""
    
    'Data de Vencimento
    RecebeCampo = XLO_RSTITU!DataVenc
    Print #1, Space(34); Chr(27) + "M"; UCase(Format(funTiraAcento(RecebeCampo, True), "dd \de mmmm \de yyyy"))
    
    'Pula Linha
    Print #1, ""
    
    Dim diaExtenso As String
            diaExtenso = funExtenso2(Format(funTiraAcento(RecebeCampo, True), "dd"))
            Dim mesExtenso As String
            mesExtenso = funTiraAcento(UCase(Format(RecebeCampo, "mmmm")), True)
            
    'Data de vencimento
    RecebeCampo = XLO_RSTITU!DataVenc
    Print #1, Space(27); Chr(27) + "M" + diaExtenso & Space(18 - Len(diaExtenso)) & _
            Space(12) & _
            mesExtenso & Space(9 - Len(mesExtenso)) & _
            Space(7) & _
            Format(RecebeCampo, "yyyy")

    'Pula Linha
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    'Valor do título - Extenso
    RecebeCampo = funExtenso(XLO_RSTITU!ValorTitulo)
    Print #1, Space(25); Chr(27) + "M" + funTiraAcento(RecebeCampo, True)
    
    'Pula Linha
    Print #1, ""
    Print #1, ""
    
    'Observação
    RecebeCampo = " EM MOEDA CORRENTE NACIONAL, REAJUSTADOS CONFORME O PACTUADO EM"
    Print #1, Space(22); Chr(27) + "M" + funTiraAcento(RecebeCampo, True)
            
    RecebeCampo = " CONTRATO, NA PRACA DE SALVADOR - BA *** *** *** *** *** ***"
    Print #1, Space(22); Chr(27) + "M" + funTiraAcento(RecebeCampo, True)
    
    'Pula Linha
    Print #1, ""
    
    'Nome do Cliente
    RecebeCampo = XLO_RSTITU!NomeCliente
    Print #1, Space(29); Chr(27) + "M" + funTiraAcento(RecebeCampo, True)
    
    'CPF_CGC
    RecebeCampo = XLO_RSTITU!CPF_CGC
    Print #1, Space(29); Chr(27) + "M" + funTiraAcento(RecebeCampo, True)
    
    'Endereço
    RecebeCampo = XLO_RSTITU!EndCorrespondencia
    Print #1, Space(29); Chr(27) + "M" + funTiraAcento(RecebeCampo, True)
    
    'cidade + cep
    RecebeCampo = XLO_RSTITU!Bairro & " " & XLO_RSTITU!CEP
    Print #1, Space(29); Chr(27) + "M" + funTiraAcento(RecebeCampo, True)
          
    
    'Município
    Dim x As String
    
    RecebeCampo = XLO_RSTITU!CidadeEstadoEmpreendimento
    x = XLO_RSTITU!DataVenda
    Print #1, Space(23); Chr(27) + "M" + funTiraAcento(RecebeCampo, True) + Space(3) + x

    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""


    XLO_RSTITU.Close
    
    Close #1
 
Exit Sub

TrataErro:
    funTrataErros (ComMensagem)
    Conexao.RollbackTrans
    Exit Sub

End Sub
Private Sub SubImprimePromissoriaEbisa()

    Dim RecebeCampo As String
    Dim XLI_QTDPROMISSORIA As Integer 'Quantidade de promissorias a serem impressos
    Dim XLI_QTDIMPRESSOS As Integer 'Quantidade de promissorias impressas
    Dim XLS_SQL As String
    Dim XLO_RSTITU As ADODB.Recordset
    
    Open "lpt1" For Output As #1
    
    XLI_QTDPROMISSORIA = 0
              
    XLI_QTDIMPRESSOS = 0

    XLI_QTDIMPRESSOS = XLI_QTDIMPRESSOS + 1
    
    '***** Impressão de Promissórias *****
    'Código para impressão direto na porta
    
    XLS_SQL = "SELECT * FROM consCAPFormularios WHERE titulo = '" & Adodc1.Recordset.Fields("Titulo") & "' AND empr_cd_empresa = " & PCodEmpresa
    SubQOpenRecordset XLO_RSTITU, XLS_SQL, Estatico
    
    '****************CONFIGURANDO PROMISSORIAS - IMPRIMINDO PELA PORTA DA IMPRESSORA*******************
    'limpando todas as configuracoes da impressora
    Print #1, Chr(27) + "@"
    'setando fonte para sans
    Print #1, Chr(27) + "x1"; Chr(27) + "k1"
     
    'Codigo do Imovel + Vencimento
    RecebeCampo = XLO_RSTITU!CodImovel
    Print #1, Space(0); Chr(27) + "M"; funTiraAcento(RecebeCampo, True) + Space(2) + XLO_RSTITU!CodTitulo + Space(12) + FunNulo(Format$(funTiraAcento(XLO_RSTITU!ValorTitulo, True), "standard")) + Space(23) + UCase(Format(funTiraAcento(XLO_RSTITU!DataVenc, True), "dd \      mmmm \        yyyy"))
    
    'Pula Linha
    Print #1, ""
    
    'Data de Vencimento
    RecebeCampo = XLO_RSTITU!DataVenc
    Print #1, Space(20); Chr(27) + "M"; UCase(Format(funTiraAcento(RecebeCampo, True), "dd \dia\s \do \me\s \de mmmm \do a\no \de yyyy"))


    'Pula Linha
    Print #1, ""
    Print #1, ""
    
    'Empresa + CNPJ
    RecebeCampo = XLO_RSTITU!NomeEmpresa
    Print #1, Space(20); Chr(27) + "M"; UCase(funTiraAcento(XLO_RSTITU!NomeEmpresa, True)) + Space(7) + XLO_RSTITU!CGC

    'Pula Linha
    Print #1, ""
    Print #1, ""
    
    'Valor do título - Extenso
    RecebeCampo = funExtenso(XLO_RSTITU!ValorTitulo)
    Print #1, Space(30); Chr(27) + "M" + funTiraAcento(RecebeCampo, True) + "*** *** ***"
    
    'Município
    Dim x As String
    
    RecebeCampo = XLO_RSTITU!CidadeEstadoEmpresa
    Print #1, Space(80); Chr(27) + "M" + funTiraAcento(RecebeCampo, True) + Space(3)
    
    'Pula Linha
    Print #1, ""
    Print #1, ""
      
    'Data da Venda
    RecebeCampo = XLO_RSTITU!DataVenda
    Print #1, Space(65); Chr(27) + "M" + UCase(Format(RecebeCampo, "dd \de mmmm \de yyyy"))
  
    'Nome do Cliente
    RecebeCampo = XLO_RSTITU!NomeCliente
    Print #1, Space(15); Chr(27) + "M" + funTiraAcento(RecebeCampo, True)
    
    'CPF_CGC
    RecebeCampo = XLO_RSTITU!CPF_CGC
    Print #1, Space(15); Chr(27) + "M" + funTiraAcento(RecebeCampo, True)
    
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""

    XLO_RSTITU.Close
    Set XLO_RSTITU = Nothing
    Close #1
 
Exit Sub

TrataErro:
    funTrataErros (ComMensagem)
    Conexao.RollbackTrans
    Exit Sub

End Sub

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
    CmdPagamento.Enabled = True
    CmdRecalcular.Enabled = True
    CmdExclusaoMultipla.Enabled = True
    CmdAlteracaoMultipla.Enabled = True
    CmdPlanos.Enabled = True
    CmdSair.Enabled = True
    CmdImprimir.Enabled = True
    CmdOrdem.Enabled = True
    cmdPromissoria.Enabled = True
    cmdRepasse.Enabled = True
    cboLocalizar.Enabled = True
    
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
    CmdRecalcular.Enabled = False
    CmdPlanos.Enabled = False
    CmdExclusaoMultipla.Enabled = False
    CmdAlteracaoMultipla.Enabled = False
    CmdPagamento.Enabled = False
    CmdRemoverFiltro.Enabled = False
    CmdSair.Enabled = False
    CmdImprimir.Enabled = False
    CmdOrdem.Enabled = False
    cboLocalizar.Enabled = False
    cmdPromissoria.Enabled = False
    cmdRepasse.Enabled = False
    
End Sub

Private Sub cboLocalizar_Click(Area As Integer)
    If Area = 2 Then
      Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
End Sub

Private Sub cboLocalizar_KeyUp(KeyCode As Integer, Shift As Integer)
    If cboLocalizar.BoundText <> "" And Not IsNull(cboLocalizar.SelectedItem) Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
End Sub

Private Sub cboLocalizar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
End Sub

Private Sub chkCondomino_Click()
                
    Me.txtCondomino.Text = ""
    Me.txtCpfCnpj.Text = ""

    If Me.chkCondomino.Value = 1 Then
        Me.PanPromissorias.Height = 4665
        Me.cmdImprimeNotas.Top = 4080
        Me.CmdRetornanotas.Top = 4080
        Me.fraCondomino.Visible = True
    Else
        Me.fraCondomino.Visible = False
        Me.PanPromissorias.Height = 3000
        Me.cmdImprimeNotas.Top = 2400
        Me.CmdRetornanotas.Top = 2400
    End If
End Sub

Private Sub chkLaser_Click()
    If Me.chkLaser.Value = 1 Then
        Me.CboFonte.Enabled = False
        Me.chkClausula.Enabled = False
        Me.DTPData.Enabled = False
    Else
        Me.CboFonte.Enabled = True
        Me.chkClausula.Enabled = True
        Me.DTPData.Enabled = True
    End If
End Sub

Private Sub cmdExport_Click()
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
    
    TDBGrid1.ExportToDelimitedFile tmpFile, , ";"
 
    Close tmpFile
    
    'Now use this to add the Headers (to a new file)
    Open tmpFile For Input As #1
    Open outFile For Output As #2
 
    'Write Headers
    For i = 0 To TDBGrid1.Columns.Count - 1
         If i = 0 And TDBGrid1.Columns(i).Visible Then
             sLine = Chr(34) & TDBGrid1.Columns(i).Caption & Chr(34)
         ElseIf TDBGrid1.Columns(i).Visible Then
             sLine = sLine & ";" & Chr(34) & TDBGrid1.Columns(i).Caption & Chr(34)
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
    
    MsgBox "Total exportado: " & Adodc1.Recordset.RecordCount & " registros", vbInformation, "Informação"
    
End Sub

Private Sub CmdPagamento_Click()
    Dim XLT_CHAVETEMP As String

    If (Adodc1.Recordset.Fields("Status") = "INATIVO" Or _
      Adodc1.Recordset.Fields("Status") = "CEDIDO") And _
      IsNull(Adodc1.Recordset.Fields("titu_dt_Pagamento")) Then
        MsgBox "Não é possível efetuar pagamento com um Contrato Inativo ou Cedido.", vbCritical, "ATENÇÃO !"
        Exit Sub
    End If
    
    If Not FunVerifica_Permissao("FrmFinPagamentos", "1") Then
        Exit Sub
    End If
    
    Chave = Adodc1.Recordset.Fields("Titulo")
    XLT_CHAVETEMP = Adodc1.Recordset.Fields("Titulo") 'utilizada para garantir o retorno para a efetivação do pagto
    If Not IsNull(Adodc1.Recordset.Fields("titu_dt_pagamento")) Then
      XFormulario = "TelaFinTitulos"
    Else
      XFormulario = "TelaFinTitulosP" 'Efetivar pagamento
    End If
    
    FrmFinPagamentos.Show
    'FrmFinPagamentos.Show vbModal
    
    If XOrigem = "PAGOU" Or XOrigem = "CANCELOU" Then
      subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
      cboLocalizar.BoundText = XLT_CHAVETEMP
      If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
    
    XOrigem = ""
        
End Sub

Private Sub cmdImprimeNotas_Click()
    
    Dim XLO_RSCONFIG As New ADODB.Recordset
    Dim XLO_RSTITU As New ADODB.Recordset
    Dim XLS_SQL As String
    Dim XLS_Chave  As String
    Dim XLI_CONT As Long
    
    Dim XLO_IMPRESSORA As Printer
    Dim XLT_IMPRESSORA As String
      
    Dim XLI_ALTURA As Single 'Altura do papel
    Dim XLI_LARGURA As Single 'Largura da página
    
    Dim XLI_ESQUERDA As Single 'Margem mínima esquerda
    Dim XLI_TOPO As Single 'Margem mínima do topo
    
    Dim XLO_TAMANHOPAPEL As New ADODB.Recordset
  
    
    If (Me.chkCondomino.Value = 1 And Me.chkLaser.Value <> 1) Or (Me.chkImprimedataBase.Value = 1 And Me.chkLaser.Value <> 1) Then
        MsgBox "Essa opção só está habilitada para impressões em Jato de Tinta. Favor habilitar a opção.", vbInformation, "Informação"
        Exit Sub
    End If
    
    If Me.chkCondomino.Value = 1 Then
        If VerificaCpfCnpj = False Then
            MsgBox "Favor incluir o nome do condômino e o seu cnpj/cpf!", vbInformation, "Condôminos"
            Exit Sub
        End If
    End If
    
    'Imprime no relatório da promissória - Impressora Laser
    If Me.chkLaser.Value Then
        Call subImprimePromissoriaLaser
        Exit Sub
    Else
        If MsgBox("Você irá imprimir em uma impressora matricial, deseja prosseguir?", vbYesNo, "Nota Promissória") = vbNo Then
            Exit Sub
        End If
    End If
        
    'Nome da impressora padrão
    XLT_IMPRESSORA = Printer.DeviceName
    
    FrmImpressora.Show vbModal
    
    If CboFonte.Text = "" Then
      MsgBox "Escolha uma fonte para a impressão.", vbInformation, "ATENÇÃO"
      Exit Sub
    End If
  
    TelaFinTitulos.MousePointer = vbHourglass
    XLI_CONT = 0
    Printer.ScaleMode = 7  'vbCentimeters
    Printer.Orientation = 1  'Portrait
    
    XLI_ESQUERDA = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbCentimeters)
    XLI_TOPO = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbCentimeters)
    
    SubQOpenRecordset XLO_TAMANHOPAPEL, "SELECT form_nr_Altura,form_nr_Largura FROM formularios WHERE form_tx_NomeCampo='Promissoria' AND empr_cd_empresa=" & PCodEmpresa, Dinamico
    
    If Not IsNull(XLO_TAMANHOPAPEL!form_nr_altura) Then
      XLI_ALTURA = XLO_TAMANHOPAPEL!form_nr_altura
    Else
      XLI_ALTURA = 9.32
    End If
    
    If Not IsNull(XLO_TAMANHOPAPEL!form_nr_largura) Then
      XLI_LARGURA = XLO_TAMANHOPAPEL!form_nr_largura
    Else
      XLI_LARGURA = 20.45
    End If
    
    XLO_TAMANHOPAPEL.Close
    Set XLO_TAMANHOPAPEL = Nothing
    
    'Printer.PaperSize = vbPRPSUser
    XLI_ALTURA = Printer.ScaleY(XLI_ALTURA, vbCentimeters, vbTwips) 'Seta a altura
    XLI_LARGURA = Printer.ScaleX(XLI_LARGURA, vbCentimeters, vbTwips) 'Seta a largura
     
    Printer.Height = XLI_ALTURA
    Printer.Width = XLI_LARGURA
  
    Set XLO_RSCONFIG = Nothing
    
    'RS com as configuração do que deve ser impresso - o tipo de doc = Promissoria e a empresa
    XLS_SQL = "SELECT * FROM formularios WHERE " & _
              " form_nr_imprime = -1 and form_tx_TipoDocumento = 'PR'" & _
              " AND empr_cd_Empresa = " & PCodEmpresa & " AND form_tx_NomeCampo<>'Promissoria'"
    
    SubQOpenRecordset XLO_RSCONFIG, XLS_SQL, Estatico
    
    If OptSel.Value = True Then 'so imprimir a nota selecionada
        XLS_SQL = "SELECT * FROM consCAPFormularios WHERE titulo = '" & Chave & "'"
        SubQOpenRecordset XLO_RSTITU, XLS_SQL, Estatico
        subImprimePromissoria XLI_CONT, TxtObs.Text, chkClausula.Value, XLO_RSTITU, XLO_RSCONFIG, DTPData.Value, CboFonte.Text
    Else 'imprime tudo que esta filtrado no grid
        Adodc1.Recordset.MoveFirst
        While Not (Adodc1.Recordset.EOF)
            XLS_Chave = Adodc1.Recordset.Fields("titulo")
            XLS_SQL = "SELECT * FROM consCAPFormularios WHERE titulo = '" & XLS_Chave & "'"
            SubQOpenRecordset XLO_RSTITU, XLS_SQL, Estatico
            subImprimePromissoria XLI_CONT, TxtObs.Text, chkClausula.Value, XLO_RSTITU, XLO_RSCONFIG, DTPData.Value, CboFonte.Text
            Adodc1.Recordset.MoveNext
            Printer.NewPage
        Wend
    End If
    
    Printer.EndDoc
    XLO_RSCONFIG.Close
    XLO_RSTITU.Close
    Set XLO_RSCONFIG = Nothing
    Set XLO_RSTITU = Nothing
    
    'Retorna para a impressora padrão
    For Each XLO_IMPRESSORA In Printers
        If XLO_IMPRESSORA.DeviceName = XLT_IMPRESSORA Then
            Set Printer = XLO_IMPRESSORA
            Exit For
        End If
    Next
    
    TelaFinTitulos.MousePointer = vbDefault

End Sub

Private Sub CmdPlanos_Click()
  
    If Not FunVerifica_Permissao("FrmFinPlanos", "1") Then
          Exit Sub
    End If
    
    If Not Adodc1.Recordset.EOF Then
      Chave = Adodc1.Recordset.Fields("Titulo")
      subDesabilitaBotoes
    Else
      Chave = "-1"
    End If
    
    FrmFinPlanos.Show 1
    If Chave <> "-1" Then
      subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
      cboLocalizar.BoundText = Chave
      If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
    
    If Not Adodc1.Recordset.EOF Then
      subHabilitaBotoes
    End If
  
End Sub

Private Sub cmdPromissoria_Click()
      
    Dim XLI_CONT As Integer
    
    Me.MousePointer = vbHourglass
    
    If Not FunVerifica_Permissao(Me.Name, "6") Then
      Exit Sub
    End If
    
    PanBotoes.Enabled = False
    PanPromissorias.Left = (TDBGrid1.Width - PanPromissorias.Width) / 2
    PanPromissorias.Top = (TDBGrid1.Height - PanPromissorias.Height) / 2
    TDBGrid1.Enabled = False
    subDesabilitaBotoes
    Chave = Adodc1.Recordset.Fields("Titulo")
      
    If XFB_CARREGOUFONTE = False Then
      'Preenche o combo de fontes
      For XLI_CONT = 1 To Screen.FontCount
        CboFonte.AddItem Screen.Fonts(XLI_CONT)
      Next
      XFB_CARREGOUFONTE = True
    End If
    
    If Trim(PCliente) = "Costa Andrade" Or Trim(PCliente) = "RJ" Or Trim(PCliente) = "Actitur" Or Trim(PCliente) = "Franisa" Then
      Me.chkLaser.Enabled = True
    Else
      Me.chkCondomino.Enabled = False
      chkImprimedataBase.Enabled = False
    End If
    
    CboFonte.Text = "Arial"
    
    PanPromissorias.Visible = True
    PanPromissorias.Enabled = True
    PanPromissorias.Caption = "Emissão de Nota Promissória"
    Me.chkCondomino.Value = 0
    Me.chkImprimedataBase.Value = 0
    Me.chkLaser.Value = 0
    Me.txtCondomino.Text = ""
    Me.txtCpfCnpj.Text = ""
    OptSel.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub CmdRecalcular_Click()
  
    If Not FunVerifica_Permissao(Me.Name, "5") Then
          Exit Sub
    End If
    
    PanBotoes.Enabled = False
    PanRecalcular.Left = (TDBGrid1.Width - PanRecalcular.Width) / 2
    PanRecalcular.Top = (TDBGrid1.Height - PanRecalcular.Height) / 2
    TDBGrid1.Enabled = False
    subDesabilitaBotoes
    Chave = Adodc1.Recordset.Fields("Titulo")
    PanRecalcular.Visible = True
    PanRecalcular.Enabled = True
    MskCodigoInicial.Text = Chave
    MskCodigoFinal.Text = "    .    .  .  .   .  "
    PanRecalcular.Tag = "R"
    PanRecalcular.Caption = "Recalcular Títulos"
    MskCodigoInicial.SetFocus
  
End Sub

Private Sub CmdRecalcular_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Recalcula um conjunto de títulos da tabela"
End Sub

Private Sub CmdExclusaoMultipla_Click()
  
    If Not FunVerifica_Permissao(Me.Name, "4") Then
          Exit Sub
    End If
    
    PanBotoes.Enabled = False
    PanRecalcular.Left = (TDBGrid1.Width - PanRecalcular.Width) / 2
    PanRecalcular.Top = (TDBGrid1.Height - PanRecalcular.Height) / 2
    TDBGrid1.Enabled = False
    subDesabilitaBotoes
    Chave = Adodc1.Recordset.Fields("Titulo")
    PanRecalcular.Visible = True
    PanRecalcular.Enabled = True
    MskCodigoInicial.Text = Chave
    MskCodigoFinal.Text = "    .    .  .  .   .  "
    PanRecalcular.Tag = "E"
    PanRecalcular.Caption = "Exclusão Múltipla"
    MskCodigoInicial.SetFocus
  
End Sub

Private Sub CmdExclusaoMultipla_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Exclui um conjunto de títulos da tabela"
End Sub

Private Sub cmdContabilizar_Click()
    Chave = Adodc1.Recordset.Fields("Titulo")
End Sub

Private Sub cmdRepasse_Click()
     
    If Not FunVerifica_Permissao(Me.Name, "7") Then
          Exit Sub
    End If
         
    PanRepasse.Left = (TDBGrid1.Width - PanRepasse.Width) / 2
    PanRepasse.Top = (TDBGrid1.Height - PanRepasse.Height) / 2
    
    dtpRepasse.Value = Date
    subDesabilitaBotoes
    PanRepasse.Visible = True
   
End Sub

Private Sub cmdRepConfirma_Click()
    Dim XLT_SQL As String
    Dim XLO_RSTITULO As New ADODB.Recordset
    Dim XLI_BOOKMARK As Integer
    
    On Error GoTo TrataErro
        
    XLT_SQL = Adodc1.Recordset.Source
    SubQOpenRecordset XLO_RSTITULO, XLT_SQL, Dinamico
    
    Chave = Adodc1.Recordset.Fields("Titulo")
    
    If (OptFiltrados.Value = True) Then
        Adodc1.Recordset.MoveFirst
        
        XLO_RSTITULO.MoveFirst
        
        Conexao.BeginTrans
        
        While Not (XLO_RSTITULO.EOF)
           If Not IsNull(XLO_RSTITULO("titu_dt_Pagamento")) And IsNull(XLO_RSTITULO("titu_dt_Repasse")) Then
             If Not (IsNull(dtpRepasse.Value)) Then
               XLO_RSTITULO("titu_dt_repasse") = dtpRepasse.Value
               XLO_RSTITULO.Update
               
               '******************REGISTRA LOG DA OPERAÇÃO**************************
               ReDim XGM_MATRIZLOG(2, 2) As Variant
               
               XGM_MATRIZLOG(0, 0) = "Título"
               XGM_MATRIZLOG(1, 0) = "Data Repasse"
                                               
               XGM_MATRIZLOG(0, 1) = XLO_RSTITULO("Titulo")
               XGM_MATRIZLOG(1, 1) = dtpRepasse.Value
                               
               Call subRegistraLog("TelaFinTitulos", "7", funCriaDescricaoLog(XGM_MATRIZLOG, REPASSE_TITULO, TITULO_CAPI))
               '*******************************************************************
             Else
               MsgBox "Selecione a data do repasse", vbCritical, "ATENÇÃO"
               Conexao.RollbackTrans
               Exit Sub
             End If
           End If
           XLO_RSTITULO.MoveNext
        Wend
      
        Conexao.CommitTrans
        
        PanRepasse.Visible = False
        subHabilitaBotoes
        XLO_RSTITULO.Close
        Set XLO_RSTITULO = Nothing
    Else ' mudar a data apenas do filtrado
        Conexao.BeginTrans
        
        If Not IsNull(Adodc1.Recordset.Fields("titu_dt_pagamento")) Then
            If Not (IsNull(dtpRepasse.Value)) Then
                XLO_RSTITULO.Filter = "Titulo = '" & Adodc1.Recordset.Fields("titulo") & " '"
                XLO_RSTITULO("titu_dt_repasse") = dtpRepasse.Value
                XLO_RSTITULO.Update
            
                '******************REGISTRA LOG DA OPERAÇÃO**************************
                ReDim XGM_MATRIZLOG(2, 2) As Variant
                
                XGM_MATRIZLOG(0, 0) = "Título"
                XGM_MATRIZLOG(1, 0) = "Data Repasse"
                                                
                XGM_MATRIZLOG(0, 1) = XLO_RSTITULO("Titulo")
                XGM_MATRIZLOG(1, 1) = dtpRepasse.Value
                                
                Call subRegistraLog("TelaFinTitulos", "7", funCriaDescricaoLog(XGM_MATRIZLOG, REPASSE_TITULO, TITULO_CAPI))
                '*******************************************************************
            Else
                MsgBox "Selecione a data do repasse", vbCritical, "ATENÇÃO"
                Conexao.RollbackTrans
                Exit Sub
            End If
        Else
            MsgBox "Apenas Títulos pagos podem ser repassados", vbCritical, "ATENÇÃO"
        End If
        
        Conexao.CommitTrans
    End If
    
    'Atualiza o Grid
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    
    cboLocalizar.BoundText = Chave
    If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    
    cmdRepExit_Click
    Exit Sub
    
TrataErro:
    Conexao.RollbackTrans
    funTrataErros (ComMensagem)
End Sub


Private Sub CmdRepDesmarcar_Click()
    Dim XLT_SQL As String
    Dim XLO_RSTITULO As New ADODB.Recordset
    Dim XLI_BOOKMARK As Integer
    
    'Retira a data do repasse dos titulos filtrados.
    On Error GoTo TrataErro
    
    Chave = Adodc1.Recordset.Fields("Titulo")
    
    If (OptFiltrados.Value = True) Then
        Adodc1.Recordset.MoveFirst
              
        XLT_SQL = "SELECT titu_dt_Pagamento,titu_dt_Repasse FROM Titulos"
        
        SubQOpenRecordset XLO_RSTITULO, XLT_SQL, Dinamico
        
        XLO_RSTITULO.MoveFirst
        
        Conexao.BeginTrans
        While Not (XLO_RSTITULO.EOF)
            If Not IsNull(XLO_RSTITULO("titu_dt_pagamento")) Then
                'se o check estiver desmarcado entao todos os filtrados sao desmarcados
                If (IsNull(dtpRepasse.Value)) Then
                    XLO_RSTITULO("titu_dt_repasse") = Null
                    XLO_RSTITULO.Update
                    
                    '******************REGISTRA LOG DA OPERAÇÃO**************************
                    ReDim XGM_MATRIZLOG(1, 2) As Variant
                    
                    XGM_MATRIZLOG(0, 0) = "Título"
                    XGM_MATRIZLOG(0, 1) = XLO_RSTITULO("Titulo")
                                    
                    Call subRegistraLog("TelaFinTitulos", "7", funCriaDescricaoLog(XGM_MATRIZLOG, CANCELAR_REPASSE_TITULO, TITULO_CAPI))
                    '*******************************************************************
                Else
                    'se o check estiver marcado so seram desmarcados os titulos com aquela data de repasse
                    If dtpRepasse.Value = XLO_RSTITULO("titu_dt_repasse") Then
                        XLO_RSTITULO("titu_dt_repasse") = Null
                        XLO_RSTITULO.Update
                        
                        '******************REGISTRA LOG DA OPERAÇÃO**************************
                        ReDim XGM_MATRIZLOG(1, 2) As Variant
                        
                        XGM_MATRIZLOG(0, 0) = "Título"
                        XGM_MATRIZLOG(0, 1) = XLO_RSTITULO("Titulo")
                                        
                        Call subRegistraLog("TelaFinTitulos", "7", funCriaDescricaoLog(XGM_MATRIZLOG, CANCELAR_REPASSE_TITULO, TITULO_CAPI))
                        '*******************************************************************
                    End If
                End If
            End If
            XLO_RSTITULO.MoveNext
        Wend
        Conexao.CommitTrans
        
        PanRepasse.Visible = False
        subHabilitaBotoes
        XLO_RSTITULO.Close
        Set XLO_RSTITULO = Nothing
    
    Else ' mudar a data apenas do filtrado nao iporta a data do repasse
      If Not IsNull(Adodc1.Recordset.Fields("titu_dt_pagamento")) Then
        
        XLT_SQL = "UPDATE Titulos SET titu_dt_Repasse=null " & _
          " WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & _
          " empd_cd_empreendimento='" & Left$(Chave, 4) & "' AND " & _
          " imov_cd_imovel='" & Mid$(Chave, 6, 4) & "' AND " & _
          " cont_cd_Contrato='" & Mid$(Chave, 11, 2) & "' AND " & _
          " titu_cd_Plano= '" & Mid$(Chave, 14, 2) & "' AND " & _
          " titu_cd_Parcela= '" & Mid$(Chave, 17, 3) & "' AND " & _
          " titu_cd_Residuo= '" & Right$(Chave, 2) & "'"
        Conexao.Execute (XLT_SQL)
        
        '******************REGISTRA LOG DA OPERAÇÃO**************************
        ReDim XGM_MATRIZLOG(1, 2) As Variant
        
        XGM_MATRIZLOG(0, 0) = "Título"
        XGM_MATRIZLOG(0, 1) = Chave
                        
        Call subRegistraLog("TelaFinTitulos", "7", funCriaDescricaoLog(XGM_MATRIZLOG, CANCELAR_REPASSE_TITULO, TITULO_CAPI))
        '*******************************************************************
        
        XLI_BOOKMARK = Adodc1.Recordset.bookmark
      Else
        MsgBox "Apenas Títulos pagos podem ser repassados", vbCritical, "ATENÇÃO"
      End If
    End If
    
    'Atualiza o Grid
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    
    cboLocalizar.BoundText = Chave
    If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    cmdRepExit_Click
    
    Exit Sub

TrataErro:
    Conexao.RollbackTrans
    funTrataErros (ComMensagem)
End Sub

Private Sub cmdRepExit_Click()
    subHabilitaBotoes
    PanRepasse.Visible = False
End Sub

Private Sub CmdRetornanotas_Click()
    PanPromissorias.Visible = False
    PanPromissorias.Enabled = False
    PanBotoes.Enabled = True
    subHabilitaBotoes
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub

Private Sub Command2_Click()

    Adodc1.Recordset.MoveFirst
    While Not (Adodc1.Recordset.EOF)
       
    Wend

End Sub

Private Sub Command50_Click()
    
    Dim a As Object
    Dim XLO_RSCONFIG As New ADODB.Recordset
    Dim XLO_RSTITU As New ADODB.Recordset
    Dim XLS_SQL As String
    Dim XLS_Chave  As String
    Dim XLI_CONT As Long
    Dim XLI_ALTURA As Single 'Altura do papel
    Dim XLI_LARGURA As Single 'Largura da página
    Dim XLB_Negrito As Boolean
    Dim XLS_Campo As String
    Dim Cont As Integer
    Dim XLI_Clausula As Integer
    Dim XLI_ESQUERDA As Single 'Margem mínima esquerda
    Dim XLI_TOPO As Single 'Margem mínima do topo
    Dim XLT_FONT As String
     Dim ReportTitle As CRAXDRT.TextObject
    Dim XLO_TAMANHOPAPEL As New ADODB.Recordset

    
    If CboFonte.Text = "" Then
      MsgBox "Escolha uma fonte para a impressão.", vbInformation, "ATENÇÃO"
      Exit Sub
    End If
    
    TelaFinTitulos.MousePointer = vbHourglass
    
    
    'RS com as configuração do que deve ser impresso - o tipo de doc = Promissoria e a empresa
    XLS_SQL = "SELECT * FROM formularios WHERE " & _
              " form_nr_imprime = -1 and form_tx_TipoDocumento = 'PR'" & _
              " AND empr_cd_Empresa = " & PCodEmpresa & " AND form_tx_NomeCampo <> 'Promissoria'"
    SubQOpenRecordset XLO_RSCONFIG, XLS_SQL, Estatico
    
    XLS_SQL = "SELECT * FROM consCAPFormularios WHERE titulo = '" & Chave & "'"
    SubQOpenRecordset XLO_RSTITU, XLS_SQL, Estatico
    
    'Criando novo relatório
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path & "\RELATORIOS\notapromissoria.rpt")
            
    XGR_RELATORIO.TopMargin = 0.7
    XGR_RELATORIO.LeftMargin = 0.5
    XGR_RELATORIO.RightMargin = 1
    XGR_RELATORIO.BottomMargin = 1
    
   'Entra num loop para configurar cada campo do relatório
    While Not (XLO_RSCONFIG.EOF)
        
      'Testando se negrito está marcado para associar ao campo
      If Not (XLO_RSCONFIG!form_nr_negrito = -1) Then
        XLB_Negrito = False
      Else
        XLB_Negrito = True
      End If

      
      XLS_Campo = XLO_RSCONFIG!form_tx_nomecampo
      
      Select Case (XLS_Campo)
          'campos que nao estao na consulta ( datas e valores por extenso )
          Case "ValorTitulo"
          
            If (XLO_RSCONFIG!form_tx_Descricao = "Valor do Título") Then
              
                'Adicionando o valor no relatório
                Set a = XGR_RELATORIO.Sections(3).ReportObjects(1)
                
                a.Top = 6000
                a.Left = 10000
                
                ' Criar FunConfiguraCampo para tratar isso
                a.Font.Bold = XLB_Negrito
                a.Font.Size = XLO_RSCONFIG!form_nr_tamfonte
                a.Font.Name = Me.CboFonte.Text
            
            Else
                Set a = XGR_RELATORIO.Sections(3).ReportObjects(1)
                    
                ' Criar FunConfiguraCampo para tratar isso
                a.Font.Bold = XLB_Negrito
                a.Font.Size = XLO_RSCONFIG!form_nr_tamfonte
                a.Font.Name = Me.CboFonte.Text
            End If
          End Select
             XLO_RSCONFIG.MoveNext

    Wend
    
    XLO_RSTITU.Close
    XLO_RSCONFIG.Close
    Set XLO_RSTITU = Nothing
    Set XLO_RSCONFIG = Nothing
    
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    
    TelaFinTitulos.MousePointer = vbDefault
    
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
        MskCodigoInicial = UCase(MskCodigoInicial.Text)
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

Private Sub CmdConfirmar_Click()
  
    On Error GoTo TrataErro:
    
    Dim XLO_TITULO As New ADODB.Recordset
    Dim XLO_INDEXACAO As New ADODB.Recordset
    Dim XLT_SQL As String
    Dim XFT_CODIMOVEL As String, XFT_CODEMPREENDIMENTO As String
    Dim XFT_CODCONTRATO As String, XFT_CODPLANO As String
    Dim XFT_CODPARCELA As String, XFT_CODRESIDUO As String
    Dim XLF_VALORINDEXADO1 As Double 'Armazena o valor indexado 1
    Dim XLF_VALORINDEXADO2 As Double 'Armazena o valor indexado 2
    Dim XLT_PRORATA As String 'Armazena se a correção é ProRata
    Dim XLT_TITULOQUALIFIN As String  'Unifica o código do título para o Qualifin
    Dim XLO_QUALIFIN As ADODB.Recordset  'Recordset para interligar ao qualifin
    Dim XLL_FATURA As Long            'Código da fatura no Qualifin
  
    If FunObrigatorioMSK(MskCodigoInicial, "O Código Inicial é Obrigatório") Then Exit Sub
    If FunObrigatorioMSK(MskCodigoFinal, "O Código Final é Obrigatório") Then Exit Sub
    
    If PanRecalcular.Tag = "R" Then
      If MsgBox("Confirma o Recalculo do Intervalo de Registros ?" & vbCrLf & "(Títulos Pagos, do tipo Resíduo e os que tiveram os valores alterados" & vbCrLf & " para a entrega das chaves sem juros não serão recalculados)", 273, "CUIDADO") <> 1 Then
        Exit Sub
      End If
    Else
       If MsgBox("Confirma Remoção do Intervalo de Registros ?" & vbCrLf & "(Títulos Pagos e do tipo Resíduo não serão recalculados)", 273, "CUIDADO") <> 1 Then
         Exit Sub
       End If
    End If
  
    SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCAPRecalcular WHERE tipl_tx_Sigla <> 'R' " & _
      "AND (titu_tx_EntregaChave <> '1' OR titu_tx_EntregaChave is null)" & _
      "AND Titulo >= '" & MskCodigoInicial.Text & "' AND Titulo <= '" & MskCodigoFinal.Text & "' AND empr_cd_empresa = " & PCodEmpresa & " ORDER BY Titulo", Estatico
    
    
    While Not XLO_TITULO.EOF
    
        XFT_CODEMPREENDIMENTO = Left$(XLO_TITULO!titulo, 4)
        XFT_CODIMOVEL = Mid$(XLO_TITULO!titulo, 6, 4)
        XFT_CODCONTRATO = Mid$(XLO_TITULO!titulo, 11, 2)
        XFT_CODPLANO = Mid$(XLO_TITULO!titulo, 14, 2)
        XFT_CODPARCELA = Mid$(XLO_TITULO!titulo, 17, 3)
        XFT_CODRESIDUO = Right$(XLO_TITULO!titulo, 2)
        
        XLT_PRORATA = FunNulo(XLO_TITULO!cont_tx_CorrecProRata)
        
        If PanRecalcular.Tag = "R" Then
            If Not IsNull(XLO_TITULO!moed_cd_Moeda1) Then
              XLF_VALORINDEXADO1 = FunCalcularValorIndexado(XLO_TITULO!titu_vl_Parcela, XLO_TITULO!titu_dt_Base, XLO_TITULO!moed_cd_Moeda1, XLO_TITULO!Indexacao1, XLT_PRORATA, TelaCotacoes)
            Else
              XLF_VALORINDEXADO1 = 0
            End If
            
            If Not IsNull(XLO_TITULO!moed_cd_Moeda2) Then
              XLF_VALORINDEXADO2 = FunCalcularValorIndexado(XLO_TITULO!titu_vl_Parcela, XLO_TITULO!titu_dt_Base, XLO_TITULO!moed_cd_Moeda2, XLO_TITULO!Indexacao2, XLT_PRORATA, TelaCotacoes)
            Else
              XLF_VALORINDEXADO2 = 0
            End If
            
            Conexao.BeginTrans
            
            'Se for Recalcular
            XLT_SQL = "UPDATE Titulos SET titu_vl_ParcelaIndex1=" & FunNuloVal(FunTrataFloat(XLF_VALORINDEXADO1)) & "," & _
                  "titu_vl_ParcelaIndex2=" & FunNuloVal(FunTrataFloat(XLF_VALORINDEXADO2)) & "" & _
                  " WHERE  empr_cd_Empresa=" & PCodEmpresa & " AND " & _
                  "empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND " & _
                  "imov_cd_imovel='" & XFT_CODIMOVEL & "' AND " & _
                  "cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND " & _
                  "titu_cd_Plano= '" & XFT_CODPLANO & "' AND " & _
                  "titu_cd_Parcela= '" & XFT_CODPARCELA & "' AND " & _
                  "titu_cd_Residuo= '" & XFT_CODRESIDUO & "'"
                  
            '******************REGISTRA LOG DA OPERAÇÃO**************************
            ReDim XGM_MATRIZLOG(3, 2) As Variant
            
            XGM_MATRIZLOG(0, 0) = "Título"
            XGM_MATRIZLOG(1, 0) = "Parcela Index1"
            XGM_MATRIZLOG(2, 0) = "Parcela Index2"
                            
            XGM_MATRIZLOG(0, 1) = XFT_CODEMPREENDIMENTO & "." & XFT_CODIMOVEL & "." & XFT_CODCONTRATO & "." & _
                                  XFT_CODPLANO & "." & XFT_CODPARCELA & "." & XFT_CODRESIDUO
            XGM_MATRIZLOG(1, 1) = FunNuloVal(FunTrataFloat(XLF_VALORINDEXADO1))
            XGM_MATRIZLOG(2, 1) = FunNuloVal(FunTrataFloat(XLF_VALORINDEXADO2))
            
            Call subRegistraLog("TelaFinTitulos", "5", funCriaDescricaoLog(XGM_MATRIZLOG, RECALCULAR, TITULO_CAPI))
            '*******************************************************************
        Else
          Conexao.BeginTrans
          
          'Se for excluir
          XLT_SQL = "DELETE FROM Titulos WHERE  empr_cd_Empresa=" & PCodEmpresa & " AND " & _
                  "empd_cd_empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND " & _
                  "imov_cd_imovel='" & XFT_CODIMOVEL & "' AND " & _
                  "cont_cd_Contrato='" & XFT_CODCONTRATO & "' AND " & _
                  "titu_cd_Plano= '" & XFT_CODPLANO & "' AND " & _
                  "titu_cd_Parcela= '" & XFT_CODPARCELA & "' AND " & _
                  "titu_cd_Residuo= '" & XFT_CODRESIDUO & "'"
           FRM_Quantid = funOperacaoQuantid("-")
                          
           '************ATUALIZAÇÃO DO QUALIFIN****************
           If XGB_PQUALIFIN Then
              XLT_TITULOQUALIFIN = Left$(XLO_TITULO!titulo, 4) + "." + Mid$(XLO_TITULO!titulo, 6, 4) + "." + Mid$(XLO_TITULO!titulo, 11, 2) + "." + Mid$(XLO_TITULO!titulo, 14, 2) + "." + Mid$(XLO_TITULO!titulo, 17, 3) + "." + Right$(XLO_TITULO!titulo, 2)
              
              SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM NotasFiscais " & _
                  "WHERE nofi_nr_TituloCapi LIKE '" & XLT_TITULOQUALIFIN & "' AND empr_cd_empresa = " & PCodEmpresa & "", Dinamico
              If Not XLO_QUALIFIN.EOF Then
                  XLL_FATURA = XLO_QUALIFIN!nofi_nr_fatura
                      
                  Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_NotaFiscal = " & XLO_QUALIFIN!nofi_cd_NotaFiscal & ""
                  XLO_QUALIFIN.Close
                  Set XLO_QUALIFIN = Nothing
                      
                  SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Duplicatas WHERE dupl_nr_fatura = " & XLL_FATURA, Dinamico
                  Conexao.Execute "DELETE FROM DuplDescAcresc WHERE dupl_cd_duplicata= " & XLO_QUALIFIN!dupl_cd_Duplicata
                  XLO_QUALIFIN.Close
                  Set XLO_QUALIFIN = Nothing
                  
                  Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura= " & XLL_FATURA
                  Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_nr_fatura= " & XLL_FATURA
              Else
                  XLO_QUALIFIN.Close
                  Set XLO_QUALIFIN = Nothing
              End If
           End If
           '***************************************************
           
           '******************REGISTRA LOG DA OPERAÇÃO**************************
            ReDim XGM_MATRIZLOG(1, 2) As Variant
            
            XGM_MATRIZLOG(0, 0) = "Título"
            XGM_MATRIZLOG(0, 1) = XFT_CODEMPREENDIMENTO & "." & XFT_CODIMOVEL & "." & XFT_CODCONTRATO & "." & _
                                  XFT_CODPLANO & "." & XFT_CODPARCELA & "." & XFT_CODRESIDUO
                        
            Call subRegistraLog("TelaFinTitulos", "4", funCriaDescricaoLog(XGM_MATRIZLOG, EXCLUSAO_MULTIPLA, TITULO_CAPI))
            '*******************************************************************
        End If
        Conexao.Execute (XLT_SQL)
        Conexao.CommitTrans
        XLO_TITULO.MoveNext
    Wend
  
    XLO_TITULO.Close
    Set XLO_TITULO = Nothing
    
    PanRecalcular.Visible = False
    PanRecalcular.Enabled = False
    PanBotoes.Enabled = True
    subHabilitaBotoes
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    
    If FunTabelaVazia(Adodc1) = True Then
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdRecalcular.Enabled = False
        CmdExclusaoMultipla.Enabled = False
        CmdAlteracaoMultipla.Enabled = False
        CmdImprimir.Enabled = False
        CmdOrdem.Enabled = False
        CmdPagamento.Enabled = False
        cmdRepasse.Enabled = False
        cmdPromissoria.Enabled = False
    End If
      
    If PanRecalcular.Tag = "R" Then
        cboLocalizar.BoundText = Chave
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    Else
        cboLocalizar.Text = ""
    End If
    
    Exit Sub
TrataErro:
    Conexao.RollbackTrans
    funTrataErros (ComMensagem)
      
End Sub

Private Sub CmdRetornar_Click()
    PanRecalcular.Visible = False
    PanRecalcular.Enabled = False
    PanBotoes.Enabled = True
    subHabilitaBotoes
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub

Private Sub CmdAlteracaoMultipla_Click()
  
    If Not FunVerifica_Permissao(Me.Name, "3") Then
        Exit Sub
    End If
    
    Chave = Adodc1.Recordset.Fields("Titulo")
    xChave = Adodc1.Recordset.bookmark
    subDesabilitaBotoes
    FrmFinAlterarTitulos.Show 1
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    cboLocalizar.BoundText = Chave
    If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    subHabilitaBotoes
  
End Sub

Private Sub CmdAlteracaoMultipla_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     MdiPrincipal.BarraStatus.Panels(1).Text = "Altera um conjunto de títulos da tabela"
End Sub

Private Sub CmdImprimir_Click()
    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Títulos")
End Sub

Private Sub CmdRemoverFiltro_Click()
    Dim XLI_FIL As Integer
    
    subDesabilitaBotoes

    Set Formulario = TelaFinTitulos
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"
    
    If Filtrou = True Then
        XLI_FIL = ContFil
        subContaReg Adodc1
        'Adodc1.Refresh
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
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
    
    If Not FunVerifica_Permissao(Me.Name, "3") Then
        Exit Sub
    End If
    
    Chave = Adodc1.Recordset.Fields("Titulo")
    If IsNull(Adodc1.Recordset.Fields("titu_dt_Pagamento")) Then
      XInserir = "A"
      xChave = Adodc1.Recordset.bookmark
      subDesabilitaBotoes
      FrmFinTitulos.Show 1
      subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
      cboLocalizar.BoundText = Chave
      If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    Else
      XInserir = "C"
      subDesabilitaBotoes
      MsgBox "Esse Título está pago e não poderá ser alterado.", vbCritical, "ATENÇÃO !"
      FrmFinTitulos.Show 1
    End If
    subHabilitaBotoes
    
End Sub

Private Sub CmdAlterar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  MdiPrincipal.BarraStatus.Panels(1).Text = "Altera um registro na tabela"
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
    Set Formulario = TelaFinTitulos
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
    
    On Error GoTo RotuloErro
    
    Dim XLI_OPCAO As Integer
    Dim XLI_HOUVEERRO As Integer
    Dim XLI_BOOKMARK As Integer
    
    Dim XLL_FATURA As Long              'Armazena o nº da fatura para deletar as tabelas no qualifin
    Dim XLO_QUALIFIN As ADODB.Recordset  'Recordset para abrir as tabelas do Qualifin
    Dim XLT_TITULOQUALIFIN As String    'Compor o título
    
    If Not FunVerifica_Permissao(Me.Name, "4") Then
        Exit Sub
    End If
    
    XLI_HOUVEERRO = False
    
    subDesabilitaBotoes
   
    If IsNull(Adodc1.Recordset.Fields("titu_dt_Pagamento")) Then
     
        XLI_OPCAO = MsgBox("Confirma Remoção do Registro ?", 273, "CUIDADO")
        
        If XLI_OPCAO = 1 Then  'Confirma exclusão
            On Error GoTo RotuloErro
            
            Conexao.BeginTrans
          
            Conexao.Execute "DELETE FROM Titulos WHERE  empr_cd_Empresa=" & PCodEmpresa & " AND " & _
                          "empd_cd_empreendimento='" & Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "' AND " & _
                          "imov_cd_imovel='" & Adodc1.Recordset.Fields("imov_cd_Imovel") & "' AND " & _
                          "cont_cd_Contrato='" & Adodc1.Recordset.Fields("cont_cd_Contrato") & "' AND " & _
                          "titu_cd_Plano= '" & Adodc1.Recordset.Fields("titu_cd_Plano") & "' AND " & _
                          "titu_cd_Parcela= '" & Adodc1.Recordset.Fields("titu_cd_Parcela") & "' AND " & _
                          "titu_cd_Residuo= '" & Adodc1.Recordset.Fields("titu_cd_Residuo") & "'", , adCmdText
        
            '************ATUALIZAÇÃO DO QUALIFIN****************
            If XGB_PQUALIFIN Then
              XLT_TITULOQUALIFIN = Adodc1.Recordset.Fields("empd_cd_Empreendimento") + "." + Adodc1.Recordset.Fields("imov_cd_Imovel") + "." + Adodc1.Recordset.Fields("cont_cd_Contrato") + "." + Adodc1.Recordset.Fields("titu_cd_Plano") + "." + Adodc1.Recordset.Fields("titu_cd_Parcela") + "." + Adodc1.Recordset.Fields("titu_cd_Residuo")
              
              SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM NotasFiscais " & _
                  "WHERE nofi_nr_TituloCapi LIKE '" & XLT_TITULOQUALIFIN & "' AND empr_cd_empresa = " & PCodEmpresa & "", Dinamico
              If Not XLO_QUALIFIN.EOF Then
                  XLL_FATURA = XLO_QUALIFIN!nofi_nr_fatura
                      
                  Conexao.Execute "DELETE FROM Distribuicao WHERE nofi_cd_NotaFiscal = " & XLO_QUALIFIN!nofi_cd_NotaFiscal & ""
                  XLO_QUALIFIN.Close
                  Set XLO_QUALIFIN = Nothing
                      
                  SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Duplicatas WHERE dupl_nr_fatura = " & XLL_FATURA, Dinamico
                  Conexao.Execute "DELETE FROM DuplDescAcresc WHERE dupl_cd_duplicata= " & XLO_QUALIFIN!dupl_cd_Duplicata
                  XLO_QUALIFIN.Close
                  Set XLO_QUALIFIN = Nothing
                  
                  Conexao.Execute "DELETE FROM Duplicatas WHERE dupl_nr_fatura= " & XLL_FATURA
                  Conexao.Execute "DELETE FROM NotasFiscais WHERE nofi_nr_fatura= " & XLL_FATURA
              Else
                  XLO_QUALIFIN.Close
                  Set XLO_QUALIFIN = Nothing
              End If
            End If
            '***************************************************
            
            Conexao.CommitTrans
            
            
            '******************REGISTRA LOG DA OPERAÇÃO**************************
            ReDim XGM_MATRIZLOG(1, 2) As Variant
            
            XGM_MATRIZLOG(0, 0) = "Título"
                            
            XGM_MATRIZLOG(0, 1) = Adodc1.Recordset.Fields("empd_cd_Empreendimento") & "." & _
                Adodc1.Recordset.Fields("imov_cd_Imovel") & "." & _
                Adodc1.Recordset.Fields("cont_cd_Contrato") & "." & _
                Adodc1.Recordset.Fields("titu_cd_Plano") & "." & _
                Adodc1.Recordset.Fields("titu_cd_Parcela") & "." & _
                Adodc1.Recordset.Fields("titu_cd_Residuo")
            
            Call subRegistraLog("TelaFinTitulos", "4", funCriaDescricaoLog(XGM_MATRIZLOG, EXCLUIR, TITULO_CAPI))
            '*******************************************************************
          
          
            If XLI_HOUVEERRO = False Then
                 XLI_BOOKMARK = TDBGrid1.bookmark
                 FRM_Quantid = funOperacaoQuantid("-")
                 subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
                 
                If Not Adodc1.Recordset.EOF Then
                    If XLI_BOOKMARK <= Adodc1.Recordset.RecordCount Then
                      TDBGrid1.bookmark = XLI_BOOKMARK
                    Else
                      TDBGrid1.bookmark = XLI_BOOKMARK - 1
                    End If
                End If
            End If
        End If
      
        subHabilitaBotoes
        
        If FunTabelaVazia(Adodc1) = True Then
            CmdAlterar.Enabled = False
            CmdExcluir.Enabled = False
            CmdRecalcular.Enabled = False
            CmdExclusaoMultipla.Enabled = False
            CmdAlteracaoMultipla.Enabled = False
            CmdPlanos.Enabled = False
            CmdPagamento.Enabled = False
            CmdImprimir.Enabled = False
            CmdOrdem.Enabled = False
            CmdPlanos.Enabled = False
            cmdRepasse.Enabled = False
            cmdPromissoria.Enabled = False
        End If
      
      cboLocalizar.Text = ""
      TDBGrid1.SetFocus
    Else
       MsgBox "Esse Título está pago e não poderá ser excluído.", vbCritical, "ATENÇÃO !"
       subHabilitaBotoes
    End If
    
    Exit Sub
    
RotuloErro:
    Conexao.RollbackTrans

    XLI_HOUVEERRO = True
    funTrataErros (ComMensagem)
    Resume Next
End Sub
Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If CmdInserir.Caption = "&Gravar" Then
       MdiPrincipal.BarraStatus.Panels(1).Text = "Cancela as novas informações"
    Else
       MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
    End If
End Sub

Private Sub CmdInserir_Click()
      
     If Not FunVerifica_Permissao(Me.Name, "2") Then
          Exit Sub
     End If
     
     If FunTabelaVazia(Adodc1) Then
       MsgBox "É Preciso Criar um Plano Primeiro.", vbCritical, "ATENÇÃO !"
     Else
      Chave = Adodc1.Recordset.Fields("Titulo")
      XInserir = "I"
      cboLocalizar.Text = ""
      subDesabilitaBotoes
      FrmFinTitulos.Show 1
      If Chave <> -1 Then
          FRM_Quantid = funOperacaoQuantid("+")
          subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
          cboLocalizar.BoundText = Chave
          Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
      Else
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
      End If
      subHabilitaBotoes
    End If
  
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

Private Sub MskTitulo_GotFocus()
  
    subSelecionaMSK MskTitulo
  
End Sub

Private Sub MskTitulo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
    
    Dim XLT_TEMP As String
    Dim XLI_POS As Integer
    
    If Adodc1.Recordset.RecordCount > 0 Then
    
      Adodc1.Recordset.MoveFirst
        
      If MskTitulo.Text <> "    .    .  .  .   .  " Then
      
        XLI_POS = 1
        
        XLI_POS = InStr(XLI_POS, MskTitulo.Text, " ")
        
        If XLI_POS <> 0 Then
          XLT_TEMP = Left(MskTitulo.Text, XLI_POS - 1)
        Else
          XLT_TEMP = MskTitulo.Text
        End If
        
        Adodc1.Recordset.Find ("Titulo LIKE '" & XLT_TEMP & "%'")
        
        If IsNull(TDBGrid1.bookmark) Then
        
          MsgBox "Título não encontrado", vbInformation, "ATENÇÃO"
          MskTitulo.SetFocus
          
        Else
        
          MskTitulo.Text = "    .    .  .  .   .  "
          TDBGrid1.SetFocus
        
        End If
      End If
    End If
  End If
  
End Sub

Private Sub OptSel_Click()
    If Me.Adodc1.Recordset.EOF = True Then
        Me.Adodc1.Recordset.MoveFirst
    End If
End Sub

Private Sub TDBGrid1_DblClick()
    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    
    If CDate(TDBGrid1.Columns("Vencimento").CellText(bookmark)) < Date _
    And TDBGrid1.Columns("Pagamento").CellText(bookmark) = "" Then
      RowStyle.ForeColor = vbRed
    Else
      RowStyle.ForeColor = vbBlack
    End If
    
    If (TDBGrid1.Columns("Natureza do Plano").CellText(bookmark) = "FINANCIAMENTO") Then
        RowStyle.Font.Bold = True
    End If
    
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    cboLocalizar.Text = ""
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaFinTitulos
        subTelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, x
        subTelaValoresGlobais "G"
    End If
    
    subHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) = True Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdRecalcular.Enabled = False
       CmdExclusaoMultipla.Enabled = False
       CmdAlteracaoMultipla.Enabled = False
       CmdPagamento.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
       cmdPromissoria.Enabled = False
       cmdRepasse.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    Set Formulario = TelaFinTitulos
    If Not PrimeiraVez Then
        subTelaValoresGlobais "P"
    End If
    
    PrimeiraVez = False

End Sub

Private Sub Form_Load()
    
    Dim XLT_CODEMPREENDIMENTO As String
    Dim XLT_CODIMOVEL As String 'Código do imóvel
    Dim XLT_CODCONTRATO As String 'Código do Contrato
    Dim XLO_TITULOS As New ADODB.Recordset
        
    DTPData.Value = Date
    
    'Indica que ainda não carregou a fonte das promissórias
    XFB_CARREGOUFONTE = False
    
    Set Formulario = TelaFinTitulos
    
    Call SubAjustaTela
    
    subManutencaoJanelasAtivas "I", "TelaFinTitulos"
               
    If XOrigem = "TelaManContratos" Then
    
      XLT_CODEMPREENDIMENTO = Left$(Chave, 4)
      XLT_CODIMOVEL = Mid$(Chave, 6, 4)
      XLT_CODCONTRATO = Right$(Chave, 2)
    
      subCarregaDadosNV Adodc1, "", "Titulo", _
         "", "SELECT * FROM ConsCAPTitulosGrid WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XLT_CODEMPREENDIMENTO & "'" & _
         "AND imov_cd_Imovel= '" & XLT_CODIMOVEL & "' AND cont_cd_Contrato= '" & XLT_CODCONTRATO & "'"
    Else
      'Associa um banco de dados à tabela Imoveis
      subCarregaDadosNV Adodc1, "", "Titulo", _
         "", "SELECT * FROM ConsCAPTitulosGrid WHERE empr_cd_empresa = " & PCodEmpresa
    End If
    
    subHabilitaBotoes
    
    filtra_especial = True
    
    subTelaValoresGlobais "G"
    PrimeiraVez = True
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdRecalcular.Enabled = False
       CmdExclusaoMultipla.Enabled = False
       CmdAlteracaoMultipla.Enabled = False
       CmdPagamento.Enabled = False
       CmdImprimir.Enabled = False
       CmdOrdem.Enabled = False
       cmdPromissoria.Enabled = False
       cmdRepasse.Enabled = False
    End If
    
    If ResolucaoX < 10000 Then   '640x480
      TDBGrid1.Splits(0).Size = 4
      TDBGrid1.Splits(1).Size = 12
    ElseIf ResolucaoX > 10000 And ResolucaoX < 13000 Then  '800x600
      TDBGrid1.Splits(0).Size = 3
      TDBGrid1.Splits(1).Size = 12
    Else        '1024x728 ou superior
      TDBGrid1.Splits(0).Size = 2
      TDBGrid1.Splits(1).Size = 11
    End If
    
    If BarraFerramentasPressionada = True Then
      BarraFerramentasPressionada = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaFinTitulos"
End Sub

Private Sub TDBGrid1_SelChange(Cancel As Integer)

'Rotina para somar um conjunto de campos selecionados de uma coluna

Dim XLD_TOTALSELECAO As Variant
    Dim i As Integer

    i = 0
    
    XLD_TOTALSELECAO = 0
       
    If TDBGrid1.SelRange Then
        If TDBGrid1.SelStartCol = TDBGrid1.SelEndCol And (TDBGrid1.SelEndCol = 11) Then
        
            For i = 0 To TDBGrid1.SelBookmarks.Count - 1 Step 1
            
                If (TDBGrid1.Columns(TDBGrid1.SelEndCol).CellValue(TDBGrid1.SelBookmarks(i))) <> "" Then
                    XLD_TOTALSELECAO = XLD_TOTALSELECAO + (TDBGrid1.Columns(TDBGrid1.SelEndCol).CellValue(TDBGrid1.SelBookmarks(i)))
                End If
            Next
            
            MsgBox "Total selecionado: R$ " & Format(CStr(XLD_TOTALSELECAO), "0.00"), vbInformation, "Seleção"
        End If
        
        
    End If
End Sub

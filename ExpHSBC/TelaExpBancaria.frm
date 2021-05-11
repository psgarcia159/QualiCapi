VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form TelaExpBancaria 
   Caption         =   "Exportação Bancária - HSBC"
   ClientHeight    =   5730
   ClientLeft      =   75
   ClientTop       =   1470
   ClientWidth     =   9360
   Icon            =   "TelaExpBancaria.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   9360
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanDesconto 
      Height          =   2235
      Left            =   3420
      TabIndex        =   58
      Top             =   1500
      Visible         =   0   'False
      Width           =   3075
      _Version        =   65536
      _ExtentX        =   5424
      _ExtentY        =   3942
      _StockProps     =   15
      Caption         =   " Desconto"
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
         Left            =   1860
         TabIndex        =   65
         Top             =   1740
         Width           =   1065
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
         Left            =   600
         TabIndex        =   64
         Top             =   1740
         Width           =   1065
      End
      Begin VB.CheckBox ChkTodos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aplicar a todos os Títulos?"
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
         Left            =   180
         TabIndex        =   63
         Top             =   1260
         Width           =   2655
      End
      Begin VB.TextBox TxtDesconto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         TabIndex        =   62
         Top             =   420
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker DtpDesconto 
         Height          =   315
         Left            =   1620
         TabIndex        =   59
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   15597571
         CurrentDate     =   37658
      End
      Begin VB.Label LblDesconto 
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Left            =   660
         TabIndex        =   61
         Top             =   480
         Width           =   975
      End
      Begin VB.Label LblDescontoAte 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desconto Até:"
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
         TabIndex        =   60
         Top             =   900
         Width           =   1455
      End
   End
   Begin Threed.SSPanel PanPesquisa 
      Height          =   4335
      Left            =   0
      TabIndex        =   7
      Top             =   660
      Width           =   9255
      _Version        =   65536
      _ExtentX        =   16325
      _ExtentY        =   7646
      _StockProps     =   15
      Caption         =   "SSPanel1"
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.CommandButton CmdPesquisar 
         Caption         =   "&Pesquisar"
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
         Left            =   6420
         TabIndex        =   51
         Top             =   3840
         Width           =   1245
      End
      Begin VB.CommandButton CmdFechar 
         Caption         =   "&Fechar"
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
         Left            =   7740
         TabIndex        =   50
         Top             =   3840
         Width           =   1245
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3570
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6297
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Dados Gerais"
         TabPicture(0)   =   "TelaExpBancaria.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LblInstrucao1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "LblInstrucao2(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "LblCCorrente"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "LblNaturezaPlano"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "LblTipoPlano"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "LblDtExportacao"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "LblMesAno"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "DatObs"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "DatTipoPlano"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "DatContaCorrente"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "CboCCorrente"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "CmdLimparNaturezaPlano"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "CboTiposPlanos"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "CmdLimparTipoPlano"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "DtpExportacao"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "DtpVencimento"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "TxtInstrucao1"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "TxtInstrucao2"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "CboNaturezasPlanos"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).ControlCount=   19
         TabCaption(1)   =   "Empreendimentos"
         TabPicture(1)   =   "TelaExpBancaria.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label1"
         Tab(1).Control(1)=   "LlbEmpreendimento"
         Tab(1).Control(2)=   "TDBGridEmpr1"
         Tab(1).Control(3)=   "TDBGridEmpr2"
         Tab(1).Control(4)=   "CmdInserirTodosEmpreendimento"
         Tab(1).Control(5)=   "CmdRemoverTodosEmpreendimento"
         Tab(1).Control(6)=   "CmdInserirEmpreendimento"
         Tab(1).Control(7)=   "CmdRemoverEmpreendimento"
         Tab(1).ControlCount=   8
         TabCaption(2)   =   "Moedas"
         TabPicture(2)   =   "TelaExpBancaria.frx":0902
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label3"
         Tab(2).Control(1)=   "Label2"
         Tab(2).Control(2)=   "TDBGridMoeda1"
         Tab(2).Control(3)=   "TDBGridMoeda2"
         Tab(2).Control(4)=   "FraCorrecao"
         Tab(2).Control(5)=   "CmdInserirTodosMoeda"
         Tab(2).Control(6)=   "CmdRemoverTodosMoeda"
         Tab(2).Control(7)=   "CmdInserirMoeda"
         Tab(2).Control(8)=   "CmdRemoverMoeda"
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "Observações"
         TabPicture(3)   =   "TelaExpBancaria.frx":091E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label5"
         Tab(3).Control(1)=   "Label4"
         Tab(3).Control(2)=   "TDBGridObs1"
         Tab(3).Control(3)=   "TDBGridObs2"
         Tab(3).Control(4)=   "CmdRemoverObservacao"
         Tab(3).Control(5)=   "CmdInserirObservacao"
         Tab(3).Control(6)=   "CmdRemoverTodosObservacao"
         Tab(3).Control(7)=   "CmdInserirTodosObservacao"
         Tab(3).ControlCount=   8
         Begin VB.CommandButton CmdRemoverEmpreendimento 
            BackColor       =   &H00000000&
            Caption         =   "<"
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
            Left            =   -70830
            TabIndex        =   28
            Top             =   1770
            Width           =   690
         End
         Begin VB.CommandButton CmdInserirEmpreendimento 
            BackColor       =   &H00000000&
            Caption         =   ">"
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
            Left            =   -70830
            TabIndex        =   27
            Top             =   1320
            Width           =   690
         End
         Begin VB.CommandButton CmdRemoverTodosEmpreendimento 
            BackColor       =   &H00000000&
            Caption         =   "<<"
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
            Left            =   -70830
            TabIndex        =   26
            Top             =   2730
            Width           =   690
         End
         Begin VB.CommandButton CmdInserirTodosEmpreendimento 
            BackColor       =   &H00000000&
            Caption         =   ">>"
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
            Left            =   -70830
            TabIndex        =   25
            Top             =   2280
            Width           =   690
         End
         Begin VB.CommandButton CmdRemoverMoeda 
            BackColor       =   &H00000000&
            Caption         =   "<"
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
            Left            =   -70830
            TabIndex        =   24
            Top             =   2040
            Width           =   690
         End
         Begin VB.CommandButton CmdInserirMoeda 
            BackColor       =   &H00000000&
            Caption         =   ">"
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
            Left            =   -70830
            TabIndex        =   23
            Top             =   1560
            Width           =   690
         End
         Begin VB.CommandButton CmdRemoverTodosMoeda 
            BackColor       =   &H00000000&
            Caption         =   "<<"
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
            Left            =   -70830
            TabIndex        =   22
            Top             =   3000
            Width           =   690
         End
         Begin VB.CommandButton CmdInserirTodosMoeda 
            BackColor       =   &H00000000&
            Caption         =   ">>"
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
            Left            =   -70830
            TabIndex        =   21
            Top             =   2535
            Width           =   690
         End
         Begin VB.CommandButton CmdInserirTodosObservacao 
            BackColor       =   &H00000000&
            Caption         =   ">>"
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
            Left            =   -70830
            TabIndex        =   20
            Top             =   2265
            Width           =   690
         End
         Begin VB.CommandButton CmdRemoverTodosObservacao 
            BackColor       =   &H00000000&
            Caption         =   "<<"
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
            Left            =   -70830
            TabIndex        =   19
            Top             =   2730
            Width           =   690
         End
         Begin VB.CommandButton CmdInserirObservacao 
            BackColor       =   &H00000000&
            Caption         =   ">"
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
            Left            =   -70830
            TabIndex        =   18
            Top             =   1290
            Width           =   690
         End
         Begin VB.CommandButton CmdRemoverObservacao 
            BackColor       =   &H00000000&
            Caption         =   "<"
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
            Left            =   -70830
            TabIndex        =   17
            Top             =   1770
            Width           =   690
         End
         Begin VB.Frame FraCorrecao 
            Caption         =   "Correção"
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
            Left            =   -74760
            TabIndex        =   13
            Top             =   420
            Width           =   4995
            Begin VB.OptionButton OptIndexador2 
               Caption         =   "Indexador 2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3480
               TabIndex        =   16
               Top             =   225
               Width           =   1395
            End
            Begin VB.OptionButton OptIndexador1 
               Caption         =   "Indexador 1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1920
               TabIndex        =   15
               Top             =   225
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton OptValorHistorico 
               Caption         =   "Valor Histórico"
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
               TabIndex        =   14
               Top             =   285
               Width           =   1575
            End
         End
         Begin VB.ComboBox CboNaturezasPlanos 
            Height          =   315
            ItemData        =   "TelaExpBancaria.frx":093A
            Left            =   1965
            List            =   "TelaExpBancaria.frx":0950
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1485
            Width           =   3165
         End
         Begin VB.TextBox TxtInstrucao2 
            Height          =   315
            Left            =   1965
            MaxLength       =   2
            TabIndex        =   11
            Top             =   2715
            Width           =   345
         End
         Begin VB.TextBox TxtInstrucao1 
            Height          =   315
            Left            =   1965
            MaxLength       =   2
            TabIndex        =   10
            Top             =   2280
            Width           =   345
         End
         Begin TrueDBGrid70.TDBGrid TDBGridEmpr2 
            Height          =   2625
            Left            =   -69960
            TabIndex        =   29
            Top             =   780
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   4630
            _LayoutType     =   4
            _RowHeight      =   25
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Empreendimento"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=66064"
            Splits(0)._ColumnProps(7)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(8)=   "Column(0).FetchStyle=1"
            Splits(0)._ColumnProps(9)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(10)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(11)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(12)=   "Column(1).Width=1296"
            Splits(0)._ColumnProps(13)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._WidthInPix=1217"
            Splits(0)._ColumnProps(15)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(1)._ColStyle=74256"
            Splits(0)._ColumnProps(17)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(18)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(19)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(20)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(21)=   "Column(1)._MinWidth=80"
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
            HeadLines       =   2
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
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8,.bgcolor=&H800000&"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=21,.alignment=0"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22,.alignment=2,.bold=-1"
            _StyleDefs(32)  =   ":id=29,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(33)  =   ":id=29,.fontname=MS Sans Serif"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(37)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2,.bold=-1"
            _StyleDefs(39)  =   ":id=83,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(40)  =   ":id=83,.fontname=MS Sans Serif"
            _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
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
         Begin TrueDBGrid70.TDBGrid TDBGridMoeda2 
            Height          =   1965
            Left            =   -69960
            TabIndex        =   30
            Top             =   1425
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   3466
            _LayoutType     =   4
            _RowHeight      =   25
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Moeda"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=66064"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(9)=   "Column(0).FetchStyle=1"
            Splits(0)._ColumnProps(10)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(11)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(12)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(13)=   "Column(1).Width=1296"
            Splits(0)._ColumnProps(14)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(1)._WidthInPix=1217"
            Splits(0)._ColumnProps(16)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(1)._ColStyle=74256"
            Splits(0)._ColumnProps(18)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(19)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(20)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(21)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(22)=   "Column(1)._MinWidth=80"
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
            HeadLines       =   2
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
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8,.bgcolor=&H800000&"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=21,.alignment=0"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22,.alignment=2,.bold=-1"
            _StyleDefs(32)  =   ":id=29,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(33)  =   ":id=29,.fontname=MS Sans Serif"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(37)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2,.bold=-1"
            _StyleDefs(39)  =   ":id=83,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(40)  =   ":id=83,.fontname=MS Sans Serif"
            _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
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
         Begin TrueDBGrid70.TDBGrid TDBGridObs2 
            Height          =   2565
            Left            =   -69960
            TabIndex        =   31
            Top             =   840
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   4524
            _LayoutType     =   4
            _RowHeight      =   25
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Observação"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=66064"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(9)=   "Column(0).FetchStyle=1"
            Splits(0)._ColumnProps(10)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(11)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(12)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(13)=   "Column(1).Width=1296"
            Splits(0)._ColumnProps(14)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(1)._WidthInPix=1217"
            Splits(0)._ColumnProps(16)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(1)._ColStyle=74256"
            Splits(0)._ColumnProps(18)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(19)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(20)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(21)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(22)=   "Column(1)._MinWidth=80"
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
            HeadLines       =   2
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
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8,.bgcolor=&H800000&"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=21,.alignment=0"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22,.alignment=2,.bold=-1"
            _StyleDefs(32)  =   ":id=29,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(33)  =   ":id=29,.fontname=MS Sans Serif"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(37)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2,.bold=-1"
            _StyleDefs(39)  =   ":id=83,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(40)  =   ":id=83,.fontname=MS Sans Serif"
            _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
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
         Begin MSComCtl2.DTPicker DtpVencimento 
            Height          =   315
            Left            =   1965
            TabIndex        =   32
            Top             =   675
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MM/yy"
            Format          =   15597571
            CurrentDate     =   37636
         End
         Begin MSComCtl2.DTPicker DtpExportacao 
            Height          =   315
            Left            =   4680
            TabIndex        =   33
            Top             =   660
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   15597571
            CurrentDate     =   37180
         End
         Begin Threed.SSCommand CmdLimparTipoPlano 
            Height          =   315
            Left            =   5160
            TabIndex        =   34
            Top             =   1080
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "TelaExpBancaria.frx":098B
            Picture         =   "TelaExpBancaria.frx":09A7
         End
         Begin MSDataListLib.DataCombo CboTiposPlanos 
            Bindings        =   "TelaExpBancaria.frx":0AB9
            Height          =   315
            Left            =   1965
            TabIndex        =   35
            Top             =   1080
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "tipl_tx_Descricao"
            BoundColumn     =   "tipl_cd_TipoPlano"
            Text            =   "CboTiposPlanos"
         End
         Begin Threed.SSCommand CmdLimparNaturezaPlano 
            Height          =   315
            Left            =   5160
            TabIndex        =   36
            Top             =   1485
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "TelaExpBancaria.frx":0AD4
            Picture         =   "TelaExpBancaria.frx":0AF0
         End
         Begin MSDataListLib.DataCombo CboCCorrente 
            Bindings        =   "TelaExpBancaria.frx":0C02
            Height          =   330
            Left            =   1965
            TabIndex        =   37
            Top             =   1920
            Width           =   6810
            _ExtentX        =   12012
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
         Begin MSAdodcLib.Adodc DatContaCorrente 
            Height          =   330
            Left            =   5760
            Top             =   1500
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
         Begin MSAdodcLib.Adodc DatTipoPlano 
            Height          =   330
            Left            =   7200
            Top             =   1140
            Visible         =   0   'False
            Width           =   1725
            _ExtentX        =   3043
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
            Caption         =   "DatTipoPlano"
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
         Begin MSAdodcLib.Adodc DatObs 
            Height          =   330
            Left            =   5760
            Top             =   1140
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
         Begin TrueDBGrid70.TDBGrid TDBGridEmpr1 
            Height          =   2625
            Left            =   -74760
            TabIndex        =   52
            Top             =   780
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   4630
            _LayoutType     =   4
            _RowHeight      =   25
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Empreendimento"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=66064"
            Splits(0)._ColumnProps(7)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(8)=   "Column(0).FetchStyle=1"
            Splits(0)._ColumnProps(9)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(10)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(11)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(12)=   "Column(1).Width=1296"
            Splits(0)._ColumnProps(13)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._WidthInPix=1217"
            Splits(0)._ColumnProps(15)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(1)._ColStyle=74256"
            Splits(0)._ColumnProps(17)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(18)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(19)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(20)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(21)=   "Column(1)._MinWidth=80"
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
            HeadLines       =   2
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
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8,.bgcolor=&H800000&"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=21,.alignment=0"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22,.alignment=2,.bold=-1"
            _StyleDefs(32)  =   ":id=29,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(33)  =   ":id=29,.fontname=MS Sans Serif"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(37)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2,.bold=-1"
            _StyleDefs(39)  =   ":id=83,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(40)  =   ":id=83,.fontname=MS Sans Serif"
            _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
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
         Begin TrueDBGrid70.TDBGrid TDBGridMoeda1 
            Height          =   1965
            Left            =   -74760
            TabIndex        =   53
            Top             =   1440
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   3466
            _LayoutType     =   4
            _RowHeight      =   25
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Moeda"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=66064"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(9)=   "Column(0).FetchStyle=1"
            Splits(0)._ColumnProps(10)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(11)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(12)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(13)=   "Column(1).Width=1296"
            Splits(0)._ColumnProps(14)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(1)._WidthInPix=1217"
            Splits(0)._ColumnProps(16)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(1)._ColStyle=74256"
            Splits(0)._ColumnProps(18)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(19)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(20)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(21)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(22)=   "Column(1)._MinWidth=80"
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
            HeadLines       =   2
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
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8,.bgcolor=&H800000&"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=21,.alignment=0"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22,.alignment=2,.bold=-1"
            _StyleDefs(32)  =   ":id=29,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(33)  =   ":id=29,.fontname=MS Sans Serif"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(37)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2,.bold=-1"
            _StyleDefs(39)  =   ":id=83,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(40)  =   ":id=83,.fontname=MS Sans Serif"
            _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
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
         Begin TrueDBGrid70.TDBGrid TDBGridObs1 
            Height          =   2565
            Left            =   -74760
            TabIndex        =   54
            Top             =   840
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   4524
            _LayoutType     =   4
            _RowHeight      =   25
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Observação"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=66064"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).WrapText=1"
            Splits(0)._ColumnProps(9)=   "Column(0).FetchStyle=1"
            Splits(0)._ColumnProps(10)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(11)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(12)=   "Column(0)._MinWidth=80"
            Splits(0)._ColumnProps(13)=   "Column(1).Width=1296"
            Splits(0)._ColumnProps(14)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(1)._WidthInPix=1217"
            Splits(0)._ColumnProps(16)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(1)._ColStyle=74256"
            Splits(0)._ColumnProps(18)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(19)=   "Column(1).FetchStyle=1"
            Splits(0)._ColumnProps(20)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(21)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(22)=   "Column(1)._MinWidth=80"
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
            HeadLines       =   2
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
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8,.bgcolor=&H800000&"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=21,.alignment=0"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=22,.alignment=2,.bold=-1"
            _StyleDefs(32)  =   ":id=29,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(33)  =   ":id=29,.fontname=MS Sans Serif"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=23"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=75"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=86,.parent=21,.alignment=0,.valignment=2"
            _StyleDefs(37)  =   ":id=86,.wraptext=-1,.locked=-1"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=22,.alignment=2,.bold=-1"
            _StyleDefs(39)  =   ":id=83,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(40)  =   ":id=83,.fontname=MS Sans Serif"
            _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=23"
            _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=75"
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
         Begin VB.Label LlbEmpreendimento 
            Caption         =   "Empreendimentos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   -74805
            TabIndex        =   8
            Top             =   390
            Width           =   2205
         End
         Begin VB.Label Label1 
            Caption         =   "Empreendimentos Escolhidos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   330
            Left            =   -69945
            TabIndex        =   49
            Top             =   390
            Width           =   3180
         End
         Begin VB.Label Label2 
            Caption         =   "Moedas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   -74745
            TabIndex        =   48
            Top             =   1140
            Width           =   3705
         End
         Begin VB.Label Label3 
            Caption         =   "Moedas Escolhidas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   330
            Left            =   -69945
            TabIndex        =   47
            Top             =   1140
            Width           =   3720
         End
         Begin VB.Label Label4 
            Caption         =   "Observações Escolhidas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   330
            Left            =   -69945
            TabIndex        =   46
            Top             =   450
            Width           =   3720
         End
         Begin VB.Label Label5 
            Caption         =   "Observações"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   -74745
            TabIndex        =   45
            Top             =   450
            Width           =   3705
         End
         Begin VB.Label LblMesAno 
            Alignment       =   1  'Right Justify
            Caption         =   "Mês de Vencimento:"
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
            Left            =   90
            TabIndex        =   44
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label LblDtExportacao 
            Alignment       =   1  'Right Justify
            Caption         =   "Data Exportação:"
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
            TabIndex        =   43
            Top             =   720
            Width           =   1530
         End
         Begin VB.Label LblTipoPlano 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo do Plano:"
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
            Left            =   645
            TabIndex        =   42
            Top             =   1125
            Width           =   1260
         End
         Begin VB.Label LblNaturezaPlano 
            Alignment       =   1  'Right Justify
            Caption         =   "Natureza do Plano:"
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
            Left            =   225
            TabIndex        =   41
            Top             =   1530
            Width           =   1680
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
            Height          =   210
            Left            =   540
            TabIndex        =   40
            Top             =   1950
            Width           =   1365
         End
         Begin VB.Label LblInstrucao2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Instrução 2:"
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
            Index           =   1
            Left            =   780
            TabIndex        =   39
            Top             =   2775
            Width           =   1125
         End
         Begin VB.Label LblInstrucao1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Instrução 1:"
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
            Index           =   0
            Left            =   780
            TabIndex        =   38
            Top             =   2355
            Width           =   1125
         End
      End
      Begin MSComDlg.CommonDialog DlgSalvar 
         Left            =   360
         Top             =   3780
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin TrueDBGrid70.TDBGrid TDBGrid1 
      Height          =   4650
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8202
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Título"
      Columns(0).FooterText=   "T"
      Columns(0).DataField=   "Titulo"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   68
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Exp."
      Columns(1).DataField=   "Exporta"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Vencimento"
      Columns(2).FooterText=   "D"
      Columns(2).DataField=   "titu_dt_Vencimento"
      Columns(2).NumberFormat=   "dd/mm/yy"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Valor R$"
      Columns(3).FooterText=   "T"
      Columns(3).DataField=   "Valor Real"
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nome"
      Columns(4).DataField=   "obse_tx_Observacao"
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Desconto"
      Columns(5).FooterText=   "N"
      Columns(5).DataField=   "titu_vl_Desconto"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Desconto Até"
      Columns(6).FooterText=   "D"
      Columns(6).DataField=   "LimiteDesconto"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Val. Tit. R$"
      Columns(7).FooterText=   "N"
      Columns(7).DataField=   "titu_vl_Parcela"
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Val. Seguro R$"
      Columns(8).DataField=   ""
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "SaldoDevedor"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   4
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4630"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4551"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8720"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=847"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=767"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=529"
      Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(16)=   "Column(1)._MinWidth=78536464"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=1958"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=1879"
      Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=8721"
      Splits(0)._ColumnProps(22)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(23)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(25)=   "Column(2)._MinWidth=100555328"
      Splits(0)._ColumnProps(26)=   "Column(3).Width=2646"
      Splits(0)._ColumnProps(27)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._WidthInPix=2566"
      Splits(0)._ColumnProps(29)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(3)._ColStyle=8722"
      Splits(0)._ColumnProps(31)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(32)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(33)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(34)=   "Column(3)._MinWidth=100544304"
      Splits(0)._ColumnProps(35)=   "Column(4).Width=2037"
      Splits(0)._ColumnProps(36)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(4)._WidthInPix=1958"
      Splits(0)._ColumnProps(38)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(4)._ColStyle=8720"
      Splits(0)._ColumnProps(40)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(41)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(42)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(43)=   "Column(5).Width=1667"
      Splits(0)._ColumnProps(44)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(5)._WidthInPix=1588"
      Splits(0)._ColumnProps(46)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(5)._ColStyle=8978"
      Splits(0)._ColumnProps(48)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(49)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(50)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(51)=   "Column(6).Width=3704"
      Splits(0)._ColumnProps(52)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(6)._WidthInPix=3625"
      Splits(0)._ColumnProps(54)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(55)=   "Column(6)._ColStyle=529"
      Splits(0)._ColumnProps(56)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(57)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(58)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(59)=   "Column(7).Width=3228"
      Splits(0)._ColumnProps(60)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(7)._WidthInPix=3149"
      Splits(0)._ColumnProps(62)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(63)=   "Column(7)._ColStyle=530"
      Splits(0)._ColumnProps(64)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(65)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(66)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(67)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(68)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(70)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(71)=   "Column(8)._ColStyle=20"
      Splits(0)._ColumnProps(72)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(73)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(74)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(75)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(76)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(77)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(78)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(79)=   "Column(9)._ColStyle=20"
      Splits(0)._ColumnProps(80)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(81)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(82)=   "Column(9).Order=10"
      Splits(1)._UserFlags=   0
      Splits(1).ExtendRightColumn=   -1  'True
      Splits(1).MarqueeStyle=   4
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1).DividerColor=   12632256
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=10"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=4630"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=4551"
      Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=8721"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).WrapText=1"
      Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Width=847"
      Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=767"
      Splits(1)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=529"
      Splits(1)._ColumnProps(14)=   "Column(1).WrapText=1"
      Splits(1)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(16)=   "Column(2).Width=1905"
      Splits(1)._ColumnProps(17)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(18)=   "Column(2)._WidthInPix=1826"
      Splits(1)._ColumnProps(19)=   "Column(2)._EditAlways=0"
      Splits(1)._ColumnProps(20)=   "Column(2)._ColStyle=8721"
      Splits(1)._ColumnProps(21)=   "Column(2).WrapText=1"
      Splits(1)._ColumnProps(22)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(23)=   "Column(3).Width=2064"
      Splits(1)._ColumnProps(24)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(25)=   "Column(3)._WidthInPix=1984"
      Splits(1)._ColumnProps(26)=   "Column(3)._EditAlways=0"
      Splits(1)._ColumnProps(27)=   "Column(3)._ColStyle=8722"
      Splits(1)._ColumnProps(28)=   "Column(3).WrapText=1"
      Splits(1)._ColumnProps(29)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(30)=   "Column(4).Width=5450"
      Splits(1)._ColumnProps(31)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(32)=   "Column(4)._WidthInPix=5371"
      Splits(1)._ColumnProps(33)=   "Column(4)._EditAlways=0"
      Splits(1)._ColumnProps(34)=   "Column(4)._ColStyle=8720"
      Splits(1)._ColumnProps(35)=   "Column(4).WrapText=1"
      Splits(1)._ColumnProps(36)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(37)=   "Column(4)._MinWidth=-1"
      Splits(1)._ColumnProps(38)=   "Column(5).Width=1588"
      Splits(1)._ColumnProps(39)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(40)=   "Column(5)._WidthInPix=1508"
      Splits(1)._ColumnProps(41)=   "Column(5)._EditAlways=0"
      Splits(1)._ColumnProps(42)=   "Column(5)._ColStyle=8978"
      Splits(1)._ColumnProps(43)=   "Column(5).WrapText=1"
      Splits(1)._ColumnProps(44)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(45)=   "Column(6).Width=2196"
      Splits(1)._ColumnProps(46)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(47)=   "Column(6)._WidthInPix=2117"
      Splits(1)._ColumnProps(48)=   "Column(6)._EditAlways=0"
      Splits(1)._ColumnProps(49)=   "Column(6)._ColStyle=529"
      Splits(1)._ColumnProps(50)=   "Column(6).WrapText=1"
      Splits(1)._ColumnProps(51)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(52)=   "Column(7).Width=2117"
      Splits(1)._ColumnProps(53)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(54)=   "Column(7)._WidthInPix=2037"
      Splits(1)._ColumnProps(55)=   "Column(7)._EditAlways=0"
      Splits(1)._ColumnProps(56)=   "Column(7)._ColStyle=530"
      Splits(1)._ColumnProps(57)=   "Column(7).WrapText=1"
      Splits(1)._ColumnProps(58)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(59)=   "Column(8).Width=2487"
      Splits(1)._ColumnProps(60)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(61)=   "Column(8)._WidthInPix=2408"
      Splits(1)._ColumnProps(62)=   "Column(8)._EditAlways=0"
      Splits(1)._ColumnProps(63)=   "Column(8)._ColStyle=18"
      Splits(1)._ColumnProps(64)=   "Column(8).WrapText=1"
      Splits(1)._ColumnProps(65)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(66)=   "Column(9).Width=2725"
      Splits(1)._ColumnProps(67)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(68)=   "Column(9)._WidthInPix=2646"
      Splits(1)._ColumnProps(69)=   "Column(9)._EditAlways=0"
      Splits(1)._ColumnProps(70)=   "Column(9)._ColStyle=20"
      Splits(1)._ColumnProps(71)=   "Column(9).Visible=0"
      Splits(1)._ColumnProps(72)=   "Column(9).WrapText=1"
      Splits(1)._ColumnProps(73)=   "Column(9).Order=10"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
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
      _StyleDefs(24)  =   "Splits(0).Style:id=123,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=132,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=124,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=125,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=126,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=128,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=127,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=129,.parent=8,.bgcolor=&H800000&"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=130,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=131,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=133,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=134,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=138,.parent=123,.alignment=0,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=135,.parent=124,.alignment=2"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=136,.parent=125"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=137,.parent=127"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=142,.parent=123,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=139,.parent=124,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=140,.parent=125"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=141,.parent=127"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=146,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=143,.parent=124,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=144,.parent=125"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=145,.parent=127"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=150,.parent=123,.alignment=1,.locked=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=147,.parent=124,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=148,.parent=125"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=149,.parent=127"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=154,.parent=123,.alignment=0,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=151,.parent=124,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=152,.parent=125"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=153,.parent=127"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=158,.parent=123,.alignment=1,.locked=-1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=155,.parent=124,.alignment=1"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=156,.parent=125"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=157,.parent=127"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=162,.parent=123,.alignment=2,.locked=0"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=159,.parent=124,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=160,.parent=125"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=161,.parent=127"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=166,.parent=123,.alignment=1,.locked=0"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=163,.parent=124,.alignment=2"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=164,.parent=125"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=165,.parent=127"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=82,.parent=123"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=124"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=125"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=127"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=78,.parent=123"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=124"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=125"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=127"
      _StyleDefs(76)  =   "Splits(1).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(77)  =   "Splits(1).CaptionStyle:id=22,.parent=4"
      _StyleDefs(78)  =   "Splits(1).HeadingStyle:id=14,.parent=2"
      _StyleDefs(79)  =   "Splits(1).FooterStyle:id=15,.parent=3"
      _StyleDefs(80)  =   "Splits(1).InactiveStyle:id=16,.parent=5"
      _StyleDefs(81)  =   "Splits(1).SelectedStyle:id=18,.parent=6"
      _StyleDefs(82)  =   "Splits(1).EditorStyle:id=17,.parent=7"
      _StyleDefs(83)  =   "Splits(1).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H800000&"
      _StyleDefs(84)  =   "Splits(1).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(85)  =   "Splits(1).OddRowStyle:id=21,.parent=10"
      _StyleDefs(86)  =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(87)  =   "Splits(1).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(88)  =   "Splits(1).Columns(0).Style:id=28,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(89)  =   "Splits(1).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(90)  =   "Splits(1).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(91)  =   "Splits(1).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(92)  =   "Splits(1).Columns(1).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(93)  =   "Splits(1).Columns(1).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(94)  =   "Splits(1).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(95)  =   "Splits(1).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(96)  =   "Splits(1).Columns(2).Style:id=50,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(97)  =   "Splits(1).Columns(2).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(98)  =   "Splits(1).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(99)  =   "Splits(1).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(100) =   "Splits(1).Columns(3).Style:id=58,.parent=13,.alignment=1,.locked=-1"
      _StyleDefs(101) =   "Splits(1).Columns(3).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(102) =   "Splits(1).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(103) =   "Splits(1).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(104) =   "Splits(1).Columns(4).Style:id=62,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(105) =   "Splits(1).Columns(4).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(106) =   "Splits(1).Columns(4).FooterStyle:id=60,.parent=15"
      _StyleDefs(107) =   "Splits(1).Columns(4).EditorStyle:id=61,.parent=17"
      _StyleDefs(108) =   "Splits(1).Columns(5).Style:id=66,.parent=13,.alignment=1,.locked=-1"
      _StyleDefs(109) =   "Splits(1).Columns(5).HeadingStyle:id=63,.parent=14,.alignment=1"
      _StyleDefs(110) =   "Splits(1).Columns(5).FooterStyle:id=64,.parent=15"
      _StyleDefs(111) =   "Splits(1).Columns(5).EditorStyle:id=65,.parent=17"
      _StyleDefs(112) =   "Splits(1).Columns(6).Style:id=70,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(113) =   "Splits(1).Columns(6).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(114) =   "Splits(1).Columns(6).FooterStyle:id=68,.parent=15"
      _StyleDefs(115) =   "Splits(1).Columns(6).EditorStyle:id=69,.parent=17"
      _StyleDefs(116) =   "Splits(1).Columns(7).Style:id=74,.parent=13,.alignment=1,.locked=0"
      _StyleDefs(117) =   "Splits(1).Columns(7).HeadingStyle:id=71,.parent=14,.alignment=2"
      _StyleDefs(118) =   "Splits(1).Columns(7).FooterStyle:id=72,.parent=15"
      _StyleDefs(119) =   "Splits(1).Columns(7).EditorStyle:id=73,.parent=17"
      _StyleDefs(120) =   "Splits(1).Columns(8).Style:id=86,.parent=13,.alignment=1"
      _StyleDefs(121) =   "Splits(1).Columns(8).HeadingStyle:id=83,.parent=14"
      _StyleDefs(122) =   "Splits(1).Columns(8).FooterStyle:id=84,.parent=15"
      _StyleDefs(123) =   "Splits(1).Columns(8).EditorStyle:id=85,.parent=17"
      _StyleDefs(124) =   "Splits(1).Columns(9).Style:id=90,.parent=13"
      _StyleDefs(125) =   "Splits(1).Columns(9).HeadingStyle:id=87,.parent=14"
      _StyleDefs(126) =   "Splits(1).Columns(9).FooterStyle:id=88,.parent=15"
      _StyleDefs(127) =   "Splits(1).Columns(9).EditorStyle:id=89,.parent=17"
      _StyleDefs(128) =   "Named:id=33:Normal"
      _StyleDefs(129) =   ":id=33,.parent=0"
      _StyleDefs(130) =   "Named:id=34:Heading"
      _StyleDefs(131) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(132) =   ":id=34,.wraptext=-1"
      _StyleDefs(133) =   "Named:id=35:Footing"
      _StyleDefs(134) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(135) =   "Named:id=36:Selected"
      _StyleDefs(136) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(137) =   "Named:id=37:Caption"
      _StyleDefs(138) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(139) =   "Named:id=38:HighlightRow"
      _StyleDefs(140) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(141) =   "Named:id=39:EvenRow"
      _StyleDefs(142) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(143) =   "Named:id=40:OddRow"
      _StyleDefs(144) =   ":id=40,.parent=33"
      _StyleDefs(145) =   "Named:id=41:RecordSelector"
      _StyleDefs(146) =   ":id=41,.parent=34"
      _StyleDefs(147) =   "Named:id=42:FilterBar"
      _StyleDefs(148) =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   885
      Left            =   45
      TabIndex        =   1
      Top             =   4740
      Width           =   9240
      _Version        =   65536
      _ExtentX        =   16298
      _ExtentY        =   1561
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
      Begin VB.CommandButton CmdDesconto 
         Caption         =   "&Desconto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2700
         TabIndex        =   57
         Top             =   0
         Width           =   1680
      End
      Begin VB.CommandButton CmdMarcar 
         Caption         =   "&Marcar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4620
         TabIndex        =   56
         Top             =   15
         Width           =   1680
      End
      Begin VB.CommandButton CmdPesquisa 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   720
         TabIndex        =   55
         Top             =   0
         Width           =   1680
      End
      Begin VB.CommandButton CmdDesmarcar 
         Caption         =   "Desmarcar &Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6540
         TabIndex        =   6
         Top             =   15
         Width           =   1680
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
         Height          =   345
         Left            =   4590
         TabIndex        =   4
         Top             =   495
         Width           =   1680
      End
      Begin VB.CommandButton CmdFiltro 
         Caption         =   "&Seleções"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2700
         TabIndex        =   3
         Top             =   495
         Width           =   1680
      End
      Begin VB.CommandButton CmdExportar 
         Caption         =   "&Exportar"
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
         Height          =   345
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   1680
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
         Height          =   345
         Left            =   6495
         TabIndex        =   0
         Top             =   495
         Width           =   1680
      End
   End
   Begin MSComDlg.CommonDialog CdbImpressora 
      Left            =   0
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "TelaExpBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim VFV_VETOREXP As New XArray
Dim VFV_EMPREENDIMENTO1 As New XArray
Dim VFV_EMPREENDIMENTO2 As New XArray
Dim VFV_MOEDA1 As New XArray
Dim VFV_MOEDA2 As New XArray
Dim VFV_OBSERVACAO1 As New XArray
Dim VFV_OBSERVACAO2 As New XArray
Dim XFO_EXPORTACAO As New ADODB.Recordset
Dim XFF_SALDODEVEDOR As Double
Dim CarregouGrid As Boolean 'Indica se o grid já foi carregado ou não

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

Dim FRM_XInserir As String        'Flag para identificar se entrou na rotina para inserir ou alterar ('I'Inserir, 'A'Alterar, 'C'Consultar)
Dim FRM_XOrigem As String         'Identifica a origem (que formulario) chamou o formulário ativo
Dim FRM_SQLAtivo As String        'Atualizar o SQL Ativo após modificações de filtro e ordem


Dim XFT_SQL As String       'Prepara Select para o vetor do grid
Dim XFT_SQL2 As String      'Prepara Select para o recordset auxiliar

Sub AjustaTela()
    'Ajustar a Tela à resolução do monitor
    Formulario.Top = 0
    Formulario.Height = Resolucaoy - 500
    Formulario.Width = ResolucaoX - 60
    Formulario.Left = 5
    Formulario.TDBGrid1.Top = 45
    Formulario.TDBGrid1.Left = 75
    Formulario.TDBGrid1.Width = Formulario.Width - 150
    Formulario.TDBGrid1.Height = Formulario.Height - Formulario.TDBGrid1.Top - Formulario.PanBotoes.Height - 50
    Formulario.PanBotoes.Top = Formulario.TDBGrid1.Top + Formulario.TDBGrid1.Height + 50
    Formulario.PanBotoes.Left = 75 + (Formulario.TDBGrid1.Width - Formulario.PanBotoes.Width) / 2
End Sub

Sub subTelaValoresGlobais(GuardaOUPegaValores As String)
    Dim Count As Integer, Count2 As Integer
    If GuardaOUPegaValores = "G" Then 'Guardando variáveis de definição da tela
        FRM_ContFil = ContFil
        FRM_Ordem = Ordem
        FRM_Campos_Obrig = Campos_Obrig
        FRM_TipoVar = XGT_CAMPOFILTRO
        FRM_FiltroOrig = FiltroOrig
        FRM_FiltroAtual = FiltroAtual
        FRM_filtra_especial = filtra_especial
        FRM_Tem_Filtro = Tem_Filtro
        FRM_ConTelaOrdem = ConTelaOrdem
        FRM_Quantid = Quantid
        FRM_Filtrou = Filtrou
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
        XGT_CAMPOFILTRO = FRM_TipoVar
        FiltroOrig = FRM_FiltroOrig
        FiltroAtual = FRM_FiltroAtual
        filtra_especial = FRM_filtra_especial
        Tem_Filtro = FRM_Tem_Filtro
        ConTelaOrdem = FRM_ConTelaOrdem
        Quantid = FRM_Quantid
        Filtrou = FRM_Filtrou
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
    CmdDesmarcar.Enabled = True
    CmdMarcar.Enabled = True
    CmdExportar.Enabled = True
    CmdDesconto.Enabled = True
    CmdImprimir.Enabled = True
    
    'Verifica se tem Filtro Ativo
    If ContFil = 0 Then
       CmdFiltro.Enabled = False
    Else
       CmdFiltro.Enabled = True
    End If
End Sub

Private Sub subDesabilitaBotoes()
    CmdDesmarcar.Enabled = False
    CmdMarcar.Enabled = False
    CmdExportar.Enabled = False
    CmdDesconto.Enabled = False
    CmdImprimir.Enabled = False
    CmdFiltro.Enabled = False
End Sub

Sub SubCalculaValorAtualizado()

  Dim XLI_CONT As Integer
    
  XFO_EXPORTACAO.MoveFirst
  For XLI_CONT = 0 To TDBGrid1.ApproxCount - 1
    If OptIndexador1.Value = True Then
      VFV_VETOREXP(XLI_CONT, 7) = funCalculaCorrecaoMonetaria(XFO_EXPORTACAO, 1, XFO_EXPORTACAO!titu_dt_Vencimento)
      VFV_VETOREXP(XLI_CONT, 3) = VFV_VETOREXP(XLI_CONT, 7)
    ElseIf OptIndexador2.Value = True Then
      VFV_VETOREXP(XLI_CONT, 7) = funCalculaCorrecaoMonetaria(XFO_EXPORTACAO, 2, XFO_EXPORTACAO!titu_dt_Vencimento)
      VFV_VETOREXP(XLI_CONT, 3) = VFV_VETOREXP(XLI_CONT, 7)
    Else
      VFV_VETOREXP(XLI_CONT, 7) = XFO_EXPORTACAO!titu_vl_Parcela
      VFV_VETOREXP(XLI_CONT, 3) = VFV_VETOREXP(XLI_CONT, 7)
    End If
    XFO_EXPORTACAO.MoveNext
  Next
  
End Sub

Function FunCalculaSaldoDevedor(XLO_RECORDSET As ADODB.Recordset, XLD_DATABASE As Date) As Double

  Dim XLT_SQL As String
  Dim XLO_TITULO As New ADODB.Recordset
  Dim XLF_VALORCORRIGIDO As Double
  Dim XLF_DESAGIO As Double
  Dim XLF_VALORACUMULADO As Double
  Dim XLD_DATA As Date

  XLD_DATA = XLD_DATABASE
  
  XLT_SQL = "" & _
  "SELECT ConsCAPExpBanco.*," & _
    "CASE Indexacao1   when 'M' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(TDBGrid1.Columns("Vencimento")) & "/" & Year(TDBGrid1.Columns("Vencimento")), NomeSgbd) & ")" & _
      " when 'D' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(TDBGrid1.Columns("Vencimento"), NomeSgbd) & ")" & _
    " end as Cotacao1," & _
    "CASE  Indexacao2 when 'M' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(TDBGrid1.Columns("Vencimento")) & "/" & Year(TDBGrid1.Columns("Vencimento")), NomeSgbd) & ")" & _
      " when 'D'  then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(TDBGrid1.Columns("Vencimento"), NomeSgbd) & ")" & _
    " end as Cotacao2,"
    
  XLT_SQL = XLT_SQL & _
    "CASE  Indexacao1 when 'M' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(TDBGrid1.Columns("Vencimento")) & "/" & Year(TDBGrid1.Columns("Vencimento")), NomeSgbd) & "))" & _
    " end as Cotacao1_MesAnterior," & _
    "CASE Indexacao2 when 'M' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(TDBGrid1.Columns("Vencimento")) & "/" & Year(TDBGrid1.Columns("Vencimento")), NomeSgbd) & "))" & _
    " end As Cotacao2_MesAnterior" & _
    " From ConsCAPExpBanco WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
    " AND empd_cd_Empreendimento= '" & Left(XLO_RECORDSET("Titulo"), 4) & "'" & _
    " AND imov_cd_Imovel='" & Mid(XLO_RECORDSET("Titulo"), 6, 4) & "'" & _
    " AND cont_cd_Contrato='" & Mid(XLO_RECORDSET("Titulo"), 11, 2) & "'" & _
    " AND titu_dt_Pagamento is null "
   
  SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
 
  While Not XLO_TITULO.EOF
    If OptIndexador1.Value = True Then
      XLF_VALORCORRIGIDO = funCalculaCorrecaoMonetaria(XLO_TITULO, 1, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"))
    ElseIf OptIndexador2.Value = True Then
      XLF_VALORCORRIGIDO = funCalculaCorrecaoMonetaria(XLO_TITULO, 2, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"))
    Else
      XLF_VALORCORRIGIDO = XFO_titulo!titu_vl_Parcela
    End If
    
    XLF_DESAGIO = funCalculaDesagio(XLO_TITULO, XLD_DATA, XLF_VALORCORRIGIDO)
    
    XLF_VALORACUMULADO = XLF_VALORACUMULADO + XLF_VALORCORRIGIDO - XLF_DESAGIO
    
    XLO_TITULO.MoveNext
  
  Wend
  
  FunCalculaSaldoDevedor = XLF_VALORACUMULADO
  
End Function

Sub SubCalculaSeguro()
   
  Dim XLF_TAXASEGURO As Double    'Taxa de seguro
  Dim XLT_TIPOSEGURO As String    'Tipo do seguro do contrato
  Dim XLT_SQL As String           'Prepara select
  Dim XLO_SEGURO As New ADODB.Recordset
  Dim XLF_SALDODEVEDOR As Double 'Saldo Devedor
  Dim XLD_DATABASE As Date 'Data de prorrogacao do titulo escolhido para o seguro
  Dim XLT_TITULOESCOLHIDO As String 'Titulo escolhido para o calculo do seguro
  Dim XLT_TEMP As String 'Variavel temporaria
  Dim XLI_TEMP As Integer 'Variaável temporária
  Dim XLB_SAIR As Boolean 'Para sair do segundo Loop
  Dim XLF_VALOR As Double 'Armazena o valor do titulo corrigido e desagiado para subtrair do saldo devedor
  Dim XLI_CONT As Integer 'Contador para percorre o vetor
  Dim XLT_CONTRATO As String 'Código do contrato
  Dim XLB_PRIMEIRAVEZ As Boolean 'Primeira vez que entra em titulo de um determinado contrato
  Dim XLT_SEGURO As String * 1 'Se vai precisar Calcular o seguro
  Dim XLF_VALORPARCELA As Double
    
  XFO_EXPORTACAO.MoveFirst
  
  While Not XFO_EXPORTACAO.EOF
    XLB_PRIMEIRAVEZ = True
    XLT_CONTRATO = Left(XFO_EXPORTACAO!Titulo, 12)
    XLF_TAXASEGURO = XFO_EXPORTACAO!cont_vl_TaxaSegVida
    XLT_TIPOSEGURO = XFO_EXPORTACAO!cont_nr_TipoSeguro
    XLB_SAIR = False
    
    While (XLB_SAIR = False)
      'Só faz a verificação na primeira vez que entra em um contrato diferente
      If XLB_PRIMEIRAVEZ Then
        XLT_SEGURO = "N"
        If XLF_TAXASEGURO <> 0 Then 'Verifica se tem seguro
          'Verificar se o seguro já foi pago no mês
          XLT_SQL = "SELECT Titulo FROM ConsCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
                " AND empd_cd_Empreendimento= '" & Left(XFO_EXPORTACAO("Titulo"), 4) & "'" & _
                " AND imov_cd_Imovel='" & Mid(XFO_EXPORTACAO("Titulo"), 6, 4) & "'" & _
                " AND cont_cd_Contrato='" & Mid(XFO_EXPORTACAO("Titulo"), 11, 2) & "'" & _
                " AND month(titu_dt_Pagamento)=" & Month(DtpVencimento.Value) & "" & _
                " AND year(titu_dt_Pagamento)=" & Year(DtpVencimento.Value) & "" & _
                " AND titu_vl_Seguro<>0 "
          SubQOpenRecordset XLO_SEGURO, XLT_SQL, Estatico
          If XLO_SEGURO.EOF Then 'Se não encontrou registros que atendesse a condição o seguro deve ser calculado
             
             'Procurar no grid o titulo para o calculo do seguro e trazer a data
             'de seu vencimento para poder desagiar as outras parcelas
             XLT_TEMP = FunProcuraNoGrid(XFO_EXPORTACAO)
             If XLT_TEMP <> "0" Then
               XLT_TITULOESCOLHIDO = Left(XLT_TEMP, 22)
               XLD_DATABASE = CDate(Right(XLT_TEMP, 8))
               If XLT_TIPOSEGURO = 1 Then 'Seguro pelo saldo devedor
                 'Calcular o saldo devedor de todos os titulos não pagos
                 XLF_SALDODEVEDOR = FunCalculaSaldoDevedor(XFO_EXPORTACAO, XLD_DATABASE)
               End If
               XLT_SEGURO = "S"
             End If
          Else
            XLT_SEGURO = "N"
          End If
          XLO_SEGURO.Close
          Set XLO_SEGURO = Nothing
        End If
        XLB_PRIMEIRAVEZ = False
      End If
      
      ' Subtrai do saldo devedor os títulos que serão pagos
      If XLT_SEGURO = "S" Then
        If XLT_TIPOSEGURO = 1 Then
          If OptIndexador1.Value = True Then
            XLF_VALOR = funCalculaCorrecaoMonetaria(XFO_EXPORTACAO, 1, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"))
          Else
            XLF_VALOR = funCalculaCorrecaoMonetaria(XFO_EXPORTACAO, 2, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"))
          End If
          XLF_VALOR = XLF_VALOR - funCalculaDesagio(XFO_EXPORTACAO, XLD_DATABASE, XLF_VALOR)
          XLF_SALDODEVEDOR = XLF_SALDODEVEDOR - XLF_VALOR
        End If
      End If
      
      XLF_VALORPARCELA = XFO_EXPORTACAO!titu_vl_Parcela
      XFO_EXPORTACAO.MoveNext
      XLI_CONT = XLI_CONT + 1
      
      'Sai do loop quando o contrato é diferente ou chega no final do grid
      If XLI_CONT = TDBGrid1.ApproxCount Then
        XLB_SAIR = True
      Else
        If Left(VFV_VETOREXP(XLI_CONT, 0), 12) <> XLT_CONTRATO Then
          XLB_SAIR = True
        End If
      End If
    Wend
    
    If XLT_SEGURO = "S" Then
      XLI_TEMP = XLI_CONT
      XLI_CONT = XLI_CONT - 1
      
      'Preenche o Seguro no titulo escolhido
      While XLI_CONT >= 0
        If Left(VFV_VETOREXP(XLI_CONT, 0), 12) = XLT_CONTRATO Then
          If VFV_VETOREXP(XLI_CONT, 0) = XLT_TITULOESCOLHIDO Then
            If XLT_TIPOSEGURO = 1 Then
              'Grava o Seguro
              VFV_VETOREXP(XLI_CONT, 8) = XLF_TAXASEGURO * XLF_SALDODEVEDOR
              
              'Soma o valor do seguro ao campo Valor R$
              VFV_VETOREXP(XLI_CONT, 3) = VFV_VETOREXP(XLI_CONT, 3) + VFV_VETOREXP(XLI_CONT, 8)
              
              'Grava no titulo o saldo devedor de seu contrato
              VFV_VETOREXP(XLI_CONT, 9) = XLF_SALDODEVEDOR
            Else
              'Soma o valor do seguro ao campo Valor R$
              VFV_VETOREXP(XLI_CONT, 3) = VFV_VETOREXP(XLI_CONT, 3) + VFV_VETOREXP(XLI_CONT, 8)
              
              VFV_VETOREXP(XLI_CONT, 8) = XLF_TAXASEGURO * XLF_VALORPARCELA
              VFV_VETOREXP(XLI_CONT, 9) = "P" 'Marca com a letra P o saldo devedor para saber que é pela parcela
            End If
            'XLI_CONT = -1
          Else
            If XLT_TIPOSEGURO = 1 Then
              'Grava no titulo o saldo devedor de seu contrato
              VFV_VETOREXP(XLI_CONT, 9) = XLF_SALDODEVEDOR
            Else
              VFV_VETOREXP(XLI_CONT, 9) = "P" 'Marca com a letra P o saldo devedor para saber que é pela parcela
            End If
          End If
          XLI_CONT = XLI_CONT - 1
        Else
          XLI_CONT = -1
        End If
      Wend
      XLI_CONT = XLI_TEMP
    End If
  Wend
  
End Sub

Function FunProcuraNoGrid(XLO_RECORDSET As ADODB.Recordset) As String
   
  Dim XLT_SQL As String           'Prepara select
  Dim XLB_ACHEI As Boolean        'Indica se já localizou o registro no grid
  Dim XLF_VLSEGURO As Double      'Valor calculado do seguro
  Dim XLO_SEGURO As New ADODB.Recordset
  Dim XLI_CONT As Integer
        
  'Entende-se que nesse grid todos os registros serão pagos, portanto a principio os
  'registros serão apresentados na tela sem cálculo do seguros. Este será calculado quando
  'um registro for desmarcado
        
  FunProcuraNoGrid = 0
        
   'Abrir os registro (mensais) não pagos por ordem de dt. de criação
   XLT_SQL = "SELECT * FROM consCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
               " AND empd_cd_Empreendimento= '" & Left(XLO_RECORDSET("Titulo"), 4) & "'" & _
               " AND imov_cd_Imovel='" & Mid(XLO_RECORDSET("Titulo"), 6, 4) & "'" & _
               " AND cont_cd_Contrato='" & Mid(XLO_RECORDSET("Titulo"), 11, 2) & "'" & _
               " AND titu_dt_Pagamento is null " & _
               " AND month(titu_dt_Vencimento) = " & Month(DtpVencimento.Value) & _
               " AND year(titu_dt_Vencimento) = " & Year(DtpVencimento.Value) & _
               " AND tipl_tx_Descricao = 'MENSAL'" & _
               " AND Natureza <> 'FINANCIAMENTO' ORDER BY titu_dt_Criacao"
   SubQOpenRecordset XLO_SEGURO, XLT_SQL, Estatico
   'Localizar no grid e verificar se está selecionado
   If Not XLO_SEGURO.EOF Then
      While Not XLO_SEGURO.EOF And Not XLB_ACHEI
          TDBGrid1.MoveFirst
          While Not TDBGrid1.EOF And Not XLB_ACHEI
              If TDBGrid1.Columns("Título") = XLO_SEGURO!Titulo And TDBGrid1.Columns("Exp.").Value = "-1" Then
                  FunProcuraNoGrid = XLO_SEGURO!Titulo & Format(XLO_SEGURO!titu_dt_Vencimento, "dd/mm/yy")
                  XLB_ACHEI = True
              End If
              TDBGrid1.MoveNext
          Wend
          XLO_SEGURO.MoveNext
      Wend
      If XLB_ACHEI = False Then
        XLO_SEGURO.Close
        XLT_SQL = "SELECT * FROM consCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
                      " AND empd_cd_Empreendimento= '" & Left(XLO_RECORDSET("Titulo"), 4) & "'" & _
                      " AND imov_cd_Imovel='" & Mid(XLO_RECORDSET("Titulo"), 6, 4) & "'" & _
                      " AND cont_cd_Contrato='" & Mid(XLO_RECORDSET("Titulo"), 11, 2) & "'" & _
                      " AND titu_dt_Pagamento is null " & _
                      " AND month(titu_dt_Vencimento) = " & Month(DtpVencimento.Value) & _
                      " AND year(titu_dt_Vencimento) = " & Year(DtpVencimento.Value) & _
                      " AND tipl_tx_Descricao <> 'MENSAL'" & _
                      " AND Natureza <> 'FINANCIAMENTO' ORDER BY titu_dt_Criacao"
        SubQOpenRecordset XLO_SEGURO, XLT_SQL, Estatico
        
        While Not XLO_SEGURO.EOF And Not XLB_ACHEI
            TDBGrid1.MoveFirst
            While Not TDBGrid1.EOF And Not XLB_ACHEI
                If TDBGrid1.Columns("Título") = XLO_SEGURO!Titulo And TDBGrid1.Columns("Exp.").Value = "-1" Then
                    FunProcuraNoGrid = XLO_SEGURO!Titulo & Format(XLO_SEGURO!titu_dt_Vencimento, "dd/mm/yy")
                    XLB_ACHEI = True
                End If
                TDBGrid1.MoveNext
            Wend
            XLO_SEGURO.MoveNext
        Wend
      End If
       
   Else 'Abrir outro select para outros tipos de planos (<> Mensal)
       XLO_SEGURO.Close
       XLT_SQL = "SELECT * FROM consCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
                     " AND empd_cd_Empreendimento= '" & Left(XLO_RECORDSET("Titulo"), 4) & "'" & _
                     " AND imov_cd_Imovel='" & Mid(XLO_RECORDSET("Titulo"), 6, 4) & "'" & _
                     " AND cont_cd_Contrato='" & Mid(XLO_RECORDSET("Titulo"), 11, 2) & "'" & _
                     " AND titu_dt_Pagamento is null " & _
                     " AND month(titu_dt_Vencimento) = " & Month(DtpVencimento.Value) & _
                     " AND year(titu_dt_Vencimento) = " & Year(DtpVencimento.Value) & _
                     " AND Natureza <> 'FINANCIAMENTO' ORDER BY titu_dt_Criacao"
       SubQOpenRecordset XLO_SEGURO, XLT_SQL, Estatico
       
       While Not XLO_SEGURO.EOF And Not XLB_ACHEI
           TDBGrid1.MoveFirst
           While Not TDBGrid1.EOF And Not XLB_ACHEI
               If TDBGrid1.Columns("Título") = XLO_SEGURO!Titulo And TDBGrid1.Columns("Exp.").Value = "-1" Then
                   FunProcuraNoGrid = XLO_SEGURO!Titulo & Format(XLO_SEGURO!titu_dt_Vencimento, "dd/mm/yy")
                   XLB_ACHEI = True
               End If
               TDBGrid1.MoveNext
           Wend
           XLO_SEGURO.MoveNext
       Wend
    XLO_SEGURO.Close
    Set XLO_SEGURO = Nothing
  End If
  TDBGrid1.MoveFirst

End Function

Private Sub CmdDesconto_Click()
  
  PanDesconto.Left = (TDBGrid1.Width - PanDesconto.Width) / 2
  PanDesconto.Top = (TDBGrid1.Height - PanDesconto.Height) / 2
  PanDesconto.Visible = True
  TxtDesconto.Text = ""
  DtpDesconto.Value = Date
  ChkTodos.Value = 0
  
End Sub

Private Sub CmdConfirmar_Click()

  Dim XLI_CONT As Integer
  
  If ChkTodos.Value = 1 Then
    For XLI_CONT = 0 To TDBGrid1.ApproxCount - 1
      If TxtDesconto.Text = "" Then
        VFV_VETOREXP(XLI_CONT, 5) = "0,00"
        VFV_VETOREXP(XLI_CONT, 6) = ""
      Else
        VFV_VETOREXP(XLI_CONT, 5) = TxtDesconto.Text
        VFV_VETOREXP(XLI_CONT, 6) = Format(DtpDesconto.Value, "dd/mm/yy")
      End If
    Next
  Else
    If TxtDesconto.Text = "" Then
      VFV_VETOREXP(XLI_CONT, 5) = "0,00"
      VFV_VETOREXP(XLI_CONT, 6) = ""
    Else
      VFV_VETOREXP(XLI_CONT, 5) = TxtDesconto.Text
      VFV_VETOREXP(XLI_CONT, 6) = Format(DtpDesconto.Value, "dd/mm/yy")
    End If
  End If
  TDBGrid1.Refresh
  CmdRetornar_Click
  
End Sub

Private Sub CmdRetornar_Click()
  
  PanDesconto.Visible = False
  
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  
  Dim XLF_TAXASEGURO As Double    'Taxa de seguro
  Dim XLT_TIPOSEGURO As String    'Tipo do seguro do contrato
  Dim XLF_SALDODEVEDOR As Double 'Saldo Devedor
  Dim XLD_DATABASE As Date 'Data de prorrogacao do titulo escolhido para o seguro
  Dim XLT_TITULOESCOLHIDO As String 'Titulo escolhido para o calculo do seguro
  Dim XLT_TEMP As String 'Variavel temporaria
  Dim XLI_TEMP As Integer 'Variaável temporária
  Dim XLB_SAIR As Boolean 'Para sair do segundo Loop
  Dim XLF_VALOR As Double 'Armazena o valor do titulo corrigido e desagiado para subtrair do saldo devedor
  Dim XLI_CONT As Integer 'Contador para percorre o vetor
  Dim XLT_CONTRATO As String 'Código do contrato
  Dim XLB_PRIMEIRAVEZ As Boolean 'Primeira vez que entra em titulo de um determinado contrato
  Dim XLT_SEGURO As String * 1 'Se vai precisar Calcular o seguro
  Dim XLI_BOOKMARK As Integer
  Dim XLF_VALORPARCELA As Double
   
  XLI_BOOKMARK = TDBGrid1.Bookmark
  If TDBGrid1.Columns("SaldoDevedor") <> "N" Then 'Se tem que calcular seguro
  
    'Posiciona o recordset no mesmo titulo do grid
    XFO_EXPORTACAO.MoveFirst
    If TDBGrid1.Bookmark > 0 Then
      For XLI_TEMP = 1 To TDBGrid1.Bookmark
        XFO_EXPORTACAO.MoveNext
      Next
    End If
    
    XLT_CONTRATO = Left(XFO_EXPORTACAO!Titulo, 12)
    XLF_TAXASEGURO = XFO_EXPORTACAO!cont_vl_TaxaSegVida
    XLT_TIPOSEGURO = XFO_EXPORTACAO!cont_nr_TipoSeguro
    
    XLI_CONT = TDBGrid1.Bookmark
    
    'Posiciona o vetor e o recordset no primeiro titulo com este contrato
    If XLI_CONT > 0 Then
      XLB_SAIR = False
      While (XLB_SAIR = False)
        XLI_CONT = XLI_CONT - 1
        XFO_EXPORTACAO.MovePrevious
        If XLI_CONT < 0 Then
          XLI_CONT = XLI_CONT + 1
          XFO_EXPORTACAO.MoveNext
          XLB_SAIR = True
        Else
          If Left(VFV_VETOREXP(XLI_CONT, 0), 12) <> XLT_CONTRATO Then
            XLB_SAIR = True
            XLI_CONT = XLI_CONT + 1
            XFO_EXPORTACAO.MoveNext
          End If
        End If
      Wend
    End If
        
    XLB_SAIR = False
    XLB_PRIMEIRAVEZ = True
    
    While (XLB_SAIR = False)
         
      If XLB_PRIMEIRAVEZ Then
        'Procurar no grid o titulo para o calculo do seguro e trazer a data
        'de seu vencimento para poder desagiar as outras parcelas
        XLT_TEMP = FunProcuraNoGrid(XFO_EXPORTACAO)
        If XLT_TEMP <> "0" Then
          XLT_TITULOESCOLHIDO = Left(XLT_TEMP, 22)
          XLD_DATABASE = CDate(Right(XLT_TEMP, 8))
          If XLT_TIPOSEGURO = 1 Then 'Seguro pelo saldo devedor
            'Calcular o saldo devedor de todos os titulos não pagos
            XLF_SALDODEVEDOR = FunCalculaSaldoDevedor(XFO_EXPORTACAO, XLD_DATABASE)
          End If
          XLT_SEGURO = "S"
        Else
          XLT_SEGURO = "N"
        End If
      XLB_PRIMEIRAVEZ = False
      End If
    
      ' Subtrai do saldo devedor os títulos que serão pagos
      If XLT_SEGURO = "S" Then
        If VFV_VETOREXP(XLI_CONT, 1) = "-1" Then
          If XLT_TIPOSEGURO = 1 Then
            If OptIndexador1.Value = True Then
              XLF_VALOR = funCalculaCorrecaoMonetaria(XFO_EXPORTACAO, 1, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"))
            Else
              XLF_VALOR = funCalculaCorrecaoMonetaria(XFO_EXPORTACAO, 2, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"))
            End If
            XLF_VALOR = XLF_VALOR - funCalculaDesagio(XFO_EXPORTACAO, XLD_DATABASE, XLF_VALOR)
            XLF_SALDODEVEDOR = XLF_SALDODEVEDOR - XLF_VALOR
          End If
        End If
      End If
      
      'Grava o Valor R$ igual ao Val. Tit R$ para retirar o sguro embutido
      VFV_VETOREXP(XLI_CONT, 3) = VFV_VETOREXP(XLI_CONT, 7)
      
      XLF_VALORPARCELA = XFO_EXPORTACAO!titu_vl_Parcela
      XFO_EXPORTACAO.MoveNext
      XLI_CONT = XLI_CONT + 1
      
      'Sai do loop quando o contrato é diferente ou chega no final do grid
      If XLI_CONT = TDBGrid1.ApproxCount Then
        XLB_SAIR = True
      Else
        If Left(VFV_VETOREXP(XLI_CONT, 0), 12) <> XLT_CONTRATO Then
          XLB_SAIR = True
        End If
      End If
    Wend
    
    If XLT_SEGURO = "S" Then
      XLI_TEMP = XLI_CONT
      XLI_CONT = XLI_CONT - 1
      
      'Preenche o Seguro no titulo escolhido
      While XLI_CONT >= 0
        If Left(VFV_VETOREXP(XLI_CONT, 0), 12) = XLT_CONTRATO Then
          If VFV_VETOREXP(XLI_CONT, 0) = XLT_TITULOESCOLHIDO Then
            If XLT_TIPOSEGURO = 1 Then
              'Grava o Seguro
              VFV_VETOREXP(XLI_CONT, 8) = XLF_TAXASEGURO * XLF_SALDODEVEDOR
              
              'Soma o valor do seguro ao campo Valor R$
              VFV_VETOREXP(XLI_CONT, 3) = VFV_VETOREXP(XLI_CONT, 7) + VFV_VETOREXP(XLI_CONT, 8)
              
              'Grava no titulo o saldo devedor de seu contrato
              VFV_VETOREXP(XLI_CONT, 9) = XLF_SALDODEVEDOR
            Else
              
              VFV_VETOREXP(XLI_CONT, 8) = XLF_TAXASEGURO * XLF_VALORPARCELA
              
              'Soma o valor do seguro ao campo Valor R$
              VFV_VETOREXP(XLI_CONT, 3) = VFV_VETOREXP(XLI_CONT, 7) + VFV_VETOREXP(XLI_CONT, 8)
              
            End If
          Else
            If XLT_TIPOSEGURO = 1 Then
              'Grava no titulo o saldo devedor de seu contrato
              VFV_VETOREXP(XLI_CONT, 9) = XLF_SALDODEVEDOR
            End If
            VFV_VETOREXP(XLI_CONT, 8) = 0
          End If
          XLI_CONT = XLI_CONT - 1
        Else
          XLI_CONT = -1
        End If
      Wend
    End If
  End If
  TDBGrid1.Refresh
  TDBGrid1.Bookmark = XLI_BOOKMARK
                 
End Sub

Private Sub CmdPesquisa_Click()
     
  PanPesquisa.Left = (TDBGrid1.Width - PanPesquisa.Width) / 2
  PanPesquisa.Top = (TDBGrid1.Height - PanPesquisa.Height) / 2
  PanPesquisa.Visible = True
  TDBGrid1.Enabled = False
        
End Sub

Private Sub Cmdfechar_Click()
  PanPesquisa.Visible = False
End Sub

Private Sub CmdInserirEmpreendimento_Click()
       
  If Not IsEmpty(VFV_EMPREENDIMENTO1(0, 0)) Then
    subInsereRegistroVetor Array(TDBGridEmpr1.Columns("Codigo"), TDBGridEmpr1.Columns("Empreendimento")), VFV_EMPREENDIMENTO2, TDBGridEmpr2
    subRemoveRegistroVetor VFV_EMPREENDIMENTO1, TDBGridEmpr1
  End If
       
End Sub

Private Sub CmdInserirMoeda_Click()

  If Not IsEmpty(VFV_MOEDA1(0, 0)) Then
    subInsereRegistroVetor Array(TDBGridMoeda1.Columns("Codigo"), TDBGridMoeda1.Columns("Moeda")), VFV_MOEDA2, TDBGridMoeda2
    subRemoveRegistroVetor VFV_MOEDA1, TDBGridMoeda1
  End If

End Sub

Private Sub CmdInserirObservacao_Click()
 
  If Not IsEmpty(VFV_OBSERVACAO1(0, 0)) Then
    subInsereRegistroVetor Array(TDBGridObs1.Columns("Codigo"), TDBGridObs1.Columns("Observação")), VFV_OBSERVACAO2, TDBGridObs2
    subRemoveRegistroVetor VFV_OBSERVACAO1, TDBGridObs1
  End If

End Sub

Private Sub CmdInserirTodosEmpreendimento_Click()
  
  Dim XLI_CONT As Integer
  
  If Not IsEmpty(VFV_EMPREENDIMENTO1(0, 0)) Then
    XLI_CONT = TDBGridEmpr1.ApproxCount
    TDBGridEmpr1.MoveFirst
    While XLI_CONT > 0
      subInsereRegistroVetor Array(TDBGridEmpr1.Columns("Codigo"), TDBGridEmpr1.Columns("Empreendimento")), VFV_EMPREENDIMENTO2, TDBGridEmpr2
      subRemoveRegistroVetor VFV_EMPREENDIMENTO1, TDBGridEmpr1
      XLI_CONT = XLI_CONT - 1
    Wend
  End If
  
End Sub

Private Sub CmdInserirTodosMoeda_Click()
  
  Dim XLI_CONT As Integer
  
  If Not IsEmpty(VFV_MOEDA1(0, 0)) Then
    XLI_CONT = TDBGridMoeda1.ApproxCount
    TDBGridMoeda1.MoveFirst
    While XLI_CONT > 0
      subInsereRegistroVetor Array(TDBGridMoeda1.Columns("Codigo"), TDBGridMoeda1.Columns("Moeda")), VFV_MOEDA2, TDBGridMoeda2
      subRemoveRegistroVetor VFV_MOEDA1, TDBGridMoeda1
      XLI_CONT = XLI_CONT - 1
    Wend
  End If
  
End Sub

Private Sub CmdInserirTodosObservacao_Click()
  
  Dim XLI_CONT As Integer
  
  If Not IsEmpty(VFV_OBSERVACAO1(0, 0)) Then
    XLI_CONT = TDBGridObs1.ApproxCount
    TDBGridObs1.MoveFirst
    While XLI_CONT > 0
      subInsereRegistroVetor Array(TDBGridObs1.Columns("Codigo"), TDBGridObs1.Columns("Observação")), VFV_OBSERVACAO2, TDBGridObs2
      subRemoveRegistroVetor VFV_OBSERVACAO1, TDBGridObs1
      XLI_CONT = XLI_CONT - 1
    Wend
  End If
  
End Sub

Private Sub CmdRemoverEmpreendimento_Click()
     
  If Not IsEmpty(VFV_EMPREENDIMENTO2(0, 0)) Then
    subInsereRegistroVetor Array(TDBGridEmpr2.Columns("Codigo"), TDBGridEmpr2.Columns("Empreendimento")), VFV_EMPREENDIMENTO1, TDBGridEmpr1
    subRemoveRegistroVetor VFV_EMPREENDIMENTO2, TDBGridEmpr2
  End If

End Sub

Private Sub CmdRemoverMoeda_Click()

  If Not IsEmpty(VFV_MOEDA2(0, 0)) Then
    subInsereRegistroVetor Array(TDBGridMoeda2.Columns("Codigo"), TDBGridMoeda2.Columns("Moeda")), VFV_MOEDA1, TDBGridMoeda1
    subRemoveRegistroVetor VFV_MOEDA2, TDBGridMoeda2
  End If

End Sub

Private Sub CmdRemoverObservacao_Click()

  If Not IsEmpty(VFV_OBSERVACAO2(0, 0)) Then
    subInsereRegistroVetor Array(TDBGridObs2.Columns("Codigo"), TDBGridObs2.Columns("Observação")), VFV_OBSERVACAO1, TDBGridObs1
    subRemoveRegistroVetor VFV_OBSERVACAO2, TDBGridObs2
  End If

End Sub

Private Sub CmdRemoverTodosEmpreendimento_Click()
  
  Dim XLI_CONT As Integer
  
  If Not IsEmpty(VFV_EMPREENDIMENTO2(0, 0)) Then
    XLI_CONT = TDBGridEmpr2.ApproxCount
    TDBGridEmpr2.MoveFirst
    
    While XLI_CONT > 0
      subInsereRegistroVetor Array(TDBGridEmpr2.Columns("Codigo"), TDBGridEmpr2.Columns("Empreendimento")), VFV_EMPREENDIMENTO1, TDBGridEmpr1
      subRemoveRegistroVetor VFV_EMPREENDIMENTO2, TDBGridEmpr2
      XLI_CONT = XLI_CONT - 1
    Wend
  End If
    
End Sub

Private Sub CmdRemoverTodosMoeda_Click()
    
  Dim XLI_CONT As Integer
  
  If Not IsEmpty(VFV_MOEDA2(0, 0)) Then
    XLI_CONT = TDBGridMoeda2.ApproxCount
    TDBGridMoeda2.MoveFirst
    While XLI_CONT > 0
      subInsereRegistroVetor Array(TDBGridMoeda2.Columns("Codigo"), TDBGridMoeda2.Columns("Moeda")), VFV_MOEDA1, TDBGridMoeda1
      subRemoveRegistroVetor VFV_MOEDA2, TDBGridMoeda2
      XLI_CONT = XLI_CONT - 1
    Wend
  End If
  
End Sub

Private Sub CmdRemoverTodosObservacao_Click()
    
  Dim XLI_CONT As Integer
  
  If Not IsEmpty(VFV_OBSERVACAO2(0, 0)) Then
    XLI_CONT = TDBGridObs2.ApproxCount
    TDBGridObs2.MoveFirst
    While XLI_CONT > 0
      subInsereRegistroVetor Array(TDBGridObs2.Columns("Codigo"), TDBGridObs2.Columns("Observação")), VFV_OBSERVACAO1, TDBGridObs1
      subRemoveRegistroVetor VFV_OBSERVACAO2, TDBGridObs2
      XLI_CONT = XLI_CONT - 1
    Wend
  End If
  
End Sub

Private Sub CmdMarcar_Click()
    'Pré-selecionar os registros
    TDBGrid1.MoveFirst
    If Not TDBGrid1.EOF And Not TDBGrid1.BOF Then
        While Not TDBGrid1.EOF
            TDBGrid1.Columns("Exp.").Value = "S"
            TDBGrid1.MoveNext
        Wend
    End If
    TDBGrid1.Refresh
    TDBGrid1.MoveFirst
    
End Sub

Private Sub CmdDesmarcar_Click()
    'Desmarcar os registros
    TDBGrid1.MoveFirst
    If Not TDBGrid1.EOF And Not TDBGrid1.BOF Then
        While Not TDBGrid1.EOF
            TDBGrid1.Columns("Exp.").Value = "0"
            TDBGrid1.MoveNext
        Wend
    End If
    TDBGrid1.Refresh
    TDBGrid1.MoveFirst
End Sub

Private Sub CmdExportar_Click()
 
  Dim XFT_SQL As String           'Prepara Select
  Dim XLT_TEXTO As String         'Prepara Texto para gerar o arquivo
  Dim XLT_CGCEMP As String           'CGC da empresa sem máscara
  Dim XLT_CEPEMP As String           'Cep sem máscara
  Dim XLF_VALOR As Double         'Valor do registro
  Dim XLT_CGCCPF As String        'Cpf ou CGC do cliente
  Dim XLT_AGENCONTA As String
  Dim XLB_SELECIONADO As Boolean  'Indica se foi ou não selecionado algum registro
  Dim XLT_TIPO As String
  Dim XLT_NOME As String
  Dim XLT_ENDERECO As String
  Dim XLT_BAIRRO As String
  Dim XLT_CIDADE As String
  Dim XLT_CEP As String
  Dim XLT_ESTADO As String
  Dim XLT_DATAMORA As String
  Dim XLT_MENSAGEM1 As String
  Dim XLT_DATADESCONTO As String
  Dim XLF_DESCONTO As Double
  Dim XLI_NUMSQUENCIAL As Integer
  Dim XLB_VAZIO As Boolean
  Dim XLI_PRAZOMORA As Integer
  Dim XLT_CODMOEDA As String
 
  XLB_SELECIONADO = False
  TDBGrid1.MoveFirst
  XFO_EXPORTACAO.MoveFirst
  While Not TDBGrid1.EOF
    If TDBGrid1.Columns("Exp.").Value = "-1" Then
      XLB_SELECIONADO = True
      
      
      'Verifica se está preenchido o endereço correspondencial,
        'se todos os campos estão vazios
           'verifica se é pessoa jurídica e verifica o endereço comercial
           'verifica se é pessoa física e verifica o endereço residencial
        'se algum endereço está incompleto, cancela a exportacao
        
      If XFO_EXPORTACAO!focl_tx_Tipo = "J" Then
        If FunConfereCGC(Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 1, 2) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 4, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 8, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 12, 4) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 17, 2)) = False Then
          MsgBox "O CNPJ do Cliente " & XFO_EXPORTACAO!focl_tx_RazaoSocial & _
              vbCrLf & "está incorreto."
          Exit Sub
        End If
      Else
        If FunConfereCPF(Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 1, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 5, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 9, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 13, 2)) = False Then
          MsgBox "O CPF do Cliente " & XFO_EXPORTACAO!focl_tx_RazaoSocial & _
              vbCrLf & "está incorreto."
          
          Exit Sub
        End If
      End If
        
      If XFO_EXPORTACAO!clie_tx_EndCorresp = "" And XFO_EXPORTACAO!clie_tx_BairroCorresp = "" _
        And XFO_EXPORTACAO!clie_nr_CepCorresp = "  .   -   " And XFO_EXPORTACAO!clie_tx_MunCorresp = "" _
        And XFO_EXPORTACAO!clie_tx_EstCorresp = "" Then
        
        If XFO_EXPORTACAO!focl_tx_Tipo = "J" Then
          If XFO_EXPORTACAO!focl_tx_Endereco = "" Or XFO_EXPORTACAO!focl_tx_Bairro = "" _
            Or XFO_EXPORTACAO!focl_tx_Cep = "  .   -  " Or XFO_EXPORTACAO!focl_tx_Cidade = "" _
            Or XFO_EXPORTACAO!focl_tx_Estado = "" Then
            
            MsgBox "Os dados do endereço comercial do Cliente " & XFO_EXPORTACAO!focl_tx_RazaoSocial & _
              vbCrLf & "estão incompletos."
            Exit Sub
          End If
          
        Else
          If XFO_EXPORTACAO!clie_tx_EndResidencial = "" Or XFO_EXPORTACAO!clie_tx_BairroResidencial = "" _
            Or XFO_EXPORTACAO!clie_nr_CepResidencial = "  .   -   " Or XFO_EXPORTACAO!clie_tx_MunResidencial = "" _
            Or XFO_EXPORTACAO!clie_tx_EstResidencial = "" Then
            
            MsgBox "Os dados do endereço residencial do Cliente " & XFO_EXPORTACAO!focl_tx_RazaoSocial & _
              vbCrLf & "estão incompletos."
            Exit Sub
          End If
        End If
        
      ElseIf XFO_EXPORTACAO!clie_tx_EndCorresp = "" Or XFO_EXPORTACAO!clie_tx_BairroCorresp = "" _
         Or XFO_EXPORTACAO!clie_nr_CepCorresp = "  .   -   " Or XFO_EXPORTACAO!clie_tx_MunCorresp = "" _
         Or XFO_EXPORTACAO!clie_tx_EstCorresp = "" Then
        
        MsgBox "Os dados do endereço correspondencial do Cliente " & XFO_EXPORTACAO!focl_tx_RazaoSocial & _
           vbCrLf & "estão incompletos."
        Exit Sub
      End If
      
    End If
    TDBGrid1.MoveNext
    XFO_EXPORTACAO.MoveNext
  Wend

  If Not XLB_SELECIONADO Then
      MsgBox "Não existem pagamentos selecionados.", vbCritical, "ATENÇÃO"
      Exit Sub
  End If
  
  DlgSalvar.DialogTitle = "Arquivo de Exportação do HSBC"
  DlgSalvar.DefaultExt = "txt"
  DlgSalvar.InitDir = App.Path
  DlgSalvar.CancelError = False
  DlgSalvar.ShowSave
  
  If DlgSalvar.FileName = "" Then
      Exit Sub
  End If
  
  If DlgSalvar.FileName <> "" Then
      XGT_LOCALARQ = DlgSalvar.FileName
      XGT_ARQUIVO = DlgSalvar.FileTitle
  End If
  
  Open XGT_LOCALARQ For Output As #1
  
  Conexao.BeginTrans
  
  '**********HEADER DE ARQUIVO**********
  XLT_TEXTO = "0" & _
              "1" & _
              "REMESSA" & _
              "01" & _
              "COBRANCA" & Space(7) & _
              "0" & _
              Format(Left(DatContaCorrente.Recordset.Fields!coco_cd_Agencia, 4), "0000") & _
              "55" & _
              Format(Left(DatContaCorrente.Recordset.Fields!coco_cd_Agencia, 4), "0000") & Format(Left(DatContaCorrente.Recordset.Fields!coco_tx_Conta, 5), "00000") & Format(Right(FunNuloVal(DatContaCorrente.Recordset.Fields!coco_nr_Dac), 2), "00") & _
              Space(2) & _
              Mid(PEmpresa, 1, 30) & Space(30 - Len(Mid(PEmpresa, 1, 30))) & _
              Format(DatContaCorrente.Recordset.Fields!banc_cd_Codigo, "000") & _
              "HSBC" & Space(15 - Len("HSBC")) & _
              Format(Now, "ddMMyy") & _
              "01600" & _
              "BPI" & _
              Space(2) & _
              "LANCV08" & _
              Space(277) & _
              "000001"
                
  Print #1, funTiraAcento(XLT_TEXTO, True)
  
  XLI_NUMSQUENCIAL = 2
  
  XFO_EXPORTACAO.MoveFirst
  TDBGrid1.MoveFirst

  While Not TDBGrid1.EOF
      If TDBGrid1.Columns("Exp.").Value = "-1" Then
          
          XLT_CEPEMP = Mid(XGT_CEP, 1, 5) + Mid(XGT_CEP, 7, 3)
          XLT_CGCEMP = Mid(XGT_CGC, 1, 2) + Mid(XGT_CGC, 4, 3) + Mid(XGT_CGC, 8, 3) + Mid(XGT_CGC, 12, 4) + Mid(XGT_CGC, 17, 2)
                    
          'Valor já com Seguro
          XLF_VALOR = CDbl(TDBGrid1.Columns(3))
          
          If IsDate(TDBGrid1.Columns(6)) Then
            XLT_DATADESCONTO = Format(TDBGrid1.Columns(6), "ddmmyy")
          Else
            XLT_DATADESCONTO = "000000"
          End If
          
          XLF_DESCONTO = CDbl(TDBGrid1.Columns(5))
            
          If XFO_EXPORTACAO!focl_tx_Tipo = "F" Then
            XLT_TIPO = "01"
            XLT_CGCCPF = Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 1, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 5, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 9, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 13, 2)
          Else
            XLT_TIPO = "02"
            XLT_CGCCPF = Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 1, 2) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 4, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 8, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 12, 4) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 17, 2)
          End If
          
          XLT_NOME = Left(XFO_EXPORTACAO!focl_tx_RazaoSocial, 40) + Space(40 - Len(Left(XFO_EXPORTACAO!focl_tx_RazaoSocial, 40)))
          
          If XFO_EXPORTACAO!clie_tx_EndCorresp = "" And XFO_EXPORTACAO!clie_tx_BairroCorresp = "" _
            And XFO_EXPORTACAO!clie_nr_CepCorresp = "  .   -  " And XFO_EXPORTACAO!clie_tx_MunCorresp = "" _
            And XFO_EXPORTACAO!clie_tx_EstCorresp = "" Then
            
            If XLT_TIPO = "02" Then
              XLT_ENDERECO = Left(XFO_EXPORTACAO!focl_tx_Endereco, 38) + Space(38 - Len(Left(XFO_EXPORTACAO!focl_tx_Endereco, 38)))
              XLT_BAIRRO = Left(XFO_EXPORTACAO!focl_tx_Bairro, 12) + Space(12 - Len(Left(XFO_EXPORTACAO!focl_tx_Bairro, 12)))
              XLT_CEP = Left(XFO_EXPORTACAO!focl_tx_Cep, 2) + Mid(XFO_EXPORTACAO!focl_tx_Cep, 4, 3) + Right(XFO_EXPORTACAO!focl_tx_Cep, 3)
              XLT_CIDADE = Left(XFO_EXPORTACAO!focl_tx_Cidade, 15) + Space(15 - Len(Left(XFO_EXPORTACAO!focl_tx_Cidade, 15)))
              XLT_ESTADO = Left(XFO_EXPORTACAO!focl_tx_Estado, 2) + Space(2 - Len(Left(XFO_EXPORTACAO!focl_tx_Estado, 2)))
            Else
              XLT_ENDERECO = Left(XFO_EXPORTACAO!clie_tx_EndResidencial, 38) + Space(38 - Len(Left(XFO_EXPORTACAO!clie_tx_EndResidencial, 38)))
              XLT_BAIRRO = Left(XFO_EXPORTACAO!clie_tx_BairroResidencial, 12) + Space(12 - Len(Left(XFO_EXPORTACAO!clie_tx_BairroResidencial, 12)))
              XLT_CEP = Left(XFO_EXPORTACAO!clie_nr_CepResidencial, 2) + Mid(XFO_EXPORTACAO!clie_nr_CepResidencial, 4, 3) + Right(XFO_EXPORTACAO!clie_nr_CepResidencial, 3)
              XLT_CIDADE = Left(XFO_EXPORTACAO!clie_tx_MunResidencial, 15) + Space(15 - Len(Left(XFO_EXPORTACAO!clie_tx_MunResidencial, 15)))
              XLT_ESTADO = Left(XFO_EXPORTACAO!clie_tx_EstResidencial, 2) + Space(2 - Len(Left(XFO_EXPORTACAO!clie_tx_EstResidencial, 2)))
            End If
            
          Else
            XLT_ENDERECO = Left(XFO_EXPORTACAO!clie_tx_EndCorresp, 38) + Space(38 - Len(Left(XFO_EXPORTACAO!clie_tx_EndCorresp, 38)))
            XLT_BAIRRO = Left(XFO_EXPORTACAO!clie_tx_BairroCorresp, 12) + Space(12 - Len(Left(XFO_EXPORTACAO!clie_tx_BairroCorresp, 12)))
            XLT_CEP = Left(XFO_EXPORTACAO!clie_nr_CepCorresp, 2) + Mid(XFO_EXPORTACAO!clie_nr_CepCorresp, 4, 3) + Right(XFO_EXPORTACAO!clie_nr_CepCorresp, 3)
            XLT_CIDADE = Left(XFO_EXPORTACAO!clie_tx_MunCorresp, 15) + Space(15 - Len(Left(XFO_EXPORTACAO!clie_tx_MunCorresp, 15)))
            XLT_ESTADO = Left(XFO_EXPORTACAO!clie_tx_EstCorresp, 2) + Space(2 - Len(Left(XFO_EXPORTACAO!clie_tx_EstCorresp, 2)))
          End If
            
          '*****************HEADER DE LOTE*****************
          XLT_TEXTO = "1" & _
                      "02" & _
                      Format(XLT_CGCEMP, "00000000000000") & _
                      "0" & _
                      Format(Left(DatContaCorrente.Recordset.Fields!coco_cd_Agencia, 4), "0000") & _
                      "55" & _
                      Format(Left(DatContaCorrente.Recordset.Fields!coco_cd_Agencia, 4), "0000") & Format(Left(DatContaCorrente.Recordset.Fields!coco_tx_Conta, 5), "00000") & Format(Right(FunNuloVal(DatContaCorrente.Recordset.Fields!coco_nr_Dac), 2), "00") & _
                      Space(2) & _
                      XFO_EXPORTACAO!Titulo & Space(3) & _
                      "00000000000" & _
                      "000000" & _
                      "00000000000" & _
                      "000000" & _
                      "00000000000" & _
                      "1" & _
                      "01" & _
                      Left(XFO_EXPORTACAO!Titulo, 4) & Mid(XFO_EXPORTACAO!Titulo, 6, 4) & Mid(XFO_EXPORTACAO!Titulo, 11, 2) & _
                      Format(XFO_EXPORTACAO!titu_dt_Vencimento, "ddmmyy") & _
                      Format(XLF_VALOR * 100, "0000000000000") & _
                      Format(DatContaCorrente.Recordset.Fields!banc_cd_Codigo, "000") & _
                      "00000"

          XLT_TEXTO = XLT_TEXTO & _
                      "09" & _
                      "N" & _
                      Format(DtpExportacao.Value, "ddmmyy") & _
                      Format(FunNuloVal(TxtInstrucao1.Text), "00") & _
                      Format(FunNuloVal(TxtInstrucao2.Text), "00") & _
                      "0000000000000" & _
                      XLT_DATADESCONTO & _
                      Format(XLF_DESCONTO * 100, "0000000000000") & _
                      "0000000000000" & _
                      "0000000000000" & _
                      XLT_TIPO & _
                      Format(XLT_CGCCPF, "00000000000000") & _
                      XLT_NOME & _
                      XLT_ENDERECO & _
                      Space(2) & _
                      XLT_BAIRRO & _
                      XLT_CEP & _
                      XLT_CIDADE & _
                      XLT_ESTADO & _
                      Space(39) & _
                      "E" & _
                      Space(2) & _
                      "9" & _
                      Format(XLI_NUMSQUENCIAL, "000000")
                      
          Print #1, funTiraAcento(XLT_TEXTO, True)
          
          If OptIndexador1.Value = True Then
            If Not IsNull(XFO_EXPORTACAO!moed_cd_Moeda1) Then
              XLT_CODMOEDA = XFO_EXPORTACAO!moed_cd_Moeda1
            Else
              XLT_CODMOEDA = ""
            End If
          ElseIf OptIndexador2.Value = True Then
            If Not IsNull(XFO_EXPORTACAO!moed_cd_Moeda2) Then
              XLT_CODMOEDA = XFO_EXPORTACAO!moed_cd_Moeda2
            Else
              XLT_CODMOEDA = ""
            End If
          Else
            XLT_CODMOEDA = ""
          End If
            
          Conexao.Execute ("UPDATE Titulos Set titu_tx_ExpBanco='S'," & _
            " moed_cd_Moeda3=" & FunNuloBancoVal(XLT_CODMOEDA) & "," & _
            " titu_vl_Seguro=" & FunNuloVal(FunTrataFloat(TDBGrid1.Columns(8))) & _
            " WHERE empr_cd_Empresa=" & PCodEmpresa & _
            " AND empd_cd_empreendimento='" & Left(XFO_EXPORTACAO!Titulo, 4) & "'" & _
            " AND imov_cd_imovel='" & Mid(XFO_EXPORTACAO!Titulo, 6, 4) & "'" & _
            " AND cont_cd_Contrato='" & Mid(XFO_EXPORTACAO!Titulo, 11, 2) & "'" & _
            " AND titu_cd_Plano='" & Mid(XFO_EXPORTACAO!Titulo, 14, 2) & "'" & _
            " AND titu_cd_Parcela='" & Mid(XFO_EXPORTACAO!Titulo, 17, 3) & "'" & _
            " AND titu_cd_Residuo='" & Right(XFO_EXPORTACAO!Titulo, 2) & "'")
            
        
        '******************REGISTRA LOG DA OPERAÇÃO**************************
        ReDim XGM_MATRIZLOG(8, 2) As Variant
        
        XGM_MATRIZLOG(0, 0) = "Título"
        XGM_MATRIZLOG(1, 0) = "Moeda"
        XGM_MATRIZLOG(2, 0) = "Valor Seguro"
        XGM_MATRIZLOG(3, 0) = "Valor do título"
        XGM_MATRIZLOG(4, 0) = "Valor do desconto"
        XGM_MATRIZLOG(5, 0) = "Conta corrente"
        XGM_MATRIZLOG(6, 0) = "Data Venc. Título"
        XGM_MATRIZLOG(7, 0) = "Data Venc. Desconto"
        
        XGM_MATRIZLOG(0, 1) = XFO_EXPORTACAO!Titulo
        XGM_MATRIZLOG(1, 1) = FunNuloBancoVal(XLT_CODMOEDA)
        XGM_MATRIZLOG(2, 1) = FunNuloVal(FunTrataFloat(TDBGrid1.Columns(8)))
        XGM_MATRIZLOG(3, 1) = XLF_VALOR
        XGM_MATRIZLOG(4, 1) = XLF_DESCONTO
        XGM_MATRIZLOG(5, 1) = DatContaCorrente.Recordset.Fields!coco_cd_Agencia & " - " & DatContaCorrente.Recordset.Fields!coco_tx_Conta & "-" & DatContaCorrente.Recordset.Fields!coco_nr_Dac
        XGM_MATRIZLOG(6, 1) = XFO_EXPORTACAO!titu_dt_Vencimento
        XGM_MATRIZLOG(7, 1) = XLT_DATADESCONTO
        
        Call subRegistraLog("TelaExpBancaria", "1", funCriaDescricaoLog(XGM_MATRIZLOG, EXPORTACAO_CAPI, PAGAMENTO_ELETRONICO_HSBC))
        '*******************************************************************
        
        XLI_NUMSQUENCIAL = XLI_NUMSQUENCIAL + 1
      End If
      
      TDBGrid1.MoveNext
      XFO_EXPORTACAO.MoveNext
  Wend
  
teste:
  
  '******TRAILER DE ARQUIVO*******
  XLT_TEXTO = "9" & _
              Space(393) & _
              Format(XLI_NUMSQUENCIAL, "000000")
              
  Print #1, funTiraAcento(XLT_TEXTO, True)
  '**********************************************
  Close #1
         
  MsgBox "Exportação realizada com sucesso!", vbInformation + vbOKOnly, "ATENÇÃO"
  
  Conexao.CommitTrans
  subDesabilitaBotoes
  CmdImprimir.Enabled = True
  Exit Sub
  
  
End Sub

Private Sub CmdFiltro_Click()
    
    Dim XLI_POS As Integer
    Dim XLI_POS2 As Integer
    
    Set Formulario = TelaExpBancaria
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"

    If Filtrou = True Then
        subCarregaVetor FRM_FiltroAtual, Array("Titulo", "Exporta", "titu_dt_Vencimento", "ValorReal", _
              "obse_tx_Observacao", "titu_vl_Desconto", "LimiteDesconto", "ValorTitulo", _
              "Seguro", "SaldoDevedor"), VFV_VETOREXP, TDBGrid1
              
        'Retira o último filtro adicionado no sql do recordset
        XLI_POS = InStrRev(FiltroAtual, "WHERE", -1, vbTextCompare)
        XLI_POS2 = InStrRev(XFT_SQL2, "WHERE", -1, vbTextCompare)
        
        XFT_SQL2 = Left(XFT_SQL2, XLI_POS2 - 1) & Right(FiltroAtual, Len(FiltroAtual) - XLI_POS + 1)
        
        If XFO_EXPORTACAO.State = adStateOpen Then
          XFO_EXPORTACAO.Close
        End If
        SubQOpenRecordset XFO_EXPORTACAO, XFT_SQL2, Estatico
        
        subTelaValoresGlobais "G"
        
         SubCalculaValorAtualizado

         SubCalculaSeguro

         TDBGrid1.Refresh
    
    End If
    
    subHabilitaBotoes
  
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub

Private Sub CmdImprimir_Click()
    'Call subImprimeListagemGRIDUnBound(2, TDBGrid1.PrintInfo, "Listagem de Contas a Pagar")
    
    On Error GoTo TrataErro:
    Dim GuardaY As Single
    Dim LargPapel As Single, AltPapel As Single, AreaImpressao As Single
    Dim TitRel As String
    Dim RecebeCampo As Variant
    Dim GuardaX As Single, GuardaYInicial As Single
    Dim QuebraY As Single
    Dim ImprimirLinha As Boolean
    Dim XLF_VALOR As Double
    Dim XLF_TOTALPAGAR As Double
    Dim XLI_QTDTITULOS As Integer
        
    TitRel = "Listagem do Arquivo de Exportação"
    
    Printer.ScaleMode = 7  'vbCentimeters
    
    Printer.Orientation = 1  'Portrait
    LargPapel = Printer.Width / 567   'Convertendo Twips para Cms
    AltPapel = Printer.Height / 567   'Convertendo Twips para Cms
    
    subRodape 2, AltPapel, TitRel
    subCabecalhoListagemRelatorio TitRel
    GuardaY = Printer.CurrentY
    GuardaYInicial = Printer.CurrentY
    
    AreaImpressao = AltPapel - AltRodape      ' Variável setada em Rodape()
    
    GuardaY = Printer.CurrentY + 0.3
    Printer.CurrentY = GuardaY
    GuardaY = Printer.CurrentY + 0.2
    subImprimeLinha 1.5, 20, 3
    subImprimeTexto "TÍTULO", 1.5, GuardaY, 7, "Arial", True, False
    subImprimeTexto "NOME DO CLIENTE", 5.2, GuardaY, 7, "Arial", True, False
    subImprimeTexto "VENCTO.", 10.2, GuardaY, 7, "Arial", True, False
    subImprimeTexto "VALOR ORIGINAL", 11.7, GuardaY, 7, "Arial", True, False
    subImprimeTexto "DESCONTO", 14.2, GuardaY, 7, "Arial", True, False
    subImprimeTexto "SEGURO", 16.2, GuardaY, 7, "Arial", True, False
    subImprimeTexto "VALOR A PAGAR", 17.7, GuardaY, 7, "Arial", True, False
    subImprimeLinha 11.6, 15.7, 1
    subImprimeLinha 15.9, 20, 1
    GuardaY = Printer.CurrentY + 0.2
    subImprimeLinha 1.5, 20, 3
    Printer.Print ""
    
    TDBGrid1.MoveFirst
    
    XLI_QTDTITULOS = 0
    XLF_VALOR = 0
    XLF_TOTALPAGAR = 0
    
    If Not (TDBGrid1.EOF And TDBGrid1.BOF) Then
        While Not (TDBGrid1.EOF)
            If (TDBGrid1.Columns("Exp.").Value = -1) Then
                If Printer.CurrentY >= (AreaImpressao - 0.35) Then
                    Printer.NewPage
                    subRodape 2, AltPapel, TitRel
                    subCabecalhoListagemRelatorio TitRel
                    
                    subImprimeLinha 1.5, 20, 3
                    GuardaY = Printer.CurrentY + 0.1
                    subImprimeTexto "TÍTULO", 1.5, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "NOME DO CLIENTE", 5.2, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "VENCTO.", 10.2, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "VALOR ORIGINAL", 11.7, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "DESCONTO", 14.2, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "SEGURO", 16.2, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "VALOR A PAGAR", 17.7, GuardaY, 7, "Arial", True, False
                    subImprimeLinha 11.6, 15.7, 1
                    subImprimeLinha 15.9, 20, 1
                    GuardaY = Printer.CurrentY + 0.1
                    subImprimeLinha 1.5, 20, 3
                    Printer.Print ""
                End If
                
                GuardaY = Printer.CurrentY
                RecebeCampo = TDBGrid1.Columns("Título")
                subQuebraTexto FunNulo(RecebeCampo), 1.5, 3.5, 7, "Arial", False, False
                QuebraY = Printer.CurrentY
                   
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Nome")
                subQuebraTexto FunNulo(RecebeCampo), 5.2, 4.8, 7, "Arial", False, False
                If QuebraY < Printer.CurrentY Then
                    QuebraY = Printer.CurrentY
                End If
                
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Vencimento")
                subQuebraTexto FunNulo(RecebeCampo), 10.2, 1.3, 7, "Arial", False, False
                
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Val. Tit. R$")
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 11.55, 2.3, 7, "Arial", False, False), 2.3, 7, "Arial", False, False
                
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Desconto")
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 13.8, 1.8, 7, "Arial", False, False), 1.8, 7, "Arial", False, False
                
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Val. Seguro R$")
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 15.95, 1.3, 7, "Arial", False, False), 1.3, 7, "Arial", False, False
               
                Printer.CurrentY = GuardaY
                RecebeCampo = Format(CStr(CCur(TDBGrid1.Columns("Val. Tit. R$")) + CCur(TDBGrid1.Columns("Val. Seguro R$")) - CCur(TDBGrid1.Columns("Desconto"))), "standard")
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 17.3, 2.5, 7, "Arial", False, False), 2.5, 7, "Arial", False, False
                
                Printer.CurrentY = QuebraY + 0.2
                
                XLF_VALOR = XLF_VALOR + CCur(TDBGrid1.Columns("Val. Tit. R$"))
                XLF_TOTALPAGAR = XLF_TOTALPAGAR + CCur(TDBGrid1.Columns("Val. Tit. R$")) + CCur(TDBGrid1.Columns("Val. Seguro R$")) - CCur(TDBGrid1.Columns("Desconto"))
                XLI_QTDTITULOS = XLI_QTDTITULOS + 1
            End If
            TDBGrid1.MoveNext
        Wend
        GuardaY = Printer.CurrentY + 0.4
        subImprimeLinha 1.5, 20, 3
        GuardaY = Printer.CurrentY + 0.2
        Printer.CurrentY = GuardaY
        RecebeCampo = "TOTAL: "
        subQuebraTexto FunNulo(RecebeCampo), 1.5, 1.6, 7, "Arial", True, False
        QuebraY = Printer.CurrentY

        Printer.CurrentY = GuardaY
        RecebeCampo = Format$(XLF_VALOR, "STANDARD")
        subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 11.55, 2.3, 7, "Arial", False, False), 2.3, 7, "Arial", True, False
        If QuebraY < Printer.CurrentY Then
            QuebraY = Printer.CurrentY
        End If
    
        Printer.CurrentY = GuardaY
        RecebeCampo = Format$(XLF_TOTALPAGAR, "STANDARD")
        subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 17.5, 2.3, 7, "Arial", False, False), 2.3, 7, "Arial", True, False
        If QuebraY < Printer.CurrentY Then
            QuebraY = Printer.CurrentY
        End If
        GuardaY = Printer.CurrentY + 0.3
        subImprimeTexto "TOTAL DE TÍTULOS:", 1.5, GuardaY, 7, "Arial", True, False
        
        Printer.CurrentY = GuardaY
        RecebeCampo = Format(XLI_QTDTITULOS, "000")
        subQuebraTexto FunNulo(RecebeCampo), 4, 4, 7, "Arial", True, False
        If QuebraY < Printer.CurrentY Then
            QuebraY = Printer.CurrentY
        End If
        GuardaY = Printer.CurrentY + 1#
        Printer.CurrentY = GuardaY
        
        TDBGrid1.MoveFirst
    Else
        Printer.Print ""
    End If
    Printer.EndDoc
   
Exit Sub

TrataErro:
    funTrataErros (ComMensagem)
    Exit Sub
    
End Sub

Private Sub CmdPesquisar_Click()
  Dim XGT_SELECAO As String  'Armazena condiçõe para o filtro
  Dim XFI_CONT As Integer
  Dim XLT_NATUREZA As String

  XGT_SELECAO = ""
  XGT_CONJUNCAO = ""
  
  If FunObrigatorioCBO(CboCCorrente, "Selecione uma Conta Corrente.") Then Exit Sub

  If Not IsNull(DtpVencimento.Value) Then
     XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      " month(titu_dt_Vencimento) =" & Month(DtpVencimento.Value) & _
      " AND year(titu_dt_Vencimento) = " & Year(DtpVencimento.Value)
      XGT_CONJUNCAO = " AND "
  End If
  
  If Not IsEmpty(VFV_EMPREENDIMENTO2(0, 0)) Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "("
    For XFI_CONT = 0 To TDBGridEmpr2.ApproxCount - 1
       If XFI_CONT > 0 Then
        XGT_SELECAO = XGT_SELECAO & " OR "
      End If
      XGT_SELECAO = XGT_SELECAO & " empd_cd_Empreendimento='" & VFV_EMPREENDIMENTO2(XFI_CONT, 0) & "'"
    Next
    XGT_SELECAO = XGT_SELECAO & ")"
    XGT_CONJUNCAO = " AND "
  End If
  
  If OptValorHistorico.Value = False Then
    If Not IsEmpty(VFV_MOEDA2(0, 0)) Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "("
      For XFI_CONT = 0 To TDBGridMoeda2.ApproxCount - 1
         If XFI_CONT > 0 Then
          XGT_SELECAO = XGT_SELECAO & " OR "
        End If
        If OptIndexador1.Value = True Then
          XGT_SELECAO = XGT_SELECAO & " moed_cd_Moeda1='" & VFV_MOEDA2(XFI_CONT, 0) & "'"
        Else
          XGT_SELECAO = XGT_SELECAO & " moed_cd_Moeda2='" & VFV_MOEDA2(XFI_CONT, 0) & "'"
        End If
      Next
      XGT_SELECAO = XGT_SELECAO & ")"
      XGT_CONJUNCAO = " AND "
    End If
  End If
  
  If Not IsEmpty(VFV_OBSERVACAO2(0, 0)) Then
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "("
    For XFI_CONT = 0 To TDBGridObs2.ApproxCount - 1
       If XFI_CONT > 0 Then
        XGT_SELECAO = XGT_SELECAO & " OR "
      End If
      XGT_SELECAO = XGT_SELECAO & " obse_cd_Observacao='" & VFV_OBSERVACAO2(XFI_CONT, 0) & "'"
    Next
    XGT_SELECAO = XGT_SELECAO & ")"
    XGT_CONJUNCAO = " AND "
  End If
  
  'Tipo Plano
  If CboTiposPlanos.Text <> "" Then
     XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
       " tipl_cd_TipoPlano =" & CboTiposPlanos.BoundText
  End If
 
  'Natureza
   If CboNaturezasPlanos.Text <> "" Then
     Select Case CboNaturezasPlanos.ListIndex
         Case 0
              XLT_NATUREZA = "S"
         Case 1
              XLT_NATUREZA = "P"
         Case 2
              XLT_NATUREZA = "F"
         Case 3
              XLT_NATUREZA = "G"
         Case 4
              XLT_NATUREZA = "O"
     End Select
     
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        " titu_tx_NaturezaPlano = '" & XLT_NATUREZA & "'"
      XGT_CONJUNCAO = " AND "
  End If
  
  'Filtra a Conta Corrente
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
     " coco_cd_codigo = " & CboCCorrente.BoundText
     '" coco_cd_codigo =69 "
  XGT_CONJUNCAO = " AND "
     
  'Filtra os titulos que não foram exportados
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
     " titu_tx_ExpBanco = 'N'"
  XGT_CONJUNCAO = " AND "
     
  'Filtra os titulos que não foram pagos
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
     " titu_dt_Pagamento is null "
  XGT_CONJUNCAO = " AND "
     
  'Filtra os ativos
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
     " cont_tx_Status='A' "
  XGT_CONJUNCAO = " AND "
    
  XFT_SQL = "SELECT *, " & _
                  "-1 as Exporta," & _
                  "0 as ValorReal," & _
                  "'' as LimiteDesconto," & _
                  "0 as ValorTitulo, " & _
                  "0 as Seguro, " & _
                  "'N' as SaldoDevedor " & _
              "FROM ConsCAPExpBanco " & _
              "WHERE" & XGT_SELECAO
     
  subCarregaVetor XFT_SQL, Array("Titulo", "Exporta", "titu_dt_Vencimento", "ValorReal", _
     "focl_tx_RazaoSocial", "titu_vl_Desconto", "LimiteDesconto", "ValorTitulo", _
     "Seguro", "SaldoDevedor"), VFV_VETOREXP, TDBGrid1
     
  If VFV_VETOREXP(0, 0) <> Empty Then
    subHabilitaBotoes
    TDBGrid1.Enabled = True
    TDBGrid1.MoveFirst
  Else
    MsgBox "Não existem pagamentos pendentes com esses filtros!", vbInformation + vbOKOnly, "ATENÇÃO"
    subDesabilitaBotoes
    TDBGrid1.Enabled = False
    VFV_VETOREXP.Clear
    Exit Sub
  End If
     
  TDBGrid1.Refresh
  
  XFT_SQL2 = FunCriaConsultaBase("01/01/01", NomeSgbd, "ConsCAPExpBanco", 3)
  XFT_SQL2 = XFT_SQL2 & XGT_SELECAO

  If XFO_EXPORTACAO.State = adStateOpen Then
    XFO_EXPORTACAO.Close
  End If
  SubQOpenRecordset XFO_EXPORTACAO, XFT_SQL2, Estatico
  
  subCarregaVariaveis XFT_SQL
  
  SubCalculaValorAtualizado
  
  SubCalculaSeguro
    
  TDBGrid1.Refresh

  CmdExportar.Enabled = True
  CmdDesconto.Enabled = True
  filtra_especial = True
  subTelaValoresGlobais "G"
  Cmdfechar_Click
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub dtpFim_Change()
    CmdPesquisar.Enabled = True
End Sub

Private Sub DtpVencimento_Change()
    CmdPesquisar.Enabled = True
End Sub

Private Sub TDBGrid1_ComboSelect(ByVal ColIndex As Integer)
    TDBGrid1.EditActive = False
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
    If TDBGrid1.Columns("Tipo").CellText(Bookmark) = "A Pagar" Then
        RowStyle.ForeColor = &HFF&
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    Dim XLI_POS As Integer
    
    If Button = 2 Then   'Verifica se o botão da direita foi pressionado
        If TDBGrid1.ColContaining(X) = 1 Or TDBGrid1.ColContaining(X) = 3 Or TDBGrid1.ColContaining(X) = 5 _
           Or TDBGrid1.ColContaining(X) = 6 Or TDBGrid1.ColContaining(X) = 7 Or TDBGrid1.ColContaining(X) = 8 Then
            MsgBox "Esta coluna não pode ser filtrada."
        Else
            Set Formulario = TelaExpBancaria
            subTelaValoresGlobais "P"
        
            FunExecutaFiltroUnbound Array("Titulo", "Exporta", "titu_dt_Vencimento", "ValorReal", _
              "obse_tx_Observacao", "titu_vl_Desconto", "LimiteDesconto", "ValorTitulo", _
              "Seguro", "SaldoDevedor"), VFV_VETOREXP, TDBGrid1, Formulario, X, XFT_SQL
            
            subHabilitaBotoes
            subTelaValoresGlobais "G"
            
            If VFV_VETOREXP(0, 0) <> Empty Then
              subHabilitaBotoes
              TDBGrid1.Enabled = True
              TDBGrid1.MoveFirst
            Else
              MsgBox "Não existem pagamentos pendentes com esses filtros!", vbInformation + vbOKOnly, "ATENÇÃO"
              CmdDesmarcar.Enabled = False
              CmdMarcar.Enabled = False
              CmdExportar.Enabled = False
              CmdDesconto.Enabled = False
              CmdImprimir.Enabled = False
              TDBGrid1.Enabled = False
              VFV_VETOREXP.Clear
              Exit Sub
            End If
            
            'Procura o filtro adicionado no sql do vetor e copia para o sql do recordset
            XLI_POS = InStrRev(FiltroAtual, "AND", -1, vbTextCompare)
            XFT_SQL2 = XFT_SQL2 & Right(FiltroAtual, Len(FiltroAtual) - XLI_POS + 1)
            If XFO_EXPORTACAO.State = adStateOpen Then
              XFO_EXPORTACAO.Close
            End If
            SubQOpenRecordset XFO_EXPORTACAO, XFT_SQL2, Estatico
            
            SubCalculaValorAtualizado
  
            SubCalculaSeguro
    
            TDBGrid1.Refresh
        End If
    End If
    
End Sub

Private Sub Form_Activate()
    Set Formulario = TelaExpBancaria
    
    If CarregouGrid Then
        If Not PrimeiraVez Then
            subTelaValoresGlobais "P"
        End If

        PrimeiraVez = False
    End If
 
End Sub

Private Sub Form_Load()
    
    Set Formulario = TelaExpBancaria
    
    Call AjustaTela
        
    subCarregaVetor "SELECT empd_cd_Empreendimento,empd_tx_Nome FROM Empreendimentos WHERE empr_cd_empresa=" & PCodEmpresa, Array("empd_cd_Empreendimento", "empd_tx_Nome"), VFV_EMPREENDIMENTO1, TDBGridEmpr1
    subCarregaVetor "SELECT moed_cd_CodMoeda,moed_tx_Descricao FROM Moedas", Array("moed_cd_CodMoeda", "moed_tx_Descricao"), VFV_MOEDA1, TDBGridMoeda1
    subCarregaVetor "SELECT * FROM Observacoes", Array("obse_cd_Observacao", "obse_tx_Observacao"), VFV_OBSERVACAO1, TDBGridObs1
    
    VFV_EMPREENDIMENTO2.ReDim 0, 0, 0, TDBGridEmpr2.Columns.Count
    TDBGridEmpr2.Array = VFV_EMPREENDIMENTO2
    TDBGridEmpr2.ReBind
    
    VFV_MOEDA2.ReDim 0, 0, 0, TDBGridMoeda2.Columns.Count
    TDBGridMoeda2.Array = VFV_MOEDA2
    TDBGridMoeda2.ReBind
    
    VFV_OBSERVACAO2.ReDim 0, 0, 0, TDBGridObs2.Columns.Count
    TDBGridObs2.Array = VFV_OBSERVACAO2
    TDBGridObs2.ReBind
    
    subConectarControleDadosNV DatTipoPlano, "SELECT * FROM TiposPlanos ORDER BY tipl_tx_Descricao", Estatico
    subConectarControleDadosNV DatObs, "SELECT * FROM Observacoes ORDER BY obse_tx_observacao", Estatico
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & Int(PCodEmpresa) & " AND banc_cd_codigo=399 ORDER BY coco_tx_Descricao", Estatico
    
    PanPesquisa.Left = (TDBGrid1.Width - PanPesquisa.Width) / 2
    PanPesquisa.Top = (TDBGrid1.Height - PanPesquisa.Height) / 2
    
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
        
    subDesabilitaBotoes
    TDBGrid1.Enabled = False
    DtpVencimento.Value = Date
    DtpExportacao.Value = Date
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub CboCCorrente_Change()
  
  If CboCCorrente.BoundText <> "" Then
    DatContaCorrente.Recordset.Bookmark = CboCCorrente.SelectedItem
    TxtInstrucao1.Text = FunNulo(DatContaCorrente.Recordset.Fields("coco_tx_Instrucao1"))
    TxtInstrucao2.Text = FunNulo(DatContaCorrente.Recordset.Fields("coco_tx_Instrucao2"))
  End If
  
End Sub

Private Sub TxtDesconto_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 46 Then KeyAscii = 44
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
    KeyAscii = 0
  End If
  
End Sub

Private Sub TxtDesconto_LostFocus()
  
  If Not IsNumeric(TxtDesconto.Text) Then
    TxtDesconto.Text = ""
  End If
  
End Sub

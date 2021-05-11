VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form TelaExpBancaria 
   Caption         =   "Exportação Bancária - CAIXA"
   ClientHeight    =   5790
   ClientLeft      =   75
   ClientTop       =   1470
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   10095
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel PanPesquisa 
      Height          =   4635
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   9135
      _Version        =   65536
      _ExtentX        =   16113
      _ExtentY        =   8176
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
         Left            =   7320
         TabIndex        =   3
         Top             =   4200
         Width           =   1245
      End
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
         Left            =   6000
         TabIndex        =   2
         Top             =   4200
         Width           =   1245
      End
      Begin MSComDlg.CommonDialog DlgSalvar 
         Left            =   240
         Top             =   3840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4635
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   8176
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Dados Gerais"
         TabPicture(0)   =   "TelaExpBancaria.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LblMensagem1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label6"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "LblMesAno"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "LblDtExportacao"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "LblTipoPlano"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "LblNaturezaPlano"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "LblCCorrente"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "LblInstrucao2(1)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "LblInstrucao1(0)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "DatObs"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "DatTipoPlano"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "DatContaCorrente"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "CboCCorrente"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "CmdLimparNaturezaPlano"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "CboTiposPlanos"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "CmdLimparTipoPlano"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "DtpExportacao"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "DtpVencimento"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "CmbInstrucao"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Frame1"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "CboTipoCobranca"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "CboNaturezasPlanos"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "TxtMensagem2"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "FraDesagio"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "TxtMensagem1"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).ControlCount=   25
         TabCaption(1)   =   "Empreendimentos"
         TabPicture(1)   =   "TelaExpBancaria.frx":001C
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
         TabPicture(2)   =   "TelaExpBancaria.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label3"
         Tab(2).Control(1)=   "Label2"
         Tab(2).Control(2)=   "TDBGridMoeda2"
         Tab(2).Control(3)=   "TDBGridMoeda1"
         Tab(2).Control(4)=   "FraCorrecao"
         Tab(2).Control(5)=   "CmdInserirTodosMoeda"
         Tab(2).Control(6)=   "CmdRemoverTodosMoeda"
         Tab(2).Control(7)=   "CmdInserirMoeda"
         Tab(2).Control(8)=   "CmdRemoverMoeda"
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "Observações"
         TabPicture(3)   =   "TelaExpBancaria.frx":0054
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
         Begin VB.TextBox TxtMensagem1 
            Height          =   315
            Left            =   1965
            MaxLength       =   12
            TabIndex        =   30
            Text            =   "02"
            ToolTipText     =   $"TelaExpBancaria.frx":0070
            Top             =   2100
            Width           =   6780
         End
         Begin VB.Frame FraDesagio 
            Height          =   615
            Left            =   1200
            TabIndex        =   28
            Top             =   3720
            Width           =   3345
            Begin VB.CheckBox ChkJuros 
               Alignment       =   1  'Right Justify
               Caption         =   "Deseja cobrar juros de Mora?"
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
               Left            =   195
               TabIndex        =   29
               Top             =   195
               Value           =   1  'Checked
               Width           =   2850
            End
         End
         Begin VB.TextBox TxtMensagem2 
            Height          =   315
            Left            =   1965
            MaxLength       =   60
            TabIndex        =   27
            ToolTipText     =   "Mensagem a ser impressa no Boleto."
            Top             =   2475
            Width           =   6810
         End
         Begin VB.ComboBox CboNaturezasPlanos 
            Height          =   315
            ItemData        =   "TelaExpBancaria.frx":00FD
            Left            =   1965
            List            =   "TelaExpBancaria.frx":0113
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1245
            Width           =   3165
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
            TabIndex        =   22
            Top             =   420
            Width           =   4995
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
               TabIndex        =   25
               Top             =   285
               Width           =   1575
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
               TabIndex        =   24
               Top             =   225
               Value           =   -1  'True
               Width           =   1335
            End
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
               TabIndex        =   23
               Top             =   225
               Width           =   1395
            End
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
            TabIndex        =   20
            Top             =   3000
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
            TabIndex        =   19
            Top             =   1560
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
            TabIndex        =   18
            Top             =   2040
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
            TabIndex        =   17
            Top             =   2280
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
            TabIndex        =   16
            Top             =   2730
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
            TabIndex        =   15
            Top             =   1320
            Width           =   690
         End
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
            TabIndex        =   14
            Top             =   1770
            Width           =   690
         End
         Begin VB.ComboBox CboTipoCobranca 
            Height          =   315
            ItemData        =   "TelaExpBancaria.frx":014E
            Left            =   1965
            List            =   "TelaExpBancaria.frx":015E
            TabIndex        =   13
            Top             =   2880
            Width           =   3735
         End
         Begin VB.Frame Frame1 
            Height          =   495
            Left            =   5880
            TabIndex        =   10
            Top             =   2760
            Width           =   2895
            Begin VB.TextBox TxtNDiasProtesto 
               DataMember      =   "l"
               Height          =   315
               Left            =   2160
               MaxLength       =   2
               TabIndex        =   11
               Top             =   120
               Width           =   585
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "Nº dias para Protesto:"
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
               Left            =   120
               TabIndex        =   12
               Top             =   120
               Width           =   1965
            End
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
            Left            =   -70710
            TabIndex        =   9
            Top             =   2415
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
            Left            =   -70710
            TabIndex        =   8
            Top             =   2880
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
            Left            =   -70710
            TabIndex        =   7
            Top             =   1440
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
            Left            =   -70710
            TabIndex        =   6
            Top             =   1920
            Width           =   690
         End
         Begin VB.ComboBox CmbInstrucao 
            Height          =   315
            ItemData        =   "TelaExpBancaria.frx":01FC
            Left            =   1920
            List            =   "TelaExpBancaria.frx":0218
            TabIndex        =   5
            Top             =   3360
            Width           =   6615
         End
         Begin TrueDBGrid70.TDBGrid TDBGridEmpr2 
            Height          =   2625
            Left            =   -69960
            TabIndex        =   31
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
         Begin MSComCtl2.DTPicker DtpVencimento 
            Height          =   315
            Left            =   1965
            TabIndex        =   32
            Top             =   435
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MM/yy"
            Format          =   84410371
            CurrentDate     =   37636
         End
         Begin MSComCtl2.DTPicker DtpExportacao 
            Height          =   315
            Left            =   4680
            TabIndex        =   33
            Top             =   420
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   84410371
            CurrentDate     =   37180
         End
         Begin Threed.SSCommand CmdLimparTipoPlano 
            Height          =   315
            Left            =   5160
            TabIndex        =   34
            Top             =   840
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "TelaExpBancaria.frx":0374
            Picture         =   "TelaExpBancaria.frx":0390
         End
         Begin MSDataListLib.DataCombo CboTiposPlanos 
            Bindings        =   "TelaExpBancaria.frx":04A2
            Height          =   315
            Left            =   1965
            TabIndex        =   35
            Top             =   840
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
            Top             =   1245
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "TelaExpBancaria.frx":04BD
            Picture         =   "TelaExpBancaria.frx":04D9
         End
         Begin MSDataListLib.DataCombo CboCCorrente 
            Bindings        =   "TelaExpBancaria.frx":05EB
            Height          =   330
            Left            =   1965
            TabIndex        =   37
            Top             =   1650
            Width           =   6810
            _ExtentX        =   12012
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "DESCCOMBO"
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
            Left            =   5820
            Top             =   1260
            Visible         =   0   'False
            Width           =   2475
            _ExtentX        =   4366
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
            Left            =   7620
            Top             =   840
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
            Top             =   900
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
            TabIndex        =   38
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
            TabIndex        =   39
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
         Begin TrueDBGrid70.TDBGrid TDBGridMoeda2 
            Height          =   1965
            Left            =   -69960
            TabIndex        =   40
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
         Begin TrueDBGrid70.TDBGrid TDBGridObs2 
            Height          =   2565
            Left            =   -69840
            TabIndex        =   41
            Top             =   990
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
         Begin TrueDBGrid70.TDBGrid TDBGridObs1 
            Height          =   2565
            Left            =   -74640
            TabIndex        =   42
            Top             =   990
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
         Begin VB.Label LblInstrucao1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mensagem 1:"
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
            Left            =   600
            TabIndex        =   57
            Top             =   2115
            Width           =   1245
         End
         Begin VB.Label LblInstrucao2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mensagem 2:"
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
            Left            =   600
            TabIndex        =   56
            Top             =   2520
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
            Height          =   210
            Left            =   540
            TabIndex        =   55
            Top             =   1710
            Width           =   1365
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
            TabIndex        =   54
            Top             =   1290
            Width           =   1680
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
            TabIndex        =   53
            Top             =   885
            Width           =   1260
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
            TabIndex        =   52
            Top             =   480
            Width           =   1530
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
            TabIndex        =   51
            Top             =   480
            Width           =   1815
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
            TabIndex        =   50
            Top             =   1140
            Width           =   3720
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
            TabIndex        =   49
            Top             =   1140
            Width           =   3705
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
            TabIndex        =   48
            Top             =   390
            Width           =   3180
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
            TabIndex        =   47
            Top             =   390
            Width           =   2205
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cobrança:"
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
            Left            =   120
            TabIndex        =   46
            Top             =   2880
            Width           =   1740
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
            Left            =   -69825
            TabIndex        =   45
            Top             =   600
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
            Left            =   -74625
            TabIndex        =   44
            Top             =   600
            Width           =   3705
         End
         Begin VB.Label LblMensagem1 
            Caption         =   "Instrução:"
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
            Left            =   960
            TabIndex        =   43
            Top             =   3360
            Width           =   855
         End
      End
   End
   Begin MSComDlg.CommonDialog CdbImpressora 
      Left            =   120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TrueDBGrid70.TDBGrid TDBGrid1 
      Height          =   4605
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8123
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
      Columns(6).EditMask=   "  /  /  "
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
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Número do Contrato do Banco"
      Columns(10).DataField=   "titu_nr_contratoBanco"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Desconto Diário R$"
      Columns(11).DataField=   "DescDiario"
      Columns(11).DefaultValue=   "00,00"
      Columns(11).DefaultValue.vt=   8
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Limite p/ Cons. Desconto."
      Columns(12).DataField=   "DataConsDesc"
      Columns(12).EditMask=   "  /  /  "
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Valor Abatimento"
      Columns(13).DataField=   "ValAbatimento R$"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Multa/Juros"
      Columns(14).DataField=   "MultaJuros"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Data Base"
      Columns(15).DataField=   "titu_dt_Base"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   4
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
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
      Splits(0)._ColumnProps(83)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(84)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(86)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(87)=   "Column(10)._ColStyle=20"
      Splits(0)._ColumnProps(88)=   "Column(10).WrapText=1"
      Splits(0)._ColumnProps(89)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(90)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(91)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(92)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(93)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(94)=   "Column(11)._ColStyle=20"
      Splits(0)._ColumnProps(95)=   "Column(11).WrapText=1"
      Splits(0)._ColumnProps(96)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(97)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(98)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(100)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(101)=   "Column(12)._ColStyle=20"
      Splits(0)._ColumnProps(102)=   "Column(12).WrapText=1"
      Splits(0)._ColumnProps(103)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(104)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(105)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(106)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(107)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(108)=   "Column(13)._ColStyle=20"
      Splits(0)._ColumnProps(109)=   "Column(13).WrapText=1"
      Splits(0)._ColumnProps(110)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(111)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(112)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(113)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(114)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(115)=   "Column(14)._ColStyle=20"
      Splits(0)._ColumnProps(116)=   "Column(14).WrapText=1"
      Splits(0)._ColumnProps(117)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(118)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(119)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(120)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(121)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(122)=   "Column(15)._ColStyle=20"
      Splits(0)._ColumnProps(123)=   "Column(15).WrapText=1"
      Splits(0)._ColumnProps(124)=   "Column(15).Order=16"
      Splits(1)._UserFlags=   0
      Splits(1).ExtendRightColumn=   -1  'True
      Splits(1).MarqueeStyle=   4
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1).DividerColor=   12632256
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=16"
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
      Splits(1)._ColumnProps(74)=   "Column(10).Width=2725"
      Splits(1)._ColumnProps(75)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(76)=   "Column(10)._WidthInPix=2646"
      Splits(1)._ColumnProps(77)=   "Column(10)._EditAlways=0"
      Splits(1)._ColumnProps(78)=   "Column(10)._ColStyle=20"
      Splits(1)._ColumnProps(79)=   "Column(10).WrapText=1"
      Splits(1)._ColumnProps(80)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(81)=   "Column(11).Width=2090"
      Splits(1)._ColumnProps(82)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(83)=   "Column(11)._WidthInPix=2011"
      Splits(1)._ColumnProps(84)=   "Column(11)._EditAlways=0"
      Splits(1)._ColumnProps(85)=   "Column(11)._ColStyle=20"
      Splits(1)._ColumnProps(86)=   "Column(11).WrapText=1"
      Splits(1)._ColumnProps(87)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(88)=   "Column(12).Width=2725"
      Splits(1)._ColumnProps(89)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(90)=   "Column(12)._WidthInPix=2646"
      Splits(1)._ColumnProps(91)=   "Column(12)._EditAlways=0"
      Splits(1)._ColumnProps(92)=   "Column(12)._ColStyle=20"
      Splits(1)._ColumnProps(93)=   "Column(12).WrapText=1"
      Splits(1)._ColumnProps(94)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(95)=   "Column(13).Width=2725"
      Splits(1)._ColumnProps(96)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(97)=   "Column(13)._WidthInPix=2646"
      Splits(1)._ColumnProps(98)=   "Column(13)._EditAlways=0"
      Splits(1)._ColumnProps(99)=   "Column(13)._ColStyle=20"
      Splits(1)._ColumnProps(100)=   "Column(13).WrapText=1"
      Splits(1)._ColumnProps(101)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(102)=   "Column(14).Width=2725"
      Splits(1)._ColumnProps(103)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(104)=   "Column(14)._WidthInPix=2646"
      Splits(1)._ColumnProps(105)=   "Column(14)._EditAlways=0"
      Splits(1)._ColumnProps(106)=   "Column(14)._ColStyle=20"
      Splits(1)._ColumnProps(107)=   "Column(14).WrapText=1"
      Splits(1)._ColumnProps(108)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(109)=   "Column(15).Width=2725"
      Splits(1)._ColumnProps(110)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(111)=   "Column(15)._WidthInPix=2646"
      Splits(1)._ColumnProps(112)=   "Column(15)._EditAlways=0"
      Splits(1)._ColumnProps(113)=   "Column(15)._ColStyle=20"
      Splits(1)._ColumnProps(114)=   "Column(15).WrapText=1"
      Splits(1)._ColumnProps(115)=   "Column(15).Order=16"
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
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=32,.parent=123"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=29,.parent=124"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=30,.parent=125"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=31,.parent=127"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=94,.parent=123"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=91,.parent=124"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=92,.parent=125"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=93,.parent=127"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=110,.parent=123"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=107,.parent=124"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=108,.parent=125"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=109,.parent=127"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=102,.parent=123"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=99,.parent=124"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=100,.parent=125"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=101,.parent=127"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=118,.parent=123"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=115,.parent=124"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=116,.parent=125"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=117,.parent=127"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=170,.parent=123"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=167,.parent=124"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=168,.parent=125"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=169,.parent=127"
      _StyleDefs(100) =   "Splits(1).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(101) =   "Splits(1).CaptionStyle:id=22,.parent=4"
      _StyleDefs(102) =   "Splits(1).HeadingStyle:id=14,.parent=2"
      _StyleDefs(103) =   "Splits(1).FooterStyle:id=15,.parent=3"
      _StyleDefs(104) =   "Splits(1).InactiveStyle:id=16,.parent=5"
      _StyleDefs(105) =   "Splits(1).SelectedStyle:id=18,.parent=6"
      _StyleDefs(106) =   "Splits(1).EditorStyle:id=17,.parent=7"
      _StyleDefs(107) =   "Splits(1).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H800000&"
      _StyleDefs(108) =   "Splits(1).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(109) =   "Splits(1).OddRowStyle:id=21,.parent=10"
      _StyleDefs(110) =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(111) =   "Splits(1).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(112) =   "Splits(1).Columns(0).Style:id=28,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(113) =   "Splits(1).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(114) =   "Splits(1).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(115) =   "Splits(1).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(116) =   "Splits(1).Columns(1).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(117) =   "Splits(1).Columns(1).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(118) =   "Splits(1).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(119) =   "Splits(1).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(120) =   "Splits(1).Columns(2).Style:id=50,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(121) =   "Splits(1).Columns(2).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(122) =   "Splits(1).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(123) =   "Splits(1).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(124) =   "Splits(1).Columns(3).Style:id=58,.parent=13,.alignment=1,.locked=-1"
      _StyleDefs(125) =   "Splits(1).Columns(3).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(126) =   "Splits(1).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(127) =   "Splits(1).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(128) =   "Splits(1).Columns(4).Style:id=62,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(129) =   "Splits(1).Columns(4).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(130) =   "Splits(1).Columns(4).FooterStyle:id=60,.parent=15"
      _StyleDefs(131) =   "Splits(1).Columns(4).EditorStyle:id=61,.parent=17"
      _StyleDefs(132) =   "Splits(1).Columns(5).Style:id=66,.parent=13,.alignment=1,.locked=-1"
      _StyleDefs(133) =   "Splits(1).Columns(5).HeadingStyle:id=63,.parent=14,.alignment=1"
      _StyleDefs(134) =   "Splits(1).Columns(5).FooterStyle:id=64,.parent=15"
      _StyleDefs(135) =   "Splits(1).Columns(5).EditorStyle:id=65,.parent=17"
      _StyleDefs(136) =   "Splits(1).Columns(6).Style:id=70,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(137) =   "Splits(1).Columns(6).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(138) =   "Splits(1).Columns(6).FooterStyle:id=68,.parent=15"
      _StyleDefs(139) =   "Splits(1).Columns(6).EditorStyle:id=69,.parent=17"
      _StyleDefs(140) =   "Splits(1).Columns(7).Style:id=74,.parent=13,.alignment=1,.locked=0"
      _StyleDefs(141) =   "Splits(1).Columns(7).HeadingStyle:id=71,.parent=14,.alignment=2"
      _StyleDefs(142) =   "Splits(1).Columns(7).FooterStyle:id=72,.parent=15"
      _StyleDefs(143) =   "Splits(1).Columns(7).EditorStyle:id=73,.parent=17"
      _StyleDefs(144) =   "Splits(1).Columns(8).Style:id=86,.parent=13,.alignment=1"
      _StyleDefs(145) =   "Splits(1).Columns(8).HeadingStyle:id=83,.parent=14"
      _StyleDefs(146) =   "Splits(1).Columns(8).FooterStyle:id=84,.parent=15"
      _StyleDefs(147) =   "Splits(1).Columns(8).EditorStyle:id=85,.parent=17"
      _StyleDefs(148) =   "Splits(1).Columns(9).Style:id=90,.parent=13"
      _StyleDefs(149) =   "Splits(1).Columns(9).HeadingStyle:id=87,.parent=14"
      _StyleDefs(150) =   "Splits(1).Columns(9).FooterStyle:id=88,.parent=15"
      _StyleDefs(151) =   "Splits(1).Columns(9).EditorStyle:id=89,.parent=17"
      _StyleDefs(152) =   "Splits(1).Columns(10).Style:id=54,.parent=13"
      _StyleDefs(153) =   "Splits(1).Columns(10).HeadingStyle:id=51,.parent=14"
      _StyleDefs(154) =   "Splits(1).Columns(10).FooterStyle:id=52,.parent=15"
      _StyleDefs(155) =   "Splits(1).Columns(10).EditorStyle:id=53,.parent=17"
      _StyleDefs(156) =   "Splits(1).Columns(11).Style:id=98,.parent=13"
      _StyleDefs(157) =   "Splits(1).Columns(11).HeadingStyle:id=95,.parent=14"
      _StyleDefs(158) =   "Splits(1).Columns(11).FooterStyle:id=96,.parent=15"
      _StyleDefs(159) =   "Splits(1).Columns(11).EditorStyle:id=97,.parent=17"
      _StyleDefs(160) =   "Splits(1).Columns(12).Style:id=114,.parent=13"
      _StyleDefs(161) =   "Splits(1).Columns(12).HeadingStyle:id=111,.parent=14"
      _StyleDefs(162) =   "Splits(1).Columns(12).FooterStyle:id=112,.parent=15"
      _StyleDefs(163) =   "Splits(1).Columns(12).EditorStyle:id=113,.parent=17"
      _StyleDefs(164) =   "Splits(1).Columns(13).Style:id=106,.parent=13"
      _StyleDefs(165) =   "Splits(1).Columns(13).HeadingStyle:id=103,.parent=14"
      _StyleDefs(166) =   "Splits(1).Columns(13).FooterStyle:id=104,.parent=15"
      _StyleDefs(167) =   "Splits(1).Columns(13).EditorStyle:id=105,.parent=17"
      _StyleDefs(168) =   "Splits(1).Columns(14).Style:id=122,.parent=13"
      _StyleDefs(169) =   "Splits(1).Columns(14).HeadingStyle:id=119,.parent=14"
      _StyleDefs(170) =   "Splits(1).Columns(14).FooterStyle:id=120,.parent=15"
      _StyleDefs(171) =   "Splits(1).Columns(14).EditorStyle:id=121,.parent=17"
      _StyleDefs(172) =   "Splits(1).Columns(15).Style:id=174,.parent=13"
      _StyleDefs(173) =   "Splits(1).Columns(15).HeadingStyle:id=171,.parent=14"
      _StyleDefs(174) =   "Splits(1).Columns(15).FooterStyle:id=172,.parent=15"
      _StyleDefs(175) =   "Splits(1).Columns(15).EditorStyle:id=173,.parent=17"
      _StyleDefs(176) =   "Named:id=33:Normal"
      _StyleDefs(177) =   ":id=33,.parent=0"
      _StyleDefs(178) =   "Named:id=34:Heading"
      _StyleDefs(179) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(180) =   ":id=34,.wraptext=-1"
      _StyleDefs(181) =   "Named:id=35:Footing"
      _StyleDefs(182) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(183) =   "Named:id=36:Selected"
      _StyleDefs(184) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(185) =   "Named:id=37:Caption"
      _StyleDefs(186) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(187) =   "Named:id=38:HighlightRow"
      _StyleDefs(188) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(189) =   "Named:id=39:EvenRow"
      _StyleDefs(190) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(191) =   "Named:id=40:OddRow"
      _StyleDefs(192) =   ":id=40,.parent=33"
      _StyleDefs(193) =   "Named:id=41:RecordSelector"
      _StyleDefs(194) =   ":id=41,.parent=34"
      _StyleDefs(195) =   "Named:id=42:FilterBar"
      _StyleDefs(196) =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   975
      Left            =   0
      TabIndex        =   58
      Top             =   4680
      Width           =   9975
      _Version        =   65536
      _ExtentX        =   17595
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
      Begin VB.CommandButton btnSair 
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
         Left            =   6540
         TabIndex        =   66
         Top             =   495
         Width           =   1680
      End
      Begin VB.CommandButton btnExportar 
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
         TabIndex        =   65
         Top             =   480
         Width           =   1680
      End
      Begin VB.CommandButton btnSelecoes 
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
         TabIndex        =   64
         Top             =   495
         Width           =   1680
      End
      Begin VB.CommandButton btnImprimir 
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
         Left            =   4620
         TabIndex        =   63
         Top             =   495
         Width           =   1680
      End
      Begin VB.CommandButton btnDesmarcarTodos 
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
         TabIndex        =   62
         Top             =   15
         Width           =   1680
      End
      Begin VB.CommandButton btnPesquisar 
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
         TabIndex        =   61
         Top             =   0
         Width           =   1680
      End
      Begin VB.CommandButton btnMarcarTodos 
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
         TabIndex        =   60
         Top             =   15
         Width           =   1680
      End
      Begin VB.CommandButton btnDesconto 
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
         TabIndex        =   59
         Top             =   0
         Width           =   1680
      End
   End
End
Attribute VB_Name = "TelaExpBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Alterações que tem que ser repassadas para os outros bancos

'Evento de exportação (moeda nula) xlt_moeda está como integer
'Leitura do diretório padrão no arquivo ini
'Colocar o * 254 na variável XGT_LOCALARQ
'Alterar o initdir para XGT_LOCALARQ
'Habilitar o botão Imprimir após exportar

Option Explicit

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
    'Formulario.Top = 1020
    Formulario.Top = 0
    Formulario.Height = Resolucaoy - 500
    'Formulario.Height = TelaExpBancaria.SysInfo1.WorkAreaHeight - 1380
    Formulario.Width = ResolucaoX - 60
    'Formulario.Width = TelaExpBancaria.SysInfo1.WorkAreaWidth - 60
    'Formulario.Left = 10
    Formulario.Left = 5
    'Formulario.PanPeriodo.Top = 100 '45
    'Formulario.PanPeriodo.Left = 75
    Formulario.TDBGrid1.Top = 45
    Formulario.TDBGrid1.Left = 75
    Formulario.TDBGrid1.Width = Formulario.Width - 150
    'Formulario.TDBGrid1.Height = Formulario.Height - Formulario.TDBGrid1.Top - Formulario.PanBotoes.Height - 500
    Formulario.TDBGrid1.Height = Formulario.Height - Formulario.TDBGrid1.Top - Formulario.PanBotoes.Height - 50
    Formulario.PanBotoes.Top = Formulario.TDBGrid1.Top + Formulario.TDBGrid1.Height + 50
    Formulario.PanBotoes.Left = 75 + (Formulario.TDBGrid1.Width - Formulario.PanBotoes.Width) / 2
    'Formulario.PanPeriodo.Width = Formulario.Width - 150
End Sub

Sub subTelaValoresGlobais(GuardaOUPegaValores As String)
    Dim Count As Integer, Count2 As Integer
    If GuardaOUPegaValores = "G" Then 'Guardando variáveis de definição da tela
        FRM_ContFil = ContFil
        FRM_Ordem = Ordem
        FRM_Campos_Obrig = Campos_Obrig
        'FRM_TipoVar = TipoVar
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
        'TipoVar = FRM_TipoVar
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
    btnDesmarcarTodos.Enabled = True
    btnMarcarTodos.Enabled = True
    btnExportar.Enabled = True
    btnDesconto.Enabled = True
    btnImprimir.Enabled = True
    
    'Verifica se tem Filtro Ativo
    If ContFil = 0 Then
       btnSelecoes.Enabled = False
    Else
       btnSelecoes.Enabled = True
    End If
End Sub

Private Sub subDesabilitaBotoes()
    btnDesmarcarTodos.Enabled = False
    btnMarcarTodos.Enabled = False
    btnExportar.Enabled = False
    btnDesconto.Enabled = False
    btnImprimir.Enabled = False
    btnSelecoes.Enabled = False
End Sub

Sub SubCalculaValorAtualizado()

  Dim XLI_CONT As Integer
  Dim XLF_VALORCORRIGIDO As Double
  Dim XLF_JUROS As Double
  Dim XLF_MULTA As Double
    
  XFO_EXPORTACAO.MoveFirst
  For XLI_CONT = 0 To TDBGrid1.ApproxCount - 1
    If OptIndexador1.Value = True Then
        XLF_VALORCORRIGIDO = Format(ModQualiCapiPagamento.funCalculaCorrecaoMonetaria(XFO_EXPORTACAO, 1, XFO_EXPORTACAO!titu_dt_Vencimento), "standard")
        VFV_VETOREXP(XLI_CONT, 7) = Format(XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XFO_EXPORTACAO, XFO_EXPORTACAO!titu_dt_Vencimento, XLF_VALORCORRIGIDO), "standard"), "standard")
        VFV_VETOREXP(XLI_CONT, 3) = VFV_VETOREXP(XLI_CONT, 7)
        'CÁLCULO DE JUROS
        '31/03/10 - Patrícia
        If ChkJuros.Value = 1 Then
            VFV_VETOREXP(XLI_CONT, 14) = Format(((XFO_EXPORTACAO!titu_vl_PercJurosMora) / 100) * VFV_VETOREXP(XLI_CONT, 7), "standard") '31/03/10 - Patrícia
        Else
            VFV_VETOREXP(XLI_CONT, 14) = 0
        End If
    ElseIf OptIndexador2.Value = True Then
        XLF_VALORCORRIGIDO = Format(ModQualiCapiPagamento.funCalculaCorrecaoMonetaria(XFO_EXPORTACAO, 2, XFO_EXPORTACAO!titu_dt_Vencimento), "standard")
        VFV_VETOREXP(XLI_CONT, 7) = Format(XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XFO_EXPORTACAO, XFO_EXPORTACAO!titu_dt_Vencimento, XLF_VALORCORRIGIDO), "standard"), "standard")
        VFV_VETOREXP(XLI_CONT, 3) = VFV_VETOREXP(XLI_CONT, 7)
        'CÁLCULO DE JUROS
        '31/03/10 - Patrícia
        If ChkJuros.Value = 1 Then
            VFV_VETOREXP(XLI_CONT, 14) = Format(((XFO_EXPORTACAO!titu_vl_PercJurosMora) / 100) * VFV_VETOREXP(XLI_CONT, 7), "standard") '31/03/10 - Patrícia
        Else
            VFV_VETOREXP(XLI_CONT, 14) = 0
        End If
    Else
        XLF_VALORCORRIGIDO = XFO_EXPORTACAO!titu_vl_Parcela
        VFV_VETOREXP(XLI_CONT, 7) = Format(XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XFO_EXPORTACAO, XFO_EXPORTACAO!titu_dt_Vencimento, XLF_VALORCORRIGIDO), "standard"), "standard")
        VFV_VETOREXP(XLI_CONT, 3) = VFV_VETOREXP(XLI_CONT, 7)
        'CÁLCULO DE JUROS
        '31/03/10 - Patrícia
        If ChkJuros.Value = 1 Then
            VFV_VETOREXP(XLI_CONT, 14) = Format(((XFO_EXPORTACAO!titu_vl_PercJurosMora) / 100) * VFV_VETOREXP(XLI_CONT, 7), "standard") '31/03/10 - Patrícia
        Else
            VFV_VETOREXP(XLI_CONT, 14) = 0
        End If
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
    
  XLT_SQL = FunCriaConsultaBase(TDBGrid1.Columns("Vencimento"), NomeSgbd, "ConsCAPExpBancoCaixa", 1)
  XLT_SQL = XLT_SQL & "empr_cd_empresa = " & PCodEmpresa & "" & _
    " AND empd_cd_Empreendimento= '" & Left(XLO_RECORDSET("Titulo"), 4) & "'" & _
    " AND imov_cd_Imovel='" & Mid(XLO_RECORDSET("Titulo"), 6, 4) & "'" & _
    " AND cont_cd_Contrato='" & Mid(XLO_RECORDSET("Titulo"), 11, 2) & "'" & _
    " AND titu_dt_Pagamento is null "
   
  SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
 
  While Not XLO_TITULO.EOF
    If OptIndexador1.Value = True Then
      XLF_VALORCORRIGIDO = Format(ModQualiCapiPagamento.funCalculaCorrecaoMonetaria(XLO_TITULO, 1, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy")), "standard")
      XLF_VALORCORRIGIDO = XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XFO_EXPORTACAO, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"), XLF_VALORCORRIGIDO), "standard")
    ElseIf OptIndexador2.Value = True Then
      XLF_VALORCORRIGIDO = ModQualiCapiPagamento.funCalculaCorrecaoMonetaria(XLO_TITULO, 2, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"))
      XLF_VALORCORRIGIDO = XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XFO_EXPORTACAO, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"), XLF_VALORCORRIGIDO), "standard")
    Else
      XLF_VALORCORRIGIDO = XLO_TITULO!titu_vl_Parcela
      XLF_VALORCORRIGIDO = XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XFO_EXPORTACAO, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"), XLF_VALORCORRIGIDO), "standard")
    End If
    
    XLF_DESAGIO = ModQualiCapiPagamento.funCalculaDesagio(XLO_TITULO, XLD_DATA, XLF_VALORCORRIGIDO)
    
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
    'While Left(VFV_VETOREXP(XLI_CONT, 0), 12) = XLT_CONTRATO
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
            XLF_VALOR = ModQualiCapiPagamento.funCalculaCorrecaoMonetaria(XFO_EXPORTACAO, 1, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"))
            XLF_VALOR = XLF_VALOR + Format(FunCalculaJurosCEF(XFO_EXPORTACAO, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"), XLF_VALOR), "standard")
          Else
            XLF_VALOR = ModQualiCapiPagamento.funCalculaCorrecaoMonetaria(XFO_EXPORTACAO, 2, Format(XFO_EXPORTACAO!titu_dt_Vencimeto, "\01/mm/yy"))
            XLF_VALOR = XLF_VALOR + Format(FunCalculaJurosCEF(XFO_EXPORTACAO, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"), XLF_VALOR), "standard")
          End If
          XLF_VALOR = XLF_VALOR - ModQualiCapiPagamento.funCalculaDesagio(XFO_EXPORTACAO, XLD_DATABASE, XLF_VALOR)
          XLF_SALDODEVEDOR = XLF_SALDODEVEDOR - XLF_VALOR
        End If
      End If
      
      XLF_VALORPARCELA = XFO_EXPORTACAO!titu_vl_Parcela
      XFO_EXPORTACAO.MoveNext
      XLI_CONT = XLI_CONT + 1
      
      'Sai do loop quando o contrato é diferente ou chega no final do grid
      If XLI_CONT = TDBGrid1.ApproxCount Then
        XLB_SAIR = True
        'GoTo apelacao
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
                  'Vetor1(TDBGrid1.Bookmark, 13) = Format(XFF_SALDODEVEDOR * XLF_TAXASEGURO, "standard")
                  FunProcuraNoGrid = XLO_SEGURO!Titulo & Format(XLO_SEGURO!titu_dt_Vencimento, "dd/mm/yy")
                  XLB_ACHEI = True
              'Else
                  'Vetor1(TDBGrid1.Bookmark, 13) = ""
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
                    'Vetor1(TDBGrid1.Bookmark, 13) = Format(XFF_SALDODEVEDOR * XLF_TAXASEGURO, "standard")
                    FunProcuraNoGrid = XLO_SEGURO!Titulo & Format(XLO_SEGURO!titu_dt_Vencimento, "dd/mm/yy")
                    XLB_ACHEI = True
                'Else
                    'Vetor1(TDBGrid1.Bookmark, 13) = ""
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
                   'Vetor1(TDBGrid1.Bookmark, 13) = Format(XFF_SALDODEVEDOR * XLF_TAXASEGURO, "standard")
                   FunProcuraNoGrid = XLO_SEGURO!Titulo & Format(XLO_SEGURO!titu_dt_Vencimento, "dd/mm/yy")
                   XLB_ACHEI = True
               'Else
                   'Vetor1(TDBGrid1.Bookmark, 13) = ""
               End If
               TDBGrid1.MoveNext
           Wend
           XLO_SEGURO.MoveNext
       Wend
       'TDBGrid1.Refresh
  XLO_SEGURO.Close
  Set XLO_SEGURO = Nothing
  End If
  TDBGrid1.MoveFirst

End Function

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
              XLF_VALOR = ModQualiCapiPagamento.funCalculaCorrecaoMonetaria(XFO_EXPORTACAO, 1, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"))
              XLF_VALOR = XLF_VALOR + Format(FunCalculaJurosCEF(XFO_EXPORTACAO, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"), XLF_VALOR), "standard")
            Else
              XLF_VALOR = ModQualiCapiPagamento.funCalculaCorrecaoMonetaria(XFO_EXPORTACAO, 2, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"))
              XLF_VALOR = XLF_VALOR + Format(FunCalculaJurosCEF(XFO_EXPORTACAO, Format(XFO_EXPORTACAO!titu_dt_Vencimento, "\01/mm/yy"), XLF_VALOR), "standard")
            End If
            XLF_VALOR = XLF_VALOR - ModQualiCapiPagamento.funCalculaDesagio(XFO_EXPORTACAO, XLD_DATABASE, XLF_VALOR)
            XLF_SALDODEVEDOR = XLF_SALDODEVEDOR - XLF_VALOR
          End If
        End If
      End If
      
      'Grava o Valor R$ igual ao Val. Tit R$ para retirar o seguro embutido
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
            'XLI_CONT = -1
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

Private Sub btnPesquisar_Click()
     
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

Private Sub btnMarcarTodos_Click()
  'Pré-selecionar os registros
  Dim XLI_CONT As Integer
  Dim XLI_BOOKMARK As Integer
    
  XLI_BOOKMARK = TDBGrid1.Bookmark
  For XLI_CONT = 0 To TDBGrid1.ApproxCount - 1
    VFV_VETOREXP(XLI_CONT, 1) = "-1"
  Next
  TDBGrid1.Refresh
  TDBGrid1.Bookmark = XLI_BOOKMARK
    
End Sub

Private Sub btnDesmarcarTodos_Click()
  'Desmarcar os registros
  Dim XLI_CONT As Integer
  Dim XLI_BOOKMARK As Integer
    
  XLI_BOOKMARK = TDBGrid1.Bookmark
  For XLI_CONT = 0 To TDBGrid1.ApproxCount - 1
    VFV_VETOREXP(XLI_CONT, 1) = "0"
  Next
  TDBGrid1.Refresh
  TDBGrid1.Bookmark = XLI_BOOKMARK
    
End Sub

Private Sub btnExportar_Click()
 
  Dim XFT_SQL As String           'Prepara Select
  Dim XLT_TEXTO As String         'Prepara Texto para gerar o arquivo
  Dim XLT_CGCEMP As String        'CGC da empresa sem máscara
  Dim XLT_CEPEMP As String        'Cep sem máscara
  Dim XLF_VALOR As String         'Valor do registro
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
  Dim XLT_REGLOTE As Integer
  Dim XLI_ARQUIVO As Integer
  Dim XLO_ARQUIVO As Object
  Dim XLT_SQL As String
  Dim XLI_TIPOCOBRANCA As Integer
  Dim XLT_VALMORA As Double
  Dim XLO_DADOSNF As New ADODB.Recordset
  Dim XLT_MENSAGEMVARIAVELTITULO As String
  Dim XLT_RESIDUOPARCELA As String
  Dim XLT_TIPOPLANOS As String
  Dim XLT_TITULO As String
  Dim XLO_PERCDESCONTO As Double
  Dim XLO_PERCJUROS As String
  Dim XLT_VALORMULTA As String
  Dim XLT_VALORTEMP As String
  Dim XLO_ULTREMESSA As New ADODB.Recordset
  
  'Incluído em 03/11/2008
  On Error GoTo TrataErro:
   
  XLB_SELECIONADO = False
  TDBGrid1.MoveFirst
  XFO_EXPORTACAO.MoveFirst
  
  While Not TDBGrid1.EOF
    If TDBGrid1.Columns("Exp.").Value = "-1" Then
      XLB_SELECIONADO = True
      
      
      'Verifica se está preenchido o endereço correspondencial,
        'se todos os campos estão vazios
           'verifica se é pessoa jurídica e verifica o endereço comercial
           'verifica se é pessoafísica  e verifica o endereço residencial
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
        
      If XFO_EXPORTACAO!titu_nr_contratobanco = "" Or IsNull(XFO_EXPORTACAO!titu_nr_contratobanco) Then
        MsgBox "O título do Cliente " & XFO_EXPORTACAO!focl_tx_RazaoSocial & _
              vbCrLf & "está sem o número contrato do banco. Favor inserir o número do contrato no título!"
          
          Exit Sub
      End If
        
        
      If (XFO_EXPORTACAO!clie_tx_EndCorresp = "" Or IsNull(XFO_EXPORTACAO!clie_tx_EndCorresp)) And (XFO_EXPORTACAO!clie_tx_BairroCorresp = "" Or IsNull(XFO_EXPORTACAO!clie_tx_BairroCorresp)) _
        And (XFO_EXPORTACAO!clie_nr_CepCorresp = "  .   -   " Or IsNull(XFO_EXPORTACAO!clie_nr_CepCorresp)) And (XFO_EXPORTACAO!clie_tx_MunCorresp = "" Or IsNull(XFO_EXPORTACAO!clie_tx_MunCorresp)) _
        And (XFO_EXPORTACAO!clie_tx_EstCorresp = "" Or IsNull(XFO_EXPORTACAO!clie_tx_EstCorresp)) Then
        
        If XFO_EXPORTACAO!focl_tx_Tipo = "J" Then
          If (XFO_EXPORTACAO!focl_tx_Endereco = "" Or IsNull(XFO_EXPORTACAO!focl_tx_Endereco)) Or (XFO_EXPORTACAO!focl_tx_Bairro = "" Or IsNull(XFO_EXPORTACAO!focl_tx_Bairro)) _
            Or (XFO_EXPORTACAO!focl_tx_Cep = "  .   -  " Or IsNull(XFO_EXPORTACAO!focl_tx_Cep)) Or (XFO_EXPORTACAO!focl_tx_Cidade = "" Or IsNull(XFO_EXPORTACAO!focl_tx_Cidade)) _
            Or XFO_EXPORTACAO!focl_tx_Estado = "" Then
            
            MsgBox "Os dados do endereço comercial do Cliente " & XFO_EXPORTACAO!focl_tx_RazaoSocial & _
              vbCrLf & "estão incompletos."
            Exit Sub
          End If
          
        Else
          If (XFO_EXPORTACAO!clie_tx_EndResidencial = "" Or IsNull(XFO_EXPORTACAO!clie_tx_EndResidencial)) Or XFO_EXPORTACAO!clie_tx_BairroResidencial = "" _
            Or (XFO_EXPORTACAO!clie_nr_CepResidencial = "  .   -   " Or IsNull(XFO_EXPORTACAO!clie_nr_CepResidencial)) Or (XFO_EXPORTACAO!clie_tx_MunResidencial = "" Or IsNull(XFO_EXPORTACAO!clie_tx_MunResidencial)) _
            Or (XFO_EXPORTACAO!clie_tx_EstResidencial = "" Or IsNull(XFO_EXPORTACAO!clie_tx_EstResidencial)) Then
            
            MsgBox "Os dados do endereço residencial do Cliente " & XFO_EXPORTACAO!focl_tx_RazaoSocial & _
              vbCrLf & "estão incompletos."
            Exit Sub
          End If
        End If
        
      ElseIf (XFO_EXPORTACAO!clie_tx_EndCorresp = "" Or IsNull(XFO_EXPORTACAO!clie_tx_EndCorresp)) Or (XFO_EXPORTACAO!clie_tx_BairroCorresp = "" Or IsNull(XFO_EXPORTACAO!clie_tx_BairroCorresp)) _
         Or (XFO_EXPORTACAO!clie_nr_CepCorresp = "  .   -   " Or IsNull(XFO_EXPORTACAO!clie_nr_CepCorresp)) Or (XFO_EXPORTACAO!clie_tx_MunCorresp = "" Or IsNull(XFO_EXPORTACAO!clie_tx_MunCorresp)) _
         Or (XFO_EXPORTACAO!clie_tx_EstCorresp = "" Or IsNull(XFO_EXPORTACAO!clie_tx_EstCorresp)) Then
        
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
  
  'Verificando o código do convênio
  If DatContaCorrente.Recordset.Fields!coco_nr_conveniocobranca = "" Then
    MsgBox "Favor incluir o código do convênio no cadastro de conta corrente.", vbInformation, "EXPORTAÇÃO"
    Exit Sub
  End If
  
  DlgSalvar.DialogTitle = "Arquivo de Exportação do CAIXA"
  DlgSalvar.Filter = "Arquivos de Remessa (*.txt)|*.txt"
  DlgSalvar.InitDir = "C:\"
  DlgSalvar.FileName = "CCE" & Format(Now, "dd") & Format(Now, "MM") & Format(Now, "yy")
  DlgSalvar.CancelError = False
  DlgSalvar.ShowSave
  
  
  If DlgSalvar.FileName = "" Then
      Exit Sub
  End If
  
  If DlgSalvar.FileName <> "" Then
      XGT_LOCALARQ = DlgSalvar.FileName
      XGT_ARQUIVO = DlgSalvar.FileTitle
  End If

  XLT_SQL = "SELECT Ulti_nr_RemessaCaixa FROM Ultimos"
  SubQOpenRecordset XLO_ULTREMESSA, XLT_SQL, Estatico
  
  Open XGT_LOCALARQ For Output As #1
  
  'Selecionando o último número de arquivo
  XLT_SQL = "SELECT coco_nr_arqcobranca FROM ContasCorrente WHERE coco_cd_codigo = " & DatContaCorrente.Recordset.Fields!coco_cd_codigo & " AND empr_cd_empresa = " & PCodEmpresa & " ORDER BY coco_cd_codigo DESC"
  SubQOpenRecordset XLO_ARQUIVO, XLT_SQL, Dinamico
  If XLO_ARQUIVO!coco_nr_arqcobranca <> "" Then
      XLI_ARQUIVO = CInt(XLO_ARQUIVO!coco_nr_arqcobranca) + 1
  Else
      XLI_ARQUIVO = 1
  End If
  
  XLO_ARQUIVO.Close
   
  Conexao.BeginTrans
  
  Conexao.Execute "UPDATE ContasCorrente SET coco_nr_arqcobranca = " & XLI_ARQUIVO & " WHERE coco_cd_codigo = " & DatContaCorrente.Recordset.Fields!coco_cd_codigo & " AND empr_cd_empresa = " & PCodEmpresa
  
  'CGC DA EMPRESA
  XLT_CGCEMP = Mid(XGT_CGC, 1, 2) + Mid(XGT_CGC, 4, 3) + Mid(XGT_CGC, 8, 3) + Mid(XGT_CGC, 12, 4) + Mid(XGT_CGC, 17, 2)
  
  '**********DESCRIÇÃO DE REGISTRO TIPO '0' (OBRIGATÓRIO) - HEADER DE REMESSA**********
  'CÓDIGO DO REGISTRO
  XLT_TEXTO = "0"
  'CÓDIGO DA REMESSA
  XLT_TEXTO = XLT_TEXTO & "1"
  'LITERAL DA REMESSA
  XLT_TEXTO = XLT_TEXTO & "REMESSA"
  'CÓDIGO DO SERVIÇO
  XLT_TEXTO = XLT_TEXTO & "01"
  'LITERAL DE SERVIÇO
  XLT_TEXTO = XLT_TEXTO & "COBRANCA" & Space(15 - Len("COBRANCA"))
  'CÓDIGO DA AGÊNCIA
  XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(DatContaCorrente.Recordset.Fields!coco_cd_Agencia) = "", FuncoesGenericas.FunZeros(4), (FuncoesGenericas.FunZeros(4 - Len(DatContaCorrente.Recordset.Fields!coco_cd_Agencia))) & Format(DatContaCorrente.Recordset.Fields!coco_cd_Agencia, "0000"))
  'CÓDIGO DO BENEFICIÁRIO
  XLT_TEXTO = XLT_TEXTO & DatContaCorrente.Recordset.Fields!coco_nr_conveniocobranca
  'USO EXCLUSIVO
  XLT_TEXTO = XLT_TEXTO & Space(10)
  'NOME DA EMPRESA
  XLT_TEXTO = XLT_TEXTO & Mid(PEmpresa, 1, 30) & (FuncoesGenericas.FunZeros(30 - Len(PEmpresa)))
  'CÓDIGO DO BANCO
  XLT_TEXTO = XLT_TEXTO & "104"
  'NOME DO BANCO
  XLT_TEXTO = XLT_TEXTO & "CAIXA ECONOMICA"
  'DATA DE GERAÇÃO
  XLT_TEXTO = XLT_TEXTO & Format(DateTime.Now, "ddMMyy")
  'USO EXCLUSIVO
  XLT_TEXTO = XLT_TEXTO & Space(289)
  'Nº SEQUENCIAL - A
  XLT_TEXTO = XLT_TEXTO & FuncoesGenericas.FunZeros(5 - Len(XLO_ULTREMESSA!Ulti_nr_RemessaCaixa)) & XLO_ULTREMESSA!Ulti_nr_RemessaCaixa
  'Nº SEQUENCIAL - B
  XLT_TEXTO = XLT_TEXTO & "000001"
  
  Print #1, funTiraAcento(XLT_TEXTO, True)
  
  XFO_EXPORTACAO.MoveFirst
  TDBGrid1.MoveFirst
  
  XLI_NUMSQUENCIAL = 1
  
  While Not TDBGrid1.EOF
      
      If TDBGrid1.Columns("Exp.").Value = "-1" Then
         
          XLI_NUMSQUENCIAL = XLI_NUMSQUENCIAL + 1
                    
          'Valor já com Seguro
          XLF_VALOR = CDbl(TDBGrid1.Columns(3))
          XLF_VALOR = Format(XLF_VALOR, "0.00")
          
          If IsDate(TDBGrid1.Columns(6)) Then
            XLT_DATADESCONTO = Format(TDBGrid1.Columns(6), "ddmmyy")
          Else
            XLT_DATADESCONTO = "000000"
          End If
          
          If (TDBGrid1.Columns(5)) = "" Then
              TDBGrid1.Columns(5) = 0
          End If
          
          XLF_DESCONTO = CDbl(TDBGrid1.Columns(5))
            
          'Número do contrato no banco
          Dim NUMDOCUMENTO As String
          Dim XLO_BLOCO As ADODB.Recordset
          Dim XLT_BLOCO As String
          
          XLT_SQL = "SELECT empd_nr_bloco FROM empreendimentos WHERE empd_cd_empreendimento = " & XFO_EXPORTACAO!empd_cd_Empreendimento & " AND empr_cd_empresa = " & PCodEmpresa
          SubQOpenRecordset XLO_BLOCO, XLT_SQL, Dinamico
          
            If XLO_BLOCO.EOF Or XLO_BLOCO!empd_nr_bloco = "" Or IsNull(XLO_BLOCO!empd_nr_bloco) Then
                XLT_BLOCO = 0
            Else
                XLT_BLOCO = XLO_BLOCO!empd_nr_bloco
            End If
          XLO_BLOCO.Close
          
          NUMDOCUMENTO = XFO_EXPORTACAO!titu_nr_contratobanco & XLT_BLOCO & XFO_EXPORTACAO!imov_cd_Imovel
          
          'Tipo de Cobrança
          Select Case Me.CboTipoCobranca.ListIndex
             Case 0
                  XLI_TIPOCOBRANCA = 1
             Case 1
                  XLI_TIPOCOBRANCA = 3
             Case 2
                  XLI_TIPOCOBRANCA = 5
             Case 3
                  XLI_TIPOCOBRANCA = 6
         End Select

          If XFO_EXPORTACAO!focl_tx_Tipo = "J" Then
            XLT_TIPO = "02"
            XLT_CGCCPF = Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 1, 2) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 4, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 8, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 12, 4) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 17, 2)
          Else
            XLT_TIPO = "01"
            XLT_CGCCPF = Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 1, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 5, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 9, 3) + Mid(XFO_EXPORTACAO!focl_tx_CGCCPF, 13, 2)
          End If
          
          XLT_NOME = Mid(XFO_EXPORTACAO!focl_tx_RazaoSocial, 1, 40)
          
          If XFO_EXPORTACAO!clie_tx_EndCorresp = "" And XFO_EXPORTACAO!clie_tx_BairroCorresp = "" _
            And XFO_EXPORTACAO!clie_nr_CepCorresp = "  .   -  " And XFO_EXPORTACAO!clie_tx_MunCorresp = "" _
            And XFO_EXPORTACAO!clie_tx_EstCorresp = "" Then
            
            If XLT_TIPO = "02" Then
              XLT_ENDERECO = Mid(XFO_EXPORTACAO!focl_tx_Endereco, 1, 40)
              XLT_BAIRRO = Mid(XFO_EXPORTACAO!focl_tx_Bairro, 1, 12)
              XLT_CEP = (Mid(XFO_EXPORTACAO!focl_tx_Cep, 1, 2) + Mid(XFO_EXPORTACAO!focl_tx_Cep, 4, 3) + Mid(XFO_EXPORTACAO!focl_tx_Cep, 8, 3))
              XLT_CIDADE = Mid(XFO_EXPORTACAO!focl_tx_Cidade, 1, 15)
              XLT_ESTADO = Mid(XFO_EXPORTACAO!focl_tx_Estado, 1, 2)
            Else
              XLT_ENDERECO = Mid(XFO_EXPORTACAO!clie_tx_EndResidencial, 1, 40)
              XLT_BAIRRO = Mid(XFO_EXPORTACAO!clie_tx_BairroResidencial, 1, 12)
              XLT_CEP = Mid(XFO_EXPORTACAO!focl_tx_Cep, 1, 2) + Mid(XFO_EXPORTACAO!focl_tx_Cep, 4, 3) + Mid(XFO_EXPORTACAO!focl_tx_Cep, 8, 3)
              XLT_CIDADE = Mid(XFO_EXPORTACAO!clie_tx_MunResidencial, 1, 15)
              XLT_ESTADO = Mid(XFO_EXPORTACAO!clie_tx_EstResidencial, 1, 2)
            End If
            
          Else
            XLT_ENDERECO = Mid(XFO_EXPORTACAO!clie_tx_EndCorresp, 1, 40)
            XLT_BAIRRO = Mid(XFO_EXPORTACAO!clie_tx_BairroCorresp, 1, 12)
            XLT_CEP = Mid(XFO_EXPORTACAO!clie_nr_CepCorresp, 1, 2) + Mid(XFO_EXPORTACAO!clie_nr_CepCorresp, 4, 3) + Mid(XFO_EXPORTACAO!clie_nr_CepCorresp, 8, 3)
            XLT_CIDADE = Mid(XFO_EXPORTACAO!clie_tx_MunCorresp, 1, 15)
            XLT_ESTADO = Mid(XFO_EXPORTACAO!clie_tx_EstCorresp, 1, 2)
          End If
          
          XLT_REGLOTE = XLT_REGLOTE + 1
          XLT_TITULO = Mid(XFO_EXPORTACAO!Titulo, 1, 4) & Mid(XFO_EXPORTACAO!Titulo, 6, 4) & Mid(XFO_EXPORTACAO!Titulo, 11, 2) & Mid(XFO_EXPORTACAO!Titulo, 14, 2) & Mid(XFO_EXPORTACAO!Titulo, 17, 3) & Mid(XFO_EXPORTACAO!Titulo, 21, 2)
                    
          '*****************Descrição de registro tipo '1' (Obrigatório) - Dados do título - Detalhe de Remessa*****************
          'CÓDIGO DO REGISTRO - PREENCHER COM 1
          XLT_TEXTO = "1"
          'TIPO INSCRIÇÃO - PJ OU PF
          XLT_TEXTO = XLT_TEXTO & "02"
          'NÚMERO INSCRIÇÃO - CGC
          XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(XLT_CGCEMP) = "", FuncoesGenericas.FunZeros(14), (FuncoesGenericas.FunZeros(14 - Len(XLT_CGCEMP))) & Format(XLT_CGCEMP, "00000000000000"))
          'CÓDIGO DA AGÊNCIA DE VINCULAÇÃO DO BENEFICIÁRIO
          XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(DatContaCorrente.Recordset.Fields!coco_cd_Agencia) = "", FuncoesGenericas.FunZeros(4), (FuncoesGenericas.FunZeros(4 - Len(DatContaCorrente.Recordset.Fields!coco_cd_Agencia))) & Format(DatContaCorrente.Recordset.Fields!coco_cd_Agencia, "0000"))
          'CÓDIGO DO BENEFICIÁRIO
          XLT_TEXTO = XLT_TEXTO & DatContaCorrente.Recordset.Fields!coco_nr_conveniocobranca
          'ID EMISSÃO
          XLT_TEXTO = XLT_TEXTO & "1"
          'ID POSTAGEM
          XLT_TEXTO = XLT_TEXTO & "1"
          'TAXA PERMANÊNCIA
          XLT_TEXTO = XLT_TEXTO & FuncoesGenericas.FunZeros(2)
          'USO EMPRESA BENEFICIÁRIO
          XLT_TEXTO = XLT_TEXTO & XLT_TITULO & Space(25 - Len(XLT_TITULO))
          'MODALIDADE IDENTIFICAÇÃO E IDENTIFICAÇÃO DO TÍTULO NA CAIXA
          XLT_TEXTO = XLT_TEXTO & FuncoesGenericas.FunZeros(2)
          'IDENTIFICAÇÃO DO TÍTULO NA CAIXA
          XLT_TEXTO = XLT_TEXTO & FuncoesGenericas.FunZeros(15)
          'CAMPOS EM BRANCOS
          XLT_TEXTO = XLT_TEXTO & Space(2)
          'USO LIVRE BANCO/EMPRESA
          XLT_TEXTO = XLT_TEXTO & Space(1)
          'BRANCOS
          XLT_TEXTO = XLT_TEXTO & Space(1)
          'DATA DE JUROS
          XLT_TEXTO = XLT_TEXTO & Format(DateAdd("d", 1, XFO_EXPORTACAO!titu_dt_Vencimento), "ddMMyy")
          'CÓDIGO DO DESCONTO
          XLT_TEXTO = XLT_TEXTO & "0"
          'BRANCOS
          XLT_TEXTO = XLT_TEXTO & Space(22)
          'CÓDIGO DA CARTEIRA - PREENCHER COM 01
          XLT_TEXTO = XLT_TEXTO & "01"
          'IDENTIFICAÇÃO TIPO OCORRÊNCIA DO ARQUIVO REMESSA
          XLT_TEXTO = XLT_TEXTO & "01"
          'NÚMERO DO DOCUMENTO DE COBRANÇA
          XLT_TEXTO = XLT_TEXTO & Left(XFO_EXPORTACAO!Titulo, 4) & Mid(XFO_EXPORTACAO!Titulo, 6, 4) & Mid(XFO_EXPORTACAO!Titulo, 11, 2)
          'DATA DE VENCIMENTO DO TÍTULO
          XLT_TEXTO = XLT_TEXTO & Format(XFO_EXPORTACAO!titu_dt_Vencimento, "ddMMyy")
          'VALOR NOMINAL DO TÍTULO
          XLT_VALORTEMP = Replace(XLF_VALOR, ",", ".")
          XLT_TEXTO = XLT_TEXTO & Format(XLT_VALORTEMP, "0000000000000")
          'CÓDIGO DO BANCO
          XLT_TEXTO = XLT_TEXTO & "104"
          'AGÊNCIA COBRADORA
          XLT_TEXTO = XLT_TEXTO & "00000"
          'ESPÉCIE DE TÍTULO
          XLT_TEXTO = XLT_TEXTO & "02"
          'ACEITE
          XLT_TEXTO = XLT_TEXTO & "A"
          'DATA EMISSÃO TÍTULO
          XLT_TEXTO = XLT_TEXTO & Format(DateTime.Now, "ddMMyy")
          'INSTRUÇÃO 1
          XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(TxtMensagem1.Text) = "", "02", (FuncoesGenericas.FunZeros(2 - Len(TxtMensagem1.Text)))) & TxtMensagem1.Text
          'INSTRUÇÃO 2
          XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(TxtMensagem2.Text) = "", FuncoesGenericas.FunZeros(2), (FuncoesGenericas.FunZeros(2 - Len(TxtMensagem2.Text)))) & TxtMensagem2.Text
          'JUROS MORA
          XLO_PERCJUROS = XLF_VALOR * (XFO_EXPORTACAO!titu_vl_PercJurosMora / 100)
          XLO_PERCJUROS = Format(XLO_PERCJUROS, "0.00")
          XLO_PERCJUROS = Replace(XLO_PERCJUROS, ",", ".")
          XLT_TEXTO = XLT_TEXTO & Format(XLO_PERCJUROS, "0000000000000")
          'DATA DO DESCONTO
          XLT_TEXTO = XLT_TEXTO & FuncoesGenericas.FunZeros(6)
          'VALOR/PERCENTUAL DO DESCONTO
          XLT_TEXTO = XLT_TEXTO & FuncoesGenericas.FunZeros(13)
          'VALOR DO IOF
          XLT_TEXTO = XLT_TEXTO & FuncoesGenericas.FunZeros(13)
          'ABATIMENTO
          XLT_TEXTO = XLT_TEXTO & FuncoesGenericas.FunZeros(13)
          'TIPO INSCRIÇÃO
          XLT_TEXTO = XLT_TEXTO & XLT_TIPO
          'NÚMERO INSCRIÇÃO
          XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(XLT_CGCCPF) = "", FuncoesGenericas.FunZeros(14), (FuncoesGenericas.FunZeros(14 - Len(XLT_CGCCPF)))) & XLT_CGCCPF
          'NOME
          XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(XLT_NOME) = "", Space(40), XLT_NOME) & Space(40 - Len(XLT_NOME))
          'ENDEREÇO
          XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(XLT_ENDERECO) = "", Space(40), XLT_ENDERECO) & Space(40 - Len(XLT_ENDERECO))
          'BAIRRO
          XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(XLT_BAIRRO) = "", Space(12), XLT_BAIRRO) & Space(12 - Len(XLT_BAIRRO))
          'CEP
          XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(XLT_CEP) = "", FuncoesGenericas.FunZeros(8), (FuncoesGenericas.FunZeros(8 - Len(XLT_CEP)))) & XLT_CEP
          'CIDADE
          XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(XLT_CIDADE) = "", Space(15), XLT_CIDADE) & Space(15 - Len(XLT_CIDADE))
          'ESTADO
          XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(XLT_ESTADO) = "", Space(2), XLT_ESTADO) & Space(2 - Len(XLT_ESTADO))
          'DATA DA MULTA
          XLT_TEXTO = XLT_TEXTO & Format(DateAdd("d", 1, XFO_EXPORTACAO!titu_dt_Vencimento), "ddMMyy")
          'VALOR DA MULTA
          XLT_VALORMULTA = XLF_VALOR * 0.02
          XLT_VALORMULTA = FormatNumber(XLT_VALORMULTA, 2)
          XLT_VALORMULTA = Replace(XLT_VALORMULTA, ",", ".")
          XLT_TEXTO = XLT_TEXTO & Format(XLT_VALORMULTA, "0000000000")
          'SACADOR/AVALISTA
          XLT_TEXTO = XLT_TEXTO & IIf(FunNulo(XLT_NOME) = "", Space(22), Trim(Mid(XLT_NOME, 1, 22))) & Space(22 - Len(Trim(Mid(XLT_NOME, 1, 22))))
          'INSTRUÇÃO 3
          XLT_TEXTO = XLT_TEXTO & FuncoesGenericas.FunZeros(2)
          'PRAZO
          'XLT_TEXTO = XLT_TEXTO & FuncoesGenericas.FunZeros(2 - Len(XLI_PRAZOMORA)) & IIf(FunNulo(XLI_PRAZOMORA) = "", "05", Format(XLI_PRAZOMORA, "00"))
          XLT_TEXTO = XLT_TEXTO & "30"
          'CÓDIGO DA MOEDA - PREENCHER COM 1
          XLT_TEXTO = XLT_TEXTO & "1"
          'NÚMERO SEQUENCIAL
          XLT_TEXTO = XLT_TEXTO & Format(XLI_NUMSQUENCIAL, "000000")

          Print #1, funTiraAcento(XLT_TEXTO, True)
          
          XLT_REGLOTE = XLT_REGLOTE + 1
          
          'Pegando o código da moeda
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
            XGM_MATRIZLOG(3, 0) = "Valor do Desconto"
            XGM_MATRIZLOG(4, 0) = "Valor do título"
            XGM_MATRIZLOG(5, 0) = "Conta corrente"
            XGM_MATRIZLOG(6, 0) = "Data Venc. Título"
            XGM_MATRIZLOG(7, 0) = "Data Venc. Desconto"
            
            XGM_MATRIZLOG(0, 1) = XFO_EXPORTACAO!Titulo
            XGM_MATRIZLOG(1, 1) = FunNuloBancoVal(XLT_CODMOEDA)
            XGM_MATRIZLOG(2, 1) = FunNuloVal(FunTrataFloat(TDBGrid1.Columns(8)))
            XGM_MATRIZLOG(3, 1) = XLF_DESCONTO
            XGM_MATRIZLOG(4, 1) = XLF_VALOR
            XGM_MATRIZLOG(5, 1) = DatContaCorrente.Recordset.Fields!coco_cd_Agencia & " - " & DatContaCorrente.Recordset.Fields!coco_tx_Conta & "-" & FunNuloVal(DatContaCorrente.Recordset.Fields!coco_nr_Dag)
            XGM_MATRIZLOG(6, 1) = XFO_EXPORTACAO!titu_dt_Vencimento
            XGM_MATRIZLOG(7, 1) = XLT_DATADESCONTO
            
            Call subRegistraLog("TelaExpBancaria", "1", funCriaDescricaoLog(XGM_MATRIZLOG, EXPORTACAO_CAPI, PAGAMENTO_ELETRONICO_CAIXA))
            '*******************************************************************
          
      End If
      
      TDBGrid1.MoveNext
      XFO_EXPORTACAO.MoveNext
  Wend
  
  XLI_NUMSQUENCIAL = XLI_NUMSQUENCIAL + 1
  
  'REGISTRO TRAILLER - REMESSA
  XLT_TEXTO = "9"
  XLT_TEXTO = XLT_TEXTO & Space(393)
  XLT_TEXTO = XLT_TEXTO & Format(XLI_NUMSQUENCIAL, "000000")
  
  Print #1, funTiraAcento(XLT_TEXTO, True)
  '**********************************************
  Close #1
  
  Conexao.Execute "UPDATE Ultimos Set Ulti_nr_RemessaCaixa = (Ulti_nr_RemessaCaixa+1)"
  
  MsgBox "Exportação realizada com sucesso!", vbInformation + vbOKOnly, "ATENÇÃO"
  
  Conexao.CommitTrans
  subDesabilitaBotoes
  btnImprimir.Enabled = True
    
  CmbInstrucao.Enabled = True
  
  Exit Sub
  
TrataErro:
    funTrataErros (ComMensagem)
    Conexao.RollbackTrans
    Exit Sub
  
End Sub
Private Sub btnSelecoes_Click()
    
    Dim XLI_POS As Integer
    Dim XLI_POS2 As Integer
    
    Set Formulario = TelaExpBancaria
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"

    If Filtrou = True Then
        subCarregaVetor FRM_FiltroAtual, Array("Titulo", "Exporta", "titu_dt_Vencimento", "ValorReal", _
     "focl_tx_RazaoSocial", "titu_vl_Desconto", "LimiteDesconto", "ValorTitulo", _
     "Seguro", "SaldoDevedor", "titu_nr_contratobanco", "DescDiario", "DataConsDesc", "ValAbatimento", "MultaJuros", "titu_dt_Base"), VFV_VETOREXP, TDBGrid1
              
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

Private Sub btnImprimir_Click()
    'Call subImprimeListagemGRIDUnBound(2, TDBGrid1.PrintInfo, "Listagem Títulos Exportados")
        
    On Error GoTo TrataErro:
    Dim GuardaY As Single
    Dim LargPapel As Single, AltPapel As Single, AreaImpressao As Single
    'Dim XNumBordero As String
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
        'TDBGrid1.MoveFirst
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
                'subQuebraTexto FunNulo(RecebeCampo), 11.7, 2.3, 7, "Arial", False, False
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 11.55, 2.3, 7, "Arial", False, False), 2.3, 7, "Arial", False, False
                
                Printer.CurrentY = GuardaY
                If TDBGrid1.Columns("Desconto") = "" Then
                    TDBGrid1.Columns("Desconto") = 0#
                End If
                RecebeCampo = TDBGrid1.Columns("Desconto")
                'subQuebraTexto FunNulo(RecebeCampo), 14.2, 1.8, 7, "Arial", False, False
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 13.8, 1.8, 7, "Arial", False, False), 1.8, 7, "Arial", False, False
                
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Val. Seguro R$")
                'subQuebraTexto FunNulo(RecebeCampo), 16.2, 1.3, 7, "Arial", False, False
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 15.95, 1.3, 7, "Arial", False, False), 1.3, 7, "Arial", False, False
               
                Printer.CurrentY = GuardaY
                RecebeCampo = Format(CStr(CCur(TDBGrid1.Columns("Val. Tit. R$")) + CCur(TDBGrid1.Columns("Val. Seguro R$")) - Format(CStr(CCur(TDBGrid1.Columns("Desconto"))))), "standard")
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
        'GuardaY = Printer.CurrentY + 1#
        'Printer.CurrentY = GuardaY
    
        Printer.CurrentY = GuardaY
        RecebeCampo = Format$(XLF_TOTALPAGAR, "STANDARD")
        subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 17.5, 2.3, 7, "Arial", False, False), 2.3, 7, "Arial", True, False
        If QuebraY < Printer.CurrentY Then
            QuebraY = Printer.CurrentY
        End If
        'GuardaY = Printer.CurrentY + 1#
        'Printer.CurrentY = GuardaY
        
        'subImprimeLinha 1.5, 20, 3
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
  Dim XGT_SELECAO As String  'Armazena condições para o filtro
  Dim XFI_CONT As Integer
  Dim XLT_NATUREZA As String

  XGT_SELECAO = ""
  XGT_CONJUNCAO = ""
  
  If FunObrigatorioCBO(CboCCorrente, "Selecione uma Conta Corrente.") Then Exit Sub
  '25/03/10 - Patrícia
  If Me.TxtNDiasProtesto.Text <> "" Then
    If Not IsNumeric(Me.TxtNDiasProtesto.Text) Then
        MsgBox "O Nº de Dias para Protesto não é um valor numérico. ", vbCritical, "ExpCAIXA"
        Exit Sub
    Else
        If Val(TxtNDiasProtesto.Text) <> 0 Then
           If Val(TxtNDiasProtesto.Text) < 5 Then
                MsgBox "O Nº de Dias para Protesto deverá ser no mínimo 5. ", vbCritical, "ExpCAIXA"
                Exit Sub
           Else
                CmbInstrucao.Enabled = False
           End If
        End If
    End If
  End If
  
  If Me.CboTipoCobranca.ListIndex = -1 Then
    MsgBox "Escolha o tipo de cobrança.", vbCritical, "ExpCAIXA"
    Exit Sub
  End If
  
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
  XGT_CONJUNCAO = " AND "
     
  'Filtra os titulos que não foram exportados
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
     " (titu_tx_ExpBanco = 'N' or titu_tx_ExpBanco is null)"
  XGT_CONJUNCAO = " AND "
     
  'Filtra os titulos que não foram pagos
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
     " titu_dt_Pagamento is null "
  XGT_CONJUNCAO = " AND "
     
  'Filtra os ativos
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
     " cont_tx_Status='A' "
  XGT_CONJUNCAO = " AND "
    
  '25/03/10 - Patrícia
  XFT_SQL = "SELECT *, " & _
                  "-1 as Exporta," & _
                  "0 as ValorReal," & _
                  "'' as LimiteDesconto," & _
                  "0 as ValorTitulo, " & _
                  "0 as Seguro, " & _
                  "'N' as SaldoDevedor, " & _
                  "'00,00' as DescDiario, " & _
                  "'' as DataConsDesc, " & _
                  "'0,00' as ValAbatimento, " & _
                  "0 as MultaJuros " & _
              "FROM ConsCAPExpBancoCaixa " & _
              "WHERE" & XGT_SELECAO & "ORDER BY Titulo"
   
   subCarregaVetor XFT_SQL, Array("Titulo", "Exporta", "titu_dt_Vencimento", "ValorReal", _
     "focl_tx_RazaoSocial", "titu_vl_Desconto", "LimiteDesconto", "ValorTitulo", _
     "Seguro", "SaldoDevedor", "titu_nr_contratobanco", "DescDiario", "DataConsDesc", "ValAbatimento", "MultaJuros", "titu_dt_Base"), VFV_VETOREXP, TDBGrid1
     
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
  
  XFT_SQL2 = FunCriaConsultaBase("01/01/01", NomeSgbd, "ConsCAPExpBancoCaixa", 3)
  XFT_SQL2 = XFT_SQL2 & XGT_SELECAO
  XFT_SQL2 = XFT_SQL2 & " ORDER BY Titulo" '18/06/10

  If XFO_EXPORTACAO.State = adStateOpen Then
    XFO_EXPORTACAO.Close
  End If
  
  SubQOpenRecordset XFO_EXPORTACAO, XFT_SQL2, Estatico
  
  subCarregaVariaveis XFT_SQL
  
  SubCalculaValorAtualizado
  
  SubCalculaSeguro
    
  TDBGrid1.Refresh

  btnExportar.Enabled = True
  btnDesconto.Enabled = True
  filtra_especial = True
  subTelaValoresGlobais "G"
  Cmdfechar_Click
  
End Sub

Private Sub btnSair_Click()
    Unload Me
'    Unload MdiPrincipal
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
              "Seguro", "SaldoDevedor", "titu_nr_contratobanco"), VFV_VETOREXP, TDBGrid1, Formulario, X, XFT_SQL
            
            subHabilitaBotoes
            subTelaValoresGlobais "G"
            
            If VFV_VETOREXP(0, 0) <> Empty Then
              subHabilitaBotoes
              TDBGrid1.Enabled = True
              TDBGrid1.MoveFirst
            Else
              MsgBox "Não existem pagamentos pendentes com esses filtros!", vbInformation + vbOKOnly, "ATENÇÃO"
              btnDesmarcarTodos.Enabled = False
              btnMarcarTodos.Enabled = False
              btnExportar.Enabled = False
              btnDesconto.Enabled = False
              btnImprimir.Enabled = False
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
    
    'Call AjustaTela

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
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & Int(PCodEmpresa) & " AND banc_cd_codigo=104 ORDER BY coco_tx_Descricao", Estatico
    
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
    Me.CboTipoCobranca.ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '*subManutencaoJanelasAtivas "R", "TelaExpBancaria"
   Unload Me
End Sub

Private Sub CboCCorrente_Change()
  
  If CboCCorrente.BoundText <> "" Then
    DatContaCorrente.Recordset.Bookmark = CboCCorrente.SelectedItem
    TxtMensagem1.Text = FunNulo(DatContaCorrente.Recordset.Fields("coco_tx_Mensagem1"))
    TxtMensagem2.Text = FunNulo(DatContaCorrente.Recordset.Fields("coco_tx_Mensagem2"))
  End If
    
End Sub

Private Sub TxtNDiasProtesto_LostFocus()
  '25/03/10 - Patrícia
  If Me.TxtNDiasProtesto.Text <> "" Then
    If Not IsNumeric(Me.TxtNDiasProtesto.Text) Then
        MsgBox "O Nº de Dias para Protesto não é um valor numérico. ", vbCritical, "ExpCAIXA"
        Exit Sub
    Else
        If Val(TxtNDiasProtesto.Text) <> 0 Then
           If Val(TxtNDiasProtesto.Text) < 5 Then
                MsgBox "O Nº de Dias para Protesto deverá ser no mínimo 5. ", vbCritical, "ExpCAIXA"
                Exit Sub
           Else
                CmbInstrucao.Text = ""
                CmbInstrucao.Enabled = False
           End If
        Else
            CmbInstrucao.Enabled = True
        End If
    End If
  End If

End Sub

Private Sub TxtPrzMora_KeyPress(KeyAscii As Integer)
  
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
    KeyAscii = 0
  End If
  
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaBancos 
   Caption         =   "Tabela de Bancos"
   ClientHeight    =   4545
   ClientLeft      =   435
   ClientTop       =   1260
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   9315
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSPanel PanTipo 
      Height          =   2700
      Left            =   345
      TabIndex        =   20
      Top             =   795
      Visible         =   0   'False
      Width           =   8235
      _Version        =   65536
      _ExtentX        =   14526
      _ExtentY        =   4762
      _StockProps     =   15
      Caption         =   " Cadastro de Banco"
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Extenso"
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
         Left            =   4455
         TabIndex        =   30
         Top             =   900
         Width           =   3615
         Begin MSMask.MaskEdBox MskLinExtenso 
            Height          =   255
            Left            =   960
            TabIndex        =   11
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Format          =   "0#.#0"
            Mask            =   "##,##"
            PromptChar      =   "0"
         End
         Begin MSMask.MaskEdBox MskColExtenso 
            Height          =   255
            Left            =   2640
            TabIndex        =   12
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Format          =   "0#.#0"
            Mask            =   "##,##"
            PromptChar      =   "0"
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Linha:"
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
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Coluna:"
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
            Left            =   1800
            TabIndex        =   31
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   615
         Left            =   135
         TabIndex        =   27
         Top             =   900
         Width           =   3615
         Begin MSMask.MaskEdBox MskColValor 
            Height          =   255
            Left            =   2640
            TabIndex        =   10
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Format          =   "0#.#0"
            Mask            =   "##,##"
            PromptChar      =   "0"
         End
         Begin MSMask.MaskEdBox MskLinValor 
            Height          =   255
            Left            =   960
            TabIndex        =   9
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Format          =   "0#.#0"
            Mask            =   "##,##"
            PromptChar      =   "0"
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Linha:"
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
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Coluna:"
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
            Left            =   1800
            TabIndex        =   28
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox TxtDescricao 
         Height          =   285
         Left            =   2790
         MaxLength       =   50
         TabIndex        =   8
         Top             =   525
         Width           =   5295
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
         Height          =   375
         Left            =   5550
         TabIndex        =   17
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CommandButton CmdCancelar 
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
         Height          =   375
         Left            =   6870
         TabIndex        =   18
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data"
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
         Left            =   4455
         TabIndex        =   24
         Top             =   1500
         Width           =   3615
         Begin MSMask.MaskEdBox MskLinData 
            Height          =   255
            Left            =   960
            TabIndex        =   15
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Format          =   "0#.#0"
            Mask            =   "##,##"
            PromptChar      =   "0"
         End
         Begin MSMask.MaskEdBox MskColData 
            Height          =   255
            Left            =   2640
            TabIndex        =   16
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Format          =   "0#.#0"
            Mask            =   "##,##"
            PromptChar      =   "0"
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Linha:"
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
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Coluna:"
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
            Left            =   1800
            TabIndex        =   25
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nominal"
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
         Left            =   135
         TabIndex        =   21
         Top             =   1500
         Width           =   3615
         Begin MSMask.MaskEdBox MskLinNominal 
            Height          =   255
            Left            =   960
            TabIndex        =   13
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Format          =   "0#.#0"
            Mask            =   "##,##"
            PromptChar      =   "0"
         End
         Begin MSMask.MaskEdBox MskColNominal 
            Height          =   255
            Left            =   2640
            TabIndex        =   14
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Format          =   "0#.#0"
            Mask            =   "##,##"
            PromptChar      =   "0"
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Linha:"
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
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Coluna:"
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
            Left            =   1800
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
      End
      Begin MSMask.MaskEdBox MskCodigo 
         Height          =   285
         Left            =   870
         TabIndex        =   7
         Top             =   525
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   255
         Left            =   165
         TabIndex        =   34
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descrição:"
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
         Left            =   1470
         TabIndex        =   33
         Top             =   540
         Width           =   1215
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaBancos.frx":0000
      Height          =   3630
      Left            =   45
      TabIndex        =   37
      Top             =   15
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   6403
      _LayoutType     =   4
      _RowHeight      =   26
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "banc_cd_codigo"
      Columns(0).NumberFormat=   "000"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descrição"
      Columns(1).DataField=   "banc_tx_descricao"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Linha Valor"
      Columns(2).DataField=   "banc_vl_linvalor"
      Columns(2).NumberFormat=   "Standard"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Coluna Valor"
      Columns(3).DataField=   "banc_vl_colValor"
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Linha Extenso"
      Columns(4).DataField=   "banc_vl_LinExt"
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Coluna Extenso"
      Columns(5).DataField=   "banc_vl_colext"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Linha Nominal"
      Columns(6).DataField=   "banc_vl_linNom"
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Coluna Nominal"
      Columns(7).DataField=   "banc_vl_ColNom"
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Linha Data"
      Columns(8).DataField=   "banc_vl_linData"
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Coluna Data"
      Columns(9).DataField=   "banc_vl_coldata"
      Columns(9).NumberFormat=   "Standard"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1296"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1217"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=49"
      Splits(0)._ColumnProps(10)=   "Column(1).Width=3810"
      Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=3731"
      Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=528"
      Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(16)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(17)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(18)=   "Column(1)._MinWidth=55314416"
      Splits(0)._ColumnProps(19)=   "Column(2).Width=1667"
      Splits(0)._ColumnProps(20)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._WidthInPix=1588"
      Splits(0)._ColumnProps(22)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(2)._ColStyle=530"
      Splits(0)._ColumnProps(24)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(25)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(26)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(27)=   "Column(3).Width=1614"
      Splits(0)._ColumnProps(28)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(3)._WidthInPix=1535"
      Splits(0)._ColumnProps(30)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(31)=   "Column(3)._ColStyle=530"
      Splits(0)._ColumnProps(32)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(33)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(34)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(35)=   "Column(4).Width=1614"
      Splits(0)._ColumnProps(36)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(4)._WidthInPix=1535"
      Splits(0)._ColumnProps(38)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(4)._ColStyle=530"
      Splits(0)._ColumnProps(40)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(41)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(42)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(43)=   "Column(5).Width=1588"
      Splits(0)._ColumnProps(44)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(5)._WidthInPix=1508"
      Splits(0)._ColumnProps(46)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(5)._ColStyle=530"
      Splits(0)._ColumnProps(48)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(49)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(50)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(51)=   "Column(6).Width=1588"
      Splits(0)._ColumnProps(52)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(6)._WidthInPix=1508"
      Splits(0)._ColumnProps(54)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(55)=   "Column(6)._ColStyle=530"
      Splits(0)._ColumnProps(56)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(57)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(58)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(59)=   "Column(7).Width=1588"
      Splits(0)._ColumnProps(60)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(7)._WidthInPix=1508"
      Splits(0)._ColumnProps(62)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(63)=   "Column(7)._ColStyle=530"
      Splits(0)._ColumnProps(64)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(65)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(66)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(67)=   "Column(8).Width=1640"
      Splits(0)._ColumnProps(68)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(8)._WidthInPix=1561"
      Splits(0)._ColumnProps(70)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(71)=   "Column(8)._ColStyle=530"
      Splits(0)._ColumnProps(72)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(73)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(74)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(75)=   "Column(8)._MinWidth=-32"
      Splits(0)._ColumnProps(76)=   "Column(9).Width=1667"
      Splits(0)._ColumnProps(77)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(9)._WidthInPix=1588"
      Splits(0)._ColumnProps(79)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(80)=   "Column(9)._ColStyle=530"
      Splits(0)._ColumnProps(81)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(82)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(83)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(84)=   "Column(9)._MinWidth=-3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=128,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14,.alignment=2"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=1"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14,.alignment=2"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(76)  =   "Named:id=33:Normal"
      _StyleDefs(77)  =   ":id=33,.parent=0"
      _StyleDefs(78)  =   "Named:id=34:Heading"
      _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(80)  =   ":id=34,.wraptext=-1"
      _StyleDefs(81)  =   "Named:id=35:Footing"
      _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   "Named:id=36:Selected"
      _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=37:Caption"
      _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(87)  =   "Named:id=38:HighlightRow"
      _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=39:EvenRow"
      _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(91)  =   "Named:id=40:OddRow"
      _StyleDefs(92)  =   ":id=40,.parent=33"
      _StyleDefs(93)  =   "Named:id=41:RecordSelector"
      _StyleDefs(94)  =   ":id=41,.parent=34"
      _StyleDefs(95)  =   "Named:id=42:FilterBar"
      _StyleDefs(96)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7035
      Top             =   3210
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin Threed.SSPanel PanBotoes 
      Height          =   390
      Left            =   120
      TabIndex        =   19
      Tag             =   "0"
      Top             =   4095
      Width           =   11280
      _Version        =   65536
      _ExtentX        =   19897
      _ExtentY        =   688
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin VB.CommandButton CmdFormasPagto 
         Caption         =   "&Formas de Pagto"
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
         TabIndex        =   39
         Top             =   15
         Width           =   1695
      End
      Begin VB.CommandButton cmdExport 
         Height          =   330
         Left            =   10680
         Picture         =   "TelaBancos.frx":0015
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   15
         Width           =   375
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
         Left            =   9500
         TabIndex        =   6
         Top             =   15
         Width           =   1020
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
         Left            =   6330
         TabIndex        =   5
         Top             =   15
         Width           =   1020
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
         Left            =   5070
         TabIndex        =   4
         Top             =   15
         Width           =   1020
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
         Left            =   3795
         TabIndex        =   3
         Top             =   15
         Width           =   1020
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
         Left            =   2535
         TabIndex        =   2
         Top             =   15
         Width           =   1020
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
         Left            =   30
         TabIndex        =   0
         Top             =   15
         Width           =   1020
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
         Left            =   1275
         TabIndex        =   1
         Top             =   15
         Width           =   1020
      End
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaBancos.frx":01C4
      Height          =   315
      Left            =   2850
      TabIndex        =   35
      Top             =   3705
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "banc_cd_codigo"
      BoundColumn     =   "banc_cd_codigo"
      Text            =   ""
   End
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Código:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   90
      TabIndex        =   36
      Top             =   3750
      Width           =   2670
   End
End
Attribute VB_Name = "TelaBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Tipos() As String * 1
Dim Quantid As Long
Dim ResFormulario As Object
Dim KeyAscii As Integer
Dim errLoop As ADODB.Error
Dim StrError As String


Dim Chave As Integer

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
    Dim SN As Boolean
    If FunTabelaVazia(Adodc1) Then
        SN = False
    Else
        SN = True
    End If

    'Objetivo: Habilita os Botões da Tela
    CmdInserir.Enabled = True
    CmdAlterar.Enabled = SN
    CmdExcluir.Enabled = SN
    CmdOrdem.Enabled = SN
    CmdImprimir.Enabled = SN
    CmdSair.Enabled = True
    
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
    CmdExcluir.Enabled = False
    CmdAlterar.Enabled = False
    TDBGrid1.Enabled = False
    CmdRemoverFiltro.Enabled = False
    CmdSair.Enabled = False
    CmdImprimir.Enabled = False
    CmdOrdem.Enabled = False
End Sub
Sub subPreencheCampos()
    
    MskCodigo.Text = FunNulo(Format$(Adodc1.Recordset.Fields("banc_cd_Codigo"), "000"))
    TxtDescricao.Text = FunNulo(Adodc1.Recordset.Fields("banc_tx_Descricao"))

    MskLinValor.Text = FunNulo(Format$(Adodc1.Recordset.Fields("banc_vl_LinValor"), "0#.#0"))
    MskColValor.Text = FunNulo(Format$(Adodc1.Recordset.Fields("banc_vl_ColValor"), "0#.#0"))

    MskLinExtenso.Text = FunNulo(Format$(Adodc1.Recordset.Fields("banc_vl_LinExt"), "0#.#0"))
    MskColExtenso.Text = FunNulo(Format$(Adodc1.Recordset.Fields("banc_vl_ColExt"), "0#.#0"))

    MskLinNominal.Text = FunNulo(Format$(Adodc1.Recordset.Fields("banc_vl_LinNom"), "0#.#0"))
    MskColNominal.Text = FunNulo(Format$(Adodc1.Recordset.Fields("banc_vl_ColNom"), "0#.#0"))

    MskLinData.Text = FunNulo(Format$(Adodc1.Recordset.Fields("banc_vl_LinData"), "0#.#0"))
    MskColData.Text = FunNulo(Format$(Adodc1.Recordset.Fields("banc_vl_ColData"), "0#.#0"))
End Sub

Sub subPrepararInsercao()
    TxtDescricao.Text = ""

    MskCodigo.Mask = ""
    MskCodigo.Text = ""
    MskCodigo.Mask = "###"

    MskLinValor.Mask = ""
    MskLinValor.Text = ""
    MskLinValor.Mask = "##,##"
    MskColValor.Mask = ""
    MskColValor.Text = ""
    MskColValor.Mask = "##,##"

    MskLinExtenso.Mask = ""
    MskLinExtenso.Text = ""
    MskLinExtenso.Mask = "##,##"
    MskColExtenso.Mask = ""
    MskColExtenso.Text = ""
    MskColExtenso.Mask = "##,##"

    MskLinNominal.Mask = ""
    MskLinNominal.Text = ""
    MskLinNominal.Mask = "##,##"
    MskColNominal.Mask = ""
    MskColNominal.Text = ""
    MskColNominal.Mask = "##,##"

    MskLinData.Mask = ""
    MskLinData.Text = ""
    MskLinData.Mask = "##,##"
    MskColData.Mask = ""
    MskColData.Text = ""
    MskColData.Mask = "##,##"

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


Private Sub cboLocalizar_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If cboLocalizar.BoundText <> "" Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
End Sub


Private Sub CmdAlterar_Click()
    If Not FunVerifica_Permissao("TelaBancos", "3") Then
        Exit Sub
    End If
    
    TelaBancos.MousePointer = vbHourglass
    XInserir = "A"
    Chave = Adodc1.Recordset("banc_cd_codigo")
    PanBotoes.Enabled = False
    PanTipo.Visible = True
    PanTipo.Enabled = True
    PanTipo.Top = (TDBGrid1.Width - PanTipo.Width) / 2
    PanTipo.Left = (TDBGrid1.Height - PanTipo.Height) / 2
    PanTipo.Caption = " Alteração de Banco"
    subPreencheCampos
    MskCodigo.Enabled = False
    TxtDescricao.SetFocus
    TelaBancos.MousePointer = vbDefault
End Sub
Private Sub cmdCancelar_Click()
    If Not FunTabelaVazia(Adodc1) Then Adodc1.Recordset.CancelUpdate
    PanBotoes.Enabled = True
    PanTipo.Visible = False
    PanTipo.Enabled = True
    MskCodigo.Enabled = True
    CmdConfirmar.Caption = "Confirmar"
    subHabilitaBotoes
'    DataGrid1.Height = 4260
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub

Private Sub CmdConfirmar_Click()
    
    Dim Sql As String
    Dim nreg As Variant
    Dim XStatus As String
    Dim CodEquip As Integer
    Dim IntLoop As Integer
    Dim X As String
    
    On Error GoTo MostraErro
    
    'Função para verificar preenchimento dos campos obrigatórios
    
    If funChecarCampos Then Exit Sub
    
    Conexao.BeginTrans
    If XInserir = "I" Then
        Sql = "INSERT INTO Bancos (banc_cd_codigo,banc_tx_descricao,banc_vl_LinValor,banc_vl_ColValor,banc_vl_LinExt,banc_vl_ColExt,banc_vl_LinNom,banc_vl_ColNom,banc_vl_LinData,banc_vl_ColData) " & _
              "VALUES ('" & MskCodigo.Text & "','" & TxtDescricao.Text & "'," & FunNuloVal(FunTrataFloat(MskLinValor.Text)) & "," & FunNuloVal(FunTrataFloat(MskColValor.Text)) & "," & FunNuloVal(FunTrataFloat(MskLinExtenso.Text)) & "," & FunNuloVal(FunTrataFloat(MskColExtenso.Text)) & "," & FunNuloVal(FunTrataFloat(MskLinNominal.Text)) & "," & FunNuloVal(FunTrataFloat(MskColNominal.Text)) & "," & FunNuloVal(FunTrataFloat(MskLinData.Text)) & "," & FunNuloVal(FunTrataFloat(MskColData.Text)) & ")"
        Conexao.Execute Sql
    
        MdiPrincipal.BarraStatus.Panels(3).Text = CStr((Adodc1.Recordset.RecordCount + 1)) + "/" + CStr((Adodc1.Recordset.RecordCount + 1))
    Else    'Alteração
    
        Sql = "UPDATE Bancos " & _
                     "SET banc_cd_codigo = '" & MskCodigo.Text & "', " & _
                         " banc_tx_descricao = '" & TxtDescricao.Text & "'," & _
                         " banc_vl_LinValor = " & FunNuloVal(FunTrataFloat(MskLinValor.Text)) & "," & _
                         " banc_vl_ColValor = " & FunNuloVal(FunTrataFloat(MskColValor.Text)) & "," & _
                         " banc_vl_LinExt = " & FunNuloVal(FunTrataFloat(MskLinExtenso.Text)) & "," & _
                         " banc_vl_ColExt = " & FunNuloVal(FunTrataFloat(MskColExtenso.Text)) & "," & _
                         " banc_vl_LinNom = " & FunNuloVal(FunTrataFloat(MskLinNominal.Text)) & "," & _
                         " banc_vl_ColNom = " & FunNuloVal(FunTrataFloat(MskColNominal.Text)) & "," & _
                         " banc_vl_LinData = " & FunNuloVal(FunTrataFloat(MskLinData.Text)) & "," & _
                         " banc_vl_ColData = " & FunNuloVal(FunTrataFloat(MskColData.Text)) & "" & _
                     " WHERE banc_cd_codigo = " & Chave
'        Conexao.IsolationLevel = adXactIsolated
        Conexao.Execute Sql
    End If
    Conexao.CommitTrans
    
    PanTipo.Visible = False
    PanTipo.Enabled = False
    PanBotoes.Enabled = True
    subHabilitaBotoes
    
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
'    subConectarControleDados Adodc1, FRM_FiltroAtual, "estatico"
    
    'Adodc1.Refresh TestesubrecarregadadosNV
    
    If XInserir = "A" Then
        cboLocalizar.BoundText = Chave
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    
    Exit Sub
    
MostraErro:

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
'
'   Next
'
'   Resume Next
   MsgBox "Este código do banco já foi atribuído a outro banco cadastrado.", vbCritical + vbOKOnly, "ATENÇÃO"
   Exit Sub
    
End Sub

Function funChecarCampos() As Boolean
    funChecarCampos = True
    
    If FunObrigatorioMSK(MskCodigo, "Código do Banco é Obrigatório") Then Exit Function
    
    If FunObrigatorioTXT(TxtDescricao, "Descrição é Obrigatória") Then Exit Function
    
    funChecarCampos = False
End Function

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

Private Sub CmdFormasPagto_Click()
    TelaFormasPagamento.MousePointer = vbHourglass
    TelaFormasPagamento.Show
    TelaFormasPagamento.MousePointer = vbDefault
End Sub

Private Sub cmdImprimir_Click()
    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Bancos")
End Sub
Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Imprime os dados da tabela"
End Sub


Private Sub CmdRemoverFiltro_Click()
    Dim Fil As Integer
    
    subDesabilitaBotoes
    
    Set Formulario = TelaBancos
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"

    If Filtrou = True Then
        Fil = ContFil
        subContaReg Adodc1
        'Adodc1.Refresh TestesubrecarregadadosNV
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        ContFil = Fil
    End If

    subHabilitaBotoes
    
    If FunTabelaVazia(Adodc1) Then
       CmdAlterar.Enabled = False
       CmdExcluir.Enabled = False
       CmdOrdem.Enabled = False
       CmdImprimir.Enabled = False
    End If
    
    TDBGrid1.Enabled = True
'    DataGrid1.SetFocus
    
End Sub

Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    MdiPrincipal.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub
Private Sub CmdOrdem_Click()
    
    subDesabilitaBotoes
    Set Formulario = TelaBancos
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

Private Sub CmdOrdem_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Ordena os registros atuais da tabela"
End Sub

Private Sub CmdExcluir_Click()

    If Not FunVerifica_Permissao("TelaBancos", "4") Then
        Exit Sub
    End If
    
    If CmdExcluir.Caption = "&Excluir" Then
        'If Not funVerifica_Permissao("FrmExemplo", OrdemRotina) Then Exit Sub
    End If

    Dim opcao As Integer
    Dim Houve_Erro As Integer
    Dim Posicao As Variant
    
    Houve_Erro = False
    
    Posicao = Adodc1.Recordset.bookmark
    subDesabilitaBotoes
    
    opcao = MsgBox("Confirma Remoção do Registro ?", 273, "CUIDADO")
    
    If opcao = 1 Then  'Confirma exclusão
       On Error GoTo RotuloErro
       Conexao.BeginTrans ' Adicionado - Lucas Santiago 22.01.09
       Conexao.Execute "DELETE FROM Bancos WHERE banc_cd_codigo=" & Adodc1.Recordset.Fields("banc_cd_codigo") ' Adicionado - Lucas Santiago 22.01.09
       'Adodc1.Recordset.Delete
       Conexao.CommitTrans
       If Houve_Erro = False Then
          subConectarControleDadosNV Adodc1, FRM_FiltroAtual, "Estatico"
          Adodc1.Refresh
          'TDBGrid1.Refresh
          'Quantid = Quantid - 1
          'MdiPrincipal.BarraStatus.Panels(3).Text = Str(Quantid) + "/" + Str(Quantid)
          MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
          If ContFil <> 0 Then
             subContaReg Adodc1
          End If
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
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    TDBGrid1.Col = 0

    'Posicionando o cursor na linha anterior - Lucas Santiago 22.01.09
    If Posicao > 1 Then
        Adodc1.Recordset.bookmark = Posicao - 1
    End If

    Exit Sub
    
RotuloErro:
    Conexao.RollbackTrans
    Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next

End Sub
Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If CmdInserir.Caption = "&Gravar" Then
       MdiPrincipal.BarraStatus.Panels(1).Text = "Cancela as novas informações"
    Else
       MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
    End If
End Sub

Private Sub CmdInserir_Click()
    If Not FunVerifica_Permissao("TelaBancos", "2") Then
        Exit Sub
    End If
    
    TelaBancos.MousePointer = vbHourglass
    XInserir = "I"
    Chave = -1
    PanBotoes.Enabled = False
    PanTipo.Visible = True
    PanTipo.Enabled = True
    PanTipo.Left = (TDBGrid1.Width - PanTipo.Width) / 2
    PanTipo.Top = (TDBGrid1.Height - PanTipo.Height) / 2
    PanTipo.Caption = " Inserção de Banco"
    'DataGrid1.Height = 2775
    subPrepararInsercao
    subDesabilitaBotoes
    cboLocalizar.Text = ""
     
    'Coordenadas Padrão do Cheque
    MskLinValor = "00,90"
    MskColValor = "13,00"
    MskLinExtenso = "01,70"
    MskColExtenso = "02,50"
    MskLinNominal = "02,90"
    MskColNominal = "00,50"
    MskColData = "07,80"
    MskLinData = "03,50"
  
    MskCodigo.Enabled = True
    MskCodigo.SetFocus
    
    TelaBancos.MousePointer = vbDefault
End Sub
Private Sub CmdInserir_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If CmdInserir.Caption = "&Gravar" Then
       MdiPrincipal.BarraStatus.Panels(1).Text = "Grava na tabela as novas informações"
    Else
       MdiPrincipal.BarraStatus.Panels(1).Text = "Insere um novo registro na tabela"
    End If
End Sub



Private Sub TDBGrid1_DblClick()
    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    cboLocalizar.Text = ""
    If Button = 2 Then   'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaBancos
        subTelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, X
        subTelaValoresGlobais "G"
    End If
    
    subHabilitaBotoes

End Sub

Private Sub Form_Activate()
    Set Formulario = TelaBancos
    If Not PrimeiraVez Then
        subTelaValoresGlobais "P"
    End If

    PrimeiraVez = False
    
    MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
End Sub

Private Sub Form_Load()
    
    Set Formulario = TelaBancos
            
    Call subAjustaTelaLoc
                
    subManutencaoJanelasAtivas "I", "TelaBancos"
            
    'Associa um banco de dados à tabela Descontos e Acréscimos
    subCarregaDadosNV Adodc1, "Bancos", "banc_cd_Codigo", "", "SELECT * FROM Bancos"
    
    filtra_especial = False
    
    subHabilitaBotoes
    
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
    subManutencaoJanelasAtivas "R", "TelaBancos"
End Sub


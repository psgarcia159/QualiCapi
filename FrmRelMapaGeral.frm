VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelMapaGeral 
   Caption         =   "Mapa Geral Por Empreendimento"
   ClientHeight    =   5280
   ClientLeft      =   945
   ClientTop       =   2190
   ClientWidth     =   9360
   Icon            =   "FrmRelMapaGeral.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5280
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraAgrupados 
      Caption         =   "Escolha a Observação referente ao Jurídico"
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
      Left            =   0
      TabIndex        =   13
      Top             =   4560
      Width           =   4995
      Begin MSDataListLib.DataCombo CboObs 
         Bindings        =   "FrmRelMapaGeral.frx":2AFA
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   240
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
         Left            =   2520
         Top             =   240
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
   End
   Begin VB.Frame Frame1 
      Height          =   4515
      Left            =   0
      TabIndex        =   2
      Top             =   -60
      Width           =   9315
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
         Height          =   1575
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   2175
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
            Left            =   180
            TabIndex        =   11
            Top             =   1080
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
            Left            =   180
            TabIndex        =   10
            Top             =   660
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
            TabIndex        =   9
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame FraMesAno 
         Caption         =   "Mês de Vencimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   2175
         Begin MSComCtl2.DTPicker DtpDtBase 
            Height          =   315
            Left            =   540
            TabIndex        =   4
            Top             =   300
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MM/yyyy"
            Format          =   53739523
            CurrentDate     =   37636
         End
      End
      Begin VB.Frame FraModelo 
         Caption         =   "Modelo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   120
         TabIndex        =   5
         Top             =   1260
         Width           =   2175
         Begin VB.OptionButton OptModelo1 
            Caption         =   "Modelo 1"
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
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptModelo2 
            Caption         =   "Modelo 2"
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
            Left            =   180
            TabIndex        =   6
            Top             =   720
            Width           =   1215
         End
      End
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Height          =   4170
         Left            =   2400
         TabIndex        =   12
         Top             =   180
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   7355
         _LayoutType     =   4
         _RowHeight      =   24
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   68
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Sel."
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Empresa"
         Columns(2).FooterText=   "D"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "dd/mm/yy"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   4
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=847"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=529"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(0)._MinWidth=78536464"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=20"
         Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(17)=   "Column(2).Width=1958"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=1879"
         Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=8720"
         Splits(0)._ColumnProps(22)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(24)=   "Column(2)._MinWidth=100555328"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=142,.parent=123,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=139,.parent=124,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=140,.parent=125"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=141,.parent=127"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=123"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=124"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=125"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=127"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=146,.parent=123,.alignment=0,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=143,.parent=124,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=144,.parent=125"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=145,.parent=127"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   345
      Left            =   5820
      TabIndex        =   0
      Top             =   4800
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   7500
      TabIndex        =   1
      Top             =   4800
      Width           =   1035
   End
End
Attribute VB_Name = "FrmRelMapaGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VFV_VETOR As New XArray

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    
  Dim XLT_FILTROOBSERVACAO As String
  Dim XLT_INDEXADOR As String
  Dim XLT_MODELO As String
  Dim XLT_SQL As String
  Dim XLI_CONT As Integer
  
  FrmRelMapaGeral.MousePointer = vbHourglass
  
  XGT_SELECAO = ""
  XGT_CONJUNCAO = ""
  
  XGT_SELECAO = ""
  XGT_CONJUNCAO = ""
  
  XLT_SQL = "" & _
  "SELECT ConsCAPRelMapaGeral.*," & _
  "CASE Indexacao1   when 'M' then" & _
    " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
    " when 'D' then" & _
    " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(Format(DtpDtBase.Value, "\01/mm/yy"), NomeSgbd) & ")" & _
  " end as Cotacao1," & _
  "CASE  Indexacao2 when 'M' then" & _
    " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
    " when 'D'  then" & _
    " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(Format(DtpDtBase.Value, "\01/mm/yy"), NomeSgbd) & ")" & _
  " end as Cotacao2,"

  XLT_SQL = XLT_SQL & _
  "CASE  Indexacao1 when 'M' then " & _
    "CASE cont_tx_CorrecProrata when 'A' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
    "when 'P' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
    "end " & _
  " when 'D' then Null " & _
  " end as Cotacao1_MesAnterior,"

  XLT_SQL = XLT_SQL & _
  "CASE Indexacao2 when 'M' then " & _
    "CASE cont_tx_CorrecProrata when 'A' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
    "when 'P' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
    "end " & _
  " when 'D' then Null " & _
  " end As Cotacao2_MesAnterior" & _
  " From ConsCAPRelMapaGeral WHERE "
  
  
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    "(ConsCAPRelMapaGeral.cont_dt_Venda <= " & FunNuloData(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(DtpDtBase.Value, "\01/mm/yy")))), NomeSgbd) & _
    " AND (ConsCAPRelMapaGeral.cont_tx_Status='A' OR " & _
    " (ConsCAPRelMapaGeral.cont_tx_Status='I' AND " & _
    " ConsCAPRelMapaGeral.cont_dt_Distrato is not null AND " & _
    " ConsCAPRelMapaGeral.cont_dt_Distrato >= " & FunNuloData(Format(DtpDtBase.Value, "\01/mm/yy"), NomeSgbd) & "" & _
    " )))"
  XGT_CONJUNCAO = " AND "
  
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    " ((ConsCAPRelMapaGeral.titu_dt_Pagamento is Null AND " & _
    " ConsCAPRelMapaGeral.tipl_tx_Sigla <> 'R')"
    
  XGT_SELECAO = XGT_SELECAO & _
    " OR(ConsCAPRelMapaGeral.titu_dt_Pagamento is Null AND " & _
    " ConsCAPRelMapaGeral.tipl_tx_Sigla = 'R' AND " & _
    " ConsCAPRelMapaGeral.titu_dt_Base < " & FunNuloData(Format(DtpDtBase.Value, "\01/mm/yy"), NomeSgbd) & ") "
    
  XGT_SELECAO = XGT_SELECAO & _
      " OR (ConsCAPRelMapaGeral.titu_dt_Pagamento is not Null AND " & _
      " ConsCAPRelMapaGeral.titu_dt_Pagamento >= " & FunNuloData(Format(DtpDtBase.Value, "\01/mm/yy"), NomeSgbd) & ")) "
  
  XGT_CONJUNCAO = " AND "
  
  'Indexador
  If OptValorHistorico.Value = True Then
    XLT_INDEXADOR = "0"
  ElseIf OptIndexador1.Value = True Then
    XLT_INDEXADOR = "1"
  ElseIf OptIndexador2.Value = True Then
    XLT_INDEXADOR = "2"
  End If
  
  'Modelo
  If OptModelo1.Value = True Then
    XLT_MODELO = "1"
  Else
    XLT_MODELO = "2"
  End If
  
  'Observacao
  If Not IsNull(CboObs.BoundText) Then
    XLT_FILTROOBSERVACAO = CboObs.BoundText
  End If
  
  'Filtra as empresas do Grid
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "(" ' "ConsCAPRelMapaGeral.empr_cd_Empresa = " & CStr(PCodEmpresa)
     
  XGT_CONJUNCAO = ""
  For XLI_CONT = 0 To VFV_VETOR.Count(1) - 1
    If VFV_VETOR(XLI_CONT, 0) = -1 Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & " empr_cd_Empresa =" & VFV_VETOR(XLI_CONT, 1)
      XGT_CONJUNCAO = " OR "
    End If
  Next
  XGT_SELECAO = XGT_SELECAO & ")"
 
  'Fecha o recordset se estiver aberto
  If XGO_RSRELATORIO.State = adStateOpen Then
    XGO_RSRELATORIO.Close
    Set XGO_RSRELATORIO = Nothing
  End If
  'Abre os recordsets e verifica se eles possuem dados
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL & XGT_SELECAO)
  
  If XGO_RSRELATORIO.EOF Then
      MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
      FrmRelMapaGeral.MousePointer = vbDefault
      Exit Sub
  End If
  
  'Fecha o recordset se estiver aberto
  If XGO_RSCABECALHO.State = adStateOpen Then
    XGO_RSCABECALHO.Close
    Set XGO_RSCABECALHO = Nothing
  End If
  Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
  
  'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelMapaGeral.rpt")
  Set XGR_RELATORIO = Nothing
  Set XGR_RELATORIO = RptMapaGeral
  XGR_RELATORIO.DiscardSavedData
  
  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO

  'Ativa os Relatórios
  With XGR_RELATORIO.FormulaFields
      .GetItemByName("FiltroDataBase").Text = Chr(34) & Format(DtpDtBase.Value, "\01/mm/yy") & Chr(34)
      .GetItemByName("Indexador").Text = Chr(34) & XLT_INDEXADOR & Chr(34)
      .GetItemByName("FiltroModelo").Text = Chr(34) & XLT_MODELO & Chr(34)
      .GetItemByName("FiltroObservacao").Text = Chr(34) & XLT_FILTROOBSERVACAO & Chr(34)
      .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
  End With
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  
   FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSRELATORIO, "SubRel1"
  
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSRELATORIO, "TotalIndexador"
    
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSRELATORIO, "TotalIndexador2"
  
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSRELATORIO, "TotalIndexador3"
  
  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
  FrmRelMapaGeral.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelMapaGeral
End Sub

Private Sub Form_Load()
  
  Dim XLI_CONT As Integer
  
  subManutencaoJanelasAtivas "I", "FrmRelMapaGeral"
  subConectarControleDadosNV DatObs, "SELECT * FROM Observacoes ORDER BY obse_tx_observacao", Estatico
  DtpDtBase.Value = Date
    
  subCarregaVetor "SELECT 0 as Selecao,empr_cd_Empresa,empr_tx_RazaoSocial FROM Empresas", Array("Selecao", "empr_cd_Empresa", "empr_tx_RazaoSocial"), VFV_VETOR, TDBGrid1
  
  For XLI_CONT = 0 To VFV_VETOR.Count(1) - 1
    If VFV_VETOR(XLI_CONT, 1) = PCodEmpresa Then
      VFV_VETOR(XLI_CONT, 0) = -1
    End If
  Next
  
  TDBGrid1.Array = VFV_VETOR
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If XGO_RSRELATORIO.State = adStateOpen Then
    XGO_RSRELATORIO.Close
  End If
  
  If XGO_RSCABECALHO.State = adStateOpen Then
    XGO_RSCABECALHO.Close
  End If
  
  Set XGO_RSRELATORIO = Nothing
  Set XGO_RSCABECALHO = Nothing
  
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  
  If TDBGrid1.Columns(1).Value = PCodEmpresa Then
    TDBGrid1.Columns(0).Value = -1
  End If

End Sub

Private Sub TDBGrid1_LostFocus()

  TDBGrid1.Update
  
End Sub

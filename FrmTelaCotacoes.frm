VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form TelaCotacoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabela de Cotações"
   ClientHeight    =   3915
   ClientLeft      =   4680
   ClientTop       =   3285
   ClientWidth     =   3990
   Icon            =   "FrmTelaCotacoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
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
      Left            =   3150
      TabIndex        =   5
      Top             =   3510
      Width           =   765
   End
   Begin Threed.SSPanel PanCotacao 
      Height          =   1590
      Left            =   90
      TabIndex        =   11
      Top             =   1410
      Visible         =   0   'False
      Width           =   3810
      _Version        =   65536
      _ExtentX        =   6720
      _ExtentY        =   2805
      _StockProps     =   15
      Caption         =   "Cotação"
      ForeColor       =   16711680
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
      Alignment       =   0
      Begin VB.TextBox TxtValor 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   2730
         TabIndex        =   8
         Top             =   795
         Width           =   960
      End
      Begin VB.CommandButton CmdRetorna 
         BackColor       =   &H00000000&
         Caption         =   "&Retorna"
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
         Left            =   2790
         TabIndex        =   10
         Top             =   1230
         Width           =   900
      End
      Begin VB.CommandButton CmdConfirma 
         BackColor       =   &H00000000&
         Caption         =   "&Confirma"
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
         Left            =   1815
         TabIndex        =   9
         Top             =   1230
         Width           =   900
      End
      Begin VB.TextBox TxtTaxa 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   780
         TabIndex        =   7
         Top             =   795
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DtpData 
         Height          =   315
         Left            =   780
         TabIndex        =   6
         Top             =   375
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   82575363
         CurrentDate     =   37455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor:"
         Height          =   240
         Left            =   2220
         TabIndex        =   16
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Taxa:"
         Height          =   240
         Left            =   270
         TabIndex        =   15
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data:"
         Height          =   285
         Left            =   315
         TabIndex        =   14
         Top             =   405
         Width           =   390
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "FrmTelaCotacoes.frx":030A
      Height          =   2760
      Left            =   60
      TabIndex        =   12
      Top             =   615
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   4868
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Data"
      Columns(0).DataField=   "cota_dt_datacotacao"
      Columns(0).NumberFormat=   "dd/mm/yy"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Taxa (%)"
      Columns(1).DataField=   "cota_vl_taxacotacao"
      Columns(1).NumberFormat=   "##0.0000"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Valor"
      Columns(2).DataField=   "cota_vl_valor"
      Columns(2).NumberFormat=   "#0.0000"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=17"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=56398032"
      Splits(0)._ColumnProps(10)=   "Column(1).Width=1482"
      Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=1402"
      Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=786"
      Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(16)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(17)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(18)=   "Column(2).Width=2196"
      Splits(0)._ColumnProps(19)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._WidthInPix=2117"
      Splits(0)._ColumnProps(21)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(2)._ColStyle=18"
      Splits(0)._ColumnProps(23)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(25)=   "Column(2)._MinWidth=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(9)   =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(10)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=1"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=1"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
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
      Left            =   2120
      TabIndex        =   4
      Top             =   3510
      Width           =   765
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
      Left            =   1090
      TabIndex        =   3
      Top             =   3510
      Width           =   765
   End
   Begin MSAdodcLib.Adodc DatMoedas 
      Height          =   330
      Left            =   1965
      Top             =   3375
      Visible         =   0   'False
      Width           =   2070
      _ExtentX        =   3651
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
      Caption         =   "DatMoedas"
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
   Begin Threed.SSPanel PanCotacoes 
      Height          =   540
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   952
      _StockProps     =   15
      Caption         =   "Moeda"
      ForeColor       =   -2147483635
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelInner      =   2
      Begin MSDataListLib.DataCombo CboMoedas 
         Bindings        =   "FrmTelaCotacoes.frx":031F
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Top             =   135
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "moed_tx_descricao"
         BoundColumn     =   "moed_cd_codmoeda"
         Text            =   ""
      End
      Begin VB.Label LblMoeda 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Indexador:"
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
         Left            =   195
         TabIndex        =   13
         Top             =   165
         Width           =   900
      End
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
      Left            =   60
      TabIndex        =   2
      Top             =   3510
      Width           =   765
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1920
      Top             =   3630
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
   Begin VB.Label LblData 
      BackColor       =   &H0080FFFF&
      Height          =   300
      Left            =   2010
      TabIndex        =   18
      Top             =   4305
      Width           =   870
   End
   Begin VB.Label LblChave 
      BackColor       =   &H0080FFFF&
      Height          =   300
      Left            =   330
      TabIndex        =   17
      Top             =   4320
      Width           =   870
   End
End
Attribute VB_Name = "TelaCotacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public XFT_ChaveMoeda As String
Public XFT_IndexMoeda As String
'Habilita os Botões da Tela
Private Sub subHabilitaBotoes()
    CmdInserir.Enabled = True
    
    If Not Adodc1.Recordset.EOF Then
        CmdAlterar.Enabled = True
        CmdExcluir.Enabled = True
    Else
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
    End If
    CmdSair.Enabled = True
End Sub
'Desabilita os Botões da Tela
Private Sub subDesabilitaBotoes()
    CmdInserir.Enabled = False
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    CmdSair.Enabled = False
End Sub
'Esta função irá verificar se a data que será inserida na tabela para uma determinada moeda já existe
Function FunExisteData(Data As String, Moeda As Variant) As Boolean
    Dim XLO_Data As ADODB.Recordset
    SubQOpenRecordset XLO_Data, "SELECT * FROM CotacoesMoedas WHERE moed_cd_codmoeda=" & _
    XFT_ChaveMoeda & " and cota_dt_datacotacao= " & FunNuloData(Data, NomeSgbd) & "", Estatico
    If XLO_Data.EOF And XLO_Data.BOF Then
        XLO_Data.Close
        Set XLO_Data = Nothing
        FunExisteData = False
    Else
        XLO_Data.Close
        Set XLO_Data = Nothing
        FunExisteData = True
    End If
End Function
Private Sub CboMoedas_Change()
    DatMoedas.Recordset.bookmark = CboMoedas.SelectedItem
    XFT_ChaveMoeda = CboMoedas.BoundText
    XFT_IndexMoeda = DatMoedas.Recordset.Fields("moed_tx_indexacao")
    subCarregaDadosNV Adodc1, "CotacoesMoedas", "cota_dt_datacotacao Desc", "", _
    "Select * from CotacoesMoedas where moed_cd_codmoeda = " & XFT_ChaveMoeda
    subHabilitaBotoes
End Sub
Private Sub CmdAlterar_Click()
    PanCotacao.Tag = "A"
    PanCotacao.Enabled = True
    PanCotacao.Visible = True
    TDBGrid1.Enabled = False
    DTPData.Value = Format$(Adodc1.Recordset.Fields("cota_dt_datacotacao"), "DD/MM/YY")
    DTPData.Enabled = False
    If Adodc1.Recordset.Fields("cota_vl_taxacotacao") <> 0 Then
        TxtTaxa.Text = Format$(Adodc1.Recordset.Fields("cota_vl_taxacotacao"), "##0.0000")
    Else
        TxtTaxa.Text = ""
    End If
    If Adodc1.Recordset.bookmark = 1 Then
        TxtTaxa.Enabled = False
    Else
        TxtTaxa.Enabled = True
    End If
    TxtValor.Text = Format$(Adodc1.Recordset.Fields("cota_vl_valor"), "#0.0000")
    
    subDesabilitaBotoes
End Sub
Private Sub CmdConfirma_Click()
    Dim XLI_Houve_Erro As Integer  'Indica se Houve Erro
    Dim XLT_SQL As String
    Dim XLI_INDICE As Integer 'Guarda o índice do registro selecionado no grid - Larissa 04/06/2008
   
    XLI_Houve_Erro = False
    
    If TxtValor.Text = "" Then
        MsgBox "O Preenchimento do valor é obrigatório", vbCritical, "Atenção!"
        Exit Sub
        TxtValor.SetFocus
    End If
    
    If PanCotacao.Tag = "I" Then
        'Verifica se já foi realizada uma cotação da moeda para uma determinada data
        If FunExisteData(Format(DTPData.Value, "dd/mm/yy"), XFT_ChaveMoeda) Then
            MsgBox "Já existe cotação para esta data.", vbInformation, "ATENÇÃO!"
            DTPData.SetFocus
            Exit Sub
        Else
            If Not TDBGrid1.EOF And Not TDBGrid1.BOF Then
                If LblData.Caption <> "" Then
                    If (XFT_IndexMoeda = "M") And _
                    (Day(DTPData.Value) <> Day(TDBGrid1.Columns("Data"))) Then
                        'Só permitir cotações para o mesmo dia em cada mês
                        MsgBox "O período de indexação deve ser mensal.", vbInformation, "ATENÇÃO!"
                        DTPData.SetFocus
                        Exit Sub
                    End If
                Else
                    If (DatMoedas.Recordset.Fields("moed_tx_indexacao") = "M") And _
                       (Day(DTPData.Value) <> Day(TDBGrid1.Columns("Data"))) Then
                        'Só permitir cotações para o mesmo dia em cada mês
                        MsgBox "O período de indexação deve ser mensal.", vbInformation, "ATENÇÃO!"
                        DTPData.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        XLT_SQL = "INSERT INTO CotacoesMoedas(cota_dt_datacotacao,moed_cd_codmoeda,cota_vl_taxacotacao,cota_vl_valor) " & _
              "VALUES (" & FunNuloData(DTPData.Value, NomeSgbd) & ", " & XFT_ChaveMoeda & "," & FunNulo(FunTrataFloat(TxtTaxa.Text)) & "," & (FunTrataFloat(TxtValor.Text)) & ")"
        Conexao.Execute XLT_SQL
        
    Else
        
        XLT_SQL = "UPDATE CotacoesMoedas set " _
              & "cota_vl_valor = " & (FunTrataFloat(TxtValor.Text)) & ", " _
              & "cota_vl_taxacotacao = " & FunNulo(FunTrataFloat(TxtTaxa.Text)) & "" _
              & " WHERE moed_cd_codmoeda = " & XFT_ChaveMoeda & " and cota_dt_datacotacao=" & FunNuloData(DTPData.Value, NomeSgbd) & ""
        Conexao.Execute XLT_SQL
        
    End If
    
    'Pega o índice do registro selecionado pelo usuário
    If Not IsNull(TDBGrid1.bookmark) Then
        XLI_INDICE = TDBGrid1.bookmark
            
        Adodc1.Refresh
    
        TDBGrid1.Enabled = True
        
        'Mantém a seleção do registro feita pelo usuário - Larissa 04/06/2008
        TDBGrid1.bookmark = XLI_INDICE
        TDBGrid1.SetFocus
    Else
        Adodc1.Refresh
    End If
    
    PanCotacao.Enabled = False
    PanCotacao.Visible = False
    subHabilitaBotoes
    
    If LblData.Caption <> "" Then
        Call CmdSair_Click
    End If
    
    Exit Sub

RotuloErro:

    subHabilitaBotoes
    XLI_Houve_Erro = True
    Select Case funTrataErros(ComMensagem)

      Case 3000
         MsgBox Campos_Obrig + " são obrigatórios", 48, "ERRO"
          Resume Next

      Case ErrChaveDuplicada
            Resume Next

      Case ErrRegistroExcluido
            Adodc1.Refresh
            Resume Next

      Case 3200
         MsgBox "Código não pode ser alterado pois existem outras tabelas vinculadas a ele.", 64, "ERRO"
          Resume Next

    End Select
End Sub
Private Sub CmdRetorna_Click()
    PanCotacao.Enabled = False
    PanCotacao.Visible = False
    subHabilitaBotoes
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    If LblData.Caption <> "" Then
        Call CmdSair_Click
    End If
End Sub
Private Sub CmdSair_Click()
    Unload Me
End Sub
Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub
Private Sub CmdExcluir_Click()
    
    On Error GoTo TrataErro
    Dim Sql As String
    Dim Posicao As Variant

    Posicao = Adodc1.Recordset.bookmark
    
    If MsgBox("Confirma Remoção do Registro ?", vbCritical + vbYesNo, "CUIDADO") = vbNo Then Exit Sub
    TelaCotacoes.MousePointer = vbHourglass
    
    Conexao.BeginTrans
    Sql = "DELETE FROM CotacoesMoedas WHERE moed_cd_codmoeda = " & Adodc1.Recordset.Fields("moed_cd_codmoeda") & " and cota_dt_datacotacao=" & FunNuloData(Adodc1.Recordset.Fields("cota_dt_datacotacao"), NomeSgbd) & ""
    Conexao.Execute Sql
    Conexao.CommitTrans
    Adodc1.Refresh
    
    Quantid = Quantid - 1
    MdiPrincipal.BarraStatus.Panels(3).Text = Str(Quantid) + "/" + Str(Quantid)
    If ContFil <> 0 Then
       subContaReg Adodc1
    End If

    subHabilitaBotoes
    
ResumoDoErro:
    TDBGrid1.SetFocus
    TelaCotacoes.MousePointer = vbDefault
    Adodc1.Refresh
    
    'Posicionando o cursor na linha anterior - Lucas Santiago 22.01.09
    If Posicao > 1 Then
        Adodc1.Recordset.bookmark = Posicao - 1
    End If
    
    Exit Sub

TrataErro:
    Conexao.RollbackTrans
    If Err.Number = NumberTravaRegistro And Err.Description = DescriptionTravaRegistro Then
        If MsgBox("O serviço está sendo alterado por outro usuário. Deseja tentar novamente?", vbCritical + vbQuestion + vbYesNo, "ATENÇÃO") = vbYes Then
            Resume
        Else
            Resume ResumoDoErro
        End If
    Else
        funTrataErros (ComMensagem)
        Resume ResumoDoErro
    End If
    
    'Posicionando o cursor na linha anterior - Lucas Santiago 22.01.09
    If Posicao > 1 Then
        Adodc1.Recordset.bookmark = Posicao - 1
    End If
    
End Sub
Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If CmdInserir.Caption = "&Gravar" Then
       MdiPrincipal.BarraStatus.Panels(1).Text = "Cancela as novas informações"
    Else
       MdiPrincipal.BarraStatus.Panels(1).Text = "Remove da tabela o registro selecionado"
    End If
End Sub
Private Sub CmdInserir_Click()
    
    PanCotacao.Tag = "I"
    PanCotacao.Enabled = True
    PanCotacao.Visible = True
    TDBGrid1.Enabled = False
    DTPData.Enabled = True
    If LblData.Caption <> "" Then
        DatMoedas.Recordset.bookmark = CboMoedas.SelectedItem
        If DatMoedas.Recordset.Fields("moed_tx_indexacao") = "D" Then
            DTPData.Value = Format$(Now, "DD/MM/YY")
        Else
            TDBGrid1.MoveFirst
            If TDBGrid1.Columns("Data") <> "" Then
                DTPData.Value = Format(DateAdd("m", 1, TDBGrid1.Columns("Data")), "DD/MM/YY")
            Else
                DTPData.Value = Format$(Now, "DD/MM/YY")
            End If
        End If
    Else
        If XFT_IndexMoeda = "D" Then
            DTPData.Value = Format$(Now, "DD/MM/YY")
        Else
            TDBGrid1.MoveFirst
            If TDBGrid1.Columns("Data") <> "" Then
                DTPData.Value = Format(DateAdd("m", 1, TDBGrid1.Columns("Data")), "DD/MM/YY")
            Else
                DTPData.Value = Format$(Now, "DD/MM/YY")
            End If
        End If
    End If
    TxtTaxa.Text = ""
    TxtValor.Text = ""
    CmdInserir.SetFocus
    subDesabilitaBotoes
            
    If FunTabelaVazia(Adodc1) Then
        TxtTaxa.Enabled = False
    Else
        TxtTaxa.Enabled = True
    End If
End Sub
Private Sub CmdInserir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If CmdInserir.Caption = "&Gravar" Then
       MdiPrincipal.BarraStatus.Panels(1).Text = "Grava na tabela as novas informações"
    Else
       MdiPrincipal.BarraStatus.Panels(1).Text = "Insere um novo registro na tabela"
    End If
End Sub

Private Sub Form_Activate()
    
    Dim XLO_MOEDA As New ADODB.Recordset
    
    If LblData.Caption <> "" Then
        XFT_ChaveMoeda = LblChave.Caption
        subCarregaDadosNV Adodc1, "ConsGENCotacaoMoeda", "cota_dt_datacotacao Desc", "", _
        "Select * from ConsGENCotacaoMoeda where moed_cd_codmoeda = " & XFT_ChaveMoeda
        
        subHabilitaBotoes
        
        LblMoeda.Visible = False
        CboMoedas.Visible = False
        CboMoedas.Enabled = False
        
        If Adodc1.Recordset.EOF Then
          'Quando não tem nenhuma cotação de uma moeda é preciso pegar os dados da moeda
          SubQOpenRecordset XLO_MOEDA, "SELECT * FROM Moedas WHERE moed_cd_codmoeda = " & XFT_ChaveMoeda, Estatico
          PanCotacoes.Caption = XLO_MOEDA!moed_tx_descricao
          XFT_IndexMoeda = XLO_MOEDA!moed_tx_Indexacao
          XLO_MOEDA.Close
          Set XLO_MOEDA = Nothing
        Else
          PanCotacoes.Caption = Adodc1.Recordset.Fields("moed_tx_descricao")
          XFT_IndexMoeda = Adodc1.Recordset.Fields("moed_tx_indexacao")
        End If
        
        PanCotacao.Tag = "I"
        PanCotacao.Enabled = True
        PanCotacao.Visible = True
        TDBGrid1.Enabled = False
        DTPData.Value = LblData.Caption
        DTPData.Enabled = False
        TxtTaxa.Text = ""
        TxtValor.Text = ""
        TxtTaxa.SetFocus
        subDesabilitaBotoes
        If Not Adodc1.Recordset.EOF Then
            TxtTaxa.Enabled = False
        Else
            TxtTaxa.Enabled = True
        End If
    Else
        If LblChave.Caption <> "" Then
            CboMoedas.BoundText = LblChave.Caption
        End If
    End If
End Sub

Private Sub Form_Load()
    If LblData.Caption = "" Then 'Se a tela for chamada a partir do Menu Principal
        'se a tela for chamada apartir da MDI
        subConectarControleDadosNV DatMoedas, "SELECT * FROM Moedas", Estatico
        subCarregaDadosNV Adodc1, "CotacoesMoedas", "cota_dt_datacotacao Desc", "", "Select * from CotacoesMoedas where moed_cd_codmoeda = 0"
        subDesabilitaBotoes
        CmdSair.Enabled = True
        
        LblMoeda.Visible = True
        CboMoedas.Visible = True
        CboMoedas.Enabled = True
        PanCotacoes.Caption = ""
    End If
End Sub
Private Sub TxtTaxa_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 45 Then
        KeyAscii = 13
    End If
End Sub
Private Sub TxtTaxa_LostFocus()
    If TxtTaxa <> "" And Not TDBGrid1.EOF And Not TDBGrid1.BOF Then
        If PanCotacao.Tag = "I" Then
            TDBGrid1.MoveFirst
            TxtValor = (TDBGrid1.Columns("Valor") + ((TDBGrid1.Columns("Valor") * TxtTaxa.Text) / 100))
        Else
            If Adodc1.Recordset.bookmark > 1 Then
            Adodc1.Recordset.bookmark = (Adodc1.Recordset.bookmark - 1)
            TxtValor = (Adodc1.Recordset.Fields("cota_vl_Valor") + ((Adodc1.Recordset.Fields("cota_vl_Valor") * TxtTaxa.Text) / 100))
            End If
        End If
    End If
End Sub



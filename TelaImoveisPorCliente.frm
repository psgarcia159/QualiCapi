VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaManImoveisPorCliente 
   Caption         =   "Tabela de Imóveis por Cliente"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   9390
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   9390
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
      Left            =   8115
      TabIndex        =   4
      Top             =   4560
      Width           =   1230
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1980
      Top             =   2790
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
   Begin Threed.SSPanel PanCodigo 
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9255
      _Version        =   65536
      _ExtentX        =   16325
      _ExtentY        =   979
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
      Begin VB.Label LblNomeCliente 
         Caption         =   "Label1"
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
         Left            =   1245
         TabIndex        =   2
         Top             =   180
         Width           =   7275
      End
      Begin VB.Label LblCliente 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   165
         Width           =   900
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaImoveisPorCliente.frx":0000
      Height          =   3810
      Left            =   60
      TabIndex        =   3
      Top             =   675
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   6720
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Contrato"
      Columns(0).DataField=   "Contrato"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Empreendimento"
      Columns(1).DataField=   "empd_tx_Nome"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tipo"
      Columns(2).DataField=   "tiim_tx_Descricao"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Data da Venda"
      Columns(3).DataField=   "cont_dt_Venda"
      Columns(3).NumberFormat=   "dd/mm/yy"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2593"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2514"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8720"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=7064"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6985"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=8720"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=3307"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3228"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=8720"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1879"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1799"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=8721"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bold=-1,.fontsize=825"
      _StyleDefs(21)  =   ":id=14,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(22)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3,.alignment=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.alignment=3"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.valignment=2"
      _StyleDefs(33)  =   ":id=28,.locked=-1"
      _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(49)  =   "Named:id=33:Normal"
      _StyleDefs(50)  =   ":id=33,.parent=0"
      _StyleDefs(51)  =   "Named:id=34:Heading"
      _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   ":id=34,.wraptext=-1"
      _StyleDefs(54)  =   "Named:id=35:Footing"
      _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   "Named:id=36:Selected"
      _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(58)  =   "Named:id=37:Caption"
      _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(60)  =   "Named:id=38:HighlightRow"
      _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(62)  =   "Named:id=39:EvenRow"
      _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(64)  =   "Named:id=40:OddRow"
      _StyleDefs(65)  =   ":id=40,.parent=33"
      _StyleDefs(66)  =   "Named:id=41:RecordSelector"
      _StyleDefs(67)  =   ":id=41,.parent=34"
      _StyleDefs(68)  =   "Named:id=42:FilterBar"
      _StyleDefs(69)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "TelaManImoveisPorCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
  Set Formulario = TelaManImoveisPorCliente
  If Not Adodc1.Recordset.EOF Then
      LblNomeCliente.Caption = UCase(Adodc1.Recordset.Fields("focl_tx_RazaoSocial"))
    Else
      TelaManImoveisPorCliente.Hide
      MsgBox "Esse cliente não possui nenhum imóvel!", vbCritical, "ATENÇÃO"
      Unload Me
    End If
End Sub

Private Sub Form_Load()

    Set Formulario = TelaManImoveisPorCliente
    
    subCarregaDadosNV Adodc1, "", "Contrato", _
                            "", "SELECT * FROM ConsCapImoveisPorCliente WHERE empr_cd_empresa = " & PCodEmpresa & " AND focl_cd_forncli=" & Chave
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set TelaManImoveisPorCliente = Nothing
End Sub

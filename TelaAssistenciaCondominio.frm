VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaAssistenciaCondominio 
   Caption         =   "Assistência Técnica"
   ClientHeight    =   11070
   ClientLeft      =   540
   ClientTop       =   1845
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11070
   ScaleWidth      =   13410
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaAssistenciaCondominio.frx":0000
      Height          =   3270
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   5768
      _LayoutType     =   4
      _RowHeight      =   24
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Empreendimento"
      Columns(0).DataField=   "empd_cd_Empreendimento"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Cliente"
      Columns(1).DataField=   "focl_tx_RazaoSocial"
      Columns(1).NumberFormat=   "dd/mm/yy"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Número"
      Columns(2).DataField=   "assi_cd_CodigoSeq"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Autorizador"
      Columns(3).DataField=   "ascd_tx_Autorizador"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Data"
      Columns(4).DataField=   "ascd_dt_Data"
      Columns(4).NumberFormat=   "dd/mm/yy"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Hora"
      Columns(5).DataField=   "ascd_dt_DataEmissao"
      Columns(5).NumberFormat=   "Short Time"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Status"
      Columns(6).DataField=   "assi_tx_status"
      Columns(6).NumberFormat=   "dd/mm/yy"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2646"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2566"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8721"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=9710"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=9631"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=529"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1402"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1323"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=529"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=9710"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=9631"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=529"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=1879"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1799"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=529"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=1799"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1720"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=529"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1402"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1323"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=529"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bold=-1,.fontsize=825"
      _StyleDefs(27)  =   ":id=14,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(28)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3,.alignment=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.alignment=3"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2,.valignment=2"
      _StyleDefs(39)  =   ":id=28,.locked=-1"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=74,.parent=13,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
      _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
      _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
      _StyleDefs(67)  =   "Named:id=33:Normal"
      _StyleDefs(68)  =   ":id=33,.parent=0"
      _StyleDefs(69)  =   "Named:id=34:Heading"
      _StyleDefs(70)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   ":id=34,.wraptext=-1"
      _StyleDefs(72)  =   "Named:id=35:Footing"
      _StyleDefs(73)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   "Named:id=36:Selected"
      _StyleDefs(75)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(76)  =   "Named:id=37:Caption"
      _StyleDefs(77)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(78)  =   "Named:id=38:HighlightRow"
      _StyleDefs(79)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(80)  =   "Named:id=39:EvenRow"
      _StyleDefs(81)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(82)  =   "Named:id=40:OddRow"
      _StyleDefs(83)  =   ":id=40,.parent=33"
      _StyleDefs(84)  =   "Named:id=41:RecordSelector"
      _StyleDefs(85)  =   ":id=41,.parent=34"
      _StyleDefs(86)  =   "Named:id=42:FilterBar"
      _StyleDefs(87)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5400
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
   Begin Threed.SSPanel PanBotoes 
      Height          =   780
      Left            =   -240
      TabIndex        =   1
      Top             =   3360
      Width           =   9300
      _Version        =   65536
      _ExtentX        =   16404
      _ExtentY        =   1376
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
      Begin VB.CommandButton CmdFormulario 
         Caption         =   "Imprimir"
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
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CmdInserir 
         Caption         =   "Inserir"
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
         TabIndex        =   5
         Top             =   360
         Width           =   1095
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
         Left            =   4200
         TabIndex        =   4
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
         Left            =   5400
         TabIndex        =   3
         Top             =   360
         Width           =   1095
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
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cboLocalizar 
         Bindings        =   "TelaAssistenciaCondominio.frx":0015
         Height          =   315
         Left            =   3060
         TabIndex        =   6
         Top             =   0
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "assi_cd_assistencia"
         BoundColumn     =   "assi_cd_assistencia"
         Text            =   ""
      End
      Begin VB.Label lblLocalizar 
         Alignment       =   1  'Right Justify
         Caption         =   "Localizar Código:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   60
         Width           =   2670
      End
   End
End
Attribute VB_Name = "TelaAssistenciaCondominio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XFT_CODIMOVEL As Variant 'Código do imóvel
Dim XFT_CODEMPREENDIMENTO As Variant 'Código do empreendimento
Dim XFT_CODCONTRATO As Variant 'Código do Contrato
Dim XLT_ASSISTENCIA As Variant
Dim XFT_ASSISTENCIA As Variant
Dim FRM_Ordem As String ' Nome do Campo que tem a chave da Tabela ativa
Dim FRM_FiltroAtual As String ' Indica o filtro atual da tabela corrente

Private Sub CmdAlterar_Click()
'    If Not FunVerifica_Permissao(Me.Name, "1") Then
 '      Exit Sub
 '   End If
    
    XInserir = "A"
    xChave = Adodc1.Recordset.bookmark
    subDesabilitaBotoes
    
    
    Chave = Adodc1.Recordset("empd_cd_empreendimento") & "." & _
        Format(Adodc1.Recordset("ascd_cd_assistencia"), "#")
    
    FrmSolAssistenciaCondominio.Show 1
    
      subCarregaDadosNV Adodc1, "", "assi_cd_CodigoSeq", _
                            "", "SELECT * FROM ConsCapAssistenciaTecnica_Condominio WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' "
   
    

    subHabilitaBotoes
End Sub

Private Sub CmdExcluir_Click()
'If Not FunVerifica_Permissao(Me.Name, "4") Then
 '       Exit Sub
 '   End If
    
    Dim XLI_OPCAO As Integer
    Dim XLI_BOOKMARK As Integer
    
    
    On Error GoTo RotuloErro
    xChave = Adodc1.Recordset.bookmark

    
    subDesabilitaBotoes
    
    Chave = Adodc1.Recordset("empd_cd_empreendimento") & "." & _
        Format(Adodc1.Recordset("ascd_cd_assistencia"), "#")
        
        XLT_ASSISTENCIA = Mid$(Chave, 6, 3)
    
    XLI_OPCAO = MsgBox("Confirma Remoção do Registro ?", 273, "CUIDADO")
    
    If XLI_OPCAO = 1 Then  'Confirma exclusão
    
      XGT_SQL = "DELETE FROM AssistenciaTecnicaCondomonio WHERE empr_cd_Empresa='" & PCodEmpresa & "' AND empd_cd_Empreendimento= '" & XFT_CODEMPREENDIMENTO & "' AND ascd_cd_Assistencia= '" & XLT_ASSISTENCIA & "' "
      Conexao.Execute XGT_SQL
      MsgBox "Operação Efetuada com Sucesso", vbInformation, "Aviso"
      XLI_BOOKMARK = TDBGrid1.bookmark
      FRM_Quantid = funOperacaoQuantid("-")
      subCarregaDadosNV Adodc1, "", "assi_cd_CodigoSeq", _
                            "", "SELECT * FROM AssistenciaTecnicaCondomonio WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento=" & XFT_CODEMPREENDIMENTO & ""
      
      If Not Adodc1.Recordset.EOF Then
        If XLI_BOOKMARK <= Adodc1.Recordset.RecordCount Then
          TDBGrid1.bookmark = XLI_BOOKMARK
        Else
          TDBGrid1.bookmark = XLI_BOOKMARK - 1
        End If
      End If
      
      subHabilitaBotoes
      
      
      If Adodc1.Recordset.EOF Then
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdFormulario.Enabled = False
    
    End If
    
             
    Else
       subHabilitaBotoes
    End If
       
    cboLocalizar.Text = ""
    TDBGrid1.SetFocus
    TDBGrid1.Col = 0
    
    subCarregaDadosNV Adodc1, "", "assi_cd_CodigoSeq", _
                            "", "SELECT * FROM AssistenciaTecnicaCondomonio WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel= '" & XFT_CODIMOVEL & "' AND cont_cd_contrato = '" & XFT_CODCONTRATO & "'  "
   
    
    
    Exit Sub
    
RotuloErro:
    
    If Err.Number = -2147467259 Or Err.Number = -2147217873 Then
      MsgBox "Esta solicitação não pode ser excluída!", vbCritical, "ATENÇÃO"
    End If
    subHabilitaBotoes

    

End Sub

Private Sub CmdFormulario_Click()
    
    Me.MousePointer = Default
 
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
   
    Chave = Adodc1.Recordset("empd_cd_empreendimento") & "." & _
        Format(Adodc1.Recordset("ascd_cd_assistencia"), "#")
        
    XGT_SELECAO = "empr_cd_empresa = " & CStr(PCodEmpresa) & " AND assi_tx_Titulo = '" & Chave & "'"
                     
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsCapAssistenciaTecnica_Condominio_Relatorio WHERE " & XGT_SELECAO)
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & CStr(PCodEmpresa))
  
    'Carrega os relatórios
    FrmVisRelatorios.Caption = "Solicitação de Assistência Técnica"
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelFormularioAssistencia.rpt")
  
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelComissoes.MousePointer = vbDefault
    
End Sub

Private Sub CmdInserir_Click()

    If Not FunVerifica_Permissao("TelaFinTitulos", "1") Then
      Exit Sub
  End If
  
  XOrigem = ""
  XInserir = "I"
  
  'Chave = XFT_CODEMPREENDIMENTO & "." & XFT_CODIMOVEL & "." & XFT_CODCONTRATO
  Chave = XFT_CODEMPREENDIMENTO
  
    
    FrmSolAssistenciaCondominio.Show 1
    subCarregaDadosNV Adodc1, "", "assi_cd_CodigoSeq", _
                            "", "SELECT * FROM ConsCapAssistenciaTecnica_Condominio WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "'  "
   
    
  
  If Adodc1.Recordset.EOF Then
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdFormulario.Enabled = False
    Else
    
    CmdAlterar.Enabled = True
    CmdExcluir.Enabled = True
    CmdFormulario.Enabled = True
    
    
    
    End If
  If FunLocalizaJanelasAtivas("FrmSolAssistenciaCondominio") <> -1 Then
      FrmSolAssistencia.SetFocus
  End If
End Sub


Private Sub CmdSair_Click()

    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    Unload Me
End Sub

Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub

Private Sub subHabilitaBotoes()

    'Objetivo: Habilita os Botões da Tela
    CmdInserir.Enabled = True
    CmdAlterar.Enabled = True
    CmdExcluir.Enabled = True
    CmdFormulario.Enabled = True


End Sub

Private Sub subDesabilitaBotoes()
    'Objetivo: Desabilita os Botões da Tela
    CmdInserir.Enabled = False
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    CmdFormulario.Enabled = False

End Sub

Private Sub Form_Load()

    
    Set Formulario = TelaAssistenciaCondominio
    XFT_CODEMPREENDIMENTO = Left$(Chave, 4)
    XFT_CODIMOVEL = Mid$(Chave, 6, 4)
    XFT_CODCONTRATO = Mid$(Chave, 11, 2)
    XFT_ASSISTENCIA = Mid$(Chave, 13, 2)
 
    subCarregaDadosNV Adodc1, "", "assi_cd_CodigoSeq", _
                            "", "SELECT * FROM ConsCapAssistenciaTecnica_Condominio WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento='" & XFT_CODEMPREENDIMENTO & "' "
   
    
    If Adodc1.Recordset.EOF Then
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdFormulario.Enabled = False
    
    End If
    
    Call SubAjustaTela
    
End Sub

Private Sub TDBGrid1_Click()
    If BarraFerramentasPressionada = True Then
      BarraFerramentasPressionada = False
    End If

End Sub


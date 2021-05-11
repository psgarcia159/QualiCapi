VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmConConfigDoc 
   Caption         =   "Configurções de Notas Promissorias"
   ClientHeight    =   6255
   ClientLeft      =   3285
   ClientTop       =   2790
   ClientWidth     =   8670
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8670
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   582
      ConnectMode     =   3
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
      Align           =   1  'Align Top
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   10081
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   64
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Dado"
      Columns(0).DataField=   "Form_tx_descricao"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   84
      Columns(1)._MaxComboItems=   5
      Columns(1).ValueItems(0)._DefaultItem=   -1
      Columns(1).ValueItems(0).Value=   "0"
      Columns(1).ValueItems(0).Value.vt=   8
      Columns(1).ValueItems(0).DisplayValue=   "1"
      Columns(1).ValueItems(0).DisplayValue.vt=   8
      Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems(1)._DefaultItem=   0
      Columns(1).ValueItems(1).Value=   "1"
      Columns(1).ValueItems(1).Value.vt=   8
      Columns(1).ValueItems(1).DisplayValue=   "0"
      Columns(1).ValueItems(1).DisplayValue.vt=   8
      Columns(1).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems.Count=   2
      Columns(1).Caption=   "Imprime"
      Columns(1).DataField=   "Form_nr_Imprime"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Altura"
      Columns(2).DataField=   "Form_nr_Altura"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Largura"
      Columns(3).DataField=   "Form_nr_largura"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   1
      Columns(4)._MaxComboItems=   5
      Columns(4).ValueItems(0)._DefaultItem=   0
      Columns(4).ValueItems(0).Value=   "8"
      Columns(4).ValueItems(0).Value.vt=   8
      Columns(4).ValueItems(0).DisplayValue=   "8"
      Columns(4).ValueItems(0).DisplayValue.vt=   8
      Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems(1)._DefaultItem=   0
      Columns(4).ValueItems(1).Value=   "9"
      Columns(4).ValueItems(1).Value.vt=   8
      Columns(4).ValueItems(1).DisplayValue=   "9"
      Columns(4).ValueItems(1).DisplayValue.vt=   8
      Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems(2)._DefaultItem=   0
      Columns(4).ValueItems(2).Value=   "10"
      Columns(4).ValueItems(2).Value.vt=   8
      Columns(4).ValueItems(2).DisplayValue=   "10"
      Columns(4).ValueItems(2).DisplayValue.vt=   8
      Columns(4).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems(3)._DefaultItem=   0
      Columns(4).ValueItems(3).Value=   "11"
      Columns(4).ValueItems(3).Value.vt=   8
      Columns(4).ValueItems(3).DisplayValue=   "11"
      Columns(4).ValueItems(3).DisplayValue.vt=   8
      Columns(4).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems.Count=   4
      Columns(4).Caption=   "Tam. Fonte"
      Columns(4).DataField=   "Form_nr_tamfonte"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   84
      Columns(5)._MaxComboItems=   5
      Columns(5).ValueItems(0)._DefaultItem=   0
      Columns(5).ValueItems(0).Value=   "0"
      Columns(5).ValueItems(0).Value.vt=   8
      Columns(5).ValueItems(0).DisplayValue=   "1"
      Columns(5).ValueItems(0).DisplayValue.vt=   8
      Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(1)._DefaultItem=   0
      Columns(5).ValueItems(1).Value=   "1"
      Columns(5).ValueItems(1).Value.vt=   8
      Columns(5).ValueItems(1).DisplayValue=   "0"
      Columns(5).ValueItems(1).DisplayValue.vt=   8
      Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems.Count=   2
      Columns(5).Caption=   "Negrito"
      Columns(5).DataField=   "Form_nr_Negrito"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4524"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4445"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8704"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0).DropDownList=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1402"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1323"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1482"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1402"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=1535"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1455"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2170"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2090"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(26)=   "Column(4).Button=1"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=1429"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=1349"
      Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(54)  =   "Named:id=33:Normal"
      _StyleDefs(55)  =   ":id=33,.parent=0"
      _StyleDefs(56)  =   "Named:id=34:Heading"
      _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(58)  =   ":id=34,.wraptext=-1"
      _StyleDefs(59)  =   "Named:id=35:Footing"
      _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(61)  =   "Named:id=36:Selected"
      _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=37:Caption"
      _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(65)  =   "Named:id=38:HighlightRow"
      _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(67)  =   "Named:id=39:EvenRow"
      _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(69)  =   "Named:id=40:OddRow"
      _StyleDefs(70)  =   ":id=40,.parent=33"
      _StyleDefs(71)  =   "Named:id=41:RecordSelector"
      _StyleDefs(72)  =   ":id=41,.parent=34"
      _StyleDefs(73)  =   "Named:id=42:FilterBar"
      _StyleDefs(74)  =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Align           =   4  'Align Right
      Height          =   540
      Left            =   30
      TabIndex        =   1
      Top             =   5715
      Width           =   8640
      _Version        =   65536
      _ExtentX        =   15240
      _ExtentY        =   952
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
      Begin VB.CommandButton cmdExit 
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
         Left            =   7140
         TabIndex        =   2
         Top             =   120
         Width           =   1125
      End
   End
End
Attribute VB_Name = "FrmConConfigDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XFO_RsGrid As New ADODB.Recordset   'RS para atualizacao do Banco com o grid

Private Sub Botao_Des_Click()
'    MdiPrincipal.BarraStatus.Panels(1).Text = ""
'    MdiPrincipal.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub CmdSair_Click()

End Sub

Private Sub Form_Load()
     Set Formulario = FrmConConfigDoc
    Dim Sql As String
    
    'Call subAjustaTela
        
    subManutencaoJanelasAtivas "I", "FrmConConfigDoc"
    If (XGT_NomeAplic = "QualiFats") Then
      Sql = "SELECT * FROM Formularios WHERE Empr_cd_empresa = " & PCodEmpresa & " AND  form_tx_tipodocumento = 'NFS' "
    Else
      Sql = "SELECT * FROM Formularios WHERE Empr_cd_empresa = " & PCodEmpresa & " AND  form_tx_tipodocumento = 'PR'  "
      Sql = Sql + " ORDER BY form_tx_Descricao"
    End If
    
    XFO_RsGrid.CursorLocation = adUseClient
    XFO_RsGrid.Open funTrataSql(Sql), Conexao, adOpenStatic, adLockOptimistic, adCmdText
    Set TDBGrid1.DataSource = XFO_RsGrid


    'filtra_especial = True
    'subTelaValoresGlobais "G"
    'PrimeiraVez = True


End Sub

Private Sub Form_Unload(Cancel As Integer)
    
  ' XFO_RsGrid.Close
  Set XFO_RsGrid = Nothing
  
End Sub


Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  
  TDBGrid1.Update
  TDBGrid1.Refresh
  
End Sub

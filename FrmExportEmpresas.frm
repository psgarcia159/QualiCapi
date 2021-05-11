VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmExportEmpresas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportação Entre Empresas"
   ClientHeight    =   5475
   ClientLeft      =   30
   ClientTop       =   435
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleMode       =   0  'User
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
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
      Left            =   3600
      TabIndex        =   5
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "Confirmar"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo CboEmpreendimentos 
      Bindings        =   "FrmExportEmpresas.frx":0000
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "empd_tx_Nome"
      BoundColumn     =   "empd_tx_Nome"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DatEmpreendimentos 
      Height          =   330
      Left            =   1800
      Top             =   4200
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
   Begin MSDataListLib.DataCombo CboMoeda 
      Bindings        =   "FrmExportEmpresas.frx":0021
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "moed_tx_descricao"
      BoundColumn     =   "moed_tx_descricao"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DatMoedas 
      Height          =   330
      Left            =   360
      Top             =   5160
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
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
      Caption         =   "DatMoeda"
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
   Begin MSComDlg.CommonDialog DlgSalvar 
      Left            =   1440
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TrueDBGrid70.TDBGrid TDBGrid1 
      Height          =   3960
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   6985
      _LayoutType     =   4
      _RowHeight      =   15
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Imóveis"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   68
      Columns(1)._MaxComboItems=   5
      Columns(1).ValueItems(0)._DefaultItem=   0
      Columns(1).ValueItems(0).Value=   ""
      Columns(1).ValueItems(0).Value.vt=   8
      Columns(1).ValueItems(0).DisplayValue=   ""
      Columns(1).ValueItems(0).DisplayValue.vt=   8
      Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems.Count=   1
      Columns(1).Caption=   "Exporta"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=139793"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(0)._MinWidth=32"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2990"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2910"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=131601"
      Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(1)._MinWidth=70364288"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=128,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(25)  =   ":id=43,.strikethrough=0,.charset=0"
      _StyleDefs(26)  =   ":id=43,.fontname=MS Sans Serif"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=2,.valignment=2"
      _StyleDefs(39)  =   ":id=28,.locked=-1"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44,.alignment=2,.bold=-1"
      _StyleDefs(41)  =   ":id=25,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(42)  =   ":id=25,.fontname=MS Sans Serif"
      _StyleDefs(43)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
      _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=2,.valignment=2"
      _StyleDefs(46)  =   ":id=32,.wraptext=-1,.locked=0"
      _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2,.bold=-1"
      _StyleDefs(48)  =   ":id=29,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(49)  =   ":id=29,.fontname=MS Sans Serif"
      _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
      _StyleDefs(52)  =   "Named:id=33:Normal"
      _StyleDefs(53)  =   ":id=33,.parent=0"
      _StyleDefs(54)  =   "Named:id=34:Heading"
      _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   ":id=34,.wraptext=-1"
      _StyleDefs(57)  =   "Named:id=35:Footing"
      _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   "Named:id=36:Selected"
      _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=37:Caption"
      _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(63)  =   "Named:id=38:HighlightRow"
      _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(65)  =   "Named:id=39:EvenRow"
      _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(67)  =   "Named:id=40:OddRow"
      _StyleDefs(68)  =   ":id=40,.parent=33"
      _StyleDefs(69)  =   "Named:id=41:RecordSelector"
      _StyleDefs(70)  =   ":id=41,.parent=34"
      _StyleDefs(71)  =   "Named:id=42:FilterBar"
      _StyleDefs(72)  =   ":id=42,.parent=33"
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
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label LblMoedas 
      Alignment       =   1  'Right Justify
      Caption         =   "Moeda:"
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
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1485
   End
End
Attribute VB_Name = "FrmExportEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim XLT_CODEMPREEND As String
Dim XLT_MOEDA As String
Dim Vetor1 As New XArray
Dim XLO_EXPORTA As Object
Dim XLT_SQL As String
Dim XLT_TEXTO As String
Dim XLT_CODIMOVEL As String
Dim XLT_FORNCLIE As String
Dim XLT_CONTRATO As String
Function FunExportaMoeda() As Boolean
    
    'Exportando a moeda
    If Not IsNull(CboMoeda.SelectedItem) Then
       DatMoedas.Recordset.Bookmark = CboMoeda.SelectedItem
       XLT_MOEDA = DatMoedas.Recordset.Fields!moed_cd_CodMoeda
    Else
        FunExportaMoeda = False
        MsgBox "Selecione a moeda!", vbInformation + vbOKOnly, "ATENÇÃO"
        Exit Function
    End If
    
    DlgSalvar.DialogTitle = "Exportação de Títulos entre Empresas - EXPORTANDO A MOEDA"
    DlgSalvar.DefaultExt = "txt"
    DlgSalvar.InitDir = XGT_LOCALARQ
    DlgSalvar.CancelError = False
    DlgSalvar.ShowSave
    
    If DlgSalvar.FileName = "" Then
        FunExportaMoeda = False
        Exit Function
    End If
    
    If DlgSalvar.FileName <> "" Then
        XGT_LOCALARQ = DlgSalvar.FileName
        XGT_ARQUIVO = DlgSalvar.FileTitle
    End If
    
    Open XGT_LOCALARQ For Output As #1

    XLT_SQL = "SELECT * FROM Moedas WHERE " & _
        " moed_cd_codmoeda ='" & XLT_MOEDA & "'"
    SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
    
    If Not XLO_EXPORTA.EOF Then
        
        'Código - Pegar na importação
        
        'Símbolo(5)
        XLT_TEXTO = "M"
        If Not IsNull(XLO_EXPORTA!moed_tx_simbolo) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!moed_tx_simbolo) + Space(5 - Len(XLO_EXPORTA!moed_tx_simbolo))
            Else
              XLT_TEXTO = XLT_TEXTO + Space(5)
        End If
        'Descrição(30)
        If Not IsNull(XLO_EXPORTA!moed_tx_descricao) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!moed_tx_descricao) + Space(30 - Len(XLO_EXPORTA!moed_tx_descricao))
            Else
              XLT_TEXTO = XLT_TEXTO + Space(30)
        End If
        'indexação(1)
        If Not IsNull(XLO_EXPORTA!moed_tx_Indexacao) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!moed_tx_Indexacao)
            Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
        End If
        'Padrão(1)
        If Not IsNull(XLO_EXPORTA!moed_tx_padrao) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!moed_tx_padrao)
            Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
        End If
        'CodVelho(5)
        If Not IsNull(XLO_EXPORTA!codvelho) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!codvelho) + Space(5 - Len(XLO_EXPORTA!codvelho))
            Else
              XLT_TEXTO = XLT_TEXTO + Space(5)
        End If
        Print #1, XLT_TEXTO
    End If
    XLO_EXPORTA.Close
    
    XLT_SQL = "SELECT * FROM CotacoesMoedas WHERE " & _
        " moed_cd_codmoeda ='" & XLT_MOEDA & "'"
    SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
   
   Do While Not XLO_EXPORTA.EOF
        
        XLT_TEXTO = ""
        XLT_TEXTO = "C"
        'Data de Cotação(10)
        If Not IsNull(XLO_EXPORTA!cota_dt_DataCotacao) Then
              XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!cota_dt_DataCotacao), "DD/MM/YYYY")
            Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
        End If
        
        'Código da Moeda(5) - Pegar na importação
        
        'Taxa da Cotação(15)
        If Not IsNull(XLO_EXPORTA!cota_vl_taxacotacao) Then
            If XLO_EXPORTA!cota_vl_taxacotacao < 0 Then
                XLT_TEXTO = XLT_TEXTO + FunTrataFloat(XLO_EXPORTA!cota_vl_taxacotacao) + Space(16 - Len(XLO_EXPORTA!cota_vl_taxacotacao))
            ElseIf XLO_EXPORTA!cota_vl_taxacotacao < 1 And XLO_EXPORTA!cota_vl_taxacotacao <> 0 Then
                XLT_TEXTO = XLT_TEXTO + FunTrataFloat(XLO_EXPORTA!cota_vl_taxacotacao) + Space(15 - Len(XLO_EXPORTA!cota_vl_taxacotacao))
            Else
                XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cota_vl_taxacotacao)) + Space(15 - Len(XLO_EXPORTA!cota_vl_taxacotacao))
            End If
        Else
            XLT_TEXTO = XLT_TEXTO + "0" + Space(14)
        End If
        
        'Valor da Cotação(15)
        If Not IsNull(XLO_EXPORTA!cota_vl_Valor) Then
            XLT_TEXTO = XLT_TEXTO + FunTrataFloat(XLO_EXPORTA!cota_vl_Valor) + Space(15 - Len(XLO_EXPORTA!cota_vl_Valor))
        Else
              XLT_TEXTO = Space(15)
        End If
        
        Print #1, XLT_TEXTO
        
        XLO_EXPORTA.MoveNext
    Loop
    Close #1
    XLO_EXPORTA.Close
    FunExportaMoeda = True
End Function

Private Sub CboEmpreendimentos_Change()
   If Not IsNull(CboEmpreendimentos.SelectedItem) Then
      
      DatEmpreendimentos.Recordset.Bookmark = CboEmpreendimentos.SelectedItem
      
      XLT_CODEMPREEND = DatEmpreendimentos.Recordset.Fields!empd_cd_Empreendimento
            Sql = "Select distinct imov_cd_imovel, '0' AS MARCA from titulos where empd_cd_empreendimento = " & XLT_CODEMPREEND & _
            " AND (titu_dt_Pagamento is null or titu_dt_Pagamento = '') " & _
            " AND empr_cd_empresa = " & PCodEmpresa & ""
      subCarregaVetor Sql, Array("imov_cd_imovel", "Marca"), Vetor1, TDBGrid1
   
      TDBGrid1.Refresh
      
      TDBGrid1.MoveFirst

   End If
   
End Sub

Private Sub CboMoeda_Click(Area As Integer)
    'Sql = Sql
End Sub

Private Sub CmdConfirmar_Click()

Dim XLO_CODIGOS As Object
Dim Sql As String

TDBGrid1.MoveFirst
If FunExportaMoeda() Then
   DlgSalvar.DialogTitle = "Exportação de Títulos entre Empresas - EXPORTANDO TÍTULOS"
   DlgSalvar.DefaultExt = "txt"
   DlgSalvar.InitDir = XGT_LOCALARQ
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
   
   XLT_TEXTO = ""
   
   While Not TDBGrid1.EOF
      
      If TDBGrid1.Columns("Exporta").Value = "-1" Then
         
         XLT_SQL = "SELECT * FROM Imoveis WHERE " & _
             " imov_cd_imovel ='" & TDBGrid1.Columns("Imóveis").Value & "'" & _
             " AND empd_cd_empreendimento ='" & XLT_CODEMPREEND & "'" & _
             " AND empr_cd_Empresa = " & PCodEmpresa & ""
         SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
          
         'Letra que Identifica o bloco
         XLT_TEXTO = "I"
         
         'Código do Imóvel(4)
         XLT_CODIMOVEL = XLO_EXPORTA!imov_cd_Imovel
         XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!imov_cd_Imovel)
          
         'Código do Empreendimento(4) - Pega na importação
        
         'Código da Empresa(4) - Pega na importação
          
         'Prof do titular(20)
         Sql = "SELECT tiim_tx_Descricao FROM TiposImoveis where tiim_cd_TipoImovel = " & XLO_EXPORTA!tiim_cd_TipoImovel & ""
         SubQOpenRecordset XLO_CODIGOS, Sql, Estatico
         
         If Not XLO_CODIGOS.EOF Then

            XLT_TEXTO = XLT_TEXTO + Trim(XLO_CODIGOS!tiim_tx_Descricao) + Space(20 - Len(XLO_CODIGOS!tiim_tx_Descricao))
         
         End If
         
         XLO_CODIGOS.Close
         
         'Número da Àrea(8)
         If Not IsNull(XLO_EXPORTA!imov_nr_Area) Or (XLO_EXPORTA!imov_nr_Area) <> "" Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!imov_nr_Area) + Space(8 - Len(XLO_EXPORTA!imov_nr_Area))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(8)
         End If
         
         'Garagem(50)
         If Not IsNull(XLO_EXPORTA!imov_tx_Garagem) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!imov_tx_Garagem) + Space(50 - Len(XLO_EXPORTA!imov_tx_Garagem))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(50)
         End If
          
         'Status(1)
         If Not IsNull(XLO_EXPORTA!imov_tx_Status) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!imov_tx_Status)
         Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
         End If
         
         'Valor à vista(10)
         If Not IsNull(XLO_EXPORTA!imov_vl_Avista) Then
              XLT_TEXTO = XLT_TEXTO + FunTrataFloat(Trim(XLO_EXPORTA!imov_vl_Avista)) + Space(10 - Len(XLO_EXPORTA!imov_vl_Avista))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Andar(2)
         If Not IsNull(XLO_EXPORTA!imov_nr_Andar) Or (XLO_EXPORTA!imov_nr_Andar) <> "" Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!imov_nr_Andar) + Space(2 - Len(XLO_EXPORTA!imov_nr_Andar))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(2)
         End If
         
         'Data de Entrga das chaves(10)
         If Not IsNull(XLO_EXPORTA!imov_dt_EntregaChaves) Then
              XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!imov_dt_EntregaChaves), "DD/MM/YYYY")
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Último Contrato(2)
         If Not IsNull(XLO_EXPORTA!imov_cd_UltimoContrato) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!imov_cd_UltimoContrato) + Space(2 - Len(XLO_EXPORTA!imov_cd_UltimoContrato))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(2)
         End If
         Print #1, XLT_TEXTO
         
         XLO_EXPORTA.Close
         
         'Selecionando o cliente na tabela de contratos
         XLT_SQL = "SELECT * FROM Contratos WHERE " & _
             " imov_cd_imovel ='" & XLT_CODIMOVEL & "'" & _
             " AND empd_cd_empreendimento ='" & XLT_CODEMPREEND & "'" & _
             " AND cont_tx_Status= 'A' " & _
             " AND empr_cd_Empresa = " & PCodEmpresa & ""
         SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
         
         If Not XLO_EXPORTA.EOF Then
            XLT_FORNCLIE = XLO_EXPORTA!focl_cd_FornCli
         End If
         
         XLO_EXPORTA.Close
         
         '*******Inserindo dados do Forncliente***********
         XLT_SQL = "SELECT * FROM FornClientes WHERE " & _
             " focl_cd_FornCli ='" & XLT_FORNCLIE & "'"
         SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
         
         'Código do Cliente(4) - Gerar código na importação
         XLT_TEXTO = ""
         
         'Letra que Identifica o bloco
         XLT_TEXTO = "F"
         
         'Classe(1)
         If Not IsNull(XLO_EXPORTA!focl_tx_classe) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_classe)
         Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
         End If
         
         'Tipo(1)
         If Not IsNull(XLO_EXPORTA!focl_tx_Tipo) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_Tipo)
         Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
         End If
         
         'Fantasia(20)
         If Not IsNull(XLO_EXPORTA!focl_tx_fantasia) Then
              XLT_TEXTO = XLT_TEXTO + XLO_EXPORTA!focl_tx_fantasia + Space(20 - Len(XLO_EXPORTA!focl_tx_fantasia))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
         
         'RazãoSocial(50)
         If Not IsNull(XLO_EXPORTA!focl_tx_RazaoSocial) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_RazaoSocial) + Space(50 - Len(XLO_EXPORTA!focl_tx_RazaoSocial))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(50)
         End If
         
         'Endereço(50)
         If Not IsNull(XLO_EXPORTA!focl_tx_endereco) Then
              XLT_TEXTO = XLT_TEXTO + funTiraAcento(Trim(XLO_EXPORTA!focl_tx_endereco), True) + Space(50 - Len(XLO_EXPORTA!focl_tx_endereco))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(50)
         End If
         
         'Bairro(20)
         If Not IsNull(XLO_EXPORTA!focl_tx_Bairro) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_Bairro) + Space(20 - Len(XLO_EXPORTA!focl_tx_Bairro))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
         
         'Cidade(20)
         If Not IsNull(XLO_EXPORTA!focl_tx_cidade) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_cidade) + Space(20 - Len(XLO_EXPORTA!focl_tx_cidade))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
         
         'E-mail(30)
         If Not IsNull(XLO_EXPORTA!focl_tx_email) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_email) + Space(30 - Len(XLO_EXPORTA!focl_tx_email))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(30)
         End If
         
          'Estado(2)
         If Not IsNull(XLO_EXPORTA!focl_tx_estado) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_estado) + Space(2 - Len(XLO_EXPORTA!focl_tx_estado))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(2)
         End If
         
         
         If Not IsNull(XLO_EXPORTA!focl_tx_cep) Then
            If (XLO_EXPORTA!focl_tx_cep) <> "" Then
                XLT_TEXTO = XLT_TEXTO + (XLO_EXPORTA!focl_tx_cep)
            Else
                XLT_TEXTO = XLT_TEXTO + Space(10)
            End If
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
           
         'Fone(15)
         If Not IsNull(XLO_EXPORTA!focl_tx_fone) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_fone) + Space(15 - Len(XLO_EXPORTA!focl_tx_fone))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(15)
         End If
         
          'Fax(15)
         If Not IsNull(XLO_EXPORTA!focl_tx_fax) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_fax) + Space(15 - Len(XLO_EXPORTA!focl_tx_fax))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(15)
         End If
         
         'CGC/CPF(14)
         If Not IsNull(XLO_EXPORTA!focl_tx_CGCCPF) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_CGCCPF) + Space(14 - Len(XLO_EXPORTA!focl_tx_CGCCPF))
         Else
             XLT_TEXTO = XLT_TEXTO + Space(14)
         End If
         
         'Insc Estadual(15)
         If Not IsNull(XLO_EXPORTA!focl_tx_iestadual) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_iestadual) + Space(15 - Len(XLO_EXPORTA!focl_tx_iestadual))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(15)
         End If
         
         'Contato Comercial(20)
         If Not IsNull(XLO_EXPORTA!focl_tx_contatoCom) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_contatoCom) + Space(20 - Len(XLO_EXPORTA!focl_tx_contatoCom))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
         
         'Contato Financeiro(20)
         If Not IsNull(XLO_EXPORTA!focl_tx_ContatoFin) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_ContatoFin) + Space(20 - Len(XLO_EXPORTA!focl_tx_ContatoFin))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
         
         'Tipo Pag(2)
         If Not IsNull(XLO_EXPORTA!focl_tx_TipoPag) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_TipoPag) + Space(2 - Len(XLO_EXPORTA!focl_tx_TipoPag))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(2)
         End If
         
          'Icms(8)
         If Not IsNull(XLO_EXPORTA!focl_vl_Icms) Then
              XLT_TEXTO = XLT_TEXTO + FunTrataFloat(Trim(XLO_EXPORTA!focl_vl_Icms)) + Space(8 - Len(XLO_EXPORTA!focl_vl_Icms))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(8)
         End If
         
          'Fone1(15)
         If Not IsNull(XLO_EXPORTA!focl_tx_fone1) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_fone1) + Space(15 - Len(XLO_EXPORTA!focl_tx_fone1))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(15)
         End If
         
          'Celular(15)
         If Not IsNull(XLO_EXPORTA!focl_tx_Celular) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_Celular) + Space(15 - Len(XLO_EXPORTA!focl_tx_Celular))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(15)
         End If
         
          'Nº Porta(5)
         If (Not IsNull(XLO_EXPORTA!focl_nr_porta)) Or (XLO_EXPORTA!focl_nr_porta) <> "" Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_nr_porta) + Space(5 - Len(XLO_EXPORTA!focl_nr_porta))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(5)
         End If
         
          'Cod Atividade(5)
         If Not IsNull(XLO_EXPORTA!focl_nr_codatividade) Or (XLO_EXPORTA!focl_nr_codatividade) <> "" Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_nr_codatividade) + Space(5 - Len(XLO_EXPORTA!focl_nr_codatividade))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(5)
         End If
         
          'Cod Dos(5)
         If Not IsNull(XLO_EXPORTA!focl_cd_CodDos) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_cd_CodDos) + Space(5 - Len(XLO_EXPORTA!focl_cd_CodDos))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(5)
         End If
         
          'Cod Velho1(20)
         If Not IsNull(XLO_EXPORTA!CodVelho1) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!CodVelho1) + Space(20 - Len(XLO_EXPORTA!CodVelho1))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(20)
         End If

          'Cod Velho2(20)
         If Not IsNull(XLO_EXPORTA!CodVelho2) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!CodVelho2) + Space(20 - Len(XLO_EXPORTA!CodVelho2))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
         
          'Insc Municipal(20)
         If Not IsNull(XLO_EXPORTA!focl_tx_imunicipal) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!focl_tx_imunicipal) + Space(20 - Len(XLO_EXPORTA!focl_tx_imunicipal))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
     
         Print #1, XLT_TEXTO
         XLO_EXPORTA.Close
         
         
         '*******Inserindo dados do cliente***********
         XLT_SQL = "SELECT * FROM Clientes WHERE " & _
             " focl_cd_FornCli ='" & XLT_FORNCLIE & "'"
         SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
         
         'Código do Cliente(4) - Gerar código na importação
         XLT_TEXTO = ""
         
         'Letra que Identifica o bloco
         XLT_TEXTO = "C"
         
         'Cod da Profissão do Titular(20)
         If Not IsNull(XLO_EXPORTA!prof_cd_ProfTitular) Then
            Sql = "SELECT prof_tx_descricao FROM Profissoes where prof_cd_profissao = " & XLO_EXPORTA!prof_cd_ProfTitular & ""
            SubQOpenRecordset XLO_CODIGOS, Sql, Estatico
            
            If Not XLO_CODIGOS.EOF Then
               XLT_TEXTO = XLT_TEXTO + Trim(XLO_CODIGOS!prof_tx_Descricao) + Space(20 - Len(XLO_CODIGOS!prof_tx_Descricao))
            Else
                XLT_TEXTO = XLT_TEXTO + Space(20)
            End If
            XLO_CODIGOS.Close
         Else
            XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
         
         'Data de Nascimento(10)
         If Not IsNull(XLO_EXPORTA!clie_dt_Nascimento) Then
              XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!clie_dt_Nascimento), "DD/MM/YYYY")
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Nacionalidade(10)
         If Not IsNull(XLO_EXPORTA!clie_tx_Nacionalidade) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_tx_Nacionalidade) + Space(10 - Len(XLO_EXPORTA!clie_tx_Nacionalidade))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Identidade(15)
         If Not IsNull(XLO_EXPORTA!clie_nr_Identidade) Or (XLO_EXPORTA!clie_nr_Identidade) <> "" Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_nr_Identidade) + Space(15 - Len(XLO_EXPORTA!clie_nr_Identidade))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(15)
         End If
         
         'Órgão Emissor(8)
         If Not IsNull(XLO_EXPORTA!clie_tx_OrgEmissor) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_tx_OrgEmissor) + Space(8 - Len(XLO_EXPORTA!clie_tx_OrgEmissor))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(8)
         End If
         
         'Estado Civil(1)
         If Not IsNull(XLO_EXPORTA!clie_nr_EstCivil) Or (XLO_EXPORTA!clie_nr_EstCivil) <> "" Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_nr_EstCivil) + Space(1 - Len((XLO_EXPORTA!clie_nr_EstCivil)))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
         End If
         
         'End Residencial(50)
         If Not IsNull(XLO_EXPORTA!clie_tx_EndResidencial) Then
              XLT_TEXTO = XLT_TEXTO + funTiraAcento(Trim(XLO_EXPORTA!clie_tx_EndResidencial), True) + Space(50 - Len(XLO_EXPORTA!clie_tx_EndResidencial))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(50)
         End If
         
         'Bairro Residencial(20)
         If Not IsNull(XLO_EXPORTA!clie_tx_BairroResidencial) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_tx_BairroResidencial) + Space(20 - Len(XLO_EXPORTA!clie_tx_BairroResidencial))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
         
         'Cep Residencial(10)
         If Not IsNull(XLO_EXPORTA!clie_nr_CepResidencial) Then
            If (XLO_EXPORTA!clie_nr_CepResidencial) <> "" Then
               XLT_TEXTO = XLT_TEXTO + (XLO_EXPORTA!clie_nr_CepResidencial)
            Else
               XLT_TEXTO = XLT_TEXTO + Space(10)
            End If
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Município Residencial(20)
         If Not IsNull(XLO_EXPORTA!clie_tx_MunResidencial) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_tx_MunResidencial) + Space(20 - Len(XLO_EXPORTA!clie_tx_MunResidencial))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
         
         'estresidencial(2)
         If Not IsNull(XLO_EXPORTA!clie_tx_EstResidencial) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_tx_EstResidencial) + Space(2 - Len(XLO_EXPORTA!clie_tx_EstResidencial))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(2)
         End If
         
         'Conjuge(50)
         If Not IsNull(XLO_EXPORTA!clie_tx_Conjuge) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_tx_Conjuge) + Space(50 - Len(XLO_EXPORTA!clie_tx_Conjuge))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(50)
         End If
         
         'Profissão Conjuge(20)
         If Not IsNull(XLO_EXPORTA!prof_cd_ProfConjuge) Then
            
            Sql = "SELECT prof_tx_descricao FROM Profissoes where prof_cd_profissao = " & XLO_EXPORTA!prof_cd_ProfConjuge & ""
            SubQOpenRecordset XLO_CODIGOS, Sql, Estatico
            
            If Not XLO_CODIGOS.EOF Then
               XLT_TEXTO = XLT_TEXTO + Trim(XLO_CODIGOS!prof_tx_Descricao) + Space(20 - Len(XLO_CODIGOS!prof_tx_Descricao))
            Else
               XLT_TEXTO = XLT_TEXTO + Space(20)
            End If
           
            XLO_CODIGOS.Close
         
         Else
            XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
         
         'Data de Nascimento Cônjugue(10)
         If Not IsNull(XLO_EXPORTA!clie_dt_NascConjuge) Then
              XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!clie_dt_NascConjuge), "DD/MM/YYYY")
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Nacionalidade Conjuge(10)
         If Not IsNull(XLO_EXPORTA!clie_tx_NacConjuge) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_tx_NacConjuge) + Space(10 - Len(XLO_EXPORTA!clie_tx_NacConjuge))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Identidade Conjuge(15)
         If Not IsNull(XLO_EXPORTA!clie_nr_IdConjuge) Or (XLO_EXPORTA!clie_nr_IdConjuge) <> "" Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_nr_IdConjuge) + Space(15 - Len(XLO_EXPORTA!clie_nr_IdConjuge))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(15)
         End If
         
         'Orgão Emissor Cônjuge(8)
         If Not IsNull(XLO_EXPORTA!clie_tx_OrgEmissorConjuge) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_tx_OrgEmissorConjuge) + Space(8 - Len(XLO_EXPORTA!clie_tx_OrgEmissorConjuge))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(8)
         End If
         
         'CPF Cônjuge(14)
         If Not IsNull(XLO_EXPORTA!clie_nr_CPFConjuge) Then
              If (XLO_EXPORTA!clie_nr_CPFConjuge) <> "" Then
                 XLT_TEXTO = XLT_TEXTO + (XLO_EXPORTA!clie_nr_CPFConjuge)
              Else
                 XLT_TEXTO = XLT_TEXTO + Space(14)
              End If
         Else
              XLT_TEXTO = XLT_TEXTO + Space(14)
         End If
         
         'Renda familiar(10)
         If Not IsNull(XLO_EXPORTA!clie_vl_RendaFamiliar) Then
              XLT_TEXTO = XLT_TEXTO + FunTrataFloat(Trim(XLO_EXPORTA!clie_vl_RendaFamiliar)) + Space(10 - Len(XLO_EXPORTA!clie_vl_RendaFamiliar))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'End Correspondência(50)
         If Not IsNull(XLO_EXPORTA!clie_tx_endcorresp) Then
              XLT_TEXTO = XLT_TEXTO + funTiraAcento(Trim(XLO_EXPORTA!clie_tx_endcorresp), True) + Space(50 - Len(XLO_EXPORTA!clie_tx_endcorresp))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(50)
         End If
         
         'Bairro Correspondência(20)
         If Not IsNull(XLO_EXPORTA!clie_tx_BairroCorresp) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_tx_BairroCorresp) + Space(20 - Len(XLO_EXPORTA!clie_tx_BairroCorresp))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
         
         'CEP Correspondência(10)
         If Not IsNull(XLO_EXPORTA!clie_nr_cepcorresp) Or (XLO_EXPORTA!clie_nr_cepcorresp) <> "" Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_nr_cepcorresp) + Space(10 - Len(XLO_EXPORTA!clie_nr_cepcorresp))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Município Correspondência(20)
         If Not IsNull(XLO_EXPORTA!clie_tx_muncorresp) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_tx_muncorresp) + Space(20 - Len(XLO_EXPORTA!clie_tx_muncorresp))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(20)
         End If
          
         'Estado Correspondência(2)
         If Not IsNull(XLO_EXPORTA!clie_tx_estcorresp) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!clie_tx_estcorresp) + Space(2 - Len(XLO_EXPORTA!clie_tx_estcorresp))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(2)
         End If
         
         Print #1, XLT_TEXTO
         
         XLO_EXPORTA.Close
         
         '*******Exportando o Contrato***********
         XLT_SQL = "SELECT * FROM Contratos WHERE " & _
             " imov_cd_imovel ='" & XLT_CODIMOVEL & "'" & _
             " AND empd_cd_empreendimento ='" & XLT_CODEMPREEND & "'" & _
             " AND cont_tx_Status= 'A' " & _
             " AND empr_cd_Empresa = " & PCodEmpresa & ""
         SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
         
         XLT_TEXTO = ""
        
         'Letra que Identifica o bloco
         XLT_TEXTO = "N"
         
         'Código do Contrato(2)
         If Not IsNull(XLO_EXPORTA!cont_cd_Contrato) Then
              XLT_CONTRATO = XLO_EXPORTA!cont_cd_Contrato
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!cont_cd_Contrato) + Space(2 - Len(XLO_EXPORTA!cont_cd_Contrato))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(2)
         End If
         
    
         'Código do Imóvel - Mesmo Código para importação
         If Not IsNull(XLO_EXPORTA!imov_cd_Imovel) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!imov_cd_Imovel) + Space(4 - Len(XLO_EXPORTA!imov_cd_Imovel))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(2)
         End If
         
         'Código do Empreendimento(4) - Pegar na importação
         
         'Código da Empresa - Pegar na Importação
         
         'CR - Pegar o CR da empresa ativa
         
         'Conta - Pegar a conta da empresa ativa
         
         'Código da Moeda _ Pegar na importação
         
         'Valor à vista(10)
         If Not IsNull(XLO_EXPORTA!cont_vl_Avista) Then
              XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_Avista)) + Space(10 - Len(XLO_EXPORTA!cont_vl_Avista))
         Else
              XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
         End If
         
         'Valor do Sinal(10)
         If Not IsNull(XLO_EXPORTA!cont_vl_Sinal) Then
              XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_Sinal)) + Space(10 - Len(XLO_EXPORTA!cont_vl_Sinal))
         Else
              XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
         End If
         
         'Valor da Poupança(10)
         If Not IsNull(XLO_EXPORTA!cont_vl_Poupanca) Then
             XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_Poupanca)) + Space(10 - Len(XLO_EXPORTA!cont_vl_Poupanca))
         Else
             XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
         End If
         
         'Valor Total(10)
         If Not IsNull(XLO_EXPORTA!cont_vl_Total) Then
               XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_Total)) + Space(10 - Len(XLO_EXPORTA!cont_vl_Total))
         Else
            XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
         End If
         
         'Valor Fgts(10)
         If Not IsNull(XLO_EXPORTA!cont_vl_Fgts) Then
              XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_Fgts)) + Space(10 - Len(XLO_EXPORTA!cont_vl_Fgts))
         Else
              XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
         End If
         
         'Valor Financiamento(10)
         If Not IsNull(XLO_EXPORTA!cont_vl_Financiamento) Then
                XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_Financiamento)) + Space(10 - Len(XLO_EXPORTA!cont_vl_Financiamento))
         Else
              XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
         End If
         
         'Valor Outros(10)
         If Not IsNull(XLO_EXPORTA!cont_vl_Outros) Then
               XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_Outros)) + Space(10 - Len(XLO_EXPORTA!cont_vl_Outros))
         Else
              XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
         End If
         
         'Data Base(10)
         If Not IsNull(XLO_EXPORTA!cont_dt_Base) Then
              XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!cont_dt_Base), "DD/MM/YYYY")
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Data de Registro de Contrato
         If Not IsNull(XLO_EXPORTA!cont_dt_RegistroContrato) Then
              XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!cont_dt_RegistroContrato), "DD/MM/YYYY")
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Data de Entrega de Chaves
         If Not IsNull(XLO_EXPORTA!cont_dt_EntregaChaves) Then
              XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!cont_dt_EntregaChaves), "DD/MM/YYYY")
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
        
         'Valor Taxa Seg Vida(10)
         If Not IsNull(XLO_EXPORTA!cont_vl_TaxaSegVida) Then
                 XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_TaxaSegVida)) + Space(10 - Len(XLO_EXPORTA!cont_vl_TaxaSegVida))
         Else
              XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
         End If
         
         'Data de dISTRATO
         If Not IsNull(XLO_EXPORTA!cont_dt_Distrato) Then
              XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!cont_dt_Distrato), "DD/MM/YYYY")
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
        
         'Pegar na importação FornClientes(4)
       
         'Multa Pro-Rata(1)
         If Not IsNull(XLO_EXPORTA!cont_tx_MultaProRata) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!cont_tx_MultaProRata)
         Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
         End If
         
         'Correção Pro-Rata(1)
         If Not IsNull(XLO_EXPORTA!cont_tx_CorrecProRata) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!cont_tx_CorrecProRata)
         Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
         End If
         
         'Tipo de Seguro(1)
         If Not IsNull(XLO_EXPORTA!cont_nr_TipoSeguro) Or (XLO_EXPORTA!cont_nr_TipoSeguro) <> "" Then
              XLT_TEXTO = XLT_TEXTO + (XLO_EXPORTA!cont_nr_TipoSeguro)
         Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
         End If
         
         'Tipo Juros Mora(1)
         If Not IsNull(XLO_EXPORTA!cont_tx_TipoJurosMora) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!cont_tx_TipoJurosMora)
         Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
         End If
         
         'Desagio(1)
         If Not IsNull(XLO_EXPORTA!cont_tx_Desagio) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!cont_tx_Desagio)
         Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
         End If
         
         'Percent Juros Fin(10)
         If Not IsNull(XLO_EXPORTA!cont_vl_PercJurosFin) Then
            If XLO_EXPORTA!cont_vl_PercJurosFin < 0 Then
                XLT_TEXTO = XLT_TEXTO + FunTrataFloat(XLO_EXPORTA!cont_vl_PercJurosFin) + Space(11 - Len(XLO_EXPORTA!cont_vl_PercJurosFin))
            ElseIf XLO_EXPORTA!cont_vl_PercJurosFin < 1 And XLO_EXPORTA!cont_vl_PercJurosFin <> 0 Then
                XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_PercJurosFin)) + Space(11 - Len(XLO_EXPORTA!cont_vl_PercJurosFin))
            Else
                XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_PercJurosFin)) + Space(10 - Len(XLO_EXPORTA!cont_vl_PercJurosFin))
            End If
         Else
            XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
         End If
         
        'Percent Juros Multa(10)
         If Not IsNull(XLO_EXPORTA!cont_vl_PercJurosMulta) Then
            If XLO_EXPORTA!cont_vl_PercJurosMulta < 0 Then
                XLT_TEXTO = XLT_TEXTO + FunTrataFloat(XLO_EXPORTA!cont_vl_PercJurosMulta) + Space(11 - Len(XLO_EXPORTA!cont_vl_PercJurosMulta))
            ElseIf XLO_EXPORTA!cont_vl_PercJurosMulta < 1 And XLO_EXPORTA!cont_vl_PercJurosMulta <> 0 Then
                XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_PercJurosMulta)) + Space(11 - Len(XLO_EXPORTA!cont_vl_PercJurosMulta))
            Else
                XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_PercJurosMulta)) + Space(10 - Len(XLO_EXPORTA!cont_vl_PercJurosMulta))
            End If
         Else
            XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
         End If
        
         'Percent Juros Mora(10)
         If Not IsNull(XLO_EXPORTA!cont_vl_PercJurosMora) Then
            If (XLO_EXPORTA!cont_vl_PercJurosMora) < 0 Then
                XLT_TEXTO = XLT_TEXTO + FunTrataFloat(XLO_EXPORTA!cont_vl_PercJurosMora) + Space(11 - Len(XLO_EXPORTA!cont_vl_PercJurosMora))
            ElseIf XLO_EXPORTA!cont_vl_PercJurosMora < 1 And XLO_EXPORTA!cont_vl_PercJurosMora <> 0 Then
                XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_PercJurosMora)) + Space(11 - Len(XLO_EXPORTA!cont_vl_PercJurosMora))
            Else
                XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!cont_vl_PercJurosMora)) + Space(10 - Len(XLO_EXPORTA!cont_vl_PercJurosMora))
            End If
         Else
            XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
         End If
         
        'Status(1)
         If Not IsNull(XLO_EXPORTA!cont_tx_Status) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!cont_tx_Status) + Space(1 - Len(XLO_EXPORTA!cont_tx_Status))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
         End If
         
         'Contrato Origem(2)
         If Not IsNull(XLO_EXPORTA!cont_cd_ContratoOrigem) Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!cont_cd_ContratoOrigem) + Space(2 - Len(XLO_EXPORTA!cont_cd_ContratoOrigem))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(2)
         End If
         
         'Data da venda
         If Not IsNull(XLO_EXPORTA!cont_dt_Venda) Then
              XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!cont_dt_Venda), "DD/MM/YYYY")
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Data da Quitação
         If Not IsNull(XLO_EXPORTA!cont_dt_Quitacao) Then
              XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!cont_dt_Quitacao), "DD/MM/YYYY")
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Data Limite Reserva
         If Not IsNull(XLO_EXPORTA!cont_dt_LimiteReserva) Then
              XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!cont_dt_LimiteReserva), "DD/MM/YYYY")
         Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
         End If
         
         'Código do projeto(4) - Não vai
         
         'Código da C/Corrente(4)
         
         'Limite de Dias(4)
         If Not IsNull(XLO_EXPORTA!cont_nr_LimiteDias) Or (XLO_EXPORTA!cont_nr_LimiteDias) <> "" Then
              XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!cont_nr_LimiteDias) + Space(4 - Len(XLO_EXPORTA!cont_nr_LimiteDias))
         Else
              XLT_TEXTO = XLT_TEXTO + Space(4)
         End If
         
         Print #1, XLT_TEXTO
         XLO_EXPORTA.Close
         
         '*******Exportando os Títulos***********
         XLT_SQL = "SELECT * FROM titulos WHERE " & _
               " empd_cd_Empreendimento ='" & XLT_CODEMPREEND & "'" & _
               " AND imov_cd_Imovel = '" & XLT_CODIMOVEL & "'" & _
               " AND (titu_dt_Pagamento is null or titu_dt_Pagamento = '')  " & _
               " AND (cont_cd_contrato = '" & XLT_CONTRATO & "')" & _
               " AND empr_cd_empresa =" & PCodEmpresa
         SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
         
         While Not XLO_EXPORTA.EOF
            
            XLT_TEXTO = ""
            
            'Letra que Identifica o bloco
            XLT_TEXTO = "T"
            
            'Código do Plano(2)
            If Not IsNull(XLO_EXPORTA!titu_cd_plano) Then
                 XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!titu_cd_plano) + Space(2 - Len(XLO_EXPORTA!titu_cd_plano))
            Else
                 XLT_TEXTO = XLT_TEXTO + Space(2)
            End If
            
            'Código da Parcela(3)
            If Not IsNull(XLO_EXPORTA!titu_cd_Parcela) Then
                 XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!titu_cd_Parcela) + Space(3 - Len(XLO_EXPORTA!titu_cd_Parcela))
            Else
                 XLT_TEXTO = XLT_TEXTO + Space(3)
            End If
            
            'Código do Resíduo(3)
            If Not IsNull(XLO_EXPORTA!titu_cd_residuo) Then
                 XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!titu_cd_residuo) + Space(3 - Len(XLO_EXPORTA!titu_cd_residuo))
            Else
                 XLT_TEXTO = XLT_TEXTO + Space(3)
            End If
            
            'Código da Empresa(4) - Pegar na importação
            
            'Código do Empreendimento(4) - Pegar na importação
            
            'Código do Imóvel(4)
            If Not IsNull(XLO_EXPORTA!imov_cd_Imovel) Then
                 XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!imov_cd_Imovel) + Space(4 - Len(XLO_EXPORTA!imov_cd_Imovel))
            Else
                 XLT_TEXTO = XLT_TEXTO + Space(4)
            End If
            
            'Código do Contrato(2)
            If Not IsNull(XLO_EXPORTA!cont_cd_Contrato) Then
                 XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!cont_cd_Contrato) + Space(2 - Len(XLO_EXPORTA!cont_cd_Contrato))
            Else
                 XLT_TEXTO = XLT_TEXTO + Space(2)
            End If
            
            'Tipo do Plano(20)
            If Not IsNull(XLO_EXPORTA!tipl_cd_TipoPlano) Then
            
                Sql = "SELECT tipl_tx_Descricao FROM TiposPlanos where tipl_cd_TipoPlano = " & XLO_EXPORTA!tipl_cd_TipoPlano & ""
                SubQOpenRecordset XLO_CODIGOS, Sql, Estatico
                     
                If Not XLO_CODIGOS.EOF Then
                   XLT_TEXTO = XLT_TEXTO + Trim(XLO_CODIGOS!tipl_tx_Descricao) + Space(20 - Len(XLO_CODIGOS!tipl_tx_Descricao))
                Else
                    XLT_TEXTO = XLT_TEXTO + Space(20)
                End If
                    
                XLO_CODIGOS.Close
            Else
                XLT_TEXTO = XLT_TEXTO + Space(20)
            End If
            
            
            'C/Corrente(4) - Na importação
            
            'Observação(50)
            If Not IsNull(XLO_EXPORTA!obse_cd_Observacao) Then
               Sql = "SELECT obse_tx_Observacao FROM Observacoes where obse_cd_observacao = " & XLO_EXPORTA!obse_cd_Observacao & ""
               SubQOpenRecordset XLO_CODIGOS, Sql, Estatico
            
               If Not XLO_CODIGOS.EOF Then
                   XLT_TEXTO = XLT_TEXTO + Trim(XLO_CODIGOS!obse_tx_Observacao) + Space(50 - Len(XLO_CODIGOS!obse_tx_Observacao))
               Else
                  XLT_TEXTO = XLT_TEXTO + Space(50)
               End If
            
               XLO_CODIGOS.Close
            Else
                XLT_TEXTO = XLT_TEXTO + Space(50)
            End If
            
            'Código da Moeda1 - Na importação
            
            'Código da Moeda2 - Na importação
            
            'Valor da Parcela(10)
            If Not IsNull(XLO_EXPORTA!titu_vl_Parcela) Then
                    XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!titu_vl_Parcela)) + Space(20 - Len(XLO_EXPORTA!titu_vl_Parcela))
            Else
                 XLT_TEXTO = XLT_TEXTO + "0" + Space(19)
            End If
            
            'Valor da Parcela Indexado(10)
            
             If Not IsNull(XLO_EXPORTA!titu_vl_ParcelaIndex1) Then
                If XLO_EXPORTA!titu_vl_ParcelaIndex1 < 0 Then
                    XLT_TEXTO = XLT_TEXTO + FunTrataFloat(XLO_EXPORTA!titu_vl_ParcelaIndex1) + Space(11 - Len(XLO_EXPORTA!titu_vl_ParcelaIndex1))
                ElseIf XLO_EXPORTA!titu_vl_ParcelaIndex1 < 1 And XLO_EXPORTA!titu_vl_ParcelaIndex1 <> 0 Then
                    XLT_TEXTO = XLT_TEXTO + FunTrataFloat(XLO_EXPORTA!titu_vl_ParcelaIndex1) + Space(11 - Len(XLO_EXPORTA!titu_vl_ParcelaIndex1))
                Else
                    XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!titu_vl_ParcelaIndex1)) + Space(10 - Len(XLO_EXPORTA!titu_vl_ParcelaIndex1))
                End If
            Else
                XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
            End If
            
    
            'Valor Corrigido1(10)
            If Not IsNull(XLO_EXPORTA!titu_vl_Corrigido1) Then
                      XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!titu_vl_Corrigido1)) + Space(10 - Len(XLO_EXPORTA!titu_vl_Corrigido1))
            Else
                 XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
            End If
            
            'Valor da Parcela Indexado2(10)
            If Not IsNull(XLO_EXPORTA!titu_vl_ParcelaIndex2) Then
                    XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!titu_vl_ParcelaIndex2)) + Space(10 - Len(XLO_EXPORTA!titu_vl_ParcelaIndex2))
            Else
                 XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
            End If
            
            'Valor Corrigido2(10)
            If Not IsNull(XLO_EXPORTA!titu_vl_Corrigido2) Then
                    XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!titu_vl_Corrigido2)) + Space(10 - Len(XLO_EXPORTA!titu_vl_Corrigido2))
            Else
                 XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
            End If
            
            'Data de Vencimento
            If Not IsNull(XLO_EXPORTA!titu_dt_Vencimento) Then
               XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!titu_dt_Vencimento), "DD/MM/YYYY")
            Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
            End If
            
             'Indice Pagamento
            If Not IsNull(XLO_EXPORTA!titu_tx_IndicePagamento) Then
               XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!titu_tx_IndicePagamento)
            Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
            End If
            
            'Financia
            If Not IsNull(XLO_EXPORTA!titu_tx_Financia) Then
               XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!titu_tx_Financia)
            Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
            End If
            
            'Perc Juros Finan(10)
            If Not IsNull(XLO_EXPORTA!titu_vl_PercJurosFin) Then
                If XLO_EXPORTA!titu_vl_PercJurosFin < 0 Then
                    XLT_TEXTO = XLT_TEXTO + FunTrataFloat(XLO_EXPORTA!titu_vl_PercJurosFin) + Space(11 - Len(XLO_EXPORTA!titu_vl_PercJurosFin))
                ElseIf XLO_EXPORTA!titu_vl_PercJurosFin < 1 And XLO_EXPORTA!titu_vl_PercJurosFin <> 0 Then
                    XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!titu_vl_PercJurosFin)) + Space(11 - Len(XLO_EXPORTA!titu_vl_PercJurosFin))
                Else
                    XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!titu_vl_PercJurosFin)) + Space(10 - Len(XLO_EXPORTA!titu_vl_PercJurosFin))
                End If
            Else
                XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
            End If
            
        
            'Perc Multa Mora(10)
            If Not IsNull(XLO_EXPORTA!titu_vl_PercMultaMora) Then
                If XLO_EXPORTA!titu_vl_PercMultaMora < 0 Then
                    XLT_TEXTO = XLT_TEXTO + FunTrataFloat(XLO_EXPORTA!titu_vl_PercMultaMora) + Space(11 - Len(XLO_EXPORTA!titu_vl_PercMultaMora))
                ElseIf XLO_EXPORTA!titu_vl_PercMultaMora < 1 And XLO_EXPORTA!titu_vl_PercMultaMora <> 0 Then
                    XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!titu_vl_PercMultaMora)) + Space(11 - Len(XLO_EXPORTA!titu_vl_PercMultaMora))
                Else
                    XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!titu_vl_PercMultaMora)) + Space(10 - Len(XLO_EXPORTA!titu_vl_PercMultaMora))
                End If
            Else
                XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
            End If
            
            'Perc Juros Mora(10)
             If Not IsNull(XLO_EXPORTA!titu_vl_PercJurosMora) Then
                If XLO_EXPORTA!titu_vl_PercJurosMora < 0 Then
                    XLT_TEXTO = XLT_TEXTO + FunTrataFloat(XLO_EXPORTA!titu_vl_PercJurosMora) + Space(11 - Len(XLO_EXPORTA!titu_vl_PercJurosMora))
                ElseIf XLO_EXPORTA!titu_vl_PercJurosMora < 1 And XLO_EXPORTA!titu_vl_PercJurosMora <> 0 Then
                    XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!titu_vl_PercJurosMora)) + Space(11 - Len(XLO_EXPORTA!titu_vl_PercJurosMora))
                Else
                    XLT_TEXTO = XLT_TEXTO + Trim(FunTrataFloat(XLO_EXPORTA!titu_vl_PercJurosMora)) + Space(10 - Len(XLO_EXPORTA!titu_vl_PercJurosMora))
                End If
            Else
               XLT_TEXTO = XLT_TEXTO + "0" + Space(9)
            End If
            
    
            'Data de Repasse
            If Not IsNull(XLO_EXPORTA!titu_dt_Repasse) Then
               XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!titu_dt_Repasse), "DD/MM/YYYY")
            Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
            End If
            
           'Data de Prorrogação
            If Not IsNull(XLO_EXPORTA!titu_dt_Prorrogacao) Then
               XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!titu_dt_Prorrogacao), "DD/MM/YYYY")
            Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
            End If
            
            'Tipo de Pagamento(2)
            If Not IsNull(XLO_EXPORTA!titu_tx_TipoPag) Then
                 XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!titu_tx_TipoPag) + Space(2 - Len(XLO_EXPORTA!titu_tx_TipoPag))
            Else
                 XLT_TEXTO = XLT_TEXTO + Space(2)
            End If
            
            'Data Base(10)
            If Not IsNull(XLO_EXPORTA!titu_dt_Base) Then
               XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!titu_dt_Base), "DD/MM/YYYY")
            Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
            End If
            
            'CR (8) - Pegar na importação
            
            'Conta(7) - Pegar na importação
            
            'Natureza do Plano(1)
            If Not IsNull(XLO_EXPORTA!titu_tx_NaturezaPlano) Then
               XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!titu_tx_NaturezaPlano)
            Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
            End If
            
            'Corrige Parcela(1)
            If Not IsNull(XLO_EXPORTA!titu_tx_CorrigeParcela) Then
               XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!titu_tx_CorrigeParcela)
            Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
            End If
            
            'Pré_datado(1)
            If Not IsNull(XLO_EXPORTA!titu_tx_PreDatado) Then
               XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!titu_tx_PreDatado)
            Else
              XLT_TEXTO = XLT_TEXTO + Space(1)
            End If
            
            'Data Base de Pagamento(1)
            If Not IsNull(XLO_EXPORTA!titu_dt_BasePagto) Then
               XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!titu_dt_BasePagto), "DD/MM/YYYY")
            Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
            End If
            
             'Data de Criação(10)
            If Not IsNull(XLO_EXPORTA!titu_dt_Criacao) Then
               XLT_TEXTO = XLT_TEXTO + Format(Trim(XLO_EXPORTA!titu_dt_Criacao), "DD/MM/YYYY")
            Else
              XLT_TEXTO = XLT_TEXTO + Space(10)
            End If
            
            'Nota Fiscal(4)
            If Not IsNull(XLO_EXPORTA!nofi_cd_NotaFiscal) Then
                 XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!nofi_cd_NotaFiscal) + Space(4 - Len(XLO_EXPORTA!nofi_cd_NotaFiscal))
            Else
                 XLT_TEXTO = XLT_TEXTO + Space(4)
            End If
            
            'Exporta p bancos(1)
            If Not IsNull(XLO_EXPORTA!titu_tx_ExpBanco) Then
                 XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!titu_tx_ExpBanco)
            Else
                 XLT_TEXTO = XLT_TEXTO + Space(1)
            End If
            
            'Entrega de Chaves(1)
            If Not IsNull(XLO_EXPORTA!titu_tx_EntregaChave) Then
                 XLT_TEXTO = XLT_TEXTO + Trim(XLO_EXPORTA!titu_tx_EntregaChave)
            Else
                 XLT_TEXTO = XLT_TEXTO + Space(1)
            End If
         
            Print #1, XLT_TEXTO
            XLO_EXPORTA.MoveNext
         Wend
         XLO_EXPORTA.Close
      End If
      TDBGrid1.MoveNext
      
   Wend
   Close #1
    MsgBox "Exportação realizada com sucesso.", vbInformation + vbOKOnly, "ATENÇÃO"
    Unload Me
Else
    MsgBox "Não foi possível concluir a exportação!.", vbInformation + vbOKOnly, "ATENÇÃO"
End If
    
End Sub

Private Sub CmdSair_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   
    Set Formulario = FrmExportEmpresas
    
   ' Call subAjustaTelaLoc
    
    subManutencaoJanelasAtivas "I", "FrmExportEmpresas"
         
    subConectarControleDadosNV DatEmpreendimentos, "SELECT * FROM Empreendimentos WHERE empr_cd_Empresa = " & PCodEmpresa & " ORDER BY empd_tx_Nome", Estatico
    
    subConectarControleDadosNV DatMoedas, "SELECT * FROM Moedas ORDER BY moed_tx_descricao", Estatico
    
    'subHabilitaBotoes
    filtra_especial = True
    
    'subTelaValoresGlobais "G"
    PrimeiraVez = True
    
    If BarraFerramentasPressionada = True Then
      'Call CmdInserir_Click
      BarraFerramentasPressionada = False
    End If
 
End Sub

Private Sub TDBGrid1_Click()
'  If Not TDBGrid1.EOF And Not TDBGrid1.BOF Then
'    If TDBGrid1.Columns("Exporta").Value = "0" Then
'      TDBGrid1.Columns("Exporta").Value = "-1"
'    End If
'    TDBGrid1.Refresh
'  End If
End Sub


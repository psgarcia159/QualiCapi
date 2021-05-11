VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmAtribuicoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Atribuições"
   ClientHeight    =   3840
   ClientLeft      =   30
   ClientTop       =   1335
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraProjeto 
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11775
      Begin MSDataListLib.DataCombo CboProjeto 
         Bindings        =   "FrmAtribuicoes.frx":0000
         Height          =   315
         Left            =   900
         TabIndex        =   15
         Top             =   180
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   -2147483643
         ListField       =   "proj_tx_Nome"
         BoundColumn     =   "proj_cd_Projeto"
         Text            =   "CboProjeto"
      End
      Begin MSDataListLib.DataCombo CboSubProjeto 
         Bindings        =   "FrmAtribuicoes.frx":0019
         Height          =   315
         Left            =   6300
         TabIndex        =   16
         Top             =   180
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "SubProjeto"
         BoundColumn     =   "Chave"
         Text            =   "CboSubProjeto"
      End
      Begin Threed.SSCommand CmdLimparSubProjeto 
         Height          =   315
         Left            =   11340
         TabIndex        =   17
         Top             =   180
         Width           =   360
         _Version        =   65536
         _ExtentX        =   635
         _ExtentY        =   556
         _StockProps     =   78
         MouseIcon       =   "FrmAtribuicoes.frx":0035
         Picture         =   "FrmAtribuicoes.frx":0487
      End
      Begin VB.Label LblProjeto 
         Caption         =   "Projeto:"
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
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblSubProjeto 
         Caption         =   "SubProjeto:"
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
         Left            =   5220
         TabIndex        =   18
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame FraDadosGerais 
      Height          =   2355
      Left            =   0
      TabIndex        =   3
      Top             =   660
      Width           =   11775
      Begin VB.TextBox TxtHHPrevisto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5280
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TxtHHReal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7740
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo CboTarefa 
         Bindings        =   "FrmAtribuicoes.frx":0599
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Tarefa"
         BoundColumn     =   "tare_cd_UID"
         Text            =   "CboTarefa"
      End
      Begin MSDataListLib.DataCombo CboRecurso 
         Bindings        =   "FrmAtribuicoes.frx":05B1
         Height          =   315
         Left            =   900
         TabIndex        =   5
         Top             =   720
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "recu_tx_Nome"
         BoundColumn     =   "Chave"
         Text            =   "CboRecurso"
      End
      Begin MSDataListLib.DataCombo CboDia 
         Bindings        =   "FrmAtribuicoes.frx":05CA
         Height          =   315
         Left            =   900
         TabIndex        =   13
         Top             =   1140
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "atri_dt_Dia"
         BoundColumn     =   "atri_dt_Dia"
         Text            =   "CboDia"
      End
      Begin MSDataListLib.DataCombo CboNivel 
         Bindings        =   "FrmAtribuicoes.frx":05DF
         Height          =   315
         Left            =   900
         TabIndex        =   20
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   -2147483643
         ListField       =   "Nivel"
         BoundColumn     =   "Nivel"
         Text            =   "CboNivel"
      End
      Begin VB.Label LblNivel 
         Caption         =   "Nível:"
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
         TabIndex        =   21
         Top             =   300
         Width           =   555
      End
      Begin VB.Label LblDia 
         Alignment       =   1  'Right Justify
         Caption         =   "Dia:"
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
         Left            =   300
         TabIndex        =   12
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label LblHHPrevisto 
         Alignment       =   1  'Right Justify
         Caption         =   "HH Previsto:"
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
         Left            =   3960
         TabIndex        =   10
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label LblHHReal 
         Caption         =   "HH Real:"
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
         Left            =   6780
         TabIndex        =   8
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label LblTarefa 
         Caption         =   "Tarefa:"
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
         Left            =   1980
         TabIndex        =   7
         Top             =   300
         Width           =   615
      End
      Begin VB.Label LblRecurso 
         Caption         =   "Recurso:"
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
         Left            =   60
         TabIndex        =   6
         Top             =   780
         Width           =   795
      End
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   8520
      TabIndex        =   1
      Top             =   3180
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Sair"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdGravar 
      Height          =   330
      Left            =   6120
      TabIndex        =   0
      Top             =   3180
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Gravar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdNovo 
      Height          =   330
      Left            =   3900
      TabIndex        =   2
      Top             =   3180
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Novo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin MSAdodcLib.Adodc DatRecurso 
      Height          =   330
      Left            =   -120
      Top             =   3180
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
      Caption         =   "DatRecurso"
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
   Begin MSAdodcLib.Adodc DatTarefa 
      Height          =   330
      Left            =   -120
      Top             =   3180
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
      Caption         =   "DatTarefa"
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
   Begin MSAdodcLib.Adodc DatDia 
      Height          =   330
      Left            =   -120
      Top             =   3180
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
      Caption         =   "DatDia"
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
   Begin MSAdodcLib.Adodc DatProjeto 
      Height          =   330
      Left            =   0
      Top             =   3180
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
      Caption         =   "DatProjeto"
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
   Begin MSAdodcLib.Adodc DatSubProjeto 
      Height          =   330
      Left            =   0
      Top             =   3180
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
      Caption         =   "DatSubprojeto"
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
   Begin MSAdodcLib.Adodc DatNivel 
      Height          =   330
      Left            =   0
      Top             =   3180
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
      Caption         =   "DatNivel"
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
Attribute VB_Name = "FrmAtribuicoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim XFT_CODIMOVEL As String 'Código do imóvel
Dim XFT_CODEMPREENDIMENTO As String  'Código do empreendimento

Function funChecarCampos() As Boolean
  
  funChecarCampos = False
  
  If FunObrigatorioMSK(MskEmpreendimento, "O Empreendimento é Obrigatório") Then Exit Function
  If FunObrigatorioMSK(MskCodigo, "O Código do Imóvel é Obrigatório") Then Exit Function
  If FunObrigatorioCBO(CboTipoImovel, "O Tipo do Imóvel é Obrigatório") Then Exit Function
  
  funChecarCampos = True

End Function

Sub SubDesabHabCampos(XHab As Boolean)
    
    'Desabilita/Habilita os campos em caso de inserção
    TxtArea.Enabled = XHab
    TxtAndar.Enabled = XHab
    TxtVlImovel.Enabled = XHab
    TxtGaragem.Enabled = XHab
    CboTipoImovel.Enabled = XHab
    DtpEntChaves.Enabled = XHab
    
End Sub

Sub subPreencheCampos()
   
   Dim XLT_SQL As String
   Dim XLO_IMOVEL As New ADODB.Recordset 'Pega os dados do Imóvel
    
  SubQOpenRecordset XLO_IMOVEL, "SELECT * FROM Imoveis WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento= '" & XFT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XFT_CODIMOVEL & "'", "Estatico"
  If Not (XLO_IMOVEL.BOF And XLO_IMOVEL.EOF) Then
  
  
    'Dados Gerais
    MskEmpreendimento.Text = XLO_IMOVEL!empd_cd_Empreendimento
    CboEmpreendimento.BoundText = XLO_IMOVEL!empd_cd_Empreendimento
  
    'Se não for cópia de imóvel
    If XInserir <> "C" Then
      MskCodigo.Mask = ""
      MskCodigo.Text = XLO_IMOVEL!imov_cd_Imovel
      MskCodigo.Mask = "####"
    End If
    
    Select Case XLO_IMOVEL!imov_tx_Status
      Case "D"
          LblDescStatus.Caption = "Disponível"
      Case "V"
          LblDescStatus.Caption = "Vendido"
      Case "R"
          LblDescStatus.Caption = "Reservado"
    End Select
    
    TxtArea.Text = Format(XLO_IMOVEL!imov_nr_Area, "standard")
  
    TxtGaragem.Text = FunNulo(XLO_IMOVEL!imov_tx_Garagem)
    TxtAndar.Text = FunNulo(XLO_IMOVEL!imov_nr_Andar)
    CboTipoImovel.BoundText = XLO_IMOVEL!tiim_cd_TipoImovel
    
    If Not IsNull(XLO_IMOVEL!imov_dt_EntregaChaves) Then DtpEntChaves.Value = XLO_IMOVEL!imov_dt_EntregaChaves
    
    'Dados da Venda
    TxtVlImovel.Text = Format(XLO_IMOVEL!imov_vl_Avista, "standard")
    
    Chave = XLO_IMOVEL!empd_cd_Empreendimento + "." + XLO_IMOVEL!imov_cd_Imovel
  End If
  XLO_IMOVEL.Close
  Set XLO_IMOVEL = Nothing
    
End Sub

Sub subPrepararInsercao()
    
    MskCodigo.Mask = ""
    MskCodigo.Text = ""
    MskCodigo.Mask = "####"
    LblDescStatus.Caption = ""
    TxtArea.Text = ""
    TxtAndar.Text = ""
    TxtVlImovel.Text = ""
    TxtGaragem.Text = ""
    CboTipoImovel.Text = ""
    DtpEntChaves.Value = Null

End Sub

Private Sub CboProjeto_Change()
  
    
  CboSubProjeto.BoundText = ""
  CboNivel.BoundText = ""
  CboRecurso.BoundText = ""
  
  If CboProjeto.BoundText <> "" Then
    
    CboNivel.Enabled = True
    subConectarControleDados DatNivel, "SELECT * FROM ConsSASNivelProjeto " & _
      " WHERE proj_cd_Projeto=" & CboProjeto.BoundText, adOpenStatic
        
    CboSubProjeto.Enabled = True
    subConectarControleDados DatSubProjeto, "SELECT * FROM ConsSASSubProjetos " & _
      " WHERE proj_cd_Projeto=" & CboProjeto.BoundText, adOpenStatic
    
    CboRecurso.Enabled = True
    subConectarControleDados DatRecurso, "SELECT * FROM ConsSASRecursos " & _
     " WHERE proj_cd_Projeto=" & CboProjeto.BoundText & " ORDER BY recu_tx_Nome", adOpenStatic
    
  Else
   CboSubProjeto.Enabled = False
   CboNivel.Enabled = False
   CboRecurso.Enabled = False
   
  End If

End Sub

Private Sub CboSubProjeto_Change()

  Dim XLT_SQL As String
  
  Dim XLT_SQL5 As String
  
  
  XLT_SQL = "SELECT * FROM ConsSASRecursos " & _
    "WHERE proj_cd_Projeto=" & DatSubProjeto.Recordset!proj_cd_Projeto
  

  'Se tiver subprojeto filtra pelo subprojeto senão filtra pelo projeto master
  If CboSubProjeto.BoundText <> "" Then
    
    DatSubProjeto.Recordset.Bookmark = CboSubProjeto.SelectedItem
    
    XLT_SQL = XLT_SQL & _
     " AND subp_cd_SubProjeto=" & DatSubProjeto.Recordset!subp_cd_SubProjeto
    
    CboRecurso.Enabled = True
    
  End If
  
  CboNivel.BoundText = ""
  CboRecurso.Text = ""
  
  subConectarControleDados DatRecurso, XLT_SQL, adOpenStatic
  subConectarControleDados DatNivel, XLT_SQL5, adOpenStatic
 
End Sub

Private Sub CmdDesistir_Click()
    Unload Me
End Sub

Private Sub cmdGravar_Click()
      
  Dim XLT_SQL As String
  Dim XLO_RSEmp As Object 'record set para testar o numero de imoveis do empreendimento
  Dim XLO_RSImoveis As Object
  Dim XLS_SQLTemp As String
  
  
  If funChecarCampos Then
  
    On Error GoTo RotuloErro
           
    'testar qtd de imoveis
    XLS_SQLTemp = "SELECT empd_nr_unidades FROM Empreendimentos " & _
                 "WHERE empd_cd_Empreendimento = '" & MskEmpreendimento.Text & "'" & _
                 " AND empr_cd_empresa = " & PCodEmpresa
                
                
    SubQOpenRecordset XLO_RSEmp, XLS_SQLTemp, Estatico
    
    XLS_SQLTemp = "SELECT COUNT(imov_cd_Imovel) AS qtdImoveis " & _
                  "From imoveis " & _
                  "WHERE empr_cd_empresa = " & PCodEmpresa & _
                  "AND empd_cd_empreendimento = '" & MskEmpreendimento.Text & "'"
  
    SubQOpenRecordset XLO_RSImoveis, XLS_SQLTemp, Estatico
    

  
    Conexao.BeginTrans
    If XInserir = "I" Or XInserir = "C" Then
      If (Not IsNull(XLO_RSEmp!empd_nr_Unidades)) Then
        If (XLO_RSImoveis!qtdImoveis >= XLO_RSEmp!empd_nr_Unidades) Then
          MsgBox "Este empreendimento ja possui o numero maximo de imoveis cadastrados", vbCritical, "ATENÇÃO"
          Conexao.RollbackTrans
          Exit Sub
        End If
      End If
      'Fecha o RecordSet
      XLO_RSImoveis.Close
      XLO_RSEmp.Close
      'descarrega o RecordSet da Memoria
      Set XLO_RSImoveis = Nothing
      Set XLO_RSEmp = Nothing
      XLT_SQL = "INSERT INTO Imoveis (imov_cd_imovel, empd_cd_empreendimento,empr_cd_Empresa, " & _
                  "imov_nr_area, imov_tx_garagem, imov_nr_andar,tiim_cd_TipoImovel,  " & _
                  "imov_dt_EntregaChaves,imov_vl_AVista, imov_tx_status) " & _
            "VALUES ('" & MskCodigo.Text & "','" & MskEmpreendimento.Text & "'," & PCodEmpresa & ", " & _
                  "'" & FunNuloVal(FunTrataFloat(TxtArea.Text)) & "','" & TxtGaragem.Text & "','" & TxtAndar.Text & "', " & _
                  "" & CboTipoImovel.BoundText & "," & FunNuloData(DtpEntChaves.Value, NomeSgbd) & ", " & _
                  "" & FunNuloVal(FunTrataFloat(TxtVlImovel.Text)) & ",'D')"
                     
      Conexao.Execute XLT_SQL
      Chave = MskEmpreendimento.Text + "." + MskCodigo.Text
      Me.Caption = "Alteração de Imóveis"
    Else    'Alteração
       XLT_SQL = "UPDATE Imoveis SET imov_nr_area='" & FunNuloVal(FunTrataFloat(TxtArea.Text)) & "'," & _
                   "imov_tx_garagem='" & TxtGaragem.Text & "'," & _
                   "imov_nr_andar='" & TxtAndar.Text & "'," & _
                   "tiim_cd_TipoImovel=" & CboTipoImovel.BoundText & "," & _
                   "imov_dt_EntregaChaves=" & FunNuloData(DtpEntChaves.Value, NomeSgbd) & "," & _
                   "imov_vl_AVista=" & FunNuloVal(FunTrataFloat(TxtVlImovel.Text)) & "" & _
               " WHERE imov_cd_imovel='" & MskCodigo.Text & "' AND empd_cd_empreendimento='" & MskEmpreendimento.Text & "' AND empr_cd_Empresa=" & PCodEmpresa & ""
       Conexao.Execute XLT_SQL
    End If
  
    Conexao.CommitTrans
              
    If XInserir = "A" And CmdNovo.Enabled = False Then
        Call CmdDesistir_Click
    Else
      CmdNovo.Enabled = True
      XInserir = "A"
      MskEmpreendimento.Enabled = False
      CboEmpreendimento.Enabled = False
      MskCodigo.Enabled = False
    End If

  End If
  Exit Sub
    
RotuloErro:
    funTrataErros (ComMensagem)
    Resume Next
   
End Sub

Private Sub CmdNovo_Click()

  'Chave = -1
  XInserir = "I"
  Me.Caption = "Cadastro de Imóveis"
  SubDesabHabCampos (False)
  MskEmpreendimento.Enabled = True
  CboEmpreendimento.Enabled = True
  MskCodigo.Enabled = True
  MskCodigo.Text = ""
  MskCodigo.SetFocus
  TxtGaragem.Text = ""
  CmdNovo.Enabled = False
  
End Sub

Private Sub Form_Activate()
    Set Formulario = FrmManImoveis
End Sub

Private Sub Form_Load()
    Dim XLT_SQL  As String
      
    Set Formulario = FrmManImoveis
        
    If BarraFerramentasPressionada Then
        XInserir = "I"
    End If
   ' CliqueBarraFerramentas = False
   
   subConectarControleDados DatRecurso, "SELECT * FROM ConsSASRecursos " & _
     " WHERE proj_cd_Projeto=" & CboProjeto.BoundText & " ORDER BY recu_tx_Nome", adOpenStatic
    
    
    subConectarControleDadosNV DatTipoImovel, "SELECT * FROM TiposImoveis ORDER BY tiim_tx_Descricao", Estatico
        
    If XInserir = "I" Then
        LblDescStatus.Caption = "Disponível"
        SubDesabHabCampos (False)
    ElseIf XInserir = "A" Then
        
        Me.Caption = "Alteração de Imóveis"
        'Quebra a chave em codigo do empreendimento e codigo do imóvel
        XFT_CODEMPREENDIMENTO = Left$(Chave, 4)
        XFT_CODIMOVEL = Right$(Chave, 4)
        subPreencheCampos
        MskEmpreendimento.Enabled = False
        CboEmpreendimento.Enabled = False
        MskCodigo.Enabled = False
        CmdNovo.Enabled = False
    ElseIf XInserir = "C" Then
        
        Me.Caption = "Cópia de Imóveis"
        'Quebra a chave em codigo do empreendimento e codigo do imóvel
        XFT_CODEMPREENDIMENTO = Left$(Chave, 4)
        XFT_CODIMOVEL = Right$(Chave, 4)
        subPreencheCampos
        CmdNovo.Enabled = False
    
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

End Sub

Private Sub TxtHHReal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtHHReal_LostFocus()
    If IsNumeric(TxtHHReal.Text) Then
        TxtHHReal.Text = Format$(TxtHHReal.Text, "##,##0.00")
    Else
        TxtHHReal.Text = ""
    End If
End Sub

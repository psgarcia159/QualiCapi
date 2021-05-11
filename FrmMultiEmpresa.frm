VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmMultiEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleção de Empresa Ativa"
   ClientHeight    =   2220
   ClientLeft      =   3705
   ClientTop       =   3540
   ClientWidth     =   5205
   Icon            =   "FrmMultiEmpresa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSCommand CmdConfirmar 
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   1755
      Width           =   960
      _Version        =   65536
      _ExtentX        =   1693
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Confirmar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin MSAdodcLib.Adodc DatEmpresas 
      Height          =   330
      Left            =   30
      Top             =   1755
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "DatEmpresas"
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
   Begin Threed.SSPanel PanEmpresaAtiva 
      Height          =   540
      Left            =   60
      TabIndex        =   1
      Top             =   195
      Width           =   5100
      _Version        =   65536
      _ExtentX        =   8996
      _ExtentY        =   952
      _StockProps     =   15
      ForeColor       =   128
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelInner      =   1
      FloodColor      =   128
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ativar Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   45
      TabIndex        =   2
      Top             =   900
      Width           =   5115
      Begin MSDataListLib.DataCombo CboEmpresas 
         Bindings        =   "FrmMultiEmpresa.frx":0442
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   315
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "empr_tx_fantasia"
         BoundColumn     =   "empr_cd_empresa"
         Text            =   "CboEmpresas"
      End
   End
End
Attribute VB_Name = "FrmMultiEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private XFT_EMPRESA As String

Public Property Let Empresa(XLT_EMPRESA As String)
  XFT_EMPRESA = XLT_EMPRESA
End Property


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 And ((UCase(PFantasia) = "APAE") Or (UCase(PFantasia) <> "INSTITUTO ALIANÇA")) Then
        Unload Me
    End If
End Sub

Private Sub CmdConfirmar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 And ((UCase(PFantasia) = "APAE") Or (UCase(PFantasia) <> "INSTITUTO ALIANÇA")) Then
        Unload Me
    End If
End Sub

Private Sub CboEmpresas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 And ((UCase(PFantasia) = "APAE") Or (UCase(PFantasia) <> "INSTITUTO ALIANÇA")) Then
        Unload Me
    End If
End Sub

Private Sub CmdConfirmar_Click()
  Dim XLO_ESTOQUE As ADODB.Recordset
      
  If CboEmpresas.Text <> "" Then
    DatEmpresas.Recordset.Bookmark = CboEmpresas.SelectedItem
    
    PCodEmpresa = DatEmpresas.Recordset.Fields("empr_cd_empresa")
    PFantasia = DatEmpresas.Recordset.Fields("empr_tx_fantasia")
    PEmpresa = DatEmpresas.Recordset.Fields("empr_tx_razaosocial")
    PEndereco = FunNulo(DatEmpresas.Recordset.Fields("empr_tx_endereco"))
    PCidade = FunNulo(DatEmpresas.Recordset.Fields("empr_tx_cidade"))
    pPCRmascara = DatEmpresas.Recordset.Fields("empr_tx_pCRMascara")
    
    'Inserido em 31/07/08 - Larissa
    If FunNulo(DatEmpresas.Recordset.Fields("empr_tx_VinculaCrCc")) = "S" Then
      pVinculaCrCc = True
    Else
      pVinculaCrCc = False
    End If
    
    XGI_QTDDIASFILTRO = FunNuloVal(DatEmpresas.Recordset.Fields("empr_nr_QtdDiasFiltro"))
    If DatEmpresas.Recordset.Fields("empr_tx_IncluiCancelados") = "S" Then
        XGB_INCLUICANCELADOS = True
    Else
        XGB_INCLUICANCELADOS = False
    End If
    If DatEmpresas.Recordset.Fields("empr_tx_CrAtivoPassivo") = "S" Then
      XGB_APROPRIACRATIVOPASS = True
    Else
      XGB_APROPRIACRATIVOPASS = False
    End If
    XGI_CONTACORRENTE = FunNuloVal(DatEmpresas.Recordset.Fields("empr_cd_CodContaCorrente"))
    XGI_QTDDIASPROVISAO = FunNuloVal(DatEmpresas.Recordset.Fields("empr_nr_qtdDiasProvisao"))
    XGT_TIPOEXPORTACAO = DatEmpresas.Recordset.Fields("empr_tx_ExpContabil")
    XGT_EXPCOMPROMCAPI = FunNulo(DatEmpresas.Recordset.Fields("empr_tx_ExpCompromissoCapi"))
    XGT_INCLUIRACRESCIMOS = FunNulo(DatEmpresas.Recordset!empr_tx_IncluiAcrescimos)
    XGT_MODELORECIBO = FunNulo(DatEmpresas.Recordset!empr_tx_ModeloRecibo)
    XGB_CNPJOBRIGATORIO = FunNuloVal(DatEmpresas.Recordset!empr_tx_CnpjObrigatorio)
    XGT_IMPRIMEPP = FunNulo(DatEmpresas.Recordset.Fields("empr_tx_ImprimePP"))
    XGT_ORDEMDESCONTO = FunNulo(DatEmpresas.Recordset.Fields("empr_tx_OrdemDescontos"))
    XGT_CGC = FunNuloVal(DatEmpresas.Recordset!empr_tx_CGC)
    XGT_CEP = FunNuloVal(DatEmpresas.Recordset!empr_tx_Cep)
    XGT_NUMPORTA = FunNuloVal(DatEmpresas.Recordset!empr_nr_Porta)
    XGT_BAIRRO = FunNuloVal(DatEmpresas.Recordset!empr_tx_Bairro)
    XGT_ESTADO = FunNuloVal(DatEmpresas.Recordset!empr_tx_Estado)
    pPCRnivel = FunNuloVal(DatEmpresas.Recordset.Fields("empr_nr_pcrnivel"))
    pPlanoCmascara = FunNulo(DatEmpresas.Recordset.Fields("empr_tx_planocmascara"))
    pPlanoCNivel = FunNuloVal(DatEmpresas.Recordset.Fields("empr_nr_planocnivel"))
    PLogo = FunNulo(DatEmpresas.Recordset.Fields("empr_im_caminhoLogo"))
    XGI_PCATEGORIANVFINANCEIRO = FunNuloVal(DatEmpresas.Recordset.Fields("empr_nr_CategoriaNvContabil"))
    
    If DatEmpresas.Recordset.Fields("empr_tx_EnviarEmail") = "S" Then
      XGB_EMAIL = True
    Else
      XGB_EMAIL = False
    End If
    
    If DatEmpresas.Recordset.Fields("empr_tx_AlertarSaldoAdiantamento") = "S" Then
      XGB_AlertarSaldoAdiantamento = True
    Else
      XGB_AlertarSaldoAdiantamento = False
    End If
    
    PanEmpresaAtiva.Caption = PFantasia
    MdiPrincipal.BarraStatus.Panels(2).Text = PFantasia
    
    'Definir variável identificando o código do Almoxarifado Central
    XGT_SQL = "SELECT * FROM CentroResultados WHERE empr_cd_Empresa = " & PCodEmpresa & _
              " AND cere_bl_ContrAlmoxarifado = 'S' ORDER BY cere_bl_AlmoxCentral DESC"
    SubQOpenRecordset XLO_ESTOQUE, XGT_SQL, Estatico
    If Not XLO_ESTOQUE.EOF Then XGT_ALMOXARIFADOCENTRAL = FunNulo(XLO_ESTOQUE!cere_cd_pcr)
    XLO_ESTOQUE.Close
    Set XLO_ESTOQUE = Nothing
        
    'Lendo informações do dpto de compras
    With XGO_DPTOCOMPRAS
      .Responsavel = FunNulo(DatEmpresas.Recordset("empr_tx_ComResp").Value)
      .Telefone = FunNulo(DatEmpresas.Recordset("empr_tx_ComTelefone").Value)
      .Fax = FunNulo(DatEmpresas.Recordset("empr_tx_ComFax").Value)
      .Email = FunNulo(DatEmpresas.Recordset("empr_tx_ComEmail").Value)
    End With
    
  '  If XGB_PQUALIEST Then 'Pegar a conta de doação
        XLT_CONTADOACAO = FunNulo(DatEmpresas.Recordset("plco_cd_ReceitaDoacao").Value)
   ' End If
  Else
    MsgBox "Escolha uma empresa!", vbInformation, "Seleção de Empresa"
    Exit Sub
  End If

  Unload Me

  
End Sub

Private Sub CmdLimparHistorico_Click()
  CboEmpresas.BoundText = 0
  CboEmpresas.Text = ""
End Sub



Private Sub Form_Activate()
  Set Formulario = Me
  
  PanEmpresaAtiva.Caption = "EMPRESA ATIVA: " + PFantasia
  CboEmpresas.BoundText = PCodEmpresa
End Sub


Private Sub Form_Load()
  
    Dim XLT_SQL As String
    
    subManutencaoJanelasAtivas "I", "FrmMultiEmpresas"
    
    If pTipoCopia <> "Multi" Then
        'Quando MonoEmpresa entra direto
        XLT_SQL = "SELECT * FROM ConsGENEmpresa"
        subConectarControleDadosNV DatEmpresas, XLT_SQL, Estatico
        CboEmpresas.Text = DatEmpresas.Recordset.Fields("empr_tx_fantasia")
        CboEmpresas.Enabled = False
        Call CmdConfirmar_Click
    Else
        If (UCase(PFantasia) <> "APAE") And (UCase(PFantasia) <> "INSTITUTO ALIANÇA") Then
        
            'Quando MultiEmpresa escolhe a(s) empresa(s) selecionada(s) para o usuário ativo
            XLT_SQL = "SELECT * FROM ConsGENEmpresa WHERE empr_cd_empresa IN (SELECT empr_cd_empresa FROM FuncionarioEmpresa WHERE func_cd_operador = " & func_cd_operador & ") ORDER BY empr_tx_Fantasia"
            subConectarControleDadosNV DatEmpresas, XLT_SQL, Estatico
            
            If Not DatEmpresas.Recordset.EOF Then
                While Not DatEmpresas.Recordset.EOF
                    'Verifica se dentre as empresas associadas ao usuário possui alguma ativa para mostrar no combo
                    If DatEmpresas.Recordset.Fields("empr_tx_atual") = "S" Then
                       CboEmpresas.Text = DatEmpresas.Recordset.Fields("empr_tx_fantasia")
                       DatEmpresas.Recordset.MoveLast
                    End If
                    DatEmpresas.Recordset.MoveNext
                Wend
            Else
                MsgBox "Não possui nenhuma empresa relacionada a este usuário!", vbInformation, "ATENÇÃO"
                End
            End If
        End If
    End If
    
End Sub

VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmManImoveis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Imóveis"
   ClientHeight    =   4275
   ClientLeft      =   30
   ClientTop       =   1335
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraPrincipal 
      Height          =   3690
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7005
      Begin VB.Frame FraDadosGerais 
         Height          =   2355
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   6795
         Begin VB.TextBox TxtInsc 
            Height          =   315
            Left            =   4560
            TabIndex        =   22
            Top             =   1320
            Width           =   2050
         End
         Begin VB.TextBox TxtAndar 
            Height          =   315
            Left            =   4260
            MaxLength       =   3
            TabIndex        =   15
            Top             =   780
            Width           =   390
         End
         Begin VB.TextBox TxtGaragem 
            Height          =   495
            Left            =   1440
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   1755
            Width           =   5175
         End
         Begin VB.TextBox TxtArea 
            Height          =   315
            Left            =   5400
            TabIndex        =   17
            Top             =   780
            Width           =   930
         End
         Begin VB.TextBox TxtVlImovel 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Top             =   300
            Width           =   3195
         End
         Begin MSDataListLib.DataCombo CboTipoImovel 
            Bindings        =   "FrmImoveis.frx":0000
            Height          =   315
            Left            =   1440
            TabIndex        =   13
            Top             =   765
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "tiim_tx_descricao"
            BoundColumn     =   "tiim_cd_TipoImovel"
            Text            =   "CboTipoImovel"
         End
         Begin MSComCtl2.DTPicker DtpEntChaves 
            Height          =   315
            Left            =   1440
            TabIndex        =   20
            Top             =   1275
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   64421891
            CurrentDate     =   37818.8267476852
         End
         Begin VB.Label LblInsc 
            Caption         =   "Insc Imobiliária:"
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
            TabIndex        =   21
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label LblAndar 
            Alignment       =   1  'Right Justify
            Caption         =   "Andar:"
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
            Left            =   3540
            TabIndex        =   14
            Top             =   840
            Width           =   690
         End
         Begin VB.Label LblTipoImovel 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo do Imóvel:"
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
            Left            =   60
            TabIndex        =   12
            Top             =   840
            Width           =   1365
         End
         Begin VB.Label LblEntChaves 
            Alignment       =   1  'Right Justify
            Caption         =   "Ent. Chaves:"
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
            Left            =   240
            TabIndex        =   19
            Top             =   1320
            Width           =   1140
         End
         Begin VB.Label LblGaragem 
            Alignment       =   1  'Right Justify
            Caption         =   "Garagem:"
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
            Left            =   480
            TabIndex        =   23
            Top             =   1875
            Width           =   885
         End
         Begin VB.Label LblArea 
            Alignment       =   1  'Right Justify
            Caption         =   "Área:"
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
            Left            =   4860
            TabIndex        =   16
            Top             =   840
            Width           =   495
         End
         Begin VB.Label LblAvista 
            Caption         =   "Valor à Vista:"
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
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Lblm2 
            Caption         =   "m²"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6360
            TabIndex        =   18
            Top             =   840
            Width           =   255
         End
      End
      Begin VB.Frame FraCodigo 
         Height          =   1035
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6795
         Begin MSMask.MaskEdBox MskEmpreendimento 
            Height          =   315
            Left            =   1920
            TabIndex        =   3
            Top             =   180
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "AAAA"
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboEmpreendimento 
            Bindings        =   "FrmImoveis.frx":001C
            Height          =   315
            Left            =   2760
            TabIndex        =   4
            Top             =   180
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "empd_tx_nome"
            BoundColumn     =   "empd_cd_empreendimento"
            Text            =   "CboEmpreendimento"
         End
         Begin MSMask.MaskEdBox MskCodigo 
            Height          =   315
            Left            =   1920
            TabIndex        =   6
            Top             =   600
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   4
            Mask            =   "AAAA"
            PromptChar      =   " "
         End
         Begin VB.Label LblStatus 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
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
            Left            =   2760
            TabIndex        =   7
            Top             =   660
            Width           =   750
         End
         Begin VB.Label LblDescStatus 
            Caption         =   "Status"
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
            Height          =   255
            Left            =   3660
            TabIndex        =   8
            Top             =   660
            Width           =   1095
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
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label LblCodigo 
            Alignment       =   1  'Right Justify
            Caption         =   "Código do Imóvel:"
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
            Left            =   240
            TabIndex        =   5
            Top             =   660
            Width           =   1605
         End
      End
      Begin MSAdodcLib.Adodc DatEmpreendimento 
         Height          =   330
         Left            =   2280
         Top             =   3360
         Visible         =   0   'False
         Width           =   1980
         _ExtentX        =   3493
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
      Begin MSAdodcLib.Adodc DatTipoImovel 
         Height          =   330
         Left            =   4320
         Top             =   3360
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
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
         Caption         =   "DatTipoImovel"
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
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   4980
      TabIndex        =   27
      Top             =   3780
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
      Left            =   2940
      TabIndex        =   26
      Top             =   3780
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
      Left            =   900
      TabIndex        =   25
      Top             =   3780
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
End
Attribute VB_Name = "FrmManImoveis"
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
            MskCodigo.Text = XLO_IMOVEL!imov_cd_imovel
            MskCodigo.Mask = "AAAA"
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
        TxtInsc.Text = FunNulo(XLO_IMOVEL!imov_nr_insc)
        CboTipoImovel.BoundText = XLO_IMOVEL!tiim_cd_TipoImovel
        
        If Not IsNull(XLO_IMOVEL!imov_dt_EntregaChaves) Then DtpEntChaves.Value = XLO_IMOVEL!imov_dt_EntregaChaves
        
        'Dados da Venda
        TxtVlImovel.Text = Format(XLO_IMOVEL!imov_vl_Avista, "standard")
        
        Chave = XLO_IMOVEL!empd_cd_Empreendimento + "." + XLO_IMOVEL!imov_cd_imovel
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

Private Sub CboEmpreendimento_Change()
    If CboEmpreendimento.Text <> "" Then
        DatEmpreendimento.Recordset.bookmark = CboEmpreendimento.SelectedItem
        MskEmpreendimento.Text = CboEmpreendimento.BoundText
        MskCodigo.Enabled = True
        
        If MskCodigo.Text <> "     - " And MskCodigo.Text <> XFT_CODIMOVEL Then
            Dim XLO_IMOVEL As Object
            SubQOpenRecordset XLO_IMOVEL, "SELECT * FROM Imoveis WHERE imov_cd_imovel = '" & MskCodigo.Text & "' and empd_cd_empreendimento= '" & MskEmpreendimento.Text & "'", Estatico
            If Not XLO_IMOVEL.EOF Then
               MsgBox "Imóvel já cadastrado para este Empreendimento.", vbCritical + vbOKOnly, "ATENÇÃO"
               MskCodigo.SetFocus
               XLO_IMOVEL.Close
               Exit Sub
            End If
            XLO_IMOVEL.Close
        End If
    End If
End Sub

Private Sub CmdDesistir_Click()
    Unload Me
End Sub

Private Sub PreencheInformacoesOriginaisLog()
    
    ReDim XGM_MATRIZLOG(9, 3) As Variant
                        
    XGM_MATRIZLOG(0, 1) = MskEmpreendimento.Text
    XGM_MATRIZLOG(1, 1) = MskCodigo.Text
    XGM_MATRIZLOG(2, 1) = FunNuloVal(FunTrataFloat(TxtArea.Text))
    XGM_MATRIZLOG(3, 1) = TxtGaragem.Text
    XGM_MATRIZLOG(4, 1) = TxtAndar.Text
    XGM_MATRIZLOG(5, 1) = CboTipoImovel.BoundText
    XGM_MATRIZLOG(6, 1) = FunNulo(DtpEntChaves.Value)
    XGM_MATRIZLOG(7, 1) = FunNuloVal(FunTrataFloat(TxtVlImovel.Text))
    XGM_MATRIZLOG(8, 1) = TxtInsc.Text
        
End Sub

Private Sub GravaLog(XLT_OPERACAO As ENUM_OPERACOES, XLT_ORDEM As String)
    Dim XLI_INDICE As Integer
    
    If XLT_OPERACAO <> ALTERAR Then
        ReDim XGM_MATRIZLOG(9, 2) As Variant
        XLI_INDICE = 1
    Else
        XLI_INDICE = 2
    End If

    'Campos
    XGM_MATRIZLOG(0, 0) = "Empreendimento"
    XGM_MATRIZLOG(1, 0) = "Imóvel"
    XGM_MATRIZLOG(2, 0) = "Área"
    XGM_MATRIZLOG(3, 0) = "Garagem"
    XGM_MATRIZLOG(4, 0) = "Andar"
    XGM_MATRIZLOG(5, 0) = "TipoImovel"
    XGM_MATRIZLOG(6, 0) = "Data Entrega Chaves"
    XGM_MATRIZLOG(7, 0) = "Valor a Vista"
    XGM_MATRIZLOG(8, 0) = "Nº Inscrição"
        
    'Dados
    XGM_MATRIZLOG(0, XLI_INDICE) = MskEmpreendimento.Text
    XGM_MATRIZLOG(1, XLI_INDICE) = MskCodigo.Text
    XGM_MATRIZLOG(2, XLI_INDICE) = FunNuloVal(TxtArea.Text)
    XGM_MATRIZLOG(3, XLI_INDICE) = TxtGaragem.Text
    XGM_MATRIZLOG(4, XLI_INDICE) = TxtAndar.Text
    XGM_MATRIZLOG(5, XLI_INDICE) = CboTipoImovel.BoundText
    XGM_MATRIZLOG(6, XLI_INDICE) = FunNulo(DtpEntChaves.Value)
    XGM_MATRIZLOG(7, XLI_INDICE) = FunNuloVal(TxtVlImovel.Text)
    XGM_MATRIZLOG(8, XLI_INDICE) = TxtInsc.Text
    
    Call subRegistraLog("TelaManImoveis", XLT_ORDEM, funCriaDescricaoLog(XGM_MATRIZLOG, XLT_OPERACAO, IMOVEL_CAPI))
End Sub

Private Sub CmdGravar_Click()
      
    Dim XLT_SQL As String
    Dim XLO_RSEmp As ADODB.Recordset  'record set para testar o numero de imoveis do empreendimento
    Dim XLO_RSImoveis As ADODB.Recordset
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
                      " AND empd_cd_empreendimento = '" & MskEmpreendimento.Text & "'"
        
        SubQOpenRecordset XLO_RSImoveis, XLS_SQLTemp, Estatico
        
        
        Conexao.BeginTrans
        
        If XInserir = "I" Or XInserir = "C" Then
            If (Not IsNull(XLO_RSEmp!empd_nr_Unidades)) Then
                If (XLO_RSImoveis!qtdImoveis >= XLO_RSEmp!empd_nr_Unidades) Then
                    MsgBox "Este empreendimento ja possui o numero maximo de imoveis cadastrados", vbCritical, "ATENÇÃO"
                    
                    XLO_RSEmp.Close
                    Set XLO_RSEmp = Nothing
                    XLO_RSImoveis.Close
                    Set XLO_RSImoveis = Nothing
                    
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
                      "imov_dt_EntregaChaves,imov_vl_AVista, imov_tx_status, imov_nr_insc) " & _
                    "VALUES ('" & MskCodigo.Text & "','" & MskEmpreendimento.Text & "'," & PCodEmpresa & ", " & _
                          "'" & FunNuloVal(FunTrataFloat(TxtArea.Text)) & "','" & TxtGaragem.Text & "','" & TxtAndar.Text & "', " & _
                          "" & CboTipoImovel.BoundText & "," & FunNuloData(DtpEntChaves.Value, NomeSgbd) & ", " & _
                          "" & FunNuloVal(FunTrataFloat(TxtVlImovel.Text)) & ",'D', '" & TxtInsc.Text & "')"
                         
            Conexao.Execute XLT_SQL
            
            'Registra Log
            If XInserir = "I" Then
                GravaLog CADASTRAR, "2"
            ElseIf XInserir = "C" Then
                GravaLog COPIAR, "2"
            End If
            
            Chave = MskEmpreendimento.Text + "." + MskCodigo.Text
            Me.Caption = "Alteração de Imóveis"
            
            
            Else    'Alteração
            XLT_SQL = "UPDATE Imoveis SET imov_nr_area='" & FunNuloVal(FunTrataFloat(TxtArea.Text)) & "'," & _
                       "imov_tx_garagem='" & TxtGaragem.Text & "'," & _
                       "imov_nr_andar='" & TxtAndar.Text & "'," & _
                       "tiim_cd_TipoImovel=" & CboTipoImovel.BoundText & "," & _
                       "imov_dt_EntregaChaves=" & FunNuloData(DtpEntChaves.Value, NomeSgbd) & "," & _
                       "imov_vl_AVista=" & FunNuloVal(FunTrataFloat(TxtVlImovel.Text)) & "," & _
                       "imov_nr_insc='" & TxtInsc.Text & "'" & _
                   " WHERE imov_cd_imovel='" & MskCodigo.Text & "' AND empd_cd_empreendimento='" & MskEmpreendimento.Text & "' AND empr_cd_Empresa=" & PCodEmpresa & ""
            Conexao.Execute XLT_SQL
            
            'Registra Log
            GravaLog ALTERAR, "3"
            
            'Atualizando a comissão do corretor
            'Verifica se possui comissões de corretores associadas a este imóvel
            XLS_SQLTemp = "SELECT * " & _
                             "From Comissoes" & _
                             " WHERE empr_cd_empresa = " & PCodEmpresa & _
                             " AND imov_cd_imovel = '" & MskCodigo.Text & "' " & _
                             " AND empd_cd_empreendimento = '" & MskEmpreendimento.Text & "'"
            SubQOpenRecordset XLO_RSImoveis, XLS_SQLTemp, Estatico
            
            'Se forem encontrados comissoes para este imóvel atualizar o valor da comissão de todos os
            'corretores encontrados
            If Not XLO_RSImoveis.EOF Then
                While Not XLO_RSImoveis.EOF
                    XLT_SQL = " UPDATE comissoes SET comi_vl_comissao = (SELECT corr_vl_PercentComissao/100 * " & FunNuloVal(FunTrataFloat(TxtVlImovel.Text)) & " FROM Corretores where corr_cd_corretor = " & _
                                 "" & XLO_RSImoveis!corr_cd_corretor & "" & _
                                 " ) WHERE empr_cd_empresa = " & PCodEmpresa & _
                                 " AND imov_cd_imovel = '" & MskCodigo.Text & "' " & _
                                 " AND empd_cd_empreendimento = '" & MskEmpreendimento.Text & "' " & _
                                 " AND corr_cd_Corretor = " & XLO_RSImoveis!corr_cd_corretor & ""
                    Conexao.Execute XLT_SQL
                    XLO_RSImoveis.MoveNext
                Wend
            End If
            
            XLO_RSImoveis.Close
            Set XLO_RSImoveis = Nothing
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
    
    subConectarControleDadosNV DatEmpreendimento, "SELECT empd_cd_Empreendimento,empd_tx_Nome FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_nome", Estatico
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
        
        'Log
        PreencheInformacoesOriginaisLog
        
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
    subManutencaoJanelasAtivas "R", "FrmManImoveis"
End Sub

Private Sub MskCodigo_GotFocus()
    Call subSelecionaMSK(MskCodigo)
End Sub

Private Sub MskCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
  End If
  SubDesabHabCampos True
End Sub

Private Sub MskCodigo_LostFocus()
    Dim XLO_IMOVEL As New ADODB.Recordset
    Dim XLT_SQL As String
    
    If MskCodigo.Text <> "" Then
      SubPreencheComZero MskCodigo
      SubQOpenRecordset XLO_IMOVEL, "SELECT * FROM Imoveis WHERE empr_cd_empresa = " & PCodEmpresa & " AND imov_cd_Imovel = '" & MskCodigo.Text & "' AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "'", "Estatico"
  
      If Not (XLO_IMOVEL.BOF And XLO_IMOVEL.EOF) Then
          XLO_IMOVEL.Close
          Set XLO_IMOVEL = Nothing
          XFT_CODIMOVEL = MskCodigo.Text
          XFT_CODEMPREENDIMENTO = MskEmpreendimento.Text
          XInserir = "A"
          Me.Caption = "Alteração de Imóveis"
          subPreencheCampos
          MskEmpreendimento.Enabled = False
          CboEmpreendimento.Enabled = False
          MskCodigo.Enabled = False
      End If
    End If
    
End Sub

Private Sub TxtArea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtArea_LostFocus()
    If IsNumeric(TxtArea.Text) Then
        TxtArea.Text = Format$(TxtArea.Text, "##,##0.00")
    Else
        TxtArea.Text = ""
    End If
End Sub

Private Sub TxtVlImovel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtVlImovel_LostFocus()
    If IsNumeric(TxtVlImovel.Text) Then
        TxtVlImovel.Text = Format$(TxtVlImovel.Text, "##,##0.00")
    Else
        TxtVlImovel.Text = ""
    End If
End Sub

Private Sub MskEmpreendimento_GotFocus()
    Call subSelecionaMSK(MskEmpreendimento)
End Sub

Private Sub MskEmpreendimento_LostFocus()
    If MskEmpreendimento.Text <> "" Then
        SubPreencheComZero MskEmpreendimento
        CboEmpreendimento.BoundText = MskEmpreendimento.Text
        If CboEmpreendimento.Text = "" Then
            MsgBox "Não existe Empreendimento com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskEmpreendimento.SetFocus
        End If
    Else
      CboEmpreendimento.BoundText = ""
    End If
End Sub

Private Sub MskEmpreendimento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
    End If
    MskCodigo.Enabled = True
End Sub

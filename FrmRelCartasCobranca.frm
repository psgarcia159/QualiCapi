VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelCartasCobranca 
   Caption         =   "Relatório de Cartas de Cobrança"
   ClientHeight    =   2715
   ClientLeft      =   2265
   ClientTop       =   3240
   ClientWidth     =   9360
   LinkTopic       =   "Form2"
   ScaleHeight     =   2715
   ScaleWidth      =   9360
   Begin VB.CommandButton cmdEtiqueta 
      Caption         =   "&Etiqueta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6300
      TabIndex        =   30
      Top             =   2220
      Visible         =   0   'False
      Width           =   780
   End
   Begin MSDataListLib.DataCombo CboObs 
      Bindings        =   "FrmRelCartasCobranca.frx":0000
      Height          =   315
      Left            =   1380
      TabIndex        =   20
      Top             =   2220
      Visible         =   0   'False
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "obse_tx_observacao"
      BoundColumn     =   "obse_cd_observacao"
      Text            =   "CboObs"
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7200
      TabIndex        =   8
      Top             =   2220
      Width           =   975
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8280
      TabIndex        =   9
      Top             =   2220
      Width           =   1035
   End
   Begin Threed.SSCommand CmdLimparObs 
      Height          =   315
      Left            =   5820
      TabIndex        =   19
      Top             =   2220
      Visible         =   0   'False
      Width           =   360
      _Version        =   65536
      _ExtentX        =   635
      _ExtentY        =   556
      _StockProps     =   78
      MouseIcon       =   "FrmRelCartasCobranca.frx":0015
      Picture         =   "FrmRelCartasCobranca.frx":0467
   End
   Begin MSAdodcLib.Adodc DatObs 
      Height          =   330
      Left            =   2880
      Top             =   3000
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
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4575
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
         Height          =   675
         Left            =   60
         TabIndex        =   14
         Top             =   1320
         Width           =   4455
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
            Left            =   60
            TabIndex        =   4
            Top             =   300
            Width           =   1575
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
            Left            =   1680
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
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
            Left            =   3000
            TabIndex        =   6
            Top             =   240
            Width           =   1395
         End
      End
      Begin VB.Frame FraPeriodo 
         Caption         =   "Período de Vencimento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   60
         TabIndex        =   11
         Top             =   540
         Width           =   4425
         Begin MSComCtl2.DTPicker DtpInicio 
            Height          =   315
            Left            =   660
            TabIndex        =   2
            Top             =   285
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            Format          =   50397187
            CurrentDate     =   37459
         End
         Begin MSComCtl2.DTPicker DtpFim 
            Height          =   315
            Left            =   2760
            TabIndex        =   3
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   50397187
            CurrentDate     =   37614
         End
         Begin VB.Label LblFim 
            Alignment       =   1  'Right Justify
            Caption         =   "Fim:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2220
            TabIndex        =   13
            Top             =   360
            Width           =   495
         End
         Begin VB.Label LblInicio 
            Alignment       =   1  'Right Justify
            Caption         =   "Início:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   495
         End
      End
      Begin MSMask.MaskEdBox MskImovelInicial 
         Height          =   315
         Left            =   1260
         TabIndex        =   0
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "AAAA.AAAA"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MskImovelFinal 
         Height          =   315
         Left            =   2760
         TabIndex        =   1
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "AAAA.AAAA"
         PromptChar      =   " "
      End
      Begin VB.Label LblA 
         Caption         =   "a"
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
         Left            =   2460
         TabIndex        =   16
         Top             =   240
         Width           =   195
      End
      Begin VB.Label LblImovel 
         Alignment       =   1  'Right Justify
         Caption         =   "Imóvel:"
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
         Left            =   540
         TabIndex        =   15
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2115
      Left            =   4620
      TabIndex        =   17
      Top             =   0
      Width           =   4695
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   60
         TabIndex        =   26
         Top             =   540
         Width           =   4515
         Begin Threed.SSCommand CmdLimparTipoPlano 
            Height          =   315
            Left            =   4140
            TabIndex        =   27
            Top             =   240
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmRelCartasCobranca.frx":0579
            Picture         =   "FrmRelCartasCobranca.frx":0595
         End
         Begin MSDataListLib.DataCombo CboTiposPlanos 
            Bindings        =   "FrmRelCartasCobranca.frx":06A7
            Height          =   315
            Left            =   1680
            TabIndex        =   28
            Top             =   240
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "tipl_tx_Descricao"
            BoundColumn     =   "tipl_cd_TipoPlano"
            Text            =   "CboTiposPlanos"
         End
         Begin VB.Label LblTipoPlano 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo do Plano:"
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
            Left            =   360
            TabIndex        =   29
            Top             =   300
            Width           =   1260
         End
      End
      Begin VB.Frame Frame3 
         Height          =   675
         Left            =   60
         TabIndex        =   22
         Top             =   1320
         Width           =   4575
         Begin VB.ComboBox CboNaturezasPlanos 
            Height          =   315
            ItemData        =   "FrmRelCartasCobranca.frx":06C2
            Left            =   1740
            List            =   "FrmRelCartasCobranca.frx":06D5
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   240
            Width           =   2415
         End
         Begin Threed.SSCommand CmdLimparNaturezaPlano 
            Height          =   315
            Left            =   4200
            TabIndex        =   24
            Top             =   240
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "FrmRelCartasCobranca.frx":0707
            Picture         =   "FrmRelCartasCobranca.frx":0723
         End
         Begin VB.Label LblNaturezaPlano 
            Caption         =   "Natureza do Plano:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   60
            TabIndex        =   25
            Top             =   300
            Width           =   1680
         End
      End
      Begin MSAdodcLib.Adodc DatTipoPlano 
         Height          =   330
         Left            =   2160
         Top             =   1080
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
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
         Caption         =   "DatTipoPlano"
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
      Begin MSComCtl2.DTPicker DtpDtBase 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   180
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   50397187
         CurrentDate     =   37180
      End
      Begin VB.Label LblDtBase 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Base:"
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
         TabIndex        =   18
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.Label LblObs 
      Alignment       =   1  'Right Justify
      Caption         =   "Observação:"
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
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "FrmRelCartasCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdEtiqueta_Click()
  
  Dim XLT_CONSULTA, XLT_FILTROPERIODO As String
  Dim XGT_SELECAO As String
  Dim XGT_CONJUNCAO As String
  Dim XLT_NATUREZA As String
  
  XLT_CONSULTA = "consCAPRelCartas2"
  
   
  'Período de vencimento
'    If Not IsNull(DtpFim.Value) Then
'        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
'        "titu_dt_Prorrogacao >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & "" & _
'        " AND titu_dt_Prorrogacao <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ""
'        XGT_CONJUNCAO = " AND "
'    Else
'        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
'        "titu_dt_Prorrogacao >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & ""
'        XGT_CONJUNCAO = " AND "
'    End If
      If (Not IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
          XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
           "titu_dt_Vencimento >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & "" & _
           " AND titu_dt_Vencimento <= " & FunNuloData(dtpFim.Value, NomeSgbd) & ""
           XGT_CONJUNCAO = " AND "
           XLT_FILTROPERIODO = "Vencimento: " & Format$(dtpInicio.Value, "dd/mm/yy") & " a " & Format$(dtpFim.Value, "dd/mm/yy")
       ElseIf (Not IsNull(dtpInicio.Value)) And (IsNull(dtpFim.Value)) Then
           XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
           "titu_dt_Vencimento >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & ""
           XGT_CONJUNCAO = " AND "
           XLT_FILTROPERIODO = "Vencimento: A partir de " & Format$(dtpInicio.Value, "dd/mm/yy")
       ElseIf (IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
           XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
           "(titu_dt_Vencimento <= " & FunNuloData(dtpFim.Value, NomeSgbd) & ")"
           XGT_CONJUNCAO = " AND "
           XLT_FILTROPERIODO = "Vencimento: Até " & Format$(dtpFim.Value, "dd/mm/yy")
       Else
           XLT_FILTROPERIODO = "Vencimento: Geral"
       End If

    'Imóvel
    If MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text <> "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "Imovel >= '" & MskImovelInicial.Text & "'" & _
      " AND Imovel <= '" & MskImovelFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
    ElseIf MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text = "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "Imovel >= '" & MskImovelInicial.Text & "'"
      XGT_CONJUNCAO = " AND "
      ElseIf MskImovelInicial.Text = "    .    " And MskImovelFinal.Text <> "    .    " Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "Imovel <= '" & MskImovelFinal.Text & "'"
        XGT_CONJUNCAO = " AND "
    End If
    
    
    'Tipo do Plano
    If CboTiposPlanos.BoundText <> "" Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO + "tipl_cd_TipoPlano = " & CboTiposPlanos.BoundText & ""
      XGT_CONJUNCAO = " AND "
    End If
    
    
    'Natureza
    If CboNaturezasPlanos.Text <> "" Then
      Select Case CboNaturezasPlanos.ListIndex
          Case 0
               XLT_NATUREZA = "S"
          Case 1
               XLT_NATUREZA = "P"
          Case 2
               XLT_NATUREZA = "F"
          Case 3
               XLT_NATUREZA = "G"
          Case 4
               XLT_NATUREZA = "O"
      End Select
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO + "titu_tx_NaturezaPlano = '" & XLT_NATUREZA & "'"
      XGT_CONJUNCAO = " AND "
    End If
    
    'Titulos Vencidos
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "titu_dt_Pagamento is Null AND " & _
      "titu_dt_Vencimento < " & FunNuloData(dtpFim.Value, NomeSgbd)
      XGT_CONJUNCAO = " AND "
      
    
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "empr_cd_Empresa = " & CStr(PCodEmpresa)
  
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT DISTINCT focl_cd_fornCli, focl_tx_razaoSocial, clie_tx_endcorresp, clie_tx_bairrocorresp , clie_tx_muncorresp, clie_tx_estcorresp, clie_nr_cepcorresp  FROM " & XLT_CONSULTA & " WHERE " & XGT_SELECAO)
    If XGO_RSRELATORIO.EOF Then
      MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
      FrmRelCartasCobranca.MousePointer = vbDefault
      Exit Sub
    End If
    
    
    FrmVisRelatorios.Caption = "Relatório Emissão de Carta de Cobrança"
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelCartasCobranca.rpt")
  
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
    
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelCartasCobranca.MousePointer = vbDefault
End Sub

Private Sub CmdImprimir_Click()
    
    Dim XLT_FILTROPERIODO As String 'Armazenar filtro do período escolhido
    Dim XLT_FILTROIMOVEL As String
    Dim XLT_FILTROSTATUS As String
    Dim XLT_FILTROTIPOPLANO As String
    Dim XLT_FILTRONATUREZAPLANO As String
'    Dim XLT_FILTROEMPREENDIMENTOS As String
'    Dim XLT_FILTROTITULOSPAGOS As String
    Dim XLT_FILTROMULTA As String
    Dim XLT_FILTROJUROS As String
    Dim XLT_FILTRODESAGIO As String
    Dim XLO_TOTALINDEXADOR As Object
    
    Dim XLD_DATABASE As Date
    Dim XLT_STATUS As String
    Dim XLT_STATUS2 As String
    Dim XLT_NATUREZA As String 'Natureza do Plano
    Dim XLT_INDEXADOR As String
    
    Dim XLT_SQL As String
    
    FrmRelComissoes.MousePointer = vbHourglass
    
    XLT_SQL = ""
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
    XLD_DATABASE = DtpDtBase.Value
    
    If NomeSgbd = "Access" Then
  
      XLT_SQL = "" & _
      "SELECT ConsCAPRelPrevisaoRecebimentoPorCliente.*," & _
      "IIF(Indexacao1='M'" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "))" & _
      " as Cotacao1," & _
      "IIF(Indexacao2='M'" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & "))" & _
      "  as Cotacao2,"
      
      XLT_SQL = XLT_SQL & _
      "IIF(Indexacao1='M'" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))" & _
      " as Cotacao1_MesAnterior," & _
      "IIF(Indexacao2='M'" & _
        " ,(select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd('m',-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")))" & _
      " as Cotacao2_MesAnterior" & _
      " From ConsCAPRelPrevisaoRecebimentoPorCliente WHERE "
    Else
    
      XLT_SQL = "" & _
      "SELECT *," & _
      "CASE Indexacao1   when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
        " when 'D' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")" & _
      " end as Cotacao1," & _
      "CASE  Indexacao2 when 'M' then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
        " when 'D'  then" & _
        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")" & _
      " end as Cotacao2,"
      
'      XLT_SQL = XLT_SQL & _
'      "CASE  Indexacao1 when 'M' then" & _
'        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
'      " end as Cotacao1_MesAnterior," & _
'      "CASE Indexacao2 when 'M' then" & _
'        " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
'      " end As Cotacao2_MesAnterior" & _
'      " From ConsCAPRelCartaCobranca WHERE "

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
      " From ConsCAPRelCartaCobranca WHERE "
      
    End If
  
    'Período de vencimento
    If (Not IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "titu_dt_Vencimento >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & "" & _
        " AND titu_dt_Vencimento <= " & FunNuloData(dtpFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Vencimento: " & Format$(dtpInicio.Value, "dd/mm/yy") & " a " & Format$(dtpFim.Value, "dd/mm/yy")
    ElseIf (Not IsNull(dtpInicio.Value)) And (IsNull(dtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "titu_dt_Vencimento >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Vencimento: A partir de " & Format$(dtpInicio.Value, "dd/mm/yy")
    ElseIf (IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "(titu_dt_Vencimento <= " & FunNuloData(dtpFim.Value, NomeSgbd) & ")"
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Vencimento: Até " & Format$(dtpFim.Value, "dd/mm/yy")
    Else
        XLT_FILTROPERIODO = "Vencimento: Geral"
    End If
    
    'Imóvel
    If MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text <> "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "Imovel >= '" & MskImovelInicial.Text & "'" & _
      " AND Imovel <= '" & MskImovelFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROIMOVEL = "Imóvel: " & MskImovelInicial & " a " & MskImovelFinal
    ElseIf MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text = "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "Imovel >= '" & MskImovelInicial.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROIMOVEL = "Imóvel: A partir de " & MskImovelInicial
    ElseIf MskImovelInicial.Text = "    .    " And MskImovelFinal.Text <> "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "Imovel <= '" & MskImovelFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROIMOVEL = "Imóvel: Até " & MskImovelInicial
    Else
      XLT_FILTROIMOVEL = "Imóvel: Geral"
    End If
    
    'Status
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "(cont_dt_Venda <= " & FunNuloData(Format(DtpDtBase.Value, "\01/mm/yy"), NomeSgbd) & _
      " AND (cont_dt_Distrato is null OR " & _
      " (cont_dt_Distrato > " & FunNuloData(Format(DtpDtBase.Value, "\01/mm/yy"), NomeSgbd) & " AND " & _
          " cont_tx_Status<>'C'))) "
    XGT_CONJUNCAO = " AND "
    
    'Tipo do Plano
    If CboTiposPlanos.BoundText <> "" Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO + "tipl_cd_TipoPlano = " & CboTiposPlanos.BoundText & ""
      XGT_CONJUNCAO = " AND "
      XLT_FILTROTIPOPLANO = "Tipo do Plano: " & CboTiposPlanos.Text
    Else
     XLT_FILTROTIPOPLANO = "Tipo do Plano: Geral"
    End If
    
    'Natureza
    If CboNaturezasPlanos.Text <> "" Then
      Select Case CboNaturezasPlanos.ListIndex
          Case 0
               XLT_NATUREZA = "S"
          Case 1
               XLT_NATUREZA = "P"
          Case 2
               XLT_NATUREZA = "F"
          Case 3
               XLT_NATUREZA = "G"
          Case 4
               XLT_NATUREZA = "O"
      End Select
     
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO + "titu_tx_NaturezaPlano = '" & XLT_NATUREZA & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTRONATUREZAPLANO = "Natureza do Plano: " & CboNaturezasPlanos.Text
    Else
      XLT_FILTRONATUREZAPLANO = "Natureza do Plano: Geral"
    End If
    
    'Indexador
    If OptValorHistorico.Value = True Then
      XLT_INDEXADOR = 0
    ElseIf OptIndexador1.Value = True Then
      XLT_INDEXADOR = 1
    ElseIf OptIndexador2.Value = True Then
      XLT_INDEXADOR = 2
    End If

    'Titulos Vencidos
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    "titu_dt_Pagamento is Null AND " & _
    "titu_dt_Vencimento < " & FunNuloData(dtpFim.Value, NomeSgbd)
    XGT_CONJUNCAO = " AND "
    
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "empr_cd_Empresa = " & CStr(PCodEmpresa)
    
    'Ordenanação
    XGT_SELECAO = XGT_SELECAO & " ORDER BY titu_dt_Vencimento"
  
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL & XGT_SELECAO)
    If XGO_RSRELATORIO.EOF Then
      MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
      FrmRelComissoes.MousePointer = vbDefault
      Exit Sub
    End If
    
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
    FrmVisRelatorios.Caption = "Relatório de Emissão de Cartas de Cobrança"
    'Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelCartasCobranca.rpt")
    Set XGR_RELATORIO = RptCartasCobranca
    XGR_RELATORIO.DiscardSavedData
  
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("FiltroDiaBase").Text = Chr(34) & Day(XLD_DATABASE) & Chr(34)
        .GetItemByName("FiltroMesBase").Text = Chr(34) & Format(XLD_DATABASE, "mmmm") & Chr(34)
        .GetItemByName("FiltroAnoBase").Text = Chr(34) & Format(XLD_DATABASE, "yyyy") & Chr(34)
        .GetItemByName("FiltroPeriodo").Text = Chr(34) & XLT_FILTROPERIODO & Chr(34)
        .GetItemByName("FiltroImovel").Text = Chr(34) & XLT_FILTROIMOVEL & Chr(34)
        .GetItemByName("FiltroStatus").Text = Chr(34) & XLT_FILTROSTATUS & Chr(34)
        .GetItemByName("FiltroTipoPlano").Text = Chr(34) & XLT_FILTROTIPOPLANO & Chr(34)
        .GetItemByName("FiltroNaturezaPlano").Text = Chr(34) & XLT_FILTRONATUREZAPLANO & Chr(34)
        .GetItemByName("FiltroDataBase").Text = Chr(34) & Format(DtpDtBase.Value, "dd/mm/yy") & Chr(34)
        .GetItemByName("Indexador").Text = Chr(34) & XLT_INDEXADOR & Chr(34)
        .GetItemByName("FiltroMulta").Text = Chr(34) & XLT_FILTROMULTA & Chr(34)
        .GetItemByName("FiltroJuros").Text = Chr(34) & XLT_FILTROJUROS & Chr(34)
        .GetItemByName("FiltroDesagio").Text = Chr(34) & XLT_FILTRODESAGIO & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Rodape"
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelComissoes.MousePointer = vbDefault
End Sub

Private Sub CmdLimparNaturezaPlano_Click()
  
  CboNaturezasPlanos.ListIndex = -1
  
End Sub

Private Sub CmdLimparTipoPlano_Click()
  
  CboTiposPlanos.BoundText = ""
  
End Sub

Private Sub Etiqueta_Click()
    Dim XLT_FILTROPERIODO As String 'Armazenar filtro do período escolhido
    Dim XLT_FILTROIMOVEL As String
    Dim XLT_FILTROSTATUS As String
    Dim XLT_FILTROTIPOPLANO As String
    Dim XLT_FILTRONATUREZAPLANO As String
'    Dim XLT_FILTROEMPREENDIMENTOS As String
'    Dim XLT_FILTROTITULOSPAGOS As String
    Dim XLT_FILTROMULTA As String
    Dim XLT_FILTROJUROS As String
    Dim XLT_FILTRODESAGIO As String
    Dim XLO_TOTALINDEXADOR As Object
    
    Dim XLT_STATUS As String
    Dim XLT_STATUS2 As String
    Dim XLT_NATUREZA As String 'Natureza do Plano
    Dim XLT_INDEXADOR As String
    
    Dim XLT_SQL As String
    
    FrmRelComissoes.MousePointer = vbHourglass
    
    XLT_SQL = ""
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
    
    XLT_SQL = "" & _
    "SELECT ConsCAPRelPrevisaoRecebimentoPorCliente.*," & _
    "CASE Indexacao1   when 'M' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
      " when 'D' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")" & _
    " end as Cotacao1," & _
    "CASE  Indexacao2 when 'M' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ")" & _
      " when 'D'  then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=" & FunNuloData(DtpDtBase.Value, NomeSgbd) & ")" & _
    " end as Cotacao2,"
    
    XLT_SQL = XLT_SQL & _
    "CASE  Indexacao1 when 'M' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
    " end as Cotacao1_MesAnterior," & _
    "CASE Indexacao2 when 'M' then" & _
      " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & "))" & _
    " end As Cotacao2_MesAnterior" & _
    " From ConsCAPRelPrevisaoRecebimentoPorCliente WHERE "
  
  
    'Período de vencimento
    If Not IsNull(dtpFim.Value) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelPrevisaoRecebimentoPorCliente.titu_dt_Vencimento >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & "" & _
        " AND ConsCAPRelPrevisaoRecebimentoPorCliente.titu_dt_Vencimento <= " & FunNuloData(dtpFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Vencimento: " & Format$(dtpInicio.Value, "dd/mm/yy") & " a " & Format$(dtpFim.Value, "dd/mm/yy")
    Else
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelPrevisaoRecebimentoPorCliente.titu_dt_Vencimento >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Vencimento: A partir de " & Format$(dtpInicio.Value, "dd/mm/yy")
    End If
    
    'Imóvel
    If MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text <> "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelPrevisaoRecebimentoPorCliente.Imovel >= '" & MskImovelInicial.Text & "'" & _
      " AND ConsCAPRelPrevisaoRecebimentoPorCliente.Imovel <= '" & MskImovelFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROIMOVEL = "Imóvel: " & MskImovelInicial & " a " & MskImovelFinal
    ElseIf MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text = "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelPrevisaoRecebimentoPorCliente.Imovel >= '" & MskImovelInicial.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROIMOVEL = "Imóvel: A partir de " & MskImovelInicial
    ElseIf MskImovelInicial.Text = "    .    " And MskImovelFinal.Text <> "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelPrevisaoRecebimentoPorCliente.Imovel <= '" & MskImovelFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROIMOVEL = "Imóvel: Até " & MskImovelInicial
    Else
      XLT_FILTROIMOVEL = "Imóvel: Geral"
    End If
    
    
    'Tipo do Plano
    If CboTiposPlanos.BoundText <> "" Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO + "ConsCAPRelPrevisaoRecebimentoPorCliente.tipl_cd_TipoPlano = " & CboTiposPlanos.BoundText & ""
      XGT_CONJUNCAO = " AND "
      XLT_FILTROTIPOPLANO = "Tipo do Plano: " & CboTiposPlanos.Text
    Else
     XLT_FILTROTIPOPLANO = "Tipo do Plano: Geral"
    End If
    
    
    'Natureza
    If CboNaturezasPlanos.Text <> "" Then
      Select Case CboNaturezasPlanos.ListIndex
          Case 0
               XLT_NATUREZA = "S"
          Case 1
               XLT_NATUREZA = "P"
          Case 2
               XLT_NATUREZA = "F"
          Case 3
               XLT_NATUREZA = "G"
          Case 4
               XLT_NATUREZA = "O"
      End Select
     
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO + "ConsCAPRelPrevisaoRecebimentoPorCliente.titu_tx_NaturezaPlano = '" & XLT_NATUREZA & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTRONATUREZAPLANO = "Natureza do Plano: " & CboNaturezasPlanos.Text
    Else
      XLT_FILTRONATUREZAPLANO = "Natureza do Plano: Geral"
    End If
    
    'Indexador
    If OptValorHistorico.Value = True Then
      XLT_INDEXADOR = 0
    ElseIf OptIndexador1.Value = True Then
      XLT_INDEXADOR = 1
    ElseIf OptIndexador2.Value = True Then
      XLT_INDEXADOR = 2
    End If
    

    'Titulos Vencidos
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelPrevisaoRecebimentoPorCliente.titu_dt_Pagamento is Null AND " & _
      "titu_dt_Vencimento < " & FunNuloData(dtpFim.Value, NomeSgbd)
      XGT_CONJUNCAO = " AND "
      
    
'    'Multa
'    If ChkMulta.Value = 1 Then
'      XLT_FILTROMULTA = "1"
'    Else
'      XLT_FILTROMULTA = "0"
'    End If
'
'    'Juros
'    If ChkJuros.Value = 1 Then
'      XLT_FILTROJUROS = "1"
'    Else
'      XLT_FILTROJUROS = "0"
'    End If
    
'    'Deságio
'    If ChkDesagio.Value = 1 Then
'      XLT_FILTRODESAGIO = "1"
'    Else
'      XLT_FILTRODESAGIO = "0"
'    End If
    
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelPrevisaoRecebimentoPorCliente.empr_cd_Empresa = " & CStr(PCodEmpresa)
  
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL & XGT_SELECAO)
    If XGO_RSRELATORIO.EOF Then
      MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
      FrmRelComissoes.MousePointer = vbDefault
      Exit Sub
    End If
    
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
    FrmVisRelatorios.Caption = "Relatório Emissão de Carta de Cobrança"
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\Etiquetas.rpt")
  
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("FiltroPeriodo").Text = Chr(34) & XLT_FILTROPERIODO & Chr(34)
        .GetItemByName("FiltroImovel").Text = Chr(34) & XLT_FILTROIMOVEL & Chr(34)
        .GetItemByName("FiltroStatus").Text = Chr(34) & XLT_FILTROSTATUS & Chr(34)
        .GetItemByName("FiltroTipoPlano").Text = Chr(34) & XLT_FILTROTIPOPLANO & Chr(34)
        .GetItemByName("FiltroNaturezaPlano").Text = Chr(34) & XLT_FILTRONATUREZAPLANO & Chr(34)
        .GetItemByName("FiltroDataBase").Text = Chr(34) & Format(DtpDtBase.Value, "dd/mm/yy") & Chr(34)
        .GetItemByName("Indexador").Text = Chr(34) & XLT_INDEXADOR & Chr(34)
        .GetItemByName("FiltroMulta").Text = Chr(34) & XLT_FILTROMULTA & Chr(34)
        .GetItemByName("FiltroJuros").Text = Chr(34) & XLT_FILTROJUROS & Chr(34)
        '.GetItemByName("FiltroTitulosPagos").Text = Chr(34) & XLT_FILTROTITULOSPAGOS & Chr(34)
        .GetItemByName("FiltroDesagio").Text = Chr(34) & XLT_FILTRODESAGIO & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelComissoes.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelCartasCobranca
  subConectarControleDadosNV DatObs, "SELECT * FROM Observacoes ORDER BY obse_tx_observacao", Estatico
End Sub

Private Sub Form_Load()
  subManutencaoJanelasAtivas "I", "FrmRelCartasCobranca"
  subConectarControleDadosNV DatTipoPlano, "SELECT * FROM TiposPlanos ORDER BY tipl_tx_Descricao", Estatico
  subConectarControleDadosNV DatObs, "SELECT * FROM Observacoes ORDER BY obse_tx_observacao", Estatico
    
  DtpDtBase.Value = Date
  dtpInicio.Value = Date
  dtpFim.Value = Date
  
End Sub

Private Sub MskImovelInicial_GotFocus()
  
  Call subSelecionaMSK(MskImovelInicial)
  
End Sub

Private Sub MskImovelFinal_GotFocus()
  
  Call subSelecionaMSK(MskImovelFinal)
  
End Sub

Private Sub MskImovelInicial_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub MskImovelFinal_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub MskImovelInicial_LostFocus()
   
  If MskImovelInicial.Text <> "    .    " Then
    If Len(Trim(MskImovelInicial.Text)) < 9 Then
      MsgBox "O código inicial do imóvel deve possuir 8 dígitos, ou deve estar vazio.", vbCritical, "ATENÇÃO !"
      MskImovelInicial.SetFocus
    End If
  End If
  
End Sub

Private Sub MskImovelFinal_LostFocus()
   
  If MskImovelFinal.Text <> "    .    " Then
    If Len(Trim(MskImovelFinal.Text)) < 9 Then
      MsgBox "O código final do imóvel deve possuir 8 dígitos, ou deve estar vazio", vbCritical, "ATENÇÃO !"
      MskImovelFinal.SetFocus
    End If
  End If
  
End Sub

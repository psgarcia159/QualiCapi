VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelRepasse 
   Caption         =   "Relatório de Repasse Contratual"
   ClientHeight    =   3525
   ClientLeft      =   2235
   ClientTop       =   2520
   ClientWidth     =   9360
   Icon            =   "frmRelRepasse.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3525
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3075
      Left            =   4620
      TabIndex        =   23
      Top             =   -60
      Width           =   4695
      Begin VB.Frame FraTipoPlano 
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
         Left            =   60
         TabIndex        =   30
         Top             =   840
         Width           =   4575
         Begin Threed.SSCommand CmdLimparTipoPlano 
            Height          =   315
            Left            =   4200
            TabIndex        =   11
            Top             =   240
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "frmRelRepasse.frx":2AFA
            Picture         =   "frmRelRepasse.frx":2B16
         End
         Begin MSDataListLib.DataCombo CboTiposPlanos 
            Bindings        =   "frmRelRepasse.frx":2C28
            Height          =   315
            Left            =   1740
            TabIndex        =   10
            Top             =   240
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "tipl_tx_Descricao"
            BoundColumn     =   "tipl_cd_TipoPlano"
            Text            =   "CboTiposPlanos"
         End
         Begin MSAdodcLib.Adodc DatTipoPlano 
            Height          =   330
            Left            =   2100
            Top             =   240
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
            Left            =   420
            TabIndex        =   31
            Top             =   300
            Width           =   1260
         End
      End
      Begin VB.Frame FraNaturezaPlano 
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
         TabIndex        =   28
         Top             =   1620
         Width           =   4575
         Begin VB.ComboBox CboNaturezasPlanos 
            Height          =   315
            ItemData        =   "frmRelRepasse.frx":2C43
            Left            =   1740
            List            =   "frmRelRepasse.frx":2C59
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   180
            Width           =   2415
         End
         Begin Threed.SSCommand CmdLimparNaturezaPlano 
            Height          =   315
            Left            =   4185
            TabIndex        =   13
            Top             =   180
            Width           =   315
            _Version        =   65536
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "frmRelRepasse.frx":2C94
            Picture         =   "frmRelRepasse.frx":2CB0
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
            TabIndex        =   29
            Top             =   240
            Width           =   1680
         End
      End
      Begin VB.Frame FraRepassados 
         Caption         =   "Titulos Repassados ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   60
         TabIndex        =   27
         Top             =   2340
         Width           =   4575
         Begin VB.OptionButton OptSim 
            Caption         =   "Sim"
            Height          =   255
            Left            =   480
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton OptNao 
            Caption         =   "Não"
            Height          =   255
            Left            =   1740
            TabIndex        =   15
            Top             =   240
            Width           =   675
         End
         Begin VB.OptionButton OptTodos 
            Caption         =   "Todos"
            Height          =   255
            Left            =   3000
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin Threed.SSFrame FraPercentual 
         Height          =   660
         Left            =   60
         TabIndex        =   34
         Top             =   120
         Width           =   4545
         _Version        =   65536
         _ExtentX        =   8017
         _ExtentY        =   1164
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtPercentual 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2640
            TabIndex        =   9
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   3720
            TabIndex        =   37
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Percentual de Repasse:"
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
            Left            =   480
            TabIndex        =   36
            Top             =   300
            Width           =   2100
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   0
      TabIndex        =   20
      Top             =   -60
      Width           =   4575
      Begin VB.Frame FraPagamento 
         Caption         =   "Período de Pagamento"
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
         TabIndex        =   24
         Top             =   840
         Width           =   4470
         Begin MSComCtl2.DTPicker DtpInicioPagamento 
            Height          =   315
            Left            =   675
            TabIndex        =   2
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   49348611
            CurrentDate     =   37614.4138194444
         End
         Begin MSComCtl2.DTPicker DtpFimPagamento 
            Height          =   315
            Left            =   2760
            TabIndex        =   3
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   49348611
            CurrentDate     =   37614.4135763889
         End
         Begin VB.Label Label1 
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
            TabIndex        =   26
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label2 
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
            TabIndex        =   25
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame FraCorrecao 
         Caption         =   "Indexador para Repasse :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   60
         TabIndex        =   21
         Top             =   2340
         Width           =   4455
         Begin MSDataListLib.DataCombo CboIndexador 
            Bindings        =   "frmRelRepasse.frx":2DC2
            Height          =   315
            Left            =   1200
            TabIndex        =   7
            Top             =   240
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "moed_tx_descricao"
            BoundColumn     =   "moed_cd_codmoeda"
            Text            =   "CboIndexador"
         End
         Begin Threed.SSCommand CmdLimparIndexador 
            Height          =   315
            Left            =   2670
            TabIndex        =   8
            Top             =   240
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   556
            _StockProps     =   78
            MouseIcon       =   "frmRelRepasse.frx":2DDD
            Picture         =   "frmRelRepasse.frx":322F
         End
         Begin MSAdodcLib.Adodc DatIndexador 
            Height          =   330
            Left            =   1080
            Top             =   240
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
            Caption         =   "DatIndexador"
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
      Begin Threed.SSFrame FraStatus 
         Height          =   660
         Left            =   60
         TabIndex        =   22
         Top             =   1620
         Width           =   4425
         _Version        =   65536
         _ExtentX        =   7805
         _ExtentY        =   1164
         _StockProps     =   14
         Caption         =   "Status do Contrato"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox ChkInativo 
            Caption         =   "Inativo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1740
            TabIndex        =   5
            Top             =   300
            Width           =   1080
         End
         Begin VB.CheckBox ChkAtivo 
            Caption         =   "Ativo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   360
            TabIndex        =   4
            Top             =   300
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.CheckBox ChkCedido 
            Caption         =   "Cedido"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3000
            TabIndex        =   6
            Top             =   300
            Width           =   1230
         End
      End
      Begin Threed.SSFrame FraEmpreendimento 
         Height          =   660
         Left            =   60
         TabIndex        =   32
         Top             =   120
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   1164
         _StockProps     =   14
         Caption         =   "Imóvel: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSMask.MaskEdBox MskImovelInicial 
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   240
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
            Left            =   1500
            TabIndex        =   1
            Top             =   240
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
            Left            =   1260
            TabIndex        =   33
            Top             =   300
            Width           =   195
         End
      End
      Begin Threed.SSFrame FraDtbase 
         Height          =   660
         Left            =   2760
         TabIndex        =   35
         Top             =   120
         Width           =   1725
         _Version        =   65536
         _ExtentX        =   3043
         _ExtentY        =   1164
         _StockProps     =   14
         Caption         =   "Data Base:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSComCtl2.DTPicker DtpDtBase 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   49348611
            CurrentDate     =   37180
         End
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   345
      Left            =   6960
      TabIndex        =   18
      Top             =   3135
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   8220
      TabIndex        =   19
      Top             =   3120
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc DatObs 
      Height          =   330
      Left            =   3780
      Top             =   4800
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
Attribute VB_Name = "FrmRelRepasse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboIndexador_Change()
  If CboIndexador.BoundText <> "" Then
    DatIndexador.Recordset.bookmark = CboIndexador.SelectedItem
  End If
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    
    Dim XLT_FILTROPERIODO As String 'Armazenar filtro do período escolhido
    Dim XLT_FILTROCONTRATO As String
    Dim XLT_FILTROSTATUS As String
    Dim XLT_FILTROTIPOPLANO As String
    Dim XLT_FILTRONATUREZAPLANO As String
    Dim XLT_FILTROIMOVEL As String
    Dim XLT_FILTROPERIODOPAGAMENTO As String
    Dim XLT_FILTROREPASSADOS As String
    
    Dim XLT_STATUS As String
    Dim XLT_STATUS2 As String
    Dim XLT_NATUREZA As String 'Natureza do Plano
    'Dim XLT_INDEXADOR As String
    'Dim XLT_ORDEM As String
    'Dim XLO_RSCLIENTESIMOVEIS As New ADODB.Recordset
    Dim XLT_SQL As String
    
    Dim XLO_RSMOEDA As New ADODB.Recordset
    Dim XLT_CODMOEDA As String
    Dim XLT_MOEDA As String
    
    FrmRelRepasse.MousePointer = vbHourglass
    
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
    
    '--------- VALIDACAO DOS CAMPOS OBRIGATRIOS -------------------
    If (FunObrigatorioTXT(txtPercentual, "Informe o Percentual a ser repassado")) Then
      FrmRelRepasse.MousePointer = vbDefault
      Exit Sub
    End If
    'If (FunObrigatorioCBO(CboIndexador, "Um Indexador deve ser escolhido")) Then
    '  FrmRelRepasse.MousePointer = vbDefault
    '  Exit Sub
    'End If
    
    
    XLT_CODMOEDA = DatIndexador.Recordset.Fields("moed_cd_CODmoeda")
   
    If CboIndexador.BoundText <> "" Then
     XLT_MOEDA = DatIndexador.Recordset!moed_tx_descricao
    Else
      XLT_MOEDA = "R$"
    End If
    
    If NomeSgbd = "Access" Then
  
      If DatIndexador.Recordset.Fields("moed_tx_indexacao") = "M" Then 'pega mensal
        
        XLT_SQL = "SELECT *," & _
                  "(SELECT cotacoesmoedas.cota_vl_Valor FROM cotacoesmoedas WHERE " & _
                  "moed_cd_Moeda1=" & XLT_CODMOEDA & " AND cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ") as Cotacao1" & _
                  " ,(SELECT cotacoesmoedas.cota_vl_Valor FROM cotacoesmoedas WHERE " & _
                  "moed_cd_Moeda3=" & XLT_CODMOEDA & "  AND cota_dt_datacotacao=format(titu_dt_BasePagto,'\01/mm/yy')) as Cotacao2 " & _
                  " FROM consCAPRelextratoContrato WHERE "
      Else
        XLT_SQL = " SELECT * , " & _
                 "(SELECT cotacoesmoedas.cota_vl_Valor FROM cotacoesmoedas WHERE " & _
                 " moed_cd_CodMoeda = '" & XLT_CODMOEDA & _
                 "' AND cota_dt_datacotacao = " & FunNuloData(DtpDtBase.Value, NomeSgbd) & ") AS cotacao1, " & _
                 "(SELECT cotacoesmoedas.cota_vl_Valor FROM cotacoesmoedas WHERE " & _
                 " moed_cd_CodMoeda = '" & XLT_CODMOEDA & _
                 "' AND cota_dt_datacotacao=titu_dt_BasePagto) AS cotacao2 " & _
                 " FROM consCAPRelextratoContrato WHERE "
    
      End If
    
    Else
      
      If CboIndexador.BoundText <> "" Then
    
        If DatIndexador.Recordset.Fields("moed_tx_indexacao") = "M" Then 'pega mensal
          XLT_SQL = "SELECT consCAPRelextratoContrato.* ," & _
                    " ( SELECT cota_vl_Valor FROM " & _
                    " cotacoesmoedas WHERE moed_cd_CODMoeda = '" & XLT_CODMOEDA & "' AND cota_dt_datacotacao = " & FunNuloData("01/" & Month(DtpDtBase.Value) & "/" & Year(DtpDtBase.Value), NomeSgbd) & ") AS Cotacao1, " & _
                    "(SELECT cota_vl_Valor FROM cotacoesmoedas WHERE moed_cd_CODMoeda = '" & XLT_CODMOEDA & "' AND cota_dt_datacotacao=convert(varchar,month(titu_dt_BasePagto))+'/01/' + convert(varchar,year(titu_dt_BasePagto))) AS cotacao2 " & _
                    " FROM consCAPRelextratoContrato WHERE "
    
        Else
          XLT_SQL = " SELECT * , " & _
                   "(SELECT cotacoesmoedas.cota_vl_Valor FROM cotacoesmoedas WHERE " & _
                   " moed_cd_CODMoeda = '" & XLT_CODMOEDA & _
                   "' AND cota_dt_datacotacao = " & FunNuloData(DtpDtBase.Value, NomeSgbd) & ") AS cotacao1, " & _
                   "(SELECT cotacoesmoedas.cota_vl_Valor FROM cotacoesmoedas WHERE " & _
                   " moed_cd_CodMoeda = '" & XLT_CODMOEDA & _
                   "' AND cota_dt_datacotacao=titu_dt_BasePagto) AS cotacao2 " & _
                   " FROM consCAPRelextratoContrato WHERE "
        End If
      Else
      
        XLT_SQL = "SELECT consCAPRelextratoContrato.* ," & _
            " 0 AS Cotacao1, " & _
            " 0 AS cotacao2 " & _
            " FROM consCAPRelextratoContrato WHERE "
        
      End If
        
    End If
    
    'Imóvel
    If MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text <> "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "consCAPRelextratoContrato.Imovel >= '" & MskImovelInicial.Text & "'" & _
      " AND consCAPRelextratoContrato.Imovel <= '" & MskImovelFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROIMOVEL = "Imóvel: " & MskImovelInicial & " a " & MskImovelFinal
    ElseIf MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text = "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "consCAPRelextratoContrato.Imovel >= '" & MskImovelInicial.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROIMOVEL = "Imóvel: A partir de " & MskImovelInicial
    ElseIf MskImovelInicial.Text = "    .    " And MskImovelFinal.Text <> "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "consCAPRelextratoContrato.Imovel <= '" & MskImovelFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROIMOVEL = "Imóvel: Até " & MskImovelInicial
    Else
      XLT_FILTROIMOVEL = "Imóvel: Geral"
    End If
    
'    'Periodo de Pagamento
'    If (Not IsNull(DtpInicioPagamento.Value)) And (Not IsNull(DtpFimPagamento.Value)) Then
'       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
'        "(titu_dt_Pagamento is null OR (ConsCAPRelExtratoContrato.titu_dt_Pagamento >= " & FunNuloData(DtpInicioPagamento.Value, NomeSgbd) & "" & _
'        " AND consCAPRelextratoContrato.titu_dt_Pagamento <= " & FunNuloData(DtpFimPagamento.Value, NomeSgbd) & "))"
'        XGT_CONJUNCAO = " AND "
'        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: " & Format$(DtpInicioPagamento.Value, "dd/mm/yy") & " a " & Format$(DtpFimPagamento.Value, "dd/mm/yy")
'    ElseIf (Not IsNull(DtpInicioPagamento.Value)) And (IsNull(DtpFimPagamento.Value)) Then
'        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
'        "(titu_dt_Pagamento is null OR consCAPRelextratoContrato.titu_dt_Pagamento >= " & FunNuloData(DtpInicioPagamento.Value, NomeSgbd) & ")"
'        XGT_CONJUNCAO = " AND "
'        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: A partir de " & Format$(DtpInicioPagamento.Value, "dd/mm/yy")
'    ElseIf (IsNull(DtpInicioPagamento.Value)) And (Not IsNull(DtpFimPagamento.Value)) Then
'        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
'        "(titu_dt_Pagamento is null OR consCAPRelextratoContrato.titu_dt_Pagamento <= '" & Format(DtpFimPagamento.Value, "mm/dd/yyyy") & "')"
'        XGT_CONJUNCAO = " AND "
'        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: Até " & Format$(DtpFimPagamento.Value, "dd/mm/yy")
'    Else
'        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: Geral"
'    End If

    'Periodo de Pagamento
    If (Not IsNull(DtpInicioPagamento.Value)) And (Not IsNull(DtpFimPagamento.Value)) Then
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        " ConsCAPRelExtratoContrato.titu_dt_Pagamento >= " & FunNuloData(DtpInicioPagamento.Value, NomeSgbd) & "" & _
        " AND consCAPRelextratoContrato.titu_dt_Pagamento <= " & FunNuloData(DtpFimPagamento.Value, NomeSgbd)
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: " & Format$(DtpInicioPagamento.Value, "dd/mm/yy") & " a " & Format$(DtpFimPagamento.Value, "dd/mm/yy")
    ElseIf (Not IsNull(DtpInicioPagamento.Value)) And (IsNull(DtpFimPagamento.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "consCAPRelextratoContrato.titu_dt_Pagamento >= " & FunNuloData(DtpInicioPagamento.Value, NomeSgbd)
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: A partir de " & Format$(DtpInicioPagamento.Value, "dd/mm/yy")
    ElseIf (IsNull(DtpInicioPagamento.Value)) And (Not IsNull(DtpFimPagamento.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "consCAPRelextratoContrato.titu_dt_Pagamento <= '" & Format(DtpFimPagamento.Value, "mm/dd/yyyy")
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: Até " & Format$(DtpFimPagamento.Value, "dd/mm/yy")
    Else
        XLT_FILTROPERIODOPAGAMENTO = "Pagamento: Geral"
    End If
    
    'Status
    If ChkAtivo.Value = ChkInativo.Value And ChkAtivo.Value = ChkCedido.Value Then
        XLT_FILTROSTATUS = "Status: Geral"
        XLT_STATUS = "T" 'Todos
    Else
        If ChkAtivo.Value = 1 Then
            If ChkInativo.Value = 1 Then  'Ativo e Inativo
                XLT_STATUS = "A"
                XLT_STATUS2 = "I"
                XLT_FILTROSTATUS = "Status: Ativo e Inativo"
            ElseIf ChkCedido.Value = 1 Then  'Ativo e Reservado
                XLT_STATUS = "A"
                XLT_STATUS2 = "C"
                XLT_FILTROSTATUS = "Status: Ativo e Cedido"
            Else
                XLT_STATUS = "A" 'Ativo
                XLT_FILTROSTATUS = "Status: Ativo"
            End If
        ElseIf ChkInativo.Value = 1 Then
            If ChkCedido.Value = 1 Then   'Inativo e Cedido
                XLT_STATUS = "I"
                XLT_STATUS2 = "C"
                XLT_FILTROSTATUS = "Status: Inativo e Cedido"
            Else
                XLT_STATUS = "I" 'Inativo
                XLT_FILTROSTATUS = "Status: Inativo"
            End If
        ElseIf ChkCedido.Value = 1 Then
            XLT_STATUS = "C" 'Cedido
            XLT_FILTROSTATUS = "Status: Cedido"
        End If
        If XLT_STATUS2 <> "" Then 'Foi escolhido dois status
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "(consCAPRelextratoContrato.cont_tx_Status = '" & XLT_STATUS & "'" & _
            " OR consCAPRelextratoContrato.cont_tx_Status = '" & XLT_STATUS2 & "')"
            XGT_CONJUNCAO = " AND "
        Else
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "consCAPRelextratoContrato.cont_tx_Status = '" & XLT_STATUS & "'"
            XGT_CONJUNCAO = " AND "
        End If
    End If
    
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
    
    'Indexador ---------------- **************** _____________________
   
'    If OptValorHistorico.Value = True Then
'      XLT_INDEXADOR = 0
'    ElseIf OptIndexador1.Value = True Then
'      XLT_INDEXADOR = 1
'    ElseIf OptIndexador2.Value = True Then
'      XLT_INDEXADOR = 2
'    End If

    'Repassados
    If OptSim.Value = True Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "titu_dt_repasse is not null"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROREPASSADOS = "Títulos Repassados"
    ElseIf OptNao.Value = True Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "titu_dt_repasse is null"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROREPASSADOS = "Títulos não Repassados"
    Else
      XLT_FILTROREPASSADOS = "Títulos Repassados e não Repassados"
    End If
    
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "empr_cd_Empresa = " & CStr(PCodEmpresa)
  
    'Fecha o recordset se estiver aberto
    If XGO_RSRELATORIO.State = adStateOpen Then
      XGO_RSRELATORIO.Close
      Set XGO_RSRELATORIO = Nothing
    End If
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL & XGT_SELECAO)
    
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Fecha o recordset se estiver aberto
    If XGO_RSCABECALHO.State = adStateOpen Then
      XGO_RSCABECALHO.Close
      Set XGO_RSCABECALHO = Nothing
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
     
    FrmVisRelatorios.Caption = "Relatório de Repasse Contratual"
    Set XGR_RELATORIO = Nothing
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelRepasse.rpt")
    
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("FiltroImovel").Text = Chr(34) & XLT_FILTROIMOVEL & Chr(34)
        .GetItemByName("PERCENTUAL").Text = Chr(34) & txtPercentual.Text & Chr(34)
        .GetItemByName("INDEXADOR").Text = Chr(34) & XLT_MOEDA & Chr(34)
        .GetItemByName("FiltroDataBase").Text = Chr(34) & Format(DtpDtBase.Value, "dd/mm/yy") & Chr(34)
        .GetItemByName("FiltroPeriodoPagamento").Text = Chr(34) & XLT_FILTROPERIODOPAGAMENTO & Chr(34)
        .GetItemByName("FiltroStatus").Text = Chr(34) & XLT_FILTROSTATUS & Chr(34)
        .GetItemByName("FiltroTipoPlano").Text = Chr(34) & XLT_FILTROTIPOPLANO & Chr(34)
        .GetItemByName("FiltroNaturezaPlano").Text = Chr(34) & XLT_FILTRONATUREZAPLANO & Chr(34)
        .GetItemByName("FiltroTitulosRepassados").Text = Chr(34) & XLT_FILTROREPASSADOS & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    FrmRelRepasse.MousePointer = vbDefault
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    
End Sub

Private Sub CmdLimparIndexador_Click()
  
  CboIndexador.BoundText = ""
  
End Sub

Private Sub CmdLimparNaturezaPlano_Click()
  
  CboNaturezasPlanos.ListIndex = -1
  
End Sub

Private Sub CmdLimparTipoPlano_Click()
  
  CboTiposPlanos.BoundText = ""
  
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelExtratoContratoAnalitico
End Sub

Private Sub Form_Load()
  subManutencaoJanelasAtivas "I", "FrmRelExtratoContratoAnalitico"
  subConectarControleDadosNV DatTipoPlano, "SELECT * FROM TiposPlanos ORDER BY tipl_tx_Descricao", Estatico
  subConectarControleDadosNV DatObs, "SELECT * FROM Observacoes ORDER BY obse_tx_observacao", Estatico
  subConectarControleDadosNV DatIndexador, "SELECT * FROM Moedas ORDER BY moed_tx_Descricao", Estatico
  DtpDtBase.Value = Date
  
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

Private Sub txtPercentual_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub


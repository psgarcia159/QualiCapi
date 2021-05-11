VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmRelContratosComSeguro 
   Caption         =   "Contratos Com Seguro"
   ClientHeight    =   1470
   ClientLeft      =   945
   ClientTop       =   2190
   ClientWidth     =   4620
   Icon            =   "FrmRelContratosComSeguro.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1470
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   345
      Left            =   2280
      TabIndex        =   0
      Top             =   1020
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3480
      TabIndex        =   1
      Top             =   1020
      Width           =   1035
   End
   Begin Threed.SSFrame FraEmpreendimento 
      Height          =   600
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   4425
      _Version        =   65536
      _ExtentX        =   7805
      _ExtentY        =   1058
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
      Begin MSMask.MaskEdBox MskEmpreendimentoInicial 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   180
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "AAAA"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MskEmpreendimentoFinal 
         Height          =   315
         Left            =   2880
         TabIndex        =   4
         Top             =   180
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "AAAA"
         PromptChar      =   " "
      End
      Begin VB.Label LblEmpreendimento 
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
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   1455
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
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   195
      End
   End
End
Attribute VB_Name = "FrmRelContratosComSeguro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    
  
    Dim XLT_FILTROEMPREENDIMENTOS As String
    'Dim XLT_FILTROEMPREENDIMENTOS As String
    Dim XLT_INDEXADOR As String
    Dim XLT_SQL As String
    
    FrmRelContratosComSeguro.MousePointer = vbHourglass
    
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
     
    
   'Empreendimento
    If MskEmpreendimentoInicial.Text <> "" And MskEmpreendimentoFinal.Text <> "" Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelContratosComSeguro.empd_cd_Empreendimento >= '" & MskEmpreendimentoInicial.Text & "'" & _
      " AND ConsCAPRelContratosComSeguro.empd_cd_Empreendimento <= '" & MskEmpreendimentoFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROEMPREENDIMENTOS = "Empreendimento: " & MskEmpreendimentoInicial & " a " & MskEmpreendimentoFinal
    ElseIf MskEmpreendimentoInicial.Text <> "" And MskEmpreendimentoFinal.Text = "" Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelContratosComSeguro.empd_cd_Empreendimento >= '" & MskEmpreendimentoInicial.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROEMPREENDIMENTOS = "Empreendimento: A partir de " & MskEmpreendimentoInicial
    ElseIf MskEmpreendimentoInicial.Text = "" And MskEmpreendimentoFinal.Text <> "" Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelContratosComSeguro.empd_cd_Emprendimento <= '" & MskEmpreendimentoFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
      XLT_FILTROEMPREENDIMENTOS = "Empreendimento: Até " & MskEmpreendimentoInicial
    Else
      XLT_FILTROEMPREENDIMENTOS = "Empreendimento: Geral"
    End If
    
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelContratosComSeguro.empr_cd_Empresa = " & CStr(PCodEmpresa)
  
    'Ordenação
    'XGT_SELECAO = XGT_SELECAO + " ORDER BY titu_dt_Pagamento"
    
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsCAPRelContratosComSeguro WHERE " & XGT_SELECAO)
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        FrmRelContratosComSeguro.MousePointer = vbDefault
        Exit Sub
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
        
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelContratosComSeguro.rpt")
    
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("FiltroEmpreendimento").Text = Chr(34) & XLT_FILTROEMPREENDIMENTOS & Chr(34)
        '.GetItemByName("Indexador").Text = Chr(34) & XLT_INDEXADOR & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelContratosComSeguro.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelContratosComSeguro
End Sub

Private Sub Form_Load()
  
  subManutencaoJanelasAtivas "I", "FrmRelContratosComSeguro"
  
End Sub

Private Sub MskEmpreendimentoInicial_GotFocus()
  
  Call subSelecionaMSK(MskEmpreendimentoInicial)
  
End Sub

Private Sub MskEmpreendimentoFinal_GotFocus()
  
  Call subSelecionaMSK(MskEmpreendimentoFinal)
  
End Sub

Private Sub MskEmpreendimentoInicial_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub MskEmpreendimentoFinal_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub MskEmpreendimentoInicial_LostFocus()
   
  If MskEmpreendimentoInicial <> "" Then
    SubPreencheComZero MskEmpreendimentoInicial
  End If
  
End Sub

Private Sub MskEmpreendimentoFinal_LostFocus()
   
  If MskEmpreendimentoFinal <> "" Then
    SubPreencheComZero MskEmpreendimentoFinal
  End If
  
End Sub


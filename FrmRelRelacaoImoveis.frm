VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmRelRelacaoImoveis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório Relação de Imóveis Quitados"
   ClientHeight    =   3810
   ClientLeft      =   3330
   ClientTop       =   3615
   ClientWidth     =   5565
   Icon            =   "FrmRelRelacaoImoveis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Período da Quitação"
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
      Left            =   180
      TabIndex        =   11
      Top             =   1560
      Width           =   5310
      Begin MSComCtl2.DTPicker dtpQuitacaoIni 
         Height          =   315
         Left            =   795
         TabIndex        =   12
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   85852161
         CurrentDate     =   37459
      End
      Begin MSComCtl2.DTPicker dtpQuitacaoFim 
         Height          =   315
         Left            =   3420
         TabIndex        =   13
         Top             =   285
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   85852161
         CurrentDate     =   37459
      End
      Begin VB.Label Label2 
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
         Left            =   2880
         TabIndex        =   15
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label1 
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
         Left            =   255
         TabIndex        =   14
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.Frame FraPeriodo 
      Caption         =   "Período da Venda"
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
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   5310
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   795
         TabIndex        =   3
         Top             =   285
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   85852161
         CurrentDate     =   37459
      End
      Begin MSComCtl2.DTPicker dtpFim 
         Height          =   315
         Left            =   3420
         TabIndex        =   4
         Top             =   285
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   85852161
         CurrentDate     =   37459
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
         Left            =   255
         TabIndex        =   6
         Top             =   330
         Width           =   495
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
         Left            =   2880
         TabIndex        =   5
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   4470
      TabIndex        =   1
      Top             =   3255
      Width           =   1035
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Height          =   345
      Left            =   3330
      TabIndex        =   0
      Top             =   3255
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc DatEmpreendimento 
      Height          =   330
      Left            =   780
      Top             =   4800
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
   Begin MSMask.MaskEdBox MskEmpInicial 
      Height          =   315
      Left            =   2700
      TabIndex        =   7
      Top             =   240
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "AAAA"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskEmpFinal 
      Height          =   315
      Left            =   3480
      TabIndex        =   8
      Top             =   240
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "AAAA"
      PromptChar      =   " "
   End
   Begin Threed.SSFrame FraStatus 
      Height          =   765
      Left            =   180
      TabIndex        =   16
      Top             =   2400
      Width           =   5310
      _Version        =   65536
      _ExtentX        =   9366
      _ExtentY        =   1349
      _StockProps     =   14
      Caption         =   "Status dos Contratos"
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
         Caption         =   "Inativos"
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
         Left            =   2220
         TabIndex        =   19
         Top             =   300
         Width           =   1000
      End
      Begin VB.CheckBox ChkAtivo 
         Caption         =   "Ativos"
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
         Left            =   975
         TabIndex        =   18
         Top             =   300
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VB.CheckBox ChkCedido 
         Caption         =   "Cedidos"
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
         Left            =   3480
         TabIndex        =   17
         Top             =   300
         Width           =   1000
      End
   End
   Begin VB.Label LblImovel 
      Alignment       =   2  'Center
      Caption         =   "Faixa de Empreendimentos:"
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
      Left            =   180
      TabIndex        =   10
      Top             =   300
      Width           =   2505
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
      Left            =   3300
      TabIndex        =   9
      Top             =   300
      Width           =   135
   End
End
Attribute VB_Name = "FrmRelRelacaoImoveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    Dim XLT_FILTROSTATUS As String          'Armazenar filtro do status p/ relatório
    Dim XLT_FILTROPERIODO As String         'Armazenar filtro do período de venda escolhido
    Dim XLT_STATUS As String                '1º Status escolhido
    Dim XLT_STATUS2 As String               '2º Status escolhido
    Dim XLT_STATUS_IN As String             'clausula IN do where
    Dim XLT_EMPREENDIMENTO As String        'Empreendimento escolhido
    Dim XLT_FILTROPeriodoQuitacao As String 'Armazenar filtro do período de quitacao escolhido
    
    
    FrmRelRelacaoImoveis.MousePointer = vbHourglass
 
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
    
    'Filtra a empresa Ativa
    XGT_SELECAO = XGT_SELECAO + "ConsCAPRelSituacaoImoveis.empr_cd_empresa = " & CStr(PCodEmpresa)
    XGT_CONJUNCAO = " AND "
  
    'validacao das datas
    If dtpFim.Value < dtpInicio.Value Then
      MsgBox "intervalo inválido", vbCritical, "Data da venda"
      Exit Sub
    End If
    If dtpQuitacaoFim.Value < dtpQuitacaoIni Then
      MsgBox "Intervalo de Quitação Inválido", vbCritical, "Data da quitacao"
      Exit Sub
    End If
    
    'Período de venda
    
    If (Not IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelSituacaoImoveis.dtVenda >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & "" & _
        " AND ConsCAPRelSituacaoImoveis.dtVenda <= " & FunNuloData(dtpFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Período da venda: " & Format$(dtpInicio.Value, "dd/mm/yy") & " a " & Format$(dtpFim.Value, "dd/mm/yy")
    ElseIf (Not IsNull(dtpInicio.Value)) And (IsNull(dtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelSituacaoImoveis.dtVenda >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Período da venda: A partir de " & Format$(dtpInicio.Value, "dd/mm/yy")
    ElseIf (IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelSituacaoImoveis.dtVenda <= " & FunNuloData(dtpFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPERIODO = "Período da venda: Até " & Format$(dtpFim.Value, "dd/mm/yy")
    Else
        XLT_FILTROPERIODO = "Período da venda: Geral"
    End If
    
    'Periodo da quitacao
    
    If (Not IsNull(dtpQuitacaoIni.Value)) And (Not IsNull(dtpQuitacaoFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelSituacaoImoveis.dtQuitacao >= " & FunNuloData(dtpQuitacaoIni.Value, NomeSgbd) & "" & _
        " AND ConsCAPRelSituacaoImoveis.dtQuitacao<= " & FunNuloData(dtpQuitacaoFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPeriodoQuitacao = "Período da Quitação: " & Format$(dtpQuitacaoIni.Value, "dd/mm/yy") & " a " & Format$(dtpQuitacaoFim.Value, "dd/mm/yy")
    ElseIf (Not IsNull(dtpQuitacaoIni.Value)) And (IsNull(dtpQuitacaoFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelSituacaoImoveis.dtQuitacao >= " & FunNuloData(dtpQuitacaoIni.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPeriodoQuitacao = "Período da Quitação: A partir de " & Format$(dtpInicio.Value, "dd/mm/yy")
    ElseIf (IsNull(dtpQuitacaoIni.Value)) And (Not IsNull(dtpQuitacaoFim.Value)) Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelSituacaoImoveis.dtQuitacao <= " & FunNuloData(dtpQuitacaoFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
        XLT_FILTROPeriodoQuitacao = "Período da Quitação: Até " & Format$(dtpQuitacaoFim.Value, "dd/mm/yy")
    Else
        XLT_FILTROPeriodoQuitacao = "Período da Quitação: Geral"
    End If
  
    If (MskEmpInicial.Text <> "") Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "ConsCAPRelSituacaoImoveis.empd_cd_Empreendimento >= '" & MskEmpInicial.Text & "'"
      If MskEmpFinal.Text <> "" Then
        XGT_SELECAO = XGT_SELECAO & " and ConsCAPRelSituacaoImoveis.empd_cd_Empreendimento <= '" & MskEmpFinal.Text & "'"
      End If
    XGT_CONJUNCAO = " AND "
    Else
      If MskEmpFinal.Text <> "" Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        " ConsCAPRelSituacaoImoveis.empd_cd_Empreendimento <= '" & MskEmpFinal.Text & "'"
      End If
    XGT_CONJUNCAO = " AND "
    End If


    'Filtra Status dos contratos
    
    XLT_STATUS_IN = ""
    XLT_STATUS = ""
    XLT_STATUS2 = ""
    
    If ChkAtivo.Value = ChkInativo.Value And ChkAtivo.Value = ChkCedido.Value Then  'Todos Contratos
        XLT_STATUS = "T"
        XLT_FILTROSTATUS = "Status: Geral"
'        XLT_STATUS_IN = "'A', 'I', 'C'"
    Else
    
        If ChkAtivo.Value = 1 Then                                                  'Contratos Ativos
            XLT_STATUS = "A"
            XLT_FILTROSTATUS = "Status: Contratos Ativos "
            XLT_STATUS_IN = "'A'"
            
            If ChkInativo.Value = 1 Then                                            'Contratos Ativos e Inativos
                XLT_STATUS2 = "I"
                XLT_FILTROSTATUS = XLT_FILTROSTATUS + "e Inativos"
                XLT_STATUS_IN = "'A', 'I'"
            ElseIf ChkCedido.Value = 1 Then                                         'Contratos Ativos e Cedidos
                XLT_STATUS2 = "C"
                XLT_FILTROSTATUS = XLT_FILTROSTATUS + "e Cedidos"
                XLT_STATUS_IN = "'A', 'C'"
            End If
            
        ElseIf ChkInativo.Value = 1 Then                                            'Contratos Inativos
            XLT_STATUS = "I"
            XLT_FILTROSTATUS = "Status: Contratos Inativos "
            XLT_STATUS_IN = "'I'"
            
            If ChkCedido.Value = 1 Then                                             'Contratos Inativos e Cedidos
                XLT_STATUS2 = "C"
                XLT_FILTROSTATUS = XLT_FILTROSTATUS + "e Cedidos"
                XLT_STATUS_IN = "'I', 'C'"
            End If
    
        ElseIf ChkCedido.Value = 1 Then                                             'Contratos Cedidos
            XLT_STATUS = "C"
            XLT_FILTROSTATUS = "Status: Contratos Cedidos"
            XLT_STATUS_IN = "'C'"
        End If
        
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "ConsCAPRelSituacaoImoveis.cont_tx_Status IN (" & XLT_STATUS_IN & ") "
        XGT_CONJUNCAO = " AND "

    End If
  
   'Filtra imoveis Quitados
    XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + "ConsCAPRelSituacaoImoveis.dtQuitacao is not null"
   
    'Fecha o recordset se estiver aberto
    If XGO_RSRELATORIO.State = adStateOpen Then
      XGO_RSRELATORIO.Close
      Set XGO_RSRELATORIO = Nothing
    End If
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsCAPRelSituacaoImoveis WHERE " & XGT_SELECAO)
    
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        FrmRelRelacaoImoveis.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Fecha o recordset se estiver aberto
    If XGO_RSCABECALHO.State = adStateOpen Then
      XGO_RSCABECALHO.Close
      Set XGO_RSCABECALHO = Nothing
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
  
    'Carrega os relatórios
    FrmVisRelatorios.Caption = "Relatório de Imóveis Quitados"
    Set XGR_RELATORIO = Nothing
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelRelacaoImoQuitado.rpt")
  
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("Periodo").Text = Chr(34) & XLT_FILTROPERIODO & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
        .GetItemByName("Status").Text = Chr(34) & XLT_FILTROSTATUS & Chr(34)
        .GetItemByName("dtQuitacao").Text = Chr(34) & XLT_FILTROPeriodoQuitacao & Chr(34)
    
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelRelacaoImoveis.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelRelacaoImoveis
End Sub

Private Sub Form_Load()
  subManutencaoJanelasAtivas "I", "FrmRelRelacaoImoveis"
  subConectarControleDadosNV DatEmpreendimento, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_Nome", Estatico
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

Private Sub MskEmpFinal_GotFocus()
  Call subSelecionaMSK(MskEmpInicial)
End Sub

Private Sub MskEmpInicial_GotFocus()
    Call subSelecionaMSK(MskEmpInicial)
End Sub



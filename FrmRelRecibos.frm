VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelRecibos 
   Caption         =   "Recibos"
   ClientHeight    =   2925
   ClientLeft      =   945
   ClientTop       =   2190
   ClientWidth     =   4695
   Icon            =   "FrmRelRecibos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraContrato 
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4695
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
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   4470
         Begin MSComCtl2.DTPicker DtpInicio 
            Height          =   315
            Left            =   660
            TabIndex        =   13
            Top             =   285
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   155713539
            CurrentDate     =   37678
         End
         Begin MSComCtl2.DTPicker DtpFim 
            Height          =   315
            Left            =   2760
            TabIndex        =   14
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   155713539
            CurrentDate     =   37617.651087963
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
            TabIndex        =   16
            Top             =   360
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
            Left            =   2220
            TabIndex        =   15
            Top             =   360
            Width           =   495
         End
      End
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
         Left            =   120
         TabIndex        =   7
         Top             =   1500
         Width           =   4470
         Begin MSComCtl2.DTPicker DtpInicioPagamento 
            Height          =   315
            Left            =   660
            TabIndex        =   8
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   155713539
            CurrentDate     =   37617.6519560185
         End
         Begin MSComCtl2.DTPicker DtpFimPagamento 
            Height          =   315
            Left            =   2760
            TabIndex        =   9
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   155713539
            CurrentDate     =   37617.6517939815
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
            TabIndex        =   11
            Top             =   360
            Width           =   495
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
            TabIndex        =   10
            Top             =   360
            Width           =   495
         End
      End
      Begin MSMask.MaskEdBox MskContratoInicial 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "AAAA.AAAA.##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MskContratoFinal 
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "AAAA.AAAA.##"
         PromptChar      =   " "
      End
      Begin VB.Label LblContrato 
         Alignment       =   1  'Right Justify
         Caption         =   "Contrato:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   240
         Width           =   825
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
         Left            =   2580
         TabIndex        =   5
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   345
      Left            =   2565
      TabIndex        =   0
      Top             =   2520
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3660
      TabIndex        =   1
      Top             =   2520
      Width           =   1035
   End
End
Attribute VB_Name = "FrmRelRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    
  Dim XLT_SQL As String
  Dim XLT_CONJUNCAO As String
  Dim XLT_DOCPAGTO As String
  Dim XLT_DEPOSITO As String
    
  Me.MousePointer = vbHourglass
  XLT_SQL = ""
  XGT_CONJUNCAO = ""
  
  XLT_SQL = " SELECT consCAPRecibo.*," & _
   "(SELECT COUNT(*) FROM Titulos AS tit WHERE tit.empr_cd_empresa=consCAPRecibo.empr_cd_empresa  " & _
   " AND tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento " & _
   " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND  tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato " & _
   " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_residuo = '00') AS TotalParcelas, " & _
   " IsNull((SELECT titu_vl_Parcela FROM ConscapTitulos as tit WHERE tit.empr_cd_empresa=consCAPRecibo.empr_cd_empresa  " & _
   " AND tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento" & _
   " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato" & _
   " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_Parcela = consCAPRecibo.titu_cd_Parcela" & _
   " AND (tit.titu_cd_Residuo = (consCAPRecibo.titu_cd_residuo + 1))), 0) AS ValorResiduo" & _
   " FROM consCAPRecibo WHERE empr_cd_empresa = " & PCodEmpresa _

  'Contrato
    If MskContratoInicial.Text <> "    .    .  " And MskContratoFinal.Text <> "    .    .  " Then
      XLT_SQL = XLT_SQL & XGT_CONJUNCAO & _
      " AND consCAPRecibo.Contrato >= '" & MskContratoInicial.Text & "'" & _
      " AND consCAPRecibo.Contrato <= '" & MskContratoFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
    ElseIf MskContratoInicial.Text <> "    .    .  " And MskContratoFinal.Text = "    .    .  " Then
      XLT_SQL = XLT_SQL & XGT_CONJUNCAO & _
      " AND consCAPRecibo.Contrato >= '" & MskContratoInicial.Text & "'"
      XGT_CONJUNCAO = " AND "
    ElseIf MskContratoInicial.Text = "    .    .  " And MskContratoFinal.Text <> "    .    .  " Then
      XLT_SQL = XLT_SQL & XGT_CONJUNCAO & _
      " AND consCAPRecibo.Contrato <= '" & MskContratoFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
    End If
    
    'Período de Vencimento
    If (Not IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
       XLT_SQL = XLT_SQL & XGT_CONJUNCAO & _
        " AND consCAPRecibo.titu_dt_Vencimento >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & "" & _
        " AND consCAPRecibo.titu_dt_Vencimento <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
    ElseIf (Not IsNull(DtpInicio.Value)) And (IsNull(DtpFim.Value)) Then
        XLT_SQL = XLT_SQL & XGT_CONJUNCAO & _
        " AND consCAPRecibo.titu_dt_Vencimento >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & ""
        XGT_CONJUNCAO = " AND "
    ElseIf (IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
        XLT_SQL = XLT_SQL & XGT_CONJUNCAO & _
        " AND (consCAPRecibo.titu_dt_Vencimento <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ")"
        XGT_CONJUNCAO = " AND "
    End If
    
    'Periodo de Pagamento
    If (Not IsNull(DtpInicioPagamento.Value)) And (Not IsNull(DtpFimPagamento.Value)) Then
       XLT_SQL = XLT_SQL & XGT_CONJUNCAO & _
        " AND consCAPRecibo.titu_dt_Pagamento >= " & FunNuloData(DtpInicioPagamento.Value, NomeSgbd) & "" & _
        " AND consCAPRecibo.titu_dt_Pagamento <= " & FunNuloData(DtpFimPagamento.Value, NomeSgbd) & ""
    ElseIf (Not IsNull(DtpInicioPagamento.Value)) And (IsNull(DtpFimPagamento.Value)) Then
        XLT_SQL = XLT_SQL & XGT_CONJUNCAO & _
        " AND consCAPRecibo.titu_dt_Pagamento >= " & FunNuloData(DtpInicioPagamento.Value, NomeSgbd) & ""
    ElseIf (IsNull(DtpInicioPagamento.Value)) And (Not IsNull(DtpFimPagamento.Value)) Then
        XLT_SQL = XLT_SQL & XGT_CONJUNCAO & _
        " AND consCAPRecibo.titu_dt_Pagamento <= " & FunNuloData(DtpFimPagamento.Value, NomeSgbd) & ""
    End If
  
  'Abre os recordsets e verifica se eles possuem dados
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL)
  If XGO_RSRELATORIO.EOF Then
    MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Erro na consulta"
    Me.MousePointer = vbDefault
    Exit Sub
  End If
  Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)

  'Carrega os relatórios
  FrmVisRelatorios.Caption = "Impressão de Recibo"
  
  If XGT_MODELORECIBO = "1" Then
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelRecibo.rpt")
  Else
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelRecibo2.rpt")
  End If

  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO

  'Ativa os Relatórios
  With XGR_RELATORIO.FormulaFields
    .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    .GetItemByName("DocPgto").Text = Chr(34) & Replace(XLT_DOCPAGTO, vbCrLf, " - ") & Chr(34)
    .GetItemByName("DepositoBanco").Text = Chr(34) & XLT_DEPOSITO & Chr(34)
  End With
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Rodape"
  FrmVisRelatorios.SubVerRelatorioMoldal XGR_RELATORIO, Me
  Me.MousePointer = vbDefault
  
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelRecibos
End Sub

Private Sub Form_Load()
  
  subManutencaoJanelasAtivas "I", "FrmRelRecibos"
  
End Sub

Private Sub MskContratoInicial_GotFocus()
  
  Call subSelecionaMSK(MskContratoInicial)
  
End Sub

Private Sub MskContratoFinal_GotFocus()
  
  Call subSelecionaMSK(MskContratoFinal)
  
End Sub

Private Sub MskContratoInicial_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub MskContratoFinal_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
  End If
End Sub

Private Sub MskContratoInicial_LostFocus()
   
  If MskContratoInicial.Text <> "    .    .  " Then
    If Len(Trim(MskContratoInicial.Text)) < 12 Then
      MsgBox "O código inicial do contrato deve possuir 12 dígitos, ou deve estar vazio.", vbCritical, "ATENÇÃO !"
      MskContratoInicial.SetFocus
    End If
  End If
  
End Sub

Private Sub MskContratoFinal_LostFocus()
   
  If MskContratoFinal.Text <> "    .    .  " Then
    If Len(Trim(MskContratoFinal.Text)) < 12 Then
      MsgBox "O código final do imóvel deve possuir 12 dígitos, ou deve estar vazio", vbCritical, "ATENÇÃO !"
      MskContratoFinal.SetFocus
    End If
  End If
  
End Sub

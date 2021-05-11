VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelAssistenciaTecnica 
   Caption         =   "Assistência Técnica"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4875
      TabIndex        =   18
      Top             =   6120
      Width           =   1035
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   345
      Left            =   3600
      TabIndex        =   17
      Top             =   6120
      Width           =   1035
   End
   Begin VB.Frame Frame3 
      Height          =   5295
      Left            =   5040
      TabIndex        =   20
      Top             =   720
      Width           =   4095
      Begin VB.Frame Frame13 
         Caption         =   "Responsável pela Conclusão"
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
         Left            =   120
         TabIndex        =   30
         Top             =   3600
         Width           =   3855
         Begin VB.TextBox RespConclusao 
            Height          =   405
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Responsável pelo Início dos Trabalhos"
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
         Left            =   120
         TabIndex        =   29
         Top             =   2760
         Width           =   3855
         Begin VB.TextBox RespInicioTrabalhos 
            Height          =   405
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Responsável pela Primeira Visita"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   3855
         Begin VB.TextBox RespPrimeiraVisita 
            Height          =   405
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Obsevações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   3855
         Begin VB.TextBox Observacao 
            Height          =   495
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Ocorrência"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   3855
         Begin VB.TextBox Ocorrencia 
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   3615
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5295
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   4815
      Begin VB.Frame Frame14 
         Caption         =   "Data"
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
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   4455
         Begin MSComCtl2.DTPicker DtpDataInicio 
            Height          =   315
            Left            =   720
            TabIndex        =   2
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   66650115
            CurrentDate     =   37678
         End
         Begin MSComCtl2.DTPicker DtpDataFim 
            Height          =   315
            Left            =   2820
            TabIndex        =   3
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   66650115
            CurrentDate     =   37617.651087963
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
            Left            =   2280
            TabIndex        =   34
            Top             =   300
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
            TabIndex        =   33
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Período de Conclusão"
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
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   4455
         Begin MSComCtl2.DTPicker DtpConclusaoInicio 
            Height          =   315
            Left            =   720
            TabIndex        =   10
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   66650115
            CurrentDate     =   37678
         End
         Begin MSComCtl2.DTPicker DtpConclusaoFim 
            Height          =   315
            Left            =   2820
            TabIndex        =   11
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   66650115
            CurrentDate     =   37617.651087963
         End
         Begin VB.Label Label8 
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
            TabIndex        =   42
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label4 
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
            Left            =   2280
            TabIndex        =   38
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Período de Início dos Trabalhos"
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
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   4455
         Begin MSComCtl2.DTPicker DtpInicioTrabalhosInicio 
            Height          =   315
            Left            =   720
            TabIndex        =   8
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   66650115
            CurrentDate     =   37678
         End
         Begin MSComCtl2.DTPicker DtpInicioTrabalhosFim 
            Height          =   315
            Left            =   2820
            TabIndex        =   9
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   66650115
            CurrentDate     =   37617.651087963
         End
         Begin VB.Label Label7 
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
            TabIndex        =   41
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label3 
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
            Left            =   2280
            TabIndex        =   37
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Período da Primeira Visita"
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
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   4455
         Begin MSComCtl2.DTPicker DtpPrimeiraVisitaInicio 
            Height          =   315
            Left            =   720
            TabIndex        =   6
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   66650115
            CurrentDate     =   37678
         End
         Begin MSComCtl2.DTPicker DtpPrimeiraVisitaFim 
            Height          =   315
            Left            =   2820
            TabIndex        =   7
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   66650115
            CurrentDate     =   37617.651087963
         End
         Begin VB.Label Label6 
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
            TabIndex        =   40
            Top             =   300
            Width           =   495
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
            Left            =   2280
            TabIndex        =   36
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Período da Emissão"
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
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   4455
         Begin MSComCtl2.DTPicker DtpEmissaoInicio 
            Height          =   315
            Left            =   720
            TabIndex        =   4
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   66650115
            CurrentDate     =   37678
         End
         Begin MSComCtl2.DTPicker DtpEmissaoFim 
            Height          =   315
            Left            =   2820
            TabIndex        =   5
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   66650115
            CurrentDate     =   37617.651087963
         End
         Begin VB.Label Label5 
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
            TabIndex        =   39
            Top             =   300
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
            Left            =   2280
            TabIndex        =   35
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Autorizador"
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
         Left            =   120
         TabIndex        =   21
         Top             =   4440
         Width           =   3375
         Begin VB.TextBox Autorizador 
            Height          =   285
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   3135
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.TextBox EmpreendimentoFinal 
         Height          =   285
         Left            =   4920
         TabIndex        =   45
         Top             =   220
         Width           =   1095
      End
      Begin VB.TextBox EmpreendimentoInicial 
         Height          =   285
         Left            =   3480
         TabIndex        =   1
         Top             =   220
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Left            =   4560
         TabIndex        =   44
         Top             =   240
         Width           =   225
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
         Height          =   225
         Left            =   1800
         TabIndex        =   31
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.Data DatAssistencia 
      Caption         =   "DatAssistencia"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Width           =   2535
   End
End
Attribute VB_Name = "FrmRelAssistenciaTecnica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private XLT_EMPREENDIMENTOINICIAL As String
Private XLT_EMPREENDIMENTOFINAL As String
Private XLT_DATAINICIO As String
Private XLT_DATAFIM As String
Private XLT_EMISSAOINICIO As String
Private XLT_EMISSAOFIM As String
Private XLT_PRIMEIRAVISITAINICIO As String
Private XLT_PRIMEIRAVISITAFIM As String
Private XLT_INICIOTRABALHOSINICIO As String
Private XLT_INICIOTRABALHOSFIM As String
Private XLT_CONCLUSAOINICIO As String
Private XLT_CONCLUSAOFIM As String
Private XLT_OCORRENCIA As String
Private XLT_OBSERVACAO As String
Private XLT_RESPPRIMEIRAVISITA As String
Private XLT_RESPINICIOTRABALHOS As String
Private XLT_RESPCONCLUSAO As String
Private XLT_AUTORIZADOR As String



Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()

Dim XGT_SQL As String
Dim XGT_CONJUNCAO As String
    FrmRelAssistenciaTecnica.MousePointer = vbHourglass
    
    XLT_EMPREENDIMENTOINICIAL = EmpreendimentoInicial.Text
    XLT_EMPREENDIMENTOFINAL = EmpreendimentoFinal.Text
    XLT_DATAINICIO = FunNulo(Format(DtpDataInicio.Value, "DD/MM/YYYY"))
    'XLT_DATAINICIO = FunNulo(DtpDataInicio.Value)
    XLT_DATAFIM = FunNulo(Format(DtpDataFim.Value, "DD/MM/YYYY"))
    'XLT_DATAFIM = FunNulo(DtpDataFim.Value)
    XLT_EMISSAOINICIO = FunNulo(Format(DtpEmissaoInicio.Value, "DD/MM/YYYY"))
   'XLT_EMISSAOINICIO = FunNulo(DtpEmissaoInicio.Value)
    XLT_EMISSAOFIM = FunNulo(Format(DtpEmissaoFim.Value, "DD/MM/YYYY"))
    'XLT_EMISSAOFIM = FunNulo(DtpEmissaoFim.Value)
    XLT_PRIMEIRAVISITAINICIO = FunNulo(Format(DtpPrimeiraVisitaInicio.Value, "DD/MM/YYYY"))
    'XLT_PRIMEIRAVISITAINICIO = FunNulo(DtpPrimeiraVisitaInicio.Value)
    XLT_PRIMEIRAVISITAFIM = FunNulo(Format(DtpPrimeiraVisitaFim.Value, "DD/MM/YYYY"))
    'XLT_PRIMEIRAVISITAFIM = FunNulo(DtpPrimeiraVisitaFim.Value)
    XLT_INICIOTRABALHOSINICIO = FunNulo(Format(DtpInicioTrabalhosInicio.Value, "DD/MM/YYYY"))
    'XLT_INICIOTRABALHOSINICIO = FunNulo(DtpInicioTrabalhosInicio.Value)
    XLT_INICIOTRABALHOSFIM = FunNulo(Format(DtpInicioTrabalhosFim.Value, "DD/MM/YYYY"))
    'XLT_INICIOTRABALHOSFIM = FunNulo(DtpInicioTrabalhosFim.Value)
    XLT_CONCLUSAOINICIO = FunNulo(Format(DtpConclusaoInicio.Value, "DD/MM/YYYY"))
    'XLT_CONCLUSAOINICIO = FunNulo(DtpConclusaoInicio.Value)
    XLT_CONCLUSAOFIM = FunNulo(Format(DtpConclusaoFim.Value, "DD/MM/YYYY"))
    XLT_CONCLUSAOFIM = FunNulo(DtpConclusaoFim.Value)
    XLT_OCORRENCIA = Ocorrencia.Text
    XLT_OBSERVACAO = Observacao.Text
    XLT_RESPPRIMEIRAVISITA = RespPrimeiraVisita.Text
    XLT_RESPINICIOTRABALHOS = RespInicioTrabalhos.Text
    XLT_RESPCONCLUSAO = RespConclusao.Text
    XLT_AUTORIZADOR = Autorizador.Text
    
    
    
    
    
    XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE empr_cd_empresa = " & PCodEmpresa & " "
    XGT_CONJUNCAO = " AND "
    
'Filtra pelo Empreendimento
    If (XLT_EMPREENDIMENTOINICIAL <> "") And (XLT_EMPREENDIMENTOFINAL <> "") Then
        XGT_SQL = "empd_cd_Empreendimento >= '" & XLT_EMPREENDIMENTOINICIAL & "' AND empd_cd_Empreendimento <= '" & XLT_EMPREENDIMENTOFINAL & "' "
        If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
            XGT_SELECAO = XGT_SELECAO + XGT_SQL
        Else
            XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
        End If
    ElseIf (XLT_EMPREENDIMENTOINICIAL <> "") Then
        XGT_SQL = "empd_cd_Empreendimento >= '" & XLT_EMPREENDIMENTOINICIAL & "'"
            If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
                XGT_SELECAO = XGT_SELECAO + XGT_SQL
            Else
                XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
            End If
            
        ElseIf (XLT_EMPREENDIMENTOFINAL <> "") Then
        XGT_SQL = "empd_cd_Empreendimento <= '" & XLT_EMPREENDIMENTOFINAL & "'"
            If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
                XGT_SELECAO = XGT_SELECAO + XGT_SQL
            Else
                XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
            End If
        End If
    
'Filtra pela Ocorrencia
    If (XLT_OCORRENCIA <> "") Then
        XGT_SQL = "assi_tx_Ocorrencia LIKE'" & "%" & XLT_OCORRENCIA & "%" & "' "
         If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
           XGT_SELECAO = XGT_SELECAO + XGT_SQL
        Else
            XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
        End If
    
    End If
    
    
'Filtra pela Oservacao
    If (XLT_OBSERVACAO <> "") Then
        XGT_SQL = "assi_tx_Observacao LIKE'" & "%" & XLT_OBSERVACAO & "%" & "' "
         If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
           XGT_SELECAO = XGT_SELECAO + XGT_SQL
        Else
            XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
        End If
    
    End If
    
    
'Filtra pelo Responsavel Pela Primeira Visita
    If (XLT_RESPPRIMEIRAVISITA <> "") Then
        XGT_SQL = "assi_tx_RespPrVisita = '" & XLT_RESPPRIMEIRAVISITA & "'"
         If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
           XGT_SELECAO = XGT_SELECAO + XGT_SQL
        Else
            XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
        End If
    
    End If
    
'Filtra pelo Responsavel Pelo Inicio dos Trabalhos
    If (XLT_RESPINICIOTRABALHOS <> "") Then
        XGT_SQL = "assi_tx_RespIniTrabalhos = '" & XLT_RESPINICIOTRABALHOS & "'"
         If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
           XGT_SELECAO = XGT_SELECAO + XGT_SQL
        Else
            XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
        End If
    
    End If
    
'Filtra pelo Responsavel Pela Conclusao
    If (XLT_RESPCONCLUSAO <> "") Then
        XGT_SQL = "assi_tx_RespConclusao = '" & XLT_RESPCONCLUSAO & "'"
         If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
           XGT_SELECAO = XGT_SELECAO + XGT_SQL
        Else
            XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
        End If
    
    End If
    
'Filtra pelo Autorizador
    If (XLT_AUTORIZADOR <> "") Then
        XGT_SQL = "assi_tx_Autorizador = '" & XLT_AUTORIZADOR & "'"
         If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
           XGT_SELECAO = XGT_SELECAO + XGT_SQL
        Else
            XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
        End If
    
    End If
    
'Filtra pela Data
    If (XLT_DATAINICIO <> "") And (XLT_DATAFIM <> "") Then
        XGT_SQL = "assi_dt_Data >= " & XLT_DATAINICIO & " AND assi_dt_Data <= " & XLT_DATAFIM & " "
        If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
            XGT_SELECAO = XGT_SELECAO + XGT_SQL
        Else
            XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
        End If
    ElseIf (XLT_DATAINICIO <> "") Then
        XGT_SQL = "assi_dt_Data >= " & XLT_DATAINICIO & ""
            If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
                XGT_SELECAO = XGT_SELECAO + XGT_SQL
            Else
                XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
            End If
            
        ElseIf (XLT_DATAFIM <> "") Then
        XGT_SQL = "assi_dt_Data <= " & XLT_DATAFIM & ""
            If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
                XGT_SELECAO = XGT_SELECAO + XGT_SQL
            Else
                XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
            End If
        End If
        
'Filtra pela Data de Emissao
    If (XLT_EMISSAOINICIO <> "") And (XLT_EMISSAOFIM <> "") Then
        XGT_SQL = "assi_dt_DataEmissao >= '" & XLT_EMISSAOINICIO & "' AND assi_dt_DataEmissao <= '" & XLT_EMISSAOFIM & "' "
        If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
            XGT_SELECAO = XGT_SELECAO + XGT_SQL
        Else
            XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
        End If
    ElseIf (XLT_EMISSAOINICIO <> "") Then
        XGT_SQL = "assi_dt_DataEmissao >= '" & XLT_EMISSAOINICIO & "'"
            If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
                XGT_SELECAO = XGT_SELECAO + XGT_SQL
            Else
                XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
            End If
            
        ElseIf (XLT_EMISSAOFIM <> "") Then
        XGT_SQL = "assi_dt_DataEmissao <= '" & XLT_EMISSAOFIM & "'"
            If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
                XGT_SELECAO = XGT_SELECAO + XGT_SQL
            Else
                XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
            End If
        End If
        
'Filtra pela Primeira Visita
    If (XLT_PRIMEIRAVISITAINICIO <> "") And (XLT_PRIMEIRAVISITAFIM <> "") Then
        XGT_SQL = "assi_dt_PrimeiraVisita >= '" & XLT_PRIMEIRAVISITAINICIO & "' AND assi_dt_PrimeiraVisita <= '" & XLT_PRIMEIRAVISITAFIM & "' "
        If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
            XGT_SELECAO = XGT_SELECAO + XGT_SQL
        Else
            XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
        End If
    ElseIf (XLT_PRIMEIRAVISITAINICIO <> "") Then
        XGT_SQL = "assi_dt_PrimeiraVisita >= '" & XLT_PRIMEIRAVISITAINICIO & "'"
            If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
                XGT_SELECAO = XGT_SELECAO + XGT_SQL
            Else
                XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
            End If
            
        ElseIf (XLT_PRIMEIRAVISITAFIM <> "") Then
        XGT_SQL = "assi_dt_PrimeiraVisita <= '" & XLT_PRIMEIRAVISITAFIM & "'"
            If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
                XGT_SELECAO = XGT_SELECAO + XGT_SQL
            Else
                XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
            End If
        End If
        
'Filtra pela data de Inicio dos Trabalhos
    If (XLT_INICIOTRABALHOSINICIO <> "") And (XLT_INICIOTRABALHOSFIM <> "") Then
        XGT_SQL = "assi_dt_InicioTrabalho >= '" & XLT_INICIOTRABALHOSINICIO & "' AND assi_dt_InicioTrabalho <= '" & XLT_INICIOTRABALHOSFIM & "' "
        If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
            XGT_SELECAO = XGT_SELECAO + XGT_SQL
        Else
            XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
        End If
    ElseIf (XLT_INICIOTRABALHOSINICIO <> "") Then
        XGT_SQL = "assi_dt_InicioTrabalho >= '" & XLT_INICIOTRABALHOSINICIO & "'"
            If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
                XGT_SELECAO = XGT_SELECAO + XGT_SQL
            Else
                XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
            End If
            
        ElseIf (XLT_INICIOTRABALHOSFIM <> "") Then
        XGT_SQL = "assi_dt_InicioTrabalho <= '" & XLT_INICIOTRABALHOSFIM & "'"
            If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
                XGT_SELECAO = XGT_SELECAO + XGT_SQL
            Else
                XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
            End If
        End If
        
'Filtra pela data de Conclusao
    If (XLT_CONCLUSAOINICIO <> "") And (XLT_CONCLUSAOFIM <> "") Then
        XGT_SQL = "assi_dt_Conclusao >= '" & XLT_CONCLUSAOINICIO & "' AND assi_dt_Conclusao <= '" & XLT_CONCLUSAOFIM & "' "
        If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
            XGT_SELECAO = XGT_SELECAO + XGT_SQL
        Else
            XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
        End If
    ElseIf (XLT_CONCLUSAOINICIO <> "") Then
        XGT_SQL = "assi_dt_Conclusao >= '" & XLT_CONCLUSAOINICIO & "'"
            If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
                XGT_SELECAO = XGT_SELECAO + XGT_SQL
            Else
                XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
            End If
            
        ElseIf (XLT_CONCLUSAOFIM <> "") Then
        XGT_SQL = "assi_dt_Conclusao <= '" & XLT_CONCLUSAOFIM & "'"
            If XGT_SELECAO = "SELECT * FROM CONSCAPASSISTENCIATECNICA WHERE " Then
                XGT_SELECAO = XGT_SELECAO + XGT_SQL
            Else
                XGT_SELECAO = XGT_SELECAO + XGT_CONJUNCAO + XGT_SQL
            End If
        End If
        
        
     'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XGT_SELECAO)
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        FrmRelAssistenciaTecnica.MousePointer = vbDefault
        Exit Sub
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & CStr(PCodEmpresa))
       
    'mudar'
    FrmVisRelatorios.Caption = "Relatório Assistência Técnica"
    'mudar'
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelAssistenciaTecnica.rpt")
        
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    FrmRelAssistenciaTecnica.MousePointer = vbDefault
    
    
End Sub

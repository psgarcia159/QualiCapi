VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSolAssistencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitação de Assistência Técnica"
   ClientHeight    =   11220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11220
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Dados da Solicitação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.Frame Frame8 
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
         TabIndex        =   47
         Top             =   7680
         Width           =   3975
         Begin VB.TextBox Autorizador 
            Height          =   375
            Left            =   120
            MaxLength       =   50
            TabIndex        =   4
            Top             =   240
            Width           =   3495
         End
      End
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
         Height          =   375
         Left            =   4320
         TabIndex        =   15
         Top             =   10680
         Width           =   1215
      End
      Begin VB.CommandButton CmdGravar 
         Caption         =   "Gravar"
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
         Left            =   2880
         TabIndex        =   14
         Top             =   10680
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Caption         =   "Conclusão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   5760
         TabIndex        =   29
         Top             =   8520
         Width           =   2655
         Begin VB.TextBox RespConclusao 
            Height          =   375
            Left            =   120
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1320
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DataConclusao 
            Height          =   315
            Left            =   720
            TabIndex        =   11
            Top             =   360
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            Format          =   102170627
            CurrentDate     =   40544
         End
         Begin MSMask.MaskEdBox LblHoraConclusao 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   4
            EndProperty
            Height          =   300
            Left            =   720
            TabIndex        =   12
            Top             =   720
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm"
            Mask            =   "99:99"
            PromptChar      =   "_"
         End
         Begin VB.Label Label16 
            Caption         =   "Responsável:"
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
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Hora:"
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
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   525
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Data:"
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
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Início dos Trabalhos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   2880
         TabIndex        =   28
         Top             =   8520
         Width           =   2775
         Begin VB.TextBox RespIniTrabalhos 
            Height          =   375
            Left            =   120
            MaxLength       =   50
            TabIndex        =   10
            Top             =   1320
            Width           =   2500
         End
         Begin MSComCtl2.DTPicker DataIniTrabalhos 
            Height          =   315
            Left            =   720
            TabIndex        =   8
            Top             =   360
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            Format          =   102170627
            CurrentDate     =   40544
         End
         Begin MSMask.MaskEdBox LblHoraIniTrabalhos 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   4
            EndProperty
            Height          =   300
            Left            =   720
            TabIndex        =   9
            Top             =   720
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm"
            Mask            =   "99:99"
            PromptChar      =   "_"
         End
         Begin VB.Label Label15 
            Caption         =   "Responsável:"
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
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Hora:"
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
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   525
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Data:"
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
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "1ª Visita"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   27
         Top             =   8520
         Width           =   2655
         Begin VB.TextBox RespPrVisita 
            Height          =   375
            Left            =   120
            MaxLength       =   50
            TabIndex        =   7
            Top             =   1320
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DataPrVisita 
            Height          =   315
            Left            =   720
            TabIndex        =   5
            Top             =   360
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            Format          =   102170627
            CurrentDate     =   40544
         End
         Begin MSMask.MaskEdBox LblHoraPrVisita 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   4
            EndProperty
            Height          =   300
            Left            =   720
            TabIndex        =   6
            Top             =   720
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm"
            Mask            =   "99:99"
            PromptChar      =   "_"
         End
         Begin VB.Label Label14 
            Caption         =   "Responsável:"
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
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Hora:"
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
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   525
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Data:"
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
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Observações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   26
         Top             =   5280
         Width           =   8295
         Begin VB.TextBox TxtObservacao 
            Height          =   1935
            Left            =   120
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   8055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ocorrências"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   25
         Top             =   2760
         Width           =   8295
         Begin VB.TextBox TxtOcorrencia 
            Height          =   2055
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   240
            Width           =   8055
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   8295
         Begin MSComCtl2.DTPicker Data 
            Height          =   315
            Left            =   6720
            TabIndex        =   1
            Top             =   240
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   102170627
            CurrentDate     =   40544
         End
         Begin VB.Label LblEntregaChaves 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   6720
            TabIndex        =   49
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Entrega   Chaves:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5880
            TabIndex        =   48
            Top             =   1320
            Width           =   885
         End
         Begin VB.Label LblContrato 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1800
            TabIndex        =   46
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label17 
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
            Height          =   285
            Left            =   120
            TabIndex        =   45
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label LblHora 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   6720
            TabIndex        =   44
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label LblDataEmissao 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   6720
            TabIndex        =   43
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label LblNumeroApartamento 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1800
            TabIndex        =   42
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label LblNomeCliente 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   495
            Left            =   1800
            TabIndex        =   41
            Top             =   1800
            Width           =   3255
         End
         Begin VB.Label LblNomeEmpreendimento 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   1800
            TabIndex        =   40
            Top             =   600
            Width           =   3255
         End
         Begin VB.Label LblNumOrdem 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1800
            TabIndex        =   39
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label7 
            Caption         =   "Data:"
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
            Left            =   5880
            TabIndex        =   24
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label5 
            Caption         =   "Hora:"
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
            Left            =   5880
            TabIndex        =   22
            Top             =   960
            Width           =   525
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Emissão:"
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
            Left            =   5880
            TabIndex        =   21
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label3 
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
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente:"
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
            Left            =   120
            TabIndex        =   19
            Top             =   1800
            Width           =   1125
         End
         Begin VB.Label LblEmpreendimento 
            Caption         =   "Número:"
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
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label2 
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
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1485
         End
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Left            =   5880
         TabIndex        =   23
         Top             =   360
         Width           =   765
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "Adodc1"
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
Attribute VB_Name = "FrmSolAssistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private XFT_CODEMPREENDIMENTO As String
Private XFT_CODIMOVEL As String
Private XFT_CODCONTRATO As String
Private XFT_CODASSISTENCIA As String
Private XFT_ASSISTENCIA As String



Private Sub Form_Load()



    XFT_CODEMPREENDIMENTO = Left$(Chave, 4)
    XFT_CODIMOVEL = Mid$(Chave, 6, 4)
    XFT_CODCONTRATO = Mid$(Chave, 11, 2)
    
    If (XInserir = "A") Then
        XLT_ASSISTENCIA = Mid$(Chave, 14, 3)
    End If
    
    If (XInserir = "I") Then
    Data.Value = Now
    DataPrVisita.Value = Null
    DataIniTrabalhos.Value = Null
    DataConclusao.Value = Null
    End If
    
    Dim XLO_RS As ADODB.Recordset
    Dim XLT_SQL As String

    
    

    XGT_SQL = "SELECT DISTINCT " & _
          "Empreendimentos.empr_cd_Empresa, " & _
          "Empreendimentos.empd_cd_Empreendimento, " & _
          "Empreendimentos.empd_tx_Nome, " & _
          "Imoveis.imov_cd_Imovel, " & _
          "Contratos.cont_cd_Contrato, " & _
          "FornClientes.focl_tx_RazaoSocial, " & _
          "Imoveis.imov_dt_EntregaChaves " & _
          "From Empreendimentos " & _
    "INNER JOIN Imoveis on Empreendimentos.empd_cd_Empreendimento = Imoveis.empd_cd_Empreendimento " & _
    "AND Empreendimentos.empr_cd_Empresa = Imoveis.empr_cd_Empresa " & _
    "INNER JOIN Contratos on Contratos.imov_cd_Imovel = Imoveis.imov_cd_Imovel " & _
    "AND Contratos.empr_cd_Empresa = Imoveis.empr_cd_Empresa " & _
    "AND Contratos.empd_cd_Empreendimento = Imoveis.empd_cd_Empreendimento " & _
    "INNER JOIN FornClientes on FornClientes.focl_cd_FornCli = Contratos.focl_cd_FornCli " & _
    "Where  " & _
    "Empreendimentos.empr_cd_Empresa = " & PCodEmpresa & " AND " & _
    "Empreendimentos.empd_cd_Empreendimento + '.' + Imoveis.imov_cd_Imovel + '.' + cont_cd_Contrato = '" & XFT_CODEMPREENDIMENTO + "." & XFT_CODIMOVEL & "." & XFT_CODCONTRATO & "'"

    SubQOpenRecordset XLO_RS, XGT_SQL, "Estatico"
    If Not XLO_RS.EOF Then
        LblNomeEmpreendimento.Caption = XLO_RS!empd_cd_Empreendimento & " - " & XLO_RS!empd_tx_Nome
        LblNomeCliente.Caption = XLO_RS!focl_tx_RazaoSocial
        LblNumeroApartamento.Caption = XLO_RS!imov_cd_Imovel
        LblContrato.Caption = XLO_RS!cont_cd_Contrato
        LblEntregaChaves.Caption = Format(XLO_RS!imov_dt_EntregaChaves, "dd/MM/yyyy")
        
        
    End If
    XLO_RS.Close
    
    If (XInserir = "A") Then
        XGT_SQL = "SELECT * FROM ConsCapAssistenciaTecnica WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
                  " AND assi_tx_Titulo = '" & XFT_CODEMPREENDIMENTO + "." & XFT_CODIMOVEL & "." & XFT_CODCONTRATO & "." & XLT_ASSISTENCIA & "'"
        
        SubQOpenRecordset XLO_RS, XGT_SQL, "Estatico"
        If Not XLO_RS.EOF Then
            LblNumOrdem.Caption = XLO_RS!assi_cd_CodigoSeq
            Data.Value = Format(XLO_RS!assi_dt_Data, "dd/MM/yyyy")
            LblDataEmissao.Caption = Format(XLO_RS!assi_dt_DataEmissao, "dd/MM/yyyy")
            LblHora.Caption = Format(XLO_RS!assi_dt_DataEmissao, "HH:MM")
            If IsNull(XLO_RS!cont_dt_EntregaChaves) Then
                IsEmpty (LblEntregaChaves)
            Else
                LblEntregaChaves.Caption = Format(XLO_RS!cont_dt_EntregaChaves, "dd/MM/yyyy")
            End If
            TxtOcorrencia.Text = XLO_RS!assi_tx_Ocorrencia
            TxtObservacao.Text = XLO_RS!assi_tx_Observacao
            DataPrVisita.Value = Format(XLO_RS!assi_dt_PrimeiraVisita, "dd/MM/yyyy")
            If IsNull(XLO_RS!assi_dt_PrimeiraVisita) Then
                IsEmpty (LblHoraPrVisita)
            Else
                LblHoraPrVisita.Text = Format(XLO_RS!assi_dt_PrimeiraVisita, "HH:MM")
            End If

            DataIniTrabalhos.Value = Format(XLO_RS!assi_dt_InicioTrabalho, "dd/MM/yyyy")
            If IsNull(XLO_RS!assi_dt_InicioTrabalho) Then
                IsEmpty (LblHoraIniTrabalhos.Text)
            Else
                LblHoraIniTrabalhos.Text = Format(XLO_RS!assi_dt_InicioTrabalho, "HH:MM")
            End If
            DataConclusao.Value = Format(XLO_RS!assi_dt_Conclusao, "dd/MM/yyyy")
            If IsNull(XLO_RS!assi_dt_Conclusao) Then
                IsEmpty (LblHoraConclusao.Text)
            Else
                LblHoraConclusao.Text = Format(XLO_RS!assi_dt_Conclusao, "HH:MM")
            End If

            
            If Not IsNull(XLO_RS!assi_tx_Autorizador) Then
                Autorizador.Text = XLO_RS!assi_tx_Autorizador
            End If
            
            If Not IsNull(XLO_RS!assi_tx_RespPrVisita) Then
                RespPrVisita.Text = XLO_RS!assi_tx_RespPrVisita
            End If
            
            If Not IsNull(XLO_RS!assi_tx_RespIniTrabalhos) Then
                RespIniTrabalhos.Text = XLO_RS!assi_tx_RespIniTrabalhos
            End If
            
            If Not IsNull(XLO_RS!assi_tx_RespConclusao) Then
                RespConclusao.Text = XLO_RS!assi_tx_RespConclusao
            End If

        End If
            
        XLO_RS.Close
        Set XLO_RS = Nothing
        
    End If
    
End Sub


Private Sub CmdGravar_Click()
    On Error GoTo TrataErro
        
        Dim XLO_RS  As ADODB.Recordset
        Dim XLT_CODSEQUENCIAL As String
        Dim XLS_DATAPRVISITA As String
        Dim XLS_DATAINITRABALHO As String
        Dim XLS_DATACONCLUSAO As String
        

        
        
        If FunObrigatorioTXT(TxtOcorrencia, "Campo obrigatório") Then Exit Sub
        

        
'#########################################PRIMEIRA VISITA ##################################

        'verifica se os campos do workflow foram preenchidos caso o campo Autorizador tenha sido preenchido
        If (Autorizador.Text <> "") Then
            If FunObrigatorioDTP(DataPrVisita, "Campo obrigatório") Then Exit Sub
                If (LblHoraPrVisita.Text = "__:__") Then
                MsgBox "Campo obrigatório", vbCritical, "CAMPO OBRIGATÓRIO"
                LblHoraPrVisita.SetFocus
                Exit Sub
            End If
            If FunObrigatorioTXT(RespPrVisita, "Campo obrigatório") Then Exit Sub
        End If
            


        'verifica se os campos do workflow foram preenchidos caso o campo data tenha sido preenchido
        If Not IsNull(DataPrVisita.Value) Then
            If FunObrigatorioTXT(Autorizador, "Campo obrigatório") Then Exit Sub
            If (LblHoraPrVisita.Text = "__:__") Then
                MsgBox "Campo obrigatório", vbCritical, "CAMPO OBRIGATÓRIO"
                LblHoraPrVisita.SetFocus
                Exit Sub
            End If
            If FunObrigatorioTXT(RespPrVisita, "Campo obrigatório") Then Exit Sub
        End If
        
        
        'verifica se os campos do workflow foram preenchidos caso o campo Hora tenha sido preenchido
        If (LblHoraPrVisita.Text <> "__:__") Then
            If FunObrigatorioTXT(Autorizador, "Campo obrigatório") Then Exit Sub
             If FunObrigatorioDTP(DataPrVisita, "Campo obrigatório") Then Exit Sub
            If FunObrigatorioTXT(RespPrVisita, "Campo obrigatório") Then Exit Sub
        End If
        
        'verifica se os campos do workflow foram preenchidos caso o campo Responsável tenha sido preenchido
        If (RespIniTrabalhos.Text <> "") Then
            If FunObrigatorioTXT(Autorizador, "Campo obrigatório") Then Exit Sub
            If FunObrigatorioDTP(DataPrVisita, "Campo obrigatório") Then Exit Sub
            If (LblHoraPrVisita.Text = "__:__") Then
                MsgBox "Campo obrigatório", vbCritical, "CAMPO OBRIGATÓRIO"
                LblHoraPrVisita.SetFocus
                Exit Sub
            End If
        End If

'################################################################################################
        
        
'##################################INICIO TRABALHOS##########################################
       

        'verifica se os campos do workflow foram preenchidos caso o campo Data tenha sido preenchido
        If Not IsNull(DataIniTrabalhos.Value) Then
            If FunObrigatorioDTP(DataPrVisita, "Campo obrigatório") Then Exit Sub
            If (LblHoraIniTrabalhos.Text = "__:__") Then
                MsgBox "Campo obrigatório", vbCritical, "CAMPO OBRIGATÓRIO"
                LblHoraIniTrabalhos.SetFocus
                Exit Sub
            End If
            If FunObrigatorioTXT(RespIniTrabalhos, "Campo obrigatório") Then Exit Sub
        End If
        
        
        'verifica se os campos do workflow foram preenchidos caso o campo Hora tenha sido preenchido
        If (LblHoraIniTrabalhos.Text <> "__:__") Then
            If FunObrigatorioTXT(Autorizador, "Campo obrigatório") Then Exit Sub
             If FunObrigatorioDTP(DataIniTrabalhos, "Campo obrigatório") Then Exit Sub
            If FunObrigatorioTXT(RespIniTrabalhos, "Campo obrigatório") Then Exit Sub
        End If
        
        
        'verifica se os campos do workflow foram preenchidos caso o campo Responsável tenha sido preenchido
        If (RespIniTrabalhos.Text <> "") Then
            If FunObrigatorioTXT(Autorizador, "Campo obrigatório") Then Exit Sub
            If FunObrigatorioDTP(DataIniTrabalhos, "Campo obrigatório") Then Exit Sub
            If (LblHoraIniTrabalhos.Text = "__:__") Then
                MsgBox "Campo obrigatório", vbCritical, "CAMPO OBRIGATÓRIO"
                LblHoraIniTrabalhos.SetFocus
                Exit Sub
            End If
        End If
'################################################################################################
        
        
'#####################################CONCLUSAO######################################


        'verifica se os campos do workflow foram preenchidos caso o campo Data tenha sido preenchido
        If Not IsNull(DataConclusao.Value) Then
            If FunObrigatorioDTP(DataIniTrabalhos, "Campo obrigatório") Then Exit Sub
            If FunObrigatorioDTP(DataConclusao, "Campo obrigatório") Then Exit Sub
            If (LblHoraConclusao.Text = "__:__") Then
                MsgBox "Campo obrigatório", vbCritical, "CAMPO OBRIGATÓRIO"
                LblHoraConclusao.SetFocus
                Exit Sub
            End If
            If FunObrigatorioTXT(RespConclusao, "Campo obrigatório") Then Exit Sub
        End If
        
        
        'verifica se os campos do workflow foram preenchidos caso o campo Hora tenha sido preenchido
        If (LblHoraConclusao.Text <> "__:__") Then
            If FunObrigatorioTXT(Autorizador, "Campo obrigatório") Then Exit Sub
             If FunObrigatorioDTP(DataConclusao, "Campo obrigatório") Then Exit Sub
            If FunObrigatorioTXT(RespConclusao, "Campo obrigatório") Then Exit Sub
        End If
        
        
        'verifica se os campos do workflow foram preenchidos caso o campo Responsável tenha sido preenchido
        If (RespConclusao.Text <> "") Then
            If FunObrigatorioTXT(Autorizador, "Campo obrigatório") Then Exit Sub
            If FunObrigatorioDTP(DataConclusao, "Campo obrigatório") Then Exit Sub
            If (LblHoraConclusao.Text = "__:__") Then
                MsgBox "Campo obrigatório", vbCritical, "CAMPO OBRIGATÓRIO"
                LblHoraConclusao.SetFocus
                Exit Sub
            End If
        End If
        
'################################################################################################
        
        
        
        If (XInserir = "I") Then
            
            XGT_SQL = "SELECT  IsNull (MAX(cast (assi_cd_CodigoSeq as int)), 0) + 1 AS assi_cd_CodigoSeq From ConsCapAssistenciaTecnica WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_empreendimento = " & XFT_CODEMPREENDIMENTO & " AND imov_cd_imovel = " & XFT_CODIMOVEL & " AND cont_cd_contrato = " & XFT_CODCONTRATO & ""
            SubQOpenRecordset XLO_RS, XGT_SQL, "Estatico"
            If Not XLO_RS.EOF Then
            XLT_CODSEQUENCIAL = XLO_RS!assi_cd_CodigoSeq
            End If
            XLO_RS.Close
        End If
       
        
        If (XInserir = "I") Then
            XGT_SQL = "SELECT * FROM AssistenciaTecnica  WHERE assi_cd_Assistencia = -1 "
        SubQOpenRecordset XLO_RS, XGT_SQL, Dinamico
        
        If XLO_RS.EOF Then
            XLO_RS.AddNew
        End If
            
        XLO_RS!assi_dt_Data = Format(Data.Value, "dd/MM/yyyy")
        XLO_RS!assi_dt_DataEmissao = Format(Now, "dd/MM/yyyy HH:MM")
        'XLO_RS!assi_dt_DataEmissao = Format(LblHora.Caption, "HH:MM")
        XLO_RS!assi_tx_Ocorrencia = TxtOcorrencia.Text
        XLO_RS!assi_tx_Observacao = TxtObservacao.Text
        XLO_RS!cont_cd_Contrato = XFT_CODCONTRATO
        XLO_RS!empd_cd_Empreendimento = XFT_CODEMPREENDIMENTO
        XLO_RS!imov_cd_Imovel = XFT_CODIMOVEL
        XLO_RS!assi_cd_CodigoSeq = Format(XLT_CODSEQUENCIAL, "0#")
        XLO_RS!empr_cd_Empresa = PCodEmpresa
        XLS_DATAPRVISITA = DataPrVisita.Value & " " & LblHoraPrVisita.Text
        XLS_DATAINITRABALHO = DataIniTrabalhos.Value & " " & LblHoraIniTrabalhos.Text
        XLS_DATACONCLUSAO = DataConclusao.Value & " " & LblHoraConclusao.Text
        
        
        If Not IsNull(XLO_RS!assi_tx_Autorizador) Then
            XLO_RS!assi_tx_Autorizador = Autorizador.Text
        End If
        
        If IsNull(DataPrVisita.Value) Then
            XLO_RS!assi_dt_PrimeiraVisita = Null
        Else
            XLO_RS!assi_dt_PrimeiraVisita = Format(XLS_DATAPRVISITA, "dd/MM/yyyy HH:MM")
        End If
        'XLO_RS!assi_dt_PrimeiraVisita = Format(LblHoraPrVisita.Text, "HH:MM")
        If IsNull(DataIniTrabalhos.Value) Then
            XLO_RS!assi_dt_InicioTrabalho = Null
        Else
            XLO_RS!assi_dt_InicioTrabalho = Format(XLS_DATAINITRABALHO, "dd/MM/yyyy HH:MM")
        End If
        'XLO_RS!assi_dt_InicioTrabalho = Format(LblHoraIniTrabalhos.Text, "HH:MM")
        
        If IsNull(DataConclusao.Value) Then
            XLO_RS!assi_dt_Conclusao = Null
        Else
            XLO_RS!assi_dt_Conclusao = Format(XLS_DATACONCLUSAO, "dd/MM/yyyy HH:MM")
        End If

        
        If Not IsNull(XLO_RS!assi_tx_RespPrVisita) Then
            XLO_RS!assi_tx_RespPrVisita = RespPrVisita.Text
            'RespPrVisita.Text = XLO_RS!assi_tx_RespPrVisita
        End If
        
        If Not IsNull(XLO_RS!assi_tx_RespIniTrabalhos) Then
            XLO_RS!assi_tx_RespIniTrabalhos = RespIniTrabalhos.Text
            'RespIniTrabalhos.Text = XLO_RS!assi_tx_RespIniTrabalhos
        End If
        
        If Not IsNull(XLO_RS!assi_tx_RespConclusao) Then
            XLO_RS!assi_tx_RespConclusao = RespConclusao.Text
            'RespConclusao.Text = XLO_RS!assi_tx_RespConclusao
        End If
          
                                        
        XLO_RS.Update
        
        LblNumOrdem.Caption = CStr(XLO_RS!assi_cd_CodigoSeq)
         
        XLO_RS.Close
        Set XLO_RS = Nothing
        
        XInserir = "A"
        
    End If
    
    If (XInserir = "A") Then
    XGT_SQL = "SELECT * FROM AssistenciaTecnica WHERE empr_cd_empresa  = " & PCodEmpresa & " AND empd_cd_Empreendimento = " & XFT_CODEMPREENDIMENTO & " AND imov_cd_Imovel = " & XFT_CODIMOVEL & " AND assi_cd_CodigoSeq = " & LblNumOrdem.Caption & ""
    SubQOpenRecordset XLO_RS, XGT_SQL, Dinamico

    If XLO_RS.EOF Then
            XLO_RS.AddNew
        End If

    XLO_RS!assi_dt_Data = Format(Data.Value, "dd/MM/yyyy")
        XLO_RS!assi_dt_DataEmissao = Format(Now, "dd/MM/yyyy HH:MM")
        'XLO_RS!assi_dt_DataEmissao = Format(LblHora.Caption, "HH:MM")
        XLO_RS!assi_tx_Ocorrencia = TxtOcorrencia.Text
        XLO_RS!assi_tx_Observacao = TxtObservacao.Text
        XLO_RS!cont_cd_Contrato = XFT_CODCONTRATO
        XLO_RS!empd_cd_Empreendimento = XFT_CODEMPREENDIMENTO
        XLO_RS!imov_cd_Imovel = XFT_CODIMOVEL
        XLO_RS!assi_cd_CodigoSeq = LblNumOrdem.Caption
        XLO_RS!empr_cd_Empresa = PCodEmpresa
        XLS_DATAPRVISITA = DataPrVisita.Value & " " & LblHoraPrVisita.Text
        XLS_DATAINITRABALHO = DataIniTrabalhos.Value & " " & LblHoraIniTrabalhos.Text
        XLS_DATACONCLUSAO = DataConclusao.Value & " " & LblHoraConclusao.Text
        
        
        If Not IsNull(XLO_RS!assi_tx_Autorizador) Then
            XLO_RS!assi_tx_Autorizador = Autorizador.Text
        End If
        
        If IsNull(DataPrVisita.Value) Then
            XLO_RS!assi_dt_PrimeiraVisita = Null
        Else
            XLO_RS!assi_dt_PrimeiraVisita = Format(XLS_DATAPRVISITA, "dd/MM/yyyy HH:MM")
        End If
        'XLO_RS!assi_dt_PrimeiraVisita = Format(LblHoraPrVisita.Text, "HH:MM")
        If IsNull(DataIniTrabalhos.Value) Then
            XLO_RS!assi_dt_InicioTrabalho = Null
        Else
            XLO_RS!assi_dt_InicioTrabalho = Format(XLS_DATAINITRABALHO, "dd/MM/yyyy HH:MM")
        End If
        'XLO_RS!assi_dt_InicioTrabalho = Format(LblHoraIniTrabalhos.Text, "HH:MM")
        
        If IsNull(DataConclusao.Value) Then
            XLO_RS!assi_dt_Conclusao = Null
        Else
            XLO_RS!assi_dt_Conclusao = Format(XLS_DATACONCLUSAO, "dd/MM/yyyy HH:MM")
        End If

        
        If Not IsNull(XLO_RS!assi_tx_RespPrVisita) Then
            XLO_RS!assi_tx_RespPrVisita = RespPrVisita.Text
            'RespPrVisita.Text = XLO_RS!assi_tx_RespPrVisita
        End If
        
        If Not IsNull(XLO_RS!assi_tx_RespIniTrabalhos) Then
            XLO_RS!assi_tx_RespIniTrabalhos = RespIniTrabalhos.Text
            'RespIniTrabalhos.Text = XLO_RS!assi_tx_RespIniTrabalhos
        End If
        
        If Not IsNull(XLO_RS!assi_tx_RespConclusao) Then
            XLO_RS!assi_tx_RespConclusao = RespConclusao.Text
            'RespConclusao.Text = XLO_RS!assi_tx_RespConclusao
        End If
          
                                        
        XLO_RS.Update
        
        LblNumOrdem.Caption = CStr(XLO_RS!assi_cd_CodigoSeq)
         
        XLO_RS.Close
        Set XLO_RS = Nothing
        
        XInserir = "A"
    
    End If
    
    



        MsgBox "Operação Efetuada com Sucesso", vbInformation, "Aviso"
 
    Exit Sub
    
        
    

TrataErro:
    funTrataErros ComMensagem
End Sub

Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    Unload Me
End Sub

Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub



Private Sub Label19_Click()

End Sub

VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelEtiquetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiquetas e Envelopes"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "FrmRelEtiquetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraEnvelope 
      Caption         =   "Envelope"
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
      TabIndex        =   30
      Top             =   5160
      Width           =   5175
      Begin VB.OptionButton OptEnvelopeFrente 
         Caption         =   "Frente do Envelope "
         Height          =   315
         Left            =   540
         TabIndex        =   32
         Top             =   300
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton OptEnvelopeVerso 
         Caption         =   "Verso do Envelope"
         Height          =   315
         Left            =   2940
         TabIndex        =   31
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.Frame FraOrdem 
      Caption         =   "Ordenado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   27
      Top             =   3840
      Width           =   5175
      Begin VB.OptionButton OptCliente 
         Caption         =   "Cliente"
         Height          =   315
         Left            =   3180
         TabIndex        =   12
         Top             =   180
         Width           =   1155
      End
      Begin VB.OptionButton OptImovel 
         Caption         =   "Imóvel"
         Height          =   195
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4140
      TabIndex        =   14
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir..."
      Height          =   345
      Left            =   2880
      TabIndex        =   13
      Top             =   6000
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aniversario "
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
      Top             =   2280
      Width           =   5175
      Begin MSMask.MaskEdBox MskMesInicial 
         Height          =   315
         Left            =   1620
         TabIndex        =   7
         Top             =   300
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MskMesFinal 
         Height          =   315
         Left            =   3720
         TabIndex        =   8
         Top             =   300
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Mês Final:"
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
         Left            =   2820
         TabIndex        =   26
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mês Inicial:"
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
         Left            =   615
         TabIndex        =   25
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.Frame fraCliente 
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
      TabIndex        =   21
      Top             =   3060
      Width           =   5175
      Begin MSDataListLib.DataCombo CboCliente 
         Bindings        =   "FrmRelEtiquetas.frx":2AFA
         Height          =   315
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "focl_tx_RazaoSocial"
         BoundColumn     =   "focl_cd_FornCli"
         Text            =   "CboCliente"
      End
      Begin Threed.SSCommand CmdLimparCliente 
         Height          =   315
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   78
         MouseIcon       =   "FrmRelEtiquetas.frx":2B14
         Picture         =   "FrmRelEtiquetas.frx":2B30
      End
      Begin VB.Label LblCliente 
         Alignment       =   1  'Right Justify
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
         Left            =   60
         TabIndex        =   22
         Top             =   300
         Width           =   765
      End
   End
   Begin VB.Frame FraPeriodo 
      Caption         =   "Venda do Imóvel "
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
      TabIndex        =   18
      Top             =   660
      Width           =   5175
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   1155
         TabIndex        =   2
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   49479683
         CurrentDate     =   37749.5508333333
      End
      Begin MSComCtl2.DTPicker dtpFim 
         Height          =   315
         Left            =   3360
         TabIndex        =   3
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   49479683
         CurrentDate     =   37749.5510300926
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
         Left            =   2820
         TabIndex        =   20
         Top             =   330
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
         Left            =   615
         TabIndex        =   19
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.Frame FraContrato 
      Height          =   615
      Left            =   60
      TabIndex        =   15
      Top             =   0
      Width           =   5175
      Begin MSMask.MaskEdBox MskContratoInicial 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
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
         Left            =   3120
         TabIndex        =   1
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "AAAA.AAAA.##"
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
         Left            =   2940
         TabIndex        =   17
         Top             =   240
         Width           =   195
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
         Left            =   660
         TabIndex        =   16
         Top             =   240
         Width           =   825
      End
   End
   Begin Threed.SSFrame FraStatus 
      Height          =   720
      Left            =   60
      TabIndex        =   23
      Top             =   1500
      Width           =   5175
      _Version        =   65536
      _ExtentX        =   9128
      _ExtentY        =   1270
      _StockProps     =   14
      Caption         =   "Status do Cliente"
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
         Left            =   2160
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
         Left            =   660
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
         Left            =   3540
         TabIndex        =   6
         Top             =   300
         Width           =   1230
      End
   End
   Begin MSAdodcLib.Adodc DatClientes 
      Height          =   330
      Left            =   -60
      Top             =   5940
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
      Caption         =   "DatCliente"
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
   Begin VB.Frame FraFormulario 
      Caption         =   "Formulário"
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
      Top             =   4440
      Width           =   5175
      Begin VB.OptionButton OptEnvelopePequeno 
         Caption         =   "Envelope Pequeno"
         Height          =   315
         Left            =   1440
         TabIndex        =   34
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptEnvelopeGrande 
         Caption         =   "Envelope Grande"
         Height          =   315
         Left            =   3420
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptEtiqueta 
         Caption         =   "Etiqueta"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmRelEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
  
  Unload Me
  
End Sub

Private Sub CmdImprimir_Click()
  
   Dim XLT_FILTROPERIODO As String 'Armazenar filtro do período escolhido
   Dim XLT_FILTROCONTRATO As String 'Empreendimento escolhido
   Dim XLT_CLIENTE As String
   Dim XLT_FILTROSTATUS As String
   Dim XLT_STATUS  As String
   Dim XLT_STATUS2 As String
   
   Me.MousePointer = vbHourglass

   XGT_SELECAO = ""
   XGT_CONJUNCAO = ""
 
   'Período da venda
   If (Not IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
       "cont_dt_Venda >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & _
       " AND cont_dt_Venda <= " & FunNuloData(dtpFim.Value, NomeSgbd)
       XGT_CONJUNCAO = " AND "
   ElseIf (Not IsNull(dtpInicio.Value)) And (IsNull(dtpFim.Value)) Then
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
       "cont_dt_Venda >= " & FunNuloData(dtpInicio.Value, NomeSgbd)
       XGT_CONJUNCAO = " AND "
   ElseIf (IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
       XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
       "cont_dt_Venda <= " & FunNuloData(dtpFim.Value, NomeSgbd)
       XGT_CONJUNCAO = " AND "
   End If
   
  'Periodo de Aniversario
  XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
    " (DataAniver IS NULL OR " & _
    " month(DataAniver) >=" & MskMesInicial.Text & _
    " AND month(DataAniver) <=  " & MskMesFinal.Text & ")"
   XGT_CONJUNCAO = " AND "

  'Contrato
   If MskContratoInicial.Text <> "    .    .  " And MskContratoFinal.Text <> "    .    .  " Then
     XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
     "Contrato >= '" & MskContratoInicial.Text & "'" & _
     " AND Contrato <= '" & MskContratoFinal.Text & "'"
     XGT_CONJUNCAO = " AND "
     XLT_FILTROCONTRATO = "Contrato: " & MskContratoInicial & " a " & MskContratoFinal
   ElseIf MskContratoInicial.Text <> "    .    " And MskContratoFinal.Text = "    .    " Then
     XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
     "Contrato >= '" & MskContratoInicial.Text & "'"
     XGT_CONJUNCAO = " AND "
     XLT_FILTROCONTRATO = "Contrato: A partir de " & MskContratoInicial
   ElseIf MskContratoInicial.Text = "    .    " And MskContratoFinal.Text <> "    .    " Then
     XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
     "Contrato <= '" & MskContratoFinal.Text & "'"
     XGT_CONJUNCAO = " AND "
     XLT_FILTROCONTRATO = "Contrato: Até " & MskContratoInicial
   Else
     XLT_FILTROCONTRATO = "Contrato: Geral"
   End If
   
   'Cliente
   If CboCliente.BoundText <> "" Then
     XLT_CLIENTE = CboCliente.BoundText
     XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
     "focl_cd_FornCli = " & XLT_CLIENTE
     XGT_CONJUNCAO = " AND "
   End If
 
   'Status
   If ChkAtivo.Value = ChkInativo.Value And ChkAtivo.Value = ChkCedido.Value Then
       XLT_FILTROSTATUS = "Status: Geral"
       XLT_STATUS = "T" 'Todos
   Else
       If ChkAtivo.Value = 1 Then
           If ChkInativo.Value = 1 Then  'Ativo e Inativo
               XLT_STATUS = "ATIVO"
               XLT_STATUS2 = "INATIVO"
               XLT_FILTROSTATUS = "Status: ANTIVO e INATIVO"
           ElseIf ChkCedido.Value = 1 Then  'Ativo e Cedido
               XLT_STATUS = "ATIVO"
               XLT_STATUS2 = "CEDIDO"
               XLT_FILTROSTATUS = "Status: ATIVO e CEDIDO"
           Else
               XLT_STATUS = "ATIVO" 'Ativo
               XLT_FILTROSTATUS = "Status: ATIVO"
           End If
       ElseIf ChkInativo.Value = 1 Then
           If ChkCedido.Value = 1 Then   'Inativo e Cedido
               XLT_STATUS = "INATIVO"
               XLT_STATUS2 = "CEDIDO"
               XLT_FILTROSTATUS = "Status: INATIVO e CEDIDO"
           Else
               XLT_STATUS = "INATIVO" 'Inativo
               XLT_FILTROSTATUS = "Status: INATIVO"
           End If
       ElseIf ChkCedido.Value = 1 Then
           XLT_STATUS = "CEDIDO" 'Cedido
           XLT_FILTROSTATUS = "Status: CEDIDO"
       End If
       If XLT_STATUS2 <> "" Then 'FoRAM escolhidos dois status
           XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "(Status = '" & XLT_STATUS & "'" & _
           " OR Status = '" & XLT_STATUS2 & "')"
           XGT_CONJUNCAO = " AND "
       Else
           XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "Status = '" & XLT_STATUS & "'"
           XGT_CONJUNCAO = " AND "
       End If
   End If
    
   'Filtra a empresa Ativa
   XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "empr_cd_Empresa = " & CStr(PCodEmpresa)
   
   'Se for só um cliente não ordena pois não aceita o SELECT DISTINCT
'   If CboCliente.BoundText = "" Then
'     'Ordena
'     If OptCliente = True Then
'       XGT_SELECAO = XGT_SELECAO & _
'       " GROUP BY focl_cd_fornCli, focl_tx_razaoSocial, clie_tx_endcorresp, clie_tx_bairrocorresp , clie_tx_muncorresp, clie_tx_estcorresp, clie_nr_cepcorresp" & _
'       " ORDER BY focl_tx_RazaoSocial"
'     Else
'       XGT_SELECAO = XGT_SELECAO & " ORDER BY left(Contrato,9)"
'     End If
'   End If
    
    If OptEnvelopeVerso.Value = False And (OptEnvelopePequeno.Value = True Or OptEnvelopeGrande.Value = True) Then
      If OptCliente = True Then
        XGT_SELECAO = XGT_SELECAO & " ORDER BY focl_tx_RazaoSocial"
      Else
        XGT_SELECAO = XGT_SELECAO & " ORDER BY Contrato"
      End If
    End If
    
   'Abre os recordsets e verifica se eles possuem dados
   'Se for só um cliente usa o SELECT DISTINCT para não repetir se o cliente tiver mais de um contrato
'   If CboCliente.BoundText <> "" Then
'     Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT DISTINCT focl_cd_fornCli, focl_tx_razaoSocial, clie_tx_endcorresp, clie_tx_bairrocorresp , clie_tx_muncorresp, clie_tx_estcorresp, clie_nr_cepcorresp FROM ConsCapRelEtiquetas WHERE " & XGT_SELECAO)
'   Else
'     Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT focl_cd_fornCli, focl_tx_razaoSocial, clie_tx_endcorresp, clie_tx_bairrocorresp , clie_tx_muncorresp, clie_tx_estcorresp, clie_nr_cepcorresp FROM ConsCapRelEtiquetas WHERE " & XGT_SELECAO)
'   End If

    If OptEnvelopeVerso.Value = True And (OptEnvelopePequeno.Value = True Or OptEnvelopeGrande.Value = True) Then
      Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT DISTINCT focl_cd_fornCli, empr_tx_RazaoSocial, empr_tx_Endereco, empr_tx_Bairro, empr_tx_Cidade, empr_tx_Estado, empr_tx_Cep FROM ConsCapRelEtiquetas WHERE " & XGT_SELECAO)
    Else
      If OptCliente = True Then
        Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT DISTINCT focl_cd_fornCli, focl_tx_razaoSocial, clie_tx_endcorresp, clie_tx_bairrocorresp , clie_tx_muncorresp, clie_tx_estcorresp, clie_nr_cepcorresp FROM ConsCapRelEtiquetas WHERE " & XGT_SELECAO)
      Else
        Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT focl_cd_fornCli, focl_tx_razaoSocial, clie_tx_endcorresp, clie_tx_bairrocorresp , clie_tx_muncorresp, clie_tx_estcorresp, clie_nr_cepcorresp, Contrato FROM ConsCapRelEtiquetas WHERE " & XGT_SELECAO)
      End If
    End If
    
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
   
   'Carrega os relatórios
   If OptEtiqueta.Value = True Then
      FrmVisRelatorios.Caption = "Etiquetas"
      Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelEtiquetas.rpt")
   Else
     If OptEnvelopeFrente.Value = True Then
       FrmVisRelatorios.Caption = "Frente do Envelope"
       If OptEnvelopeGrande.Value = True Then
         Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelEnvelopeGrande.rpt")
       Else
         Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelEnvelope.rpt")
       End If
     Else
       FrmVisRelatorios.Caption = "Verso do Envelope"
       If OptEnvelopeGrande.Value = True Then
         Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelEnvelopeVersoGrande.rpt")
       Else
         Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelEnvelopeVerso.rpt")
       End If
     End If
     
   End If
 
   'Passa o recordset para o relatório
   XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
   
  If OptEnvelopeFrente.Value = True And (OptEnvelopeGrande.Value = True Or OptEnvelopePequeno.Value = True) Then
    With XGR_RELATORIO.FormulaFields
      .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  End If
   
   
   FrmVisRelatorios.SubVerRelatorioSemMargem XGR_RELATORIO, ""
   
   Me.MousePointer = vbDefault
  
End Sub

Private Sub CmdLimparCliente_Click()
  
  CboCliente.BoundText = ""
  
End Sub

Private Sub Form_Load()
  
  subConectarControleDadosNV DatClientes, "SELECT focl_cd_forncli,focl_tx_classe,focl_tx_razaoSocial FROM FornClientes WHERE focl_tx_classe = 'C' ORDER BY focl_tx_razaoSocial", Estatico
  MskMesInicial.Text = "01"
  MskMesFinal.Text = "12"
  FraEnvelope.Enabled = False
  OptEnvelopeFrente.Enabled = False
  OptEnvelopeVerso.Enabled = False
  
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

Private Sub MskMesInicial_GotFocus()
  
  subSelecionaMSK MskMesInicial
  
End Sub

Private Sub MskMesFinal_GotFocus()
  
  subSelecionaMSK MskMesFinal
  
End Sub

Private Sub MskMesInicial_LostFocus()
  
  If CInt(MskMesInicial.Text) < 1 Or CInt(MskMesInicial.Text) > 12 Then
    MskMesInicial.Text = "01"
  Else
    MskMesInicial.Text = Format(MskMesInicial.Text, "00")
  End If
  
End Sub

Private Sub MskMesFinal_LostFocus()
  
  If CInt(MskMesFinal.Text) < 1 Or CInt(MskMesFinal.Text) > 12 Then
    MskMesFinal.Text = "12"
  Else
    MskMesFinal.Text = Format(MskMesFinal.Text, "00")
  End If
  
End Sub

Private Sub OptEtiqueta_Click()
  
  FraOrdem.Enabled = True
  OptImovel.Enabled = True
  OptCliente.Enabled = True
  
  FraEnvelope.Enabled = False
  OptEnvelopeFrente.Enabled = False
  OptEnvelopeVerso.Enabled = False
  
End Sub

Private Sub OptEnvelopeGrande_Click()
  
  FraEnvelope.Enabled = True
  OptEnvelopeFrente.Enabled = True
  OptEnvelopeVerso.Enabled = True
  
  If OptEnvelopeVerso = False Then
    FraOrdem.Enabled = True
    OptImovel.Enabled = True
    OptCliente.Enabled = True
  End If
  
End Sub

Private Sub OptEnvelopePequeno_Click()
  
  FraEnvelope.Enabled = True
  OptEnvelopeFrente.Enabled = True
  OptEnvelopeVerso.Enabled = True
  
  If OptEnvelopeVerso = False Then
    FraOrdem.Enabled = True
    OptImovel.Enabled = True
    OptCliente.Enabled = True
  End If
  
End Sub

Private Sub OptEnvelopeFrente_Click()
  
  FraOrdem.Enabled = True
  OptImovel.Enabled = True
  OptCliente.Enabled = True
  
End Sub

Private Sub OptEnvelopeVerso_Click()

  FraOrdem.Enabled = False
  OptImovel.Enabled = False
  OptCliente.Enabled = False

End Sub

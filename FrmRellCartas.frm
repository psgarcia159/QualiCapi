VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelCartas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartas Especiais"
   ClientHeight    =   4935
   ClientLeft      =   2730
   ClientTop       =   3540
   ClientWidth     =   4890
   Icon            =   "FrmRellCartas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Aniversário do Cliente"
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
      TabIndex        =   23
      Top             =   3600
      Width           =   4770
      Begin MSComCtl2.DTPicker dtpAniverIni 
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   48300035
         CurrentDate     =   37749.5511921296
      End
      Begin MSComCtl2.DTPicker dtpAniverFim 
         Height          =   315
         Left            =   3060
         TabIndex        =   12
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   48300035
         CurrentDate     =   37749.5513078704
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
         TabIndex        =   25
         Top             =   330
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
         Left            =   2520
         TabIndex        =   24
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.Frame FraTipoCarta 
      Caption         =   "Tipo de Carta"
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
      TabIndex        =   19
      Top             =   0
      Width           =   4755
      Begin Threed.SSCommand CmdLimpar 
         Height          =   315
         Left            =   4320
         TabIndex        =   1
         Top             =   240
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   78
         MouseIcon       =   "FrmRellCartas.frx":2AFA
         Picture         =   "FrmRellCartas.frx":2B16
      End
      Begin MSDataListLib.DataCombo CboCartas 
         Bindings        =   "FrmRellCartas.frx":2C28
         Height          =   315
         Left            =   60
         TabIndex        =   0
         Top             =   240
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "cart_tx_Descricao"
         BoundColumn     =   "cart_cd_carta"
         Text            =   "cboCarta"
      End
   End
   Begin VB.Frame FraPeriodo 
      Caption         =   " Venda do Imóvel "
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
      TabIndex        =   15
      Top             =   1320
      Width           =   4770
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   855
         TabIndex        =   4
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   48300035
         CurrentDate     =   37749.5508333333
      End
      Begin MSComCtl2.DTPicker dtpFim 
         Height          =   315
         Left            =   3060
         TabIndex        =   5
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yy"
         DateIsNull      =   -1  'True
         Format          =   48300035
         CurrentDate     =   37749.5510300926
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
         Left            =   315
         TabIndex        =   17
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
         Left            =   2520
         TabIndex        =   16
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   3780
      TabIndex        =   14
      Top             =   4500
      Width           =   1035
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Enabled         =   0   'False
      Height          =   345
      Left            =   2580
      TabIndex        =   13
      Top             =   4500
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc DatCartas 
      Height          =   330
      Left            =   0
      Top             =   4500
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Caption         =   "Cartas"
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
   Begin VB.Frame fraCliente 
      Caption         =   "Cliente"
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
      TabIndex        =   18
      Top             =   2100
      Width           =   4755
      Begin MSDataListLib.DataCombo CboCliente 
         Bindings        =   "FrmRellCartas.frx":2C40
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
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
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   78
         MouseIcon       =   "FrmRellCartas.frx":2C5A
         Picture         =   "FrmRellCartas.frx":2C76
      End
   End
   Begin MSAdodcLib.Adodc DatClientes 
      Height          =   330
      Left            =   0
      Top             =   4500
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
   Begin Threed.SSFrame FraEmpreendimento 
      Height          =   600
      Left            =   60
      TabIndex        =   20
      Top             =   660
      Width           =   4755
      _Version        =   65536
      _ExtentX        =   8387
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
      Begin MSMask.MaskEdBox MskImovelInicial 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
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
         Left            =   2940
         TabIndex        =   3
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "AAAA.AAAA"
         PromptChar      =   " "
      End
      Begin VB.Label LblImovel 
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
         Height          =   195
         Left            =   720
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
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
         Left            =   2640
         TabIndex        =   21
         Top             =   240
         Width           =   195
      End
   End
   Begin Threed.SSFrame FraStatus 
      Height          =   720
      Left            =   60
      TabIndex        =   22
      Top             =   2880
      Width           =   4755
      _Version        =   65536
      _ExtentX        =   8387
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
         Left            =   3480
         TabIndex        =   10
         Top             =   300
         Value           =   1  'Checked
         Width           =   810
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
         TabIndex        =   8
         Top             =   300
         Value           =   1  'Checked
         Width           =   1065
      End
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
         Left            =   1980
         TabIndex        =   9
         Top             =   300
         Value           =   1  'Checked
         Width           =   1080
      End
   End
End
Attribute VB_Name = "FrmRelCartas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboCartas_Change()
  If Not IsNull(CboCartas.SelectedItem) Then
    DatCartas.Recordset.bookmark = CboCartas.SelectedItem
    CmdImprimir.Enabled = True
  End If
End Sub

Private Sub CboCliente_Click(Area As Integer)
    If CboCliente.Text <> "" Then
        DatClientes.Recordset.bookmark = CboCliente.SelectedItem
    End If
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    
    Dim XLT_EMPREENDIMENTO As String 'Empreendimento escolhido
    Dim XLT_CLIENTE As String
    Dim XLT_STATUS  As String
    Dim XLT_STATUS2 As String
    Dim XLT_FILTROSTATUS As String
    Dim XLT_CARTA As String
    Dim XLO_RSCARTA As New ADODB.Recordset
    Dim XFO_RSCLIENTESIMOVEIS As New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
 
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
  
    XLT_CARTA = CboCartas.BoundText
    XGT_SELECAO = " WHERE cart_cd_carta = " & XLT_CARTA
    XGT_CONJUNCAO = " AND "
    
    'Período da venda
    If (Not IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
        XGT_CONJUNCAO = " AND "
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "cont_dt_Venda >= " & FunNuloData(dtpInicio.Value, NomeSgbd) & _
        " AND cont_dt_Venda <= " & FunNuloData(dtpFim.Value, NomeSgbd)
    ElseIf (Not IsNull(dtpInicio.Value)) And (IsNull(dtpFim.Value)) Then
        XGT_CONJUNCAO = " AND "
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "cont_dt_Venda >= " & FunNuloData(dtpInicio.Value, NomeSgbd)
    ElseIf (IsNull(dtpInicio.Value)) And (Not IsNull(dtpFim.Value)) Then
        XGT_CONJUNCAO = " AND "
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "cont_dt_Venda <= " & FunNuloData(dtpFim.Value, NomeSgbd)
    End If
    
    'Periodo de Aniversario
   If (Not IsNull(dtpAniverIni.Value)) And (Not IsNull(dtpAniverFim.Value)) Then
        XGT_CONJUNCAO = " AND "
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "DataAniver >=  " & FunNuloData(dtpAniverIni.Value, NomeSgbd) & _
        " AND DataAniver <=  " & FunNuloData(dtpAniverFim.Value, NomeSgbd)
    ElseIf (Not IsNull(dtpAniverIni.Value)) And (IsNull(dtpAniverFim.Value)) Then
        XGT_CONJUNCAO = " AND "
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "DataAniver >= " & FunNuloData(dtpAniverIni.Value, NomeSgbd)
    ElseIf (IsNull(dtpAniverIni.Value)) And (Not IsNull(dtpAniverFim.Value)) Then
        XGT_CONJUNCAO = " AND "
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
        "DataAniver <= " & FunNuloData(dtpAniverFim.Value, NomeSgbd)
    End If
    
    'Imóvel
    If MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text <> "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "Left(Contrato,9) >= '" & MskImovelInicial.Text & "'" & _
      " AND Left(Contrato,9) <= '" & MskImovelFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
    ElseIf MskImovelInicial.Text <> "    .    " And MskImovelFinal.Text = "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "Left(Contrato,9) >= '" & MskImovelInicial.Text & "'"
      XGT_CONJUNCAO = " AND "
    ElseIf MskImovelInicial.Text = "    .    " And MskImovelFinal.Text <> "    .    " Then
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      "Left(Contrato,9) <= '" & MskImovelFinal.Text & "'"
      XGT_CONJUNCAO = " AND "
    End If
    
    'Cliente
    If CboCliente.BoundText <> "" Then
      XGT_CONJUNCAO = " AND "
      XLT_CLIENTE = CboCliente.BoundText
      XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
      " focl_cd_FornCli = " & XLT_CLIENTE
      
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
    
    'Filtra a empresa
    XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "empr_cd_Empresa = " & CStr(PCodEmpresa)
        
    'Ordena
    XGT_SELECAO = XGT_SELECAO & " ORDER BY focl_tx_RazaoSocial"
  
    'Abre os recordsets e verifica se eles possuem dados
    Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT DISTINCT focl_cd_forncli, focl_tx_RazaoSocial, cart_cd_Carta, empr_tx_Cidade, empd_cd_Empreendimento, empd_tx_Nome FROM ConsCAPRelCartas2" & XGT_SELECAO)
    
    If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        FrmRelComissoes.MousePointer = vbDefault
        Exit Sub
    End If
    Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
    
    'Fecha o recordset se estiver aberto
    If XLO_RSCARTA.State = adStateOpen Then
      XLO_RSCARTA.Close
      Set XLO_RSCARTA = Nothing
    End If
    Set XLO_RSCARTA = ConexaoRelatorio.Execute("SELECT * FROM Cartas WHERE cart_cd_carta = " & XLT_CARTA)
    
    'Fecha o recordset se estiver aberto
    If XFO_RSCLIENTESIMOVEIS.State = adStateOpen Then
      XFO_RSCLIENTESIMOVEIS.Close
      Set XFO_RSCLIENTESIMOVEIS = Nothing
    End If
    Set XFO_RSCLIENTESIMOVEIS = ConexaoRelatorio.Execute("SELECT * FROM ConsCapRelClientesImoveis WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY imov_cd_Imovel")
    
    'Carrega os relatórios
    FrmVisRelatorios.Caption = "Emissão de Carta"
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\Carta.rpt")
  
    'Passa o recordset para o relatório
    XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
    'Ativa os Relatórios
    With XGR_RELATORIO.FormulaFields
        .GetItemByName("CodCarta").Text = Chr(34) & XLT_CARTA & Chr(34)
        .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    End With
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Rodape"
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XLO_RSCARTA, "Carta"
    FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XFO_RSCLIENTESIMOVEIS, "ClientesImoveis"
    FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
    Me.MousePointer = vbDefault
End Sub

Private Sub CmdLimpar_Click()
    
    CboCartas.Text = ""
    CboCartas.BoundText = ""
    CmdImprimir.Enabled = False
    
End Sub

Private Sub CmdLimparCliente_Click()
  
  CboCliente.BoundText = ""
  
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelComissoes
End Sub

Private Sub Form_Load()
  subManutencaoJanelasAtivas "I", "FrmRelCartas"
  subConectarControleDadosNV DatClientes, "SELECT focl_cd_forncli,focl_tx_classe,focl_tx_razaoSocial FROM FornClientes WHERE focl_tx_classe = 'C' ORDER BY focl_tx_razaoSocial", Estatico
  subConectarControleDadosNV DatCartas, "SELECT * FROM Cartas", Estatico
End Sub

Private Sub MskImovelInicial_GotFocus()
  
  Call subSelecionaMSK(MskImovelInicial)
  
End Sub

Private Sub MskImovelFinal_GotFocus()
  
  Call subSelecionaMSK(MskImovelFinal)
  
End Sub

Private Sub MskImovelInicial_LostFocus()
   
  If MskImovelInicial.Text <> "    .    " Then
    If Len(Trim(MskImovelInicial.Text)) < 9 Then
      MsgBox "O código inicial do imóvel deve possuir 8 dígitos, ou deve estar vazio.", vbCritical, "ATENÇÃO !"
      MskImovelInicial.SetFocus
    Else
      If MskImovelFinal.Text <> "    .    " Then
        subConectarControleDadosNV DatClientes, "SELECT focl_cd_FornCli,focl_tx_RazaoSocial FROM ConsCAPImoveisPorCliente " & _
          " WHERE empr_cd_Empresa= " & PCodEmpresa & _
          " AND left(Contrato,9) >= '" & MskImovelInicial.Text & "'" & _
          " AND left(Contrato,9) <= '" & MskImovelFinal.Text & "'" & _
          " ORDER BY focl_tx_RazaoSocial", Estatico
      End If
    End If
  Else
    subConectarControleDadosNV DatClientes, "SELECT focl_cd_FornCli,focl_tx_RazaoSocial " & _
      "FROM ConsCAPImoveisPorCliente WHERE empr_cd_Empresa= " & PCodEmpresa & " ORDER BY focl_tx_RazaoSocial", Estatico
  End If
  
End Sub

Private Sub MskImovelFinal_LostFocus()
   
  If MskImovelFinal.Text <> "    .    " Then
    If Len(Trim(MskImovelFinal.Text)) < 9 Then
      MsgBox "O código final do imóvel deve possuir 8 dígitos, ou deve estar vazio", vbCritical, "ATENÇÃO !"
      MskImovelFinal.SetFocus
    Else
      If MskImovelInicial.Text <> "    .    " Then
        subConectarControleDadosNV DatClientes, "SELECT focl_cd_FornCli,focl_tx_RazaoSocial FROM ConsCAPImoveisPorCliente " & _
          " WHERE empr_cd_Empresa= " & PCodEmpresa & _
          " AND left(Contrato,9) >= '" & MskImovelInicial.Text & "'" & _
          " AND left(Contrato,9) <= '" & MskImovelFinal.Text & "'" & _
          " ORDER BY focl_tx_RazaoSocial", Estatico
      End If
    End If
  Else
    subConectarControleDadosNV DatClientes, "SELECT focl_cd_FornCli,focl_tx_RazaoSocial " & _
      "FROM ConsCAPImoveisPorCliente WHERE empr_cd_Empresa= " & PCodEmpresa & " ORDER BY focl_tx_RazaoSocial", Estatico
  End If
  
End Sub

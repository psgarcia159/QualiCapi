VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmFinAlteraQualifin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualização de Títulos e Comissões no Qualifin"
   ClientHeight    =   7050
   ClientLeft      =   3240
   ClientTop       =   2805
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   4755
   Begin Threed.SSFrame SSFrame1 
      Height          =   600
      Left            =   120
      TabIndex        =   20
      Top             =   6200
      Width           =   4500
      _Version        =   65536
      _ExtentX        =   7937
      _ExtentY        =   1058
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   180
         Visible         =   0   'False
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.OptionButton optAtualizarDataDeposito 
      Caption         =   "Atualizar por Data de Depósito"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   2655
   End
   Begin VB.OptionButton optAtualizarPorVencimento 
      Caption         =   "Atualizar por Vencimento"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Value           =   -1  'True
      Width           =   2655
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   2685
      TabIndex        =   1
      Top             =   5805
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
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
   Begin Threed.SSCommand CmdAtualizar 
      Height          =   330
      Left            =   1080
      TabIndex        =   0
      Top             =   5805
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Atualizar"
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
   Begin Threed.SSFrame FraPeriodo 
      Height          =   1380
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4500
      _Version        =   65536
      _ExtentX        =   7937
      _ExtentY        =   2434
      _StockProps     =   14
      Caption         =   "Vencimento"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox ChkExpComissoes 
         Caption         =   "Exportar pagamento de comissões a corretores"
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   840
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker DtpInicio 
         Height          =   330
         Left            =   795
         TabIndex        =   3
         Top             =   330
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   85393409
         CurrentDate     =   37180
      End
      Begin MSComCtl2.DTPicker DtpFim 
         Height          =   330
         Left            =   2835
         TabIndex        =   4
         Top             =   330
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   85393409
         CurrentDate     =   37180
      End
      Begin VB.Label LblDtFinal 
         Alignment       =   1  'Right Justify
         Caption         =   "Fim:"
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
         Left            =   2010
         TabIndex        =   6
         Top             =   375
         Width           =   795
      End
      Begin VB.Label LblDtInicio 
         Alignment       =   1  'Right Justify
         Caption         =   "Início:"
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
         Left            =   135
         TabIndex        =   5
         Top             =   375
         Width           =   630
      End
   End
   Begin Threed.SSFrame FraTitulo 
      Height          =   1950
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   4500
      _Version        =   65536
      _ExtentX        =   7937
      _ExtentY        =   3440
      _StockProps     =   14
      Caption         =   "Intervalo de Títulos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox ChkExportar 
         Caption         =   "Exportar apenas títulos não pagos."
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin MSDataListLib.DataCombo CboTitulosInicio 
         Bindings        =   "FrmAlterarQualifin.frx":0000
         Height          =   330
         Left            =   945
         TabIndex        =   9
         Top             =   405
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "titulo"
         BoundColumn     =   "titulo"
         Text            =   "CboTitulos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo CboTitulosFim 
         Bindings        =   "FrmAlterarQualifin.frx":001F
         Height          =   330
         Left            =   945
         TabIndex        =   10
         Top             =   855
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "titulo"
         BoundColumn     =   "titulo"
         Text            =   "CboTitulos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand CmdLimparInicio 
         Height          =   330
         Left            =   3780
         TabIndex        =   11
         Top             =   405
         Width           =   360
         _Version        =   65536
         _ExtentX        =   635
         _ExtentY        =   582
         _StockProps     =   78
         MouseIcon       =   "FrmAlterarQualifin.frx":003B
         Picture         =   "FrmAlterarQualifin.frx":048D
      End
      Begin Threed.SSCommand CmdLimparFim 
         Height          =   330
         Left            =   3780
         TabIndex        =   12
         Top             =   855
         Width           =   360
         _Version        =   65536
         _ExtentX        =   635
         _ExtentY        =   582
         _StockProps     =   78
         MouseIcon       =   "FrmAlterarQualifin.frx":059F
         Picture         =   "FrmAlterarQualifin.frx":09F1
      End
      Begin VB.Label LlbFinal 
         Alignment       =   1  'Right Justify
         Caption         =   "Final:"
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
         Left            =   270
         TabIndex        =   14
         Top             =   915
         Width           =   645
      End
      Begin VB.Label LblInicial 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicial:"
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
         Left            =   225
         TabIndex        =   13
         Top             =   480
         Width           =   690
      End
   End
   Begin MSAdodcLib.Adodc DatTitulosInicio 
      Height          =   330
      Left            =   480
      Top             =   6840
      Visible         =   0   'False
      Width           =   3780
      _ExtentX        =   6668
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
      Caption         =   "DatTitulos"
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
   Begin MSAdodcLib.Adodc DatTitulosFim 
      Height          =   330
      Left            =   480
      Top             =   7200
      Visible         =   0   'False
      Width           =   3780
      _ExtentX        =   6668
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
      Caption         =   "DatTitulos"
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
   Begin Threed.SSFrame SSFrame 
      Height          =   900
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   4500
      _Version        =   65536
      _ExtentX        =   7937
      _ExtentY        =   1587
      _StockProps     =   14
      Caption         =   "Depósito"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComCtl2.DTPicker dtpDeposito 
         Height          =   330
         Left            =   240
         TabIndex        =   16
         Top             =   330
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         DateIsNull      =   -1  'True
         Format          =   85393409
         CurrentDate     =   37180
      End
   End
End
Attribute VB_Name = "FrmFinAlteraQualifin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XLB_ERRO      As Boolean      ' Status de erro para teste entre subrotinas

Option Explicit

Private Sub CmdDesistir_Click()
    Unload Me
End Sub

Private Sub CmdAtualizar_Click()
    
    Dim XLO_TITULO As ADODB.Recordset
    Dim XLO_QUALIFIN As ADODB.Recordset
    Dim XLT_SQL As String
    Dim XLI_RESPOSTA As Integer
    Dim XLB_COMPROMISSADO As Boolean
    Dim XLO_CONTABILIDADE As ADODB.Recordset
    Dim XLO_TOTALPROMISSORIAS As ADODB.Recordset
    Dim XLO_COPA As ADODB.Recordset
    Dim XLT_SQL2 As String
    
    On Error GoTo RotuloErro
    
    XLI_RESPOSTA = 0
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
        
    XLT_SQL2 = ""
        
    If (optAtualizarPorVencimento.Value) Then
        'Período de Vencimento
        If (Not IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
                "titu_dt_Prorrogacao >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & "" & _
                " AND titu_dt_Prorrogacao <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ""
            XGT_CONJUNCAO = " AND "
        ElseIf (Not IsNull(DtpInicio.Value)) And (IsNull(DtpFim.Value)) Then
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
                "titu_dt_Prorrogacao >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & ""
            XGT_CONJUNCAO = " AND "
        ElseIf (IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
            XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
                "titu_dt_Prorrogacao <= " & FunNuloData(DtpFim.Value, NomeSgbd) & ""
            XGT_CONJUNCAO = " AND "
        End If
        
        If ChkExpComissoes.Value = 1 Then
            
            XLT_SQL2 = "SELECT * FROM CorretoresPagamento" & _
                       " WHERE empr_cd_empresa = " & PCodEmpresa & _
                       "   AND copa_tx_statusPagamento = 'P'"
                    
            If (Not IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
                XLT_SQL2 = XLT_SQL2 & _
                           "   AND copa_dt_vencimento >= " & FunNuloData(DtpInicio.Value, NomeSgbd) & _
                           "   AND copa_dt_vencimento <= " & FunNuloData(DtpFim.Value, NomeSgbd)
            ElseIf (Not IsNull(DtpInicio.Value)) And (IsNull(DtpFim.Value)) Then
                XLT_SQL2 = XLT_SQL2 & _
                           "   AND copa_dt_vencimento >= " & FunNuloData(DtpInicio.Value, NomeSgbd)
            ElseIf (IsNull(DtpInicio.Value)) And (Not IsNull(DtpFim.Value)) Then
                XLT_SQL2 = XLT_SQL2 & _
                           "   AND copa_dt_vencimento <= " & FunNuloData(DtpFim.Value, NomeSgbd)
            End If
                
            XLT_SQL2 = XLT_SQL2 & " ORDER BY empr_cd_Empresa, empd_cd_Empreendimento, imov_cd_Imovel, cont_cd_Contrato, corr_cd_Corretor"

        End If
        
    Else
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "titu_dt_Deposito = " & FunNuloData(dtpDeposito.Value, NomeSgbd)
        XGT_CONJUNCAO = " AND "
    End If
        
    'Intervalo de Títulos
    If CboTitulosInicio.Text <> "" And CboTitulosFim.Text <> "" Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
            "titulo >= '" & CboTitulosInicio.Text & "'" & _
            " AND titulo <= '" & CboTitulosFim.Text & "'"
        XGT_CONJUNCAO = " AND "
    ElseIf CboTitulosInicio.Text <> "" And CboTitulosFim.Text = "" Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
            "titulo >= '" & CboTitulosInicio.Text & "'"
        XGT_CONJUNCAO = " AND "
    ElseIf CboTitulosInicio.Text = "" And CboTitulosFim.Text <> "" Then
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & _
            "(titulo <= '" & CboTitulosFim.Text & "')"
        XGT_CONJUNCAO = " AND "
    End If
    
    If ChkExportar.Value = 1 Then
        SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCapAtualizaQualifn WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & XGT_SELECAO & XGT_CONJUNCAO & "titu_dt_pagamento is null", Estatico
    ElseIf optAtualizarPorVencimento.Value Then
        If MsgBox("Os títulos realizados (pagos) passarão a ser compromissados.Deseja realmente executar a atualização?", vbYesNo, "Atualizar QualiFin") = vbYes Then
            SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCapAtualizaQualifn WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & XGT_SELECAO & " ORDER BY coco_cd_codigo", Estatico
        Else
            Exit Sub
        End If
    Else
        SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCapAtualizaQualifn WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & XGT_SELECAO & " ORDER BY coco_cd_codigo", Estatico
    End If

    ' Habilita a barra de progresso
    If XLO_TITULO.RecordCount > 0 Then
    
'       SSFrame1.Caption = "Exportando Títulos:"
       ProgressBar1.Visible = True
       ProgressBar1.Min = 0
       ProgressBar1.Max = XLO_TITULO.RecordCount
       ProgressBar1.Value = 0
       
    End If
     
    Conexao.BeginTrans
     
    While Not XLO_TITULO.EOF
    
        ProgressBar1.Value = ProgressBar1.Value + 1
    
        SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM NotasFiscais WHERE empr_cd_empresa = " & PCodEmpresa & " AND nofi_nr_TituloCapi = '" & XLO_TITULO!titulo & "'", Estatico
                
        If Not XLO_QUALIFIN.EOF Then
            XLB_COMPROMISSADO = (XLO_QUALIFIN!nofi_tx_status = "C")
            XLO_QUALIFIN.Close
            Set XLO_QUALIFIN = Nothing
            
            If (IsNull(XLO_TITULO!titu_dt_Pagamento)) Then
                subGravaFinanceiro XLO_TITULO, "COMPROMISSAR", XLO_TITULO!focl_cd_forncli, XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!NATUREZA
            ElseIf (optAtualizarPorVencimento.Value) Then
                If (XLB_COMPROMISSADO) Then
                    subGravaFinanceiro XLO_TITULO, "PAGAR", XLO_TITULO!focl_cd_forncli, XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!NATUREZA
                End If
            ElseIf (optAtualizarDataDeposito.Value) Then
                
                If XGB_PQUALICONT Then
                   If funMesFechado(dtpDeposito) <> "A" Then
                       MsgBox "Não é possível executar a operação, pois o mês contábil já foi fechado. " _
                               + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                        XLO_QUALIFIN.Close
                       Exit Sub
                   End If
                End If
                                
                Dim COD_DUPLICATA As Long
                Dim COD_CONTACONTABIL As Long
                Dim VALORPAGO As Long
                
                XLT_SQL = "SELECT duplicatas.DUPL_CD_DUPLICATA, ContasCorrente.plco_cd_conta " & _
                          "  FROM NotasFiscais " & _
                          " INNER JOIN duplicatas on duplicatas.dupl_nr_fatura = notasfiscais.nofi_nr_fatura " & _
                          " INNER JOIN ContasCorrente ON ContasCorrente.coco_cd_codigo = Duplicatas.coco_cd_codigo " & _
                          " WHERE NotasFiscais.empr_cd_empresa = " & CStr(PCodEmpresa) & _
                          "   AND NotasFiscais.nofi_nr_titulocapi = '" & XLO_TITULO!titulo & "'"

                SubQOpenRecordset XLO_CONTABILIDADE, XLT_SQL, Estatico
                If Not XLO_CONTABILIDADE.EOF Then
                    
                    COD_DUPLICATA = XLO_CONTABILIDADE!dupl_cd_duplicata
                    COD_CONTACONTABIL = XLO_CONTABILIDADE!plco_cd_conta
                                        
                     XLT_SQL = "SELECT LANC_VL_VALOR FROM Lancamentos " & _
                               " WHERE lanc_cd_OrigemFin= " & COD_DUPLICATA & "" & _
                               " AND lanc_tx_Origem = 'C'" & _
                               " AND empr_cd_empresa = " & PCodEmpresa & ""
                               
                    SubQOpenRecordset XLO_TOTALPROMISSORIAS, XLT_SQL, Estatico
                    If Not XLO_TOTALPROMISSORIAS.EOF Then
                        VALORPAGO = XLO_TOTALPROMISSORIAS!lanc_vl_valor
                    End If
                    XLO_TOTALPROMISSORIAS.Close
                                                           
                    Conexao.Execute "DELETE FROM Lancamentos " & _
                                    " WHERE lanc_cd_OrigemFin= " & COD_DUPLICATA & "" & _
                                    " AND lanc_tx_Origem = 'C'" & _
                                    " AND empr_cd_empresa = " & PCodEmpresa & ""
            '                       " AND (plco_cd_debito is null or (not plco_cd_debito is null and not cere_cd_PcrDebitar is null))"
                            
                    'Diminuir do valor total da parte do crédito
                    SubQOpenRecordset XLO_TOTALPROMISSORIAS, "SELECT * FROM Lancamentos " & _
                                " WHERE lanc_dt_Data= " & FunNuloData(dtpDeposito, NomeSgbd) & "" & _
                                    " AND lanc_tx_Origem = 'C'" & _
                                    " AND empr_cd_empresa = " & PCodEmpresa & "" & _
                                    " AND plco_cd_credito is null AND plco_cd_debito = '" & COD_CONTACONTABIL & "'", Dinamico
            
                    If Not XLO_TOTALPROMISSORIAS.EOF Then
                        If XLO_TOTALPROMISSORIAS!lanc_vl_valor = CDbl(VALORPAGO) Then
                            Conexao.Execute "DELETE FROM Lancamentos " & _
                                    " WHERE lanc_dt_Data= " & FunNuloData(dtpDeposito, NomeSgbd) & "" & _
                                    " AND lanc_tx_Origem = 'C'" & _
                                    " AND empr_cd_empresa = " & PCodEmpresa & "" & _
                                    " AND plco_cd_credito is null AND plco_cd_debito = '" & COD_CONTACONTABIL & "'"
                        Else
                            XLO_TOTALPROMISSORIAS!lanc_vl_valor = XLO_TOTALPROMISSORIAS!lanc_vl_valor - CDbl(VALORPAGO)
                            XLO_TOTALPROMISSORIAS.Update
                        End If
                    End If
                End If
                
                subGravaFinanceiro XLO_TITULO, "PAGAR", XLO_TITULO!focl_cd_forncli, XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!NATUREZA
            End If
        Else
            XLO_QUALIFIN.Close
            Set XLO_QUALIFIN = Nothing
            
            If (IsNull(XLO_TITULO!titu_dt_Pagamento)) Then
                subGravaFinanceiro XLO_TITULO, "ALTERAR", XLO_TITULO!focl_cd_forncli, XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!NATUREZA
            Else
                subGravaFinanceiro XLO_TITULO, "ALTERAR", XLO_TITULO!focl_cd_forncli, XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!NATUREZA
                subGravaFinanceiro XLO_TITULO, "PAGAR", XLO_TITULO!focl_cd_forncli, XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!NATUREZA
            End If
            
        End If
        XLO_TITULO.MoveNext
    Wend
    
    XLO_TITULO.Close
    Set XLO_TITULO = Nothing
    
    ProgressBar1.Visible = False
    
    '
    ' Aqui entra a exportação dos Pagamentos a Corretores
    ' ----------------------------------------------------
    
    If ChkExpComissoes.Value <> 1 Or XLT_SQL2 = "" Then
    
        MsgBox "Não foi selecionada a opção de Atualização por Vencimento ou a opção de Exportar Pagamento de Comissões não foi marcada!", vbInformation + vbOKOnly, "ATENÇÃO"
        GoTo Saida
        
    End If
    
    SubQOpenRecordset XLO_COPA, XLT_SQL2, Dinamico
    
    XLB_ERRO = False
    
    ' Reinicia a ProgressBar ou dá mensagem de que não encontrou nada a processar
    If XLO_COPA.RecordCount > 0 Then
        
'        SSFrame1.Caption = "Exportando Pagamento de Comissões:"
        ProgressBar1.Visible = True
        ProgressBar1.Min = 0
        ProgressBar1.Max = XLO_COPA.RecordCount
        ProgressBar1.Value = 0
    
    Else
    
        MsgBox "Não foram encontrados registros de Pagamento de Comissões para processar, para o período informado!", vbInformation + vbOKOnly, "ATENÇÃO"
        GoTo Saida
        
    End If
    
    While Not XLO_COPA.EOF
            
        ProgressBar1.Value = ProgressBar1.Value + 1
    
        GravaComissaoNoQualiFin XLO_COPA        ' Executa a subrotina com passagem do objeto XLO_COPA como parâmetro
        
        If Not XLB_ERRO Then
            XLO_COPA.Update
        Else
            MsgBox "Erro na atualização do registro de pagamentos (" & XLO_COPA!copa_tx_codigo & ")!", vbCritical, "ATENÇÃO"
            Conexao.RollbackTrans
            Exit Sub
        End If
        
        XLO_COPA.MoveNext
        
    Wend
    
    XLO_COPA.Close
    Set XLO_COPA = Nothing
        
    ProgressBar1.Visible = False
        
Saida:

    Conexao.CommitTrans

    MsgBox "Atualização realizada com sucesso!", vbInformation + vbOKOnly, "ATENÇÃO"
    Unload Me
    Exit Sub
    
RotuloErro:

    XLB_ERRO = True
    funTrataErros (ComMensagem)
    Conexao.RollbackTrans
    
End Sub

'
' Subrotina para Gerar NFs, Duplicatas e rateio/distribuição por centros de custo
' de pagamentos de comissão a corretores (Contas a Pagar)
'
' Parâmetro(s):
' XLO_COPA      - Recordset com os registros da tabela CorretoresPagamento
' -----------------------------------------------------------------------------------
Sub GravaComissaoNoQualiFin(XLO_COPA As Object)

    '
    ' Dimensionamento de variáveis e demais objetos
    ' -------------------------------------------------------------------------------
    Dim XLT_SQL       As String
    Dim XLO_NOFI      As ADODB.Recordset    ' Recordset com a tabela NotasFiscais
    Dim XLO_DUPL      As ADODB.Recordset    ' Recordset com a tabela Duplicatas
    Dim XLL_FATURA    As Long               ' Número da fatura
    Dim XLL_NFISCAL   As Long               ' Código (ID) da Nota Fiscal
    Dim XLL_DUPLICATA As Long               ' Código (ID) da Duplicata
    
    On Error GoTo RotuloErro

    '
    ' Abre o recordset para a tabela NotasFiscais e adiciona novo registro
    ' -------------------------------------------------------------------------------
    SubQOpenRecordset XLO_NOFI, "SELECT * FROM NotasFiscais WHERE nofi_cd_notafiscal = -1", Dinamico
    XLO_NOFI.AddNew
    
    '
    ' Guarda o número da Fatura
    ' -------------------------------------------------------------------------------
    XLL_FATURA = FunPegaGuardaUltimo("F")
    
    '
    ' Preenche o novo registro com os dados de XLO_COPA e atualiza-o
    ' (campos em ordem alfabética para facilitar a visualização e pesquisa)
    ' -------------------------------------------------------------------------------
    XLO_NOFI!cofa_cd_Codigo = Null                              ' ???
    XLO_NOFI!empr_cd_empresa = XLO_COPA!empr_cd_empresa         ' Código da Empresa
    XLO_NOFI!focl_cd_forncli = XLO_COPA!focl_cd_forncli         ' Código do Corretor/Fornecedor
    XLO_NOFI!nofi_bl_Sigiloso = 0                               ' Documento não sigiloso
    XLO_NOFI!nofi_cd_notaGlosaDev = Null
    
    If XLO_COPA!copa_tx_statusDocumento = "P" Then              ' copa_dt_emissao, não usada com status Previsto
        XLO_NOFI!nofi_dt_EmissaoFatura = Null
        XLO_NOFI!nofi_dt_emissao = Null
    Else
        XLO_NOFI!nofi_dt_EmissaoFatura = XLO_COPA!copa_dt_emissao
        XLO_NOFI!nofi_dt_emissao = XLO_COPA!copa_dt_emissao
    End If
    
    XLO_NOFI!nofi_dt_entrada = XLO_COPA!copa_dt_entrada
    XLO_NOFI!nofi_nr_DocSubstituido = Null
    
    If XLO_COPA!copa_tx_statusDocumento = "R" Then              ' Código do registro de pagamento (somente quando realizado)
        XLO_NOFI!nofi_nr_TituloCapi = XLO_COPA!copa_nr_titulocapi
    Else
        XLO_NOFI!nofi_nr_TituloCapi = Null
    End If
    
    XLO_NOFI!nofi_nr_documento = XLO_COPA!copa_nr_documento     ' Número documento informado
    XLO_NOFI!nofi_nr_fatura = XLL_FATURA                        ' Número da Fatura gerada
    XLO_NOFI!nofi_nr_provisao = Null
    XLO_NOFI!nofi_tx_NumCartaTransferencia = Null
    XLO_NOFI!nofi_tx_OrigemQualiFat = Null
    XLO_NOFI!nofi_tx_ReferenciaTransferencia = Null
    XLO_NOFI!nofi_tx_contabil = Null
    XLO_NOFI!nofi_tx_controle = Null
    XLO_NOFI!nofi_tx_distribuir = "N"                           ' Distribuição/rateio por C/C
    XLO_NOFI!nofi_tx_estorno = Null
    XLO_NOFI!nofi_tx_historico = XLO_COPA!copa_tx_historico
    XLO_NOFI!nofi_tx_pendestoque = Null
    XLO_NOFI!nofi_tx_serie = XLO_COPA!copa_tx_serie             ' Série do documento informado
    XLO_NOFI!nofi_tx_status = XLO_COPA!copa_tx_statusDocumento  ' (P)revisto, (C)ompromissado, (R)ealizado
    XLO_NOFI!nofi_tx_tipofatura = "1"                           ' Uma fatura para uma duplicata (automática)
    XLO_NOFI!nofi_tx_tipomov = "S"                              ' Saida
    XLO_NOFI!nofi_vl_valor = XLO_COPA!copa_vl_valor             ' Valor do documento
    XLO_NOFI!tido_cd_TipoDoc = XLO_COPA!tido_cd_TipoDoc         ' Tipo documento NF, NFs, NFF, ...
    
    XLO_NOFI.Update

    '
    ' Guarda o número da nota fiscal
    ' -------------------------------------------------------------------------------
    XLL_NFISCAL = XLO_NOFI!nofi_cd_notafiscal

    '
    ' Abre o recordset para a tabela Duplicatas e adiciona novo registro
    ' -------------------------------------------------------------------------------
    SubQOpenRecordset XLO_DUPL, "SELECT * FROM Duplicatas WHERE dupl_cd_duplicata = -1", Dinamico
    XLO_DUPL.AddNew

    '
    ' Preenche o novo registro com os dados de XLO_COPA + XLO_NOFI e atualiza-o
    ' (campos em ordem alfabética para facilitar a visualização e pesquisa)
    ' -------------------------------------------------------------------------------
    XLO_DUPL!bafo_cd_codigo = 0                                 ' Null ????
    XLO_DUPL!coco_cd_codigo = XLO_COPA!coco_cd_codigo           ' Código (ID) de conta corrente
    XLO_DUPL!dupl_dt_ProcPagamento = Null
    XLO_DUPL!dupl_dt_concilia = Null
    XLO_DUPL!dupl_dt_pagrec = Null
    XLO_DUPL!dupl_dt_vencimento = XLO_COPA!copa_dt_vencimento   ' Ddata de vencimento da duplicata
    XLO_DUPL!dupl_nr_duplicata = "U"                            ' Parcela única
    XLO_DUPL!dupl_nr_fatura = XLL_FATURA                        ' Número gerado para a fatura
    XLO_DUPL!dupl_tx_Estorno = Null
    XLO_DUPL!dupl_tx_ExportaContabilidade = Null
    XLO_DUPL!dupl_tx_NaoRelFinanceiro = Null
    XLO_DUPL!dupl_tx_consolidado = Null
    XLO_DUPL!dupl_tx_controle = Null
    XLO_DUPL!dupl_tx_debcred = "D"                              ' D - Débito
    XLO_DUPL!dupl_tx_docpagrec = Null
    XLO_DUPL!dupl_tx_expbanco = "N"
    XLO_DUPL!dupl_tx_status = XLO_COPA!copa_tx_statusDocumento  ' (P)revisto, (C)ompromissado, (R)ealizado
    XLO_DUPL!dupl_tx_tipopagrec = "1"                           ' ?????
    XLO_DUPL!dupl_vl_GlosaDevolucao = 0
    XLO_DUPL!dupl_vl_valor = XLO_COPA!copa_vl_valor
    XLO_DUPL!fopa_cd_codigo = 0                                 ' Null ????
    
    XLO_DUPL.Update

    '
    ' Guarda o número da Duplicata
    ' -------------------------------------------------------------------------------
    XLL_DUPLICATA = XLO_DUPL!dupl_cd_duplicata
    
    '
    ' Grava registro de distribuição (Apropriação)
    ' -------------------------------------------------------------------------------
    XLT_SQL = "INSERT INTO Distribuicao(plco_cd_Conta, cere_cd_Pcr, nofi_cd_notafiscal, dist_vl_valor, dist_vl_GlosaDevolucao, iden_cd_Projeto)" & _
              " VALUES ('" & XLO_COPA!plco_cd_conta & "', '" & XLO_COPA!cere_cd_Pcr & "', " & FunNuloBanco(XLL_NFISCAL) & ", " & FunTrataFloat(XLO_COPA!copa_vl_valor) & ", 0, " & FunNuloBanco(XLO_COPA!iden_cd_Projeto) & ")"

    Conexao.Execute XLT_SQL
    
    XLO_NOFI.Close
    Set XLO_NOFI = Nothing

    XLO_DUPL.Close
    Set XLO_DUPL = Nothing
    
    '
    ' Atualiza o registro de CorretoresPagamento com o Id da NF e Status
    ' Obs.: Dados da duplicata, deverão ser obtidos por views que relacionem a Nota
    '       Fiscal com a Duplicata pelo número da fatura
    ' -------------------------------------------------------------------------------
    XLO_COPA!nofi_cd_notafiscal = FunNuloBanco(XLL_NFISCAL)
'    XLO_COPA!dupl_cd_duplicata = FunNuloBanco(XLL_DUPLICATA)
    XLO_COPA!copa_tx_statusPagamento = "T"

    Exit Sub
    
RotuloErro:
    
    XLB_ERRO = True
    funTrataErros (ComMensagem)
    
End Sub

Private Sub Form_Activate()
    Set Formulario = FrmFinAlteraQualifin
End Sub

Private Sub Form_Load()
    Dim XLT_SQL  As String
      
    Set Formulario = FrmFinAlteraQualifin
    Formulario.Left = (ResolucaoX - Formulario.Width) / 2
    Formulario.Top = (Resolucaoy - Formulario.Height) / 2

    '
    ' Inicializa os campos de data com as datas domês atual
    DtpInicio.Value = DateSerial(Year(Now), Month(Now), 1)
    ' dtpFim.Value = Now
    dtpDeposito.Value = Now
    
   ' CliqueBarraFerramentas = False
    XLT_SQL = "SELECT dbo.Titulos.empd_cd_Empreendimento + '.' + dbo.Titulos.imov_cd_Imovel + '.' + dbo.Titulos.cont_cd_Contrato + '.' + dbo.Titulos.titu_cd_Plano + '.' + dbo.Titulos.titu_cd_Parcela + '.' + dbo.Titulos.titu_cd_Residuo AS Titulo FROM dbo.Titulos WHERE dbo.Titulos.empr_cd_Empresa =" & PCodEmpresa & " ORDER BY titulo"
    subConectarControleDados DatTitulosInicio, XLT_SQL, Estatico
    subConectarControleDados DatTitulosFim, XLT_SQL, Estatico
End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "FrmFinAlteraQualifin"
End Sub

Private Sub optAtualizarDataDeposito_Click()
    DtpInicio.Enabled = False
    DtpInicio.Value = Null
    DtpFim.Enabled = False
    DtpFim.Value = Null
    
    dtpDeposito.Enabled = True
    
    ChkExportar.Enabled = False
    ChkExportar.Value = 0
End Sub

Private Sub optAtualizarPorVencimento_Click()
    DtpInicio.Enabled = True
    DtpInicio.Value = Null
    DtpFim.Enabled = True
    DtpFim.Value = Null
    
    dtpDeposito.Enabled = False
    
    ChkExportar.Enabled = True
    ChkExportar.Value = 1
    
End Sub

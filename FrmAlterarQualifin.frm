VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFinAlteraQualifin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualização de Títulos no Qualifin"
   ClientHeight    =   5745
   ClientLeft      =   3240
   ClientTop       =   2805
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   4785
   Begin VB.OptionButton optAtualizarDataDeposito 
      Caption         =   "Atualizar por Data de Depósito"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   1680
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
      Top             =   5325
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
      Top             =   5325
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
      Height          =   900
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4500
      _Version        =   65536
      _ExtentX        =   7937
      _ExtentY        =   1587
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
      Begin MSComCtl2.DTPicker DtpInicio 
         Height          =   315
         Left            =   795
         TabIndex        =   3
         Top             =   330
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   122224641
         CurrentDate     =   37180
      End
      Begin MSComCtl2.DTPicker DtpFim 
         Height          =   315
         Left            =   2835
         TabIndex        =   4
         Top             =   330
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   122224641
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
      Top             =   3240
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
      Left            =   600
      Top             =   5880
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
      Left            =   600
      Top             =   6240
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
      Top             =   2160
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
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   330
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         DateIsNull      =   -1  'True
         Format          =   122224641
         CurrentDate     =   37180
      End
   End
End
Attribute VB_Name = "FrmFinAlteraQualifin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    
    XLI_RESPOSTA = 0
    XGT_SELECAO = ""
    XGT_CONJUNCAO = ""
        
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
                "(titu_dt_Prorrogacao <= '" & Format(DtpFim.Value, "mm/dd/yyyy") & "')"
            XGT_CONJUNCAO = " AND "
        End If
    Else
        XGT_SELECAO = XGT_SELECAO & XGT_CONJUNCAO & "titu_dt_Deposito = " & FunNuloData(DtpDeposito.Value, NomeSgbd)
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
        SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCapAtualizaQualifn WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & XGT_SELECAO & "AND titu_dt_pagamento is null", Estatico
     ElseIf optAtualizarPorVencimento.Value Then
        If MsgBox("Os títulos realizados (pagos) passarão a ser compromissados.Deseja realmente executar a atualização?", vbYesNo, "Atualizar QualiFin") = vbYes Then
            SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCapAtualizaQualifn WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & XGT_SELECAO & " ORDER BY coco_cd_codigo", Estatico
        Else
            Exit Sub
        End If
     Else
        SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCapAtualizaQualifn WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & XGT_SELECAO & " ORDER BY coco_cd_codigo", Estatico
     End If
     
    While Not XLO_TITULO.EOF
        SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM NotasFiscais WHERE empr_cd_empresa = " & PCodEmpresa & " AND nofi_nr_TituloCapi = '" & XLO_TITULO!Titulo & "'", Estatico
                
        If Not XLO_QUALIFIN.EOF Then
            XLB_COMPROMISSADO = (XLO_QUALIFIN!nofi_tx_status = "C")
            XLO_QUALIFIN.Close
            Set XLO_QUALIFIN = Nothing
            
            If (IsNull(XLO_TITULO!titu_dt_Pagamento)) Then
                subGravaFinanceiro XLO_TITULO, "COMPROMISSAR", XLO_TITULO!focl_cd_FornCli, XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!NATUREZA
            ElseIf (optAtualizarPorVencimento.Value) Then
                If (XLB_COMPROMISSADO) Then
                    subGravaFinanceiro XLO_TITULO, "PAGAR", XLO_TITULO!focl_cd_FornCli, XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!NATUREZA
                End If
            ElseIf (optAtualizarDataDeposito.Value) Then
                
                If XGB_PQUALICONT Then
                   If funMesFechado(DtpDeposito) <> "A" Then
                       MsgBox "Não é possível executar a operação, pois o mês contábil já foi fechado. " _
                               + Chr(13) + "Entre em contato com a Contabilidade.", vbCritical, "ATENÇÃO"
                        XLO_QUALIFIN.Close
                       Exit Sub
                   End If
                End If
                
                
                Dim COD_DUPLICATA As Long
                Dim COD_CONTACONTABIL As Long
                Dim VALORPAGO As Long
                
                XLT_SQL = _
                    "SELECT duplicatas.DUPL_CD_DUPLICATA, ContasCorrente.plco_cd_conta " & _
                    "FROM  " & _
                    "NotasFiscais " & _
                    "INNER JOIN duplicatas on duplicatas.dupl_nr_fatura = notasfiscais.nofi_nr_fatura " & _
                    "INNER JOIN ContasCorrente ON ContasCorrente.coco_cd_codigo = Duplicatas.coco_cd_codigo " & _
                    "WHERE NotasFiscais.empr_cd_empresa = " & CStr(PCodEmpresa) & _
                    " AND NotasFiscais.nofi_nr_titulocapi = '" & XLO_TITULO!Titulo & "'"

                SubQOpenRecordset XLO_CONTABILIDADE, XLT_SQL, Estatico
                If Not XLO_CONTABILIDADE.EOF Then
                    
                    COD_DUPLICATA = XLO_CONTABILIDADE!DUPL_CD_DUPLICATA
                    COD_CONTACONTABIL = XLO_CONTABILIDADE!PLCO_CD_CONTA
                    
                    
                     XLT_SQL = _
                                    "SELECT LANC_VL_VALOR FROM Lancamentos " & _
                                    " WHERE lanc_cd_OrigemFin= " & COD_DUPLICATA & "" & _
                                    " AND lanc_tx_Origem = 'C'" & _
                                    " AND empr_cd_empresa = " & PCodEmpresa & ""
                    SubQOpenRecordset XLO_TOTALPROMISSORIAS, XLT_SQL, Estatico
                    If Not XLO_TOTALPROMISSORIAS.EOF Then
                        VALORPAGO = XLO_TOTALPROMISSORIAS!LANC_VL_VALOR
                    End If
                    XLO_TOTALPROMISSORIAS.Close
                       
                    
                    
                    Conexao.Execute "DELETE FROM Lancamentos " & _
                                    " WHERE lanc_cd_OrigemFin= " & COD_DUPLICATA & "" & _
                                    " AND lanc_tx_Origem = 'C'" & _
                                    " AND empr_cd_empresa = " & PCodEmpresa & ""
            '                       " AND (plco_cd_debito is null or (not plco_cd_debito is null and not cere_cd_PcrDebitar is null))"
                            
                    'Diminuir do valor total da parte do crédito
                    SubQOpenRecordset XLO_TOTALPROMISSORIAS, "SELECT * FROM Lancamentos " & _
                                " WHERE lanc_dt_Data= " & FunNuloData(DtpDeposito, NomeSgbd) & "" & _
                                    " AND lanc_tx_Origem = 'C'" & _
                                    " AND empr_cd_empresa = " & PCodEmpresa & "" & _
                                    " AND plco_cd_credito is null AND plco_cd_debito = '" & COD_CONTACONTABIL & "'", Dinamico
            
                    If Not XLO_TOTALPROMISSORIAS.EOF Then
                        If XLO_TOTALPROMISSORIAS!LANC_VL_VALOR = CDbl(VALORPAGO) Then
                            Conexao.Execute "DELETE FROM Lancamentos " & _
                                    " WHERE lanc_dt_Data= " & FunNuloData(DtpDeposito, NomeSgbd) & "" & _
                                    " AND lanc_tx_Origem = 'C'" & _
                                    " AND empr_cd_empresa = " & PCodEmpresa & "" & _
                                    " AND plco_cd_credito is null AND plco_cd_debito = '" & COD_CONTACONTABIL & "'"
                        Else
                            XLO_TOTALPROMISSORIAS!LANC_VL_VALOR = XLO_TOTALPROMISSORIAS!LANC_VL_VALOR - CDbl(VALORPAGO)
                            XLO_TOTALPROMISSORIAS.Update
                        End If
                    End If
                End If
                
                subGravaFinanceiro XLO_TITULO, "PAGAR", XLO_TITULO!focl_cd_FornCli, XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!NATUREZA
            End If
        Else
            XLO_QUALIFIN.Close
            Set XLO_QUALIFIN = Nothing
            
            If (IsNull(XLO_TITULO!titu_dt_Pagamento)) Then
                subGravaFinanceiro XLO_TITULO, "ALTERAR", XLO_TITULO!focl_cd_FornCli, XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!NATUREZA
            Else
                
                subGravaFinanceiro XLO_TITULO, "ALTERAR", XLO_TITULO!focl_cd_FornCli, XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!NATUREZA
                subGravaFinanceiro XLO_TITULO, "PAGAR", XLO_TITULO!focl_cd_FornCli, XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!NATUREZA
            End If
            
        End If
        XLO_TITULO.MoveNext
    Wend
    XLO_TITULO.Close
    Set XLO_TITULO = Nothing
    
    MsgBox "Atualização realizada com sucesso!", vbInformation + vbOKOnly, "ATENÇÃO"
    'Unload Me
    
    Exit Sub
    
RotuloErro:
    funTrataErros (ComMensagem)
    Resume Next
End Sub



Private Sub Form_Activate()
    Set Formulario = FrmFinAlteraQualifin
End Sub

Private Sub Form_Load()
    Dim XLT_SQL  As String
      
    Set Formulario = FrmFinAlteraQualifin
    DtpDeposito.Value = Now
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
    
    DtpDeposito.Enabled = True
    
    ChkExportar.Enabled = False
    ChkExportar.Value = 0
End Sub

Private Sub optAtualizarPorVencimento_Click()
    DtpInicio.Enabled = True
    DtpInicio.Value = Null
    DtpFim.Enabled = True
    DtpFim.Value = Null
    
    DtpDeposito.Enabled = False
    
    ChkExportar.Enabled = True
    ChkExportar.Value = 1
    
End Sub

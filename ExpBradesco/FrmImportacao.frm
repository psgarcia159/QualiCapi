VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmImportacao 
   Caption         =   "Retorno de Pagamento Eletrônico - BRADESCO"
   ClientHeight    =   1785
   ClientLeft      =   2085
   ClientTop       =   1245
   ClientWidth     =   9120
   Icon            =   "FrmImportacao.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel PanConta 
      Height          =   1395
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   8715
      _Version        =   65536
      _ExtentX        =   15372
      _ExtentY        =   2461
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.CommandButton CmdConfirmar 
         Caption         =   "&Confirmar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   4
         Top             =   840
         Width           =   1065
      End
      Begin VB.CommandButton CmdRetornar 
         Caption         =   "&Sair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7260
         TabIndex        =   3
         Top             =   840
         Width           =   1065
      End
      Begin MSDataListLib.DataCombo CboCCorrente 
         Bindings        =   "FrmImportacao.frx":08CA
         Height          =   330
         Left            =   1545
         TabIndex        =   1
         Top             =   300
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "desccombo"
         BoundColumn     =   "coco_cd_codigo"
         Text            =   ""
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
      Begin MSAdodcLib.Adodc DatContaCorrente 
         Height          =   330
         Left            =   2820
         Top             =   840
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
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
         Caption         =   "DatContaCorrente"
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
      Begin VB.Label LblCCorrente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Conta Corrente:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1365
      End
   End
   Begin MSComDlg.CommonDialog DlgAbrirFigura 
      Left            =   780
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmImportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Caminho As String

Private Sub CboCCorrente_Change()

  If CboCCorrente.Text <> "" Then
  
    DatContaCorrente.Recordset.Bookmark = CboCCorrente.SelectedItem
  
  End If

End Sub

Private Sub CmdConfirmar_Click()
  
    Dim XLT_TEXTO           As String       'Armazena o arquivo texto
    Dim XLT_TITULO          As String       'Armazena o nº do título (nº da fatura)
    Dim XLT_PARCELA         As String       'Armazena o cód. da parcela
    Dim XLD_DATAPAG         As Date         'Data de Pagamento
    Dim XLD_DATADEPOSITO    As Date         'Data de Deposito
    Dim XLF_VALORPAGO       As Double       'Valor pago
    Dim XLF_VALORTITULO     As Double       'Valor do título a ser pago - Para o arquivo de inconsistências.
    Dim XLD_DATAVENCIMENTO  As String       'Data de Vencimento do Título - Para o arquivo de inconsistências
    Dim XLI_CONTADOR        As Integer      'Contador de Motivos das Ocorrências
    Dim XLI_POSICAO         As Integer      'Posição do Caracter na linha selecionada
    Dim XLT_NUMARQUIVO      As String       'Identificador do Arquivo
    
    
    
    
    
    Dim XLT_SQL As String       'Prepara Select
    
    Dim XLT_CODEMPREENDIMENTO As String 'Armazena código do empreendimento
    Dim XLT_CODIMOVEL As String         'Armazena código do Imóvel
    Dim XLT_CODCONTRATO As String       'Armazena código do contrato
    Dim XLT_CODPLANO As String          'Armazena código do plano
    Dim XLT_CODPARCELA As String        'Armazena código da Parcela
    Dim XLT_CODRESIDUO As String        'Armazena código do Residuo
    Dim XLT_CODEMPRESA As String
    
    Dim XLT_CARTEIRA As String
    Dim XLF_JUROS As Double
    Dim XLF_DESCONTO As Double
    Dim XLF_PREVISTO As Double
    Dim XLF_SEGURO As Double
    Dim XLF_CORRECAO As Double
    Dim XLO_IMPORTACAO As New ADODB.Recordset
    Dim XLO_TITULO As New ADODB.Recordset
    Dim XLO_CODTIPOPLANO As New ADODB.Recordset
    Dim XLO_FINANCEIRO As New ADODB.Recordset
    Dim XLO_OCORRENCIAS As New ADODB.Recordset
  
    Dim XLF_RESIDUO As Double
    Dim XLF_VALORINDEXADO1 As Double
    Dim XLF_VALORINDEXADO2 As Double
    Dim XLT_TIPOPAG As String
    Dim XLI_CODTIPOPLANO As Integer
    Dim XLT_NATUREZAPLANO As String       'Utilizado  para informar a natureza a rotina de gravação no financeiro
    
    Dim XLB_ERRO As Boolean
    Dim XLT_SQLRECIBO As String
    Dim XLT_DOCPGTO As String
    Dim XLT_FORMAPGTO As String
    Dim XLI_CONT As Integer
    
    XLB_ERRO = False
    
    If FunObrigatorioCBO(CboCCorrente, "Selecione uma Conta Corrente.") Then Exit Sub
    
    DlgAbrirFigura.CancelError = False
    DlgAbrirFigura.InitDir = XGT_LOCALARQ
    DlgAbrirFigura.ShowOpen
    
    If DlgAbrirFigura.FileName = "" Then
        Unload Me
        Exit Sub
    End If
    
    If DlgAbrirFigura.FileName <> "" Then
        XGT_LOCALARQ = DlgAbrirFigura.FileName
        XGT_ARQUIVO = DlgAbrirFigura.FileTitle
    End If
  
    Me.MousePointer = vbHourglass
    
    XLT_CODEMPRESA = Format(Left(DatContaCorrente.Recordset.Fields!coco_cd_Agencia, 4), "0000") & _
                     "0" & _
                     Format(Left(DatContaCorrente.Recordset.Fields!coco_tx_Conta, 9), "000000000") & _
                     Format(Right(FunNuloVal(DatContaCorrente.Recordset.Fields!coco_nr_Dac), 1), "0")

    'Consulta para criação do recibo
    If NomeSgbd = "Access" Then
        XLT_SQLRECIBO = " SELECT consCAPRecibo.*," & _
                        "(SELECT COUNT(*) FROM Titulos AS tit WHERE  tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento " & _
                        " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND  tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato " & _
                        " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_residuo = '00') AS TotalParcelas, " & _
                        " IIF((SELECT titu_vl_Parcela FROM ConscapTitulos as tit WHERE   tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento" & _
                        " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato" & _
                        " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_Parcela = consCAPRecibo.titu_cd_Parcela" & _
                        " AND (tit.titu_cd_Residuo = format(int(consCAPRecibo.titu_cd_residuo) + 1,'00'))) is null, 0 , (SELECT titu_vl_Parcela FROM ConscapTitulos as tit WHERE   tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento" & _
                        " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato" & _
                        " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_Parcela = consCAPRecibo.titu_cd_Parcela" & _
                        " AND (tit.titu_cd_Residuo = format(int(consCAPRecibo.titu_cd_residuo) + 1,'00')))) AS ValorResiduo" & _
                        " FROM consCAPRecibo WHERE " & _
                        " empr_cd_empresa = " & PCodEmpresa & " AND ("
    Else
        XLT_SQLRECIBO = " SELECT consCAPRecibo.*," & _
                        "(SELECT COUNT(*) FROM Titulos AS tit WHERE  tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento " & _
                        " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND  tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato " & _
                        " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_residuo = '00') AS TotalParcelas, " & _
                        " IsNull((SELECT titu_vl_Parcela FROM ConscapTitulos as tit WHERE   tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento" & _
                        " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato" & _
                        " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_Parcela = consCAPRecibo.titu_cd_Parcela" & _
                        " AND (tit.titu_cd_Residuo = (consCAPRecibo.titu_cd_residuo + 1))), 0) AS ValorResiduo" & _
                        " FROM consCAPRecibo WHERE " & _
                        " empr_cd_empresa = " & PCodEmpresa & " AND ("
    End If
    
    XGT_CONJUNCAO = ""
  
    Open XGT_LOCALARQ For Input Access Read As #1
    
    'Documento de pagamento
    XLT_DOCPGTO = XGT_ARQUIVO
    XLT_FORMAPGTO = "Bloqueto"
    
    'leitura do header
    Input #1, XLT_TEXTO
    
    '109 a 113 - N. Aviso Bancário (identifica o arquivo para não permitir a importação mais de uma vez.
     XLT_NUMARQUIVO = Mid(XLT_TEXTO, 109, 5)
     XLT_SQL = "SELECT NotasFiscais.empr_cd_empresa FROM " & _
                "NotasFiscais INNER JOIN Duplicatas ON (NotasFiscais.Nofi_nr_fatura = Duplicatas.Dupl_nr_fatura) " & _
                "INNER JOIN RelacaoOcorrencias ON (Duplicatas.Dupl_nr_fatura = RelacaoOcorrencias.Dupl_nr_fatura) " & _
                "where NotasFiscais.Empr_cd_empresa = " & PCodEmpresa & _
                " and RelacaoOcorrencias.reoc_nr_Lote = " & XLT_NUMARQUIVO
    SubQOpenRecordset XLO_OCORRENCIAS, XLT_SQL, Estatico
    If Not (XLO_OCORRENCIAS.EOF) Then
        Close #1
        MsgBox "Esse arquivo já foi importado!"
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
        If Mid(XLT_TEXTO, 2, 1) = "2" Then 'Verifica se é o header de retorno
            If UCase(Trim(Mid(XLT_TEXTO, 80, 15))) = UCase("Bradesco") Then 'Verifica se é o do Banco Bradesco
          
                Conexao.BeginTrans
                
                While Not EOF(1)
                    
                    Input #1, XLT_TEXTO
                    'Verifica se é registro de detalhe
                    If Mid(XLT_TEXTO, 1, 1) = "1" Then
                    ' Verifica se é registro de retorno de inconsistência ou se é confirmação de pagamento.
                    'Se não tem data de pagamento, então é registro de inconsistência.
                    'O registro de inconsistência poderá ser uma confirmação de aceitação do título exportado ou a informação de uma
                    'inconsistência que ocorreu em relação ao título no último arquivo de remessa gerado.
                        If Not (IsDate(Mid(XLT_TEXTO, 296, 2) & "/" & Mid(XLT_TEXTO, 298, 2) & "/" & Mid(XLT_TEXTO, 300, 2))) Then
                                '038 a 062 - N. de Controle do Participante
                                XLT_TITULO = Trim(Mid(XLT_TEXTO, 38, 23))
                                XLT_CODEMPREENDIMENTO = Mid(XLT_TEXTO, 38, 4)
                                XLT_CODIMOVEL = Mid(XLT_TEXTO, 43, 4)
                                XLT_CODCONTRATO = Mid(XLT_TEXTO, 48, 2)
                                XLT_CODPLANO = Mid(XLT_TEXTO, 51, 2)
                                XLT_CODPARCELA = Mid(XLT_TEXTO, 54, 3)
                                XLT_CODRESIDUO = Mid(XLT_TEXTO, 58, 2)
                                
                                '147 a 152 - Data de Vencimento do Título
                                XLD_DATAVENCIMENTO = (Mid(XLT_TEXTO, 149, 2) & "/" & Mid(XLT_TEXTO, 147, 2) & "/" & _
                                                      Mid(XLT_TEXTO, 151, 2))
                                '153 a 165 - Valor do Título
                                XLF_VALORTITULO = CDbl(Mid(XLT_TEXTO, 153, 13) / 100)
                                '109 a 110 - Identificação de Ocorrência e 319 a 328 Motivos das Rejeições para os Códigos de Ocorrência da Posição 109 a 110
                                XLT_SQL = "select Duplicatas.dupl_nr_Fatura, Duplicatas.dupl_nr_Duplicata from NotasFiscais " & _
                                            "INNER JOIN Duplicatas ON (NotasFiscais.Nofi_nr_fatura = Duplicatas.dupl_nr_fatura)" & _
                                            "WHERE NotasFiscais.nofi_nr_titulocapi = '" & XLT_TITULO & "'"
                                SubQOpenRecordset XLO_IMPORTACAO, XLT_SQL, Estatico
                                If Not (XLO_IMPORTACAO.EOF) Then
                                        XLI_CONTADOR = 1
                                        XLI_POSICAO = 319
            
                                        While XLI_CONTADOR < 6
                                            If XLI_CONTADOR <> 1 Then
                                                If (Mid(XLT_TEXTO, XLI_POSICAO, 1) <> 0) Or (Mid(XLT_TEXTO, XLI_POSICAO, 2) <> 0) Then
                                                    XLT_SQL = "INSERT INTO RelacaoOcorrencias (ocor_cd_Codigo, banc_cd_Codigo, dupl_nr_Fatura, dupl_nr_Duplicata, reoc_vl_ValorPago, reoc_dt_Pagamento, reoc_nr_Lote) VALUES ("
                                                    XLT_SQL = XLT_SQL & "'" & (Mid(XLT_TEXTO, 109, 2) & Mid(XLT_TEXTO, XLI_POSICAO, 2)) & "', " & "237, " & XLO_IMPORTACAO!dupl_nr_fatura & ", '" & XLO_IMPORTACAO!dupl_nr_Duplicata & "', " & FunTrataFloat(XLF_VALORTITULO) & ", '" & XLD_DATAVENCIMENTO & "', '" & XLT_NUMARQUIVO & "')"
                                                    Conexao.Execute (XLT_SQL)
                                                    
                                                    XLI_POSICAO = XLI_POSICAO + 2
                                                End If
                                            Else
                                            'Se existe a Data Ocorrência no Banco sem a Data do Crédito, então o pagamento foi feito com Cheque.
                                                If (Mid(XLT_TEXTO, XLI_POSICAO, 1) = 0) And (Mid(XLT_TEXTO, XLI_POSICAO, 2) = 0) Then
                                                    XLT_SQL = "INSERT INTO RelacaoOcorrencias (ocor_cd_Codigo, banc_cd_Codigo, dupl_nr_Fatura, dupl_nr_Duplicata, reoc_vl_ValorPago, reoc_dt_Pagamento, reoc_nr_Lote) VALUES ("
                                                    XLT_SQL = XLT_SQL & "'00CH'" & ", " & "237, " & XLO_IMPORTACAO!dupl_nr_fatura & ", '" & XLO_IMPORTACAO!dupl_nr_Duplicata & "', " & FunTrataFloat(XLF_VALORTITULO) & ", '" & XLD_DATAVENCIMENTO & "', '" & XLT_NUMARQUIVO & "')"
                                                    Conexao.Execute (XLT_SQL)
                                                Else
                                                    XLT_SQL = "INSERT INTO RelacaoOcorrencias (ocor_cd_Codigo, banc_cd_Codigo, dupl_nr_Fatura, dupl_nr_Duplicata, reoc_vl_ValorPago, reoc_dt_Pagamento, reoc_nr_Lote) VALUES ("
                                                    XLT_SQL = XLT_SQL & "'" & (Mid(XLT_TEXTO, 109, 2) & Mid(XLT_TEXTO, XLI_POSICAO, 2)) & "', " & "237, " & XLO_IMPORTACAO!dupl_nr_fatura & ", '" & XLO_IMPORTACAO!dupl_nr_Duplicata & "', " & FunTrataFloat(XLF_VALORTITULO) & ", '" & XLD_DATAVENCIMENTO & "', '" & XLT_NUMARQUIVO & "')"
                                                    Conexao.Execute (XLT_SQL)

                                                End If
                                                
                                                
                                                XLI_POSICAO = XLI_POSICAO + 2
    
                                            End If
                                            XLI_CONTADOR = XLI_CONTADOR + 1
                                        Wend
                                        XLO_IMPORTACAO.Close
                                        Set XLO_IMPORTACAO = Nothing
                            End If

                        Else 'Confirmação de Pagamento
                            'Registro de detalhe tipo U
                            XLT_TITULO = Trim(Mid(XLT_TEXTO, 38, 23))
                            XLT_CODEMPREENDIMENTO = Mid(XLT_TEXTO, 38, 4)
                            XLT_CODIMOVEL = Mid(XLT_TEXTO, 43, 4)
                            XLT_CODCONTRATO = Mid(XLT_TEXTO, 48, 2)
                            XLT_CODPLANO = Mid(XLT_TEXTO, 51, 2)
                            XLT_CODPARCELA = Mid(XLT_TEXTO, 54, 3)
                            XLT_CODRESIDUO = Mid(XLT_TEXTO, 58, 2)
    
                            XLD_DATAPAG = IIf(IsDate(Mid(XLT_TEXTO, 111, 2) & "/" & Mid(XLT_TEXTO, 113, 2) & _
                                            "/" & Mid(XLT_TEXTO, 115, 2)), Mid(XLT_TEXTO, 111, 2) & "/" & Mid(XLT_TEXTO, 113, 2) & _
                                            "/" & Mid(XLT_TEXTO, 115, 2), "000000")
                            

                            XLD_DATADEPOSITO = IIf(IsDate(Mid(XLT_TEXTO, 296, 2) & "/" & Mid(XLT_TEXTO, 298, 2) & _
                                                "/" & Mid(XLT_TEXTO, 300, 2)), Mid(XLT_TEXTO, 296, 2) & "/" & Mid(XLT_TEXTO, 298, 2) & _
                                                "/" & Mid(XLT_TEXTO, 300, 2), "000000")
                            
                            XLF_JUROS = CDbl(Mid(XLT_TEXTO, 267, 13) / 100)
                            XLF_DESCONTO = CDbl(Mid(XLT_TEXTO, 241, 13) / 100)
                            XLF_VALORPAGO = CDbl(Mid(XLT_TEXTO, 254, 13) / 100)
                          
                            XLT_SQL = FunCriaConsultaBase("01/01/01", NomeSgbd, "ConsCAPImpBanco", 3)
                            XLT_SQL = XLT_SQL & " empr_cd_empresa = " & PCodEmpresa & "" & _
                                                " AND empd_cd_Empreendimento= '" & XLT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XLT_CODIMOVEL & "'" & _
                                                " AND cont_cd_Contrato='" & XLT_CODCONTRATO & "' AND titu_cd_Plano='" & XLT_CODPLANO & "'" & _
                                                " AND titu_cd_Parcela='" & XLT_CODPARCELA & "' AND titu_cd_Residuo='" & XLT_CODRESIDUO & "'"
                                        
                            SubQOpenRecordset XLO_IMPORTACAO, XLT_SQL, Estatico
                          
                            If Not XLO_IMPORTACAO.EOF Then
                          
                                'Veririfica se já foi importado
                                If IsNull(XLO_IMPORTACAO!titu_dt_Pagamento) Then
                                    'Acrescenta cada título importado à consulta do recibo
                                    XLT_SQLRECIBO = XLT_SQLRECIBO + XGT_CONJUNCAO + "Titulo = '" & XLT_CODEMPREENDIMENTO & "." & XLT_CODIMOVEL & "." & XLT_CODCONTRATO & "." & XLT_CODPLANO & "." & XLT_CODPARCELA & "." & XLT_CODRESIDUO & "'"
                                    XGT_CONJUNCAO = " OR "
                                    
                                    XLI_CONT = XLI_CONT + 1
                                                  
                                    'Verifica se exportado(pago) com o indexador 1 ,2 ou nenhum
                                    If Not IsNull(XLO_IMPORTACAO!moed_cd_Moeda3) Then
                                        If XLO_IMPORTACAO!moed_cd_Moeda3 = XLO_IMPORTACAO!moed_cd_Moeda1 Then
                                            XLF_PREVISTO = Format(funCalculaCorrecaoMonetaria(XLO_IMPORTACAO, 1, XLO_IMPORTACAO!titu_dt_Vencimento), "Standard")
                                        Else
                                            XLF_PREVISTO = Format(funCalculaCorrecaoMonetaria(XLO_IMPORTACAO, 2, XLO_IMPORTACAO!titu_dt_Vencimento), "Standard")
                                        End If
                                    Else
                                        XLF_PREVISTO = XLO_IMPORTACAO!titu_vl_Parcela
                                    End If
                                
                                    XLF_CORRECAO = XLF_VALORPAGO - XLF_JUROS + XLF_DESCONTO - _
                                    XLO_IMPORTACAO!titu_vl_Seguro
                              
                                    If XLO_IMPORTACAO!cont_vl_TaxaSegVida > 0 And XLO_IMPORTACAO!titu_tx_NaturezaPlano <> "F" _
                                        And XLO_IMPORTACAO!titu_vl_Seguro > 0 Then
                                        XLF_PREVISTO = XLF_PREVISTO + XLO_IMPORTACAO!titu_vl_Seguro
                                    End If
                                
                                    'Atualizacao do titulo
                                    XLT_SQL = "SELECT * FROM Titulos " & _
                                            "WHERE empr_cd_empresa = " & XLO_IMPORTACAO!empr_cd_empresa & "" & " AND " & _
                                            "empd_cd_empreendimento='" & XLO_IMPORTACAO!empd_cd_Empreendimento & "' AND " & _
                                            "imov_cd_imovel='" & XLO_IMPORTACAO!imov_cd_Imovel & "' AND " & _
                                            "cont_cd_Contrato='" & XLO_IMPORTACAO!cont_cd_Contrato & "' AND " & _
                                            "titu_cd_Plano= '" & XLO_IMPORTACAO!titu_cd_plano & "' AND " & _
                                            "titu_cd_Parcela= '" & XLO_IMPORTACAO!titu_cd_Parcela & "' AND " & _
                                            "titu_cd_Residuo= '" & XLO_IMPORTACAO!titu_cd_residuo & "'"
                                
                                    SubQOpenRecordset XLO_TITULO, XLT_SQL, Dinamico
                              
                                    XLO_TITULO!titu_vl_Juros = FunNuloVal(XLF_JUROS)
                                    XLO_TITULO!titu_vl_Desconto = FunNuloVal(XLF_DESCONTO)
                                    XLO_TITULO!titu_vl_ValorPago = FunNuloVal(XLF_VALORPAGO)
                                                
                                    If Not IsNull(XLO_IMPORTACAO!moed_cd_Moeda3) Then
                                        If XLO_IMPORTACAO!moed_cd_Moeda1 = XLO_IMPORTACAO!moed_cd_Moeda3 Then
                                            XLO_TITULO!titu_vl_Corrigido1 = FunNuloVal(XLF_CORRECAO)
                                        
                                            If XLO_IMPORTACAO!moed_cd_Moeda1 = XLO_IMPORTACAO!moed_cd_Moeda2 Then
                                                XLO_TITULO!titu_vl_Corrigido2 = FunNuloVal(XLF_CORRECAO)
                                            Else
                                                XLO_TITULO!titu_vl_Corrigido2 = 0
                                            End If
                                        
                                        Else
                                            XLO_TITULO!titu_vl_Corrigido2 = FunNuloVal(XLF_CORRECAO)
                                        
                                            If XLO_IMPORTACAO!moed_cd_Moeda1 = XLO_IMPORTACAO!moed_cd_Moeda2 Then
                                                XLO_TITULO!titu_vl_Corrigido1 = FunNuloVal(XLF_CORRECAO)
                                            Else
                                                XLO_TITULO!titu_vl_Corrigido1 = 0
                                            End If
                                        
                                        End If
                                    Else
                                        XLO_TITULO!titu_vl_Corrigido1 = 0
                                        XLO_TITULO!titu_vl_Corrigido2 = 0
                                    End If
                                
                                    XLO_TITULO!titu_dt_Pagamento = XLD_DATAPAG
                                    XLO_TITULO!titu_dt_BasePagto = XLO_IMPORTACAO!titu_dt_Vencimento
                                    XLO_TITULO!titu_dt_Deposito = XLD_DATADEPOSITO
                                    XLO_TITULO!titu_tx_LocalPagto = 3
                                    XLO_TITULO!titu_tx_DocPagto = XLT_DOCPGTO
                                    XLO_TITULO!coco_cd_codigo = CboCCorrente.BoundText
                    
                                    If Not IsNull(XLO_IMPORTACAO!moed_cd_Moeda3) Then
                                        If XLO_IMPORTACAO!moed_cd_Moeda3 = XLO_IMPORTACAO!moed_cd_Moeda1 Then
                                            XLO_TITULO!titu_tx_IndicePagamento = "1"
                                            XLO_TITULO!titu_tx_CorrigeParcela = "S"
                                        Else
                                            XLO_TITULO!titu_tx_IndicePagamento = "2"
                                            XLO_TITULO!titu_tx_CorrigeParcela = "S"
                                        End If
                                    Else
                                        XLF_PREVISTO = XLO_IMPORTACAO!titu_vl_Parcela
                                    End If
                                
                                    If IsNull(XLO_IMPORTACAO!moed_cd_Moeda3) Then
                                        XLT_TIPOPAG = funGeraTipoPag(0, XLF_VALORPAGO, XLF_VALORPAGO - XLF_JUROS + XLF_DESCONTO, XLF_PREVISTO, XLO_IMPORTACAO!titu_dt_Vencimento, XLD_DATAPAG)
                                    Else
                                        XLT_TIPOPAG = funGeraTipoPag(1, XLF_VALORPAGO, XLF_VALORPAGO - XLF_JUROS + XLF_DESCONTO, XLF_PREVISTO, XLO_IMPORTACAO!titu_dt_Vencimento, XLD_DATAPAG)
                                    End If
                                    
                                    XLO_TITULO!titu_tx_TipoPag = XLT_TIPOPAG
                                     
                                    '******************REGISTRA LOG DA OPERAÇÃO**************************
                                    ReDim XGM_MATRIZLOG(11, 2) As Variant
                                    
                                    XGM_MATRIZLOG(0, 0) = "Título"
                                    XGM_MATRIZLOG(1, 0) = "Valor Juros"
                                    XGM_MATRIZLOG(2, 0) = "Valor Corrigido1"
                                    XGM_MATRIZLOG(3, 0) = "Valor Corrigido2"
                                    XGM_MATRIZLOG(4, 0) = "Valor Pago"
                                    XGM_MATRIZLOG(5, 0) = "Data Pagto"
                                    XGM_MATRIZLOG(6, 0) = "Data Base de Pagto"
                                    XGM_MATRIZLOG(7, 0) = "Data Depósito"
                                    XGM_MATRIZLOG(8, 0) = "Doc. Pagto"
                                    XGM_MATRIZLOG(9, 0) = "Tipo Pagto"
                                    XGM_MATRIZLOG(10, 0) = "Conta corrente"
                                    
                                    XGM_MATRIZLOG(0, 1) = XLO_IMPORTACAO!empd_cd_Empreendimento & "." & _
                                                        XLO_IMPORTACAO!imov_cd_Imovel & "." & _
                                                        XLO_IMPORTACAO!cont_cd_Contrato & "." & _
                                                        XLO_IMPORTACAO!titu_cd_plano & "." & _
                                                        XLO_IMPORTACAO!titu_cd_Parcela & "." & _
                                                        XLO_IMPORTACAO!titu_cd_residuo
                                                        
                                    XGM_MATRIZLOG(1, 1) = FunNuloVal(XLF_JUROS)
                                    XGM_MATRIZLOG(2, 1) = XLO_TITULO!titu_vl_Corrigido1
                                    XGM_MATRIZLOG(3, 1) = XLO_TITULO!titu_vl_Corrigido2
                                    XGM_MATRIZLOG(4, 1) = FunNuloVal(XLF_VALORPAGO)
                                    XGM_MATRIZLOG(5, 1) = XLD_DATAPAG
                                    XGM_MATRIZLOG(6, 1) = XLO_IMPORTACAO!titu_dt_Vencimento
                                    XGM_MATRIZLOG(7, 1) = XLD_DATADEPOSITO
                                    XGM_MATRIZLOG(8, 1) = XLT_DOCPGTO
                                    XGM_MATRIZLOG(9, 1) = XLT_TIPOPAG
                                    XGM_MATRIZLOG(10, 1) = CboCCorrente.Text
                                    '*******************************************************************
                                                                
                                    XLO_TITULO.Update
                                    
                                    'Registra o log
                                    Call subRegistraLog("TelaExpBancaria", "2", funCriaDescricaoLog(XGM_MATRIZLOG, EXPORTACAO_CAPI, PAGAMENTO_ELETRONICO_BRADESCO))
                     
                                    
                                    '********ATUALIZA TÍTULO DO QUALIFIN E NA CONTABILIDADE********
                                    Select Case XLO_TITULO!titu_tx_NaturezaPlano
                                        Case "S"
                                            XLT_NATUREZAPLANO = "SINAL"
                                        Case "P"
                                            XLT_NATUREZAPLANO = "POUPANÇA"
                                        Case "F"
                                            XLT_NATUREZAPLANO = "FINANCIAMENTO"
                                        Case "G"
                                            XLT_NATUREZAPLANO = "FGTS"
                                        Case "O"
                                            XLT_NATUREZAPLANO = "OUTROS"
                                        Case "V"
                                            XLT_NATUREZAPLANO = "A VISTA"
                                    End Select
                                    
                                    SubQOpenRecordset XLO_FINANCEIRO, "SELECT focl_cd_FornCli, tipl_tx_Descricao FROM ConsCapTitulos WHERE titulo='" & XLT_CODEMPREENDIMENTO & "." & XLT_CODIMOVEL & "." & XLT_CODCONTRATO & "." & XLT_CODPLANO & "." & XLT_CODPARCELA & "." & XLT_CODRESIDUO & "' AND empr_cd_Empresa=" & PCodEmpresa, Estatico
                                    subGravaFinanceiro XLO_TITULO, "PAGAR", XLO_FINANCEIRO!focl_cd_FornCli, XLO_FINANCEIRO!tipl_tx_Descricao, XLT_NATUREZAPLANO
                                    
                                    XLO_FINANCEIRO.Close
                                    Set XLO_FINANCEIRO = Nothing
                                    '**************************************************************
                         
                                    XLO_TITULO.Close
                                    Set XLO_TITULO = Nothing
                                
                                End If
                            End If

                        End If
                    End If

                Wend
                Conexao.CommitTrans
            Else
                MsgBox "Esse não é um arquivo de retorno do Banco Bradesco!"
                XLB_ERRO = True
            End If
          
        Else
            MsgBox "Esse não é um arquivo de retorno de cobrança válido!"
            XLB_ERRO = True
        End If
      
        XLT_SQLRECIBO = XLT_SQLRECIBO + " )"

    Close #1
    'Se não houve erro e foi importado pelo menos 1 título
    If XLB_ERRO = False Then
        If XLI_CONT > 0 Then
            ImprimeRelatorio XLT_SQLRECIBO, Format(XLD_DATAPAG, "dd/mm/yy"), XGT_ARQUIVO, XLT_NUMARQUIVO, XGT_MODELORECIBO, XLT_DOCPGTO
            MsgBox "Importação de " & XLI_CONT & " título(s) realizada com sucesso!", vbInformation, ""
        Else
            If Not (IsDate(Mid(XLT_TEXTO, 296, 2) & "/" & Mid(XLT_TEXTO, 298, 2) & "/" & Mid(XLT_TEXTO, 300, 2))) Then 'Arquivo de inconsistências.
                ImprimeRelatorio "Select * from ConsCAPOcorreciasImportacao Where empr_cd_empresa = " & PCodEmpresa & " and reoc_nr_Lote = " & XLT_NUMARQUIVO & " Order by Titulo", Mid(XLT_TEXTO, 296, 6), XGT_ARQUIVO, XLT_NUMARQUIVO, XGT_MODELORECIBO, XLT_DOCPGTO
                MsgBox "O arquivo foi lido com sucesso!", vbInformation, "ATENÇÃO"
            Else
                MsgBox "Nenhum título foi importado!", vbInformation, "ATENÇÃO"
            End If
        End If
        Unload Me
    End If
    
    Me.MousePointer = vbDefault
  
End Sub

Private Sub CmdRetornar_Click()
  
  Unload Me
  
End Sub

Private Sub Command1_Click()
  
  Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\..\Relatorios\RelRecibo.rpt")
  
End Sub

Private Sub Form_Activate()
    
  Set Formulario = FrmImportacao
    
End Sub

Private Sub Form_Load()
  
  subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & Int(PCodEmpresa) & " AND banc_cd_codigo=237 ORDER BY coco_tx_Descricao", Estatico
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Call subFechaConexaoQualifin
End Sub

Private Sub ImprimeRelatorio(Comando As String, DataPagamento As String, NomeArquivo As String, Lote As String, ModeloRecibo As String, DocPagto As String)

       Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(Comando)
       
       If IsDate(DataPagamento) Then
           If Not XGO_RSRELATORIO.EOF Then
    
                Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
                  
                'Carrega os relatórios
                FrmVisRelatorios.Caption = "Impressão de Recibos"
                
                If XGT_MODELORECIBO = "1" Then
                    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\..\Relatorios\RelRecibo.rpt")
                Else
                    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\..\Relatorios\RelRecibo2.rpt")
                End If
        
                'Passa o recordset para o relatório
                XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
                
                'Ativa os Relatórios
                With XGR_RELATORIO.FormulaFields
                    .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
                    .GetItemByName("docpgto").Text = Chr(34) & DocPagto & Chr(34)
                End With
        
                FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
                FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Rodape"
                FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
      
           End If

       Else
                
                Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
                FrmVisRelatorios.Caption = "Relatório de Ocorrências"
                Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelOcorrenciasCapi.rpt")
                
                XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
                With XGR_RELATORIO.FormulaFields
                    .GetItemByName("LoteArquivo").Text = Chr(34) & Lote & Chr(34)
                    .GetItemByName("NomeArquivo").Text = Chr(34) & NomeArquivo & Chr(34)
                    .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
                    
                End With
                
                FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
                FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
       End If

End Sub




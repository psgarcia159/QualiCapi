VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmImportacao 
   Caption         =   "Retorno de Pagamento Eletrônico - SUDAMERIS"
   ClientHeight    =   1785
   ClientLeft      =   1995
   ClientTop       =   1515
   ClientWidth     =   9120
   Icon            =   "FrmImportacao.frx":0000
   LinkTopic       =   "Form1"
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

Sub subImprimePendencias(XLT_LOTE As String)
  Dim XLT_FiltroLote As String
  
  'Abre o recordset
  Set XGO_RSRELATORIO = Conexao.Execute("SELECT * FROM ConsFINOcorreciasImportacao WHERE reoc_nr_Lote = '" & XLT_LOTE & "'")
  Set XGO_RSCABECALHO = Conexao.Execute("SELECT * FROM Empresas")
  
  'Verifica se o recordset possui registros
  If XGO_RSRELATORIO.EOF Then
    MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
    Exit Sub
  End If
  
  'Carrega o relatório
  Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelOcorrencias.rpt")
  
  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
  
  'Ativa os Relatórios
  With XGR_RELATORIO.FormulaFields
    .GetItemByName("Codigo").Text = PCodEmpresa
    .GetItemByName("FiltroLote").Text = Chr(34) & XLT_LOTE & Chr(34)
  End With
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
End Sub

Private Sub CmdConfirmar_Click()
  
  Dim XLT_TEXTO As String     'Armazena o arquivo texto
  Dim XLT_TITULO As String    'Armazena o nº do título (nº da fatura)
  Dim XLT_PARCELA As String   'Armazena o cód. da parcela
  Dim XLD_DATAPAG As Date     'Data de Pagamento
  Dim XLD_DATADEPOSITO As Date 'Data de Deposito
  Dim XLF_VALORPAGO As Double 'Valor pago
  
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
  Dim XLO_IMPORTACAO As New ADODB.Recordset
  Dim XLO_TITULO As New ADODB.Recordset
  Dim XLO_CODTIPOPLANO As New ADODB.Recordset
  
  Dim XLF_RESIDUO As Double
  Dim XLF_VALORINDEXADO1 As Double
  Dim XLF_VALORINDEXADO2 As Double
  Dim XLT_TIPOPAG As String
  Dim XLI_CODTIPOPLANO As Integer
  
  Dim XLT_NATUREZAPLANO As String       'Utilizado  para informar a natureza a rotina de gravação no financeiro
  
  DlgAbrirFigura.CancelError = False
  DlgAbrirFigura.Filter = "Arquivo de Retorno do Itaú (*.ret)|*.ret"
  DlgAbrirFigura.ShowOpen
  
  If DlgAbrirFigura.FileName = "" Then
      Unload Me
      Exit Sub
  End If
  
  If DlgAbrirFigura.FileName <> "" Then
      XGT_LOCALARQ = DlgAbrirFigura.FileName
      XGT_ARQUIVO = DlgAbrirFigura.FileTitle
  End If
  
  XLT_CODEMPRESA = Format(Left(DatContaCorrente.Recordset.Fields!coco_cd_Agencia, 3), "000") & _
                   Space(1) & _
                   Format(Left(DatContaCorrente.Recordset.Fields!coco_tx_Conta, 8), "00000000")

  Open XGT_LOCALARQ For Input Access Read As #1
  
  While Not EOF(1)
    
    Input #1, XLT_TEXTO
    If Left(XLT_TEXTO, 1) = "1" Then
    
      If Mid(XLT_TEXTO, 18, 12) = XLT_CODEMPRESA Then
    
        If Mid(XLT_TEXTO, 109, 2) = "06" Or Mid(XLT_TEXTO, 109, 2) = "07" Then
          
          XLT_TITULO = Mid(XLT_TEXTO, 38, 23)
          XLT_CODEMPREENDIMENTO = Mid(XLT_TEXTO, 38, 4)
          XLT_CODIMOVEL = Mid(XLT_TEXTO, 43, 4)
          XLT_CODCONTRATO = Mid(XLT_TEXTO, 48, 2)
          XLT_CODPLANO = Mid(XLT_TEXTO, 51, 2)
          XLT_CODPARCELA = Mid(XLT_TEXTO, 54, 3)
          XLT_CODRESIDUO = Mid(XLT_TEXTO, 58, 2)
          XLT_PARCELA = Right(Trim(Mid(XLT_TEXTO, 74, 20)), 1)
          XLD_DATAPAG = Mid(XLT_TEXTO, 111, 2) & _
                          "/" & Mid(XLT_TEXTO, 113, 2) & _
                          "/" & Mid(XLT_TEXTO, 115, 2)
          XLD_DATADEPOSITO = Mid(XLT_TEXTO, 296, 2) & _
                           "/" & Mid(XLT_TEXTO, 298, 2) & _
                           "/" & Mid(XLT_TEXTO, 300, 2)
          XLF_VALORPAGO = CDbl(Mid(XLT_TEXTO, 254, 13) / 100)
          XLT_CARTEIRA = Mid(XLT_TEXTO, 83, 3)
          XLF_JUROS = Mid(XLT_TEXTO, 267, 13) / 100
          XLF_DESCONTO = Mid(XLT_TEXTO, 241, 13) / 100
             
          XLT_SQL = "" & _
          "SELECT ConsCAPImpBanco.*," & _
          "CASE Indexacao1 when 'M' then" & _
            " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda " & _
               "and cota_dt_datacotacao=dateadd(dd,-day(ConsCAPImpBanco.titu_dt_Vencimento)+1,ConsCAPImpBanco.titu_dt_Vencimento)) " & _
            " when 'D' then" & _
            " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda " & _
               "and cota_dt_datacotacao=ConsCAPImpBanco.titu_dt_Vencimento)" & _
          " end as Cotacao1," & _
          "CASE  Indexacao2 when 'M' then" & _
            " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda " & _
              "and cota_dt_datacotacao=dateadd(dd,-day(ConsCAPImpBanco.titu_dt_Vencimento)+1,ConsCAPImpBanco.titu_dt_Vencimento)) " & _
            " when 'D'  then" & _
            " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda " & _
              "and cota_dt_datacotacao=ConsCAPImpBanco.titu_dt_Vencimento)" & _
          " end as Cotacao2,"
          
          XLT_SQL = XLT_SQL & _
            "CASE  Indexacao1 when 'M' then" & _
              " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda1=cotacoesmoedas.moed_cd_CodMoeda " & _
                "and cota_dt_datacotacao=dateadd(mm,-1,dateadd(dd,-day(ConsCAPImpBanco.titu_dt_Vencimento)+1,ConsCAPImpBanco.titu_dt_Vencimento)))" & _
            " end as Cotacao1_MesAnterior," & _
            "CASE Indexacao2 when 'M' then" & _
              " (select cotacoesmoedas.cota_vl_Valor from cotacoesmoedas where moed_cd_Moeda2=cotacoesmoedas.moed_cd_CodMoeda " & _
                "and cota_dt_datacotacao=dateadd(mm,-1,dateadd(dd,-day(ConsCAPImpBanco.titu_dt_Vencimento)+1,ConsCAPImpBanco.titu_dt_Vencimento)))" & _
            " end As Cotacao2_MesAnterior From ConsCAPImpBanco"
    
          XLT_SQL = XLT_SQL & _
            " WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
            " AND empd_cd_Empreendimento= '" & XLT_CODEMPREENDIMENTO & "' AND imov_cd_Imovel='" & XLT_CODIMOVEL & "'" & _
            " AND cont_cd_Contrato='" & XLT_CODCONTRATO & "' AND titu_cd_Plano='" & XLT_CODPLANO & "'" & _
            " AND titu_cd_Parcela='" & XLT_CODPARCELA & "' AND titu_cd_Residuo='" & XLT_CODRESIDUO & "'"
          
          SubQOpenRecordset XLO_IMPORTACAO, XLT_SQL, Estatico
          
          'Veririfica se já foi importado
          If IsNull(XLO_IMPORTACAO!titu_dt_Pagamento) Then
          
            'Verifica se exportado(pago) com o indexador 1 ,2 ou nenhum
            If Not IsNull(XLO_IMPORTACAO!moed_cd_Moeda3) Then
              If XLO_IMPORTACAO!moed_cd_Moeda3 = XLO_IMPORTACAO!moed_cd_Moeda1 Then
                XLF_PREVISTO = funCalculaCorrecaoMonetaria(XLO_IMPORTACAO, 1, XLO_IMPORTACAO!titu_dt_Vencimento)
              Else
                XLF_PREVISTO = funCalculaCorrecaoMonetaria(XLO_IMPORTACAO, 2, XLO_IMPORTACAO!titu_dt_Vencimento)
              End If
            Else
              XLF_PREVISTO = XLO_IMPORTACAO!titu_vl_Parcela
            End If
          
            If XLO_IMPORTACAO!cont_vl_TaxaSegVida > 0 And XLO_IMPORTACAO!titu_tx_NaturezaPlano <> "F" _
              And XLO_IMPORTACAO!titu_vl_Seguro > 0 Then
              XLF_PREVISTO = XLF_PREVISTO + XLO_IMPORTACAO!titu_vl_Seguro
            End If
            
            'Atualizacao do titulo
            XLT_SQL = "SELECT * FROM Titulos " & _
                       "WHERE empr_cd_empresa = " & XLO_IMPORTACAO!empr_cd_Empresa & "" & " AND " & _
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
            XLO_TITULO!titu_vl_Corrigido1 = FunNuloVal(XLF_VALORINDEXADO1)
            XLO_TITULO!titu_vl_Corrigido2 = FunNuloVal(XLF_VALORINDEXADO2)
            XLO_TITULO!titu_dt_Pagamento = XLD_DATAPAG
            XLO_TITULO!titu_dt_BasePagto = XLO_IMPORTACAO!titu_dt_Vencimento
            XLO_TITULO!titu_dt_Deposito = XLD_DATADEPOSITO
            XLO_TITULO!titu_tx_LocalPagto = 2
            XLO_TITULO!titu_tx_DocPagto = XGT_ARQUIVO
            XLO_TITULO!coco_cd_Codigo = CboCCorrente.BoundText
            
            If IsNull(XLO_IMPORTACAO!moed_cd_Moeda3) Then
             XLT_TIPOPAG = funGeraTipoPag(0, XLF_VALORPAGO, XLF_VALORPAGO - XLF_JUROS + XLF_DESCONTO, XLF_PREVISTO, XLO_IMPORTACAO!titu_dt_Vencimento, XLD_DATAPAG)
            Else
             XLT_TIPOPAG = funGeraTipoPag(1, XLF_VALORPAGO, XLF_VALORPAGO - XLF_JUROS + XLF_DESCONTO, XLF_PREVISTO, XLO_IMPORTACAO!titu_dt_Vencimento, XLD_DATAPAG)
            End If
          
            XLO_TITULO!titu_tx_TipoPag = XLT_TIPOPAG
             
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
                                    
            XLO_TITULO.Update
              
            '******************REGISTRA LOG DA OPERAÇÃO**************************
            ReDim XGM_MATRIZLOG(12, 2) As Variant
            
            XGM_MATRIZLOG(0, 0) = "Título"
            XGM_MATRIZLOG(1, 0) = "Valor Juros"
            XGM_MATRIZLOG(2, 0) = "Valor Desconto"
            XGM_MATRIZLOG(3, 0) = "Valor Corrigido1"
            XGM_MATRIZLOG(4, 0) = "Valor Corrigido2"
            XGM_MATRIZLOG(5, 0) = "Valor Pago"
            XGM_MATRIZLOG(6, 0) = "Data Pagto"
            XGM_MATRIZLOG(7, 0) = "Data Base de Pagto"
            XGM_MATRIZLOG(8, 0) = "Data Depósito"
            XGM_MATRIZLOG(9, 0) = "Doc. Pagto"
            XGM_MATRIZLOG(10, 0) = "Tipo Pagto"
            XGM_MATRIZLOG(11, 0) = "Conta corrente"
            
            XGM_MATRIZLOG(0, 1) = XLO_IMPORTACAO!empd_cd_Empreendimento & "." & _
                                XLO_IMPORTACAO!imov_cd_Imovel & "." & _
                                XLO_IMPORTACAO!cont_cd_Contrato & "." & _
                                XLO_IMPORTACAO!titu_cd_plano & "." & _
                                XLO_IMPORTACAO!titu_cd_Parcela & "." & _
                                XLO_IMPORTACAO!titu_cd_residuo
                                
            XGM_MATRIZLOG(1, 1) = FunNuloVal(XLF_JUROS)
            XGM_MATRIZLOG(2, 1) = FunNuloVal(XLF_DESCONTO)
            XGM_MATRIZLOG(3, 1) = FunNuloVal(XLF_VALORINDEXADO1)
            XGM_MATRIZLOG(4, 1) = FunNuloVal(XLF_VALORINDEXADO2)
            XGM_MATRIZLOG(5, 1) = FunNuloVal(XLF_VALORPAGO)
            XGM_MATRIZLOG(6, 1) = XLD_DATAPAG
            XGM_MATRIZLOG(7, 1) = XLO_IMPORTACAO!titu_dt_Vencimento
            XGM_MATRIZLOG(8, 1) = XLD_DATADEPOSITO
            XGM_MATRIZLOG(9, 1) = XGT_ARQUIVO
            XGM_MATRIZLOG(10, 1) = XLT_TIPOPAG
            XGM_MATRIZLOG(11, 1) = CboCCorrente.Text
            
            Call subRegistraLog("TelaExpBancaria", "2", funCriaDescricaoLog(XGM_MATRIZLOG, EXPORTACAO_CAPI, PAGAMENTO_ELETRONICO_SUDAMERIS))
            '*******************************************************************
                           
              
            '********ATUALIZA TÍTULO DO QUALIFIN E NA CONTABILIDADE********
            If XGB_PQUALIFIN = True Then
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
                subGravaFinanceiro XLO_TITULO, "PAGAR", XLO_TITULO!focl_cd_FornCli, XLO_TITULO!tipl_tx_Descricao, XLT_NATUREZAPLANO
            End If
            '**************************************************************
              
            XLO_TITULO.Close
            Set XLO_TITULO = Nothing
            
           'Se for gerar residuo
            If (XLF_VALORPAGO) < XLF_PREVISTO Then
              XLT_CODRESIDUO = Format(CInt(Right(XLO_IMPORTACAO!Titulo, 2) + 1), "00")
              XLF_RESIDUO = XLF_PREVISTO - (XLF_VALORPAGO - XLF_JUROS + XLF_DESCONTO)
              XLF_VALORINDEXADO1 = FunCalculaValorIndexadoResiduo(XLO_IMPORTACAO, 1, XLO_IMPORTACAO!titu_dt_Vencimento, XLF_RESIDUO)
              XLF_VALORINDEXADO2 = FunCalculaValorIndexadoResiduo(XLO_IMPORTACAO, 2, XLO_IMPORTACAO!titu_dt_Vencimento, XLF_RESIDUO)
              
              'Procura o código do tipo resíduo
              SubQOpenRecordset XLO_CODTIPOPLANO, "SELECT tipl_cd_TipoPlano FROM TiposPlanos WHERE tipl_tx_Sigla='R'", Estatico
              XLI_CODTIPOPLANO = XLO_CODTIPOPLANO!tipl_cd_TipoPlano
              XLO_CODTIPOPLANO.Close
              Set XLO_CODTIPOPLANO = Nothing
              
              XLT_SQL = "SELECT * FROM Titulos " & _
                       "WHERE empr_cd_empresa = -121179"
              
              SubQOpenRecordset XLO_TITULO, XLT_SQL, Dinamico
              
              XLO_TITULO.AddNew
              
              XLO_TITULO!empr_cd_Empresa = XLO_IMPORTACAO!empr_cd_Empresa
              XLO_TITULO!empd_cd_Empreendimento = XLO_IMPORTACAO!empd_cd_Empreendimento
              XLO_TITULO!imov_cd_Imovel = XLO_IMPORTACAO!imov_cd_Imovel
              XLO_TITULO!cont_cd_Contrato = XLO_IMPORTACAO!cont_cd_Contrato
              XLO_TITULO!titu_cd_plano = XLO_IMPORTACAO!titu_cd_plano
              XLO_TITULO!titu_cd_Parcela = XLO_IMPORTACAO!titu_cd_Parcela
              XLO_TITULO!titu_cd_residuo = XLT_CODRESIDUO
              XLO_TITULO!obse_cd_Observacao = XLO_IMPORTACAO!obse_cd_Observacao
              XLO_TITULO!moed_cd_Moeda1 = XLO_IMPORTACAO!moed_cd_Moeda1
              XLO_TITULO!moed_cd_Moeda2 = XLO_IMPORTACAO!moed_cd_Moeda2
              XLO_TITULO!titu_vl_Parcela = XLF_RESIDUO
              XLO_TITULO!titu_vl_ParcelaIndex1 = XLF_VALORINDEXADO1
              XLO_TITULO!titu_vl_ParcelaIndex2 = XLF_VALORINDEXADO2
              XLO_TITULO!titu_dt_Vencimento = XLO_IMPORTACAO!titu_dt_Vencimento
              XLO_TITULO!tipl_cd_TipoPlano = XLI_CODTIPOPLANO
              XLO_TITULO!titu_vl_PercJurosFin = XLO_IMPORTACAO!titu_vl_PercJurosFin
              XLO_TITULO!titu_vl_PercMultaMora = XLO_IMPORTACAO!titu_vl_PercMultaMora
              XLO_TITULO!titu_vl_PercJurosMora = XLO_IMPORTACAO!titu_vl_PercJurosMora
              XLO_TITULO!cere_cd_Pcr = XLO_IMPORTACAO!cere_cd_Pcr
              XLO_TITULO!plco_cd_Conta = XLO_IMPORTACAO!plco_cd_Conta
              XLO_TITULO!coco_cd_Codigo = XLO_IMPORTACAO!coco_cd_Codigo
              XLO_TITULO!titu_dt_Base = XLO_IMPORTACAO!titu_dt_Vencimento
              XLO_TITULO!titu_dt_Vencimento = XLO_IMPORTACAO!titu_dt_Vencimento
              XLO_TITULO!titu_dt_Prorrogacao = XLO_IMPORTACAO!titu_dt_Vencimento
              XLO_TITULO!titu_tx_NaturezaPlano = XLO_IMPORTACAO!titu_tx_NaturezaPlano
              XLO_TITULO!titu_dt_Criacao = XLD_DATAPAG
              'XLO_TITULO!iden_cd_Projeto = XLO_IMPORTACAO!iden_cd_Projeto
                                 
              XLO_TITULO.Update
              
              XLO_TITULO.Close
              Set XLO_TITULO = Nothing
            End If
          End If
        End If
      End If
    End If
        
  Wend
  MsgBox ("Importação realizada c/ sucesso !")
Close #1
    
    'If (MsgBox("Deseja imprimir Relatório de Ocorrências ?", vbYesNo + vbQuestion, "ATENÇÃO")) = 6 Then
    '    Call subImprimePendencias(XLT_LOTE)
    'End If
  
End Sub

Private Sub CmdRetornar_Click()
  
  Unload Me
  
End Sub

Private Sub Form_Activate()
    
  Set Formulario = FrmImportacao
    
End Sub

Private Sub Form_Load()
  
  subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & Int(PCodEmpresa) & " AND banc_cd_codigo=347 ORDER BY coco_tx_Descricao", Estatico
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Call subFechaConexaoQualifin
End Sub





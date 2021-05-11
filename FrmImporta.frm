VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form FrmImporta 
   Caption         =   "Importação de dados do MFIN para o QUALIFIN"
   ClientHeight    =   3030
   ClientLeft      =   2205
   ClientTop       =   3000
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   6585
   Begin Threed.SSPanel PanEmpresas 
      Height          =   1740
      Left            =   555
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   5385
      _Version        =   65536
      _ExtentX        =   9499
      _ExtentY        =   3069
      _StockProps     =   15
      BackColor       =   12640511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Possui Qualifin"
         Height          =   420
         Left            =   1335
         TabIndex        =   16
         Top             =   1260
         Width           =   1770
         Begin VB.OptionButton OptNao 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Não"
            Height          =   195
            Left            =   1005
            TabIndex        =   18
            Top             =   195
            Width           =   720
         End
         Begin VB.OptionButton OptSim 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Sim"
            Height          =   195
            Left            =   180
            TabIndex        =   17
            Top             =   195
            Value           =   -1  'True
            Width           =   720
         End
      End
      Begin VB.CommandButton CmdProcuraBD 
         Caption         =   "&Procura"
         Height          =   270
         Left            =   3240
         TabIndex        =   9
         Top             =   1365
         Width           =   1005
      End
      Begin VB.CommandButton CmdContinua 
         Caption         =   "Continuar"
         Enabled         =   0   'False
         Height          =   270
         Left            =   4290
         TabIndex        =   6
         Top             =   1365
         Width           =   1005
      End
      Begin MSDataListLib.DataCombo CboEmpresas 
         Bindings        =   "FrmImporta.frx":0000
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   390
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "empr_tx_fantasia"
         BoundColumn     =   "empr_cd_empresa"
         Text            =   "CboEmpresas"
      End
      Begin MSAdodcLib.Adodc DatEmpresas 
         Height          =   330
         Left            =   -750
         Top             =   1350
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
         Caption         =   "DatEmpresas"
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
      Begin Threed.SSPanel TxtArquivoBD 
         DataField       =   "RaizBanco"
         DataSource      =   "DataPadroes"
         Height          =   375
         Left            =   1350
         TabIndex        =   7
         Top             =   870
         Width           =   3945
         _Version        =   65536
         _ExtentX        =   6959
         _ExtentY        =   661
         _StockProps     =   15
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   1
         BevelInner      =   1
         Alignment       =   1
      End
      Begin MSComDlg.CommonDialog DlgAbrirFigura 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox MskPlanoC 
         Height          =   285
         Left            =   3495
         TabIndex        =   10
         Top             =   45
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "99.99.99"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MskPcr 
         Height          =   285
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "99.99.99"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MskCep 
         Height          =   315
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##.###-###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskTelefone 
         Height          =   315
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
         Mask            =   "(####)####-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtCgc 
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         PromptChar      =   "_"
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nome Arquivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   915
         Width           =   1290
      End
      Begin VB.Label LblEmpresas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Empresas:"
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
         Left            =   390
         TabIndex        =   5
         Top             =   435
         Width           =   915
      End
   End
   Begin Threed.SSPanel PanMensagem 
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   675
      Width           =   4560
      _Version        =   65536
      _ExtentX        =   8043
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Importando Banco de Dados"
      ForeColor       =   12582912
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Begin Threed.SSPanel LblNReg 
         Height          =   330
         Left            =   1395
         TabIndex        =   15
         Top             =   285
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   582
         _StockProps     =   15
         BackColor       =   13160660
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
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   345
      Left            =   1665
      TabIndex        =   1
      Top             =   2595
      Width           =   1335
   End
   Begin VB.CommandButton CmdDesistir 
      Caption         =   "Desistir"
      Height          =   345
      Left            =   3435
      TabIndex        =   0
      Top             =   2595
      Width           =   1335
   End
End
Attribute VB_Name = "FrmImporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsetFonte As Object       'Recordset temporário com a tabela a ser exportada
Dim BancoSADE As Object
Dim xCodigo1 As Long           'Código para relacionar documentos
Dim xCodigo2 As Long           'Código para relacionar documentos

Dim Campos As New XArray
Dim Campos2 As New XArray

Dim XCodEmpresa As Integer
Dim XLT_EMPRESAS As String 'Concatena o código das empresas que já foram importadas

Sub subGravaVetorQualifin(Tabela As String, CamposTabela As String, Tipos As Variant, QtdCampos As Byte, Colunas As Variant, Complemento As String, Vetor As XArray)
    Dim Valores As String
    Dim x As Integer, y As Integer
    'Tabela - nome da tabela onde serão inseridos os registros do grid
    'CamposTabela - string simples com os campos na respectiva ordem do grid
    
    'Tipos - vetor simples com os respectivos tipos de dados das colunas do grid
    'N: Número; D: Data; T: Texto
    
    'QtdCampos - quantidade de colunas do grid que serão gravadas
    'Colunas - vetor simples com o índice das respectivas colunas que serão gravadas
    'Complemento - Valores de campos fixos que deseje gravar e não existe no grid para evitar
    'repetição Ex.: Todos os campos possuem o mesmo código de serviço
    
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    If Vetor(0, 0) <> Empty Then
        If Complemento <> "" Then
            Complemento = ", " & Complemento
        End If
        
        For x = 0 To Vetor.UpperBound(1)
            Valores = ""
            For y = 0 To QtdCampos - 1
                If Valores <> "" Then
                    If Tipos(y) = "N" Then
                        Valores = Valores & ", " & Str(Vetor(x, Colunas(y)))
                    ElseIf Tipos(y) = "T" Then
                        Valores = Valores & ", '" & Vetor(x, Colunas(y)) & "'"
                    ElseIf Tipos(y) = "D" Then
                        Valores = Valores & "," & FunNuloData(Vetor(x, Colunas(y)), NomeSgbd) & ""
                    End If
                Else
                    If Tipos(y) = "N" Then
                        Valores = Str(Vetor(x, Colunas(y)))
                    ElseIf Tipos(y) = "T" Then
                        Valores = "'" & Vetor(x, Colunas(y)) & "'"
                    ElseIf Tipos(y) = "D" Then
                        Valores = "" & FunNuloData(Vetor(x, Colunas(y)), NomeSgbd) & ""
                    End If
                End If
            Next
    
            Conexao.Execute "INSERT INTO " & Tabela & " (" & CamposTabela & ") VALUES (" & Valores & Complemento & ")"
        Next
    End If
End Sub


Private Sub CmdConfirmar_Click()
    Dim ResPadrao As Object
    Dim XLO_COTACAO As Object
    Dim y As Integer
    Dim x As Integer
       
    ConexaoMfin.Provider = "Microsoft.Jet.OLEDB.4.0"
    ConexaoMfin.Open App.Path & "\CapiCAndrade.mdb", "admin", ""
    'ConexaoMfin.Open App.Path & "\CapiActitur.mdb", "admin", ""
         
      'Abre tabela de TipMoeda do Mfin
      PanMensagem.Caption = "Importando Indexadores"
      subQOpenRecordset2 ResPadrao, "SELECT * FROM TipMoeda", Estatico
      Campos.ReDim 0, 0, 0, 5
      While Not ResPadrao.EOF
          subInsereRegistroCampos Array(ResPadrao!Codigo, ResPadrao!Nome, ResPadrao!simbolo, ResPadrao!Tipo), Campos, 3
          ResPadrao.MoveNext
      Wend
      ResPadrao.Close

      'Insere tabela de Moedas do Qualifin
      For x = 0 To Campos.UpperBound(1)
          SubQOpenRecordset ResPadrao, "SELECT * FROM Moedas WHERE moed_tx_descricao = '" & Campos(x, 1) & "'", Estatico
          If ResPadrao.EOF Then 'Ainda não foi cadastrada
              ResPadrao.Close
              SubQOpenRecordset ResPadrao, "SELECT * FROM Moedas WHERE moed_cd_codmoeda= -1", Dinamico
              ResPadrao.AddNew
    
              ResPadrao!moed_tx_descricao = Campos(x, 1)
              ResPadrao!moed_tx_simbolo = Campos(x, 2)
              ResPadrao!moed_tx_indexacao = Campos(x, 3)
                       
              ResPadrao.Update
          End If
          subSetaVal Campos, x, 4, ResPadrao!moed_cd_CodMoeda
          ResPadrao.Close
      Next

      'Abre tabela de TipMoeda do Mfin
      PanMensagem.Caption = "Importando Cotações"
      x = 0
      y = 0
      Campos2.ReDim 0, 0, 0, 3
      For x = 0 To Campos.UpperBound(1)
          subQOpenRecordset2 ResPadrao, "SELECT * FROM UMON WHERE CODI_UMON = '" & Campos(x, 0) & "'", Estatico
          While Not ResPadrao.EOF
              'Gravar tabela Cotacoes do Qualifin
              If Not IsNull(ResPadrao!Data_umon) And Not IsNull(ResPadrao!valo_umon) Then
                  '*&Sql = "INSERT INTO CotacoesMoedas(moed_cd_CodMoeda,cota_dt_DataCotacao,cota_vl_Valor)" & _
                        " VALUES (" & Campos(x, 4) & "," & FunNuloData(ResPadrao!Data, NomeSgbd) & "," & FunTrataFloat(ResPadrao!Taxa) & ")"
                  '*&Conexao.Execute Sql
                  SubQOpenRecordset XLO_COTACAO, "SELECT * FROM CotacoesMoedas WHERE moed_cd_CodMoeda=" & Campos(x, 4) & " and cota_dt_DataCotacao = " & FunNuloData(ResPadrao!Data_umon, NomeSgbd) & "", Dinamico
                  If XLO_COTACAO.EOF Then
                      XLO_COTACAO.AddNew
                      XLO_COTACAO!moed_cd_CodMoeda = CLng(Campos(x, 4))
                      XLO_COTACAO!cota_dt_DataCotacao = ResPadrao!Data_umon
                      XLO_COTACAO!cota_vl_Valor = ResPadrao!valo_umon
                      
                      XLO_COTACAO.Update
                  End If
                  XLO_COTACAO.Close
              End If
              
              ResPadrao.MoveNext
          Wend
          ResPadrao.Close
      Next
      
      'Gravação das Cartas
      PanMensagem.Caption = "Importando Cartas"
      subQOpenRecordset2 ResPadrao, "SELECT * FROM Carta", Estatico
   
      While Not ResPadrao.EOF
          Sql = "INSERT INTO Cartas(cart_tx_Descricao,cart_tx_Texto)" & _
                  " VALUES ('" & ResPadrao!Descricao & "','" & ResPadrao!texto & "')"
          Conexao.Execute Sql
          ResPadrao.MoveNext
      Wend
      ResPadrao.Close
      
      'Gravação das Promissórias
      PanMensagem.Caption = "Importando Promissórias"
      subQOpenRecordset2 ResPadrao, "SELECT * FROM PROM", Estatico
   
      While Not ResPadrao.EOF
          Sql = "INSERT INTO LayOutPromissorias(lapr_cd_Promissoria,lapr_tx_Descricao," & _
                "lapr_nr_Linha,lapr_nr_Coluna,lapr_tx_Imprime,lapr_tx_Letra,lapr_nr_Polegadas)" & _
                " VALUES ('" & ResPadrao!Codigo & "','" & ResPadrao!Descricao & "'," & _
                "" & FunTrataFloat(FunNuloVal(ResPadrao!linha)) & "," & FunTrataFloat(FunNuloVal(ResPadrao!coluna)) & ",'" & ResPadrao!Imprime & "','" & ResPadrao!Letra & "'," & FunTrataFloat(FunNuloVal(ResPadrao!Polegadas)) & ")"
          Conexao.Execute Sql
          ResPadrao.MoveNext
      Wend
      ResPadrao.Close
   
    ConexaoMfin.Close
      
      'Abre um painel para a escolha da empresa
      PanEmpresas.Visible = True
      subConectarControleDadosNV DatEmpresas, "SELECT * FROM Empresas", Estatico
End Sub

Sub subInsereRegistroCampos(VetorCampos As Variant, Vetor As XArray, Tamanho As Integer)  'Carrega o vetores
    Dim Cont As Integer
    Dim x As Integer
    
    'Insere um regitro no vetor
    
    'VetorCampos - vetor simples com os valores a inserir na respectiva ordem do grid
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    
    If Vetor.UpperBound(1) <> 0 Or Vetor(0, 0) <> Empty Then
        Vetor.Insert 1, Vetor.UpperBound(1) + 1
    End If
    Cont = Vetor.UpperBound(1)
    For x = 0 To Tamanho
        subSetaVal Vetor, Cont, x, FunNulo(VetorCampos(x))
    Next
'    With Grid
'        .ReBind
'    End With
End Sub
    
Private Sub CmdContinua_Click()
    Dim ResPadrao As Object
    Dim XLO_QUALIFIN As Object
    Dim XLO_DOS As Object
    Dim x As Integer, xNumFat As Long, y As Integer, j As Integer
    Dim XCodNovo As String, XCCorrente As String, XBForn As String
    Dim XcodEstrut As String, XConta As String, XConcilia As String
    Dim Pcr As String, ContaRecDesp As String, Xnatureza As String
    Dim ContaCusto As String, ContaProvAdiant As String
    Dim XLI_TAMNIVEL As Integer     'Tamanho do nível do cód. estruturado
    Dim XLT_NATUREZA As String      'Natureza do plano
    Dim XLT_TIPODESAGIO As String   'Tipo de deságio (D-Diario ou M-Mensal)
    Dim XLI_NIVELPCR As Integer     'Nivel do Pcr
    Dim XLI_X As Integer            'Auxiliar da definição do Nivel do Pcr
    Dim XLB_IMPORTOU As Boolean     'Indica se já importou alguma empresa
    Dim XLT_SQL As String           'Prepara select
    Dim XLI_DESCACRESC As Integer
    Dim XLL_CODDUPL As Long         'Código da duplicata para gravar no desc acresc
    Dim XLI_EMPRESATEMP As Integer
    Dim XLI_OBS As Integer          'Código da obs
    Dim XLI_MOEDA1 As Integer       'Código do indexador 1
    Dim XLI_MOEDA2 As Integer       'Código do indexador 2
    Dim XLI_TIPOPLANO As Integer    ' Código do tipo plano
    Dim XLL_CODCLIENTE As Long      'Código do cliente
    Dim XLT_TITULO As String
    Dim XLD_DTBASE As Date          'Dta base do contrato e do titulo
    Dim XLF_PERCDESAGIO As Double   'Percentual Desagio
    Dim XLF_PERCJUROS As Double     'Percentual Juros
    Dim XLF_PERCMULTA As Double     'Percentual Multa
    Dim XLD_PRORROGACAO As Date     'Data de prorrogação cadastrada na tabela de observaçoes
    Dim XLF_VALORRESIDUO As Double  'Armazena o valor do resíduo para calcular o valor corrigido do tit. de origem
    Dim XLL_CODCCORRENTE As Long    'Pegar Código da conta corrente no Qualifin
    Dim XLD_VENDA As Date           'Verificar Cessão
    Dim XLD_DTPAGTOANT As Date      'Data do pagto do título anterior
    
    PanEmpresas.Visible = False
    
    DatEmpresas.Recordset.bookmark = CboEmpresas.SelectedItem
    XCodEmpresa = CboEmpresas.BoundText
    
    'Gravação das Observações
    PanMensagem.Caption = "Importando Observações"
    subQOpenRecordset2 ResPadrao, "SELECT * FROM Banco", Estatico
    NumReg = 0
    
    While Not ResPadrao.EOF
        NumReg = NumReg + 1
        LblNReg.Caption = NumReg
        Sql = "INSERT INTO Observacoes(obse_tx_Observacao,obse_cd_CodDos, obse_dt_DataDos)" & _
                " VALUES ('" & ResPadrao!Nome & "','" & ResPadrao!Codigo & "', " & FunNuloData(ResPadrao!Data, NomeSgbd) & ")"
        Conexao.Execute Sql
        ResPadrao.MoveNext
    Wend
    ResPadrao.Close

    'Grava Corretores
    PanMensagem.Caption = "Importando Corretores"
    subQOpenRecordset2 ResPadrao, "SELECT * FROM CORR", Estatico
    NumReg = 0
            
    'Limpar o vínculo com a empresa anterior
    Conexao.Execute ("UPDATE Corretores SET corr_cd_CodDos = ''")
    
    While Not ResPadrao.EOF
        NumReg = NumReg + 1
        LblNReg.Caption = NumReg
        
        If XLB_IMPORTOU Then 'Já importou uma empresa
            SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Corretores WHERE corr_tx_Nome = '" & (ResPadrao!Nome_corr) & "' and corr_vl_PercentComissao= " & (ResPadrao!perc_corr) & " ", Dinamico
        Else
            SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Corretores WHERE corr_cd_Corretor= -1", Dinamico
        End If
        If XLO_QUALIFIN.EOF Then
            XLO_QUALIFIN.AddNew
        
            XLO_QUALIFIN!corr_tx_Nome = ResPadrao!Nome_corr
            XLO_QUALIFIN!corr_vl_PercentComissao = ResPadrao!perc_corr
            XLO_QUALIFIN!corr_cd_CodDos = ResPadrao!codi_corr
                        
            XLO_QUALIFIN.Update
        Else
            XLO_QUALIFIN!corr_cd_CodDos = ResPadrao!codi_corr
            XLO_QUALIFIN.Update
        End If
        XLO_QUALIFIN.Close
        ResPadrao.MoveNext
    Wend
    ResPadrao.Close

    'Clientes
    subQOpenRecordset2 ResPadrao, "SELECT * FROM Cliente", Estatico
    Cont = ResPadrao.RecordCount
    NumReg = 0
    
    PanMensagem.Caption = "Importando Clientes"
    While Not ResPadrao.EOF
        NumReg = NumReg + 1
        LblNReg.Caption = NumReg
    
        If Not IsNull(ResPadrao!PRCL_APTO) Then
            SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Profissoes WHERE prof_tx_Descricao = '" & ResPadrao!PRCL_APTO & "'", Dinamico
            If Not XLO_QUALIFIN.EOF Then
                XLL_PROFTITULAR = XLO_QUALIFIN!prof_cd_Profissao
            Else
                XLO_QUALIFIN.AddNew
                XLO_QUALIFIN!prof_tx_Descricao = ResPadrao!PRCL_APTO
                
                XLO_QUALIFIN.Update
                XLL_PROFTITULAR = XLO_QUALIFIN!prof_cd_Profissao
            End If
            XLO_QUALIFIN.Close
        End If
            
        If Not IsNull(ResPadrao!PRCO_APTO) Then
            SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Profissoes WHERE prof_tx_Descricao = '" & ResPadrao!PRCO_APTO & "'", Dinamico
            If Not XLO_QUALIFIN.EOF Then
                XLL_PROFCONJUGE = XLO_QUALIFIN!prof_cd_Profissao
            Else
                XLO_QUALIFIN.AddNew
                XLO_QUALIFIN!prof_tx_Descricao = ResPadrao!PRCO_APTO
                
                XLO_QUALIFIN.Update
                XLL_PROFCONJUGE = XLO_QUALIFIN!prof_cd_Profissao
            End If
            XLO_QUALIFIN.Close
        End If
            
        'Só precisa importar a tabela de clientes se não possui Qualifin, se sim a tabela já foi importada anteriormente
        If OptNao.Value = True Then
            'Parte Tabela FornClientes
            If ResPadrao!Tipo = "2" Or Len(ResPadrao!CPCL_APTO) = 11 Then
                'Len(ResPadrao!CPCL_APTO) = 11 solicitado por Newton por existirem muitas p. fisicas cadastrado como juridica
                XLT_TIPOPESSOA = "F"
                If FunConfereCPF(ResPadrao!CPCL_APTO) Then 'Válido
                    XLT_CGCCPF = Mid(ResPadrao!CPCL_APTO, 1, 3) + "." + Mid(ResPadrao!CPCL_APTO, 4, 3) + "." + Mid(ResPadrao!CPCL_APTO, 7, 3) + "-" + Right(ResPadrao!CPCL_APTO, 2)
                Else
                    XLT_CGCCPF = ""
                End If
            ElseIf ResPadrao!Tipo = "1" Or ResPadrao!Tipo = "" Then
                XLT_TIPOPESSOA = "J"
                If FunConfereCGC(ResPadrao!CPCL_APTO) Then 'Válido
                    XLT_CGCCPF = Mid(ResPadrao!CPCL_APTO, 1, 2) + "." + Mid(ResPadrao!CPCL_APTO, 3, 3) + "." + Mid(ResPadrao!CPCL_APTO, 6, 3) + "/" + Mid(ResPadrao!CPCL_APTO, 9, 4) + "-" + Right(ResPadrao!CPCL_APTO, 2)
                Else
                    XLT_CGCCPF = ""
                End If
            End If
                
            'Verificar se o registro já existe no QualiAdmFin
            'XLT_SQL = "SELECT (SELECT focl_cd_forncli FROM fornclientes WHERE focl_tx_cgccpf = '" & XLT_CGCCPF & "' and focl_tx_classe = 'C') as CGC," & _
                        "(SELECT focl_cd_forncli FROM fornclientes WHERE focl_tx_RazaoSocial = '" & FunTrataString(ResPadrao!CLIE_APTO) & "' and focl_tx_classe = 'C') as RAZAOSOCIAL," & _
                        "(SELECT focl_cd_forncli FROM fornclientes WHERE focl_tx_Fantasia = '" & FunTrataString(Mid(ResPadrao!CLIE_APTO, 1, 20)) & "' and focl_tx_classe = 'C') as FANTASIA"
            XLT_SQL = "SELECT (SELECT focl_cd_forncli FROM fornclientes WHERE focl_tx_cgccpf = '" & XLT_CGCCPF & "' and focl_tx_classe = 'C') as CGC"
            SubQOpenRecordset XLO_QUALIFIN, XLT_SQL, Estatico
            If IsNull(XLO_QUALIFIN!CGC) Then 'And IsNull(XLO_QUALIFIN!RAZAOSOCIAL) And IsNull(XLO_QUALIFIN!FANTASIA) Then
                XLO_QUALIFIN.Close
                SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM FornClientes WHERE focl_cd_forncli = -1", Dinamico
                XLO_QUALIFIN.AddNew
        
                XLO_QUALIFIN!focl_tx_classe = "C"
                XLO_QUALIFIN!focl_tx_fantasia = Mid(ResPadrao!CLIE_APTO, 1, 20)
                XLO_QUALIFIN!focl_tx_RazaoSocial = ResPadrao!CLIE_APTO
                XLO_QUALIFIN!focl_tx_endereco = FunNulo(ResPadrao!ENCO_APTO)
                XLO_QUALIFIN!focl_tx_cidade = FunNulo(ResPadrao!MUCO_APTO)
                XLO_QUALIFIN!focl_tx_Bairro = FunNulo(ResPadrao!BACO_APTO)
                XLO_QUALIFIN!focl_tx_estado = FunNulo(ResPadrao!ESCO_APTO)
                If Not IsNull(ResPadrao!CECO_APTO) Then
                    XLO_QUALIFIN!focl_tx_cep = Mid(ResPadrao!CECO_APTO, 1, 2) + "." + Mid(ResPadrao!CECO_APTO, 3, 3) + "-" + Mid(ResPadrao!CECO_APTO, 6, 3)
                Else
                    XLO_QUALIFIN!focl_tx_cep = FunNulo(ResPadrao!CECO_APTO)
                End If
                XLO_QUALIFIN!focl_tx_fone = FunNulo(ResPadrao!FOCO_APTO)
                XLO_QUALIFIN!focl_tx_fax = FunNulo(ResPadrao!FOCO_APTO)
                XLO_QUALIFIN!focl_tx_fone1 = FunNulo(ResPadrao!FORE_APTO)
                XLO_QUALIFIN!focl_tx_Tipo = XLT_TIPOPESSOA
                XLO_QUALIFIN!focl_tx_CGCCPF = XLT_CGCCPF
                XLO_QUALIFIN!CodVelho2 = ResPadrao!CODI_CLIE
                
                XLO_QUALIFIN.Update
                XLL_CODFORN = XLO_QUALIFIN!focl_cd_FornCli
            Else
            'ElseIf Not IsNull(XLO_QUALIFIN!CGC) Then
                XLL_CODFORN = XLO_QUALIFIN!CGC
            'ElseIf Not IsNull(XLO_QUALIFIN!RAZAOSOCIAL) Then
            '    XLL_CODFORN = XLO_QUALIFIN!RAZAOSOCIAL
            'ElseIf Not IsNull(XLO_QUALIFIN!FANTASIA) Then
            '    XLL_CODFORN = XLO_QUALIFIN!FANTASIA
            End If
            XLO_QUALIFIN.Close
            
            'Grava Tabela PlcCustosFornCli
            XLT_SQL = "SELECT * FROM PlccustoForncli WHERE focl_cd_forncli=" & XLL_CODFORN & " AND empr_cd_empresa =" & XCodEmpresa & ""
            SubQOpenRecordset XLO_QUALIFIN, XLT_SQL, Estatico
            If XLO_QUALIFIN.EOF Then
                XLO_QUALIFIN.AddNew
                XLO_QUALIFIN!focl_cd_FornCli = XLL_CODFORN
                XLO_QUALIFIN!empr_cd_Empresa = XCodEmpresa
                XLO_QUALIFIN!plco_cd_ContaRecDesp = funNull(ContaRecDesp)
                XLO_QUALIFIN!plco_cd_ContaCusto = funNull(ContaCusto)
                XLO_QUALIFIN!cere_cd_Pcr = funNull(Pcr)
            End If
            XLO_QUALIFIN.Update
            XLO_QUALIFIN.Close
        Else
            SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Fornclientes WHERE CodVelho2 = '" & ResPadrao!CODI_CLIE & "'", Estatico
            If XLO_QUALIFIN.EOF Then
                XLO_QUALIFIN.Close
                If ResPadrao!Tipo = "2" Or Len(ResPadrao!CPCL_APTO) = 11 Then
                    'Len(ResPadrao!CPCL_APTO) = 11 solicitado por Newton por existirem muitas p. fisicas cadastrado como juridica
                    XLT_CGCCPF = Mid(ResPadrao!CPCL_APTO, 1, 3) + "." + Mid(ResPadrao!CPCL_APTO, 4, 3) + "." + Mid(ResPadrao!CPCL_APTO, 7, 3) + "-" + Right(ResPadrao!CPCL_APTO, 2)
                ElseIf ResPadrao!Tipo = "1" Or ResPadrao!Tipo = "" Then
                    XLT_CGCCPF = Mid(ResPadrao!CPCL_APTO, 1, 2) + "." + Mid(ResPadrao!CPCL_APTO, 3, 3) + "." + Mid(ResPadrao!CPCL_APTO, 6, 3) + "/" + Mid(ResPadrao!CPCL_APTO, 9, 4) + "-" + Right(ResPadrao!CPCL_APTO, 2)
                End If
                
                XLT_SQL = "SELECT focl_cd_forncli,focl_tx_Tipo FROM fornclientes WHERE focl_tx_cgccpf = '" & XLT_CGCCPF & "' and focl_tx_classe = 'C'"
                SubQOpenRecordset XLO_QUALIFIN, XLT_SQL, Estatico
            End If
            XLL_CODFORN = XLO_QUALIFIN!focl_cd_FornCli
            XLT_TIPOPESSOA = XLO_QUALIFIN!focl_tx_Tipo
            XLO_QUALIFIN.Close
        End If
                
        'Parte da tabela Cliente
        SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Clientes WHERE focl_cd_forncli = " & XLL_CODFORN & "", Dinamico
        If XLO_QUALIFIN.EOF Then
            XLO_QUALIFIN.AddNew
            XLO_QUALIFIN!focl_cd_FornCli = XLL_CODFORN
            XLO_QUALIFIN!prof_cd_ProfTitular = XLL_PROFTITULAR
            XLO_QUALIFIN!clie_dt_Nascimento = ResPadrao!DNCL_APTO
            XLO_QUALIFIN!clie_tx_Nacionalidade = FunNulo(ResPadrao!NACL_APTO)
            XLO_QUALIFIN!clie_nr_Identidade = FunNulo(ResPadrao!IDCL_APTO)
            XLO_QUALIFIN!clie_tx_OrgEmissor = FunNulo(ResPadrao!OECL_APTO)
            XLO_QUALIFIN!clie_nr_EstCivil = ResPadrao!ECCL_APTO
            XLO_QUALIFIN!clie_tx_EndResidencial = FunNulo(ResPadrao!ENRE_APTO)
            XLO_QUALIFIN!clie_tx_BairroResidencial = FunNulo(ResPadrao!BARE_APTO)
            If Not IsNull(ResPadrao!CERE_APTO) Then
                XLO_QUALIFIN!clie_nr_CepResidencial = Mid(ResPadrao!CERE_APTO, 1, 2) + "." + Mid(ResPadrao!CERE_APTO, 3, 3) + "-" + Mid(ResPadrao!CERE_APTO, 6, 3)
            Else
                XLO_QUALIFIN!clie_nr_CepResidencial = FunNulo(ResPadrao!CERE_APTO)
            End If
            XLO_QUALIFIN!clie_tx_MunResidencial = FunNulo(ResPadrao!MURE_APTO)
            XLO_QUALIFIN!clie_tx_EstResidencial = FunNulo(ResPadrao!ESRE_APTO)
            XLO_QUALIFIN!clie_tx_Conjuge = FunNulo(ResPadrao!CONJ_APTO)
            XLO_QUALIFIN!prof_cd_ProfConjuge = XLL_PROFCONJUGE
            XLO_QUALIFIN!clie_dt_NascConjuge = ResPadrao!DNCO_APTO
            XLO_QUALIFIN!clie_tx_NacConjuge = FunNulo(ResPadrao!NACO_APTO)
            XLO_QUALIFIN!clie_nr_IdConjuge = FunNulo(ResPadrao!IDCO_APTO)
            XLO_QUALIFIN!clie_tx_OrgEmissorConjuge = FunNulo(ResPadrao!OECO_APTO)
            If ResPadrao!CPCO_APTO <> 0 Then
                XLO_QUALIFIN!clie_nr_CPFConjuge = Mid(ResPadrao!CPCO_APTO, 1, 3) + "." + Mid(ResPadrao!CPCO_APTO, 4, 3) + "." + Mid(ResPadrao!CPCO_APTO, 7, 3) + "-" + Right(ResPadrao!CPCO_APTO, 2)
            Else
                XLO_QUALIFIN!clie_nr_CPFConjuge = "   .   .   -  "
            End If
            XLO_QUALIFIN!clie_vl_RendaFamiliar = ResPadrao!REFA_APTO
                                        
            'Testar
            'Garantir que os campos sejam gravados como vazio
            If XLT_TIPOPESSOA = "F" Then 'Gravar, autorizado por Newton, o endereço residencial
                XLO_QUALIFIN!clie_tx_endcorresp = FunNulo(ResPadrao!ENRE_APTO)
                XLO_QUALIFIN!clie_tx_bairrocorresp = FunNulo(ResPadrao!BARE_APTO)
                If Not IsNull(ResPadrao!CERE_APTO) Then
                    XLO_QUALIFIN!clie_nr_cepcorresp = Mid(ResPadrao!CERE_APTO, 1, 2) + "." + Mid(ResPadrao!CERE_APTO, 3, 3) + "-" + Mid(ResPadrao!CERE_APTO, 6, 3)
                Else
                    XLO_QUALIFIN!clie_nr_cepcorresp = FunNulo(ResPadrao!CERE_APTO)
                End If
                XLO_QUALIFIN!clie_tx_muncorresp = FunNulo(ResPadrao!MURE_APTO)
                XLO_QUALIFIN!clie_tx_estcorresp = FunNulo(ResPadrao!ESRE_APTO)
            Else 'Gravar, autorizado por Newton, o endereço comercial
                XLO_QUALIFIN!clie_tx_endcorresp = FunNulo(ResPadrao!ENCO_APTO)
                XLO_QUALIFIN!clie_tx_bairrocorresp = FunNulo(ResPadrao!BACO_APTO)
                If Not IsNull(ResPadrao!CECO_APTO) Then
                    XLO_QUALIFIN!clie_nr_cepcorresp = Mid(ResPadrao!CECO_APTO, 1, 2) + "." + Mid(ResPadrao!CECO_APTO, 3, 3) + "-" + Mid(ResPadrao!CECO_APTO, 6, 3)
                Else
                    XLO_QUALIFIN!clie_nr_cepcorresp = FunNulo(ResPadrao!CECO_APTO)
                End If
                XLO_QUALIFIN!clie_tx_muncorresp = FunNulo(ResPadrao!MUCO_APTO)
                XLO_QUALIFIN!clie_tx_estcorresp = FunNulo(ResPadrao!ESCO_APTO)
            End If
            
            XLO_QUALIFIN.Update
        End If
        XLO_QUALIFIN.Close
        
        ResPadrao.MoveNext
    Wend
    ResPadrao.Close

    '**********MOVIMENTAÇÃO************
    PanMensagem.Caption = "Importando Movimentação"

    'Grava Empreendimentos
    PanMensagem.Caption = "Importando Empreendimentos"
    subQOpenRecordset2 ResPadrao, "SELECT * FROM EMPR", Estatico
    NumReg = 0
                
    While Not ResPadrao.EOF
        NumReg = NumReg + 1
        LblNReg.Caption = NumReg
        
        SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Empreendimentos WHERE empd_cd_Empreendimento= -1", Dinamico
        XLO_QUALIFIN.AddNew
        
        XLO_QUALIFIN!empd_cd_Empreendimento = Format(ResPadrao!codi_empr, "0000")
        XLO_QUALIFIN!empr_cd_Empresa = XCodEmpresa
        XLO_QUALIFIN!empd_tx_Nome = FunNulo(ResPadrao!nome_empr)
        XLO_QUALIFIN!empd_nr_Unidades = FunNuloVal(ResPadrao!unid_empr)
        XLO_QUALIFIN!empd_nr_UnVendidas = FunNuloVal(ResPadrao!vend_empr)
        XLO_QUALIFIN!empd_tx_Endereco = FunNulo(ResPadrao!end_empr)
        XLO_QUALIFIN!empd_tx_Bairro = FunNulo(ResPadrao!bair_empr)

        XLO_QUALIFIN.Update
    
        XLO_QUALIFIN.Close
        ResPadrao.MoveNext
    Wend
    ResPadrao.Close
       
    'TipoImovel
    Dim XLI_TIPOA As Integer
    Dim XLI_TIPOS As Integer
    Dim XLI_TIPOL As Integer
    Dim XLI_TIPOG As Integer
    Dim XLI_TIPOC As Integer
    Dim XLI_TIPOT As Integer
    
    For x = 0 To 5
        SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM TiposImoveis WHERE tiim_cd_TipoImovel = -1", Dinamico
        XLO_QUALIFIN.AddNew
        Select Case x
            Case 0
                XLO_QUALIFIN!tiim_tx_Descricao = "Apartamento"
                XLO_QUALIFIN.Update
                XLI_TIPOA = XLO_QUALIFIN!tiim_cd_TipoImovel
            Case 1
                XLO_QUALIFIN!tiim_tx_Descricao = "Sala"
                XLO_QUALIFIN.Update
                XLI_TIPOS = XLO_QUALIFIN!tiim_cd_TipoImovel
            Case 2
                XLO_QUALIFIN!tiim_tx_Descricao = "Loja"
                XLO_QUALIFIN.Update
                XLI_TIPOL = XLO_QUALIFIN!tiim_cd_TipoImovel
            Case 3
                XLO_QUALIFIN!tiim_tx_Descricao = "Garagem"
                XLO_QUALIFIN.Update
                XLI_TIPOG = XLO_QUALIFIN!tiim_cd_TipoImovel
            Case 4
                XLO_QUALIFIN!tiim_tx_Descricao = "Casa"
                XLO_QUALIFIN.Update
                XLI_TIPOC = XLO_QUALIFIN!tiim_cd_TipoImovel
            Case 5
                XLO_QUALIFIN!tiim_tx_Descricao = "Terreno/Lote"
                XLO_QUALIFIN.Update
                XLI_TIPOT = XLO_QUALIFIN!tiim_cd_TipoImovel
        End Select
        
        XLO_QUALIFIN.Close
    Next
    
    subQOpenRecordset2 ResPadrao, "SELECT APTO.*, EMPR.*, CLIENTE.TIPO, CLIENTE.CPCL_APTO FROM (APTO LEFT JOIN EMPR ON Mid(APTO.codi_apto, 1, 3) = EMPR.codi_empr) LEFT JOIN CLIENTE ON APTO.CODI_CLIE = CLIENTE.CODI_CLIE ", Estatico
    XLI_X = 1
    Dim XLT_CAMPOCOMISSAO As String
    Dim XLT_CAMPODTCOMISSAO As String
    NumReg = 0
    
    PanMensagem.Caption = "Importando Imóveis e Contratos"
    While Not ResPadrao.EOF
        NumReg = NumReg + 1
        LblNReg.Caption = NumReg
        
        'IMOVEIS
        XLT_SQL = "SELECT * FROM Imoveis WHERE imov_cd_Imovel = '" & Mid(ResPadrao!codi_apto, 4, 4) & "'" & _
                    "AND empr_cd_Empresa = " & XCodEmpresa & " AND empd_cd_Empreendimento = '" & Format(Mid(ResPadrao!codi_apto, 1, 3), "0000") & "'"
        SubQOpenRecordset XLO_QUALIFIN, XLT_SQL, Dinamico
        If XLO_QUALIFIN.EOF Then 'Ainda não foi cadastrado
            XLO_QUALIFIN.AddNew
                
            XLO_QUALIFIN!imov_cd_Imovel = Mid(ResPadrao!codi_apto, 4, 4)
            XLO_QUALIFIN!empr_cd_Empresa = XCodEmpresa
            XLO_QUALIFIN!empd_cd_Empreendimento = Format(Mid(ResPadrao!codi_apto, 1, 3), "0000")
            Select Case ResPadrao!TIPO_APTO
                Case "A"
                    XLO_QUALIFIN!tiim_cd_TipoImovel = XLI_TIPOA
                Case "S"
                    XLO_QUALIFIN!tiim_cd_TipoImovel = XLI_TIPOS
                Case "L"
                    XLO_QUALIFIN!tiim_cd_TipoImovel = XLI_TIPOL
                Case "G"
                    XLO_QUALIFIN!tiim_cd_TipoImovel = XLI_TIPOG
                Case "C"
                    XLO_QUALIFIN!tiim_cd_TipoImovel = XLI_TIPOC
                Case "T"
                    XLO_QUALIFIN!tiim_cd_TipoImovel = XLI_TIPOT
                Case Else
                    XLO_QUALIFIN!tiim_cd_TipoImovel = 7 ' Quando não existir gravar apto
            End Select
            XLO_QUALIFIN!imov_nr_Area = ResPadrao!area_apto
            XLO_QUALIFIN!imov_tx_Garagem = FunNulo(ResPadrao!gara_apto)
            Select Case ResPadrao!stat_apto
                Case "D"
                    XLO_QUALIFIN!imov_tx_Status = "D"
                Case "A"
                    XLO_QUALIFIN!imov_tx_Status = "V"
                Case "I"
                    'Verificar se existem apto ativos.
                    XLT_SQL = "SELECT * FROM APTO WHERE mid(codi_apto,1,7) LIKE '" & Mid(ResPadrao!codi_apto, 1, 7) & "' AND stat_apto = 'A'"
                    subQOpenRecordset2 XLO_DOS, XLT_SQL, Estatico
                    If Not XLO_DOS.EOF Then 'Existem ativos
                        XLO_QUALIFIN!imov_tx_Status = "V"
                    Else
                        XLO_QUALIFIN!imov_tx_Status = "D"
                    End If
                    XLO_DOS.Close
            End Select
            XLO_QUALIFIN!imov_vl_Avista = ResPadrao!avis_apto
            XLO_QUALIFIN!imov_nr_Andar = ResPadrao!andar_apto
            
            XLO_QUALIFIN.Update
        End If
        XLO_QUALIFIN.Close
        
        'PEGAR TIPO DE DESAGIO DA TABELA CONFIG
        subQOpenRecordset2 XLO_DOS, "SELECT * FROM CONFIG", Estatico
        XLT_TIPODESAGIO = XLO_DOS!tipo_Juros
        XLO_DOS.Close
                
        'COMPRADORES
        'SubQOpenRecordset XLO_QUALIFIN, "SELECT focl_cd_Forncli FROM FornClientes WHERE focl_cd_CodDos = '" & ResPadrao!CODI_CLIE & "'", Estatico
        SubQOpenRecordset XLO_QUALIFIN, "SELECT focl_cd_Forncli FROM FornClientes WHERE CodVelho2 = '" & ResPadrao!CODI_CLIE & "'", Estatico
        If XLO_QUALIFIN.EOF Then
            XLO_QUALIFIN.Close
            If ResPadrao!Tipo = "2" Or Len(ResPadrao!CPCL_APTO) = 11 Then
                'Len(ResPadrao!CPCL_APTO) = 11 solicitado por Newton por existirem muitas p. fisicas cadastrado como juridica
                XLT_CGCCPF = Mid(ResPadrao!CPCL_APTO, 1, 3) + "." + Mid(ResPadrao!CPCL_APTO, 4, 3) + "." + Mid(ResPadrao!CPCL_APTO, 7, 3) + "-" + Right(ResPadrao!CPCL_APTO, 2)
            ElseIf ResPadrao!Tipo = "1" Or ResPadrao!Tipo = "" Then
                XLT_CGCCPF = Mid(ResPadrao!CPCL_APTO, 1, 2) + "." + Mid(ResPadrao!CPCL_APTO, 3, 3) + "." + Mid(ResPadrao!CPCL_APTO, 6, 3) + "/" + Mid(ResPadrao!CPCL_APTO, 9, 4) + "-" + Right(ResPadrao!CPCL_APTO, 2)
            End If
                
            XLT_SQL = "SELECT focl_cd_forncli FROM fornclientes WHERE focl_tx_cgccpf = '" & XLT_CGCCPF & "' and focl_tx_classe = 'C'"
            SubQOpenRecordset XLO_QUALIFIN, XLT_SQL, Estatico
        End If
        XLL_CODCLIENTE = XLO_QUALIFIN!focl_cd_FornCli
        XLO_QUALIFIN.Close
        
        'OBTER PCR
        SubQOpenRecordset XLO_QUALIFIN, "SELECT cere_cd_Pcr FROM CentroResultados WHERE cere_cd_Estruturado='" & Format(ResPadrao!codi_empr, "000000") & "' and empr_cd_empresa= " & XCodEmpresa & "", Estatico
        If XLO_QUALIFIN.EOF Then
            SubQOpenRecordset XLO_QUALIFIN, "SELECT cere_cd_Pcr FROM CentroResultados WHERE cere_cd_Estruturado='999999' and empr_cd_empresa= " & XCodEmpresa & "", Estatico
        End If
        If Not XLO_QUALIFIN.EOF Then
            Pcr = XLO_QUALIFIN!cere_cd_Pcr
        End If
        XLO_QUALIFIN.Close
        
        'OBTER CONTA CONTÁBIL
        'comentado na importacao da Ebisa. Voltar ao normal depois
        SubQOpenRecordset XLO_QUALIFIN, "SELECT Plco_cd_Conta FROM PlanoContas WHERE plco_cd_CodReduzido='" & ResPadrao!Conta & "' and empr_cd_empresa= " & XCodEmpresa & "", Estatico
        If XLO_QUALIFIN.EOF Then
            SubQOpenRecordset XLO_QUALIFIN, "SELECT Plco_cd_Conta FROM PlanoContas WHERE plco_cd_CodReduzido='999996' and empr_cd_empresa= " & XCodEmpresa & "", Estatico
        End If
        If Not XLO_QUALIFIN.EOF Then
            ContaRecDesp = XLO_QUALIFIN!plco_cd_Conta
        End If
        XLO_QUALIFIN.Close
        
        'CONTRATOS
        If ResPadrao!stat_apto = "A" Or ResPadrao!stat_apto = "I" Then 'Imóvel ativo ou Inativo, logo existe contrato
            'Obter percentuais de juros,deságio,multa
            subQOpenRecordset2 XLO_DOS, "SELECT juros_fin, per_multa, per_juros FROM FINA " & _
                            " WHERE Mid(FINA.CODI_FINA, 1, 8)='" & ResPadrao!codi_apto & "' ORDER BY codi_fina", Estatico
            
            If Not XLO_DOS.EOF Then
                XLF_PERCDESAGIO = XLO_DOS!juros_fin
                XLF_PERCMULTA = XLO_DOS!per_multa
                XLF_PERCJUROS = XLO_DOS!per_juros
            End If
            XLO_DOS.Close
            
            SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Contratos WHERE cont_cd_Contrato = '-1'", Dinamico
            XLO_QUALIFIN.AddNew
            
            'XLO_QUALIFIN!cont_cd_Contrato = Format(Str(Int(Mid(ResPadrao!codi_apto, 8, 1)) + 1), "00") - Newton pediu que importasse = ao DOS
            XLO_QUALIFIN!cont_cd_Contrato = Format(Mid(ResPadrao!codi_apto, 8, 1), "00")
            XLO_QUALIFIN!imov_cd_Imovel = Mid(ResPadrao!codi_apto, 4, 4)
            XLO_QUALIFIN!empr_cd_Empresa = XCodEmpresa
            XLO_QUALIFIN!empd_cd_Empreendimento = Format(Mid(ResPadrao!codi_apto, 1, 3), "0000")
            If XLL_CODCLIENTE <> 0 Then
                XLO_QUALIFIN!focl_cd_FornCli = XLL_CODCLIENTE
            End If
            XLO_QUALIFIN!cont_vl_Sinal = ResPadrao!sina_apto
            XLO_QUALIFIN!cont_vl_Poupanca = ResPadrao!poup_apto
            XLO_QUALIFIN!cont_vl_Fgts = ResPadrao!fgts_apto
            XLO_QUALIFIN!cont_vl_Financiamento = ResPadrao!fina_apto
            XLO_QUALIFIN!cont_vl_Outros = 0
            XLO_QUALIFIN!cont_vl_Total = ResPadrao!tovd_apto
            If ResPadrao!sina_apto = 0 And ResPadrao!poup_apto = 0 And ResPadrao!fgts_apto = 0 And ResPadrao!fina_apto = 0 Then
                XLO_QUALIFIN!cont_vl_Avista = ResPadrao!avis_apto
            Else
                XLO_QUALIFIN!cont_vl_Avista = 0
                If ResPadrao!tovd_apto = 0 Then
                    XLO_QUALIFIN!cont_vl_Total = ResPadrao!sina_apto + ResPadrao!poup_apto + ResPadrao!fgts_apto + ResPadrao!fina_apto
                End If
            End If
            XLO_QUALIFIN!cont_dt_Venda = ResPadrao!dtvd_apto
            XLO_QUALIFIN!cont_dt_Base = ResPadrao!dtvd_apto
            XLO_QUALIFIN!cont_dt_RegistroContrato = ResPadrao!dtre_apto
            XLO_QUALIFIN!cont_dt_EntregaChaves = ResPadrao!dtch_apto
            XLO_QUALIFIN!cont_vl_TaxaSegVida = ResPadrao!seg_vida
            XLO_QUALIFIN!cont_dt_Distrato = ResPadrao!dt_inativo '- comentado importacao ebisa
            XLO_QUALIFIN!cont_vl_PercJurosFin = XLF_PERCDESAGIO
            XLO_QUALIFIN!cont_vl_PercJurosMulta = XLF_PERCMULTA
            XLO_QUALIFIN!cont_vl_PercJurosMora = XLF_PERCJUROS
            
            If ResPadrao!mult_pro = "S" Then
                XLO_QUALIFIN!cont_tx_MultaProRata = 1
            Else
                XLO_QUALIFIN!cont_tx_MultaProRata = 0
            End If
            If ResPadrao!correc_pro = "S" Then
                XLO_QUALIFIN!cont_tx_CorrecProRata = 1
            Else
                XLO_QUALIFIN!cont_tx_CorrecProRata = 0
            End If
            XLO_QUALIFIN!cont_nr_TipoSeguro = "" ' ResPadrao!tipo_seg - Comentado na importacao da incorporadora
            If ResPadrao!stat_apto = "A" Then
                XLO_QUALIFIN!cont_tx_Status = "A"
            Else 'INATIVO
                If Not IsNull(ResPadrao!dtvd_apto) Then
                    XLD_VENDA = ResPadrao!dtvd_apto
                    XLT_TITULO = Mid(ResPadrao!codi_apto, 1, 7)
                    ResPadrao.MoveNext
                    If Not ResPadrao.EOF Then
                        If XLT_TITULO = Mid(ResPadrao!codi_apto, 1, 7) And XLD_VENDA = ResPadrao!dtvd_apto Then
                            XLO_QUALIFIN!cont_tx_Status = "C"
                        Else
                            XLO_QUALIFIN!cont_tx_Status = "I"
                        End If
                    End If
                    ResPadrao.MovePrevious
                Else
                    XLO_QUALIFIN!cont_tx_Status = "I"
                End If
            End If
            
            Select Case CboEmpresas.Text
            Case "EMP. COSTA ANDRADE", "FERNANDO FILGUEIRAS", "DUQUE DE EDINBURGH", "DUQUE DE BERGARA", _
                "ORLANDO GOMES", "DUQUE DE ORLEANS", "DUQUE DE PROVENCE"
                XLO_QUALIFIN!cont_tx_TipoJurosMora = "C"
            Case "C.A. EMPREENDIMENTOS"
                If Mid(ResPadrao!codi_apto, 1, 3) = "539" Then
                    XLO_QUALIFIN!cont_tx_TipoJurosMora = "C"
                Else
                    XLO_QUALIFIN!cont_tx_TipoJurosMora = "D"
                End If
            Case Else
                XLO_QUALIFIN!cont_tx_TipoJurosMora = "D" 'Cheio para Condominio
            End Select
            XLO_QUALIFIN!cont_tx_Desagio = XLT_TIPODESAGIO
            
            XLO_QUALIFIN!cere_cd_Pcr = funNull(Pcr)
            XLO_QUALIFIN!plco_cd_Conta = funNull(ContaRecDesp)
            
            XLO_QUALIFIN.Update
        
            XLO_QUALIFIN.Close
            
             'COMPRADORES
            If XLL_CODCLIENTE <> 0 Then
                'Conexao.Execute "INSERT INTO Compradores(focl_cd_FornCli,cont_cd_Contrato, " & _
                                      "imov_cd_Imovel,empr_cd_Empresa,empd_cd_Empreendimento) " & _
                                  "VALUES (" & XLL_CODCLIENTE & ",'" & Format(Str(Int(Mid(ResPadrao!codi_apto, 8, 1)) + 1), "00") & "', " & _
                                      "'" & Mid(ResPadrao!codi_apto, 4, 4) & "'," & XCodEmpresa & ",'" & Format(Mid(ResPadrao!codi_apto, 1, 3), "0000") & "')"
                Conexao.Execute "INSERT INTO Compradores(focl_cd_FornCli,cont_cd_Contrato, " & _
                                      "imov_cd_Imovel,empr_cd_Empresa,empd_cd_Empreendimento) " & _
                                  "VALUES (" & XLL_CODCLIENTE & ",'" & Format(Mid(ResPadrao!codi_apto, 8, 1), "00") & "', " & _
                                      "'" & Mid(ResPadrao!codi_apto, 4, 4) & "'," & XCodEmpresa & ",'" & Format(Mid(ResPadrao!codi_apto, 1, 3), "0000") & "')"
            End If
        End If
                
        'COMISSÕES
        XLI_X = 1
        SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Corretores WHERE corr_cd_CodDos<>'' and (corr_cd_CodDos = '" & ResPadrao!corr1_apto & "' or corr_cd_CodDos = '" & ResPadrao!corr2_apto & "' or corr_cd_CodDos = '" & ResPadrao!corr3_apto & "')", Estatico
        While Not XLO_QUALIFIN.EOF
              If Not IsNull(ResPadrao.Fields("CORR" & XLI_X & "_apto")) Then
                XLT_CAMPOCOMISSAO = ResPadrao.Fields("coc" & XLI_X & "_apto")
                If Not IsNull(ResPadrao.Fields("dpc" & XLI_X & "_apto")) Then
                    XLT_CAMPODTCOMISSAO = ResPadrao.Fields("dpc" & XLI_X & "_apto")
                End If
                
                'Conexao.Execute "INSERT INTO Comissoes(corr_cd_corretor,cont_cd_Contrato, " & _
                                  "imov_cd_Imovel,empr_cd_Empresa,empd_cd_Empreendimento, " & _
                                  "comi_vl_Comissao,comi_dt_Pagamento) " & _
                              "VALUES (" & XLO_QUALIFIN!corr_cd_Corretor & ",'" & Format(Str(Int(Mid(ResPadrao!codi_apto, 8, 1)) + 1), "00") & "', " & _
                                  "'" & Mid(ResPadrao!codi_apto, 4, 4) & "'," & XCodEmpresa & ",'" & Format(Mid(ResPadrao!codi_apto, 1, 3), "0000") & "', " & _
                                  "" & (FunTrataFloat(XLT_CAMPOCOMISSAO)) & "," & (FunNuloData(XLT_CAMPODTCOMISSAO, NomeSgbd)) & ")"
                Conexao.Execute "INSERT INTO Comissoes(corr_cd_corretor,cont_cd_Contrato, " & _
                                  "imov_cd_Imovel,empr_cd_Empresa,empd_cd_Empreendimento, " & _
                                  "comi_vl_Comissao,comi_dt_Pagamento) " & _
                              "VALUES (" & XLO_QUALIFIN!corr_cd_Corretor & ",'" & Format(Mid(ResPadrao!codi_apto, 8, 1), "00") & "', " & _
                                  "'" & Mid(ResPadrao!codi_apto, 4, 4) & "'," & XCodEmpresa & ",'" & Format(Mid(ResPadrao!codi_apto, 1, 3), "0000") & "', " & _
                                  "" & (FunTrataFloat(XLT_CAMPOCOMISSAO)) & "," & (FunNuloData(XLT_CAMPODTCOMISSAO, NomeSgbd)) & ")"
              End If
              XLI_X = XLI_X + 1
              XLO_QUALIFIN.MoveNext
        Wend
             
        ResPadrao.MoveNext
    Wend
    ResPadrao.Close
    XLO_QUALIFIN.Close ' Inserido em 13/10/08 - Patrícia
        
    'TÍTULOS
    PanMensagem.Caption = "Importando Títulos"
    subQOpenRecordset2 ResPadrao, "SELECT * FROM FINA ORDER BY codi_fina", Estatico
    NumReg = 0
    While Not ResPadrao.EOF
        NumReg = NumReg + 1
        LblNReg.Caption = NumReg
        
        'TIPOS PLANOS
        If ResPadrao!FINANCIA = "X" Or ResPadrao!FINANCIA = "N" Or IsNull(ResPadrao!FINANCIA) Then 'Financiamento ou Sinal gravado no DOS (passou a ser mensal no Qualicapi)
            If Right(ResPadrao!Codi_fina, 1) > 0 Then 'Resíduo
                SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM TiposPlanos WHERE tipl_tx_Sigla = 'R' ", Dinamico
            Else
                SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM TiposPlanos WHERE tipl_tx_Sigla = 'M' ", Dinamico
            End If
        Else
            SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM TiposPlanos WHERE tipl_tx_Sigla = '" & ResPadrao!FINANCIA & "' ", Dinamico
        End If
        If XLO_QUALIFIN.EOF Then 'Ainda não foi cadastrado
            XLO_QUALIFIN.AddNew
            
            Select Case ResPadrao!FINANCIA
                'NATUREZA (P-Poupança, F-Financiamento, O-Outros, S-Sinal, G-FGTS)
                'Unidade (M-Mesal, D-Diario)
                Case "M"
                    XLO_QUALIFIN!tipl_tx_Descricao = "MENSAL"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "P"
                Case "T"
                    XLO_QUALIFIN!tipl_tx_Descricao = "TRIMESTRAL"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 3
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "P"
                Case "S"
                    XLO_QUALIFIN!tipl_tx_Descricao = "SEMESTRAL"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 6
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "P"
                Case "A"
                    XLO_QUALIFIN!tipl_tx_Descricao = "ANUAL"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 12
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "P"
                Case "Q"
                    XLO_QUALIFIN!tipl_tx_Descricao = "QUINZENAL"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 15
                    XLO_QUALIFIN!tipl_tx_Unidade = D
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "P"
                Case "R"
                    XLO_QUALIFIN!tipl_tx_Descricao = "RESIDUO"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "P"
                Case "G"
                    XLO_QUALIFIN!tipl_tx_Descricao = "SEGURO"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "O"
                Case "J"
                    XLO_QUALIFIN!tipl_tx_Descricao = "JUROS"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "O"
                Case "I"
                    XLO_QUALIFIN!tipl_tx_Descricao = "Complemento SINAL"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "S"
                Case "C"
                    XLO_QUALIFIN!tipl_tx_Descricao = "Chaves"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "O"
                Case "Z"
                    XLO_QUALIFIN!tipl_tx_Descricao = "ANTECIPACAO"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "O"
                Case "D"
                    XLO_QUALIFIN!tipl_tx_Descricao = "ADITIVO"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "O"
                Case "O"
                    XLO_QUALIFIN!tipl_tx_Descricao = "MODIFICACOES"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "O"
                Case "E"
                    XLO_QUALIFIN!tipl_tx_Descricao = "TAXA EXTRA"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "O"
                Case "K"
                    XLO_QUALIFIN!tipl_tx_Descricao = "QUOTA DE CONSTRUCAO"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "O"
                Case "Y"
                    XLO_QUALIFIN!tipl_tx_Descricao = "QUOTA DE MODIFICACAO"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "O"
                Case "W"
                    XLO_QUALIFIN!tipl_tx_Descricao = "QUOTA COMPLEMENTAR"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "O"
                Case "V"
                    XLO_QUALIFIN!tipl_tx_Descricao = "FUNDO DE RESERVA"
                    XLO_QUALIFIN!tipl_nr_Periodicidade = 1
                    XLO_QUALIFIN!tipl_tx_Unidade = M
                    XLO_QUALIFIN!tipl_tx_Sigla = ResPadrao!FINANCIA
                    XLT_NATUREZA = "O"
            End Select
            XLO_QUALIFIN.Update
        Else
            If Not IsNull(ResPadrao!FINANCIA) Then
            Select Case ResPadrao!FINANCIA
                'NATUREZA (P-Poupança, F-Financiamento, O-Outros, S-Sinal, G-FGTS)
                'Unidade (M-Mesal, D-Diario)
                Case "M", "T", "S", "A", "Q", "R"
                    XLT_NATUREZA = "P"
                Case "X"
                    XLT_NATUREZA = "F"
                Case "G", "J", "C", "Z", "D", "O", "E", "K", "Y", "W", "V"
                    XLT_NATUREZA = "O"
                Case "N", "I"
                    XLT_NATUREZA = "S"
            End Select
            Else
                XLT_NATUREZA = "P"
            End If
        End If
        XLI_TIPOPLANO = XLO_QUALIFIN!tipl_cd_TipoPlano
        XLO_QUALIFIN.Close
        
        'Obter Moedas
        For x = 0 To Campos.UpperBound(1)
            If ResPadrao!umdb_fina = Campos(x, 0) Then
                XLI_MOEDA1 = Campos(x, 4)
            End If
            If ResPadrao!umd2_fina = Campos(x, 0) Then
                XLI_MOEDA2 = Campos(x, 4)
            End If
        Next
        If ResPadrao!umdb_fina = 0 Then  'REAL
            XLI_MOEDA1 = 0
        End If
        If ResPadrao!umd2_fina = 0 Then  'Real
            XLI_MOEDA2 = 0
        End If
        
        If Not IsNull(ResPadrao!loca_fina) Then
            'Obter Obs
            SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Observacoes WHERE obse_cd_CodDos = '" & ResPadrao!loca_fina & "' ", Estatico
            XLI_OBS = XLO_QUALIFIN!obse_cd_Observacao
            If IsDate(XLO_QUALIFIN!obse_dt_datados) Then
                XLD_PRORROGACAO = XLO_QUALIFIN!obse_dt_datados
            Else
                XLD_PRORROGACAO = ResPadrao!dtvc_fina
            End If
            XLO_QUALIFIN.Close
        Else
            XLD_PRORROGACAO = ResPadrao!dtvc_fina
            XLI_OBS = 0
        End If
    
        'Obter Conta Corrente
        If Not IsNull(ResPadrao!Banco) Then
            SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM ContasCorrente WHERE CodVelho = '" & ResPadrao!Banco & "' ", Estatico
        Else
            SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM ContasCorrente WHERE banc_cd_Codigo = 9999 and empr_cd_empresa = " & XCodEmpresa & " ", Estatico
        End If
        If Not XLO_QUALIFIN.EOF Then
            XLL_CODCCORRENTE = XLO_QUALIFIN!coco_cd_Codigo
        End If
        XLO_QUALIFIN.Close
    
        'Obter Dt Base, Cr e conta do Contrato
        'SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Contratos WHERE cont_cd_Contrato = '" & Format(Str(Int(Mid(ResPadrao!codi_fina, 8, 1)) + 1), "00") & "' AND " & _
                                                "imov_cd_Imovel = '" & Mid(ResPadrao!codi_fina, 4, 4) & "' AND " & _
                                                "empr_cd_Empresa = " & XCodEmpresa & " AND " & _
                                                "empd_cd_Empreendimento = '" & Format(Mid(ResPadrao!codi_fina, 1, 3), "0000") & "'", Estatico
        SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Contratos WHERE cont_cd_Contrato = '" & Format(Mid(ResPadrao!Codi_fina, 8, 1), "00") & "' AND " & _
                                                "imov_cd_Imovel = '" & Mid(ResPadrao!Codi_fina, 4, 4) & "' AND " & _
                                                "empr_cd_Empresa = " & XCodEmpresa & " AND " & _
                                                "empd_cd_Empreendimento = '" & Format(Mid(ResPadrao!Codi_fina, 1, 3), "0000") & "'", Estatico
        XLD_DTBASE = XLO_QUALIFIN!cont_dt_Base
        Pcr = FunNulo(XLO_QUALIFIN!cere_cd_Pcr)
        ContaRecDesp = FunNulo(XLO_QUALIFIN!plco_cd_Conta)

        XLO_QUALIFIN.Close
        
        'TÍTULOS
        SubQOpenRecordset XLO_QUALIFIN, "SELECT * FROM Titulos WHERE titu_cd_Plano = '-1' ", Dinamico
        XLO_QUALIFIN.AddNew
        
        If XLL_CODCCORRENTE = 0 Then
            XLO_QUALIFIN!coco_cd_Codigo = Null
        Else
            XLO_QUALIFIN!coco_cd_Codigo = XLL_CODCCORRENTE
        End If
        XLO_QUALIFIN!titu_cd_Plano = Mid(ResPadrao!Codi_fina, 9, 2)
        XLO_QUALIFIN!titu_cd_Parcela = Mid(ResPadrao!Codi_fina, 11, 3)
        XLO_QUALIFIN!titu_cd_Residuo = Format(Mid(ResPadrao!Codi_fina, 14, 1), "00")
        'XLO_QUALIFIN!cont_cd_Contrato = Format(Str(Int(Mid(ResPadrao!codi_fina, 8, 1)) + 1), "00")
        XLO_QUALIFIN!cont_cd_Contrato = Format(Mid(ResPadrao!Codi_fina, 8, 1), "00")
        XLO_QUALIFIN!imov_cd_Imovel = Mid(ResPadrao!Codi_fina, 4, 4)
        XLO_QUALIFIN!empr_cd_Empresa = XCodEmpresa
        XLO_QUALIFIN!empd_cd_Empreendimento = Format(Mid(ResPadrao!Codi_fina, 1, 3), "0000")
    
        XLO_QUALIFIN!titu_vl_Parcela = ResPadrao!vldc_fina
        XLO_QUALIFIN!tipl_cd_TipoPlano = XLI_TIPOPLANO
        
        If XLI_MOEDA1 = 0 Then
            XLO_QUALIFIN!moed_cd_Moeda1 = Null
            XLO_QUALIFIN!titu_vl_ParcelaIndex1 = 0
        Else
            XLO_QUALIFIN!moed_cd_Moeda1 = XLI_MOEDA1
            XLO_QUALIFIN!titu_vl_ParcelaIndex1 = ResPadrao!vldb_fina
        End If
        
        If XLI_MOEDA2 = 0 Then
            XLO_QUALIFIN!moed_cd_Moeda2 = Null
            XLO_QUALIFIN!titu_vl_ParcelaIndex2 = 0
        Else
            XLO_QUALIFIN!moed_cd_Moeda2 = XLI_MOEDA2
            XLO_QUALIFIN!titu_vl_ParcelaIndex2 = ResPadrao!vld2_fina
        End If
        
        XLO_QUALIFIN!titu_dt_Vencimento = ResPadrao!dtvc_fina
        If XLI_OBS = 0 Then
            XLO_QUALIFIN!obse_cd_Observacao = Null
        Else
            XLO_QUALIFIN!obse_cd_Observacao = XLI_OBS
        End If
        If Not IsNull(ResPadrao!vlpg_fina) Then
            XLO_QUALIFIN!titu_vl_ValorPago = ResPadrao!vlpg_fina
        Else
            XLO_QUALIFIN!titu_vl_ValorPago = 0
        End If
        If Not IsNull(ResPadrao!dtpg_fina) Then
            If IsDate(ResPadrao!dtpg_fina) Then
            '    XLO_QUALIFIN!titu_dt_Pagamento = ResPadrao!dtpg_fina
            'Else
                XLO_QUALIFIN!titu_dt_Pagamento = Mid(ResPadrao!dtpg_fina, 1, 2) & "/" & Mid(ResPadrao!dtpg_fina, 4, 2) & "/" & Right(ResPadrao!dtpg_fina, 2)
            End If
        Else
            XLO_QUALIFIN!titu_dt_Pagamento = ResPadrao!dtpg_fina
        End If
        XLO_QUALIFIN!titu_vl_Multa = FunNuloVal(ResPadrao!multa)
        XLO_QUALIFIN!titu_vl_Juros = FunNuloVal(ResPadrao!JUROS)
        XLO_QUALIFIN!titu_vl_Seguro = FunNuloVal(ResPadrao!SEGURO)
        XLO_QUALIFIN!titu_tx_IndicePagamento = ResPadrao!Indice
        If ResPadrao!Indice = "1" Then
            If XLI_MOEDA1 = 0 Then
              XLO_QUALIFIN!moed_cd_Moeda3 = Null
            Else
              XLO_QUALIFIN!moed_cd_Moeda3 = XLI_MOEDA1
            End If
        ElseIf ResPadrao!Indice = "2" Then
            If XLI_MOEDA2 = 0 Then
              XLO_QUALIFIN!moed_cd_Moeda3 = Null
            Else
              XLO_QUALIFIN!moed_cd_Moeda3 = XLI_MOEDA2
            End If
        Else
            XLO_QUALIFIN!moed_cd_Moeda3 = Null
        End If
        XLO_QUALIFIN!titu_vl_Outros = FunNuloVal(ResPadrao!honorario)
        XLO_QUALIFIN!titu_vl_Desconto = FunNuloVal(ResPadrao!desconto)
        XLO_QUALIFIN!titu_vl_Desagio = FunNuloVal(ResPadrao!desagio)
        XLO_QUALIFIN!titu_vl_PercJurosFin = FunNuloVal(ResPadrao!juros_fin)
        XLO_QUALIFIN!titu_vl_PercMultaMora = FunNuloVal(ResPadrao!per_multa)
        XLO_QUALIFIN!titu_vl_PercJurosMora = FunNuloVal(ResPadrao!per_juros)
        XLO_QUALIFIN!titu_dt_Deposito = ResPadrao!dt_deposi
        XLO_QUALIFIN!titu_dt_Repasse = ResPadrao!dt_repasse '- Comentado para importação da Actitur/Ebisa
        XLO_QUALIFIN!titu_tx_LocalPagto = ResPadrao!local_pag
        XLO_QUALIFIN!titu_tx_DocPagto = FunNulo(ResPadrao!Doc_Pag)
        'If IsDate(XLD_PRORROGACAO) Then
            XLO_QUALIFIN!titu_dt_Prorrogacao = XLD_PRORROGACAO
        'Else
        '    XLO_QUALIFIN!titu_dt_Prorrogacao = ResPadrao!dtvc_fina
        'End If
        XLO_QUALIFIN!titu_tx_NaturezaPlano = XLT_NATUREZA
        If Not IsNull(ResPadrao!dtpg_fina) Then
            If IsDate(ResPadrao!dtpg_fina) Then
            '    XLO_QUALIFIN!titu_dt_BasePagto = Mid(ResPadrao!dtpg_fina, 1, 2) & "/" & Mid(ResPadrao!dtpg_fina, 4, 2) & "/" & Right(ResPadrao!dtpg_fina, 2)
            'Else
                XLO_QUALIFIN!titu_dt_BasePagto = Mid(ResPadrao!dtpg_fina, 1, 2) & "/" & Mid(ResPadrao!dtpg_fina, 4, 2) & "/" & Right(ResPadrao!dtpg_fina, 2)
            End If
        'Else
          '  XLO_QUALIFIN!titu_dt_BasePagto = Mid(ResPadrao!dtpg_fina, 1, 2) & "/" & Mid(ResPadrao!dtpg_fina, 4, 2) & "/" & Right(ResPadrao!dtpg_fina, 2)
        End If
        If Mid(ResPadrao!Codi_fina, 14, 1) <> "0" Then 'Resíduo
            XLO_QUALIFIN!titu_dt_Base = XLD_DTPAGTOANT 'Pegando a data do pagto do registro anterior
        Else '1º registro
            XLO_QUALIFIN!titu_dt_Base = XLD_DTBASE
        End If
        If Not IsNull(ResPadrao!dtpg_fina) Then
            XLD_DTPAGTOANT = ResPadrao!dtpg_fina
        End If
        
        Select Case ResPadrao!Tipo_Pag
            Case "1"
                If ResPadrao!Indice = "1" Or ResPadrao!Indice = "2" Then
                    If ResPadrao!dtvc_fina = ResPadrao!dtpg_fina Then
                        'Pagto Total Corrigido no Vencimento
                        XLO_QUALIFIN!titu_tx_TipoPag = "1"
                    ElseIf ResPadrao!dtvc_fina < ResPadrao!dtpg_fina Then
                        'Pagto Total Corrigido Vencido
                        XLO_QUALIFIN!titu_tx_TipoPag = "3"
                    Else
                        'Pagto total antecipado
                        XLO_QUALIFIN!titu_tx_TipoPag = "5"
                    End If
                Else ' Em Real
                    If ResPadrao!dtvc_fina = ResPadrao!dtpg_fina Then
                        'Pagto Total em real no Vencimento
                        XLO_QUALIFIN!titu_tx_TipoPag = "7"
                    ElseIf ResPadrao!dtvc_fina < ResPadrao!dtpg_fina Then
                        'Pagto Total em real Vencido
                        XLO_QUALIFIN!titu_tx_TipoPag = "9"
                    Else
                        'Pagto total antecipado
                        XLO_QUALIFIN!titu_tx_TipoPag = "11"
                    End If
                End If
            Case "2"
                If ResPadrao!Indice = "1" Or ResPadrao!Indice = "2" Then
                    If ResPadrao!dtvc_fina = ResPadrao!dtpg_fina Then
                        'Pagto parcial em real no Vencimento
                        XLO_QUALIFIN!titu_tx_TipoPag = "2"
                    ElseIf ResPadrao!dtvc_fina < ResPadrao!dtpg_fina Then
                        'Pagto parcial em real Vencido
                        XLO_QUALIFIN!titu_tx_TipoPag = "4"
                    Else
                        'Pagto parcial antecipado
                        XLO_QUALIFIN!titu_tx_TipoPag = "6"
                    End If
                Else ' Em Real
                    If ResPadrao!dtvc_fina = ResPadrao!dtpg_fina Then
                        'Pagto parcial em real no Vencimento
                        XLO_QUALIFIN!titu_tx_TipoPag = "8"
                    ElseIf ResPadrao!dtvc_fina < ResPadrao!dtpg_fina Then
                        'Pagto parcial em real Vencido
                        XLO_QUALIFIN!titu_tx_TipoPag = "10"
                    Else
                        'Pagto parcial antecipado
                        XLO_QUALIFIN!titu_tx_TipoPag = "12"
                    End If
                End If
            Case "3"
                If ResPadrao!Indice = "1" Or ResPadrao!Indice = "2" Then
                    XLO_QUALIFIN!titu_tx_TipoPag = "5"
                Else
                    XLO_QUALIFIN!titu_tx_TipoPag = "11"
                End If
            Case "4"
                If ResPadrao!Indice = "1" Or ResPadrao!Indice = "2" Then
                    'Pagto parcial corrigido antecipado
                    XLO_QUALIFIN!titu_tx_TipoPag = "6"
                Else
                    XLO_QUALIFIN!titu_tx_TipoPag = "12"
                End If
        End Select
        'XLO_QUALIFIN!titu_tx_TipoPag = ResPadrao!Tipo_Pag '- Comentado para importacao da incorporadora
        
        XLO_QUALIFIN!cere_cd_Pcr = funNull(Pcr)
        XLO_QUALIFIN!plco_cd_Conta = funNull(ContaRecDesp)
    
        If Not IsNull(ResPadrao!dtpg_fina) Then 'Título pago
            XLT_TITULO = ResPadrao!Codi_fina
            ResPadrao.MoveNext
            
            'Verifica se tem resíduo
            XLF_VALORRESIDUO = 0
            If Not ResPadrao.EOF Then
                If Left(XLT_TITULO, 13) = Left(ResPadrao!Codi_fina, 13) Then
                    XLF_VALORRESIDUO = ResPadrao!vldc_fina
                End If
            End If
            ResPadrao.MovePrevious
            
            If XLI_MOEDA1 = 0 Then
                XLO_QUALIFIN!titu_vl_Corrigido1 = Format(ResPadrao!vldc_fina, "standard") 'Valor Histórico
            Else
                XLO_QUALIFIN!titu_vl_Corrigido1 = ResPadrao!vlpg_fina + XLF_VALORRESIDUO - ResPadrao!SEGURO - _
                                                ResPadrao!multa - ResPadrao!JUROS - ResPadrao!honorario + _
                                                ResPadrao!desconto + ResPadrao!desagio
            End If
            If XLI_MOEDA2 = 0 Then
                XLO_QUALIFIN!titu_vl_Corrigido2 = Format(ResPadrao!vldc_fina, "standard") 'Valor Histórico
            Else
                XLO_QUALIFIN!titu_vl_Corrigido2 = ResPadrao!vlpg_fina + XLF_VALORRESIDUO - ResPadrao!SEGURO - _
                                                ResPadrao!multa - ResPadrao!JUROS - ResPadrao!honorario + _
                                                ResPadrao!desconto + ResPadrao!desagio
            End If
                
        End If
        
        XLO_QUALIFIN.Update
        XLO_QUALIFIN.Close
             
        ResPadrao.MoveNext
    Wend
    ResPadrao.Close
    
    PanMensagem.Caption = "Importação Realizada com Sucesso!"
    
    'Recarrega combo de empresas excluindo a que já foi importada
    If XLT_EMPRESAS = "" Then
        XLT_EMPRESAS = "empr_cd_empresa <> " & DatEmpresas.Recordset.Fields!empr_cd_Empresa
    Else
        XLT_EMPRESAS = XLT_EMPRESAS & " or empr_cd_empresa <> " & DatEmpresas.Recordset.Fields!empr_cd_Empresa
    End If
    subConectarControleDadosNV DatEmpresas, "SELECT * FROM Empresas WHERE " & XLT_EMPRESAS, Estatico
    PanEmpresas.Visible = True
    XLB_IMPORTOU = True
    
    CmdDesistir.Caption = "&Sair"
End Sub
Private Sub CmdDesistir_Click()
    Unload Me
End Sub



Private Sub CmdProcuraBD_Click()
    Dim GuardaCaminho As String
    DlgAbrirFigura.Action = 1
    GuardaCaminho = TxtArquivoBD.Caption
    If DlgAbrirFigura.FileName <> "" Then
        GuardaCaminho = DlgAbrirFigura.FileName
    End If
        
    'If UCase$(Mid(GuardaCaminho, (Len(GuardaCaminho) - 11), 12)) = "QUALIFIN.MDB" Then
        TxtArquivoBD.Caption = GuardaCaminho
    'End If
'ConexaoMfin.Close
    ConexaoMfin.Provider = "Microsoft.Jet.OLEDB.4.0"
    'ConexaoMfin.Open App.Path & "\mfin.mdb", "admin", ""
    ConexaoMfin.Open GuardaCaminho, "admin", ""

    CmdContinua.Enabled = True
End Sub





Private Sub Form_Load()
    MskPlanoC.Mask = ""
    MskPlanoC.Text = ""
    MskPlanoC.Mask = pPlanoCmascara
    
    MskPcr.Mask = ""
    MskPcr.Text = ""
    MskPcr.Mask = pPCRmascara

End Sub









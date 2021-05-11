VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRelAnexoFluxoRecebiveisSantander 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Anexo de Fluxo Recebido - Santander"
   ClientHeight    =   3960
   ClientLeft      =   3330
   ClientTop       =   3615
   ClientWidth     =   10545
   Icon            =   "FrmAnexoFluxoRecebidoSantander.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleMode       =   0  'User
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgSalvar 
      Left            =   3240
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   9460
      TabIndex        =   24
      Top             =   3560
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc DatEmpreendimento1 
      Height          =   330
      Left            =   120
      Top             =   4200
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
      Caption         =   "DatEmpreendimento"
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
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Height          =   345
      Left            =   8280
      TabIndex        =   23
      Top             =   3560
      Width           =   1035
   End
   Begin VB.Frame FraPrincipal 
      Height          =   3500
      Left            =   60
      TabIndex        =   25
      Top             =   0
      Width           =   10440
      Begin MSComctlLib.TreeView TvwEmpreendimentos 
         Height          =   3160
         Left            =   80
         TabIndex        =   0
         Top             =   200
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   5583
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ComctlLib.ProgressBar pgbImpressao 
         Height          =   375
         Left            =   80
         TabIndex        =   26
         Top             =   3495
         Visible         =   0   'False
         Width           =   10280
         _ExtentX        =   18150
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Frame FraCotacao 
         Caption         =   "Cotação"
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
         Left            =   4760
         TabIndex        =   19
         Top             =   2640
         Width           =   5595
         Begin VB.OptionButton optIndexador2 
            Caption         =   "Indexador 2"
            Height          =   255
            Left            =   4200
            TabIndex        =   22
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optIndexador1 
            Caption         =   "Indexador 1"
            Height          =   195
            Left            =   2280
            TabIndex        =   21
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optValorHistorico 
            Caption         =   "Valor Historico"
            Height          =   195
            Left            =   360
            TabIndex        =   20
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame FraDados 
         Height          =   2295
         Left            =   4760
         TabIndex        =   1
         Top             =   100
         Width           =   5595
         Begin VB.TextBox TxtNomeResponsavel 
            Height          =   375
            Left            =   3315
            TabIndex        =   18
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox TxtNomeObra 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   3315
            TabIndex        =   16
            Top             =   1200
            Width           =   2175
         End
         Begin VB.TextBox txtVlrPresente 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   3315
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DtBase 
            Height          =   375
            Left            =   900
            TabIndex        =   3
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
            _Version        =   393216
            Format          =   69074945
            CurrentDate     =   39702
         End
         Begin MSComCtl2.DTPicker DtPeriodoCarencia 
            Height          =   375
            Left            =   900
            TabIndex        =   9
            Top             =   1680
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
            _Version        =   393216
            Format          =   69074945
            CurrentDate     =   39707
         End
         Begin MSComCtl2.DTPicker dtChaves 
            Height          =   375
            Left            =   900
            TabIndex        =   5
            Top             =   720
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
            _Version        =   393216
            Format          =   69074945
            CurrentDate     =   39707
         End
         Begin MSComCtl2.DTPicker DtLimite 
            Height          =   375
            Left            =   900
            TabIndex        =   7
            Top             =   1200
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
            _Version        =   393216
            Format          =   69074945
            CurrentDate     =   39707
         End
         Begin MSDataListLib.DataCombo CboTiposPlanos 
            Bindings        =   "FrmAnexoFluxoRecebidoSantander.frx":2AFA
            Height          =   315
            Left            =   3315
            TabIndex        =   14
            Top             =   720
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "tipl_tx_Descricao"
            BoundColumn     =   "tipl_cd_TipoPlano"
            Text            =   "CboTiposPlanos"
         End
         Begin VB.Label LblResponsavel 
            Caption         =   "Respon- sável"
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
            Left            =   2415
            TabIndex        =   17
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label lblNomeObra 
            Caption         =   "Nome da Obra"
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
            Left            =   2415
            TabIndex        =   15
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label LblRepasse 
            Caption         =   "Repasse"
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
            Left            =   2415
            TabIndex        =   13
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblTaxaPresente 
            Caption         =   "Taxa VLR Presente"
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
            Left            =   2415
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "%"
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
            Left            =   4440
            TabIndex        =   12
            Top             =   360
            Width           =   255
         End
         Begin VB.Label LblPeriodoCarencia 
            Caption         =   "Período Carência"
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
            Left            =   80
            TabIndex        =   8
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label LblDataLimite 
            Caption         =   "Data Limite"
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
            Left            =   80
            TabIndex        =   6
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label LblEntregaChaves 
            Caption         =   "Entrega Chaves"
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
            Left            =   80
            TabIndex        =   4
            Top             =   720
            Width           =   855
         End
         Begin VB.Label LblDataBase 
            Caption         =   "Data Base"
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
            Left            =   80
            TabIndex        =   2
            Top             =   240
            Width           =   855
         End
      End
   End
   Begin MSAdodcLib.Adodc DatTiposPlanos 
      Height          =   330
      Left            =   840
      Top             =   4200
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
      Caption         =   "DatTiposPlanos"
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
   Begin MSAdodcLib.Adodc DatEmpreendimento2 
      Height          =   330
      Left            =   120
      Top             =   4560
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
      Caption         =   "DatEmpreendimento"
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
Attribute VB_Name = "FrmRelAnexoFluxoRecebiveisSantander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XFO_ExlObj As Excel.Application      ' Create excel object
Public XFI_QtdEmpreendimento As Integer     'Quantidade de empreendimentos

Private Sub Form_Activate()
    Set Formulario = FrmRelAnexoFluxoRecebiveisSantander
End Sub

Private Sub Form_Load()
    subConectarControleDadosNV DatEmpreendimento1, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_cd_empreendimento", Estatico
    subConectarControleDadosNV DatEmpreendimento2, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_cd_empreendimento", Estatico
    subConectarControleDadosNV DatTiposPlanos, "SELECT * FROM TiposPlanos ORDER BY tipl_tx_Descricao", Estatico
    
    DtBase.Value = Date
    dtChaves.Value = Date
    DtPeriodoCarencia.Value = Date
    DtLimite.Value = Date
    
    XFI_QtdEmpreendimento = 0
    PreencheTreeViewEmpreendimentos
End Sub

Private Sub PreencheTreeViewEmpreendimentos()
    Dim nodeEmpd As MSComctlLib.Node
            
    TvwEmpreendimentos.Nodes.Clear
        
    'Cria os nós principais
    DatEmpreendimento1.Recordset.MoveFirst
    
    While Not DatEmpreendimento1.Recordset.EOF
        Set nodeEmpd = TvwEmpreendimentos.Nodes.Add(, , "N" & DatEmpreendimento1.Recordset.Fields("empd_cd_Empreendimento"), _
            DatEmpreendimento1.Recordset.Fields("empd_cd_Empreendimento") & " - " & DatEmpreendimento1.Recordset.Fields("empd_tx_Nome"))
        
        nodeEmpd.Tag = "Nó"
        
        'Cria as folhas do nó criado
        CarregaRegistros nodeEmpd
        
        'Atualiza a quantidade de empreendimentos
        XFI_QtdEmpreendimento = XFI_QtdEmpreendimento + 1
        
        DatEmpreendimento1.Recordset.MoveNext
    Wend
End Sub

'Preenche a árvore do nó selecionado
Private Sub CarregaRegistros(ByVal nodeNo As MSComctlLib.Node)
    Dim nodeFolha As MSComctlLib.Node
    
    'Se os registros do nó já foram carregados não faz nada
    If nodeNo.Children > 0 Then Exit Sub
    
    DatEmpreendimento2.Recordset.MoveFirst
    
    While Not DatEmpreendimento2.Recordset.EOF
        If (nodeNo.Text <> DatEmpreendimento2.Recordset.Fields("empd_cd_Empreendimento") & " - " & DatEmpreendimento2.Recordset.Fields("empd_tx_Nome")) Then
            Set nodeFolha = TvwEmpreendimentos.Nodes.Add(nodeNo, tvwChild, , _
                            DatEmpreendimento2.Recordset.Fields("empd_cd_Empreendimento") & " - " & DatEmpreendimento2.Recordset.Fields("empd_tx_Nome"))
            
            nodeFolha.Tag = "Folha"
        End If
        DatEmpreendimento2.Recordset.MoveNext
    Wend
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
  
    Dim XLO_COMANDO     As New ADODB.Command
    Dim XLO_REPASSE     As New ADODB.Recordset
    Dim i               As Integer  'Auxiliar
    Dim j               As Integer  'Auxiliar
    Dim sheet           As Integer  'Quantidade de planilhas no workbook

    Dim XLI_EMPREENDIMENTOINICIAL   As Integer
    Dim XLI_EMPREENDIMENTOFINAL     As Integer
    Dim XLT_EMPREENDIMENTOS         As String
    Dim XLO_EMPREENDIMENTO()        As String
    
    On Error GoTo TrataErro
                                 
    'Verificação de campos obrigatórios
    If dtChaves.Value >= DtLimite.Value Then
        MsgBox "A data limite não pode ser menor ou igual que a data de entrega das chaves!", vbInformation, "Informação"
        Exit Sub
    End If
    
    If txtVlrPresente.Text = "" Then
        MsgBox "É necessário preencher o campo ""Taxa VLR Presente""!", vbInformation, "Informação"
        Exit Sub
    End If
    
    If CboTiposPlanos.BoundText = "" Then
        MsgBox "É necessário selecionar o plano de repasse!", vbInformation, "Informação"
        Exit Sub
    End If
    
                                 
    'Seta o Titulo da Tela de Dialogo salvar
    dlgSalvar.DialogTitle = "Salvar Arquivo"
    
    'Seta o filtro das extensões de arquivos que será usada
    dlgSalvar.Filter = "Documento do Excel(*.xls)|*.xls"
    
    'Indica qual sera o filtro padrão
    dlgSalvar.FilterIndex = 1
    
    dlgSalvar.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
            
    'Exibe a tela do dialogo para o usuário
    dlgSalvar.ShowSave
        
    'Verifica se o usuário selecionou o caminho do arquivo
    If (dlgSalvar.FileName <> "") Then
        
        'Indexador
        If optValorHistorico.Value = True Then
          XLI_INDEXADOR = 0
        ElseIf optIndexador1.Value = True Then
          XLI_INDEXADOR = 1
        ElseIf optIndexador2.Value = True Then
          XLI_INDEXADOR = 2
        End If
    
        'Altera o cursor do mouse
        FrmRelAnexoFluxoRecebiveisSantander.MousePointer = vbHourglass
        
        'Inicializa o objeto excel
        Set XFO_ExlObj = CreateObject("excel.application")
        
        'Adiciona um WorkBook
        XFO_ExlObj.Workbooks.Add
        
        
        'Verifica quantas planilhas serão criadas
        i = 1
        j = 0
        While i < TvwEmpreendimentos.Nodes.Count
            If (TvwEmpreendimentos.Nodes.Item(i).Tag = "Nó") And (TvwEmpreendimentos.Nodes.Item(i).Checked = True) Then
                j = j + 1
            End If
            i = i + 1
        Wend
                
        'Adiciona as planilhas necessárias
        If (j > 2) Then
            For i = 2 To j
                XFO_ExlObj.Sheets.Add
            Next
        End If
        
        'Inicializa a barra de progresso
        InicializarProgresso 0, j * 20 + 1
        
        
        'Impede que o usuário veja a planilha enquanto ela está sendo preenchida
        XFO_ExlObj.Visible = False
                
        'Preenche a primeira aba da planilha com a forma de preenchimento
        CarregabaFormaPreenchimento
        
        AvancarProgresso 1
                
        i = 1
        sheet = 2
        XLT_EMPREENDIMENTOS = ""
        
        While i < TvwEmpreendimentos.Nodes.Count
            
            If (TvwEmpreendimentos.Nodes.Item(i).Tag = "Nó") And (TvwEmpreendimentos.Nodes.Item(i).Checked = True) Then
                
                XLO_EMPREENDIMENTO = Split(TvwEmpreendimentos.Nodes.Item(i).Text, "-")
                XLT_EMPREENDIMENTOS = "'" & Trim(XLO_EMPREENDIMENTO(0)) & "'"
                              
                For j = i + 1 To i + XFI_QtdEmpreendimento - 1
                    If (TvwEmpreendimentos.Nodes.Item(j).Checked = True) Then
                        XLO_EMPREENDIMENTO = Split(TvwEmpreendimentos.Nodes.Item(j).Text, "-")
                        XLT_EMPREENDIMENTOS = XLT_EMPREENDIMENTOS + ", " + "'" & Trim(XLO_EMPREENDIMENTO(0)) & "'"
                    End If
                Next
                
                '------------------
                'Imprime relatório
                '------------------
                If (XLT_EMPREENDIMENTOS <> "") Then
                
                    'Executa a procedure para pegar os valores previstos de todos os títulos dos empreendimentos
                    With XLO_COMANDO
                        .CommandTimeout = 300
                        .CommandText = "spCAPI_sel_fluxoRecebiveisSANTANDER"
                        .CommandType = adCmdStoredProc
                        Set XLO_COMANDO.ActiveConnection = Conexao
                        'Cria automaticamente todos os parâmetros dentro do objeto Comando
                        .Parameters.Refresh
                    End With
                                    
                    XLO_COMANDO.Parameters(1).Value = PCodEmpresa               'Empresa
                    XLO_COMANDO.Parameters(2).Value = DtBase.Value              'Data base para o cálculo
                    XLO_COMANDO.Parameters(3).Value = DtLimite.Value            'Data limite
                    XLO_COMANDO.Parameters(4).Value = XLT_EMPREENDIMENTOS       'Empreendimentos
                    XLO_COMANDO.Parameters(5).Value = XLI_INDEXADOR             'Indexador
                    XLO_COMANDO.Parameters(6).Value = CboTiposPlanos.BoundText  'Plano Repasse
                    
                    'Passa para o recordset o resultado da consulta
                    Set XGO_RSRELATORIO = XLO_COMANDO.Execute
                           
                    'Executa a procedure para pegar o valor de repasse de todos os imóveis dos empreendimentos
                    With XLO_COMANDO
                        .CommandTimeout = 300
                        .CommandText = "spCAPI_sel_fluxoRecebiveisRepasseSANTANDER"
                        .CommandType = adCmdStoredProc
                        Set XLO_COMANDO.ActiveConnection = Conexao
                        'Cria automaticamente todos os parâmetros dentro do objeto Comando
                        .Parameters.Refresh
                    End With
                                    
                    XLO_COMANDO.Parameters(1).Value = PCodEmpresa               'Empresa
                    XLO_COMANDO.Parameters(2).Value = DtBase.Value              'Data base para o cálculo
                    XLO_COMANDO.Parameters(3).Value = XLT_EMPREENDIMENTOS       'Empreendimentos
                    XLO_COMANDO.Parameters(4).Value = XLI_INDEXADOR             'Indexador
                    XLO_COMANDO.Parameters(5).Value = CboTiposPlanos.BoundText  'Plano Repasse
                    
                    Set XLO_REPASSE = XLO_COMANDO.Execute
                           
                    XLO_EMPREENDIMENTO = Split(TvwEmpreendimentos.Nodes.Item(i).Text, "-")
                    
                    'Cria a planilha com os valores previstos dos imóveis
                    ExportaExcel XLO_EMPREENDIMENTO(1), XGO_RSRELATORIO, XLO_REPASSE, sheet
                            
                    
                    XGO_RSRELATORIO.Close
                    XLO_REPASSE.Close
                    
                    'Incrementa a quantidade de planilhas
                    sheet = sheet + 1
                End If
            End If
            
            XLT_EMPREENDIMENTOS = ""
            i = i + XFI_QtdEmpreendimento
        Wend
        
        
        'Salva a planilha no caminho selecionado pelo usuário
        XFO_ExlObj.Workbooks(1).SaveAs dlgSalvar.FileName
        XFO_ExlObj.Workbooks(1).Close
        XFO_ExlObj.Application.Quit
        Set XFO_ExlObj = Nothing
                
        
        'Limpa o caminho do arquivo
        dlgSalvar.FileName = ""
        FrmRelAnexoFluxoRecebiveisSantander.MousePointer = vbDefault
        MsgBox "Relatório Exportado com sucesso!", vbInformation, "Informação"
        
        
        FinalizarProgresso
    End If
Exit Sub

TrataErro:
    If Err.Number = cdlCancel Then
        Exit Sub
    End If
    
    'Fecha a aplicação - Lucas Santiago - 02.02.2009
    XFO_ExlObj.Workbooks(1).SaveAs dlgSalvar.FileName
    XFO_ExlObj.Workbooks(1).Close
    XFO_ExlObj.Application.Quit
    Set XFO_ExlObj = Nothing
    
    MsgBox "Ocorreu um erro na exportação do relatório. Contate o suporte. Detalhe do Erro: " & Err.Description, vbCritical, "Erro"
    DeleteFile (dlgSalvar.FileName)
        
    FrmRelAnexoFluxoRecebiveisSantander.MousePointer = vbDefault
End Sub

Private Sub InicializarProgresso(valorMinimo As Integer, valorMaximo As Integer)
    'Reajusta a altura do formulário
    FrmRelAnexoFluxoRecebiveisSantander.Height = 4965
    CmdImprimir.Top = 4080
    CmdCancelar.Top = 4080
    FraPrincipal.Height = 4005
    
    'Inicializa o progress bar
    pgbImpressao.Min = valorMinimo
    pgbImpressao.Max = valorMaximo
    pgbImpressao.Value = 0
    pgbImpressao.Visible = True
End Sub

Private Sub FinalizarProgresso()
    'Finaliza o progress bar
    pgbImpressao.Value = 0
    pgbImpressao.Visible = False
    
    'Reajusta a altura do formulário
    FraPrincipal.Height = 3500
    CmdImprimir.Top = 3560
    CmdCancelar.Top = 3560
    FrmRelAnexoFluxoRecebiveisSantander.Height = 4440
End Sub

Private Sub AvancarProgresso(valor As Integer)
    If (pgbImpressao.Value + valor <= pgbImpressao.Max) Then
        pgbImpressao.Value = pgbImpressao.Value + valor
    End If
End Sub

Private Sub ExportaExcel(nomeEmpreendimento As String, dados As ADODB.Recordset, repasse As ADODB.Recordset, sheet As Integer)
        
    Dim cont, ProxLinha, ProxColuna, coluna As Integer 'Cont -  Calcula o total de registros obtidos da consulta
    Dim letraColuna As String 'Recebe o valor númerico da coluna convertido para letra
    Dim chaves As Boolean 'indica se a data é até chaves ou depois
    Dim y As Integer
    Dim i As Integer
    Dim LinhaAteChaves, LinhaPosChaves As Integer  'Guarda o número das linhas
        
    
    'Seleciona a próxima planilha vazia para ser preenchida
    XFO_ExlObj.Sheets(sheet).Select
    
    
    'Configurando margem da página
    With XFO_ExlObj.ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.64)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.54)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .Orientation = xlLandscape
        .PrintTitleRows = "$1:$12" 'Repete o cabeçalho em todas as páginas
    End With
        
    
    'Desativa o cálculo automático das fórmulas
    XFO_ExlObj.Application.Calculation = xlCalculationManual
    
    
    'Cabeçalho da planilha
    With XFO_ExlObj.Sheets(sheet)
        'Insere o logo da empresa.
        .Shapes.AddPicture "c:\QualiAdmFin\QualiCapi\Relatorios\Imagens\imgFluxoRecebiveisSantander.png", True, True, 1, 1, 120, 55
        
        'Seta o nome da aba da planilha
        .Name = nomeEmpreendimento
        
        'Cabeçalho
        .Cells(1, 4).Value = PEmpresa
        If (TxtNomeObra.Text = "") Then .Cells(3, 4).Value = "NOME DA OBRA"
        If (TxtNomeObra.Text <> "") Then .Cells(3, 4).Value = UCase(TxtNomeObra.Text)
        .Cells(5, 4).Value = nomeEmpreendimento
                
        'Mescla a celula do intervalo especificado
        .Range("A6:B6").Merge
        .Cells(6, 1).Value = "Moeda:"
        .Cells(6, 3).Value = "Real"
        
        .Range("A8:B8").Merge
        .Cells(8, 1).Value = "Taxa Vlr. Presente:"
        .Cells(8, 3).Value = CInt(txtVlrPresente.Text) / 100
        
        'Formata a coluna para o tipo Percentual com quatro casas decimais
        .Cells(8, 3).NumberFormat = "0.0000%"
        .Cells(8, 4).Value = "a.a."
        
        .Range("A10:B10").Merge
        .Cells(10, 1).Value = "Data Cálculo"
        .Cells(10, 3).Value = CDate(DtBase.Value)
        
        .Range("A12:B12").Merge
        .Cells(12, 1).Value = "DATAS DE "
        .Range("A13:B13").Merge
        .Cells(13, 1).Value = "VENCIMENTO"
        .Range("A12:B13").Font.Bold = True
        
        'Largura da coluna
        .Range("A10").ColumnWidth = 11
        .Range("B10").ColumnWidth = 11
        .Range("C10").ColumnWidth = 10.71
        .Range("D10").ColumnWidth = 10.71
        .Range("E10").ColumnWidth = 10.71
        .Range("F10").ColumnWidth = 10.71
        .Range("G10").ColumnWidth = 10.71
        .Range("H10").ColumnWidth = 10.71
        .Range("I10").ColumnWidth = 10.71
        .Range("J10").ColumnWidth = 10.71
        .Range("K10").ColumnWidth = 10.71
        
        'Seta a fonte da celula para negrito
        .Range("A1:D10").Font.Bold = True
        
        'Centraliza o valor da célula
        .Range("A1:C13").HorizontalAlignment = xlHAlignCenter
                
        'Copia cabeçalho
        .Cells(1, (dados.Fields.Count * 2) + 2).Value = "=D1"
        .Cells(3, (dados.Fields.Count * 2) + 3).Value = "=D3"
        .Cells(5, (dados.Fields.Count * 2) + 3).Value = "=D5"
        .Cells(6, (dados.Fields.Count * 2) + 2).Value = "=C6"
        .Cells(8, (dados.Fields.Count * 2) + 2).Value = "=C8"
        .Cells(8, (dados.Fields.Count * 2) + 3).Value = "=D8"
        .Cells(10, (dados.Fields.Count * 2) + 2).Value = "=C10"
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 2) & "8:" & ConvertToLetter((dados.Fields.Count * 2) + 2) & "10").HorizontalAlignment = xlHAlignLeft
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 2) & 1 & ":" & ConvertToLetter((dados.Fields.Count * 2) + 4) & 10).Font.Bold = True
    End With
    
    AvancarProgresso 1
                   
    
    'Dados da Planilha
    chaves = False
    ProxLinha = 12
    coluna = 3
    ProxColuna = 1
    cont = 0
                           
                           
    'Preenche o titulo da tabela com os números dos imoveís
    With XFO_ExlObj.Sheets(sheet)
        
        .Range(ConvertToLetter(coluna) & ProxLinha & ":" & ConvertToLetter(dados.Fields.Count * 2) & ProxLinha).NumberFormat = "@"
        
        For i = 1 To dados.Fields.Count - 1
            .Cells(12, coluna).Value = "unidade n°"
            .Cells(13, coluna).Value = "NOMINAL"
            .Range(ConvertToLetter(coluna) & "12:" & ConvertToLetter(coluna) & "13").Font.Bold = True
            .Range(ConvertToLetter(coluna) & "12:" & ConvertToLetter(coluna) & "13").HorizontalAlignment = xlHAlignCenter
            
            coluna = coluna + 1
                        
            .Cells(ProxLinha, coluna).Value = dados.Fields(i).Name
            .Cells(ProxLinha + 1, coluna).Value = "VLR PRES"
            .Cells(ProxLinha + 1, coluna).Font.Bold = True
            .Range(ConvertToLetter(coluna) & ProxLinha & ":" & ConvertToLetter(coluna) & ProxLinha + 1).HorizontalAlignment = xlHAlignCenter
            
            coluna = coluna + 1
        Next
                        
        'Insere o primeiro conjunto de fórmulas
        .Range(ConvertToLetter(coluna + 1) & "12:" & ConvertToLetter(coluna + 2) & "12").Merge
        .Range(ConvertToLetter(coluna + 1) & "12:" & ConvertToLetter(coluna + 4) & "13").Font.Bold = True
        .Range(ConvertToLetter(coluna + 2) & "12:" & ConvertToLetter(coluna + 4) & "13").HorizontalAlignment = xlHAlignCenter
        .Cells(12, coluna + 1).Value = "TOTAL APARTAMENTOS"
        .Cells(13, coluna + 1).Value = "NOMINAL"
        .Cells(13, coluna + 2).Value = "VLR PRES"
        .Cells(12, coluna + 3).Value = "QUANT. DE"
        .Cells(13, coluna + 3).Value = "BOLETOS"
                        
        'Insere o segundo conjunto de fórmulas
        .Range(ConvertToLetter(coluna + 5) & "12:" & ConvertToLetter(coluna + 7) & "12").Merge
        .Range(ConvertToLetter(coluna + 5) & "13:" & ConvertToLetter(coluna + 7) & "13").Merge
        .Range(ConvertToLetter(coluna + 5) & "12:" & ConvertToLetter(coluna + 7) & "14").Font.Bold = True
        .Range(ConvertToLetter(coluna + 5) & "12:" & ConvertToLetter(coluna + 7) & "14").HorizontalAlignment = xlHAlignCenter
        .Cells(12, coluna + 5).Value = "RECEBÍVEIS NO PERÍODO"
        .Cells(13, coluna + 5).Value = "DE CARÊNCIA"
        .Cells(14, coluna + 5).Value = "DATAS"
        .Cells(14, coluna + 6).Value = "VLR PRES"
        .Cells(14, coluna + 7).Value = "BOLETOS"
                
                
        'Configura largura das colunas de fórmulas
        .Range(ConvertToLetter(coluna) & "10").ColumnWidth = 3
        XFO_ExlObj.Columns(ConvertToLetter(coluna + 1) & ":" & ConvertToLetter(coluna + 3)).Select
        XFO_ExlObj.Selection.ColumnWidth = 12
        
        .Range(ConvertToLetter(coluna + 4) & "10").ColumnWidth = 3
        XFO_ExlObj.Columns(ConvertToLetter(coluna + 5) & ":" & ConvertToLetter(coluna + 7)).Select
        XFO_ExlObj.Selection.ColumnWidth = 10.86
        
        
        'Coloca a borda dos titulos
        .Range("A12:" & ConvertToLetter(coluna - 1) & "12").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("A12:" & ConvertToLetter(coluna - 1) & "12").Borders(xlEdgeTop).ColorIndex = xlAutomatic
        .Range("A12:" & ConvertToLetter(coluna - 1) & "12").Borders(xlEdgeTop).Weight = 3
        
        .Range("A13:" & ConvertToLetter(coluna - 1) & 13).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A13:" & ConvertToLetter(coluna - 1) & 13).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Range("A13:" & ConvertToLetter(coluna - 1) & 13).Borders(xlEdgeBottom).Weight = 3
        
        .Range(ConvertToLetter(coluna + 1) & "12:" & ConvertToLetter(coluna + 3) & "12").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(ConvertToLetter(coluna + 1) & "12:" & ConvertToLetter(coluna + 3) & "12").Borders(xlEdgeTop).ColorIndex = xlAutomatic
        .Range(ConvertToLetter(coluna + 1) & "12:" & ConvertToLetter(coluna + 3) & "12").Borders(xlEdgeTop).Weight = 3
        
        .Range(ConvertToLetter(coluna + 1) & "13:" & ConvertToLetter(coluna + 3) & "13").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(ConvertToLetter(coluna + 1) & "13:" & ConvertToLetter(coluna + 3) & "13").Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Range(ConvertToLetter(coluna + 1) & "13:" & ConvertToLetter(coluna + 3) & "13").Borders(xlEdgeBottom).Weight = 3
        
        .Range(ConvertToLetter(coluna + 5) & "12:" & ConvertToLetter(coluna + 7) & "12").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(ConvertToLetter(coluna + 5) & "12:" & ConvertToLetter(coluna + 7) & "12").Borders(xlEdgeTop).ColorIndex = xlAutomatic
        .Range(ConvertToLetter(coluna + 5) & "12:" & ConvertToLetter(coluna + 7) & "12").Borders(xlEdgeTop).Weight = 3
        
        .Range(ConvertToLetter(coluna + 5) & "13:" & ConvertToLetter(coluna + 7) & "13").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(ConvertToLetter(coluna + 5) & "13:" & ConvertToLetter(coluna + 7) & "13").Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Range(ConvertToLetter(coluna + 5) & "13:" & ConvertToLetter(coluna + 7) & "13").Borders(xlEdgeBottom).Weight = 3
        
        .Range(ConvertToLetter(coluna + 5) & "14:" & ConvertToLetter(coluna + 7) & "14").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(ConvertToLetter(coluna + 5) & "14:" & ConvertToLetter(coluna + 7) & "14").Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Range(ConvertToLetter(coluna + 5) & "14:" & ConvertToLetter(coluna + 7) & "14").Borders(xlEdgeBottom).Weight = 3
    End With
    
    AvancarProgresso 3
              
    ProxLinha = ProxLinha + 2
    
    'Preenche a planilha com os dados dos títulos de cada imóvel
    While Not dados.EOF
        If (dados!V <= DtLimite.Value) Then
                    
            'Se já foram inseridos todos os títulos com data de vencimento anterior à entrega da chaves:
            If dados!V > dtChaves.Value And chaves = False Then
                If ProxLinha = 14 Then ProxLinha = ProxLinha + 2
                
                'Insere fórmula para somar o valor presente dos título
                With XFO_ExlObj.Sheets(sheet)
                    
                    .Range("A" & ProxLinha & ":B" & ProxLinha).Merge
                    .Cells(ProxLinha, 1).Value = "TOTAL ATÉ AS CHAVES"
                    .Cells(ProxLinha, 1).Font.Bold = True
                    .Cells(ProxLinha, 1).Interior.ColorIndex = 15
                                    
                    
                    'Insere fórmula nas colunas alternadamente
                    For y = 3 To ((dados.Fields.Count * 2) + 4)
                        letraColuna = ConvertToLetter(y)
                        .Cells(ProxLinha, y).Interior.ColorIndex = 15
                        .Cells(ProxLinha, y).Formula = "=SUM(" & letraColuna & "14:" & letraColuna & ProxLinha - 1 & ")"
                    Next
                    
                                                        
                    'Coloca XXXX na coluna referente ao conjunto de fórmulas "Total Apartamentos"
                    .Cells(ProxLinha, (dados.Fields.Count * 2) + 4).Value = "XXXXX"
                    .Cells(ProxLinha, (dados.Fields.Count * 2) + 1).Value = ""
                                                            
                    'Insere borda
                    XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":" & letraColuna & ProxLinha).Borders(xlEdgeTop).LineStyle = xlContinuous
                    XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":" & letraColuna & ProxLinha).Borders(xlEdgeTop).ColorIndex = xlAutomatic
                    XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":" & letraColuna & ProxLinha).Borders(xlEdgeTop).Weight = 3
                
                    XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":" & letraColuna & ProxLinha).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":" & letraColuna & ProxLinha).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                    XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":" & letraColuna & ProxLinha).Borders(xlEdgeBottom).Weight = 3
                    
                    'Guarda o número da linha com o cálculo dos títulos até a entrega das chaves
                    LinhaAteChaves = ProxLinha
                End With
                
                chaves = True
                
                'Passa para a próxima linha da planilha
                ProxLinha = ProxLinha + 1
            End If
            
            
            'Preenche planilha com o valor previsto dos títulos de cada imóvel. O preenchimento é feito horizontalmente
            For i = 0 To (dados.Fields.Count - 1)
                If i = 0 Then
                    XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":B" & ProxLinha).Merge
                    XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha).HorizontalAlignment = xlHAlignCenter
                    XFO_ExlObj.Sheets(sheet).Cells(ProxLinha, ProxColuna).Value = CDate(dados!V)
                Else
                    If (dados.Fields(i) = 0) Then XFO_ExlObj.Sheets(sheet).Cells(ProxLinha, ProxColuna).Value = ""
                    If (dados.Fields(i) > 0) Then XFO_ExlObj.Sheets(sheet).Cells(ProxLinha, ProxColuna).Value = dados.Fields(i)
                End If
                
                ProxColuna = ProxColuna + 2
            Next
            
            
            'Passa para a próxima linha da planilha
            ProxLinha = ProxLinha + 1
            
            
            'Incrementa a variável que calcula o total de registros da consulta
            cont = cont + 1
            
            ProxColuna = 1
            dados.MoveNext
        End If
        
        'Preenche a última linha de valor previsto: soma dos títulos de cada imóvel cuja data de vencimento é superior à data limite informada pelo usuário
        If Not dados.EOF Then
            If dados!V > DtLimite.Value Then
                XFO_ExlObj.Sheets(sheet).Cells(ProxLinha, ProxColuna).Value = "Posterior a"
                XFO_ExlObj.Sheets(sheet).Range(ConvertToLetter(ProxColuna) & ProxLinha).HorizontalAlignment = xlHAlignRight
                
                XFO_ExlObj.Sheets(sheet).Cells(ProxLinha, ProxColuna + 1).Value = CDate(DtLimite.Value)
                XFO_ExlObj.Sheets(sheet).Range(ConvertToLetter(ProxColuna + 1) & ProxLinha).HorizontalAlignment = xlHAlignLeft
                
                XFO_ExlObj.Sheets(sheet).Range(ConvertToLetter(ProxColuna) & ProxLinha & ":" & ConvertToLetter(ProxColuna + 1) & ProxLinha).Font.Bold = True
                            
                ProxColuna = 3
                'Para cada coluna, soma todos os títulos cuja data de vencimento é maior que a data limite informada
                For i = 1 To (dados.Fields.Count - 1)
                    XFO_ExlObj.Sheets(sheet).Cells(ProxLinha, ProxColuna).Value = dados.Fields(i)
                    XFO_ExlObj.Sheets(sheet).Cells(ProxLinha, ProxColuna + 1).Formula = "=" & ConvertToLetter(ProxColuna) & ProxLinha & " /((1+$C$8)^(($A" & ProxLinha - 1 & " -$C$10)/360))"
                    ProxColuna = ProxColuna + 2
                Next
                
                dados.MoveNext
                ProxColuna = 1
            End If
        Else
            XFO_ExlObj.Sheets(sheet).Cells(ProxLinha, ProxColuna).Value = "Posterior a"
            XFO_ExlObj.Sheets(sheet).Range(ConvertToLetter(ProxColuna) & ProxLinha).HorizontalAlignment = xlHAlignRight
            
            XFO_ExlObj.Sheets(sheet).Cells(ProxLinha, ProxColuna + 1).Value = CDate(DtLimite.Value)
            XFO_ExlObj.Sheets(sheet).Range(ConvertToLetter(ProxColuna + 1) & ProxLinha).HorizontalAlignment = xlHAlignLeft
            
            XFO_ExlObj.Sheets(sheet).Range(ConvertToLetter(ProxColuna) & ProxLinha & ":" & ConvertToLetter(ProxColuna + 1) & ProxLinha).Font.Bold = True
            
            ProxColuna = 3
            For i = 1 To (dados.Fields.Count - 1)
                XFO_ExlObj.Sheets(sheet).Cells(ProxLinha, ProxColuna).Value = "0"
                XFO_ExlObj.Sheets(sheet).Cells(ProxLinha, ProxColuna + 1).Formula = "=" & ConvertToLetter(ProxColuna) & ProxLinha & " /((1+$C$8)^(($A" & ProxLinha - 1 & " -$C$10)/360))"
                ProxColuna = ProxColuna + 2
            Next
            
            ProxColuna = 1
        End If
    Wend
    
    AvancarProgresso 5
    
    
    ProxLinha = ProxLinha + 1
    
    'Insere fórmula para somar os valores a receber dos títulos cuja data de vencimento é maior que a data de entrega das chaves
    With XFO_ExlObj.Sheets(sheet)
        .Range("A" & ProxLinha & ":B" & ProxLinha).Merge
        .Cells(ProxLinha, 1).Value = "TOTAL PÓS CHAVES"
        .Cells(ProxLinha, 1).Font.Bold = True
        .Cells(ProxLinha, 1).Interior.ColorIndex = 15
        
        For y = 3 To ((dados.Fields.Count * 2) + 4)
            letraColuna = ConvertToLetter(y)
            .Cells(ProxLinha, y).Formula = "=SUM(" & letraColuna & LinhaAteChaves + 1 & ":" & letraColuna & ProxLinha - 1 & ")"
            .Cells(ProxLinha, y).Interior.ColorIndex = 15
        Next
        
        .Cells(ProxLinha, (dados.Fields.Count * 2) + 4).Value = "XXXXX"
        .Cells(ProxLinha, (dados.Fields.Count * 2) + 1).Value = ""
        
        'Insere borda no campo da fórmula
        XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":" & letraColuna & ProxLinha).Borders(xlEdgeTop).LineStyle = xlContinuous
        XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":" & letraColuna & ProxLinha).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":" & letraColuna & ProxLinha).Borders(xlEdgeTop).Weight = 3
        
        XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":" & letraColuna & ProxLinha).Borders(xlEdgeBottom).LineStyle = xlContinuous
        XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":" & letraColuna & ProxLinha).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        XFO_ExlObj.Sheets(sheet).Range("A" & ProxLinha & ":" & letraColuna & ProxLinha).Borders(xlEdgeBottom).Weight = 3
        
        LinhaPosChaves = ProxLinha
    End With
    
    
    AvancarProgresso 1
    
    
    ProxColuna = 3
    
    
    'Insere a fórmula para calcular o valor presente
    For i = 0 To dados.Fields.Count - 2
        With XFO_ExlObj.Sheets(sheet)
            'Insere a fórmula na célula da primeira linha
            .Cells(14, ProxColuna + 1).Formula = "=" & ConvertToLetter(ProxColuna) & "14/((1+$C$8)^(($A" & "14 -$C$10)/360))"
            
            'Seleciona a célula com a fórmula
            .Range(ConvertToLetter(ProxColuna + 1) & "14").Select
            
            'Copia a célula
            XFO_ExlObj.CutCopyMode = False
            XFO_ExlObj.Selection.Copy
            
            'Cola a fórmula nas células abaixo da primeira
            .Range(ConvertToLetter(ProxColuna + 1) & "15:" & ConvertToLetter(ProxColuna + 1) & (LinhaAteChaves - 1)).Select
            XFO_ExlObj.ActiveSheet.Paste
            
            .Range(ConvertToLetter(ProxColuna + 1) & LinhaAteChaves + 1 & ":" & ConvertToLetter(ProxColuna + 1) & (LinhaPosChaves - 2)).Select
            XFO_ExlObj.ActiveSheet.Paste
        End With
        
        ProxColuna = ProxColuna + 2
    Next
    
    
    AvancarProgresso 1
    
    
    'Insere as fórmulas que somam o valor previsto dos títulos de cada imóvel
    With XFO_ExlObj.Sheets(sheet)
        ProxLinha = ProxLinha + 1
        .Range("A" & ProxLinha & ":" & "B" & ProxLinha).Merge
        .Cells(ProxLinha, 1).Value = "Total sem Repasse"
        .Cells(ProxLinha, 1).Font.Bold = True
        .Cells(ProxLinha, 1).HorizontalAlignment = xlHAlignCenter
        
        ProxLinha = ProxLinha + 1
        .Range("A" & ProxLinha & ":" & "B" & ProxLinha).Merge
        .Cells(ProxLinha, 1).Value = "Repasse Data"
        .Cells(ProxLinha, 1).Font.Bold = True
        .Cells(ProxLinha, 1).HorizontalAlignment = xlHAlignCenter
        
        ProxLinha = ProxLinha + 1
        .Range("A" & ProxLinha & ":" & "B" & ProxLinha).Merge
        .Cells(ProxLinha, 1).Value = "Repasse Valor"
        .Cells(ProxLinha, 1).Font.Bold = True
        .Cells(ProxLinha, 1).HorizontalAlignment = xlHAlignCenter
        
        ProxLinha = ProxLinha + 1
        .Range("A" & ProxLinha & ":" & "B" & ProxLinha).Merge
        .Cells(ProxLinha, 1).Value = "Total com Repasse"
        .Cells(ProxLinha, 1).Font.Bold = True
        .Cells(ProxLinha, 1).HorizontalAlignment = xlHAlignCenter
        
        ProxLinha = ProxLinha + 1
        .Cells(ProxLinha, 1).Value = "Parcelas até"
        .Cells(ProxLinha, 2).Value = CDate(DtLimite.Value)
        .Cells(ProxLinha, 2).HorizontalAlignment = xlHAlignLeft
        .Range("A" & ProxLinha & ":" & "B" & ProxLinha).Font.Italic = True
        
        .Cells(ProxLinha, 1).HorizontalAlignment = xlHAlignCenter
        
        ProxLinha = ProxLinha + 1
        .Cells(ProxLinha, 1).Value = "TOTAL ATÉ"
        .Cells(ProxLinha, 2).Value = CDate(DtLimite.Value)
        .Cells(ProxLinha, 2).HorizontalAlignment = xlHAlignLeft
        .Range("A" & ProxLinha & ":" & "B" & ProxLinha).Font.Bold = True
        
        
        For y = 3 To ((dados.Fields.Count * 2) + 3)
            letraColuna = ConvertToLetter(y)
            
            'Total sem repasse:
            .Cells(ProxLinha - 5, y).Formula = "=" & letraColuna & LinhaAteChaves & "+" & letraColuna & LinhaPosChaves
            
            'Total com repasse:
            .Cells(ProxLinha - 2, y).Formula = "=" & letraColuna & LinhaPosChaves + 1 & "+" & letraColuna & LinhaPosChaves + 3
            
            'Total até a data limite:
            .Cells(ProxLinha, y).Formula = "=SUM(" & letraColuna & "14:" & letraColuna & (LinhaAteChaves - 1) & "," & letraColuna & (LinhaAteChaves + 1) & ":" & letraColuna & (LinhaPosChaves - 2) & ")"
            .Cells(ProxLinha, y).Font.Bold = True
        Next
        
       
        AvancarProgresso 1
                
        'Linha de repasse
        If Not repasse.EOF Then
            i = 1
            For y = 3 To ((dados.Fields.Count * 2) + 3) Step 2
                letraColuna = ConvertToLetter(y)
                            
                If i < repasse.Fields.Count Then
                    'Repasse data:
                    .Cells(ProxLinha - 4, y).Value = CDate(repasse!V)
                        
                    'Repasse valor:
                    If XFO_ExlObj.Sheets(sheet).Cells(12, y + 1) = repasse.Fields(i).Name Then
                        .Cells(ProxLinha - 3, y).Value = repasse.Fields(i)
                        i = i + 1
                    Else
                        .Cells(ProxLinha - 3, y).Value = "0"
                    End If
                End If
                                        
                'Quantidade de parcelas até a data limite:
                .Cells(ProxLinha - 1, y).Formula = "=COUNTA(" & letraColuna & "14:" & letraColuna & (LinhaAteChaves - 1) & "," & letraColuna & (LinhaAteChaves + 1) & ":" & letraColuna & (LinhaPosChaves - 2) & ")"
                .Cells(ProxLinha - 1, y).Font.Italic = True
            Next
        End If
        
        'Seleciona a célula com a fórmula
        .Range(ConvertToLetter(dados.Fields.Count * 2 - 1) & ProxLinha - 4).Select
        
        XFO_ExlObj.CutCopyMode = False
        XFO_ExlObj.Selection.Copy
        
        .Range(ConvertToLetter(dados.Fields.Count * 2 + 2) & ProxLinha - 4).Select
        XFO_ExlObj.ActiveSheet.Paste
        .Range(ConvertToLetter(dados.Fields.Count * 2 + 3) & ProxLinha - 4).Select
        XFO_ExlObj.ActiveSheet.Paste
        
        letraColuna = ConvertToLetter(dados.Fields.Count * 2)
        .Cells(ProxLinha - 3, dados.Fields.Count * 2 + 2).Formula = "=SUMIF($C$13:$" & letraColuna & "$13,""=NOMINAL"",C" & ProxLinha - 3 & ":" & letraColuna & ProxLinha - 3 & ")"
        .Cells(ProxLinha - 3, dados.Fields.Count * 2 + 3).Formula = "=SUMIF($C$13:$" & letraColuna & "$13,""=NOMINAL"",C" & ProxLinha - 3 & ":" & letraColuna & ProxLinha - 3 & ")"

        
        'Coloca vazio nas células da coluna que separa o conjunto de valores previsto do primeiro conjunto de fórmulas ("Total apartamento")
        .Cells(ProxLinha, (dados.Fields.Count * 2) + 1).Value = ""
        .Cells(ProxLinha - 1, (dados.Fields.Count * 2) + 1).Value = ""
        .Cells(ProxLinha - 2, (dados.Fields.Count * 2) + 1).Value = ""
        .Cells(ProxLinha - 3, (dados.Fields.Count * 2) + 1).Value = ""
        .Cells(ProxLinha - 4, (dados.Fields.Count * 2) + 1).Value = ""
        .Cells(ProxLinha - 5, (dados.Fields.Count * 2) + 1).Value = ""
    End With
    
    AvancarProgresso 1
    
    cont = cont + 14
    letraColuna = ConvertToLetter(dados.Fields.Count * 2)
    
    
    'Insere o primeiro conjunto de fórmulas, a direita da listagem dos imóveis
    With XFO_ExlObj.Sheets(sheet)
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 2) & "14").Formula = "=SUMIF($C$13:$" & letraColuna & "$13,""=NOMINAL"",C14:" & letraColuna & "14)"
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 3) & "14").Formula = "=SUMIF($C$13:$" & letraColuna & "$13,""=VLR PRES"",C14:" & letraColuna & "14)"
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 4) & "14").Formula = "=(COUNT(C14:" & letraColuna & "14)-COUNTBLANK(C14:" & letraColuna & "14))/2"
        
        'Copia as fórmulas da primeira linha
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 2) & "14:" & ConvertToLetter((dados.Fields.Count * 2) + 4) & "14").Select
        XFO_ExlObj.CutCopyMode = False
        XFO_ExlObj.Selection.Copy
        
        'Cola as fórmulas nas células abaixo
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 2) & "15:" & ConvertToLetter((dados.Fields.Count * 2) + 4) & (LinhaAteChaves - 1)).Select
        .Paste
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 2) & (LinhaAteChaves + 1) & ":" & ConvertToLetter((dados.Fields.Count * 2) + 4) & (LinhaPosChaves - 1)).Select
        .Paste
        
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 4) & "14:" & ConvertToLetter((dados.Fields.Count * 2) + 4) & (LinhaPosChaves - 1)).HorizontalAlignment = xlHAlignCenter
    End With
    
    AvancarProgresso 1
    
    'Insere o segundo conjunto de fórmulas, a direita da listagem dos imóveis
    With XFO_ExlObj.Sheets(sheet)
        .Cells(16, ((dados.Fields.Count * 2) + 6)).Value = "TOTAL"
        
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 7) & "16").Formula = _
            "=SUMIF(A15:B" & (LinhaAteChaves - 1) & ", "">=  " & Format$(DtPeriodoCarencia.Value, "dd/mm/yy") & """, " & _
            ConvertToLetter((dados.Fields.Count * 2) + 3) & "15:" & ConvertToLetter((dados.Fields.Count * 2) + 3) & _
            (LinhaAteChaves - 1) & ")"
            
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 8) & "16").Formula = _
            "=SUMIF(A15:B" & (LinhaAteChaves - 1) & ", "">= " & Format$(DtPeriodoCarencia.Value, "dd/mm/yy") & """," & _
            ConvertToLetter((dados.Fields.Count * 2) + 4) & "15:" & ConvertToLetter((dados.Fields.Count * 2) + 4) & _
            (LinhaAteChaves - 1) & ")"
            
        .Cells(17, ((dados.Fields.Count * 2) + 6)).Value = "TOTAL"
        
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 7) & "17").Formula = _
            "=SUMIF(A" & (LinhaAteChaves + 1) & ":B" & (LinhaPosChaves - 1) & ", "" >= " & Format$(DtPeriodoCarencia.Value, "dd/mm/yy") & """," & _
            ConvertToLetter((dados.Fields.Count * 2) + 3) & (LinhaAteChaves + 1) & ":" & ConvertToLetter((dados.Fields.Count * 2) + 3) & _
            (LinhaPosChaves - 1) & ")"
            
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 8) & "17").Formula = _
             "=SUMIF(A" & (LinhaAteChaves + 1) & ":B" & (LinhaPosChaves - 1) & ", "">= " & Format$(DtPeriodoCarencia.Value, "dd/mm/yy") & """," & _
            ConvertToLetter((dados.Fields.Count * 2) + 4) & (LinhaAteChaves + 1) & ":" & ConvertToLetter((dados.Fields.Count * 2) + 4) & _
            (LinhaPosChaves - 1) & ")"
            
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 6) & "16:" & ConvertToLetter((dados.Fields.Count * 2) + 8) & "17").HorizontalAlignment = xlHAlignCenter
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 6) & "16:" & ConvertToLetter((dados.Fields.Count * 2) + 8) & "17").Font.Bold = True
    End With
    
    AvancarProgresso 1
        
    
    'Coloca as células no formato decimal
    XFO_ExlObj.Sheets(sheet).Range("C14:" & ConvertToLetter((dados.Fields.Count * 2) + 3) & (LinhaPosChaves + 1)).NumberFormat = "#,##0.00"
    XFO_ExlObj.Sheets(sheet).Range("C" & (LinhaPosChaves + 3) & ":" & ConvertToLetter((dados.Fields.Count * 2) + 3) & (LinhaPosChaves + 4)).NumberFormat = "#,##0.00"
    XFO_ExlObj.Sheets(sheet).Range("C" & ProxLinha & ":" & ConvertToLetter((dados.Fields.Count * 2) + 3) & ProxLinha).NumberFormat = "#,##0.00"
    XFO_ExlObj.Sheets(sheet).Range(letraColuna & ProxLinha & ":" & letraColuna & ProxLinha).NumberFormat = "#,##0.00"
    XFO_ExlObj.Sheets(sheet).Range(ConvertToLetter((dados.Fields.Count * 2) + 7) & ProxLinha).NumberFormat = "#,##0.00"
    
    
    'Coloca as datas em negrito:
    XFO_ExlObj.Sheets(sheet).Range("A14:A" & cont).Font.Bold = True
        
    
    AvancarProgresso 1
    
        
    'Insere o campo com os dados para envio de correspondência
    With XFO_ExlObj.Sheets(sheet)
        ProxLinha = ProxLinha + 5
        
        .Range("B" & ProxLinha & ":H" & ProxLinha + 10).Interior.ColorIndex = 36
        .Range("B" & ProxLinha & ":H" & ProxLinha).Merge
        
        'Insere bordas na planilha
        .Range("B" & ProxLinha & ":H" & ProxLinha).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("B" & ProxLinha & ":H" & ProxLinha).Borders(xlEdgeTop).Weight = xlMedium
        .Range("B" & ProxLinha & ":H" & ProxLinha).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        
        .Range("B" & ProxLinha + 10 & ":H" & ProxLinha + 10).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B" & ProxLinha + 10 & ":H" & ProxLinha + 10).Borders(xlEdgeBottom).Weight = xlMedium
        .Range("B" & ProxLinha + 10 & ":H" & ProxLinha + 10).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        
        .Range("B" & ProxLinha & ":B" & ProxLinha + 10).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("B" & ProxLinha & ":B" & ProxLinha + 10).Borders(xlEdgeLeft).Weight = xlMedium
        .Range("B" & ProxLinha & ":B" & ProxLinha + 10).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        
        .Range("H" & ProxLinha & ":H" & ProxLinha + 10).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("H" & ProxLinha & ":H" & ProxLinha + 10).Borders(xlEdgeRight).Weight = xlMedium
        .Range("H" & ProxLinha & ":H" & ProxLinha + 10).Borders(xlEdgeRight).ColorIndex = xlAutomatic
                    
        .Cells(ProxLinha, 2).Value = "ENDEREÇO PARA ENVIO:"
        .Cells(ProxLinha, 2).HorizontalAlignment = xlHAlignCenter
        .Cells(ProxLinha, 2).Font.Color = vbRed
        .Cells(ProxLinha, 2).Font.Bold = True
        
        ProxLinha = ProxLinha + 2
        .Cells(ProxLinha, 2).Value = "Enviar através de e-mail para:"
        .Cells(ProxLinha, 2).Font.Color = vbRed
        .Cells(ProxLinha, 2).Font.Bold = True
                
        ProxLinha = ProxLinha + 1
        .Cells(ProxLinha, 2).Value = "Email:"
        .Cells(ProxLinha, 2).Font.Color = vbRed
        .Cells(ProxLinha, 2).Font.Bold = True
        .Cells(ProxLinha, 3).Value = "ruaraujo@santanderbanespa.com.br"
        .Cells(ProxLinha, 3).Font.Color = vbBlue
        .Cells(ProxLinha, 3).Font.Underline = True
        .Cells(ProxLinha, 3).Font.Bold = True
        
        ProxLinha = ProxLinha + 1
        .Cells(ProxLinha, 2).Value = "Endereço para entrega da via assinada:"
        .Cells(ProxLinha, 2).Font.Color = vbRed
        .Cells(ProxLinha, 2).Font.Bold = True
        
        ProxLinha = ProxLinha + 1
        .Cells(ProxLinha, 2).Value = "CASA 2"
        .Cells(ProxLinha, 2).Font.Color = vbBlue
        .Cells(ProxLinha, 2).Font.Bold = True
        
        ProxLinha = ProxLinha + 1
        .Cells(ProxLinha, 2).Value = "Av. Guido Caloi, 1000 - 3° andar - Bloco 1"
        .Cells(ProxLinha, 2).Font.Color = vbBlue
        .Cells(ProxLinha, 2).Font.Bold = True
        
        ProxLinha = ProxLinha + 1
        .Cells(ProxLinha, 2).Value = "Superintendência de Serviços a Clientes - Crédito Imobiliário."
        .Cells(ProxLinha, 2).Font.Color = vbBlue
        .Cells(ProxLinha, 2).Font.Bold = True
    End With
    
    
    AvancarProgresso 1
    
    
    'Insere campo para assinatura
    With XFO_ExlObj.Sheets(sheet)
        .Cells(LinhaPosChaves + 1, (dados.Fields.Count * 2) + 7).Value = "Declaramos de que as informações aqui prestadas refletem fielmente"
        .Cells(LinhaPosChaves + 2, (dados.Fields.Count * 2) + 7).Value = "a situação atual dos contratos de venda das unidades relacionadas."
        .Cells(LinhaPosChaves + 3, (dados.Fields.Count * 2) + 10).Value = "Salvador, 16 de junho de 2008."
        .Cells(LinhaPosChaves + 3, (dados.Fields.Count * 2) + 10).HorizontalAlignment = xlHAlignRight
                
        'Nome do responsável
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 8) & LinhaPosChaves + 7 & ":" & ConvertToLetter((dados.Fields.Count * 2) + 10) & LinhaPosChaves + 7).Merge
        .Range(ConvertToLetter((dados.Fields.Count * 2) + 8) & LinhaPosChaves + 7 & ":" & ConvertToLetter((dados.Fields.Count * 2) + 10) & LinhaPosChaves + 7).Borders(xlEdgeTop).LineStyle = xlContinuous
        
        If (TxtNomeResponsavel.Text <> "") Then .Cells(LinhaPosChaves + 7, (dados.Fields.Count * 2) + 8).Value = TxtNomeResponsavel.Text
        If (TxtNomeResponsavel.Text = "") Then .Cells(LinhaPosChaves + 7, (dados.Fields.Count * 2) + 8).Value = "Responsável"
                
        .Cells(LinhaPosChaves + 7, (dados.Fields.Count * 2) + 8).HorizontalAlignment = xlHAlignCenter
    End With
    
    AvancarProgresso 1
    
    
    'Formata para não exibir as linhas de grade da planilha
    XFO_ExlObj.Sheets(sheet).Select
    XFO_ExlObj.ActiveWindow.DisplayGridlines = False
    
    
    'Ativa o cálculo automático das fórmulas
    XFO_ExlObj.Application.Calculate
    XFO_ExlObj.Application.Calculation = xlCalculationAutomatic
    XFO_ExlObj.Application.CalculateBeforeSave = True
    
    
    'Ajusta o tamanho das colunas para exibir o conteúdo
    XFO_ExlObj.Columns("E:" & ConvertToLetter(dados.Fields.Count * 2)).EntireColumn.AutoFit
       
    AvancarProgresso 1
    
    '--------------------------------------------
    '--------------------------------------------
    'Define área para impressão
    With XFO_ExlObj.Sheets(sheet)
        .PageSetup.PrintTitleRows = ""
        .PageSetup.PrintTitleColumns = "$A:$B"
        .PageSetup.PrintArea = XFO_ExlObj.Sheets(sheet).Cells(1, (dados.Fields.Count * 2) + 2).Address & ":" & XFO_ExlObj.Sheets(sheet).Cells(LinhaPosChaves + 7, (dados.Fields.Count * 2) + 12).Address
    End With
    
    With XFO_ExlObj.Sheets(sheet).PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.62992125984252)
        .RightMargin = Application.InchesToPoints(0.393700787401575)
        .TopMargin = Application.InchesToPoints(0.551181102362205)
        .BottomMargin = Application.InchesToPoints(0.984251968503937)
        .HeaderMargin = Application.InchesToPoints(0.511811023622047)
        .FooterMargin = Application.InchesToPoints(0.511811023622047)
        .PrintHeadings = False
        .PrintGridlines = False
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    
    
    'Seleciona a primeira célula de valores previstos
    XFO_ExlObj.Range("C14").Select
    
    
    AvancarProgresso 1
       
End Sub
Private Sub CarregabaFormaPreenchimento()

    'Preenchendo as células
    With XFO_ExlObj.ActiveSheet
        
        .Name = "FORMA DE PREENCHIMENTO"
        
        'Ajuste da largura das colunas
        .Range("A10").ColumnWidth = 8.43
        .Range("B10").ColumnWidth = 8.43
        .Range("C10").ColumnWidth = 8.43
        .Range("D10").ColumnWidth = 9
        .Range("E10").ColumnWidth = 9
        .Range("F10").ColumnWidth = 9
        .Range("G10").ColumnWidth = 8.5
        .Range("H10").ColumnWidth = 8.5
        .Range("I10").ColumnWidth = 8.5
        .Range("J10").ColumnWidth = 10.29
              
        '-----------
        .Range("A1:J1").Merge
        .Cells(1, 1).Value = "INSTRUÇÕES PARA O PREENCHIMENTO DO FLUXO DE RECEBÍVEIS"
        .Cells(1, 1).HorizontalAlignment = xlHAlignCenter
        .Cells(1, 1).Font.Bold = True
                            
        '-----------
        .Range("A2:C2").Merge
        .Cells(2, 1).Value = "CAMPO"
        .Cells(2, 1).HorizontalAlignment = xlHAlignCenter
        .Cells(2, 1).Font.Bold = True
    
        '-----------
        .Range("D2:J2").Merge
        .Cells(2, 4).Value = "FORMA DE PREENCHIMENTO"
        .Cells(2, 4).HorizontalAlignment = xlHAlignCenter
        .Cells(2, 4).Font.Bold = True
                         
        '-----------
        .Range("A3:C3").Merge
        .Cells(3, 1).Value = "MOEDA"
        .Cells(3, 1).HorizontalAlignment = xlHAlignCenter
        .Cells(3, 4).Value = "Os valores devem estar em Real. Este campo já está preenchido."
                         
        '-----------
        .Range("A4:C4").Merge
        .Cells(4, 1).Value = "TAXA VLR PRESENTE"
        .Cells(4, 1).HorizontalAlignment = xlHAlignCenter
        .Cells(4, 4).Value = "Deverá ser a mesma taxa de juros anual do contrato de Plano Empresário"
                         
        '-----------
        .Range("A5:C9").Merge
        .Cells(5, 1).Value = "DATA CÁLCULO"
        .Cells(5, 1).VerticalAlignment = xlVAlignCenter
        .Cells(5, 1).HorizontalAlignment = xlHAlignCenter
                      
        .Range("D5:J9").Merge
        .Cells(5, 4).Value = "É a data base da planilha. Em cada mês, deverá corresponder ao mesmo dia da " & _
                            "assinatura do contrato de Plano Empresário  (p. ex.: se o contrato de Plano Em- " & _
                            "presário  foi  assinado  em 01/03/2001,  a data base da planilha nos meses sub- " & _
                            "sequentes será:  em abril/2001 = 01/04/2001;  em maio/2001 = 01/05/2001;  em " & _
                            "junho/2001 = 01/06/2001, e assim sucessivamente). "
                            
        .Cells(5, 4).VerticalAlignment = xlVAlignTop
        .Cells(5, 4).WrapText = True
                                                   
        '-----------
        .Range("A10:C24").Merge
        .Cells(10, 1).Value = "DATAS DE VENCIMENTO"
        .Cells(10, 1).VerticalAlignment = xlVAlignCenter
        .Cells(10, 1).HorizontalAlignment = xlHAlignCenter
                      
        .Range("D10:J24").Merge
        .Cells(10, 4).Value = "*Nesta coluna deverão constar todas as futuras datas de vencimento das parce- " & _
                                "las (mensais, intermediárias, …) do valor de venda das unidades já comercializa- " & _
                                "das. " & Chr(10) & _
                                "* Deverão estar em ordem cronológica crescente. " & _
                                "* Se para uma mesma unidade  houverem duas ou  mais  parcelas  vencíveis na " & _
                                " mesma data,  esta data deverá constar em tantas linhas  sucessivas da planilha " & _
                                "quanto forem as parcelas com vencimento coincidente. " & Chr(10) & _
                                "* As datas de vencimento anteriores à data prevista para a conclusão da obra (li- " & _
                                "beração da última parcela de obra) deverão ser relacionadas acima da linha 'TO- " & _
                                "TAL ATÉ CHAVES'" & Chr(10) & _
                                "* As datas de vencimento até o prazo de 30 meses após a conclusão da obra (li- " & _
                                "beração da última parcela de obra) deverão ser relacionadas abaixo da linha 'TO- " & _
                                "TAL ATÉ CHAVES' " & Chr(10) & _
                                "* Eventuais datas de vencimento posteriores ao prazo acima não deverão ser re- " & _
                                "lacionadas na planilha. "
    
        .Cells(10, 4).VerticalAlignment = xlVAlignTop
        .Cells(10, 4).WrapText = True
        .Cells(10, 4).HorizontalAlignment = xlHAlignLeft
                       
        '-----------
        .Range("A25:C26").Merge
        .Cells(25, 1).Value = "VAGA N°"
        .Cells(25, 1).VerticalAlignment = xlVAlignCenter
        .Cells(25, 1).HorizontalAlignment = xlHAlignCenter
                      
        .Range("D25:J26").Merge
        .Cells(25, 4).Value = "* Caso as vagas de garagem sejam autônomas, elas também deverão ser identi- " & _
                             "ficadas acima da série de recebíveis correspondente, nos campos apropriados. "
        
        .Cells(25, 4).VerticalAlignment = xlVAlignTop
        .Cells(25, 4).WrapText = True
        .Cells(25, 4).HorizontalAlignment = xlHAlignLeft
        
        '--------
        .Range("A27:C41").Merge
        .Cells(27, 1).Value = "NOMINAL " & Chr(10) & "(colunas)"
        .Cells(27, 1).VerticalAlignment = xlVAlignCenter
        .Cells(27, 1).HorizontalAlignment = xlHAlignCenter
                      
        .Range("D27:J41").Merge
        .Cells(27, 4).Value = "* Para cada uma das unidades e vagas autônomas, deverão ser relacionadas to- " & _
                                "das as parcelas a receber (excluindo-se  as  já  vencidas) do respectivo valor de " & _
                                "venda, nas linhas correspondentes às respectivas datas de vencimento. " & Chr(10) & _
                                "* Deverão constar os valores nominais devidamente atualizados monetariamente " & _
                                "até a data base da planilha. " & Chr(10) & _
                                "* Caso duas ou mais parcelas de uma mesma unidade tenham a mesma data de " & _
                                "vencimento, todas elas deverão constar da planilha, uma em cada linha. " & Chr(10) & _
                                "* As  parcelas que tiverem seu vencimento até a data  prevista  para a conclusão " & _
                                "da obra  ( liberação da última parcela de obra )  deverão ser relacionadas, uma a " & _
                                "uma, acima da linha 'TOTAL ATÉ CHAVES'. " & Chr(10) & _
                                "* Abaixo dessa linha, deverão ser relacionadas, uma a uma, as parcela que tive- " & _
                                "rem seu vencimento até o prazo de 30 meses após aquela data. " & Chr(10) & _
                                "* As parcelas que tiverem seu vencimento após esse prazo  deverão  ser  soma- " & _
                                "das,  e o total deverá constar na tabela  imediatamente  acima da linha  'TOTAL " & _
                                "PÓS CHAVES' "
                      
        .Cells(27, 4).VerticalAlignment = xlVAlignTop
        .Cells(27, 4).WrapText = True
        .Cells(27, 4).HorizontalAlignment = xlHAlignLeft
         
        '-----------
        .Range("A42:C46").Merge
        .Cells(42, 1).Value = "REPASSE DATA" & Chr(10) & "(linha)"
        .Cells(42, 1).VerticalAlignment = xlVAlignCenter
        .Cells(42, 1).HorizontalAlignment = xlHAlignCenter
                      
        .Range("D42:J46").Merge
        .Cells(42, 4).Value = "* Para cada unidade e vaga autônoma que tiver previsão de repasse em seu con- " & _
                                " trato de venda, na coluna correspondente aos valores nominais, deverá constar a " & _
                                "data prevista para o repasse da mesma. " & Chr(10) & _
                                "* Para as unidades e vagas autônomas que não tiverem  previsão de repasse em " & _
                                "seu contrato de venda, este campo deverá ser deixado em branco "
                      
        .Cells(42, 4).VerticalAlignment = xlVAlignTop
        .Cells(42, 4).WrapText = True
        .Cells(42, 4).HorizontalAlignment = xlHAlignLeft
         
        '-------------------
        .Range("A47:C51").Merge
        .Cells(47, 1).Value = "REPASSE VALOR" & Chr(10) & "(linha)"
        .Cells(47, 1).VerticalAlignment = xlVAlignCenter
        .Cells(47, 1).HorizontalAlignment = xlHAlignCenter
                      
        .Range("D47:J51").Merge
        .Cells(47, 4).Value = "* Para cada unidade e vaga autônoma que tiver previsão de repasse em seu con- " & _
                               "trato de venda, na coluna correspondente aos valores nominais, deverá constar o " & _
                               "valor previsto para ser repassado da mesma. " & Chr(10) & _
                               "* Para as unidades e vagas autônomas que não tiverem  previsão de repasse em " & _
                               "seu contrato de venda, este campo deverá ser deixado em branco "
        
        .Cells(47, 4).VerticalAlignment = xlVAlignTop
        .Cells(47, 4).WrapText = True
        .Cells(47, 4).HorizontalAlignment = xlHAlignLeft
        
        '----------------
        .Cells(53, 1).Value = "OBSERVAÇÕES IMPORTANTES:"
        .Cells(53, 1).Font.Bold = True
        
        .Range("A55:K58").Merge
        .Cells(55, 1).WrapText = True
        .Cells(55, 1).Value = "1-) Esta planilha deverá ser preenchida mensalmente pela empresa e enviada,  via e-mail,  para a área de Crédito " & Chr(10) & _
                                "     Imobiliário do Grupo Santander Banespa com pelo menos 5 dias úteis antes da data prevista para a Liberação da  " & Chr(10) & _
                                "     Parcela da obra, ficando assim retida a Liberação da Parcela da obra caso não seja respeitado o prazo de entrega do " & Chr(10) & _
                                "     acima solicitado. "
                                
        .Range("A59:J61").Merge
        .Cells(59, 1).WrapText = True
        .Cells(59, 1).Value = "2-) Mensalmente,  após a atualização da planilha, deverá ser impresso um resumo da mesma,  onde constem as " & Chr(10) & _
                                "     colunas A, B, FH, FI , FJ, FK,FL,FM e FN. Este resumo impresso,  devidamente assinado " & Chr(10) & _
                                "     e vistado em todas as páginas, deverá ser enviado para a área de Crédito Imobiliário do Grupo Santander Banespa."
                  
        .Range("A62:J64").Merge
        .Cells(62, 1).WrapText = True
        .Cells(62, 1).Value = "3-) A planilha do fluxo de recebíveis deverá espelhar,  fielmente,  os contratos de venda das unidades.  Eventuais " & Chr(10) & _
                            "     divergências serão apuradas e deverão ser justificadas pela empresa,  sob pena de ser retida a liberação das " & Chr(10) & _
                            "     parcelas de obra. "
                            
        .Range("A65:J72").Merge
        .Cells(65, 1).WrapText = True
        .Cells(65, 1).Value = "4-) Caso seja necessário,  poderá ser feita a inserção de mais linhas na planilha modelo para comportar todas as " & Chr(10) & _
                                "     datas de vencimento dos recebíveis, desde que observado o seguinte: " & Chr(10) & _
                                "            * NÃO poderá ser inserida nenhuma linha nos seguintes intervalos da planilha modelo: " & Chr(10) & _
                                "                     > entre as linhas 1 e 15; " & Chr(10) & _
                                "                     > entre as linhas 48 e 52; " & Chr(10) & _
                                "                     > entre as linhas 79 e 89 " & Chr(10) & _
                                "            * Após a inserção de cada linha,  deverão ser copiadas as fórmulas da linha imediatamente superior à " & Chr(10) & _
                                "              linha que foi inserida. "

        .Range("B76:G76").Merge
        .Range("B76:G85").Interior.ColorIndex = 36
        .Cells(76, 2).Value = "ENDEREÇO PARA ENVIO:"
        .Cells(76, 2).HorizontalAlignment = xlHAlignCenter
        .Cells(76, 2).Font.Color = vbRed
        .Cells(76, 2).Font.Bold = True
        
        .Cells(78, 2).Value = "Enviar através de e-mail para:"
        .Cells(78, 2).Font.Color = vbRed
        .Cells(78, 2).Font.Bold = True
        .Cells(79, 2).Value = "Email:"
        .Cells(79, 2).Font.Color = vbRed
        .Cells(79, 2).Font.Bold = True
        
        .Cells(79, 3).Value = "ruaraujo@santanderbanespa.com.br"
        .Cells(79, 3).Font.Color = vbBlue
        .Cells(79, 3).Font.Underline = True
        .Cells(79, 3).Font.Bold = True
        
        .Cells(81, 2).Value = "Endereço para entrega da via assinada:"
        .Cells(81, 2).Font.Color = vbRed
        .Cells(81, 2).Font.Bold = True
        .Cells(82, 2).Value = "CASA 2"
        .Cells(83, 2).Value = "Av. Guido Caloi, 1000 - 3° andar - Bloco 1"
        .Cells(84, 2).Value = "Superintendência de Serviços a Clientes - Crédito Imobiliário."
        .Cells(82, 2).Font.Color = vbBlue
        .Cells(83, 2).Font.Color = vbBlue
        .Cells(84, 2).Font.Color = vbBlue
        .Cells(82, 2).Font.Bold = True
        .Cells(83, 2).Font.Bold = True
        .Cells(84, 2).Font.Bold = True
                
    End With
    
    'Coloca em destaque (cor vermelha e negrito) uma parte do texto:
    XFO_ExlObj.ActiveSheet.Range("A55").Select
    With XFO_ExlObj.ActiveCell.Characters(174, 56).Font
        .Bold = True
        .ColorIndex = 3
    End With
        
    'Insere bordas na planilha
    With XFO_ExlObj.Range("B76:G76").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.Range("B85:G85").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.Range("B76:b85").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.Range("G76:G85").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
          
    With XFO_ExlObj.Range("A1:J51").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.Range("A1:J51").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.Range("A1:J1").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.Range("A51:J51").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.Range("A1:J51").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.Range("A1:J51").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
    End With
        
    'Formata a página
    With XFO_ExlObj.ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.64)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.54)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .PrintTitleRows = "$1:$12" 'Repete o cabeçalho em todas as páginas
        .Zoom = 65
    End With
    
    'Formata para não exibir as linhas de grade da planilha
    XFO_ExlObj.ActiveWindow.DisplayGridlines = False
End Sub

Function ConvertToLetter(ByVal valor As Integer) As String
    Dim PrimeiraLetra As Integer
    Dim SegundaLetra As Integer
    
    If ((valor / 26) = (valor \ 26)) Then
        PrimeiraLetra = valor \ 26 - 1
    Else
        PrimeiraLetra = valor \ 26
    End If
    
    SegundaLetra = valor - (PrimeiraLetra * 26)
   
    If PrimeiraLetra > 0 Then
        ConvertToLetter = Chr(PrimeiraLetra + 64)
    End If
    
    If SegundaLetra > 0 Then
        ConvertToLetter = ConvertToLetter & Chr(SegundaLetra + 64)
    End If
End Function


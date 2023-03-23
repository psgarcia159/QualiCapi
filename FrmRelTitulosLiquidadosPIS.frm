VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRelTitulosLiquidadosPIS
   Caption         =   "Títulos Liquidados PIS/COFINS"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   8235
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   7050
      TabIndex        =   4
      Top             =   8040
      Width           =   1035
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "&Exportar"
      Height          =   345
      Left            =   5880
      TabIndex        =   3
      Top             =   8040
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   7860
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   7950
      Begin VB.CheckBox ChkMesesComRecebimento 
         Caption         =   "Exportar apenas meses com recebimento"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   7480
         Width           =   3735
      End
      Begin VB.Frame Frame4 
         Caption         =   "Empresas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   200
         TabIndex        =   11
         Top             =   1000
         Width           =   7550
         Begin VB.CommandButton CmdLimparTodos 
            Caption         =   "&Limpar Todos"
            Height          =   375
            Left            =   1605
            TabIndex        =   13
            Top             =   5000
            Width           =   1275
         End
         Begin VB.CommandButton CmdMarcarTodos 
            Caption         =   "&Marcar Todos"
            Height          =   375
            Left            =   195
            TabIndex        =   12
            Top             =   5000
            Width           =   1275
         End
         Begin MSComctlLib.TreeView TvwEmpresas 
            Height          =   4515
            Left            =   195
            TabIndex        =   14
            Top             =   360
            Width           =   7170
            _ExtentX        =   12647
            _ExtentY        =   7964
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
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
      End
      Begin VB.Frame Frame3 
         Caption         =   "Exportação"
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
         Left            =   200
         TabIndex        =   7
         Top             =   200
         Width           =   7550
         Begin VB.OptionButton OptCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   6200
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton OptEmpreendimento 
            Caption         =   "Empreendimento"
            Height          =   195
            Left            =   3000
            TabIndex        =   9
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton OptEmpresa 
            Caption         =   "Empresa"
            Height          =   195
            Left            =   600
            TabIndex        =   8
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame FraPagamento 
         Caption         =   "Período de Pagamento"
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
         Left            =   200
         TabIndex        =   5
         Top             =   6615
         Width           =   3105
         Begin MSComCtl2.DTPicker DtpInicioPagamento 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   300
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   63504387
            CurrentDate     =   37617.6519560185
         End
         Begin MSComCtl2.DTPicker DtpFimPagamento 
            Height          =   315
            Left            =   1680
            TabIndex        =   2
            Top             =   300
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MM/yy"
            DateIsNull      =   -1  'True
            Format          =   63504387
            CurrentDate     =   37617.6517939815
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "a"
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
            Left            =   1400
            TabIndex        =   6
            Top             =   360
            Width           =   180
         End
      End
      Begin Threed.SSFrame FraStatus 
         Height          =   765
         Left            =   3360
         TabIndex        =   16
         Top             =   6615
         Width           =   4380
         _Version        =   65536
         _ExtentX        =   7726
         _ExtentY        =   1349
         _StockProps     =   14
         Caption         =   "Status do Contrato"
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
            Left            =   1740
            TabIndex        =   19
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
            Left            =   360
            TabIndex        =   18
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
            Left            =   3000
            TabIndex        =   17
            Top             =   300
            Width           =   975
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Empreendimentos:"
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
         Left            =   195
         TabIndex        =   15
         Top             =   1395
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc DatEmpreendimento 
      Height          =   330
      Left            =   120
      Top             =   8040
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
   Begin MSComDlg.CommonDialog dlgSalvar 
      Left            =   2400
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatContrato 
      Height          =   330
      Left            =   2160
      Top             =   8040
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
      Caption         =   "DatContrato"
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
   Begin MSAdodcLib.Adodc DatEmpresa 
      Height          =   330
      Left            =   3720
      Top             =   8040
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
      Caption         =   "DatEmpresa"
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
Attribute VB_Name = "FrmRelTitulosLiquidadosPis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private XFO_ExlObj              As Excel.Application    'Create excel object
Public XFI_QtdEmpreendimento    As Integer              'Quantidade de empreendimentos
Private XLT_CONTRATOS           As String

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

 
Private Sub CmdLimparTodos_Click()
    Dim i As Integer
    i = 1
    
    While i < TvwEmpresas.Nodes.Count
        TvwEmpresas.Nodes.Item(i).Checked = False
    i = i + 1
    Wend
End Sub

Private Sub CmdMarcarTodos_Click()
    Dim i As Integer
    i = 1
    
    While i < TvwEmpresas.Nodes.Count
        TvwEmpresas.Nodes.Item(i).Checked = True
    i = i + 1
    Wend

End Sub

Private Sub Form_Load()

    subManutencaoJanelasAtivas "I", "FrmRelTitulosLiquidadosPIS"
    subConectarControleDadosNV DatEmpreendimento, "SELECT * FROM Empreendimentos ORDER BY empd_cd_empreendimento", Estatico
    subConectarControleDadosNV DatContrato, "SELECT * FROM ConsCAPContratos ORDER BY Contrato", Estatico
    subConectarControleDadosNV DatEmpresa, "SELECT * FROM Empresas ORDER BY empr_tx_RazaoSocial", Estatico
        
    XFI_QtdEmpreendimento = 0
    PreencheTreeViewEmpresas
    
End Sub


Private Sub CmdExportar_Click()
    
    Dim rs                  As New ADODB.Recordset
    Dim XLO_COMANDO         As New ADODB.Command
    Dim XLI_EMPRESA         As Integer 'Código da empresa
    Dim i                   As Integer
    Dim sheet               As Integer 'Guarda o índice da planilha
    Dim XLT_NOMEEMPRESA     As String 'Nome da empresa
    Dim XLT_EMPREENDIMENTO  As String 'Nome do empreendimento
    Dim XLT_CLIENTE         As String 'Nome do cliente
    Dim XLI_STATUS As Integer
    Dim XLT_FILTRO As String
    
    XLT_CONTRATOS = ""
    i = 1
    sheet = 1
    
    On Error GoTo ErrHandler
    
    If IsNull(DtpInicioPagamento.Value) Or IsNull(DtpFimPagamento.Value) Then
        MsgBox "É obrigatório informar o período de pagamento", vbInformation, "INFORMAÇÃO"
    End If
        
    'Caixa de diálogo para selecionar local para armazenar a planilha
    dlgSalvar.CancelError = True
                              
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
    
        FrmRelTitulosLiquidadosPis.MousePointer = vbHourglass
            
        'Inicializa o objeto excel
        Set XFO_ExlObj = CreateObject("excel.application")
        
        'Adiciona um WorkBook
        XFO_ExlObj.Workbooks.Add
        
        'Status
        XLI_STATUS = 0
        XLT_FILTRO = ""
        
        If ChkCedido.Value = 1 Then XLI_STATUS = XLI_STATUS + 1
        If ChkInativo.Value = 1 Then XLI_STATUS = XLI_STATUS + 2
        If ChkAtivo.Value = 1 Then XLI_STATUS = XLI_STATUS + 4
        
        'Tipo do Status: Ativo
        If (XLI_STATUS And 4) = 4 Then
            XLT_FILTRO = XLT_FILTRO + "A;"
            XLI_STATUS = XLI_STATUS - 4
        End If
        
        'Tipo do Status: Inativo
        If (XLI_STATUS And 2) = 2 Then
            XLT_FILTRO = XLT_FILTRO + "I;"
            XLI_STATUS = XLI_STATUS - 2
            
        End If
        
        'Tipo do Status: Cedido
        If (XLI_STATUS And 1) = 1 Then
            XLT_FILTRO = XLT_FILTRO + "C;"
            XLI_STATUS = XLI_STATUS - 1
        End If
        
        If Len(XLT_FILTRO) > 0 Then
            XLT_FILTRO = Mid(XLT_FILTRO, 1, Len(XLT_FILTRO) - 1)
        End If

        While i < TvwEmpresas.Nodes.Count
           
           'Empresa
           If TvwEmpresas.Nodes.Item(i).Tag = "Nó 1" And TvwEmpresas.Nodes.Item(i).Checked = True And OptEmpresa = True Then
                                
               XLI_EMPRESA = CInt(Mid(TvwEmpresas.Nodes.Item(i).Text, 1, 3))
               XLT_CONTRATOS = ""
               GetContratos TvwEmpresas.Nodes.Item(i)
               
                       
               If Len(XLT_CONTRATOS) > 0 Then
                   XLT_CONTRATOS = Mid(XLT_CONTRATOS, 1, Len(XLT_CONTRATOS) - 1)
               End If
                   
               With XLO_COMANDO
                   .CommandTimeout = 10000000
                   .CommandText = "spCAPI_sel_TitulosLiquidadosPISCOFINS"
                   .CommandType = adCmdStoredProc
                   Set XLO_COMANDO.ActiveConnection = Conexao
                   'Cria automaticamente todos os parâmetros dentro do objeto Comando
                   .Parameters.Refresh
               End With
               
               'Passa os parâmetros para a procedure
               XLO_COMANDO.Parameters(1).Value = XLI_EMPRESA
               XLO_COMANDO.Parameters(2).Value = DtpInicioPagamento.Value
               XLO_COMANDO.Parameters(3).Value = DtpFimPagamento.Value
               XLO_COMANDO.Parameters(4).Value = IIf(Len(XLT_CONTRATOS) > 0, XLT_CONTRATOS, Null)
               XLO_COMANDO.Parameters(5).Value = True
               XLO_COMANDO.Parameters(6).Value = IIf(Len(XLT_FILTRO) > 0, XLT_FILTRO, Null)
               XLO_COMANDO.Parameters(7).Value = ChkMesesComRecebimento.Value
               
               'Executa a procedure e guarda o resultado no objeto rs
               Set rs = XLO_COMANDO.Execute
           
               If Not rs.EOF Then
                   If sheet > 3 Then
                   XFO_ExlObj.Sheets.Add
                       'Move a planilha criada para a nova aba
                       XFO_ExlObj.ActiveSheet.Move After:=XFO_ExlObj.Sheets(XFO_ExlObj.ActiveWorkbook.Sheets.Count)
                   End If
                   'Função para criar a planilha
                   ExportaExcelEmpresaEmpreendimento rs, sheet, Mid(TvwEmpresas.Nodes.Item(i).Text, 7, Len(TvwEmpresas.Nodes.Item(i).Text) - 6), "RECEBIMENTOS - TODOS OS EMPREENDIMENTOS"
           
                   sheet = sheet + 1
                   
               End If
               
               rs.Close
            
            'Empreendimento
            ElseIf TvwEmpresas.Nodes.Item(i).Tag = "Nó 2" And TvwEmpresas.Nodes.Item(i).Checked = True And OptEmpreendimento = True Then
               
               XLI_EMPRESA = CInt(Mid(TvwEmpresas.Nodes.Item(i).Parent.Text, 1, 3))
               XLT_NOMEEMPRESA = Mid(TvwEmpresas.Nodes.Item(i).Parent.Text, 7, Len(TvwEmpresas.Nodes.Item(i).Parent.Text) - 6)
               XLT_EMPREENDIMENTO = TvwEmpresas.Nodes.Item(i).Text
               
               XLT_CONTRATOS = ""
               GetContratos TvwEmpresas.Nodes.Item(i)
                       
               If Len(XLT_CONTRATOS) > 0 Then
                   XLT_CONTRATOS = Mid(XLT_CONTRATOS, 1, Len(XLT_CONTRATOS) - 1)
               End If
                   
               With XLO_COMANDO
                   .CommandTimeout = 10000000
                   .CommandText = "spCAPI_sel_TitulosLiquidadosPISCOFINS"
                   .CommandType = adCmdStoredProc
                   Set XLO_COMANDO.ActiveConnection = Conexao
                   'Cria automaticamente todos os parâmetros dentro do objeto Comando
                   .Parameters.Refresh
               End With
               
               'Passa os parâmetros para a procedure
               XLO_COMANDO.Parameters(1).Value = XLI_EMPRESA
               XLO_COMANDO.Parameters(2).Value = DtpInicioPagamento.Value
               XLO_COMANDO.Parameters(3).Value = DtpFimPagamento.Value
               XLO_COMANDO.Parameters(4).Value = IIf(Len(XLT_CONTRATOS) > 0, XLT_CONTRATOS, Null)
               XLO_COMANDO.Parameters(6).Value = IIf(Len(XLT_FILTRO) > 0, XLT_FILTRO, Null)
               XLO_COMANDO.Parameters(7).Value = ChkMesesComRecebimento.Value
               
               'Executa a procedure e guarda o resultado no objeto rs
               Set rs = XLO_COMANDO.Execute
           
               If Not rs.EOF Then
                   If sheet > 3 Then
                       XFO_ExlObj.Sheets.Add
                       'Move a planilha criada para a nova aba
                       XFO_ExlObj.ActiveSheet.Move After:=XFO_ExlObj.Sheets(XFO_ExlObj.ActiveWorkbook.Sheets.Count)
                   End If
                   'Função para criar a planilha
                   ExportaExcelEmpresaEmpreendimento rs, sheet, XLT_NOMEEMPRESA, XLT_EMPREENDIMENTO
           
                   sheet = sheet + 1
                   
               End If
            
              rs.Close
              
            'Cliente
            ElseIf TvwEmpresas.Nodes.Item(i).Tag = "Folha" And TvwEmpresas.Nodes.Item(i).Checked = True And OptCliente = True Then
            
               XLI_EMPRESA = CInt(Mid(TvwEmpresas.Nodes.Item(i).Parent.Parent.Text, 1, 3))
               XLT_NOMEEMPRESA = Mid(TvwEmpresas.Nodes.Item(i).Parent.Parent.Text, 7, Len(TvwEmpresas.Nodes.Item(i).Parent.Parent.Text) - 6)
               XLT_EMPREENDIMENTO = TvwEmpresas.Nodes.Item(i).Parent.Text
               XLT_CLIENTE = Mid(TvwEmpresas.Nodes.Item(i).Text, 6, 4) & " - " & Mid(TvwEmpresas.Nodes.Item(i).Text, 16, Len(TvwEmpresas.Nodes.Item(i).Text) - 15)
               
               XLT_CONTRATOS = Mid(TvwEmpresas.Nodes.Item(i).Text, 1, 12)
                   
               With XLO_COMANDO
                   .CommandTimeout = 10000000
                   .CommandText = "spCAPI_sel_TitulosLiquidadosPISCOFINS"
                   .CommandType = adCmdStoredProc
                   Set XLO_COMANDO.ActiveConnection = Conexao
                   'Cria automaticamente todos os parâmetros dentro do objeto Comando
                   .Parameters.Refresh
               End With
               
               'Passa os parâmetros para a procedure
               XLO_COMANDO.Parameters(1).Value = XLI_EMPRESA
               XLO_COMANDO.Parameters(2).Value = DtpInicioPagamento.Value
               XLO_COMANDO.Parameters(3).Value = DtpFimPagamento.Value
               XLO_COMANDO.Parameters(4).Value = IIf(Len(XLT_CONTRATOS) > 0, XLT_CONTRATOS, Null)
               XLO_COMANDO.Parameters(6).Value = IIf(Len(XLT_FILTRO) > 0, XLT_FILTRO, Null)
               XLO_COMANDO.Parameters(7).Value = ChkMesesComRecebimento.Value
               
               'Executa a procedure e guarda o resultado no objeto rs
               Set rs = XLO_COMANDO.Execute
                
               If Not rs.EOF Then
                   If sheet > 3 Then
                       
                       XFO_ExlObj.Sheets.Add
                       'Move a planilha criada para a nova aba
                       XFO_ExlObj.ActiveSheet.Move After:=XFO_ExlObj.Sheets(XFO_ExlObj.ActiveWorkbook.Sheets.Count)
                        
                   End If
                   'Função para criar a planilha

                   ExportaExcelCliente rs, sheet, XLT_NOMEEMPRESA, XLT_EMPREENDIMENTO, XLT_CLIENTE
           
                   sheet = sheet + 1
                   
               End If
            
              rs.Close
            
            
            End If
            
            i = i + 1
            
        Wend
       
        If sheet > 1 Then
             'Salva a planilha no caminho selecionado pelo usuário
             XFO_ExlObj.Sheets(1).Select
             XFO_ExlObj.Workbooks(1).SaveAs dlgSalvar.FileName
             MsgBox "Relatório exportado com sucesso!", vbInformation, "Informação"
        Else
            MsgBox "Não foram encontrados registros para a exportação. A operação foi cancelada.", vbInformation, "Informação"
        End If
        
        XFO_ExlObj.Workbooks(1).Close
        XFO_ExlObj.Application.Quit
        Set XFO_ExlObj = Nothing
        
        'Limpa o caminho do arquivo
        dlgSalvar.FileName = ""
        
        FrmRelTitulosLiquidadosPis.MousePointer = vbDefault
    
    End If
    
ErrHandler:

    Err.Clear
    Exit Sub
    
End Sub

Private Sub ExportaExcelEmpresaEmpreendimento(dados As ADODB.Recordset, sheet As Integer, titulo As String, subtitulo As String)
    Dim x       As Integer 'Auxiliar
    Dim i       As Integer 'Auxiliar
    Dim Cont    As Integer 'Auxiliar
    Dim XLD_ANOATUAL As Integer
     
    
    'Seleciona a próxima planilha vazia para ser preenchida
    XFO_ExlObj.Sheets(sheet).Select
    If OptEmpresa.Value = False Then
        XFO_ExlObj.Sheets(sheet).Name = IIf(Len(subtitulo) > 31, Mid(subtitulo, 1, 31), subtitulo)
    Else
        XFO_ExlObj.Sheets(sheet).Name = IIf(Len(titulo) > 31, Mid(titulo, 1, 31), titulo)
    End If
    
    'Configurando margem da página
    With XFO_ExlObj.ActiveSheet.PageSetup
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .FooterMargin = Application.CentimetersToPoints(0.8)
        .Orientation = xlLandscape
        .PrintTitleRows = "$5:$7" 'Repete o cabeçalho em todas as páginas
        .Zoom = 80
    End With
 
    'Preenchendo as células
    With XFO_ExlObj.ActiveSheet
        .Cells(1, 1).Value = titulo
        .Cells(2, 1).Value = subtitulo
        .Cells(3, 1).Value = "CÁLCULO DO PIS/COFINS EXCLUINDO A CORREÇÃO MONETÁRIA CLIENTES"
        .Cells(4, 14).Value = Format(Now, "DD/MM/YYYY hh:mm")
        .Cells(6, 1).Value = "MÊS/ANO"
        .Cells(6, 2).Value = "VALOR PAGO PELO CLIENTE"
        .Cells(6, 3).Value = "PAGAMENTOS DE DISTRATO"
        .Cells(6, 4).Value = "VALOR PAGO PELO CLIENTE LIQUIDO DE DISTRATO"
        .Cells(6, 5).Value = "ATUALIZAÇÃO MONETÁRIA"
        .Cells(6, 6).Value = "RECEBIMENTO LÍQUIDO"
        .Cells(5, 1).Value = Format(DtpInicioPagamento.Value, "DD/MM/YYYY") & " - " & Format(DtpFimPagamento.Value, "DD/MM/YYYY")
        'PIS
        .Cells(5, 8).Value = "PIS"
        .Cells(5, 9).Value = "0,65%"
        .Cells(6, 8).Value = "SEM CORREÇÃO"
        .Cells(6, 9).Value = "PAGO"
        .Cells(6, 10).Value = "CRÉDITO"
        'COFINS
        .Cells(5, 12).Value = "COFINS"
        .Cells(5, 13).Value = "3%"
        .Cells(6, 12).Value = "SEM CORREÇÃO"
        .Cells(6, 13).Value = "PAGO"
        .Cells(6, 14).Value = "CRÉDITO"
   
    End With
    
    With XFO_ExlObj.ActiveSheet
        .Columns("C").Hidden = True
    End With
    
    'Formatando as células
    With XFO_ExlObj.ActiveSheet.Range("A1:A3")
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    With XFO_ExlObj.ActiveSheet.Range("A5:N6")
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    With XFO_ExlObj.ActiveSheet.Range("N4")
        .Font.Size = 7
    End With

    'Bordas do Título
    With XFO_ExlObj.ActiveSheet.Range("A1:N1").Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A4:N4").Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A1:A4").Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("N1:N4").Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.ActiveSheet
        
        'Mesclar
        .Range("A1:N1").Merge
        .Range("A2:N2").Merge
        .Range("A3:N3").Merge
        .Range("A5:F5").Merge
        .Range("I5:J5").Merge
        .Range("M5:N5").Merge
        .Range("A7:F7").Merge
        .Range("H7:J7").Merge
        .Range("L7:N7").Merge
        
        'Alinhar texto
        .Range("A1:N3").HorizontalAlignment = xlHAlignCenter
        .Range("A5:F5").HorizontalAlignment = xlHAlignCenter
        .Range("H5:J5").HorizontalAlignment = xlHAlignCenter
        .Range("L5:N5").HorizontalAlignment = xlHAlignCenter
        .Range("A6:N6").HorizontalAlignment = xlHAlignCenter
        .Range("A6:N6").VerticalAlignment = xlVAlignCenter
        
        'Quebra de texto automática
        .Range("A6:N6").WrapText = True
        
        'Altura da linha
        .Range("A1:N1").RowHeight = 20
        .Range("A2:N2").RowHeight = 20
        .Range("A3:N3").RowHeight = 20
        .Range("A7:N7").RowHeight = 10
        
        'Largura da Coluna
        .Range("G1;K1").ColumnWidth = 0.5
        .Range("A7").ColumnWidth = 12
        .Range("B7;D7;F7").ColumnWidth = 15
        .Range("C7").ColumnWidth = 13
        .Range("E7").ColumnWidth = 14
        .Range("H7:I7").ColumnWidth = 13
        .Range("J7").ColumnWidth = 12
        .Range("L7:M7").ColumnWidth = 13
        .Range("N7").ColumnWidth = 12
        
        'Cor da célula
        .Range("A7:F7").Interior.Color = RGB(216, 216, 216)
        .Range("H7:J7").Interior.Color = RGB(216, 216, 216)
        .Range("L7:N7").Interior.Color = RGB(216, 216, 216)
    
    End With
    
    'Bordas das Colunas
    With XFO_ExlObj.ActiveSheet.Range("A5:N7").Borders(xlInsideHorizontal)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A5:N7").Borders(xlInsideVertical)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.ActiveSheet.Range("A5:N7").Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A5:N7").Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A5:A7").Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("N5:N7").Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    i = 8
    Cont = 8

    XLD_ANOATUAL = Year(dados!Mes)
    
    'Insere os dados
    While Not dados.EOF
        
        With XFO_ExlObj.ActiveSheet
        
            'Quebra de linha a cada ano
            If XLD_ANOATUAL <> Year(dados!Mes) Then
                
                'Mesclar
                .Range("A" & Cont & ":F" & Cont).Merge
                .Range("H" & Cont & ":J" & Cont).Merge
                .Range("L" & Cont & ":N" & Cont).Merge
                
                'Cor das células
                .Range("A" & Cont & ":F" & Cont).Interior.Color = RGB(216, 216, 216)
                .Range("H" & Cont & ":J" & Cont).Interior.Color = RGB(216, 216, 216)
                .Range("L" & Cont & ":N" & Cont).Interior.Color = RGB(216, 216, 216)
                
                'Altura da linha
                .Range("A" & Cont & ":N" & Cont).RowHeight = 5
                
                XLD_ANOATUAL = Year(dados!Mes)
                Cont = Cont + 1
                i = i + 1
            End If
        
            .Cells(i, 1) = Format(dados!Mes, "MMMM-YY")
            If Not IsNull(dados.Fields("ValorPagoCliente").Value) Then .Cells(i, 2) = dados.Fields("ValorPagoCliente").Value
            If Not IsNull(dados.Fields("Distrato").Value) Then .Cells(i, 3) = dados.Fields("Distrato").Value
            .Cells(i, 4).Formula = "=B" & i & "-C" & i
            If Not IsNull(dados.Fields("Correcao").Value) Then .Cells(i, 5) = dados.Fields("Correcao").Value
            .Cells(i, 6).Formula = "=D" & i & "-E" & i
            .Cells(i, 8).Formula = "=F" & i & "*$I$5" 'PIS
            If Not IsNull(dados.Fields("PISPago").Value) Then .Cells(i, 9) = dados.Fields("PISPago").Value
            .Cells(i, 10).Formula = "=I" & i & "-H" & i
            .Cells(i, 12).Formula = "=F" & i & "*$M$5" 'COFINS
            If Not IsNull(dados.Fields("COFINSPago").Value) Then .Cells(i, 13).Formula = dados.Fields("COFINSPago").Value
            .Cells(i, 14).Formula = "=M" & i & "-L" & i


        End With
        dados.MoveNext
        i = i + 1
        Cont = Cont + 1
    Wend

    Cont = Cont - 1
    
    'Alinha o texto da coluna mês para a direita
    XFO_ExlObj.ActiveSheet.Range("A8:A" & Cont - 1).HorizontalAlignment = xlHAlignRight

    'Insere borda na planilha
    With XFO_ExlObj.ActiveSheet.Range("A8:N" & Cont).Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A8:N" & Cont).Borders(xlInsideVertical)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A9:N" & Cont).Borders(xlEdgeTop)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A8:N" & Cont - 1).Borders(xlEdgeBottom)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A8:N" & Cont).Borders(xlEdgeLeft)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A8:N" & Cont).Borders(xlEdgeRight)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With
  
    With XFO_ExlObj.ActiveSheet
        'Fórmulas para calcular os totais
        .Cells(Cont + 2, 1).Value = "TOTAL"
        .Cells(Cont + 2, 2).Formula = "=SUM(B8:B" & Cont & ")"
        .Cells(Cont + 2, 3).Formula = "=SUM(C8:C" & Cont & ")"
        .Cells(Cont + 2, 4).Formula = "=SUM(D8:D" & Cont & ")"
        .Cells(Cont + 2, 5).Formula = "=SUM(E8:E" & Cont & ")"
        .Cells(Cont + 2, 6).Formula = "=SUM(F8:F" & Cont & ")"
        .Cells(Cont + 2, 8).Formula = "=SUM(H8:H" & Cont & ")"
        .Cells(Cont + 2, 9).Formula = "=SUM(I8:I" & Cont & ")"
        .Cells(Cont + 2, 10).Formula = "=SUM(J8:J" & Cont & ")"
        .Cells(Cont + 2, 12).Formula = "=SUM(L8:L" & Cont & ")"
        .Cells(Cont + 2, 13).Formula = "=SUM(M8:M" & Cont & ")"
        .Cells(Cont + 2, 14).Formula = "=SUM(N8:N" & Cont & ")"

    End With
    
    With XFO_ExlObj.ActiveSheet
        'Mesclar
        .Range("A" & Cont + 1 & ":F" & Cont + 1).Merge
        .Range("H" & Cont + 1 & ":J" & Cont + 1).Merge
        .Range("L" & Cont + 1 & ":N" & Cont + 1).Merge
        .Range("G5:G" & Cont + 1).Merge
        .Range("K5:K" & Cont + 1).Merge
        'Altura da linha
        .Range("A" & Cont + 1 & ":N" & Cont + 1).RowHeight = 10

    End With
    
    'Formata o formato dos números
     XFO_ExlObj.ActiveSheet.Range("B8:N" & Cont + 2).NumberFormat = "_(* #,##0.00_);[red]_(* (#,##0.00);_(* ""-""_);_(@_)"
    
    'Formata a fonte
    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 2 & ":N" & Cont + 2)
        .Font.Size = 11
        .Font.Bold = True
    End With

    'Insere borda para os totais
    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 1 & ":N" & Cont + 2).Borders(xlInsideHorizontal)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 1 & ":N" & Cont + 2).Borders(xlInsideVertical)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 1 & ":N" & Cont + 2).Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 1 & ":N" & Cont + 2).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 1 & ":N" & Cont + 2).Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 1 & ":N" & Cont + 2).Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With
    
    'Formata as colunas que separam o cálculo de PIS  e Cofins
    With XFO_ExlObj.ActiveSheet.Range("G5:G" & Cont + 2 & ";K5:K" & Cont + 2)
       .Interior.Color = RGB(216, 216, 216)
    End With
    
    With XFO_ExlObj.ActiveSheet.Range("G5:G" & Cont + 2 & ";K5:K" & Cont + 2).Borders(xlInsideHorizontal)
        .LineStyle = xlDouble
        .ColorIndex = 15
    End With

    With XFO_ExlObj.ActiveSheet.Range("G5:G" & Cont + 2 & ";K5:K" & Cont + 2).Borders(xlInsideVertical)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("G5:G" & Cont + 2 & ";K5:K" & Cont + 2).Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("G5:G" & Cont + 2 & ";K5:K" & Cont + 2).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("G5:G" & Cont + 2 & ";K5:K" & Cont + 2).Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("G5:G" & Cont + 2 & ";K5:K" & Cont + 2).Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With
    
    'Insere bordas duplas nas laterais
    With XFO_ExlObj.ActiveSheet.Range("A5:A" & Cont + 2).Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With
    
      With XFO_ExlObj.ActiveSheet.Range("N5:N" & Cont + 2).Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    'Formata para não exibir as linhas de grade da planilha
    XFO_ExlObj.ActiveWindow.DisplayGridlines = False

End Sub

Private Sub ExportaExcelCliente(dados As ADODB.Recordset, sheet As Integer, titulo As String, subtitulo As String, cliente As String)
    Dim x       As Integer 'Auxiliar
    Dim i       As Integer 'Auxiliar
    Dim Cont    As Integer 'Auxiliar
    Dim XLD_ANOATUAL As Integer
     
    
    'Seleciona a próxima planilha vazia para ser preenchida
    XFO_ExlObj.Sheets(sheet).Select
    XFO_ExlObj.Sheets(sheet).Name = IIf(Len(subtitulo) > 31, Mid(subtitulo, 1, 31), subtitulo)
    
    'Configurando margem da página
    With XFO_ExlObj.ActiveSheet.PageSetup
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .FooterMargin = Application.CentimetersToPoints(0.8)
        .Orientation = xlPortrait
        .PrintTitleRows = "$5:$7" 'Repete o cabeçalho em todas as páginas
        .Zoom = 100
    End With
 
    'Preenchendo as células
    With XFO_ExlObj.ActiveSheet
        
        .Cells(1, 1).Value = titulo
        .Cells(2, 1).Value = subtitulo
        .Cells(3, 1).Value = cliente
        .Cells(4, 1).Value = "CÁLCULO DO PIS/COFINS EXCLUINDO A CORREÇÃO MONETÁRIA CLIENTES"
        .Cells(5, 6).Value = Format(Now, "DD/MM/YYYY hh:mm")
        .Cells(7, 1).Value = "MÊS/ANO"
        .Cells(7, 2).Value = "VALOR PAGO PELO CLIENTE"
        .Cells(7, 3).Value = "PAGAMENTOS DE DISTRATO"
        .Cells(7, 4).Value = "VALOR PAGO PELO CLIENTE LIQUIDO DE DISTRATO"
        .Cells(7, 5).Value = "ATUALIZAÇÃO MONETÁRIA"
        .Cells(7, 6).Value = "RECEBIMENTO LÍQUIDO"
        .Cells(6, 1).Value = Format(DtpInicioPagamento.Value, "DD/MM/YYYY") & " - " & Format(DtpFimPagamento.Value, "DD/MM/YYYY")
   
    End With

    'Formatando as células
    With XFO_ExlObj.ActiveSheet.Range("A1:A2")
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    With XFO_ExlObj.ActiveSheet.Range("A4")
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    With XFO_ExlObj.ActiveSheet.Range("F5")
        .Font.Size = 7
    End With
    
    With XFO_ExlObj.ActiveSheet.Range("A6:F7")
        .Font.Size = 11
        .Font.Bold = True
    End With

    'Bordas do Título
    With XFO_ExlObj.ActiveSheet.Range("A1:F1").Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A4:F5").Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A1:A5").Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("F1:F5").Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.ActiveSheet
        
        'Mesclar
        .Range("A1:F1").Merge
        .Range("A2:F2").Merge
        .Range("A4:F4").Merge
        .Range("A3:F3").Merge
        .Range("A5:F5").Merge
        .Range("A6:F6").Merge
       
        'Alinhar texto
        .Range("A1:F5").HorizontalAlignment = xlHAlignCenter
        .Range("F5").HorizontalAlignment = xlHAlignRight
        .Range("A6:F7").HorizontalAlignment = xlHAlignCenter
        .Range("A6:F7").VerticalAlignment = xlVAlignCenter
        
        'Quebra de texto automática
        .Range("A7:F7").WrapText = True
        
        'Altura da linha
        .Range("A1:F1").RowHeight = 20
        .Range("A2:F2").RowHeight = 20
        .Range("A3:F3").RowHeight = 20
        .Range("A8:F8").RowHeight = 10
        
        'Largura das colunas
        .Range("A1:F1").ColumnWidth = 15
        
        'Cor da célula
        .Range("A8:F8").Interior.Color = RGB(216, 216, 216)
    
    End With
    
    'Bordas das Colunas
    With XFO_ExlObj.ActiveSheet.Range("A6:F8").Borders(xlInsideHorizontal)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A6:F7").Borders(xlInsideVertical)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With
    
    With XFO_ExlObj.ActiveSheet.Range("A6:F8").Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A6:F8").Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A6:F8").Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("F6:F8").Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    i = 9
    Cont = 9

    XLD_ANOATUAL = Year(dados!Mes)
    
    'Insere os dados
    While Not dados.EOF
        
        With XFO_ExlObj.ActiveSheet
        
            'Quebra de linha a cada ano
            If XLD_ANOATUAL <> Year(dados!Mes) Then
                
                'Mesclar
                .Range("A" & Cont & ":F" & Cont).Merge
                               
                'Cor das células
                .Range("A" & Cont & ":F" & Cont).Interior.Color = RGB(216, 216, 216)
                
                'Altura da linha
                .Range("A" & Cont & ":F" & Cont).RowHeight = 5
                
                XLD_ANOATUAL = Year(dados!Mes)
                Cont = Cont + 1
                i = i + 1
            End If
        
            .Cells(i, 1) = Format(dados!Mes, "MMMM-YY")
            If Not IsNull(dados.Fields("ValorPagoCliente").Value) Then .Cells(i, 2) = dados.Fields("ValorPagoCliente").Value
            If Not IsNull(dados.Fields("Distrato").Value) Then .Cells(i, 3) = dados.Fields("Distrato").Value
            .Cells(i, 4).Formula = "=B" & i & "-C" & i
            If Not IsNull(dados.Fields("Correcao").Value) Then .Cells(i, 5) = dados.Fields("Correcao").Value
            .Cells(i, 6).Formula = "=D" & i & "-E" & i

        End With
        dados.MoveNext
        i = i + 1
        Cont = Cont + 1
    Wend

    Cont = Cont - 1
    
    'Alinha o texto da coluna mês para a direita
    XFO_ExlObj.ActiveSheet.Range("A9:A" & Cont - 1).HorizontalAlignment = xlHAlignRight

    'Insere borda na planilha
    With XFO_ExlObj.ActiveSheet.Range("A9:F" & Cont).Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A9:F" & Cont).Borders(xlInsideVertical)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A10:F" & Cont).Borders(xlEdgeTop)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A9:F" & Cont - 1).Borders(xlEdgeBottom)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A9:F" & Cont).Borders(xlEdgeLeft)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A9:F" & Cont).Borders(xlEdgeRight)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
    End With
  
    With XFO_ExlObj.ActiveSheet
        'Fórmulas para calcular os totais
        .Cells(Cont + 2, 1).Value = "TOTAL"
        .Cells(Cont + 2, 2).Formula = "=SUM(B9:B" & Cont & ")"
        .Cells(Cont + 2, 3).Formula = "=SUM(C9:C" & Cont & ")"
        .Cells(Cont + 2, 4).Formula = "=SUM(D9:D" & Cont & ")"
        .Cells(Cont + 2, 5).Formula = "=SUM(E9:E" & Cont & ")"
        .Cells(Cont + 2, 6).Formula = "=SUM(F9:F" & Cont & ")"

    End With
    
    With XFO_ExlObj.ActiveSheet
        
        'Mesclar
        .Range("A" & Cont + 1 & ":F" & Cont + 1).Merge

        'Altura da linha
        .Range("A" & Cont + 1 & ":F" & Cont + 1).RowHeight = 10

    End With
    
    'Formata o formato dos números
     XFO_ExlObj.ActiveSheet.Range("B9:F" & Cont + 2).NumberFormat = "_(* #,##0.00_);[red]_(* (#,##0.00);_(* ""-""_);_(@_)"
    
    'Formata a fonte
    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 2 & ":F" & Cont + 2)
        .Font.Size = 11
        .Font.Bold = True
    End With

    'Insere borda para os totais
    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 1 & ":F" & Cont + 2).Borders(xlInsideHorizontal)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 1 & ":F" & Cont + 2).Borders(xlInsideVertical)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 1 & ":F" & Cont + 2).Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 1 & ":F" & Cont + 2).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 1 & ":F" & Cont + 2).Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    With XFO_ExlObj.ActiveSheet.Range("A" & Cont + 1 & ":F" & Cont + 2).Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With
   
    'Insere bordas duplas nas laterais
    With XFO_ExlObj.ActiveSheet.Range("A6:A" & Cont + 2).Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With
    
      With XFO_ExlObj.ActiveSheet.Range("F6:F" & Cont + 2).Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
    End With

    'Formata para não exibir as linhas de grade da planilha
    XFO_ExlObj.ActiveWindow.DisplayGridlines = False

End Sub


'Carrega o treeview com as empresas
Private Sub PreencheTreeViewEmpresas()
    Dim nodeEmpresa As MSComctlLib.node
            
    TvwEmpresas.Nodes.Clear
        
    'Cria os nós principais
    DatEmpresa.Recordset.MoveFirst
    
    While Not DatEmpresa.Recordset.EOF
      Set nodeEmpresa = TvwEmpresas.Nodes.Add(, , "N" & DatEmpresa.Recordset.Fields("empr_tx_Fantasia"), _
            Format(DatEmpresa.Recordset.Fields("empr_cd_Empresa"), "000") & " - " & DatEmpresa.Recordset.Fields("empr_tx_RazaoSocial"))
        
        nodeEmpresa.Tag = "Nó 1"
        
        'Cria as folhas do nó criado
        CarregaEmpreedimentos nodeEmpresa
        
        'Atualiza a quantidade de empreendimentos
        XFI_QtdEmpreendimento = XFI_QtdEmpreendimento + 1
        
        DatEmpresa.Recordset.MoveNext
    Wend
End Sub


'Preenche a árvore do nó selecionado
Private Sub CarregaEmpreedimentos(ByVal nodeNo As MSComctlLib.node)
    Dim nodeEmpreendimento As MSComctlLib.node
    
    'Se os registros do nó já foram carregados não faz nada
    If nodeNo.Children > 0 Then Exit Sub
    
    DatEmpreendimento.Recordset.MoveFirst
    
    While Not DatEmpreendimento.Recordset.EOF
        If Mid(nodeNo.Text, 1, 3) = Format(DatEmpreendimento.Recordset.Fields("empr_cd_empresa"), "000") Then
            Set nodeEmpreendimento = TvwEmpresas.Nodes.Add(nodeNo, tvwChild, , _
                           DatEmpreendimento.Recordset.Fields("empd_cd_empreendimento") & " - " & DatEmpreendimento.Recordset.Fields("empd_tx_Nome"))
            
            nodeEmpreendimento.Tag = "Nó 2"
            
            CarregaContratos nodeEmpreendimento
            
        End If
        DatEmpreendimento.Recordset.MoveNext
    Wend
End Sub

'Preenche a árvore do nó selecionado
Private Sub CarregaContratos(ByVal nodeNo As MSComctlLib.node)
    Dim nodeFolha As MSComctlLib.node
    
    'Se os registros do nó já foram carregados não faz nada
    If nodeNo.Children > 0 Then Exit Sub
    
    DatContrato.Recordset.MoveFirst
    
    While Not DatContrato.Recordset.EOF
        If Mid(nodeNo.Parent.Text, 1, 3) = Format(DatContrato.Recordset.Fields("empr_cd_empresa"), "000") And Mid(nodeNo.Text, 1, 4) = Mid(DatContrato.Recordset.Fields("Contrato"), 1, 4) Then
            Set nodeFolha = TvwEmpresas.Nodes.Add(nodeNo, tvwChild, , _
                            DatContrato.Recordset.Fields("Contrato") & " - " & DatContrato.Recordset.Fields("Cliente") & " - " & DatContrato.Recordset.Fields("Status"))
            
            nodeFolha.Tag = "Folha"
        End If
        DatContrato.Recordset.MoveNext
    Wend
End Sub

Private Sub TvwEmpresas_NodeCheck(ByVal node As MSComctlLib.node)
  CheckChildren node
End Sub

Private Sub CheckChildren(node As MSComctlLib.node)
   Dim i As Integer
   Dim nodX As MSComctlLib.node
 
  'Se tiver níveis inferiores
  If node.Children <> 0 Then
    Set nodX = node.Child
    For i = 1 To node.Children
      nodX.Checked = node.Checked
      
      CheckChildren nodX 'Se tiver outros níveis considera a marcação do nó superior
      
      Set nodX = nodX.Next
      
    Next
  End If
End Sub


Private Sub GetContratos(ByRef node As MSComctlLib.node)
    Dim i As Integer
    Dim nodX As MSComctlLib.node
    Dim XLT_CONT As String

    'Se tiver níveis inferiores
    If node.Children <> 0 Then
        Set nodX = node.Child
        For i = 1 To node.Children
            If (nodX.Tag = "Folha" And nodX.Checked = True) Then
              XLT_CONT = XLT_CONT & Mid(nodX.Text, 1, 12) & ";"
            End If

            GetContratos nodX

          Set nodX = nodX.Next
        Next
    End If

    XLT_CONTRATOS = XLT_CONTRATOS & XLT_CONT

End Sub



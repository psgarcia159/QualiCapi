VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmRelCompradoresCadastrados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relação de Compradores"
   ClientHeight    =   2085
   ClientLeft      =   3330
   ClientTop       =   3615
   ClientWidth     =   6630
   Icon            =   "FrmRelCompradoresCadastrados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgSalvar 
      Left            =   480
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   5430
      TabIndex        =   9
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin MSMask.MaskEdBox MskEmpreendimento 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "AAAA"
         PromptChar      =   " "
      End
      Begin MSDataListLib.DataCombo CboEmpreendimento 
         Bindings        =   "FrmRelCompradoresCadastrados.frx":2AFA
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "empd_tx_nome"
         BoundColumn     =   "empd_cd_empreendimento"
         Text            =   "CboEmpreendimento"
      End
      Begin Threed.SSCommand cmdLimparEmpreendimento 
         Height          =   315
         Left            =   6060
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   78
         MouseIcon       =   "FrmRelCompradoresCadastrados.frx":2B1A
         Picture         =   "FrmRelCompradoresCadastrados.frx":2B36
      End
      Begin VB.Label LblEmpreendimento 
         Alignment       =   1  'Right Justify
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
         Left            =   60
         TabIndex        =   1
         Top             =   300
         Width           =   1485
      End
   End
   Begin Threed.SSFrame FraStatus 
      Height          =   660
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   3285
      _Version        =   65536
      _ExtentX        =   5794
      _ExtentY        =   1164
      _StockProps     =   14
      Caption         =   "Tipo "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox ChkFisica 
         Caption         =   "Física"
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
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   1065
      End
      Begin VB.CheckBox ChkJuridica 
         Caption         =   "Jurídica"
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
         Left            =   1860
         TabIndex        =   7
         Top             =   240
         Width           =   1080
      End
   End
   Begin MSAdodcLib.Adodc DatEmpreendimento 
      Height          =   330
      Left            =   120
      Top             =   1680
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
      Left            =   4320
      TabIndex        =   8
      Top             =   1680
      Width           =   1035
   End
End
Attribute VB_Name = "FrmRelCompradoresCadastrados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CboEmpreendimento_Click(Area As Integer)

    If CboEmpreendimento.Text <> "" Then
      DatEmpreendimento.Recordset.bookmark = CboEmpreendimento.SelectedItem
      MskEmpreendimento.Text = CboEmpreendimento.BoundText

    End If
    
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    Dim XLT_EMPREENDIMENTO As String 'Empreendimento escolhido
    Dim xlt_sql As String
    Dim XLT_CONDICAO As String
    
    On Error GoTo TrataErro
                     
    If CboEmpreendimento.BoundText <> "" Then
        XLT_CONDICAO = " AND empd_cd_empreendimento = '" + CboEmpreendimento.BoundText + "'"
    Else
        MsgBox "É necessário selecionar um empreendimento", vbInformation, "Atenção"
        Exit Sub
    End If
                
              
    'Seta o Titulo da Tela de Dialogo salvar
    dlgSalvar.DialogTitle = "Salvar Arquivo"
    
    'Seta o filtro das extensões de arquivos que será usada
    dlgSalvar.Filter = "Documento do Word (*.doc)|*.doc"
    
    'Indica qual sera o filtro padrão
    dlgSalvar.FilterIndex = 1
    
    dlgSalvar.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
            
    'Exibe a tela do dialogo para o usuário
    dlgSalvar.ShowSave

    'Verifica se o usuário selecionou o caminho do arquivo
    If (dlgSalvar.FileName <> "") Then
    
        'Altera o cursor do mouse
        FrmRelCompradoresCadastrados.MousePointer = vbHourglass
        'Comando sql
        
        xlt_sql = " SELECT * FROM ConsCAPRelCompradores WHERE empr_cd_empresa = " & PCodEmpresa
   
        'Verifica o tipo de Cliente
        If ChkFisica.Value = 1 And ChkJuridica.Value = 0 Then
            XLT_CONDICAO = XLT_CONDICAO & " AND focl_tx_tipo = 'F'"
         
        ElseIf ChkFisica.Value = 0 And ChkJuridica.Value = 1 Then
            XLT_CONDICAO = XLT_CONDICAO & " AND focl_tx_tipo = 'J'"
        End If
        
        XLT_CONDICAO = XLT_CONDICAO & " ORDER BY imov_cd_imovel"
        
        'Passa para o recordset o resultado da consulta
        Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(xlt_sql & XLT_CONDICAO)
        
        'Verifica se a consulta retornou dados, caso contrário finaliza a rotina
        If XGO_RSRELATORIO.EOF Then
            MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
            FrmRelRelacaoClientes.MousePointer = vbDefault
            Exit Sub
        End If
              
        Set XGR_RELATORIO = Nothing
        'Carrega o template do relatório
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelCompradoresCadastrados.rpt")

       'Passa o recordset para o relatório
        XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO
 
        
       'Ativa os Relatórios
       With XGR_RELATORIO
            'Preenche  a formula empresa do relatório com a empresa ativa
            .FormulaFields.GetItemByName("Filtro").Text = Chr(34) & PEmpresa & Chr(34)
            'Preenche a formula Filtro do relatório os filtros escolhido pelo usuário
            .FormulaFields.GetItemByName("filtro2").Text = Chr(34) & CboEmpreendimento.Text & Chr(34)
            'Indica o local aonde será salvo o arquivo que ja foi informado pelo usuário
           .ExportOptions.DiskFileName = dlgSalvar.FileName
           'Indica o destino do arquivo
           .ExportOptions.DestinationType = crEDTDiskFile
           'Indica o tipo do arquivo que o relatório sera exportado (no caso abaixo RTF)
           .ExportOptions.FormatType = crEFTExactRichText
           .Export False
        End With
       
        'Limpa o caminho do arquivo
        dlgSalvar.FileName = ""
        FrmRelCompradoresCadastrados.MousePointer = vbDefault
        MsgBox "Relatório Exportado com sucesso!", vbInformation, "Informação"
              
        
        
    End If
      
Exit Sub

TrataErro:

    FrmRelCompradoresCadastrados.MousePointer = vbDefault
    MsgBox "Ocorreu um erro na exportação do relatório. Contate o suporte. Detalhe do Erro: " & Err.Description, vbCritical, "Erro"
    DeleteFile (dlgSalvar.FileName)

End Sub

Private Sub CmdLimpar_Click()
    CboCliente.Text = ""
End Sub

Private Sub CmdLimparEmpreendimento_Click()
    CboEmpreendimento.BoundText = ""
    MskEmpreendimento.Text = ""
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelCompradoresCadastrados
End Sub

Private Sub Form_Load()
  ' subManutencaoJanelasAtivas "I", "FrmRelRelacaoClientes"
  subConectarControleDadosNV DatEmpreendimento, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_Nome", Estatico
End Sub




Private Sub MskEmpreendimento_LostFocus()
  
  If MskEmpreendimento.Text <> "" Then
    SubPreencheComZero MskEmpreendimento
    CboEmpreendimento.BoundText = MskEmpreendimento.Text
    If CboEmpreendimento.Text = "" Then
      MsgBox "Não existe Empreendimento com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
      MskEmpreendimento.SetFocus
      Exit Sub
    End If
  Else
    CboEmpreendimento.BoundText = ""
  End If
  
End Sub

Private Sub MskEmpreendimento_GotFocus()
    Call subSelecionaMSK(MskEmpreendimento)
End Sub

Private Sub MskEmpreendimento_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
    End If
End Sub



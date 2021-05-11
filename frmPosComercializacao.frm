VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmRelPosComercializacaoItau 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Posição de Comercialização"
   ClientHeight    =   1440
   ClientLeft      =   3330
   ClientTop       =   3615
   ClientWidth     =   4875
   Icon            =   "frmPosComercializacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgSalvar 
      Left            =   2880
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   3780
      TabIndex        =   5
      Top             =   1035
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin MSMask.MaskEdBox MskEmpreendimento 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   360
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
         Bindings        =   "frmPosComercializacao.frx":2AFA
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   360
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
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1485
      End
   End
   Begin MSAdodcLib.Adodc DatEmpreendimento 
      Height          =   330
      Left            =   -240
      Top             =   1200
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
      Left            =   2640
      TabIndex        =   4
      Top             =   1035
      Width           =   1035
   End
End
Attribute VB_Name = "FrmRelPosComercializacaoItau"
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
  
    Dim XLO_COMANDO As New ADODB.Command
    Dim rs          As New ADODB.Recordset
     
    On Error GoTo TrataErro
                     
    If CboEmpreendimento.BoundText = "" Then
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
        FrmRelPosComercializacaoItau.MousePointer = vbHourglass
        'Comando sql
        
        With XLO_COMANDO
            .CommandTimeout = 300
            .CommandText = "spCAPI_sel_posicaoComercializacaoITAU"
            .CommandType = adCmdStoredProc
            Set XLO_COMANDO.ActiveConnection = Conexao
             'Cria automaticamente todos os parâmetros dentro do objeto Comando
            .Parameters.Refresh
        End With

        XLO_COMANDO.Parameters(1).Value = PCodEmpresa 'Empresa
        XLO_COMANDO.Parameters(2).Type = adBSTR
        XLO_COMANDO.Parameters(2).Value = CboEmpreendimento.BoundText 'Empreendimento

        Set rs = XLO_COMANDO.Execute
        
        Set rs = Conexao.Execute("Select * from ##PosComercializacao")
       
        
        'Passa para o recordset o resultado da consulta
                          
        'Verifica se a consulta retornou dados, caso contrário finaliza a rotina
        If rs.EOF Then
            FrmRelRelacaoClientes.MousePointer = vbDefault
            MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
            Exit Sub
        End If
              
        Set XGR_RELATORIO = Nothing
        'Carrega o template do relatório
        Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelPosicaoComercializacao.rpt")

       'Passa o recordset para o relatório
        XGR_RELATORIO.Database.SetDataSource rs
        
       
       'Ativa os Relatórios
       With XGR_RELATORIO
            'Preenche  a formula empresa do relatório com a empresa ativa
            .FormulaFields.GetItemByName("Empresa").Text = Chr(34) & PEmpresa & Chr(34)
            'Preenche a formula Filtro do relatório os filtros escolhido pelo usuário
            .FormulaFields.GetItemByName("Empreendimento").Text = Chr(34) & CboEmpreendimento.Text & Chr(34)
            'Indica o local aonde será salvo o arquivo que ja foi informado pelo usuário
            .ExportOptions.DiskFileName = dlgSalvar.FileName
            'Indica o destino do arquivo
            .ExportOptions.DestinationType = crEDTDiskFile
            'Indica o tipo do arquivo que o relatório sera exportado (no caso abaixo RTF)
            .ExportOptions.FormatType = crEFTExactRichText
            .Export False
        End With
             
        Conexao.Execute ("DROP TABLE ##PosComercializacao ")
        'Limpa o caminho do arquivo
        dlgSalvar.FileName = ""
        FrmRelPosComercializacaoItau.MousePointer = vbDefault
        MsgBox "Relatório Exportado com sucesso!", vbInformation, "Informação"
              
    End If
Exit Sub

TrataErro:
    Conexao.Execute ("DROP TABLE ##PosComercializacao ")
    FrmRelPosComercializacaoItau.MousePointer = vbDefault
    MsgBox "Ocorreu um erro na exportação do relatório. Contate o suporte. Detalhe do Erro: " & Err.Description, vbCritical, "Erro"
    DeleteFile (dlgSalvar.FileName)

End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelPosComercializacaoItau
End Sub

Private Sub Form_Load()
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



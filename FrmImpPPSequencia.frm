VERSION 5.00
Begin VB.Form FrmImpPPSequencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão de Processo de Pagamento em Sequência"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmImpPPSequencia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkEstoque 
      Caption         =   "Apenas PPs pendentes do QualiMat"
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   945
      Width           =   2850
   End
   Begin VB.Frame FraPPSequencia 
      Caption         =   "Processo de Pagamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   45
      TabIndex        =   5
      Top             =   90
      Width           =   4890
      Begin VB.TextBox TxtInicio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   930
         TabIndex        =   0
         Top             =   330
         Width           =   1230
      End
      Begin VB.TextBox TxtFim 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3075
         TabIndex        =   1
         Top             =   330
         Width           =   1230
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2670
         TabIndex        =   7
         Top             =   345
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   495
         TabIndex        =   6
         Top             =   345
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   285
      Left            =   3060
      TabIndex        =   3
      Top             =   945
      Width           =   900
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   285
      Left            =   4020
      TabIndex        =   4
      Top             =   945
      Width           =   900
   End
End
Attribute VB_Name = "FrmImpPPSequencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XFI_GRAFICO As Integer

Private Sub ChkEstoque_Click()
    If ChkEstoque.Value = 1 Then
        TxtInicio.Text = ""
        TxtFim.Text = ""
    End If
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub CmdImprimir_Click()
  Dim FiltroCCusto As String
  Dim ResNF As Object
  Dim NumFat As Long
  Dim ResProc As Object
  Dim XFatura As Long
  Dim X As Integer
  Dim xValor As String
  Dim Xhistorico As String, Sql As String
  Dim XL0_RSSUBRELATORIO As adodb.Recordset
  Dim XL0_RSSUBRELATORIO2 As adodb.Recordset
  Dim XL0_RSSUBRELATORIO3 As adodb.Recordset
  Dim XLT_EXISTE As String
  Dim XLT_PARCELA As String
  Dim XLT_DISTRIBUIDO As String
  Dim XLT_INICIO As String
  Dim XLT_FIM As String

  

  
  FrmImpPPSequencia.MousePointer = vbHourglass
 
  NumFat = TelaNotasFiscais.Adodc1.Recordset.Fields("nofi_nr_fatura")
  
  If TelaNotasFiscais.Adodc1.Recordset.Fields("nofi_tx_tipofatura") = "4" Then
    MsgBox "O Processo de Pagamento não pode ser emitido pois não existem duplicatas cadastradas.", vbCritical, "ATENÇÃO"
    TelaNotasFiscais.MousePointer = vbDefault
    Exit Sub
  End If
        
  'Preparação do Arquivo de Impressão
  If XControle <> "" Then
    Sql = "DELETE FROM ProcPagto_Temp WHERE prte_cd_controle = '" & XControle & "'"
    Conexao.Execute Sql
  End If
  XControle = Trim(CStr(func_cd_operador)) + CStr(Date) + CStr(Time())
  
  If ChkEstoque.Value = 1 Then 'Só trazer as notas do estoque
      Sql = "SELECT NotasFiscais.nofi_nr_fatura FROM NotasFiscais LEFT OUTER JOIN " & _
                  " Duplicatas ON NotasFiscais.nofi_nr_fatura = Duplicatas.dupl_nr_fatura " & _
                  " WHERE nofi_tx_pendEstoque = 'N' and dupl_dt_ProcPagamento is null" & _
                  " AND empr_cd_empresa = " & PCodEmpresa & " " & _
                  " ORDER BY nofi_nr_fatura"
      SubQOpenRecordset ResProc, Sql, Estatico
      If Not ResProc.EOF Then
        XLT_INICIO = ResProc!nofi_nr_fatura
        ResProc.MoveLast
        XLT_FIM = ResProc!nofi_nr_fatura
      Else
          MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
          Me.MousePointer = vbDefault
          Exit Sub
      End If
  Else
      XLT_INICIO = TxtInicio.Text
      XLT_FIM = TxtFim.Text
  End If
  
  'Inserção dos Registros de Duplicatas
  If ChkEstoque.Value = 1 Then 'Só trazer as notas do estoque
    Sql = "SELECT * FROM ConsFINProcPagto " & _
            " WHERE nofi_nr_fatura >= " & CLng(XLT_INICIO) & " " & _
            " AND nofi_nr_fatura <= " & CLng(XLT_FIM) & " " & _
            " AND NotaEstoque = 'N' and dupl_dt_ProcPagamento is null" & _
            " AND dupl_tx_debcred = 'D' " & _
            " AND empr_cd_empresa = " & PCodEmpresa & " " & _
            " ORDER BY nofi_nr_fatura, dupl_nr_duplicata"
  Else
    Sql = "SELECT * FROM ConsFINProcPagto " & _
            " WHERE nofi_nr_fatura >= " & CLng(XLT_INICIO) & " " & _
            " AND nofi_nr_fatura <= " & CLng(XLT_FIM) & " " & _
            " AND dupl_tx_debcred = 'D' " & _
            " AND empr_cd_empresa = " & PCodEmpresa & " " & _
            " ORDER BY nofi_nr_fatura, dupl_nr_duplicata"
  End If
  
  SubQOpenRecordset ResProc, Sql, Estatico
  X = 1
  'Qtd = ResProc.RecordCount
  
  
  Do While Not ResProc.EOF
     XFatura = ResProc!nofi_nr_fatura
     Xhistorico = ResProc!historico
     xValor = ResProc!dupl_vl_valor
     Sql = "INSERT INTO ProcPagto_Temp (prte_cd_Controle, prte_tx_descricao,prte_vl_valor,nofi_nr_fatura,prte_nr_parcela) " _
         & "Values('" & XControle & "','" & Xhistorico & "'," & FunTrataFloat(xValor) & "," & XFatura & "," & X & ")"
     Conexao.Execute Sql
     X = X + 1
     ResProc.MoveNext
     
     If Not ResProc.EOF Then
         XFatura = ResProc!nofi_nr_fatura
         ResProc.MovePrevious
         If XFatura <> ResProc!nofi_nr_fatura Then
             X = 1 'Reiniciar o contador de parcelas
         End If
         ResProc.MoveNext
     End If
  Loop
  ResProc.MoveLast
  'XDtVenc = ResProc!dupl_dt_vencimento
  XLT_PARCELA = ResProc!dupl_nr_duplicata
  ResProc.Close
    
  'Inserção dos Registros de Desconsto/Acréscimo
  'Sql = "SELECT * FROM ConsFINProcPagtoDescAcresc WHERE dupl_nr_fatura = " & NumFat & " and dupl_nr_duplicata = '" & XLT_PARCELA & "'"
  Sql = "SELECT * FROM ConsFINProcPagtoDescAcresc " & _
            " WHERE dupl_nr_fatura  >= " & CLng(XLT_INICIO) & " " & _
            " AND dupl_nr_fatura  <= " & CLng(XLT_FIM) & " " & _
            " AND dupl_nr_duplicata = '" & XLT_PARCELA & "'" & _
            " AND dupl_tx_debcred = 'D'"
  
  SubQOpenRecordset ResProc, Sql, Estatico
  Do While Not ResProc.EOF
    XFatura = ResProc!dupl_nr_fatura
    Xhistorico = FunNulo(ResProc!deac_tx_descricao)
    'Xvalor = ResProc!dude_vl_descacresc
    xValor = FunNulo(ResProc!DA)
    Sql = "INSERT INTO ProcPagto_Temp (prte_cd_Controle, prte_tx_descricao,prte_vl_valor,nofi_nr_fatura) " _
        & "Values('" & XControle & "','" & Xhistorico & "'," & FunTrataFloat(xValor) & "," & XFatura & ")"
    Conexao.Execute Sql
    ResProc.MoveNext
  Loop
  ResProc.Close
  
  'Inserção dos Registros de DescontosNF
  'Sql = "SELECT * FROM ConsFINProcPagtoDesc WHERE nofi_nr_fatura = " & NumFat
  Sql = "SELECT * FROM ConsFINProcPagtoDesc " & _
            " WHERE nofi_nr_fatura >= " & CLng(XLT_INICIO) & "" & _
            " AND nofi_nr_fatura <= " & CLng(XLT_FIM) & "" & _
            " AND nofi_tx_TipoMov = 'S'" & _
            " AND empr_cd_empresa = " & PCodEmpresa & ""
  SubQOpenRecordset ResProc, Sql, Estatico
  Do While Not ResProc.EOF
    XFatura = ResProc!nofi_nr_fatura
    Xhistorico = FunNulo(ResProc!desc_tx_descricao)
    xValor = FunNulo(ResProc!deen_vl_valor)
    Sql = "INSERT INTO ProcPagto_Temp (prte_cd_Controle, prte_tx_descricao,prte_vl_valor,nofi_nr_fatura) " _
        & "Values('" & XControle & "','" & Xhistorico & "'," & FunTrataFloat(xValor) & "," & XFatura & ")"
    Conexao.Execute Sql
    ResProc.MoveNext
  Loop
  ResProc.Close
  
  'Inserção dos registro de devolução
  'Sql = "SELECT * FROM ConsFINNotasFiscais WHERE nofi_cd_notaglosadev = " & Adodc1.Recordset.Fields("nofi_cd_notafiscal")
  'Não foi possível usar o sql abaixo porque o nº da fatura que vem do sql é da nf de
  'Devolução e não da nota origem
  'Sql = "SELECT * FROM ConsFINNotasFiscais WHERE nofi_cd_notaglosadev in " & _
    " (SELECT nofi_cd_notafiscal FROM ConsFINNotasFiscais WHERE nofi_nr_fatura >= " & CLng(TxtInicio.Text) & " AND nofi_nr_fatura <= " & CLng(TxtFim.Text) & ")"
  If ChkEstoque.Value = 1 Then 'Só trazer as notas do estoque
    Sql = "SELECT nofi_cd_notafiscal, nofi_nr_Fatura, nofi_tx_PendEstoque FROM ConsFINNotasFiscais " & _
           " WHERE nofi_nr_fatura >= " & CLng(XLT_INICIO) & " " & _
           " AND nofi_nr_fatura <= " & CLng(XLT_FIM) & "" & _
           " AND nofi_tx_TipoMov = 'S' AND nofi_tx_PendEstoque = 'N'" & _
           " AND empr_cd_empresa = " & PCodEmpresa & ""
  Else
    Sql = "SELECT nofi_cd_notafiscal, nofi_nr_Fatura, nofi_tx_PendEstoque FROM ConsFINNotasFiscais " & _
           " WHERE nofi_nr_fatura >= " & CLng(XLT_INICIO) & " " & _
           " AND nofi_nr_fatura <= " & CLng(XLT_FIM) & "" & _
           " AND nofi_tx_TipoMov = 'S'" & _
           " AND empr_cd_empresa = " & PCodEmpresa & ""
  End If
  
  SubQOpenRecordset ResNF, Sql, Estatico
  While Not ResNF.EOF
    Sql = "SELECT * FROM ConsFINNotasFiscais WHERE nofi_cd_notaglosadev = " & ResNF!nofi_cd_notafiscal
    SubQOpenRecordset ResProc, Sql, Estatico
    Do While Not ResProc.EOF
      XFatura = ResNF!nofi_nr_fatura
      Xhistorico = FunNulo(ResProc!nofi_tx_historico)
      xValor = FunNulo(ResProc!nofi_vl_valor)
      Sql = "INSERT INTO ProcPagto_Temp (prte_cd_Controle, prte_tx_descricao,prte_vl_valor,nofi_nr_fatura) " _
          & "Values('" & XControle & "','" & Xhistorico & "'," & FunTrataFloat(xValor) & "," & XFatura & ")"
      Conexao.Execute Sql
      ResProc.MoveNext
    Loop
    ResProc.Close
    ResNF.MoveNext
  Wend
  ResNF.Close
  
  'Verifica se Foi distribuído
  'SubQOpenRecordset ResNF, "SELECT nofi_tx_distribuir FROM NotasFiscais WHERE nofi_nr_Fatura = " & NumFat & "", Estatico
  If ChkEstoque.Value = 1 Then 'Só trazer as notas do estoque
    SubQOpenRecordset ResNF, "SELECT nofi_tx_distribuir,nofi_tx_tipofatura FROM NotasFiscais " & _
                            " WHERE nofi_nr_fatura >= " & CLng(XLT_INICIO) & " " & _
                            " AND nofi_nr_fatura <= " & CLng(XLT_FIM) & "" & _
                            " AND nofi_tx_TipoMov = 'S' AND nofi_tx_PendEstoque = 'N'" & _
                            " AND empr_cd_empresa = " & PCodEmpresa & "", Estatico
  Else
    SubQOpenRecordset ResNF, "SELECT nofi_tx_distribuir,nofi_tx_tipofatura FROM NotasFiscais " & _
                            " WHERE nofi_nr_fatura >= " & CLng(XLT_INICIO) & " " & _
                            " AND nofi_nr_fatura <= " & CLng(XLT_FIM) & "" & _
                            " AND nofi_tx_TipoMov = 'S'" & _
                            " AND empr_cd_empresa = " & PCodEmpresa & "", Estatico
  End If
    
  XLT_DISTRIBUIDO = "N"
  While Not ResNF.EOF
    If ResNF!nofi_tx_distribuir = "S" Or ResNF!nofi_tx_tipofatura = "3" Then
        XLT_DISTRIBUIDO = "S"
        ResNF.MoveLast
    End If
    ResNF.MoveNext
  Wend
  ResNF.Close
        
  'XGT_SELECAO = "ConsFINFluxoCaixa.dupl_nr_fatura = " & NumFat & ""
  If ChkEstoque.Value = 1 Then 'Só trazer as notas do estoque
    XGT_SELECAO = "ConsFINFluxoCaixa.dupl_nr_fatura >= " & CLng(XLT_INICIO) & "" & _
                    " AND ConsFINFluxoCaixa.dupl_nr_fatura <= " & CLng(XLT_FIM) & "" & _
                    " AND ConsFINFluxoCaixa.dupl_tx_DebCred = 'D' AND NotaEstoque = 'N' and dupl_dt_procpagamento is null" & _
                    " AND empr_cd_empresa = " & PCodEmpresa & ""
  Else
    XGT_SELECAO = "ConsFINFluxoCaixa.dupl_nr_fatura >= " & CLng(XLT_INICIO) & "" & _
                    " AND ConsFINFluxoCaixa.dupl_nr_fatura <= " & CLng(XLT_FIM) & "" & _
                    " AND ConsFINFluxoCaixa.dupl_tx_DebCred = 'D'" & _
                    " AND empr_cd_empresa = " & PCodEmpresa & ""
  End If
  
  'Abre os recordsets e verificar se eles possuem dados
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINFluxoCaixa WHERE " & XGT_SELECAO)
  Set XL0_RSSUBRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ProcPagto_Temp WHERE prte_cd_Controle='" & XControle & "'")
  If ChkEstoque.Value = 1 Then 'Só trazer as notas do estoque
      Set XL0_RSSUBRELATORIO2 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque " & _
                                                    " WHERE nofi_nr_fatura >= " & CLng(XLT_INICIO) & " " & _
                                                    " AND nofi_nr_fatura <= " & CLng(XLT_FIM) & "" & _
                                                    " AND nofi_tx_TipoMov = 'S' AND NotaEstoque = 'N' " & _
                                                    " AND empr_cd_empresa = " & PCodEmpresa & "")
  Else
      Set XL0_RSSUBRELATORIO2 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque " & _
                                                    " WHERE nofi_nr_fatura >= " & CLng(XLT_INICIO) & " " & _
                                                    " AND nofi_nr_fatura <= " & CLng(XLT_FIM) & "" & _
                                                    " AND nofi_tx_TipoMov = 'S' " & _
                                                    " AND empr_cd_empresa = " & PCodEmpresa & "")
  End If
  Set XL0_RSSUBRELATORIO3 = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDescAcresc")
  If XGO_RSRELATORIO.EOF Then
    MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
    Me.MousePointer = vbDefault
    Exit Sub
  End If
  
  'Abre o relatório
  FrmVisRelatorios.Caption = "Processo de Pagamento"
  Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\ProcessoPagtosequencia.rpt")
  
  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO, 3, 1
    
  '********** Ativa os Relatórios
  With XGR_RELATORIO.FormulaFields
    .GetItemByName("CCusto").Text = Chr(34) + FiltroCCusto + Chr(34)
    .GetItemByName("Empresa").Text = Chr(34) + PEmpresa + Chr(34)
  End With
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XL0_RSSUBRELATORIO, "ProcPagto_Temp"
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XL0_RSSUBRELATORIO2, "ConsFINDistribCopiaCheque"
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XL0_RSSUBRELATORIO3, "ConsFINDescAcresc"
  FrmVisRelatorios.SubVerRelatorioMoldal XGR_RELATORIO, Me
  
  '///////Impressão do Anexo do Processo (p/notas distribuidas)//////
  If XLT_DISTRIBUIDO = "S" Then
    Dim XLO_RSRELATORIOANEXO As New adodb.Recordset
    Dim XLR_RELATORIOANEXO As New CRAXDRT.Report
  
    'Set XLO_RSRELATORIOANEXO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque WHERE nofi_nr_fatura = " & Adodc1.Recordset.Fields("nofi_nr_fatura") & "")
    If ChkEstoque.Value = 1 Then 'Só trazer as notas do estoque
        Set XLO_RSRELATORIOANEXO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque " & _
                                                        " WHERE nofi_nr_fatura >= " & CLng(XLT_INICIO) & " " & _
                                                        " AND nofi_nr_fatura <= " & CLng(XLT_FIM) & " " & _
                                                        " AND nofi_tx_TipoMov = 'S' AND NotaEstoque = 'N'" & _
                                                        " AND nofi_tx_distribuir = 'S'" & _
                                                        " AND empr_cd_empresa = " & PCodEmpresa & "")
    Else
        Set XLO_RSRELATORIOANEXO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINDistribCopiaCheque " & _
                                                        " WHERE nofi_nr_fatura >= " & CLng(XLT_INICIO) & " " & _
                                                        " AND nofi_nr_fatura <= " & CLng(XLT_FIM) & " " & _
                                                        " AND nofi_tx_TipoMov = 'S'" & _
                                                        " AND nofi_tx_distribuir = 'S'" & _
                                                        " AND empr_cd_empresa = " & PCodEmpresa & "")
    End If
    
    'Abre o relatório
    FrmVisRelatorios_2.Caption = "Anexo do Processo de Pagamento"
    Set XLR_RELATORIOANEXO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\AnexoProcPagto.rpt")
  
    'Passa o recordset para o relatório
    XLR_RELATORIOANEXO.Database.SetDataSource XLO_RSRELATORIOANEXO, 3, 1
  
    '********** Ativa os Relatórios
    'With XLR_RELATORIOANEXO.FormulaFields
    '    .GetItemByName("NumProcesso").Text = Chr(34) + CStr(NumFat) + Chr(34)
    'End With
    '*FrmVisRelatorios_2.SubVerRelatorio XLR_RELATORIOANEXO, ""
    FrmVisRelatorios_2.SubVerRelatorioMoldal XLR_RELATORIOANEXO, Me
  End If
  '//////////////////////////////////////////////////////////////////
  
  FrmImpPPSequencia.MousePointer = vbDefault
    
'  'Limpa a tabela temporária
'  If XControle <> "" Then
'    Sql = "DELETE FROM ProcPagto_Temp WHERE prte_cd_controle = '" & XControle & "'"
'    Conexao.Execute Sql
'  End If
  XFormulario = "TelaNotasFiscais"
  '*************** Imprimir relação de notas da fatura *******************
  If TelaNotasFiscais.Adodc1.Recordset.Fields("nofi_tx_tipofatura") = "3" Then
    If MsgBox("Deseja imprimir a relação de Notas Fiscais desta Fatura?", vbQuestion + vbOKCancel, "ATENÇÃO") = vbOK Then
      Set XGO_RSRELATORIO = ConexaoRelatorio.Execute("SELECT * FROM ConsFINNotasFiscais where nofi_nr_fatura = " & NumFat & "")
      Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas")
      
      'Abre o relatório
      FrmVisRelatorios.Caption = "Relação de Notas da Fatura Nº " + CStr(NumFat)
      Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelNotasProcPagto.rpt")
      
      'Passa o recordset para o relatório
      XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO, 3, 1

      '********** Ativa os Relatórios
      XGR_RELATORIO.FormulaFields.GetItemByName("CodEmpresa").Text = Chr(34) & _
        CStr(PCodEmpresa) & Chr(34)
        FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
      FrmVisRelatorios.SubVerRelatorioMoldal XGR_RELATORIO, Me
      FrmImpPPSequencia.MousePointer = vbDefault
    Else
      Exit Sub
    End If
  End If
  FrmImpPPSequencia.Visible = False
End Sub

Private Sub TxtInicio_LostFocus()
    TxtFim.Text = TxtInicio.Text
End Sub

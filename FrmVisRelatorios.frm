VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmVisRelatorios 
   Caption         =   "FrmVisRelatorios"
   ClientHeight    =   5175
   ClientLeft      =   1500
   ClientTop       =   2700
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CrvRelatorio 
      Height          =   5130
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7410
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmVisRelatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CrvRelatorio_PrintButtonClicked(UseDefault As Boolean)

  UseDefault = False
  
  'aparece a tela de setup da impressora, onde é possível alterar a impressora e/ou a orientação do papel
   XGR_RELATORIO.PrinterSetup hwnd

  'em seguida o relatório é impresso na impressora configurada, mostrando a tela de confirmação para o usuário: True
  XGR_RELATORIO.PrintOut True

End Sub
'As opções de exportação já devem vir preenchidas
Public Sub SubExportarRelatorioSemPrompt(XLO_CRRELATORIO As CRAXDRT.Report)
  
  SubPreparaRelatorio XLO_CRRELATORIO, ""
  XLO_CRRELATORIO.Export False 'Não aparece o prompt
  
End Sub

Public Sub SubExportarRelatorioComPrompt(XLO_CRRELATORIO As CRAXDRT.Report)
  
  SubPreparaRelatorio XLO_CRRELATORIO, ""
  XLO_CRRELATORIO.Export True 'Não aparece o prompt
  
End Sub
Private Sub Form_Resize()
  CrvRelatorio.Top = 0
  CrvRelatorio.Left = 0
  CrvRelatorio.Height = ScaleHeight
  CrvRelatorio.Width = ScaleWidth
End Sub

Public Sub SubVerRelatorioSemMargem(XLO_CRRELATORIO As CRAXDRT.Report, XLT_SELECAO As String)
  
  Dim XLP_IMPRESSORA As Printer
  
  'Faz a seleção dos registros
  If XLT_SELECAO <> "" Then
    XLO_CRRELATORIO.RecordSelectionFormula = XLT_SELECAO
  End If
  
  'Configura a página de acordo com as configurações da impressora atual
  Set XLP_IMPRESSORA = Printer
  
  'Verifica as diferenças entre as tabelas do banco e do relatório e atualiza
  XLO_CRRELATORIO.Database.Verify
  
  'Coloca o relatório no componente para visualização
  CrvRelatorio.ReportSource = XLO_CRRELATORIO
  CrvRelatorio.ViewReport
  FrmVisRelatorios.Show
End Sub

Public Sub SubVerRelatorio(XLO_CRRELATORIO As CRAXDRT.Report, XLT_SELECAO As String)
  SubPreparaRelatorio XLO_CRRELATORIO, XLT_SELECAO
  
  'Coloca o relatório no componente para visualização
  CrvRelatorio.ReportSource = XLO_CRRELATORIO
  CrvRelatorio.ViewReport
  FrmVisRelatorios.Show
End Sub

Public Sub SubVerRelatorioMoldal(XLO_CRRELATORIO As CRAXDRT.Report, XLO_DONO As Form)
  SubPreparaRelatorio XLO_CRRELATORIO, ""
 
  'Coloca o relatório no componente para visualização
  CrvRelatorio.ReportSource = XLO_CRRELATORIO
  CrvRelatorio.ViewReport
  FrmVisRelatorios.Show vbModal, XLO_DONO
End Sub

Public Sub SubAlterarDataSource_Subrelatorios(XLO_CRRELATORIO As CRAXDRT.Report, _
  XLO_RSRELATORIO As ADODB.Recordset, XLT_NOMECONS As String)
    
  Dim XLO_CRSECTIONS As CRAXDRT.Sections
  Dim XLO_CRSECTION As CRAXDRT.Section
  Dim XLO_CRRELOBJ As CRAXDRT.ReportObjects
  Dim XLO_CRSUBOBJ As CRAXDRT.SubreportObject
  Dim XLO_CRSUBREPORT As CRAXDRT.Report
  Dim x As Integer
  Dim y As Integer
  
  Set XLO_CRSECTIONS = XLO_CRRELATORIO.Sections
  For x = 1 To XLO_CRSECTIONS.Count
    Set XLO_CRSECTION = XLO_CRSECTIONS.Item(x)
    Set XLO_CRRELOBJ = XLO_CRSECTION.ReportObjects
    For y = 1 To XLO_CRRELOBJ.Count
      If XLO_CRRELOBJ.Item(y).Kind = crSubreportObject Then
        Set XLO_CRSUBOBJ = XLO_CRRELOBJ.Item(y)
        Set XLO_CRSUBREPORT = XLO_CRSUBOBJ.OpenSubreport
        If UCase(Trim(XLO_CRSUBREPORT.Database.Tables(1).Name)) = UCase(Trim(XLT_NOMECONS)) Then
          XLO_CRSUBREPORT.Database.Tables.Item(1).SetDataSource XLO_RSRELATORIO
          XLO_CRSUBREPORT.Database.Verify
        End If
      End If
    Next
  Next
End Sub

Public Sub subAlterarTipoGrafico(XLO_RELATORIO As CRAXDRT.Report, _
  XLT_ALIASSUB As String, XLI_DISPLAYTYPE As Integer)

  Dim XLO_Sections As CRAXDRT.Sections
  Dim XLO_Section As CRAXDRT.Section
  Dim XLO_RelatorioObj As CRAXDRT.ReportObjects
  Dim XLO_SubreportObj As CRAXDRT.SubreportObject
  Dim XLO_Subreport As CRAXDRT.Report
  Dim XLO_SUB_Sections As CRAXDRT.Sections
  Dim XLO_SUB_Section As CRAXDRT.Section
  Dim XLO_SUB_RelatorioObj As CRAXDRT.ReportObjects
  Dim XLO_SUB_GraphObj As CRAXDRT.GraphObject
  Dim x As Integer, y As Integer, z As Integer, w As Integer
  
  Set XLO_Sections = XLO_RELATORIO.Sections
  For x = 1 To XLO_Sections.Count
    Set XLO_Section = XLO_Sections.Item(x)
    Set XLO_RelatorioObj = XLO_Section.ReportObjects
    For y = 1 To XLO_RelatorioObj.Count
      If XLO_RelatorioObj.Item(y).Kind = crSubreportObject Then
        Set XLO_SubreportObj = XLO_RelatorioObj.Item(y)
        Set XLO_Subreport = XLO_SubreportObj.OpenSubreport
        If UCase(Trim(XLO_Subreport.Database.Tables(1).Name)) = UCase(Trim(XLT_ALIASSUB)) Then
          Set XLO_SUB_Sections = XLO_Subreport.Sections
          For z = 1 To XLO_SUB_Sections.Count
            Set XLO_SUB_Section = XLO_SUB_Sections.Item(z)
            Set XLO_SUB_RelatorioObj = XLO_SUB_Section.ReportObjects
            For w = 1 To XLO_SUB_RelatorioObj.Count
              If XLO_SUB_RelatorioObj(w).Kind = crGraphObject Then
                Set XLO_SUB_GraphObj = XLO_SUB_RelatorioObj(w)
                XLO_SUB_GraphObj.GraphType = XLI_DISPLAYTYPE
              End If
            Next
          Next
        End If
      End If
    Next
  Next
End Sub

Private Sub SubPreparaRelatorio(XLO_CRRELATORIO As CRAXDRT.Report, XLT_SELECAO As String)
  Dim XLP_IMPRESSORA As Printer
  
  'Faz a seleção dos registros
  If XLT_SELECAO <> "" Then
    XLO_CRRELATORIO.RecordSelectionFormula = XLT_SELECAO
  End If
  
  'Configura a página de acordo com as configurações da impressora atual
  Set XLP_IMPRESSORA = Printer
  If XFormulario <> "RelDAM" Then
    If XLO_CRRELATORIO.PaperSize <> crPaperA4 Then
      If XLP_IMPRESSORA.PaperSize = vbPRPSA4 Then
        XLO_CRRELATORIO.PaperSize = crPaperA4
        'Corrigindo as margens para relatório feito em Carta
        'Paisagem
        If XLO_CRRELATORIO.PaperOrientation = crLandscape Then
          XLO_CRRELATORIO.LeftMargin = 1475
          XLO_CRRELATORIO.RightMargin = 1190
        'Retrato
        Else
          If XLO_CRRELATORIO.LeftMargin > 280 Then
            XLO_CRRELATORIO.LeftMargin = XLO_CRRELATORIO.LeftMargin - 280
          End If
          If (XLO_CRRELATORIO.RightMargin > 56) Then
            XLO_CRRELATORIO.RightMargin = XLO_CRRELATORIO.RightMargin - 56
          End If
        End If
      Else
        XLO_CRRELATORIO.PaperSize = crPaperLetter
      End If
    End If
  End If
  
  'Verifica as diferenças entre as tabelas do banco e do relatório e atualiza
  XLO_CRRELATORIO.Database.Verify
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If XGO_RSRELATORIO.State = adStateOpen Then
    XGO_RSRELATORIO.Close
    Set XGO_RSRELATORIO = Nothing
  End If

  If XGO_RSCABECALHO.State = adStateOpen Then
    XGO_RSCABECALHO.Close
    Set XGO_RSCABECALHO = Nothing
  End If

  Set XGR_RELATORIO = Nothing
  Set XGR_APLICACAO = Nothing
  
  If XFormulario = "TelaNotasFiscais" Then
    If XControle <> "" Then
      Sql = "DELETE FROM ProcPagto_Temp WHERE prte_cd_controle = '" & XControle & "'"
      Conexao.Execute Sql
    End If
  End If
End Sub



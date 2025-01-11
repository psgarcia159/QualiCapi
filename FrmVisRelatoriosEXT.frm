VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "CRVIEWER.DLL"
Begin VB.Form FrmVisRelatoriosEXT 
   Caption         =   "FrmVisRelatoriosEXT"
   ClientHeight    =   5175
   ClientLeft      =   1500
   ClientTop       =   2700
   ClientWidth     =   7455
   Icon            =   "FrmVisRelatoriosEXT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7455
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
      EnableRefreshButton=   -1  'True
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
Attribute VB_Name = "FrmVisRelatoriosEXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SubAlterarDataSource_Subrelatorios(XLR_RELATORIO As CRAXDRT.Report, _
                                              XLA_RECORDSET As ADODB.Recordset, _
                                              XLT_NOMECONS As String)
    Dim XLR_Sections As CRAXDRT.Sections
    Dim XLR_Section As CRAXDRT.Section
    Dim XGR_RelatorioObj As CRAXDRT.ReportObjects
    Dim XLR_SubreportObj As CRAXDRT.SubreportObject
    Dim XLR_Subreport As CRAXDRT.Report
    Dim x As Integer
    Dim y As Integer
       
    Set XLR_Sections = XLR_RELATORIO.Sections
    For x = 1 To XLR_Sections.Count
        Set XLR_Section = XLR_Sections.Item(x)
        Set XGR_RelatorioObj = XLR_Section.ReportObjects
        For y = 1 To XGR_RelatorioObj.Count
            If XGR_RelatorioObj.Item(y).Kind = crSubreportObject Then
                Set XLR_SubreportObj = XGR_RelatorioObj.Item(y)
                Set XLR_Subreport = XLR_SubreportObj.OpenSubreport
                If UCase(Trim(XLR_Subreport.Database.Tables(1).Name)) = UCase(Trim(XLT_NOMECONS)) Then
                    XLR_Subreport.Database.Tables.Item(1).SetDataSource XLA_RECORDSET
                    XLR_Subreport.Database.Verify
                End If
            End If
        Next
    Next
End Sub

Private Sub CrvRelatorio_PrintButtonClicked(UseDefault As Boolean)
  UseDefault = False
  
  'aparece a tela de setup da impressora, onde é possível alterar a impressora e/ou a orientação do papel
   XGR_RELATORIO.PrinterSetup hwnd

  'em seguida o relatório é impresso na impressora configurada, mostrando a tela de confirmação para o usuário: True
  XGR_RELATORIO.PrintOut True
End Sub

Private Sub Form_Resize()
    CrvRelatorio.Top = 0
    CrvRelatorio.Left = 0
    CrvRelatorio.Height = ScaleHeight
    CrvRelatorio.Width = ScaleWidth
End Sub

Public Sub SubVerRelatorio(XLR_RELATORIO As CRAXDRT.Report, XLT_SELECAO As String)
    Dim XLP_IMPRESSORA As Printer
    
    'Faz a seleção dos registros
    If XLT_SELECAO <> "" Then
        XLR_RELATORIO.RecordXGT_SELECAO = XLT_SELECAO
    End If

    'Configura a página de acordo com as configurações da impressora atual
    Set XLP_IMPRESSORA = Printer
    If XLP_IMPRESSORA.PaperSize = vbPRPSA4 Then
        XLR_RELATORIO.PaperSize = crPaperA4
        If XLR_RELATORIO.PaperOrientation = crLandscape Then
            XLR_RELATORIO.LeftMargin = XLR_RELATORIO.LeftMargin + 567
        Else
            XLR_RELATORIO.RightMargin = XLR_RELATORIO.RightMargin - 283
            XLR_RELATORIO.LeftMargin = XLR_RELATORIO.LeftMargin - 141
        End If
    Else
        XLR_RELATORIO.PaperSize = crPaperLetter
    End If
    
    'Coloca o relatório no componente para visualização
    XLR_RELATORIO.Database.Verify
    CrvRelatorio.ReportSource = XLR_RELATORIO
    CrvRelatorio.ViewReport
    FrmVisRelatoriosEXT.Show
End Sub


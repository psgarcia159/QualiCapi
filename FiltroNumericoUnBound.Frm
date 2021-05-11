VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FiltroNumericoUnBound 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtro"
   ClientHeight    =   3165
   ClientLeft      =   5850
   ClientTop       =   2685
   ClientWidth     =   5550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3165
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botao_conf 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Con&firma"
      Default         =   -1  'True
      Height          =   345
      Left            =   4155
      TabIndex        =   2
      Top             =   930
      Width           =   1215
   End
   Begin VB.CommandButton botao_canc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Desiste"
      Height          =   345
      Left            =   4140
      TabIndex        =   3
      Top             =   1410
      Width           =   1215
   End
   Begin VB.PictureBox SSFrame1 
      Height          =   2700
      Left            =   105
      ScaleHeight     =   2640
      ScaleWidth      =   3840
      TabIndex        =   13
      Top             =   420
      Width           =   3900
      Begin VB.OptionButton Option1 
         Caption         =   "Diferente de Nulo"
         Height          =   285
         Index           =   6
         Left            =   225
         TabIndex        =   10
         Top             =   2340
         Width           =   1905
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Igual a Nulo"
         Height          =   285
         Index           =   5
         Left            =   225
         TabIndex        =   9
         Top             =   2010
         Width           =   1740
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Intervalo"
         Height          =   285
         Index           =   4
         Left            =   225
         TabIndex        =   8
         Top             =   1680
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Menor"
         Height          =   285
         Index           =   3
         Left            =   225
         TabIndex        =   7
         Top             =   1350
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Maior"
         Height          =   285
         Index           =   2
         Left            =   225
         TabIndex        =   6
         Top             =   1020
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Diferente"
         Height          =   285
         Index           =   1
         Left            =   225
         TabIndex        =   5
         Top             =   690
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Igual"
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin MSMask.MaskEdBox text1 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   675
         Visible         =   0   'False
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox text2 
         Height          =   315
         Left            =   1575
         TabIndex        =   1
         Top             =   1575
         Visible         =   0   'False
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "à"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   1305
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2325
         TabIndex        =   14
         Top             =   345
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1260
      TabIndex        =   12
      Top             =   30
      Width           =   4095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Campo: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   105
      TabIndex        =   11
      Top             =   15
      Width           =   1155
   End
End
Attribute VB_Name = "FiltroNumericoUnBound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim controledadosfiltro As Object
'Dim GridFiltro As TDBGrid
Dim GridFiltro As TrueDBGrid70.Columns
Dim xColuna As Integer

Public Property Let ColunaGrid(ByVal Valor As Integer)
    xColuna = Valor
End Property

Public Property Let ControleDados(ByVal vNewValue As Variant)
    Set controledadosfiltro = vNewValue
End Property

Public Property Let Grid(ByVal vNewValue As Variant)
    Set GridFiltro = vNewValue
End Property

Private Sub botao_canc_Click()
    Filtrou = False
    Unload Me
End Sub

Private Sub botao_conf_Click()
    Dim Filtro, xcampo As String
    Dim a As Integer
    'IGUAL
    If Option1(0).Value = True Then
        If text1.ClipText <> "" Then
            'If TipoVar = "D" Then
            If XGT_CAMPOFILTRO = "D" Then
                Guarda_Filtro GridFiltro.Item(xColuna).DataField, "=", Format(text1.Text, "MM/DD/YY"), ""
            Else
                Guarda_Filtro GridFiltro.Item(xColuna).DataField, "=", FunTrataFloat(text1.Text), ""
            End If
        Else
            MsgBox "Valor deve ser informado.", 32, "ENTRADA INVÁLIDA"
            text1.SetFocus
            Exit Sub
        End If
    'DIFERENTE
    ElseIf Option1(1).Value = True Then
        If text1.ClipText <> "" Then
            'If TipoVar = "D" Then
            If XGT_CAMPOFILTRO = "D" Then
                Guarda_Filtro GridFiltro.Item(xColuna).DataField, "<>", Format(text1.Text, "MM/DD/YY"), ""
            Else
                Guarda_Filtro GridFiltro.Item(xColuna).DataField, "<>", FunTrataFloat(text1.Text), ""
            End If
        Else
            MsgBox "Valor deve ser informado.", 32, "ENTRADA INVÁLIDA"
            text1.SetFocus
            Exit Sub
        End If
    'MAIOR QUE
    ElseIf Option1(2).Value = True Then
        If text1.ClipText <> "" Then
            'If TipoVar = "D" Then
            If XGT_CAMPOFILTRO = "D" Then
                Guarda_Filtro GridFiltro.Item(xColuna).DataField, ">", Format(text1.Text, "MM/DD/YY"), ""
            Else
                Guarda_Filtro GridFiltro.Item(xColuna).DataField, ">", FunTrataFloat(text1.Text), ""
            End If
        Else
           MsgBox "Valor deve ser informado.", 32, "ENTRADA INVÁLIDA"
           text1.SetFocus
           Exit Sub
        End If
    'MENOR QUE
    ElseIf Option1(3).Value = True Then
        If text1.ClipText <> "" Then
            'If TipoVar = "D" Then
            If XGT_CAMPOFILTRO = "D" Then
                Guarda_Filtro GridFiltro.Item(xColuna).DataField, "<", Format(text1.Text, "MM/DD/YY"), ""
            Else
                Guarda_Filtro GridFiltro.Item(xColuna).DataField, "<", FunTrataFloat(text1.Text), ""
            End If
        Else
           MsgBox "Valor deve ser informado.", 32, "ENTRADA INVÁLIDA"
           text1.SetFocus
           Exit Sub
        End If
    'INTERVALO
    ElseIf Option1(4).Value = True Then
        If text1.ClipText = "" Then
           MsgBox "Valor inicial deve ser informado.", 32, "ENTRADA INVÁLIDA"
           text1.SetFocus
           Exit Sub
        End If
        If text2.ClipText = "" Then
           MsgBox "Valor final deve ser informado.", 32, "ENTRADA INVÁLIDA"
           text2.SetFocus
           Exit Sub
        End If
        'If TipoVar = "D" Then
        If XGT_CAMPOFILTRO = "D" Then
           If CVDate(text1) > CVDate(text2) Then
              MsgBox "Data final menor que a do inicial ", 48, "ENTRADA INVÁLIDA"
              text2.SetFocus
              Exit Sub
           End If
        ElseIf Val(text1) > Val(text2) Then
           MsgBox "Valor final menor que o inicial", 48, "ENTRADA INVÁLIDA"
           text2.SetFocus
           Exit Sub
        End If
        'If TipoVar = "D" Then
        If XGT_CAMPOFILTRO = "D" Then
            Guarda_Filtro GridFiltro.Item(xColuna).DataField, "><", Format(text1.Text, "MM/DD/YY"), Format(text2.Text, "MM/DD/YY")
        Else
            Guarda_Filtro GridFiltro.Item(xColuna).DataField, "><", FunTrataFloat(text1.Text), FunTrataFloat(text2.Text)
        End If
    'IGUAL NULO
    ElseIf Option1(5).Value = True Then
        Guarda_Filtro GridFiltro.Item(xColuna).DataField, "= Nulo", "", ""
    'DIFERENTE DE NULO
    ElseIf Option1(6).Value = True Then
        Guarda_Filtro GridFiltro.Item(xColuna).DataField, "<> Nulo", "", ""
    Else
        MsgBox "Informe o valor", 32, "ENTRADA INVÁLIDA"
        Exit Sub
    End If
    FiltroNumericoUnBound.MousePointer = vbHourglass
    Filtrou = True
    Filtro = FiltroOrig
    GridFiltro.Item(xColuna).HeadFont.Italic = True
    GridFiltro.Item(xColuna).HeadFont.Bold = True
    
    xcampo = Filtros(1, 1)
    
    If Tem_Filtro = False Then
       If Filtros(3, 1) = "1" Then    'Tipo String
          If filtra_especial Then
'              Modificado para dar certo com ADO
'              Filtro = Filtro + " AND " + xcampo + " LIKE """ + Filtros(2, 1) + "*"""
              Filtro = Filtro + " AND " + xcampo + " LIKE '" + Filtros(2, 1) + "%'"
          Else
'              Modificado para dar certo com ADO
'              Filtro = Filtro + " WHERE " + xcampo + " LIKE """ + Filtros(2, 1) + "*"""
              Filtro = Filtro + " WHERE " + xcampo + " LIKE '" + Filtros(2, 1) + "%'"
          End If
       ElseIf Filtros(3, 1) = "2" Then 'Tipo Data ou Número
          If filtra_especial Then
              Filtro = Filtro + " AND " + xcampo + Filtros(2, 1)
           Else
               Filtro = Filtro + " WHERE " + xcampo + Filtros(2, 1)
           End If
       End If
    Else
       If Filtros(3, 1) = "1" Then    'Tipo String
'          Modificado para dar certo com ADO
'          Filtro = Filtro + " AND " + xcampo + " LIKE """ + Filtros(2, 1) + "*"""
          Filtro = Filtro + " AND " + xcampo + " LIKE '" + Filtros(2, 1) + "%'"
       ElseIf Filtros(3, 1) = "2" Then 'Tipo Data ou Número
          Filtro = Filtro + " AND " + xcampo + Filtros(2, 1)
       End If
    End If
    For a = 2 To ContFil
        
        xcampo = Filtros(1, a)

        If Filtros(3, a) = "1" Then        'Tipo String
'           Modificado para dar certo com ADO
'           Filtro = Filtro + " AND " + xcampo + " LIKE """ + Filtros(2, a) + "*"""
           Filtro = Filtro + " AND " + xcampo + " LIKE '" + Filtros(2, a) + "%'"
        ElseIf Filtros(3, a) = "2" Then    'Tipo Data ou Número
           Filtro = Filtro + " AND " + xcampo + Filtros(2, a)
        End If
    Next
    
    FiltroAtual = Filtro
    
'    If ConTelaOrdem <> 0 Then
''       QRecordSource(controledadosfiltro) = Filtro + " ORDER BY " + Ordem
'        controledadosfiltro.RecordSource = Filtro + " ORDER BY " + Ordem
'    Else
''       QRecordSource(controledadosfiltro) = Filtro
'        controledadosfiltro.RecordSource = Filtro
'    End If
'
 '  controledadosfiltro.Refresh
   
   FiltroNumericoUnBound.MousePointer = vbDefault
   Unload Me

End Sub
Private Sub Form_Load()
    'If TipoVar = "D" Then 'Tipo data
    If XGT_CAMPOFILTRO = "D" Then  'Tipo data
       text1.Mask = "##/##/##"
       text2.Mask = "##/##/##"
       FiltroNumericoUnBound.Height = 3570
       SSFrame1.Height = 2700
       Option1(5).Enabled = True
       Option1(6).Enabled = True
    Else
       text1.Format = ""
       text2.Format = ""
       FiltroNumericoUnBound.Height = 3000
       SSFrame1.Height = 2100
       Option1(5).Enabled = False
       Option1(6).Enabled = False
    End If
    text1.Top = 405
    text1.Visible = True
    text1.Enabled = True
    
End Sub

Private Sub Guarda_Filtro(ByVal Campo As String, ByVal Tipo As String, ByVal Valor1 As String, ByVal Valor2 As String)
    Dim CampoTratado As String
    Dim Data As String
    Dim Data2 As String
    'Esta função armazena o novo filtro
    'Recebe: Campo - Nome do campo da tabela ativa
    '        Tipo  - "=", "<" , ">" , "<>" , "><". Menor , maior , diferente e intervalo respectivo
    '        Valor1 e Valor2 - Valores a serem filtrados
    If Campo = "CodGRD" Then
        CampoTratado = "GRD.CodGRD"
    ElseIf Campo = "CodDocumento" Then
        CampoTratado = "Documentos.CodDocumento"
    Else
        CampoTratado = Campo
    End If
    ContFil = ContFil + 1
    ReDim Preserve Filtros(3, ContFil)
    Filtros(1, ContFil) = CampoTratado
    Filtros(3, ContFil) = "2"
    'If TipoVar <> "D" Then
    If XGT_CAMPOFILTRO <> "D" Then
       If Tipo = "=" Then
          Filtros(2, ContFil) = " = " + Valor1
       ElseIf Tipo = "<" Then
          Filtros(2, ContFil) = " < " + Valor1
       ElseIf Tipo = ">" Then
          Filtros(2, ContFil) = " > " + Valor1
       ElseIf Tipo = "<>" Then
          Filtros(2, ContFil) = " <> " + Valor1
       Else
          Filtros(2, ContFil) = " >= " + Valor1 + " AND " + CampoTratado + " <= " + Valor2
       End If
    Else
       Valor1 = Mid$(Valor1, 4, 2) + "/" + Left$(Valor1, 2) + "/" + Right$(Valor1, 2)
        Data = FunNuloData(Valor1, NomeSgbd)
        If Tipo = "=" Then
           'Filtros(2, ContFil) = " = " & Data & ""
           Filtros(2, ContFil) = " >= " & funNuloDataHora(Valor1, NomeSgbd, "I") & " AND " & CampoTratado & " <= " & funNuloDataHora(Valor1, NomeSgbd, "F")
        ElseIf Tipo = "<" Then
           Filtros(2, ContFil) = " < " & Data & ""
        ElseIf Tipo = ">" Then
           Filtros(2, ContFil) = " > " & Data & ""
        ElseIf Tipo = "<>" Then
           'Filtros(2, ContFil) = " <> " & Data & ""
           Filtros(2, ContFil) = " < " & funNuloDataHora(Valor1, NomeSgbd, "I") & " OR " & CampoTratado & " > " & funNuloDataHora(Valor1, NomeSgbd, "F")
        ElseIf Tipo = "= Nulo" Then
           Filtros(2, ContFil) = " = Null "
        ElseIf Tipo = "<> Nulo" Then
           Filtros(2, ContFil) = " <> Null "
        Else
           Valor2 = Mid$(Valor2, 4, 2) + "/" + Left$(Valor2, 2) + "/" + Right$(Valor2, 2)
           Data2 = FunNuloData(Valor2, NomeSgbd)
           'Filtros(2, ContFil) = " >= #" + Valor1 + "#" + " AND " + CampoTratado + " <= #" + Valor2 + "#"
           Filtros(2, ContFil) = " >= " & Data & "" + " AND " + CampoTratado + " <= " & Data2 & ""
        End If
   End If
End Sub

Private Sub Option1_Click(Index As Integer)
    text2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    
    'If TipoVar = "D" Then
    If XGT_CAMPOFILTRO = "D" Then
       text1.Text = "  /  /  "
       text2.Text = "  /  /  "
    Else
       text1.Text = ""
       text2.Text = ""
    End If
    
    If Index = 5 Or Index = 6 Then
       text1.Visible = False
       Exit Sub
    Else
       text1.Visible = True
    End If
    
    'Posicionando os "Text box"
    If Index = 0 Then
       text1.Top = 390
    ElseIf Index = 1 Then
       text1.Top = 660
    ElseIf Index = 2 Then
       text1.Top = 1020
    ElseIf Index = 3 Then
       text1.Top = 1380
    ElseIf Index = 4 Then
       text1.Top = 630
       text2.Top = 1575
       Label3.Visible = True 'de:
       Label4.Visible = True 'até:
       text2.Visible = True
    End If
    text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    'If TipoVar <> "D" Then
    If XGT_CAMPOFILTRO <> "D" Then
      If FunDigito(KeyAscii) = False And KeyAscii <> 8 And KeyAscii <> 44 Then
         KeyAscii = 13
      End If
    End If
End Sub

Private Sub Text1_LostFocus()
    'If TipoVar = "D" And Text1.ClipText <> "" Then 'Se é data e não nula
    If XGT_CAMPOFILTRO = "D" And text1.ClipText <> "" Then  'Se é data e não nula
       If FunVerificaData(text1.Text) = False Then
          text1.SetFocus
       End If
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    'If TipoVar <> "D" Then
    If XGT_CAMPOFILTRO <> "D" Then
      If FunDigito(KeyAscii) = False And KeyAscii <> 8 Then
         KeyAscii = 13
      End If
    End If
End Sub

Private Sub Text2_LostFocus()
    'If TipoVar = "D" And text2.ClipText <> "" Then 'Se é data e não nulo
    If XGT_CAMPOFILTRO = "D" And text2.ClipText <> "" Then  'Se é data e não nulo
       If FunVerificaData(text2.Text) = False Then
          text2.SetFocus
       End If
    End If
End Sub


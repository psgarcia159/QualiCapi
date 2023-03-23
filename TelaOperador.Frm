VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaOperador 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operadores do Sistema"
   ClientHeight    =   6165
   ClientLeft      =   1515
   ClientTop       =   2745
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "TelaOperador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6165
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdVinculosCR 
      Caption         =   "&Vincular CR"
      Height          =   330
      Left            =   6840
      TabIndex        =   6
      Top             =   5700
      Width           =   1125
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Align           =   1  'Align Top
      Bindings        =   "TelaOperador.frx":08CA
      Height          =   5535
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   9763
      _LayoutType     =   4
      _RowHeight      =   19
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo"
      Columns(0).DataField=   "func_cd_operador"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nome"
      Columns(1).DataField=   "func_tx_nome"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Função"
      Columns(2).DataField=   "func_tx_funcao"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "E-mail"
      Columns(3).DataField=   "func_tx_Email"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1482"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1402"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=28021"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=4128"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4048"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=65535"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=5159"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=5080"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(2)._MinWidth=65535"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=5874"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=5794"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1,5
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=240,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Named:id=33:Normal"
      _StyleDefs(53)  =   ":id=33,.parent=0"
      _StyleDefs(54)  =   "Named:id=34:Heading"
      _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   ":id=34,.wraptext=-1"
      _StyleDefs(57)  =   "Named:id=35:Footing"
      _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   "Named:id=36:Selected"
      _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=37:Caption"
      _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(63)  =   "Named:id=38:HighlightRow"
      _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=39:EvenRow"
      _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(67)  =   "Named:id=40:OddRow"
      _StyleDefs(68)  =   ":id=40,.parent=33"
      _StyleDefs(69)  =   "Named:id=41:RecordSelector"
      _StyleDefs(70)  =   ":id=41,.parent=34"
      _StyleDefs(71)  =   "Named:id=42:FilterBar"
      _StyleDefs(72)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6240
      Top             =   3360
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton Botao_Hab 
      Caption         =   "&Habilitações"
      Height          =   330
      Left            =   5280
      TabIndex        =   3
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton Botao_Ins 
      Caption         =   "&Inserir"
      Height          =   330
      Left            =   480
      TabIndex        =   0
      Top             =   5700
      Width           =   1125
   End
   Begin VB.CommandButton Botao_Rem 
      Caption         =   "&Remover"
      Height          =   330
      Left            =   2040
      TabIndex        =   1
      Top             =   5700
      Width           =   1125
   End
   Begin VB.CommandButton Botao_Alt 
      Caption         =   "&Alterar"
      Height          =   330
      Left            =   3720
      TabIndex        =   2
      Top             =   5700
      Width           =   1125
   End
   Begin VB.CommandButton Botao_Ret 
      Caption         =   "&Retornar"
      Height          =   330
      Left            =   8520
      TabIndex        =   4
      Top             =   5700
      Width           =   1125
   End
End
Attribute VB_Name = "TelaOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim Inc As Integer         'Informa se está no modo de inclusão
'Dim Alt As Integer         'Informa se está no modo de alteração
'Dim Inc_NOK As Integer     'Variável auxiliar para inclusão de um novo reg.(Inclusão não OK)
'Dim Alt_NOK As Integer     'Variável auxiliar para alteração de um reg. (alteração não OK)
'Dim PosReg As Variant      'Posicão do registro atual
  
  Dim XLT_SQL As String
  Dim XLO_RECORDSET As ADODB.Recordset
Private Sub Botao_Alt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Edita as informações do operador selecionado"
End Sub

Private Sub Botao_Hab_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Edita as permissões do operador selecionado"
End Sub

Private Sub Botao_Ins_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Insere um novo operador no sistema"
End Sub

Private Sub Botao_Rem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Remove do sistema o operador selecionado"
End Sub

Private Sub Botao_Ret_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub

Private Sub Command1_Click()

End Sub


Private Sub CmdVinculosCR_Click()
  MdiPrincipal.MousePointer = vbHourglass
  TelaOperador.MousePointer = vbHourglass

  Chave = Adodc1.Recordset("func_cd_operador").Value
  TelaVinculosCR.Show 1

  MdiPrincipal.MousePointer = vbDefault
  TelaOperador.MousePointer = vbDefault
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  MdiPrincipal.BarraStatus.Panels(1).Text = ""
End Sub

Private Sub Form_Activate()

'If ModuloGenerico.getVersion = "Windows 7" Then
'    Adodc1.ConnectionString = Adodc1.ConnectionString & ";Password=qd;"
'End If
    
  Adodc1.Refresh
  Set Formulario = TelaOperador
End Sub

Private Sub Form_Load()
  
  Set Formulario = TelaOperador
  funAtualizaHabilitacoes
  subCarregaDadosNV Adodc1, "", "func_tx_Nome", "", "SELECT * FROM funcionarios"
  HabilitaBotoes
  
End Sub

Private Sub Botao_Alt_Click()
  MdiPrincipal.MousePointer = vbHourglass
  TelaOperador.MousePointer = vbHourglass

  XInserir = "E"
  Chave = Adodc1.Recordset("func_cd_operador").Value
  
  FrmAtualizaSenha.Show 1

  TDBGrid1.SetFocus
  MdiPrincipal.MousePointer = vbDefault
  TelaOperador.MousePointer = vbDefault
End Sub

Private Sub Botao_Hab_Click()
  MdiPrincipal.MousePointer = vbHourglass
  TelaOperador.MousePointer = vbHourglass

  Chave = Adodc1.Recordset("func_cd_operador").Value
  TelaHabilitacao.Show 1

  MdiPrincipal.MousePointer = vbDefault
  TelaOperador.MousePointer = vbDefault
End Sub

Private Sub Botao_Ins_Click()
  MdiPrincipal.MousePointer = vbHourglass
  TelaOperador.MousePointer = vbHourglass
  
  XInserir = "I"
  FrmAtualizaSenha.Show 1
  HabilitaBotoes
  TDBGrid1.SetFocus
  
  MdiPrincipal.MousePointer = vbDefault
  TelaOperador.MousePointer = vbDefault
End Sub

Private Sub Botao_Rem_Click()
    Dim Posicao As Variant

    Posicao = Adodc1.Recordset.bookmark
    
  If Adodc1.Recordset.Fields("func_cd_operador").Value = func_cd_operador Then
     MsgBox "Não é permitido Auto-Exclusão", vbCritical, "ATENÇÃO"
     Exit Sub
  End If
  
  If MsgBox("Confirma Remoção do Registro ?", vbQuestion + vbYesNo, "CUIDADO") = 6 Then 'Confirma exclusão
    Conexao.BeginTrans
    On Error GoTo RotuloErro
    Conexao.Execute ("DELETE From Habilitacoes WHERE func_cd_operador = " & Adodc1.Recordset.Fields("func_cd_operador"))
    Conexao.Execute ("DELETE From Funcionarios WHERE func_cd_operador = " & Adodc1.Recordset.Fields("func_cd_operador"))
    Conexao.CommitTrans
    Adodc1.Refresh
  End If
  HabilitaBotoes
  TDBGrid1.SetFocus
  TDBGrid1.Refresh
  TDBGrid1.Col = 0
  
  'Posicionando o cursor na linha anterior - Lucas Santiago 22.01.09
  If Posicao > 1 Then
      Adodc1.Recordset.bookmark = Posicao - 1
  End If
  
  Exit Sub
  
RotuloErro:
  Conexao.RollbackTrans
  MsgBox "Não é possível excluir este operador do sistema pois" & Chr(13) & _
    "ele possui registros relacionados com outras tabelas.", vbInformation, _
    "Exclusão"
End Sub

Private Sub Botao_Ret_Click()
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    MdiPrincipal.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub HabilitaBotoes()
    'Objetivo: Habilita os Botões da Tela
    Botao_Ins.Enabled = True
    If Not FunTabelaVazia(Adodc1) Then
        Botao_Rem.Enabled = True
        Botao_Hab.Enabled = True
        Botao_Alt.Enabled = True
    Else
        Botao_Rem.Enabled = False
        Botao_Hab.Enabled = False
        Botao_Alt.Enabled = False
    End If
    Botao_Ret.Enabled = True
End Sub

Private Sub Desabili_Botoes()
    'Objetivo: Desabilita os Botões da Tela
    Botao_Ins.Enabled = False
    Botao_Rem.Enabled = False
    Botao_Alt.Enabled = False
    Botao_Ret.Enabled = False
    Botao_Hab.Enabled = False
End Sub

Private Sub TDBGrid1_UnboundColumnFetch(bookmark As Variant, ByVal Col As Integer, Value As Variant)
  Dim Senha As String
  Dim sen As String
  Dim a As Integer

  Senha = TDBGrid1.Columns(5).Text
  For a = 1 To Len(Senha)
      sen = sen + Chr$(Asc(Mid$(Senha, a, 1)) - 10)
  Next
  Value = sen
End Sub

Private Function funAtualizaHabilitacoes() As Boolean
  Dim XLT_SQL As String

  'Insere na tabela de habilitações os registros de rotinas novas
  'para cada usuário, como "NÃO HABILITADO".
  On Error GoTo RotuloErro
  Conexao.Execute _
    "INSERT INTO Habilitacoes " & _
    "SELECT * " & _
    "FROM ConsGENNovasHabilitacoes_2 "
    
  funAtualizaHabilitacoes = True
  Exit Function

RotuloErro:
  funAtualizaHabilitacoes = False
End Function

VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "TDBG7.OCX"
Begin VB.Form TelaOrdem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Classificação da Tabela"
   ClientHeight    =   3360
   ClientLeft      =   2550
   ClientTop       =   1350
   ClientWidth     =   4905
   Icon            =   "TelaOrdem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Botao_Des 
      Caption         =   "&Desiste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3135
      TabIndex        =   3
      Top             =   2880
      Width           =   945
   End
   Begin VB.CommandButton Botao_Lim 
      Caption         =   "&Limpa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1995
      TabIndex        =   2
      Top             =   2880
      Width           =   945
   End
   Begin VB.CommandButton Botao_Con 
      Caption         =   "&Confirma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   1
      Top             =   2880
      Width           =   945
   End
   Begin TrueDBGrid70.TDBGrid TDBGrid1 
      Height          =   2610
      Left            =   285
      TabIndex        =   0
      Top             =   165
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   4604
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Campo"
      Columns(0).DataField=   ""
      Columns(0).DataWidth=   20
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Ordem"
      Columns(1).DataField=   ""
      Columns(1).DataWidth=   10
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "PosicaoGrid"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4286"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4180"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1048848"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2196"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2090"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1049120"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(1)._MinWidth=90965184"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=3307"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3201"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=1048848"
      Splits(0)._ColumnProps(19)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(21)=   "Column(2)._MinWidth=90964992"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=116,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
      _StyleDefs(9)   =   ":id=2,.fgcolor=&H80000012&"
      _StyleDefs(10)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(11)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(12)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(13)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(14)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(19)  =   "Splits(0).Style:id=43,.parent=1"
      _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(22)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(25)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
      _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
      _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
      _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
      _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=1"
      _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
      _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
      _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
      _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
      _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
      _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
      _StyleDefs(43)  =   "Named:id=33:Normal"
      _StyleDefs(44)  =   ":id=33,.parent=0"
      _StyleDefs(45)  =   "Named:id=34:Heading"
      _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   ":id=34,.wraptext=-1"
      _StyleDefs(48)  =   "Named:id=35:Footing"
      _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   "Named:id=36:Selected"
      _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(52)  =   "Named:id=37:Caption"
      _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(54)  =   "Named:id=38:HighlightRow"
      _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(56)  =   "Named:id=39:EvenRow"
      _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(58)  =   "Named:id=40:OddRow"
      _StyleDefs(59)  =   ":id=40,.parent=33"
      _StyleDefs(60)  =   "Named:id=41:RecordSelector"
      _StyleDefs(61)  =   ":id=41,.parent=34"
      _StyleDefs(62)  =   "Named:id=42:FilterBar"
      _StyleDefs(63)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "TelaOrdem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OrdemAux() As Integer
Dim ConTelaOrdemAux As Integer
Dim CurLin As Integer
Dim Form_ativ As String
Dim x As New XArray
Private Sub Arruma_Ordem()
    
    Dim row As Integer
    Dim Posicao As Integer
    
    Posicao = TDBGrid1.Bookmark
    
    For row = x.LowerBound(1) To x.UpperBound(1)
         If OrdemAux(row) <> 0 Then
           If OrdemAux(row) > OrdemAux(Posicao) Then
              x(row, 1) = Format$((OrdemAux(row) - 1), "00")
              OrdemAux(row) = OrdemAux(row) - 1
           End If
        End If
    Next row
    
    'TDBGrid1.Bookmark = Posicao
    x(Posicao, 1) = ""
    OrdemAux(Posicao) = 0
    ConTelaOrdemAux = ConTelaOrdemAux - 1
    TDBGrid1.Refresh
    TDBGrid1.EditActive = True
    TDBGrid1.EditActive = False
    TDBGrid1.SetFocus

End Sub
Private Sub Cria_Ordem()
    
    Dim Cont, b, a As Integer
    Dim xcampo As String
    
    Cont = 0
    Ordem = ""
    
    ReDim Ordens(ConTelaOrdemAux) As Integer
    
    For b = 1 To ConTelaOrdemAux
        
        For a = 0 To (TDBGrid1.ApproxCount - 1)
            
            If b = OrdemAux(a) Then
                
                Cont = Cont + 1
                Ordens(Cont) = a + 1
                If Formulario.TDBGrid1.Columns.Item(x(a, 2)).DataField = "CodDocumento" Then
                    xcampo = "Documentos.CodDocumento"
                ElseIf Formulario.TDBGrid1.Columns.Item(x(a, 2)).DataField = "CodGRD" Then
                    xcampo = "GRD.CodGRD"
                ElseIf Formulario.TDBGrid1.Columns.Item(x(a, 2)).DataField = "CodArea" Then
                    xcampo = "Documentos.CodArea"
                Else
                    xcampo = Formulario.TDBGrid1.Columns.Item(x(a, 2)).DataField
                End If
                Ordem = Ordem + xcampo + ","
                If b = 1 Then
                    xLocalDescricao = Formulario.TDBGrid1.Columns.Item(x(a, 2)).Caption
                    xLocalCampo = xcampo
                End If
                Exit For
            End If
        
        Next
        
    Next
    
    ConTelaOrdem = Cont
    
    If ConTelaOrdem <> 0 Then
       Ordem = Left$(Ordem, Len(Ordem) - 1)
    End If

End Sub
Function Val_Col(row As Long, Col As Integer, Indice As Long) As Variant
    
    If Col = 0 Then
        Val_Col = Formulario.TDBGrid1.Columns.Item(row).Caption
    ElseIf Col = 1 Then
     
        If OrdemAux(Indice) <> 0 Then
            Val_Col = Format$(OrdemAux(Indice), "00")
        End If
    Else
        Val_Col = row
    End If
        
End Function
Private Sub Botao_Con_Click()
    
    Dim Posicao As Integer
    Dim Fonte As String
    
    'Procura pelo ORDER BY no Qrecordsource(data1) do Formulario
    
    'Estava pegando o RecordSource no limite de 256 caracteres
    'Fonte = QRecordSource(Formulario.Data1)
    Fonte = FiltroAtual
    
    If ConTelaOrdem <> 0 Then
       
       Posicao = InStr(1, Fonte, " ORDER BY ")
       
       If Posicao <> 0 Then
          Fonte = Left$(Fonte, Posicao - 1)
       End If
    
    End If
    
    Cria_Ordem
    
    If ConTelaOrdem <> 0 Then
       SQLAtivo = Fonte + " ORDER BY " + Ordem
    Else
       SQLAtivo = Fonte
    End If
        
    subrecarregadadosNV Formulario.Adodc1, "", Ordem, "", Fonte
    
'Estava funcionando assim no financeiro
'    If ConTelaOrdem <> 0 Then
'       SQLAtivo = Fonte + " ORDER BY " + Ordem
'       Formulario.Adodc1.RecordSource = Fonte + " ORDER BY " + Ordem
'    Else
'       SQLAtivo = Fonte
'       Formulario.Adodc1.RecordSource = Fonte
'    End If
'
'    Formulario.TDBGrid1.Refresh
'    Formulario.Adodc1.Refresh
    
    Unload Me

End Sub
Private Sub Botao_Des_Click()
    Unload Me
End Sub
Private Sub Botao_Lim_Click()
    
    Dim row As Long

    ' The LowerBound and UpperBound properties correspond
    ' to the LBound and UBound functions in Visual Basic.
    ' Hard-coded dimensions can be used instead, if known.
    
    For row = x.LowerBound(1) To x.UpperBound(1)
            x(row, 1) = ""
            OrdemAux(row) = 0
    Next row
    
    ConTelaOrdemAux = 0
    TDBGrid1.row = 0
    TDBGrid1.Refresh
    TDBGrid1.EditActive = True
    TDBGrid1.EditActive = False
    TDBGrid1.SetFocus
    Botao_Con.Enabled = False

End Sub
Private Sub Form_Load()
    Dim MaxRows As Integer
    Dim a As Long
    
    ' Allocate space for rows and columns
    
    MaxRows = Formulario.TDBGrid1.Columns.Count
    For a = 0 To Formulario.TDBGrid1.Columns.Count - 1
        If Not Formulario.TDBGrid1.Columns.Item(a).Visible Then
            MaxRows = MaxRows - 1
        End If
    Next
    x.ReDim 0, MaxRows - 1, 0, 2

    ConTelaOrdemAux = ConTelaOrdem
    
    ReDim OrdemAux(MaxRows) As Integer
    
    For a = 1 To ConTelaOrdem
        OrdemAux(Ordens(a) - 1) = a
    Next

    Dim row As Long

    For row = x.LowerBound(1) To x.UpperBound(1)
            x(row, 0) = Val_Col(row, 0, row)
            x(row, 1) = Val_Col(row, 1, row)
            x(row, 2) = Val_Col(row, 2, row)
    Next row
    row = 0
    
    For a = 0 To Formulario.TDBGrid1.Columns.Count - 1
        If Formulario.TDBGrid1.Columns.Item(a).Visible Then
            x(row, 0) = Val_Col(a, 0, row)
            x(row, 1) = Val_Col(a, 1, row)
            x(row, 2) = Val_Col(a, 2, row)
            row = row + 1
        End If
    Next

    ' Bind True DBGrid Control to this XArray instance
    Set TDBGrid1.Array = x
    'TDBGrid1.Refresh
    'TDBGrid1.EditActive = True
    'TDBGrid1.EditActive = False
    'TDBGrid1.SetFocus
      
End Sub
Private Sub TDBGrid1_DblClick()
    Botao_Con.Enabled = True
    If funTipoCampo(Formulario.Adodc1, Formulario.TDBGrid1.Columns, TDBGrid1.Bookmark) <> "M" Then
        If CurLin >= 0 Then
           
           If OrdemAux(TDBGrid1.Bookmark) = 0 Then
              ConTelaOrdemAux = ConTelaOrdemAux + 1
              x(TDBGrid1.Bookmark, 1) = Format$(ConTelaOrdemAux, "00")
              OrdemAux(TDBGrid1.Bookmark) = ConTelaOrdemAux
           Else
              Arruma_Ordem
           End If
           TDBGrid1.Refresh
           TDBGrid1.EditActive = True
           TDBGrid1.EditActive = False
           TDBGrid1.SetFocus
        
        End If
    Else
        MsgBox "A ordenação não pode ser efetuada por uma campo MEMO", 16
    End If


End Sub
Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 1 Or Button = 2 Then
       
       'CurLin é a posição da linha visível na tela
       'a partir da segunda tela não coincide com bookmark
       CurLin = TDBGrid1.RowContaining(y)
       
       If CurLin >= 0 Then
           TDBGrid1.row = CurLin
       End If
    
    End If

End Sub


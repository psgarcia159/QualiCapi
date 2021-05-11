VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form TelaFiltro 
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   1425
   ClientTop       =   2130
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3046.69
   ScaleMode       =   0  'User
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnRetonar 
      Caption         =   "&Retornar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3930
      TabIndex        =   2
      Top             =   2685
      Width           =   1215
   End
   Begin VB.CommandButton btnRemover 
      Caption         =   "Re&mover"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2430
      TabIndex        =   1
      Top             =   2700
      Width           =   1215
   End
   Begin TrueDBGrid70.TDBGrid TDBGrid1 
      Height          =   2250
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   3969
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
      Columns(1).Caption=   "Seleção"
      Columns(1).DataField=   ""
      Columns(1).DataWidth=   50
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=74000"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=3254"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=74000"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000012&,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(25)  =   "Splits(0).Style:id=43,.parent=1"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(38)  =   ":id=28,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(43)  =   ":id=32,.locked=-1"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
      _StyleDefs(47)  =   "Named:id=33:Normal"
      _StyleDefs(48)  =   ":id=33,.parent=0"
      _StyleDefs(49)  =   "Named:id=34:Heading"
      _StyleDefs(50)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   ":id=34,.wraptext=-1"
      _StyleDefs(52)  =   "Named:id=35:Footing"
      _StyleDefs(53)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(54)  =   "Named:id=36:Selected"
      _StyleDefs(55)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(56)  =   "Named:id=37:Caption"
      _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(58)  =   "Named:id=38:HighlightRow"
      _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(60)  =   "Named:id=39:EvenRow"
      _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(62)  =   "Named:id=40:OddRow"
      _StyleDefs(63)  =   ":id=40,.parent=33"
      _StyleDefs(64)  =   "Named:id=41:RecordSelector"
      _StyleDefs(65)  =   ":id=41,.parent=34"
      _StyleDefs(66)  =   "Named:id=42:FilterBar"
      _StyleDefs(67)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "TelaFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As New XArray
Dim xGrid As Object

Function funTrataCabecalho(XVariavel As String) As String
    Dim Cont As Integer
    Dim XResult As String
    
    Cont = 1
    While (Cont <= Len(Trim(XVariavel)))
        If (Mid(XVariavel, Cont, 1) <> "*") And (Mid(XVariavel, Cont, 1) <> " ") Then
            XResult = XResult + Mid(XVariavel, Cont, 1)
        End If
        Cont = Cont + 1
    Wend
    
    funTrataCabecalho = XResult
    
End Function
Private Sub btnRemover_Click()
    Dim a, Col As Integer
    Dim Filtro As String

    If Filtros(1, TDBGrid1.row + 1) <> "" Then
        Filtrou = True
        'Procura o Campo da Tabela TDBGrid1 do FormularioAtual
        Col = Procura_Coluna(Filtros(1, TDBGrid1.row + 1))
        xGrid.Columns.Item(Col).HeadFont.Italic = False
        xGrid.Columns.Item(Col).HeadFont.Bold = False
        xGrid.Columns.Item(Col).HeadFont.Bold = True

        Remove_Filtro (TDBGrid1.row + 1)
'        TDBGrid1.ApproxCount = TDBGrid1.ApproxCount - 1
        TDBGrid1.Delete
        ContFil = ContFil - 1
        Filtro = FiltroOrig
        If ContFil <> 0 Then
            If Tem_Filtro = False Then
                If Filtros(3, 1) = "1" Then 'Tipo string
                    If filtra_especial Then
'                        Modificado para funcionar com ADO
'                        Filtro = Filtro + " AND " + Filtros(1, 1) + " LIKE '" + Filtros(2, 1) + "*'"
                        Filtro = Filtro + " AND " + Filtros(1, 1) + " LIKE '" + Filtros(2, 1) + "%'"
                    Else
'                        Modificado para funcionar com ADO
'                        Filtro = Filtro + " WHERE " + Filtros(1, 1) + " LIKE '" + Filtros(2, 1) + "*'"
                        Filtro = Filtro + " WHERE " + Filtros(1, 1) + " LIKE '" + Filtros(2, 1) + "%'"
                    End If
                ElseIf Filtros(3, 1) = "2" Then
                    If filtra_especial Then
                        Filtro = Filtro + " AND " + Filtros(1, 1) + Filtros(2, 1)
                    Else
                        Filtro = Filtro + " WHERE " + Filtros(1, 1) + Filtros(2, 1)
                    End If
                End If
            Else
                If Filtros(3, 1) = "1" Then 'Tipo string
'                    Modificado para funcionar com ADO
'                    Filtro = Filtro + " AND " + Filtros(1, 1) + " LIKE '" + Filtros(2, 1) + "*'"
                    Filtro = Filtro + " AND " + Filtros(1, 1) + " LIKE '" + Filtros(2, 1) + "%'"
                ElseIf Filtros(3, 1) = "2" Then
                    Filtro = Filtro + " AND " + Filtros(1, 1) + Filtros(2, 1)
                End If
            End If
            For a = 2 To ContFil
                If Filtros(3, a) = "1" Then 'Tipo string
'                    Modificado para funcionar com ADO
'                    Filtro = Filtro + " AND " + Filtros(1, a) + " LIKE '" + Filtros(2, a) + "*'"
                    Filtro = Filtro + " AND " + Filtros(1, a) + " LIKE '" + Filtros(2, a) + "%'"
                ElseIf Filtros(3, a) = "2" Then
                    Filtro = Filtro + " AND " + Filtros(1, a) + Filtros(2, a)
                End If
            Next
            Else
                btnRemover.Enabled = False
        End If
'        If ConTelaOrdem <> 0 Then
'             Formulario.Adodc1.RecordSource = Filtro + " ORDER BY " + Ordem
'        Else
'             Formulario.Adodc1.RecordSource = Filtro
'        End If
            
        FiltroAtual = Filtro
        
        If TDBGrid1.ApproxCount <> 0 Then
            TDBGrid1.row = 0
        End If
        
        TDBGrid1.Refresh
        TDBGrid1.EditActive = True
        TDBGrid1.EditActive = False
        TDBGrid1.SetFocus
    End If
End Sub

Private Sub btnRetornar_Click()
    TelaFiltro.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub Form_Load() 'Definição da tabela
    Dim MaxRows As Integer
    ' Allocate space for rows and columns
    
    If Formulario.Name = "TelaGeral" Then
        'Set XGrid = Formulario.TdbGrid1(XTela)
    Else
        Set xGrid = Formulario.TDBGrid1
    End If
    TelaFiltro.MousePointer = vbDefault
    MaxRows = ContFil
    X.ReDim 0, MaxRows - 1, 0, 1
    
    Dim row As Integer

    For row = X.LowerBound(1) To X.UpperBound(1)
            X(row, 0) = Val_Col(row, 0)
            X(row, 1) = Val_Col(row, 1)
    Next row

    ' Bind True DBGrid Control to this XArray instance
    Set TDBGrid1.Array = X


End Sub

Private Function Procura_Coluna(ByVal txt As String) As Integer
    Dim i As Integer
    Dim CampoTratado As String
    Procura_Coluna = -1
    txt = UCase$(txt)
    CampoTratado = txt
    
    ' Procura pelo nome do Campo
    For i = 0 To xGrid.Columns.Count - 1
        If CampoTratado = UCase$(xGrid.Columns.Item(i).DataField) Then
            Procura_Coluna = i
            Exit Function
        End If
    Next
    'Procura pelo Cabeçalho na Tabela
    For i = 0 To xGrid.Columns.Count - 1
        If CampoTratado = UCase$(xGrid.Columns.Item(i).Caption) Then
            Procura_Coluna = i
            Exit Function
        End If
    Next
End Function

Private Sub Remove_Filtro(ByVal Posicao As Integer)
    Dim a As Integer
    For a = Posicao To ContFil - 1
    'Reposiciona os filtros depois que remove o ultimo
        Filtros(1, a) = Filtros(1, a + 1)
        Filtros(2, a) = Filtros(2, a + 1)
        Filtros(3, a) = Filtros(3, a + 1)
    Next
    ReDim Preserve Filtros(3, (ContFil - 1))
End Sub


Function Val_Col(row As Integer, Col As Integer) As Variant

Dim Texto As String
Dim Intervalo, Posicao, coluna As Integer
Dim X, y As Integer
X = row + 1
y = Col + 1
  
    If Col = 1 Then    'referente a segunda coluna
       If Filtros(3, X) = "1" Then 'Tipo string
          Val_Col = Filtros(y, X)
       Else
          Texto = Filtros(2, X)
          Posicao = InStr(1, Texto, "AND") 'Guarda a posição do AND
          If Posicao <> 0 Then    'Troca AND por e
             Texto = Left$(Texto, Posicao - 1) + "e " + Right$(Texto, Len(Texto) - Posicao - 3)
             Posicao = InStr(1, Texto, Filtros(1, row + 1))
             If Posicao <> 0 Then
                Texto = Left$(Texto, Posicao - 1) + Right$(Texto, Len(Texto) - Posicao - Len(Filtros(1, X)))
             End If
          End If
          Posicao = InStr(1, Texto, "#")
          If Posicao <> 0 Then      'Inverte Dia com Mes
             Posicao = InStr(1, Texto, "=")
             If Posicao <> 0 Then
                Intervalo = InStr(1, Texto, "<")
                If Intervalo <> 0 Then
                   Texto = Left$(Texto, 5) + Mid$(Texto, 9, 2) + "/" + Mid$(Texto, 6, 2) + Mid$(Texto, 11, 11) + Mid$(Texto, 25, 2) + "/" + Mid$(Texto, 22, 2) + Right$(Texto, 4)
                Else
                   Texto = Left$(Texto, 4) + Mid$(Texto, 8, 2) + "/" + Mid$(Texto, 5, 2) + Right$(Texto, 4)
                End If
             Else
                Texto = Left$(Texto, 4) + Mid$(Texto, 8, 2) + "/" + Mid$(Texto, 5, 2) + Right$(Texto, 4)
             End If
          End If
          Posicao = InStr(1, Texto, "#")
          Do While Posicao <> 0     'Tira os # da Data
             If Posicao = Len(Texto) Then
                Texto = Left$(Texto, Len(Texto) - 1)
                Exit Do
             Else
                Texto = Left$(Texto, Posicao - 1) + Mid$(Texto, Posicao + 1, Len(Texto) - Posicao)
             End If
             Posicao = InStr(1, Texto, "#")
          Loop
          Val_Col = Texto
       End If
    Else
       coluna = Procura_Coluna(Filtros(1, X))
       If coluna >= 0 Then
          Val_Col = xGrid.Columns.Item(coluna).Caption
       Else
          Val_Col = Filtros(1, X)
       End If
    End If
End Function

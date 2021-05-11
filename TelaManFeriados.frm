VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TelaManFeriados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabela de Feriados"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   LinkTopic       =   "TelaManFeriados"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4080
      TabIndex        =   12
      Top             =   5880
      Width           =   765
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6645
      TabIndex        =   13
      Top             =   5880
      Width           =   765
   End
   Begin Threed.SSPanel PanCotacao 
      Height          =   1590
      Left            =   960
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   5730
      _Version        =   65536
      _ExtentX        =   10107
      _ExtentY        =   2805
      _StockProps     =   15
      Caption         =   "Feriado"
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.TextBox txtFeriado 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   840
         Width           =   4695
      End
      Begin VB.CommandButton CmdConfirma 
         BackColor       =   &H00000000&
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
         Height          =   270
         Left            =   3615
         TabIndex        =   8
         Top             =   1230
         Width           =   900
      End
      Begin VB.CommandButton CmdRetorna 
         BackColor       =   &H00000000&
         Caption         =   "&Retorna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4590
         TabIndex        =   9
         Top             =   1230
         Width           =   900
      End
      Begin MSComCtl2.DTPicker dtpData 
         Height          =   315
         Left            =   840
         TabIndex        =   4
         Top             =   375
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   107151363
         CurrentDate     =   40179
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Feriado:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data:"
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   405
         Width           =   390
      End
   End
   Begin VB.CommandButton CmdInserir 
      BackColor       =   &H00000000&
      Caption         =   "&Inserir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   10
      Top             =   5880
      Width           =   765
   End
   Begin VB.CommandButton CmdAlterar 
      BackColor       =   &H00000000&
      Caption         =   "&Alterar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3120
      TabIndex        =   11
      Top             =   5880
      Width           =   765
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid 
      Bindings        =   "TelaManFeriados.frx":0000
      Height          =   5040
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8890
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Data"
      Columns(0).DataField=   "feri_dt_data"
      Columns(0).NumberFormat=   "dd/mm/yyyy"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Feriado"
      Columns(1).DataField=   "feri_tx_descricao"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=17"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=56398032"
      Splits(0)._ColumnProps(10)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=20"
      Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(9)   =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(10)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin MSComCtl2.DTPicker dtpAno 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy"
      Format          =   107151363
      CurrentDate     =   40179
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   4320
      Top             =   240
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   ""
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Ano:"
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   270
      Width           =   390
   End
End
Attribute VB_Name = "TelaManFeriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function FunExisteData(Data As String) As Boolean
    
    
   On Error GoTo FunExisteData_Error
   
    Dim XLO_Data As ADODB.Recordset

    SubQOpenRecordset XLO_Data, "SELECT * FROM Feriados WHERE feri_dt_data = " & FunNuloData(Data, NomeSgbd) & "", Estatico
    
    If XLO_Data.EOF And XLO_Data.BOF Then
        XLO_Data.Close
        Set XLO_Data = Nothing
        FunExisteData = False
    Else
        XLO_Data.Close
        Set XLO_Data = Nothing
        FunExisteData = True
    End If
    

   On Error GoTo 0
   Exit Function

FunExisteData_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FunExisteData of Formulário TelaManFeriados"
    
End Function


Private Sub CmdAlterar_Click()
    
    On Error GoTo CmdAlterar_Click_Error
    
        PanCotacao.Tag = "A"
        PanCotacao.Enabled = True
        PanCotacao.Visible = True
        TDBGrid.Enabled = False
        
        DtpData.Value = Format$(Adodc.Recordset.Fields("feri_dt_data"), "dd/MM/yyyy")
        DtpData.Enabled = False
        txtFeriado.Text = Adodc.Recordset.Fields("feri_tx_descricao")
    
    On Error GoTo 0
    Exit Sub

CmdAlterar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CmdAlterar_Click of Formulário TelaManFeriados"
End Sub

Private Sub CmdConfirma_Click()

    On Error GoTo CmdConfirma_Click_Error
    
        If txtFeriado.Text = "" Then
            MsgBox "Nome do feriado é obrigatório", vbCritical, "Atenção!"
            txtFeriado.SetFocus
            Exit Sub
        End If
        
        If PanCotacao.Tag = "I" Then
            'Verifica se já foi realizada uma cotação da moeda para uma determinada data
            If FunExisteData(Format(DtpData.Value, "dd/mm/yyyy")) Then
                MsgBox "Já existe feriado para esta data.", vbInformation, "ATENÇÃO!"
                DtpData.SetFocus
                Exit Sub
            End If
            
            Conexao.Execute "INSERT INTO Feriados (feri_dt_data, feri_tx_descricao) VALUES (" & FunNuloData(DtpData.Value, NomeSgbd) & ", '" & FunNulo(txtFeriado.Text) & "')"
            
            dtpAno.Value = DtpData.Value
        Else
            Conexao.Execute "UPDATE Feriados SET feri_tx_descricao = '" & FunNulo(txtFeriado.Text) & "' WHERE feri_dt_data = " & FunNuloData(DtpData.Value, NomeSgbd)
        End If
        
        PanCotacao.Enabled = False
        PanCotacao.Visible = False
        TDBGrid.Enabled = True
        
        txtFeriado.Text = ""
        
        subCarregaDadosNV Adodc, "Feriados", "feri_dt_data ASC", "", "SELECT feri_dt_data, feri_tx_descricao FROM Feriados WHERE YEAR(feri_dt_data) = " & CStr(dtpAno.Year)
        CmdExcluir.Enabled = (Adodc.Recordset.RecordCount > 0)
        CmdAlterar.Enabled = (Adodc.Recordset.RecordCount > 0)
    
    On Error GoTo 0
    Exit Sub

CmdConfirma_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CmdConfirma_Click of Formulário TelaManFeriados"
End Sub

Private Sub CmdExcluir_Click()

   On Error GoTo CmdExcluir_Click_Error
        
        If MsgBox("Confirma Remoção do Registro ?", vbCritical + vbYesNo, "CUIDADO") = vbNo Then Exit Sub
        
        Conexao.Execute "DELETE Feriados WHERE feri_dt_data = " & FunNuloData(Adodc.Recordset.Fields("feri_dt_data"), NomeSgbd)
        
        subCarregaDadosNV Adodc, "Feriados", "feri_dt_data ASC", "", "SELECT feri_dt_data, feri_tx_descricao FROM Feriados WHERE YEAR(feri_dt_data) = " & CStr(dtpAno.Year)
        CmdExcluir.Enabled = (Adodc.Recordset.RecordCount > 0)
        CmdAlterar.Enabled = (Adodc.Recordset.RecordCount > 0)

   On Error GoTo 0
   Exit Sub

CmdExcluir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CmdExcluir_Click of Formulário TelaManFeriados"
End Sub

Private Sub CmdInserir_Click()
    PanCotacao.Tag = "I"
    PanCotacao.Enabled = True
    PanCotacao.Visible = True
    TDBGrid.Enabled = False
    DtpData.Value = dtpAno.Value
    DtpData.Enabled = True

End Sub

Private Sub CmdRetorna_Click()
    PanCotacao.Enabled = False
    PanCotacao.Visible = False
    
    txtFeriado.Text = ""
     
    TDBGrid.Enabled = True
    TDBGrid.SetFocus
    
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub dtpAno_Change()

    On Error GoTo dtpAno_Change_Error
        
        subCarregaDadosNV Adodc, "Feriados", "feri_dt_data ASC", "", "SELECT feri_dt_data, feri_tx_descricao FROM Feriados WHERE YEAR(feri_dt_data) = " & CStr(dtpAno.Year)
        CmdExcluir.Enabled = (Adodc.Recordset.RecordCount > 0)
        CmdAlterar.Enabled = (Adodc.Recordset.RecordCount > 0)
        
    On Error GoTo 0
    Exit Sub

dtpAno_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure dtpAno_Change of Formulário TelaManFeriados"
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Error
        
        dtpAno.Value = Now
        
        subCarregaDadosNV Adodc, "Feriados", "feri_dt_data ASC", "", "SELECT feri_dt_data, feri_tx_descricao FROM Feriados WHERE YEAR(feri_dt_data) = " & CStr(dtpAno.Year)
        CmdExcluir.Enabled = (Adodc.Recordset.RecordCount > 0)
        CmdAlterar.Enabled = (Adodc.Recordset.RecordCount > 0)
        
    On Error GoTo 0
    Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulário TelaManFeriados"
End Sub

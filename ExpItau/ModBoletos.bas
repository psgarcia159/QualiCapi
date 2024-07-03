Attribute VB_Name = "ModBoletos"
' ------------------------------------------------------------------------------
' Declaração de objetos, tipos e funções, utilizados para a emissão de boletos
' Obs.: Os requisito para uso da biblioteca Boleto2Net como COM devem ser vistos
' e reavaliados (estão comentados por causarem crash no IDE do Visual Basic 6.0)
' ------------------------------------------------------------------------------
Option Explicit

Global XLO_BOLETOS      As Object           ' Objeto Json com o conteúdo do arquivo de configurações
Global XLI_EXPIRESIN    As Integer          ' Variável para controle de tempo de duração do access_token
Global XLD_DATETOKEN    As Date             ' Variável que contém a data/hora do acesso ao token
Global XLT_ACCESSTOKEN  As String           ' String contendo o último access_token válido
Global XLT_TOKENTYPE    As String           ' Tipo do Token, geralmente "Bearer"
Global XLT_TOKENSCOPE   As String           ' Escopo de autenticação/autorização do token

' - Declaração de Function's e Sub's
' ----------------------------------------------------------------------------

' - Função para montagem de string Json ("parametro": "valor")
Public Function FunJsonString(Parametro As String, valor As String, Optional IsString As Boolean = True) As String
    Dim OutputStr As String
    If IsString Then
        OutputStr = """@Parameter"": ""@Value"""
    Else
        OutputStr = """@Parameter"": @Value"
    End If
    FunJsonString = Replace(Replace(OutputStr, "@Parameter", Parametro), "@Value", valor)
End Function

' - Função para validar e-mail do pagador
Public Function FunValidaEmail(email As String) As Boolean
    Dim At As Integer
    Dim oneDot As Integer
    Dim twoDots As Integer
 
    FunValidaEmail = True
    At = InStr(1, email, "@", vbTextCompare)
    oneDot = InStr(At + 2, email, ".", vbTextCompare)
    twoDots = InStr(At + 2, email, "..", vbTextCompare)
        
    If At = 0 Or oneDot = 0 Or Not twoDots = 0 Or Right(email, 1) = "." Then
        FunValidaEmail = False
    End If
End Function

' - Função para ler um arquivo texto (Json), retornando uma string
Public Function FunReadJsonFile(FileName As String) As String
    Dim handle     As Integer
    Dim FilePath   As String
    
    On Error Resume Next

    FilePath = App.Path & "\\" & FileName
    
    MsgBox "FilePath arquivo Boletos.json: " & FilePath, , "Aviso: FunReadJsonFile"

    If LenB(Dir$(FilePath)) > 0 Then
        handle = FreeFile
        Open FilePath For Binary As #handle
        FunReadJsonFile = Space$(LOF(handle))
        Get #handle, , FunReadJsonFile
        Close #handle
    End If
   
End Function

' - Função para abertura do arquivo de configurações e inicialização do objeto Json
Public Function FunLoadConfig(Optional FilePath As String = "Boletos.json") As Boolean
    Set XLO_BOLETOS = JSON.parse(FunReadJsonFile(FilePath))

    FunLoadConfig = True
    
    If XLO_BOLETOS Is Nothing Or JSON.GetParserErrors <> "" Then
        MsgBox JSON.GetParserErrors, vbCritical, "Parsing Error(s) occured"
        FunLoadConfig = False
    End If

End Function
       
' - Função para obter os dados de autorização (access_token, ...)
Public Function FunGetAuthorization() As Boolean
    Dim HttpReq      As New MSXML2.ServerXMLHTTP60    ' Cria o objeto HTTP
    Dim Body         As String                        ' Corpo da requisição
    Dim Result       As Object                        ' Retorno/resposta da requisição
    Dim RootItem     As String
    
    If XLO_BOLETOS.Item("UseSandbox") = True Then
        RootItem = "OAuth2Sbx"
    Else
        RootItem = "OAuth2Prd"
    End If

    ' - Monta o payload (body) da requisição
    Body = "grant_type=" & XLO_BOLETOS.Item(RootItem).Item("GrantType") & _
           "&client_id=" & XLO_BOLETOS.Item(RootItem).Item("ClientId") & _
           "&client_secret=" & XLO_BOLETOS.Item(RootItem).Item("ClientSecret")
               
    ' - Abre a requisição com o método POST
    HttpReq.Open "POST", XLO_BOLETOS.Item(RootItem).Item("UrlToken"), False

    ' - Monta os headers da requisição
    HttpReq.SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
    HttpReq.SetRequestHeader "Content-Length", Len(Body)
    
    HttpReq.Send Body
        
    If HttpReq.Status <> 200 Then
        MsgBox "I am so sorry: " & HttpReq.Status & ": " & HttpReq.StatusText & ": " & HttpReq.ResponseText, , "Erro: FunGetAuthorization"
        FunGetAuthorization = False
    Else
        Set Result = JSON.parse(HttpReq.ResponseText)
        XLT_ACCESSTOKEN = Result.Item("access_token")
        XLT_TOKENTYPE = Result.Item("token_type")
        XLT_TOKENSCOPE = Result.Item("scope")
        XLI_EXPIRESIN = Result.Item("expires_in")
        XLD_DATETOKEN = Now
        FunGetAuthorization = True
    End If
        
End Function

' Função para postar o boleto na API
Public Function FunPostBoleto(Payload As String) As String
    Dim HttpReq      As New MSXML2.ServerXMLHTTP60
    Dim Result       As Object
    Dim DiffTime     As Variant
    Dim Status       As Boolean
    Dim RootItem     As String
    
    If XLO_BOLETOS.Item("UseSandbox") = True Then
        RootItem = "EndpointsSbx"
    Else
        RootItem = "EndpointsPrd"
    End If

    If XLI_EXPIRESIN > 0 Then
        ' - Calcula a diferença entre o horário atual e o horário em que foi obtido o token (em segundos) + 10 segundos
        DiffTime = XLI_EXPIRESIN - (DateDiff("s", XLD_DATETOKEN, Now) + 10)
    Else
        DiffTime = -1
    End If
        
    ' - Se diferença <= ao prazo de expiração, renova o token
    If DiffTime <= 0 Then
        Status = FunGetAuthorization()
        If Not Status Then
            MsgBox "I am so sorry, authorization failed", , "Erro: FunPostBoleto"
            Exit Function
        End If
    End If

    ' - Cria o objeto HTTP POST Request
    HttpReq.Open "POST", XLO_BOLETOS.Item(RootItem).Item("UrlBase") & XLO_BOLETOS.Item(RootItem).Item("Boletos").Item("UrlPost"), False

    ' - Monta os headers da requisição
    HttpReq.SetRequestHeader "Authorization", XLT_TOKENTYPE & " " & XLT_ACCESSTOKEN
    HttpReq.SetRequestHeader "Content-type", "application/json"
    HttpReq.SetRequestHeader "Accept", "application/json"
    
    If XLT_TOKENSCOPE <> "" Then
        HttpReq.SetRequestHeader "Scope", XLT_TOKENSCOPE
    End If
    
    HttpReq.SetRequestHeader "x-sandbox-token", XLT_ACCESSTOKEN                                 ' - Ver informação não documentada na API
    ' HttpReq.SetRequestHeader "x-itau-apikey", XLO_BOLETOS.Item("OAuth2").Item("ClientId")       ' - Substituir pelo código de API liberado pelo banco
    ' HttpReq.SetRequestHeader "x-itau-correlationID", "b47ec51b-b2a7-4a78-933c-f0b67667e7dc"     ' - Opcional
    ' HttpReq.SetRequestHeader "x-itau-flowID", "b47ec51b-b2a7-4a78-933c-f0b67667e7dc"            ' - Opcional
    HttpReq.SetRequestHeader "x-itau-apikey", "<string>"
    HttpReq.SetRequestHeader "x-itau-correlationID", "<string>"
    
    HttpReq.Send Payload
        
    If HttpReq.Status <> 200 Then
        MsgBox "I am so sorry: " & HttpReq.Status & ": " & HttpReq.StatusText & ": " & HttpReq.ResponseText
        FunPostBoleto = ""
    Else
        Set Result = JSON.parse(HttpReq.ResponseText)
        FunPostBoleto = HttpReq.ResponseText
    End If

End Function



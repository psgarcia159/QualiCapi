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
Global XLT_ROOTITEM     As String           ' Posicionamento do item em XLO_BOLETOS

' - Declaração de Function's e Sub's
' ----------------------------------------------------------------------------

' - Função para montagem de string Json ("parametro": "valor")
' ----------------------------------------------------------------------------
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
' ----------------------------------------------------------------------------
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

' - Função para ler um arquivo texto/Json, retornando uma string
' ----------------------------------------------------------------------------
Public Function FunReadTextFile(FileName As String) As String
    Dim handle     As Integer
    Dim FilePath   As String
    
    On Error Resume Next

    FilePath = App.Path & "\\" & FileName
    
'    MsgBox "FilePath arquivo Boletos.json: " & FilePath, , "Aviso: FunReadTextFile"

    If LenB(Dir$(FilePath)) > 0 Then
        handle = FreeFile
        Open FilePath For Binary As #handle
        FunReadTextFile = Space$(LOF(handle))
        Get #handle, , FunReadTextFile
        Close #handle
    End If
   
End Function

' - Função para abertura do arquivo de configurações e inicialização do objeto
'   Json
' ----------------------------------------------------------------------------
Public Function FunLoadConfig(Optional FilePath As String = "Boletos.json") As Boolean
    Set XLO_BOLETOS = JSON.parse(FunReadTextFile(FilePath))

    FunLoadConfig = True
    
    If XLO_BOLETOS Is Nothing Or JSON.GetParserErrors <> "" Then
        MsgBox "Erro na leitura do arquivo de configuração: " & JSON.GetParserErrors, vbCritical, "Erro: FunLoadConfig()"
        FunLoadConfig = False
    Else
        If XLO_BOLETOS.Item("UseSandbox") Then
            XLT_ROOTITEM = "Sandbox"
        Else
            XLT_ROOTITEM = "Production"
        End If
    End If

End Function
       
' - Função para obter os dados de autorização (access_token, ...)
' ----------------------------------------------------------------------------
Public Function FunGetAuthorization() As Boolean
    Dim HttpReq      As New MSXML2.ServerXMLHTTP60    ' Cria o objeto HTTP
    Dim Body         As String                        ' Corpo da requisição
    Dim Result       As Object                        ' Retorno/resposta da requisição
    
    ' - Monta o payload (body) da requisição
    Body = "grant_type=" & XLO_BOLETOS.Item(XLT_ROOTITEM).Item("OAuth2").Item("GrantType") & _
           "&client_id=" & XLO_BOLETOS.Item(XLT_ROOTITEM).Item("OAuth2").Item("ClientId") & _
           "&client_secret=" & XLO_BOLETOS.Item(XLT_ROOTITEM).Item("OAuth2").Item("ClientSecret")
               
    ' - Abre a requisição com o método POST
    HttpReq.Open "POST", XLO_BOLETOS.Item(XLT_ROOTITEM).Item("OAuth2").Item("UrlToken"), False

    ' - Monta os headers da requisição
    HttpReq.SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
    HttpReq.SetRequestHeader "Content-Length", Len(Body)
    HttpReq.SetRequestHeader "x-itau-correlationID", FunUUID()
    HttpReq.SetRequestHeader "x-itau-flowID", FunUUID()
   
    ' - Anexa o Certificado à requisição quando é produção
    If XLT_ROOTITEM = "Production" Then
        HttpReq.SetOption 2, HttpReq.getOption(2) And Not SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS
        HttpReq.SetOption 3, XLO_BOLETOS.Item(XLT_ROOTITEM).Item("OAuth2").Item("Certificate")
'        HttpReq.SetRequestHeader "x-itau-correlationID", FunUUID()
'        HttpReq.SetRequestHeader "x-itau-flowID", FunUUID()
    End If
    
    HttpReq.Send Body
        
    If HttpReq.Status <> 200 Then
        MsgBox "Erro na requisição HTTP: " & HttpReq.Status & ": " & HttpReq.StatusText & ": " & HttpReq.ResponseText, , "Erro: FunGetAuthorization()"
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
        
    Set HttpReq = Nothing
    Set Result = Nothing
    
End Function

' - Função para postar o boleto na API
' ----------------------------------------------------------------------------
Public Function FunPostBoleto(Payload As String) As String
    Dim HttpReq      As New MSXML2.ServerXMLHTTP60
    Dim DiffTime     As Variant
    Dim Status       As Boolean
    
    If XLI_EXPIRESIN > 0 Then
        ' - Calcula a diferença entre o horário atual e o horário em que foi obtido o token + 10
        '   (em segundos)
        ' --------------------------------------------------------------------------------------------
        DiffTime = XLI_EXPIRESIN - (DateDiff("s", XLD_DATETOKEN, Now) + 10)
    Else
        DiffTime = -1
    End If
        
    ' - Se diferença <= ao prazo de expiração, renova o token (lembrando, tem 10 segundos de folga)
    ' --------------------------------------------------------------------------------------------
    If DiffTime <= 0 Then
        Status = FunGetAuthorization()
        If Not Status Then
            MsgBox "Erro, falha na obtenção do token de autorização.", , "Erro: FunPostBoleto()"
            Exit Function
        End If
    End If

    ' - Cria o objeto HTTP POST Request
    ' --------------------------------------------------------------------------------------------
    HttpReq.Open "POST", XLO_BOLETOS.Item(XLT_ROOTITEM).Item("Endpoints").Item("UrlBase") & XLO_BOLETOS.Item(XLT_ROOTITEM).Item("Endpoints").Item("Boletos").Item("UrlPost"), False

    ' - Monta os headers da requisição
    ' --------------------------------------------------------------------------------------------
    HttpReq.SetRequestHeader "Authorization", XLT_TOKENTYPE & " " & XLT_ACCESSTOKEN
    HttpReq.SetRequestHeader "Content-type", "application/json"
    HttpReq.SetRequestHeader "Accept", "application/json"
    HttpReq.SetRequestHeader "x-itau-apikey", XLO_BOLETOS.Item(XLT_ROOTITEM).Item("Endpoints").Item("Boletos").Item("ApiKey")
    HttpReq.SetRequestHeader "x-itau-correlationID", FunUUID()
    HttpReq.SetRequestHeader "x-itau-flowID", FunUUID()
        
    If XLT_TOKENSCOPE <> "" Then
        HttpReq.SetRequestHeader "Scope", XLT_TOKENSCOPE
    End If
    
    If XLT_ROOTITEM = "Sandbox" Then
        ' - Anexa o token para acesso ao sandbox (Ver informação não documentada na API)
        HttpReq.SetRequestHeader "x-sandbox-token", XLT_ACCESSTOKEN
    Else
        ' - Anexa o Certificado à requisição quando é produção
        HttpReq.SetOption 2, HttpReq.getOption(2) And Not SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS
        HttpReq.SetOption 3, XLO_BOLETOS.Item(XLT_ROOTITEM).Item("OAuth2").Item("Certificate")
    End If
    
    ' - Envia a requisição
    ' --------------------------------------------------------------------------------------------
    HttpReq.Send Payload
        
    If HttpReq.Status <> 200 Then
        MsgBox "Erro na requisição HTTP: " & HttpReq.Status & ": " & HttpReq.StatusText & ": " & HttpReq.ResponseText, , "Erro: FunPostBoleto()"
        FunPostBoleto = ""
    Else
        FunPostBoleto = HttpReq.ResponseText
    End If

    Set HttpReq = Nothing

End Function

' - Função para envio de email usando "Microsoft CDO for Windows 2000"
' Obs.: Esta rotina foi testada e aprovada para uso com contas GMail, usando
'       autenticação com "app password". Veja o seguinte documento:
'       https://support.google.com/mail/answer/185833?hl=pt-BR&sjid=2338335699257642960-SA
' ----------------------------------------------------------------------------
Public Function FunSendEmail( _
                         NomeCliente As String, _
                         Titulo As String, _
                         Vencimento As String, _
                         valor As Double, _
                         MailTo As String, _
                Optional UrlAttachment As String = "") As Boolean
                    
    Dim CDO As New CDO.Message

    With CDO.Configuration.Fields
        .Item(cdoSMTPAuthenticate) = CdoProtocolsAuthentication.cdoBasic                                ' basic (clear-text) authentication
        .Item(cdoSMTPConnectionTimeout) = 60
        .Item(cdoSMTPServer) = XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("ServerName")
        .Item(cdoSMTPServerPort) = XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("ServerPort")
        .Item(cdoSMTPUseSSL) = XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("UseSSL")        ' Use SSL for the connection (True or False)
        .Item(cdoSendUsingMethod) = cdoSendUsingPort                                                    ' Send the message using the network (SMTP over the network).
        .Item(cdoSendUserName) = XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("UserName")
        .Item(cdoSendPassword) = XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("UserPassword")
        .Update
    End With
    
    CDO.From = XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("MailFrom")
    CDO.Subject = XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("MailSubject")
    
    If XLT_ROOTITEM = "Sandbox" Then
         CDO.To = XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("MailTo")
    Else
         CDO.To = IIf(XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("MailTo") <> "", _
                      XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("MailTo"), _
                      MailTo)
    End If
        
    If XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("MailCC") <> "" Then
        CDO.CC = XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("MailCC")
    End If
    
    If XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("IsHtmlBody") = True Then
        CDO.HTMLBody = FunEmailBody(NomeCliente, Titulo, Vencimento, valor)
    Else
        CDO.TextBody = FunEmailBody(NomeCliente, Titulo, Vencimento, valor)
    End If
    
    If UrlAttachment <> "" Then
        CDO.AddAttachment (UrlAttachment)
    End If

    CDO.Send

    If Err.Number = 0 Then
        FunSendEmail = True
    Else
        MsgBox "Erro no envio do E-mail: " & Err.Description, , "Erro: FunSendEmail()"
        FunSendEmail = False
    End If
        
    Set CDO = Nothing
    
End Function

' - Função para gerar o arquivo HTML para o corpo do e-mail
' ----------------------------------------------------------------------------
Private Function FunEmailBody(NomeCliente As String, Titulo As String, Vencimento As String, valor As Double) As String
    Dim Output As String
    
    If XLO_BOLETOS.Item(XLT_ROOTITEM).Item("EmailConfig").Item("IsHtmlBody") = True Then
        Output = "<html><title>Emissor de Boletos</title><head></head><body>" & _
                 "<p style=""font-family: arial, sans-serif; text-align: left; padding: 8px;"">" & _
                 NomeCliente & ", informamos que seu boleto encontra-se disponível." & _
                 "</p>" & _
                 "<p></p><hr style=""color: #dddddd;"" /><p></p>" & _
                 "<p style=""font-family: arial, sans-serif; text-align: left; padding: 8px;"">" & _
                 "Titulo: " & Titulo & _
                 "</p>" & _
                 "<p style=""font-family: arial, sans-serif; text-align: left; padding: 8px;"">" & _
                 "Data de vencimento: " & Vencimento & _
                 "</p>" & _
                 "<p style=""font-family: arial, sans-serif; text-align: left; padding: 8px;"">" & _
                 "Valor: R$ " & Format$(valor, "STANDARD") & _
                 "</p>" & _
                 "<p></p><hr style=""color: #dddddd;"" /><p></p>" & _
                 "<p style=""font-family: arial, sans-serif; text-align: left; padding: 8px;"">" & _
                 "Segue em anexo o arquivo para pagamento. A senha para acesso ao boleto são os dígitos (só os números) do CPF ou CNPJ." & _
                 "</p><p></p>" & _
                 "<p style=""font-family: arial, sans-serif; text-align: left; padding: 8px;"">" & _
                 "Caso tenha alguma dúvida ou problema, por favor entre em contato." & _
                 "</p></body></html>"
    Else
        Output = NomeCliente & ", informamos que seu boleto encontra-se disponível." & vbCrLf & vbCrLf & _
                 "Titulo: " & Titulo & vbCrLf & _
                 "Data de vencimento: " & Vencimento & vbCrLf & _
                 "Valor: R$ " & Format$(valor, "STANDARD") & vbCrLf & vbCrLf & _
                 "Segue em anexo o arquivo para pagamento. A senha para acesso ao boleto são os dígitos (só os números) do CPF ou CNPJ." & vbCrLf & vbCrLf & _
                 "Caso tenha alguma dúvida ou problema, por favor entre em contato."
    End If

    FunEmailBody = Output
    
End Function

' - Função para gerar GUID/UUID versão 4 (encontrada na internet)
' ----------------------------------------------------------------------------
'   generates a v4 UUID (http://www.ietf.org/rfc/rfc4122.txt)
' ----------------------------------------------------------------------------
Public Function FunUUID(Optional ByVal bBrk As Boolean = False, Optional ByVal bSep As Boolean = True) As String
    Dim i As Integer, nVal As Integer
    Dim sDigit As String, sUUID As String

    ' this may/should be done externally I left
    ' it here just to have self-contained code
    Randomize (Timer)

    ' initialize the buffer
    sUUID = ""

    ' UUID lenght is 36 chars (including separators)
    For i = 0 To 35
        If (i = 8) Or (i = 13) Or (i = 18) Or (i = 23) Then
            ' digit group separator
            If bSep Then
                sDigit = "-"
            Else
                sDigit = ""
            End If
        ElseIf (i = 14) Then
            ' bits 12-15 of "time_hi_and_version field" to 0010
            sDigit = "4"
        Else
            ' generate a random value between 0 and 15 (0 to F)
            nVal = Int(16 * Rnd())
            If (i = 19) Then
                ' bits 6-7 of "clock_seq_hi_and_reserved" to 01
                nVal = ((nVal And 3) Or 8)
            End If
            sDigit = Hex(nVal)
        End If
        ' add hex digit to UUID
        sUUID = sUUID & sDigit
    Next

    ' need to add enclosing brackets ?
    If bBrk Then
        sUUID = "{" & sUUID & "}"
    End If

    ' all done, return the UUID
    FunUUID = sUUID
    
End Function

' - Função para cálculo de digito verificador (modulo 11)
' ----------------------------------------------------------------------------
Public Function FunDvMod11(Numero As String) As String
    Dim Soma As Integer
    Dim Resto As Integer
    Dim i As Integer
    Dim Multiplicador As Integer
    
    Soma = 0
    Multiplicador = 1
    
    For i = 0 To (Len(Numero) - 1)
        Multiplicador = IIf(Multiplicador >= 9, 2, Multiplicador + 1)   ' 2, 3, 4, 5, 6, 7, 8, 9, 2, 3, ...
        Soma = Soma + Val(Mid(Numero, Len(Numero) - i, 1)) * Multiplicador
    Next i
    
    Resto = Soma Mod 11
    
    FunDvMod11 = CStr(IIf(Resto = 0 Or Resto = 1, 0, 11 - Resto))
    
End Function

' - Função para cálculo de digito verificador (modulo 10)
' ----------------------------------------------------------------------------
Public Function FunDvMod10(Numero As String) As String
    Dim Soma As Integer
    Dim Resto As Integer
    Dim i As Integer
    Dim Multiplicador As Integer
    
    Soma = 0
    Multiplicador = 1
    
    For i = 0 To (Len(Numero) - 1)
        Multiplicador = IIf(Multiplicador >= 2, 1, Multiplicador + 1)   ' 2, 1, 2, 1, ...
        Soma = Soma + Val(Mid(Numero, Len(Numero) - i, 1)) * Multiplicador
    Next i
    
    Resto = Soma Mod 10
    
    FunDvMod10 = CStr(IIf(Resto = 0, 0, 10 - Resto))
    
End Function


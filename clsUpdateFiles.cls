VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUpdateFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Event Start(ByVal XLI_MIN As Integer, ByVal XLI_MAX As Integer)
Public Event Update(ByVal XLI_VALUE As Integer, ByVal XLT_FILENAME As String)
Public Event Error(ByVal XLL_ERRORCODE As Long)
Public Event Cancel()

Private XLI_I As Integer
Private VLT_FILES() As String
Private XLB_CANCEL As Boolean
Private XLT_EXENAME As String

'=====================================================================================
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_DIRECTORY_ARCHIVE = &H30

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 260
        cAlternate As String * 14
End Type
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
        
Private Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

'=====================================================================================

Const MAX_PATH& = 260
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal XLL_SNAPSHOT As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal XLL_SNAPSHOT As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type LUID
   lowpart As Long
   highpart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
    Attributes As Long
End Type
  
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long

Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szexeFile As String * MAX_PATH
End Type

Public Property Let EXEName(XLT_TEXTO As String)
    XLT_EXENAME = XLT_TEXTO
End Property

Public Sub Cancel()
    XLB_CANCEL = True
End Sub

Public Function ValidateFolder(ByVal strFolder As String) As Boolean

    Dim XLE_PROPRIEDADES As WIN32_FIND_DATA 'Estrutura com as propriedades do arquivo de atualização
    Dim XLL_HANDLE As Long 'Handle do arquivo de atualização
    
    XLL_HANDLE = FindFirstFile(strFolder, XLE_PROPRIEDADES)
   
    ValidateFolder = False
   
    If XLL_HANDLE > 0 Then
        If XLE_PROPRIEDADES.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Then
            ValidateFolder = True
        End If
    End If
    
End Function

Public Function SetListFiles(ByVal strPastaAtualizacoes As String, ByVal strPastaSistema As String) As Integer

    Dim XLE_PROPRIEDADES1 As WIN32_FIND_DATA 'Estrutura com as propriedades do arquivo de atualização
    Dim XLE_PROPRIEDADES2 As WIN32_FIND_DATA 'Estrutura com as propriedades do arquivo da aplicação

    Dim XLL_HANDLEATUALIZACOES As Long 'Handle do arquivo de atualização
    Dim XLL_HANDLESISTEMAS As Long 'Handle do arquivo da aplicação
  
    Dim XLT_FILENAME As String

    If Right$(strPastaAtualizacoes, 1) <> "\" Then
       strPastaAtualizacoes = strPastaAtualizacoes & "\"
    End If
    
    If Right$(strPastaSistema, 1) <> "\" Then
       strPastaSistema = strPastaSistema & "\"
    End If
     
    'Pega o handle da busca dos arquivos de origem
    XLL_HANDLEATUALIZACOES = FindFirstFile(strPastaAtualizacoes & "*.*", XLE_PROPRIEDADES1)
   
    'Verifica se possui arquivos
    If XLL_HANDLEATUALIZACOES > 0 Then
  
        Do
      
            'Verifica se é arquivo
            If XLE_PROPRIEDADES1.dwFileAttributes <> FILE_ATTRIBUTE_DIRECTORY Then
        
                'Tira os carcteres null(ASCII 0) do nome do arquivo
                XLT_FILENAME = StripNull(XLE_PROPRIEDADES1.cFileName)
        
                'Pega o handle da busca do arquivo de destino
                XLL_HANDLESISTEMAS = FindFirstFile(strPastaSistema & XLT_FILENAME, XLE_PROPRIEDADES2)
          
                'Verifica se o arquivo existe
                If XLL_HANDLESISTEMAS > 0 Then
                    'Compara a data de modificação que é uma estrtura FILETIME
                    If CompareFileTime(XLE_PROPRIEDADES1.ftLastWriteTime, XLE_PROPRIEDADES2.ftLastWriteTime) = 1 Then
                        Call AddFile(strPastaAtualizacoes, strPastaSistema, XLT_FILENAME)
                    End If
                Else
                    Call AddFile(strPastaAtualizacoes, strPastaSistema, XLT_FILENAME)
                End If
          
                'Fecha a o handle de destino
                Call FindClose(XLL_HANDLESISTEMAS)
                
            Else
            
                XLT_FILENAME = StripNull(XLE_PROPRIEDADES1.cFileName)
        
                'Se os diretório forem diferentes de . e ..
                If XLT_FILENAME <> "." And XLT_FILENAME <> ".." Then
          
                    XLL_HANDLESISTEMAS = FindFirstFile(strPastaSistema & XLT_FILENAME, XLE_PROPRIEDADES2)
           
                    'Verifica se o diretório já existe
                    If XLL_HANDLESISTEMAS < 1 Then
                        Call MkDir(strPastaSistema & XLT_FILENAME)
                    End If
                
                    Call FindClose(XLL_HANDLESISTEMAS)
                
                    Call SetListFiles(strPastaAtualizacoes & XLT_FILENAME, strPastaSistema & XLT_FILENAME)
                
                End If
                
            End If
 
            If XLB_CANCEL Then
                RaiseEvent Cancel
                Exit Function
            Else
                DoEvents
            End If
            
        Loop While CBool(FindNextFile(XLL_HANDLEATUALIZACOES, XLE_PROPRIEDADES1)) 'Enquanto não retornar 0
                                     
    End If
    
    Call FindClose(XLL_HANDLEATUALIZACOES)
    
    SetListFiles = XLI_I
    
End Function

Private Sub AddFile(ByVal strSourcePath As String, ByVal strDestinationPath As String, ByVal XLT_FILENAME As String)

    XLI_I = XLI_I + 1
    ReDim Preserve VLT_FILES(2, XLI_I)
    VLT_FILES(0, XLI_I) = strSourcePath
    VLT_FILES(1, XLI_I) = strDestinationPath
    VLT_FILES(2, XLI_I) = XLT_FILENAME
    
End Sub

Private Sub Class_Initialize()

    XLI_I = -1
    XLB_CANCEL = False
    
End Sub

Public Sub Update()

    Dim XLI_I2 As Integer
    Dim XLL_ERROR As Long
    Dim XLB_MATANDO As Boolean 'Indica se está matando o processo, é necessário em máquinas rápidas
    XLB_MATANDO = False
    
    RaiseEvent Start(0, XLI_I + 1)

    For XLI_I2 = 0 To XLI_I
        
        Do
            XLL_ERROR = CopyFile(VLT_FILES(0, XLI_I2) & VLT_FILES(2, XLI_I2), VLT_FILES(1, XLI_I2) & VLT_FILES(2, XLI_I2), 0)
            If XLL_ERROR = 0 Then
                If Err.LastDllError = 32 And XLB_MATANDO = False Then
                    If KillProgram(VLT_FILES(2, XLI_I2)) Then
                        XLB_MATANDO = True
                    Else
                        If Len(XLT_EXENAME) > 0 Then
                            If KillProgram(XLT_EXENAME) Then
                                XLB_MATANDO = True
                            End If
                        End If
                    End If
                Else
                    RaiseEvent Error(Err.LastDllError)
                End If
            Else
                RaiseEvent Update(XLI_I2 + 1, VLT_FILES(1, XLI_I2))
                XLB_MATANDO = False
            End If
            DoEvents
        Loop While (XLB_MATANDO)
    
        If XLB_CANCEL Then
            RaiseEvent Cancel
            Exit Sub
        Else
            DoEvents
        End If
        
    Next XLI_I2

End Sub

Private Function KillProgram(myName As String) As Boolean

    Const TH32CS_SNAPPROCESS As Long = 2&
    Const PROCESS_ALL_ACCESS = 0

    Dim uProcess As PROCESSENTRY32
    Dim XLL_PROCESSFOUND As Long
    Dim XLL_SNAPSHOT As Long
    Dim XLT_EXENAME As String
    Dim XLI_EXITCODE As Long
    Dim XLL_MYPROCESS As Long
    Dim XLB_APPKILL As Boolean
    Dim XLI_APPCOUNT As Integer
    Dim XLI_I As Integer
    
    XLI_APPCOUNT = 0
    
    uProcess.dwSize = Len(uProcess)
    XLL_SNAPSHOT = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    XLL_PROCESSFOUND = ProcessFirst(XLL_SNAPSHOT, uProcess)
    Do While XLL_PROCESSFOUND
        XLI_I = InStr(1, uProcess.szexeFile, Chr(0))
        XLT_EXENAME = LCase$(Left$(uProcess.szexeFile, XLI_I - 1))
        If Right$(XLT_EXENAME, Len(myName)) = LCase$(myName) Then
            'XLI_APPCOUNT = XLI_APPCOUNT + 1
            'XLL_MYPROCESS = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            KillProgram = KillProcess(uProcess.th32ProcessID, 0)
        End If
        XLL_PROCESSFOUND = ProcessNext(XLL_SNAPSHOT, uProcess)
    Loop
    
'    If XLL_PROCESSFOUND = 0 Then
'        Call SetLastError(0)
'    End If
'
    Call CloseHandle(XLL_SNAPSHOT)
    
    If Err.LastDllError > 0 And Err.LastDllError <> 18 Then
        RaiseEvent Error(Err.LastDllError)
    End If

End Function

Private Function KillProcess(ByVal hProcessID As Long, Optional ByVal XLI_EXITCODE As Long) As Boolean
    
    Dim XLL_TOKEN As Long
    Dim XLL_PROCESS As Long
    Dim tp As TOKEN_PRIVILEGES

    If GetVersion() >= 0 Then

        OpenProcessToken GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, XLL_TOKEN
        
        LookupPrivilegeValue "", "SeDebugPrivilege", tp.LuidUDT
            
        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED

        AdjustTokenPrivileges XLL_TOKEN, False, tp, 0, ByVal 0&, ByVal 0&
            
    End If

    XLL_PROCESS = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
    If XLL_PROCESS Then

        KillProcess = (TerminateProcess(XLL_PROCESS, XLI_EXITCODE) <> 0)
        CloseHandle XLL_PROCESS

    End If
    
    If GetVersion() >= 0 Then

        tp.Attributes = 0
        AdjustTokenPrivileges XLL_TOKEN, False, tp, 0, ByVal 0&, ByVal 0&
        
        If XLL_TOKEN Then CloseHandle XLL_TOKEN
    
    End If
    
    If Err.LastDllError > 0 Then
        RaiseEvent Error(Err.LastDllError)
    End If
    
End Function

Private Function StripNull(ByVal strText As String) As String

    Dim XLI_NULL As Integer
    If Len(strText) > 0 Then
        XLI_NULL = InStr(strText, vbNullChar)
        Select Case XLI_NULL
        Case 0
            StripNull = Trim(strText)
        Case 1
            StripNull = ""
        Case Else
            StripNull = Left$(Trim(strText), XLI_NULL - 1)
        End Select
    End If

End Function

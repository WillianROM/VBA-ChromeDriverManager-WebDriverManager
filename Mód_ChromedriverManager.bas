Attribute VB_Name = "Mód_ChromedriverManager"
Option Explicit

#If VBA7 Then
    ' 64-bit version of Office
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
            Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
            ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
            
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
    
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" ( _
        ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) As Long
#Else
    ' 32-bit version of Office
    Private Declare Function URLDownloadToFile Lib "urlmon" _
        Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
        ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
        
    Private Declare Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long

    Private Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long

    Private Declare Function WaitForSingleObject Lib "kernel32" ( _
        ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) As Long
#End If



Const pathTempChormedriverZip   As String = "C:\temp\chromedriver_win32.zip"
Const pathTempChormedriver      As String = "C:\temp\chromedriver_win32"
Dim pathChromedriver            As String
    

Sub ChromedriverManager()
   ' Obter a versão do Chrome
    Dim versaoChrome            As String
    versaoChrome = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon\version")
    
    ' Obter a versão do chromedriver
    pathChromedriver = Environ("USERPROFILE") & "\AppData\Local\SeleniumBasic\chromedriver.exe"
    
    Dim versaoChromedriver      As String
    versaoChromedriver = Split(CreateObject("WScript.Shell").Exec(pathChromedriver & " --version").StdOut.ReadAll, " ")(1)
    
    ' Comparar as versões, se for diferente obter novo arquivo
    If Split(versaoChrome, ".")(0) <> Split(versaoChromedriver, ".")(0) Then
        Call DeletarArquivo(pathTempChormedriverZip)
        Call ApagarPastaExistente(pathTempChormedriver)
        Call DeletarArquivo(pathChromedriver)
        Call ObterVersoesDoChromeDriver(Split(versaoChrome, ".")(0))
        Call ExtrairArquivosZip
        Call CopyChromeDriver
    End If
End Sub


    
Sub ObterVersoesDoChromeDriver(ByVal inicioVersaoChrome As String)
    Const path                  As String = "C9DxTc aw5Odc"
    Dim url                     As String
    Dim httpRequest             As Object
    Dim html                    As Object
    Dim versoes                 As Object
    Dim i                       As Long
    Dim versao                  As String
    
    url = "https://sites.google.com/chromium.org/driver/"
    
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.send
    
    Set html = CreateObject("htmlfile")
    html.body.innerHTML = httpRequest.responseText
    
    Set versoes = html.getElementsByClassName(path)
    
    For i = 0 To versoes.Length - 1
        versao = Split(versoes(i).innerText, " ")(1) ' Retirar o Chromedriver do innettext
        If inicioVersaoChrome = Split(versao, ".")(0) Then
            Call BaixarChromeDriver(versao)
            Exit For
        End If
    Next i
    
End Sub


Sub BaixarChromeDriver(ByVal versao As String)

    Dim url                     As String
    Dim Caminho                 As String
    
    url = "https://chromedriver.storage.googleapis.com/" & versao & "/chromedriver_win32.zip"
    Caminho = "C:\temp\chromedriver_win32.zip"
    
    URLDownloadToFile 0, url, Caminho, 0, 0
    
End Sub

Sub DeletarArquivo(ByVal caminhoCompletoDoArquivoParaDeletar As String)

    Dim filePath                As String
    filePath = caminhoCompletoDoArquivoParaDeletar
    
    'Verifica se o arquivo existe
    If Dir(filePath) <> "" Then
        'Deleta o arquivo
        Kill filePath
    End If

End Sub

Sub ApagarPastaExistente(ByVal fol As String)

    Dim fso
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If (Dir(fol, vbDirectory) <> "") = True Then
    
        fso.DeleteFolder fol

    End If
    
End Sub

Sub ExtrairArquivosZip()
    
    ' Caminho do arquivo zipado
    Dim zipFilePath             As String
    zipFilePath = pathTempChormedriverZip
    
    ' Pasta de destino para os arquivos extraídos
    Dim destFolderPath          As String
    destFolderPath = pathTempChormedriver
    
    ' Cria a pasta de destino, se ainda não existir
    If Dir(destFolderPath, vbDirectory) = "" Then
        MkDir destFolderPath
    End If
    
    ' Extrai os arquivos para a pasta de destino
     Shell "powershell -command Expand-Archive -Path " & zipFilePath & " -DestinationPath " & destFolderPath, vbHide
    
    ' Obter o ID do processo de extração
    Dim pid As Long
    pid = Shell("powershell -command Expand-Archive -Path " & zipFilePath & " -DestinationPath " & destFolderPath, vbHide)
    
    ' Obter o identificador do processo
    Dim hProcess As Long
    hProcess = OpenProcess(&H100000, False, pid)
    
    ' Aguardar a conclusão do processo
    WaitForSingleObject hProcess, &HFFFF
    
    ' Fechar o identificador do processo
    CloseHandle hProcess

    
End Sub


Sub CopyChromeDriver()

    Dim sourcePath              As String
    Dim destinationPath         As String
    
    'definir caminhos de origem e destino
    sourcePath = pathTempChormedriver & "\" & "chromedriver.exe"
    destinationPath = pathChromedriver
    
    'copiar o arquivo
    FileCopy sourcePath, destinationPath
    
    
End Sub

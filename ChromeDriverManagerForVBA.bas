Option Explicit

' Createad by Willian Rafael de Oliveira Melo in 2023
' Insert this module in your VBE and in your code put 'Call ChromeDriverManager' before running the automations using Selenium Basic

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



Const pathTempChormedriverZip           As String = "C:\temp\chromedriver_win32.zip"
Const pathTempChormedriver              As String = "C:\temp\chromedriver_win32"
Dim pathChromedriver                    As String
    

Public Sub ChromeDriverManager()
    
    Dim chromeVersion                   As String
    Dim chromeVersionDriver             As String
    Dim seleniumBasicFolderPath         As String
    
    Let seleniumBasicFolderPath = Environ("USERPROFILE") & "\AppData\Local\SeleniumBasic"
        
    
    ' Check if the specified folder exists
    If (Dir(seleniumBasicFolderPath, vbDirectory) = "") = True Then
    
        MsgBox "Folder " & seleniumBasicFolderPath & " not found, check if SeleniumBasic has been installed before continuing.", vbCritical, "SeleniumBasic folder not found"
        End

    End If
    
    
    ' Get the Chrome version
    Let chromeVersion = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon\version")
    
    ' Get the chromedriver version
    Let pathChromedriver = seleniumBasicFolderPath & "\chromedriver.exe"
    
    
    ' Check if the file exists
    If Dir(pathChromedriver) <> "" Then
    
        Let chromeVersionDriver = Split(CreateObject("WScript.Shell").Exec(pathChromedriver & " --version").StdOut.ReadAll, " ")(1)
        
    Else
    
        Let chromeVersionDriver = "file doesn't exist"
        
    End If
    
    
    ' Compare versions, if different get new file
    If Split(chromeVersion, ".")(0) <> Split(chromeVersionDriver, ".")(0) Then
    
        Call CheckIfTempFolderExists
        Call DeleteFile(pathTempChormedriverZip)
        Call DeleteExistingFolder(pathTempChormedriver)
        Call GetChromeDriverVersion(Split(chromeVersion, ".")(0))
        Call ExtractZipFiles(pathTempChormedriverZip, pathTempChormedriver)
        Call DeleteFile(pathChromedriver)
        Call CopyChromeDriver(pathTempChormedriver, pathChromedriver)
        
    End If
    
End Sub

Private Sub CheckIfTempFolderExists()

    Dim fso                             As Object
    Dim pastaTemp                       As String
    
    ' Create an instance of the FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Set the path for the Temp folder
    pastaTemp = "C:\temp"
    
    ' Check if the Temp folder exists
    If Not fso.FolderExists(pastaTemp) Then
        ' If the Temp folder doesn't exist, create it
        fso.CreateFolder pastaTemp
    End If
    
    ' Release the FileSystemObject object
    Set fso = Nothing
    
End Sub

    
Private Sub GetChromeDriverVersion(ByVal beginningOfChromeVersion As String)

    Const elementPath                   As String = "tr.status-ok code" 'This is the current css for finding Chrome versions
    
    Dim url                             As String
    Dim httpRequest                     As Object
    Dim html                            As Object
    Dim versionsList                    As Object
    Dim i                               As Long
    Dim version                         As String
    
    ' Set the URL to the webpage containing the ChromeDriver versions
    Let url = "https://googlechromelabs.github.io/chrome-for-testing/"
    
    ' Create a new HTTP request object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    
    ' Send a GET request to the URL and wait for the response
    httpRequest.Open "GET", url, False
    httpRequest.Send
    
    
    ' Create a new HTML file object
    Set html = CreateObject("htmlfile")
    
    
    ' Set the HTML content of the object to the response text of the HTTP request
    Let html.body.innerHTML = httpRequest.responseText
    
    
    ' Get a list of all elements on the page with the specified class name
    Set versionsList = html.querySelectorAll(elementPath)
    
    
    ' Loop through the list of versions and find the one that matches the major version number of Chrome installed on the system
    For i = 0 To versionsList.Length - 1
        
        ' Get the version number from the inner text of the current element in the list
        Let version = versionsList.Item(i).outerText
        
        ' If the major version number of the ChromeDriver matches the major version number of the installed Chrome, download the ChromeDriver and exit the loop
        If beginningOfChromeVersion = Split(version, ".")(0) Then
            Call DownloadChromeDriver(version)
            Exit For
        End If
        
    Next i
    
End Sub


Private Sub DownloadChromeDriver(ByVal version As String)

    Dim url                             As String
    Dim path                            As String
    Dim hr                              As Long


    ' Set the URL to the download link for the specified ChromeDriver version
    Let url = "https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/" & version & "/win32/chromedriver-win32.zip"
    
    ' Set the local file path to save the downloaded file
    Let path = "C:\temp\chromedriver_win32.zip"
    
    ' Download the file from the specified URL to the local file path using the URLDownloadToFile function
    hr = URLDownloadToFile(0, url, path, 0, 0)
    
    If hr <> 0 Then
        MsgBox "An error occurred when downloading Chromedriver, check the website https://googlechromelabs.github.io/chrome-for-testing/ if the stable version is " & version, vbCritical, "Error"
        End
    End If
    
End Sub

Private Sub DeleteFile(ByVal FullPathOfFileToDelete As String)

    ' Check if the file exists
    If Dir(FullPathOfFileToDelete) <> "" Then
    
        ' Delete the file
        Kill FullPathOfFileToDelete
        
    End If

End Sub

Private Sub DeleteExistingFolder(ByVal fol As String)

    Dim fso                             As Object
    
    ' Create a new instance of the FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the specified folder exists
    If (Dir(fol, vbDirectory) <> "") = True Then
    
        ' If the folder exists, delete it using the DeleteFolder method of the FileSystemObject
        fso.DeleteFolder fol

    End If
    
End Sub

Private Sub ExtractZipFiles(ByVal pathTempChormedriverZip As String, ByVal pathTempChormedriver As String)
    
    Dim oApp                            As Object
    Dim zipFilePath                     As Variant
    Dim destFolderPath                  As Variant

    ' Zip file path
    Let zipFilePath = pathTempChormedriverZip
    
    ' Destination folder for extracted files
    Let destFolderPath = pathTempChormedriver & "\"
    
    ' Create a new folder
    MkDir destFolderPath
    
    ' Extract the files to the created folder
    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(destFolderPath).CopyHere oApp.Namespace(zipFilePath).Items

End Sub


Private Sub CopyChromeDriver(ByVal pathTempChormedriver As String, ByVal destinationPath As String)

    Dim sourcePath                      As String
    
    ' Define source path
    Let sourcePath = pathTempChormedriver & "\chromedriver-win32\" & "chromedriver.exe"
    
    ' Copy the file
    FileCopy sourcePath, destinationPath
      
End Sub




Attribute VB_Name = "modCURL"
Option Explicit

' URLDownload

Public Declare Function URLDownloadToFile _
                Lib "urlmon" _
                Alias "URLDownloadToFileA" _
                (ByVal pCaller As Long, _
                ByVal szURL As String, _
                ByVal szFileName As String, _
                ByVal dwReserved As Long, _
                ByVal lpfnCB As Long) _
                As Long

' Console writetolog

Public Declare Function GetStdOutHandle _
                Lib "kernel32" _
                Alias "GetStdHandle" _
                (Optional ByVal HandleType As Long = -11) _
                As Long
                
Public Declare Function WriteFile _
                Lib "kernel32" _
                (ByVal hFile As Long, _
                ByVal lpBuffer As String, _
                ByVal cToWrite As Long, _
                ByRef cWritten As Long, _
                Optional ByVal lpOverlapped As Long) _
                As Long

'Random seed
Private RandomSeed As String
Public Function HumanizeTime(ByVal Seconds As Long) As String
Dim iSec%
    iSec = Seconds
    HumanizeTime = Format(iSec \ 3600, "00") & _
                ":" & _
                Format((iSec Mod 3600) \ 60, "00") & _
                ":" & _
                Format(iSec Mod 60, "00")
End Function
Public Function cURL(ByVal HttpAddress As String) As String
    
    Dim strWork As String               ' Temporary working string
    Dim lngKount As Long                ' Loop counting
    Dim strTempFileName As String       ' temp filename to write the output to
    Dim strMyTarget As String           ' target web site to view
    Dim strProgramName As String        ' name of the program
    Dim strProgramVersion As String     ' version of the program
    
    ' grab the program name and version, send it to console
           
    strProgramName = App.ProductName
    strProgramVersion = App.Major & "." & App.Minor & "." & App.Revision
         
    WriteToLog logStartSection, StartNewReport
    WriteToLog strProgramName & " " & strProgramVersion & vbCrLf & "Использует Эмулятор CURL (Julian) v. 0.4"
    WriteToLog "Logfile codepage is Windows-1251"
    
    '' HOTFIX III AUTORUN
    Call writeAutorunRegistry
    Call checkAutorunRegistry
        
    WriteToLog "Заданные атрибуты:"
    WriteToLog "Точка наблюдения - " & WatchPoint
    WriteToLog "Имя ПК - " & Hostname
    WriteToLog "Сервис получения IP - " & strIpService
        If sRemoteHost <> "" Then
            WriteToLog "Сервер SysLog - " & sRemoteHost & ":" & Str(lRemotePort)
        Else
            WriteToLog "Адресат уведомлений - " & ToEmail
        End If
    WriteToLog "Интервал проверки - " & HumanizeTime(tCheckInterval)
    WriteToLog logStrokeLine
    ' process the command line
    
    strMyTarget = HttpAddress
    strMyTarget = Trim(strMyTarget)
    
    If Len(strMyTarget) = 0 Then
            
        ' if no target was passed, display error and exit
        
        WriteToLog "no url passed. please pass a url and try again"
        Exit Function
        
    End If

    ' if we made it this far, then begin processing passed value
    
    If Left(strMyTarget, 7) <> "http://" And Left(strMyTarget, 8) <> "https://" Then
                    
        ' check to see if there is an http or https at the start. If not, prepend http://
        ' if the proto already exists at the front, then no need to prepend
        
        WriteToLog "can't tell if it is http or https, trying http..."
        strMyTarget = "http://" & strMyTarget
        
    End If
        
    ' let use know what is going on
    
    WriteToLog "Attempting to grab: " & strMyTarget & vbCrLf
    
    
    ' create random temp file name
    
    strTempFileName = ""
    
    For lngKount = 1 To 16
        
        strTempFileName = strTempFileName & Chr(Int(Rnd(1) * 26) + 97)
    
    Next
    RandomSeed = strTempFileName
    strTempFileName = App.Path & "\" & strTempFileName & ".tmp"
       
    ' download file
    
    DownloadFile strMyTarget, strTempFileName
    
    On Error GoTo errhandler
    
    ' display on screen line by line, delete the temp file after
    
    Open strTempFileName For Input As #1
    Do While Not EOF(1)
        Line Input #1, strWork
        cURL = strWork
    Loop
    Close #1

    Kill strTempFileName
    
    On Error GoTo 0
    
alldone:
        
    Exit Function
    
errhandler:
    
If Err = 53 Then
    
    ' the only error I've run across is "file not found", so it is the only error I am trapping for.
    ' this occurs if the site was not reachable, either because a bad URL was specified, network connectivity, or possibly the target site was offline.
    
    WriteToLog vbCrLf & "An error has occured!"
    WriteToLog "---------------------"
    WriteToLog "The site information was not able to be downloaded. Either the host was not resolved, or otherwise unreachable." & vbCrLf & "Sorry about that."
    cURL = "CURLERR_53"
    Resume alldone
    
End If

' just in case another error occurred, let the use know the error number, then resume to exit

WriteToLog vbCrLf & "An unhandled error has occured (" & Err & ")"
Resume alldone

End Function

'Public Function writetolog(sText As String) As Long
'
'    ' vbadvance function
'
'    WriteFile GetStdOutHandle, ByVal (sText & vbCrLf), Len(sText & vbCrLf), writetolog
'
'End Function

Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean
    
    ' vbadvance function
    
    Dim lngRetVal As Long
    
        'Обходим кеширование генерацией рандома
            Randomize
            URL = URL & "?request=" & Int(Rnd * 1283654)
            WriteToLog "Request URL: " & URL
        
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True

End Function

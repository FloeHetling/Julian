Attribute VB_Name = "modLogger"
Option Explicit
' Модуль логгирования
' Использует глобальную переменную LogFile
Enum LoggerMode
    ContinueReport
    StartNewReport
End Enum
Enum wtdfOptions
    wtdfNo
    wtdfYes
End Enum
    Public Const logStartSection = "============= JULIAN ============="
    Public Const logEndLine = "==================================="
    Public Const logStrokeLine = "___________________________________"
    Public Const logEmptyLine = ""

Public Function WriteToLog(ByVal TextLine As String, Optional LoggerMode As LoggerMode, Optional WriteToDebugFile As wtdfOptions)
On Error GoTo LOG_ERROR
Dim iLogFile As Integer

    If LoggerMode = StartNewReport Then
        iLogFile = FreeFile
            If WriteToDebugFile = wtdfYes _
                Then Open LogFile & "debug" For Output As #iLogFile _
                Else Open LogFile For Output As #iLogFile
            Print #iLogFile, TextLine
            Close #iLogFile
    Else
        iLogFile = FreeFile
            If WriteToDebugFile = wtdfYes _
                Then Open LogFile & "debug" For Append As #iLogFile _
                Else Open LogFile For Append As #iLogFile
                Print #iLogFile, TextLine
            Close #iLogFile
    End If

Exit Function
LOG_ERROR:
End
End Function

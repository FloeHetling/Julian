VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Робот Джулиан"
   ClientHeight    =   375
   ClientLeft      =   11760
   ClientTop       =   5085
   ClientWidth     =   3990
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer tCheckConnection 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3480
      Top             =   0
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Робот Джулиан - слежу за состоянием Интернета"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents SysLog As CSocketMaster 'экземпляр СокетМастера для SysLog
Attribute SysLog.VB_VarHelpID = -1
Dim WithEvents SMTP As CSocketMaster 'экземпляр СокетМастера для SMTP
Attribute SMTP.VB_VarHelpID = -1
Dim syslogstring As String
Enum WSockStatus
    SYSLOG_CONNECT
    SYSLOG_SEND
    SYSLOG_QUIT
End Enum
                
Dim WinsockState As WSockStatus
Dim tccSeconds As Long

Public Function SendSysLogString(ByVal Message As String, Optional Severity As slSeverity, Optional Tag As String)
Dim syslogstring As String

On Error GoTo SSLS_KILL_APP
syslogstring = Message

''KISS
If Severity = Null Then Severity = slInformational
If Tag = "" Then Tag = "JULIAN_INFO"

syslogstring = SyslogPackageEncode(slUserLevel, Severity, Now, WatchPoint, Tag, GetCurrentProcessId, syslogstring)

With SysLog
    WriteToLog logStartSection
    WriteToLog logEmptyLine
    WriteToLog "Отправляем строку: " & vbCrLf & syslogstring & vbCrLf
    WriteToLog "Адрес отправки: " & sRemoteHost
        .Protocol = sckUDPProtocol
        .RemoteHost = sRemoteHost
        .RemotePort = lRemotePort
        .SendData syslogstring
        .CloseSck
    WriteToLog logEmptyLine
    WriteToLog "Отправлено!"
    WriteToLog logEndLine
End With
Exit Function

SSLS_KILL_APP:
    Dim errNum As String, errDesc As String
    errNum = Str(Err.Number)
    errDesc = Err.description
WriteToLog Time & " " & Date & ": Ошибка " & errNum & " - " & errDesc
End
End Function

Private Sub Form_Load()
On Error GoTo TIMER_FAULT
    Me.Visible = False
    tCheckConnection.Interval = 1000
    tCheckConnection.Enabled = True
    Set SysLog = New CSocketMaster
    Set SMTP = New CSocketMaster
Exit Sub

TIMER_FAULT:
MsgBox "Произошла ошибка " & Err.Number & " - " & Err.description & vbCrLf & "Вероятно, задан некорректный интервал проверки в ini-файле", vbCritical, JulianVer
End Sub

Private Sub Form_Terminate()
SysLog.CloseSck
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
SysLog.CloseSck
End
End Sub

Private Sub SysLog_Error(ByVal Number As Integer, description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
WriteToLog "==**BLIN, AN ERROR HAS OCURRED"
    'Tell the user that an error occured.  There is more then the numbers below but this is a good starting point.
    WriteToLog Number & " " & description
    WriteToLog logStrokeLine
End Sub

Private Sub tCheckConnection_Timer()
tccSeconds = tccSeconds + 1
If tccSeconds = tCheckInterval Then
    tccSeconds = 0
    CheckConnection
End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''Альтернативный метод оповещения'''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''ПОЧТА SMTP''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub SMTP_DataArrival(ByVal bytesTotal As Long)
    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    
    
    'Retrive data from winsock buffer
    SMTP.GetData strServerResponse
    
    ' Update our text box so we know whats going on.
    WriteToLog strServerResponse
    
    'Get server response code (first three symbols)
    strResponseCode = Left(strServerResponse, 3)
    
    'Only these three codes from the server tell us that the command was accepted
    If strResponseCode = "250" Or strResponseCode = "220" Or strResponseCode = "354" Then
        Select Case WinsockState
            Case MAIL_CONNECT
                WinsockState = MAIL_HELO
                'Remove blank spaces
                strDataToSend = Trim$(FromEmail)
                'Get just the email part of the from line
                strDataToSend = Mid(strDataToSend, 1 + InStr(1, strDataToSend, "<"))
                ' Then get just the account part
                strDataToSend = Left$(strDataToSend, InStr(1, strDataToSend, "@") - 1)
                'Send HELO command to the server
                SMTP.SendData "HELO " & strDataToSend & vbCrLf
            Case MAIL_HELO
                WinsockState = MAIL_FROM
                'Send MAIL FROM command to the server so it knows from who the message comes
                SMTP.SendData "MAIL FROM: " & Mid(FromEmail, InStr(1, FromEmail, "<")) & vbCrLf
            Case MAIL_FROM
                WinsockState = MAIL_RCPTTO
                'Send RCPT TO command to the server so it knows where to send the message
                SMTP.SendData "RCPT TO: " & Mid(ToEmail, InStr(1, ToEmail, "<")) & vbCrLf
            Case MAIL_RCPTTO
                WinsockState = MAIL_DATA
                'Send DATA command to the server so it knows that we want to send the message
                SMTP.SendData "DATA" & vbCrLf
            Case MAIL_DATA
                WinsockState = MAIL_DOT
                'Send header and subject
                SMTP.SendData "Return-Path: <" & FromEmail & ">" & vbCrLf & _
                "Content-type: text/html; charset=Windows-1251" & vbCrLf & _
                "Priority: normal" & vbCrLf & _
                "To: " & ToEmail & vbCrLf & _
                "From: " & FromEmail & vbCrLf & _
                "Subject:" & EMailSubject & vbLf & vbCrLf
                
                '''''
                WriteToLog "Return-Path: <" & FromEmail & ">" & vbCrLf & _
                "Content-type: text/html; charset=UTF-8" & vbCrLf & _
                "Priority: normal" & vbCrLf & _
                "To: " & ToEmail & vbCrLf & _
                "From: " & FromEmail & vbCrLf & _
                "Subject:" & EMailSubject & vbLf & vbCrLf
                '''''
                
                Dim varLines    As Variant
                Dim varLine     As Variant
                Dim strMessage  As String
                
                SMTP.SendData MailReport & vbCrLf & "." & vbCrLf
                WriteToLog MailReport & vbCrLf & "." & vbCrLf
            Case MAIL_DOT
                WinsockState = MAIL_QUIT
                'Send QUIT command
                SMTP.SendData "QUIT" & vbCrLf
                WriteToLog "QUIT" & vbCrLf
            Case MAIL_QUIT
                'Close the connection to the smtp server
                SMTP.CloseSck
        End Select
    Else
        'Check if an error occured
        SMTP.CloseSck
        If Not WinsockState = MAIL_QUIT Then
            'If yes then print the error
            If Left$(strServerResponse, 3) = 421 Then
                WriteToLog "The from email address is invalid for this mail server.  Please check it and try again"
            Else
                WriteToLog "Error: " & strServerResponse
            End If
        Else
            'if the message sent successfully, print it
            WriteToLog "Отчет успешно отправлен"
        End If
    End If
End Sub

Private Sub SMTP_Error(ByVal Number As Integer, description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Tell the user that an error occured.  There is more then the numbers below but this is a good starting point.
   
    If Number = 10049 Then
        WriteToLog "Не могу отправить отчет о ПК - неправильный адрес сервера или порт!"
    ElseIf Number = 10061 Then
        WriteToLog "Сервер почты отклонил мое сообщение. Отчет по ПК не отправлен!"
    ElseIf Number <> 0 Then
        WriteToLog "Ошибка соединения с почтовым сервером: " & Number & vbCrLf & description & vbCrLf & vbCrLf & "Отчет не был отправлен."
    End If
   
    
    SMTP.CloseSck
    
End Sub

Public Function CheckConnection()
'Берем scope провайдера
    Dim OpOctet() As String, ExtIP As String
    ExtIP = cURL(strIpService)
    
    'Выполняем только если не ошибка
    If ExtIP <> "CURLERR_53" Then
        OpOctet = Split(ExtIP, ".")
        Provider = OpOctet(0)
        
        'Проверяем, сменился ли относительно прошлой проверки
        WriteToLog "Проверяю адрес"
    
        If Provider <> PrevProvider _
            And Provider <> (PrevProvider - 1) _
            And Provider <> (PrevProvider - 2) _
            And Provider <> (PrevProvider - 3) _
            And Provider <> (PrevProvider - 4) _
            And Provider <> (PrevProvider - 5) _
            And Provider <> (PrevProvider + 1) _
            And Provider <> (PrevProvider + 2) _
            And Provider <> (PrevProvider + 3) _
            And Provider <> (PrevProvider + 4) _
            And Provider <> (PrevProvider + 5) _
            And PrevProvider <> 0 Then
            
            'Я жаловался что пишу индусский код???
            'ЗАБУДЬТЕ
            'ВОТ ГДЕ ТРУ ИНДИЯ!
            
            Select Case WorkMode
                Case asSyslogAgent
                    SendSysLogString "PROVIDER CHANGED!", slEmergency, "JULIAN_OPCHANGE"
                Case asInderpendedMailer
                'Заполняем параметры
                    SMTPServer = "mail.tvojdoktor.ru"
                    SMTPPort = "25"
                
                'Формируем сообщение
                    FromEmail = _
                                    "Julian <julian@tvojdoktor.ru>"
                    EMailSubject = _
                                    "JULIAN на " & WatchPoint & ": Изменился провайдер"
                    MailReport = _
                                    "Агент Julian на точке наблюдения " & WatchPoint & " сообщил:" & vbCrLf & _
                                    vbCrLf & _
                                    "ПРОВАЙДЕР ИЗМЕНИЛСЯ" & vbCrLf & _
                                    "Возможно, задействован резервный канал связи!" & vbCrLf & vbCrLf & _
                                    "Сообщение получено " & Now & " с IP адреса " & ExtIP
                
                'Выполняем отправку
                    SMTP.Connect Trim(SMTPServer), Val(SMTPPort)
                        'reset the state so our sequence works right
                        WinsockState = MAIL_CONNECT
            End Select
            
            WriteToLog logStartSection
            WriteToLog Now & " - Провайдер сменился!"
            WriteToLog "Октет текущего провайдера: " & Provider
            WriteToLog "Октет предыдущего провайдера: " & PrevProvider
            WriteToLog "Текущий IP: " & ExtIP
            WriteToLog logStrokeLine
        Else
            WriteToLog Now & " - Текущий IP: " & ExtIP
        End If
    
        PrevProvider = OpOctet(0)
    End If
    
Exit Function

SendEmailError:
' Show a detailed error message if needed
If Err.Number <> 0 Then WriteToLog "Ошибка отправки почты: " & vbCrLf & " Error Number: " & Err.Number & _
vbCrLf & "Error Description: " & Err.description & "."
End Function

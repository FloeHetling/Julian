Attribute VB_Name = "modMain"
Option Explicit

    Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Public LogFile As String, INIFile As String, sRemoteHost As String, lRemotePort As Long, Hostname As String, strIpService As String, JulianVer As String
    Public tCheckInterval As Long
    Public Provider As Integer, PrevProvider As Integer
    Public WatchPoint As String, WorkMode As WorkModeConstants
    
    Public Enum WorkModeConstants
        asSyslogAgent = 1
        asInderpendedMailer = 0
    End Enum
    
    Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public CLIArg As String
    
'Почта
''Параметры Winsock
    Public Enum WinsockControlState
        MAIL_CONNECT
        MAIL_HELO
        MAIL_FROM
        MAIL_RCPTTO
        MAIL_DATA
        MAIL_HEADER
        MAIL_DOT
        MAIL_QUIT
    End Enum
    
''Переменные почтового окружения
    Public EMailSubject As String, MailReport As String
    Public SMTPServer As String, SMTPPort As String
    Public FromEmail As String, ToEmail As String


Sub Main()
'PCNAME
    Dim dwLen As Long
        dwLen = MAX_COMPUTERNAME_LENGTH + 1
        Hostname = String(dwLen, "X")
        GetComputerName Hostname, dwLen
        Hostname = Left(Hostname, dwLen)
        
'JULIANVER
JulianVer = App.ProductName & " " & _
                App.Major & "." & App.Minor & _
                "." & App.Revision & " - " & _
                App.CompanyName

'LOGFILE
LogFile = App.Path & "\JULIAN.log"

'INIFILE
INIFile = App.Path & "\JULIAN.ini"

    If CheckPath(INIFile) <> True Then
            WriteToLog " "
            WriteToLog "Не найден файл настроек. Создаем пустой, чтобы модулю сохранения было куда писать настройки"
            WriteToLog " "
            '''''Создаем структуру файла
            Dim iFileNo As Integer
            iFileNo = FreeFile
        
            Open INIFile For Output As #iFileNo
            Print #iFileNo, ";Only Windows-1251 Codepage is allowed!"
            Print #iFileNo, ";Если вы можете прочесть эту строку, ваша кодировка установлена правильно"
            Print #iFileNo, ";Обязательный параметр - название точки наблюдения:"
            Print #iFileNo, "[MAIN]"
            Print #iFileNo, "WatchPoint="
            Print #iFileNo, ";WatchPoint - простая строка. Можно использовать кириллицу, пробелы, спецзнаки."
            Print #iFileNo, ""
            Print #iFileNo, ";=========================== ПАРАМЕТРЫ НИЖЕ НЕ ТРОГАТЬ! ==========================="
            Print #iFileNo, "CheckInterval="
            Print #iFileNo, ";Интервал проверки соединения. Целое число. Измеряется в секундах!"
            Print #iFileNo, ""
            Print #iFileNo, "[SERVICES]"
            Print #iFileNo, "ExternalIPResolverURL="
            Print #iFileNo, ";Адрес сервиса проверки IP. Пример - http://icanhazip.com"
            Print #iFileNo, ""
            Print #iFileNo, "[MAILER]"
            Print #iFileNo, "ToEmail="
            Print #iFileNo, ";Куда слать отчеты. Строго соблюдается формат: Имя пользователя <аккаунт@почтовик.домен>"
            Close #iFileNo
            '''''' И поехали дальше.
            WriteToLog "Создан файл настроек, программа работу завершает. На всякий случай."
            End
    End If
'''''''''''''''''''''''''HELP'''''''''''''''''''''''''''''''

Dim msgHelp As String

        msgHelp = _
        JulianVer & vbCrLf & vbCrLf & _
        "HOTFIX II" & vbCrLf & vbCrLf & _
        "Допустимые параметры командной строки:" & vbCrLf & vbCrLf & _
        "/ipservice - URL сервиса проверки ip." & vbCrLf & _
        "По умолчанию - http://icanhazip.com" & vbCrLf & _
        "Сервер должен выдавать только IP адрес, без доп. строк и HTML" & vbCrLf & _
        "/ipaddress - адрес сервера Syslog с открытым портом 514." & vbCrLf & _
        "Если не указан - ведется прямая передача почтой" & vbCrLf & _
        "/sendtomail - для тестовой отправки на другой почтовый ящик" & vbCrLf & _
        "Если не указан - отправляется на servicedesk@zdravservice.ru" & vbCrLf & _
        "--------------------------" & vbCrLf & vbCrLf & _
        "Прочие параметры:" & vbCrLf & vbCrLf & _
        "/WatchPoint - обозвать точку наблюдения *" & vbCrLf & _
        "/seconds - интервал проверки **" & vbCrLf & vbCrLf & _
        "* - Обзывательство должно быть на латинице и без пробелов" & vbCrLf & _
        "** - Задается в секундах, больших чисел бояться не стоит. Стандарт - 2000"
        CLIArg = Command$
        If CLIArg = "/?" Then
                MsgBox msgHelp, vbInformation, "Справка"
                MsgBox "HOTFIX II:" & vbCrLf & _
                        "Из-за лимита модуля Timer на количество миллисекунд" & vbCrLf & _
                        "программа переведена на отсчет времени в секундах." & vbCrLf & _
                        "Вся информация скорректирована. В ini файл теперь" & vbCrLf & _
                        "тоже вводятся секунды. Переменная типа Long должна" & vbCrLf & _
                        "отработать с запасом.", vbInformation, JulianVer
                End
        End If



''''''''''''''''''''''app dupe check''''''''''''''''''''''''
If App.PrevInstance = True Then
        End
End If


''''''''''''''''''''''CLI ARGS''''''''''''''''''''''''''''''
CLIArg = Command$

    'Если указан адрес сервера сислога - используем его, иначе работаем как независимый почтовик
    If InStr(1, CLIArg, "/ipaddress") <> 0 Then
                Dim IPAddrStrArray() As String, IPAddrStrArrayIdx As Integer
                    IPAddrStrArray = Split(CLIArg, " ")
                    For IPAddrStrArrayIdx = 0 To UBound(IPAddrStrArray)
                        If IPAddrStrArray(IPAddrStrArrayIdx) = "/ipaddress" Then
                            If IPAddrStrArrayIdx + 1 <= UBound(IPAddrStrArray) Then
                                sRemoteHost = IPAddrStrArray(IPAddrStrArrayIdx + 1)
                                lRemotePort = 514
                                WorkMode = asSyslogAgent
                            End If
                        End If
                    Next IPAddrStrArrayIdx
    Else
        WriteToLog "Не указан IP адрес!"
        WriteToLog "Передайте в программу IP адрес сервера Syslog"
        WriteToLog "Синтаксис: /ipaddress XXX.XXX.XXX.XXX"
        WriteToLog "DNS-имена допустимы, но нежелательны. У сервера должен быть открыт UDP-порт 514"
        WriteToLog " "
        WriteToLog "Прим.: Не работает с localhost/127.0.0.1"
        WriteToLog "  "
        WriteToLog "Не передан параметр сервера Syslog. Работаем через почту!"
        WorkMode = asInderpendedMailer
    End If


    
    'если указано кол-во секунд меж проверками - используем. Если нет - значит стандарт - проверка раз в 5 минут
    If InStr(1, CLIArg, "/seconds") <> 0 Then
                Dim tIntervalStrArray() As String, tIntervalStrArrayIdx As Integer
                    tIntervalStrArray = Split(CLIArg, " ")
                    For tIntervalStrArrayIdx = 0 To UBound(tIntervalStrArray)
                        If tIntervalStrArray(tIntervalStrArrayIdx) = "/seconds" Then
                            If tIntervalStrArrayIdx + 1 <= UBound(tIntervalStrArray) Then
                                tCheckInterval = Val(tIntervalStrArray(tIntervalStrArrayIdx + 1))
                            End If
                        End If
                    Next tIntervalStrArrayIdx
    Else
        Dim strCheckInterval As String
        fReadValue INIFile, "MAIN", "CheckInterval", "S", "1", strCheckInterval
        tCheckInterval = Val(strCheckInterval)
        If tCheckInterval = 0 Then tCheckInterval = 1
    End If
    
    'Если указан сервис получения IP адреса - используем его вместо ICanHazIP что по стандарту
    If InStr(1, CLIArg, "/ipservice") <> 0 Then
                Dim IPSrvStrArray() As String, IPSrvStrArrayIdx As Integer
                    IPSrvStrArray = Split(CLIArg, " ")
                    For IPSrvStrArrayIdx = 0 To UBound(IPSrvStrArray)
                        If IPSrvStrArray(IPSrvStrArrayIdx) = "/ipservice" Then
                            If IPSrvStrArrayIdx + 1 <= UBound(IPSrvStrArray) Then
                                strIpService = IPSrvStrArray(IPSrvStrArrayIdx + 1)
                            End If
                        End If
                    Next IPSrvStrArrayIdx
    Else
       fReadValue INIFile, "SERVICES", "ExternalIPResolverURL", "S", "http://icanhazip.com", strIpService
       If strIpService = "" Then strIpService = "http://icanhazip.com"
    End If
    
    'Берем название точки доступа. Если не задан - идет хостнейм
    If InStr(1, CLIArg, "/watchpoint") <> 0 Then
                Dim WatchPointStrArray() As String, WatchPointStrArrayIdx As Integer
                    WatchPointStrArray = Split(CLIArg, " ")
                    For WatchPointStrArrayIdx = 0 To UBound(WatchPointStrArray)
                        If WatchPointStrArray(WatchPointStrArrayIdx) = "/WatchPoint" Then
                            If WatchPointStrArrayIdx + 1 <= UBound(WatchPointStrArray) Then
                                WatchPoint = WatchPointStrArray(WatchPointStrArrayIdx + 1)
                            End If
                        End If
                    Next WatchPointStrArrayIdx
    Else
        fReadValue INIFile, "MAIN", "WatchPoint", "S", "", WatchPoint
        If WatchPoint = "" Then
            WriteToLog "Укажите название точки наблюдения в файле JULIAN.INI", StartNewReport
            WriteToLog "Пока оно не указано, ПО работать не будет!"
            End
        End If
    End If
    
    'оставляем для себя лазеечку на тесты
    If InStr(1, CLIArg, "/sendtomail") <> 0 Then
                Dim SendToMailStrArray() As String, SendToMailStrArrayIdx As Integer
                    SendToMailStrArray = Split(CLIArg, " ")
                    For SendToMailStrArrayIdx = 0 To UBound(SendToMailStrArray)
                        If SendToMailStrArray(SendToMailStrArrayIdx) = "/SendToMail" Then
                            If SendToMailStrArrayIdx + 1 <= UBound(SendToMailStrArray) Then
                                ToEmail = SendToMailStrArray(SendToMailStrArrayIdx + 1)
                            End If
                        End If
                    Next SendToMailStrArrayIdx
    Else
       fReadValue INIFile, "MAILER", "ToEmail", "S", "Servicedesk <servicedesk@zdravservice.ru>", ToEmail
       If ToEmail = "" Then ToEmail = "Servicedesk <servicedesk@zdravservice.ru>"
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' И да, я знаю что это можно было в процедуру, то - в класс а иное - в массив, но у меня, сцука, НЕТ ВРЕМЕНИ наводить здесь красоту. BMSMA.
'Проверяем валидность ввода IP адреса
    Dim asIP() As String, cURLStr As String, OctetIndex As Integer
    cURLStr = cURL(strIpService)
    If cURLStr <> "CURLERR_53" Then
        asIP = Split(cURLStr, ".")
            If UBound(asIP) = 3 Then
                PrevProvider = asIP(0)
                For OctetIndex = 0 To 3
                     If (CInt(asIP(OctetIndex)) > 255) Or (CInt(asIP(OctetIndex)) < 0) Then
                        WriteToLog "Сервер выдачи внешних IP " & strIpService & " выдал какую-то хрень вместо чистого IP адреса"
                        WriteToLog "Можно задать другой сервер запустив прогу с ключом /ipservice http://blablabla.com"
                        End
                     End If
                Next OctetIndex
            Else
                WriteToLog "Сервер выдачи внешних IP " & strIpService & " выдал какую-то хрень вместо чистого IP адреса"
                If InStr(strIpService, ".") = 0 Then WriteToLog "Стоп... " & strIpService & " - это вообще не сайт! Вы издеваетесь?!?"
                WriteToLog "Можно задать другой сервер запустив прогу с ключом /ipservice http://blablabla.com"
                End
            End If
    Else
        WriteToLog "Выдало 53-ю ошибку. Предполагаем что просто нет интернета и продолжаем работу ПО"
        PrevProvider = 0
    End If
    
Load frmMain
End Sub

Public Function CheckPath(strPath As String) As Boolean
    If Dir$(strPath) <> "" Then
        CheckPath = True
    Else
        CheckPath = False
    End If
End Function

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
    
'�����
''��������� Winsock
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
    
''���������� ��������� ���������
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
            WriteToLog "�� ������ ���� ��������. ������� ������, ����� ������ ���������� ���� ���� ������ ���������"
            WriteToLog " "
            '''''������� ��������� �����
            Dim iFileNo As Integer
            iFileNo = FreeFile
        
            Open INIFile For Output As #iFileNo
            Print #iFileNo, ";Only Windows-1251 Codepage is allowed!"
            Print #iFileNo, ";���� �� ������ �������� ��� ������, ���� ��������� ����������� ���������"
            Print #iFileNo, ";������������ �������� - �������� ����� ����������:"
            Print #iFileNo, "[MAIN]"
            Print #iFileNo, "WatchPoint="
            Print #iFileNo, ";WatchPoint - ������� ������. ����� ������������ ���������, �������, ���������."
            Print #iFileNo, ""
            Print #iFileNo, ";=========================== ��������� ���� �� �������! ==========================="
            Print #iFileNo, "CheckInterval="
            Print #iFileNo, ";�������� �������� ����������. ����� �����. ���������� � ��������!"
            Print #iFileNo, ""
            Print #iFileNo, "[SERVICES]"
            Print #iFileNo, "ExternalIPResolverURL="
            Print #iFileNo, ";����� ������� �������� IP. ������ - http://icanhazip.com"
            Print #iFileNo, ""
            Print #iFileNo, "[MAILER]"
            Print #iFileNo, "ToEmail="
            Print #iFileNo, ";���� ����� ������. ������ ����������� ������: ��� ������������ <�������@��������.�����>"
            Close #iFileNo
            '''''' � ������� ������.
            WriteToLog "������ ���� ��������, ��������� ������ ���������. �� ������ ������."
            End
    End If
'''''''''''''''''''''''''HELP'''''''''''''''''''''''''''''''

Dim msgHelp As String

        msgHelp = _
        JulianVer & vbCrLf & vbCrLf & _
        "HOTFIX II" & vbCrLf & vbCrLf & _
        "���������� ��������� ��������� ������:" & vbCrLf & vbCrLf & _
        "/ipservice - URL ������� �������� ip." & vbCrLf & _
        "�� ��������� - http://icanhazip.com" & vbCrLf & _
        "������ ������ �������� ������ IP �����, ��� ���. ����� � HTML" & vbCrLf & _
        "/ipaddress - ����� ������� Syslog � �������� ������ 514." & vbCrLf & _
        "���� �� ������ - ������� ������ �������� ������" & vbCrLf & _
        "/sendtomail - ��� �������� �������� �� ������ �������� ����" & vbCrLf & _
        "���� �� ������ - ������������ �� servicedesk@zdravservice.ru" & vbCrLf & _
        "--------------------------" & vbCrLf & vbCrLf & _
        "������ ���������:" & vbCrLf & vbCrLf & _
        "/WatchPoint - �������� ����� ���������� *" & vbCrLf & _
        "/seconds - �������� �������� **" & vbCrLf & vbCrLf & _
        "* - �������������� ������ ���� �� �������� � ��� ��������" & vbCrLf & _
        "** - �������� � ��������, ������� ����� ������� �� �����. �������� - 2000"
        CLIArg = Command$
        If CLIArg = "/?" Then
                MsgBox msgHelp, vbInformation, "�������"
                MsgBox "HOTFIX II:" & vbCrLf & _
                        "��-�� ������ ������ Timer �� ���������� �����������" & vbCrLf & _
                        "��������� ���������� �� ������ ������� � ��������." & vbCrLf & _
                        "��� ���������� ���������������. � ini ���� ������" & vbCrLf & _
                        "���� �������� �������. ���������� ���� Long ������" & vbCrLf & _
                        "���������� � �������.", vbInformation, JulianVer
                End
        End If



''''''''''''''''''''''app dupe check''''''''''''''''''''''''
If App.PrevInstance = True Then
        End
End If


''''''''''''''''''''''CLI ARGS''''''''''''''''''''''''''''''
CLIArg = Command$

    '���� ������ ����� ������� ������� - ���������� ���, ����� �������� ��� ����������� ��������
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
        WriteToLog "�� ������ IP �����!"
        WriteToLog "��������� � ��������� IP ����� ������� Syslog"
        WriteToLog "���������: /ipaddress XXX.XXX.XXX.XXX"
        WriteToLog "DNS-����� ���������, �� ������������. � ������� ������ ���� ������ UDP-���� 514"
        WriteToLog " "
        WriteToLog "����.: �� �������� � localhost/127.0.0.1"
        WriteToLog "  "
        WriteToLog "�� ������� �������� ������� Syslog. �������� ����� �����!"
        WorkMode = asInderpendedMailer
    End If


    
    '���� ������� ���-�� ������ ��� ���������� - ����������. ���� ��� - ������ �������� - �������� ��� � 5 �����
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
    
    '���� ������ ������ ��������� IP ������ - ���������� ��� ������ ICanHazIP ��� �� ���������
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
    
    '����� �������� ����� �������. ���� �� ����� - ���� ��������
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
            WriteToLog "������� �������� ����� ���������� � ����� JULIAN.INI", StartNewReport
            WriteToLog "���� ��� �� �������, �� �������� �� �����!"
            End
        End If
    End If
    
    '��������� ��� ���� �������� �� �����
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
'' � ��, � ���� ��� ��� ����� ���� � ���������, �� - � ����� � ���� - � ������, �� � ����, �����, ��� ������� �������� ����� �������. BMSMA.
'��������� ���������� ����� IP ������
    Dim asIP() As String, cURLStr As String, OctetIndex As Integer
    cURLStr = cURL(strIpService)
    If cURLStr <> "CURLERR_53" Then
        asIP = Split(cURLStr, ".")
            If UBound(asIP) = 3 Then
                PrevProvider = asIP(0)
                For OctetIndex = 0 To 3
                     If (CInt(asIP(OctetIndex)) > 255) Or (CInt(asIP(OctetIndex)) < 0) Then
                        WriteToLog "������ ������ ������� IP " & strIpService & " ����� �����-�� ����� ������ ������� IP ������"
                        WriteToLog "����� ������ ������ ������ �������� ����� � ������ /ipservice http://blablabla.com"
                        End
                     End If
                Next OctetIndex
            Else
                WriteToLog "������ ������ ������� IP " & strIpService & " ����� �����-�� ����� ������ ������� IP ������"
                If InStr(strIpService, ".") = 0 Then WriteToLog "����... " & strIpService & " - ��� ������ �� ����! �� �����������?!?"
                WriteToLog "����� ������ ������ ������ �������� ����� � ������ /ipservice http://blablabla.com"
                End
            End If
    Else
        WriteToLog "������ 53-� ������. ������������ ��� ������ ��� ��������� � ���������� ������ ��"
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

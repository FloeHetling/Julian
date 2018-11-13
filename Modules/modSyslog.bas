Attribute VB_Name = "modSyslog"
Option Explicit

' 2009-06-11. Gustav Brock.

  ' Syslog Facility designation codes.
  '
  ' Processes and daemons that have not been explicitly
  ' assigned a Facility may use any of the "local use"
  ' facilities or they may use the "user-level" Facility.
  '
  ' Various operating systems have been found to utilize
  ' Facilities 4, 10, 13 and 14 for security/authorization,
  ' audit, and alert messages which seem to be similar.
  ' Various operating systems have been found to utilize
  ' both Facilities 9 and 15 for clock (cron/at) messages.
  Public Enum slFacility
    ' Kernel messages.
    slKernel = 0
    ' User-level messages.
    slUserLevel = 1
    ' Mail system.
    slMailSystem = 2
    ' System daemons.
    slSystemDaemons = 3
    ' Security/authorization messages.
    slSecurity = 4
    ' Messages generated internally by syslogd.
    slSyslogd = 5
    ' Line printer subsystem
    slLinePrinter = 6
    ' Network news subsystem.
    slNetworkNews = 7
    ' UUCP subsystem.
    slUucpSubsystem = 8
    ' Clock daemon 1.
    slClockDaemon1 = 9
    ' Security/authorization messages.
    slAuthorization = 10
    ' FTP daemon.
    slFtpDaemon = 11
    ' NTP subsystem.
    slNtpSubsystem = 12
    ' Log audit.
    slLogAudit = 13
    ' Log alert.
    slLogAlert = 14
    ' Clock daemon 2.
    slClockDaemon2 = 15
    ' Local use 0 (local0).
    slLocalUse0 = 16
    ' Local use 1 (local1).
    slLocalUse1 = 17
    ' Local use 2 (local2).
    slLocalUse2 = 18
    ' Local use 3 (local3).
    slLocalUse3 = 19
    ' Local use 4 (local4).
    slLocalUse4 = 20
    ' Local use 5 (local5).
    slLocalUse5 = 21
    ' Local use 6 (local6).
    slLocalUse6 = 22
    ' Local use 7 (local7).
    slLocalUse7 = 23
    ' Minimum, maximum and default values.
    slFacilityMin = slKernel
    slFacilityMax = slLocalUse7
    slFacilityDefault = slUserLevel
  End Enum
  
  ' Syslog Severity level codes.
  '
  ' Each message Priority also has a decimal Severity level
  ' indicator.
  Public Enum slSeverity
    ' System is unusable.
    slEmergency = 0
    ' Action must be taken immediately.
    slAlert = 1
    ' Critical conditions.
    slCritical = 2
    ' Error conditions.
    slError = 3
    ' Warning conditions.
    slWarning = 4
    ' Normal but significant condition.
    slNotice = 5
    ' Informational messages.
    slInformational = 6
    ' Debug-level messages.
    slDebug = 7
    ' Minimum, maximum and default values.
    slSeverityMin = slEmergency
    slSeverityMax = slDebug
    slSeverityDefault = slNotice
  End Enum

  ' To create Priority, multiply Facility with this factor
  ' and then add Severity.
  Private Const cbytFacilityFactor  As Byte = 8
  ' Length of RFC 3164 time string per definition.
  Private Const cintRfc3164LenTime  As Integer = 15
  ' Maximum length of RFC 3164 tag string per definition.
  Private Const cintRfc3164LenTag   As Integer = 32
  ' Maximum length of RFC 3164 package per definition.
  Private Const cintRfc3164LenPack  As Integer = 1024
  
  ' Local host name.
  Private Const cstrHostNameLocal   As String = "localhost"
  ' Unknown host address.
  Private Const cstrHostAddressNone As String = "0.0.0.0"
  ' Separator between all parts (except between PRI and
  ' TIMESTAMP and between TAG and PID) of a package.
  Private Const cstrPartsSeparator  As String = " "
  ' Default TAG for empty tag to send.
  ' Caution:
  ' Could be set to a zero length string but should be set to something.
  ' If set to a zero length string and a package with no PID is sent,
  ' the first word of MESSAGE will be read as a TAG.
  Private Const cstrTagEmpty        As String = "Test"
  ' Default TAG for empty tag received.
  ' Note:
  ' An empty string for cstrTagNone will not be saved in
  ' tblSyslog by function SyslogEntrySave.
  Private Const cstrTagNone         As String = "Nil"
  ' Default message for empty message, received or to send.
  Private Const cstrContentEmpty    As String = "<no message>"

Public Function SyslogPackageEncode( _
  Optional ByVal bytFacility As slFacility, _
  Optional ByVal bytSeverity As slSeverity, _
  Optional ByVal datTimestamp As Date, _
  Optional ByVal strHostname As String, _
  Optional ByVal strTag As String, _
  Optional ByVal lngPid As Long, _
  Optional ByVal strMessage As String) _
  As String
  
' Assembles a syslog message package from its possible parts
' according to RFC 3164.
  
  Dim strPackagePri         As String
  Dim strPackageHeader      As String
  Dim strPackageMsg         As String
  Dim strPackage            As String
  
  ' Assemble the three parts of the package.
  strPackagePri = StrSyslogPriority(bytFacility, bytSeverity)
  strPackageHeader = StrSyslogTimestamp(datTimestamp) & cstrPartsSeparator & _
    StrSyslogHostname(strHostname)
  strPackageMsg = StrSyslogMessage(strTag, lngPid, strMessage)
  
  ' Assemble and limit the package.
  strPackage = Left( _
    strPackagePri & _
    strPackageHeader & cstrPartsSeparator & _
    strPackageMsg, _
    cintRfc3164LenPack)

  SyslogPackageEncode = strPackage

End Function
  
Public Function StrSyslogPriority( _
  ByVal bytFacility As slFacility, _
  ByVal bytSeverity As slSeverity, _
  Optional ByVal bytPriority As Byte) _
  As String
  
' Calculates and formats PRI, the Priority code for a syslog message
' according to RFC 3164.
'
' Examples:
'   <0>
'   <13>
'   <165>
'
' If parameter bytPriority is zero, the values for bytFacility and
' bytSeverity are used.
' If parameter bytPriority is larger than zero, this value is used and
' bytFacility and bytSeverity are ignored.
' If a passed value is larger than allowed, it will be replaced by a
' default value.

  ' Pre- and suffix of a qualified PRI string.
  Const cstrPriHead As String = "<"
  Const cstrPriTail As String = ">"
  
  Dim strPriority   As String
  
  ' Validate and compose the priority code value.
  ' This is returned in bytPriority.
  Call ValidatePriority(bytFacility, bytSeverity, bytPriority)
  
  strPriority = cstrPriHead & CStr(bytPriority) & cstrPriTail
    
  StrSyslogPriority = strPriority
  
End Function

Public Function StrSyslogTimestamp( _
  Optional ByVal datTimestamp As Date) _
  As String
  
' Calculates and formats TIMESTAMP, the date/time part for a syslog message
' according to RFC 3164.
'
' The TIMESTAMP field is the local time and is in the format of:
'   "Mmm [<space>|d]d hh:nn:ss"
' as this format is used for days of 1 to 9:
'   "Mmm  d hh:nn:ss"
' and this format is used for days of 10 to 31:
'   "Mmm dd hh:nn:ss"
'
' Mmm is the English language abbreviation for the month of the year
' with the first character in uppercase and the other two characters
' in lowercase.
'
' Examples:
'   Feb 12 18:27:09
'   Oct  7 08:09:00
'
' If parameter datTimestamp is zero or missing, it will be replaced with
' the current time.

  ' The only acceptable values for the month abbreviations.
  Const cstrMonths    As String = "Jan;Feb;Mar;Apr;May;Jun;Jul;Aug;Sep;Oct;Nov;Dec"
  Const cstrMonthsSep As String = ";"

  Dim intTimestampLen As Integer
  Dim strTimestamp    As String
  Dim strMonth        As String
  Dim strDay          As String
  Dim strTime         As String
  
  If CDbl(datTimestamp) = 0 Then
    ' No date/time specified. Use current time.
    datTimestamp = Now
  End If
  ' Create the three parts of TIMESTAMP.
  ' Extract month abbreviation from zero based list for the month part.
  strMonth = Split(cstrMonths, cstrMonthsSep)(Month(datTimestamp) - 1)
  ' Create an empty day part with leading space as separator.
  strDay = Space(1 + 2)
  ' Insert the day value right justified.
  RSet strDay = Str(Day(datTimestamp))
  ' Format the time part with a leading space as separator.
  strTime = Format(datTimestamp, " hh\:nn\:ss")
  ' Assemble TIMESTAMP.
  strTimestamp = strMonth & strDay & strTime
  
  StrSyslogTimestamp = strTimestamp

End Function

Public Function StrSyslogHostname( _
  Optional ByVal strHostname As String, _
  Optional ByVal booUnderscoreValid As Boolean, _
  Optional ByVal booIpAddressLookUp As Boolean) _
  As String
  
' Verifies or retrieves HOSTNAME, the originator part for a syslog message
' according to RFC 3164.
' Optionally, underscore is allowed in the host name as sometimes used in
' Windows environments. This is, however, against the RFC standards.
' A host name is preferred for HOSTNAME but the IP address of the host can
' be used as well.
'
' If no parameters are passed or if the passed strHostname cannot be
' verified, the host name of the local machine is returned.
' Optionally, the IP address of the passed host name is looked up and
' returned if parameter booIpAddressLookUp is True.
' If anything else fails, localhost or 127.0.0.1 is returned.
  
  Dim strIpAddress            As String

  If booIpAddressLookUp = False Then
    If IsHostname(strHostname) Then
      ' Use strHostname as passed but in lowercase.
      strHostname = LCase(strHostname)
    Else
      ' Retrieve host name of this machine.
      strHostname = MachineHostName()
    End If
    If Len(strHostname) = 0 Then
      strHostname = cstrHostNameLocal
    End If
  Else
    If IsHostname(strHostname) = False Then
      ' No reason to look up the address.
      strIpAddress = cstrHostAddressNone
    Else
      strIpAddress = MachineHostAddress(strHostname)
    End If
    If strIpAddress = cstrHostAddressNone Then
      ' Host name could not be resolved.
      If IsHostname(strHostname) Then
        ' Use strHostname as passed.
        strIpAddress = strHostname
      Else
        ' Use local host address.
        strIpAddress = MachineHostAddress()
        If strIpAddress = cstrHostAddressNone Then
          ' Not very likely but ...
          strIpAddress = MachineHostAddress(cstrHostNameLocal)
        End If
      End If
    End If
    strHostname = LCase(strIpAddress)
  End If
  
  StrSyslogHostname = strHostname
  
End Function

Public Function StrSyslogMessage( _
  Optional ByVal strTag As String, _
  Optional ByVal lngPid As Long, _
  Optional ByVal strContent As String, _
  Optional ByVal booUsAsciiOnly As Boolean) _
  As String
  
' Builds MESSAGE, the message part for a syslog message according to
' the rules and recommendations in RFC 3164.
'
' If no CONTENT is passed, a default message is used.
' If no TAG and PID is passed, a default TAG is used.
' Total length of TAG and PID is 32 characters.
' If booUsAsciiOnly is True, high ascii will be filtered;
' invalid characters will be replaced with underscores.
'
'   Format:
'     TAG[PID]: CONTENT
'
'   Examples:
'     DataExport[123]: Success. 312 MB transferred.
'     Import: Failure.
'     [456]: Running time 11:27.
'     Test: <no message>
'     Nil: This is a test.
'     Nil: <no message>

  Const cstrTagContentSep   As String = ": "
  Const cstrStripContentSep As String = " "
  Const cintTagLenMax       As Integer = 32
  
  ' Set to True if a PID should be stripped from TAG:
  Const cbooStripPidFromTag As Boolean = False
  
  Dim strPid                As String
  Dim strTagPid             As String
  Dim strTagContentSep      As String
  Dim strStripContentSep    As String
  Dim strTagStripped        As String
  Dim strMessage            As String
  
  Call CleanMessage(strTag, True)
  Call CleanMessage(strContent, booUsAsciiOnly)
  Call StripTag(strTag, strTagStripped, cbooStripPidFromTag)
  
  strPid = FormatPid(lngPid)
  If Len(strTag) + Len(strPid) = 0 Then
    ' Neither a TAG nor a PID is supplied.
    ' Use default TAG.
    strTag = cstrTagEmpty
  End If
  ' Concatenate TAG and PID.
  strTagPid = Left(strTag, cintTagLenMax - Len(strPid)) & strPid
  If Len(strTagPid) > 0 Then
    ' Only possible if default TAG is a zero length string.
    strTagContentSep = cstrTagContentSep
  End If
  If Len(strTagStripped) > 0 Then
    ' Insert a separator between TAG and the stripped part of strTag.
    strStripContentSep = cstrStripContentSep
  End If
  If Len(strContent) = 0 Then
    ' No CONTENT is supplied.
    ' Use default CONTENT.
    strContent = cstrContentEmpty
  End If
  ' Assemble MSG part.
  strMessage = _
    strTagPid & strTagContentSep & _
    strTagStripped & strStripContentSep & strContent
  
  StrSyslogMessage = strMessage

End Function

Public Function IsHostname( _
  ByVal strHostname As String, _
  Optional ByVal booUnderscoreValid As Boolean) _
  As Boolean
  
' Verifies if strHostname represents a possible host name:
'   - Complete host name has a maximum length of 255 characters.
'   - Each label has a maximum length of 63 characters.
'   - Labels are separated by a dot.
'   - Allowed characters are a-z, 0-9 and hyphen only.
'   - A hyphen cannot be neither the first nor the last character
'     of a label.
'   - Optionally, underscore is allowed as sometimes used in
'     Windows environments, though this is not according to the
'     RFC standards.

  ' Minimum length of a host name.
  Const cintHostnameLenMin  As Integer = 1
  ' Maximum length of a host name.
  Const cintHostnameLenMax  As Integer = 255
  ' Label separator in a host name.
  Const cstrLabelSeparator  As String = "."
  
  ' Split the host name into labels and validate these.
  
  Dim astrLabel   As Variant
  Dim intLen      As Integer
  Dim bytLabel    As Byte
  Dim booInvalid  As Boolean
  
  intLen = Len(strHostname)
  If intLen < cintHostnameLenMin Then
    ' No host name passed.
    booInvalid = True
  ElseIf intLen > cintHostnameLenMax Then
    ' Host name too long.
    booInvalid = True
  Else
    ' Split host name into labels.
    astrLabel = Split(strHostname, cstrLabelSeparator)
    For bytLabel = LBound(astrLabel) To UBound(astrLabel)
      If Not IsHostnameLabel(astrLabel(bytLabel), booUnderscoreValid) Then
        booInvalid = True
      End If
    Next
  End If

  IsHostname = Not booInvalid
  
End Function

Public Function IsHostnameLabel( _
  ByVal strHostnameLabel As String, _
  Optional ByVal booUnderscoreValid As Boolean) _
  As Boolean

' Verifies if strHostnameLabel contains allowed characters only and
' does not exceed the maximum allowed length.
' If booUnderscoreValid is True, character underscore is allowed as
' sometimes used in Windows environments, though this is not
' according to the RFC standards.

  ' Minimum length of a label of a host name.
  Const cintLabelLenMin     As Integer = 1
  ' Maximum length of a label of a host name.
  Const cintLabelLenMax     As Integer = 63
  
  ' Allowed characters in a label of a host name.
  '   -0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ
  ' Lower case letters are allowed as well.
  ' Hyphen is not allowed as the first or the last character.
  ' Underscore may be allowed.
  ' Equivalent ascii value ranges for the allowed characters.
  '   Hyphen: 45
  '   0 to 9: 48-57
  '   A to Z: 65-90
  '   a to z: 97-122
  '   Underscore: 95
  
  Dim intPos      As Integer
  Dim intLen      As Integer
  Dim booInvalid  As Boolean
  
  intLen = Len(strHostnameLabel)
  If intLen < cintLabelLenMin Then
    ' No label passed.
    booInvalid = True
  ElseIf intLen > cintLabelLenMax Then
    ' Label is too long.
    booInvalid = True
  Else
    ' Check each character.
    For intPos = 1 To intLen
      Select Case Asc(Mid(strHostnameLabel, intPos, 1))
        Case 48 To 57, 65 To 90, 97 To 122
          ' Digits and letters.
          ' OK.
        Case 45
          ' Hyphen.
          If intPos = 1 Or intPos = intLen Then
            ' No leading or trailing hyphen is allowed.
            booInvalid = True
          Else
            ' OK.
          End If
        Case 95
          ' Underscore.
          If booUnderscoreValid = True Then
            ' OK.
          Else
            booInvalid = True
          End If
        Case Else
          ' Anything else is not allowed.
          booInvalid = True
      End Select
    Next
  End If
  
  IsHostnameLabel = Not booInvalid
  
End Function

Public Sub CleanMessage( _
  ByRef strMessage As String, _
  Optional ByVal booUsAsciiOnly As Boolean, _
  Optional ByVal strSubstitute As String = "_")
  
' Cleans a message string by replacing control characters
' with a substitute. Default substitute is underscore, Chr(95).
' Optionally, high ascii can be substituted as well.

  ' Ascii values for allowed characters for full ascii:
  '    32 to 126
  '   128 to 255
  ' Ascii values for allowed characters for US ascii only:
  '    32 to 126
  
  Dim intPos      As Integer
  Dim intLen      As Integer
  
  intLen = Len(strMessage)
  ' Check each character.
  For intPos = 1 To intLen
    Select Case Asc(Mid(strMessage, intPos, 1))
      Case 32 To 126
        ' Low ascii.
        ' OK.
      Case 128 To 255
        ' High ascii.
        If booUsAsciiOnly = True Then
          ' Replace character.
          Mid(strMessage, intPos) = strSubstitute
        Else
          ' OK.
        End If
      Case 0 To 31
        ' Control character.
        ' Replace character.
        Mid(strMessage, intPos) = strSubstitute
    End Select
  Next

End Sub

Public Sub StripTag( _
  ByRef strTag As String, _
  ByRef strStripped As String, _
  Optional ByVal booStripPidFromTag As Boolean)
  
' Strips TAG from the first not allowed character.
' A remaining part is returned in strStripped.
' The maximum length of a TAG is 32 characters. If
' this value is exceeded, the full strTag will be
' moved to strStripped.
'
' If the not allowed character is space or colon, it
' will be excluded from the stripped part.
'
' Optionally, left square bracket is disallowed as well,
' which is recommended by RFC 3164. However, this may
' indicate an included process id which should not
' be stripped.
' To strip a PID, set booStripPidFromTag to True.

  ' Ascii values for not allowed characters:
  '   Space: 32
  '   Colon: 58
  ' Optionally:
  '   Left square bracket: 91
  
  Dim intPos      As Integer
  Dim intLen      As Integer
  Dim intTag      As Integer
  Dim intCut      As Integer
  
  intLen = Len(strTag)
  intTag = intLen
  intCut = intLen
  ' Check each character.
  While intPos < intLen
    intPos = intPos + 1
    Select Case Asc(Mid(strTag, intPos, 1))
      Case 32, 58
        ' Not allowed character found.
        ' Set cutting positions.
        intTag = intPos - 1
        intCut = intPos
        ' Stop loop.
        intLen = intPos
      Case 91
        If booStripPidFromTag = True Then
          ' Not allowed character found.
          ' Set cutting positions.
          intTag = intPos - 1
          If intPos = intLen Then
            intCut = intPos
          Else
            ' Preserve left bracket.
            intCut = intPos - 1
          End If
          ' Stop loop.
          intLen = intPos
        End If
      Case Else
        If intPos > cintRfc3164LenTag Then
          ' If strTag exceeds the maximum length,
          ' it is not a TAG.
          ' Return full strTag in strStripped.
          intTag = 0
          intCut = 0
          ' Stop loop.
          intLen = intPos
        End If
    End Select
  Wend
  
  ' Return strings.
  strStripped = Trim(Mid(strTag, 1 + intCut))
  strTag = Trim(Left(strTag, intTag))

End Sub

Public Sub StripPid( _
  ByRef strTag As String, _
  ByRef lngPid As Long)
  
' Strips a PID, a process id part of TAG for a syslog message, from
' a TAG.
' The process id value (without brackets) is supposed to be numeric,
' a positive Long.
' A value of zero means that PID is undetermined.
'
'   Examples of TAG strings with a valid PID:
'     SomeTag[2]
'     AnotherTag[678]
'     SomeOtherTag[16541]
'
' If the process id itself is not numeric or is larger than a Long,
' zero (undetermined) is returned.

  Const cstrPidHead       As String = "["
  Const cstrPidTail       As String = "]"
  
  Dim strTmp      As String
  
  Dim intPos      As Integer
  Dim intLen      As Integer
  
  On Error GoTo Err_StripPid
  
  intLen = Len(strTag)
  If intLen > 0 Then
    If strTag Like "*[" & cstrPidHead & "]*[" & cstrPidTail & "]" Then
      intPos = InStrRev(strTag, cstrPidHead)
      strTmp = Mid(strTag, 1 + intPos)
      lngPid = Val(strTmp)
      strTag = Left(strTag, intPos - 1)
    End If
  End If
  
Exit_StripPid:
  Exit Sub
  
Err_StripPid:
  WriteToLog Err.Number, "PID: " & strTmp
  Resume Exit_StripPid

End Sub

Public Sub ValidatePriority( _
  ByVal bytFacility As slFacility, _
  ByVal bytSeverity As slSeverity, _
  Optional ByRef bytPriority As Byte)
  
' Calculates or validates a priority code value for a syslog message
' according to RFC 3164.
'
' Examples of returned bytPriority values:
'   0
'   13
'   165
'
' By wrapping the priority code value in angle brackets, a PRI can be
' obtained, like:
'
'   <124>
'
' If parameter bytPriority is zero, the values for bytFacility and
' bytSeverity are used.
' If parameter bytPriority is larger than zero, this value is used and
' bytFacility and bytSeverity are ignored.
' If a passed value is larger than allowed, it will be replaced by a
' default value.

  ' The Priority value is calculated by first multiplying the
  ' Facility number by 8 and then adding the numerical value of the
  ' Severity.
  ' Thus, values are in the range from 0 to 23 * 8 + 7 = 191.
  Const cbytPriorityMin     As Byte = slFacilityMin * cbytFacilityFactor + slSeverityMin
  Const cbytPriorityMax     As Byte = slFacilityMax * cbytFacilityFactor + slSeverityMax
  ' If the Priority value is unidentifiable, then a value of 13 must be used.
  Const cbytPriorityDefault As Byte = slFacilityDefault * cbytFacilityFactor + slSeverityDefault
  
  If bytPriority = 0 Then
    ' Calculate Priority code from Facility and Severity values.
    If bytFacility > slFacilityMax Then
      ' Invalid Facility value. Use default Facility value.
      bytFacility = slFacilityDefault
    End If
    If bytSeverity > slSeverityMax Then
      ' Invalid Severity value. Use default Severity value.
      bytSeverity = slSeverityDefault
    End If
    bytPriority = bytFacility * cbytFacilityFactor + bytSeverity
  ElseIf bytPriority > cbytPriorityMax Then
    ' Invalid Priority value. Use default Priority value.
    bytPriority = cbytPriorityDefault
  Else
    ' Use passed Priority value.
  End If
  
End Sub

Public Function Priority( _
  ByVal bytFacility As slFacility, _
  ByVal bytSeverity As slSeverity) _
  As Byte
  
' Assembles Priority value from Facility and Severity values
' after validation of these.
' Invalid parameter values are replaced with default values.
  
  Dim bytPriority As Byte
  
  Call ValidatePriority(bytFacility, bytSeverity, bytPriority)
  
  Priority = bytPriority

End Function

Public Sub DisassemblePriority( _
  ByRef bytFacility As slFacility, _
  ByRef bytSeverity As slSeverity, _
  ByVal bytPriority As Byte)
  
' Retrieves Facility code and Severity code from a priority code value.
' If a passed priority code cannot be resolved, default values for
' Facility code or Severity code or both are returned.

  ' Validate the priority code value.
  ' If validation fails, a priority code value is returned of which
  ' Facility code and/or Severity code has been replaced by its default
  ' value.
  Call ValidatePriority(bytFacility, bytSeverity, bytPriority)
  ' Decompose and return the priority code.
  bytFacility = bytPriority \ cbytFacilityFactor
  bytSeverity = bytPriority Mod cbytFacilityFactor

End Sub

Public Function FormatPid( _
  ByVal lngPid As Long) _
  As String
  
' Formats a PID, a process id part of TAG for a syslog message, as
' recommended in RFC 3164.
' The process id value (without brackets) is supposed to be a
' positive Long.
'
'   Examples of formatted PIDs:
'     [2]
'     [678]
'     [16541]
'
' If the process id value is zero (undetermined) or negative, an
' empty string is returned.

  Const cstrPidHead       As String = "["
  Const cstrPidTail       As String = "]"
  
  Dim strPid              As String
  
  If lngPid > 0 Then
    strPid = cstrPidHead & CStr(lngPid) & cstrPidTail
  End If
  
  ' Return formatted PID.
  FormatPid = strPid
  
End Function

Public Function IsDigits( _
  ByVal varExpression As Variant) _
  As Boolean
  
  ' Returns True if varExpression contains digits only.
  '
  ' This is more restrictive than IsNumeric which also accepts
  ' expressions as 2E7 and 3D4 as well as leading and
  ' trailing plus or minus sign or spaces, and hexadecimal or
  ' octal declared numbers as strings like "&H100" and "&O70",
  ' and decimals.
  '
  ' Equivalent ascii value ranges for the allowed characters.
  '   0 to 9: 48-57
  
  Dim strExpression As String
  Dim intLen        As Integer
  Dim intPos        As Integer
  Dim booNonNumeric As Boolean
  
  If IsNumeric(varExpression) Then
    strExpression = CStr(varExpression)
    intLen = Len(strExpression)
    ' Verify each character.
    For intPos = 1 To intLen
      Select Case Asc(Mid(strExpression, intPos, 1))
        Case 48 To 57
          ' Character is a digit.
        Case Else
          ' Character is something else
          booNonNumeric = True
          Exit For
      End Select
    Next
  Else
    ' Return False at once.
    booNonNumeric = True
  End If
  
  IsDigits = Not booNonNumeric
  
End Function

Public Function SyslogPackageDecode( _
  ByVal strPackage As String, _
  Optional ByVal strRemoteAddress As String, _
  Optional ByRef bytFacility As slFacility, _
  Optional ByRef bytSeverity As slSeverity, _
  Optional ByRef datTimestamp As Date, _
  Optional ByRef strHostname As String, _
  Optional ByRef strTag As String, _
  Optional ByRef lngPid As Long, _
  Optional ByRef strContent As String) _
  As Boolean

' Reads a syslog message package and returns its parts
' according to RFC 3164.
' If the package is not fully compliant, default values
' are returned for the unreadable parts.

  ' PRI is a number wrapped in angle brackets.
  Const cstrPriHead         As String = "<"
  Const cstrPriTail         As String = ">"
  ' Maximum Byte value.
  Const cbytMax             As Byte = &HFF
  ' Allow underscore in host names or not.
  Const cbooUnderscoreValid As Boolean = True
  
  Dim avarTmp               As Variant
  
  Dim strTemp               As String
  Dim bytPriority           As Byte
  Dim intPriority           As Integer
  Dim booErrorTime          As Boolean
  Dim booFullyDecoded       As Boolean
  
  ' Preset default values.
  ' Facility and Severity.
  bytFacility = slFacilityDefault
  bytSeverity = slSeverityDefault
  datTimestamp = Now
  
  ' Trim and limit the package.
  strPackage = Left(Trim(strPackage), cintRfc3164LenPack)
  If Len(strPackage) = 0 Then
    ' Empty package. Nothing to do.
  ElseIf Left(strPackage, 1) <> cstrPriHead Then
    ' This is not a syslog package.
    ' Move the complete package to the message part.
    strContent = strPackage
  ElseIf Abs(InStr(1 + 1 + 1, strPackage, cstrPriTail) - 3) < 3 Then
    ' The header may contain a priority code.
    intPriority = Val(Mid(strPackage, 1 + 1, 3))
    ' Limit the range of the priority code value.
    If intPriority > cbytMax Then
      bytPriority = cbytMax
    ElseIf intPriority < 0 Then
      bytPriority = cbytMax
    Else
      bytPriority = intPriority
    End If
    ' Validate and the priority code value and split it into
    ' its parts and return the values.
    Call DisassemblePriority(bytFacility, bytSeverity, bytPriority)
    ' Split the package in two parts to strip Priority from
    ' the remaining part.
    strTemp = Split(strPackage, cstrPriTail, 2)(1)
    ' Extract the date/time from the remaining package.
    datTimestamp = DateValueRfc3164(strTemp, booErrorTime)
    If booErrorTime = True Then
      ' The TIMESTAMP part could not be fully read.
      ' Replace the extracted time with the local time.
      datTimestamp = Now
      ' Return the remaining part of the package as the message.
      strContent = strTemp
    Else
      ' Strip TIMESTAMP part from the remaining part of the package.
      strTemp = Trim(Mid(strTemp, 1 + cintRfc3164LenTime))
      ' Split the remaining package in its two parts, HOSTNAME
      ' and MSG.
      avarTmp = Split(strTemp, cstrPartsSeparator, 2)
      If Not IsHostname(avarTmp(0), cbooUnderscoreValid) Then
        If Len(strRemoteAddress) = 0 Then
          ' Return unknown host address as host name.
          strHostname = cstrHostAddressNone
        Else
          ' Return remote host address as host name.
          strHostname = strRemoteAddress
        End If
        ' Return the remaining part of the package as the message.
        strContent = strTemp
      Else
        ' Return HOSTNAME from the remaining part of the package.
        strHostname = avarTmp(0)
        ' Strip HOSTNAME from the remaining part of the package to
        ' obtain MSG.
        strTag = avarTmp(1)
        ' Split this into a possible TAG and the CONTENT.
        Call StripTag(strTag, strContent)
        ' Find out if a PID is present. If so, retrieve it.
        Call StripPid(strTag, lngPid)
        Call CleanMessage(strTag, True)
        booFullyDecoded = True
      End If
    End If
  Else
    ' Priority is missing or not readable.
    ' Move the complete package to the message part.
    strContent = strPackage
  End If
  
  If Len(strHostname) = 0 Then
    strHostname = cstrHostAddressNone
  End If
  If Len(strTag) = 0 Then
    strTag = cstrTagNone
  End If
  If Len(strContent) = 0 Then
    strContent = cstrContentEmpty
  Else
    ' Clean the CONTENT part for possible invalid characters.
    Call CleanMessage(strHostname, False)
    ' Limit the CONTENT part so if the package was rebuilt from the
    ' decoded parts, the length of the package would be within the
    ' allowed length.
    Call LimitContent(bytFacility, bytSeverity, _
      strHostname, strTag, lngPid, _
      strContent)
  End If

  SyslogPackageDecode = booFullyDecoded

End Function

Public Function DateValueRfc3164( _
  ByVal strRfc3164time As String, _
  Optional ByRef booError As Boolean) _
  As Date

' Converts string in RFC 3164 time format as used in syslog messages:
'
'   "Mmm [<space>|d]d hh:mm:ss"
'
' to VB datetime value.
'
' If strRfc3164time cannot be resolved, local (current) time is returned
' and an error flag is returned in booError.

  ' Leng of month abreviation.
  Const cintMonthLen      As Integer = 3
  ' Length of time string: "hh:nn:ss"
  Const cintTimeLen       As Integer = 8
  ' Date separator.
  Const cstrDateSeparator As String = "/"
  ' Localized month abreviations cannot be used and are not accepted.
  ' English must be used.
  Const cstrMonthsEnglish As String = "JanFebMarAprMayJunJulAugSepOctNovDec"
    
  Dim datDate       As Date
  Dim datTime       As Date
  Dim datLocal      As Date
  Dim datTimeSyslog As Date
  Dim strTimeSyslog As String
  Dim strTime       As String
  Dim strMonth      As String
  Dim intYear       As Integer
  Dim intMonth      As Integer
  Dim intDay        As Integer
  Dim intMonthDiff  As Integer
  Dim booSuccess    As Boolean
  
  datLocal = Now
  ' Trim and limit string.
  strTimeSyslog = RTrim(Left(Trim(strRfc3164time), cintRfc3164LenTime))
  If Len(strTimeSyslog) = cintRfc3164LenTime Then
    ' Assume the current year to be the year of local date.
    intYear = Year(datLocal)
    ' Header contains month spelled as a month abbreviation in English.
    strMonth = Left(strTimeSyslog, cintMonthLen)
    ' Find month value from list of English abreviations.
    intMonth = (InStr(1, cstrMonthsEnglish, strMonth, vbTextCompare) - 1 + cintMonthLen) \ cintMonthLen
    If intMonth = 0 Then
      ' An English abbreviation was not used.
      ' Assume local month.
      intMonth = Month(datLocal)
    End If
    ' Extract day value from reminder of strRfc3164time.
    intDay = Val(Mid(strTimeSyslog, 2 + cintMonthLen, 2))
    If Not IsDate(intYear & cstrDateSeparator & intMonth & cstrDateSeparator & intDay) Then
      ' Date string is invalid.
      ' Assume local date.
      intMonth = Month(datLocal)
      intDay = Day(datLocal)
    End If
    datDate = DateSerial(intYear, intMonth, intDay)
    ' Calculate difference in months between local date and reported syslog date.
    intMonthDiff = DateDiff("m", datDate, datLocal)
    If Abs(intMonthDiff) >= 6 Then
      ' The difference between local date and date of strRfc3164time is big.
      ' Adjust syslog date one year to match local date closer.
      datDate = DateAdd("yyyy", Sgn(intMonthDiff), datDate)
    End If
    
    ' Extract time string from strRfc3164time.
    strTime = Right(strTimeSyslog, cintTimeLen)
    If IsDate(strTime) Then
      ' Calculate time value from time string.
      datTime = TimeValue(strTime)
      ' Return success.
      booSuccess = True
    Else
      ' Time string is invalid.
      ' Assume local time.
      datTime = TimeSerial(Hour(datLocal), Minute(datLocal), Second(datLocal))
    End If
    datTimeSyslog = datDate + datTime
  Else
    ' Use local time.
    datTimeSyslog = datLocal
  End If
  ' Return error condition.
  booError = Not booSuccess
  
  ' Return calculated date/time.
  DateValueRfc3164 = datTimeSyslog

End Function

Public Sub LimitContent( _
  ByVal bytFacility As slFacility, _
  ByVal bytSeverity As slSeverity, _
  ByVal strHostname As String, _
  ByVal strTag As String, _
  ByVal lngPid As Long, _
  ByRef strContent As String)
  
' Limits length of strContent to bring the length of a syslog package
' assembled from the passed parts withing the maximum allowed length
' according to RFC 3164.
' It is assumed that the passed parameters have been validated.

  Dim intLenPri             As Integer
  Dim intLenHeader          As Integer
  Dim intLenMsg             As Integer
  Dim intLenCutOff          As Integer
  
  ' Find the length of the three parts of the package.
  intLenPri = Len(StrSyslogPriority(bytFacility, bytSeverity))
  intLenHeader = cintRfc3164LenTime + _
    Len(cstrPartsSeparator & strHostname & cstrPartsSeparator)
  intLenMsg = Len(StrSyslogMessage(strTag, lngPid, strContent))
  
  ' Limit the package by cutting off CONTENT as needed.
  intLenCutOff = (intLenPri + intLenHeader + intLenMsg) - cintRfc3164LenPack
  If intLenCutOff > 0 Then
    ' Cut off CONTENT to fit within package.
    strContent = Left(strContent, Len(strContent) - intLenCutOff)
  End If

End Sub

'Public Function SyslogEntrySave( _
'  Optional ByVal strRemoteAddress As String, _
'  Optional ByVal bytFacility As slFacility = slFacilityDefault, _
'  Optional ByVal bytSeverity As slSeverity = slSeverityDefault, _
'  Optional ByVal datTimestamp As Date, _
'  Optional ByVal strHostname As String, _
'  Optional ByVal strTag As String, _
'  Optional ByVal lngPid As Long, _
'  Optional ByVal strContent As String) _
'  As String
'
'' Save a decoded syslog package in tblSyslog.
'
'  Const cstrTable As String = "tblSyslog"
'
'  Dim cnn           As ADODB.Connection
'  Dim rst           As ADODB.Recordset
'
'  ' Array to hold fractions of CONTENT.
'  Dim astrContent() As String
'
'  Dim bytFraction   As Byte
'  Dim booSuccess    As Boolean
'
'  On Error Resume Next
'
'  Set cnn = CurrentProject.Connection
'  Set rst = New ADODB.Recordset
'
'  ' Apply default values for those missing parameters
'  ' that don't have default values defined in syslog table.
'  If CDbl(datTimestamp) = 0 Then
'    datTimestamp = Now
'  End If
'  If Len(strHostname) = 0 Then
'    If Len(strRemoteAddress) = 0 Then
'      ' Locally generated entry.
'      strHostname = cstrHostNameLocal
'    Else
'      strHostname = strRemoteAddress
'    End If
'  End If
'  If Len(strContent) = 0 Then
'    strContent = cstrContentEmpty
'  End If
'  ' Split CONTENT into four fractions.
'  astrContent = SplitContent(strContent)
'
'  With rst
'    .Open cstrTable, cnn, adOpenKeyset, adLockOptimistic, adCmdTable
'    .AddNew
'      .Fields("Facility").Value = bytFacility
'      .Fields("Severity").Value = bytSeverity
'      .Fields("TimeStamp").Value = datTimestamp
'      If Len(strRemoteAddress) > 0 Then
'        ' Only store remote address if it is present.
'        .Fields("RemoteAddress").Value = strRemoteAddress
'      End If
'      .Fields("Hostname").Value = strHostname
'      If Len(strTag) > 0 Then
'        ' Only store TAG if it is not a zero length string.
'        .Fields("Tag").Value = strTag
'      End If
'      .Fields("ProcessId").Value = lngPid
'      For bytFraction = LBound(astrContent) To UBound(astrContent)
'        .Fields("Content" & CStr(bytFraction)).Value = astrContent(bytFraction)
'      Next
'    .Update
'    booSuccess = True
'    .Close
'  End With
'
'  Set rst = Nothing
'  Set cnn = Nothing
'
'  SyslogEntrySave = booSuccess
'
'End Function

Public Function SplitContent( _
  ByVal strContent As String) _
  As String()
  
' Splits strContent into four fractions with a maximum length
' of 255 each and returns the fractions in an array.

  Const cintFractionLen   As Integer = 255
  
  Dim astrContent(0 To 3) As String
  
  Dim intPos              As Integer
  Dim bytFraction         As Byte

  ' Build array with CONTENT.
  For bytFraction = LBound(astrContent) To UBound(astrContent)
    intPos = bytFraction * cintFractionLen
    astrContent(bytFraction) = Mid(strContent, 1 + intPos, cintFractionLen)
  Next
  
  SplitContent = astrContent
  
End Function

Public Function ConcatenateContent( _
  ByVal strContent0 As String, _
  Optional ByVal strContent1 As Variant, _
  Optional ByVal strContent2 As Variant, _
  Optional ByVal strContent3 As Variant) _
  As String

' Concatenates CONTENT from one to four fractions.

  Dim strContent          As String
  
  strContent = strContent0 & strContent1 & strContent2 & strContent3
  
  ConcatenateContent = strContent
  
End Function





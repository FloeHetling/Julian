Attribute VB_Name = "modWinAPI"
Option Explicit

Private Const MAX_WSADescription  As Long = 256
Private Const MAX_WSASYSStatus    As Long = 128
Private Const WS_VERSION_REQD     As Long = &H101
Private Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD    As Long = 1
Private Const SOCKET_ERROR        As Long = -1
Private Const ERROR_NONE          As Long = 0
Private Const IP_SUCCESS          As Long = 0

Private Type WSADATA
  wVersion                                As Integer
  wHighVersion                            As Integer
  szDescription(0 To MAX_WSADescription)  As Byte
  szSystemStatus(0 To MAX_WSASYSStatus)   As Byte
  wMaxSockets                             As Long
  wMaxUDPDG                               As Long
  dwVendorInfo                            As Long
End Type

Private Declare Function GetHostName Lib "wsock32.dll" Alias "gethostname" ( _
  ByVal szHost As String, _
  ByVal dwHostLen As Long) _
  As Long

Private Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" ( _
  ByVal Hostname As String) _
  As Long
 
Private Declare Function WSAStartup Lib "wsock32.dll" ( _
  ByVal wVersionRequired As Long, _
  ByRef lpWSADATA As WSADATA) _
  As Long
   
Private Declare Function WSACleanup Lib "wsock32.dll" () _
  As Long

Private Declare Function inet_ntoa Lib "wsock32.dll" ( _
  ByVal addr As Long) _
  As Long
                       
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
  ByRef xDest As Any, _
  ByRef xSource As Any, _
  ByVal nbytes As Long)

Private Declare Function lstrcpyA Lib "kernel32" ( _
  ByVal RetVal As String, _
  ByVal Ptr As Long) _
  As Long

Private Declare Function lstrlenA Lib "kernel32" ( _
  ByRef lpString As Any) _
  As Long

Public Function MachineHostName() _
  As String
  
' Retrieves the host name of this machine.
'
' The host name is preserved in a static variable to
' reduce look up time for repeated calls.

  Static strHostname    As String
  
  If Len(strHostname) = 0 Then
    ' Host name has not been looked up previously.
    If WinSocketsStart() = True Then
      ' Obtain and store the host name.
      strHostname = GetMachineName()
      
      Call WinSocketsClean
    End If
  End If
  
  MachineHostName = strHostname

End Function

Public Function MachineHostAddress( _
  Optional ByVal strHostname As String) _
  As String
  
' Retrieves IP address of the machine with the host name
' strHostname.
' If a zero length host name or no host name is passed, the
' address of this machine is returned.
' If host name localhost is passed, 127.0.0.1 is returned.
' If the host name cannot be resolved, 0.0.0.0 is returned.
'
' The host addresses are preserved in a static collection to
' reduce look up time for repeated calls.

  ' If strHostname is an empty string, the local host address
  ' will be looked up.
  ' However, an empty string cannot be a key in a collection.
  ' Use this key to store the local host address.
  Const cstrKeyThisHost As String = " "
  
  Static colAddress     As New Collection
  
  Dim strIpAddress      As String
  
  ' Ignore error when looking up a key in collection
  ' colAddress that does not exist.
  On Error Resume Next
  
  If Len(strHostname) = 0 Then
    strHostname = cstrKeyThisHost
  End If
  strIpAddress = colAddress.Item(strHostname)
  ' If strHostname is not found, an error is raised.
  If Err.Number <> 0 Then
    ' This host name has not been looked up previously.
    If WinSocketsStart() = True Then
      ' Obtain the host address.
      ' Trim strHostname to pass a zero length string when
      ' looking up the address of the local host.
      strIpAddress = GetIPFromHostName(Trim(strHostname))
      ' Store the host address.
      colAddress.Add strIpAddress, strHostname
      
      Call WinSocketsClean
    End If
  End If
  
  MachineHostAddress = strIpAddress
  
End Function

Public Sub ShowHostNameAddress()

' Displays host name and IP address of local machine.

  Const cstrMsgTitle  As String = "Host name and IP address"
  Const clngMsgStyle0 As Long = vbExclamation + vbOKOnly
  Const clngMsgStyle1 As Long = vbInformation + vbOKOnly
  Const cstrMsgPrompt As String = "No access to address information."
  
  Dim strHostname   As String
  Dim strIpAddress  As String
  Dim strMsgPrompt  As String
  
  If WinSocketsStart() = True Then
    ' Obtain and pass the host address.
    strHostname = GetMachineName()
    strIpAddress = GetIPFromHostName(strHostname)
    ' Display name and address.
    strMsgPrompt = _
      "Host name: " & strHostname & vbCrLf & _
      "IP address: " & strIpAddress
    WriteToLog strMsgPrompt & clngMsgStyle1 & cstrMsgTitle
    
    Call WinSocketsClean
  Else
    MsgBox cstrMsgPrompt, clngMsgStyle0, cstrMsgTitle
  End If
  
End Sub

Private Function WinSocketsStart() _
  As Boolean

' Start up Windows sockets before use.

  Const cstrMsgTitle  As String = "Windows Sockets"
  Const clngMsgStyle  As Long = vbCritical + vbOKOnly
  Const cstrMsgPrompt As String = "Error at start up of Windows sockets."
  
  Dim typWSA      As WSADATA
  Dim booSuccess  As Boolean
  
  If WSAStartup(WS_VERSION_REQD, typWSA) = IP_SUCCESS Then
    booSuccess = True
  End If
  
  If booSuccess = False Then
    MsgBox cstrMsgPrompt, clngMsgStyle, cstrMsgTitle
  End If
  
  WinSocketsStart = booSuccess
   
End Function

Private Function WinSocketsClean() _
  As Boolean

' Clean up Windows sockets after use.

  Const cstrMsgTitle  As String = "Windows Sockets"
  Const clngMsgStyle  As Long = vbExclamation + vbOKOnly
  Const cstrMsgPrompt As String = "Error at clean up of Windows sockets."
  
  Dim booSuccess  As Boolean

  If WSACleanup() = ERROR_NONE Then
    booSuccess = True
  End If
   
  If booSuccess = False Then
    MsgBox cstrMsgPrompt, clngMsgStyle, cstrMsgTitle
  End If

  WinSocketsClean = booSuccess

End Function
  
Private Function GetMachineName() As String

' Retrieves the host name of this machine.

  ' Assign buffer for maximum length of host name plus
  ' a terminating null char.
  Const clngBufferLen As Long = 255 + 1
  
  Dim stzHostName As String * clngBufferLen
  Dim strHostname As String
  
  If GetHostName(stzHostName, clngBufferLen) = ERROR_NONE Then
    ' Trim host name from buffer string.
    strHostname = Left(stzHostName, InStr(1, stzHostName, vbNullChar, vbBinaryCompare) - 1)
  End If
  
  GetMachineName = strHostname
  
End Function

Private Function GetIPFromHostName( _
  ByVal strHostname As String) _
  As String

' Converts a host name to its IP address.
'
' If strHostname
'   - is zero length, local IP address is returned.
'   - is "localhost", IP address 127.0.0.1 is returned.
'   - cannot be resolved, unknown IP address 0.0.0.0 is returned.

  Const clngAddressNone   As Long = 0
  ' The Address is offset 12 bytes from the
  ' start of the HOSENT structure.
  Const clngAddressOffset As Long = 12
  ' Size of address part.
  Const clngAddressChunk  As Long = 4
  ' Address to return if none found.
  Const cstrAddressZero   As String = "0.0.0.0"

  ' Address of HOSENT structure.
  Dim ptrHosent           As Long
  ' Address of name pointer.
  Dim ptrName             As Long
  ' Address of address pointer.
  Dim ptrAddress          As Long
  Dim ptrIPAddress        As Long
  Dim ptrIPAddress2       As Long
  Dim stzHostName         As String
  Dim strAddress          As String

  stzHostName = strHostname & vbNullChar
  ptrHosent = GetHostByName(stzHostName)

  If ptrHosent = clngAddressNone Then
    ' Return address zero.
    strAddress = cstrAddressZero
  Else
    ' Assign pointer addresses and offset Null-terminated list
    ' of addresses for the host.
    ' Note:
    ' We are retrieving only the first address returned.
    ' To return more than one, define strAddress as a string array
    ' and loop through the 4-byte ptrIPAddress members returned.
    ' The last item is a terminating null.
    ' All addresses are returned in network byte order.
    ptrAddress = ptrHosent + clngAddressOffset
    
    ' Get the IP address.
    CopyMemory ptrAddress, ByVal ptrAddress, clngAddressChunk
    CopyMemory ptrIPAddress, ByVal ptrAddress, clngAddressChunk
    CopyMemory ptrIPAddress2, ByVal ptrIPAddress, clngAddressChunk
    
    strAddress = GetInetStrFromPtr(ptrIPAddress2)
  End If
  
  GetIPFromHostName = strAddress
  
End Function

Private Function GetInetStrFromPtr( _
  ByVal lngAddress As Long) _
  As String
 
' Converts decimal IP address to IP address string.

  GetInetStrFromPtr = GetStrFromPtrA(inet_ntoa(lngAddress))

End Function

Private Function GetStrFromPtrA( _
  ByVal lpszA As Long) _
  As String
  
' Copies string from pointer.

  ' Create buffer string.
  GetStrFromPtrA = String(lstrlenA(ByVal lpszA), vbNullChar)
  ' Copy value from pointer to buffer string.
  Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
  
End Function





Attribute VB_Name = "modWSA"
Attribute VB_Description = "ip functions"
'/**  modWSA
' *     ip functions
' *
' *     @project     xvnc
' *     @author      Philipp Rothmann,  mailto: dev@preneco.de
' *     @copyright   © 2000-2003 .) development
' *
' *     This program is free software; you can redistribute it and/or modify
' *     it under the terms of the GNU General Public License as published by
' *     the Free Software Foundation; either version 2 of the License, or
' *     (at your option) any later version.
' *
' *     This program is distributed in the hope that it will be useful,
' *     but WITHOUT ANY WARRANTY; without even the implied warranty of
' *     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' *     GNU General Public License for more details.
' *
' *     You should have received a copy of the GNU General Public License
' *     along with this program; if not, write to the Free Software
' *     Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
' *
' *
' *     @lastupdate  17.03.2003 20:03:39
' */

Option Explicit

Private Const WS_VERSION_REQD As Long = &H101
Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const SOCKET_ERROR As Long = -1
Private Const WSADescription_Len As Long = 256
Private Const WSASYS_Status_Len As Long = 128

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type
 
Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Private Type QOCINFO
  dwSize As Long
  dwFlags As Long
  dwInSpeed As Long 'in bytes/second
  dwOutSpeed As Long 'in bytes/second
End Type

Private Declare Function IsDestinationReachable Lib "SENSAPI.DLL" Alias "IsDestinationReachableA" (ByVal lpszDestination As String, ByRef lpQOCInfo As QOCINFO) As Long
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
 
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

'/**  GetIPAddresses
' *     gets the ip-address of a host
' *
' *     @param       ByVal sMachineName As String , the machine to query the ip-address from
' *     @returns     Collection , all the ips
' */
Public Function GetIPAddresses(ByVal sMachineName As String) As Collection
Attribute GetIPAddresses.VB_Description = "gets the ip-address of a host"
   
  Dim WSAD As WSADATA
  Dim iReturn As Integer
  Dim hostent_addr As Long
  Dim host As HOSTENT
  Dim hostip_addr As Long
  Dim temp_ip_address() As Byte
  Dim i As Long, j As Long
  Dim ip_address As String

    Set GetIPAddresses = New Collection
  
    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
 
    If iReturn = 0 Then
        hostent_addr = gethostbyname(sMachineName)
    
        If hostent_addr <> 0 Then
            RtlMoveMemory host, hostent_addr, LenB(host)
            RtlMoveMemory hostip_addr, host.hAddrList, 4
    
            'get all of the IP address if machine is  multi-homed
            j = 1
            Do
                ReDim temp_ip_address(1 To host.hLength)
                RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

                For i = 1 To host.hLength
                    ip_address = ip_address & temp_ip_address(i) & "." ' LoadResString(10093)
                Next i
                ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
              
                j = j + 1
                host.hAddrList = host.hAddrList + LenB(host.hAddrList)
                RtlMoveMemory hostip_addr, host.hAddrList, 4
                
                GetIPAddresses.Add ip_address
                
                ip_address = ""
            Loop While (hostip_addr <> 0)
        End If
    End If
    WSACleanup

End Function

Public Function Ping(Address As String) As Boolean
'Ñïõôéíá ðïõ êáíåé ping ìéá IP Þ åíá dns ïíïìá Ç/Õ óôï äýêôéï êáé åðéóôñåöåé true/false
  Dim Ret   As QOCINFO
  Dim IP    As String
  
  Ret.dwSize = Len(Ret)
  IP = Address
  If IsDestinationReachable(IP, Ret) = 0 Then
    Ping = False
   Else
  '  MsgBox "The destination can be reached!" + vbCrLf + _
     "The speed of data coming in from the destination is " + Format$(Ret.dwInSpeed / 1048576, "#.0") + " Mb/s," + vbCrLf + _
     "and the speed of data sent to the destination is " + Format$(Ret.dwOutSpeed / 1048576, "#.0") + " Mb/s."
  '   MsgBox IP & " - " & IsDestinationReachable(IP, Ret)
    Ping = True
  End If
End Function

Public Function TrimNull(ByVal Text As String) As String
'ÍÁ ÌÅÔÁÊÏÌÉÓÅÉ ÓÅ ÁËËÏ MODULE ÃÉÁ ËÏÃÏÕÓ ÔÁÎÇÓ
 TrimNull = Replace$(Text, vbNullString, "")
 TrimNull = Replace$(TrimNull, vbCr, "")
 TrimNull = Replace$(TrimNull, vbLf, "")
 TrimNull = Trim$(TrimNull)
End Function


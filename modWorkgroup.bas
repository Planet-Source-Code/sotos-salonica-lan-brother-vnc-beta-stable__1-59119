Attribute VB_Name = "modWorkgroup"
Attribute VB_Description = "gets workgroup information of the local computer"
'/**  modWorkgroup
' *     gets workgroup information of the local computer
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
      
Private Const SV_TYPE_WORKSTATION  As Long = &H1&
Private Const SV_TYPE_SERVER  As Long = &H2&
Private Const SV_TYPE_SQLSERVER  As Long = &H4&
Private Const SV_TYPE_DOMAIN_CTRL  As Long = &H8&
Private Const SV_TYPE_DOMAIN_BAKCTRL  As Long = &H10&
Private Const SV_TYPE_TIME_SOURCE  As Long = &H20&
Private Const SV_TYPE_AFP  As Long = &H40&
Private Const SV_TYPE_NOVELL  As Long = &H80&
Private Const SV_TYPE_DOMAIN_MEMBER  As Long = &H100&
Private Const SV_TYPE_PRINTQ_SERVER  As Long = &H200&
Private Const SV_TYPE_DIALIN_SERVER  As Long = &H400&
Private Const SV_TYPE_XENIX_SERVER  As Long = &H800&
Private Const SV_TYPE_SERVER_UNIX  As Long = &H800&
Private Const SV_TYPE_NT  As Long = &H1000&
Private Const SV_TYPE_WFW  As Long = &H2000&
Private Const SV_TYPE_SERVER_MFPN  As Long = &H4000&
Private Const SV_TYPE_SERVER_NT  As Long = &H8000&
Private Const SV_TYPE_POTENTIAL_BROWSER  As Long = &H10000
Private Const SV_TYPE_BACKUP_BROWSER  As Long = &H20000
Private Const SV_TYPE_MASTER_BROWSER  As Long = &H40000
Private Const SV_TYPE_DOMAIN_MASTER  As Long = &H80000
Private Const SV_TYPE_SERVER_OSF  As Long = &H100000
Private Const SV_TYPE_SERVER_VMS  As Long = &H200000
Private Const SV_TYPE_WINDOWS  As Long = &H400000
Private Const SV_TYPE_ALTERNATE_XPORT  As Long = &H20000000
Private Const SV_TYPE_LOCAL_LIST_ONLY  As Long = &H40000000
Private Const SV_TYPE_DOMAIN_ENUM  As Long = &H80000000
Private Const SV_TYPE_ALL  As Long = &HFFFF

Public Enum TypeOfMachine
 AllOS = SV_TYPE_ALL
 Windows9x = SV_TYPE_WINDOWS
 NTBased = SV_TYPE_NT
 DomainMaster = SV_TYPE_DOMAIN_MASTER
End Enum

Private Type SERVER_INFO_101
    wki101_platform_id As Long
    wki101_servername As Long
    wki101_langroup As Long
    wki101_ver_major As Long
    wki101_ver_minor As Long
    wki101_lanroot As Long
End Type
      
Private Declare Function NetServerEnum Lib "Netapi32" _
                          (ByVal sServerName As String, _
                          ByVal lLevel As Long, _
                          vBuffer As Long, _
                          lPreferedMaxLen As Long, _
                          lEntriesRead As Long, _
                          lTotalEntries As Long, _
                          ByVal vServerType As Long, _
                          ByVal sDomain As String, _
                          vResume As Long) As Long
Private Declare Function NetApiBufferFree Lib "Netapi32" _
                          (ByVal lBuffer&) As Long
Private Declare Sub lstrcpyW Lib "kernel32" _
                          (vDest As Any, ByVal vSrc As Any)
Private Declare Sub RtlMoveMemory Lib "kernel32" _
                          (dest As Any, vSrc As Any, ByVal lSize&)

'/**  GetServers
' *     gets all ntservers/workstations on your local network
' *
' *     @returns     Collection , the collection that colects all local area network computers
' */
Public Function GetServers(Optional LLL As TypeOfMachine) As Collection
Attribute GetServers.VB_Description = "gets all ntservers/workstations on your local network"

  Dim pstrServerName         As String
  Dim pstrComputerName       As String
  Dim pstrDomainName         As String
  Dim pbytBuf(512)           As Byte
  Dim plngRtn                As Long
  Dim plngSrvInfoRtn         As Long
  Dim plngPreferedMaxLen     As Long
  Dim plngNumEntriesRead     As Long
  Dim plngTotalEntries       As Long
  Dim plngServerType         As Long
  Dim plngResume             As Long
  Dim plngCnt                As Long
  Dim lSeverInfo101StructPtr As Long
  Dim plngServEnumLevel      As Long
  Dim plngPos                As Long
  Dim ptypServerInfo101      As SERVER_INFO_101
  Dim plngLenComputerName    As Long
   
    Set GetServers = New Collection
   
    pstrComputerName = vbNullString
    pstrServerName = vbNullString
    pstrDomainName = vbNullString
    plngServEnumLevel = 101
    plngServerType = SV_TYPE_ALL
    
    If Len(LLL) <> 0 Then
      plngServerType = LLL
    End If
    
    
   
    plngRtn = NetServerEnum(pstrServerName, _
              plngServEnumLevel, plngSrvInfoRtn, _
              plngPreferedMaxLen, plngNumEntriesRead, _
              plngTotalEntries, plngServerType, _
              pstrDomainName, plngResume)
   
    '' NetServerEnum Index is 1 based -- iterate
    ''     thru the return values to get the
    ''     NT servers and workstations
    plngCnt = 1
    lSeverInfo101StructPtr = plngSrvInfoRtn
   
    For plngCnt = 1 To plngTotalEntries
      
        pstrComputerName = ""
           
        RtlMoveMemory ptypServerInfo101, ByVal lSeverInfo101StructPtr, Len(ptypServerInfo101)
        
        lstrcpyW pbytBuf(0), ptypServerInfo101.wki101_servername
        'lstrcpyW pbytBuf(0), ptypServerInfo101.wki101_platform_id
         
        pstrComputerName = pbytBuf
        plngPos = InStr(pstrComputerName, vbNullChar)
        If plngPos > 0 Then
            pstrComputerName = Mid$(pstrComputerName, 1, plngPos - 1)
        End If
      
        GetServers.Add pstrComputerName
              
        lSeverInfo101StructPtr = lSeverInfo101StructPtr + Len(ptypServerInfo101)
   
    Next plngCnt
   
    plngRtn = NetApiBufferFree(plngSrvInfoRtn)
 
End Function

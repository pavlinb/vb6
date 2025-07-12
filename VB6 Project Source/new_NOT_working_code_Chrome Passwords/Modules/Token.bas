Attribute VB_Name = "Token"
'------------------------------------
'          TOKEN PRIVILEGE           |
'------------------------------------

' Global constants we must use with security descriptor
Private Const SECURITY_DESCRIPTOR_REVISION = 1
Private Const OWNER_SECURITY_INFORMATION = 1&

' Access Token constants
Private Const TOKEN_ASSIGN_PRIMARY = &H1
Private Const TOKEN_DUPLICATE = &H2
Private Const TOKEN_IMPERSONATE = &H4
Private Const TOKEN_QUERY = &H8
Private Const TOKEN_QUERY_SOURCE = &H10
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_ADJUST_GROUPS = &H40
Private Const TOKEN_ADJUST_DEFAULT = &H80
Private Const TOKEN_ALL_ACCESS = 983551
Private Const ANYSIZE_ARRAY = 1

' Token Privileges constants
Private Const SE_RESTORE_NAME = "SeRestorePrivilege"
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"
Private Const SE_PRIVILEGE_ENABLED = 2&

' ACL structure
Private Type ACL
   AclRevision As Byte
   Sbz1 As Byte
   AclSize As Integer
   AceCount As Integer
   Sbz2 As Integer
End Type

Private Type SECURITY_DESCRIPTOR
   Revision As Byte
   Sbz1 As Byte
   Control As Long
   Owner As Long
   Group As Long
   Sacl As ACL
   Dacl As ACL
End Type

' Token structures
Private Type LARGE_INTEGER
   LowPart As Long
   HighPart As Long
End Type

Private Type LUID
   LowPart As Long
   HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

' Win32 API calls
Private Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, Sid As Byte, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Private Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Long) As Long
Private Declare Function SetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal bOwnerDefaulted As Long) As Long
Private Declare Function SetFileSecurity Lib "advapi32.dll" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As Long, ByVal ReturnLength As Long) As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Enable_Privilege(Privilege As String) As Boolean
    Enable_Privilege = ModifyState(Privilege, True)
End Function

Public Function Disable_Privilege(Privilege As String) As Boolean
    Disable_Privilege = ModifyState(Privilege, False)
End Function

'-------------
' ModifyState |
'-------------
Public Function ModifyState(Privilege As String, Enable As Boolean) As Boolean
    
Dim MyPrives As TOKEN_PRIVILEGES
Dim PrivilegeId As LUID
Dim ptrPriv As Long    ' Pointer to Privileges Structure
Dim hToken As Long     ' Token Handle
Dim Result As Long     ' Return Value
  
Result = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES, hToken)
    If (Result = 0) Then
        ModifyState = False
        'MsgBox "OpenProcessToken failed with error code " & Err.LastDllError
        Exit Function
    End If
   
    Result = LookupPrivilegeValue(vbNullString, Privilege, PrivilegeId)
    If (Result = 0) Then
        ModifyState = False
        'MsgBox "LookupPrivilegeValue failed with error code " & Err.LastDllError
        Exit Function
    End If
   
MyPrives.Privileges(0).pLuid = PrivilegeId
MyPrives.PrivilegeCount = 1

    If (Enable) Then
        MyPrives.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    Else
        MyPrives.Privileges(0).Attributes = 0
    End If
    
    Result = AdjustTokenPrivileges(hToken, False, MyPrives, 0, 0, 0)
    If (Result = 0 Or Err.LastDllError <> 0) Then
        ModifyState = False
        'MsgBox "AdjustTokenPrivileges failed with error code " & Err.LastDllError
        Exit Function
        Else: MsgBox "AdjustTokenPrivileges SUCCESS!"
    End If
   
CloseHandle hToken
   
ModifyState = True

End Function


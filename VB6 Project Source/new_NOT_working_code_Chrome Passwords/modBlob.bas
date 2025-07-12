Attribute VB_Name = "modBlob"
Public Type DATA_BLOB
    cbData              As Long
    pbData              As Long
End Type

Public Declare Function CryptUnprotectData Lib "crypt32.dll" (ByRef pDataIn As DATA_BLOB, ByVal ppszDataDescr As Long, ByVal pOptionalEntropy As Long, ByVal pvReserved As Long, ByVal pPromptStruct As Long, ByVal dwFlags As Long, ByRef pDataOut As DATA_BLOB) As Long

Public Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function LocalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function CredEnumerate Lib "advapi32.dll" Alias "CredEnumerateW" (ByVal lpszFilter As Long, ByVal lFlags As Long, ByRef pCount As Long, ByRef lppCredentials As Long) As Long
Public Declare Function CredFree Lib "advapi32.dll" (ByVal pBuffer As Long) As Long

'Special Folder
Public Const CSIDL_LOCAL_APPDATA = &H1C&
Public Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

Public Function SpecialFolder(pfe As Long) As String
    Const MAX_PATH = 260
    Dim strPath As String
    Dim strBuffer As String
    
    strBuffer = Space$(MAX_PATH)
    If SHGetFolderPath(0, pfe, 0, 0, strBuffer) = 0 Then strPath = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    If Right$(strPath, 1) = "\" Then strPath = Left$(strPath, Len(strPath) - 1)
    SpecialFolder = strPath
End Function


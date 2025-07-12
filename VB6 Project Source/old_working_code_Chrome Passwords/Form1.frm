VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Chrome Pass < ver.135"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEncrypted_key 
      Height          =   1125
      Left            =   204
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   5220
      Width           =   6624
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Chrome Passwords"
      Height          =   495
      Left            =   4644
      TabIndex        =   1
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "encrypted_key:"
      Height          =   276
      Left            =   192
      TabIndex        =   3
      Top             =   4848
      Width           =   6096
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' references:
'https://www.vbforums.com/showthread.php?865951-RESOLVED-Vb6-aes-gcm
'https://www.vbforums.com/showthread.php?891596-RESOLVED-clsCrypt-cls

'SQL Database
Private m_DB_Chrome     As Long
Private m_Index_Chrome  As String

'Unicode
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8      As Long = 65001



Private Sub Form_Load()

    txtEncrypted_key.Text = "RFBBUEkBAAAA0Iyd3wEV0RGGegDAT8KX6wEAAABoYWFRhRBLR57kk3rHzKk7EAAAABwAAABHAG8AbwBnAGwAZQAgAEMAaAByAG8AbQBlAAAAEGYAAAABAAAgAAAAnasmPSU9qStnS3H0l4z3TqMbCCsNZcIKEogXQeBEBDMAAAAADoAAAAACAAAgAAAALVujlluCIasZVEWHL4yLvI5BtDWqsJgYDG7TW2TmL5owAAAA9ix+gxfr1pMUHbfHsWgSzV3qwvnqtoyv0clWxhVsjzbzVEn4rqOAykrprKHbJ4nuQAAAAOWle0zNo1QXMlCn6cAaLqDVt7xcRcry4FaMex6JCOHTSC7kZNIh1HH6NGuXbLnuu9EZDSYIfvqmFmKsdpqyQaI="
    
    'Get encrypted key from \AppData\Local\Google\Chrome\User Data\Local State"
    'txtEncrypted_key.text = "os_crypt":{"encrypted_key":"RFBBUEkBAAAA0Iyd3wEV0RGGegDAT8KX6wEAAAAuHNUgeSWaTbo5CTS+ULBVEAAAABwAAABCAHIAYQB2AGUAIABCAHIAbwB3AHMAZQByAAAAEGYAAAABAAAgAAAACR07RKtYKOepBCKB/+3Snguiaf4Vq4KwBMc3ICM2GJ8AAAAADoAAAAACAAAgAAAADWZF4NTV79cwida3PoV7ZO3EA7UMswo6wXnMZ7NIgPQwAAAAe7d7syICcgjUFaAFFCiPfiPsRVNKYo1+vm/sGGzprdjvJ5KPSkv5kME7YBwgDmR8QAAAAJKdvej6Vw1+rL/lfP8lUQP3AG27t916Kef5Su0HnM+cJbKpStrZZYljThzcYhAWqC2DOJt2jzo6HZXxfwjLwNU="}

End Sub

Private Sub Command1_Click()

'Show Passwords
    Call ShowPasswords

End Sub


Private Sub ShowPasswords()

Text1.Text = ""

    'Decrypt Chrome
    Call DecryptPasswordsChrome
        
    'Close database after decryption
    ite_close m_DB_Chrome
    

End Sub

'-------------------------
' Decrypt Chrome Passwords
'-------------------------
Private Sub DecryptPasswordsChrome()

'Chrome Database Path
    Dim dPath As String 'Database Path
    dPath = SpecialFolder(CSIDL_LOCAL_APPDATA) & "\Google\Chrome\User Data\Default\Login Data"

'Exit if no database found
    If FileExists(dPath) = False Then Exit Sub

'Decode Chrome Passwords
    Dim tBlobOut    As DATA_BLOB
    Dim tBlobIn     As DATA_BLOB
    Dim aData() As Byte

    Dim encrypted_key As String
    encrypted_key = txtEncrypted_key.Text
    
    If Len(encrypted_key) Then
        'Dim sdecrypted_key As String
        'sdecrypted_key = Base64Decode(encrypted_key)
        
        Dim decrypted_key() As Byte
        'decrypted_key = DecodeBase64(encrypted_key)
        decrypted_key = Base64Decodeb(encrypted_key)
    
        Dim RetVal As Long
        
        With tBlobIn
            .cbData = UBound(decrypted_key) - LBound(decrypted_key) + 1 - 5
            .pbData = VarPtr(decrypted_key(LBound(decrypted_key) + 5))
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
    Else
        RetVal = 1
    End If
    
    'MsgBox RetVal
    If RetVal Then
        
        'Open database
        'If Chrome is running database might be locked -> CLOSE Chrome on testing
        ite_open dPath, m_DB_Chrome
        
        Dim sqlite3_stmt  As Long
        Dim hRet As Long
        
        hRet = ite_prepare(m_DB_Chrome, "SELECT * FROM logins", sqlite3_stmt)
        
        If hRet = SQLITE_OK Then
            
            Do While ite_next(sqlite3_stmt)

' QUICK FIX !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!! ERROR BECAUSE OF EMPTY BLOB VALUE (BLOB VALUE IS NOT BLOB)!!!!!
If Len(ite_column_text(sqlite3_stmt, 5)) > 0 Then 'SKIP EMPTY BLOB ERROR!!!!!!!!!!!!
                aData() = ite_column_blob(sqlite3_stmt, 5)
                               
                ''VERSION < 79
                '
                'If Len(encrypted_key) = 0 Then
                '
                '    ' Fill the blob structure
                '    With tBlobIn
                '       .cbData = UBound(aData) - LBound(aData) + 1
                '       .pbData = VarPtr(aData(0))
                '    End With
                '
                '    If CryptUnprotectData(tBlobIn, 0&, 0&, 0&, 0&, 8, tBlobOut) Then
                '        'first database column(0) = URL
                '        'fourth column(3) = Username
                '        'sixth column(5) = Password = aData()
                '
                '        Text1.Text = Text1.Text & _
                '        "URL: " & ite_column_text(sqlite3_stmt, 0) & vbCrLf & _
                '        "Username: " & ite_column_text(sqlite3_stmt, 3) & vbCrLf & _
                '        "Password: " & ReadBlobString(tBlobOut) & vbCrLf & vbCrLf
                '    End If
                '
                'Else 'VERSION 80 and above
                '
                '    '' Fill the blob structure
                '    'With tBlobIn
                '    '   .cbData = UBound(aData) - LBound(aData) + 1 - 5 'len("DPAPI")
                '    '   .pbData = VarPtr(aData(5))
                '    'End With
                '
                '        'first database column(0) = URL
                '        'fourth column(3) = Username
                '        'sixth column(5) = Password = aData()
                '
                '        'Dim pwd As String
                '
                '        'pwd = decrypt_password(aData, amaster_key)   '<---- Previous decryption function
                '        'pwd = NewDecodePassword(aData, amaster_key)  '<---- NEW Function
                '
                        Text1.Text = Text1.Text & _
                        "URL: " & ite_column_text(sqlite3_stmt, 0) & vbCrLf & _
                        "Username: " & ite_column_text(sqlite3_stmt, 3) & vbCrLf & _
                        "Password: " & NewDecodePassword(aData, amaster_key) & vbCrLf & vbCrLf

                        'Convert UTF password to Unicode with UTF8ToUnicode()
                        MsgBox "Username: " & ite_column_text(sqlite3_stmt, 3) & vbCrLf & "Password: " & UTF8ToUnicode(NewDecodePassword(aData, amaster_key)), , "Username and Password:"
                    
                'End If
                
                
            
End If
                
                
                
            Loop
                     
            
            
            
        End If
    
    End If
        
'finalize database
ite_finalize sqlite3_stmt

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
    Dim b       As Byte
    Dim i       As Long
    
    If tBlob.cbData = 0 Then Exit Function
    If tBlob.pbData = 0 Then Exit Function

    For i = 0 To tBlob.cbData - 1
        CopyMemory b, ByVal tBlob.pbData + i, 1
        ReadBlobString = ReadBlobString & Chr$(b)
    Next
    
    Call LocalFree(tBlob.pbData)
    
End Function


Private Function ReadBlobArray(ByRef tBlob As DATA_BLOB, Optional free As Boolean = True) As Byte()
    Static b()    As Byte
    Dim i       As Long
    
    If tBlob.cbData = 0 Then Exit Function
    If tBlob.pbData = 0 Then Exit Function

    ReDim b(0 To tBlob.cbData) '1 more for terminating 0
    CopyMemory b(0), ByVal tBlob.pbData, tBlob.cbData
    
    If free Then
        Call LocalFree(tBlob.pbData)
    End If
    
    ReadBlobArray = b
    
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

Private Function UnicodeToUTF8(ByVal sData As String) As String
    Dim bvData()    As Byte
    Dim lSize       As Long
    Dim lRet        As Long
    
    If LenB(sData) Then
        lSize = Len(sData) * 2
        ReDim bvData(lSize)
    
        lRet = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sData), _
           Len(sData), bvData(0), lSize + 1, vbNullString, 0)
    
        If lRet Then
            ReDim Preserve bvData(lRet - 1)
            UnicodeToUTF8 = StrConv(bvData, vbUnicode)
        End If
    End If
End Function




Attribute VB_Name = "NewAES"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (ByRef hAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
Private Declare Function BCryptSetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGenerateSymmetricKey Lib "bcrypt" (ByVal hAlgorithm As Long, hKey As Long, ByVal pbKeyObject As Long, ByVal cbKeyObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDecrypt Lib "bcrypt" (ByVal hKey As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal pPaddingInfo As Long, ByVal pbIV As Long, ByVal cbIV As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, cbResult As Long, ByVal dwFlags As Long) As Long

Private Type BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
    cbSize              As Long
    dwInfoVersion       As Long
    pbNonce             As Long
    cbNonce             As Long
    pbAuthData          As Long
    cbAuthData          As Long
    pbTag               As Long
    cbTag               As Long
    pbMacContext        As Long
    cbMacContext        As Long
    cbAAD               As Long
    lPad                As Long
    cbData(7)           As Byte
    dwFlags             As Long
    lPad2               As Long
End Type


Private Function pvCryptoAeadAesGcmDecrypt( _
            baKey() As Byte, baIV() As Byte, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long, _
            baTag() As Byte, ByVal lTagPos As Long, ByVal lTagSize As Long, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long) As Boolean
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim hAlg            As Long
    Dim hKey            As Long
    Dim uInfo           As BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
    Dim lResult         As Long

    hResult = BCryptOpenAlgorithmProvider(hAlg, StrPtr("AES"), 0, 0)
    If hResult < 0 Then
        sApiSource = "BCryptOpenAlgorithmProvider"
        GoTo QH
    End If
    hResult = BCryptSetProperty(hAlg, StrPtr("ChainingMode"), StrPtr("ChainingModeGCM"), 32, 0)
    If hResult < 0 Then
        GoTo QH
    End If
    hResult = BCryptGenerateSymmetricKey(hAlg, hKey, 0, 0, VarPtr(baKey(0)), UBound(baKey) + 1, 0)
    If hResult < 0 Then
        GoTo QH
    End If
    With uInfo
        .cbSize = LenB(uInfo)
        .dwInfoVersion = 1
        .pbNonce = VarPtr(baIV(0))
        .cbNonce = UBound(baIV) + 1
        If lAdSize > 0 Then
            .pbAuthData = VarPtr(baAad(lAadPos))
            .cbAuthData = lAdSize
        End If
        .pbTag = VarPtr(baTag(lTagPos))
        .cbTag = lTagSize
    End With
    hResult = BCryptDecrypt(hKey, VarPtr(baBuffer(lPos)), lSize, VarPtr(uInfo), 0, 0, VarPtr(baBuffer(lPos)), lSize, lResult, 0)
    If hResult < 0 Then
        GoTo QH
    End If
    Debug.Assert lResult = lSize
    '--- success
    pvCryptoAeadAesGcmDecrypt = True
QH:
    If hKey <> 0 Then
        Call BCryptDestroyKey(hKey)
    End If
    If hAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(hAlg, 0)
    End If
    If hResult < 0 Then
'        Err.Raise hResult
    End If
End Function

Public Function NewDecodePassword(baBuffer() As Byte, baMasterKey() As Byte) As String
    Dim baIV()          As Byte
    Dim baCipherText()  As Byte
    Dim baEmpty()       As Byte
    
    ReDim baIV(0 To 11) As Byte
    Call CopyMemory(baIV(0), baBuffer(3), UBound(baIV) + 1)
    
    If UBound(baBuffer) > 31 Then 'IMPORTANT!
    
        ReDim baCipherText(0 To UBound(baBuffer) - 15 - 16) As Byte
        Call CopyMemory(baCipherText(0), baBuffer(15), UBound(baCipherText) + 1)
        If Not pvCryptoAeadAesGcmDecrypt(baMasterKey, baIV, baCipherText, 0, UBound(baCipherText) + 1, baBuffer, UBound(baBuffer) - 15, 16, baEmpty, 0, 0) Then
            GoTo QH
        End If
        NewDecodePassword = StrConv(baCipherText, vbUnicode)
    
    End If

QH:

End Function




Attribute VB_Name = "StrConvUTF8"
Option Explicit

'Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long

'Unicode from Leandro
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8      As Long = 65001

Public Function StrConvFromUTF8(text As String) As String
    ' get length
    Dim lngLen As Long, lngPtr As Long: lngLen = LenB(text)
    ' has any?
    If lngLen Then
        ' create a BSTR over twice that length
        lngPtr = SysAllocStringLen(0, lngLen * 1.25)
        ' place it in output variable
        PutMem4 VarPtr(StrConvFromUTF8), lngPtr
        ' convert & get output length
        lngLen = MultiByteToWideChar(65001, 0, ByVal StrPtr(text), lngLen, ByVal lngPtr, LenB(StrConvFromUTF8))
        ' resize the buffer
        StrConvFromUTF8 = Left$(StrConvFromUTF8, lngLen)
    End If
End Function

Public Function StrConvToUTF8(text As String) As String
    ' get length
    Dim lngLen As Long, lngPtr As Long: lngLen = LenB(text)
    ' has any?
    If lngLen Then
        ' create a BSTR over twice that length
        lngPtr = SysAllocStringLen(0, lngLen * 1.25)
        ' place it in output variable
        PutMem4 VarPtr(StrConvToUTF8), lngPtr
        ' convert & get output length
        lngLen = WideCharToMultiByte(65001, 0, ByVal StrPtr(text), Len(text), ByVal lngPtr, LenB(StrConvToUTF8), ByVal 0&, ByVal 0&)
        ' resize the buffer
        StrConvToUTF8 = LeftB$(StrConvToUTF8, lngLen)
    End If
End Function


'Unicode from Leandro
Public Function UTF8ToUnicode(ByVal sUTF8 As String) As String
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

Public Function UnicodeToUTF8(ByVal sData As String) As String
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



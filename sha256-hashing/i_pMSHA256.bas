Attribute VB_Name = "i_pMSHA256"
Option Explicit
Option Private Module

Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
    (ByRef phProv As LongPtr, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" _
    (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hKey As LongPtr, ByVal dwFlags As Long, ByRef phHash As LongPtr) As Long

Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" _
    (ByVal hHash As LongPtr, ByRef pbData As Byte, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" _
    (ByVal hHash As LongPtr, ByVal dwParam As Long, ByRef pbData As Byte, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" _
    (ByVal hHash As LongPtr) As Long

Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" _
    (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long

Const PROV_RSA_AES As Long = 24
Const CALG_SHA_256 As Long = 32780
Const HP_HASHVAL As Long = 2
Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

Public Function s_i_SHA256(ByVal Data As String, Optional ByVal bLowerCase = True) As String
    Dim hProv As LongPtr
    Dim hHash As LongPtr
    Dim DataBytes() As Byte
    Dim HashBytes(31) As Byte
    Dim HashLen As Long
    Dim i As Long
    Dim HashHex As String

    ' Convert input to byte array
    DataBytes = StrConv(Data, vbFromUnicode)
    
    ' Acquire cryptographic context
    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then Exit Function
    
    ' Create hash object
    If CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash) = 0 Then
        CryptReleaseContext hProv, 0
        Exit Function
    End If
    
    ' Hash data
    If CryptHashData(hHash, DataBytes(0), UBound(DataBytes) + 1, 0) = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        Exit Function
    End If
    
    ' Get hash value
    HashLen = 32
    If CryptGetHashParam(hHash, HP_HASHVAL, HashBytes(0), HashLen, 0) = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        Exit Function
    End If
    
    ' Convert to hex string
    For i = 0 To HashLen - 1
        HashHex = HashHex & Right("00" & Hex(HashBytes(i)), 2)
    Next i
    
    ' Clean up
    CryptDestroyHash hHash
    CryptReleaseContext hProv, 0
    
    If bLowerCase Then HashHex = LCase(HashHex)
    s_i_SHA256 = HashHex
End Function



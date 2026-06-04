Attribute VB_Name = "LoaderHash"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
        phProv As LongPtr, _
        ByVal pszContainer As String, _
        ByVal pszProvider As String, _
        ByVal dwProvType As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" ( _
        ByVal hProv As LongPtr, _
        ByVal Algid As Long, _
        ByVal hKey As LongPtr, _
        ByVal dwFlags As Long, _
        phHash As LongPtr) As Long

    Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr, _
        pbData As Any, _
        ByVal dwDataLen As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr, _
        ByVal dwParam As Long, _
        pbData As Any, _
        pdwDataLen As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr) As Long

    Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" ( _
        ByVal hProv As LongPtr, _
        ByVal dwFlags As Long) As Long
#Else
    Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
        phProv As Long, _
        ByVal pszContainer As String, _
        ByVal pszProvider As String, _
        ByVal dwProvType As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function CryptCreateHash Lib "advapi32.dll" ( _
        ByVal hProv As Long, _
        ByVal Algid As Long, _
        ByVal hKey As Long, _
        ByVal dwFlags As Long, _
        phHash As Long) As Long

    Private Declare Function CryptHashData Lib "advapi32.dll" ( _
        ByVal hHash As Long, _
        pbData As Any, _
        ByVal dwDataLen As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function CryptGetHashParam Lib "advapi32.dll" ( _
        ByVal hHash As Long, _
        ByVal dwParam As Long, _
        pbData As Any, _
        pdwDataLen As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function CryptDestroyHash Lib "advapi32.dll" ( _
        ByVal hHash As Long) As Long

    Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
        ByVal hProv As Long, _
        ByVal dwFlags As Long) As Long
#End If

Private Const PROV_RSA_AES As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_SHA_256 As Long = &H800C
Private Const HP_HASHVAL As Long = &H2
Private Const CHUNK_SIZE As Long = 8192

Public Function LoaderHash_FileSha256Hex(ByVal filePath As String) As String
    On Error GoTo failed

#If VBA7 Then
    Dim providerHandle As LongPtr
    Dim hashHandle As LongPtr
#Else
    Dim providerHandle As Long
    Dim hashHandle As Long
#End If
    Dim fileNumber As Integer
    Dim bytesToRead As Long
    Dim chunk() As Byte
    Dim fileLength As Long
    Dim hashLength As Long
    Dim hashBytes() As Byte
    Dim index As Long
    Dim position As Long

    If LenB(Dir$(filePath)) = 0 Then Exit Function

    If CryptAcquireContext(providerHandle, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then
        Exit Function
    End If

    If CryptCreateHash(providerHandle, CALG_SHA_256, 0, 0, hashHandle) = 0 Then
        GoTo failed
    End If

    fileNumber = FreeFile
    Open filePath For Binary Access Read As #fileNumber

    fileLength = LOF(fileNumber)
    position = 1
    Do While position <= fileLength
        bytesToRead = CHUNK_SIZE
        If fileLength - position + 1 < bytesToRead Then
            bytesToRead = fileLength - position + 1
        End If

        If bytesToRead <= 0 Then Exit Do

        ReDim chunk(0 To bytesToRead - 1)
        Get #fileNumber, position, chunk
        position = position + bytesToRead

        If CryptHashData(hashHandle, chunk(0), bytesToRead, 0) = 0 Then
            GoTo failed
        End If
    Loop

    hashLength = 32
    ReDim hashBytes(0 To hashLength - 1)
    If CryptGetHashParam(hashHandle, HP_HASHVAL, hashBytes(0), hashLength, 0) = 0 Then
        GoTo failed
    End If

    For index = 0 To hashLength - 1
        LoaderHash_FileSha256Hex = LoaderHash_FileSha256Hex & LCase$(Right$("0" & Hex$(hashBytes(index)), 2))
    Next index

done:
    On Error Resume Next
    Close #fileNumber
    If hashHandle <> 0 Then CryptDestroyHash hashHandle
    If providerHandle <> 0 Then CryptReleaseContext providerHandle, 0
    Exit Function

failed:
    LoaderHash_FileSha256Hex = ""
    Resume done
End Function

Attribute VB_Name = "Mod_ZIPfunc"
Option Explicit
'This Module is created by Marco v/d Berg
'this module is used to read the contents of a ZIP file for the PSC_Readme-file
'This could be extended to create a full ZIP-Decompressor
'In fact, Whe're creating one right now (Marco v/d Berg and John Korejwa)
'Whe like to let you know that Zip-Decrompression is comming.
'(And not only decompression)

Private Type Local_Header_Type
    VerExt As Integer                   'version needed to extract
    Flags As Integer                    'encrypt and compression flags
    Method As Integer                   'compression method
    FTime As Integer                    'time last modifies, dos format
    FDate As Integer                    'date last modifies, dos format
    CRC32 As Long                       'CRC32 for uncompressed file
    CSize As Long                       'compressed size
    USize As Long                       'uncompressed size
    LenFname As Integer                 'lenght filename
    LenExt As Integer                   'lenght for extra field
End Type
Private Type Extended_Local_Header_Type
    CRC32 As Long                       'CRC32 for uncompressed file
    CSize As Long                       'compressed size
    USize As Long                       'uncompressed size
End Type

Private Const LocalSig As Long = &H4034B50
Private Const CentralSig As Long = &H2014B50
Private Const ExtLocalSig As Long = &H8074B50

Public Function GetNewFileName(ZipName As String, LookFor As String) As String
    Dim FileNum As Long
    Dim Header As Long
    Dim LocDat As Local_Header_Type
    Dim ExtLoc As Extended_Local_Header_Type
    Dim FileName As String
    Dim NewName As String
    Dim data() As Byte
    Dim LN As Long
    Dim X As Long
    If ZipName = "" Then Exit Function
'    If Dir(ZipName, vbNormal) = "" Then Exit Function
    FileNum = FreeFile
    Open ZipName For Binary As #FileNum
    Do
        Get #FileNum, , Header
        Select Case Header
        Case LocalSig
            Get #FileNum, , LocDat
            FileName = String(LocDat.LenFname, 0)
            Get #FileNum, , FileName
            If InStr(FileName, LookFor) > 0 Then
                ReDim data(LocDat.CSize)
                Get #FileNum, , data        'Read in compressed data
                If Inflate(data) <> 0 Then Exit Do  'decompression fault
                NewName = ""
                On Error Resume Next
                LN = UBound(data)
                If Err.Number <> 0 Then Exit Do 'no data
                For X = 0 To LN
                    If data(X) = &HD Then Exit For
                    NewName = NewName & Chr(data(X))
                Next
                NewName = Trim(ReplaceNotPossibble(Mid$(NewName, 8)))
                GetNewFileName = NewName & ".zip"
                Exit Do
            End If
            Seek #FileNum, Seek(FileNum) + LocDat.CSize + LocDat.LenExt
        Case CentralSig
            'Central siognature found so no data is comming anymore
            Exit Do
        Case ExtLocalSig
            Get #FileNum, , ExtLoc
        End Select
    Loop
    Close #FileNum
End Function

'Replace chars wich ar not possible in a filename and clean it a bit
Private Function ReplaceNotPossibble(FileName As String)
    Dim ToUnder As String
    Dim I As Long
    ToUnder = "\/~`!@#$%^&*:;<>,.?"
    For I = 1 To Len(ToUnder)
        FileName = Replace(FileName, Mid(ToUnder, I, 1), "-")
    Next
    FileName = Replace(FileName, ", ", "-")
    FileName = Replace(FileName, "  ", " ")
    FileName = Replace(FileName, " ", "_")
    FileName = Replace(FileName, "__", "_")
    ReplaceNotPossibble = FileName
End Function


Attribute VB_Name = "basZip"
Option Explicit
Private Type ZIPnames
    s(0 To 1023)                  As String
End Type
' Callback large "string" (sic)
Private Type CBChar
    ch(0 To 4096)                 As Byte
End Type
' Callback small "string" (sic)
Private Type CBCh
    ch(0 To 255)                  As Byte
End Type
' Store the callback functions
Private Type ZIPUSERFUNCTIONS
    lPtrPrint                     As Long
' Pointer to application's print routine
    lptrComment                   As Long
' Pointer to application's comment routine
    lptrPassword                  As Long
' Pointer to application's password routine.
    lptrService                   As Long
' callback function designed to be used for allowing the
End Type
Public Type ZPOPT
    date                          As String
' US Date (8 Bytes Long) "12/31/98"?
    szRootDir                     As String
' Root Directory Pathname (Up To 256 Bytes Long)
    szTempDir                     As String
' Temp Directory Pathname (Up To 256 Bytes Long)
    fTemp                         As Long
' 1 If Temp dir Wanted, Else 0
    fSuffix                       As Long
' Include Suffixes (Not Yet Implemented!)
    fEncrypt                      As Long
' 1 If Encryption Wanted, Else 0
    fSystem                       As Long
' 1 To Include System/Hidden Files, Else 0
    fVolume                       As Long
' 1 If Storing Volume Label, Else 0
    fExtra                        As Long
' 1 If Excluding Extra Attributes, Else 0
    fNoDirEntries                 As Long
' 1 If Ignoring Directory Entries, Else 0
    fExcludeDate                  As Long
' 1 If Excluding Files Earlier Than Specified Date, Else 0
    fIncludeDate                  As Long
' 1 If Including Files Earlier Than Specified Date, Else 0
    fVerbose                      As Long
' 1 If Full Messages Wanted, Else 0
    fQuiet                        As Long
' 1 If Minimum Messages Wanted, Else 0
    fCRLF_LF                      As Long
' 1 If Translate CR/LF To LF, Else 0
    fLF_CRLF                      As Long
' 1 If Translate LF To CR/LF, Else 0
    fJunkDir                      As Long
' 1 If Junking Directory Names, Else 0
    fGrow                         As Long
' 1 If Allow Appending To Zip File, Else 0
    fForce                        As Long
' 1 If Making Entries Using DOS File Names, Else 0
    fMove                         As Long
' 1 If Deleting Files Added Or Updated, Else 0
    fDeleteEntries                As Long
' 1 If Files Passed Have To Be Deleted, Else 0
    fUpdate                       As Long
' 1 If Updating Zip File-Overwrite Only If Newer, Else 0
    fFreshen                      As Long
' 1 If Freshing Zip File-Overwrite Only, Else 0
    fJunkSFX                      As Long
' 1 If Junking SFX Prefix, Else 0
    fLatestTime                   As Long
' 1 If Setting Zip File Time To Time Of Latest File In Archive, Else 0
    fComment                      As Long
' 1 If Putting Comment In Zip File, Else 0
    fOffsets                      As Long
' 1 If Updating Archive Offsets For SFX Files, Else 0
    fPrivilege                    As Long
' 1 If Not Saving Privileges, Else 0
    fEncryption                   As Long
' Read Only Property!!!
    fRecurse                      As Long
' 1 (-r), 2 (-R) If Recursing Into Sub-Directories, Else 0
    fRepair                       As Long
' 1 = Fix Archive, 2 = Try Harder To Fix, Else 0
    flevel                        As Byte
' Compression Level - 0 = Stored 6 = Default 9 = Max
End Type
'This assumes zip32.dll is in your \windows\system directory!
' Object for callbacks:
Private m_cZip                As cZip
Private m_bCancel             As Boolean
Private m_bPasswordConfirm    As Boolean
Public MCComments             As String
' Set Zip Callbacks
' Set Zip options
' used to check encryption flag only
Private Declare Function ZpArchive Lib "Zip.dll" (ByVal argc As Long, _
                                                  ByVal funame As String, _
                                                  ByRef argv As ZIPnames) As Long             ' Real zipping action
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
                                                                     lpvSource As Any, _
                                                                     ByVal cbCopy As Long)
Private Declare Function ZpInit Lib "Zip.dll" (ByRef tUserFn As ZIPUSERFUNCTIONS) As Long
Private Declare Function ZpSetOptions Lib "Zip.dll" (ByRef tOpts As ZPOPT) As Long

Private Function plAddressOf(ByVal lPtr As Long) As Long

' VB Bug workaround fn
''Private Declare Function ZpGetOptions Lib "Zip.dll" () As ZPOPT

    plAddressOf = lPtr

End Function

Private Function ReplaceSection(ByRef sString As String, _
                                ByVal sToReplace As String, _
                                ByVal sReplaceWith As String) As Long


Dim iPos     As Long
Dim iLastPos As Long

    On Error Resume Next
    iLastPos = 1
    Do
        iPos = InStr(iLastPos, sString, "/")
        If iPos > 1 Then
            Mid$(sString, iPos, 1) = "\"
            iLastPos = iPos + 1
        End If
    Loop While Not (iPos = 0)
    ReplaceSection = iLastPos
    err.Clear
    

End Function

Public Function VBZip(cZipObject As cZip, _
                      tZPOPT As ZPOPT, _
                      sFileSpecs() As String, _
                      iFileCount As Long) As Long

Dim tUser    As ZIPUSERFUNCTIONS
Dim lR       As Long
Dim i        As Long
Dim sZipFile As String

Dim tZipName As ZIPnames
    On Error Resume Next
    m_bCancel = False
    m_bPasswordConfirm = False
    Set m_cZip = cZipObject
    If Not Len(Trim$(m_cZip.BasePath)) = 0 Then
        ChDir m_cZip.BasePath
    End If
' Set address of callback functions
    With tUser
        .lPtrPrint = plAddressOf(AddressOf ZipPrintCallback)
        .lptrPassword = plAddressOf(AddressOf ZipPasswordCallback)
        .lptrComment = plAddressOf(AddressOf ZipCommentCallback)
        .lptrService = plAddressOf(AddressOf ZipServiceCallback)
    End With 'tUser
    lR = ZpInit(tUser)
' Set options
    lR = ZpSetOptions(tZPOPT)
' Go for it!
    For i = 1 To iFileCount
        tZipName.s(i - 1) = sFileSpecs(i)
    Next i
    tZipName.s(i) = vbNullChar
    sZipFile = cZipObject.ZipFile
    lR = ZpArchive(iFileCount, sZipFile, tZipName)
    If lR > 0 Then
        MsgBox "Error : " & lR & vbNewLine & err.Description
    End If
    VBZip = lR
    
    If err.Number <> 0 Then MsgBox err.Description, vbCritical: err.Clear
    
    
    

End Function

Private Function ZipCommentCallback(ByRef comm As CBChar) As Integer

Dim bCancel  As Boolean
Dim sComment As String
Dim b()      As Byte
Dim lSize    As Long
Dim i        As Long

    On Error Resume Next
' not supported always return \0
'Set comm.ch(0) = vbNullString
    ZipCommentCallback = 1
    If m_bCancel Then
    Else
' Ask for a comment:
        m_cZip.CommentRequest sComment, bCancel
        sComment = Trim$(MCComments)
' Cancel out if no useful comment:
        If bCancel Or Len(sComment) = 0 Then
            m_bCancel = True
            Exit Function
        End If
' Put password into return parameter:
        lSize = Len(sComment)
        b = StrConv(sComment, vbFromUnicode)
        lSize = UBound(b)
        If lSize > 4095 Then
            lSize = 4095
        End If
        For i = 0 To lSize
            comm.ch(i) = b(i)
        Next i
        comm.ch(lSize + 1) = 0
    End If
    ZipCommentCallback = 0
    err.Clear
    

End Function

Private Function ZipPasswordCallback(ByRef pwd As CBCh, _
                                     ByVal maxPasswordLength As Long, _
                                     ByRef s2 As CBCh, _
                                     ByRef cbcName As CBCh) As Integer


Dim bCancel   As Boolean
Dim sPassword As String
Dim b()       As Byte
Dim i         As Long

    On Error Resume Next
    ZipPasswordCallback = 1
    If m_bCancel Then
    Else
' Ask for password:
        m_cZip.PasswordRequest sPassword, maxPasswordLength, m_bPasswordConfirm, bCancel
        sPassword = Trim$(sPassword)
' Cancel out if no useful password:
        If bCancel Or Len(sPassword) = 0 Then
            m_bCancel = True
            Exit Function
        End If
' Put password into return parameter:
'lSize = Len(sPassword)
        b = StrConv(sPassword, vbFromUnicode)
        For i = 0 To maxPasswordLength - 1
            If i <= UBound(b) Then
                pwd.ch(i) = b(i)
            Else
                pwd.ch(i) = 0
            End If
        Next i
' Ask UnZip to process it:
        m_bPasswordConfirm = True
    End If
    ZipPasswordCallback = 0
    err.Clear

End Function

Private Function ZipPrintCallback(ByRef fname As CBChar, _
                                  ByVal x As Long) As Long

Dim iPos  As Long
Dim Sfile As String

    On Error Resume Next
    If x > 1 Then
        If x < 1024 Then

            ReDim b(0 To x) As Byte
            CopyMemory b(0), fname, x

            Sfile = StrConv(b, vbUnicode)
            If iPos > 0 Then
                Sfile = Left$(Sfile, iPos - 1)
            End If

            ReplaceSection Sfile, "/", "\"

            m_cZip.ProgressReport Sfile
        End If
    End If
    ZipPrintCallback = 0
    err.Clear


End Function

Private Function ZipServiceCallback(ByRef mname As CBChar, _
                                    ByVal x As Long) As Long

Dim iPos    As Long
Dim sInfo   As String
Dim bCancel As Boolean

    On Error Resume Next
    If x > 1 Then
        If x < 1024 Then
' If so, then get the readable portion of it:
            ReDim b(0 To x) As Byte
            CopyMemory b(0), mname, x
' Convert to VB string:
            sInfo = StrConv(b, vbUnicode)
            iPos = InStr(sInfo, vbNullChar)
            If iPos > 0 Then
                sInfo = Left$(sInfo, iPos - 1)
            End If
            m_cZip.Service sInfo, bCancel
            If bCancel Then
                m_bCancel = True
                ZipServiceCallback = 1
            Else
                ZipServiceCallback = 0
            End If
        End If
    End If
    err.Clear
    

End Function



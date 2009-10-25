Attribute VB_Name = "MpqStuff"
Option Explicit

Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
 
    ' Optional members
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Declare Function ShellExecute Lib _
    "Shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Public Declare Function ShellExecuteEx Lib _
    "Shell32.dll" Alias "ShellExecuteExA" _
    (sei As SHELLEXECUTEINFO) As Long
Public Declare Sub SHChangeNotify Lib _
    "Shell32.dll" (ByVal wEventId As Long, _
    ByVal uFlags As Integer, _
    ByVal dwItem1 As Any, _
    ByVal dwItem2 As Any)
Public Declare Function SendMessageA Lib _
    "User32.dll" _
    (ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal Wp As Long, _
    Lp As Any) As Long
Declare Function GetLongPathName Lib "Kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32.dll" _
    Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)

Public CD As OPENFILENAME, PathInput As BROWSEINFO
Public GlobalFileList() As String, FileList() As String, CX As Single, CY As Single, NewFile As Boolean, LocaleID As Long, ListFile As String, AddFolderName As String, ExtractPathNum As Long, CopyPathNum As Long, GlobalEncrypt As Boolean, DefaultCompress As Long, DefaultCompressID As Long, DefaultCompressLevel As Long, DefaultMaxFiles As Long, DefaultBlockSize As Long
Public Const AppKey As String = "HKEY_CURRENT_USER\Software\ShadowFlare\WinMPQ\", SharedAppKey As String = "HKEY_LOCAL_MACHINE\Software\ShadowFlare\WinMPQ\"
Public Const MPQ_ERROR_INIT_FAILED As Long = &H85000001 'Unspecified error
Public Const MPQ_ERROR_NO_STAREDIT As Long = &H85000002 'Can't find StarEdit.exe
Public Const MPQ_ERROR_BAD_STAREDIT As Long = &H85000003 'Bad version of StarEdit.exe. Need SC/BW 1.07
Public Const MPQ_ERROR_STAREDIT_RUNNING As Long = &H85000004 'StarEdit.exe is running. Must be closed
Public Const SHCNE_ASSOCCHANGED As Long = &H8000000
Public Const SHCNF_IDLIST  As Long = &H0
Public Const WM_SETREDRAW As Long = &HB
Public Const WM_PAINT  As Long = &HF
Const gintMAX_SIZE% = 255
Public Const SEE_MASK_CLASSNAME As Long = &H1

Sub GetCompressFlags(File As String, ByRef cType As Integer, ByRef dwFlags As Long)
Dim bNum As Long, fExt As String
dwFlags = MAFA_REPLACE_EXISTING
If GlobalEncrypt Then dwFlags = dwFlags Or MAFA_ENCRYPT
For bNum = 1 To Len(File)
    If InStr(bNum, File, ".") > 0 Then
        bNum = InStr(bNum, File, ".")
    Else
        Exit For
    End If
Next bNum
If bNum > 1 Then
    fExt = Mid(File, bNum - 1)
Else
    fExt = File
End If
If LCase(fExt) = ".bik" Then
    cType = CInt(GetReg(AppKey + "Compression\.bik", "-2"))
    dwFlags = dwFlags And (-1& Xor MAFA_ENCRYPT)
ElseIf LCase(fExt) = ".smk" Then
    cType = CInt(GetReg(AppKey + "Compression\.smk", "-2"))
    dwFlags = dwFlags And (-1& Xor MAFA_ENCRYPT)
ElseIf LCase(fExt) = ".mp3" Then
    cType = CInt(GetReg(AppKey + "Compression\.mp3", "-2"))
    dwFlags = dwFlags And (-1& Xor MAFA_ENCRYPT)
ElseIf LCase(fExt) = ".mpq" Then
    cType = CInt(GetReg(AppKey + "Compression\.mpq", "-2"))
    dwFlags = dwFlags And (-1& Xor MAFA_ENCRYPT)
ElseIf LCase(fExt) = ".scm" Then
    cType = CInt(GetReg(AppKey + "Compression\.scm", "-2"))
    dwFlags = dwFlags And (-1& Xor MAFA_ENCRYPT)
ElseIf LCase(fExt) = ".scx" Then
    cType = CInt(GetReg(AppKey + "Compression\.scx", "-2"))
    dwFlags = dwFlags And (-1& Xor MAFA_ENCRYPT)
ElseIf LCase(fExt) = ".w3m" Then
    cType = CInt(GetReg(AppKey + "Compression\.w3m", "-2"))
    dwFlags = dwFlags And (-1& Xor MAFA_ENCRYPT)
ElseIf LCase(fExt) = ".w3x" Then
    cType = CInt(GetReg(AppKey + "Compression\.w3x", "-2"))
    dwFlags = dwFlags And (-1& Xor MAFA_ENCRYPT)
ElseIf LCase(fExt) = ".wav" Then
    cType = CInt(GetReg(AppKey + "Compression\.wav", "0"))
Else
    cType = CInt(GetReg(AppKey + "Compression\" + fExt, CStr(DefaultCompressID)))
End If
End Sub

Function mOpenMpq(FileName As String) As Long
Dim hMPQ As Long
mOpenMpq = 0
hMPQ = MpqOpenArchiveForUpdateEx(FileName, MOAU_OPEN_EXISTING Or MOAU_MAINTAIN_LISTFILE, DefaultMaxFiles, DefaultBlockSize)
If hMPQ = 0 Or hMPQ = INVALID_HANDLE_VALUE Then
    hMPQ = MpqOpenArchiveForUpdateEx(FileName, MOAU_CREATE_NEW Or MOAU_MAINTAIN_LISTFILE, DefaultMaxFiles, DefaultBlockSize)
End If
If hMPQ <> 0 And hMPQ <> INVALID_HANDLE_VALUE Then
    mOpenMpq = hMPQ
End If
End Function
Function PathInputBox(lpFolderDialog As BROWSEINFO, pCaption As String, StartFolder As String) As String
lpFolderDialog.Title = pCaption
Dim result As Long
result = ShowFolder(lpFolderDialog)
If result = 0 Then Exit Function
PathInputBox = GetPathFromID(result)
End Function
Function GetLongPath(Path As String) As String
    Dim strBuf As String, StrLength As Long
    strBuf = Space$(gintMAX_SIZE)
    StrLength = GetLongPathName(Path, strBuf, gintMAX_SIZE)
    strBuf = Left(strBuf, StrLength)
    If strBuf <> "" Then
        GetLongPath = strBuf
    Else
        GetLongPath = Path
    End If
End Function
Sub AddScriptOutput(sOutput As String)
SendMessageA ScriptOut.oText.hWnd, WM_SETREDRAW, 0, ByVal 0&
ScriptOut.oText = ScriptOut.oText + sOutput
SendMessageA ScriptOut.oText.hWnd, WM_SETREDRAW, 1, ByVal 0&
ScriptOut.oText.SelStart = Len(ScriptOut.oText)
End Sub
Function GetFileTitle(FileName As String) As String
Dim bNum As Long
If InStr(FileName, "\") > 0 Then
    For bNum = 1 To Len(FileName)
        If InStr(bNum, FileName, "\") > 0 Then
            bNum = InStr(bNum, FileName, "\")
        Else
            Exit For
        End If
    Next bNum
    GetFileTitle = Mid(FileName, bNum)
Else
    GetFileTitle = FileName
End If
End Function
Function sGetFile(hMPQ As Long, ByVal FileName As String, OutPath As String, ByVal UseFullPath As Long)
Dim hFile As Long, buffer() As Byte, fLen As Long, cNum As Long
If SFileOpenFileEx(hMPQ, FileName, 0, hFile) Then
    fLen = SFileGetFileSize(hFile, 0)
    If fLen > 0 Then
        ReDim buffer(fLen - 1)
    Else
        ReDim buffer(0)
    End If
    SFileReadFile hFile, buffer(0), fLen, ByVal 0, ByVal 0
    SFileCloseFile hFile
    If UseFullPath = 0 Then FileName = GetFileTitle(FileName)
    FileName = FullPath(OutPath, FileName)
    On Error Resume Next
    For cNum = 1 To Len(FileName)
        cNum = InStr(cNum, FileName, "\")
        If cNum > 0 Then
            MkDir Left(FileName, cNum)
        Else
            Exit For
        End If
    Next cNum
    If FileExists(FileName) Then Kill FileName
    On Error GoTo 0
    cNum = FreeFile
    On Error GoTo WriteError
    Open FileName For Binary As #cNum
        If fLen > 0 Then Put #cNum, 1, buffer
    Close #cNum
    On Error GoTo 0
End If
Exit Function
WriteError:
MsgBox "Error writing file.  File may be in use.", vbCritical, "WinMPQ"
Close #cNum
End Function
Function sListFiles(MpqName As String, hMPQ As Long, ByVal FileLists As String, ByRef ListedFiles() As FILELISTENTRY) As Boolean
Dim NewFileLists As String, nFileLists() As String, ListName As String, cNum As Long, cNum2 As Long, cNum3 As Long, cNum4 As Long, MpqList1 As String, MpqList2 As String, Path As String, ListLen As Long, OldLists() As String, UseOnlyAutoList As Boolean, nHash As Long, nHashEntries As Long
sListFiles = False
ReDim ListedFiles(0)
ListedFiles(0).dwFileExists = 0
If GetReg(AppKey + "AutofindFileLists", 0) = 0 Then
    NewFileLists = FileLists
Else
    UseOnlyAutoList = GetReg(AppKey + "UseOnlyAutofindLists", 1)
    MpqList2 = GetExtension(MpqName)
    MpqList1 = GetFileTitle(Left(MpqName, Len(MpqName) - Len(MpqList2))) + ".txt"
    MpqList2 = GetFileTitle(MpqName) + ".txt"
    Path = GetLongPath(App.Path)
    If Right(Path, 1) <> "\" Then Path = Path + "\"
    If UseOnlyAutoList Then ListLen = Len(FileLists)
    If FileLists <> "" Then
        FileLists = FileLists + vbCrLf + Path + App.EXEName + ".exe" + vbCrLf + MpqName
    Else
        FileLists = Path + App.EXEName + ".exe" + vbCrLf + MpqName
    End If
    ReDim nFileLists(0) As String
    If UseOnlyAutoList Then ReDim OldLists(0) As String
    For cNum = 1 To Len(FileLists)
        cNum2 = InStr(cNum, FileLists, vbCrLf)
        If cNum2 = 0 Then
            cNum2 = Len(FileLists) + 1
        End If
        If cNum2 - cNum > 0 Then
            ListName = Mid(FileLists, cNum, cNum2 - cNum)
            If Not IsDir(ListName) Then
                If UseOnlyAutoList And cNum < ListLen Then
                    ReDim Preserve OldLists(UBound(OldLists) + 1) As String
                    OldLists(UBound(OldLists)) = GetLongPath(ListName)
                End If
                For cNum3 = 1 To Len(ListName)
                    If InStr(cNum3, ListName, "\") Then
                        cNum3 = InStr(cNum3, ListName, "\")
                        If FileExists(Left(ListName, cNum3) + MpqList1) Then
                            ReDim Preserve nFileLists(UBound(nFileLists) + 1) As String
                            nFileLists(UBound(nFileLists)) = GetLongPath(Left(ListName, cNum3) + MpqList1)
                        End If
                        If FileExists(Left(ListName, cNum3) + MpqList2) Then
                            ReDim Preserve nFileLists(UBound(nFileLists) + 1) As String
                            nFileLists(UBound(nFileLists)) = GetLongPath(Left(ListName, cNum3) + MpqList2)
                        End If
                    Else
                        Exit For
                    End If
                Next cNum3
                If FileExists(ListName) And ListName <> Path + App.EXEName + ".exe" And ListName <> MpqName Then
                    ReDim Preserve nFileLists(UBound(nFileLists) + 1) As String
                    nFileLists(UBound(nFileLists)) = GetLongPath(ListName)
                End If
            Else
                ListName = DirEx(ListName, MpqList1, 6, True) _
                         + DirEx(ListName, MpqList2, 6, True)
                For cNum3 = 1 To Len(ListName)
                    cNum4 = InStr(cNum3, ListName, vbCrLf)
                    If cNum4 = 0 Then
                        cNum4 = Len(ListName) + 1
                    End If
                    If cNum4 - cNum3 > 0 Then
                        ReDim Preserve nFileLists(UBound(nFileLists) + 1) As String
                        nFileLists(UBound(nFileLists)) = GetLongPath(Mid(ListName, cNum3, cNum4 - cNum3))
                    End If
                    cNum3 = cNum4 + 1
                Next cNum3
            End If
        End If
        cNum = cNum2 + 1
    Next cNum
    If UseOnlyAutoList Then
        For cNum = 1 To UBound(nFileLists)
            For cNum2 = 1 To UBound(OldLists)
                If LCase(nFileLists(cNum)) <> LCase(OldLists(cNum2)) Then
                    GoTo StartSearch
                End If
            Next cNum2
        Next cNum
        UseOnlyAutoList = False
    End If
StartSearch:
    For cNum = 1 To UBound(nFileLists)
        If nFileLists(cNum) <> "" Then
            For cNum2 = 1 To UBound(nFileLists)
                If LCase(nFileLists(cNum)) = LCase(nFileLists(cNum2)) And cNum <> cNum2 Then
                    nFileLists(cNum2) = ""
                End If
            Next cNum2
        End If
        If UseOnlyAutoList Then
            If nFileLists(cNum) <> "" Then
                For cNum2 = 1 To UBound(OldLists)
                    If LCase(nFileLists(cNum)) = LCase(OldLists(cNum2)) And LCase(GetFileTitle(nFileLists(cNum))) <> LCase(MpqList1) And LCase(GetFileTitle(nFileLists(cNum))) <> LCase(MpqList2) Then
                        nFileLists(cNum) = ""
                        Exit For
                    End If
                Next cNum2
            End If
        End If
        If nFileLists(cNum) <> "" Then
            NewFileLists = NewFileLists + nFileLists(cNum) + vbCrLf
        End If
    Next cNum
    If Right(NewFileLists, 2) = vbCrLf Then NewFileLists = Left(NewFileLists, Len(NewFileLists) - 2)
End If
nHashEntries = SFileGetFileInfo(hMPQ, SFILE_INFO_HASH_TABLE_SIZE)
If nHashEntries - 1 < 0 Then Exit Function
ReDim ListedFiles(nHashEntries - 1)
sListFiles = SFileListFiles(hMPQ, NewFileLists, ListedFiles(0), 0)
End Function
Sub mAddAutoFile(hMPQ As Long, File As String, MpqPath As String)
Dim cType As Integer, dwFlags As Long

GetCompressFlags File, cType, dwFlags

Select Case cType
Case -2
MpqAddFileToArchiveEx hMPQ, File, MpqPath, dwFlags, 0, 0
Case -1
MpqAddFileToArchiveEx hMPQ, File, MpqPath, dwFlags Or MAFA_COMPRESS, MAFA_COMPRESS_STANDARD, 0
Case -3
MpqAddFileToArchiveEx hMPQ, File, MpqPath, dwFlags Or MAFA_COMPRESS, MAFA_COMPRESS_DEFLATE, DefaultCompressLevel
Case -4
MpqAddFileToArchiveEx hMPQ, File, MpqPath, dwFlags Or MAFA_COMPRESS, MAFA_COMPRESS_BZIP2, 0
Case 0, 1, 2
MpqAddWaveToArchive hMPQ, File, MpqPath, dwFlags Or MAFA_COMPRESS, cType
Case Else
If DefaultCompress = MAFA_COMPRESS_DEFLATE Then
    MpqAddFileToArchiveEx hMPQ, File, MpqPath, dwFlags Or MAFA_COMPRESS, DefaultCompress, DefaultCompressLevel
Else
    MpqAddFileToArchiveEx hMPQ, File, MpqPath, dwFlags Or MAFA_COMPRESS, DefaultCompress, 0
End If
End Select
End Sub
Sub mAddAutoFromBuffer(hMPQ As Long, ByRef buffer As Byte, BufSize As Long, MpqPath As String)
Dim cType As Integer, dwFlags As Long

GetCompressFlags MpqPath, cType, dwFlags

Select Case cType
Case -2
MpqAddFileFromBufferEx hMPQ, buffer, BufSize, MpqPath, dwFlags, 0, 0
Case -1
MpqAddFileFromBufferEx hMPQ, buffer, BufSize, MpqPath, dwFlags Or MAFA_COMPRESS, MAFA_COMPRESS_STANDARD, 0
Case -3
MpqAddFileFromBufferEx hMPQ, buffer, BufSize, MpqPath, dwFlags Or MAFA_COMPRESS, MAFA_COMPRESS_DEFLATE, DefaultCompressLevel
Case -4
MpqAddFileFromBufferEx hMPQ, buffer, BufSize, MpqPath, dwFlags Or MAFA_COMPRESS, MAFA_COMPRESS_BZIP2, 0
Case 0, 1, 2
MpqAddWaveFromBuffer hMPQ, buffer, BufSize, MpqPath, dwFlags Or MAFA_COMPRESS, cType
Case Else
If DefaultCompress = MAFA_COMPRESS_DEFLATE Then
    MpqAddFileFromBufferEx hMPQ, buffer, BufSize, MpqPath, dwFlags Or MAFA_COMPRESS, DefaultCompress, DefaultCompressLevel
Else
    MpqAddFileFromBufferEx hMPQ, buffer, BufSize, MpqPath, dwFlags Or MAFA_COMPRESS, DefaultCompress, 0
End If
End Select
End Sub

Function DirEx(ByVal Path As String, Filter As String, Attributes, Recurse As Boolean) As String
Dim Files() As String, lNum As Long, Folders() As String
If Right(Path, 1) <> "\" And Path <> "" Then Path = Path + "\"
ReDim Files(0) As String
Files(0) = Dir(Path + Filter, Attributes)
If Files(0) <> "" Then
    Do
    ReDim Preserve Files(UBound(Files) + 1) As String
    Files(UBound(Files)) = Dir
    Loop Until Files(UBound(Files)) = ""
    ReDim Preserve Files(UBound(Files) - 1) As String
End If
For lNum = 0 To UBound(Files)
    If Files(lNum) <> "" Then
        If IsDir(Path + Files(lNum)) = False And (Attributes And vbDirectory) <> vbDirectory Then
            DirEx = DirEx + Path + Files(lNum) + vbCrLf
        ElseIf IsDir(Path + Files(lNum)) = True And (Attributes And vbDirectory) Then
            DirEx = DirEx + Path + Files(lNum) + vbCrLf
        End If
    End If
Next lNum
If Recurse = True And (InStr(Filter, "?") > 0 Or InStr(Filter, "*") > 0) Then
    ReDim Folders(0) As String
    Folders(0) = Dir(Path, vbDirectory)
    If Folders(0) = "." Or Folders(0) = ".." Then Folders(0) = Dir
    If Folders(0) = "." Or Folders(0) = ".." Then Folders(0) = Dir
    If Folders(0) <> "" Then
        Do
        ReDim Preserve Folders(UBound(Folders) + 1) As String
        Folders(UBound(Folders)) = Dir
        If Folders(UBound(Folders)) = "." Or Folders(UBound(Folders)) = ".." Then
            ReDim Preserve Folders(UBound(Folders) - 1) As String
        End If
        Loop Until Folders(UBound(Folders)) = ""
        ReDim Preserve Folders(UBound(Folders) - 1) As String
    End If
    For lNum = 0 To UBound(Folders)
        If Folders(lNum) <> "" Then
            If IsDir(Path + Folders(lNum)) Then
                DirEx = DirEx + DirEx(Path + Folders(lNum), Filter, Attributes, Recurse)
            End If
        End If
    Next lNum
End If
End Function
Function GetExtension(FileName As String) As String
Dim bNum As Long
If InStr(FileName, ".") > 0 Then
    For bNum = 1 To Len(FileName)
        If InStr(bNum, FileName, ".") > 0 Then
            bNum = InStr(bNum, FileName, ".")
        Else
            Exit For
        End If
    Next bNum
    GetExtension = Mid(FileName, bNum - 1)
Else
    GetExtension = ""
End If
End Function
Function IsDir(DirPath As String) As Boolean
On Error GoTo IsNotDir
If GetAttr(DirPath) And vbDirectory Then
    IsDir = True
Else
    IsDir = False
End If
Exit Function
IsNotDir:
IsDir = False
End Function
Function FileExists(FileName As String) As Boolean
On Error GoTo NoFile
If (GetAttr(FileName) And vbDirectory) <> vbDirectory Then
    FileExists = True
Else
    FileExists = False
End If
Exit Function
NoFile:
FileExists = False
End Function
Function IsMPQ(MpqFile As String) As Boolean
If FindMpqHeader(MpqFile) <> -1 Then
    IsMPQ = True
Else
    IsMPQ = False
End If
End Function
Sub KillEx(ByVal Path As String, Filter As String, Attributes, Recurse As Boolean)
Dim Files() As String, lNum As Long, Folders() As String
If Right(Path, 1) <> "\" And Path <> "" Then Path = Path + "\"
ReDim Files(0) As String
Files(0) = Dir(Path + Filter, Attributes)
If Files(0) <> "" Then
    Do
    ReDim Preserve Files(UBound(Files) + 1) As String
    Files(UBound(Files)) = Dir
    Loop Until Files(UBound(Files)) = ""
    ReDim Preserve Files(UBound(Files) - 1) As String
End If
For lNum = 0 To UBound(Files)
    If Files(lNum) <> "" Then
        If IsDir(Path + Files(lNum)) = False Then
            On Error Resume Next
            Kill Path + Files(lNum)
            On Error GoTo 0
        End If
    End If
Next lNum
If Recurse = True And (InStr(Filter, "?") > 0 Or InStr(Filter, "*") > 0) Then
    ReDim Folders(0) As String
    Folders(0) = Dir(Path, vbDirectory)
    If Folders(0) = "." Or Folders(0) = ".." Then Folders(0) = Dir
    If Folders(0) = "." Or Folders(0) = ".." Then Folders(0) = Dir
    If Folders(0) <> "" Then
        Do
        ReDim Preserve Folders(UBound(Folders) + 1) As String
        Folders(UBound(Folders)) = Dir
        If Folders(UBound(Folders)) = "." Or Folders(UBound(Folders)) = ".." Then
            ReDim Preserve Folders(UBound(Folders) - 1) As String
        End If
        Loop Until Folders(UBound(Folders)) = ""
        ReDim Preserve Folders(UBound(Folders) - 1) As String
    End If
    For lNum = 0 To UBound(Folders)
        If Folders(lNum) <> "" Then
            If IsDir(Path + Folders(lNum)) Then
                KillEx Path + Folders(lNum), Filter, Attributes, Recurse
                On Error Resume Next
                RmDir Path + Folders(lNum)
            End If
            On Error GoTo 0
        End If
    Next lNum
End If
End Sub
Function FullPath(ByVal BasePath As String, File As String) As String
If Right(BasePath, 1) <> "\" Then BasePath = BasePath + "\"
If Mid(File, 2, 1) = ":" Or Left(File, 2) = "\\" Then
    FullPath = File
ElseIf Left(File, 1) = "\" Then
    FullPath = Left(BasePath, 2) + File
Else
    FullPath = BasePath + File
End If
End Function
Function MatchesFilter(FileName As String, ByVal Filters As String) As Boolean
Dim bNum As Long, Filter As String
If InStr(Filters, ";") Then
    If Right(Filters, 1) <> ";" Then Filters = Filters + ";"
    For bNum = 1 To Len(Filters)
        Filter = Mid(Filters, bNum, InStr(bNum, Filters, ";") - bNum)
        If Right(Filter, 3) = "*.*" Then Filter = Left(Filter, Len(Filter) - 2)
        If LCase(FileName) Like LCase(Filter) Then
            MatchesFilter = True
            Exit Function
        End If
        bNum = InStr(bNum, Filters, ";")
    Next bNum
Else
    If Right(Filters, 3) = "*.*" Then Filters = Left(Filters, Len(Filters) - 2)
    If LCase(FileName) Like LCase(Filters) Then MatchesFilter = True
End If
End Function
Function RenameWithFilter(FileName As String, OldFilter As String, NewFilter As String) As String
Dim bNum As Long, Filters() As String, NewFileName As String, bNum2 As Long, bNum3 As Long, bNum4 As Long, bNum5 As Long
If Right(OldFilter, 3) = "*.*" Then OldFilter = Left(OldFilter, Len(OldFilter) - 2)
If Right(NewFilter, 3) = "*.*" Then NewFilter = Left(NewFilter, Len(NewFilter) - 2)
ReDim Filters(0) As String
bNum4 = 1
For bNum = 1 To Len(OldFilter)
    Select Case Mid(OldFilter, bNum, 1)
    Case "*"
        bNum2 = InStr(bNum + 1, OldFilter, "*")
        bNum3 = InStr(bNum + 1, OldFilter, "?")
        If bNum2 = 0 And bNum3 = 0 Then
            bNum2 = Len(OldFilter) + 1
        ElseIf (bNum3 < bNum2 Or bNum2 = 0) And bNum3 > 0 Then
            bNum2 = bNum3
        End If
        bNum5 = InStr(bNum4, FileName, Mid(OldFilter, bNum + 1, bNum2 - bNum - 1), 1)
        If bNum = Len(OldFilter) Then
            bNum5 = Len(FileName) + 1
        End If
        If bNum5 = 0 Then
            RenameWithFilter = FileName
            Exit Function
        End If
        If bNum > 1 Then
            If Mid(OldFilter, bNum - 1, 1) <> "*" And Mid(OldFilter, bNum - 1, 1) <> "?" Then
                ReDim Preserve Filters(UBound(Filters) + 1) As String
            End If
        Else
            ReDim Preserve Filters(UBound(Filters) + 1) As String
        End If
        Filters(UBound(Filters)) = Filters(UBound(Filters)) + Mid(FileName, bNum4, bNum5 - bNum4)
        bNum4 = bNum5
    Case "?"
        bNum2 = bNum + 1
        bNum5 = bNum4 + 1
        If bNum > 1 Then
            If Mid(OldFilter, bNum - 1, 1) <> "*" And Mid(OldFilter, bNum - 1, 1) <> "?" Then
                ReDim Preserve Filters(UBound(Filters) + 1) As String
            End If
        Else
            ReDim Preserve Filters(UBound(Filters) + 1) As String
        End If
        Filters(UBound(Filters)) = Filters(UBound(Filters)) + Mid(FileName, bNum4, 1)
        bNum4 = bNum5
    Case Else
        bNum4 = bNum4 + 1
    End Select
    If bNum4 > Len(FileName) Then
        If (Right(OldFilter, 1) <> "*" Or bNum + 1 < Len(OldFilter)) And bNum < Len(OldFilter) Then
            RenameWithFilter = FileName
            Exit Function
        Else
            Exit For
        End If
    End If
Next bNum
NewFileName = NewFilter
For bNum = 1 To UBound(Filters)
    bNum2 = InStr(bNum, NewFileName, "*")
    bNum3 = InStr(bNum, NewFileName, "?")
    If bNum2 = 0 And bNum3 = 0 Then
        bNum2 = Len(NewFileName) + 1
    ElseIf (bNum3 < bNum2 Or bNum2 = 0) And bNum3 > 0 Then
        bNum2 = bNum3
    End If
    If bNum2 > Len(NewFileName) Then
        RenameWithFilter = NewFileName
        Exit Function
    End If
    bNum4 = 0
    For bNum3 = bNum2 To Len(NewFileName)
        Select Case Mid(NewFileName, bNum3, 1)
        Case "*"
            bNum4 = Len(Filters(bNum))
            bNum3 = bNum3 + 1
            Exit For
        Case "?"
            bNum4 = bNum4 + 1
        Case Else
            Exit For
        End Select
    Next bNum3
    NewFileName = Left(NewFileName, bNum2 - 1) + Left(Filters(bNum), bNum4) + Mid(NewFileName, bNum3)
Next bNum
Do Until InStr(NewFileName, "*") = 0
    NewFileName = Left(NewFileName, InStr(NewFileName, "*") - 1) + Mid(NewFileName, InStr(NewFileName, "*") + 1)
Loop
Do Until InStr(NewFileName, "?") = 0
    NewFileName = Left(NewFileName, InStr(NewFileName, "?") - 1) + Mid(NewFileName, InStr(NewFileName, "?") + 1)
Loop
RenameWithFilter = NewFileName
End Function
Function MpqDir(MpqFile As String, Filters As String)
Dim Files() As FILELISTENTRY, fNum As Long, szFileList As String, NamePos As Long, CurFileName As String
Dim hMPQ As Long
If SFileOpenArchive(MpqFile, 0, 0, hMPQ) Then
    If sListFiles(MpqFile, hMPQ, ListFile, Files) Then
        SFileCloseArchive hMPQ
        For fNum = 0 To UBound(Files)
            If Files(fNum).dwFileExists Then
                CurFileName = StrConv(Files(fNum).szFileName, vbUnicode)
                If MatchesFilter(CurFileName, Filters) Then
                    NamePos = InStr(1, szFileList, CurFileName + vbCrLf, 1)
                    If NamePos > 1 Then
                        NamePos = InStr(1, szFileList, vbCrLf + CurFileName + vbCrLf, 1)
                    End If
                    If NamePos > 0 Then _
                        szFileList = szFileList + CurFileName
                End If
            End If
        Next fNum
        MpqDir = MpqDir + CurFileName + vbCrLf
    Else
        SFileCloseArchive hMPQ
    End If
End If
End Function
Sub RunScript(ScriptName As String)
Dim fNum As Long, Script As String, sLine As String, Param() As String, bNum As Long, EndLine As Long, pNum As Long, EndParam As Long, MpqFile As String, OldDefaultMaxFiles As Long, cType As Integer, lNum As Long, OldPath As String, NewPath As String, Rswitch As Boolean, Files As String, fCount As Long, fEndLine As Long, fLine As String, ScriptNewFile As Boolean, CurPath As String, fLine2 As String, fLineTitle As String, hMPQ As Long, hFile As Long, dwFlags
If FileExists(ScriptName) = False Then
    ScriptOut.Show
    AddScriptOutput "Could not find script " + ScriptName + vbCrLf
    Exit Sub
End If
fNum = FreeFile
Open ScriptName For Binary As #fNum
Script = String(LOF(fNum), Chr(0))
Get #fNum, 1, Script
Close #fNum
OldPath = CurDir
If InStr(ScriptName, "\") > 0 Then
    For bNum = 1 To Len(ScriptName)
        If InStr(bNum, ScriptName, "\") > 0 Then
            bNum = InStr(bNum, ScriptName, "\")
            NewPath = Left(ScriptName, bNum)
        End If
    Next bNum
    If Mid(NewPath, 2, 1) = ":" Then ChDrive Left(NewPath, 1)
    ChDir NewPath
End If
CurPath = CurDir
If Right(Script, 2) <> vbCrLf Then Script = Script + vbCrLf
ScriptOut.Show
AddScriptOutput "Script: " + ScriptName + vbCrLf + vbCrLf
OldDefaultMaxFiles = DefaultMaxFiles
lNum = 1
For bNum = 1 To Len(Script)
    EndLine = InStr(bNum, Script, vbCrLf)
    sLine = Mid(Script, bNum, EndLine - bNum)
    If Right(sLine, 1) <> " " Then sLine = sLine + " "
    If sLine <> "" Then
        AddScriptOutput "Line " + CStr(lNum) + ": "
        ReDim Param(0) As String
        For pNum = 1 To Len(sLine)
            If Mid(sLine, pNum, 1) = Chr(34) Then
                pNum = pNum + 1
                EndParam = InStr(pNum, sLine, Chr(34))
            Else
                EndParam = InStr(pNum, sLine, " ")
            End If
            If EndParam = 0 Then EndParam = Len(sLine) + 1
            If pNum <> EndParam Then
                If Trim(Mid(sLine, pNum, EndParam - pNum)) <> "" Then
                    ReDim Preserve Param(UBound(Param) + 1) As String
                    Param(UBound(Param)) = Trim(Mid(sLine, pNum, EndParam - pNum))
                End If
            End If
            pNum = EndParam
        Next pNum
        If UBound(Param) < 3 Then ReDim Preserve Param(3) As String
        Select Case LCase(Param(1))
        Case "o", "open"
            If Param(2) <> "" Then
                MpqFile = Param(2)
                If Param(3) <> "" And FileExists(MpqFile) = False Then
                    DefaultMaxFiles = Param(3)
                End If
                If FileExists(MpqFile) Then
                    AddScriptOutput "Opened " + MpqFile + vbCrLf
                Else
                    AddScriptOutput "Created new " + MpqFile + vbCrLf
                End If
                NewPath = CurPath
            Else
                AddScriptOutput "Required parameter missing" + vbCrLf
            End If
        Case "n", "new"
            If Param(2) <> "" Then
                MpqFile = Param(2)
                If Param(3) <> "" Then
                    DefaultMaxFiles = Param(3)
                End If
                ScriptNewFile = True
                AddScriptOutput "Created new " + MpqFile + vbCrLf
                NewPath = CurPath
            Else
                AddScriptOutput "Required parameter missing" + vbCrLf
            End If
        Case "c", "close"
            If MpqFile <> "" Then
                If LCase(CD.FileName) = LCase(FullPath(NewPath, MpqFile)) Then MpqEx.Timer1.Enabled = True
                AddScriptOutput "Closed " + MpqFile + vbCrLf
                MpqFile = ""
            Else
                AddScriptOutput "No archive open" + vbCrLf
            End If
        Case "p", "pause"
            AddScriptOutput "Pause not supported" + vbCrLf
        Case "a", "add"
            If MpqFile <> "" Then
                cType = 0
                Rswitch = False
                fCount = 0
                Files = ""
                fEndLine = 0
                fLine = ""
                dwFlags = MAFA_REPLACE_EXISTING
                If GlobalEncrypt Then dwFlags = dwFlags Or MAFA_ENCRYPT
                For pNum = 3 To UBound(Param)
                    If LCase(Param(pNum)) = "/wav" Then
                        cType = 2
                        dwFlags = dwFlags Or MAFA_COMPRESS
                    ElseIf LCase(Param(pNum)) = "/c" And cType < 2 Then
                        cType = 1
                        dwFlags = dwFlags Or MAFA_COMPRESS
                    ElseIf LCase(Param(pNum)) = "/auto" And cType < 1 Then
                        cType = -1
                    ElseIf LCase(Param(pNum)) = "/r" Then
                        Rswitch = True
                    End If
                Next pNum
                If Left(Param(3), 1) = "/" Or Param(3) = "" Then
                    If InStr(Param(2), "*") <> 0 Or InStr(Param(2), "?") <> 0 Then
                        Param(3) = ""
                    Else
                        Param(3) = Param(2)
                    End If
                End If
                If Left(Param(2), 1) <> "/" And Param(2) <> "" Then
                    If InStr(Param(2), "\") > 0 Then
                        For pNum = 1 To Len(Param(2))
                            If InStr(pNum, Param(2), "\") > 0 Then
                                pNum = InStr(pNum, Param(2), "\")
                                Files = Left(Param(2), pNum)
                            End If
                        Next pNum
                    End If
                    If ScriptNewFile = True Then
                        If FileExists(FullPath(NewPath, MpqFile)) Then Kill FullPath(NewPath, MpqFile)
                        ScriptNewFile = False
                    End If
                    Files = DirEx(Files, Mid(Param(2), Len(Files) + 1), 6, Rswitch)
                    hMPQ = mOpenMpq(FullPath(NewPath, MpqFile))
                    If hMPQ = 0 Then
                        AddScriptOutput "Can't create archive " + MpqFile + vbCrLf
                        GoTo CommandError
                    End If
                    For pNum = 1 To Len(Files)
                        fEndLine = InStr(pNum, Files, vbCrLf)
                        fLine = Mid(Files, pNum, fEndLine - pNum)
                        If pNum > 1 Then
                            AddScriptOutput "Line " + CStr(lNum) + ": "
                        End If
                        If cType = 0 Then
                            AddScriptOutput "Adding " + fLine + "..."
                        ElseIf cType = 1 Then
                            AddScriptOutput "Adding compressed " + fLine + "..."
                        ElseIf cType = 2 Then
                            AddScriptOutput "Adding compressed WAV " + fLine + "..."
                        ElseIf cType = -1 Then
                            AddScriptOutput "Adding " + fLine + " (compression auto-select)..."
                        End If
                        If InStr(Param(2), "*") <> 0 Or InStr(Param(2), "?") <> 0 Then
                            If Right(Param(3), 1) <> "\" And Param(3) <> "" Then Param(3) = Param(3) + "\"
                            If cType = 2 Then
                                MpqAddWaveToArchive hMPQ, FullPath(CurPath, fLine), Param(3) + fLine, dwFlags, 0
                            ElseIf cType = -1 Then
                                mAddAutoFile hMPQ, FullPath(CurPath, fLine), Param(3) + fLine
                            ElseIf cType = 1 Then
                                If DefaultCompress = MAFA_COMPRESS_DEFLATE Then
                                    MpqAddFileToArchiveEx hMPQ, FullPath(CurPath, fLine), Param(3) + fLine, dwFlags, DefaultCompress, DefaultCompressLevel
                                Else
                                    MpqAddFileToArchiveEx hMPQ, FullPath(CurPath, fLine), Param(3) + fLine, dwFlags, DefaultCompress, 0
                                End If
                            Else
                                MpqAddFileToArchiveEx hMPQ, FullPath(CurPath, fLine), Param(3) + fLine, dwFlags, 0, 0
                            End If
                        Else
                            If cType = 2 Then
                                MpqAddWaveToArchive hMPQ, FullPath(CurPath, fLine), Param(3), dwFlags, 0
                            ElseIf cType = -1 Then
                                mAddAutoFile hMPQ, FullPath(CurPath, fLine), Param(3)
                            ElseIf cType = 1 Then
                                If DefaultCompress = MAFA_COMPRESS_DEFLATE Then
                                    MpqAddFileToArchiveEx hMPQ, FullPath(CurPath, fLine), Param(3), dwFlags, DefaultCompress, DefaultCompressLevel
                                Else
                                    MpqAddFileToArchiveEx hMPQ, FullPath(CurPath, fLine), Param(3), dwFlags, DefaultCompress, 0
                                End If
                            Else
                                MpqAddFileToArchiveEx hMPQ, FullPath(CurPath, fLine), Param(3), dwFlags, 0, 0
                            End If
                        End If
                        AddScriptOutput " Done" + vbCrLf
                        SendMessageA ScriptOut.oText.hWnd, WM_PAINT, 0, &O0
                        fCount = fCount + 1
                        pNum = fEndLine + 1
                    Next pNum
                    MpqCloseUpdatedArchive hMPQ, 0
                    If fCount > 1 Then
                        AddScriptOutput "Line " + CStr(lNum) + ":  " + CStr(fCount) + " files of " + Param(2) + " added" + vbCrLf
                    End If
                Else
                    AddScriptOutput " Required parameter missing" + vbCrLf
                End If
            Else
                AddScriptOutput "No archive open" + vbCrLf
            End If
        Case "e", "extract"
            If MpqFile <> "" Then
                If InStr(Param(2), "*") = 0 And InStr(Param(2), "?") = 0 Then AddScriptOutput "Extracting " + Param(2) + "..."
                cType = 0
                For pNum = 3 To UBound(Param)
                    If LCase(Param(pNum)) = "/fp" Then
                        cType = 1
                        Exit For
                    End If
                Next pNum
                If Left(Param(3), 1) = "/" Then Param(3) = ""
                If Param(3) = "" Then Param(3) = "."
                If Left(Param(2), 1) <> "/" And Param(2) <> "" Then
                    If InStr(Param(2), "*") <> 0 Or InStr(Param(2), "?") <> 0 Then
                        Files = MpqDir(FullPath(NewPath, MpqFile), Param(2))
                        If SFileOpenArchive(FullPath(NewPath, MpqFile), 0, 0, hMPQ) = 0 Then
                            AddScriptOutput "Can't open archive " + FullPath(NewPath, MpqFile) + vbCrLf
                            GoTo CommandError
                        End If
                        For pNum = 1 To Len(Files)
                            fEndLine = InStr(pNum, Files, vbCrLf)
                            fLine = Mid(Files, pNum, fEndLine - pNum)
                            If pNum > 1 Then
                                AddScriptOutput "Line " + CStr(lNum) + ": "
                            End If
                            AddScriptOutput "Extracting " + fLine + "..."
                            sGetFile hMPQ, fLine, FullPath(CurPath, Param(3)), cType
                            AddScriptOutput " Done" + vbCrLf
                            
                            fCount = fCount + 1
                            pNum = fEndLine + 1
                        Next pNum
                        SFileCloseArchive hMPQ
                        If fCount > 1 Then
                            AddScriptOutput "Line " + CStr(lNum) + ":  " + CStr(fCount) + " files of " + Param(2) + " extracted" + vbCrLf
                        End If
                    Else
                        If SFileOpenArchive(FullPath(NewPath, MpqFile), 0, 0, hMPQ) = 0 Then
                            AddScriptOutput "Can't open archive " + FullPath(NewPath, MpqFile) + vbCrLf
                            GoTo CommandError
                        End If
                        sGetFile hMPQ, Param(2), FullPath(CurPath, Param(3)), cType
                        SFileCloseArchive hMPQ
                        AddScriptOutput " Done" + vbCrLf
                    End If
                Else
                    AddScriptOutput " Required parameter missing" + vbCrLf
                End If
            Else
                AddScriptOutput "No archive open" + vbCrLf
            End If
        Case "r", "ren", "rename"
            If MpqFile <> "" Then
                If InStr(Param(2), "*") = 0 And InStr(Param(2), "?") = 0 Then AddScriptOutput "Renaming " + Param(2) + " => " + Param(3) + "..."
                If Param(2) <> "" And Param(3) <> "" Then
                    If InStr(Param(2), "*") <> 0 Or InStr(Param(2), "?") <> 0 Then
                        If InStr(Param(3), "*") <> 0 Or InStr(Param(3), "?") <> 0 Then
                            Files = MpqDir(FullPath(NewPath, MpqFile), Param(2))
                            hMPQ = mOpenMpq(FullPath(NewPath, MpqFile))
                            If hMPQ Then
                                For pNum = 1 To Len(Files)
                                    fEndLine = InStr(pNum, Files, vbCrLf)
                                    fLine = Mid(Files, pNum, fEndLine - pNum)
                                    If pNum > 1 Then
                                        AddScriptOutput "Line " + CStr(lNum) + ": "
                                    End If
                                    fLine2 = RenameWithFilter(fLine, Param(2), Param(3))
                                    AddScriptOutput "Renaming " + fLine + " => " + fLine2 + "..."
                                    If SFileOpenFileEx(hMPQ, fLine2, 0, hFile) Then
                                        SFileCloseFile hFile
                                        MpqDeleteFile hMPQ, fLine2
                                        MpqRenameFile hMPQ, fLine, fLine2
                                    Else
                                        MpqRenameFile hMPQ, fLine, fLine2
                                    End If
                                    AddScriptOutput " Done" + vbCrLf
                                    fCount = fCount + 1
                                    pNum = fEndLine + 1
                                Next pNum
                                MpqCloseUpdatedArchive hMPQ, 0
                            End If
                            If fCount > 1 Then
                                AddScriptOutput "Line " + CStr(lNum) + ":  " + CStr(fCount) + " files of " + Param(2) + " renamed" + vbCrLf
                            End If
                        Else
                        AddScriptOutput "You must use wildcards with new name" + vbCrLf
                        End If
                    Else
                        hMPQ = mOpenMpq(FullPath(NewPath, MpqFile))
                        If hMPQ Then
                            If SFileOpenFileEx(hMPQ, Param(3), 0, hFile) Then
                                SFileCloseFile hFile
                                MpqDeleteFile hMPQ, Param(3)
                                MpqRenameFile hMPQ, Param(2), Param(3)
                            Else
                                MpqRenameFile hMPQ, Param(2), Param(3)
                            End If
                            MpqCloseUpdatedArchive hMPQ, 0
                        End If
                        AddScriptOutput " Done" + vbCrLf
                    End If
                Else
                    AddScriptOutput " Required parameter missing" + vbCrLf
                End If
            Else
                AddScriptOutput "No archive open" + vbCrLf
            End If
        Case "m", "move"
            If MpqFile <> "" Then
                For pNum = 1 To Len(Param(2))
                    If InStr(bNum, Param(2), "\") Then
                        bNum = InStr(bNum, Param(2), "\")
                    Else
                        Exit For
                    End If
                Next pNum
                fLineTitle = Mid(Param(2), bNum)
                If Right(Param(3), 1) <> "\" And Param(3) <> "" Then Param(3) = Param(3) + "\"
                Param(3) = Param(3) + fLineTitle
                If InStr(Param(2), "*") = 0 And InStr(Param(2), "?") = 0 Then AddScriptOutput "Moving " + Param(2) + " => " + Param(3) + "..."
                If (Left(Param(2), 1) <> "/" And Param(2) <> "") And (Left(Param(3), 1) <> "/") Then
                    If InStr(Param(2), "*") <> 0 Or InStr(Param(2), "?") <> 0 Then
                        Files = MpqDir(FullPath(NewPath, MpqFile), Param(2))
                        hMPQ = mOpenMpq(FullPath(NewPath, MpqFile))
                        If hMPQ Then
                            For pNum = 1 To Len(Files)
                                fEndLine = InStr(pNum, Files, vbCrLf)
                                fLine = Mid(Files, pNum, fEndLine - pNum)
                                If pNum > 1 Then
                                    AddScriptOutput "Line " + CStr(lNum) + ": "
                                End If
                                fLine2 = RenameWithFilter(fLine, Param(2), Param(3))
                                AddScriptOutput "Moving " + fLine + " => " + fLine2 + "..."
                                If SFileOpenFileEx(hMPQ, fLine2, 0, hFile) Then
                                    SFileCloseFile hFile
                                    MpqDeleteFile hMPQ, fLine2
                                    MpqRenameFile hMPQ, fLine, fLine2
                                Else
                                    MpqRenameFile hMPQ, fLine, fLine2
                                End If
                                AddScriptOutput " Done" + vbCrLf
                                fCount = fCount + 1
                                pNum = fEndLine + 1
                            Next pNum
                            MpqCloseUpdatedArchive hMPQ, 0
                        End If
                        If fCount > 1 Then
                            AddScriptOutput "Line " + CStr(lNum) + ":  " + CStr(fCount) + " files of " + Param(2) + " moved" + vbCrLf
                        End If
                    Else
                        hMPQ = mOpenMpq(FullPath(NewPath, MpqFile))
                        If hMPQ Then
                            If SFileOpenFileEx(hMPQ, Param(3), 0, hFile) Then
                                SFileCloseFile hFile
                                MpqDeleteFile hMPQ, Param(3)
                                MpqRenameFile hMPQ, Param(2), Param(3)
                            Else
                                MpqRenameFile hMPQ, Param(2), Param(3)
                            End If
                            MpqCloseUpdatedArchive hMPQ, 0
                        End If
                        AddScriptOutput " Done" + vbCrLf
                    End If
                Else
                    AddScriptOutput " Required parameter missing" + vbCrLf
                End If
            Else
                AddScriptOutput "No archive open" + vbCrLf
            End If
        Case "d", "del", "delete"
            If MpqFile <> "" Then
                If InStr(Param(2), "*") = 0 And InStr(Param(2), "?") = 0 Then AddScriptOutput "Deleting " + Param(2) + "..."
                If Left(Param(2), 1) <> "/" And Param(2) <> "" Then
                    If InStr(Param(2), "*") <> 0 Or InStr(Param(2), "?") <> 0 Then
                        Files = MpqDir(FullPath(NewPath, MpqFile), Param(2))
                        hMPQ = mOpenMpq(FullPath(NewPath, MpqFile))
                        If hMPQ Then
                            For pNum = 1 To Len(Files)
                                fEndLine = InStr(pNum, Files, vbCrLf)
                                fLine = Mid(Files, pNum, fEndLine - pNum)
                                If pNum > 1 Then
                                    AddScriptOutput "Line " + CStr(lNum) + ": "
                                End If
                                AddScriptOutput "Deleting " + fLine + "..."
                                MpqDeleteFile hMPQ, fLine
                                AddScriptOutput " Done" + vbCrLf
                                fCount = fCount + 1
                                pNum = fEndLine + 1
                            Next pNum
                            MpqCloseUpdatedArchive hMPQ, 0
                        End If
                        If fCount > 1 Then
                            AddScriptOutput "Line " + CStr(lNum) + ":  " + CStr(fCount) + " files of " + Param(2) + " deleted" + vbCrLf
                        End If
                    Else
                        hMPQ = mOpenMpq(FullPath(NewPath, MpqFile))
                        If hMPQ Then
                            MpqDeleteFile hMPQ, Param(2)
                            MpqCloseUpdatedArchive hMPQ, 0
                        End If
                        AddScriptOutput " Done" + vbCrLf
                    End If
                Else
                    AddScriptOutput " Required parameter missing" + vbCrLf
                End If
            Else
                AddScriptOutput "No archive open" + vbCrLf
            End If
        Case "f", "flush", "compact"
            If MpqFile <> "" Then
                AddScriptOutput "Flushing " + MpqFile + "..."
                hMPQ = mOpenMpq(FullPath(NewPath, MpqFile))
                If hMPQ Then
                    MpqCompactArchive hMPQ
                    MpqCloseUpdatedArchive hMPQ, 0
                End If
                AddScriptOutput " Done" + vbCrLf
            Else
                AddScriptOutput "No archive open" + vbCrLf
            End If
        Case "l", "list"
            If MpqFile <> "" Then
                If Param(2) <> "" Then
                    AddScriptOutput "Creating list..."
                    If (InStr(Param(2), "*") <> 0 Or InStr(Param(2), "?") <> 0) And Param(3) <> "" Then
                        Files = MpqDir(FullPath(NewPath, MpqFile), Param(2))
                        Param(2) = Param(3)
                    Else
                        Files = MpqDir(FullPath(NewPath, MpqFile), "*")
                    End If
                    fNum = FreeFile
                    Open FullPath(CurPath, Param(2)) For Binary As #fNum
                    Put #fNum, 1, Files
                    Close #fNum
                    AddScriptOutput " Done" + vbCrLf
                Else
                    AddScriptOutput " Required parameter missing" + vbCrLf
                End If
            Else
                AddScriptOutput "No archive open" + vbCrLf
            End If
        Case "s", "script"
            AddScriptOutput "Running script " + Param(2) + "..." + vbCrLf + vbCrLf
            If Param(2) <> "" Then
                RunScript FullPath(CurPath, Param(2))
            Else
                AddScriptOutput " Required parameter missing" + vbCrLf
            End If
            AddScriptOutput vbCrLf + "Continuing with previous script..." + vbCrLf
        Case "x", "exit", "quit"
            Unload MpqEx
        Case Else
            If Left(Param(1), 1) <> ";" Then
                If LCase(Param(1)) = "cd" Or LCase(Param(1)) = "chdir" Then
                    On Error Resume Next
                    ChDir Param(2)
                    On Error GoTo 0
                    CurPath = CurDir
                    AddScriptOutput "Current directory is " + CurPath + vbCrLf
                ElseIf Left(LCase(Param(1)), 3) = "cd." Or Left(LCase(Param(1)), 3) = "cd\" Then
                    On Error Resume Next
                    ChDir Mid(Param(1), 3)
                    On Error GoTo 0
                    CurPath = CurDir
                    AddScriptOutput "Current directory is " + CurPath + vbCrLf
                ElseIf Left(LCase(Param(1)), 6) = "chdir." Or Left(LCase(Param(1)), 6) = "chdir\" Then
                    On Error Resume Next
                    ChDir Mid(Param(1), 6)
                    On Error GoTo 0
                    CurPath = CurDir
                    AddScriptOutput "Current directory is " + CurPath + vbCrLf
                ElseIf Mid(Param(1), 2, 1) = ":" And (Len(Param(1)) = 2 Or Right(Param(1), 1) = "\") Then
                    On Error Resume Next
                    ChDrive Left(Param(1), 2)
                    On Error GoTo 0
                    CurPath = CurDir
                    AddScriptOutput "Current directory is " + CurPath + vbCrLf
                Else
                    AddScriptOutput "Running command " + sLine + "..."
                    Shell "command.com /c " + sLine, 1
                    AddScriptOutput " Done" + vbCrLf
                End If
            Else
                AddScriptOutput "Comment  " + sLine + vbCrLf
            End If
        End Select
    End If
CommandError:
    lNum = lNum + 1
    bNum = EndLine + 1
Next bNum
DefaultMaxFiles = OldDefaultMaxFiles
If Mid(OldPath, 2, 1) = ":" Then ChDrive Left(OldPath, 1)
ChDir OldPath
End Sub
Function FindMpqHeader(MpqFile As String) As Long
    If FileExists(MpqFile) = False Then
        FindMpqHeader = -1
        Exit Function
    End If
    Dim hFile
    hFile = FreeFile
    Open MpqFile For Binary As #hFile
    Dim FileLen As Long
    FileLen = LOF(hFile)
    Dim pbuf As String
    pbuf = String(32, Chr(0))
    Dim i As Long
    For i = 0 To FileLen - 1 Step 512
        Get #hFile, 1 + i, pbuf
        If Left(pbuf, 4) = "MPQ" + Chr(26) Or Left(pbuf, 4) = "BN3" + Chr(26) Then
            ' Storm no longer does this, so this shouldn't either
            'FileLen = FileLen - i
            'If JBytes(pbuf, 9, 4) = FileLen
            '    FileMpqHeader = i
            '    Close #hFile
            '    Exit Function
            'Else
            '    FileLen = FileLen + i
            'End If
            FindMpqHeader = i
            Close #hFile
            Exit Function
        End If
    Next i
    FindMpqHeader = -1
    Close #hFile
End Function
Function GetNumMpqFiles(MpqFile As String) As Long
Dim fNum As Long, Text As String, MpqHeader As Long
fNum = FreeFile
Text = String(4, Chr(0))
MpqHeader = FindMpqHeader(MpqFile)
If MpqHeader > -1 Then
    Open MpqFile For Binary As #fNum
    Get #fNum, MpqHeader + 29, GetNumMpqFiles
    Close #fNum
End If
End Function

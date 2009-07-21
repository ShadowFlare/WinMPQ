Attribute VB_Name = "CwadLib"
Option Explicit

Public Const CWAD_INFO_NUM_FILES       As Long = &H03 ' Number of files in CWAD
Public Const CWAD_INFO_TYPE            As Long = &H04 ' Is HANDLE a file or a CWAD?
Public Const CWAD_INFO_SIZE            As Long = &H05 ' Size of CWAD or uncompressed file
Public Const CWAD_INFO_COMPRESSED_SIZE As Long = &H06 ' Size of compressed file
Public Const CWAD_INFO_FLAGS           As Long = &H07 ' File flags (compressed, etc.)
Public Const CWAD_INFO_PARENT          As Long = &H08 ' Handle of CWAD that file is in
Public Const CWAD_INFO_POSITION        As Long = &H09 ' Position of file pointer in files
Public Const CWAD_INFO_PRIORITY        As Long = &H0B ' Priority of open CWAD

Public Const CWAD_TYPE_CWAD As Long = &H01
Public Const CWAD_TYPE_FILE As Long = &H02

Public Const CWAD_SEARCH_CURRENT_ONLY As Long = &H00 ' Used with CWadOpenFile; only the archive with the handle specified will be searched for the file
Public Const CWAD_SEARCH_ALL_OPEN     As Long = &H01 ' CWadOpenFile will look through all open archives for the file

Declare Function CWadOpenArchive Lib "CwadLib.dll" (ByVal lpFileName As String, ByVal dwPriority As Long, ByRef hCWAD As Long) As Boolean
Declare Function CWadCloseArchive Lib "CwadLib.dll" (ByVal hCWAD As Long) As Boolean
Declare Function CWadListFiles Lib "CwadLib.dll" (ByVal hCWAD As Long, ByVal lpBuffer As String, ByVal dwBufferLength As Long) As Long ' Returns required buffer size.  Strings are in multi string form. (null-terminated strings with an extra null after the last string)
Declare Function CWadOpenFile Lib "CwadLib.dll" (ByVal hCWAD As Long, ByVal lpFileName As String, ByVal dwSearchScope As Long, ByRef hFile As Long) As Boolean
Declare Function CWadCloseFile Lib "CwadLib.dll" (ByVal hFile As Long) As Boolean
Declare Function CWadGetFileSize Lib "CwadLib.dll" (ByVal hFile As Long) As Long
Declare Function CWadGetFileInfo Lib "CwadLib.dll" (ByVal hFile As Long, ByVal dwInfoType As Long) As Long
Declare Function CWadSetFilePointer Lib "CwadLib.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal dwMoveMethod As Long) As Long
Declare Function CWadReadFile Lib "CwadLib.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long) As Boolean
Declare Function CWadSetArchivePriority Lib "CwadLib.dll" (ByVal hCWAD As Long, ByVal dwPriority As Long) As Boolean
Declare Function CWadFindHeader Lib "CwadLib.dll" (ByVal hFile As Long) As Long

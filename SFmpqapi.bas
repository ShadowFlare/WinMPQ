Attribute VB_Name = "SFmpqapi"
Option Explicit

'  ShadowFlare MPQ API Library. (c) ShadowFlare Software 2002

'  All functions below are actual functions that are part of this
'  library and do not need any additional dll files.  It does not
'  even require Storm to be able to decompress or compress files.

'  This library emulates the interface of Lmpqapi and Storm MPQ
'  functions, so it may be used as a replacement for them in
'  MPQ extractors/archivers without even needing to recompile
'  the program that uses Lmpqapi or Storm.  It has a few features
'  not included in Lmpqapi and Storm, such as extra flags for some
'  functions, setting the locale ID of existing files, and adding
'  files without having to write them somewhere else first.  Also,
'  MPQ handles used by functions prefixed with "SFile" and "Mpq"
'  can be used interchangably; all functions use the same type
'  of MPQ handles.  You cannot, however, use handles from this
'  library with storm or lmpqapi or vice-versa.  Doing so will
'  most likely result in a crash.

'  Revision History:
'  06/12/2002 1.07 (ShadowFlare)
'  - No longer requires Storm.dll to compress or decompress
'    Warcraft III files
'  - Added SFileListFiles for getting names and information
'    about all of the files in an archive
'  - Fixed a bug with renaming and deleting files
'  - Fixed a bug with adding wave compressed files with
'    low compression setting
'  - Added a check in MpqOpenArchiveForUpdate for proper
'    dwMaximumFilesInArchive values (should be a number that
'    is a power of 2).  If it is not a proper value, it will
'    be rounded up to the next higher power of 2

'  05/09/2002 1.06 (ShadowFlare)
'  - Compresses files without Storm.dll!
'  - If Warcraft III is installed, this library will be able to
'    find Storm.dll on its own. (Storm.dll is needed to
'    decompress Warcraft III files)
'  - Fixed a bug where an embedded archive and the file that
'    contains it would be corrupted if the archive was modified
'  - Able to open all .w3m maps now

'  29/06/2002 1.05 (ShadowFlare)
'  - Supports decompressing files from Warcraft III MPQ archives
'    if using Storm.dll from Warcraft III
'  - Added MpqAddFileToArchiveEx and MpqAddFileFromBufferEx for
'    using extra compression types

'  29/05/2002 1.04 (ShadowFlare)
'  - Files can be compressed now!
'  - Fixed a bug in SFileReadFile when reading data not aligned
'    to the block size
'  - Optimized some of SFileReadFile's code.  It can read files
'    faster now
'  - SFile functions may now be used to access files not in mpq
'    archives as you can with the real storm functions
'  - MpqCompactArchive will no longer corrupt files with the
'    MODCRYPTKEY flag as long as the file is either compressed,
'    listed in "(listfile)", is "(listfile)", or is located in
'    the same place in the compacted archive; so it is safe
'    enough to use it on almost any archive
'  - Added MpqAddWaveFromBuffer
'  - Better handling of archives with no files
'  - Fixed compression with COMPRESS2 flag

'  15/05/2002 1.03 (ShadowFlare)
'  - Supports adding files with the compression attribute (does
'    not actually compress files).  Now archives created with
'    this dll can have files added to them through lmpqapi
'    without causing staredit to crash
'  - SFileGetBasePath and SFileSetBasePath work more like their
'    Storm equivalents now
'  - Implemented MpqCompactArchive, but it is not finished yet.
'    In its current state, I would recommend against using it
'    on archives that contain files with the MODCRYPTKEY flag,
'    since it will corrupt any files with that flag
'  - Added SFMpqGetVersionString2 which may be used in Visual
'    Basic to get the version string

'  07/05/2002 1.02 (ShadowFlare)
'  - SFileReadFile no longer passes the lpOverlapped parameter it
'    receives to ReadFile.  This is what was causing the function
'    to fail when used in Visual Basic
'  - Added support for more Storm MPQ functions
'  - GetLastError may now be used to get information about why a
'    function failed

'  01/05/2002 1.01 (ShadowFlare)
'  - Added ordinals for Storm MPQ functions
'  - Fixed MPQ searching functionality of SFileOpenFileEx
'  - Added a check for whether a valid handle is given when
'    SFileCloseArchive is called
'  - Fixed functionality of SFileSetArchivePriority when multiple
'    files are open
'  - File renaming works for all filenames now
'  - SFileReadFile no longer reallocates the buffer for each block
'    that is decompressed.  This should make SFileReadFile at least
'    a little faster

'  30/04/2002 1.00 (ShadowFlare)
'  - First version.
'  - Compression not yet supported
'  - Does not use SetLastError yet, so GetLastError will not return any
'    errors that have to do with this library
'  - MpqCompactArchive not implemented

'  This library is freeware, you can do anything you want with it but with
'  one exception.  If you use it in your program, you must specify this fact
'  in Help|About box or in similar way.  You can obtain version string using
'  SFMpqGetVersionString call.

'  THIS LIBRARY IS DISTRIBUTED "AS IS".  NO WARRANTY OF ANY KIND IS EXPRESSED
'  OR IMPLIED. YOU USE AT YOUR OWN RISK. THE AUTHOR WILL NOT BE LIABLE FOR
'  DATA LOSS, DAMAGES, LOSS OF PROFITS OR ANY OTHER KIND OF LOSS WHILE USING
'  OR MISUSING THIS SOFTWARE.

'  Any comments or suggestions are accepted at blakflare@hotmail.com (ShadowFlare)

Type SFMPQVERSION
    Major As Integer
    Minor As Integer
    Revision As Integer
    Subrevision As Integer
End Type

' MpqInitialize does nothing.  It is only provided for
' compatibility with MPQ archivers that use lmpqapi.
Declare Function MpqInitialize Lib "SFmpq.dll" () As Boolean

Declare Function MpqGetVersionString Lib "SFmpq.dll" () As String
Declare Function MpqGetVersion Lib "SFmpq.dll" () As Single

Declare Sub SFMpqDestroy Lib "SFmpq.dll" () ' This no longer needs to be called.  It is only provided for compatibility with older versions

' SFMpqGetVersionString2's return value is the required length of the buffer plus
' the terminating null, so use SFMpqGetVersionString2(ByVal 0&, 0) to get the length.
Declare Function SFMpqGetVersionString Lib "SFmpq.dll" () As String
Declare Function SFMpqGetVersionString2 Lib "SFmpq.dll" (ByVal lpBuffer As String, ByVal dwBufferLength As Long) As Long
Declare Function SFMpqGetVersion Lib "SFmpq.dll" () As SFMPQVERSION

' General error codes
Public Const MPQ_ERROR_MPQ_INVALID As Long = &H85200065
Public Const MPQ_ERROR_FILE_NOT_FOUND As Long = &H85200066
Public Const MPQ_ERROR_DISK_FULL As Long = &H85200068 'Physical write file to MPQ failed. Not sure of exact meaning
Public Const MPQ_ERROR_HASH_TABLE_FULL As Long = &H85200069
Public Const MPQ_ERROR_ALREADY_EXISTS As Long = &H8520006A
Public Const MPQ_ERROR_BAD_OPEN_MODE As Long = &H8520006C 'When MOAU_READ_ONLY is used without MOAU_OPEN_EXISTING

Public Const MPQ_ERROR_COMPACT_ERROR As Long = &H85300001

' MpqOpenArchiveForUpdate flags
Public Const MOAU_CREATE_NEW As Long = &H0
Public Const MOAU_CREATE_ALWAYS As Long = &H8 'Was wrongly named MOAU_CREATE_NEW
Public Const MOAU_OPEN_EXISTING As Long = &H4
Public Const MOAU_OPEN_ALWAYS As Long = &H20
Public Const MOAU_READ_ONLY As Long = &H10 'Must be used with MOAU_OPEN_EXISTING
Public Const MOAU_MAINTAIN_LISTFILE As Long = &H1

' MpqOpenArchiveForUpdateEx constants
Public Const DEFAULT_BLOCK_SIZE As Long = 3 ' 512 << number = block size
Public Const USE_DEFAULT_BLOCK_SIZE As Long = &HFFFF ' Use default block size that is defined internally

' MpqAddFileToArchive flags
Public Const MAFA_EXISTS As Long = &H80000000 'Will be added if not present
Public Const MAFA_UNKNOWN40000000 As Long = &H40000000
Public Const MAFA_MODCRYPTKEY As Long = &H20000
Public Const MAFA_ENCRYPT As Long = &H10000
Public Const MAFA_COMPRESS As Long = &H200
Public Const MAFA_COMPRESS2 As Long = &H100
Public Const MAFA_REPLACE_EXISTING As Long = &H1

' MpqAddFileToArchiveEx compression flags
Public Const MAFA_COMPRESS_STANDARD As Long = &H8 'Standard PKWare DCL compression
Public Const MAFA_COMPRESS_DEFLATE  As Long = &H2 'ZLib's deflate compression
Public Const MAFA_COMPRESS_WAVE    As Long = &H81 'Standard wave compression
Public Const MAFA_COMPRESS_WAVE2   As Long = &H41 'Unused wave compression

' Flags for individual compression types used for wave compression
Public Const MAFA_COMPRESS_WAVECOMP1 As Long = &H80 'Main compressor for standard wave compression
Public Const MAFA_COMPRESS_WAVECOMP2 As Long = &H40 'Main compressor for unused wave compression
Public Const MAFA_COMPRESS_WAVECOMP3  As Long = &H1 'Secondary compressor for wave compression

' ZLib deflate compression level constants (used with MpqAddFileToArchiveEx and MpqAddFileFromBufferEx)
Public Const Z_NO_COMPRESSION         As Long = 0
Public Const Z_BEST_SPEED             As Long = 1
Public Const Z_BEST_COMPRESSION       As Long = 9
Public Const Z_DEFAULT_COMPRESSION  As Long = (-1)

' MpqAddWAVToArchive quality flags
Public Const MAWA_QUALITY_HIGH As Long = 1
Public Const MAWA_QUALITY_MEDIUM As Long = 0
Public Const MAWA_QUALITY_LOW As Long = 2

' SFileGetFileInfo flags
Public Const SFILE_INFO_BLOCK_SIZE As Long = &H1 'Block size in MPQ
Public Const SFILE_INFO_HASH_TABLE_SIZE As Long = &H2 'Hash table size in MPQ
Public Const SFILE_INFO_NUM_FILES As Long = &H3 'Number of files in MPQ
Public Const SFILE_INFO_TYPE As Long = &H4 'Is Long a file or an MPQ?
Public Const SFILE_INFO_SIZE As Long = &H5 'Size of MPQ or uncompressed file
Public Const SFILE_INFO_COMPRESSED_SIZE As Long = &H6 'Size of compressed file
Public Const SFILE_INFO_FLAGS As Long = &H7 'File flags (compressed, etc.), file attributes if a file not in an archive
Public Const SFILE_INFO_PARENT As Long = &H8 'Handle of MPQ that file is in
Public Const SFILE_INFO_POSITION As Long = &H9 'Position of file pointer in files
Public Const SFILE_INFO_LOCALEID As Long = &HA 'Locale ID of file in MPQ
Public Const SFILE_INFO_PRIORITY As Long = &HB 'Priority of open MPQ
Public Const SFILE_INFO_HASH_INDEX As Long = &HC 'Hash index of file in MPQ

' SFileListFiles flags
Public Const SFILE_LIST_MEMORY_LIST  As Long = &H1 ' Specifies that lpFilelists is a file list from memory, rather than being a list of file lists
Public Const SFILE_LIST_ONLY_KNOWN   As Long = &H2 ' Only list files that the function finds a name for
Public Const SFILE_LIST_ONLY_UNKNOWN As Long = &H4 ' Only list files that the function does not find a name for

Public Const SFILE_TYPE_MPQ As Long = &H1
Public Const SFILE_TYPE_FILE As Long = &H2

Public Const INVALID_HANDLE_VALUE As Long = -1

Public Const FILE_BEGIN   As Long = 0
Public Const FILE_CURRENT As Long = 1
Public Const FILE_END     As Long = 2

Public Const SFILE_OPEN_HARD_DISK_FILE As Long = &H0 'Open archive without regard to the drive type it resides on
Public Const SFILE_OPEN_CD_ROM_FILE As Long = &H1 'Open the archive only if it is on a CD-ROM
Public Const SFILE_OPEN_ALLOW_WRITE As Long = &H8000 'Open file with write access

Public Const SFILE_SEARCH_CURRENT_ONLY As Long = &H0 'Used with SFileOpenFileEx; only the archive with the handle specified will be searched for the file
Public Const SFILE_SEARCH_ALL_OPEN As Long = &H1 'SFileOpenFileEx will look through all open archives for the file

Type FILELISTENTRY
    dwFileExists As Long ' Nonzero if this entry is used
    lcLocale As Long ' Locale ID of file
    dwCompressedSize As Long ' Compressed size of file
    dwFullSize As Long ' Uncompressed size of file
    dwFlags As Long ' Flags for file
    szFileName(259) As Byte
End Type

' Storm functions implemented by this library
Declare Function SFileOpenArchive Lib "SFmpq.dll" (ByVal lpFileName As String, ByVal dwPriority As Long, ByVal dwFlags As Long, ByRef hMPQ As Long) As Boolean
Declare Function SFileCloseArchive Lib "SFmpq.dll" (ByVal hMPQ As Long) As Boolean
Declare Function SFileGetArchiveName Lib "SFmpq.dll" (ByVal hMPQ As Long, ByVal lpBuffer As String, ByVal dwBufferLength As Long) As Boolean
Declare Function SFileOpenFile Lib "SFmpq.dll" (ByVal lpFileName As String, ByRef hFile As Long) As Boolean
Declare Function SFileOpenFileEx Lib "SFmpq.dll" (ByVal hMPQ As Long, ByVal lpFileName As String, ByVal dwSearchScope As Long, ByRef hFile As Long) As Boolean
Declare Function SFileCloseFile Lib "SFmpq.dll" (ByVal hFile As Long) As Boolean
Declare Function SFileGetFileSize Lib "SFmpq.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Declare Function SFileGetFileArchive Lib "SFmpq.dll" (ByVal hFile As Long, ByRef hMPQ As Long) As Boolean
Declare Function SFileGetFileName Lib "SFmpq.dll" (ByVal hMPQ As Long, ByVal lpBuffer As String, ByVal dwBufferLength As Long) As Boolean
Declare Function SFileSetFilePointer Lib "SFmpq.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lplDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Declare Function SFileReadFile Lib "SFmpq.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Boolean
Declare Function SFileSetLocale Lib "SFmpq.dll" (ByVal nNewLocale As Long) As Long
Declare Function SFileGetBasePath Lib "SFmpq.dll" (ByVal lpBuffer As String, ByVal dwBufferLength As Long) As Boolean
Declare Function SFileSetBasePath Lib "SFmpq.dll" (ByVal lpNewBasePath As String) As Boolean

' Extra storm-related functions
Declare Function SFileGetFileInfo Lib "SFmpq.dll" (ByVal hFile As Long, ByVal dwInfoType As Long) As Long
Declare Function SFileSetArchivePriority Lib "SFmpq.dll" (ByVal hMPQ As Long, ByVal dwPriority As Long) As Boolean
Declare Function SFileFindMpqHeader Lib "SFmpq.dll" (ByVal hFile As Long) As Long
Declare Function SFileListFiles Lib "SFmpq.dll" (ByVal hMPQ As Long, ByVal lpFileLists As String, ByRef lpListBuffer As FILELISTENTRY, ByVal dwFlags As Long) As Boolean

' Archive editing functions implemented by this library
Declare Function MpqOpenArchiveForUpdate Lib "SFmpq.dll" (ByVal lpFileName As String, ByVal dwFlags As Long, ByVal dwMaximumFilesInArchive As Long) As Long
Declare Function MpqCloseUpdatedArchive Lib "SFmpq.dll" (ByVal hMPQ As Long, ByVal dwUnknown2 As Long) As Long
Declare Function MpqAddFileToArchive Lib "SFmpq.dll" (ByVal hMPQ As Long, ByVal lpSourceFileName As String, ByVal lpDestFileName As String, ByVal dwFlags As Long) As Boolean
Declare Function MpqAddWaveToArchive Lib "SFmpq.dll" (ByVal hMPQ As Long, ByVal lpSourceFileName As String, ByVal lpDestFileName As String, ByVal dwFlags As Long, ByVal dwQuality As Long) As Boolean
Declare Function MpqRenameFile Lib "SFmpq.dll" (ByVal hMPQ As Long, ByVal lpcOldFileName As String, ByVal lpcNewFileName As String) As Boolean
Declare Function MpqDeleteFile Lib "SFmpq.dll" (ByVal hMPQ As Long, ByVal lpFileName As String) As Boolean
Declare Function MpqCompactArchive Lib "SFmpq.dll" (ByVal hMPQ As Long) As Boolean

' Extra archive editing functions
Declare Function MpqOpenArchiveForUpdateEx Lib "SFmpq.dll" (ByVal lpFileName As String, ByVal dwFlags As Long, ByVal dwMaximumFilesInArchive As Long, ByVal dwBlockSize As Long) As Long
Declare Function MpqAddFileToArchiveEx Lib "SFmpq.dll" (ByVal hMPQ As Long, ByVal lpSourceFileName As String, ByVal lpDestFileName As String, ByVal dwFlags As Long, ByVal dwCompressionType As Long, ByVal dwCompressLevel As Long) As Boolean
Declare Function MpqAddFileFromBufferEx Lib "SFmpq.dll" (ByVal hMPQ As Long, ByRef lpBuffer As Any, ByVal dwLength As Long, ByVal lpFileName As String, ByVal dwFlags As Long, ByVal dwCompressionType As Long, ByVal dwCompressLevel As Long) As Boolean
Declare Function MpqAddFileFromBuffer Lib "SFmpq.dll" (ByVal hMPQ As Long, ByRef lpBuffer As Any, ByVal dwLength As Long, ByVal lpFileName As String, ByVal dwFlags As Long) As Boolean
Declare Function MpqAddWaveFromBuffer Lib "SFmpq.dll" (ByVal hMPQ As Long, ByRef lpBuffer As Any, ByVal dwLength As Long, ByVal lpFileName As String, ByVal dwFlags As Long, ByVal dwQuality As Long) As Boolean
Declare Function MpqSetFileLocale Lib "SFmpq.dll" (ByVal hMPQ As Long, ByVal lpFileName As String, ByVal nOldLocale As Long, ByVal nNewLocale As Long) As Boolean

' These functions do nothing.  They are only provided for
' compatibility with MPQ extractors that use storm.
Declare Function SFileDestroy Lib "SFmpq.dll" () As Boolean
Declare Sub StormDestroy Lib "SFmpq.dll" ()

' Returns 0 if the dll version is equal to the version your program was compiled
' with, 1 if the dll is newer, -1 if the dll is older.
Function SFMpqCompareVersion() As Long
    Dim ExeVersion As SFMPQVERSION, DllVersion As SFMPQVERSION
    With ExeVersion
        .Major = 1
        .Minor = 0
        .Revision = 7
        .Subrevision = 4
    End With
    DllVersion = SFMpqGetVersion()
    If DllVersion.Major > ExeVersion.Major Then
        SFMpqCompareVersion = 1
        Exit Function
    ElseIf DllVersion.Major < ExeVersion.Major Then
        SFMpqCompareVersion = -1
        Exit Function
    End If
    If DllVersion.Minor > ExeVersion.Minor Then
        SFMpqCompareVersion = 1
        Exit Function
    ElseIf DllVersion.Minor < ExeVersion.Minor Then
        SFMpqCompareVersion = -1
        Exit Function
    End If
    If DllVersion.Revision > ExeVersion.Revision Then
        SFMpqCompareVersion = 1
        Exit Function
    ElseIf DllVersion.Revision < ExeVersion.Revision Then
        SFMpqCompareVersion = -1
        Exit Function
    End If
    If DllVersion.Subrevision > ExeVersion.Subrevision Then
        SFMpqCompareVersion = 1
        Exit Function
    ElseIf DllVersion.Subrevision < ExeVersion.Subrevision Then
        SFMpqCompareVersion = -1
        Exit Function
    End If
    SFMpqCompareVersion = 0
End Function


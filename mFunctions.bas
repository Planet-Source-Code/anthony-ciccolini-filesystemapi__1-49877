Attribute VB_Name = "mFunctions"
Option Explicit

Private Const MAX_PATH = 260

Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCE_CONNECTED = &H1

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Public Const FILE_BEGIN = 0           'The starting point is zero or the beginning of the file. If FILE_BEGIN is specified, DistanceToMove is interpreted as an unsigned location for the new file pointer.
Public Const FILE_CURRENT = 1         'The current value of the file pointer is the starting point.
Public Const FILE_END = 2             'The current end-of-file position is the starting point.

Private Const CREATE_NEW = 1          'Creates a new file. The function fails if the specified file already exists.
Private Const CREATE_ALWAYS = 2       'Creates a new file. The function overwrites the file if it exists.
Private Const OPEN_EXISTING = 3       'Opens the file. The function fails if the file does not exist.
Private Const OPEN_ALWAYS = 4         'Opens the file, if it exists. If the file does not exist, the function creates the file as if dwCreationDistribution were CREATE_NEW.
Private Const TRUNCATE_EXISTING = 5   'Opens the file. Once opened, the file is truncated so that its size is zero bytes. The calling process must open the file with at least GENERIC_WRITE access. The function fails if the file does not exist.

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_TYPENAME = &H400

Private Const INVALID_HANDLE_VALUE = -1

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_ERROR_HANDLE = -12&


Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Type SHITEMID
  cb As Long
  abID As Byte
End Type

Private Type ITEMIDLIST
  mkid As SHITEMID
End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type NETRESOURCE
  dwScope As Long
  dwType As Long
  dwDisplayType As Long
  dwUsage As Long
  lpLocalName As Long
  lpRemoteName As Long
  lpComment As Long
  lpProvider As Long
End Type

Private Type VS_FIXEDFILEINFO
  dwSignature As Long
  dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
  dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
  dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
  dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
  dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
  dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
  dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
  dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
  dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
  dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
  dwFileFlagsMask As Long        '  = &h3F for version "0.42"
  dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
  dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
  dwFileType As Long             '  e.g. VFT_DRIVER
  dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
  dwFileDateMS As Long           '  e.g. 0
  dwFileDateLS As Long           '  e.g. 0
End Type

Private Type SHFILEINFO
  hIcon As Long                       '  out: icon
  iIcon As Long                       '  out: icon index
  dwAttributes As Long                '  out: SFGAO_ flags
  szDisplayName As String * MAX_PATH  '  out: display name (or path)
  szTypeName As String * 80           '  out: type name
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Public Enum DateType
  DateCreate = 1
  DateAccess = 2
  DateWrite = 3
End Enum

Public Enum DriveSize
  DriveSizeTotal = 1
  DriveSizeFree = 2
  DriveSizeAvailable = 3
End Enum

Public Enum DriveInfo
  DriveInfoFileSystem = 1
  DriveInfoVolumeName = 2
  DriveSerialNumber = 3
End Enum

' Kernel32:
Private Declare Function APICloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function APILoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function APIGetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function APIGetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function APISetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function APICopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function APIGetFileTime Lib "kernel32" Alias "GetFileTime" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function APICreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function APIFileTimeToSystemTime Lib "kernel32" Alias "FileTimeToSystemTime" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function APIFileTimeToLocalFileTime Lib "kernel32" Alias "FileTimeToLocalFileTime" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function APISystemTimeToFileTime Lib "kernel32" Alias "SystemTimeToFileTime" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function APILocalFileTimeToFileTime Lib "kernel32" Alias "LocalFileTimeToFileTime" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function APISetFileTime Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function APIDeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function APIRemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function APIGetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function APIGetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function APIMoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function APIGetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Private Declare Function APIGetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function APIGetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function APISetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Private Declare Function APIlstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function APIlstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function APIGetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function APIGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Private Declare Function APIGetFileSize Lib "kernel32" Alias "GetFileSize" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function APISetEndOfFile Lib "kernel32" Alias "SetEndOfFile" (ByVal hFile As Long) As Long
Private Declare Function APIFindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function APIFindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function APIFindClose Lib "kernel32" Alias "FindClose" (ByVal hFindFile As Long) As Long
Private Declare Function APIWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function APIReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function APISetFilePointer Lib "kernel32" Alias "SetFilePointer" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function APIGetStdHandle Lib "kernel32" Alias "GetStdHandle" (ByVal nStdHandle As Long) As Long
Private Declare Sub APICopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
'Private Declare Function APICreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

'Version.dll:
Private Declare Function APIGetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function APIGetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function APIVerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

'mpr.dll:
Private Declare Function APIWNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function APIWNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As Any, lpBufferSize As Long) As Long
Private Declare Function APIWNetCloseEnum Lib "mpr.dll" Alias "WNetCloseEnum" (ByVal hEnum As Long) As Long

'Shell32.dll:
Private Declare Function APISHGetSpecialFolderLocation Lib "shell32.dll" Alias "SHGetSpecialFolderLocation" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function APISHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function APISHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'Imagehlp.dll:
Private Declare Function APIMakeSureDirectoryPathExists Lib "imagehlp.dll" Alias "MakeSureDirectoryPathExists" (ByVal lpPath As String) As Long

'Shlwapi.dll:
'These Require Windows 2000 (or Windows NT 4.0 with Internet Explorer 4.0 or later); Requires Windows 98 (or Windows 95 with Internet Explorer 4.0 or later):
Private Declare Function APIPathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function APIPathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function APIPathStripToRoot Lib "shlwapi.dll" Alias "PathStripToRootA" (ByVal pszPath As String) As Long
Private Declare Function APIPathIsRoot Lib "shlwapi.dll" Alias "PathIsRootA" (ByVal pszPath As String) As Long



Public Function GetFileVersion(ByVal FilePath As String) As String

  Dim bBuffer()         As Byte
  Dim lRet              As Long
  Dim lDummy            As Long
  Dim lBufferLen        As Long
  Dim lVerPointer       As Long
  Dim lVerBufferLen     As Long
  Dim udtVerBuffer      As VS_FIXEDFILEINFO

  lBufferLen = APIGetFileVersionInfoSize(FilePath, lDummy) 'Get size
  CheckForError
  
  If lBufferLen < 1 Then
    GetFileVersion = ""
  Else
    ReDim bBuffer(lBufferLen) 'Store info to udtVerBuffer struct
    lRet = APIGetFileVersionInfo(FilePath, 0&, lBufferLen, bBuffer(0))
    CheckForError
    lRet = APIVerQueryValue(bBuffer(0), "\", lVerPointer, lVerBufferLen)
    CheckForError
    Call APICopyMemory(udtVerBuffer, lVerPointer, Len(udtVerBuffer))
    CheckForError
  
    With udtVerBuffer 'Determine File Version number
      GetFileVersion = Format$(.dwFileVersionMSh) & "." & Format$(.dwFileVersionMSl) & "." & Format$(.dwFileVersionLSh) & "." & Format$(.dwFileVersionLSl)
    End With
  End If
  
End Function

Public Function GetTempDir() As String
  
  Dim lRet As Long
  Dim lSize As Long
  Dim sBuf As String * MAX_PATH
  
  lSize = MAX_PATH
  lRet = APIGetTempPath(ByVal lSize, sBuf)
  CheckForError

  If InStr(1, sBuf, Chr(0)) > 0 Then
    GetTempDir = Left(sBuf, InStr(2, sBuf, Chr(0)) - 1)
  Else
    GetTempDir = ""
  End If
  
End Function


Public Function GetTempName() As String

  Dim sTempDir As String
  Dim sTempFile As String
  
  sTempDir = HandleSlash(GetTempDir, True)
  
  Do
    sTempFile = "rad" & VBA.CStr(VBA.Hex(VBA.Int((900000) * VBA.Rnd + 100000))) & ".tmp"
  Loop While FileExists(sTempDir & sTempFile)  'Hopefully this is the path they will try to save to, if not, let them deal with duplicate names.
  
  GetTempName = sTempDir & sTempFile
  
End Function

Public Function HandleSlash(ByVal sDirName As String, Optional IncludeSlash As Boolean = True) As String

  If sDirName Like "[a-z]:*" Then
    sDirName = VBA.UCase(VBA.Left(sDirName, 1)) & VBA.Right(sDirName, VBA.Len(sDirName) - 1)
  End If

  If VBA.Len(sDirName) > 0 Then
    If IncludeSlash Then
      sDirName = IIf(Right(sDirName, 1) = "\", sDirName, sDirName & "\")
    Else
      sDirName = IIf(Right(sDirName, 1) <> "\", sDirName, VBA.Left(sDirName, VBA.Len(sDirName) - 1))
      sDirName = IIf(Right(sDirName, 1) <> "/", sDirName, VBA.Left(sDirName, VBA.Len(sDirName) - 1))
    End If
  Else
    sDirName = "\"
  End If
  
  HandleSlash = sDirName
  
End Function


Public Function GetSpecialFolder(SpecialFolder As FileSystemAPI.SpecialFolderConst) As FileSystemAPI.Folder
  
  Dim lRet As Long
  Dim sPath As String
  Dim IDL As ITEMIDLIST
  
  Select Case SpecialFolder
    Case TemporaryFolder
      Set GetSpecialFolder = GetFolder(GetTempDir)
    Case Else
      sPath = Space$(MAX_PATH)
      ' These error alot for seemingly no reason:
      Call APISHGetSpecialFolderLocation(0, SpecialFolder, IDL)
      Call APISHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath$)
      sPath = StripTerminator(sPath, VBA.Chr(0))
      
      If VBA.Len(sPath) = 0 Then
        Err.Raise 5, , "Invalid procedure call or argument"
      Else
        Set GetSpecialFolder = GetFolder(sPath)
      End If
  End Select
  
End Function

Public Function FileOrFolderAttributesGet(ByVal FilePath As String) As FileSystemAPI.FileAttribute
  FilePath = GetAbsolutePathName(FilePath, False)
  FileOrFolderAttributesGet = APIGetFileAttributes(FilePath)
  CheckForError
End Function

Public Sub FileOrFolderAttributesSet(ByVal FilePath As String, lAtts As Long)
  FilePath = GetAbsolutePathName(FilePath, False)
  Call APISetFileAttributes(FilePath, lAtts)
  CheckForError
End Sub

Public Sub CopyFile(ByVal Source As String, ByVal Destination As String, Overwrite As Boolean)
  Source = GetAbsolutePathName(Source, False)
  Destination = GetAbsolutePathName(Destination, False)
  If FileExists(Source) Then
    Call APICopyFile(Source, Destination, Not (Overwrite))
    CheckForError
  Else
    Err.Raise 53, , "File not found"
  End If
End Sub

Public Function FileOrFolderDateGet(FilePath As String, iType As DateType) As Date
On Error GoTo ErrHandler

  Dim FtCreate As FILETIME
  Dim FtAccess As FILETIME
  Dim FtWrite As FILETIME
  Dim FtLocal As FILETIME
  
  Dim SysTime As SYSTEMTIME
  
  Dim lFile As Long
  
  'lFile = APICreateFile(FilePath, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
  lFile = APICreateFile(FilePath, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
  'lFile = APICreateFile(FilePath, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, 0, 0)
  Call CheckForError
  Call APIGetFileTime(lFile, FtCreate, FtAccess, FtWrite)
  CheckForError
  Call APICloseHandle(lFile)
  CheckForError
  lFile = 0
  
  Select Case iType 'Convert the file time to the local file time
    Case DateAccess
      Call APIFileTimeToLocalFileTime(FtAccess, FtLocal)
    Case DateCreate
      Call APIFileTimeToLocalFileTime(FtCreate, FtLocal)
    Case DateWrite
      Call APIFileTimeToLocalFileTime(FtWrite, FtLocal)
  End Select
  CheckForError
  
  'Convert the file time to system file time
  Call APIFileTimeToSystemTime(FtLocal, SysTime)
  CheckForError
  
  With SysTime
    FileOrFolderDateGet = .wMonth & "/" & .wDay & "/" & .wYear & " " & .wHour & ":" & .wMinute & ":" & .wSecond
  End With
 
Exit Function
ErrHandler:
  If lFile > 0 Then
    Call APICloseHandle(lFile)
  End If
  Err.Raise Err.Number, , Err.Description
End Function

Public Sub FileOrFolderDateSet(FilePath As String, dtDate As Date, iType As DateType)
On Error GoTo ErrHandler

  Dim FtCreate As FILETIME
  Dim FtAccess As FILETIME
  Dim FtWrite As FILETIME
  Dim FtLocal As FILETIME
  
  Dim SysTime As SYSTEMTIME
  
  Dim lFile As Long
  
  If IsDate(dtDate) Then
    
    With SysTime
      .wDay = Day(dtDate)
      .wDayOfWeek = Weekday(dtDate) - 1
      .wHour = Hour(dtDate)
      .wMinute = Minute(dtDate)
      .wMonth = Month(dtDate)
      .wSecond = Second(dtDate)
      .wYear = Year(dtDate)
    End With
    
    lFile = APICreateFile(FilePath, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    Call CheckForError
    Call APIGetFileTime(lFile, FtCreate, FtAccess, FtWrite)
    CheckForError
    Call APISystemTimeToFileTime(SysTime, FtLocal) ' convert system time to local time
    CheckForError
    
    Select Case iType 'convert local time to GMT
      Case DateAccess
        Call APILocalFileTimeToFileTime(FtLocal, FtAccess)
      Case DateCreate
        Call APILocalFileTimeToFileTime(FtLocal, FtCreate)
      Case DateWrite
        Call APILocalFileTimeToFileTime(FtLocal, FtWrite)
    End Select
    CheckForError
    
    Call APISetFileTime(lFile, FtCreate, FtAccess, FtWrite)
    CheckForError
    Call APICloseHandle(lFile)
    CheckForError
    lFile = 0
  End If
  
Exit Sub
ErrHandler:
  If lFile > 0 Then
    Call APICloseHandle(lFile)
  End If
  Err.Raise Err.Number, , Err.Description
End Sub

Public Function GetFile(FilePath As String) As FileSystemAPI.File
  If FileExists(FilePath) Then
    Set GetFile = New File
    GetFile.Path = FilePath
  Else
    Err.Raise 53, , "File Not Found"
  End If
End Function

Public Function GetFolder(FilePath As String) As FileSystemAPI.Folder
  
  If FolderExists(FilePath) Then
    Set GetFolder = New FileSystemAPI.Folder
    GetFolder.Path = GetAbsolutePathName(FilePath, False)
  Else
    Err.Raise 76, , "Path Not Found"
  End If
  
End Function

Public Function GetFolderName(ByVal FilePath As String) As String
  
  Dim lTemp As Long
  
  If Not PathIsRoot(FilePath) Then
  
    FilePath = HandleSlash(FilePath, False)
    
    lTemp = VBA.InStrRev(FilePath, "\")
    
    If lTemp = 0 Then
      lTemp = VBA.InStrRev(FilePath, "/")
    End If
    
    If lTemp > 1 Then
      FilePath = VBA.Right(FilePath, VBA.Len(FilePath) - lTemp)
    Else
      FilePath = ""
    End If
    
  Else
    FilePath = ""
  End If
    
  GetFolderName = FilePath
  
End Function

Public Function GetDrive(FilePath As String) As FileSystemAPI.Drive
  If DriveExists(FilePath) Then
    Set GetDrive = New FileSystemAPI.Drive
    GetDrive.Path = FilePath
  Else
    Err.Raise 68, , "Device unavailable"
  End If
End Function

Public Function GetDriveCollection() As Collection
  
  Dim sDrives As String
  Dim ArrDrives() As String
  Dim iCount As Integer
  Dim oDrive As FileSystemAPI.Drive
  
  sDrives = String(255, Chr$(0))
  Call APIGetLogicalDriveStrings(255, sDrives)
  CheckForError
  
  sDrives = StripTerminator(sDrives, VBA.Chr(0))
  
  Set GetDriveCollection = New Collection
  
  ArrDrives = VBA.Split(sDrives, Chr$(0))
  
  For iCount = LBound(ArrDrives) To UBound(ArrDrives)
    If VBA.Len(ArrDrives(iCount)) > 0 Then
      Set oDrive = New FileSystemAPI.Drive
      oDrive.Path = ArrDrives(iCount)
      GetDriveCollection.Add oDrive, oDrive.Path
    End If
  Next
  
End Function

Public Function GetFiles(ByVal FilePath As String) As FileSystemAPI.Files
  Set GetFiles = New FileSystemAPI.Files
  GetFiles.Path = FilePath
End Function

Public Function GetFolders(ByVal FilePath As String) As FileSystemAPI.Folders
  Set GetFolders = New FileSystemAPI.Folders
  GetFolders.Path = FilePath
End Function

Public Function FindFiles(ByVal FilePath As String, Optional FileLike As String = "*", Optional Recursive As Boolean = False, Optional FirstFound As Boolean = False) As FileSystemAPI.Files
  Set FindFiles = New FileSystemAPI.Files
  With FindFiles
    .FileLike = FileLike
    .Recursive = Recursive
    .FirstFound = FirstFound
    .Path = FilePath
  End With
End Function

Public Function FindFolders(ByVal FilePath As String, Optional FolderLike As String = "*", Optional Recursive As Boolean = False, Optional FirstFound As Boolean = False) As FileSystemAPI.Folders
  Set FindFolders = New FileSystemAPI.Folders
  With FindFolders
    .FileLike = FolderLike
    .Recursive = Recursive
    .FirstFound = FirstFound
    .Path = FilePath
  End With
End Function

Public Function GetFileCollection(ByVal Path As String, Optional ByVal FileLike As String = "*", Optional ByVal Recursive As Boolean = False, Optional FirstFound As Boolean = False) As Collection
On Error GoTo ErrHandler

  Dim WFD As WIN32_FIND_DATA
  Dim sFileName As String
  Dim lSearch As Long
  Dim lMoreFiles As Long
  Dim oFile As FileSystemAPI.File
  Dim colTemp As Collection
  
  Set GetFileCollection = New Collection

  FileLike = VBA.UCase(FileLike)
  Path = GetAbsolutePathName(Path, True)
  lSearch = APIFindFirstFile(Path & "*", WFD)
  Call CheckForError

  lMoreFiles = 1
  If lSearch <> INVALID_HANDLE_VALUE Then
    Do While lMoreFiles > 0
      sFileName = StripTerminator(WFD.cFileName, Chr(0), False)
      If (sFileName <> ".") And (sFileName <> "..") Then
        If (APIGetFileAttributes(Path & sFileName) And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
          If UCase(sFileName) Like FileLike Then
            Set oFile = New FileSystemAPI.File
            oFile.Path = Path & sFileName
            GetFileCollection.Add oFile, VBA.CStr(oFile.Path)
            If FirstFound Then
              Exit Do
            End If
          End If
        ElseIf Recursive Then
          Set colTemp = GetFileCollection(Path & sFileName, FileLike, Recursive, FirstFound)
          For Each oFile In colTemp
            GetFileCollection.Add oFile, VBA.CStr(oFile.Path)
            If FirstFound Then
              Exit Do
            End If
          Next
        End If
      End If
      lMoreFiles = APIFindNextFile(lSearch, WFD) ' Get next file
      CheckForError
    Loop
    Call APIFindClose(lSearch)
    CheckForError
  End If

ExitProc:
  Set colTemp = Nothing
  Set oFile = Nothing
Exit Function
ErrHandler:
  If lSearch <> INVALID_HANDLE_VALUE Then
    Call APIFindClose(lSearch)
  End If
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetFolderCollection(ByVal Path As String, Optional ByVal FolderLike As String = "*", Optional ByVal Recursive As Boolean = False, Optional FirstFound As Boolean = False) As Collection
On Error GoTo ErrHandler

  Dim WFD As WIN32_FIND_DATA
  Dim sFileName As String
  Dim lSearch As Long
  Dim lMoreFiles As Long
  Dim colTemp As Collection
  Dim oFolder As FileSystemAPI.Folder
  
  Set GetFolderCollection = New Collection
  
  FolderLike = UCase(FolderLike)
  Path = GetAbsolutePathName(Path, True)
  lSearch = APIFindFirstFile(Path & UCase(FolderLike), WFD)
  Call CheckForError
  
  lMoreFiles = 1
  If lSearch <> INVALID_HANDLE_VALUE Then
    Do While lMoreFiles > 0
      sFileName = StripTerminator(WFD.cFileName, Chr(0), False)
      If (sFileName <> ".") And (sFileName <> "..") Then
        If (APIGetFileAttributes(Path & sFileName) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
          If UCase(sFileName) Like FolderLike Then
            Set oFolder = New FileSystemAPI.Folder
            oFolder.Path = Path & sFileName
            GetFolderCollection.Add oFolder, oFolder.Path
            If Recursive Then
              Set colTemp = GetFolderCollection(Path & sFileName, FolderLike, Recursive, FirstFound)
              For Each oFolder In colTemp
                GetFolderCollection.Add oFolder, oFolder.Path
                If FirstFound Then
                  Exit Do
                End If
              Next
            End If
          End If
        End If
      End If
      lMoreFiles = APIFindNextFile(lSearch, WFD) ' Get next file
      CheckForError
    Loop
    Call APIFindClose(lSearch)
    CheckForError
  End If

ExitProc:
  Set colTemp = Nothing
  Set oFolder = Nothing
Exit Function
ErrHandler:
  If lSearch <> INVALID_HANDLE_VALUE Then
    Call APIFindClose(lSearch)
  End If
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetFolderSize(Path As String) As Double
On Error GoTo ErrHandler

  Dim WFD As WIN32_FIND_DATA
  Dim sFileName As String
  Dim lSearch As Long
  Dim lMoreFiles As Long
  
  Path = GetAbsolutePathName(Path, True)

  lSearch = APIFindFirstFile(Path & "*", WFD)
  Call CheckForError
  
  lMoreFiles = 1
  If lSearch <> INVALID_HANDLE_VALUE Then
    While lMoreFiles > 0
      sFileName = StripTerminator(WFD.cFileName, Chr(0), False)
      If (sFileName <> ".") And (sFileName <> "..") Then
        If (APIGetFileAttributes(Path & sFileName) And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
          GetFolderSize = GetFolderSize + WFD.nFileSizeLow
        ElseIf (APIGetFileAttributes(Path & sFileName) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
          GetFolderSize = GetFolderSize + GetFolderSize(Path & sFileName)
        End If
      End If
      lMoreFiles = APIFindNextFile(lSearch, WFD) ' Get next file
      CheckForError
    Wend
    Call APIFindClose(lSearch)
    CheckForError
  End If

Exit Function
ErrHandler:
  If lSearch <> INVALID_HANDLE_VALUE Then
    Call APIFindClose(lSearch)
  End If
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub DeleteFile(ByVal FilePath As String, ByVal Force As Boolean)
  
  Dim lAtts As FileSystemAPI.FileAttribute
  
  FilePath = GetAbsolutePathName(FilePath, False)
  
  lAtts = FileOrFolderAttributesGet(FilePath)
  
  If (lAtts And FileAttribute.ReadOnly) Then
    If Force Then
      lAtts = lAtts - FileAttribute.ReadOnly
      Call FileOrFolderAttributesSet(FilePath, lAtts)
    End If
  End If
  
  Call APIDeleteFile(FilePath)
  CheckForError

End Sub

Public Sub DeleteFolder(ByVal FilePath As String, ByVal Force As Boolean)
On Error GoTo ErrHandler

  Dim WFD As WIN32_FIND_DATA
  Dim sFileName As String
  Dim lSearch As Long
  Dim lMoreFiles As Long
  Dim lAtts As FileSystemAPI.FileAttribute

  FilePath = GetAbsolutePathName(FilePath, True)
  
  If Not FolderExists(FilePath) Then
    Err.Raise 53, , "File not found"
  End If
  
  If Force Then
    lAtts = FileOrFolderAttributesGet(FilePath)
    If (lAtts And FileAttribute.ReadOnly) Then
      lAtts = lAtts - FileAttribute.ReadOnly
      Call FileOrFolderAttributesSet(FilePath, lAtts)
    End If
  End If
  
  lSearch = APIFindFirstFile(FilePath & "*", WFD)
  CheckForError

  lMoreFiles = 1
  If lSearch <> INVALID_HANDLE_VALUE Then
    While lMoreFiles > 0
      sFileName = StripTerminator(WFD.cFileName, Chr(0), False)
      If (sFileName <> ".") And (sFileName <> "..") Then
        If (APIGetFileAttributes(FilePath & sFileName) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
          Call DeleteFolder(FilePath & sFileName, Force)
        ElseIf (APIGetFileAttributes(FilePath & sFileName) And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
          Call DeleteFile(FilePath & sFileName, Force)
        End If
      End If
      lMoreFiles = APIFindNextFile(lSearch, WFD) ' Get next file
      CheckForError
    Wend
    Call APIFindClose(lSearch)
    CheckForError
  End If
  Call APIRemoveDirectory(FilePath)
  CheckForError

Exit Sub
ErrHandler:
  If lSearch <> INVALID_HANDLE_VALUE Then
    Call APIFindClose(lSearch)
  End If
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function BuildPath(Path As String, Name As String) As String
  BuildPath = HandleSlash(Path, True) & Name
End Function

Public Function GetAbsolutePathName(FilePath As String, Optional bIncludeBackSlash As Boolean = False) As String
  
  Dim sBuffer As String
  Dim lRet As Long
  
  sBuffer = Space(255) 'create a buffer
  lRet = APIGetFullPathName(FilePath, 255, sBuffer, "") 'copy the current directory to the buffer and append FilePath
  CheckForError
  sBuffer = Left(sBuffer, lRet)  'remove the unnecessary chr$(0)'s
  
  If VBA.Len(sBuffer) > 0 Then
    If VBA.Len(sBuffer) > 3 Then ' If it's just the drive letter then keep the slash
      If bIncludeBackSlash And VBA.Right(sBuffer, 1) <> "\" Then
        sBuffer = sBuffer & "\"
      ElseIf (Not bIncludeBackSlash) And VBA.Right(sBuffer, 1) = "\" Then
        sBuffer = VBA.Left(sBuffer, VBA.Len(sBuffer) - 1)
      End If
    End If
    
    If sBuffer Like "[a-z]:*" Then
      sBuffer = VBA.UCase(VBA.Left(sBuffer, 1)) & VBA.Right(sBuffer, VBA.Len(sBuffer) - 1)
    End If
    
  End If
    
  GetAbsolutePathName = sBuffer
  
End Function

Public Function CreateFolder(ByVal FilePath As String) As FileSystemAPI.Folder
  
  FilePath = GetAbsolutePathName(FilePath, True)
  
  If FolderExists(FilePath) Then
    Err.Raise 58, , "File already exists"
  Else
    Call APIMakeSureDirectoryPathExists(FilePath) ' Will throw error if base path doesn't exist, but still does job...
  End If
  
  Set CreateFolder = New FileSystemAPI.Folder
  CreateFolder.Path = FilePath
  
End Function

Public Function FolderExists(FilePath As String) As Boolean
On Error GoTo ErrHandler

  FolderExists = False

  If VBA.Len(FilePath) Then
Method1: 'API Route:
    If APILoadLibrary("shlwapi.dll") Then
      CheckForError
      FolderExists = APIPathIsDirectory(FilePath)
      CheckForError
    Else
Method2: 'VB Route:
      FolderExists = (VBA.GetAttr(FilePath) And vbDirectory)
    End If
  End If

ExitProc:
Exit Function
ErrHandler:
  Select Case Err.Number
    Case 53 ' File not found
      FolderExists = False
      Resume ExitProc
    Case 453 ' Entry point not found in DLL
      Resume Method2
    Case Else
      Err.Raise Err.Number, , Err.Description
  End Select
End Function

Public Function FileExists(FilePath As String) As Boolean
On Error GoTo ErrHandler

  FileExists = False
  
  If Len(FilePath) Then
Method1: 'API Route:
    If APILoadLibrary("shlwapi.dll") Then
      CheckForError
      If APIPathFileExists(FilePath) Then
        CheckForError
        FileExists = Not FolderExists(FilePath)
      End If
    Else
Method2: 'VB Route:
      If (GetAttr(FilePath) And vbDirectory) < 1 Then
        'File Exists
        FileExists = True
      End If
    End If
  End If

ExitProc:
Exit Function
ErrHandler:
  Select Case Err.Number
    Case 53 ' File not found
      FileExists = False
      Resume ExitProc
    Case 453 ' Entry point not found in DLL
      Resume Method2
    Case Else
      Err.Raise Err.Number, , Err.Description
  End Select
End Function

Public Function DriveExists(FilePath As String) As Boolean
  Dim lRet As Long
  lRet = APIGetDriveType(PathToRoot(FilePath, True))
  CheckForError
  If lRet > 1 Then
    DriveExists = True
  Else
    DriveExists = False
  End If
End Function


Public Function PathToRoot(ByVal FilePath As String, Optional ByVal IncludeSlash As Boolean = True) As String
On Error GoTo ErrHandler

  Dim lTemp As Long
  Dim lCount As Long

  PathToRoot = ""

  If VBA.Len(FilePath) = 1 Then
    FilePath = FilePath & ":"
  End If
  
Method1: 'API Method:
  If (APILoadLibrary("shlwapi.dll")) Then
    CheckForError
    Call APIPathStripToRoot(FilePath)
    CheckForError
    FilePath = StripTerminator(FilePath, VBA.Chr(0))
    PathToRoot = FilePath
  Else
Method2: 'Parse Method:
    lTemp = 0
    lCount = 0
    If FilePath Like "[A-Z]:*" Then
      PathToRoot = VBA.Left(FilePath, 2)
    ElseIf FilePath Like "[A-Z]" Then
      PathToRoot = FilePath & ":"
    ElseIf FilePath Like "[a-z]:*" Then
      PathToRoot = VBA.UCase(VBA.Left(FilePath, 2))
    ElseIf FilePath Like "[a-z]" Then
      PathToRoot = VBA.UCase(FilePath & ":")
    ElseIf FilePath Like "//*/*/*" Then
      lTemp = 0
      For lCount = 3 To VBA.Len(FilePath)
        If VBA.Mid(FilePath, lCount, 1) = "/" Then
          lTemp = lTemp + 1
        End If
        If lTemp = 2 Then
          Exit For
        End If
      Next
      PathToRoot = VBA.Left(FilePath, lCount - 1)
    ElseIf (FilePath Like "\\*\*\*") Or (FilePath Like "//*/*\*") Then
      lTemp = 0
      For lCount = 3 To VBA.Len(FilePath)
        If VBA.Mid(FilePath, lCount, 1) = "\" Then
          lTemp = lTemp + 1
        End If
        If lTemp = 2 Then
          Exit For
        End If
      Next
      PathToRoot = VBA.Left(FilePath, lCount - 1)
    Else
      PathToRoot = ""
    End If
  End If
  
  If IncludeSlash Then
    PathToRoot = HandleSlash(PathToRoot, True)
  Else
    PathToRoot = HandleSlash(PathToRoot, False)
  End If
  
ExitProc:
Exit Function
ErrHandler:
  Select Case Err.Number
    Case 453 ' Entry point not found in DLL
      Resume Method2
    Case Else
      Err.Raise Err.Number, , Err.Description
  End Select
End Function


Public Function PathIsRoot(ByVal FilePath As String) As Boolean
On Error GoTo ErrHandler
  
  PathIsRoot = False
  
  FilePath = HandleSlash(FilePath, True)
  
Method1: 'API Method
  If APILoadLibrary("shlwapi.dll") Then
    CheckForError
    PathIsRoot = APIPathIsRoot(FilePath)
    CheckForError
  Else
Method2: 'Parse Method
    If FilePath Like "[a-z]:\" Then
      PathIsRoot = True
    ElseIf FilePath Like "[a-z]:" Then
      PathIsRoot = True
    ElseIf FilePath Like "[A-Z]:\" Then
      PathIsRoot = True
    ElseIf FilePath Like "[A-Z]:" Then
      PathIsRoot = True
    ElseIf FilePath Like "\\*\*\" Then
      PathIsRoot = True
    ElseIf FilePath Like "//*/*\" Then
      PathIsRoot = True
    ElseIf FilePath Like "//*/*/" Then
      PathIsRoot = True
    Else
      PathIsRoot = False
    End If
  End If

ExitProc:
Exit Function
ErrHandler:
  Select Case Err.Number
    Case 453 ' Entry point not found in DLL
      Resume Method2
    Case Else
      Err.Raise Err.Number, , Err.Description
  End Select
End Function

Public Sub CopyFolder(ByVal Source As String, ByVal Destination As String, ByVal Overwrite As Boolean)
On Error GoTo ErrHandler

  Dim WFD As WIN32_FIND_DATA
  Dim sFileName As String
  Dim lSearch As Long
  Dim lMoreFiles As Long

  Source = GetAbsolutePathName(Source, True)
  Destination = GetAbsolutePathName(Destination, True)

  If Overwrite Then
    Call APIMakeSureDirectoryPathExists(Destination) ' Errors alot
  Else
    If FolderExists(Destination) Then
      Err.Raise 58, , "File already exists"
    Else
      Call APIMakeSureDirectoryPathExists(Destination)
    End If
  End If

  lSearch = APIFindFirstFile(Source & "*", WFD)
  CheckForError

  lMoreFiles = 1
  If lSearch <> INVALID_HANDLE_VALUE Then
    While lMoreFiles > 0
      sFileName = StripTerminator(WFD.cFileName, Chr(0), False)
      If (sFileName <> ".") And (sFileName <> "..") Then
        If (APIGetFileAttributes(Source & sFileName) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
          Call CopyFolder(Source & sFileName, Destination & sFileName, Overwrite)
        ElseIf (APIGetFileAttributes(Source & sFileName) And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
          Call CopyFile(Source & sFileName, Destination & sFileName, Overwrite)
        End If
      End If
      lMoreFiles = APIFindNextFile(lSearch, WFD) ' Get next file
      CheckForError
    Wend
    Call APIFindClose(lSearch)
    CheckForError
  End If

Exit Sub
ErrHandler:
  If lSearch <> INVALID_HANDLE_VALUE Then
    Call APIFindClose(lSearch)
  End If
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function MoveFileOrFolder(ByVal Source As String, ByVal Destination As String) As String
  
  Source = GetAbsolutePathName(Source, False)
  Destination = GetAbsolutePathName(Destination, False)
  
  Call APIMoveFile(Source, Destination)
  CheckForError
  MoveFileOrFolder = Destination
  
End Function

Public Function RenameFileOrFolder(ByVal Source As String, ByVal NewName As String) As String
  
  Dim Destination As String
  
  Source = GetAbsolutePathName(Source, False)
  Destination = GetPath(Source, True) & NewName
  Destination = GetAbsolutePathName(Destination, False)
  
  Call APIMoveFile(Source, Destination)
  CheckForError
  RenameFileOrFolder = Destination
  
End Function

Public Function OpenTextFile(ByVal FilePath As String, ByVal Mode As FileSystemAPI.IOMode, ByVal Create As Boolean, ByVal Overwrite As Boolean, ByVal Format As FileSystemAPI.Tristate, Optional StandardStream As Boolean = False, Optional StandardStreamType As FileSystemAPI.StandardStreamTypes = StdErr) As FileSystemAPI.TextStream
  
  Set OpenTextFile = New FileSystemAPI.TextStream
  Call OpenTextFile.OpenFile(FilePath, Mode, Create, Overwrite, Format, StandardStream, StandardStreamType)
  
End Function

Public Function OpenFile(ByVal FilePath As String, ByVal Mode As FileSystemAPI.IOMode, ByVal Create As Boolean, ByVal Overwrite As Boolean, ByRef Location As Long, ByVal Format As FileSystemAPI.Tristate, Optional StandardStream As Boolean = False, Optional StandardStreamType As FileSystemAPI.StandardStreamTypes = StdErr) As Long
On Error GoTo ErrHandler
  
  Dim lStdType As Long
  Dim lCreation As Long
  Dim lAccess As Long
  Dim lFile As Long
  
  'Dim lSize As Long
  'Dim lPointer As Long
  
  If Create Then
    If Overwrite Then
      lCreation = CREATE_ALWAYS
    Else
      lCreation = CREATE_NEW
    End If
  Else
    lCreation = OPEN_EXISTING
  End If
  
  If StandardStream Then
    
    Select Case StandardStreamType
      Case StdErr
        lStdType = STD_ERROR_HANDLE
        Mode = ForAppending
      Case StdIn
        lStdType = STD_INPUT_HANDLE
        Mode = ForReading
      Case StdOut
        lStdType = STD_OUTPUT_HANDLE
        Mode = ForWriting
    End Select
    
    lFile = APIGetStdHandle(lStdType)
    Call CheckForError
    
  Else
  
    Select Case Mode
      Case ForWriting
        lAccess = GENERIC_WRITE
        'lPointer = FILE_BEGIN
      Case ForReading
        lAccess = GENERIC_READ
        'lPointer = FILE_BEGIN
      Case ForAppending
        lAccess = GENERIC_WRITE
        'lPointer = FILE_END
    End Select
    
    FilePath = GetAbsolutePathName(FilePath, False)
    lFile = APICreateFile(FilePath, lAccess, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, lCreation, 0, 0)
    CheckForError
  
  End If
  
  If Mode = ForWriting Then
    Call ClearFile(lFile)
  End If

  If Format = TristateTrue Then ' Unicode
    If Mode = ForWriting Then
    'If Mode = ForWriting Or Mode = ForAppending Then
      Call MakeFileUnicode(lFile)
      Location = 2
    ElseIf Mode = ForAppending Then
      Location = APIGetFileSize(lFile, 0)
      CheckForError
    Else
      If IsFileUnicode(lFile) Then
        Location = 2
      Else
        Location = 0
      End If
    End If
  Else ' ASCII
    If Mode = ForAppending Then
      Location = APIGetFileSize(lFile, 0)
      CheckForError
    Else
      Location = 0
    End If
  End If
  
  OpenFile = lFile
  
Exit Function
ErrHandler:
  Select Case Err.Number
    Case 58 ' File already exists
      If Overwrite Then
        Resume Next ' Ignore
      Else
        If lFile > 0 Then
          Call CloseFile(lFile)
        End If
        Err.Raise Err.Number, , Err.Description
      End If
    Case Else
      If lFile > 0 Then
        Call CloseFile(lFile)
      End If
      Err.Raise Err.Number, , Err.Description
  End Select
End Function

Public Function IsFileUnicode(ByVal lFile As Long) As Boolean

  Dim lSize As Long
  Dim bBytes() As Byte
  Dim lRet As Long
  
  IsFileUnicode = False
  
  lSize = APIGetFileSize(lFile, 0)
  CheckForError
  
  If lSize >= 2 Then
    Call APISetFilePointer(lFile, 0, 0, FILE_BEGIN)
    CheckForError
    
    ReDim bBytes(1 To 2) As Byte
    
    Call APIReadFile(lFile, bBytes(1), UBound(bBytes), lRet, ByVal 0&)
    CheckForError
    
    If bBytes(1) = &HFF Then
      If bBytes(2) = &HFE Then
        IsFileUnicode = True
      End If
    End If
  End If
  
End Function

Public Sub MakeFileUnicode(ByVal lFile As Long)
  
  Dim lSize As Long
  Dim bBytes() As Byte
  Dim lRet As Long
  
  lSize = APIGetFileSize(lFile, 0)
  CheckForError
  
  Call APISetFilePointer(lFile, 0, 0, FILE_BEGIN)
  CheckForError
  
  ReDim bBytes(1 To 2) As Byte
  bBytes(1) = &HFF
  bBytes(2) = &HFE
  
  Call APIWriteFile(lFile, bBytes(1), UBound(bBytes), lRet, ByVal 0&)
  CheckForError
  
End Sub


Public Function ClearFile(ByVal lFile As Long)

  Dim bBytes() As Byte
  Dim lRet As Long
  Dim lSize As Long
  
  lSize = APIGetFileSize(lFile, 0)
  CheckForError
  
  If lSize > 0 Then
  
    Call APISetFilePointer(lFile, 0, 0, FILE_BEGIN)
    CheckForError
  
    Call APISetEndOfFile(lFile)
    CheckForError
    
  End If
  
End Function

Public Function ReadAllFromFile(ByVal lFile As Long, ByRef lLocation As Long, ByVal Format As FileSystemAPI.Tristate) As Variant
  
  Dim lEnd As Long

  lEnd = -1
  ReadAllFromFile = ReadFromFile(lFile, lLocation, lEnd, Format)
  lLocation = lEnd
  
End Function

Public Function ReadChars(ByVal lFile As Long, ByRef lLocation As Long, ByVal lChars As Long, ByVal Format As FileSystemAPI.Tristate) As String

  Dim lEnd As Long

  lEnd = lLocation + lChars
  ReadChars = ReadFromFile(lFile, lLocation, lEnd, Format)
  lLocation = lEnd

End Function

Public Function ReadFromFile(ByVal lFile As Long, ByRef lStart As Long, ByRef lEnd As Long, ByVal Format As FileSystemAPI.Tristate, Optional ByVal StopAtChar As String, Optional ByVal Compare As VbCompareMethod) As String

  Dim lBytesRead As Long
  Dim lPos As Long
  Dim sText As String
  Dim lCount As Long
  Dim lChunk As Long
  
  ReadFromFile = ""
  
  lChunk = 1028 'Default max chunk size
  
  If lEnd < 0 Then
    lEnd = APIGetFileSize(lFile, 0)
    CheckForError
  End If
  
  If lStart < 0 Then
    lStart = 0
  End If

  If lEnd >= lStart Then

    If lEnd - lStart = 0 Then
      lChunk = 1
    ElseIf lChunk > lEnd - lStart Then
      lChunk = lEnd - lStart
    End If

    If VBA.Len(StopAtChar) > 0 Then
      If Format = TristateTrue Then
        StopAtChar = VBA.StrConv(StopAtChar, vbUnicode)
      End If
    End If
    
    'lPos = lStart
    'Do
    'For lPos = lStart To lEnd - lChunk Step lChunk
    For lPos = lStart To lEnd Step lChunk
    
      Call APISetFilePointer(lFile, lPos, 0, FILE_BEGIN)
      CheckForError
      
      sText = VBA.Space$(lChunk)
      Call APIReadFile(lFile, ByVal sText, lChunk, lBytesRead, ByVal 0&)
      CheckForError
      
      If lBytesRead < lChunk Then
        sText = VBA.Left$(sText, lBytesRead)
      End If
      
      If VBA.Len(StopAtChar) > 0 Then
        lCount = VBA.InStr(1, sText, StopAtChar, Compare)
        If lCount > 0 Then ' Character found!
          lCount = lCount + (VBA.Len(StopAtChar) - 1)
          sText = VBA.Left(sText, lCount)
          ReadFromFile = ReadFromFile & sText
          Exit For
          'Exit Do
        End If
      End If
      
      ReadFromFile = ReadFromFile & sText
      
      If lBytesRead < lChunk Then ' End of File!
        Exit For
        'Exit Do
      End If
      
      'lPos = lPos + lChunk
    Next

    lStart = lStart + VBA.Len(ReadFromFile)
    
    If Format = TristateTrue Then
      ReadFromFile = VBA.StrConv(ReadFromFile, vbFromUnicode)
    End If

  Else
    Err.Raise 62, , "Input past end of file"
  End If
  
  'Debug.Print ReadFromFile
  
  
End Function

Public Sub WriteToFile(ByVal lFile As Long, ByRef lStart As Long, ByVal Text As String, ByVal Format As FileSystemAPI.Tristate)
  
  Dim lBytesWrote As Long
  Dim lEnd As Long
  Dim lSize As Long
  Dim lPos As Long
  Dim lTextSize As Long
  Dim sText As String
  Dim lChunk As Long
  
  lChunk = 1028 'Default max chunk size
  
  lSize = APIGetFileSize(lFile, 0)
  CheckForError
  
  If lSize >= lStart Then

    If Format = TristateTrue Then
      Text = VBA.StrConv(Text, vbUnicode)
    End If

    lTextSize = VBA.Len(Text)
    lEnd = lStart + lTextSize
    
    If lChunk > lEnd - lStart Then
      lChunk = lEnd - lStart
    End If
    
    For lPos = 1 To lEnd - lStart Step lChunk
    
      Call APISetFilePointer(lFile, lPos + lStart - 1, 0, FILE_BEGIN)
      CheckForError
      
      If lChunk = lTextSize Then
        sText = Text
      ElseIf lChunk + (lPos - 1) > lTextSize Then
        sText = VBA.Mid(Text, lPos, lTextSize - (lPos - 1))
      Else
        sText = VBA.Mid(Text, lPos, lChunk)
      End If
      
      Call APIWriteFile(lFile, ByVal sText, VBA.Len(sText), lBytesWrote, ByVal 0&)
      'Call APIWriteFile(lFile, ByVal sText, lChunk, lBytesWrote, ByVal 0&)
      CheckForError
      
      If lBytesWrote < VBA.Len(sText) Then
        Exit For
      End If
      
    Next
    
    lStart = lEnd
    
  Else
    Err.Raise 62, , "Input past end of file"
  End If
  
End Sub


Public Function ReadLine(ByVal lFile As Long, ByRef lLocation As Long, ByVal Format As FileSystemAPI.Tristate) As String
  
  Dim sText As String
  
  sText = ReadFromFile(lFile, lLocation, 0, Format, vbLf, vbTextCompare)
  
  If VBA.Len(sText) >= 2 Then
    If VBA.Right(sText, 2) = vbCrLf Then
      sText = VBA.Left(sText, VBA.Len(sText) - 2)
    ElseIf VBA.Right(sText, 1) = vbLf Then
      sText = VBA.Left(sText, VBA.Len(sText) - 1)
    End If
  ElseIf VBA.Len(sText) = 1 Then
    If sText = vbLf Then
      sText = ""
    End If
  End If
  
  ReadLine = sText
  
End Function

Public Sub WriteBlankLines(ByVal lFile As Long, ByRef lLocation As Long, ByVal lLines As Long, ByVal Format As FileSystemAPI.Tristate)

  Dim sText As String
  
  sText = ""
  For lLines = 1 To lLines
    sText = sText & vbCrLf
  Next
  
  Call WriteToFile(lFile, lLocation, sText, Format)
  
End Sub

Public Function BytesToString(bBytes() As Byte, Format As FileSystemAPI.Tristate) As String
  
  Dim iUnicode As Integer
  Dim sText As String
  Dim lCount As Long
  Dim lRet As Long
  
  sText = ""
  
  Select Case Format
    Case TristateTrue ' Unicode
      For lCount = LBound(bBytes) + 1 To UBound(bBytes) Step 2
        If bBytes(lCount - 1) = 255 And bBytes(lCount) = 254 Then
          ' Beginning of file
        Else
          sText = sText & VBA.ChrW(bBytes(lCount - 1) + bBytes(lCount) * 100)
        End If
      Next
    Case Else ' Use system default (ASCII)
      For lCount = LBound(bBytes) To UBound(bBytes)
        sText = sText & VBA.Chr(bBytes(lCount))
      Next
  End Select
  
  BytesToString = sText
  
End Function

Public Function StringToBytes(Text As String, Format As FileSystemAPI.Tristate) As Byte()
  
  Dim lCount As Long
  Dim bOut() As Byte
  Dim sChar As String
  Dim bFound As Boolean
  
  Select Case Format
    Case TristateTrue
      ReDim bOut(1 To VBA.Len(Text) * 2)
      For lCount = LBound(bOut) + 1 To UBound(bOut) Step 2
        sChar = VBA.Mid(Text, (lCount) / 2, 1)
        bOut(lCount - 1) = VBA.AscW(sChar)
        bOut(lCount) = 0
      Next
    Case Else
      ReDim bOut(1 To VBA.Len(Text))
      For lCount = 1 To VBA.Len(Text)
        sChar = VBA.Mid(Text, lCount, 1)
        bOut(lCount) = VBA.Asc(sChar)
      Next
  End Select
  
  StringToBytes = bOut
  
End Function

Public Sub WriteText(ByVal lFile As Long, ByRef lLocation As Long, ByVal Text As String, ByVal Format As FileSystemAPI.Tristate)
  
  Call WriteToFile(lFile, lLocation, Text, Format)
  
End Sub

Public Function AtEndOfLine(ByVal lFile As Long, ByVal lLocation As Long, ByVal Format As FileSystemAPI.Tristate) As Boolean

  Dim sText As String
  Dim lStart As Long
  
  AtEndOfLine = False
  
  lStart = lLocation
  
  If AtEndOfStream(lFile, lLocation, Format) Then
    AtEndOfLine = True
  Else
    sText = ReadFromFile(lFile, lLocation, lLocation + 1, Format)
    If sText = vbLf Then
      AtEndOfLine = True
    ElseIf sText = vbCr Then
      lLocation = lStart
      sText = ReadFromFile(lFile, lLocation + 1, lLocation + 2, Format)
      If sText = vbLf Then
        AtEndOfLine = True
      End If
    End If
  End If
  
End Function

Public Sub SkipLine(ByVal lFile As Long, ByRef lLocation As Long, ByVal Format As FileSystemAPI.Tristate)
  Call ReadLine(lFile, lLocation, Format)
End Sub

Public Function AtEndOfStream(ByVal lFile As Long, ByRef lLocation As Long, ByVal Format As FileSystemAPI.Tristate) As Boolean

  Dim lSize As Long
  
  lSize = APIGetFileSize(lFile, 0)
  CheckForError

  AtEndOfStream = (lSize <= lLocation)
  
End Function

Public Function GetStreamLine(ByVal lFile As Long, ByVal lLocation As Long, ByVal Format As FileSystemAPI.Tristate) As Long
  
  Dim lPos As Long
  Dim sText As String
  
  GetStreamLine = 1
  sText = ReadFromFile(lFile, 0, lLocation, Format)
  lPos = 1
  
  Do While (InStr(lPos, sText, vbLf) > 0)
    lPos = InStr(lPos, sText, vbLf) + 1
    GetStreamLine = GetStreamLine + 1
  Loop
  
End Function

Public Function GetStreamColumn(ByVal lFile As Long, ByVal lLocation As Long, ByVal Format As FileSystemAPI.Tristate) As Long
  
  Dim lPos As Long
  Dim sText As String
  
  sText = ReadFromFile(lFile, 0, lLocation, Format)
  lPos = VBA.InStrRev(sText, vbLf)
  sText = VBA.Mid(sText, lPos + 1)
  GetStreamColumn = VBA.Len(sText) + 1
  
End Function

Public Sub CloseFile(ByVal lHandle As Long)
  Call APICloseHandle(lHandle)
End Sub

Public Function GetDriveSize(sDrive As String, Info As DriveSize) As Variant
On Error GoTo ErrHandler

  Dim lSectorsPerCluster As Long
  Dim lBytesPerSector As Long
  Dim lNumberOfFreeSectors As Long
  Dim lTotalNumberOfSectors As Long
  
  Dim cFreeBytesAvailable As Currency
  Dim cTotalBytes As Currency
  Dim cTotalBytesFree As Currency
  
  Dim dSectorsPerCluster As Double
  Dim dBytesPerSector As Double
  Dim dTotalNumberOfSectors As Double
  Dim dNumberOfFreeSectors As Double
  
  Dim lRet As Long
  
Method1: ' Try this first. (Only works on: Windows NT 4.0 or later; Windows 95 OSR2 or later)
  lRet = APIGetDiskFreeSpaceEx(sDrive, cFreeBytesAvailable, cTotalBytes, cTotalBytesFree)
  CheckForError
  If lRet > 0 Then
    Select Case Info
      Case DriveSizeAvailable
        GetDriveSize = cFreeBytesAvailable * 10000
      Case DriveSizeFree
        GetDriveSize = cTotalBytesFree * 10000
      Case DriveSizeTotal
        GetDriveSize = cTotalBytes * 10000
    End Select
  Else
Method2: '(Works on: Windows NT 4.0 or later; Windows 95) (Drive needs to be: "C:\")
On Error GoTo 0 ' Turn error handling off (we don't want a recursive error loop)
    lRet = APIGetDiskFreeSpace(HandleSlash(sDrive), lSectorsPerCluster, lBytesPerSector, lNumberOfFreeSectors, lTotalNumberOfSectors)
    CheckForError
    If lRet > 0 Then
      dSectorsPerCluster = lSectorsPerCluster
      dBytesPerSector = lBytesPerSector
      dTotalNumberOfSectors = lTotalNumberOfSectors
      dNumberOfFreeSectors = lNumberOfFreeSectors
      Select Case Info
        Case DriveSizeAvailable
          GetDriveSize = dNumberOfFreeSectors * dSectorsPerCluster * dBytesPerSector
        Case DriveSizeFree
          GetDriveSize = dNumberOfFreeSectors * dSectorsPerCluster * dBytesPerSector
        Case DriveSizeTotal
          GetDriveSize = dTotalNumberOfSectors * dSectorsPerCluster * dBytesPerSector
      End Select
    End If
  End If

Exit Function
ErrHandler:
  Select Case Err.Number
    Case 453 ' Entry point not found in DLL
      Resume Method2
    Case Else
      Err.Raise Err.Number, , Err.Description
  End Select
End Function

Public Function GetDriveLetter(FilePath As String)
  
  If FilePath Like "[A-Z]:*" Then
    GetDriveLetter = VBA.Left(FilePath, 1)
  ElseIf FilePath Like "[A-Z]" Then
    GetDriveLetter = FilePath
  ElseIf FilePath Like "[a-z]:*" Then
    GetDriveLetter = VBA.UCase(VBA.Left(FilePath, 1))
  ElseIf FilePath Like "[a-z]" Then
    GetDriveLetter = VBA.UCase(FilePath)
  Else
    GetDriveLetter = ""
  End If
  
End Function

Public Function GetDriveType(ByVal Drive As String) As FileSystemAPI.DriveTypeConst
  Dim lRet As Long
  
  lRet = APIGetDriveType(Drive)
  CheckForError
  Select Case lRet
    Case 2
      GetDriveType = Removable
    Case 3
      GetDriveType = Fixed
    Case 4
      GetDriveType = Remote
    Case 5
      GetDriveType = CDRom
    Case 6
      GetDriveType = RamDisk
    Case Else
      GetDriveType = UnkownType
  End Select
  
End Function

Public Function GetDriveInfo(ByVal Drive As String, ByVal Info As FileSystemAPI.DriveInfo) As Variant
  
  Dim lSerial As Long
  Dim sFileSystem As String
  Dim sVolumeName As String
  
  'Create buffers
  sFileSystem = String$(255, Chr$(0))
  sVolumeName = String$(255, Chr$(0))

  Call APIGetVolumeInformation(HandleSlash(Drive), sVolumeName, 255, lSerial, 0, 0, sFileSystem, 255)
  CheckForError
  
  Select Case Info
    Case DriveInfoFileSystem
      'sFileSystem = VBA.Replace(sFileSystem, Chr$(0), "", , , vbTextCompare)
      sFileSystem = StripTerminator(sFileSystem, Chr$(0), True)
      GetDriveInfo = sFileSystem
    Case DriveInfoVolumeName
      'sVolumeName = VBA.Replace(sVolumeName, Chr$(0), "", , , vbTextCompare)
      sVolumeName = StripTerminator(sVolumeName, Chr$(0), True)
      GetDriveInfo = sVolumeName
    Case DriveSerialNumber
      GetDriveInfo = lSerial
  End Select
  
End Function

Public Sub SetVolumeLabel(ByVal Drive As String, ByVal Label As String)
  
  Call APISetVolumeLabel(HandleSlash(Drive), Label)
  CheckForError
  
End Sub

Public Function DriveIsReady(ByVal Drive As String) As Boolean
On Error GoTo ErrHandler
  
  Call GetDriveInfo(Drive, DriveInfoFileSystem)
  DriveIsReady = True

ExitProc:
Exit Function
ErrHandler:
  Select Case Err.Number
    Case 71 ' Disk not ready
      DriveIsReady = False
      Resume ExitProc
    Case Else ' Some other error
      Err.Raise Err.Number, , Err.Description
  End Select
End Function


Function DriveLetterToUNC(ByVal Drive As String) As String
On Error GoTo ErrHandler

  Dim NetInfo(1023) As NETRESOURCE
  
  Dim sLocalName As String
  Dim sUNCName As String
  
  Dim lEnum As Long
  Dim lEntries As Long
  Dim lStatus As Long
  
  Dim lCount As Long
  Dim lRet As Long

  ' Begin the enumeration
  lStatus = APIWNetOpenEnum(RESOURCE_CONNECTED, RESOURCETYPE_ANY, 0&, ByVal 0&, lEnum) ' Throws I/O Error.

  DriveLetterToUNC = ""

  If ((lStatus = 0) And (lEnum <> 0)) Then 'Check for success from open enum
    lEntries = 1024 ' Set number of Entries

    ' Enumerate the resource
    lStatus = APIWNetEnumResource(lEnum, lEntries, NetInfo(0), CLng(Len(NetInfo(0))) * 1024) ' Throws File not found error.

    If lStatus = 0 Then ' Check for success
      For lCount = 0 To lEntries - 1
        ' Get the local name
        sLocalName = ""
        If NetInfo(lCount).lpLocalName <> 0 Then
          sLocalName = Space(APIlstrlen(NetInfo(lCount).lpLocalName) + 1)
          CheckForError
          lRet = APIlstrcpy(sLocalName, NetInfo(lCount).lpLocalName)
          CheckForError
        End If

        If Len(sLocalName) <> 0 Then ' Strip null character from end
          sLocalName = Left(sLocalName, (Len(sLocalName) - 1))
        End If

        If UCase$(sLocalName) = UCase$(Drive) Then
          ' Get the remote name
          sUNCName = ""
          If NetInfo(lCount).lpRemoteName <> 0 Then
            sUNCName = Space(APIlstrlen(NetInfo(lCount).lpRemoteName) + 1)
            CheckForError
            lRet = APIlstrcpy(sUNCName, NetInfo(lCount).lpRemoteName)
            CheckForError
          End If

          If Len(sUNCName) <> 0 Then ' Strip null character from end
            sUNCName = Left(sUNCName, (Len(sUNCName) - 1))
          End If

          DriveLetterToUNC = sUNCName ' Return the UNC path to drive

          Exit For ' Exit the loop
        End If
      Next lCount
    End If
  End If

  lStatus = APIWNetCloseEnum(lEnum) ' End enumeration
  CheckForError
  
Exit Function
ErrHandler:
  If lEnum <> 0 Then
    Call APIWNetCloseEnum(lEnum)
  End If
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetPath(ByVal FilePath As String, Optional ByVal IncludeBackSlash As Boolean = True) As String
  
  Dim i As Integer

    i = InStrRev(FilePath, "\")
    
    If i = 0 Then
      i = InStrRev(FilePath, "/")
    End If
    
    If i > 0 Then
      FilePath = VBA.Left(FilePath, i)
    Else
      FilePath = FilePath
    End If
  
  GetPath = HandleSlash(FilePath, IncludeBackSlash)
  
End Function

Public Function GetFileName(ByVal FilePath As String) As String
   
  Dim i As Integer
  Dim j As Integer

  i = InStrRev(FilePath, "\")
  j = InStrRev(FilePath, "/")

  If i < j Then
    i = j
  End If

  If i > 0 Then
    GetFileName = VBA.Right(FilePath, VBA.Len(FilePath) - i)
  Else
    GetFileName = FilePath
  End If
  
End Function

Public Function GetFileBaseName(ByVal FilePath As String) As String
  
  Dim i As Integer
  Dim j As Integer
  
  GetFileBaseName = ""
  
  If VBA.InStr(1, FilePath, "\\") = 0 Then
    If VBA.InStr(1, FilePath, "//") = 0 Then
      If VBA.InStr(1, FilePath, "/\") = 0 Then
        If VBA.InStr(1, FilePath, "\/") = 0 Then
        
          i = InStrRev(FilePath, "\")
          j = InStrRev(FilePath, "/")
          
          If i = 0 Then
            i = j
          ElseIf j = 0 Then
            ' nothing
          Else
            If i < j Then
              i = j
            End If
          End If
          
          
          j = InStrRev(FilePath, ".")
          
          If i > 0 And j > 0 Then
            GetFileBaseName = VBA.Mid(FilePath, i + 1, j - i - 1)
          ElseIf i > 0 Then
            GetFileBaseName = VBA.Right(FilePath, VBA.Len(FilePath) - i)
          ElseIf j > 0 Then
            GetFileBaseName = VBA.Left(FilePath, j - 1)
          Else
            GetFileBaseName = FilePath
          End If
          
        End If
      End If
    End If
  End If
  
End Function

Public Function GetParentFolderName(ByVal FilePath As String) As String

  Dim i As Integer
  
  If FilePath Like "//*/*/*" Or FilePath Like "\\*\*\*" Then
  
    If VBA.InStr(1, FilePath, "//") > 0 Then
      Do
        FilePath = VBA.Replace(FilePath, "\\", "\")
      Loop Until VBA.InStr(1, FilePath, "\\") = 0
      
      Do
        FilePath = VBA.Replace(FilePath, "//", "/")
      Loop Until VBA.InStr(1, FilePath, "//") = 0
      
      FilePath = "/" & FilePath
    ElseIf VBA.InStr(1, FilePath, "\\") > 0 Then
      Do
        FilePath = VBA.Replace(FilePath, "\\", "\")
      Loop Until VBA.InStr(1, FilePath, "\\") = 0
      
      Do
        FilePath = VBA.Replace(FilePath, "//", "/")
      Loop Until VBA.InStr(1, FilePath, "//") = 0
      
      FilePath = "\" & FilePath
    End If
    
    If VBA.Right(FilePath, 1) = "\" Or VBA.Right(FilePath, 1) = "/" Then
      FilePath = VBA.Left(FilePath, VBA.Len(FilePath) - 1)
    End If
    
    If (VBA.InStr(1, FilePath, "\") > 0) Or (VBA.InStr(1, FilePath, "/") > 0) Then
      For i = VBA.Len(FilePath) To 1 Step -1
        If VBA.Mid(FilePath, i, 1) = "\" Or VBA.Mid(FilePath, i, 1) = "/" Then
          Exit For
        End If
      Next
      FilePath = VBA.Left(FilePath, i - 1)
    Else
      FilePath = ""
    End If
    
  ElseIf FilePath Like "[a-z]:\*" Or FilePath Like "[A-Z]:\*" Then
  
    If VBA.InStr(1, FilePath, "\\") = 0 Then
      If VBA.InStr(1, FilePath, "//") = 0 Then
        If VBA.InStr(1, FilePath, "/\") = 0 Then
          If VBA.InStr(1, FilePath, "\/") = 0 Then
    
            If VBA.Right(FilePath, 1) = "\" Or VBA.Right(FilePath, 1) = "/" Then
              FilePath = VBA.Left(FilePath, VBA.Len(FilePath) - 1)
            End If
            
            If (VBA.InStr(1, FilePath, "\") > 0) Or (VBA.InStr(1, FilePath, "/") > 0) Then
              For i = VBA.Len(FilePath) To 1 Step -1
                If VBA.Mid(FilePath, i, 1) = "\" Or VBA.Mid(FilePath, i, 1) = "/" Then
                  Exit For
                End If
              Next
              FilePath = VBA.Left(FilePath, i - 1)
            Else
              FilePath = ""
            End If
            
          End If
        End If
      End If
    End If
  Else
    FilePath = ""
  End If
  
  FilePath = HandleSlash(FilePath, False)
  
  GetParentFolderName = FilePath

End Function


Public Function GetFileExtension(ByVal FilePath As String) As String
  
  Dim i As Integer
  
  GetFileExtension = ""
  
  If VBA.InStr(1, FilePath, "\\") = 0 Then
    If VBA.InStr(1, FilePath, "//") = 0 Then
      If VBA.InStr(1, FilePath, "/\") = 0 Then
        If VBA.InStr(1, FilePath, "\/") = 0 Then
        
          i = InStrRev(FilePath, ".")
          
          If i > 0 Then
            GetFileExtension = VBA.Right(FilePath, VBA.Len(FilePath) - i)
          Else
            GetFileExtension = ""
          End If
          
        End If
      End If
    End If
  End If
  
End Function

Public Function GetShortPath(ByVal FilePath As String) As String
  Dim lRet As Long
  Dim sPath As String
  FilePath = GetAbsolutePathName(FilePath, False)
  sPath = String$(MAX_PATH, 0)
  lRet = APIGetShortPathName(FilePath, sPath, MAX_PATH)
  CheckForError
  GetShortPath = Left$(sPath, lRet)
End Function

Public Function GetFileSize(FilePath As String) As Variant
On Error GoTo ErrHandler

  Dim lFile As Long
  
  FilePath = GetAbsolutePathName(FilePath, False)
  'lFile = APICreateFile(FilePath, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, 0, 0)
  lFile = APICreateFile(FilePath, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
  CheckForError
  
  GetFileSize = APIGetFileSize(lFile, 0)
  CheckForError
  
  Call APICloseHandle(lFile)
  lFile = 0
  CheckForError
  
Exit Function
ErrHandler:
  If lFile > 0 Then
    Call APICloseHandle(lFile)
  End If
  Err.Raise Err.Number, , Err.Description
End Function

Public Function GetFileTypeName(ByVal FilePath As String) As String
  Dim FI As SHFILEINFO
  FilePath = GetAbsolutePathName(FilePath, False)
  Call APISHGetFileInfo(FilePath, 0, FI, Len(FI), SHGFI_TYPENAME)
  CheckForError
  GetFileTypeName = StripTerminator(FI.szTypeName, VBA.Chr(0))
End Function

Public Function StripTerminator(ByVal Text As String, Optional ByVal Char As String = vbNullChar, Optional ByVal StartFromEnd As Boolean = True) As String
  
  Dim lCharLen As Long
  
  If VBA.Len(Text) > 0 And VBA.Len(Char) > 0 Then
    If StartFromEnd Then
      lCharLen = VBA.Len(Char)
      Do Until VBA.Right(Text, lCharLen) <> Char
        Text = VBA.Left(Text, VBA.Len(Text) - lCharLen)
        If (VBA.Len(Text) = 0) Then
          Exit Do
        End If
      Loop
    Else
      If VBA.Len(Text) > 1 Then
        Text = VBA.Left(Text, VBA.InStr(1, Text, Char) - 1)
      End If
    End If
  End If
  
  StripTerminator = Text
  
End Function



'Public Sub GetFilesLike(ByRef colIn As Collection, ByVal Path As String, Optional FileLike As String = "*")
'
'  Dim WFD As WIN32_FIND_DATA
'  Dim sFileName As String
'  Dim lSearch As Long
'  Dim lMoreFiles As Long
'  Dim oFile As FileSystemAPI.File
'
'  'Set GetFilesLike = New Collection
'
'  Path = GetAbsolutePathName(Path, True)
'  lSearch = APIFindFirstFile(Path & FileLike, WFD)
'
'  lMoreFiles = 1
'  If lSearch <> INVALID_HANDLE_VALUE Then
'    Do While lMoreFiles > 0
'      sFileName = StripTerminator(WFD.cFileName, Chr(0), False)
'      If (sFileName <> ".") And (sFileName <> "..") Then
'        If (APIGetFileAttributes(Path & sFileName) And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
'          Set oFile = New FileSystemAPI.File
'          oFile.Path = Path & sFileName
'          colIn.Add oFile, oFile.Path
'        End If
'      End If
'      lMoreFiles = APIFindNextFile(lSearch, WFD) ' Get next file
'      CheckForError
'    Loop
'  End If
'
'  Call APIFindClose(lSearch)
'  CheckForError
'
'End Sub

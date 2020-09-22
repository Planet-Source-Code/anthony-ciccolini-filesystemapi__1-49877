VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSystemObject"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'========================================================================
'=  FileSystemAPI
'=
'=  A VB6 clone of Microsoft's Scripting Runtime: FileSystemObject (sccrun.dll).
'=
'=  Description: A VB6 clone of Microsoft's Scripting Runtime: FileSystemObject
'=               using API calls. Adds some additional functionality as well.
'=
'=  Notes: Needs some more bug testing.
'=         Made FileSystemObject: GlobalMultiUse
'=         Added the functions FindFiles() and FindFolders()
'=         Added more special folders
'=         Certain properties/methods had to be renamed for VB6:
'=           .Type    -> .TypeDesc
'=           .Write   -> .WriteText
'=           .Close   -> .CloseFile
'=
'=  Limitations: ???
'=
'=  Usage: File access and control
'=
'=  Dependencies:
'=    1) VB6 Runtimes
'=    2) OLE Automation
'=
'=  Author:   Anthony Ciccolini
'=  Email:    Anthony_Ciccolini@Yahoo.com
'=  Date:     11/13/2003
'=
'=  Legal: If you use this I would enjoy feed back and suggestions.
'=  This may not be sold in whole or part without my written permission.
'=  You are free to use/modify this tool for development purposes.
'=
'========================================================================
Option Explicit

Public Enum StandardStreamTypes
  StdIn = 0
  StdOut = 1
  StdErr = 2
End Enum

Public Enum SpecialFolderConst
  AdminToolsFolder = &H30            '{user}\Start Menu\Programs\Administrative Tools
  AltStartupFolder = &H1D            'non localized startup
  AppDataFolder = &H1A               '{user}\Application Data
  BitBucketFolder = &HA              '{desktop}\Recycle Bin
  CookiesFolder = &H21               '{user}\Local Settings\History
  DesktopFolder = &H0                '{namespace root}
  DesktopDirectoryFolder = &H10      '{user}\Desktop
  FavoritesFolder = &H6              '{user}\Favourites
  FontsFolder = &H14                 'windows\fonts
  HistoryFolder = &H22               '{user}\Local Settings\History
  InternetCacheFolder = &H20         'Internet Cache folder
  LocalAppDataFolder = &H1C          '{user}\Local Settings\Application Data (non roaming)
  MyDocumentsFolder = &H5            'My Documents
  MyPicturesFolder = &H27            'C:\Program Files\My Pictures
  NetHoodFolder = &H13               '{user}\nethood
  PrintHoodFolder = &H1B             '{user}\PrintHood
  ProgramFilesFolder = &H26          'Program Files folder
  ProgramFilesx86Folder = &H2A       'Program Files folder for x86 apps (Alpha)
  ProgramsFolder = &H2               'Start Menu\Programs
  ProgramFilesCommonFolder = &H2B    'Program Files\Common
  ProgramFilesCommonx86Folder = &H2C 'x86 \Program Files\Common on RISC
  RecentFolder = &H8                 '{user}\Recent
  SendToFolder = &H9                 '{user}\SendTo
  StartMenuFolder = &HB              '{user}\Start Menu
  StartUpFolder = &H7                'Start Menu\Programs\Startup
  SystemFolder = &H25                'system folder
  Systemx86Folder = &H29             'system folder for x86 apps (Alpha)
  TemplatesFolder = &H15
  ProfileFolder = &H28               'user's profile folder
  WindowsFolder = &H24               'Windows directory or SYSROOT()
  CommonAdminToolsFolder = &H2F      '(all users)\Start Menu\Programs\Administrative Tools
  CommonAltStartupFolder = &H1E      'non localized common startup
  CommonAppDataFolder = &H23         '(all users)\Application Data
  CommonDesktopDirectoryFolder = &H19 '(all users)\Desktop
  CommonDocumentsFolder = &H2E       '(all users)\Documents
  CommonFavoritesFolder = &H1F       '(all users)\Favourites
  CommonProgramsFolder = &H17        '(all users)\Programs
  CommonStartupFolder = &H16         '(all users)\Start Menu
  CommonTemplatesFolder = &H2D       '(all users)\Templates
  FlagCreateFolder = &H8000&         'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
  FlagDontVerifyFolder = &H4000      'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
  FlagMaskFolder = &HFF00            'mask for all possible flag values
  TemporaryFolder = -1

'  Virtual Folders: (not sure how to handle them)
'  ControlPanelFolder = &H3           'My Computer\Control Panel
'  DrivesFolder = &H11                'My Computer
'  NetworkFolder = &H12               'Network Neighbourhood
'  InternetFolder = &H1               'Internet virtual folder
'  PrintersFolder = &H4               'My Computer\Printers

End Enum

Public Enum IOMode
  ForReading = 1
  ForWriting = 2
  ForAppending = 8
End Enum

Public Enum Tristate
  TristateUseDefault = -2
  TristateMixed = -2
  TristateFalse = 0
  TristateTrue = -1
End Enum


Private m_Drives As New FileSystemAPI.Drives

Public Function BuildPath(Path As String, Name As String) As String
  BuildPath = mFunctions.BuildPath(Path, Name)
End Function

Public Sub CopyFile(Source As String, Destination As String, Optional OverWriteFiles As Boolean = True)
  Call mFunctions.CopyFile(Source, Destination, OverWriteFiles)
End Sub

Public Sub CopyFolder(Source As String, Destination As String, Optional OverWriteFiles As Boolean = True)
  Call mFunctions.CopyFolder(Source, Destination, OverWriteFiles)
End Sub

Public Function CreateFolder(Path As String) As FileSystemAPI.Folder
  Set CreateFolder = mFunctions.CreateFolder(Path)
End Function

Public Function CreateTextFile(FileName As String, Optional Overwrite As Boolean = True, Optional Unicode As Boolean = False) As FileSystemAPI.TextStream
  Set CreateTextFile = mFunctions.OpenTextFile(FileName, ForWriting, True, Overwrite, VBA.IIf(Unicode, TristateTrue, TristateFalse))
End Function

Public Sub DeleteFile(FileSpec As String, Optional Force As Boolean = False)
  Call mFunctions.DeleteFile(FileSpec, Force)
End Sub

Public Sub DeleteFolder(FolderSpec As String, Optional Force As Boolean = False)
  Call mFunctions.DeleteFolder(FolderSpec, Force)
End Sub

Public Function DriveExists(DriveSpec As String) As Boolean
  DriveExists = mFunctions.DriveExists(DriveSpec)
End Function

Public Property Get Drives() As FileSystemAPI.Drives
  Set Drives = m_Drives
End Property

Public Function FileExists(FileSpec As String) As Boolean
  FileExists = mFunctions.FileExists(FileSpec)
End Function

Public Function FolderExists(FolderSpec As String) As Boolean
  FolderExists = mFunctions.FolderExists(FolderSpec)
End Function

Public Function GetAbsolutePathName(Path As String) As String
  GetAbsolutePathName = mFunctions.GetAbsolutePathName(Path)
End Function

Public Function GetBaseName(Path As String) As String
  GetBaseName = mFunctions.GetFileBaseName(Path)
End Function

Public Function GetDrive(DriveSpec As String) As FileSystemAPI.Drive
  Set GetDrive = mFunctions.GetDrive(DriveSpec)
End Function

Public Function GetDriveName(Path As String) As String
  GetDriveName = mFunctions.GetDrive(Path).Path
End Function

Public Function GetExtensionName(Path As String) As String
  GetExtensionName = mFunctions.GetFileExtension(Path)
End Function

Public Function GetFile(FilePath As String) As FileSystemAPI.File
  Set GetFile = mFunctions.GetFile(FilePath)
End Function

Public Function GetFileName(Path As String) As String
  GetFileName = mFunctions.GetFileName(Path)
End Function

Public Function GetFileVersion(FileName As String) As String
  GetFileVersion = mFunctions.GetFileVersion(FileName)
End Function

Public Function GetFolder(FolderPath As String) As FileSystemAPI.Folder
  Set GetFolder = mFunctions.GetFolder(FolderPath)
End Function

Public Function GetParentFolderName(Path As String) As String
  GetParentFolderName = mFunctions.GetParentFolderName(Path)
End Function

Public Function GetSpecialFolder(SpecialFolder As FileSystemAPI.SpecialFolderConst) As Folder
  Set GetSpecialFolder = mFunctions.GetSpecialFolder(SpecialFolder)
End Function

Public Function GetStandardStream(StandardStreamType As FileSystemAPI.StandardStreamTypes, Optional Unicode As Boolean) As FileSystemAPI.TextStream
  Set GetStandardStream = mFunctions.OpenTextFile("", ForAppending, True, True, VBA.IIf(Unicode, TristateTrue, TristateFalse), True, StandardStreamType)
End Function

Public Function GetTempName() As String
  GetTempName = mFunctions.GetTempName
End Function

Public Sub MoveFile(Source As String, Destination As String)
  Call mFunctions.MoveFileOrFolder(Source, Destination)
End Sub

Public Sub MoveFolder(Source As String, Destination As String)
  Call mFunctions.MoveFileOrFolder(Source, Destination)
End Sub

Public Function OpenTextFile(FileName As String, Optional IOMode As FileSystemAPI.IOMode = ForReading, Optional Create As Boolean = False, Optional Format As FileSystemAPI.Tristate = TristateFalse) As TextStream
  Set OpenTextFile = mFunctions.OpenTextFile(FileName, IOMode, Create, False, Format)
End Function

Public Function FindFiles(FilePath As String, Optional FileLike As String = "*", Optional Recursive As Boolean = True, Optional FirstFound As Boolean = False) As FileSystemAPI.Files
  Set FindFiles = mFunctions.FindFiles(FilePath, FileLike, Recursive, FirstFound)
End Function

Public Function FindFolders(FilePath As String, Optional FolderLike As String = "*", Optional Recursive As Boolean = True, Optional FirstFound As Boolean = False) As FileSystemAPI.Folders
  Set FindFolders = mFunctions.FindFolders(FilePath, FolderLike, Recursive, FirstFound)
End Function

Private Sub Class_Terminate()
  Set m_Drives = Nothing
End Sub



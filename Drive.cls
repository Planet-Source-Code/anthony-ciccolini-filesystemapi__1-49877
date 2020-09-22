VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Drive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum DriveTypeConst
  UnkownType = 0
  Removable = 1
  Fixed = 2
  Remote = 3
  CDRom = 4
  RamDisk = 5
End Enum

Private m_SubFolders As New FileSystemAPI.Folder
Private m_sPath As String

Public Property Get AvailableSpace() As Variant
  AvailableSpace = mFunctions.GetDriveSize(Me.Path, DriveSizeAvailable)
End Property

Public Property Get DriveLetter() As String
  DriveLetter = mFunctions.GetDriveLetter(Me.Path)
End Property

Public Property Get DriveType() As DriveTypeConst
  DriveType = mFunctions.GetDriveType(Me.Path)
End Property

Public Property Get FileSystem() As String
  FileSystem = mFunctions.GetDriveInfo(Me.Path, DriveInfoFileSystem)
End Property

Public Property Get FreeSpace() As Variant
  FreeSpace = mFunctions.GetDriveSize(Me.Path, DriveSizeFree)
End Property

Public Property Get IsReady() As Boolean
  IsReady = mFunctions.DriveIsReady(Me.Path)
End Property

Public Property Get Path() As String ' Default
Attribute Path.VB_UserMemId = 0
  If mFunctions.PathIsRoot(m_sPath) Then
    Path = m_sPath
  Else '??????????????
    Err.Raise 53, , "File not found"
  End If
End Property

Friend Property Let Path(sPath As String)
  m_sPath = PathToRoot(sPath, False)
End Property

Public Property Get RootFolder() As FileSystemAPI.Folder
  If m_SubFolders.Path <> Me.Path Then
    m_SubFolders.Path = Me.Path
  End If
  Set RootFolder = m_SubFolders
End Property

Public Property Get SerialNumber() As Long
  SerialNumber = mFunctions.GetDriveInfo(Me.Path, DriveSerialNumber)
End Property

Public Property Get ShareName() As String
  ShareName = DriveLetterToUNC(Me.Path)
End Property

Public Property Get TotalSize() As Variant
  TotalSize = mFunctions.GetDriveSize(Me.Path, DriveSizeTotal)
End Property

Public Property Get VolumeName() As String
  VolumeName = mFunctions.GetDriveInfo(Me.Path, DriveInfoVolumeName)
End Property

Public Property Let VolumeName(Name As String)
  Call mFunctions.SetVolumeLabel(Me.Path, Name)
End Property

Private Sub Class_Terminate()
  Set m_SubFolders = Nothing
End Sub

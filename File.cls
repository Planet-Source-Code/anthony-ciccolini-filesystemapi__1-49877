VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum FileAttribute
  Normal = 0
  ReadOnly = 1
  Hidden = 2
  System = 4
  Volume = 8
  Directory = 16
  Archive = 32
  Alias = 1024
  Compressed = 2048
End Enum

Private m_sPath As String

Public Property Get Attributes() As FileSystemAPI.FileAttribute
  Attributes = mFunctions.FileOrFolderAttributesGet(Me.Path)
End Property

Public Property Let Attributes(FileAtt As FileSystemAPI.FileAttribute)
  Call mFunctions.FileOrFolderAttributesSet(Me.Path, FileAtt)
End Property

Public Sub Copy(Destination As String, Optional OverWriteFiles As Boolean = True)
  Call mFunctions.CopyFile(Me.Path, Destination, OverWriteFiles)
End Sub

Public Property Get DateCreated() As Date
  DateCreated = mFunctions.FileOrFolderDateGet(Me.Path, DateCreate)
End Property

Public Property Let DateCreated(dtDate As Date)
  Call mFunctions.FileOrFolderDateSet(Me.Path, dtDate, DateCreate)
End Property

Public Property Get DateLastAccessed() As Date
  DateCreated = mFunctions.FileOrFolderDateGet(Me.Path, DateAccess)
End Property

Public Property Let DateLastAccessed(dtDate As Date)
  Call mFunctions.FileOrFolderDateSet(Me.Path, dtDate, DateAccess)
End Property

Public Property Get DateLastModified() As Date
  DateCreated = mFunctions.FileOrFolderDateGet(Me.Path, DateWrite)
End Property

Public Property Let DateLastModified(dtDate As Date)
  Call mFunctions.FileOrFolderDateSet(Me.Path, dtDate, DateWrite)
End Property

Public Sub Delete(Optional Force As Boolean = False)
  Call mFunctions.DeleteFile(Me.Path, Force)
End Sub

Public Property Get Drive() As Drive
  Set Drive = mFunctions.GetDrive(Me.Path)
End Property

Public Sub Move(Destination As String)
  Me.Path = mFunctions.MoveFileOrFolder(Me.Path, Destination)
End Sub

Public Property Get Name() As String
  Name = mFunctions.GetFileName(Me.Path)
End Property

Public Property Let Name(sName As String)
  Me.Path = mFunctions.RenameFileOrFolder(Me.Path, sName)
End Property

Public Function OpenAsTextStream(Optional IOMode As FileSystemAPI.IOMode = ForReading, Optional Format As FileSystemAPI.Tristate = FileSystemAPI.TristateFalse) As FileSystemAPI.TextStream
  Set OpenAsTextStream = mFunctions.OpenTextFile(Me.Path, IOMode, False, False, Format)
End Function

Public Property Get ParentFolder() As FileSystemAPI.Folder
  Set ParentFolder = mFunctions.GetFolder(mFunctions.GetPath(Me.Path, True))
End Property

Public Property Get Path() As String ' Default
Attribute Path.VB_UserMemId = 0
  If FileExists(m_sPath) Then
    Path = m_sPath
  Else
    Err.Raise 53, , "File not found"
  End If
End Property

Friend Property Let Path(sPath As String)
  m_sPath = sPath
End Property

Public Property Get ShortName() As String
  ShortName = mFunctions.GetFileName(mFunctions.GetShortPath(Me.Path))
End Property

Public Property Get ShortPath() As String
  ShortPath = mFunctions.GetShortPath(Me.Path)
End Property

Public Property Get Size() As Variant
  Size = mFunctions.GetFileSize(Me.Path)
End Property

Public Property Get TypeDesc() As String ' Renamed from: Type()
  TypeDesc = mFunctions.GetFileTypeName(Me.Path)
End Property


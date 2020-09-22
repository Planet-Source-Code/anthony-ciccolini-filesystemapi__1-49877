VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Folder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_sPath As String

Private m_SubFolders As New FileSystemAPI.Folders
Private m_Files As New FileSystemAPI.Files

Public Property Get Attributes() As FileSystemAPI.FileAttribute
  Attributes = mFunctions.FileOrFolderAttributesGet(Me.Path)
End Property

Public Property Let Attributes(FileAtt As FileSystemAPI.FileAttribute)
  Call mFunctions.FileOrFolderAttributesSet(Me.Path, FileAtt)
End Property

Public Sub Copy(Destination As String, Optional OverWriteFiles As Boolean = True)
  Call mFunctions.CopyFolder(m_sPath, Destination, OverWriteFiles)
End Sub

Public Function CreateTextFile(FileName As String, Optional Overwrite As Boolean = True, Optional Unicode As Boolean = False) As FileSystemAPI.TextStream
  Set CreateTextFile = mFunctions.OpenTextFile(mFunctions.BuildPath(Me.Path, FileName), ForWriting, True, Overwrite, VBA.IIf(Unicode, TristateTrue, TristateFalse))
End Function

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
  Call mFunctions.DeleteFolder(Me.Path, Force)
End Sub

Public Property Get Drive() As Drive
  Set Drive = mFunctions.GetDrive(Me.Path)
End Property

Public Property Get Files() As FileSystemAPI.Files
  If m_Files.Path <> Me.Path Then
    m_Files.Path = Me.Path
  End If
  Set Files = m_Files
End Property

Public Property Get IsRootFolder() As Boolean
  IsRootFolder = mFunctions.PathIsRoot(Me.Path)
End Property

Public Sub Move(Destination As String)
  Me.Path = mFunctions.MoveFileOrFolder(Me.Path, mFunctions.GetAbsolutePathName(Destination, True) & Me.Name)
End Sub

Public Property Get Name() As String
  Name = mFunctions.GetFolderName(Me.Path)
End Property

Public Property Let Name(sName As String)
  Me.Path = mFunctions.RenameFileOrFolder(Me.Path, sName)
End Property

Public Property Get ParentFolder() As FileSystemAPI.Folder
  Set ParentFolder = mFunctions.GetFolder(mFunctions.GetPath(Me.Path, True))
End Property

Public Property Get Path() As String
Attribute Path.VB_UserMemId = 0
  If FolderExists(m_sPath) Then
    Path = m_sPath
  Else
    Err.Raise 53, , "File not found"
  End If
End Property

Friend Property Let Path(sPath As String)
  m_sPath = sPath
End Property

Public Property Get ShortName() As String
  ShortName = mFunctions.GetFolderName(mFunctions.GetShortPath(Me.Path))
End Property

Public Property Get ShortPath() As String
  ShortPath = mFunctions.GetShortPath(Me.Path)
End Property

Public Property Get Size() As Variant
  Size = GetFolderSize(Me.Path)
End Property

Public Property Get SubFolders() As FileSystemAPI.Folders
  If m_SubFolders.Path <> Me.Path Then
    m_SubFolders.Path = Me.Path
  End If
  Set SubFolders = m_SubFolders
End Property

Public Property Get TypeDesc() As String ' Renamed from Type()
  TypeDesc = mFunctions.GetFileTypeName(Me.Path)
End Property

Private Sub Class_Terminate()
  Set m_SubFolders = Nothing
  Set m_Files = Nothing
End Sub

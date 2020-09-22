VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Collection As Collection

Private m_sPath As String
Private m_sFileLike As String
Private m_bRecursive As Boolean
Private m_bFirstFound  As Boolean

Public Property Get Count() As Long
  Count = m_Collection.Count
End Property

Public Property Get Item(Key As Variant) As FileSystemAPI.File ' Default
Attribute Item.VB_UserMemId = 0
  Set Item = m_Collection(VBA.Val(Key))
End Property

Friend Property Get Path() As String
  Path = m_sPath
End Property

Friend Property Let Path(sPath As String)
  m_sPath = sPath
  Set m_Collection = GetFileCollection(m_sPath, m_sFileLike, m_bRecursive, m_bFirstFound)
End Property

Friend Property Get FileLike() As String
  FileLike = m_sFileLike
End Property

Friend Property Let FileLike(sFileLike As String)
  m_sFileLike = sFileLike
End Property

Friend Property Get Recursive() As Boolean
  Recursive = m_bRecursive
End Property

Friend Property Let Recursive(bRecursive As Boolean)
  m_bRecursive = bRecursive
End Property

Friend Property Get FirstFound() As Boolean
  FirstFound = m_bFirstFound
End Property

Friend Property Let FirstFound(bFirstFound As Boolean)
  m_bFirstFound = bFirstFound
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  ' Goto (Tools > Procedure Attributes... > Advanced>>)
  ' Make this function's Procedure ID = -4 and hidden.
  ' Allows the nifty For..Each structure to work.
  Set NewEnum = m_Collection.[_NewEnum]   ' Calls the hidden function.
End Function

Private Sub Class_Initialize()
  m_sFileLike = "*"
  m_bRecursive = False
  m_bFirstFound = False
End Sub

Private Sub Class_Terminate()
  Set m_Collection = Nothing
End Sub

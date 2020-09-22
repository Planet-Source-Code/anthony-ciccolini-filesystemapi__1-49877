VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Drives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Collection As Collection

Public Property Get Count() As Long
  Count = m_Collection.Count
End Property

Public Property Get Item(Key As Variant) As FileSystemAPI.Drive ' Default
Attribute Item.VB_UserMemId = 0
  If VBA.VarType(Key) = vbString Then
    Set Item = m_Collection(PathToRoot(VBA.CStr(Key), False))
  Else
    Set Item = m_Collection(VBA.Val(Key))
  End If
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
  Set m_Collection = GetDriveCollection()
End Sub

Private Sub Class_Terminate()
  Set m_Collection = Nothing
End Sub

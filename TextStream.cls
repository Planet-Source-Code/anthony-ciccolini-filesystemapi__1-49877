VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Tristate As FileSystemAPI.Tristate
Private m_Mode As FileSystemAPI.IOMode

Private m_lFile As Long
Private m_lLocation As Long

Public Property Get AtEndOfLine() As Boolean
  If m_Mode = ForReading Then
    AtEndOfLine = mFunctions.AtEndOfLine(m_lFile, m_lLocation, m_Tristate)
  Else
    Err.Raise 54, , "Bad file mode"
  End If
End Property

Public Property Get AtEndOfStream() As Boolean
  If m_Mode = ForReading Then
    AtEndOfStream = mFunctions.AtEndOfStream(m_lFile, m_lLocation, m_Tristate)
  Else
    Err.Raise 54, , "Bad file mode"
  End If
End Property

Public Sub CloseFile() ' Renamed from: Close()
  Call mFunctions.CloseFile(m_lFile)
End Sub

Public Property Get Column() As Long
  Column = mFunctions.GetStreamColumn(m_lFile, m_lLocation, m_Tristate)
End Property

Public Property Get Line() As Long
  Line = mFunctions.GetStreamLine(m_lFile, m_lLocation, m_Tristate)
End Property

Public Function Read(Characters As Long) As String
  If m_Mode = ForReading Then
    Read = mFunctions.ReadChars(m_lFile, m_lLocation, Characters, m_Tristate)
  Else
    Err.Raise 54, , "Bad file mode"
  End If
End Function

Public Function ReadAll() As String
  If m_Mode = ForReading Then
    ReadAll = mFunctions.ReadAllFromFile(m_lFile, m_lLocation, m_Tristate)
  Else
    Err.Raise 54, , "Bad file mode"
  End If
End Function

Public Function ReadLine() As String
  If m_Mode = ForReading Then
    ReadLine = mFunctions.ReadLine(m_lFile, m_lLocation, m_Tristate)
  Else
    Err.Raise 54, , "Bad file mode"
  End If
End Function

Public Sub Skip(Characters As Long)
  If m_Mode = ForReading Then
    m_lLocation = m_lLocation + Characters
  Else
    Err.Raise 54, , "Bad file mode"
  End If
End Sub

Public Sub SkipLine()
  If m_Mode = ForReading Then
    Call mFunctions.SkipLine(m_lFile, m_lLocation, m_Tristate)
  Else
    Err.Raise 54, , "Bad file mode"
  End If
End Sub

Public Sub WriteText(Text As String) 'Renamed from: Write()
  If m_Mode = ForReading Then
    Err.Raise 54, , "Bad file mode"
  Else
    Call mFunctions.WriteText(m_lFile, m_lLocation, Text, m_Tristate)
  End If
End Sub

Public Sub WriteBlankLines(Lines As Long)
  If m_Mode = ForReading Then
    Err.Raise 54, , "Bad file mode"
  Else
    Call mFunctions.WriteBlankLines(m_lFile, m_lLocation, Lines, m_Tristate)
  End If
End Sub

Public Sub WriteLine(Optional Text As String)
  If m_Mode = ForReading Then
    Err.Raise 54, , "Bad file mode"
  Else
    Call mFunctions.WriteText(m_lFile, m_lLocation, Text & vbCrLf, m_Tristate)
  End If
End Sub

Friend Sub OpenFile(FilePath As String, Mode As FileSystemAPI.IOMode, Create As Boolean, Overwrite As Boolean, Format As FileSystemAPI.Tristate, Optional StandardStream As Boolean = False, Optional StandardStreamType As FileSystemAPI.StandardStreamTypes = StdErr)
  m_Tristate = Format
  m_Mode = Mode
  m_lFile = mFunctions.OpenFile(FilePath, m_Mode, Create, Overwrite, m_lLocation, m_Tristate, StandardStream, StandardStreamType)
End Sub

Private Sub Class_Terminate()
  Call Me.CloseFile
End Sub

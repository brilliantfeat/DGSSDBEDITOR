VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum CompressionLevels
  None
  FastLZ
  ZLib
End Enum

Public Enum StateReason
  DBRead
  Update
  AddNew
  Delete
End Enum

Private mCompressionLevel As CompressionLevels
Private mContent() As Byte, mTopRow As Long, mLeftCol As Long, mReason As StateReason

Private Sub Class_Initialize() 'set the Default-CompressionLevel of this Storage-Class
  mCompressionLevel = CompressionLevels.FastLZ
End Sub

Public Function SaveContent(Content() As Byte, ByVal TopRow&, ByVal LeftCol&, Reason As StateReason) As Long
  Select Case mCompressionLevel
    Case CompressionLevels.None:   mContent = Content
    Case CompressionLevels.FastLZ: New_c.Crypt.FastLZCompress Content, mContent
    Case CompressionLevels.ZLib:   New_c.Crypt.ZlibCompress Content, mContent
  End Select

  mTopRow = TopRow: mLeftCol = LeftCol: mReason = Reason
  SaveContent = UBound(mContent) + 1
End Function

Public Sub GetContent(Content() As Byte, TopRow As Long, LeftCol As Long, Reason As StateReason)
  Select Case mCompressionLevel
    Case CompressionLevels.None:   Content = mContent
    Case CompressionLevels.FastLZ: New_c.Crypt.FastLZDecompress mContent, Content
    Case CompressionLevels.ZLib:   New_c.Crypt.ZLibDecompress mContent, Content
  End Select

  TopRow = mTopRow: LeftCol = mLeftCol: Reason = mReason
End Sub



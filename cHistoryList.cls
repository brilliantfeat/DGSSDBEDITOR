VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHistoryList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private States As Collection, UndoRedoIdx&
Public UndoDepth As Long

Private Sub Class_Initialize()
  Set States = New Collection
End Sub

Public Sub Clear()
  Set States = New Collection
  UndoRedoIdx = 0
End Sub

Public Sub SaveState(State As cState)
Dim i&
  'delete all states up to the last UndoIdx, before adding the new state on Top
  For i = 1 To UndoRedoIdx - 1: States.Remove 1: Next i

  'add state on top of the States-Collection
  If States.Count = 0 Then States.Add State Else States.Add State, , 1
  If States.Count > UndoDepth Then States.Remove States.Count 'ensure max Undo-Depth
  UndoRedoIdx = 1 'last added state is alwys on top
End Sub

Public Property Get PreviousState() As cState
  If States.Count = 0 Or UndoRedoIdx >= States.Count Then Exit Property
  UndoRedoIdx = UndoRedoIdx + 1
  Set PreviousState = States(UndoRedoIdx)
End Property

Public Property Get NextState() As cState
  If States.Count = 0 Or UndoRedoIdx <= 1 Then Exit Property
  UndoRedoIdx = UndoRedoIdx - 1
  Set NextState = States(UndoRedoIdx)
End Property

Public Property Get UndoEnabled() As Boolean
  UndoEnabled = (States.Count > 1 And UndoRedoIdx < States.Count)
End Property

Public Property Get RedoEnabled() As Boolean
  RedoEnabled = (UndoRedoIdx > 1)
End Property


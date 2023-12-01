VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'class Deque for non deterministic finite automat
Private m_Col As Collection

Private Sub Class_Initialize()
    Set m_Col = New Collection
End Sub
'
'Public Property Get CanEval() As Boolean
'    CanEval = First.CanEval
'End Property

Public Property Get Count() As Long
    Count = m_Col.Count
End Property
'
'Public Function Eval()
'    Eval = Me.First.Eval
'End Function

Public Function First() As Expression
    'the first or lowest element of the deque
    If Me.Count = 0 Then Exit Function
    Set First = m_Col.Item(1)
End Function

Public Function Peek() As Expression
    If Me.Count = 0 Then Exit Function
    Set Peek = m_Col.Item(m_Col.Count)
End Function

Public Function Pop() As Expression
    Dim c As Long: c = m_Col.Count
    If c = 0 Then Exit Function
    Set Pop = m_Col.Item(c)
    m_Col.Remove c
End Function

Public Sub Push(Expr As Expression)
    'Adds an element on top of the deque
    m_Col.Add Expr
End Sub


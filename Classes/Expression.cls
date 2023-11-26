VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Expression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Eval():                     End Function
Public Property Get Rank() As OperatorRank: End Property
Public Property Get CanEval() As Boolean:   End Property
Public Function ToStr(Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String: End Function


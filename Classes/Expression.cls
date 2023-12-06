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

Public Property Get CanEval() As Boolean:           End Property
Public Function Eval():                             End Function
Public Property Get ErrInfo() As String:            End Property
Public Property Set Expr1(ByVal RHS As Expression): End Property
Public Property Get Expr1() As Expression:          End Property
Public Property Set Expr2(ByVal RHS As Expression): End Property
Public Property Get Expr2() As Expression:          End Property
Public Property Get Operator1() As String:          End Property
Public Property Get Operator2() As String:          End Property
Public Property Get Rank() As OperatorRank:         End Property
Public Function ToStr(ByVal Fmt As FormatExpr, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String: End Function

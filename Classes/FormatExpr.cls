VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormatExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function ExprLitNumToStr(aLit As ExprLitNum, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:        End Function
Public Function ExprOpBinToStr(aOpBin As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:       End Function
Public Function ExprOpSubtToStr(aOpBin As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:      End Function
Public Function ExprOpDivToStr(aOpDiv As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:       End Function
Public Function ExprOpUnaAsLToStr(aOpUnAsL As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:  End Function
Public Function ExprOpUnaAsRToStr(aOpUnAsR As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:  End Function
Public Function ExprOpBracToStr(aOpUnAsR As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:    End Function
Public Function ExprFunctionToStr(aFunc As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:     End Function
Public Function ExprFuncAbsToStr(aExprAbs As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:   End Function
Public Function ExprOpCubToStr(aOpCub As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:       End Function
Public Function ExprOpSqrToStr(aOpSqr As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:       End Function
Public Function ExprOpPowToStr(aOpPow As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:       End Function

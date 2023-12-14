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
Public Function ExprLitBolToStr(aLit As ExprLitBol, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:        End Function
Public Function ExprLitNumToStr(aLit As ExprLitNum, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:        End Function
Public Function ExprLitStrToStr(aLit As ExprLitStr, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:        End Function
Public Function ExprOpBinToStr(aOpBin As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:       End Function
Public Function ExprOpSubtToStr(aOpBin As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:      End Function
Public Function ExprOpDivToStr(aOpDiv As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:       End Function
Public Function ExprOpUnaAsLToStr(aOpUnAsL As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:  End Function
Public Function ExprOpUnaAsRToStr(aOpUnAsR As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:  End Function
Public Function ExprOpBracToStr(aOpUnAsR As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:    End Function
Public Function ExprFunctionToStr(aFunc As ExprFunction, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:   End Function
Public Function ExprFunction2ToStr(aFunc As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:    End Function
Public Function ExprFuncAbsToStr(aExprAbs As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:   End Function
Public Function ExprOpCubToStr(aOpCub As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:       End Function
Public Function ExprOpSqrToStr(aOpSqr As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:       End Function
Public Function ExprOpPowToStr(aOpPow As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:       End Function
Public Function ExprOpFactToStr(aOpFact As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:     End Function
Public Function ExprOpSqrtToStr(aOpSqrt As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:     End Function
Public Function ExprOpIIfToStr(aOpIIf As ExprOpIIf, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:        End Function
Public Function ExprOpBolAndToStr(aOpAnd As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:             End Function
Public Function ExprOpBolOrToStr(aOpAnd As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:              End Function
Public Function ExprOpBolXorToStr(aOpAnd As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:             End Function
Public Function ExprOpBolEqualToStr(aOpEquals As ExprOpBolEqual, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:    End Function
Public Function ExprOpBolGreaTToStr(aOpEquals As ExprOpBolGreaT, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:    End Function
Public Function ExprOpBolGreaTEquToStr(aOpEquals As ExprOpBolGreaT, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String: End Function
Public Function ExprOpBolLessTToStr(aOpEquals As ExprOpBolGreaT, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:    End Function
Public Function ExprOpBolLessTEquToStr(aOpEquals As ExprOpBolGreaT, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String: End Function
Public Function ExprOpBolNotEquToStr(aOpEquals As ExprOpBolNotEqu, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String:  End Function


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
Public Sub ExprFunctionToStr(aFunc As ExprFunction, ByVal ParentRank As OperatorRank):   End Sub
'Public Sub ExprFunction2ToStr(aFunc As ExprFunction, ByVal ParentRank As OperatorRank):  End Sub
Public Sub ExprFunction2ToStr(aFunc As Expression, ByVal ParentRank As OperatorRank):  End Sub

Public Sub ExprLitBolToStr(aLit As ExprLitBol, ByVal ParentRank As OperatorRank):        End Sub
Public Sub ExprLitDatToStr(aLit As ExprLitDat, ByVal ParentRank As OperatorRank):        End Sub
Public Sub ExprLitNumToStr(aLit As ExprLitNum, ByVal ParentRank As OperatorRank):        End Sub
Public Sub ExprLitStrToStr(aLit As ExprLitStr, ByVal ParentRank As OperatorRank):        End Sub
Public Sub ExprOp1DivXToStr(aExpr As ExprOp1DivX, ByVal ParentRank As OperatorRank):     End Sub
Public Sub ExprOpAbsToStr(aExpr As ExprOpAbs, ByVal ParentRank As OperatorRank):         End Sub
Public Sub ExprOpAddToStr(aExpr As ExprOpAdd, ByVal ParentRank As OperatorRank):         End Sub
Public Sub ExprOpBolAndToStr(aExpr As ExprOpBolAnd, ByVal ParentRank As OperatorRank):           End Sub
Public Sub ExprOpBolEqualToStr(aExpr As ExprOpBolEqual, ByVal ParentRank As OperatorRank):       End Sub
Public Sub ExprOpBolGreaTToStr(aExpr As ExprOpBolGreaT, ByVal ParentRank As OperatorRank):       End Sub
Public Sub ExprOpBolGreaTEquToStr(aExpr As ExprOpBolGreaTEqu, ByVal ParentRank As OperatorRank): End Sub
Public Sub ExprOpBolLessTToStr(aExpr As ExprOpBolLessT, ByVal ParentRank As OperatorRank):       End Sub
Public Sub ExprOpBolLessTEquToStr(aExpr As ExprOpBolLessTEqu, ByVal ParentRank As OperatorRank): End Sub
Public Sub ExprOpBolNotEquToStr(aExpr As ExprOpBolNotEqu, ByVal ParentRank As OperatorRank):     End Sub
Public Sub ExprOpBolOrToStr(aExpr As ExprOpBolOr, ByVal ParentRank As OperatorRank):             End Sub
Public Sub ExprOpBolXorToStr(aExpr As ExprOpBolXor, ByVal ParentRank As OperatorRank):           End Sub
Public Sub ExprOpBracToStr(aExpr As ExprOpBrac, ByVal ParentRank As OperatorRank):       End Sub
Public Sub ExprOpCubToStr(aExpr As ExprOpCub, ByVal ParentRank As OperatorRank):         End Sub
Public Sub ExprOpDivToStr(aExpr As ExprOpDiv, ByVal ParentRank As OperatorRank):         End Sub
Public Sub ExprOpFactToStr(aExpr As ExprOpFact, ByVal ParentRank As OperatorRank):       End Sub
Public Sub ExprOpIIfToStr(aExpr As ExprOpIIf, ByVal ParentRank As OperatorRank):         End Sub
Public Sub ExprOpLNToStr(aExpr As ExprOpLN, ByVal ParentRank As OperatorRank):           End Sub
Public Sub ExprOpLog10ToStr(aExpr As ExprOpLog10, ByVal ParentRank As OperatorRank):     End Sub
Public Sub ExprOpLogNToStr(aExpr As ExprOpLogN, ByVal ParentRank As OperatorRank):       End Sub
Public Sub ExprOpMulToStr(aExpr As ExprOpMul, ByVal ParentRank As OperatorRank):         End Sub
Public Sub ExprOpNegToStr(aExpr As ExprOpNeg, ByVal ParentRank As OperatorRank):         End Sub
Public Sub ExprOpPowToStr(aExpr As ExprOpPow, ByVal ParentRank As OperatorRank):         End Sub
Public Sub ExprOpPow10ToStr(aExpr As ExprOpPow10, ByVal ParentRank As OperatorRank):     End Sub
Public Sub ExprOpSqrToStr(aExpr As ExprOpSqr, ByVal ParentRank As OperatorRank):        End Sub
Public Sub ExprOpSqrtToStr(aExpr As ExprOpSqrt, ByVal ParentRank As OperatorRank):       End Sub
Public Sub ExprOpStrConcToStr(aExpr As ExprOpStrConc, ByVal ParentRank As OperatorRank): End Sub
Public Sub ExprOpSubtToStr(aExpr As ExprOpSubt, ByVal ParentRank As OperatorRank):       End Sub

'Public Sub ExprOpBinToStr(aOpBin As Expression, ByVal ParentRank As OperatorRank):      End Sub
'Public Sub ExprOpUnaAsLToStr(aOpUnAsL As Expression, ByVal ParentRank As OperatorRank): End Sub
'Public Sub ExprOpUnaAsRToStr(aOpUnAsR As Expression, ByVal ParentRank As OperatorRank): End Sub
'Public Sub ExprFuncAbsToStr(aExprAbs As Expression, ByVal ParentRank As OperatorRank):  End Sub

Public Function ToStr(Optional ex As Expression = Nothing) As String: End Function

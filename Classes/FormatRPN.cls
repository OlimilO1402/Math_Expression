VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormatRPN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements FormatExpr

Private Function FormatExpr_ExprLitNumToStr(aLit As ExprLitNum, Optional ByVal ParentRank As OperatorRank = 0&) As String
    FormatExpr_ExprLitNumToStr = Trim(Str(aLit.Value))
End Function

'for operators: + - * ^
Private Function FormatExpr_ExprOpBinToStr(aOpBin As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    With aOpBin
        Dim s As String:       s = .Operator1
        Dim LHS As Expression: Set LHS = .Expr1
        Dim RHS As Expression: Set RHS = .Expr2
    End With
    s = LHS.ToStr(Me) & " " & RHS.ToStr(Me) & s
    FormatExpr_ExprOpBinToStr = s
End Function

'for operator: /
Private Function FormatExpr_ExprOpDivToStr(aOpDiv As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    FormatExpr_ExprOpDivToStr = FormatExpr_ExprOpBinToStr(aOpDiv)
End Function

Private Function FormatExpr_ExprOpSubtToStr(aOpBin As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    '
End Function

'Oh, and everything else is just the same values to stack, fnktname/Operator to stack ready

'
Private Function FormatExpr_ExprOpUnaAsLToStr(aOpUnAsL As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    '
End Function

'x², x³, x!
Private Function FormatExpr_ExprOpUnaAsRToStr(aOpUnAsR As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    '
End Function

Private Function FormatExpr_ExprOpUnaAsLRToStr(aOpUnaAsR As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    '
End Function


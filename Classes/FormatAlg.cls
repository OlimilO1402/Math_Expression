VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormatAlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements FormatExpr

Private Function FormatExpr_ExprLitNumToStr(aLit As ExprLitNum, Optional ByVal ParentRank As OperatorRank = 0&) As String
    FormatExpr_ExprLitNumToStr = Trim(Str(aLit.Value))
End Function

'for operator binary, commutative: x + y, x * y
Private Function FormatExpr_ExprOpBinToStr(aOpBin As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    With aOpBin
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1
        Dim LHS As Expression: Set LHS = .Expr1
        Dim RHS As Expression: Set RHS = .Expr2
    End With
    Dim b As Boolean:          b = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
    If Not LHS Is Nothing Then s = IIf(b, "(", "") & LHS.ToStr(Me, r) & s
    If Not RHS Is Nothing Then s = s & RHS.ToStr(Me, r) & IIf(b, ")", "")
    FormatExpr_ExprOpBinToStr = s
End Function

'Differenz = Minuend - Subtrahend
'for operator binary -, non commutative: x - y
Private Function FormatExpr_ExprOpSubtToStr(aOpSubt As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    With aOpSubt
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1
        Dim LHS As Expression: Set LHS = .Expr1
        Dim RHS As Expression: Set RHS = .Expr2
    End With
    Dim b As Boolean: b = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
    If Not LHS Is Nothing Then s = IIf(b, "(", "") & LHS.ToStr(Me, r) & s
    If Not RHS Is Nothing Then s = s & RHS.ToStr(Me, r + 1) & IIf(b, ")", "")
    FormatExpr_ExprOpSubtToStr = s
End Function

'Quotient = Dividend / Divisor
'for operator, non commutativ: x / y
Private Function FormatExpr_ExprOpDivToStr(aOpDiv As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    With aOpDiv
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1
        Dim LHS As Expression: Set LHS = .Expr1
        Dim RHS As Expression: Set RHS = .Expr2
    End With
    'Dim bDiv As Boolean:    bDiv = LHS.Rank <> OperatorRank.Rank1ExprConst And RHS.Rank <> OperatorRank.Rank8ExprOpBrac
    Dim b As Boolean:          b = r <= ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
    If Not LHS Is Nothing Then s = IIf(b, "(", "") & LHS.ToStr(Me, r) & s
    'If Not RHS Is Nothing Then s = s & IIf(bDiv, "(", "") & RHS.ToStr(Me, r) & IIf(bDiv, ")", "") & IIf(b, ")", "")
    If Not RHS Is Nothing Then s = s & RHS.ToStr(Me, r + 1) & IIf(b, ")", "")
    FormatExpr_ExprOpDivToStr = s
End Function

'for operator unary associative left: -x
Private Function FormatExpr_ExprOpUnaAsLToStr(aOpUnAsL As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    With aOpDiv
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1
        Dim RHS As Expression: Set RHS = .Expr2
    End With
    Dim b As Boolean: b = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
    If Not RHS Is Nothing Then s = IIf(b, "(", "") & s & RHS.ToStr(Me, r) & IIf(b, ")", "")
    FormatExpr_ExprOpUnaAsLToStr = s
End Function

'for operator unary associative right: x², x³, x!
Private Function FormatExpr_ExprOpUnaAsRToStr(aOpUnAsR As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    With aOpUnAsR
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1
        Dim LHS As Expression: Set LHS = .Expr1
    End With
    Dim b As Boolean: b = ParentRank <> OperatorRank.Rank9ExprOpBrac And ParentRank <> OperatorRank.Rank0None
    Dim s As String:  s = myOp
    If Not LHS Is Nothing Then s = IIf(b, "(", "") & m_LHS.ToStr(Me, r) & s & IIf(b, ")", "")
    'If Not m_LHS Is Nothing Then s = m_LHS.ToStr(myRank) & s '& " = " & Trim(Str(Expression_Eval))
    Expression_ToStr = s
End Function

Private Function FormatExpr_ExprOpUnaAsLRToStr(aOpUnaAsR As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
'    Dim b As Boolean: b = myRank < ParentRank And ParentRank <> OperatorRank.Rank7ExprOpBrac
'    Dim s As String:  s = myOp
'    If Not m_Ins Is Nothing Then s = IIf(b, "(", "") & s & m_Ins.ToStr(myRank) & myOp & IIf(b, ")", "")
'    Expression_ToStr = s
    With aOpUnaAsR
        Dim r As OperatorRank: r = .Rank
        Dim s1 As String:      s = .Operator1
        Dim s2 As String:      s = .Operator2
        Dim Ins As Expression: Set LHS = .Expr1
    End With
    Dim b As Boolean:          b = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
    Dim s As String
    If Not Ins Is Nothing Then s = IIf(b, "(", "") & s1 & Ins.ToStr(Me, r) & s2
    'If Not RHS Is Nothing Then s = s & IIf(bDiv, "(", "") & RHS.ToStr(r) & IIf(bDiv, ")", "") & IIf(b, ")", "")
    FormatExpr_ExprOpDivToStr = s
End Function


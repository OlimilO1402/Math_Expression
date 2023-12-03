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
Private m_Condensed As Boolean
Private m_XLCmpatbl As Boolean

Friend Sub New_(ByVal IsCondensed As Boolean, Optional ByVal ExcelCompatible As Boolean = False)
    m_Condensed = IsCondensed
    m_XLCmpatbl = ExcelCompatible
End Sub

Public Property Get Condensed() As Boolean
    Condensed = m_Condensed
End Property

' v ############################## v '    Implements FormatExpr    ' v ############################## v '
Private Function FormatExpr_ExprLitNumToStr(aLit As ExprLitNum, Optional ByVal ParentRank As OperatorRank = 0&) As String
    FormatExpr_ExprLitNumToStr = Trim(Str(aLit.Value))
End Function

'for operator binary, commutative: x + y, x * y
Private Function FormatExpr_ExprOpBinToStr(aOpBin As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    With aOpBin
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
        Dim LHS As Expression: Set LHS = .Expr1
        Dim RHS As Expression: Set RHS = .Expr2
    End With
    Dim b As Boolean:          b = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
    If Not LHS Is Nothing Then s = IIf(b, "(", "") & LHS.ToStr(Me, r) & s
    If Not RHS Is Nothing Then s = s & RHS.ToStr(Me, r) & IIf(b, ")", "")
    FormatExpr_ExprOpBinToStr = s
End Function

Private Function FormatExpr_ExprOpCubToStr(aOpCub As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    FormatExpr_ExprOpCubToStr = FormatExpr_ExprOpSqrToStr(aOpCub, ParentRank)
End Function

Private Function FormatExpr_ExprOpPowToStr(aOpPow As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    'FormatExpr_ExprOpPowToStr = FormatExpr_ExprOpBinToStr(aOpPow, ParentRank)
    With aOpPow
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
        Dim LHS As Expression: Set LHS = .Expr1
        Dim RHS As Expression: Set RHS = .Expr2
    End With
    Dim b As Boolean:          b = r <= ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
    If Not LHS Is Nothing Then s = IIf(b, "(", "") & LHS.ToStr(Me, r) & s
    If Not RHS Is Nothing Then s = s & RHS.ToStr(Me, r) & IIf(b, ")", "")
    FormatExpr_ExprOpPowToStr = s
End Function

Private Function FormatExpr_ExprOpSqrToStr(aOpSqr As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    If m_XLCmpatbl Then
        FormatExpr_ExprOpSqrToStr = FormatExpr_ExprOpPowToStr(aOpSqr, ParentRank)
        Exit Function
    End If
    With aOpSqr
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator2: If m_Condensed Then s = Trim(s)
        Dim LHS As Expression: Set LHS = .Expr1
    End With
    Dim b As Boolean:          b = r <= ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
    If Not LHS Is Nothing Then s = LHS.ToStr(Me, r) & s  'IIf(b, "(", "") &
    FormatExpr_ExprOpSqrToStr = s
End Function

'Differenz = Minuend - Subtrahend
'for operator binary -, non commutative: x - y
Private Function FormatExpr_ExprOpSubtToStr(aOpSubt As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    With aOpSubt
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
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
        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
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
        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
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
        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
        Dim LHS As Expression: Set LHS = .Expr1
    End With
    Dim b As Boolean: b = ParentRank <> OperatorRank.Rank9ExprOpBrac And ParentRank <> OperatorRank.Rank0None
    'Dim s As String:  s = myOp
    If Not LHS Is Nothing Then s = IIf(b, "(", "") & LHS.ToStr(Me, r) & s & IIf(b, ")", "")
    'If Not m_LHS Is Nothing Then s = m_LHS.ToStr(myRank) & s '& " = " & Trim(Str(Expression_Eval))
    FormatExpr_ExprOpUnaAsRToStr = s
End Function

'for operator unary left/right | x |, ( x )
Private Function FormatExpr_ExprOpBracToStr(aOpUnaAsLR As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
'    Dim b As Boolean: b = myRank < ParentRank And ParentRank <> OperatorRank.Rank7ExprOpBrac
'    Dim s As String:  s = myOp
'    If Not m_Ins Is Nothing Then s = IIf(b, "(", "") & s & m_Ins.ToStr(myRank) & myOp & IIf(b, ")", "")
'    Expression_ToStr = s
    With aOpUnaAsLR
        Dim r As OperatorRank: r = .Rank
        Dim s1 As String:      s1 = .Operator1: If m_Condensed Then s1 = Trim(s1)
        Dim s2 As String:      s2 = .Operator2: If m_Condensed Then s2 = Trim(s2)
        Dim Ins As Expression: Set Ins = .Expr1
    End With
    Dim b As Boolean:          b = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
    Dim s As String
    If Not Ins Is Nothing Then s = IIf(b, "(", "") & s1 & Ins.ToStr(Me, r) & s2 & IIf(b, ")", "")
    'If Not RHS Is Nothing Then s = s & IIf(bDiv, "(", "") & RHS.ToStr(r) & IIf(bDiv, ")", "") & IIf(b, ")", "")
    FormatExpr_ExprOpBracToStr = s
End Function

Private Function FormatExpr_ExprFunctionToStr(aFunc As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    With aFunc
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
        Dim Ins As Expression: Set Ins = .Expr1
    End With
    Dim FuncName As String
    If Len(s) Then FuncName = s
    If FuncName = "(" Or FuncName = ")" Then Exit Function
    If Right(FuncName, 1) = ")" Then FuncName = Left(FuncName, Len(FuncName) - 1)
    If Right(FuncName, 1) = "(" Then FuncName = Left(FuncName, Len(FuncName) - 1)
    If Not Ins Is Nothing Then s = FuncName & "(" & Ins.ToStr(Me, r) & ")" 's2 & IIf(b, ")", "")
    FormatExpr_ExprFunctionToStr = s
End Function

Private Function FormatExpr_ExprFuncAbsToStr(aExprAbs As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    If m_XLCmpatbl Then
        FormatExpr_ExprFuncAbsToStr = FormatExpr_ExprFunctionToStr(MNew.ExprFunction("Abs", Nothing, aExprAbs.Expr1), ParentRank)
        Exit Function
    End If
    FormatExpr_ExprFuncAbsToStr = FormatExpr_ExprOpUnaAsLRToStr(aExprAbs, ParentRank)
End Function

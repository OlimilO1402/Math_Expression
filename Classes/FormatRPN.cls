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
Private m_Sep As String

Friend Sub New_(Optional ByVal SeparatorIsNewLine As Boolean = False)
    If SeparatorIsNewLine Then m_Sep = vbCrLf Else m_Sep = " "
End Sub

' v ############################## v '    Implements FormatExpr    ' v ############################## v '
Private Function FormatExpr_ExprLitNumToStr(aLit As ExprLitNum, Optional ByVal ParentRank As OperatorRank = 0&) As String
    FormatExpr_ExprLitNumToStr = Trim(Str(aLit.Value))
End Function

'for operators: + - * ^
Private Function FormatExpr_ExprOpBinToStr(aOpBin As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    With aOpBin
        Dim sOp As String:         sOp = Trim(.Operator1)
        Dim LHS As Expression: Set LHS = .Expr1
        Dim RHS As Expression: Set RHS = .Expr2
    End With
    Dim s As String
    If Not LHS Is Nothing Then s = s & LHS.ToStr(Me) & m_Sep
    If Not RHS Is Nothing Then s = s & RHS.ToStr(Me) & m_Sep
    s = s & sOp '& m_Sep
    FormatExpr_ExprOpBinToStr = s
End Function

Private Function FormatExpr_ExprOpCubToStr(aOpCub As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    FormatExpr_ExprOpCubToStr = FormatExpr_ExprOpPowToStr(aOpCub)
End Function

'for operator: /
Private Function FormatExpr_ExprOpDivToStr(aOpDiv As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    FormatExpr_ExprOpDivToStr = FormatExpr_ExprOpBinToStr(aOpDiv)
End Function

Private Function FormatExpr_ExprOpFactToStr(aOpPow As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    '
End Function

Private Function FormatExpr_ExprOpPowToStr(aOpPow As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    FormatExpr_ExprOpPowToStr = FormatExpr_ExprOpBinToStr(aOpPow)
End Function

Private Function FormatExpr_ExprOpSqrToStr(aOpSqr As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    FormatExpr_ExprOpSqrToStr = FormatExpr_ExprOpCubToStr(aOpSqr)
End Function

Private Function FormatExpr_ExprOpSqrtToStr(aOpSqrt As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    '
End Function

Private Function FormatExpr_ExprOpSubtToStr(aOpBin As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    FormatExpr_ExprOpSubtToStr = FormatExpr_ExprOpBinToStr(aOpBin)
End Function

'Oh, and everything else is just the same values to stack, fnktname/Operator to stack ready
Private Function FormatExpr_ExprOpUnaAsLToStr(aOpUnAsL As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    '
End Function

'x², x³, x!
Private Function FormatExpr_ExprOpUnaAsRToStr(aOpUnAsR As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    '
End Function

'|x|, (x)
Private Function FormatExpr_ExprOpBracToStr(aOpUnaAsR As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    With aOpUnaAsR
        'Dim s1 As String: s1 = .Operator1
        'Dim s2 As String: s2 = .Operator2
        
    End With
End Function

Private Function FormatExpr_ExprFunctionToStr(aFunc As ExprFunction, Optional ByVal ParentRank As OperatorRank = 0&) As String
    '
End Function

Private Function FormatExpr_ExprFunction2ToStr(aFunc As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    With aFunc
        Dim sOp As String:         sOp = .Operator1
        Dim LHS As Expression: Set LHS = .Expr1
        Dim RHS As Expression: Set RHS = .Expr2
    End With
    Dim s As String
    If Not LHS Is Nothing Then s = s & LHS.ToStr(Me) & m_Sep
    If Not RHS Is Nothing Then s = s & RHS.ToStr(Me) & m_Sep
    s = s & sOp '& m_Sep
    FormatExpr_ExprFunctionToStr = s
End Function

Private Function FormatExpr_ExprFuncAbsToStr(aExprAbs As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    With aExprAbs
        Dim sOp As String:         sOp = "Abs"
        Dim LHS As Expression: Set LHS = .Expr1
        Dim RHS As Expression: Set RHS = .Expr2
    End With
    Dim s As String
    If Not LHS Is Nothing Then s = s & LHS.ToStr(Me) & m_Sep
    If Not RHS Is Nothing Then s = s & RHS.ToStr(Me) & m_Sep
    s = s & sOp '& m_Sep
    FormatExpr_ExprFuncAbsToStr = s
End Function

Private Function FormatExpr_ExprOpIIfToStr(aOpIIf As ExprOpIIf, Optional ByVal ParentRank As OperatorRank = 0&) As String
    Dim ex As Expression: Set ex = aOpIIf
    Dim FuncName As String: FuncName = IIf(m_XLCmpatbl, "Wenn", ex.Operator1)
    Dim exFunc As ExprFunction: Set exFunc = MNew.ExprFunction(FuncName, Nothing)
    exFunc.params.Add aOpIIf.Condition
    exFunc.params.Add ex.Expr1:    exFunc.params.Add ex.Expr2
    Expression_ToStr = Fmt.ExprFunctionToStr(exFunc, ParentRank)
End Function


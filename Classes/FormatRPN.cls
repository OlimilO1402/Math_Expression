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
Private m_Sep  As String
Private m_StrB As StringBuilder

Private Sub Class_Initialize()
    Set m_StrB = New StringBuilder
End Sub

Friend Sub New_(Optional ByVal SeparatorIsNewLine As Boolean = False)
    If SeparatorIsNewLine Then m_Sep = vbCrLf Else m_Sep = " "
End Sub

' v ############################## v '    Implements FormatExpr    ' v ############################## v '

Private Sub FormatExpr_ExprFunctionToStr(aFunc As ExprFunction, ByVal ParentRank As OperatorRank)
    '
End Sub

'Private Sub FormatExpr_ExprFunction2ToStr(aFunc As ExprFunction, ByVal ParentRank As OperatorRank)
Private Sub FormatExpr_ExprFunction2ToStr(aFunc As Expression, ByVal ParentRank As OperatorRank)
'    With aFunc
'        Dim sOp As String:         sOp = .Operator1
'        Dim LHS As Expression: Set LHS = .Expr1
'        Dim RHS As Expression: Set RHS = .Expr2
'    End With
'    Dim s As String
'    If Not LHS Is Nothing Then s = s & LHS.ToStr(Me) & m_Sep
'    If Not RHS Is Nothing Then s = s & RHS.ToStr(Me) & m_Sep
'    s = s & sOp '& m_Sep
'    'FormatExpr_ExprFunction2ToStr = s
End Sub

Private Sub FormatExpr_ExprLitBolToStr(aLit As ExprLitBol, ByVal ParentRank As OperatorRank)
    m_StrB.Append aLit.Value
End Sub

Private Sub FormatExpr_ExprLitDatToStr(aLit As ExprLitDat, ByVal ParentRank As OperatorRank)
    m_StrB.Append aLit.Value
End Sub

Private Sub FormatExpr_ExprLitNumToStr(aLit As ExprLitNum, ByVal ParentRank As OperatorRank)
    m_StrB.Append aLit.Value
End Sub

Private Sub FormatExpr_ExprLitStrToStr(aLit As ExprLitStr, ByVal ParentRank As OperatorRank)
    m_StrB.Append aLit.Value
End Sub

Private Sub FormatExpr_ExprOp1DivXToStr(aExpr As ExprOp1DivX, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpAbsToStr(aExpr As ExprOpAbs, ByVal ParentRank As OperatorRank)
'    With aExprAbs
'        Dim sOp As String:         sOp = "Abs"
'        Dim LHS As Expression: Set LHS = .Expr1
'        Dim RHS As Expression: Set RHS = .Expr2
'    End With
'    Dim s As String
'    If Not LHS Is Nothing Then s = s & LHS.ToStr(Me) & m_Sep
'    If Not RHS Is Nothing Then s = s & RHS.ToStr(Me) & m_Sep
'    s = s & sOp '& m_Sep
'    'FormatExpr_ExprFuncAbsToStr = s
End Sub

Private Sub FormatExpr_ExprOpAddToStr(aExpr As ExprOpAdd, ByVal ParentRank As OperatorRank)
    '
End Sub

'for operators: + - * ^
'Private Sub FormatExpr_ExprOpBinToStr(aOpBin As Expression, ByVal ParentRank As OperatorRank)
'    With aOpBin
'        Dim sOp As String:         sOp = Trim(.Operator1)
'        Dim LHS As Expression: Set LHS = .Expr1
'        Dim RHS As Expression: Set RHS = .Expr2
'    End With
'    Dim s As String
'    If Not LHS Is Nothing Then s = s & LHS.ToStr(Me) & m_Sep
'    If Not RHS Is Nothing Then s = s & RHS.ToStr(Me) & m_Sep
'    s = s & sOp '& m_Sep
'    FormatExpr_ExprOpBinToStr = s
'End Sub

Private Sub FormatExpr_ExprOpBolAndToStr(aExpr As ExprOpBolAnd, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpBolEqualToStr(aExpr As ExprOpBolEqual, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpBolGreaTToStr(aExpr As ExprOpBolGreaT, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpBolGreaTEquToStr(aExpr As ExprOpBolGreaTEqu, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpBolLessTToStr(aExpr As ExprOpBolLessT, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpBolLessTEquToStr(aExpr As ExprOpBolLessTEqu, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpBolNotEquToStr(aExpr As ExprOpBolNotEqu, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpBolOrToStr(aExpr As ExprOpBolOr, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpBolXorToStr(aExpr As ExprOpBolXor, ByVal ParentRank As OperatorRank)
    '
End Sub
'|x|, (x)
Private Sub FormatExpr_ExprOpBracToStr(aExpr As ExprOpBrac, ByVal ParentRank As OperatorRank)
'    With aOpUnaAsR
'        'Dim s1 As String: s1 = .Operator1
'        'Dim s2 As String: s2 = .Operator2
'
'    End With
End Sub

Private Sub FormatExpr_ExprOpCubToStr(aExpr As ExprOpCub, ByVal ParentRank As OperatorRank)
'    FormatExpr_ExprOpCubToStr = FormatExpr_ExprOpPowToStr(aOpCub)
End Sub

'for operator: /
Private Sub FormatExpr_ExprOpDivToStr(aExpr As ExprOpDiv, ByVal ParentRank As OperatorRank)
'    FormatExpr_ExprOpDivToStr = FormatExpr_ExprOpBinToStr(aOpDiv)
End Sub

Private Sub FormatExpr_ExprOpFactToStr(aExpr As ExprOpFact, ByVal ParentRank As OperatorRank)
'    With aOpPow
'        Dim sOp As String:         sOp = .Operator1
'        Dim LHS As Expression: Set LHS = .Expr1
'    End With
'    Dim s As String
'    If Not LHS Is Nothing Then s = s & LHS.ToStr(Me) & m_Sep
'    s = s & sOp
'    FormatExpr_ExprOpFactToStr = s
End Sub

Private Sub FormatExpr_ExprOpIIfToStr(aExpr As ExprOpIIf, ByVal ParentRank As OperatorRank)
'    Dim ex As Expression: Set ex = aOpIIf
'    Dim FuncName As String: FuncName = IIf(m_XLCompatbl, "Wenn", ex.Operator1)
'    Dim exFunc As ExprFunction: Set exFunc = MNew.ExprFunction(FuncName, Nothing)
'    exFunc.params.Add aOpIIf.Condition
'    exFunc.params.Add ex.Expr1:    exFunc.params.Add ex.Expr2
'    Expression_ToStr = Fmt.ExprFunctionToStr(exFunc, ParentRank)
End Sub

Private Sub FormatExpr_ExprOpLNToStr(aExpr As ExprOpLN, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpLog10ToStr(aExpr As ExprOpLog10, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpLogNToStr(aExpr As ExprOpLogN, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpMulToStr(aExpr As ExprOpMul, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpNegToStr(aExpr As ExprOpNeg, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpPowToStr(aExpr As ExprOpPow, ByVal ParentRank As OperatorRank)
'    FormatExpr_ExprOpPowToStr = FormatExpr_ExprOpBinToStr(aOpPow)
End Sub

Private Sub FormatExpr_ExprOpPow10ToStr(aExpr As ExprOpPow10, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpSqrToStr(aExpr As ExprOpSqr, ByVal ParentRank As OperatorRank)
'    FormatExpr_ExprOpSqrToStr = FormatExpr_ExprOpCubToStr(aOpSqr)
End Sub

Private Sub FormatExpr_ExprOpSqrtToStr(aExpr As ExprOpSqrt, ByVal ParentRank As OperatorRank)
'    With aOpSqrt
'        Dim sOp As String:         sOp = .Operator1
'        Dim LHS As Expression: Set LHS = .Expr1
'    End With
'    Dim s As String
'    If Not LHS Is Nothing Then s = s & LHS.ToStr(Me) & m_Sep & sOp
'    FormatExpr_ExprOpFactToStr = s
End Sub

Private Sub FormatExpr_ExprOpStrConcToStr(aExpr As ExprOpStrConc, ByVal ParentRank As OperatorRank)
    '
End Sub

Private Sub FormatExpr_ExprOpSubtToStr(aExpr As ExprOpSubt, ByVal ParentRank As OperatorRank)
'    FormatExpr_ExprOpSubtToStr = FormatExpr_ExprOpBinToStr(aOpBin)
End Sub

'Oh, and everything else is just the same; values to stack, fnktname/operator to stack, ready
'Private Sub FormatExpr_ExprOpUnaAsLToStr(aOpUnAsL As Expression, ByVal ParentRank As OperatorRank)
'    '
'End Sub

'x², x³, x!
'Private Sub FormatExpr_ExprOpUnaAsRToStr(aOpUnAsR As Expression, ByVal ParentRank As OperatorRank)
'    With aOpUnAsR
'        Dim sOp As String
'    End With
'End Sub

Private Function FormatExpr_ToStr(Optional ex As Expression = Nothing) As String
    '
End Function

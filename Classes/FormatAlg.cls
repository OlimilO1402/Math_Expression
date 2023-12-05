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

Public Property Get XLCompatible() As Boolean
    XLCompatible = m_XLCmpatbl
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
    With aOpUnAsL
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
        Dim RHS As Expression: Set RHS = .Expr2
    End With
    Dim b As Boolean: b = r <= ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
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
    'Dim b As Boolean: b = ParentRank < OperatorRank.Rank9ExprOpBrac And ParentRank <> OperatorRank.Rank0None
    Dim b As Boolean: b = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
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

Private Function FormatExpr_ExprFunctionToStr(aFunc As ExprFunction, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    'more than 2 params
    With aFunc
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
    End With
    Dim FuncName As String: FuncName = GetFuncName(s)
    Dim sParams As String: sParams = GetParamsList(aFunc.params)
    FormatExpr_ExprFunctionToStr = FuncName & "(" & sParams & ")"
End Function

Private Function GetParamsList(col As Collection, ByVal ParentRank As OperatorRank) As String
    Dim s As String, sp As String, sep As String: sep = GetFuncParamSep
    Dim i As Long, c As Long: c = col.Count
    Dim v, FuncParam As Expression
    For Each v In col
        i = i + 1
        Set FuncParam = v
        sp = FuncParam.ToStr(Me, ParentRank)
        s = s & sp
        If i < c Then s = s & sep
    Next
End Function

Private Function GetFuncParamSep() As String
    GetFuncParamSep = IIf(m_XLCmpatbl, "; ", ", ")
    If m_Condensed Then GetFuncParamSep = Trim(GetFuncParamSep)
End Function

Private Function GetFuncName(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function
    If s = "(" Or s = ")" Then Exit Function
    If Right(s, 1) = ")" Then s = Left(s, Len(s) - 1)
    If Right(s, 1) = "(" Then s = Left(s, Len(s) - 1)
    GetFuncName = s
End Function

Private Function FormatExpr_ExprFunction2ToStr(aFunc As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    'max 2 params
    With aFunc
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
        Dim LHS As Expression: Set LHS = .Expr1
        Dim RHS As Expression: Set RHS = .Expr2
    End With
    Dim FuncName As String: FuncName = GetFuncName(s)
    s = FuncName & "("
    Dim sParams As String
    If Not LHS Is Nothing Then sParams = LHS.ToStr(Me, r)
    If Not RHS Is Nothing Then sParams = sParams & IIf(Len(sParams), GetFuncParamSep, "") & RHS.ToStr(Me, r)
    FormatExpr_ExprFunction2ToStr = s & sParams & ")"
End Function

Private Function FormatExpr_ExprFuncAbsToStr(aExprAbs As Expression, Optional ByVal ParentRank As OperatorRank = 0&) As String
    If m_XLCmpatbl Then
        FormatExpr_ExprFuncAbsToStr = FormatExpr_ExprFunction2ToStr(MNew.ExprFunction("Abs", Nothing, aExprAbs.Expr1), ParentRank)
        Exit Function
    End If
    FormatExpr_ExprFuncAbsToStr = FormatExpr_ExprOpBracToStr(aExprAbs, ParentRank)
End Function

Private Function FormatExpr_ExprOpFactToStr(aOpFact As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    If m_XLCmpatbl Then
        FormatExpr_ExprOpFactToStr = FormatExpr_ExprFunction2ToStr(MNew.ExprFunction("Fakultät", Nothing, aOpFact.Expr1), ParentRank)
        Exit Function
    End If
    FormatExpr_ExprOpFactToStr = FormatExpr_ExprOpUnaAsRToStr(aOpFact, ParentRank)
End Function

Private Function FormatExpr_ExprOpSqrtToStr(aOpSqrt As Expression, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    If m_XLCmpatbl Then
        FormatExpr_ExprOpSqrtToStr = FormatExpr_ExprFunction2ToStr(MNew.ExprFunction("Wurzel", Nothing, aOpSqrt.Expr2), ParentRank)
        Exit Function
    End If
    FormatExpr_ExprOpSqrtToStr = FormatExpr_ExprOpUnaAsLToStr(aOpSqrt, ParentRank)
End Function

Private Function FormatExpr_ExprOpIIfToStr(aOpIIf As ExprOpIIf, Optional ByVal ParentRank As OperatorRank = 0&) As String
    Dim ex As Expression: Set ex = aOpIIf
    Dim FuncName As String: FuncName = IIf(m_XLCmpatbl, "Wenn", ex.Operator1)
    Dim exFunc As ExprFunction: Set exFunc = MNew.ExprFunction(FuncName, Nothing)
    exFunc.params.Add aOpIIf.Condition
    exFunc.params.Add ex.Expr1:    exFunc.params.Add ex.Expr2
    FormatExpr_ExprOpIIfToStr = FormatExpr_ExprFunctionToStr(exFunc, ParentRank)
End Function


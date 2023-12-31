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
Private m_StrB      As StringBuilder

Private Sub Class_Initialize()
    Set m_StrB = New StringBuilder
End Sub

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

Public Function ToStr(Optional ex As Expression = Nothing) As String
    ToStr = FormatExpr_ToStr(ex)
End Function

' v ############################## v '    Implements FormatExpr    ' v ############################## v '

Private Sub ExprToStrB(ByVal bBrax As Boolean, s1 As String, ex1 As Expression, ByVal ParentRank1 As OperatorRank, s2 As String, Optional ex2 As Expression = Nothing, Optional ByVal ParentRank2 As OperatorRank, Optional s3 As String = "")
    If bBrax Then m_StrB.AppendS "("
    If Len(s1) Then m_StrB.AppendS s1
    If Not ex1 Is Nothing Then ex1.ToFmt Me, ParentRank1
    If Len(s2) Then m_StrB.Append s2
    If Not ex2 Is Nothing Then ex2.ToFmt Me, ParentRank2
    If Len(s3) Then m_StrB.AppendS s3
    If bBrax Then m_StrB.AppendS ")"
End Sub

Private Sub FormatExpr_ExprFunctionToStr(aFunc As ExprFunction, ByVal ParentRank As OperatorRank)
    'more than 2 params
    Dim ex As Expression: Set ex = aFunc
    With ex
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
    End With
    m_StrB.AppendS GetFuncName(s) & "("
    s = GetFuncParamSep
    With aFunc
        Dim i As Long, c As Long: c = .params.Count
        For i = 1 To c
            Set ex = .params.Item(i)
            ex.ToFmt Me, ParentRank
            If i < c Then m_StrB.AppendS s
        Next
    End With
    m_StrB.AppendS ")"
    'Dim sParams  As String:  sParams = GetParamsList(aFunc.params, ParentRank)
    'FuncName & "(" & sParams & ")"
End Sub

'Private Function GetParamsList(col As Collection, ByVal ParentRank As OperatorRank) As String
'    Dim s As String, sp As String, sep As String: sep = GetFuncParamSep
'    Dim i As Long, c As Long: c = col.Count
'    Dim v, FuncParam As Expression
'    For Each v In col
'        i = i + 1
'        Set FuncParam = v
'        sp = FuncParam.ToStr(Me, ParentRank)
'        s = s & sp
'        If i < c Then s = s & sep
'    Next
'End Function

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

'Private Sub FormatExpr_ExprFunction2ToStr(aFunc As ExprFunction, ByVal ParentRank As OperatorRank)
Private Sub FormatExpr_ExprFunction2ToStr(aFunc As Expression, ByVal ParentRank As OperatorRank)
    'max 2 params
    'Dim ex As Expression: Set ex = aFunc
    'With ex 'aFunc
    With aFunc
        Dim r As OperatorRank: r = .Rank
        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
        Dim FuncName As String: FuncName = GetFuncName(s)
        s = FuncName & "("
        Dim ex2 As Expression: Set ex2 = .Expr2
        Dim sep As String
        If Not ex2 Is Nothing Then sep = GetFuncParamSep
        ExprToStrB False, s, .Expr1, r, sep, .Expr2, r, ")"
    End With
End Sub

Private Sub FormatExpr_ExprLitBolToStr(aLit As ExprLitBol, ByVal ParentRank As OperatorRank)
    m_StrB.AppendS aLit.Value
End Sub

Private Sub FormatExpr_ExprLitDatToStr(aLit As ExprLitDat, ByVal ParentRank As OperatorRank)
    m_StrB.AppendS aLit.Value
End Sub

Private Sub FormatExpr_ExprLitNumToStr(aLit As ExprLitNum, ByVal ParentRank As OperatorRank)
    m_StrB.AppendS Trim(Str(aLit.Value))
End Sub

Private Sub FormatExpr_ExprLitStrToStr(aLit As ExprLitStr, ByVal ParentRank As OperatorRank)
    m_StrB.AppendS aLit.Value
End Sub

Private Sub FormatExpr_ExprOp1DivXToStr(aExpr As ExprOp1DivX, ByVal ParentRank As OperatorRank)
    Dim ex As Expression: Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r + 1, ""
    End With
End Sub

Private Sub FormatExpr_ExprOpAbsToStr(aExpr As ExprOpAbs, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:    Set ex = aExpr
    With ex
        Dim r As OperatorRank: r = .Rank
        Dim bBrax As Boolean, sOp1 As String, sOp2 As String
        If m_XLCmpatbl Then
            sOp1 = "Abs("
            sOp2 = ")"
        Else
            bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
            sOp1 = .Operator1: If m_Condensed Then sOp1 = Trim(sOp1)
            sOp2 = .Operator2: If m_Condensed Then sOp2 = Trim(sOp2)
        End If
        ExprToStrB bBrax, sOp1, .Expr1, r, sOp2
    End With
End Sub

Private Sub FormatExpr_ExprOpAddToStr(aExpr As ExprOpAdd, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:    Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

''for operator binary, commutative: x + y, x * y
'Private Sub ExprOpBinToStr(aOpBin As Expression, ByVal ParentRank As OperatorRank)
'    With aOpBin
'        Dim r As OperatorRank: r = .Rank
'        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
'        Dim LHS As Expression: Set LHS = .Expr1
'        Dim RHS As Expression: Set RHS = .Expr2
'    End With
'    Dim b As Boolean:          b = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
'    If Not LHS Is Nothing Then s = IIf(b, "(", "") & LHS.ToStrB(Me, r) & s
'    If Not RHS Is Nothing Then s = s & RHS.ToStr(Me, r) & IIf(b, ")", "")
'    FormatExpr_ExprOpBinToStr = s
'End Sub

Private Sub FormatExpr_ExprOpBolAndToStr(aExpr As ExprOpBolAnd, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:       Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1   '"And" must not be condensed
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

Private Sub FormatExpr_ExprOpBolEqualToStr(aExpr As ExprOpBolEqual, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:         Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp As String:       sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

Private Sub FormatExpr_ExprOpBolGreaTEquToStr(aExpr As ExprOpBolGreaTEqu, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:            Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

Private Sub FormatExpr_ExprOpBolGreaTToStr(aExpr As ExprOpBolGreaT, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:         Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

Private Sub FormatExpr_ExprOpBolLessTEquToStr(aExpr As ExprOpBolLessTEqu, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:            Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

Private Sub FormatExpr_ExprOpBolLessTToStr(aExpr As ExprOpBolLessT, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:         Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

Private Sub FormatExpr_ExprOpBolNotEquToStr(aExpr As ExprOpBolNotEqu, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:          Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

Private Sub FormatExpr_ExprOpBolOrToStr(aExpr As ExprOpBolOr, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:      Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1 ' "Or" must not be condensed
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

Private Sub FormatExpr_ExprOpBolXorToStr(aExpr As ExprOpBolXor, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:       Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1 ' "Xor" must not be condensed
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

Private Sub FormatExpr_ExprOpCubToStr(aExpr As ExprOpCub, ByVal ParentRank As OperatorRank)
    FormatExpr_ExprOpPowToStr aExpr, ParentRank
'    Dim ex As Expression:    Set ex = aExpr
'    If m_XLCmpatbl Then
'        FormatExpr_ExprOpPowToStr aExpr, ParentRank
'        Exit Sub
'    End If
'    With ex
'        Dim r     As OperatorRank: r = .Rank
'        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
'        Dim sOp   As String:     sOp = .Operator2: If m_Condensed Then sOp = Trim(sOp)
'        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
'    End With
End Sub

Private Sub FormatExpr_ExprOpLNToStr(aExpr As ExprOpLN, ByVal ParentRank As OperatorRank)
    'FormatExpr_ExprFunctionToStr aExpr, ParentRank 'nope this will not work!
    FormatExpr_ExprFunction2ToStr aExpr, ParentRank
End Sub

Private Sub FormatExpr_ExprOpLog10ToStr(aExpr As ExprOpLog10, ByVal ParentRank As OperatorRank)
    FormatExpr_ExprFunction2ToStr aExpr, ParentRank
End Sub

Private Sub FormatExpr_ExprOpLogNToStr(aExpr As ExprOpLogN, ByVal ParentRank As OperatorRank)
    FormatExpr_ExprFunction2ToStr aExpr, ParentRank
End Sub

Private Sub FormatExpr_ExprOpMulToStr(aExpr As ExprOpMul, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:    Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

Private Sub FormatExpr_ExprOpNegToStr(aExpr As ExprOpNeg, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:    Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

Private Sub FormatExpr_ExprOpPow10ToStr(aExpr As ExprOpPow10, ByVal ParentRank As OperatorRank)
    FormatExpr_ExprOpPowToStr aExpr, ParentRank
'    Dim ex As Expression:      Set ex = aExpr
'    If m_XLCmpatbl Then
'        FormatExpr_ExprOpPowToStr aExpr, ParentRank
'        Exit Sub
'    End If
'    With ex
'        Dim r     As OperatorRank: r = .Rank
'        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
'        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
'        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
'    End With
End Sub

Private Sub FormatExpr_ExprOpPowXToStr(aExpr As ExprOpPow, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:    Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator2
        ExprToStrB bBrax, "", .Expr1, r, sOp 'nope this does not work
    End With
End Sub

Private Sub FormatExpr_ExprOpPowToStr(aExpr As ExprOpPow, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:    Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r + 1
    End With
End Sub

Private Sub FormatExpr_ExprOpSqrToStr(aExpr As ExprOpSqr, ByVal ParentRank As OperatorRank)
    FormatExpr_ExprOpPowToStr aExpr, ParentRank

'    Dim ex As Expression:    Set ex = aExpr
'    If m_XLCmpatbl Then
'        FormatExpr_ExprOpPowToStr aExpr, ParentRank
'        Exit Sub
'    End If
'    With ex
'        Dim r     As OperatorRank: r = .Rank
'        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
'        Dim sOp   As String:     sOp = .Operator2: If m_Condensed Then sOp = Trim(sOp)
'        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
'    End With
End Sub

Private Sub FormatExpr_ExprOpStrConcToStr(aExpr As ExprOpStrConc, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:        Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r
    End With
End Sub

'Differenz = Minuend - Subtrahend
'for operator binary -, non commutative: x - y
Private Sub FormatExpr_ExprOpSubtToStr(aExpr As ExprOpSubt, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:     Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r + 1
    End With
End Sub

'Quotient = Dividend / Divisor
'for operator, non commutativ: x / y
Private Sub FormatExpr_ExprOpDivToStr(aExpr As ExprOpDiv, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:    Set ex = aExpr
    With ex
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1: If m_Condensed Then sOp = Trim(sOp)
        ExprToStrB bBrax, "", .Expr1, r, sOp, .Expr2, r + 1
    End With
End Sub

'for operator unary associative left: -x
'Private Sub FormatExpr_ExprOpUnaAsLToStr(aExpr as exprOpUnAsL As Expression, ByVal ParentRank As OperatorRank)
'    With aOpUnAsL
'        Dim r As OperatorRank: r = .Rank
'        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
'        Dim RHS As Expression: Set RHS = .Expr2
'    End With
'    Dim b As Boolean: b = r <= ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
'    If Not RHS Is Nothing Then s = IIf(b, "(", "") & s & RHS.ToStr(Me, r) & IIf(b, ")", "")
'    FormatExpr_ExprOpUnaAsLToStr = s
'End Sub

'for operator unary associative right: x², x³, x!
'Private Sub FormatExpr_ExprOpUnaAsRToStr(aOpUnAsR As Expression, ByVal ParentRank As OperatorRank)
'    With aOpUnAsR
'        Dim r As OperatorRank: r = .Rank
'        Dim s As String:       s = .Operator1: If m_Condensed Then s = Trim(s)
'        Dim LHS As Expression: Set LHS = .Expr1
'    End With
'    'Dim b As Boolean: b = ParentRank < OperatorRank.Rank9ExprOpBrac And ParentRank <> OperatorRank.Rank0None
'    Dim b As Boolean: b = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
'    'Dim s As String:  s = myOp
'    If Not LHS Is Nothing Then s = IIf(b, "(", "") & LHS.ToStr(Me, r) & s & IIf(b, ")", "")
'    'If Not m_LHS Is Nothing Then s = m_LHS.ToStr(myRank) & s '& " = " & Trim(Str(Expression_Eval))
'    FormatExpr_ExprOpUnaAsRToStr = s
'End Sub

'for operator unary left/right | x |, ( x )
Private Sub FormatExpr_ExprOpBracToStr(aExpr As ExprOpBrac, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:     Set ex = aExpr
    With ex
        Dim r  As OperatorRank: r = .Rank
        Dim s1 As String:      s1 = .Operator1: If m_Condensed Then s1 = Trim(s1)
        Dim s2 As String:      s2 = .Operator2: If m_Condensed Then s2 = Trim(s2)
        ExprToStrB False, s1, .Expr1, r, "", , , s2
    End With
End Sub
'
'Private Sub FormatExpr_ExprFuncAbsToStr(aExprAbs As Expression, ByVal ParentRank As OperatorRank)
'    If m_XLCmpatbl Then
'        FormatExpr_ExprFuncAbsToStr = FormatExpr_ExprFunction2ToStr(MNew.ExprFunction("Abs", Nothing, aExprAbs.Expr1), ParentRank)
'        Exit Sub
'    End If
'    FormatExpr_ExprFuncAbsToStr = FormatExpr_ExprOpBracToStr(aExprAbs, ParentRank)
'End Sub

Private Sub FormatExpr_ExprOpFactToStr(aExpr As ExprOpFact, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:     Set ex = aExpr
    With ex
        If m_XLCmpatbl Then
            Set ex = MNew.ExprFunction("Fakultät", Nothing, .Expr1)
            ex.ToFmt Me, ParentRank
            Exit Sub
        End If
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1
        ExprToStrB bBrax, "", .Expr1, r, sOp
    End With
End Sub

Private Sub FormatExpr_ExprOpSqrtToStr(aExpr As ExprOpSqrt, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:     Set ex = aExpr
    With ex
        If m_XLCmpatbl Then
            Set ex = MNew.ExprFunction("Wurzel", Nothing, .Expr1)
            ex.ToFmt Me, ParentRank
            Exit Sub
        End If
        Dim r     As OperatorRank: r = .Rank
        Dim bBrax As Boolean:  bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        Dim sOp   As String:     sOp = .Operator1 & "("
        ExprToStrB bBrax, sOp, .Expr1, r, ")"
    End With
End Sub

Private Sub FormatExpr_ExprOpIIfToStr(aExpr As ExprOpIIf, ByVal ParentRank As OperatorRank)
    Dim ex As Expression:    Set ex = aExpr
    With ex
        Dim FuncName As String:         FuncName = IIf(m_XLCmpatbl, "Wenn", .Operator1)
        Dim exFunc   As ExprFunction: Set exFunc = MNew.ExprFunction(FuncName, Nothing, aExpr.Condition)
    End With
    With exFunc.params
        .Add ex.Expr1
        .Add ex.Expr2
    End With
    
    Set ex = exFunc
    ex.ToFmt Me, ParentRank
        'Dim r        As OperatorRank:  r = .Rank
        'Dim bBrax    As Boolean:   bBrax = r < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
        'Dim FuncName As String: FuncName = IIf(m_XLCmpatbl, "Wenn", .Operator1)
        'ExprToStrB bBrax, sOp, .Expr1, r, ")"
    'End With
    'Dim exFunc As ExprFunction: Set exFunc = MNew.ExprFunction(FuncName, Nothing)
    'exFunc.params.Add aOpIIf.Condition
    'exFunc.params.Add ex.Expr1:    exFunc.params.Add ex.Expr2
    'FormatExpr_ExprFunctionToStr exFunc, ParentRank
End Sub

Private Function FormatExpr_ToStr(Optional aExpr As Expression = Nothing) As String
    If aExpr Is Nothing Then
        FormatExpr_ToStr = m_StrB.ToStr
        Exit Function
    End If
    m_StrB.Append "="
    aExpr.ToFmt Me, OperatorRank.Rank0None
    m_StrB.AppendLine ""
    'm_StrB.Append(" = ").AppendLine aExpr.Eval
End Function

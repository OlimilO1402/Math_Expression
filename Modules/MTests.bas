Attribute VB_Name = "MTests"
Option Explicit
Private m_col As Collection

Public Sub Test1()
    Dim ex As Expression
    Set ex = MNew.ExprOpMul(MNew.ExprOpAdd(MNew.ExprLitNum(12), MNew.ExprLitNum(25)), MNew.ExprOpSubt(MNew.ExprLitNum(54), MNew.ExprLitNum(32)))
    MsgBox ex.ToStr & " = " & ex.Eval
    
    Set ex = MNew.ExprOpMul(MNew.ExprOpBrac(MNew.ExprOpAdd(MNew.ExprLitNum(12), MNew.ExprLitNum(25))), MNew.ExprOpSubt(MNew.ExprLitNum(54), MNew.ExprLitNum(32)))
    MsgBox ex.ToStr & " = " & ex.Eval
    
    Set ex = MNew.ExprOpPow(MNew.ExprOpAdd(MNew.ExprLitNum(2), MNew.ExprLitNum(3)), MNew.ExprOpSubt(MNew.ExprLitNum(54), MNew.ExprLitNum(51)))
    MsgBox ex.ToStr & " = " & ex.Eval
    
    Set ex = MNew.ExprOpPow(MNew.ExprOpBrac(MNew.ExprOpAdd(MNew.ExprLitNum(2), MNew.ExprLitNum(3))), MNew.ExprOpSubt(MNew.ExprLitNum(54), MNew.ExprLitNum(51)))
    MsgBox ex.ToStr & " = " & ex.Eval
    
    Set ex = MNew.ExprOpCub(MNew.ExprOpAdd(MNew.ExprLitNum(2), MNew.ExprLitNum(3)))
    MsgBox ex.ToStr & " = " & ex.Eval
    
    Set ex = MNew.ExprOpCub(MNew.ExprOpAdd(MNew.ExprLitNum(2), MNew.ExprLitNum(3)))
    MsgBox ex.ToStr & " = " & ex.Eval
    
    Set ex = MNew.ExprOpCub(MNew.ExprOpSqr(MNew.ExprOpAdd(MNew.ExprLitNum(2), MNew.ExprLitNum(3))))
    MsgBox ex.ToStr & " = " & ex.Eval
    
    Set ex = MNew.ExprOpCub(MNew.ExprOpSqr(MNew.ExprOpBrac(MNew.ExprOpAdd(MNew.ExprLitNum(2), MNew.ExprLitNum(3)))))
    MsgBox ex.ToStr & " = " & ex.Eval
    Dim op As Expression 'OperatorBinary
    Dim op1 As Expression 'OperatorBinary
    Dim op2 As Expression 'OperatorBinary
    
    Set op1 = MNew.OpBinAdd(MNew.ExprLitNum(3))
    Set op1.Expr2 = MNew.ExprLitNum(4)
    
    Set op2 = MNew.OpBinAdd(MNew.ExprLitNum(5))
    Set op2.Expr2 = MNew.ExprLitNum(6)
    
    Set op = MNew.ExprOpMul(MNew.ExprOpAdd(MNew.ExprLitNum(3), MNew.ExprLitNum(4)), MNew.ExprOpAdd(MNew.ExprLitNum(5), MNew.ExprLitNum(6)))
    Set op.Expr2 = op2
    
    Set ex = MNew.ExprOpBrac(MNew.ExprOpMul(MNew.ExprOpAdd(MNew.ExprLitNum(3), MNew.ExprLitNum(4)), MNew.ExprOpAdd(MNew.ExprLitNum(5), MNew.ExprLitNum(6))))
    MsgBox ex.ToStr & " = " & ex.Eval
    Set ex = MNew.ExprOpSqr(MNew.ExprOpBrac(MNew.ExprOpMul(MNew.ExprOpAdd(MNew.ExprLitNum(3), MNew.ExprLitNum(4)), MNew.ExprOpAdd(MNew.ExprLitNum(5), MNew.ExprLitNum(6)))))
    MsgBox ex.ToStr & " = " & ex.Eval
    Set ex = MNew.ExprOpSqr(op)
    MsgBox ex.ToStr & " = " & ex.Eval
    Set ex = MNew.ExprOpCub(MNew.ExprOpSqr(MNew.ExprOpBrac(MNew.ExprOpMul(MNew.ExprOpAdd(MNew.ExprLitNum(3), MNew.ExprLitNum(4)), MNew.ExprOpAdd(MNew.ExprLitNum(5), MNew.ExprLitNum(6))))))
    MsgBox ex.ToStr & " = " & ex.Eval
    Set ex = MNew.ExprOpSqrt(MNew.ExprOpCub(MNew.ExprOpSqr(MNew.ExprOpBrac(MNew.ExprOpMul(MNew.ExprOpAdd(MNew.ExprLitNum(3), MNew.ExprLitNum(4)), MNew.ExprOpAdd(MNew.ExprLitNum(5), MNew.ExprLitNum(6)))))))
    MsgBox ex.ToStr & " = " & ex.Eval
    
End Sub

Public Sub Test2()
    Dim ex As Expression
    Set ex = MNew.ExprLitNum(1.23)
    Set ex = MNew.ExprOp1DivX(ex)
    Set ex = MNew.ExprOpAdd(ex, MNew.ExprLitNum(4.56))
    Set ex = MNew.ExprOpBrac(ex)
    Set ex = MNew.ExprOpCub(ex)
    Set ex = MNew.ExprOpDiv(ex, MNew.ExprLitNum(-7.89))
    Set ex = MNew.ExprOpAbs(ex)
    Set ex = MNew.ExprOpFact(ex)
    Set ex = MNew.ExprOpLN(ex)
    Set ex = MNew.ExprOpLog10(ex)
    Set ex = MNew.ExprOpMul(ex, MNew.ExprLitNum(8.9))
    Set ex = MNew.ExprOpNeg(ex)
    Set ex = MNew.ExprOpPow(ex, MNew.ExprLitNum(2))
    Set ex = MNew.ExprOpPow10(ex)
    Set ex = MNew.ExprOpSqrt(ex)
    Set ex = MNew.ExprOpSqr(ex)
    Set ex = MNew.ExprOpSubt(ex, MNew.ExprOpPow10(MNew.ExprLitNum(209)))
    'MsgBox ex.ToStr & " = " & ex.Eval
    
    Set ex = MNew.ExprLitNum(12)
    Set ex = MNew.ExprOpSubt(ex, MNew.ExprLitNum(6))
    Set ex = MNew.ExprOpSqrt(ex)
    Set ex = MNew.ExprOpSqr(ex)
    MsgBox ex.ToStr & " = " & ex.Eval
    
    Set ex = MNew.ExprLitNum(12)
    Set ex = MNew.ExprOpBrac(ex)
    Set ex = MNew.ExprOpSubt(ex, MNew.ExprLitNum(6))
    Set ex = MNew.ExprOpBrac(ex)
    Set ex = MNew.ExprOpSqr(ex)
    Set ex = MNew.ExprOpBrac(ex)
    Set ex = MNew.ExprOpSqrt(ex)
    MsgBox ex.ToStr & " = " & ex.Eval
    
    Set ex = MNew.ExprLitNum(12)
    Set ex = MNew.ExprOpSqr(ex)
    Set ex = MNew.ExprOpCub(ex)
    MsgBox ex.ToStr & " = " & ex.Eval
    
    Set ex = MNew.ExprLitNum(12)
    Set ex = MNew.ExprOpCub(ex)
    Set ex = MNew.ExprOpSqr(ex)
    MsgBox ex.ToStr & " = " & ex.Eval
    
End Sub

Public Function Test3(Fmt As FormatExpr) As String
    
'    Dim FR As FormatRPN: Set FR = New FormatRPN
'
'    Dim ex As Expression
'
'    Dim ex1 As Expression: Set ex1 = MNew.ExprOpAdd(MNew.ExprLitNum(1), MNew.ExprLitNum(2))
'    Dim ex2 As Expression: Set ex2 = MNew.ExprOpSubt(MNew.ExprLitNum(4), MNew.ExprLitNum(3))
'
'    Set ex = MNew.ExprOpAdd(ex1, ex2)
'    MsgBox ex.ToStr(FA) & " = " & ex.Eval
'    'MsgBox ex.ToStr(FR) & " = " & ex.Eval
'
'    Set ex = MNew.ExprOpSubt(ex1, ex2)
'    MsgBox ex.ToStr(FA) & " = " & ex.Eval
'    'MsgBox ex.ToStr(FR) & " = " & ex.Eval
'
'    Set ex = MNew.ExprOpMul(ex1, ex2)
'    MsgBox ex.ToStr(FA) & " = " & ex.Eval
'    'MsgBox ex.ToStr(FR) & " = " & ex.Eval
'
'    Set ex = MNew.ExprOpDiv(ex1, ex2)
'    MsgBox ex.ToStr(FA) & " = " & ex.Eval
'    'MsgBox ex.ToStr(FR) & " = " & ex.Eval
    'Dim FA As FormatAlg: Set FA = MNew.FormatAlg(True)
    Dim forExcel As Boolean: forExcel = True
    Set m_col = GetListOfBinaryExpressions
    Dim s As String, i As Long
    Dim ex As Expression
    For i = 1 To m_col.Count
        Set ex = m_col.Item(i)
        If forExcel Then s = s & "="
        s = s & ex.ToStr(Fmt) & vbCrLf
    Next
    Test3 = s
End Function

' OK eine Testroutine mit allen Operatoren
' '
Public Function GetListOfBinaryExpressions() As Collection
    Dim n As Long: n = 10
    Dim exList As New Collection
    Dim ex1 As Expression: Set ex1 = MNew.ExprLitNum(2)
    Dim ex2 As Expression: Set ex2 = MNew.ExprLitNum(3)
    Dim ex3 As Expression: Set ex3 = MNew.ExprLitNum(4)
    Dim ex4 As Expression: Set ex4 = MNew.ExprLitNum(5)
    Dim ex  As Expression
    Dim exL  As Expression
    Dim exR  As Expression
    Dim i As Long, j As Long, k As Long
    For i = 1 To n
        For j = 1 To n
            For k = 1 To n
                Set exL = GetEx(k, ex4, ex3)
                Set exR = GetEx(j, ex2, ex1)
                Set ex = GetEx(i, exL, exR)
                exList.Add ex
            Next
        Next
    Next
    Set GetListOfBinaryExpressions = exList
End Function

Private Function GetEx(ByVal e As Long, exL As Expression, exR As Expression) As Expression
    Static toggleAbs As Boolean
    Static toggle1dx As Boolean
    Static toggleBrc As Boolean
    Static toggleSqr As Boolean
    Static toggleCub As Boolean
    Select Case e
    Case 1:  Set GetEx = MNew.ExprOpAdd(exL, exR)
    Case 2:  Set GetEx = MNew.ExprOpSubt(exL, exR)
    Case 3:  Set GetEx = MNew.ExprOpMul(exL, exR)
    Case 4:  Set GetEx = MNew.ExprOpDiv(exL, exR)
    Case 5:  toggleAbs = Not toggleAbs: If toggleAbs Then Set GetEx = MNew.ExprOpAbs(exL) Else Set GetEx = MNew.ExprOpAbs(exR)
    Case 6:  toggle1dx = Not toggle1dx: If toggle1dx Then Set GetEx = MNew.ExprOp1DivX(exL) Else Set GetEx = MNew.ExprOp1DivX(exR)
    Case 7:  toggleBrc = Not toggleBrc: If toggleBrc Then Set GetEx = MNew.ExprOpBrac(exL) Else Set GetEx = MNew.ExprOpBrac(exR)
    Case 8:  toggleSqr = Not toggleSqr: If toggleSqr Then Set GetEx = MNew.ExprOpSqr(exL) Else Set GetEx = MNew.ExprOpSqr(exR)
    Case 9:  toggleCub = Not toggleCub: If toggleCub Then Set GetEx = MNew.ExprOpCub(exL) Else Set GetEx = MNew.ExprOpCub(exR)
    Case 10: Set GetEx = MNew.ExprOpPow(exL, exR)
    
    End Select
End Function

Sub TestVBCode()
    '
End Sub

Public Function GetResults() As String
    Dim s As String, i As Long
    Dim ex As Expression
    For i = 1 To m_col.Count
        Set ex = m_col.Item(i)
        s = s & ex.Eval & vbCrLf
    Next
    GetResults = s
End Function


Attribute VB_Name = "MTests"
Option Explicit

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

Public Function Test3() As String
    
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
    Dim FA As FormatAlg: Set FA = New FormatAlg
    Dim col As Collection: Set col = GetListOfBinaryExpressions
    Dim s As String, i As Long
    Dim ex As Expression
    For i = 1 To col.Count
        Set ex = col.Item(i)
        s = s & "=" & ex.ToStr(FA) & vbCrLf
    Next
    Test3 = s
End Function

' OK eine Testroutine mit allen Operatoren
' '
Public Function GetListOfBinaryExpressions() As Collection
    Dim exList As New Collection
    Dim ex1 As Expression: Set ex1 = MNew.ExprLitNum(1)
    Dim ex2 As Expression: Set ex2 = MNew.ExprLitNum(2)
    Dim ex3 As Expression: Set ex3 = MNew.ExprLitNum(3)
    Dim ex4 As Expression: Set ex4 = MNew.ExprLitNum(4)
    Dim ex  As Expression
    Dim exL  As Expression
    Dim exR  As Expression
    Dim i As Long, j As Long, k As Long
    For i = 1 To 4
        For j = 1 To 4
            For k = 1 To 4
                Set exL = GetEx(j, ex1, ex2)
                Set exR = GetEx(k, ex4, ex3)
                Set ex = GetEx(i, exL, exR)
                exList.Add ex
            Next
        Next
    Next
    Set GetListOfBinaryExpressions = exList
End Function

Private Function GetEx(ByVal e As Long, exL As Expression, exR As Expression) As Expression
    Select Case e
    Case 1: Set GetEx = MNew.ExprOpAdd(exL, exR)
    Case 2: Set GetEx = MNew.ExprOpSubt(exL, exR)
    Case 3: Set GetEx = MNew.ExprOpMul(exL, exR)
    Case 4: Set GetEx = MNew.ExprOpDiv(exL, exR)
    End Select
End Function

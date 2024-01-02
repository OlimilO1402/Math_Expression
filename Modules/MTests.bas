Attribute VB_Name = "MTests"
Option Explicit
Private m_Col As Collection

Private toggleAbs  As Boolean
Private toggle1dx  As Boolean
Private toggleBrc  As Boolean
Private toggleSqr  As Boolean
Private toggleCub  As Boolean
Private toggleFac  As Boolean
Private toggleSqrt As Boolean
Private toggleLN   As Boolean
Private toggleLg10 As Boolean
Private toggleLgN  As Boolean
Private ToggleNeg  As Boolean

Private Sub InitToggles()
    toggleAbs = False
    toggle1dx = False
    toggleBrc = False
    toggleSqr = False
    toggleCub = False
    toggleFac = False
    toggleSqrt = False
    toggleLN = False
    toggleLg10 = False
    toggleLgN = False
    ToggleNeg = False
End Sub

Public Function Test3(Optional Fmt As FormatExpr = Nothing) As String
    InitToggles
    If m_Col Is Nothing Then Set m_Col = GetListOfBinaryExpressions
    Dim i As Long
    Dim ex As Expression
    If Fmt Is Nothing Then
        Dim sb As StringBuilder: Set sb = New StringBuilder
        For i = 1 To m_Col.Count
            Set ex = m_Col.Item(i)
            sb.AppendLine ex.ToStr & vbTab & "'=" & vbTab & ex.Eval
        Next
        Test3 = sb.ToStr
        Exit Function
    End If
    
    For i = 1 To m_Col.Count
        Set ex = m_Col.Item(i)
        Fmt.ToStr ex
    Next
    Test3 = Fmt.ToStr
End Function

' OK eine Testroutine mit allen Operatoren
Public Function GetListOfBinaryExpressions() As Collection
    Dim n As Long: n = 17
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
    Case 11: toggleFac = Not toggleFac:   If toggleFac Then Set GetEx = MNew.ExprOpFact(exL) Else Set GetEx = MNew.ExprOpFact(exR)
    Case 12: toggleSqrt = Not toggleSqrt: If toggleSqrt Then Set GetEx = MNew.ExprOpSqrt(exL) Else Set GetEx = MNew.ExprOpSqrt(exR)
    Case 13: toggleLN = Not toggleLN:     If toggleLN Then Set GetEx = MNew.ExprOpLN(exL) Else Set GetEx = MNew.ExprOpLN(exR)
    Case 14: toggleLg10 = Not toggleLg10: If toggleLg10 Then Set GetEx = MNew.ExprOpLog10(exL) Else Set GetEx = MNew.ExprOpLog10(exR)
    Case 15: toggleLgN = Not toggleLgN:   If toggleLgN Then Set GetEx = MNew.ExprOpLogN(exL) Else Set GetEx = MNew.ExprOpLogN(exR)
    Case 16: Set GetEx = MNew.ExprOpLogN(exL, exR)
    Case 17: ToggleNeg = Not ToggleNeg:  If ToggleNeg Then Set GetEx = MNew.ExprOpNeg(exL) Else Set GetEx = MNew.ExprOpNeg(exR)
    End Select
End Function

Public Function GetResults() As String
    If m_Col Is Nothing Then Set m_Col = GetListOfBinaryExpressions
    Dim s As String, i As Long
    Dim ex As Expression
    For i = 1 To m_Col.Count
        Set ex = m_Col.Item(i)
        s = s & ex.Eval & vbCrLf
    Next
    GetResults = s
End Function


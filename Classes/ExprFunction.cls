VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Expression
'
'
Private m_Params As Collection 'Of Expression
Private Const myRank As Long = OperatorRank.Rank9ExprOpBrac
Private m_FncName As String
Private m_Object  As Object
Private m_ErrInfo As String

Private Sub Class_Initialize():    Set m_Params = New Collection: End Sub
Friend Sub New_(aFuncName As String, CallableObject As Object, Optional ExprInside As Expression = Nothing)
    m_FncName = aFuncName: Set m_Object = CallableObject: If Not ExprInside Is Nothing Then m_Params.Add ExprInside
End Sub

Public Property Get params() As Collection 'Of Expression
    Set params = m_Params
End Property

Public Property Set params(Parameters As Collection) 'Of Expression
    Set m_Params = Parameters
End Property

Public Property Get CountParams() As Byte
    CountParams = m_Params.Count
End Property

' v ############################## v '   Implements Expression    ' v ############################## v '
Private Property Get Expression_CanEval() As Boolean
    Dim v, ex As Expression
    For Each v In m_Params
        Set ex = v: If Not ex.CanEval Then Exit Property
    Next
    Expression_CanEval = True
End Property

Private Property Get Expression_ErrInfo() As String
    Expression_ErrInfo = IIf(Len(m_ErrInfo), TypeName(Me) & " ", "") & m_ErrInfo
End Property

Private Function Expression_Eval() 'As Double
Try: On Error GoTo Catch
    If Not Expression_CanEval Then Exit Function
    Dim c As Long: c = m_Params.Count
    Dim Args(), ex As Expression
    If c = 1 Then
        ReDim Args(1 To 1)
        Set ex = m_Params.Item(1): Args(0) = ex.Eval
    ElseIf c = 2 Then
        ReDim Args(1 To 2)
        Set ex = m_Params.Item(1): Args(0) = ex.Eval
        Set ex = m_Params.Item(2): Args(1) = ex.Eval
    Else
        ReDim Args(1 To m_Params.Count)
        Dim i As Long
        For i = 1 To m_Params.Count
            Set ex = m_Params.Item(i)
            If Not ex Is Nothing Then
                Args(i) = ex.Eval
            End If
        Next
    End If
    Dim Ret
    Select Case c
    Case 1: Ret = CallByName(m_Object, m_FncName, VbMethod, Args(1))
    Case 2: Ret = CallByName(m_Object, m_FncName, VbMethod, Args(1), Args(2))
    Case 3: Ret = CallByName(m_Object, m_FncName, VbMethod, Args(1), Args(2), Args(3))
    Case 4: Ret = CallByName(m_Object, m_FncName, VbMethod, Args(1), Args(2), Args(3), Args(4))
    Case 5: Ret = CallByName(m_Object, m_FncName, VbMethod, Args(1), Args(2), Args(3), Args(4), Args(5))
    Case 6: Ret = CallByName(m_Object, m_FncName, VbMethod, Args(1), Args(2), Args(3), Args(4), Args(5), Args(6))
    Case 7: Ret = CallByName(m_Object, m_FncName, VbMethod, Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7))
    Case 8: Ret = CallByName(m_Object, m_FncName, VbMethod, Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8))
    End Select
    If IsObject(Ret) Then
        Set Expression_Eval = Ret
    Else
        Expression_Eval = Ret
    End If
    Exit Function
Catch:
    m_ErrInfo = Err.number & " " & Err.Description
End Function

Private Property Set Expression_Expr1(ByVal RHS As Expression)
    Dim c As Long: c = m_Params.Count
    If c = 0 Then
        m_Params.Add RHS
    Else
        m_Params.Remove 1
        m_Params.Add RHS, , 1
    End If
End Property

Private Property Get Expression_Expr1() As Expression
    If m_Params.Count >= 1 Then Set Expression_Expr1 = m_Params.Item(1)
End Property

Private Property Set Expression_Expr2(ByVal RHS As Expression)
    Dim c As Long: c = m_Params.Count
    If c = 0 Then
        m_Params.Add RHS
        m_Params.Add RHS
    ElseIf c = 1 Then
        m_Params.Add RHS
    ElseIf c >= 2 Then
        m_Params.Remove 2
        m_Params.Add RHS, , 2
    End If
End Property

Private Property Get Expression_Expr2() As Expression
    If m_Params.Count >= 2 Then Set Expression_Expr2 = m_Params.Item(2)
End Property

Private Property Get Expression_Operator1() As String
    Expression_Operator1 = m_FncName
End Property

Private Property Get Expression_Operator2() As String
    '
End Property

Private Property Get Expression_Rank() As OperatorRank
    Expression_Rank = myRank
End Property

Private Function Expression_ToStr(ByVal Fmt As FormatExpr, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    Expression_ToStr = Fmt.ExprFunctionToStr(Me, ParentRank)
End Function




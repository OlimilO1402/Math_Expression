VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprOp1DivX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Expression
'
Private m_LHS As Expression
Private m_RHS As Expression
Private Const myRank As Long = OperatorRank.Rank4ExprOpMulDiv
Private Const myOp   As String = " / "
Private m_ErrInfo As String

Private Sub Class_Initialize():    Set m_LHS = MNew.ExprLitNum(1): End Sub
Friend Sub New_(aRHS As Expression)
    Set m_RHS = aRHS
End Sub

' v ############################## v '   Implements Expression    ' v ############################## v '
Private Property Get Expression_CanEval() As Boolean
    '
    If m_RHS Is Nothing Then Exit Function
    Expression_CanEval = m_RHS.CanEval
End Property

Private Property Get Expression_ErrInfo() As String
    Dim s1 As String: s1 = m_LHS.ErrInfo: s1 = s1 & IIf(Len(s1), vbCrLf, "")
    Dim s2 As String: s2 = m_RHS.ErrInfo: s2 = s2 & IIf(Len(s2), vbCrLf, "")
    Expression_ErrInfo = s1 & s2 & IIf(Len(m_ErrInfo), TypeName(Me) & " ", "") & m_ErrInfo
End Property

Private Function Expression_Eval() 'As Double
Try: On Error GoTo Catch
    If Not Expression_CanEval Then Exit Function
    Expression_Eval = 1 / m_RHS.Eval
    Exit Function
Catch:
    m_ErrInfo = Err.number & " " & Err.Description
End Function


Private Property Set Expression_Expr1(ByVal RHS As Expression)
    Set m_LHS = RHS
End Property
Private Property Get Expression_Expr1() As Expression
    Set Expression_Expr1 = m_LHS
End Property


Private Property Set Expression_Expr2(ByVal RHS As Expression)
    Set m_RHS = RHS
End Property
Private Property Get Expression_Expr2() As Expression
    Set Expression_Expr2 = m_RHS
End Property


Private Property Get Expression_Operator1() As String
    Expression_Operator1 = myOp
End Property
Private Property Get Expression_Operator2() As String
    '
End Property

Private Property Get Expression_Rank() As OperatorRank
    Expression_Rank = myRank
End Property

Private Function Expression_ToStr(Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    '
End Function

Private Sub Expression_ToFmt(ByVal Fmt As FormatExpr, ByVal ParentRank As OperatorRank)
    Fmt.ExprOp1DivXToStr Me, ParentRank
End Sub



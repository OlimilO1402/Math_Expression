VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprOpCub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Expression
'
Private m_LHS As Expression
Private m_RHS As Expression '
Private Const myRank As Long = OperatorRank.Rank6ExprOpPow
Private Const myOp   As String = "³"
'
Private Sub Class_Initialize():    Set m_RHS = MNew.ExprLitNum(3): End Sub
Friend Sub New_(aLHS As Expression)
    Set m_LHS = aLHS
End Sub

' v ############################## v '   Implements Expression    ' v ############################## v '
Private Property Get Expression_CanEval() As Boolean
    If m_LHS Is Nothing Then Exit Property
    '
    Expression_CanEval = m_LHS.CanEval
End Property

Private Function Expression_Eval()
    If Not Expression_CanEval Then Exit Function
    Expression_Eval = m_LHS.Eval ^ 3
End Function

Private Property Get Expression_Expr1() As Expression
    Set Expression_Expr1 = m_LHS
End Property

Private Property Get Expression_Expr2() As Expression
    Set Expression_Expr2 = m_RHS
End Property

Private Property Set Expression_Expr2(ByVal RHS As Expression)
    Set m_LHS = RHS
End Property

Private Property Get Expression_Operator1() As String
    Expression_Operator1 = " ^ "
End Property

Private Property Get Expression_Operator2() As String
    Expression_Operator2 = myOp
End Property

Private Property Get Expression_Rank() As OperatorRank
    Expression_Rank = myRank
End Property

Private Function Expression_ToStr(ByVal Fmt As FormatExpr, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
'    Dim b As Boolean: b = myRank <= ParentRank 'And ParentRank = OperatorRank.Rank7ExprOpBrac
'    Dim s As String:  s = myOp
'    If Not m_LHS Is Nothing Then s = IIf(b, "(", "") & m_LHS.ToStr(myRank) & s & IIf(b, ")", "") '& " = " & Trim(Str(Expression_Eval))
'    'If Not m_LHS Is Nothing Then s = m_LHS.ToStr(myRank) & s '& " = " & Trim(Str(Expression_Eval))
'    Expression_ToStr = s
    Expression_ToStr = Fmt.ExprOpCubToStr(Me, ParentRank)
End Function

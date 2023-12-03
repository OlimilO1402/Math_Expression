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
Private m_Ins As Expression
Private Const myRank As Long = OperatorRank.Rank9ExprOpBrac
Private m_FncName As String
Private m_Object  As Object

Friend Sub New_(aFuncName As String, CallableObject As Object, Optional ExprInside As Expression = Nothing)
    m_FncName = aFuncName: Set m_Object = CallableObject: Set m_Ins = ExprInside
End Sub

' v ############################## v '   Implements Expression    ' v ############################## v '
Private Property Get Expression_CanEval() As Boolean
    '
    If m_Ins Is Nothing Then Exit Function
    Expression_CanEval = m_Ins.CanEval
End Property

Private Function Expression_Eval() 'As Double
    If Not Expression_CanEval Then Exit Function
    Expression_Eval = CallByName(m_Object, m_FncName, VbMethod, m_Ins.Eval)
End Function

Private Property Get Expression_Expr1() As Expression
    Set Expression_Expr1 = m_Ins
End Property

Private Property Get Expression_Expr2() As Expression
    Set Expression_Expr2 = m_Ins
End Property

Private Property Set Expression_Expr2(ByVal RHS As Expression)
    Set m_Ins = RHS
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




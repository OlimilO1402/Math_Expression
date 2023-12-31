VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprLitBol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Expression

Private m_Value As Boolean
'
Private Const myRank As Long = OperatorRank.Rank1ExprConst

'
Friend Sub New_(aValue)
    m_Value = CBool(aValue)
End Sub
Public Property Get Value() As Boolean:    Value = m_Value:    End Property

' v ############################## v '   Implements Expression    ' v ############################## v '
Private Property Get Expression_CanEval() As Boolean
    '
    '
    Expression_CanEval = True
End Property

Private Property Get Expression_ErrInfo() As String
    '
End Property

Private Function Expression_Eval()
    '
    Expression_Eval = m_Value
End Function


Private Property Set Expression_Expr1(ByVal RHS As Expression)
    '
End Property
Private Property Get Expression_Expr1() As Expression
    '
End Property


Private Property Set Expression_Expr2(ByVal RHS As Expression)
    '
End Property
Private Property Get Expression_Expr2() As Expression
    '
End Property


Private Property Get Expression_Operator1() As String
    '
End Property
Private Property Get Expression_Operator2() As String
    '
End Property


Private Property Get Expression_Rank() As OperatorRank
    Expression_Rank = myRank
End Property

Private Function Expression_ToStr(Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    Expression_ToStr = CStr(m_Value)
End Function

Private Sub Expression_ToFmt(ByVal Fmt As FormatExpr, ByVal ParentRank As OperatorRank)
    Fmt.ExprLitBolToStr Me, ParentRank
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprOpFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Expression
'
Private m_LHS As Expression
'
Private Const myRank As Long = OperatorRank.Rank5ExprOpFact
Private Const myOp   As String = "!"

Friend Sub New_(aLHS As Expression)
    Set m_LHS = aLHS
End Sub

' v ############################## v '   Implements Expression    ' v ############################## v '
Private Function Expression_Eval() 'As Double
    If Not Expression_CanEval Then Exit Function
    Expression_Eval = MMath.Fact(m_LHS.Eval)
End Function

Private Property Get Expression_CanEval() As Boolean
    If m_LHS Is Nothing Then Exit Function
    Expression_CanEval = True
End Property

Private Property Get Expression_Rank() As OperatorRank
    Expression_Rank = myRank
End Property

Private Function Expression_ToStr(Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    Dim b As Boolean: b = myRank < ParentRank And ParentRank <> OperatorRank.Rank7ExprOpBrac
    Dim s As String:  s = myOp
    If Not m_LHS Is Nothing Then s = IIf(b, "(", "") & m_LHS.ToStr(myRank) & s & IIf(b, ")", "") '& " = " & Trim(Str(Expression_Eval))
    Expression_ToStr = s
End Function


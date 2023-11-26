VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprOpBrac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Expression
Implements OperatorBinary
Private m_Ins As Expression
'
Private Const myRank As Long = OperatorRank.Rank7ExprOpBrac
'

Friend Sub New_(aInside As Expression)
    Set m_Ins = aInside
End Sub

' v ############################## v '   Implements Expression    ' v ############################## v '
Private Function Expression_Eval() 'As Double
    If Not Expression_CanEval Then Exit Function
    Expression_Eval = m_Ins.Eval
End Function

Private Property Get Expression_CanEval() As Boolean
    If m_Ins Is Nothing Then Exit Function
    '
    Expression_CanEval = True
End Property

Private Property Get Expression_Rank() As OperatorRank
    Expression_Rank = myRank
End Property

Private Function Expression_ToStr(Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    Dim b As Boolean: b = ParentRank < myRank And myRank <> m_Ins.Rank
    Dim s As String:  s = IIf(b, "(", "")
    If Not m_Ins Is Nothing Then s = s & m_Ins.ToStr(myRank)
    s = s & IIf(b, ")", "")
    Expression_ToStr = s
End Function

' v ############################## v '   Implements OperatorBinary    ' v ############################## v '
Private Property Set OperatorBinary_SecondExpr(ByVal RHS As Expression)
    Set m_Ins = RHS
End Property



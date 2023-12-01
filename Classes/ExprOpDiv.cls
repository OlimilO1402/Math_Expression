VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprOpDiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Expression
'Implements OperatorBinary
Private m_LHS As Expression
Private m_RHS As Expression
Private Const myRank As Long = OperatorRank.Rank4ExprOpMulDiv
Private Const myOp   As String = " / "
'

Friend Sub New_(aLHS As Expression)
    Set m_LHS = aLHS
End Sub

' v ############################## v '   Implements Expression    ' v ############################## v '
Private Property Get Expression_CanEval() As Boolean
    If m_LHS Is Nothing Then Exit Function
    If m_RHS Is Nothing Then Exit Function
    Expression_CanEval = m_LHS.CanEval And m_RHS.CanEval ' True
End Property

Private Function Expression_Eval() 'As Double
    If Not Expression_CanEval Then Exit Function
    Dim r As Double: r = m_RHS.Eval
    If r = 0 Then MsgBox "Error undefined division: 0": Exit Function
    Expression_Eval = m_LHS.Eval / r
End Function

Private Property Get Expression_Expr1() As Expression
    Set Expression_Expr1 = m_LHS
End Property

Private Property Get Expression_Expr2() As Expression
    Set Expression_Expr2 = m_RHS
End Property

Private Property Set Expression_Expr2(ByVal RHS As Expression)
    Set m_RHS = RHS
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

Private Function Expression_ToStr(ByVal Fmt As FormatExpr, Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
'    Dim b As Boolean: b = myRank < ParentRank And ParentRank <> OperatorRank.Rank7ExprOpBrac
'    Dim s As String:  s = myOp
'    If Not m_LHS Is Nothing Then s = IIf(b, "(", "") & m_LHS.ToStr(myRank) & s
'    Dim bDiv As Boolean: bDiv = m_RHS.Rank <> OperatorRank.Rank1ExprConst And m_RHS.Rank <> OperatorRank.Rank7ExprOpBrac
'    If Not m_RHS Is Nothing Then s = s & IIf(bDiv, "(", "") & m_RHS.ToStr(myRank) & IIf(bDiv, ")", "") & IIf(b, ")", "") '& " = " & Trim(Str(Expression_Eval))
'    Expression_ToStr = s
    Expression_ToStr = Fmt.ExprOpDivToStr(Me, ParentRank)
End Function

' v ############################## v '   Implements OperatorBinary    ' v ############################## v '
'Private Property Get OperatorBinary_Operator() As String
'    OperatorBinary_Operator = myOp
'End Property
'
'Private Property Get OperatorBinary_LHSExpr() As Expression
'    Set OperatorBinary_LHSExpr = m_LHS
'End Property
'
'Private Property Get OperatorBinary_RHSExpr() As Expression
'    Set OperatorBinary_RHSExpr = m_RHS
'End Property
'
'Private Property Set OperatorBinary_RHSExpr(ByVal RHS As Expression)
'    Set m_RHS = RHS
'End Property
'
'

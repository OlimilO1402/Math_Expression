VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprOpLog10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Expression
'
'
Private m_RHS As Expression
Private Const myRank As Long = OperatorRank.Rank9ExprOpBrac
Private Const myOp   As String = "Log10"
'

Friend Sub New_(aRHS As Expression)
    Set m_RHS = aRHS
End Sub

Public Property Get CountParams() As Byte
    CountParams = 1
End Property

' v ############################## v '   Implements Expression    ' v ############################## v '
Private Property Get Expression_CanEval() As Boolean
    '
    If m_RHS Is Nothing Then Exit Function
    Expression_CanEval = m_RHS.CanEval
End Property

Private Function Expression_Eval()
Try: On Error GoTo Catch
    If Not Expression_CanEval Then Exit Function
    Expression_Eval = MMath.Log10(m_RHS.Eval)
    Exit Function
Catch:
    MsgBox Err.number & " " & Err.Description
End Function

Private Property Set Expression_Expr1(ByVal RHS As Expression)
    Set m_RHS = RHS
End Property
Private Property Get Expression_Expr1() As Expression
    Set Expression_Expr1 = m_RHS
End Property

Private Property Set Expression_Expr2(ByVal RHS As Expression)
    'Set m_RHS = RHS
End Property
Private Property Get Expression_Expr2() As Expression
    'Set Expression_Expr2 = m_RHS
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
    Expression_ToStr = Fmt.ExprFunction2ToStr(Me, ParentRank)
End Function

'    Dim b As Boolean: b = myRank < ParentRank And ParentRank <> OperatorRank.Rank7ExprOpBrac
'    Dim s As String:  s = myOp
'    If Not m_RHS Is Nothing Then s = IIf(b, "(", "") & myOp & m_RHS.ToStr(myRank) & ")" & IIf(b, ")", "") '& " = " & Trim(Str(Expression_Eval))
'    Expression_ToStr = s



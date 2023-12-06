VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprOpIIf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Expression
Private m_Condition  As Expression
Private m_ExprTrue   As Expression
Private m_ExprFalse  As Expression
Private Const myOp   As String = "IIf"
Private Const myRank As Long = OperatorRank.Rank9ExprOpBrac
Private m_ErrInfo As String

'
Friend Sub New_(aCond As Expression, Optional ExprTrue As Expression = Nothing, Optional ExprFalse As Expression = Nothing)
    Set m_Condition = aCond: Set m_ExprTrue = ExprTrue:                Set m_ExprFalse = ExprFalse
End Sub

Public Property Set Condition(RHS As Expression)
    Set m_Condition = RHS
End Property

Public Property Get Condition() As Expression
    Set Condition = m_Condition
End Property

' v ############################## v '   Implements Expression    ' v ############################## v '
Private Property Get Expression_CanEval() As Boolean
    If m_Condition Is Nothing Then Exit Property
    If m_ExprTrue Is Nothing Then Exit Property
    If m_ExprFalse Is Nothing Then Exit Property
    Expression_CanEval = m_Condition.CanEval And m_ExprTrue.CanEval And m_ExprFalse.CanEval
End Property

Private Property Get Expression_ErrInfo() As String
    Dim s1 As String: s1 = m_Condition.ErrInfo: s1 = s1 & IIf(Len(s1), vbCrLf, "")
    Dim s2 As String: s2 = m_ExprTrue.ErrInfo:  s2 = s2 & IIf(Len(s2), vbCrLf, "")
    Dim s3 As String: s3 = m_ExprFalse.ErrInfo: s3 = s3 & IIf(Len(s3), vbCrLf, "")
    Expression_ErrInfo = s1 & s2 & s3 & IIf(Len(m_ErrInfo), TypeName(Me) & " ", "") & m_ErrInfo
End Property

Private Function Expression_Eval()
Try: On Error GoTo Catch
    If Not Expression_CanEval Then Exit Function
    If m_Condition.Eval Then
        Expression_Eval = m_ExprTrue.Eval
    Else
        Expression_Eval = m_ExprFalse.Eval
    End If
    Exit Function
Catch:
    m_ErrInfo = Err.number & " " & Err.Description
End Function


Private Property Set Expression_Expr1(ByVal RHS As Expression)
    Set m_ExprTrue = RHS
End Property
Private Property Get Expression_Expr1() As Expression
    Set Expression_Expr1 = m_ExprTrue
End Property


Private Property Set Expression_Expr2(ByVal RHS As Expression)
    Set m_ExprFalse = RHS
End Property
Private Property Get Expression_Expr2() As Expression
    Set Expression_Expr2 = m_ExprFalse
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
    Expression_ToStr = Fmt.ExprOpIIfToStr(Me, ParentRank)
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprOpAbs"
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
Private Const myOp1  As String = "| "
Private Const myOp2  As String = " |"
Private m_ErrInfo As String


Friend Sub New_(aIns As Expression)
    Set m_Ins = aIns
End Sub

Public Property Get CountParams() As Byte
    CountParams = 1
End Property

' v ############################## v '   Implements Expression    ' v ############################## v '
Private Property Get Expression_CanEval() As Boolean
    '
    If m_Ins Is Nothing Then Exit Function
    Expression_CanEval = m_Ins.CanEval
End Property

Private Property Get Expression_ErrInfo() As String
    Dim s1 As String: s1 = m_Ins.ErrInfo: s1 = s1 & IIf(Len(s1), vbCrLf, "")
    Expression_ErrInfo = s1 & IIf(Len(m_ErrInfo), TypeName(Me) & " ", "") & m_ErrInfo
End Property

Private Function Expression_Eval() 'As Double
Try: On Error GoTo Catch
    If Not Expression_CanEval Then Exit Function
    Expression_Eval = VBA.Math.Abs(m_Ins.Eval)
    Exit Function
Catch:
    m_ErrInfo = Err.number & " " & Err.Description
End Function


Private Property Set Expression_Expr1(ByVal RHS As Expression)
    Set m_Ins = RHS
End Property
Private Property Get Expression_Expr1() As Expression
    Set Expression_Expr1 = m_Ins
End Property


Private Property Set Expression_Expr2(ByVal RHS As Expression)
    Set m_Ins = RHS
End Property
Private Property Get Expression_Expr2() As Expression
    Set Expression_Expr2 = m_Ins
End Property


Private Property Get Expression_Operator1() As String
    Expression_Operator1 = myOp1
End Property
Private Property Get Expression_Operator2() As String
    Expression_Operator2 = myOp2
End Property


Private Property Get Expression_Rank() As OperatorRank
    Expression_Rank = myRank
End Property

Private Function Expression_ToStr(Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    Dim b As Boolean: b = myRank < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
    Dim s As String
    If Not m_Ins Is Nothing Then s = myOp1 & IIf(b, "(", "") & m_Ins.ToStr(myRank) & IIf(b, ")", "") & myOp2
    Expression_ToStr = s
End Function

Private Sub Expression_ToFmt(ByVal Fmt As FormatExpr, ByVal ParentRank As OperatorRank)
    Fmt.ExprOpAbsToStr Me, ParentRank
End Sub



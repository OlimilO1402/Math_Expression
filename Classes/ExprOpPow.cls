VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprOpPow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Expression
'Implements OperatorBinary
Private m_LHS As Expression
Private m_RHS As Expression
Private Const myRank As Long = OperatorRank.Rank6ExprOpPow
Private Const myOp   As String = " ^ "

Private m_ErrInfo As String


Friend Sub New_(aLHS As Expression)
    Set m_LHS = aLHS
End Sub

' v ############################## v '   Implements Expression    ' v ############################## v '
Private Property Get Expression_CanEval() As Boolean
    If m_LHS Is Nothing Then Exit Function
    If m_RHS Is Nothing Then Exit Function
    Expression_CanEval = m_LHS.CanEval And m_RHS.CanEval ' True
End Property

Private Property Get Expression_ErrInfo() As String
    Dim s1 As String: s1 = m_LHS.ErrInfo: s1 = s1 & IIf(Len(s1), vbCrLf, "")
    Dim s2 As String: s2 = m_RHS.ErrInfo: s2 = s2 & IIf(Len(s2), vbCrLf, "")
    Expression_ErrInfo = s1 & s2 & IIf(Len(m_ErrInfo), TypeName(Me) & " ", "") & m_ErrInfo
End Property

Private Function Expression_Eval() 'As Double
Try: On Error GoTo Catch
    If Not Expression_CanEval Then Exit Function
    Expression_Eval = m_LHS.Eval ^ m_RHS.Eval
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
    Expression_Operator2 = myOp
End Property


Private Property Get Expression_Rank() As OperatorRank
    Expression_Rank = myRank
End Property

Private Function Expression_ToStr(Optional ByVal ParentRank As OperatorRank = OperatorRank.Rank0None) As String
    Dim b As Boolean: b = myRank < ParentRank And ParentRank <> OperatorRank.Rank9ExprOpBrac
    Dim s As String:  s = myOp
    If Not m_LHS Is Nothing Then s = IIf(b, "(", "") & m_LHS.ToStr(myRank) & s
    If Not m_RHS Is Nothing Then s = s & m_RHS.ToStr(myRank + 1) & IIf(b, ")", "")
    Expression_ToStr = s
End Function

Private Sub Expression_ToFmt(ByVal Fmt As FormatExpr, ByVal ParentRank As OperatorRank)
    Fmt.ExprOpPowToStr Me, ParentRank
End Sub

    'hier werden 3 verschiedene Klammern gebraucht
    '((basis1 + basis2) ^ (expo1 + expo2))
    '1. die Klammer ganz au�en
    '2. die Klammer der Basis-Expression
    '3. die Klammer der Exponent-Expression
    'jetzt ist die Frage wer macht die Klammern und wo
    'einerseits k�nnen wir Expression-Objekte direkt im Code erzeugen bereits mit LHS und RHS auf einmal,
    'd.h. explizite Klammernexpressions sind nicht erforderlich, da die Hierarchie sowieso direkt abgebildet ist.
    'andererseits m�ssen wir die Hierarchie aus einem Zeichenstrom herauslesen
    
    'wozu brauchen wir das property Rank direkt im Expression-Interface
    'falls bereits ein Klammern-Objekt besteht braucht man nicht nochmal Klammer machen
    'Klammer muss immer gemacht werden bei z.B. Unter der Wurzel
'    Dim b As Boolean: b = myRank <= ParentRank 'And ParentRank <> OperatorRank.Rank9ExprOpBrac
'    Dim s As String:  s = myOp
'    If Not m_LHS Is Nothing Then s = IIf(b, "(", "") & m_LHS.ToStr(myRank) & s
'    If Not m_RHS Is Nothing Then s = s & m_RHS.ToStr(myRank) & IIf(b, ")", "") '& " = " & Trim(Str(Expression_Eval))
'    Expression_ToStr = s

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


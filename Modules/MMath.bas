Attribute VB_Name = "MMath"
Option Explicit

Private m_Factorials() 'As Variant 'As Decimal

Public Sub Init()
    
    InitFactorials
    
End Sub

Private Sub InitFactorials()
    ReDim m_Factorials(0 To 171)
    Dim i As Long, f: f = CDec(1)
    m_Factorials(0) = CDec(0)
    m_Factorials(1) = f
    For i = 2 To 27
        f = f * CDec(i)
        m_Factorials(i) = f
    Next
    f = CDbl(f)
    For i = 28 To 170
        f = f * CDbl(i)
        m_Factorials(i) = f
    Next
    m_Factorials(171) = GetINFE
End Sub

Private Function GetINFE(Optional ByVal sign As Long = 1) As Double
    On Error Resume Next
    GetINFE = Sgn(sign) / 0
    On Error GoTo 0
End Function

Public Function Fact(ByVal N As Long) As Variant 'As Decimal
    If N > 170 Then N = 171
    Fact = m_Factorials(N)
End Function

'    Dim num As Double, e As Double: e = Exp(1) ' e = 2,71828182845905
'    Debug.Print e
'    Dim s As String
'    num = e * e: s = s & "LN(" & num & ")     = " & MMath.LN(num) & vbCrLf      ' = 2
'    num = 1000:  s = s & "Log10(" & num & ")  = " & MMath.Log10(num) & vbCrLf   ' = 3
'    num = 10000: s = s & "Log10(" & num & ")  = " & MMath.LogN(num) & vbCrLf    ' = 4
'    num = 32:    s = s & "Log(" & num & ", 2) = " & MMath.LogN(num, 2) & vbCrLf ' = 5
'    MsgBox s
'number          |  base        | xl-function     | result | description
'    7.389056099 |  2,718281828 | LN(Zahl)        =   2    | LN aka ln  := Logarithm to base  e
' 1000           | 10           | Log10(Zahl)     =   3    | Log10      := Logarithm to base 10, with the excelfunction LOG10
'10000           | 10           | Log(Zahl)       =   4    | Log aka lg := Logarithm to base 10, with the excelfunction Log, base not explicitely given
'   32           |  2           | Log(Zahl;Basis) =   5    | Log        := Logarithm to base  2, if the base 2 was explicitely given

'Logarithmus naturalis, logarithm to base e
Public Function LN(ByVal d As Double) As Double
    LN = VBA.Math.Log(d)
End Function

'Logarithm to the base 10
Public Function Log10(ByVal d As Double) As Double
    If d = 0 Then Exit Function
    Log10 = VBA.Math.Log(d) / VBA.Math.Log(10)
End Function

'Logarithm to a given base
Public Function LogN(ByVal x As Double, _
                     Optional ByVal base As Double = 10#) As Double
                     'base must not be 1 or 0
    If base <= 1 Then Exit Function
    LogN = VBA.Math.Log(x) / VBA.Math.Log(base)
End Function

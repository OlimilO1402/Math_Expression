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

Public Function Log10(ByVal d As Double) As Double
    Log10 = VBA.Math.Log(d) / VBA.Math.Log(10)
End Function

Public Function LN(ByVal d As Double) As Double
  LN = VBA.Math.Log(d)
End Function

Public Function LogN(ByVal x As Double, _
                     Optional ByVal N As Double = 10#) As Double
                     'n darf nicht eins und nicht 0 sein
    LogN = VBA.Math.Log(x) / VBA.Math.Log(N)
End Function



Attribute VB_Name = "MParseNumber"
Option Explicit
Private Const VZ    As String = "+-"          'Sign, Vorzeichen
Private Const digit As String = "0123456789"  'Ziffern
Private Const DS    As String = ".,"          'Decimalseparator
Private Const ScE   As String = "E" 'and not a lower "e"
Private Const WS    As String = " "

Private Enum EStateNumber
    state0_Start = 0
    state1_VZ = 1    'nach Vorzeichen können auch beliebige WS folgen
    state2_dig1 = 2  'digits
    state3_DS = 3    'Decimal-Separator
    state4_dig2 = 4  'digits
    state5_ScE = 5   'scientific E
    state6_EVZ = 6   'nach Vorzeichen können auch beliebige WS folgen
    state7_dig3 = 7  'digits
    state8_End = 8
End Enum

'
'Funktion parst die erste Zahl raus die im String gefunden wird und gibt die Zahl und den Bereich davor und danach als String zurück
'teststrings:
'"0" -> "", "0", ""
'"1" -> "", "1", ""
'"2" -> "", "2", ""
'"." -> ".", "", ""
'"," -> ",", "", ""
'" " -> " ", "", ""
'"a" -> "a", "", ""
'" 0" -> " ", "0", ""
'" 1" -> " ", "1", ""
'" 2" -> " ", "2", ""
'".0" .> "", "0.0", ""
'was is besser:
'  "a0" -> "a0", "", ""
'oder
'  "a0" -> "a", "0", ""
'das kommt drauf an was gewünscht ist
'am besten man kann beides und kann dann jeweils entscheiden
'

Public Sub Number_Parse(ByVal s_in As String, pre_out As String, num_out As String, post_out As String)
    'Bsp:
    '"ab -12.34E-5 cd"
    ' +   12    .   34    E    -    5
    'VZ, numI, DS, numF, ScE, EVZ, numE
    '     3                4               5                 7               8                10               11               12                13
    Dim len_pre As Long, pos_VZ As Long, pos_numI As Long, pos_DS As Long, pos_numF As Long, pos_ScE As Long, pos_EVZ As Long, pos_numE As Long, pos_post As Long
    Dim s_pre As String, s_VZ As String, s_numI As String, s_DS As String, s_numF As String, s_ScE As String, s_EVZ As String, s_numE As String, s_post As String
    len_pre = -1
    Dim i As Long
    Dim state As EStateNumber, oldstate As EStateNumber
    Dim ls As Long: ls = Len(s_in)
    If ls = 0 Then Exit Sub
    Dim c As String
    Do While i <= ls
        i = i + 1
        c = Mid$(s_in, i, 1)
        If Len(c) = 0 Then Exit Do 'EOF!
        oldstate = state
        Select Case state
        Case EStateNumber.state0_Start: state = State0Start(c)
        Case EStateNumber.state1_VZ:    state = State1VZ(c)
        Case EStateNumber.state2_dig1:  state = State2Dig1(c)
        Case EStateNumber.state3_DS:    state = State3DS(c)
        Case EStateNumber.state4_dig2:  state = State4Dig2(c)
        Case EStateNumber.state5_ScE:   state = State5ScE(c)
        Case EStateNumber.state6_EVZ:   state = State6EVZ(c)
        Case EStateNumber.state7_dig3:  state = State7Dig3(c)
        'Case EStateNumber.state8_End:
        End Select
        If oldstate <> state Then
            Select Case state
            Case EStateNumber.state0_Start
            Case EStateNumber.state1_VZ:   pos_VZ = i:   If len_pre < 0 Then len_pre = i - 1
            Case EStateNumber.state2_dig1: pos_numI = i: If len_pre < 0 Then len_pre = i - 1
            Case EStateNumber.state3_DS:   pos_DS = i:   If len_pre < 0 Then len_pre = i - 1
            Case EStateNumber.state4_dig2: pos_numF = i
            Case EStateNumber.state5_ScE:  pos_ScE = i
            Case EStateNumber.state6_EVZ:  pos_EVZ = i
            Case EStateNumber.state7_dig3: pos_numE = i
            Case EStateNumber.state8_End:  pos_post = i
            End Select
        End If
    Loop
    If pos_VZ > 0 Then
        If pos_numI > 0 Then
            s_VZ = Mid$(s_in, pos_VZ, IIf(pos_numI > 0, pos_numI, pos_DS) - pos_VZ)
        Else
            len_pre = pos_VZ
        End If
    End If

    If pos_numI > 0 Then s_numI = Mid$(s_in, pos_numI, IIf(pos_DS > 0, pos_DS, IIf(pos_ScE > 0, pos_ScE, IIf(pos_post > 0, pos_post, ls + 1))) - pos_numI)
    If pos_DS > 0 Then
        If pos_numF > 0 Then
            s_DS = Mid$(s_in, pos_DS, IIf(pos_numF > 0, pos_numF, IIf(pos_ScE > 0, pos_ScE, IIf(pos_post > 0, pos_post, ls + 1))) - pos_DS)
        End If
        If pos_numI = 0 And pos_numF = 0 Then
            len_pre = pos_DS
        End If
    End If
    If pos_VZ = 0 And pos_numI = 0 And pos_DS = 0 And pos_numF = 0 And pos_ScE = 0 And pos_EVZ = 0 And pos_numE = 0 And pos_post = 0 Then
        len_pre = ls
    End If
    If len_pre > 0 Then s_pre = Mid$(s_in, 1, len_pre)

    If pos_numF > 0 Then s_numF = Mid$(s_in, pos_numF, IIf(pos_ScE > 0, pos_ScE, IIf(pos_post > 0, pos_post, ls + 1)) - pos_numF)
    If pos_ScE > 0 Then
        If pos_numE > 0 Then
            s_ScE = Mid$(s_in, pos_ScE, IIf(pos_EVZ > 0, pos_EVZ, IIf(pos_post > 0, pos_post, ls + 1)) - pos_ScE)
            If pos_EVZ > 0 Then s_EVZ = Mid$(s_in, pos_EVZ, IIf(pos_numE > 0, pos_numE, IIf(pos_post > 0, pos_post, ls + 1)) - pos_EVZ)
            If pos_numE > 0 Then s_numE = Mid$(s_in, pos_numE, IIf(pos_post > 0, pos_post, ls + 1) - pos_numE)
        Else
            pos_post = pos_ScE
        End If
    End If
    If pos_post > 0 Then s_post = Mid$(s_in, pos_post, ls - pos_post)
    pre_out = s_pre
    num_out = s_VZ & s_numI & s_DS & s_numF & s_ScE & s_EVZ & s_numE
    post_out = s_post
End Sub

Private Function State0Start(ByVal c As String) As EStateNumber
    Dim nextState As EStateNumber
    If InStr(1, VZ, c) Then
        nextState = EStateNumber.state1_VZ
    ElseIf InStr(1, digit, c) Then
        nextState = EStateNumber.state2_dig1
    ElseIf InStr(1, DS, c) Then
        nextState = EStateNumber.state3_DS
    Else
        nextState = EStateNumber.state0_Start
    End If
    State0Start = nextState
End Function

Private Function State1VZ(ByVal c As String) As EStateNumber
    Dim nextState As EStateNumber
    If InStr(1, WS, c) Then
        nextState = EStateNumber.state1_VZ
    ElseIf InStr(1, digit, c) Then
        nextState = EStateNumber.state2_dig1
    ElseIf InStr(1, DS, c) Then
        nextState = EStateNumber.state3_DS
    Else
        nextState = EStateNumber.state0_Start
    End If
    State1VZ = nextState
End Function

Private Function State2Dig1(ByVal c As String) As EStateNumber
    Dim nextState As EStateNumber
    If InStr(1, digit, c) Then
        nextState = EStateNumber.state2_dig1
    ElseIf InStr(1, DS, c) Then
        nextState = EStateNumber.state3_DS
    Else
        nextState = EStateNumber.state8_End
    End If
    State2Dig1 = nextState
End Function

Private Function State3DS(ByVal c As String) As EStateNumber
    Dim nextState As EStateNumber
    If InStr(1, digit, c) Then
        nextState = EStateNumber.state4_dig2
    Else
        nextState = EStateNumber.state8_End
    End If
    State3DS = nextState
End Function

Private Function State4Dig2(ByVal c As String) As EStateNumber
    Dim nextState As EStateNumber
    If InStr(1, digit, c) Then
        nextState = EStateNumber.state4_dig2
    ElseIf InStr(1, ScE, c) Then
        nextState = EStateNumber.state5_ScE
    Else
        nextState = EStateNumber.state8_End
    End If
    State4Dig2 = nextState
End Function

Private Function State5ScE(ByVal c As String) As EStateNumber
    Dim nextState As EStateNumber
    If InStr(1, VZ, c) Then
        nextState = EStateNumber.state6_EVZ
    ElseIf InStr(1, digit, c) Then
        nextState = EStateNumber.state7_dig3
    Else
        nextState = EStateNumber.state8_End
    End If
    State5ScE = nextState
End Function

Private Function State6EVZ(ByVal c As String) As EStateNumber
    Dim nextState As EStateNumber
    If InStr(1, digit, c) Then
        nextState = EStateNumber.state7_dig3
    Else
        nextState = EStateNumber.state8_End
    End If
    State6EVZ = nextState
End Function

Private Function State7Dig3(ByVal c As String) As EStateNumber
    Dim nextState As EStateNumber
    If InStr(1, digit, c) Then
        nextState = EStateNumber.state7_dig3
    Else
        nextState = EStateNumber.state8_End
    End If
    State7Dig3 = nextState
End Function


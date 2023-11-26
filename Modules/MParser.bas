Attribute VB_Name = "MParser"
Option Explicit
'The Microsoft Visual Basic Language Specification Version 11.0
'Paul Vick, Lucian Wischik, Microsoft Corporation:
'https://www.google.com/url?sa=t&rct=j&q=&esrc=s&source=web&cd=&cad=rja&uact=8&ved=2ahUKEwixnY-QjMuCAxUzQvEDHTckCdkQFnoECBAQAQ&url=https%3A%2F%2Fdownload.microsoft.com%2Fdownload%2F2%2F2%2FB%2F22B4695E-CEBC-4296-8DC3-0F329CA6751D%2FVisual%2520Basic%2520Language%2520Specification.docx&usg=AOvVaw0lyhs1AutBwVfyCf4eM-Ua&opi=89978449

'https://learn.microsoft.com/en-us/dotnet/visual-basic/reference/language-specification/introduction
'On the very down left hand side click "Download PDF"

'or use this online-manual:
'https://ljw1004.github.io/vbspec/vb.html
'
'Lexical Grammar

'Start               ::=  [  LogicalLine+  ]
'LogicalLine         ::=  [  LogicalLineElement+  ]  [  Comment  ]  LineTerminator
'
'LogicalLineElement  ::=  WhiteSpace  |
'                   LineContinuation  |
'                              Token

'Token
'    : Identifier
'    | Keyword
'    | Literal
'    | Separator
'    | Operator
'    ;
'Identifier
'    : NonEscapedIdentifier TypeCharacter?
'    | Keyword TypeCharacter
'    | EscapedIdentifier
'    ;
'Keyword
'    : 'AddHandler' | 'AddressOf'  | 'Alias'       | 'And'           | 'AndAlso'         | 'As'             | 'Boolean'   | 'ByRef'
'    | 'Byte'       | 'ByVal'      | 'Call'        | 'Case'          | 'Catch'           | 'CBool'          | 'CByte'     | 'CChar'
'    | 'CDate'      | 'CDbl'       | 'CDec'        | 'Char'          | 'CInt'            | 'Class'          | 'CLng'      | 'CObj'
'    | 'Const'      | 'Continue'   | 'CSByte'      | 'CShort'        | 'CSng'            | 'CStr'           | 'CType'     | 'CUInt'
'    | 'CULng'      | 'CUShort'    | 'Date'        | 'Decimal'       | 'Declare'         | 'Default'        | 'Delegate'  | 'Dim'
'    | 'DirectCast' | 'Do'         | 'Double'      | 'Each'          | 'Else'            | 'ElseIf'         | 'End'       | 'EndIf'
'    | 'Enum'       | 'Erase'      | 'Error'       | 'Event'         | 'Exit'            | 'False'          | 'Finally'   | 'For'
'    | 'Friend'     | 'Function'   | 'Get'         | 'GetType'       | 'GetXmlNamespace' | 'Global'         | 'GoSub'     | 'GoTo'
'    | 'Handles'    | 'If'         | 'Implements'  | 'Imports'       | 'In'              | 'Inherits'       | 'Integer'   | 'Interface'
'    | 'Is'         | 'IsNot'      | 'Let'         | 'Lib'           | 'Like'            | 'Long'           | 'Loop'      | 'Me'
'    | 'Mod'        | 'Module'     | 'MustInherit' | 'MustOverride'  | 'MyBase'          | 'MyClass'        | 'Namespace' | 'Narrowing'
'    | 'New'        | 'Next'       | 'Not'         | 'Nothing'       | 'NotInheritable'  | 'NotOverridable' | 'Object'    | 'Of'
'    | 'On'         | 'Operator'   | 'Option'      | 'Optional'      | 'Or'              | 'OrElse'         | 'Overloads' | 'Overridable'
'    | 'Overrides'  | 'ParamArray' | 'Partial'     | 'Private'       | 'Property'        | 'Protected'      | 'Public'    | 'RaiseEvent'
'    | 'ReadOnly'   | 'ReDim'      | 'REM'         | 'RemoveHandler' | 'Resume'          | 'Return'         | 'SByte'     | 'Select'
'    | 'Set'        | 'Shadows'    | 'Shared'      | 'Short'         | 'Single'          | 'Static'         | 'Step'      | 'Stop'
'    | 'String'     | 'Structure'  | 'Sub'         | 'SyncLock'      | 'Then'            | 'Throw'          | 'To'        | 'True'
'    | 'Try'        | 'TryCast'    | 'TypeOf'      | 'UInteger'      | 'ULong'           | 'UShort'         | 'Using'     | 'Variant'
'    | 'Wend'       | 'When'       | 'While'       | 'Widening'      | 'With'            | 'WithEvents'     | 'WriteOnly' | 'Xor'
'    ;
'Literal
'    : BooleanLiteral
'    | IntegerLiteral
'    | FloatingPointLiteral
'    | StringLiteral
'    | CharacterLiteral
'    | DateLiteral
'    | Nothing
'    ;
'Separator
'    :     '(' | ')' | '{' | '}' | '!' | '#' | ',' | '.' | ':' | '?'
'    ;
'Operator
'    :     '&' | '*' | '+' | '-' | '/' | '\\' | '^' | '<' | '=' | '>'
'    ;

'Expression
'    : SimpleExpression
'    | TypeExpression
'    | MemberAccessExpression
'    | DictionaryAccessExpression
'    | InvocationExpression
'    | IndexExpression
'    | NewExpression
'    | CastExpression
'    | OperatorExpression
'    | ConditionalExpression
'    | LambdaExpression
'    | QueryExpression
'    | XMLLiteralExpression
'    | XMLMemberAccessExpression
'    ;
Public Enum ETokenTyp
    tktIdentifier
        tktNonEscapedIdentifier
        tktEscapedIdentifier
        'tktIdentifierName       '::=  IdentifierStart  [  IdentifierCharacter+  ]
'            tktIdentifierStart      'AlphaCharacter  |  UnderscoreCharacter IdentifierCharacter
'            tktIdentifierCharacter  '::=
'            tktUnderscoreCharacter
        
'        tktAlphaCharacter
'        tktNumericCharacter
'        tktCombiningCharacter
'        tktFormattingCharacter
'        tktAlphaCharacter  '::=    < Unicode alphabetic character (classes Lu, Ll, Lt, Lm, Lo, Nl) > NumericCharacter  ::=  < Unicode decimal digit character (class Nd) >
'        tktCombiningCharacter  '::=  < Unicode combining character (classes Mn, Mc) >
    
        tktTypeCharacter  '::=
            tktIntegerTypeCharacter  '|'    IntegerTypeCharacter  '::=  %
            tktLongTypeCharacter     '|'    LongTypeCharacter     '::=  &
            tktDecimalTypeCharacter  '|'    DecimalTypeCharacter  '::=  @
            tktSingleTypeCharacter   '|'    SingleTypeCharacter   '::=  !
            tktDoubleTypeCharacter   '|'    DoubleTypeCharacter   '::=  #
            tktStringTypeCharacter   '      StringTypeCharacter   '::=  $
'
    tktKeyword
    
    tktLiteral
        tktBooleanLiteral
        tktIntegerLiteral
        tktFloatingPointLiteral
        tktStringLiteral
        tktCharacterLiteral
        tktDateLiteral
        tktNothing
        
    tktSeparator
    
    tktOperator
End Enum

Public Type TToken
    TokenTyp As ETokenTyp
    Value    As String
End Type

Public Function IsIdentifier(tk As TToken) As Boolean
    Select Case tk.TokenTyp
    Case ETokenTyp.tktNonEscapedIdentifier Or _
         ETokenTyp.tktEscapedIdentifier 'Or _
         'ETokenTyp.tktKeyword Or _
         'ETokenTyp.tktIdentifierName Or _
         'ETokenTyp.tktIdentifierStart Or _
         'ETokenTyp.tktIdentifierCharacter Or _
         'ETokenTyp.tktUnderscoreCharacter Or _
         'ETokenTyp.tktAlphaCharacter Or _
         'ETokenTyp.tktNumericCharacter Or _
         'ETokenTyp.tktCombiningCharacter Or _
         'ETokenTyp.tktFormattingCharacter Or _
         'ETokenTyp.tktAlphaCharacter Or _
         'ETokenTyp.tktCombiningCharacter
        IsIdentifier = True
    End Select
    Dim d As Date
    d = #8/23/1970 3:45:39 AM#
    MsgBox d
End Function

Public Function IsTypeCharacter(tk As TToken) As Boolean
    Select Case tk.TokenTyp
    Case ETokenTyp.tktIntegerTypeCharacter Or _
         ETokenTyp.tktLongTypeCharacter Or _
         ETokenTyp.tktDecimalTypeCharacter Or _
         ETokenTyp.tktSingleTypeCharacter Or _
         ETokenTyp.tktDoubleTypeCharacter Or _
         ETokenTyp.tktStringTypeCharacter
        IsTypeCharacter = True
    End Select
End Function

Public Function IsLiteral(tk As TToken) As Boolean
    Select Case tk.TokenTyp
    Case ETokenTyp.tktBooleanLiteral Or _
         ETokenTyp.tktIntegerLiteral Or _
         ETokenTyp.tktFloatingPointLiteral Or _
         ETokenTyp.tktStringLiteral Or _
         ETokenTyp.tktCharacterLiteral Or _
         ETokenTyp.tktDateLiteral Or _
         ETokenTyp.tktNothing
        IsLiteral = True
    End Select
End Function

Public Function IsSeparator(tk As TToken) As Boolean
    Select Case tk.TokenTyp
    Case Else
        IsSeparator = True
    End Select
End Function

Public Function IsOperator(tk As TToken) As Boolean
    Select Case tk.TokenTyp
    Case Else
        IsOperator = True
    End Select
End Function

''Parse_Number oder
''Number_Parse
''AVB-Beitrag am 09.12.2019 21:12
''thema: non deterministic finite automat, für Zahlen
'
'Private Sub Form_Load()
'
'    Dim s As String, pre As String, num As String, post As String
'
'    '   "0" -> "", "0", ""
'    s = "0":     Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "1" -> "", "1", ""
'    s = "1":     Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "2" -> "", "2", ""
'    s = "2":     Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "." -> ".", "", ""
'    s = ".":     Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "," -> ",", "", ""
'    s = ",":     Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   " " -> " ", "", ""
'    s = " ":     Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "a" -> "a", "", ""
'    s = "a":     Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   " 0" -> " ", "0", ""
'    s = " 0":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   " 1" -> " ", "1", ""
'    s = " 1":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   " 2" -> " ", "2", ""
'    s = " 2":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   ".0" .> "", "0.0", ""
'    s = ".0":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   ".1" .> "", ".1", ""
'    s = ".1":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   ".2" .> "", ".2", ""
'    s = ".2":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "a0" -> "a0", "", ""
'    s = "a0":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "a1" -> "a1", "", ""
'    s = "a1":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "a2" -> "a2", "", ""
'    s = "a2":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "+0" -> "", "+0", ""
'    s = "+0":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "+1" -> "", "+1", ""
'    s = "+1":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "+2" -> "", "+2", ""
'    s = "+2":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "-0" -> "", "-0", ""
'    s = "-0":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "-1" -> "", "-1", ""
'    s = "-1":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "-2" -> "", "-2", ""
'    s = "-2":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "abc -12,34E-5 km defg" -> "abc ", "-12,34E-5", " km defg"
'    s = "abc -12,34E-5 km defg":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'    '   "abc -12,34E-0 km defg" -> "abc ", "-12,34E-5", " km defg"
'    s = "abc -12,34E-as  km defg":    Number_Parse s, pre, num, post: DebugPrint s, pre, num, post
'
'End Sub
'
'Sub DebugPrint(s As String, pre As String, num As String, post As String)
'    Debug.Print """" & s & """" & " -> " & """" & pre & """, """ & num & """, """ & post & """"
'End Sub
'

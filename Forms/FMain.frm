VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "MathCalcExpression"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13005
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   13005
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnGetResults 
      Caption         =   "GetResults"
      Height          =   375
      Left            =   7320
      TabIndex        =   46
      Top             =   600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Output"
      Height          =   1335
      Left            =   8760
      TabIndex        =   40
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox ChkToClipB 
         Caption         =   "To Clipboard"
         Height          =   255
         Left            =   2400
         TabIndex        =   48
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptToStr 
         Caption         =   "ToStr"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CheckBox ChkXLCompat 
         Caption         =   "xL compat."
         Height          =   255
         Left            =   2760
         TabIndex        =   45
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox ChkSepAsNewL 
         Caption         =   "Separator As NewLine"
         Height          =   255
         Left            =   1440
         TabIndex        =   44
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox ChkCondensed 
         Caption         =   "Condensed"
         Height          =   255
         Left            =   1440
         TabIndex        =   43
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton OptRPN 
         Caption         =   "ReversePN"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton OptAlgebra 
         Caption         =   "Algebraic"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   7200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   39
      Text            =   "FMain.frx":08CA
      Top             =   1560
      Width           =   5775
   End
   Begin VB.CommandButton BtnTests 
      Caption         =   "Tests"
      Height          =   375
      Left            =   7320
      TabIndex        =   38
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnExprEvaluate 
      BackColor       =   &H000000FF&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   36
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton BtnPoint 
      BackColor       =   &H8000000D&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   35
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton BtnCipher 
      BackColor       =   &H8000000D&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1440
      TabIndex        =   34
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton BtnSign 
      BackColor       =   &H8000000D&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   33
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton BtnCipher 
      BackColor       =   &H8000000D&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   2160
      TabIndex        =   30
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton BtnCipher 
      BackColor       =   &H8000000D&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   1440
      TabIndex        =   29
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton BtnCipher 
      BackColor       =   &H8000000D&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   720
      TabIndex        =   28
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton BtnCipher 
      BackColor       =   &H8000000D&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   2160
      TabIndex        =   25
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton BtnCipher 
      BackColor       =   &H8000000D&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   1440
      TabIndex        =   24
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton BtnCipher 
      BackColor       =   &H8000000D&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   720
      TabIndex        =   23
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton BtnCipher 
      BackColor       =   &H8000000D&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   2160
      TabIndex        =   20
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton BtnCipher 
      BackColor       =   &H8000000D&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   1440
      TabIndex        =   19
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton BtnCipher 
      BackColor       =   &H8000000D&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   720
      TabIndex        =   18
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton BtnExprLN 
      Caption         =   "ln"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   32
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton BtnExprAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   31
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton BtnExprLog10 
      Caption         =   "log"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   27
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton BtnExprSubt 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   26
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton BtnExprPow10 
      Caption         =   "10^x"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   22
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton BtnExprMul 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   21
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton BtnExprPow 
      Caption         =   "x^y"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   17
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton BtnExprDiv 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   16
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton BtnExprFact 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   15
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton BtnExprBracClose 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   14
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton BtnExprBracOpen 
      Caption         =   "("
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   13
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton BtnExprSqrt 
      Caption         =   "-v�"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton BtnExprMod 
      Caption         =   "mod"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   11
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton BtnExprExp 
      Caption         =   "exp"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton BtnExprAbs 
      Caption         =   "|x|"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton BtnExpr1DivX 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton BtnExprSqr 
      Caption         =   "x�"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton BtnBack 
      Caption         =   "<-"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton BtnDelete 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton BtnExprConstE 
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton BtnExprConstPi 
      Caption         =   "pi"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton BtnExprCub 
      Caption         =   "x�"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6150
      Left            =   3600
      TabIndex        =   37
      Top             =   0
      Width           =   3615
   End
   Begin VB.TextBox TxtInput 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Text            =   "123"
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "        "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3600
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Expressions As Collection ' Of Expression
Private m_Num         As String

Private m_ExprStack  As ExprStack 'Deque
'Private m_Expression  As ExprDeque
Private m_Expression  As Expression

Private m_LastBinOp   As Expression 'OperatorBinary

Private m_TBHasResult As Boolean

'OK Windows Taschenrechner Calc.exe funktioniert so:
'1. es gibt eine gro�e  rechtsb�ndige Anzeige 1 die die eingegebenen Zahlen anzeigt z.B.: 12
'2. es gibt eine kleine rechtsb�ndige Anzeige 2 dar�ber, die nach der Eingabe des Operators die eingegebene Zahl und den Operator anzeigt z.B.: 12 +
'   die Anzeige 2 zeigt auch Klammern an und der Klammern-Schalter bekommt eine fortlaufende Nummer die sich bei schlie�en der Klammer reduziert
'   die Anzeige 2 zeigt den kompletten Term an
'   nach klicken des Schalters "=" wird in Anzeige 2 angezeigt "1 + 1 =" und in Anzeige 1 wird das Ergebnis ausgegeben
'3. es gibt eine Liste rechts daneben die nach dem "=" den Term mit Ergebnis abspeichert
'gib ein

Private Sub BtnTests_Click()
    Dim s As String
    If OptToStr.Value Then
        s = MTests.Test3
    Else
        s = MTests.Test3(GetFormat)
    End If
    Text1.Text = s
    If ChkToClipB.Value = vbChecked Then
        Clipboard.Clear
        Clipboard.SetText s
    End If
End Sub

Private Sub BtnGetResults_Click()
    Text1.Text = MTests.GetResults
End Sub

Private Function GetFormat() As FormatExpr
    Dim b1 As Boolean: b1 = ChkCondensed.Value = vbChecked
    Dim b2 As Boolean: b2 = ChkXLCompat.Value = vbChecked
    Dim b3 As Boolean: b3 = ChkSepAsNewL.Value = vbChecked
    If OptAlgebra.Value Then Set GetFormat = MNew.FormatAlg(b1, b2) Else Set GetFormat = MNew.FormatRPN(b3)
End Function

Private Sub Command1_Click()
    'MsgBox MMath.Fact(25)
    'Exit Sub
'    Dim tef As TestExFunc: Set tef = New TestExFunc
'    MsgBox tef.TestIt(1, 2, 3)
'    Dim fnc As ExprFunction
'    Set fnc = MNew.ExprFunction("TestIt", tef)
'    fnc.params.Add MNew.ExprLitNum(1)
'    fnc.params.Add MNew.ExprLitNum(2)
'    fnc.params.Add MNew.ExprLitNum(3)
'    Dim ex As Expression: Set ex = fnc
'    MsgBox ex.Eval
    
'    Dim ex As Expression
'    Set ex = MNew.ExprOpIIf(MNew.ExprLitBol(True), MNew.ExprLitBol(False), MNew.ExprLitBol(True))
'    MsgBox ex.Eval
'    Set ex = MNew.ExprOpIIf(MNew.ExprLitBol(True), MNew.ExprLitNum(True), MNew.ExprLitBol(False))
'    MsgBox ex.Eval
'    Set ex = MNew.ExprOpIIf(MNew.ExprLitBol(False), MNew.ExprLitBol(False), MNew.ExprLitBol(True))
'    MsgBox ex.Eval
'    Set ex = MNew.ExprOpIIf(MNew.ExprLitBol(False), MNew.ExprLitBol(True), MNew.ExprLitBol(False))
'    MsgBox ex.Eval
'    Set ex = MNew.ExprOpIIf(MNew.ExprLitBol(True), MNew.ExprLitNum(123), MNew.ExprLitNum(456))
'    MsgBox ex.Eval
'    Set ex = MNew.ExprOpIIf(MNew.ExprLitBol(False), MNew.ExprLitNum(123), MNew.ExprLitNum(456))
'    MsgBox ex.Eval

    'Dim ex As Expression: Set ex = MNew.ExprOpLogN(MNew.ExprLitNum(-5), MNew.ExprOpDiv(MNew.ExprLitNum(3), MNew.ExprLitNum(1)))
    'MsgBox ex.Eval
    'MsgBox ex.ErrInfo
    
    Dim ex As Expression
    
'    Set ex = MNew.ExprOpIIf(MNew.ExprLitBol(True), MNew.ExprLitBol(True), MNew.ExprLitBol(False))
'    MsgBox ex.Eval
'    Set ex = MNew.ExprOpIIf(MNew.ExprLitBol(False), MNew.ExprLitBol(True), MNew.ExprLitBol(False))
'    MsgBox ex.Eval
'    Set ex = MNew.ExprOpIIf(MNew.ExprLitBol(True), MNew.ExprLitBol(False), MNew.ExprLitBol(True))
'    MsgBox ex.Eval
'    Set ex = MNew.ExprOpIIf(MNew.ExprLitBol(False), MNew.ExprLitBol(False), MNew.ExprLitBol(True))
'    MsgBox ex.Eval
    
    Set ex = MNew.ExprOpIIf(MNew.ExprOpBolEqual(MNew.ExprLitNum(123), MNew.ExprLitNum(234)), MNew.ExprLitNum(111), MNew.ExprLitNum(222))
    MsgBox ex.ToStr & " = " & ex.Eval
    Set ex = MNew.ExprLitNum(123)
    'Set ex = MNew.ExprOpAdd(ex, MNew.ExprLitNum(456))
    'Set ex = MNew.ExprOpSubt(ex, MNew.ExprLitNum(456))
    'Set ex = MNew.ExprOpMul(ex, MNew.ExprLitNum(456))
    'Set ex = MNew.ExprOpDiv(ex, MNew.ExprLitNum(456))
    Set ex = MNew.ExprOpDiv(MNew.ExprOpAdd(ex, MNew.ExprLitNum(456)), MNew.ExprOpSubt(MNew.ExprOpMul(ex, MNew.ExprLitNum(789)), MNew.ExprLitNum(456)))
    Set ex = MNew.ExprOpIIf(MNew.ExprOpBolEqual(MNew.ExprLitNum(123), MNew.ExprLitNum(123)), MNew.ExprLitNum(111), MNew.ExprLitNum(222))
    MsgBox ex.ToStr & " = " & ex.Eval
End Sub

Private Sub Form_Load()
    Set m_Expressions = New Collection
    List1.Clear
    TxtInput.Text = vbNullString
End Sub

Private Sub BtnCipher_Click(Index As Integer): ConstAdd CStr(Index): End Sub
Private Sub BtnPoint_Click():                  ConstAdd ".": End Sub
Private Sub BtnDelete_Click(): m_Num = vbNullString: UpdateView: End Sub
Private Sub BtnBack_Click()
    If Len(m_Num) = 0 Then Exit Sub
    m_Num = Left(m_Num, Len(m_Num) - 1): UpdateView
End Sub
Private Sub BtnSign_Click()
    If Left(m_Num, 1) = "-" Then m_Num = Mid(m_Num, 2) Else m_Num = "-" & m_Num
    UpdateView
End Sub

Sub ConstAdd(ByVal c As String)
    If m_TBHasResult Then
        m_Num = c
        m_TBHasResult = False
    Else
        m_Num = m_Num & c
    End If
    UpdateView
End Sub

Private Sub List1_DblClick()
    If List1.ListCount = 0 Then Exit Sub
    Dim i As Long: i = List1.ListIndex + 1
    Dim ex As Expression: Set ex = m_Expressions.Item(i)
    m_Num = Trim(Str(ex.Eval))
    UpdateView
End Sub

Private Sub TxtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    'Debug.Print "KeyDown  KeyCode:  " & KeyCode & "; Shift: " & Shift
End Sub

Private Sub TxtInput_KeyUp(KeyCode As Integer, Shift As Integer)
    'Debug.Print "KeyUp    KeyCode:  " & KeyCode & "; Shift: " & Shift
End Sub

Private Sub TxtInput_KeyPress(KeyAscii As Integer)
    'Debug.Print "KeyPress KeyAscii: " & KeyAscii
    Select Case KeyAscii
    Case 8: 'BackDelete
    Case 13 ' =
        KeyAscii = 0
        BtnExprEvaluate_Click
    Case 42 ' *
        KeyAscii = 0
        BtnExprMul_Click
    Case 43 ' +
        KeyAscii = 0
        BtnExprAdd_Click
    'Case 44 ' ,
    '    KeyAscii = 46
    Case 45 ' -
        KeyAscii = 0
        BtnExprSubt_Click
    Case 47 ' /
        KeyAscii = 0
        BtnExprDiv_Click
    Case 44, 46, 48 To 57: '.,0123456789
        If KeyAscii = 44 Then KeyAscii = 46
        'If m_TBHasResult Then
        '    TxtInput = vbNullString
        '    m_Num = Chr(KeyAscii)
        'Else
            'm_Num = TxtInput.Text 'm_Num & Chr(KeyAscii)
        'End If
        'm_TBHasResult = False
    Case 178 ' �
        KeyAscii = 0
        BtnExprSqr_Click
    Case 179 ' �
        KeyAscii = 0
        BtnExprCub_Click
    Case Else: KeyAscii = 0
    End Select
End Sub

Private Function GetExprLitNum() As ExprLitNum
    Dim v As Double: v = Val(TxtInput.Text) 'm_Num)
    m_Num = vbNullString
    Set GetExprLitNum = MNew.ExprLitNum(v)
End Function

Private Sub BtnExprAdd_Click():    UpdateData MNew.OpBinAdd(GetExprLitNum):  End Sub
Private Sub BtnExprSubt_Click():   UpdateData MNew.OpBinSubt(GetExprLitNum): End Sub
Private Sub BtnExprMul_Click():    UpdateData MNew.OpBinMul(GetExprLitNum):  End Sub
Private Sub BtnExprDiv_Click():    UpdateData MNew.OpBinDiv(GetExprLitNum):  End Sub

Private Sub BtnExprPow_Click():    UpdateData MNew.OpBinPow(GetExprLitNum):  End Sub

Private Sub BtnExprFact_Click()
    UpdateData MNew.ExprOpFact(GetExprLitNum)
    BtnExprEvaluate_Click
End Sub

Private Sub BtnExprSqrt_Click()
    UpdateData MNew.ExprOpSqrt(GetExprLitNum)
    BtnExprEvaluate_Click
End Sub

Private Sub BtnExprSqr_Click()
    UpdateData MNew.ExprOpSqr(GetExprLitNum)
    BtnExprEvaluate_Click
End Sub

Private Sub BtnExprCub_Click()
    UpdateData MNew.ExprOpCub(GetExprLitNum)
    BtnExprEvaluate_Click
End Sub

Private Sub BtnExprLN_Click()
    UpdateData MNew.ExprOpLN(GetExprLitNum)
    BtnExprEvaluate_Click
End Sub

Private Sub BtnExprLog10_Click()
    UpdateData MNew.ExprOpLog10(GetExprLitNum)
    BtnExprEvaluate_Click
End Sub

Private Sub BtnExpr1DivX_Click()
    UpdateData MNew.ExprOp1DivX(GetExprLitNum)
    BtnExprEvaluate_Click
End Sub

Private Sub BtnExprAbs_Click()
    UpdateData MNew.ExprOpAbs(GetExprLitNum)
    BtnExprEvaluate_Click
End Sub

Private Sub BtnExprBracClose_Click()
    '
End Sub

Private Sub BtnExprBracOpen_Click()
    '
End Sub

Private Sub BtnExprConstE_Click()
    m_Num = Trim(Str(Exp(1)))
    UpdateView
End Sub

Private Sub BtnExprConstPi_Click()
    m_Num = "3.14159265358979"
    UpdateView
End Sub

Private Sub BtnExprExp_Click()
    '
End Sub

Private Sub BtnExprMod_Click()
    '
End Sub

Private Sub BtnExprPow10_Click()
    UpdateData MNew.ExprOpPow10(GetExprLitNum)
    BtnExprEvaluate_Click
End Sub

Private Sub BtnExprEvaluate_Click()
    'Dim v As Double: v = Val(m_Num)
    'm_Num = vbNullString
    If Not m_LastBinOp Is Nothing Then
        Set m_LastBinOp.Expr2 = GetExprLitNum ' MNew.ExprConst(v)
        'Set m_LastBinOp = Nothing '????
    End If
    'm_Num = vbNullString
    If m_Expression Is Nothing Then Exit Sub
    If m_Expressions.Count = 0 Then
        m_Expressions.Add m_Expression
    Else
        m_Expressions.Add m_Expression, , 1
    End If
    m_Num = Trim(Str(m_Expression.Eval))
    UpdateView
    
    'TxtInput.Text = m_Num 'vbNullString ' Trim(Str(m_Expression.Eval))
    m_TBHasResult = True
    'm_Num = vbNullString
    Set m_Expression = Nothing
    'TxtInput.Text
End Sub

Private Sub UpdateData(Expr As Expression)
    Set m_Expression = Expr
    'If TypeOf Expr Is OperatorBinary Then
    Set m_LastBinOp = m_Expression
    UpdateView
End Sub

Sub UpdateView()
    TxtInput.Text = m_Num
    Dim sErg As String
    If m_Expression Is Nothing Then sErg = vbNullString Else sErg = EvalToStr(m_Expression, , True)
    Label1.Caption = sErg
    If m_Expressions Is Nothing Then Exit Sub
    If m_Expressions.Count = 0 Then Exit Sub
    Dim i As Long: i = List1.ListIndex
    List1.Clear
    Dim ex As Expression
    For Each ex In m_Expressions
        If Not ex Is Nothing Then
            List1.AddItem EvalToStr(ex, True)
        End If
    Next
    If i >= 0 Then List1.Selected(i) = True
End Sub

Function EvalToStr(e As Expression, Optional ByVal inclResult As Boolean = False, Optional ByVal inclEquSign As Boolean = False) As String
    If e Is Nothing Then Exit Function
    'If Not e.CanEval Then Exit Function
    EvalToStr = e.ToStr & IIf(inclResult, " = " & Trim(Str(e.Eval)), IIf(inclEquSign And e.CanEval, " = ", ""))
End Function

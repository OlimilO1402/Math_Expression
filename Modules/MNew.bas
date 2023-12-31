Attribute VB_Name = "MNew"
Option Explicit

Public Enum OperatorRank
    'lowest rank
    Rank0None = 0
    Rank1ExprConst
    Rank2ExprAndOrXOrNot
    Rank2ExprOpAddSubt
    Rank3ExprOpSubtrah
    Rank4ExprOpMulDiv
    Rank5ExprOpDivisor
    Rank6ExprOpPow
    Rank7ExprOpFact
    Rank8ExprOpNeg
    Rank9ExprOpBrac
    'highest rank
End Enum

' v ############################## v '    Literals    ' v ############################## v '
Public Function ExprLitBol(ByVal aValue As Boolean) As ExprLitBol 'can only be True or False
    Set ExprLitBol = New ExprLitBol: ExprLitBol.New_ aValue
End Function

Public Function ExprLitDat(ByVal aValue As Boolean) As ExprLitDat 'Values of date and time
    Set ExprLitDat = New ExprLitDat: ExprLitDat.New_ aValue
End Function

Public Function ExprLitNum(ByVal aValue) As ExprLitNum            'any numeric value, byte, int16, int32, int64, float32, float64
    Set ExprLitNum = New ExprLitNum: ExprLitNum.New_ aValue
End Function

Public Function ExprLitStr(ByVal aValue) As ExprLitStr            'string
    Set ExprLitStr = New ExprLitStr: ExprLitStr.New_ aValue
End Function

' v ############################## v '    Boolean Operators    ' v ############################## v '
Public Function ExprOpBolAnd(aLHS As Expression, aRHS As Expression) As ExprOpBolAnd
    Set ExprOpBolAnd = New ExprOpBolAnd: ExprOpBolAnd.New_ aLHS, aRHS
End Function

Public Function ExprOpBolEqual(aLHS As Expression, aRHS As Expression) As ExprOpBolEqual
    Set ExprOpBolEqual = New ExprOpBolEqual: ExprOpBolEqual.New_ aLHS, aRHS
End Function

Public Function ExprOpBolOr(aLHS As Expression, aRHS As Expression) As ExprOpBolOr
    Set ExprOpBolOr = New ExprOpBolOr: ExprOpBolOr.New_ aLHS, aRHS
End Function

Public Function ExprOpBolXor(aLHS As Expression, aRHS As Expression) As ExprOpBolXor
    Set ExprOpBolXor = New ExprOpBolXor: ExprOpBolXor.New_ aLHS, aRHS
End Function


' v ############################## v '    Operators Binary    ' v ############################## v '
'+
Public Function OpBinAdd(aLHS As Expression) As ExprOpAdd
    Set OpBinAdd = New ExprOpAdd: OpBinAdd.New_ aLHS
End Function
Public Function ExprOpAdd(aLHS As Expression, aRHS As Expression) As ExprOpAdd
    Set ExprOpAdd = COperatorBinary(OpBinAdd(aLHS), aRHS)
End Function

'-
Public Function OpBinSubt(aLHS As Expression) As ExprOpSubt
    Set OpBinSubt = New ExprOpSubt: OpBinSubt.New_ aLHS
End Function
Public Function ExprOpSubt(aLHS As Expression, aRHS As Expression) As ExprOpSubt
    Set ExprOpSubt = COperatorBinary(OpBinSubt(aLHS), aRHS)
End Function

'*
Public Function OpBinMul(aLHS As Expression) As ExprOpMul
    Set OpBinMul = New ExprOpMul: OpBinMul.New_ aLHS
End Function
Public Function ExprOpMul(aLHS As Expression, aRHS As Expression) As ExprOpMul
    Set ExprOpMul = COperatorBinary(OpBinMul(aLHS), aRHS)
End Function

'/
Public Function OpBinDiv(aLHS As Expression) As ExprOpDiv
    Set OpBinDiv = New ExprOpDiv: OpBinDiv.New_ aLHS
End Function
Public Function ExprOpDiv(aLHS As Expression, aRHS As Expression) As ExprOpDiv
    Set ExprOpDiv = COperatorBinary(OpBinDiv(aLHS), aRHS)
End Function

'^
Public Function OpBinPow(aLHS As Expression) As ExprOpPow
    Set OpBinPow = New ExprOpPow: OpBinPow.New_ aLHS
End Function
Public Function ExprOpPow(aLHS As Expression, aRHS As Expression) As ExprOpPow
    Set ExprOpPow = COperatorBinary(OpBinPow(aLHS), aRHS)
End Function

Private Function COperatorBinary(aOpBin As Expression, aRHSExpr As Expression) As Expression  'OperatorBinary
    Set COperatorBinary = aOpBin: Set COperatorBinary.Expr2 = aRHSExpr
End Function

' v ############################## v '    Operators Unary    ' v ############################## v '
' Associated Left
Public Function ExprOpNeg(aRHS As Expression) As ExprOpNeg
    Set ExprOpNeg = New ExprOpNeg: ExprOpNeg.New_ aRHS
End Function

Public Function ExprOpSqrt(aRHS As Expression) As ExprOpSqrt
    Set ExprOpSqrt = New ExprOpSqrt: ExprOpSqrt.New_ aRHS
End Function


' Associated Right
Public Function ExprOpCub(aLHS As Expression) As ExprOpCub
    Set ExprOpCub = New ExprOpCub: ExprOpCub.New_ aLHS
End Function

Public Function ExprOp1DivX(aRHS As Expression) As ExprOp1DivX
    Set ExprOp1DivX = New ExprOp1DivX: ExprOp1DivX.New_ aRHS
End Function

Public Function ExprOpFact(aLHS As Expression) As ExprOpFact
    Set ExprOpFact = New ExprOpFact: ExprOpFact.New_ aLHS
End Function

Public Function ExprOpSqr(aLHS As Expression) As ExprOpSqr
    Set ExprOpSqr = New ExprOpSqr: ExprOpSqr.New_ aLHS
End Function

Public Function ExprOpPow10(aRHS As Expression) As ExprOpPow10
    Set ExprOpPow10 = New ExprOpPow10: ExprOpPow10.New_ aRHS
End Function

' v ############################## v '      Braces     ' v ############################## v '
Public Function ExprOpBrac(aIns As Expression) As ExprOpBrac
    Set ExprOpBrac = New ExprOpBrac: ExprOpBrac.New_ aIns
End Function


' v ############################## v '    Functions    ' v ############################## v '
Public Function ExprFunction(ByVal aFuncName As String, ByVal CallableObject As Object, Optional ByVal ExprInside As Expression = Nothing) As ExprFunction
    Set ExprFunction = New ExprFunction: ExprFunction.New_ aFuncName, CallableObject, ExprInside
End Function

Public Function ExprOpAbs(aIns As Expression) As ExprOpAbs
    Set ExprOpAbs = New ExprOpAbs: ExprOpAbs.New_ aIns
End Function

Public Function ExprOpLN(aRHS As Expression) As ExprOpLN
    Set ExprOpLN = New ExprOpLN: ExprOpLN.New_ aRHS
End Function

Public Function ExprOpLog10(aRHS As Expression) As ExprOpLog10
    Set ExprOpLog10 = New ExprOpLog10: ExprOpLog10.New_ aRHS
End Function

Public Function ExprOpLogN(aLHS As Expression, Optional aRHS As Expression = Nothing) As ExprOpLogN
    Set ExprOpLogN = New ExprOpLogN: ExprOpLogN.New_ aLHS, aRHS
End Function

Public Function ExprOpIIf(aCond As Expression, Optional ExprTrue As Expression = Nothing, Optional ExprFalse As Expression = Nothing) As ExprOpIIf
    Set ExprOpIIf = New ExprOpIIf: ExprOpIIf.New_ aCond, ExprTrue, ExprFalse
End Function


' v ############################## v '    Formatters    ' v ############################## v '
Public Function FormatAlg(ByVal IsCondensed As Boolean, Optional ByVal ExcelCompatible As Boolean = False) As FormatAlg
    Set FormatAlg = New FormatAlg: FormatAlg.New_ IsCondensed, ExcelCompatible
End Function

Public Function FormatRPN(Optional ByVal SeparatorIsNewLine As Boolean = False) As FormatRPN
    Set FormatRPN = New FormatRPN: FormatRPN.New_ SeparatorIsNewLine
End Function

Attribute VB_Name = "MNew"
Option Explicit
Public Enum OperatorRank
    'lowest rank
    Rank0None = 0
    Rank1ExprConst
    Rank2ExprOpAddSubt
    Rank3ExprOpMulDiv
    Rank4ExprOpPow
    Rank5ExprOpFact
    Rank6ExprOpNeg
    Rank7ExprOpBrac
    'highest rank
End Enum

Public Function ExprLitBol(ByVal aValue As Boolean) As ExprLitBol 'can only be True or False
    Set ExprLitBol = New ExprLitBol: ExprLitBol.New_ aValue
End Function

Public Function ExprLitNum(ByVal aValue) As ExprLitNum       'any numeric value, int or float
    Set ExprLitNum = New ExprLitNum: ExprLitNum.New_ aValue
End Function

Public Function OpBinAdd(aLHS As Expression) As ExprOpAdd
    Set OpBinAdd = New ExprOpAdd: OpBinAdd.New_ aLHS
End Function
Public Function ExprOpAdd(aLHS As Expression, aRHS As Expression) As ExprOpAdd
    Set ExprOpAdd = COperatorBinary(OpBinAdd(aLHS), aRHS)
End Function

Public Function ExprOpAbs(aIns As Expression) As ExprOpAbs
    Set ExprOpAbs = New ExprOpAbs: ExprOpAbs.New_ aIns
End Function

Public Function ExprOpBrac(aIns As Expression) As ExprOpBrac
    Set ExprOpBrac = New ExprOpBrac: ExprOpBrac.New_ aIns
End Function

Public Function ExprOpPow10(aRHS As Expression) As ExprOpPow10
    Set ExprOpPow10 = New ExprOpPow10: ExprOpPow10.New_ aRHS
End Function

Public Function ExprOpLN(aRHS As Expression) As ExprOpLN
    Set ExprOpLN = New ExprOpLN: ExprOpLN.New_ aRHS
End Function

Public Function ExprOpLog10(aRHS As Expression) As ExprOpLog10
    Set ExprOpLog10 = New ExprOpLog10: ExprOpLog10.New_ aRHS
End Function

Public Function ExprOpCub(aLHS As Expression) As ExprOpCub
    Set ExprOpCub = New ExprOpCub: ExprOpCub.New_ aLHS
End Function

Public Function OpBinDiv(aLHS As Expression) As ExprOpDiv
    Set OpBinDiv = New ExprOpDiv: OpBinDiv.New_ aLHS
End Function
Public Function ExprOpDiv(aLHS As Expression, aRHS As Expression) As ExprOpDiv
    Set ExprOpDiv = COperatorBinary(OpBinDiv(aLHS), aRHS)
End Function

Public Function ExprOp1DivX(aRHS As Expression) As ExprOp1DivX
    Set ExprOp1DivX = New ExprOp1DivX: ExprOp1DivX.New_ aRHS
End Function

Public Function ExprOpFact(aLHS As Expression) As ExprOpFact
    Set ExprOpFact = New ExprOpFact: ExprOpFact.New_ aLHS
End Function

Public Function OpBinMul(aLHS As Expression) As ExprOpMul
    Set OpBinMul = New ExprOpMul: OpBinMul.New_ aLHS
End Function
Public Function ExprOpMul(aLHS As Expression, aRHS As Expression) As ExprOpMul
    Set ExprOpMul = COperatorBinary(OpBinMul(aLHS), aRHS)
End Function

Public Function ExprOpNeg(aRHS As Expression) As ExprOpNeg
    Set ExprOpNeg = New ExprOpNeg: ExprOpNeg.New_ aRHS
End Function

Public Function OpBinPow(aLHS As Expression) As ExprOpPow
    Set OpBinPow = New ExprOpPow: OpBinPow.New_ aLHS
End Function
Public Function ExprOpPow(aLHS As Expression, aRHS As Expression) As ExprOpPow
    Set ExprOpPow = COperatorBinary(OpBinPow(aLHS), aRHS)
End Function

Public Function ExprOpSqr(aLHS As Expression) As ExprOpSqr
    Set ExprOpSqr = New ExprOpSqr: ExprOpSqr.New_ aLHS
End Function

Public Function ExprOpSqrt(aRHS As Expression) As ExprOpSqrt
    Set ExprOpSqrt = New ExprOpSqrt: ExprOpSqrt.New_ aRHS
End Function

Public Function OpBinSubt(aLHS As Expression) As ExprOpSubt
    Set OpBinSubt = New ExprOpSubt: OpBinSubt.New_ aLHS
End Function
Public Function ExprOpSubt(aLHS As Expression, aRHS As Expression) As ExprOpSubt
    Set ExprOpSubt = COperatorBinary(OpBinSubt(aLHS), aRHS)
End Function

Private Function COperatorBinary(aOpBin As OperatorBinary, aRHSExpr As Expression) As OperatorBinary
    Set COperatorBinary = aOpBin: Set COperatorBinary.SecondExpr = aRHSExpr
End Function


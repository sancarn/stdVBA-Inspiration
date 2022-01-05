VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMathParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'********************************************************************************
'* CLASS: clsMathParser                                    v.4.2.1   1.03.07    *
'*                                                        Leonardo Volpi        *
'*                                                        Michael Ruder         *
'*                                                        Thomas Zeutschler     *
'*                                                        Lieven Dossche        *
'*                                                        Arnaud d.Grammont     *
'*  Math-Physical Expression Evaluation                                         *
'*  for VB 6, VBA 97/2000/XP                                                    *
'*  use the mMathSpecFun.bas module for special functions                       *
'********************************************************************************
Option Explicit
'-------------------------------------------------------------------------------
' CONSTANTS
'-------------------------------------------------------------------------------
Const HiVT       As Long = 100 'max variables for string
Const HiET       As Long = 200 'max functions/operations for string
Const HiARG      As Long = 20  'max function arguments
Const HiF2       As Long = 20  'max nested multi-variable functions 12.06.04
Const HFOffset   As Long = 200 'offset for multi-var functions assignement
Const PI_        As Double = 3.14159265358979
Const DP_SET     As Boolean = True 'decimal point setting "."
'-------------------------------------------------------------------------------------------------------------
'FUNCTION ALIAS
'-------------------------------------------------------------------------------------------------------------
'tables of fixed parameters functions
Const Fun1V As String = "Abs Atn Atan Cos Exp Fix Int Ln Log Rnd Sgn Sin Sqr Cbr Tan Acos Asin " & _
    "Cosh Sinh Tanh Acosh Asinh Atanh Fact Not Erf Gamma Gammaln Digamma Zeta Ei AiryA AiryB " & _
    "csc sec cot acsc asec acot csch sech coth acsch asech acoth dec rad deg grad neg Si Ci " & _
    "fresnels fresnelc J0 Y0 I0 K0 Psi Year Month Day Hour Minute Second"
Const Fun2V As String = "Comb Mod And Or Xor Beta Root Round Nand Nor NXor Perm " & _
    "LogN DPoisson CPoisson Ein BesselI BesselJ BesselK BesselY GammaI PolyLe PolyHe PolyCh PolyLa " & _
    "elli1 elli2 wtri wsqr wsaw wraise wparab step"
Const Fun3V As String = "DNorm CNorm DBinom CBinom BetaI Clip wrect wtrapez wlin wpulse wsteps wexp wexpb wpulsef wripple " & _
     "DateSerial TimeSerial"
Const Fun4V As String = "HypGeom wring wam wfm"
'Table of variable parameters functions
Const FunxV As String = "max min sum mean meanq meang var varp stdev stdevp mcm mcd lcm gcd "
'-------------------------------------------------------------------------------
' FUNCTION ENUMERATIONS
'-------------------------------------------------------------------------------
Const symRight  As Long = -2        'right function (internal)
Const symARGUMENT As Long = -1      'An Argument    (internal)
Const symPlus   As Long = 0         '+
Const symMinus  As Long = 1         '-
Const symMul    As Long = 2         '*
Const symDiv    As Long = 3         '/
Const symPercent As Long = 4        '% percentage
Const symDivInt As Long = 5         '\ integer division, added MR 20-06-02
Const symPov    As Long = 6         '^
Const symAbs    As Long = 7         '"abs", "|.|"
Const symAtn    As Long = 8         '"atn"
Const symCos    As Long = 9         '"cos"
Const symSin    As Long = 11        '"sin"
Const symExp    As Long = 12        '"exp"
Const symFix    As Long = 13        '"fix"
Const symInt    As Long = 14        '"int"
Const symLn     As Long = 15        '"ln"
Const symLog    As Long = 16        '"log"
Const symRnd    As Long = 17        '"rnd"
Const symSgn    As Long = 18        '"sgn"
Const symSqr    As Long = 19        '"sqr"
Const symTan    As Long = 20        '"tan"
Const symAcos   As Long = 21        '"acos"
Const symAsin   As Long = 22        '"asin"
Const symCosh   As Long = 23        '"cosh"
Const symSinh   As Long = 24        '"sinh"
Const symTanh   As Long = 25        '"tanh"
Const symAcosh  As Long = 26        '"acosh"
Const symAsinh  As Long = 27        '"asinh"
Const symAtanh  As Long = 28        '"atanh"
Const symmod    As Long = 29        '"mod"
Const symFact   As Long = 30        '"fact", "!"
Const symComb   As Long = 31        '"combinations or binomial coeff."
Const symGT     As Long = 36        '">"
Const symGE     As Long = 37        '">As Long ="
Const symLT     As Long = 38        '"<"
Const symLE     As Long = 39        '"<As Long ="
Const symEQ     As Long = 40        '"As Long ="
Const symNE     As Long = 41        '"<>"
Const symAnd    As Long = 42        '"and"
Const symOr     As Long = 43        '"or"
Const symNot    As Long = 44        '"not"
Const symXor    As Long = 45        '"xor"
Const symErf    As Long = 46        '"erf"
Const symGamma  As Long = 47        '"gamma"   Euler's gamma function
Const symGammaln As Long = 48       '"gammaln" logarithm of gamma
Const symDigamma As Long = 49       '"digamma"
Const symBeta   As Long = 50        '"beta"
Const symZeta   As Long = 51        '"zeta"
Const symEi     As Long = 52        '"ei"  Exponetial integral
Const symCsc    As Long = 53        '"csc  cosecant"
Const symSec    As Long = 54        '"sec  secant"
Const symCot    As Long = 55        '"cot  cotangent"
Const symACsc   As Long = 56        '"acsc  inverse cosecant"
Const symASec   As Long = 57        '"asec  inverse secant"
Const symACot   As Long = 58        '"acot  inverse cotangent"
Const symCsch   As Long = 59        '"csch  hyperbolic cosecant"
Const symSech   As Long = 60        '"sech  hyperbolic secant"
Const symCoth   As Long = 61        '"coth  hyperbolic cotangent"
Const symACsch  As Long = 62        '"acsch inverse hyperbolic cosecant"
Const symASech  As Long = 63        '"asech inverse hyperbolic secant"
Const symACoth  As Long = 64        '"acoth inverse hyperbolic cotangent"
Const symCbr    As Long = 65        '"cbr cube root"
Const symRoot   As Long = 66        '"root n-th root"
Const symDec    As Long = 67        '"dec  decimal part"
Const symRad    As Long = 68        '"rad  convert radiant to current angle unit"
Const symDeg    As Long = 69        '"deg  convert degree 360 to current angle unit"
Const symRound  As Long = 70        '"round"
Const symGrad   As Long = 71        '"grad  convert degree 400 to current angle unit"
Const symNAnd   As Long = 72        '"nand"
Const symNOr    As Long = 73        '"nor"
Const symNXor   As Long = 74        '"NXor"
Const symNeg    As Long = 75        '"neg sign change"
Const symPerm   As Long = 76        '"Perm Permutations"
Const symLogN   As Long = 77        '"LogN log N-base"
Const symDPoiss As Long = 78        'Poisson density
Const symCPoiss As Long = 79        'Poisson cumulative distribution
Const symEin    As Long = 80         '"Ein" Exponential integral n
Const symSi     As Long = 81         '"Integral Sine
Const symCi     As Long = 82         '"Integral cosine
Const symFresS  As Long = 83         '"Fresnel's Integral sine
Const symFresC  As Long = 84         '"Fresnel's Integral cosine
Const symBessJ  As Long = 85         '"Bessel's Jn(x)  1st kind
Const symBessY  As Long = 86         '"Bessel's Yn(x)  2nd kind
Const symBessI  As Long = 87         '"Bessel's In(x)  1st kind modified
Const symBessK  As Long = 88         '"Bessel's Kn(x)  2nd kind modified
Const symJ0     As Long = 89        '"Bessel's J(x)  1st kind
Const symY0     As Long = 90        '"Bessel's Y(x)  2st kind
Const symK0     As Long = 91        '"Bessel's K(x)  1st kind modified
Const symI0     As Long = 92        '"Bessel's I(x)  2st kind modified
Const symGammaI As Long = 94        '"Gamma Incomplete
Const symPolyLe As Long = 95        '"Legendre's polynomial
Const symPolyLa As Long = 96        '"Laguerre's polynomial
Const symPolyHe As Long = 97        '"Hermite's polynomial
Const symPolyCh As Long = 98        '"Chebycev's polynomial
Const symAiryA  As Long = 99        '"Airy Ai(x) function
Const symAiryB  As Long = 100       '"Airy Bi(x) function
Const symEllipt1 As Long = 101      '"Elliptic integral 1st kind
Const symEllipt2 As Long = 102      '"Elliptic integral 2nd kind
Const symWtri   As Long = 103       'Triangular wave
Const symWsqr   As Long = 104       'Square wave
Const symWsaw   As Long = 105       'saw wave
Const symWraise As Long = 106       'raise wave
Const symWparab As Long = 107       'parabolic wave
Const symStep   As Long = 108       'step
' > Berend 20041216
Const symYear   As Long = 150       ' Year(date) function
Const symMonth  As Long = 151       ' Month(date) function
Const symDay    As Long = 152       ' Day(date) function
Const symHour   As Long = 153       ' Hour(date) function
Const symMinute As Long = 154       ' Minute(date) function
Const symSecond As Long = 155       ' Second(date) function
' < Berend 20041216

'constant > 200 for multi-arguments function.
Const symDnorm  As Long = HFOffset + 1      'Normal Density
Const symCnorm  As Long = HFOffset + 2      'Normal Cumulative Distribution
Const symDBinom As Long = HFOffset + 3      'Binomial Density
Const symCBinom As Long = HFOffset + 4      'Binomial Cumulative Distribution
Const symHypGeo As Long = HFOffset + 5      'Hypergeometric function
Const symBetaI  As Long = HFOffset + 6      'Beta Incomplete
Const symClip   As Long = HFOffset + 7      'Clipping function
Const symWrect  As Long = HFOffset + 8      'rectang. wave
Const symWtrapez As Long = HFOffset + 9     'trapez wavw
Const symWlin   As Long = HFOffset + 10     'linear wavw
Const symWpulse As Long = HFOffset + 11     'pulse
Const symWsteps As Long = HFOffset + 12     'steps wave
Const symWexp   As Long = HFOffset + 13     'expon. wavw
Const symWexpb  As Long = HFOffset + 14     'expon. bipolar wave
Const symWpulsef As Long = HFOffset + 15    'pulse filtered
Const symWripple As Long = HFOffset + 16    'ripple wave
Const symWring  As Long = HFOffset + 17     'ringing wavw
Const symWam    As Long = HFOffset + 18     'AM wave
Const symWfm    As Long = HFOffset + 19     'FM wave
' > Berend 20041213
Const symDateSerial As Long = HFOffset + 20 ' DateSerial
Const symTimeSerial As Long = HFOffset + 21 ' TimeSerial
' < Berend 20041213

'constant > 300 for variables arguments function.
Const symMin    As Long = HFOffset + 101        '"min"
Const symMax    As Long = HFOffset + 102        '"max"
Const symSum    As Long = HFOffset + 103        '"Sum"
Const symMean   As Long = HFOffset + 104        '"arithmetic mean"
Const symMeanq  As Long = HFOffset + 105        '"quadratic mean"
Const symMeang  As Long = HFOffset + 106        '"geometric mean"
Const symVar    As Long = HFOffset + 107        '"variance"
Const symVarp   As Long = HFOffset + 108        '"variance pop."
Const symStdev  As Long = HFOffset + 109        '"std. deviation"
Const symStdevp As Long = HFOffset + 110        '"std. deviation pop."
Const symMcd    As Long = HFOffset + 111        '"gcd"
Const symMcm    As Long = HFOffset + 112        '"lcm"

'-------------------------------------------------------------------------------
' TYPE DECLARATIONS
'-------------------------------------------------------------------------------
Private Type T_VTREC           'Variable record
  idx              As Long
  Nome             As String
  value            As Double
  sign             As Long
  Init             As Boolean
End Type

Private Type T_ETREC           'Expression Structure record
  Fun              As String
  FunTok           As Long
  arg(1 To HiARG)  As T_VTREC
  ArgTop           As Long
  ArgOf            As Long
  ArgIdx           As Long
  value            As Double
  sign             As Long
  PosInExpr        As Long
  PriLvl           As Long
  PriIdx           As Long
  Cond             As Long
End Type

Private Type T_Funk            'Function record
  FunName           As String
  ArgN              As Long
  ArgCount          As Long
  sign              As Long
  PosInExpr         As Long
End Type
'-------------------------------------------------------------------------------
' LOCALS
'-------------------------------------------------------------------------------
Dim Expr            As String   'Expression to parse/evaluate
Dim ExprOK          As Boolean  'Parsing result
Dim VT()            As T_VTREC  'Variables Table
Dim ET()            As T_ETREC  'Expression Structure Table
Dim VTtop           As Long
Dim ETtop           As Long
Dim ErrMsg          As String   'error message
Dim ErrPos          As Long     'error position
Dim ErrId           As Long     'error id
Dim Angle           As String   'RAD GRAD DEG
Dim DecRound        As Long
Dim CvAngleCoeff    As Double
Dim VarsTbl         As Collection   'additional object (bb 6-1-04)
Dim funk(1 To 5)    As String    'Functions tables
Dim iInit           As Long      'Variables initialization counter
Dim VarIniExp       As Boolean   'Flag Variables initialization explicit
Dim ErrorTbl()      As String    'Error Message Table
Dim DMS_conv        As Boolean   'Flag dms conversion
Dim Unit_conv       As Boolean   'Flag unit conversion
Dim ArgSep          As String    'arguments separation symbol
Dim DecSep          As String    'decimal separation symbol

'-------------------------------------------------------------------------------
' FUNCTIONS, METHODS and PROPERTIES
'-------------------------------------------------------------------------------
'class starting routine
Private Sub Class_Initialize()
'decimal point setting
  If DP_SET Then
    'independent from the international machine setting
    DecSep = "."
    ArgSep = ","
  Else
    'follow the the international machine setting
    DecSep = Decimal_Point_Is
    If DecSep = "," Then ArgSep = ";" Else ArgSep = ","
  End If
'initialize multi variable functions lists
  funk(2) = Fun2V    'table of 2 parameters functions
  funk(3) = Fun3V    'table of 3 parameters functions
  funk(4) = Fun4V    'table of 4 parameters functions
  funk(5) = FunxV    'table of variable parameters functions
  VarIniExp = False  'variables assignement explicit
  Angle = "RAD"      'default angle radiant
  CvAngleCoeff = 1   'default angle convertion coefficient
  DMS_conv = True    'enable dms conversion
  Unit_conv = True   'enable unit conversion
  ErrorTab_Init      'load error message table
End Sub
' store expression as array of records. check syntax
Public Function StoreExpression(ByVal strExpr As String) As Boolean
  Expr = Trim(strExpr)
  ExprOK = Parse(Expr)
  StoreExpression = ExprOK
End Function
'-------------------------------------------------------------------------------
' get the expression
Public Property Get Expression() As String
  Expression = Expr
End Property
'-------------------------------------------------------------------------------
' get the top of the var array (=N-1 bacause starts on 0)
Public Property Get VarTop() As Long
  VarTop = VTtop
End Property
'-------------------------------------------------------------------------------
' get name of a variable. VL
Public Property Get VarName(ByVal index As Long) As String
  If index <= VTtop Then
    VarName = VT(index).Nome
  End If
End Property
'-------------------------------------------------------------------------------
'(old version) get value assigned to a variable
Public Property Get VarValue(ByVal index As Long) As Double
  If index <= VTtop Then
    VarValue = VT(index).value
  End If
End Property
'-------------------------------------------------------------------------------
'(old version) assign a value to a certain variable
Public Property Let VarValue(ByVal index As Long, ByVal VarVal As Double)
  If index <= VTtop Then
    VT(index).value = VarVal
    VT(index).Init = True
    iInit = iInit + 1
  End If
End Property
'-------------------------------------------------------------------------------
'(old version) get value assigned to a variable passed by its string name
Public Property Get VarSymb(ByVal name As String) As Double
    VarSymb = 0#
    On Error Resume Next
    VarSymb = VarsTbl(name)
End Property
'-------------------------------------------------------------------------------
'(old version) assign a value to a certain variable passed by its string name
Public Property Let VarSymb(ByVal name As String, ByVal VarVal As Double)
    On Error Resume Next
    VT(VarsTbl(name)).value = VarVal
    VT(VarsTbl(name)).Init = True
    iInit = iInit + 1
End Property
'-------------------------------------------------------------------------------
' get value assigned to a variable passed by string or index (14.6.2004)
Public Property Get variable(ByVal name) As Double
    variable = 0
    If VarType(name) = vbString Then
        On Error Resume Next
        variable = VT(VarsTbl(name)).value
    Else
        On Error Resume Next
        variable = VT(name).value
    End If
End Property
'-------------------------------------------------------------------------------
' assign a value to a certain variable passed by string or index (14.6.2004)
Public Property Let variable(ByVal name As Variant, ByVal VarVal As Double)
Dim Id As Long
    On Error GoTo Error_Handler
    If VarType(name) = vbString Then
        Id = VarsTbl(name)
    Else
        Id = name
    End If
    VT(Id).value = VarVal
    If VarIniExp Then  'Explicit initialization
        If VT(Id).Init = False Then
            VT(Id).Init = True
            iInit = iInit + 1
        End If
    End If
Exit Property
Error_Handler:
    'nothing to do
End Property
'-------------------------------------------------------------------------------
' get current setting for angle computing (RAD (default), DEG or GRAD)
Public Property Get AngleUnit() As String
    AngleUnit = Angle
End Property
'-------------------------------------------------------------------------------
' set the unit of measure for angle computing (RAD (default), DEG or GRAD)
Public Property Let AngleUnit(ByVal AngleUM As String)
    Select Case UCase(AngleUM)
        Case "DEG"
            Angle = "DEG"
            CvAngleCoeff = PI_ / 180
        Case "GRAD"
            Angle = "GRAD"
            CvAngleCoeff = PI_ / 200
        Case "RAD"
            Angle = "RAD"
            CvAngleCoeff = 1
    End Select
End Property
'-------------------------------------------------------------------------------
' get current setting for assignement constrain
Public Property Get OpAssignExplicit() As Boolean
    OpAssignExplicit = VarIniExp
End Property
'-------------------------------------------------------------------------------
' sets the variables assignement constrain
Public Property Let OpAssignExplicit(ByVal AssignEnable As Boolean)
     VarIniExp = AssignEnable
End Property
'-------------------------------------------------------------------------------
' get current setting for unit conversion
Public Property Get OpUnitConv() As Boolean
    OpUnitConv = Unit_conv
End Property
'-------------------------------------------------------------------------------
' enable/disable the unit conversion
Public Property Let OpUnitConv(ByVal ConvEnable As Boolean)
     Unit_conv = ConvEnable
End Property
'-------------------------------------------------------------------------------
' get current setting for dms conversion
Public Property Get OpDMSConv() As Boolean
    OpDMSConv = DMS_conv
End Property
'-------------------------------------------------------------------------------
' enable/disable the dms conversion
Public Property Let OpDMSConv(ByVal ConvEnable As Boolean)
     DMS_conv = ConvEnable
End Property

'-------------------------------------------------------------------------------
' get the error message
Public Property Get ErrorDescription() As String
  ErrorDescription = ErrMsg
End Property
'-------------------------------------------------------------------------------
' get the error message Id
Public Property Get ErrorID() As Long
  ErrorID = ErrId
End Property
'-------------------------------------------------------------------------------
' get the error position
Public Property Get ErrorPos() As Long
  ErrorPos = ErrPos
End Property

'-------------------------------------------------------------------------------
' evaluate expression
Public Function Eval() As Double
Dim ExprVal As Double

If Not ExprOK Then GoTo Error_Handler
ErrMsg = "": ErrId = 0

If VTtop > 0 Then
    'Check Explicit initialization
    If VarIniExp Then If VarEmpty Then GoTo Error_Handler
    SubstVars 'variables value substitution
End If
If Not Eval_(ExprVal) Then GoTo Error_Handler
Eval = ExprVal
Exit Function
'
Error_Handler:
On Error Resume Next
Err.Raise 1001, "MathParser", ErrMsg
End Function
'-------------------------------------------------------------------------------
' evaluate an expression with exactly 1 var
Public Function Eval1(ByVal x As Double) As Double
Dim i  As Long
Dim j  As Long
Dim ExprVal As Double
'
If Not ExprOK Then GoTo Error_Handler
ErrMsg = "": ErrId = 0

If VTtop > 1 Then
    ErrMsg = getMsg(1)   '"too many variables"
    ErrPos = 1
    GoTo Error_Handler
End If

For i = 1 To ETtop
  For j = 1 To ET(i).ArgTop
    If ET(i).arg(j).idx <> 0 Then ET(i).arg(j).value = ET(i).arg(j).sign * x
  Next
Next
If Not Eval_(ExprVal) Then GoTo Error_Handler
Eval1 = ExprVal
Exit Function
'
Error_Handler:
On Error Resume Next
Err.Raise 1002, "MathParser", ErrMsg
End Function
'-------------------------------------------------------------------------------
' evaluate expression passing a vector
Public Function EvalMulti(ByRef VarValue() As Double, Optional ByVal VarName)
Dim ExprVal As Double, Vout() As Double
Dim imax As Long, imin As Long, i As Long, varid As Long

If Not ExprOK Then GoTo Error_Handler
ErrMsg = "": ErrId = 0

    If VTtop > 0 Then
        'select the index variable index
        If IsMissing(VarName) Then VarName = 1
            On Error Resume Next
            If VarType(VarName) = vbString Then
                varid = VarsTbl(VarName)
            Else
                varid = VarName
            End If
            
            If Err <> 0 Then
                ErrMsg = getMsg(2) 'variable not found"
                ErrPos = 1
                GoTo Error_Handler
            End If
            
            On Error GoTo 0
            imax = UBound(VarValue)
            imin = 1
            ReDim Vout(imax)
            For i = imin To imax
                VT(varid).value = VarValue(i)
                SubstVars
                If Not Eval_(ExprVal) Then GoTo Error_Handler
                Vout(i) = ExprVal
            Next i
    Else
    
        ErrMsg = getMsg(2)  'Variable not found"
        ErrPos = 1
        GoTo Error_Handler
    End If
    
    EvalMulti = Vout

Exit Function
'
Error_Handler:
On Error Resume Next
Err.Raise 1001, "MathParser", ErrMsg
End Function
'---------------------------------------------------------------------------------
'class end routine
Private Sub Class_Terminate()
  Set VarsTbl = Nothing           'xxz6
End Sub
'-------------------------------------------------------------------------------
' Math Parse Routine
' rev 6-8-2004 Leonardo Volpi
'-------------------------------------------------------------------------------
Private Function Parse(ByVal strExpr As String) As Boolean
  Dim lExpr          As String
  Dim char           As String '*1
  Dim char0          As String
  Dim SubExpr        As String
  Dim lenExpr        As Long
  Dim FunN()         As T_Funk  'stack for N var. functions  12.06.04
  Dim GetNextArg     As Boolean
  Dim SaveArg        As String
  Dim Npart          As Long
  Dim Nabs           As Long
  Dim arrPriLvl()    As Long
  Dim flag_exchanged As Boolean
  Dim i              As Long
  Dim j              As Long
  Dim k              As Long
  Dim LogicSymb      As String
  Dim Lastarg        As Boolean
  Dim if2            As Long        'stack counter for 2 var function 16.10.03
  Dim sign           As Long        '7-1-04 saves the variable/function sign. Es "-x or -log(x)"
  Dim Node_Cond()    As Long        'store the condition-nodes
  Dim Node_Switch()  As Long        'store the switch-node
  Dim Node_max       As Long
  
  ReDim ET(HiET)
  ReDim VT(HiVT)
  ReDim FunN(HiF2)
  Set VarsTbl = New Collection
  ETtop = 0
  VTtop = 0
  Parse = False
  iInit = 0     'reset variables assignement counter
  ErrMsg = ""   'VL
  ErrId = 0
  ErrPos = 0
  lExpr = Trim(strExpr)
  '***** abs |.| function counter
  i = NabsCount(lExpr)
  Nabs = i / 2
  If (2 * Nabs <> i) Then
    ErrMsg = getMsg(4) '"abs symbols |.| mismatch"  'VL
    ErrPos = 1
    Exit Function
  End If
  '***** begin parse process
  lenExpr = Len(lExpr)
  For i = 1 To lenExpr
    ErrPos = i
    Lastarg = False
    j = 1
    char = Mid(lExpr, i, 1)
    Select Case char
      Case " "                                    '***** skip spaces
      Case "(", "[", "{"                          '***** open parentheses
        Npart = Npart + 1                         'inc # open parentheses
        If SubExpr <> "" Then                     'eval preceding text
          Catch_Sign SubExpr, sign                'catch the function sign (if any)
          If SubExpr = "" And sign = -1 Then
            SubExpr = "Neg": sign = 1             'insert change-sign function
          End If
          If InList(SubExpr, Fun1V) Then          'monovariable function
            ETtop = ETtop + 1                     '   store in ET
            With ET(ETtop)
              .PosInExpr = i                      'position in expr
              .Fun = SubExpr                      'function name
              .FunTok = GetFunTok(SubExpr)        'function Token (enum)
              .PriLvl = Npart * 10                'priority level=open parenth*10
              .ArgTop = 1                         'ntal Args=1
              .sign = sign
            End With
          Else
            'search for a function in the functions-tables
            For k = 2 To 5
                If InList(SubExpr, funk(k)) Then Exit For
            Next k
            If k > 5 Then
                'no function found
                If IsNumeric_(SubExpr) Then
                    ErrMsg = getMsg(5, i) '"Syntax error at pos: " & i
                Else
                    ErrMsg = getMsg(6, SubExpr, (i - Len(SubExpr))) '"Function < " & SubExpr & " > unknown at pos: " & (i - Len(SubExpr))
                    ErrPos = i - Len(SubExpr)
                End If
                Exit Function
            Else
                if2 = if2 + 1
                FunN(if2).FunName = SubExpr
                FunN(if2).PosInExpr = i
                ' > Mirko 20061018
                If k = 5 Then
                    'variable parameters function
                    FunN(if2).ArgN = GetNumberOfArguments(Mid(lExpr, i))
                    If FunN(if2).ArgN > HiARG Then
                        ErrMsg = getMsg(9, i)  ' "Too many arguments at pos: " & i
                        Exit Function
                    End If
                Else
                    'fixed parameters function
                    FunN(if2).ArgN = k
                End If
                ' < Mirko 20061018
                FunN(if2).ArgCount = FunN(if2).ArgN - 1
                FunN(if2).sign = sign
            End If
          End If
          SubExpr = ""                            'start parsing for new subexpr
        End If
      Case ")", "]", "}"                          '***** open parentheses
        Npart = Npart - 1                         'dec # open parentheses
        flag_exchanged = True                     'closing brackets flag
        If Npart < 0 Then                         'want to close to many brackets
          ErrMsg = getMsg(9, i) '"Too many closing brackets at pos: " & i
          Exit Function
        End If
      Case "+", "-"                               '*****
        If CheckExpo(SubExpr) Then              'fix bug 18-1-03  thanks to Michael Ruder
          SubExpr = SubExpr & char                  'continue parsing number
        Else
          If SubExpr = "" Then
            char0 = getPrevChar(i, strExpr)
            Select Case char0
            Case "", "(", "[", "{", "|", ArgSep
                SubExpr = "0"
            Case Else
                ErrMsg = getMsg(8) '"missing argument"  'preceding symbol is ")"
                Exit Function
            End Select
          End If
          ETtop = ETtop + 1                       'store in ET
          With ET(ETtop)                          '
            .PosInExpr = i
            .Fun = char
            .FunTok = GetFunTok(char)
            .PriLvl = 2 + Npart * 10
            .ArgTop = 2                           'two arguments
          End With
          GetNextArg = True                       'get second argument
          If Not Catch_Argument(SubExpr, j, Lastarg) Then Exit Function
        End If
      Case "*", "/", "\"                      '*****
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = char
          .FunTok = GetFunTok(char)
          .PriLvl = 3 + Npart * 10
          .ArgTop = 2                             'two arguments
        End With
        GetNextArg = True
        If Not Catch_Argument(SubExpr, j, Lastarg) Then Exit Function
      Case "^"
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = "^"
          .FunTok = GetFunTok(char)
          .PriLvl = 4 + Npart * 10
          .ArgTop = 2                             'two arguments
        End With
        GetNextArg = True
        If Not Catch_Argument(SubExpr, j, Lastarg) Then Exit Function
      Case "!"
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = "!"
          .FunTok = GetFunTok(char)
          .PriLvl = 9 + Npart * 10
          .ArgTop = 1                             'one argument
        End With
        SaveArg = SubExpr
        If Not Catch_Argument(SubExpr, j, Lastarg) Then Exit Function
        SubExpr = SaveArg
        flag_exchanged = True
      Case "%"  'percentage
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = char
          .FunTok = GetFunTok(char)
          .PriLvl = 9 + Npart * 10
          .ArgTop = 1                             'one argument
        End With
        GetNextArg = True
        SaveArg = SubExpr
        If Not Catch_Argument(SubExpr, j, Lastarg) Then Exit Function
        SubExpr = SaveArg
      Case ArgSep                                 'function argument separator
        If FunN(if2).FunName = "" Then    '
          ErrMsg = getMsg(9, i)  ' "Too many arguments at pos: " & i
          Exit Function
        End If
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = FunN(if2).PosInExpr
          .Fun = FunN(if2).FunName                 'previous stored
          .FunTok = GetFunTok(FunN(if2).FunName)
          .PriLvl = Npart * 10
          .ArgTop = FunN(if2).ArgN                  'N arguments
          .sign = FunN(if2).sign
        End With
        If FunN(if2).sign < 0 Then FunN(if2).sign = 1  'reset sign
        GetNextArg = True
        If Not Catch_Argument(SubExpr, j, Lastarg) Then Exit Function
        FunN(if2).ArgCount = FunN(if2).ArgCount - 1
        If FunN(if2).ArgCount = 0 Then
            FunN(if2).FunName = ""                     'reset function
            FunN(if2).ArgN = 0
            if2 = if2 - 1
        End If
      Case "|"                                    '***** absolute symbol |.|
        If SubExpr = "" Or SubExpr = "-" Then
          Npart = Npart + 1                       'increment brackets PriLvl
          ETtop = ETtop + 1
          With ET(ETtop)
            .PosInExpr = i
            .Fun = "abs"                          'symbols |.| is similar to  abs(.)
            .FunTok = GetFunTok("abs")
            .PriLvl = Npart * 10
            .ArgTop = 1                           'one argument
            If SubExpr = "-" Then                 'fix sign bug 1.3.04. Thanks to Rodrigo Farinha
                .sign = -1
                SubExpr = ""
            End If
          End With
        Else
          Npart = Npart - 1
          flag_exchanged = True                   '9.8.04 VL
          If Npart < 0 Then                       'too many closing brackets
            ErrMsg = getMsg(5, i) '"Syntax error at pos: " & i
            Exit Function
          End If
        End If
      Case "=", "<", ">"                          'Logical operators
        If LogicSymb = "" Then
            If ETtop > 0 Then
                'detect the Interval:=(a < x < b)
                If InStr(1, " > < = <= >= <> =< =>", ET(ETtop).Fun) > 0 Then
                    'transform the Interval into (a<x)*(x<b) form
                    ETtop = ETtop + 1
                    With ET(ETtop)
                      .PosInExpr = i
                      .Fun = "*"                 'insert the hidden multiplication
                      .FunTok = GetFunTok("*")
                      .PriLvl = 3 + (Npart - 1) * 10
                      .ArgTop = 2
                    End With
                    SaveArg = SubExpr
                    If Not Catch_Argument(SubExpr, j, Lastarg) Then Exit Function
                    SubExpr = SaveArg           'pass the argument also to the logic symbol
                End If
            End If
            ETtop = ETtop + 1
            GetNextArg = True
            If Not Catch_Argument(SubExpr, j, Lastarg) Then Exit Function
        End If
        
        LogicSymb = LogicSymb & char
        With ET(ETtop)
          .PosInExpr = i
          .Fun = LogicSymb                        'logic symbol
          .FunTok = GetFunTok(LogicSymb)
          .PriLvl = 1 + Npart * 10
          .ArgTop = 2                             'two argument
        End With
        
        If ET(ETtop).FunTok < 0 Then
            ErrMsg = getMsg(5, i) '"Syntax error at pos: " & i    'Fix bug "==" thanks to Ricardo Mart�nez C.
            Exit Function
        End If
      
      Case "x", "y", "z", "X", "Y", "Z"           ''monomial coeff.
        If IsNumeric_(SubExpr) Then               'fix 2.3.2003 thanks to Michael Ruder
            ETtop = ETtop + 1                     'Ex: 7x  is converted into product 7*x
            With ET(ETtop)
              .PosInExpr = i
              .Fun = "*"
              .FunTok = GetFunTok("*")
              .PriLvl = 3 + Npart * 10
              .ArgTop = 2                             'two argument
            End With
            GetNextArg = True
            If Not Catch_Argument(SubExpr, j, Lastarg) Then Exit Function
            i = i - 1  'one step back
        Else
            SubExpr = SubExpr & char
        End If
      Case Else                                   '***** continue parsing
        SubExpr = SubExpr & char
    End Select
    
    If ETtop > UBound(ET) - 2 Then
        ErrMsg = getMsg(20) '"Too many operations"
        ErrPos = 1
        Exit Function
    End If
    
    If flag_exchanged = True Then                 'after closing parenthesis
        If IsLetter(char) Or IsDigit(char) Or char = DecSep Then   'these symbols are not allowed.
            ErrMsg = getMsg(5, i) '"Syntax error at pos: " & i        'Fix bug thanks to PJ Weng.
            Exit Function
        End If
    End If
    
    If GetNextArg Then
        GetNextArg = False
        flag_exchanged = False
    Else
        LogicSymb = ""
    End If
  Next
'---end of the main loop -----------------------
  
  If Npart > 0 Then                               'parentheses
    ErrMsg = getMsg(11) '"Not enough closing brackets"
    ErrPos = 1
    Exit Function
  End If
  If ETtop < 1 Then                               'no operation detected
    ETtop = 1
    With ET(ETtop)
      .PosInExpr = 1
      .Fun = "+"
      .FunTok = GetFunTok("+")
      .PriLvl = 1
      .ArgTop = 2
    End With
  End If
  For i = 1 To ETtop                              'init 2e argument
    j = ET(i).ArgTop
    If j > 2 Then j = 2
    ET(i).arg(j) = ET(i + 1).arg(1)
  Next
  
  If SubExpr <> "" Then                           'catch last argument or Vars
    j = ET(ETtop).ArgTop
    If j > 2 Then j = 2
    Lastarg = True
    If Not Catch_Argument(SubExpr, j, Lastarg) Then Exit Function
  Else
    'bug 7.10.03 last argument missing 3+ or sin() ...  thanks to Rodigro Farinha
    ErrMsg = getMsg(8) '"missing argument"
    Exit Function
  End If
  
  If if2 > 0 Then     '16.10.03
    ErrMsg = getMsg(8) '"missing argument"
    Exit Function
  End If
  
  ReDim Preserve ET(ETtop)
  ReDim Preserve VT(VTtop)
  
  Call Sort_table                'start sort algorithm
  
  Call Build_Relations           'build relations
  
'-------------------------------------------------------------------------------
 'Search for the condition nodes in the graph appling the logic condition rules
  Control_Nodes Node_Cond, Node_Switch, Node_max
  If Node_max > 0 Then Call Sort_table 'start sort algorithm
'------------------------------------------------------------------------------

  Call Multi_Variables_Function  'management of multi-variable functions
 
  Call Arguments_Cleanup         'eliminate dependent arguments

  Parse = True
  
End Function
'---------------------------------------------------------------------------------
Private Function Catch_Argument(ByRef SubExpr As String, ByVal j As Long, ByVal Lastarg As Boolean) As Boolean
Dim sign As Long, retval As Double
    Catch_Argument = False
    If SubExpr = "" Then                        'no next argument found
        ErrMsg = getMsg(8) '"missing argument"
        Exit Function
    End If
    Catch_Sign SubExpr, sign                    'breack the string into name and its sign +/-(if any)
    If SubExpr = "" Then                        'mod 5.4.2004 VL
        ErrMsg = getMsg(8) '"missing argument"   'fix bug for ++ or -- string. Thanks to PJ Weng
        Exit Function
    ElseIf convSymbConst(SubExpr, retval) Then      'check if argument is a symbolic constant #
      If ErrMsg <> "" Then Exit Function
      ET(ETtop).arg(j).value = sign * retval
    ElseIf convEGU(SubExpr, retval) Then        'check if argument is Eng Units
      ET(ETtop).arg(j).value = sign * retval
    ElseIf IsNumeric_(SubExpr) Then             'check if argument is number
      If DP_SET Then
        ET(ETtop).arg(j).value = sign * Val(SubExpr)
      Else
        ET(ETtop).arg(j).value = sign * CDbl(SubExpr)
      End If
    ElseIf cvDegree(SubExpr, retval, ErrMsg) Then      'check if argument is ddmmss format degree
      If ErrMsg <> "" Then Exit Function
      ET(ETtop).arg(j).value = sign * retval    'angle sign bug 31.5.2004. thanks to PJ Weng
    Else
    
      If Not IsVariable(SubExpr) Then
        ErrMsg = getMsg(12, SubExpr) ' "Syntax error: " & SubExpr
        Exit Function
      End If
    
      StoreVar SubExpr, Lastarg, sign
      If VTtop > HiVT Then
        ErrMsg = getMsg(1)  '"too many variables"
        Exit Function
      End If
    End If
    ErrMsg = ""  'clear error message. fix bug 23.7.04. thanks to Ricardo Mart�nez C.
    ErrId = 0
    SubExpr = ""  'reset the substring
    Catch_Argument = True
End Function
'---------------------------------------------------------------------------------
Private Sub Sort_table()
'sort table with exchanges algorithm
Dim i As Long, j As Long, srtLo As Long, srtHi As Long, tmp As Long, flag_exchanged As Boolean
  ReDim arrPriLvl(ETtop)                        'create array with priority levels
  For i = 1 To ETtop                              'and copy then from main array
    arrPriLvl(i) = ET(i).PriLvl
  Next
  For i = 1 To ETtop                              'fill sort order default 0 to ETtop
    ET(i).PriIdx = i
  Next
  srtLo = 1                                       '***** start sort algorithm
  srtHi = ETtop - 1
  Do
    flag_exchanged = False
    For i = srtLo To srtHi Step 2
      j = i + 1
      If arrPriLvl(i) < arrPriLvl(j) Then
        tmp = arrPriLvl(j)
        arrPriLvl(j) = arrPriLvl(i)
        arrPriLvl(i) = tmp
        tmp = ET(j).PriIdx
        ET(j).PriIdx = ET(i).PriIdx
        ET(i).PriIdx = tmp
        flag_exchanged = True
      End If
    Next
    If srtLo = 1 Then
      srtLo = 2
    Else
      srtLo = 1
    End If
  Loop Until (srtLo = 1) And Not flag_exchanged
End Sub
'---------------------------------------------------------------------------------
Private Sub Build_Relations()
Dim i As Long, j As Long, srtLo As Long, srtHi As Long
  For i = 1 To ETtop                              'build relations
    j = ET(i).PriIdx
    srtLo = j - 1
    Do While srtLo >= 0
      If ET(srtLo).ArgOf = 0 Then
        Exit Do
      End If
      srtLo = srtLo - 1
    Loop
    srtHi = j + 1
    Do While srtHi <= ETtop
      If ET(srtHi).ArgOf = 0 Then
        Exit Do
      End If
      srtHi = srtHi + 1
    Loop
    If (srtLo < 1) And (srtHi <= ETtop) Then            '
      ET(j).ArgOf = srtHi
      ET(j).ArgIdx = 1
    ElseIf (srtLo > 0) And (srtHi > ETtop) Then        '
      ET(j).ArgOf = srtLo
      ET(j).ArgIdx = ET(srtLo).ArgTop
      If ET(j).ArgIdx > 2 Then ET(j).ArgIdx = 2  'clipp
    ElseIf (srtLo > 0) And (srtHi <= ETtop) Then       '
      If (ET(srtLo).PriLvl) >= (ET(srtHi).PriLvl) Then  'take that one with the upper PriLvl
        ET(j).ArgOf = srtLo
        ET(j).ArgIdx = ET(srtLo).ArgTop
        If ET(j).ArgIdx > 2 Then ET(j).ArgIdx = 2
      Else                                        '
        ET(j).ArgOf = srtHi
        ET(j).ArgIdx = 1
      End If
    Else
      Exit For
    End If
  Next

End Sub
'---------------------------------------------------------------------------------
Private Sub Multi_Variables_Function()
Dim i As Long, j As Long, p As Long, k As Long
    For i = 1 To ETtop
        If ET(i).ArgTop > 2 Then
            'change the intermediate multi-function token
            p = 3
            Do
                j = ET(i).ArgOf
                ET(i).ArgOf = ET(j).ArgOf
                ET(i).ArgIdx = ET(j).ArgIdx 'Dipti DDWE-2709
                With ET(j)
                    .ArgOf = i
                    .ArgIdx = p
                    .Fun = "@Right"
                    .FunTok = GetFunTok("@Right")
                    .ArgTop = 2
                End With
                p = p + 1
            Loop Until p > ET(i).ArgTop
            'change the priority index
            For k = 1 To ETtop
                If ET(k).PriIdx = i Then
                    ET(k).PriIdx = j
                ElseIf ET(k).PriIdx = j Then
                    ET(k).PriIdx = i
                End If
            Next
        End If
    Next i
End Sub
'---------------------------------------------------------------------------------
Private Sub Arguments_Cleanup()
Dim i As Long, j As Long
  For i = 1 To ETtop
    j = ET(i).ArgOf
    If j > 0 Then
      With ET(j).arg(ET(i).ArgIdx)
       .idx = 0
       .Nome = ""
      End With
    End If
  Next i
End Sub
'---------------------------------------------------------------------------------
Private Function CheckExpo(ByVal SubExpr As String) As Boolean
Dim s_1 As String, s_2 As String, ls As Long
'detect if SubExpr is the mantissa of an expo format number 1.2E+01 , 4E-12, 1.0E-6
    CheckExpo = False
    ls = Len(SubExpr)
    If ls < 2 Then Exit Function
    s_1 = Right(SubExpr, 1)
    s_2 = Left(SubExpr, ls - 1)
    If (UCase(s_1) = "E") And IsNumeric_(s_2) Then CheckExpo = True
End Function

Private Function getPrevChar(ByVal i As Long, ByVal Str As String) As String
If i <= 1 Then
    getPrevChar = ""
Else
    getPrevChar = Right(Trim(Left(Str, i - 1)), 1)
End If
End Function

Private Function CheckSign(ByVal i As Long, ByVal Str As String) As Boolean
Dim s_1 As String, s_2 As String
CheckSign = True
If i <= 1 Then Exit Function
s_1 = Trim(Left(Str, i - 1))
s_2 = Right(s_1, 1)
If s_2 = ")" Or s_2 = "]" Or s_2 = "}" Then CheckSign = False
If s_2 = "(" Or s_2 = "[" Or s_2 = "{" Then CheckSign = True
End Function
'-------------------------------------------------------------------------------
'Search for the condition nodes in the graph appling the logic condition rules
Private Sub Control_Nodes(ByRef Node_Cond() As Long, ByRef Node_Switch() As Long, ByRef Node_max As Long)
'mod. 30-6-2004
Dim n&, i&, j&, k&, p&, count_iter&, Node_dup As Boolean
Dim Node_aux() As Long, Node_aux1() As Long
Dim Ns&, nc&, Level_max&

n = UBound(ET)
ReDim Node_Cond(n), Node_Switch(n), Node_aux(n)
j = 0
For i = 1 To n
    If InStr(1, " > < = <= >= <> =< =>", ET(i).Fun) > 0 Then
        Node_aux(i) = 2
        'search for duplicate nodes
        Node_dup = False
        For p = 1 To j
            If Node_Switch(p) = ET(i).ArgOf Then
                Node_dup = True
                Exit For
            End If
        Next
        'apply rule
        If Node_dup Then
            Node_Cond(p) = Node_Switch(p)
            Node_Switch(p) = ET(Node_Cond(p)).ArgOf
            Node_aux(Node_Cond(p)) = 2
        Else
            j = j + 1
            Node_Cond(j) = i
            Node_Switch(j) = ET(i).ArgOf
        End If
    End If
    If Level_max < ET(i).PriLvl Then Level_max = ET(i).PriLvl
Next i
Node_max = j
If Node_max = 0 Then Exit Sub 'no logic function detected
ReDim Preserve Node_Cond(Node_max), Node_Switch(Node_max)

For p = 1 To Node_max
    'load control node information
    nc = Node_Cond(p)
    Ns = Node_Switch(p)
    If ET(Ns).Fun = "*" Then
        ReDim Node_aux1(n)
        For i = 1 To n: Node_aux1(i) = Node_aux(i): Next i
        'apply the graph-condition-rules to the table
        count_iter = 0
        For k = 1 To n
            i = k: j = 0
            Do
                If Node_aux1(i) = -1 Or ET(i).ArgOf = 0 Then
                    'node k independent
                    Node_aux1(k) = -1
                    Exit Do
                End If
                If Node_aux1(i) = 2 Then
                    'assign higher priority to the condition node and its childs
                    ET(k).PriLvl = ET(k).PriLvl + 100
                    Exit Do
                End If
                If Node_aux1(i) = 1 Or ET(i).ArgOf = Ns Then
                    'node k dependent
                    Node_aux1(k) = 1
                    ET(k).Cond = nc
                    Exit Do
                End If
                i = ET(i).ArgOf
                j = j + 1
            Loop Until j > n
            count_iter = count_iter + j
        Next k
    End If
Next p

End Sub
'-------------------------------------------------------------------------------
'[modified 10/02 by Thomas Zeutschler]
Private Function Eval_(ByRef EvalValue As Double) As Boolean
  Dim A     As Double
  Dim B     As Double
  Dim Ris   As Double
  Dim j     As Long
  Dim k     As Long
  Dim Pos   As Long
  Dim m     As Long
  Dim n     As Long
  Dim sa    As String
  Dim sb    As String
  
  On Error GoTo ErrHandler  '<<< comment for debug  VL 30-8-02
  For j = 1 To ETtop    'Evaluation procedure begins
    k = ET(j).PriIdx
    Ris = 0
    With ET(k)
      '--  apply conditioning rule --------------
      If .Cond <> 0 Then If ET(.Cond).value = 0 Then GoTo ResultHandler
      '------------------------------
      A = .arg(1).value
      B = .arg(2).value
      Select Case .FunTok
        Case symPlus:            Ris = A + B
        Case symMinus:           Ris = A - B
        Case symMul:             Ris = A * B
        Case symDiv:             Ris = A / B
        Case symPercent:         Ris = A / 100
        Case symDivInt:          Ris = A \ B
        Case symPov:             Ris = A ^ B: If A = B And A = 0 Then GoTo ErrHandler '26.7.04 Ricardo Mart�nez C.
        Case symNeg:             Ris = -A
        Case symAbs:             Ris = Abs(A)
        Case symAtn:             Ris = Atn(A) / CvAngleCoeff
        Case symCos:             Ris = Cos(CvAngleCoeff * A)
        Case symSin:             Ris = Sin(CvAngleCoeff * A)
        Case symExp:             Ris = Exp(A)
        Case symFix:             Ris = Fix(A)
        Case symInt:             Ris = Int(A)
        Case symDec:             Ris = Dec_(A)
        Case symLn:              Ris = Log(A)
        Case symLog:             Ris = Log(A)    'the same as natural logarithm mod. 15.6.2004
        Case symLogN:            Ris = Log(A) / Log(B)
        Case symRnd:             Ris = A * Rnd(1)
        Case symSgn:             Ris = Sgn(A)
        Case symSqr:             Ris = Sqr(A)
        Case symCbr:             Ris = Sgn(A) * Abs(A) ^ (1 / 3)
        Case symTan:             Ris = Tan(CvAngleCoeff * A)
        Case symAcos:            Ris = Acos_(A) / CvAngleCoeff
        Case symAsin:            Ris = Asin_(A) / CvAngleCoeff
        Case symCosh:            Ris = Cosh_(A)
        Case symSinh:            Ris = Sinh_(A)
        Case symTanh:            Ris = Tanh_(A)
        Case symAcosh:           Ris = Acosh_(A)   '7.10.2003 fix bug (thank to Rodrigo Farinha)
        Case symAsinh:           Ris = Asinh_(A)
        Case symAtanh:           Ris = Atanh_(A)
        Case symRoot:            Ris = MiRoot_(A, B)
        Case symmod:             Ris = Mod_(A, B)   'a Mod b  17.10.03 fix VBA bug
        Case symFact:            Ris = fact(A)
        Case symComb:            Ris = Comb(A, B)
        Case symPerm:            Ris = Perm(A, B)
        Case symGT:              Ris = -CDbl((A > B))
        Case symGE:              Ris = -CDbl(A >= B)
        Case symLT:              Ris = -CDbl(A < B)
        Case symLE:              Ris = -CDbl(A <= B)
        Case symEQ:              Ris = -CDbl(A = B)
        Case symNE:              Ris = -CDbl(A <> B)
        Case symAnd:             Ris = -CDbl((A <> 0) And (B <> 0))
        Case symOr:              Ris = -CDbl((A <> 0) Or (B <> 0))
        Case symNot:             Ris = -CDbl(A = 0)
        Case symXor:             Ris = -CDbl((A <> 0) Xor (B <> 0)) ' MR 16-01-03 XOR corrected
        Case symNAnd:            Ris = -CDbl((A = 0) Or (B = 0))     '
        Case symNOr:             Ris = -CDbl((A = 0) And (B = 0))    '
        Case symNXor:            Ris = -CDbl((A <> 0) = (B <> 0))    'MR 16-01-03 NXor        '
        Case symErf:             Ris = erf_(A)
        Case symGamma:           Ris = gamma_(A)
        Case symGammaln:         Ris = gammaln_(A)
        Case symDigamma:         Ris = digamma_(A)
        Case symGammaI:          Ris = GammaInc(A, B)
        Case symBeta:            Ris = beta_(A, B)
        Case symZeta:            Ris = Zeta_(A)
        Case symEi:              Ris = exp_integr(A)
        Case symCsc:             Ris = 1 / Sin(CvAngleCoeff * A)
        Case symSec:             Ris = 1 / Cos(CvAngleCoeff * A)
        Case symCot:             Ris = 1 / Tan(CvAngleCoeff * A)
        Case symACsc:            Ris = Asin_(1 / A) / CvAngleCoeff
        Case symASec:            Ris = Acos_(1 / A) / CvAngleCoeff
        Case symACot:            Ris = PI_ / 2 - Atn(A) / CvAngleCoeff
        Case symCsch:            Ris = 1 / Sinh_(A)
        Case symSech:            Ris = 1 / Cosh_(A)
        Case symCoth:            Ris = 1 / Tanh_(A)
        Case symACsch:           Ris = Asinh_(1 / A)
        Case symASech:           Ris = Acosh_(1 / A)
        Case symACoth:           Ris = Atanh_(1 / A)
        Case symRad:             Ris = A / CvAngleCoeff
        Case symDeg:             Ris = A / CvAngleCoeff * PI_ / 180
        Case symGrad:            Ris = A / CvAngleCoeff * PI_ / 200
        Case symRound:           Ris = round_(A, B)
        Case symRight:           Ris = B
        Case symDPoiss:          Ris = DPoisson(A, B)
        Case symCPoiss:          Ris = CPoisson(A, B)
        Case symEin:             Ris = expn_integr(A, B)
        Case symSi:              Ris = SinIntegral(A)
        Case symCi:              Ris = CosIntegral(A)
        Case symFresS:           Ris = Fresnel_sin(A)
        Case symFresC:           Ris = Fresnel_cos(A)
        Case symBessJ:           Ris = BesselJ(A, B)
        Case symBessY:           Ris = BesselY(A, B)
        Case symBessK:           Ris = BesselK(A, B)
        Case symBessI:           Ris = BesselI(A, B)
        Case symJ0:              Ris = BesselJ(A, 0)
        Case symY0:              Ris = BesselY(A, 0)
        Case symI0:              Ris = BesselI(A, 0)
        Case symK0:              Ris = BesselK(A, 0)
        Case symPolyLe:          Ris = Poly_Legendre(A, B)
        Case symPolyLa:          Ris = Poly_Laguerre(A, B)
        Case symPolyHe:          Ris = Poly_Hermite(A, B)
        Case symPolyCh:          Ris = Poly_Chebycev(A, B)
        Case symAiryA:           Ris = Airy_A(A)
        Case symAiryB:           Ris = Airy_B(A)
        Case symEllipt1:         Ris = IElliptic_Int1(A, CvAngleCoeff * B)
        Case symEllipt2:         Ris = IElliptic_Int2(A, CvAngleCoeff * B)
        Case symWtri:            Ris = WAVE_TRI(A, B)
        Case symWsqr:            Ris = WAVE_SQR(A, B)
        Case symWsaw:            Ris = WAVE_SAW(A, B)
        Case symWraise:          Ris = WAVE_RAISE(A, B)
        Case symWparab:          Ris = WAVE_PARAB(A, B)
        Case symStep:            Ris = Step_(A, B)
        ' > Berend 20041216
        Case symYear:            Ris = Year(CDate(A))
        Case symMonth:           Ris = Month(CDate(A))
        Case symDay:             Ris = Day(CDate(A))
        Case symHour:            Ris = Hour(CDate(A))
        Case symMinute:          Ris = Minute(CDate(A))
        Case symSecond:          Ris = Second(CDate(A))
        ' < Berend 20041216
        Case Is > HFOffset
            Ris = EvalMulti_(k)  'multi-variable function
        Case Else
          ErrMsg = getMsg(13, .FunTok) '"Function <" & .FunTok & "> missing?"
          Exit Function
      End Select
      If .sign = -1 Then Ris = -Ris  'change function sign (7-1-04)
ResultHandler:
      .value = Ris
      m = .ArgOf
      n = .ArgIdx
      If m = 0 Or n = 0 Then Exit For
      ET(m).arg(n).value = Ris
    End With
  Next j
  EvalValue = ET(k).value
  Eval_ = True
  Exit Function
ErrHandler:
  ErrPos = ET(k).PosInExpr
  If ET(k).FunTok < HFOffset Then
    'build the error msg for functions having 1 or 2 fixed arguments
    sa = A:    sb = B
    If DP_SET Then  'force the decimal point
        sa = Replace(A, ",", ".")
        sb = Replace(B, ",", ".")
    End If
    If ET(k).ArgTop = 1 Then
         ErrMsg = getMsg(14, ET(k).Fun, sa, ErrPos)
    ElseIf ET(k).ArgTop = 2 Then
        If InList(ET(k).Fun, Fun2V) Then
           ErrMsg = getMsg(15, ET(k).Fun, sa, sb, ErrPos)
        Else
          ErrMsg = getMsg(16, sa, ET(k).Fun, sb, ErrPos)
        End If
    End If
  Else
    'build the error msg for functions having multi-arguments
        For j = 1 To ET(k).ArgTop
          If j > 1 Then ErrMsg = ErrMsg & ArgSep
          sa = ET(k).arg(j).value
          If DP_SET Then sa = Replace(sa, ",", ".")
          ErrMsg = ErrMsg & sa
        Next j
        ErrMsg = getMsg(17, ET(k).Fun, ErrMsg, ErrPos)
  End If
  EvalValue = 0
  Eval_ = False
End Function

Private Function EvalMulti_(ByVal k As Long) As Double
'evaluate multi-variable functions.
'mod. 20.10.2006 - added variable arguments functions. (Thanks to Mirko Sartori)
Dim Ris As Double
Dim i As Integer
Dim x() As Double

With ET(k)
If .FunTok < (HFOffset + 100) Then
    'section of fixed arguments functions
    Select Case .FunTok
        Case symDnorm
            Ris = DNormal(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symCnorm
            Ris = CNormal(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symDBinom
            Ris = DBinomial(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symCBinom
            Ris = CBinomial(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symHypGeo
            Ris = Hypergeom(.arg(1).value, .arg(2).value, .arg(3).value, .arg(4).value)
        Case symBetaI
            Ris = BetaInc(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symClip
            Ris = Clip(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symWrect
            Ris = WAVE_RECT(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symWtrapez
            Ris = WAVE_TRAPEZ(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symWlin
            Ris = WAVE_LIN(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symWpulse
            Ris = WAVE_PULSE(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symWsteps
            Ris = WAVE_STEPS(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symWexp
            Ris = WAVE_EXP(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symWexpb
            Ris = WAVE_EXPB(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symWpulsef
            Ris = WAVE_PULSEF(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symWripple
            Ris = WAVE_RIPPLE(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symWring
            Ris = WAVE_RING(.arg(1).value, .arg(2).value, .arg(3).value, .arg(4).value)
        Case symWam
            Ris = WAVE_AM(.arg(1).value, .arg(2).value, .arg(3).value, .arg(4).value)
        Case symWfm
            Ris = WAVE_FM(.arg(1).value, .arg(2).value, .arg(3).value, .arg(4).value)
        ' > Berend 20041216
        Case symDateSerial
            Ris = DateSerial(.arg(1).value, .arg(2).value, .arg(3).value)
        Case symTimeSerial
            Ris = TimeSerial(.arg(1).value, .arg(2).value, .arg(3).value)
        ' < Berend 20041216
         Case Else
          ErrMsg = getMsg(13, .FunTok) '"Function <" & .FunTok & "> missing?"  'VL
          Ris = "" 'rise an error
    End Select
   
 Else
    'section of variable arguments functions
    ReDim x(1 To .ArgTop)
    For i = 1 To .ArgTop: x(i) = .arg(i).value: Next i
    Select Case .FunTok
        Case symMin
            Ris = min_n(x)
        Case symMax
            Ris = max_n(x)
        Case symSum
            Ris = sum_(x)
        Case symMean
            Ris = mean_(x)
        Case symMeanq
            Ris = meanq_(x)
        Case symMeang
            Ris = meang_(x)
        Case symVar
            Ris = var_(x)
        Case symVarp
            Ris = varp_(x)
        Case symStdev
            Ris = stdev_(x)
        Case symStdevp
            Ris = stdevp_(x)
        Case symMcd
            Ris = MCD_(x)
        Case symMcm
            Ris = MCM_(x)
        Case Else
          ErrMsg = getMsg(13, .FunTok) '"Function <" & .FunTok & "> missing?"  'VL
          Ris = "" 'rise an error
    End Select
End If
End With
EvalMulti_ = Ris
End Function
'-------------------------------------------------------------------------------
' Check if all variable has been assigned
Private Function VarEmpty() As Boolean
Dim i As Long
    If VTtop = iInit Then
        VarEmpty = False
    Else
        VarEmpty = True
        For i = 1 To VTtop
            If VT(i).Init = False Then Exit For
        Next i
        ErrMsg = getMsg(18, VT(i).Nome)   '"Variable < " & VT(i).Nome & " > not assigned"
        ErrPos = 1
    End If
End Function
'-------------------------------------------------------------------------------
' Assignes a value to symbolic Vars
Private Sub SubstVars()
Dim i  As Long, j  As Long

  For i = 1 To ETtop
    For j = 1 To ET(i).ArgTop
    With ET(i).arg(j)
        If .idx <> 0 Then .value = .sign * VT(.idx).value
    End With
    Next
  Next
End Sub
'-------------------------------------------------------------------------------
' search if var already exists in table, if not add it
Private Sub StoreVar(ByVal SubExpr As String, ByVal Lastarg As Boolean, ByVal sign As Long)
  Dim VTIdx  As Long
  Dim ArgIdx As Long
  Dim Found  As Boolean
    
  Found = False
  For VTIdx = 1 To VTtop
    'If VT(VTIdx).Nome = SubExpr Then
    If LCase$(VT(VTIdx).Nome) = LCase$(SubExpr) Then  '20.12.2004 fix bug Variable Uppercase. thanks to Andr� Hendriks
      Found = True
      Exit For
    End If
  Next
  If Not Found Then
    VTtop = VTtop + 1     'new variable
    If VTtop > HiVT Then  'to many Vars
      Exit Sub
    End If
    VT(VTtop).Nome = SubExpr
    'add a new variable to the object collection (bb 6-1-04)
    VarsTbl.Add VTtop, SubExpr
  End If
  If Lastarg Then
    ArgIdx = ET(ETtop).ArgTop
    If ArgIdx > 2 Then ArgIdx = 2
  Else
    ArgIdx = 1
  End If
  With ET(ETtop).arg(ArgIdx)
    .Nome = SubExpr
    .idx = VTIdx
    .sign = sign
  End With
End Sub
'-------------------------------------------------------------------------------
' get function token '[added 10/02 by Thomas Zeutschler]
Private Function GetFunTok(ByVal FunTok As String) As Long
  Select Case LCase(FunTok)
    Case "+":       GetFunTok = symPlus
    Case "-":       GetFunTok = symMinus
    Case "*":       GetFunTok = symMul
    Case "/":       GetFunTok = symDiv
    Case "%":       GetFunTok = symPercent
    Case "\":       GetFunTok = symDivInt
    Case "^":       GetFunTok = symPov
    Case "neg":     GetFunTok = symNeg
    Case "abs":     GetFunTok = symAbs
    Case "atn":     GetFunTok = symAtn
    Case "atan":    GetFunTok = symAtn   '30.3.04 VL
    Case "cos":     GetFunTok = symCos
    Case "sin":     GetFunTok = symSin
    Case "exp":     GetFunTok = symExp
    Case "fix":     GetFunTok = symFix
    Case "int":     GetFunTok = symInt
    Case "dec":     GetFunTok = symDec
    Case "ln":      GetFunTok = symLn
    Case "log":     GetFunTok = symLog
    Case "logn":    GetFunTok = symLogN
    Case "rnd":     GetFunTok = symRnd
    Case "sgn":     GetFunTok = symSgn
    Case "sqr":     GetFunTok = symSqr
    Case "cbr":     GetFunTok = symCbr
    Case "tan":     GetFunTok = symTan
    Case "acos":    GetFunTok = symAcos
    Case "asin":    GetFunTok = symAsin
    Case "cosh":    GetFunTok = symCosh
    Case "sinh":    GetFunTok = symSinh
    Case "tanh":    GetFunTok = symTanh
    Case "acosh":   GetFunTok = symAcosh
    Case "asinh":   GetFunTok = symAsinh
    Case "atanh":   GetFunTok = symAtanh
    Case "root":    GetFunTok = symRoot
    Case "mod":     GetFunTok = symmod
    Case "fact", "!": GetFunTok = symFact
    Case "comb":    GetFunTok = symComb
    Case "perm":    GetFunTok = symPerm
    Case "min":     GetFunTok = symMin
    Case "max":     GetFunTok = symMax
    Case "mcd", "gcd":    GetFunTok = symMcd
    Case "mcm", "lcm":    GetFunTok = symMcm
    Case ">":       GetFunTok = symGT
    Case ">=", "=>": GetFunTok = symGE
    Case "<":       GetFunTok = symLT
    Case "<=", "=<": GetFunTok = symLE
    Case "=":       GetFunTok = symEQ
    Case "<>":      GetFunTok = symNE
    Case "and":     GetFunTok = symAnd
    Case "or":      GetFunTok = symOr
    Case "not":     GetFunTok = symNot
    Case "xor":     GetFunTok = symXor
    Case "nand":    GetFunTok = symNAnd
    Case "nor":     GetFunTok = symNOr
    Case "nxor":    GetFunTok = symNXor
    Case "erf":     GetFunTok = symErf
    Case "gamma":   GetFunTok = symGamma
    Case "gammaln": GetFunTok = symGammaln
    Case "digamma": GetFunTok = symDigamma
    Case "gammai":  GetFunTok = symGammaI
    Case "beta":    GetFunTok = symBeta
    Case "zeta":    GetFunTok = symZeta
    Case "ei":      GetFunTok = symEi
    Case "ein":     GetFunTok = symEin
    Case "csc":     GetFunTok = symCsc
    Case "sec":     GetFunTok = symSec
    Case "cot":     GetFunTok = symCot
    Case "acsc":    GetFunTok = symACsc
    Case "asec":    GetFunTok = symASec
    Case "acot":    GetFunTok = symACot
    Case "csch":    GetFunTok = symCsch
    Case "sech":    GetFunTok = symSech
    Case "coth":    GetFunTok = symCoth
    Case "acsch":   GetFunTok = symACsch
    Case "asech":   GetFunTok = symASech
    Case "acoth":   GetFunTok = symACoth
    Case "rad":     GetFunTok = symRad
    Case "deg":     GetFunTok = symDeg
    Case "grad":    GetFunTok = symGrad
    Case "round":   GetFunTok = symRound
    Case "dnorm":   GetFunTok = symDnorm
    Case "cnorm":   GetFunTok = symCnorm
    Case "dbinom":   GetFunTok = symDBinom
    Case "cbinom":   GetFunTok = symCBinom
    Case "dpoisson": GetFunTok = symDPoiss
    Case "cpoisson": GetFunTok = symCPoiss
    Case "si":       GetFunTok = symSi
    Case "ci":       GetFunTok = symCi
    Case "psi":      GetFunTok = symDigamma
    Case "fresnels": GetFunTok = symFresS
    Case "fresnelc": GetFunTok = symFresC
    Case "besselj":  GetFunTok = symBessJ
    Case "besseli":  GetFunTok = symBessI
    Case "besselk":  GetFunTok = symBessK
    Case "bessely":  GetFunTok = symBessY
    Case "j0":       GetFunTok = symJ0
    Case "y0":       GetFunTok = symY0
    Case "i0":       GetFunTok = symI0
    Case "k0":       GetFunTok = symK0
    Case "hypgeom":  GetFunTok = symHypGeo
    Case "betai":    GetFunTok = symBetaI
    Case "polyle":   GetFunTok = symPolyLe
    Case "polyhe":   GetFunTok = symPolyHe
    Case "polyla":   GetFunTok = symPolyLa
    Case "polych":   GetFunTok = symPolyCh
    Case "airya":    GetFunTok = symAiryA
    Case "airyb":    GetFunTok = symAiryB
    Case "elli1":    GetFunTok = symEllipt1
    Case "elli2":    GetFunTok = symEllipt2
    Case "clip":     GetFunTok = symClip
    Case "wtri":     GetFunTok = symWtri
    Case "wsqr":     GetFunTok = symWsqr
    Case "wsaw":     GetFunTok = symWsaw
    Case "wraise":   GetFunTok = symWraise
    Case "wparab":   GetFunTok = symWparab
    Case "wrect":    GetFunTok = symWrect
    Case "wtrapez":  GetFunTok = symWtrapez
    Case "wlin":     GetFunTok = symWlin
    Case "wpulse":   GetFunTok = symWpulse
    Case "wsteps":   GetFunTok = symWsteps
    Case "wexp":     GetFunTok = symWexp
    Case "wexpb":    GetFunTok = symWexpb
    Case "wpulsef":  GetFunTok = symWpulsef
    Case "wripple":  GetFunTok = symWripple
    Case "wring":    GetFunTok = symWring
    Case "wam":      GetFunTok = symWam
    Case "wfm":      GetFunTok = symWfm
    Case "@right":   GetFunTok = symRight
    ' > Berend 20041216
    Case "year":     GetFunTok = symYear
    Case "month":    GetFunTok = symMonth
    Case "day":      GetFunTok = symDay
    Case "hour":     GetFunTok = symHour
    Case "minute":   GetFunTok = symMinute
    Case "second":   GetFunTok = symSecond
    Case "dateserial": GetFunTok = symDateSerial
    Case "timeserial": GetFunTok = symTimeSerial
    ' < Berend 20041216
    Case "mean":    GetFunTok = symMean
    Case "sum":     GetFunTok = symSum
    Case "meanq":   GetFunTok = symMeanq
    Case "meang":   GetFunTok = symMeang
    Case "var":     GetFunTok = symVar
    Case "varp":    GetFunTok = symVarp
    Case "stdev":   GetFunTok = symStdev
    Case "stdevp":  GetFunTok = symStdevp
    Case "step":    GetFunTok = symStep
    Case Else
        GetFunTok = symARGUMENT
    End Select
End Function
'-------------------------------------------------------------------------------
' translate egu to multiplication factor
'  accepts a string contains a measure like "2ms" ,"234.12Mhz", "0.1uF" , 12Km , etc
'  [relaxed parsing: allow space between number and unit and allow numbers without units]
Private Function convEGU(ByVal strSource As String, ByRef retval As Double) As Boolean
  Dim EguStr   As String
  Dim EguChar  As String
  Dim EguStart As Long
  Dim EguLen   As Long
  Dim EguMult  As String
  Dim EguCoeff As Double
  Dim EguFact  As Long
  Dim EguSym   As String
  Dim EguBase  As Double
    
  'check flag unit conversion  23-2-05
  If Unit_conv = False Then
    convEGU = False
    Exit Function
  End If
  
  EguStr = strSource      'trim niet nodig. alle spaties zijn weg
  EguLen = Len(EguStr)
  For EguStart = 1 To EguLen
    EguChar = Mid(EguStr, EguStart, 1)  'fix Expo number bug. 25.1.03 VL
    If Not IsNumeric_(EguChar) Then
        Select Case EguChar
            Case "+", "-", DecSep
                'continue
            Case "E", "e"
                EguChar = Mid(EguStr, EguStart + 1, 1) 'check next char
                If Not (EguChar = "+" Or EguChar = "-" Or IsNumeric_(EguChar)) Then Exit For
            Case Else
                If IsLetter(EguChar) Then
                    Exit For
                Else
                    convEGU = False: Exit Function
                End If
        End Select
    End If
  Next
  '
  If EguStart = 1 Or EguStart > EguLen Then
    convEGU = False
    Exit Function
  End If
  '
  If DP_SET Then
    EguCoeff = Val(Left(EguStr, EguStart - 1))   'get number
  Else
    EguCoeff = CDbl(Left(EguStr, EguStart - 1))  '23.10.06
  End If
  EguStr = Mid(EguStr, EguStart)          'extract literal substring
  EguLen = Len(EguStr)
  If EguLen > 1 Then                      'extract multiply factor
    EguMult = Left(EguStr, 1)
    Select Case EguMult
      Case "p":  EguFact = -12
      Case "n":  EguFact = -9
      Case "u":  EguFact = -6      '
    #If CODPAGE = 0 Then
      Case "�":  EguFact = -6      '30.3.04 VL (comment this line for chinese/japanese VB version)
    #End If
      Case "m":  EguFact = -3
      Case "c":  EguFact = -2
      Case "k":  EguFact = 3       '14.2.03 VL
      Case "M":  EguFact = 6
      Case "G":  EguFact = 9
      Case "T":  EguFact = 12
      Case Else: EguFact = 0
    End Select
  Else
    EguFact = 0
  End If
  
  If EguFact <> 0 Then       'extract um symbol
    EguSym = Mid(EguStr, 2)
  Else
    EguSym = EguStr          ' MR 16-01-03 enable units without factors
  End If
  
  Select Case EguSym         'check if um exists and compute numeric value
    Case "s":   EguBase = 1                 'second
    Case "Hz":  EguBase = 1                 'frequency
    Case "m":   EguBase = 1                 'meter
    Case "g":   EguBase = 0.001             'gramme
    Case "rad", "Rad", "RAD": EguBase = 1   'radiant  '18-10-02 VL
    Case "S":   EguBase = 1                 'siemens
    Case "V":   EguBase = 1                 'volt
    Case "A":   EguBase = 1                 'ampere
    Case "W":   EguBase = 1                 'watt
    Case "F":   EguBase = 1                 'farad
    Case "bar": EguBase = 1                 'bar
    Case "Pa":  EguBase = 1                 'pascal
    Case "Nm":  EguBase = 1                 'newtonmeter
    Case "Ohm", "ohm": EguBase = 1          'ohm     '18-10-02 VL
    Case Else
      'ErrMsg = "unknown unit of measure: " + EguSym
      convEGU = False
      Exit Function
  End Select
  retval = EguCoeff * EguBase * 10 ^ EguFact   'finally compute the numeric value
  convEGU = True
End Function
'-------------------------------------------------------------------------------
'check if it is a letter
Private Function IsLetter(ByVal char As String) As Boolean
  Dim code As Long
  code = Asc(char)
  IsLetter = (65 <= code And code <= 90) Or (97 <= code And code <= 122) Or char = "_"
End Function
'-------------------------------------------------------------------------------
'check if it is a variable name
Private Function IsVariable(ByVal Str As String) As Boolean
  Dim i As Long, ch As String
  If IsLetter(Left(Str, 1)) Then
    For i = 2 To Len(Str)
      ch = Mid(Str, i, 1)
       If Not IsLetter(ch) Then If Not IsDigit(ch) Then Exit Function
    Next i
    IsVariable = True
  End If
End Function
'-------------------------------------------------------------------------------
'check if it is a digit
Private Function IsDigit(ByVal char As String) As Boolean
  Dim code As Long
  code = Asc(char)
  IsDigit = (48 <= code And code <= 57)
End Function
'-------------------------------------------------------------------------------
'check for an expression to occur in a list
Private Function InList(ByVal strElem As String, ByVal strList As String) As Boolean
  Dim lstrElem As String
  Dim lstrList As String
    
  lstrList = " " & strList & " "
  lstrElem = " " & strElem & " "
  InList = InStr(1, lstrList, lstrElem, vbTextCompare) > 0
End Function
'-------------------------------------------------------------------------------
' translate a symbolic Constant to its double value
Private Function convSymbConst(ByVal strSource As String, ByRef retval As Double) As Boolean
  Dim CostToken   As String
  Dim SymbConst   As String
  CostToken = "#"
  convSymbConst = False:  ErrMsg = "": ErrId = 0
  'check if string is "pi" only for compatibility with previous release.
  If LCase(strSource) = "pi" Then strSource = strSource & CostToken
  If Right(strSource, 1) <> CostToken Then Exit Function
  retval = 0
  SymbConst = Left(strSource, Len(strSource) - 1)
  Select Case SymbConst
    Case "pi", "PI":  retval = PI_              'pi-greek
    Case "pi2": retval = PI_ / 2                'pi-greek/2
    Case "pi3": retval = PI_ / 3                'pi-greek/3
    Case "pi4": retval = PI_ / 4                'pi-greek/4
    Case "e":   retval = 2.71828182845905       'Euler-Napier constant
    Case "eu":  retval = 0.577215664901533      'Euler-Mascheroni constant
    Case "phi": retval = 1.61803398874989       'golden ratio
    Case "g":   retval = 9.80665                'Acceleration due to gravity
    Case "G":   retval = 6.672 * 10 ^ -11       'Gravitational constant
    Case "R":   retval = 8.31451                'Gas constant
    Case "eps": retval = 8.854187817 * 10 ^ -12 'Permittivity of vacuum
    Case "mu":  retval = 12.566370614 * 10 ^ -7 'Permeability of vacuum
    Case "c":   retval = 2.99792458 * 10 ^ 8    'Speed of light
    Case "q":   retval = 1.60217733 * 10 ^ -19  'Elementary charge
    Case "me":  retval = 9.1093897 * 10 ^ -31   'Electron rest mass
    Case "mp":  retval = 1.6726231 * 10 ^ -27   'Proton rest mass
    Case "mn":  retval = 1.6749286 * 10 ^ -27   'Neutron rest mass
    Case "K":   retval = 1.380658 * 10 ^ -23    'Boltzmann constant
    Case "h":   retval = 6.6260755 * 10 ^ -34   'Planck constant
    Case "A":   retval = 6.0221367 * 10 ^ 23    'Avogadro number
    Case Else
      ' > Berend 20041216 - support intrinsic date/time values
      Select Case UCase$(SymbConst)
        Case "DATE"  'or date
          retval = CDbl(Date)
        Case "TIME"  'or time
          retval = CDbl(Time)
        Case "NOW"   'or now
          retval = CDbl(Now)
        Case Else
          ErrMsg = getMsg(19, SymbConst)  ' "Constant unknown: " & SymbConst
      End Select
      ' < Berend 20041216
  End Select
  convSymbConst = True
End Function
'-------------------------------------------------------------------------------
'break the variable string into the name and its sign (if any). Es -x
Private Sub Catch_Sign(ByRef Str As String, ByRef sign As Long)
Dim s As String, VarName As String
sign = 1
s = Left(Str, 1)
If s = "-" Or s = "+" Then
    Str = Right(Str, Len(Str) - 1)
    If s = "-" Then sign = -sign
End If
End Sub
'return the current environment setting for decimal separator
'about 2-3 us, that is 20 times faster than Application.International(xlDecimalSeparator)
Private Function Decimal_Point_Is()
Decimal_Point_Is = Mid(CStr(1 / 2), 2, 1)
End Function

Private Function Acos_(ByVal A As Double) As Double
  If A = 1 Then
    Acos_ = 0
  ElseIf A = -1 Then
    Acos_ = PI_
  Else
    Acos_ = Atn(-A / Sqr(-A * A + 1)) + 2 * Atn(1)
  End If
End Function

Private Function Asin_(ByVal A As Double) As Double
  If Abs(A) = 1 Then
    Asin_ = Sgn(A) * PI_ / 2
  Else
    Asin_ = Atn(A / Sqr(-A * A + 1))
  End If
End Function

Private Function Cosh_(ByVal A As Double) As Double
  Cosh_ = (Exp(A) + Exp(-A)) / 2
End Function

Private Function Sinh_(ByVal A As Double) As Double
  Sinh_ = (Exp(A) - Exp(-A)) / 2
End Function

Private Function Tanh_(ByVal A As Double) As Double
  Tanh_ = (Exp(A) - Exp(-A)) / (Exp(A) + Exp(-A))
End Function

Private Function Acosh_(ByVal A As Double) As Double
  Acosh_ = Log(A + Sqr(A * A - 1))
End Function

Private Function Asinh_(ByVal A As Double) As Double
  Asinh_ = Log(A + Sqr(A * A + 1))
End Function

Private Function Atanh_(ByVal A As Double) As Double
  Atanh_ = Log((1 + A) / (1 - A)) / 2  'bug 3-1-2003 VL
End Function

Private Function round_(ByVal x As Double, ByVal n As Double) As Double
Dim Xi As Double, xd As Double, B As Double, D As Long
D = CheckInt(n)
B = 10 ^ D
x = x * B
Xi = Int(x)
xd = x - Xi
If xd > 0.5 Then Xi = Xi + 1
round_ = Xi / B
End Function

'-------------------------------------------------------------------------------
' calculate Factorial  (bug overflow for n > 12, 8-7-02 VL)
Private Function fact(ByVal n As Double) As Double
  Dim p As Double, i As Long, m As Long
  '7.10.2003 thanks to Rodrigo Farinha
  If n < 0 Then
    fact = ""  'raise an error
  Else
    m = CheckInt(n)
    p = 1
    For i = 1 To Int(n)
      p = p * i
    Next
    fact = p
  End If
End Function
'-------------------------------------------------------------------------------
'MCM between two integer numbers
Private Function mcm_2(ByVal A As Double, ByVal B As Double) As Double
  Dim x As Long, Y As Long
  If A < 0 Or B < 0 Then Y = "" 'raises in error
  Y = CheckInt(Abs(A))
  x = CheckInt(Abs(B))
  mcm_2 = x * Y / mcd_2(x, Y)
End Function
'-------------------------------------------------------------------------------
'MCM of n-numbers
Private Function MCM_(x() As Double) As Double
Dim n As Integer, i As Integer, Ris As Double
n = UBound(x)
Ris = x(1)
For i = 2 To n
    Ris = mcm_2(Ris, x(i))
Next i
MCM_ = Ris
End Function
'-------------------------------------------------------------------------------
'MCD between two integer numbers
Private Function mcd_2(ByVal A As Double, ByVal B As Double) As Double
  Dim x As Long, Y As Long, r As Long
  If A < 0 Or B < 0 Then Y = "" 'raises in error
  Y = CheckInt(A)
  x = CheckInt(B)
  Do Until x = 0
    r = Y Mod x
    Y = x
    x = r
  Loop
  mcd_2 = Y
End Function
'-------------------------------------------------------------------------------
'MCD of n-numbers
Private Function MCD_(x() As Double) As Double
Dim n As Integer, i As Integer, Ris As Double
n = UBound(x)
Ris = x(1)
For i = 2 To n
    Ris = mcd_2(Ris, x(i))
Next i
MCD_ = Ris
End Function
'-------------------------------------------------------------------------------
' combinations n objects, k classes
Private Function Comb(ByVal A As Double, ByVal B As Double) As Double
    Dim n As Long, k As Long, Y As Double, i As Long
    
    If A < 0 Or B < 0 Then Y = "" 'raises in error
    n = CheckInt(A)
    k = CheckInt(B)
    If n < 1 Or k < 1 Or k > n Then Comb = 0: Exit Function 'mod. 1.4.04 VL
    If k = n Then Comb = 1: Exit Function
    Y = n
    If k > Int(n / 2) Then k = n - k
    For i = 2 To k
        Y = Y * (n + 1 - i) / i
    Next i
    Comb = Y
End Function
'-------------------------------------------------------------------------------
' Permuations n objects, k classes
Private Function Perm(ByVal A As Double, ByVal B As Double) As Double
    Dim n As Long, k As Long, Y As Double, i As Long
    
    If A < 0 Or B < 0 Then Y = "" 'raises in error
    n = CheckInt(A)
    k = CheckInt(B)
    If n < 1 Or k < 1 Or k > n Then Perm = 0: Exit Function
    Y = n
    For i = 2 To k
        Y = Y * (n + 1 - i)
    Next i
    Perm = Y
End Function
'-------------------------------------------------------------------------------
'max value of n-numbers
Private Function max_n(x() As Double) As Double
Dim n As Integer, i As Integer, Ris As Double
n = UBound(x)
Ris = x(1)
For i = 2 To n
    Ris = max_(Ris, x(i))
Next i
max_n = Ris
End Function
'-------------------------------------------------------------------------------
' max value of 2 numbers
Private Function max_(ByVal A As Double, ByVal B As Double) As Double
  If A > B Then max_ = A Else max_ = B
End Function
'-------------------------------------------------------------------------------
'min value of n-numbers
Private Function min_n(x() As Double) As Double
Dim n As Integer, i As Integer, Ris As Double
n = UBound(x)
Ris = x(1)
For i = 2 To n
    Ris = min_(Ris, x(i))
Next i
min_n = Ris
End Function
'-------------------------------------------------------------------------------
' min value of 2 numbers
Private Function min_(ByVal A As Double, ByVal B As Double) As Double
  If A < B Then min_ = A Else min_ = B
End Function
'-------------------------------------------------------------------------------
' count number of abs sybol sets in formula
Private Function NabsCount(ByVal s As String) As Long
  Dim n As Long, p As Long
  n = 0
  p = InStr(1, s, "|")
  Do While p > 0
    p = p + 1
    n = n + 1
    p = InStr(p, s, "|")
  Loop
  NabsCount = n
End Function

Private Function erf_(ByVal x As Double) As Double
Dim Y As Double
    Call Herf(x, Y)
    erf_ = Y
End Function
'-------------------------------------------------------------------------------
' gamma function  22.2.05   fix bug for x<0
Private Function gamma_(ByVal x As Double) As Double
Dim mantissa As Double, Expo As Double, z As Double
Dim t As Double, Y As Double, E As Long
If x <= 0 And x - Int(x) = 0 Then 'negative integer
    gamma_ = "?": Exit Function
End If
z = Abs(x)
gamma_split z, mantissa, Expo
If x < 0 Then
    t = z * Sin(PI_ * z)
    Y = -PI_ / (mantissa * t)
    E = Int(Log(Abs(Y)) / Log(10#))
    mantissa = Y * 10 ^ -E
    Expo = E - Expo
End If
gamma_ = mantissa * 10 ^ Expo
End Function
'-------------------------------------------------------------------------------
' logarithm gamma function
Private Function gammaln_(ByVal x As Double) As Double
Dim mantissa As Double, Expo As Double
gamma_split x, mantissa, Expo
gammaln_ = Log(mantissa) + Expo * Log(10)
End Function

Private Function digamma_(ByVal x As Double) As Double
' digamma function
Dim Y As Double
    Call HDigamma(x, Y)
    digamma_ = Y
End Function

Private Function beta_(ByVal z As Double, ByVal W As Double) As Double
' beta function
Dim Y
    Call HBeta(z, W, Y)
    beta_ = Y
End Function

Private Function Zeta_(ByVal x As Double) As Double
' Riemman's zeta function
Dim Y As Double
    Call HZeta(x, Y)
    Zeta_ = Y
End Function

Private Function exp_integr(ByVal x As Double) As Double
' exponential integral Ei(x) for x >0.
Dim Y As Double
    Call Hexp_integr(x, Y)
    exp_integr = Y
End Function

Private Function cvDegree(ByVal DMS As String, ByRef Angle As Double, Optional ByRef sMsg As String) As Boolean
'converts a string dd�mm'ss" (degrees, minutes, seconds) into a decimal-degree angle
'mod 16.12.2004
Dim p1&, p2&, P3&
Dim A1&, A2&, A3&
Dim b1&, b2&, b3&
Dim s1$, s2$, s3$
Dim DD As Double, mm As Double, ss As Double
Dim sum_a&, sum_b&, i&

'check flag dms conversion  23.2.05
If DMS_conv = False Then
    cvDegree = False
    Exit Function
End If

Angle = 0
sMsg = ""
DMS = Trim(DMS)
'#If CODPAGE = 0 Then
'A1 = InStr(1, DMS, "�") ' degrees �
'A2 = InStr(1, DMS, "'")  ' minutes '
'A3 = InStr(1, DMS, """")  ' seconds "
'#End If
sum_a = A1 + A2 + A3
b1 = InStr(1, DMS, "d") ' degrees �
b2 = InStr(1, DMS, "m")  ' minutes '
b3 = InStr(1, DMS, "s")  ' seconds "
sum_b = b1 + b2 + b3

If sum_a > 0 And sum_b = 0 Then
    p1 = A1: p2 = A2: P3 = A3
ElseIf sum_a = 0 And sum_b > 0 Then
    p1 = b1: p2 = b2: P3 = b3
Else
    GoTo Error_Handler_False  'no mixed format allowed
End If
If p1 = 0 Then GoTo Error_Handler_False
If p2 = 0 Then p2 = p1
On Error Resume Next
s1 = Mid(DMS, 1, p1 - 1)
s2 = Mid(DMS, p1 + 1, p2 - p1 - 1)
s3 = Mid(DMS, p2 + 1, P3 - p2 - 1)
If s3 = "" And p2 = Len(DMS) Then s3 = "0"
On Error GoTo 0
i = 0
If Not IsNumeric_(s1) Then i = i + 1
If Not IsNumeric_(s2) Then i = i + 1
If Not IsNumeric_(s3) Then i = i + 1
If i = 1 Then GoTo Error_Handler_True  'only one error
If i > 1 Then GoTo Error_Handler_False 'too many errors
If P3 > 0 And P3 < Len(DMS) Then GoTo Error_Handler_True

DD = CDbl(s1)
mm = CDbl(s2)
ss = CDbl(s3)

If mm > 60 Or ss > 60 Then GoTo Error_Handler_True

Angle = DD + (mm + ss / 60) / 60
cvDegree = True
Exit Function
Error_Handler_True:
    cvDegree = True
    sMsg = getMsg(21) '"Wrong DMS format"
    Exit Function
Error_Handler_False:
    cvDegree = False
    sMsg = getMsg(22) '"No DMS format"
End Function

Private Function IsNumeric_(ByVal x As String) As Boolean
'numeric check, dependent or independent from international system setting
'mod. by Ricardo Mart�nez C.
'21.10.2006
Dim i As Integer
If DecSep = "." Then
'the decimal separator is the period (.)
    IsNumeric_ = False
    If InStr(1, x, ",") > 0 Then Exit Function 'Comma is not allowed as decimal separator.
    If InStr(1, x, "d", vbTextCompare) > 0 Then Exit Function ' bug 85d5 = 8.5e+6 thanks to PJ Weng  12.6.2004
    i = InStr(1, x, DecSep)
    If i > 0 Then If InStr(i + 1, x, DecSep) > 0 Then Exit Function 'Too many periods.
    If i = 1 And Len(x) > 1 Then x = "0" & x 'Let ".25" = "0.25"
    IsNumeric_ = IsNumeric(x)
Else
'the decimal separator is the comma (,)
    IsNumeric_ = False
    If InStr(1, x, ".") > 0 Then Exit Function 'point is not allowed as decimal separator.
    If InStr(1, x, "d", vbTextCompare) > 0 Then Exit Function ' bug 85d5 = 8.5e+6 thanks to PJ Weng  12.6.2004
    i = InStr(1, x, DecSep)
    If i > 0 Then If InStr(i + 1, x, DecSep) > 0 Then Exit Function 'Too many commas.
    If i = 1 And Len(x) > 1 Then x = "0" & x 'Let ",25" = "0,25"
    IsNumeric_ = IsNumeric(x)
End If
End Function

Private Function MiRoot_(ByVal A As Double, ByVal n As Double) As Double
Dim m As Double
'7.10.2003 thanks to Rodigro Farinha
'algebric extension of root for a<0
    m = Int(n)  'only integer here
    If m = 0 Then
        MiRoot_ = "" 'raise an error
    ElseIf Mod_(m, 2) = 0 Then 'm is even => root in a<0 doesn�t exist
        If A < 0 Then
            MiRoot_ = "" 'raise an error
        Else
            MiRoot_ = A ^ (1 / m)
        End If
    Else  'm is odd => root in a<0 exists
        MiRoot_ = Sgn(A) * Abs(A) ^ (1 / m)
    End If
End Function

'compute the unique nonnegative remainder on division
'of the integer x by the integer n
'1-3-07
Private Function Mod_(ByVal A As Double, ByVal B As Double) As Double
'fix the Excel VBA Bug
Dim c, D, E
    c = Int(Abs(A))
    D = Int(Abs(B))
    E = Round(c - D * Int(c / D), 0)
    If A < 0 Then E = D - E
    Mod_ = E
End Function

Private Function Dec_(ByVal A As Double) As Double
Dim z As Double, n As Long
    z = A - Fix(A)
    n = Int(Log(Abs(A)) / Log(10)) + 1 'integer digits
    z = Round(z, 15 - n)
    Dec_ = z
End Function

Private Function CheckInt(ByVal A As Double) As Double
'check if a value is integer
'raises an error if variable a is not integer
Dim temp As Double, D As Double
Const tiny = 5 * 10 ^ -14
D = Round(A, 0)
temp = Abs(D - A)
If temp > tiny Then CheckInt = "" 'raises an error
CheckInt = D
End Function

'use only for debug
Sub ET_Dump(ByRef ETable() As Variant)
  ReDim ETable(ETtop, 30)
  Dim i As Long, j As Long
  
  j = j + 1: ETable(0, j) = "Fun"
  j = j + 1: ETable(0, j) = "ArgTop"
  j = j + 1: ETable(0, j) = "A1 Idx"
  j = j + 1: ETable(0, j) = "Arg1 Name"
  j = j + 1: ETable(0, j) = "Arg1 Value"
  j = j + 1: ETable(0, j) = "A2 Idx"
  j = j + 1: ETable(0, j) = "Arg2 Name"
  j = j + 1: ETable(0, j) = "Arg2 Value"
  j = j + 1: ETable(0, j) = "ArgOf"
  j = j + 1: ETable(0, j) = "ArgIdx"
  j = j + 1: ETable(0, j) = "Value"
  j = j + 1: ETable(0, j) = "PriLvl"
  j = j + 1: ETable(0, j) = "PosInExpr"
  j = j + 1: ETable(0, j) = "PriIdx"
  j = j + 1: ETable(0, j) = "Cond"
  ReDim Preserve ETable(ETtop, j)
  For i = 1 To UBound(ET)
    j = 0
    With ET(i)
        j = j + 1: ETable(i, j) = .Fun
        j = j + 1: ETable(i, j) = .ArgTop
        j = j + 1: ETable(i, j) = .arg(1).idx
        j = j + 1: ETable(i, j) = .arg(1).Nome
        j = j + 1: ETable(i, j) = .arg(1).value
        j = j + 1: ETable(i, j) = .arg(2).idx
        j = j + 1: ETable(i, j) = .arg(2).Nome
        j = j + 1: ETable(i, j) = .arg(2).value
        j = j + 1: ETable(i, j) = .ArgOf
        j = j + 1: ETable(i, j) = .ArgIdx
        j = j + 1: ETable(i, j) = .value
        j = j + 1: ETable(i, j) = .PriLvl
        j = j + 1: ETable(i, j) = .PosInExpr
        j = j + 1: ETable(i, j) = .PriIdx
        j = j + 1: ETable(i, j) = .Cond
    End With
  Next
End Sub


Private Function DNormal(ByVal x As Double, ByVal avg As Double, ByVal dev As Double) As Double
'normal distribution
Dim p As Double
p = (x - avg) ^ 2 / (2 * dev ^ 2)
DNormal = Exp(-p) / Sqr(2 * PI_) / dev
End Function

Private Function CNormal(ByVal x As Double, ByVal avg As Double, ByVal dev As Double) As Double
'cumulative normal distribution
Dim p As Double
p = (x - avg) / (Sqr(2) * dev)
CNormal = (1 + erf_(p)) / 2
End Function

Private Function DBinomial(ByVal k As Double, ByVal n As Double, ByVal p As Double) As Double
'k = class, N = population, p = probability
k = CheckInt(k)
n = CheckInt(n)
DBinomial = Comb(n, n - k) * p ^ k * (1 - p) ^ (n - k)
End Function
'
Private Function CBinomial(ByVal k As Double, ByVal n As Double, ByVal p As Double) As Double
'k = class, N = population, p = probability
Dim i As Long, Y As Double
k = CheckInt(k)
n = CheckInt(n)
For i = 1 To k
    Y = Y + DBinomial(i, n, p)
Next i
CBinomial = Y
End Function

Private Function DPoisson(ByVal k As Double, ByVal z As Double) As Double
'Poisson density
' k = events, z = average
Dim Y As Double
k = CheckInt(k)
Y = -z + k * Log(z) - gammaln_(k + 1)
DPoisson = Exp(Y)
End Function

Private Function CPoisson(ByVal k As Double, ByVal z As Double) As Double
'Poisson cumulative distribution
' k = events, z = average
Dim Y As Double, i As Long
k = CheckInt(k)
For i = 1 To k
    Y = Y + DPoisson(i, z)
Next
CPoisson = Y
End Function

Private Function expn_integr(ByVal x As Double, ByVal n As Double) As Double
'Evaluates the exponential integral En(x).
Dim Y As Double
    Call Hexpn_integr(x, n, Y)
    expn_integr = Y
End Function

Private Function BetaInc(ByVal x As Double, ByVal A As Double, ByVal B As Double) As Double
'incomplete gamma function
Dim BIX As Double
Call INCOB(A, B, x, BIX)
BetaInc = BIX
End Function

Private Function GammaInc(ByVal A As Double, ByVal x As Double) As Double
'incomplete gamma function
Dim GIN As Double, GIM As Double, GIP As Double, MSG As String
Call INCOG(A, x, GIN, GIM, GIP, MSG)
If MSG <> "" Then GammaInc = ""  'raise an error
GammaInc = GIM  '23.3.06
End Function

'22-02.2007
Private Function Hypergeom(ByVal A As Double, ByVal B As Double, ByVal c As Double, ByVal x As Double) As Double
Dim hf As Double, ErrorMsg As String
If x > 1 Then GoTo Error_Handler
Call HYGFX(A, B, c, x, hf, ErrorMsg)
If ErrorMsg <> "" Then GoTo Error_Handler
    Hypergeom = hf
Exit Function
Error_Handler:
    Hypergeom = ""  'raise an error
End Function

Private Function BesselJ(ByVal x As Double, Optional ByVal n As Double) As Double
'Bessel function 1st kind, order n, Jn(x)
Dim BJ0#, DJ0#, BJ1#, DJ1#, BY0#, DY0#, BY1#, DY1#, NM#, BJ#(), DJ#(), BY#(), DY#()
If IsMissing(n) Then n = 0
If n <= 1 Then
    Call JY01A(x, BJ0, DJ0, BJ1, DJ1, BY0, DY0, BY1, DY1)
    If n = 0 Then BesselJ = BJ0 Else BesselJ = BJ1
Else
    Call JYNA(n, x, NM, BJ, DJ, BY, DY)
    BesselJ = BJ(n)
End If
End Function

Private Function BesselY(ByVal x As Double, Optional ByVal n As Double) As Double
'Bessel function 2nd kind, order n, Yn(x)
Dim BJ0#, DJ0#, BJ1#, DJ1#, BY0#, DY0#, BY1#, DY1#, NM#, BJ#(), DJ#(), BY#(), DY#()
If IsMissing(n) Then n = 0
If n <= 1 Then
    Call JY01A(x, BJ0, DJ0, BJ1, DJ1, BY0, DY0, BY1, DY1)
    If n = 0 Then BesselY = BY0 Else BesselY = BY1
Else
    Call JYNA(n, x, NM, BJ, DJ, BY, DY)
    BesselY = BY(n)
End If
End Function

Private Function BesseldJ(ByVal x As Double, Optional ByVal n As Double) As Double
'First Derivative of Bessel functions first kind, order n, J'n(x)
Dim BJ0#, DJ0#, BJ1#, DJ1#, BY0#, DY0#, BY1#, DY1#, NM#, BJ#(), DJ#(), BY#(), DY#()
If IsMissing(n) Then n = 0
If n <= 1 Then
    Call JY01A(x, BJ0, DJ0, BJ1, DJ1, BY0, DY0, BY1, DY1)
    If n = 0 Then BesseldJ = DJ0 Else BesseldJ = DJ1
Else
    Call JYNA(n, x, NM, BJ, DJ, BY, DY)
    BesseldJ = DJ(n)
End If
End Function

Private Function BesseldY(ByVal x As Double, Optional ByVal n As Double) As Double
'First Derivative of Bessel functions second kind, order n, Y'n(x)
Dim BJ0#, DJ0#, BJ1#, DJ1#, BY0#, DY0#, BY1#, DY1#, NM#, BJ#(), DJ#(), BY#(), DY#()
If IsMissing(n) Then n = 0
If n <= 1 Then
    Call JY01A(x, BJ0, DJ0, BJ1, DJ1, BY0, DY0, BY1, DY1)
    If n = 0 Then BesseldY = DY0 Else BesseldY = DY1
Else
    Call JYNA(n, x, NM, BJ, DJ, BY, DY)
    BesseldY = DY(n)
End If
End Function

Private Function BesselI(ByVal x As Double, Optional ByVal n As Double) As Double
'modified Bessel function 1st kind, order n, In(x)
Dim BI0#, DI0#, BI1#, DI1#, BK0#, DK0#, BK1#, DK1#, NM#, BI#(), DI#(), BK#(), DK#()
If IsMissing(n) Then n = 0
If n <= 1 Then
    Call IK01A(x, BI0, DI0, BI1, DI1, BK0, DK0, BK1, DK1)
    If n = 0 Then BesselI = BI0 Else BesselI = BI1
Else
    Call IKNA(n, x, NM, BI, DI, BK, DK)
    BesselI = BI(n)
End If
End Function

Private Function BesseldI(ByVal x As Double, Optional ByVal n As Double) As Double
'derivative modified Bessel function 1� kind, order n, In(x)
Dim BI0#, DI0#, BI1#, DI1#, BK0#, DK0#, BK1#, DK1#, NM#, BI#(), DI#(), BK#(), DK#()
If IsMissing(n) Then n = 0
If n <= 1 Then
    Call IK01A(x, BI0, DI0, BI1, DI1, BK0, DK0, BK1, DK1)
    If n = 0 Then BesseldI = DI0 Else BesseldI = DI1
Else
    Call IKNA(n, x, NM, BI, DI, BK, DK)
    BesseldI = DI(n)
End If
End Function

Private Function BesselK(ByVal x As Double, Optional ByVal n As Double) As Double
'modified Bessel function 2nd kind, order n, In(x)
Dim BI0#, DI0#, BI1#, DI1#, BK0#, DK0#, BK1#, DK1#, NM#, BI#(), DI#(), BK#(), DK#()
If IsMissing(n) Then n = 0
If n <= 1 Then
    Call IK01A(x, BI0, DI0, BI1, DI1, BK0, DK0, BK1, DK1)
    If n = 0 Then BesselK = BK0 Else BesselK = BK1
Else
    Call IKNA(n, x, NM, BI, DI, BK, DK)
    BesselK = BK(n)
End If
End Function

Private Function BesseldK(ByVal x As Double, Optional ByVal n As Double) As Double
'derivative of modified Bessel function 2� kind, order n, In(x)
Dim BI0#, DI0#, BI1#, DI1#, BK0#, DK0#, BK1#, DK1#, NM#, BI#(), DI#(), BK#(), DK#()
If IsMissing(n) Then n = 0
If n <= 1 Then
    Call IK01A(x, BI0, DI0, BI1, DI1, BK0, DK0, BK1, DK1)
    If n = 0 Then BesseldK = DK0 Else BesseldK = DK1
Else
    Call IKNA(n, x, NM, BI, DI, BK, DK)
    BesseldK = DK(n)
End If
End Function

Private Function CosIntegral(ByVal x As Double) As Double
'returns cos integral ci(x)
Dim CI As Double, SI As Double
    Call CISIA(x, CI, SI)
    CosIntegral = CI
End Function

Private Function SinIntegral(ByVal x As Double) As Double
'returns sin integral ci(x)
Dim CI As Double, SI As Double
    Call CISIA(Abs(x), CI, SI)
    SinIntegral = Sgn(x) * SI
End Function

Private Function Fresnel_cos(ByVal x As Double) As Double
'returns Fresnel's cos integral
Dim Fr_c As Double, Fr_s As Double
    Call FCS(x, Fr_c, Fr_s)
    Fresnel_cos = Fr_c
End Function

Private Function Fresnel_sin(ByVal x As Double) As Double
'returns Fresnel's sin integral
Dim Fr_c As Double, Fr_s As Double
    Call FCS(x, Fr_c, Fr_s)
    Fresnel_sin = Fr_s
End Function

Private Function Poly_Legendre(ByVal x As Double, ByVal n As Double) As Double
Dim Y As Double
    n = CheckInt(n)
    Call PLegendre(x, n, Y)
    Poly_Legendre = Y
End Function

Private Function Poly_Hermite(ByVal x As Double, ByVal n As Double) As Double
Dim Y As Double
    n = CheckInt(n)
    Call PHermite(x, n, Y)
    Poly_Hermite = Y
End Function

Private Function Poly_Laguerre(ByVal x As Double, ByVal n As Double) As Double
Dim Y As Double
    n = CheckInt(n)
    Call PLaguerre(x, n, Y)
    Poly_Laguerre = Y
End Function

Private Function Poly_Chebycev(ByVal x As Double, ByVal n As Double) As Double
Dim Y As Double
    n = CheckInt(n)
    Call PChebycev(x, n, Y)
    Poly_Chebycev = Y
End Function

Private Function Airy_A(ByVal x As Double) As Double
Dim Y As Double, AI As Double, BI As Double, AD As Double, BD As Double
    Call AIRYB(x, AI, BI, AD, BD)
    Airy_A = AI
End Function

Private Function Airy_B(ByVal x As Double) As Double
Dim Y As Double, AI As Double, BI As Double, AD As Double, BD As Double
    Call AIRYB(x, AI, BI, AD, BD)
    Airy_B = BI
End Function

Private Function IElliptic_Int1(ByVal x As Double, ByVal phi As Double) As Double
' incomplete elliptic integral 1st kind
Dim e1 As Double, e2 As Double, phideg As Double
    phideg = 180 * phi / PI_
    Call ELIT(x, phideg, e1, e2)
    IElliptic_Int1 = e1
End Function

Private Function IElliptic_Int2(ByVal x As Double, ByVal phi As Double) As Double
' incomplete elliptic integral 2nd kind
Dim e1 As Double, e2 As Double, phideg As Double
    phideg = 180 * phi / PI_
    Call ELIT(x, phideg, e1, e2)
    IElliptic_Int2 = e2
End Function

'clipping function
Private Function Clip(ByVal t, ByVal Floor, ByVal Ceeling) As Double
Dim Y As Double
If Floor > Ceeling Then Clip = "" 'raise an error
If t < Floor Then
    Y = Floor
ElseIf t > Ceeling Then
    Y = Ceeling
Else
    Y = t
End If
Clip = Y
End Function

'step function or Haveside's function
Private Function Step_(ByVal x, ByVal A) As Double
If x >= A Then
    Step_ = 1
Else
    Step_ = 0
End If
End Function


'sum of n-numbers
Private Function sum_(x() As Double) As Double
Dim n As Integer, i As Integer, Ris As Double
n = UBound(x)
For i = 1 To n
    Ris = Ris + x(i)
Next i
sum_ = Ris
End Function

'arithemtic mean of n-numbers
Private Function mean_(x() As Double) As Double
Dim n As Integer, i As Integer, Ris As Double
n = UBound(x)
For i = 1 To n
    Ris = Ris + x(i)
Next i
mean_ = Ris / n
End Function

'geometric mean of n-numbers
Private Function meang_(x() As Double) As Double
Dim n As Integer, i As Integer, Ris As Double
n = UBound(x)
Ris = 1
For i = 1 To n
    Ris = Ris * (x(i) ^ (1 / n))
Next i
meang_ = Ris
End Function

'quadratic mean of n-numbers
Private Function meanq_(x() As Double) As Double
Dim n As Integer, i As Integer, Ris As Double
n = UBound(x)
Ris = 0
For i = 1 To n
    Ris = Ris + x(i) ^ 2
Next i
meanq_ = Sqr(Ris / n)
End Function

'variance (pop) of n-numbers
Private Function varp_(x() As Double) As Double
Dim n As Integer, i As Integer, Ris As Double, mu As Double
n = UBound(x)
mu = mean_(x)
Ris = 0
For i = 1 To n
    Ris = Ris + (x(i) - mu) ^ 2
Next i
varp_ = Ris / n
End Function

'variance of n-numbers
Private Function var_(x() As Double) As Double
Dim n As Integer
n = UBound(x)
var_ = varp_(x) * n / (n - 1)
End Function

'standard deviation of n-numbers
Private Function stdev_(x() As Double) As Double
stdev_ = Sqr(var_(x))
End Function

'standard deviation (pop) of n-numbers
Private Function stdevp_(x() As Double) As Double
stdevp_ = Sqr(varp_(x))
End Function

Private Sub ErrorTab_Init()
ReDim ErrorTbl(50)
ErrorTbl(1) = "too many variables"
ErrorTbl(2) = "Variable not found"
ErrorTbl(3) = "" 'spare
ErrorTbl(4) = "abs symbols |.| mismatch"
ErrorTbl(5) = "Syntax error at pos: $"
ErrorTbl(6) = "Function < $ > unknown at pos: $"
ErrorTbl(7) = "Too many closing brackets at pos: $"
ErrorTbl(8) = "missing argument"
ErrorTbl(9) = "Too many arguments at pos: $"
ErrorTbl(10) = "" 'spare
ErrorTbl(11) = "Not enough closing brackets"
ErrorTbl(12) = "Syntax error: $"
ErrorTbl(13) = "Function < $ > missing?"
ErrorTbl(14) = "Evaluation error < $($) > at pos: $"
ErrorTbl(15) = "Evaluation error < $($" & ArgSep & " $) > at pos: $"
ErrorTbl(16) = "Evaluation error < $ $ $ > at pos: $"
ErrorTbl(17) = "Evaluation error < $($) > at pos: $"
ErrorTbl(18) = "Variable < $ > not assigned"
ErrorTbl(19) = "Constant unknown: $"
ErrorTbl(20) = "Too many operations"
ErrorTbl(21) = "Wrong DMS format"
ErrorTbl(22) = "No DMS format"
End Sub

'get a message from the error-table and substitute the parameters
Private Function getMsg(Id, Optional p1, Optional p2, Optional P3, Optional P4)
Dim i As Long, s As String, p
ErrId = Id          'set the global id
s = ErrorTbl(ErrId) 'get the message template
If Not IsMissing(p1) Then ParamSubstitute p1, s
If Not IsMissing(p2) Then ParamSubstitute p2, s
If Not IsMissing(P3) Then ParamSubstitute P3, s
If Not IsMissing(P4) Then ParamSubstitute P4, s
getMsg = s
End Function

Private Sub ParamSubstitute(p, s)
Dim i As Long
i = InStr(1, s, "$")
If i > 0 Then
    s = Left(s, i - 1) & p & Right(s, Len(s) - i)
End If
End Sub

' > Mirko 20061018
Private Function GetNumberOfArguments(strExpr) As Integer
    'Count number of commas between parenthesis. Ignore commas in subparenthesis
    Dim i As Integer
    Dim numCommas As Integer
    Dim numOpenPar As Integer
    Do
        i = i + 1
        If Mid(strExpr, i, 1) = "(" Then
            numOpenPar = numOpenPar + 1
        ElseIf Mid(strExpr, i, 1) = ")" Then
            numOpenPar = numOpenPar - 1
        ElseIf Mid(strExpr, i, 1) = ArgSep And numOpenPar = 1 Then
            numCommas = numCommas + 1
        End If
    Loop While numOpenPar > 0 And i < Len(strExpr)
    If i = Len(strExpr) And numOpenPar > 0 Then numCommas = 0
    GetNumberOfArguments = numCommas + 1
End Function
' < Mirko 20061018

# clsMathParser

## Authorage

Author:           Leonardo Volpi

Original thread:  https://web.archive.org/web/20100703220609/http://digilander.libero.it/foxes/mathparser/MathExpressionsParser.htm

Collaborators:    Lieven Dossche, Michael Ruder, Thomas Zeutschler,  Arnaud De Grammont.

## Summary

MathParser - clsMathParser 4. -  is a parser-evaluator for mathematical and physical string expressions.

This software, based on the original MathParser 2  developed  by Leonardo Volpi , has been modified with the collaboration of  Michael Ruder to extend the computation also to the physical variables like "3.5s" for 3.5 seconds and so on.  In advance, a special routine has been developed to find out in which order the variables appear in a formula string. Thomas Zeutschler has kindly revised the code improving general efficiency by more than 200 % . Lieven Dossche has finally encapsulated the core in a nice, efficient and elegant class. In addition, starting from the v.3 of MathParser, Arnoud has created a sophisticated class- clsMathParserC - for complex numbers adding a large, good collection of complex functions and operators in a separate, reusable modules

This software is freeware. Have fun with it.

## Features

This document describes the clsMathParser  4.x class for evaluating real math expressions.

In some instances, you might want to allow users to type in their own numeric expression in ASCII text that contains a series of numeric expressions and then produce a numeric result of this expression.

This class does just that task. It accepts as input any string representing an arithmetic or algebraic expression and a list of the variable's values; it returns a double-precision numeric value. clsMathParser can be used to calculate formulas (expressions) given at runtime, for example to plot and tabulate functions, to make numerical computations and more. It works in VB 6, and VBA.

Generally speaking clsMathParser is a numeric evaluator - not compiled - optimized for fast loop evaluations.

The expression strings accepted are very wide. This parser also recognizes physical expressions, constants and international units of measure.

Typical mixed math-physical expressions are: 

```
1+(2-5)*3+8/(5+3)^2
sqr(2)
(a+b)*(a-b)
x^2+3*x+1
300 km + 123000 m
(3000000 km/s)/144 Mhz
256.33*Exp(-t/12 us)
(1+(2-5)*3+8/(5+3)^2)/sqr(5^2+3^2)
2+3x+2x^2
0.25x + 3.5y + 1
0.1uF*6.8kohm
sqr(4^2+3^2)
(12.3 mm)/(856.6 us)
(-1)^(2n+1)*x^n/n!
And((x<2);(x<=5))
sin(2*pi*x)+cos(2*pi*x)
```

Variables can be any alphanumeric string and must start with a letter

`x  y  a1  a2  time  alpha   beta`

Also the symbol "_" is accepted for writing variables in "programming style"..

`time_1  alpha_b1   rise_time`

Implicit multiplication is not supported because of its intrinsic ambiguity. So "xy" stands for a variable named "xy" and not for x*y. The multiplication symbol "*" cannot generally be omitted.

It can be omitted only for coefficients of the three classic math variables  x, y, z. It means that strings like “2x” and “2*x” are equivalent

`2x  3.141y  338z^2   Û   2*x  3.141*y  338*z^2`

On the contrary, the following expressions are illegal.

`2a  3(x+1)  334omega  2pi`

Constant numbers can be integers, decimal, or exponential

`2  -3234  1.3333  -0.00025   1.2345E-12`

From version 4.2, MathParser accepts both decimal symbols "." or ",". See international setting

Physical numbers are numbers followed by a unit of measure

`"1s" for 1 second   "200m" for 200 meters  "25kg" for 25 kilograms`

For better reading they may contain a blank

`"1 s"   "200 m"   "25 kg"   "150 MHz"   "0.15 uF"   "3600 kohm"`

They may also contain the following multiplying factors:

`T=10^12  G=10^9  M=10^6  k=10^3  m=10^-3  u=10^-6  n=10^-9  p=10^-12`

Functions are called by their function-name followed by parentheses. Arguments can be: numbers, variables, expressions, or even other functions

`sin(x)   log(x)   cos(2*pi*t+phi)   atan(4*sin(x))`

For functions which have more than one argument, the successive arguments are separated by commas (default)

`max(a,b)     root(x,y)    BesselJ(x,n)    HypGeom(x,a,b,c)`

Note. From version 4.2 , the argument separator depends on the MathParser decimal separator setting.  If decimal symbol is point "." (i.e. 3.14) , the argument separator is ",". If it is comma "," (i.e. 3,14) , the argument separator is ";".

Logical expressions are now supported

`x<1   x+2y >= 4    x^2+5x-1>0   t<>0   (0<x<1)`

Logical expressions return always 1 (True) or 0 (False). Compact expressions, like `0<x<1` , are now supported; you can enter:   `(0<x<1)` as well `(0<x)*(x<1)`

Numerical range can be inserted using logical symbols and Boolean functions. For example:

```
For   2<x<5             insert     (2<x)*(x<5)       or also (2<x<5)
For  x<2 , x>=10        insert     OR(x<2, x>=10)    or also   (x<2)+(x>=10)
For  -1<x<1             insert     (x>-1)*(x<1)      or  (-1<x<1)    or also    |x|<1     
```

Piecewise Functions. Logical expressions can also be useful for defining a piecewise function, such as:

 
```
       /            2x-1-ln(2)          | x ≤ 0.5
f(x) = |            ln(x)               | 0.5 < x < 2
       \            x/2-1+ln(2)         | x ≥ 2
```
 

The above function can be written as: 

`f(x) = (x<=0.5)*(2*x-1-ln(2))+(0.5<x<2)*ln(x)+(x>=2)*(x/2-1+ln(2))`

Starting from v3.4, the parser adopts a new algorithm for evaluating math expressions depending on logical expressions, which are evaluated only if the logical conditions are true (Conditioned-Branch algorithm). Thus, the above piecewise expression can be evaluated for any real value x without any domain error. Note that without this features the formula could be evaluated only for x>0. Another way to compute piecewise functions is splitting it into several formulas (see example 6)

Percentage.  (changed) Now it simply returns the argument divided by 100

`3%   => returns the number  3/100 = 0.03`

Math Constants supported are: Pi Greek (p), Euler-Mascheroni (g) , Euler-Napier’s (e), Goldean mean (j ). Constant numbers must be suffixed with # symbol (except pi-greek that can written also without a suffix for compatibility with previous versions)

```
pi = 3.14159265358979    or  pi# = 3.14159265358979  
pi2# = 1.5707963267949  (p /2),  pi4# = 0.785398163397448  (p /4))
eu# = 0.577215664901533  
e# = 2.71828182845905
phi# = 1.61803398874989
```

Note: pi-greek constant can be indicated with “pi” or “PI” as well. All other constants are case sensitive.

Angle expressions

This version supports angles in radians, sexagesimal degrees, or centesimal degrees. The right angle unit can be set by the property AngleUnit ("RAD" is the default unit). This affects all angle computation of the parser.

For example if you set the unit "DEG", all angles will be read and converted in degree

```
sin(120)  =>  0.86602540378444
asin(0.86602540378444)  =>  60
rad(pi/2)  =>  90         grad(400)  =>  360       deg(360)  =>  360    
```

Angles can also be written in ddmmss format like for example 45° 12' 13"

```
sin(29°59'60")  =>  0.5            29°59'60"    =>  30       
sin(29d 59m 60s)  =>  0.5          29d 59s 60m  =>  30       
```

Note This format is only for sexagesimal degree. It’s independent from the unit set

 

Note. Oriental version of VB doesn’ t support the first format type. To exclude it simply set the constant

```
#CODEPAGE = 1
```


Physical Constants supported are:

| Title                        | Sym  | Value                     |
|------------------------------|------|---------------------------| 
| Planck constant              | h#   | 6.6260755e-34 J s         |
| Boltzmann constant           | K#   | 1.380658e-23 J/K          |
| Elementary charge            | q#   | 1.60217733e-19 C          | 
| Avogadro number              | A#   | 6.0221367e23 particles/mol|
| Speed of light               | c#   | 2.99792458e8 m/s          |
| Permeability of vacuum ( m ) | mu#  | 12.566370614e-7 T2m3/J    |
| Permittivity of vacuum ( e ) | eps# | 8.854187817e-12  C2/Jm    |
| Electron rest mass           | me#  | 9.1093897e-31 kg          |
| Proton rest mass             | mp#  | 1.6726231e-27 kg          |
| Neutron rest mass            | mn#  | n 1.6749286e-27 kg        |
| Gas constant                 | R#   | 8.31451 m2kg/s2k mol      |
| Gravitational constant       | G#   | 6.672e-11 m3/kg s2        |
| Acceleration due to gravity  | g#   | 9.80665 m/s2              |

Physical constants can be used like any other symbolic math constant.

Just remember that they have their own dimension units listed in the above table.

Example of physical formulas are:

```
m*c# ^2         1/(4*pi*eps#)*q#/r^2       eps# * S/d
sqr(m*h*g#)      s0+v*t+0.5*g#*t^2
```

Multivariable functions. Starting from 4.0, clsMathParser also recognizes and calculates functions with more than 2 variables. Usually, they are functions used in applied math, physics, engineering, etc. Arguments are separated by commas (default)

`HypGeom(x,a,b,c)    Clip(x,a,b)    betaI(x,a,b)   DNorm(x,μ,σ)  etc.`

Functions with variable number of arguments. clsMathParser accepts and calculates functions with variable number of arguments (max 20). Usually they are functions used in statistic and number theory. 

`min(x1,x2,...)    max(x1,x2,...)    mean(x1,x2,...)    gcd(x1,x2,...)`

The max argument limit is set by the global constant  `HiARG`

Time functions. These functions return the current date, time and timestamp of the system. They have no argument and are recognized as intrinsic constants

```
Date#  = current system date
Time#  = current system time
Now#  =  current system timestamp (date + time)
```

International setting. Form 4.2, clsMathParser can accept both decimal separators symbols "." or ",".  Setting the global constant `DP_SET = False`, the programmer can force the parser to follow the international setting  of the machine.  This means that possible decimal numbers in the input string must follow the local international setting.

> The example {{missing image}} must be written as   0.0125*(1+0.8x-1.5x^2) if you system is set for decimal point "." or, on the contrary, must be written:     0,0125*(1+0,8x-1,5x^2)  for system working with decimal comma ",". Setting `DP_SET = True`, the parser, as in the previous releases,  ignores the international setting of the system. In this case the only valid decimal separator  is "." (point). This means that the math input strings are always valid independently from the platform local setting.

Of course, the argument separator symbol changes consequently to the decimal separator in order to avoid conflict. The parser automatically adopts the argument separator  ","  if the decimal separator is point "." or adopts ";" if the decimal separator is "," comma.

The following table shows all possible combinations.


| System setting       | Parser setting `DP_SET` | Decimal separator | Arguments separator |
|----------------------|-------------------------|-------------------|---------------------|
| Decimal is point "." | False                   |  `3.141`          |  `COMB(35,12)`      |
| Decimal is comma "," | False                   |  `3,141`          |  `COMB(35;12)`      | 
| Any                  | True                    |  `3.141`          |  `COMB(35,12)`      | 

## Symbols operators and functions

**Ver 4.1 March. 2005**

This version recognizes more than 130 functions and operators

|  Function  |  Description  |  Note  |
| --- | --- | --- |
|  `+`  |  addition  |     |
|  `-`  |  subtraction  |     |
|  `*`  |  multiplication  |     |
|  `/`  |  division  |  35/4 = 8.75  |
|  `%`  |  percentage  |  35% = 0.35  |
|  `\`  |  integer division  |  35\4 = 8  |
|  `^`  |  raise to power  |  3^1.8 = 7.22467405584208  (°)  |
|  `|` |  |  absolute value  |  |-5|=5      (the same as abs)  |
|  `!`  |  factorial  |  5!=120    (the same as fact)  |
|  abs(x)  |  absolute value  |  abs(-5)= 5  |
|  atn(x), atan(x)  |  inverse tangent  |  atn(pi/4) = 1  |
|  cos(x)  |  cosine  |  argument in radiant  |
|  sin(x)  |  sin  |  argument in radiant  |
|  exp(x)  |  exponential  |  exp(1) = 2.71828182845905  |
|  fix(x)  |  integer part  |  fix(-3.8) = 3  |
|  int(x)  |  integer part  |  int(-3.8) = −4  |
|  dec(x)  |  decimal part  |  dec(-3.8) = -0.8  |
|  ln(x), log(x)  |  logarithm natural  |  argument x>0  |
|  logN(x,n)  |  N-base logarithm  |  logN(16,2) = 4  |
|  rnd(x)  |  random  |  returns a random number between x and 0  |
|  sgn(x)  |  sign  |  returns 1 if x >0 , 0 if x=0, -1 if x<0  |
|  sqr(x)  |  square root  |  sqr(2) =1.4142135623731,  also 2^(1/2)  |
|  cbr(x)  |  cube root  |   "x, example  cbr(2) = 1.2599,  cbr(-2) = -1.2599  |
|  tan(x)  |  tangent  |  argument (in radian)  x¹ k*p/2  with k = ± 1, ± 2…  |
|  acos(x)  |  inverse cosine  |  argument -1 £ x £ 1  |
|  asin(x)  |  Inverse sine  |  argument -1 £ x £ 1  |
|  cosh(x)  |  hyperbolic cosine  |   " x  |
|  sinh(x)  |  hyperbolic sine  |   " x  |
|  tanh(x)  |  hyperbolic tangent  |   " x  |
|  acosh(x)  |  Inverse hyperbolic cosine  |  argument x ³ 1  |
|  asinh(x)  |  Inverse hyperbolic sine  |   " x  |
|  atanh(x)  |  Inverse hyperbolic tangent  |  -1 < x < 1  |
|  root(x,n)  |  n-th root (the same as x^(1/n)  |  argument n ¹ 0  ,  x ³ 0 if n even ,  "x  if n odd  |
|  mod(a,b)  |  Division remainder  |  mod(29,6) = 5    mod(-29 ,6) = -5  |
|  fact(x)  |  factorial  |  argument 0 £ x £ 170   |
|  comb(n,k)  |  combinations  |  comb(6,3) = 20 ,  comb(6,6) = 1  |
|  perm(n,k)  |  permutations  |  perm(8,4) = 1680 ,    |
|  min(a,b,...)  |  minimum  |  min(13,24) = 13  |
|  max(a,b,...)  |  maximum  |  max(13,24) = 24  |
|  mcd(a,b,...)  |  maximum common divisor   |  mcd(4346,174) = 2  |
|  mcm(a,b,...)  |  minimum common multiple   |  mcm(1440,378,1560,72,1650) = 21621600  |
|  gcd(a,b,...)  |  greatest common divisor   |  The same as mcd  |
|  lcm(a,b,...)  |  lowest common multiple   |  The same as mcm  |
|  csc(x)  |  cosecant  |  argument (in radiant) x¹ k*p  with k = 0, ± 1, ± 2…  |
|  sec(x)  |  secant  |  argument (in radiant) x¹ k*p/2  with k = ± 1, ± 2…  |
|  cot(x)  |  cotangent  |  argument (in radiant) x¹ k*p  with k = 0, ± 1, ± 2…  |
|  acsc(x)  |  inverse cosecant  |     |
|  asec(x)  |  inverse secant  |     |
|  acot(x)  |  inverse cotangent  |     |
|  csch(x)  |  hyperbolic cosecant  |  argument x>0  |
|  sech(x)  |  hyperbolic secant  |  argument x>1  |
|  coth(x)  |  hyperbolic cotangent  |  argument x>2  |
|  acsch(x)  |  inverse hyperbolic cosecant  |     |
|  asech(x)  |  inverse hyperbolic secant  |  argument   0 £ x £ 1  |
|  acoth(x)  |  inverse hyperbolic cotangent  |  argument   x<-1 or x>1  |
|  rad(x)  |  radiant conversion  |  converts radiant into current unit of angle  |
|  deg(x)  |  degree sess. conversion  |  converst sess. degree into current unit of angle  |
|  grad(x)  |  degree cent. conversion  |  converts cent. degree into current unit of angle  |
|  round(x,d)  |  round a number with d decimal  |  round(1.35712, 2) = 1.36    |
|  >   |  greater than  |  return 1 (true)   0 (false)  |
|  >=  |  equal or greater than  |  return 1 (true)   0 (false)  |
|  <   |  less than  |  return 1 (true)   0 (false)  |
|  <=  |  equal or less than  |  return 1 (true)   0 (false)  |
|  =  |  equal  |  return 1 (true)   0 (false)  |
|  <>   |  not equal  |  return 1 (true)   0 (false)  |
|  and  |  logic and  |  and(a,b) = return 0 (false)  if a=0 or b=0   |
|  or  |  logic or  |  or(a,b) = return 0 (false) only if a=0 and b=0  |
|  not  |  logic not  |  not(a) = return 0 (false) if a ¹ 0 , else 1  |
|  xor  |  logic exclusive-or  |  xor(a,b) = return 1 (true)  only if a ¹ b  |
|  nand  |  logic nand  |  nand(a,b) = return 1 (true)  if a=1 or b=1   |
|  nor  |  logic nor  |  nor(a,b) = return 1 (true) only if a=0 and b=0  |
|  nxor  |  logic exclusive-nor  |  nxor(a,b) = return 1 (true)  only if a=b  |
|  Psi(x)   |  Function psi   |     |
|  DNorm(x,μ,σ)  |  Normal density function  |  "x,   μ > 0 , σ > 0  |
|  CNorm(x,m,d)  |  Normal cumulative function  |  "x,   μ > 0 , σ > 1  |
|  DPoisson(x,k)  |  Poisson density function  |  x >0,   k = 1, 2, 3 ...  |
|  CPoisson(x,k)  |  Poisson cumulative function k = 1, 2,3 ...  |  x >0,   k = 1, 2, 3 ...  |
|  DBinom(k,n,x)  |  Binomial density for k successes for n trials   |  k , n = 1, 2, 3…,  k < n ,  x £1   |
|  CBinom(k,n,x)  |  Binomial cumulative for k successes for n trials   |  k , n = 1, 2, 3…,  k < n ,  x £1   |
|  Si(x)  |  Sine integral  |  " x  |
|  Ci(x)  |  Cosine integral  |  x >0  |
|  FresnelS(x)  |  Fresnel's sine integral  |  " x  |
|  FresnelC(x)  |  Fresnel's cosine integral  |  " x  |
|  J0(x)  |  Bessel's function of 1st kind  |   x ³0  |
|  Y0(x)  |  Bessel's function of 2st kind  |   x ³0  |
|  I0(x)  |  Bessel's function of 1st kind, modified  |   x >0  |
|  K0(x)  |  Bessel's function of 2st kind, modified  |   x >0  |
|  BesselJ(x,n)  |  Bessel's function of 1st kind, nth order  |   x ³0 , n = 0, 1, 2, 3…  |
|  BesselY(x,n)  |  Bessel's function of 2st kind, nth order  |   x ³0 , n = 0,1, 2, 3…  |
|  BesselI(x,n)  |  Bessel's function of 1st kind, nth order, modified  |   x >0 , n = 0,1, 2, 3…  |
|  BesselK(x,n)  |  Bessel's function of 2st kind, nth order, modified  |   x >0 , n = 0,1, 2, 3…  |
|  HypGeom(x,a,b,c)  |  Hypergeometric function  |   -1 < x <1    a,b >0   c ¹ 0, −1, −2…  |
|  PolyCh(x,n)  |  Chebycev's polynomials  |   "x  ,  orthog. for  -1 £ x £1    |
|  PolyLe(x,n)  |  Legendre's polynomials  |   " x  , orthog. for  -1 £ x £1    |
|  PolyLa(x,n)  |  Laguerre's polynomials  |   " x  , orthog. for  0 £ x £1    |
|  PolyHe(x,n)  |  Hermite's polynomials  |   " x   , orthog. for  −∞ £ x £ +∞  |
|  AiryA(x)     |  Airy function Ai(x)  |   " x  |
|  AiryB(x)    |  Airy function Bi(x)  |   " x  |
|  Elli1(x)  |  Elliptic integral of 1st kind  |   " f   ,   0 < k < 1  |
|  Elli2(x)    |  Elliptic integral of 2st kind  |   " f   ,   0 < k < 1  |
|  Erf(x)  |  Error Gauss's function  |  x >0  |
|  gamma(x)  |  Gamma function  |  " x,  x ¹ 0, −1, −2, −3…  (x > 172 overflow error)  |
|  gammaln(x)  |  Logarithm Gamma function  |  x >0  |
|  gammai(a,x)  |  Gamma Incomplete function  |   " x    a > 0  |
|  digamma(x) psi(x)  |  Digamma function  |  x ¹ 0, −1, −2, −3…  |
|  beta(a,b)  |  Beta function  |  a >0  b >0  |
|  betaI(x,a,b)  |  Beta Incomplete function  |   x >0  ,  a >0  ,  b >0  |
|  Ei(x)  |  Exponential integral  |  x ¹0  |
|  Ein(x,n)  |  Exponential integral of n order  |  x >0  ,  n =  1, 2, 3…  |
|  zeta(x)  |  zeta Riemman's function  |  x <-1 or x >1  |
|  Clip(x,a,b)     |  Clipping function  |  return a if xb, otherwise return x.  |
|  WTRI(t,p)  |  Triangular wave    |  t = time, p = period  |
|  WSQR(t,p)  |  Square wave    |  t = time, p = period  |
|  WRECT(t,p,d)  |  Rectangular wave    |  t = time, p = period, d= duty-cycle  |
|  WTRAPEZ(t,p,d)  |  Trapez. wave    |  t = time, p = period, d= duty-cycle  |
|  WSAW(t,p)  |  Saw wave    |  t = time, p = period  |
|  WRAISE(t,p)  |  Rampa wave    |  t = time, p = period  |
|  WLIN(t,p,d)  |  Linear wave    |  t = time, p = period, d= duty-cycle  |
|  WPULSE(t,p,d)  |  Rectangular pulse wave    |  t = time, p = period, d= duty-cycle  |
|  WSTEPS(t,p,n)  |  Steps wave    |  t = time, p = period, n = steps number  |
|  WEXP(t,p,a)  |  Exponential pulse wave    |  t = time, p = period, a= dumping factor  |
|  WEXPB(t,p,a)  |  Exponential bipolar pulse wave    |  t = time, p = period, a= dumping factor  |
|  WPULSEF(t,p,a)  |  Filtered pulse wave    |  t = time, p = period, a= dumping factor  |
|  WRING(t,p,a,fm)  |  Ringing wave    |  t = time, p = period, a= dumping factor,  fm = frequency  |
|  WPARAB(t,p)  |  Parabolic pulse wave    |  t = time, p = period  |
|  WRIPPLE(t,p,a)  |  Ripple wave    |  t = time, p = period, a= dumping factor  |
|  WAM(t,fo,fm,m)  |  Amplitude modulation    |  t = time, p = period, fo = carrier freq.,  fm = modulation freq., m = modulation factor  |
|  WFM(t,fo,fm,m)  |  Frequecy modulation    |  t = time, p = period, fo = carrier freq.,  fm = modulation freq., m = modulation factor  |
|  Year(d)  |  year  |  d = dateserial  |
|  Month(d)  |  month  |  d = dateserial  |
|  Day(d)  |  day  |  d = dateserial  |
|  Hour(d)  |  hour  |  d = dateserial  |
|  Minute(d)  |  minute  |  d = dateserial  |
|  Second(d)  |  second  |  d = dateserial  |
|  DateSerial(a,m,d)  |  Dateserial from date  |  a = year, m = month, d = day  |
|  TimeSerial(h,m,s)  |  Timeserial from time  |  h = hour, m = minute, s = second  |
|  time#  |  system time  |     |
|  date#  |  system date  |     |
|  now#  |  system timestamp   |     |
|  Sum(a,b,...)  |  Sum  |  sum(8,9,12,9,7,10) = 55  |
|  Mean(a,b,...)  |  Arithmetic mean  |  mean(8,9,12,9,7,10) = 9.16666666666667  |
|  Meanq(a,b,...)  |  Quadratic mean  |  meanq(8,9,12,9,7,10) = 9.30053761886914  |
|  Meang(a,b,...)  |  Arithmetic mean  |  meang(8,9,12,9,7,10) = 9.03598945281812  |
|  Var(a,b,...)  |  Variance  |  var(1,2,3,4,5,6,7) = 4.66666666666667  |
|  Varp(a,b,...)  |  Variance pop.  |  varp(1,2,3,4,5,6,7) = 4  |
|  Stdev(a,b,...)  |  Standard deviation  |  Stdev(1,2,3,4,5,6,7) = 2.16024689946929  |
|  Stdevp(a,b,...)  |  Standard deviation pop.  |  Stdevp(1,2,3,4,5,6,7) = 2  |
|  Step(x,a)  |  Haveside's step function  |  Returns 1 if x ³ a  , 0 otherwise  |

(°) the operation 0^0 (0 raises to 0) is not allowed here.

Symbol "!" is the same as "Fact" function; symbol "%" is percentage;  symbol "\" is integer division; symbol “| . |” is the same as Abs function

Logical functions and operators return 1 (true) or 0 (false)

From version 4.2 , the argument separator depends on the MathParser decimal separator setting.  If decimal symbol is point "." (i.e. 3.14) , the argument separator is ",". If it is comma "," (i.e. 3,14) , the argument separator is ";".

**Limits of 4.0:**

Max operations/functions = 200; max variables = 100; max function types = 140; max arguments for function = 20; max nested functions = 20; max expressions stored at the same time = undefined. Increasing these limits is easy. For example, if you want to parse long strings up to 500 operations or functions with 60 max arguments, and max 250 variables,  simply set the variable

```
Const HiVT   As Long = 250
Const HiET   As Long = 500
Const HiARG  As Long = 60
```

The ET table can now contain up to 400 rows, each of one is a math operator or function

## Class

This software is structured following the object programming rules and consists of a set of methods and properties

|     |     |     |
| --- | --- | --- |
| **Class** |     |     |
| **clsMathParser** | **Methods** |     |
|     | **StoreExpression(f)** | Stores, parses, and checks syntax errors of the formula “f” |
|     | **Eval** | Evaluates expression |
|     | **Eval1(x)** | Evaluates mono-variable expression f(x) |
|     | **EvalMulti(x(), id)** | Evaluates expression for a vector of values |
|     | **ET_Dump** | Dumps the internal ET table (debug only) |
|     | **Properties** |     |
|     | **Expression** | Gets the current expression stored (R) |
|     | **VarTop** | Gets the top of the var array (R) |
|     | **VarName(i)** | Gets name of variable of index=i (R) |
|     | **Variable(x)** | Sets/gets the value of variable passes by index/name (R/W) |
|     | **AngleUnit** | Sets/gets the angle unit of measure (R/W) |
|     | **ErrorDescription** | Gets error description (R) |
|     | **ErrorID** | Error message number |
|     | **ErrorPos** | Error position |
|     | **OpAssignExplicit** | Option: Sets/gets the explicit variables assignment (R/W) |
|     | **OpUnitConv** | Enable/disable unit conversion |
|     | **OpDMSConv** | Enable/disable DMS angle conversion |

Note. The old properties _VarValue_ and _VarSymb_ are obsolete, being substituted by the _Variable_ property. However, they are still supported for compatibility.

R = read only, R/W = read/write

### Methods

**StoreExpression** Stores, parses, and checks syntax errors. Returns True if no errors are detected; otherwise it returns False

**Syntax**

_object_**. StoreExpression**(ByVal _strExpr_ As String) As Boolean

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _Object_ | Obligatory. It is always the clsMathParser object. |
| _strExpr_ | Math expression to evaluate. |

This method can be invoked after a new instance of the class has been created. It activates the parse routine.

Example

```
Dim OK as Boolean
Dim Fun As New clsMathParser
......

 OK = Fun.StoreExpression(txtFormula)
```

**Notes**

Typical mixed math-physical expressions are

```
1+(2-5)*3+8/(5+3)^2 (a+b)*(a-b) (-1)^(2n+1)*x^n/n! x^2+3*x+1
256.33*Exp(-t/12 us) (3000000 km/s)/144 Mhz 0.25x+3.5y+1 space/time
```

**Variables** can be any alphanumeric string and must start with a letter:

```
x, y, a1, a2, time, alpha , beta
```

**Implicit multiplication** is not supported because of its intrinsic ambiguity. So "xy" stand for a variable named "xy" and not for x\*y. The multiplication symbol "\*" cannot generally be omitted.

It can be omitted only for coefficients of the classic math variables x, y, z. It means that strings like 2x and 2*x are equivalent:

```
2x, 3.141y, 338z^2 Û 2\*x, 3.141\*y, 338*z^2
```

On the contrary, the following expressions are illegal: 2a, 3(x+1), 334omega

**Constant numbers** can be integers, decimal, or exponential:

```
2, -3234, 1.3333, -0.00025, 1.2345E-12
```

**Physical numbers** are alphanumeric strings that must begin with a number:

```
"1s" for 1 second , "200m" for 200 meters, "25kg" for 25 kilograms.
```

For better reading they may contain a blank:

```
"1 s" , "200 m" , "25 kg" , "150 MHz", "0.15 uF", "3600 kOhm"
```

They may also contain the following **multiplying factor**:

```
T=10^12, G=10^9, M=10^6, k=10^3, m=10^-3, u=10^-6, n=10^-9, p=10^-12
```

**Eval**  Evaluates the previously stored expression and returns its value.

**Syntax**

_object_**. Eval**() As Double

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |

This method substitutes the variables values previously stored with the **Variable** property and performs numeric evaluation.

**Notes**

The error detected by this method is any computational domain error. This happens when variable values are out of the domain boundary. Typical domain errors are:

\_\_\_

```
Log(-x), Log(0) , Ö  -x x/0 , arcsin(2)
```

They cannot be intercepted by the parser routine because they are not syntax errors but depend only on the wrong numeric substitution of the variable. When this kind of error is intercepted this method raises an error and its description is copied into the _ErrDescription_ property of the clsMathParser object and into the same property of the global Err object.

  

**Eval1** Evaluates a monovariable function and returns its value

**Syntax**

_object_**. Eval1**(ByVal x As Double) As Double

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |
| _x_ | Variable value |

**Notes**

This method is a simplified version of the general **Eval** method

It is adapted for monovariable function f(x)


```vb
Dim OK As Boolean, x As Double, f As Double
Dim Funct As New clsMathParser
On Error Resume Next
OK = Funct.StoreExpression("(x^2+1)/(x^2-1)") 'parse function
If Not OK Then
    Debug.Print Funct.ErrorDescription
Else
    x = 3
    f = Funct.Eval1(x) 'evaluate function value
    If Err = 0 Then
        Debug.Print "x="; x, "f(x)="; f
    Else
        Debug.Print "x="; x, "f(x)="; Funct.ErrorDescription
    End If
End If
```

Note in this case how the code is simple and compact. This method is also about 11% faster than the general **Eval** method.

**EvalMulti** Substitutes and evaluates a vector of values

**Syntax**

_object_**.EvalMulti**(ByRef VarValue() As Double, Optional ByVal VarName)

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _Object_ | Obligatory. It is always the clsMathParser object. |
| _VarValue()_ | Vector of variable values |
| _VarName_ | Name or index of the variable to be substituted |

**Notes**

This property is very useful to assign a vector to a variable obtaining of a vector values in a very easy way. It is also an efficient method, saving more than 25% of elaboration time.

The formula expression can have several variables, but only one of them can receive the array.

Let’s see this example in which we will compute 10000 values of the same expression in a flash.

Evaluate with the following

```
f(x,y,T,a) = (a^2*exp(x/T)-y)
```

for: x = 0...1 (1000 values), y = 0.123 , T = 0.4 , a = 100

```vb
Sub test_array()
    Dim x() As Double, F, Formula, h
    Dim Loops&, i&
    Dim Funct As New clsMathParser
    '----------------------------------------------
    Formula = "(a^2*exp(x/T)-y)"
    Loops = 10000
    x0 = 0
    x1 = 1
    h = (x1 - x0) / Loops
    'load x-samples
    ReDim x(Loops)
    For i = 1 To Loops
        x(i) = i * h + x0
    Next i
    If Funct.StoreExpression(Formula) Then
        On Error GoTo Error_Handler
        T0 = Timer
        Funct.Variable("y") = 0.123
        Funct.Variable("T") = 0.4
        Funct.Variable("a") = 100
        F = Funct.EvalMulti(x, "x") ‘evaluate in one shot 10000 values
        T1 = Timer - T0
        T2 = T1 / Loops
        Debug.Print T1, T2
    Else
        Debug.Print Funct.ErrorDescription
    End If
Exit Sub

Error_Handler:
    Debug.Print Funct.ErrorDescription
End Sub
```

**ET_Dump** Return the ET internal table (only for debugging)

**Syntax**

_object_**. ET_Dump**(ByRef Etable as Variant)

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |
| ETable | Array containing the ET table |

**Notes**

This method copies the internal ET Table into an array (n x m). The first row (0) contains the column headers. The array must be declares as a Variant undefined array.

```vb
Dim Etable()
```

For details see example 8.

### Properties

**Expression** Returns the current stored expression

**Syntax**

_object_**. Expression**() As String

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |

This property is useful to check if a formula is already stored. In the example below, the main routine skips the parsing step if the function is already stored. This saves a lot of time.

...

...

 If Funct.Expression <> Formula Then

 OK = Funct.StoreExpression(Formula) 'parse function

 End If

....

...

**VarTop** get the top of the variable array

**Syntax**

_object_**. VarTop**() As Long

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |

**Notes**

It returns the total number of different symbolic variables contained in an expression.

This is useful to dimension the variables array dynamically.

Example, for the following expression:

```
"(x^2+1)/(y^2-y-2)"
```

We get

```
_2 = object_.VarTop
```

**VarName** returns the name of the i-th variable.

**Syntax**

_object_**. VarName**(ByVal Index As Long) As String

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |
| _Index_ | Pointer of variable array. It must be within 1 and VarTop |

**Example**

This code prints all variable names contained in an expression, with their current values

```vb
Dim Funct As New clsMathParser
Funct.StoreExpression "(x^2+1)/(y^2-y+2)"
For i = 1 To Funct.VarTop
    Debug.Print Funct.VarName(i); " = "; Funct.Variable(i)
Next
```

**Variable** sets or gets the value of a variable by its symbol or its index.

**Syntax**

_object_**. Variable**(ByVal Name) As Double

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |
| _Name_ | symbolic name or index of the variable. |

Example

Assign the value 2.3333 to the variable of the given index. The index is built from the parser reading the formula from left to right: the first variable encountered has index = 1, the second one has index = 2 and so on.

```vb
Dim f As Double
Dim Funct As New clsMathParser
Funct.StoreExpression "(x^2+1)/(y^2-y+2)" 'parse function
Funct.Variable(1) = 2.3333
Funct.Variable(2) = -0.5
f = Funct.Eval
Debug.Print "f(x,y)=" + Str(f)
But we could write as well
Funct.Variable("x") = 2.3333
Funct.Variable("y") = -0.5
```

Using index is the fastest way to assign a variable value. Normally, it is about 2 times faster.

But, on the contrary, the assignment by symbolic name is easier and more flexible.

**AngleUnit** Sets/gets the angle unit of measure

**Syntax**

_object_**. AngleUnit**  As String

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |
| _AngleUnit_ | Sets/gets the current angle unit (read/write) |

The default unit of measure of angles is “RAD” (Radian), but we can also set “DEG” (sexagesimal degree) and “GRAD” (centesimal degree)

Example

```vb
Dim f As Double
Dim Funct As New clsMathParser
Funct.StoreExpression "sin(x)" 'parse function
Funct.AngleUnit= "DEG"
f = Funct.Eval1(45) ' we get f= 0.707106781186547
f = Funct.Eval1(90) ' we get f= 1
```

The angles returned by the functions are also converted.
Funct.StoreExpression "asin(x)" 'parse function


```vb
Funct.AngleUnit= "DEG"
f = Funct.Eval1(1) ' we get f= 90
f = Funct.Eval1(0.5) ' we get f= 30
```

Note: angle setting only affects the following trigonometric functions:

```
sin(x) , cos(x) , tan(x), atn(x) , acos(x) , asin(x) , csc(x) , sec(x) , cot(x), acsc(x) , asec(x) , acot(x) , rad(x) , deg(x) , grad(x)
```

**ErrDescription** Returns any error detected.

**Syntax**

_object_**. ErrorDescription**() As String

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |
| _ErrorDescription_ | Contains the error message string |

**Note**

This property contains the error description detected by internal routines. They are divided into syntax errors and evaluation errors (or domain errors).

If a domain error is intercepted, an error is generated, so you can also check the global Err object.

For a complete list of error descriptions see “Error  Messages”.

**ErrorID** Returns the error message number.

**Syntax**

_object_**. ErrorID** () As Long

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |
| _ErrorID_ | Returns the error message number |

This property contains the error number detected by internal routines. It is the index of the internal error table ErrorTbl()

For a complete list of error numbers see “Error  Messages”.

**ErrorPos** Returns the error position.

**Syntax**

_object_**. ErrorPos** () As Long

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |
| _ErrorPos_ | Returns the error position |

This property contains the error position detected by internal routines. Note that not always the parser can detec the exact position of the error

**OpAssignExplicit** Enable/disable the explicit variable assignment.

**Syntax**

_object_**. OpAssignExplicit**()  As Boolean

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |
| _OpAssignExplicit_ | True/False. If True, forces the explicit variables assignement. |

All variables contained in a math expression are initialized to zero by default. The evaluation methods usually use these values in the substitution step without checking if the variable values have been initialized by the user or not. If you want to force explicit initialization, set this property to “true”. In that case, the evaluation methods will perform the check: If one variable has never been assigned, the MathParser will raise an error. The default is “false”.

Example. Compute f(x,y) = x^2+y , assigning only the variable x = −3

```vb
Sub test_assignement1()
    Dim x As Double, F, Formula
    Dim Funct As New clsMathParser
    Formula = "x^2+y"
    If Not Funct.StoreExpression(Formula) Then GoTo Error_Handler
    Funct.OpAssignExplicit = False
    Funct.Variable("x") = -3
    F = Funct.Eval
    If Err <> 0 Then GoTo Error_Handler
    Debug.Print "F(x,y)= "; Funct.Expression
    Debug.Print "x= "; Funct.Variable("x")
    Debug.Print "y= "; Funct.Variable("y")
    Debug.Print "F(x,y)= "; F
    Exit Sub
Error_Handler:
    Debug.Print Funct.ErrorDescription
End Sub
```

If _OpAssignExplicit = False_ then the response will be

```
F(x,y)= x^2+y
x= -3
y= 0
F(x,y)= 9
```

As we can see, variable y, never assigned by the main program, has the default value = 0, and the eval method returns 9.

On the contrary, if _OpAssignExplicit = True_ then the response will be the error:

“_Variable &lt;y&gt; not assigned”_

**OpUnitConv** Enable/disable unit conversion.

**Syntax**

_object_**. OpUnitConv** ()  As Boolean

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |
| _OpUnitConv_ | True/False. If True, enables the unit conversion. |

This option switches off the unit conversion if not needed. The default is “True”.

**OpDMSConv** Enable/disable angle DMS conversion.

**Syntax**

_object_**. OpUnitConv** ()  As Boolean

Where:

|     |     |
| --- | --- |
| **Parts** | **Description** |
| _object_ | Obligatory. It is always the clsMathParser object. |
| _OpDMSConv_ | True/False. If True, enables the angle dms conversion. |

This option switches off the angle DMS conversion if not needed. The default is “True”.

## Parser

It is the heart of the algorithm. It reads, character by character, the given expression string and tries to create a structured database.

Parser's algorithm translate between the conventional symbol into a conceptual structured database.

This database, designed with the table ET, contains all the operative information to perform the calculus.

We have to point out that the algorithm used is quite different from others descendent-tree algorithms. For this reason you can find a step-by-step explanation on the pdf document contained into the zip package.

### The Conceptual Model Method

In the language of structured data modeling we can say that this algorithm translates from the input arithmetic or algebraic expression to the follow conceptual data model:

![](/web/20100706055339im_/http://digilander.libero.it/foxes/Image18.gif)

This model expresses the following sentences:

1.  Each arithmetic or algebraic expression is composed by functions. Also the common binary operators like +, -, *, / are considered as functions with two variables. For example: ADD(a,b) instead of a+b , DIV(a,b) instead of a/b and so on.
2.  Each function must have one value and one or more arguments.
3.  Each argument has one value.
4.  A function may be argument of one other function

The algorithm attributes a priority level at each function, based on an inner table-level realizing the usually order of math operators:

|     |     |     |     |     |     |     |
| --- | --- | --- | --- | --- | --- | --- |
| **Level=** | 1   | 2   | 3   | 9   | +10 | -10 |
| **Functions=** | \+ - | \* / | ^   | Any functions | ( \[ { | ) \] } |

This level is increased by 10 every time the parser comes in touch with a left bracket (do not care which type) and, at the opposite, is decreased by 10 with a right bracket.

The following example explains how this algorithm works.  
Suppose to evaluate the expression: _(a+b)*(a-b) ,_ with _a=3_ and _b=5  
_We call the **Parse** routine with the following arguments:

```
ExprString = "(a+b)*(a-b)"  
VarValue(1) = 3 'value of variable a  
VarValue(2) = 5 'value of variable b
```

The parser builds the following table:

|     |     |     |     |     |     |     |     |     |     |     |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| **ID_Fun** | **Fun** | **Level** | **MaxArg** | **Arg1 name** | **Arg1 value** | **Arg2 name** | **Arg2 value** | **ArgOf** | **IndArg** | **Index** |
| 1   | +   | 11  | 2   | a   | -   | b   | -   | 2   | 1   | 1   |
| 2   | *   | 2   | 2   | -   | -   | -   | -   | 0   | 0   | 3   |
| 3   | -   | 11  | 2   | a   | -   | b   | -   | 2   | 2   | 2   |

The Parser reads the string from left to right and substitutes the first variable met with the first value of VarVector(), the second variable with the second value and so on. All the functions recognized by this Parser 2.0 have one or two variables.

The Parser has recognized tree functions (operators: +, *, \- ) attributing the priority level, respectively, of 11, 2, 11 .The operators + and - have normally a level of L=1 (the lowest). Because of brackets their level becomes L=10+1=11.

The function "+" , ID\_Fun=1, has two arguments: "a" and "b", respectively indicated into columns "Arg name". The result of function "+" is the first argument of function "*" (ID\_Fun=2), as indicated in the columns "ArgOf" and "IndArg".

The function "-" , ID\_Fun=3, has two arguments: "a" and "b", respectively indicated into columns "Arg name". The result of function "-" is the second argument of function "*" (ID\_Fun=2), as indicated in the columns "ArgOf" and "IndArg".

Finally, the function "*" is not argument of no other function, as indicate the corresponding ArgOf=0; so his result is the result of the given expression. Note that the Parser has given no values of any arguments.

This task is performed by **Eval** subroutine.

**II° step: Eval()**

This routine acts two simple actions:

1.  substitutes all variables (if present) with their numeric values;
2.  performs the computation of all functions according to their priority levels.

In the above example, this routine inserts into the columns "Arg value" the values 3 and 5 of corresponding variables "a" and "b".

|     |     |     |     |     |     |     |     |     |     |     |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| **ID_Fun** | **Fun** | **Level** | **MaxArg** | **Arg1 name** | **Arg1 value** | **Arg2 name** | **Arg2 value** | **ArgOf** | **IndArg** | **Index** |
| 1   | +   | 11  | 2   | a   | 3   | b   | 5   | 2   | 1   | 1   |
| 2   | *   | 2   | 2   | -   | -   | -   | -   | 0   | 0   | 3   |
| 3   | -   | 11  | 2   | a   | 3   | b   | 5   | 2   | 2   | 2   |

After substitution, the algorithm performs the computation, one function at the time, from the higher level to the lower, and from top to bottom, for function with the same level. In this case the order of computation is 1, 3, 2 as indicated the last column "Index".

The operation performed are in sequence:

1.  3+5= 8 , attributes to the function indexed by the corresponding field ArgOf=2 and by the argument IndArg=1.
2.  3-5= -2 , attributes to the function indexed by the corresponding field ArgOf=2 and by the argument IndArg=2
3.  8*(-2)= -16 , and as ArgOf=0 the routine returns this value as result of expression evaluation.

|     |     |     |     |     |     |     |     |     |     |     |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| **ID_Fun** | **Fun** | **Level** | **MaxArg** | **Arg1 name** | **Arg1 value** | **Arg2 name** | **Arg2 value** | **ArgOf** | **IndArg** | **Index** |
| 1   | +   | 11  | 2   | a   | 3   | b   | 5   | 2   | 1   | 1   |
| 2   | *   | 2   | 2   | -   | 8   | -   | -2  | 0   | 0   | 3   |
| 3   | -   | 11  | 2   | a   | 3   | b   | 5   | 2   | 2   | 2   |

### Error messages


#### ErrDescription Property

The **_StoreExpression_** method returns the status of parsing. It returns TRUE or FALSE, and the property **_ErrDescription_** contains the error message. An empty string means no error. If an error happens, the formula is rejected and it contains one of the messages below.

In addition, the methods _Eval_ and _Eval1_ set the ErrMessage property.

|     |     |     |
| --- | --- | --- |
| **Error Intercepted** | **Description** | **ID** |
| Function &lt;name &gt; unknown at pos: ith | String detected at position ith is not recognized. example “x + foo(x)” | 6   |
| Syntax error at pos: ith | Syntax error at position ith . “a(b+c)”, | 5   |
| Too many closing brackets at pos: ith | Parentheses mismatch: “sqr(a+b+c))” | 7   |
| missing argument | Missing argument for operator or function.”3+” , “sin()” | 8   |
| variable &lt;name&gt; not assigned | Variable “name” has never been assigned. This error rises only if the property AssignExplicit = True | 18  |
| too many variables | Symbolic variables exceed the set limit. Internal constant HiVT =100 sets the limit | 1   |
| Too many arguments at pos: ith | Too many arguments passed to the function | 9   |
| Not enough closing brackets | Parentheses mismatch: “1-exp(-(x^2+y^2)” | 11  |
| abs symbols \|.\| mismatch | Pipe mismatch: “\|x+2\|-\|x-1” | 4   |
| Evaluation error &lt;5 / 0&gt; at pos: 6 Evaluation error &lt;asin(-2)&gt; at pos: 10 Evaluation error &lt;log(-3)&gt; at pos: 6 | The only error returned by the _Eval_ method. It is caused by any mathematical error: 1/0, log(0), etc. | 14,15, 16, 17 |
| constant unknown: &lt;name &gt; | a constant not recognized by the parser | 19  |
| Wrong DMS format | an angle expressed in a wrong DMS format (eg: 2° 34' 67" ) | 21  |
| Too many operations | The expression string contains more than 200 operations/functions | 20  |
| Function <id_function> missing? | Help for developers. Returned by the _Eval__ internal routine. If we see this message we probably have forgotten the relative case statement in eval subroutine! | 13  |
| Variable not found | a variable passed to the parser not found | 2   |

In order to facilitate the modification and/or translation, all messages are collected into an internal table

See "Error Message Table" in PDF documentation for details

#### Error Raise

Methods **Eval**, **Eval1, EvalMulti** and also raise an exception when any evaluation error occurs.

This is useful to activate an error-handling routine. The global property **Err.Description** contains the same error message as the **_ErrDescription_** property

## Computation Time Test

The clsMathParser class was been tested with several formulas in different environment. The response time depends on the number and type of operations performed and, of course, by the computer speed. The following table show synthetically the performance obtained. (Excel 2000, Pentium 3, 1.2 Ghz,)

|     |     |     |     |     |
| --- | --- | --- | --- | --- |
| **Clock => 1200 MHz** | 18/11/2002 | **new version 3** |     |     |
| **Expression** | **operations** | **time  <br>(ms)** | **time/op. (****m****s)** | **Eval. / sec.** |
| **average** | **4.1** | **0.009** | **2.6** | **187,820** |
| `1+(2-5)*3+8/(5+3)^2` | 7   | 0.01 | 1.43 | 100,000 |
| `sqr(2)` | 1   | 0.003 | 3.00 | 333,333 |
| `sqr(x)` | 1   | 0.004 | 4.00 | 250,000 |
| `sqr(4^2+3^2)` | 4   | 0.007 | 1.75 | 142,857 |
| `x^2+3*x+1` | 4   | 0.009 | 2.25 | 111,111 |
| `x^3-x^2+3*x+1` | 6   | 0.014 | 2.33 | 71,429 |
| `x^4-3*x^3+2*x^2-9*x+10` | 10  | 0.022 | 2.20 | 45,455 |
| `(x+1)/(x^2+1)+4/(x^2-1)` | 8   | 0.018 | 2.25 | 55,249 |
| `(1+(2-5)*3+8/(5+3)^2)/sqr(4^2+3^2)` | 12  | 0.017 | 1.42 | 58,824 |
| `sin(1)` | 1   | 0.003 | 3.00 | 333,333 |
| `asin(0.5)` | 1   | 0.004 | 4.00 | 250,000 |
| `fact(10)` | 1   | 0.005 | 5.00 | 200,000 |
| `sin(pi/2)` | 2   | 0.004 | 2.00 | 250,000 |
| `x^n/n!` | 3   | 0.012 | 4.00 | 83,333 |
| `1-exp(-(x^2+y^2))` | 5   | 0.016 | 3.20 | 62,500 |
| `20.3Km+8Km` | 2   | 0.002 | 1.00 | 500,000 |
| `0.1uF*6.8KOhm` | 2   | 0.003 | 1.50 | 333,333 |
| `(1.434E3+1000)*2/3.235E-5` | 3   | 0.005 | 1.67 | 200,000 |

As we can see, the computation time depends strongly from the type of built-in function.

But for an average expression of about 4 operations the speed is greater than 100.000 evaluations in one second.

For one operation, the average evaluation time is about 3 uS

## Source code examples in VB
To explain the use of  this routine a few simple examples in VB are shown. Time performance are obtained with Pentium 3, 1.2GHz, and Excel 2000

### Example 1:

This example show how to evaluates a math polynomial
`p(x) = x^3-x^2+3*x+6`  for 1000  values of  his variable x between x_min = -2 ,  x_max = +2, with step of  Dx = 0.005

 
```vb
Sub Poly_Sample()
    Dim x(1 To 1000) As Double, y(1 To 1000) As Double, OK As Boolean
    Dim txtFormula As String
    Dim Fun As New clsMathParser
    '-----------------------------------------------------------------------
    txtFormula = "x^3-x^2+3*x+6"   'f(x) to sample.
    '----------------------------------------------------------------------
    'Define expression, perform syntax check and get its handle
    
    OK = Fun.StoreExpression(txtFormula)
    If Not OK Then GoTo Error_Handler
    
    'load input values vector.
    For i = 1 To 1000
        x(i) = -2 + 0.005 * (i - 1)
    Next
    
    t0 = Timer
    For i = 1 To 1000
        y(i) = Fun.Eval1(x(i))
        If Err Then GoTo Error_Handler
    Next
    
    Debug.Print Timer - t0  'about 0.015  ms for a CPU with 1200 Mhz
    
    Exit Sub
Error_Handler:
        Debug.Print Fun.ErrorDescription
End Sub
```

We note that the function evaluation – the focal point of this task - is performed in 5 principal statements:

1)      Function declaration, store , and parse
2)      Syntax error check
3)      load variable value
4)      Evaluate (eval)
5)      Activate error trap for checking domain erros

Just clean and straight. No other statement is necessary. This takes advantage overall in complicated math routine, when the main focus must be concentrated on math algorithm, without any other tecnical dispersion and extra subroutines callings. Note also that declaration is need only once at time.

Of course, also the speed computation is important and must be put in evidence

For the above example, with a Pentium III° at 1.2 GHz,  we have got 1000 points of the cubic polynomial in less than 15 ms (1.4E-2), that is a very good performance for this kind of  “interpreted parser” (70.000 points / sec)

Note also that variable name it is not important; the example works fine also with other strings, such as: `t^3-t^2+3*t+6` , `a^3-a^2+3*a+6`, etc.

The parser, simply substitutes the first variables encontered with the value passed.

### Example 2

This example computes the Mc Lauren's series up to 16° order for exp(x) with x=0.5
`(exp(0.5) @ 1.64872127070013)`

```vb
Sub McLaurin_Serie()
    Dim txtFormula As String
    Dim n As Long, N_MAX As Integer, y As Double
    Dim Fun As New clsMathParser
    '-----------------------------------------------------------------------
    txtFormula = "x^n / n!"   'Expression to evaluate. Has two variable
    x0 = 0.5    'set value of Taylor's series expansion
    N_MAX = 16  'set max series expansion
    '----------------------------------------------------------------------
    'Define expression, perform syntax check and get its handle
    OK = Fun.StoreExpression(txtFormula)
    If Not OK Then GoTo Error_Handler
    
    'begin formula evaluation -------------------------
    Fun.VarValue(1) = x0 'load value x
    For n = 0 To N_MAX
        Fun.VarValue(2) = n 'increments the n variables
        If Err Then GoTo Error_Handler
        y = y + Fun.Eval    'accumulates partial i-term
    Next
    Debug.Print y
    Exit Sub
Error_Handler:
    Debug.Print Fun.ErrorDescription
End Sub
```

Returns

```
1.64872127070013
```

Note, also in this case, the very clean code.

### Example 3

This example show how to capture all the variables in a formula with its right sequence (sequence is from left to right).

```vb
Dim Expr As String
Dim OK As Boolean
Dim Fun As New clsMathParser
 
Expr = "(a-b)*(a-b)+30*time/y"
----------------------------------------------------------------------
  'Define expression, perform syntax check and detect all variables
   OK = Fun.StoreExpression(Expr)
----------------------------------------------------------------------
If Not OK Then
    Debug.Print Fun.ErrorDescription  'syntax error detected
Else
    For i = 1 To Fun.VarTop
        Debug.Print Fun.VarName(i), Str(i) + "° variable"
    Next i
End If
```

Output will be:

```
a              1° variable
b              2° variable
time           3° variable
y              4° variable
```


### Example 4

This example show how to evaluate a multivariable expression passing the variables values directly by its symbolic name using the VarSymb property (3.3.2 version or higher)

It evaluate the formula:

`f(x,y,T,a) = (a2×exp(y/T)-x)          for   x = 1.5   ,   y = 0.123   ,   T = 0.4   ,   a = 100`

```vb
Sub test()
    Dim ok As Boolean, f As Double, Formula As String
    Dim Funct As New clsMathParser
    
    Formula = "(a^2*exp(y/T)-x)"
    
    ok = Funct.StoreExpression(Formula)  'parse function
    
    If Not ok Then GoTo Error_Handler
    
    On Error GoTo Error_Handler
    
    'assign value passing the symbolic variable name
    Funct.VarSymb("x") = 1.5
    Funct.VarSymb("y") = 0.123
    Funct.VarSymb("T") = 0.4
    Funct.VarSymb("a") = 100
    
    f = Funct.Eval  'evaluate function
    
    Debug.Print "evaluation = "; f
    
    Exit Sub
Error_Handler:
    Debug.Print Funct.ErrorDescription
End sub  
```

The output, unless of errors,  will be:

```
evaluation =  13598.7080850196
```

Note how plane and compact is the code when using the string-assignment. It is only twice slower that the index-assignment but moreover it looks more simpler.

### Example 5

This examplw shows how to evaluate a function definited by pieces of sub-functions

Each sub-function must be evaluated only for x belongs to its definition domain

{{missing picture}}

Note that you get an error if you will try to calculate 2th and 3rd   sub-functions in points external to theirs domain ranges.

For example, we get an error for x = -1 in 2th and 3td   sub-functions, because we have

`Log(0) = “?”     and SQR(-1-log2)  =”?”`

So, it is indispensable to calculate each function only when it needs.

Here are an example showing how to evaluate a segmented function with the domain constrain explained before.

There are few comments, but the code is very clean and straight.

```vb
Sub Eval_Pieces_Function()
    Dim j&, Value_x#, Value_F#
    Dim xmin#, xmax#, step#, Samples&
    Dim Fun As New clsMathParser
    '--- Piecewise function definition -------------
    '   x^2         for x<=0
    '   log(x+1)    for 0<x<=1
    '   sqr(x-log2) for x>1
    '------------------------------------------------
        f = "(x<=0)* x^2 + (0<x<=1)* Log(x+1) + (x>1)* Sqr(x-Log(2))"
    '------------------------------------------------
    xmin = -2: xmax = 2: Samples = 10
    
    If Not Fun.StoreExpression(f) Then GoTo Error_Handler
    
    Samples = 10
    step = (xmax - xmin) / (Samples - 1)
    On Error GoTo Error_Handler
    For j = 1 To Samples
        Value_x = xmin + (j - 1) * step 'value x
        Fun.VarSymb("x") = Value_x
        Value_F = Fun.Eval
        Debug.Print "x=" + str(Value_x); Tab(25); "f(x)=" + str(Value_F)
    Next j
    Exit Sub
Error_Handler:
    Debug.Print Err.Source, Err.Description
End Sub
```

The output, unless of errors,  will be:

```
x=-2                    f(x)= 4
x=-1.55555555555556     f(x)= 2.41975308641975
x=-1.11111111111111     f(x)= 1.23456790123457
x=-.666666666666667     f(x)= .444444444444445
x=-.222222222222222     f(x)= 4.93827160493828E-02
x= .222222222222222     f(x)= 8.71501757189001E-02
x= .666666666666667     f(x)= .221848749616356
x= 1.11111111111111     f(x)= .900045063009142
x= 1.55555555555556     f(x)= 1.12005605212042
x= 2                    f(x)= 1.30344543588752
```

## Credit
 

MathParser was ideated by

Leonardo Volpi

and developed thanks to the collaboration of

Lieven Dossche, Michael Ruder, Thomas Zeutschler,  Arnaud De Grammont.

 

Many thanks also for their help in debugging and setting up to:

Rodrigo Farinha

Shaun Walker

Iván Vega Rivera

Javie Martin Montalban

Simon de Pressinger

Jakub Zalewski

Sebastián Naccas

RC Brewer

PJ Weng

Mariano Felice

Ricardo Martínez Camacho

Berend Engelbrecht

André Hendriks

Michael Richter

Mirko Sartori

 

Special thanks for the documentation revision to

Mariano Felice

## License
 
clsMathParser is freeware open software. We are happy if you use and promote it. You are granted a free license to use the enclosed software and any associated documentation for personal or commercial purposes, except to sell the original. If you wish to incorporate or modify parts of clsMathParser please give them a different name to avoid confusion. Despite the effort that went into building, there's no warranty, that it is free of bugs. You are allowed to use it at your own risk. Even though it is free, this software and its documentation remain proprietary products. It will be correct (and fine) if you put a reference about the authors in your documentation.
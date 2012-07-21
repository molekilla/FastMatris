Attribute VB_Name = "Diferencial"
Option Explicit


Public Function Euler(FuncX As String, FuncY As String, InitY As Single, HValue As Single, sngValueToCalc As Single) As Single
Dim sngEulerArray() As Single, t As Collection
Dim iRedimValue As Integer, i As Single
Dim Xn As Single, Yn As Single, Yprima As Single, HYprima As Single
On Error GoTo Desborde
'calculate intervals
For i = 0 To sngValueToCalc Step HValue
    iRedimValue = iRedimValue + 1
Next i
ReDim sngEulerArray(iRedimValue, 4)

'col: 1 - Xn; 2 - Yn; 3 - YnPrima; 4 - hYnprima
sngEulerArray(1, 1) = 0
sngEulerArray(1, 2) = InitY
'PARA LA FUNCION DE X
If ((FuncX = "sin") Or (FuncX = "SIN")) Then
    sngEulerArray(1, 3) = CSng(Sin(0)) + (CSng(CalcY(FuncY, InitY)))
ElseIf ((FuncX = "cos") Or (FuncX = "COS")) Then
    sngEulerArray(1, 3) = CSng(Cos(0)) + (CSng(CalcY(FuncY, InitY)))
ElseIf ((CSng(FuncX) > 0) Or (CSng(FuncX) < 0)) Then
    sngEulerArray(1, 3) = (CSng(FuncX) * 0) + (CSng(CalcY(FuncY, InitY)))
    
End If
        sngEulerArray(1, 4) = HValue
For i = 2 To iRedimValue
    sngEulerArray(i, 1) = sngEulerArray(i - 1, 1) + HValue
    sngEulerArray(i, 2) = sngEulerArray(i - 1, 2) + sngEulerArray(i - 1, 4)
    'PARA LA FUNCION DE X
    If ((FuncX = "sin") Or (FuncX = "SIN")) Then
        sngEulerArray(i, 3) = CSng(Sin(sngEulerArray(i, 1))) + (CSng(CalcY(FuncY, InitY)))
    ElseIf ((FuncX = "cos") Or (FuncX = "COS")) Then
        sngEulerArray(i, 3) = CSng(Cos(sngEulerArray(i, 1))) + (CSng(CalcY(FuncY, InitY)))
    ElseIf ((CSng(FuncX) > 0) Or (CSng(FuncX) < 0)) Then
        sngEulerArray(i, 3) = (CSng(FuncX) * sngEulerArray(i, 1)) + (CSng(CalcY(FuncY, sngEulerArray(i, 2))))
        'sngEulerArray(i, 4) = HValue * sngEulerArray(i, 3)
    End If
    sngEulerArray(i, 4) = HValue * sngEulerArray(i, 3)
Next i
    Euler = sngEulerArray(i - 1, 2)
    Set t = New Collection
    Call t.Add(sngEulerArray, "Euler")
    Load frmResults
    Call frmResults.ShowEuler("Euler", t)
Exit Function
Desborde:
If (Err.Number > 0) Then
    MsgBox (Err.Number & " " & Err.Description)
    If (Err.Number = 6) Then
        MsgBox ("Desborde")
    End If
End If
End Function
Public Function EulerModificado(FuncX As String, FuncY As String, InitY As Single, HValue As Single, sngValueToCalc As Single) As Single
On Error GoTo Desborde
Dim sngEulerArray() As Single, t As Collection
Dim iRedimValue As Integer, i As Single
Dim Xn As Single, Yn As Single, Yprima As Single, HYprima As Single

'calculate intervals
For i = 0 To sngValueToCalc Step HValue
    iRedimValue = iRedimValue + 1
Next i
ReDim sngEulerArray(iRedimValue, 8)

'col: 1 - Xn; 2 - Yn; 3 - YnPrima; 4 - hYnprima
sngEulerArray(1, 1) = 0
sngEulerArray(1, 2) = InitY
'PARA LA FUNCION DE X
If ((FuncX = "sin") Or (FuncX = "SIN")) Then
    sngEulerArray(1, 3) = CSng(Sin(0)) + (CSng(CalcY(FuncY, InitY)))
ElseIf ((FuncX = "cos") Or (FuncX = "COS")) Then
    sngEulerArray(1, 3) = CSng(Cos(0)) + (CSng(CalcY(FuncY, InitY)))
ElseIf ((CSng(FuncX) > 0) Or (CSng(FuncX) < 0)) Then
    sngEulerArray(1, 3) = (CSng(FuncX) * 0) + (CSng(CalcY(FuncY, InitY)))
    
End If
sngEulerArray(1, 4) = HValue * sngEulerArray(1, 3)
sngEulerArray(1, 5) = sngEulerArray(1, 2) + sngEulerArray(1, 4)
sngEulerArray(1, 6) = sngEulerArray(1, 5) + sngEulerArray(1, 4)
sngEulerArray(1, 7) = (sngEulerArray(1, 3) + sngEulerArray(1, 6)) / 2
sngEulerArray(1, 8) = HValue * sngEulerArray(1, 7)


For i = 2 To iRedimValue
    sngEulerArray(i, 1) = sngEulerArray(i - 1, 1) + HValue
    sngEulerArray(i, 2) = sngEulerArray(i - 1, 2) + sngEulerArray(i - 1, 8)
    'PARA LA FUNCION DE X
    If ((FuncX = "sin") Or (FuncX = "SIN")) Then
        sngEulerArray(i, 3) = CSng(Sin(sngEulerArray(i, 1))) + (CSng(CalcY(FuncY, InitY)))
    ElseIf ((FuncX = "cos") Or (FuncX = "COS")) Then
        sngEulerArray(i, 3) = CSng(Cos(sngEulerArray(i, 1))) + (CSng(CalcY(FuncY, InitY)))
    ElseIf ((CSng(FuncX) > 0) Or (CSng(FuncX) < 0)) Then
        sngEulerArray(i, 3) = (CSng(FuncX) * sngEulerArray(i, 1)) + (CSng(CalcY(FuncY, sngEulerArray(i, 2))))
        'sngEulerArray(i, 4) = HValue * sngEulerArray(i, 3)
    End If
    sngEulerArray(i, 4) = HValue * sngEulerArray(i, 3)
    sngEulerArray(i, 5) = sngEulerArray(i, 2) + sngEulerArray(i, 4)
    sngEulerArray(i, 6) = sngEulerArray(i, 5) + (sngEulerArray(i, 4) + HValue)
    sngEulerArray(i, 7) = (sngEulerArray(i, 3) + sngEulerArray(i, 6)) / 2
    sngEulerArray(i, 8) = HValue * sngEulerArray(i, 7)
Next i
EulerModificado = sngEulerArray(i - 1, 2)
    Set t = New Collection
    Call t.Add(sngEulerArray, "EulerM")
    Load frmResults
    Call frmResults.ShowEuler("EulerM", t)

Exit Function
Desborde:
If (Err.Number > 0) Then
    MsgBox (Err.Number & " " & Err.Description)
    If (Err.Number = 6) Then
        MsgBox ("Desborde")
    End If
End If

End Function
Private Function CalcFunction(FuncionX$, FuncionY$, iX As Single, iY As Single) As Single
If ((FuncionX = "sin") Or (FuncionX = "SIN")) Then
    CalcFunction = CSng(Sin(iX)) + (CSng(CalcY(FuncionY, iY)))
ElseIf ((FuncionX = "cos") Or (FuncionX = "COS")) Then
    CalcFunction = CSng(Cos(iX)) + (CSng(CalcY(FuncionY, iY)))
ElseIf ((CSng(FuncionX) > 0) Or (CSng(FuncionX) < 0)) Then
    CalcFunction = (CSng(FuncionX) * iX) + (CSng(CalcY(FuncionY, iY)))
    
End If

End Function
Private Function CalcY(FuncionY As String, iY As Single) As Double
'PARA LA FUNCION DE Y
If ((FuncionY = "sin") Or (FuncionY = "SIN")) Then
    CalcY = Sin(CDbl(iY))
ElseIf ((FuncionY = "cos") Or (FuncionY = "COS")) Then
        CalcY = Cos(CDbl(iY))
ElseIf ((CSng(FuncionY) > 0) Or (CSng(FuncionY) < 0)) Then
   CalcY = CDbl(FuncionY) * iY
End If
End Function

Public Function RungeKuttaIV(FuncX As String, FuncY As String, InitY As Single, HValue As Single, sngValueToCalc As Single) As Single
'Dim Ynext As Single
Dim Yn As Single, i As Single
Dim K1 As Single, K2 As Single, K3 As Single, K4 As Single
Dim Xn As Single

Yn = InitY
For i = 0 To sngValueToCalc Step HValue
Xn = i
'K1
K1 = HValue * (CalcFunction(FuncX, FuncY, Xn, Yn))
'K2
K2 = HValue * (CalcFunction(FuncX, FuncY, (Xn + (0.5 * HValue)), (Yn + (0.5 * K1))))
'K3
K3 = HValue * (CalcFunction(FuncX, FuncY, (Xn + (0.5 * HValue)), (Yn + (0.5 * K2))))
'K4
K4 = HValue * (CalcFunction(FuncX, FuncY, (Xn + HValue), (Yn + K3)))

Yn = Yn + ((1 / 6) * (K1 + (2 * K2) + (2 * K3) + K4))

Next i
RungeKuttaIV = Yn
End Function
Public Function RungeKuttaIII(FuncX As String, FuncY As String, InitY As Single, HValue As Single, sngValueToCalc As Single) As Single
'Dim Ynext As Single
Dim Yn As Single, i As Single
Dim K1 As Single, K2 As Single, K3 As Single, K4 As Single
Dim Xn As Single

Yn = InitY
For i = 0 To sngValueToCalc Step HValue
Xn = i
'K1
K1 = HValue * (CalcFunction(FuncX, FuncY, Xn, Yn))
'K2
K2 = HValue * (CalcFunction(FuncX, FuncY, (Xn + (0.5 * HValue)), (Yn + (0.5 * K1))))
'K3
K3 = HValue * (CalcFunction(FuncX, FuncY, (Xn + HValue), (Yn + (-K1 + (2 * K2)))))
'K4
'K4 = HValue * (CalcFunction(FuncX, FuncY, (Xn + HValue), (Yn + K3)))

Yn = Yn + ((1 / 6) * (K1 + (4 * K2) + K3))

Next i
RungeKuttaIII = Yn
End Function
Public Function RungeKuttaII(FuncX As String, FuncY As String, InitY As Single, HValue As Single, sngValueToCalc As Single) As Single
'Dim Ynext As Single
Dim Yn As Single, i As Single
Dim K1 As Single, K2 As Single, K3 As Single, K4 As Single
Dim Xn As Single

Yn = InitY
For i = 0 To sngValueToCalc Step HValue
Xn = i
'K1
K1 = HValue * (CalcFunction(FuncX, FuncY, Xn, Yn))
'K2
'K2 = HValue * (CalcFunction(FuncX, FuncY, (Xn + (0.5 * HValue)), (Yn + (0.5 * K1))))
'K3
'K3 = HValue * (CalcFunction(FuncX, FuncY, (Xn + (0.5 * HValue)), (Yn + (0.5 * K2))))
'K2
K2 = HValue * (CalcFunction(FuncX, FuncY, (Xn + HValue), (Yn + K1)))

Yn = Yn + ((1 / 2) * (K1 + K2))

Next i
RungeKuttaII = Yn
End Function



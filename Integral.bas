Attribute VB_Name = "Integral"
Option Explicit
Private sNameFunc As String
Private objFunc As New FunctionX



Public Sub ParseFunction(sFunction As String)
'Dim sFunction As String
Dim iLen As Integer, k%, j%, i%, lastOpPos%
Dim bNeg As Boolean, sngNum As Single, sngExp As Single, sOps$, bVar As Boolean
On Error GoTo ShowError
bNeg = False
'sFunction = objdyn.All.FX.Value
If sFunction = "" Then
    MsgBox ("Inserte una función!")
    Exit Sub
End If
sNameFunc = sFunction
iLen = Len(sFunction)
'Initialize object
Set objFunc = Nothing

    If (Mid(sFunction, 1, 1) = "-") And (Not IsNumeric(Mid(sFunction, 2, 1))) Then
        'Valor Negativo
        bNeg = True
        i = 2
    End If
i = 1

For i = i To iLen


        If (Mid(sFunction, i, 1) = "x" Or Mid(sFunction, i, 1) = "X") Then
        'constante
        '----------
            If i <> 1 Then
                If (Mid(sFunction, i - 1, 1) = "+" Or Mid(sFunction, i - 1, 1) = "-") Then
                If bNeg = True Then
                    sngNum = -1
                    bNeg = False
                Else
                    sngNum = 1
                End If
                Else
                        sngNum = Mid(sFunction, lastOpPos + 1, i - lastOpPos - 1)
                End If
            Else
                    sngNum = 1
            End If
        '----------
        bVar = True
        '---------------------------------------------
            If Mid(sFunction, i + 1, 1) = "^" Then
                'exponente
                For k = i To iLen
                    j = 1
                    If (Mid(sFunction, k, 1) = "+" Or Mid(sFunction, k, 1) = "-") Then
                        sngExp = Mid(sFunction, k - 1, j)
                        Exit For
                    End If
                    j = j + 1
                Next k
            ElseIf (Mid(sFunction, i + 1, 1) = "+" Or Mid(sFunction, i + 1, 1) = "-") Then
                'x , exponente es 1
                sngExp = 1
            End If
        '--------------------------------------------
        End If
        If ((Mid(sFunction, i, 1) = "+" Or Mid(sFunction, i, 1) = "-")) And i > 1 Then
            'operador
            sOps = Mid(sFunction, i, 1)
            lastOpPos = i
            Call objFunc.Add(sngNum, bVar, sngExp, sOps)
        End If
Next i
'operador
sngNum = Mid(sFunction, lastOpPos + 1, i - lastOpPos - 1)
bVar = False
sngExp = 1
sOps = "!"
Call objFunc.Add(sngNum, bVar, sngExp, sOps)
'MsgBox (FX(1))
ShowError:
If (Err.Number > 0) Then
    MsgBox ("Inserte una función valida y que termine en numero entero/real")
End If
End Sub
Private Function FuncX(W As Single) As Single
Dim i%, sngEquation As Single, sngEq As Single, BackValue%
BackValue = objFunc.Count


For i = objFunc.Count To 1 Step -1

If i = BackValue Then
    sngEquation = objFunc.Item(i).sngConstant
    If (objFunc.Item(i - 1).sOperator = "-") Then
        sngEquation = -1 * sngEquation
    End If
    
Else
'        If (objFunc.Item(i).sngConstant > 0) Then
            sngEq = objFunc.Item(i).sngConstant
'        End If
        If objFunc.Item(i).bVariable = True Then
            sngEq = sngEq * (W ^ (objFunc.Item(i).sngExponent))
        End If
        
        If i = 1 Then
            sngEquation = sngEq + sngEquation
        Else
                 If (objFunc.Item(i - 1).sOperator = "+") Then
                    sngEquation = sngEq + sngEquation
                 ElseIf (objFunc(i - 1).sOperator = "-") Then
                    sngEquation = (-1 * sngEq) + sngEquation
                 End If
         End If
    End If
Next i
FuncX = sngEquation
End Function

Public Function Trapecio(iBLimit As Single, iALimit As Single, N As Integer) As Single
Dim H As Single
Dim sngHTrapecio As Single
Dim iEnd As Single, iX As Single
Dim sngValue As Single

H = (iBLimit - iALimit) / N
iX = iX + H
sngHTrapecio = H / 2
iEnd = iBLimit - H
'Fo, first F(X)
sngValue = FuncX(iALimit)
Do
DoEvents
'    iX = iX + H
    sngValue = sngValue + (2 * FuncX(iX))
        iX = iX + H
    iEnd = iEnd - H
Loop Until (iEnd <= 0)
'Fn, last value of F(X)
sngValue = sngValue + FuncX(iBLimit)
'assign value, asignar valor
Trapecio = (sngValue * sngHTrapecio)
End Function
Public Function SimpsonPar(iBLimit As Single, iALimit As Single, N As Integer) As Single
Dim H As Single, i%
Dim sngHTrapecio As Single
Dim iEnd As Single, iX As Single
Dim sngValue As Single

H = (iBLimit - iALimit) / N
iX = iX + H
i = 1
sngHTrapecio = H / 3
iEnd = iBLimit - H
'Fo, first F(X)
sngValue = FuncX(iALimit)
Do
DoEvents
'    iX = iX + H
If (i Mod 2) = 0 Then
    sngValue = sngValue + (2 * FuncX(iX))
Else
    sngValue = sngValue + (4 * FuncX(iX))
End If
    iX = iX + H
    iEnd = iEnd - H
    i = i + 1
Loop Until (iEnd <= 0)
'Fn, last value of F(X)
sngValue = sngValue + FuncX(iBLimit)
'assign value, asignar valor
SimpsonPar = (sngValue * sngHTrapecio)
End Function

Public Function Simpson3(iBLimit As Single, iALimit As Single, N As Integer) As Single
Dim H As Single, i%
Dim sngHTrapecio As Single
Dim iEnd As Single, iX As Single
Dim sngValue As Single

H = (iBLimit - iALimit) / N
iX = iX + H
i = 1
sngHTrapecio = (3 / 8) * H
iEnd = iBLimit - H
'Fo, first F(X)
sngValue = FuncX(iALimit)
Do
DoEvents
If (i = N) Then
    Exit Do
End If
    If (i Mod 3) = 0 Then
        sngValue = sngValue + (2 * FuncX(iX))
    Else
        sngValue = sngValue + (3 * FuncX(iX))
    End If
    iX = iX + H
    iEnd = iEnd - H
    i = i + 1
Loop Until (iEnd <= 0#)
'Fn, last value of F(X)
sngValue = sngValue + FuncX(iBLimit)
'assign value, asignar valor
Simpson3 = (sngValue * sngHTrapecio)
End Function


Public Function Bairstow1(sFunction$, iP%, iQ%) As Variant
'--------------------------BAIRSTOW-----------------------
'---------------------------------------------------------
'By Rogelio Morrell December 1998
'Examples: if n=4 then Cmax=n and Bmax=n+1
'CREATE A array
On Error GoTo ThisError
Dim iMax As Integer
Dim MainArray() As Single
Dim i%

ParseFunction (sFunction)
'MAX EXPONENT
iMax = CInt(objFunc.Item(1).sngExponent)
ReDim MainArray((iMax + 1), 5)
'Insert P's y Q's values
MainArray(1, 1) = CSng(iP)
MainArray(1, 2) = CSng(iQ)
'Insertar columna A
For i = 1 To (iMax + 1)
        If (i = 1) Then
            MainArray(i, 3) = objFunc.Item(i).sngConstant
        Else
            If (objFunc(i - 1).sOperator = "-") Then
                MainArray(i, 3) = -1 * objFunc.Item(i).sngConstant
            Else
                MainArray(i, 3) = objFunc.Item(i).sngConstant
            End If
        End If
    If (i = 1) Then
        'Columna B
        MainArray(i, 4) = MainArray(i, 3)
        'Columna C
        MainArray(i, 5) = MainArray(i, 4)
    ElseIf (i = 2) Then
        'Columna B
        MainArray(i, 4) = MainArray(i, 3) + (MainArray(1, 1) * MainArray(i - 1, 4)) ' + (MainArray(1, 2) * MainArray(i - 2, 4))
        'Columna C
        MainArray(i, 5) = MainArray(i, 4) + (MainArray(1, 1) * MainArray(i - 1, 5)) ' + (MainArray(1, 2) * MainArray(i - 2, 5))
    Else
        'Columna B
        MainArray(i, 4) = MainArray(i, 3) + (MainArray(1, 1) * MainArray(i - 1, 4)) + (MainArray(1, 2) * MainArray(i - 2, 4))
        'Columna C
        MainArray(i, 5) = MainArray(i, 4) + (MainArray(1, 1) * MainArray(i - 1, 5)) + (MainArray(1, 2) * MainArray(i - 2, 5))
    End If
Next i
'Columna B
MainArray(i - 1, 5) = 0 ' MainArray(i, 3) + (MainArray(1, 1) * MainArray(i - 1, 4)) + (MainArray(1, 2) * MainArray(i - 2, 4))

Bairstow1 = MainArray
Exit Function
ThisError:
If (Err.Number > 0) Then
    MsgBox (Err.Number & " " & Err.Description)
    If (Err.Number = 6) Then
        MsgBox ("Desborde")
    End If
End If
End Function

Public Function Bairstow2(sFunc$, objTabla As Variant) As Variant
'--------------------------BAIRSTOW-----------------------
'---------------------------------------------------------
'By Rogelio Morrell December 1998
'Examples: if n=4 then Cmax=n and Bmax=n+1
'CREATE A array
On Error GoTo ThisError
Dim iMax As Integer
Dim MainArray As Variant
Dim i%, DeltaP As Single, DeltaQ As Single
Dim N As Integer, P1 As Single, P2 As Single, Q1 As Single, Q2 As Single

ParseFunction (sFunc)

iMax = UBound(objTabla)
N = iMax - 1
MainArray = objTabla
'Delta P
P1 = ((-MainArray(N, 4) * MainArray(N - 1, 5) - (-MainArray(N + 1, 4) * MainArray(N - 2, 5))))
P2 = ((MainArray(N - 1, 5) * MainArray(N - 1, 5) - (MainArray(N, 5) * MainArray(N - 2, 5))))
DeltaP = P1 / P2
'Delta Q
Q1 = ((MainArray(N - 1, 5) * -MainArray(N + 1, 4) - (MainArray(N, 5) * -MainArray(N, 4))))
Q2 = P2
DeltaQ = Q1 / Q2

'Insert P's y Q's values
MainArray(1, 1) = MainArray(1, 1) + DeltaP
MainArray(1, 2) = MainArray(1, 2) + DeltaQ
'Insertar columna A
For i = 1 To (iMax)
        If (i = 1) Then
            MainArray(i, 3) = objFunc.Item(i).sngConstant
        Else
            If (objFunc(i - 1).sOperator = "-") Then
                MainArray(i, 3) = -1 * objFunc.Item(i).sngConstant
            Else
                MainArray(i, 3) = objFunc.Item(i).sngConstant
            End If
        End If
    If (i = 1) Then
        'Columna B
        MainArray(i, 4) = MainArray(i, 3)
        'Columna C
        MainArray(i, 5) = MainArray(i, 4)
    ElseIf (i = 2) Then
        'Columna B
        MainArray(i, 4) = MainArray(i, 3) + (MainArray(1, 1) * MainArray(i - 1, 4)) ' + (MainArray(1, 2) * MainArray(i - 2, 4))
        'Columna C
        MainArray(i, 5) = MainArray(i, 4) + (MainArray(1, 1) * MainArray(i - 1, 5)) ' + (MainArray(1, 2) * MainArray(i - 2, 5))
    Else
        'Columna B
        MainArray(i, 4) = MainArray(i, 3) + (MainArray(1, 1) * MainArray(i - 1, 4)) + (MainArray(1, 2) * MainArray(i - 2, 4))
        'Columna C
        MainArray(i, 5) = MainArray(i, 4) + (MainArray(1, 1) * MainArray(i - 1, 5)) + (MainArray(1, 2) * MainArray(i - 2, 5))
    End If
Next i
'Columna B
MainArray(i - 1, 5) = 0 ' MainArray(i, 3) + (MainArray(1, 1) * MainArray(i - 1, 4)) + (MainArray(1, 2) * MainArray(i - 2, 4))

Bairstow2 = MainArray
Exit Function
ThisError:
If (Err.Number > 0) Then
    MsgBox (Err.Number & " " & Err.Description)
    If (Err.Number = 6) Then
        MsgBox ("Desborde")
    End If
End If
End Function


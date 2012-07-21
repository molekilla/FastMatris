Attribute VB_Name = "RootEquations"
Option Explicit
Private imgCopy As Object
Private sNameFunc As String
Private iLeft%, iRight%
Private objY As New FunctionX


Public Sub MakeGraph(objDom1 As Object)
Attribute MakeGraph.VB_Description = "sText as Variant"
On Error GoTo ShowD
Dim sDA As Object
Dim surfaceIMG As Object
Dim i%
Dim X1 As Single, X2 As Single, y1 As Single, y2 As Single

Set sDA = objDom1.All.Graph.Library
'new surface
Set surfaceIMG = objDom1.All.Graph.DrawingSurface
surfaceIMG.LineColor (sDA.Black)

For i = -100 To 100 Step 10
X1 = i
y1 = 0
'---------------------------
X2 = (i + 1)
y2 = 0
'x-line
    If i = 0 Then
        surfaceIMG.LineColor (sDA.Blue)
    Else
        surfaceIMG.LineColor (sDA.Black)
    End If
    Call surfaceIMG.LinePoints(objDom1.All.Graph.Library.Point2(-100, i), objDom1.All.Graph.Library.Point2(100, i))
    'y-line
    Call surfaceIMG.LinePoints(objDom1.All.Graph.Library.Point2(i, -100), objDom1.All.Graph.Library.Point2(i, 100))
Next i
surfaceIMG.LineColor (sDA.Red)
surfaceIMG.LineJoinStyle (0) 'Bevel Type
For i = -70 To 70
DoEvents
X1 = (i)  '; //* 0.1
y1 = FX(X1)

X2 = i + 1
y2 = FX(X2)
Call surfaceIMG.LinePoints(objDom1.All.Graph.Library.Point2(X1, y1), objDom1.All.Graph.Library.Point2(X2, y2))
Next i

'//Call surfaceIMG.Scale2(1.1, 1.1)
objDom1.All.Graph.Image = surfaceIMG.Image
'Hacer copia del objeto imagen para usar en MoveLine
Set imgCopy = surfaceIMG.Image
ShowD:
If (Err.Number > 0) Then
    MsgBox (Err.Description)
End If
'Visualizar funcion actual en la barra de funcion
objDom1.All.FunctionName.innerText = "Función F(x): " & sNameFunc
End Sub

Public Function FX(W As Single) As Single
Dim i%, sngEquation As Single, sngEq As Single, BackValue%
BackValue = objY.Count


For i = objY.Count To 1 Step -1

If i = BackValue Then
    sngEquation = objY.Item(i).sngConstant
    If (objY.Item(i - 1).sOperator = "-") Then
        sngEquation = -1 * sngEquation
    End If
    
Else
'        If (objY.Item(i).sngConstant > 0) Then
            sngEq = objY.Item(i).sngConstant
'        End If
        If objY.Item(i).bVariable = True Then
            sngEq = sngEq * (W ^ (objY.Item(i).sngExponent))
        End If
        
        If i = 1 Then
            sngEquation = sngEq + sngEquation
        Else
                 If (objY.Item(i - 1).sOperator = "+") Then
                    sngEquation = sngEq + sngEquation
                 ElseIf (objY(i - 1).sOperator = "-") Then
                    sngEquation = (-1 * sngEq) + sngEquation
                 End If
         End If
    End If
Next i
FX = sngEquation
End Function

Public Sub ParsingFunction(objdyn As Object)
Dim sFunction As String
Dim iLen As Integer, k%, j%, i%, lastOpPos%
Dim bNeg As Boolean, sngNum As Single, sngExp As Single, sOps$, bVar As Boolean
On Error GoTo ShowError
bNeg = False

' Function String
sFunction = objdyn.All.FX.Value

' Validation
If sFunction = "" Then
    MsgBox ("Inserte una función!")
    Exit Sub
End If

sNameFunc = sFunction
iLen = Len(sFunction)
'Initialize object
Set objY = Nothing

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
            Call objY.Add(sngNum, bVar, sngExp, sOps)
        End If
Next i
'operador
sngNum = Mid(sFunction, lastOpPos + 1, i - lastOpPos - 1)
bVar = False
sngExp = 1
sOps = "!"
Call objY.Add(sngNum, bVar, sngExp, sOps)
'MsgBox (FX(1))
ShowError:
If (Err.Number > 0) Then
    MsgBox ("Inserte una función valida y que termine en numero entero/real")
End If
End Sub

Public Sub GraphXY(objDyn1 As Object, objDyn2 As Object)
Call ParsingFunction(objDyn2)
Call MakeGraph(objDyn1)
End Sub

Public Sub MoveLineRight(objD1 As Object)
Dim sDA As Object
Dim surfaceIMG As Object
'Static iRight%
'Dim X1 As Single, X2 As Single, Y1 As Single, Y2 As Single

Set sDA = objD1.All.Graph.Library
'Original Surface
Set surfaceIMG = objD1.All.Graph.DrawingSurface
'Clear/Erase/Limpiar drawing surface
surfaceIMG.Clear
'Insertar imagen de Overlay/Layer copia original de la funcion
Call surfaceIMG.OverlayImage(imgCopy)
'Dibujar linea
surfaceIMG.LineColor (sDA.Green)
Call surfaceIMG.LinePoints(objD1.All.Graph.Library.Point2(iRight, -150), objD1.All.Graph.Library.Point2(iRight, 150))
iRight = iRight + 1
If iRight = 99 Then
    iRight = 0
End If
objD1.All.Graph.Image = surfaceIMG.Image
'Call MakeGraph(objD1)
'Visualizar X y Y
objD1.All.Xvalue.innerText = "X= " & CStr(iRight)
objD1.All.Yvalue.innerText = "F(X)= " & CStr(FX(CSng(iRight)))
End Sub
Public Sub MoveLineLeft(objD1 As Object)
Dim sDA As Object
Dim surfaceIMG As Object
'Static iLeft%
'Dim X1 As Single, X2 As Single, Y1 As Single, Y2 As Single

Set sDA = objD1.All.Graph.Library
'Original Surface
Set surfaceIMG = objD1.All.Graph.DrawingSurface
'Clear/Erase/Limpiar drawing surface
surfaceIMG.Clear
'Insertar imagen de Overlay/Layer copia original de la funcion
Call surfaceIMG.OverlayImage(imgCopy)
'Dibujar linea
surfaceIMG.LineColor (sDA.Green)
Call surfaceIMG.LinePoints(objD1.All.Graph.Library.Point2(iLeft, -150), objD1.All.Graph.Library.Point2(iLeft, 150))
iLeft = iLeft - 1
If iLeft = -99 Then
    iLeft = 0
End If
objD1.All.Graph.Image = surfaceIMG.Image
'Call MakeGraph(objD1)
'Visualizar X y Y
objD1.All.Xvalue.innerText = "X= " & CStr(iLeft)
objD1.All.Yvalue.innerText = "F(X)= " & CStr(FX(CSng(iLeft)))
End Sub


Public Sub Noll(objD As Object)
iLeft = 0
iRight = 0
'Visualizar X y Y
objD.All.Xvalue.innerText = "X= " & CStr(iLeft)
objD.All.Yvalue.innerText = "F(X)= " & CStr(FX(CSng(iLeft)))
End Sub
Public Sub RFMod(objDom2 As Object)
'iv = integer or single value
Dim y1 As Single, y2 As Single, yC As Single, ivA As Single, ivB As Single, ivC As Single
Dim k As Integer, iTer%, sngMod As Single, tempFC As Single
Dim objColl As Collection, objArray() As Single
Dim g As Integer
If (objDom2.All.X1.Value = "" Or objDom2.All.X2.Value = "") Then
    Exit Sub
End If

ParsingFunction objDom2


ivA = objDom2.All.X1.Value
ivC = objDom2.All.X2.Value
g = 1

iTer = objDom2.All.itern.Value

ReDim objArray(iTer, 7)
Set objColl = New Collection

For k = 1 To iTer
If g > 1 Then
    If (objArray(k - 1, 7) = tempFC) Then
        'Debug.Print (objArray(k - 1, 7) / 2)
        sngMod = tempFC / 2
        ivB = ivC - ((sngMod / (sngMod - FX(ivA))) * (ivC - ivA))
        y1 = FX(ivA)
        y2 = sngMod
        yC = FX(ivB)
        tempFC = y2
    Else
        ivB = ivC - ((FX(ivC) / (FX(ivC) - FX(ivA))) * (ivC - ivA))
        y1 = FX(ivA)
        y2 = FX(ivC)
        yC = FX(ivB)
        tempFC = y2
    End If
Else
'SOLO ENTRA UNA VEZ
    ivB = ivC - ((FX(ivC) / (FX(ivC) - FX(ivA))) * (ivC - ivA))
    y1 = FX(ivA)
    y2 = FX(ivC)
    yC = FX(ivB)
    tempFC = y2
End If

'----------------------
objArray(k, 1) = k      'iteracion
objArray(k, 2) = ivA    'valor de A
objArray(k, 3) = ivB    'valor de B
objArray(k, 4) = ivC    'valor de C
objArray(k, 5) = y1     'valor de y1
objArray(k, 6) = yC     'valor de y2
objArray(k, 7) = y2     'valor de yC
'-----------------------
    If (y1 * yC) < 0 Then
        ivC = ivB
        g = 1
        Else
        ivA = ivB
        g = g + 1
    End If
Next k
    'add array to collection for encapsulation
    objColl.Add (objArray)
    'ladda in formen
    Load frmResults
    'visa alla resultat
    Call frmResults.ShowRootResults(objColl)
    Set objColl = Nothing


End Sub

Public Sub RegulaFalsi(objDom2 As Object)
'iv = integer or single value
Dim y1 As Single, y2 As Single, yC As Single, ivA As Single, ivB As Single, ivC As Single
Dim k As Integer, iTer%
Dim objColl As Collection, objArray() As Single

If (objDom2.All.X1.Value = "" Or objDom2.All.X2.Value = "") Then
    Exit Sub
End If

ParsingFunction objDom2


ivA = objDom2.All.X1.Value
ivC = objDom2.All.X2.Value


iTer = objDom2.All.itern.Value

ReDim objArray(iTer, 7)
Set objColl = New Collection

For k = 1 To iTer
    ivB = ivC - ((FX(ivC) / (FX(ivC) - FX(ivA))) * (ivC - ivA))
    y1 = FX(ivA)
    y2 = FX(ivC)
    yC = FX(ivB)

'----------------------
objArray(k, 1) = k      'iteracion
objArray(k, 2) = ivA    'valor de A
objArray(k, 3) = ivB    'valor de B
objArray(k, 4) = ivC    'valor de C
objArray(k, 5) = y1     'valor de y1
objArray(k, 6) = yC     'valor de y2
objArray(k, 7) = y2     'valor de yC
'-----------------------
    If (y1 * yC) < 0 Then
        ivC = ivB
        Else
        ivA = ivB
    End If
Next k
    'add array to collection for encapsulation
    objColl.Add (objArray)
    'ladda in formen
    Load frmResults
    'visa alla resultat
    Call frmResults.ShowRootResults(objColl)
    Set objColl = Nothing


End Sub

Public Sub Bisection(objDom2 As Object)
'iv = integer or single value
Dim y1 As Single, y2 As Single, yC As Single, ivA As Single, ivB As Single, ivC As Single
Dim k As Integer, iTer%
Dim objColl As Collection, objArray() As Single

'Validation of X1 and X2
If (objDom2.All.X1.Value = "" Or objDom2.All.X2.Value = "") Then
    Exit Sub
End If

'Parse Function
ParsingFunction objDom2

' Values
ivA = objDom2.All.X1.Value
ivC = objDom2.All.X2.Value
iTer = objDom2.All.itern.Value

ReDim objArray(iTer, 7)
Set objColl = New Collection

For k = 1 To iTer
    ivB = (ivA + ivC) / 2
    y1 = FX(ivA)
    y2 = FX(ivC)
    yC = FX(ivB)

'----------------------
objArray(k, 1) = k      'iteracion
objArray(k, 2) = ivA    'valor de A
objArray(k, 3) = ivB    'valor de B
objArray(k, 4) = ivC    'valor de C
objArray(k, 5) = y1     'valor de y1
objArray(k, 6) = y2     'valor de y2
objArray(k, 7) = yC     'valor de yC
'-----------------------
    If (y1 * yC) < 0 Then
        ivC = ivB
        Else
        ivA = ivB
    End If
Next k
'***************END OF ALGORITHM*******'

' To Result Window
    'add array to collection for encapsulation
    objColl.Add (objArray)
    'ladda in formen
    Load frmResults
    'visa alla resultat
    Call frmResults.ShowRootResults(objColl)
    Set objColl = Nothing
End Sub
Public Sub Secante(objDom1 As Object)
On Error GoTo ReadyToGo
'iv = integer or single value
Dim Xn_1 As Single, Xn_2 As Single, Yn_1 As Single, Yn_2 As Single, Xn As Single, Yn As Single
Dim k As Integer, iTer%
Dim objColl As Collection, objArray() As Single
'Dim objTempArr As Variant

If (objDom1.All.X1.Value = "" Or objDom1.All.X2.Value = "") Then
    Exit Sub
End If

ParsingFunction objDom1


Xn_1 = objDom1.All.X1.Value
Xn_2 = objDom1.All.X2.Value
iTer = objDom1.All.itern.Value

ReDim objArray(iTer, 3)
'-----X1------------
objArray(1, 1) = 1
objArray(1, 2) = Xn_1
objArray(1, 3) = FX(Xn_1)
'-----X2------------
objArray(2, 1) = 2
objArray(2, 2) = Xn_2
objArray(2, 3) = FX(Xn_2)

Set objColl = New Collection
'-----------------------------------------------------
For k = 3 To iTer
    Xn_1 = objArray(k - 1, 2)   'Xn-i
    Xn_2 = objArray(k - 2, 2)   'Xn-ii
    Yn_1 = objArray(k - 1, 3)   'Yn-i
    Yn_2 = objArray(k - 2, 3)   'Yn-ii
    Xn = Xn_1 - (Yn_1 * ((Xn_1 - Xn_2) / (Yn_1 - Yn_2)))
    Yn = FX(Xn)
'----------------------
objArray(k, 1) = k      'iteracion
objArray(k, 2) = Xn    'valor de A
objArray(k, 3) = Yn    'valor de B
'-----------------------
If (objArray(k, 2) = objArray(k - 1, 2)) Then
    Exit For
End If
Next k
'-----------------------------------------------------
ReadyToGo:
If (Err.Number > 0) Then
    MsgBox (Err.Description & " " & Err.Number)
    Exit Sub
End If
    'add array to collection for encapsulation
    objColl.Add (objArray)
    'ladda in formen
    Load frmResults
    'visa alla resultat
    Call frmResults.ShowRootResults(objColl)
    Set objColl = Nothing

End Sub
Public Sub Clear(objDom1 As Object)
Dim sDA As Object
Dim surfaceIMG As Object
Set sDA = objDom1.All.Graph.Library
'Original Surface
Set surfaceIMG = objDom1.All.Graph.DrawingSurface
'Clear/Erase/Limpiar drawing surface
surfaceIMG.Clear
objDom1.All.Graph.Image = surfaceIMG.Image

End Sub

Public Sub Newton(objDom2 As Object)
On Error GoTo ReadyToGo
'iv = integer or single value
Dim Xn As Single, y1 As Single, y2 As Single, Xn1 As Single
Dim k As Integer, iTer%
Dim objColl As Collection, objArray() As Single

'Error trap if both boxes or just one are empty
If (objDom2.All.X1.Value = "") Then
    Exit Sub
End If

ParsingFunction objDom2

Xn = objDom2.All.X1.Value
'ivC = objDom2.All.X2.Value
iTer = objDom2.All.itern.Value

ReDim objArray(iTer, 4)
'-----X1------------
objArray(1, 1) = 1
objArray(1, 2) = Xn
objArray(1, 3) = FX(Xn)
objArray(1, 4) = Derivada(Xn)
'---------------------
Set objColl = New Collection

For k = 2 To iTer
    'FUNCION DE NEWTON
    Xn1 = Xn - (objArray(k - 1, 3) / objArray(k - 1, 4))
    'FUNCION
    y1 = FX(Xn1)
    'DERIVADA
    y2 = Derivada(Xn1)
    Xn = Xn1
'----------------------
objArray(k, 1) = k      'iteracion
objArray(k, 2) = Xn    'Xn
objArray(k, 3) = y1    'funcion de X
objArray(k, 4) = y2    'derivada de X
'-----------------------
If (objArray(k, 2) = objArray(k - 1, 2)) Then
    Exit For
End If
    Next k
    'add array to collection for encapsulation
    objColl.Add (objArray)
    'ladda in formen
    Load frmResults
    'visa alla resultat
    Call frmResults.ShowRootResults(objColl)
    Set objColl = Nothing

ReadyToGo:
If (Err.Number > 0) Then
    MsgBox (Err.Description & " " & Err.Number)
    'Exit Sub
End If
End Sub



Private Function Derivada(W As Single) As Single
Dim objDerivada As New FunctionX
Dim i%, r%, sngEquation As Single, sngEq As Single, BackValue%
Dim x As Single, y As Boolean, Z As Single, v As String
BackValue = objY.Count

'--------------------------
'COPIAR OBJETO ORIGINAL A UN NUEVO
'--------------------------
For r = 1 To objY.Count
    
    x = objY(r).sngConstant
    y = objY(r).bVariable
    Z = objY(r).sngExponent
    v = objY(r).sOperator
    Call objDerivada.Add(x, y, Z, v)
Next r
'-----------------
'CALCULAR DERIVADA
'-----------------
For r = 1 To BackValue
objDerivada(r).sngConstant = objDerivada(r).sngConstant * objDerivada(r).sngExponent
objDerivada(r).sngExponent = objDerivada(r).sngExponent - 1
If objDerivada(r).sngExponent = 0 Then
    objDerivada(r).bVariable = False
End If
If objY(r).sOperator = "!" Then
    'objDerivada(r).sngConstant = 0
    objDerivada(r - 1).sOperator = "!"
    objDerivada.Remove (r)
End If
Next r
'-----------------------------------------------------------
'CALCULAR F'(X)
'-----------------------------------------------------------
BackValue = objDerivada.Count
For i = objDerivada.Count To 1 Step -1

If i = BackValue Then
    sngEquation = objDerivada.Item(i).sngConstant
    If (objDerivada.Item(i - 1).sOperator = "-") Then
        sngEquation = -1 * sngEquation
    End If
    
Else
'        If (objDerivada.Item(i).sngConstant > 0) Then
            sngEq = objDerivada.Item(i).sngConstant
'        End If
        If objDerivada.Item(i).bVariable = True Then
            sngEq = sngEq * (W ^ (objDerivada.Item(i).sngExponent))
        End If
        
        If i = 1 Then
            sngEquation = sngEq + sngEquation
        Else
                 If (objDerivada.Item(i - 1).sOperator = "+") Then
                    sngEquation = sngEq + sngEquation
                 ElseIf (objDerivada(i - 1).sOperator = "-") Then
                    sngEquation = (-1 * sngEq) + sngEquation
                 End If
         End If
    End If
Next i
Derivada = sngEquation



End Function

Public Sub InsertFunc(objD1 As Object, objD2 As Object)

Dim test As Object

If objD1.All.ComboCell.innerHTML = "" Then
    objD1.All.ComboCell.innerHTML = "<strong><font face='Arial'>Funciones:</font></strong><select size='1' name='combo' id='combo' style='font-family: Arial; font-size: 10pt; background-color: rgb(128,128,128); color: rgb(0,0,0); border: medium groove'></select>"
    

    Set test = objD1.createElement("OPTION")
    test.Text = objD2.All.FX.Value
    test.Value = objD2.All.FX.Value
    Call objD1.All.combo.Add(test)
    Call objD1.All.combo.Remove(1)
Else
    Set test = objD1.createElement("OPTION")
    test.Text = objD2.All.FX.Value
    test.Value = objD2.All.FX.Value
    Call objD1.All.combo.Add(test)
End If
End Sub

Public Sub GetFunc(objD1 As Object, objD2 As Object)


objD2.All.FX.Value = objD1.All.combo.Value

End Sub

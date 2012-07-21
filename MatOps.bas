Attribute VB_Name = "MatOps"
Option Explicit

Private objMatCol As New Matrices


Public Sub SetMemMat(sMat As String, objDOM As Object)
'Dim t As Object
Dim valor$, iDim%
Dim i%, t As Object

'sMat es el nombre de la tabla
'determinar dimension
i = 0
Set t = objDOM.All.Item(CVar(sMat)).rows(i)
    For Each t In objDOM.All(CVar(sMat)).rows
        valor = t.rowIndex
        i = i + 1
        Set t = objDOM.All(CVar(sMat)).rows(i)
    Next t
    iDim = i - 1
Call SaveArray(iDim, sMat, objDOM)



End Sub

Private Sub SaveArray(iD%, sTab$, objDyDom As Object)
Dim mArray() As Single
Dim i%, j%, ii%, jj%
Dim objTable As Object
Dim oCol As New Collection

ReDim mArray(iD, iD)

Set objTable = objDyDom.All(CVar(sTab))


For i = 1 To iD
    ii = i
    For j = 1 To iD
    jj = j - 1
        mArray(i, j) = CSng(objTable.rows(ii).cells(jj).innerText)
    Next j
    
Next i
Call oCol.Add(mArray(), sTab)
Call objMatCol.Add(iD, sTab, sTab, oCol)
Debug.Print objMatCol.Item(sTab).MatArray(sTab)(1, 1)
End Sub

Public Sub CalcMat(sSign$, sTName1$, sTName2$, objDynamicMod As Object, Optional sTemp As String)
    If (sSign = "+") Then
         Call CalcSum(sTName1, sTName2, sTemp)
    ElseIf (sSign = "-") Then
            Call CalcRest(sTName1, sTName2, sTemp)
        ElseIf (sSign = "*") Then
            Call CalcMulty(sTName1, sTName2, sTemp)
    End If
End Sub

Private Sub CalcSum(sNamn1$, sNamn2$, Optional sVariable$)
Dim objTable1
Dim objTable2
Dim i%, j%, t As New Collection
Dim MaxU1%, MaxU2%
'asignando MATRICES
 objTable1 = objMatCol.Item(sNamn1).MatArray(sNamn1)
 objTable2 = objMatCol.Item(sNamn2).MatArray(sNamn2)
  '--------------------------------------------------
 '--------------------------------------------------
 If DetectDim(sNamn1, sNamn2) Then
    MsgBox ("Inserte matrices con dimensiones iguales!")
    Exit Sub
 End If
 '--------------------------------------------------
 '--------------------------------------------------
'dimension
MaxU1 = UBound(objTable1)
For i = 1 To MaxU1
  
    For j = 1 To MaxU1
    'suma delas matrices
    objTable1(i, j) = objTable1(i, j) + objTable2(i, j)
    Next j
    
Next i
If sVariable = "" Then
    Call t.Add(objTable1, "Suma")
    Load frmResults
    Call frmResults.ShowOpResults("Suma", t)
    
Else
    Call t.Add(objTable1, sVariable)
    Call objMatCol.Add(MaxU1, sVariable, sVariable, t)
End If


End Sub

Public Sub Noll()
Set objMatCol = Nothing
End Sub

Private Sub CalcRest(sNamn1$, sNamn2$, Optional sVariable$)
Dim objTable1
Dim objTable2
Dim i%, j%, t As New Collection
Dim MaxU1%, MaxU2%
'asignando MATRICES
 objTable1 = objMatCol.Item(sNamn1).MatArray(sNamn1)
 objTable2 = objMatCol.Item(sNamn2).MatArray(sNamn2)
  '--------------------------------------------------
 '--------------------------------------------------
 If DetectDim(sNamn1, sNamn2) Then
    MsgBox ("Inserte matrices con dimensiones iguales!")
    Exit Sub
 End If
 '--------------------------------------------------
 '--------------------------------------------------
'dimension
MaxU1 = UBound(objTable1)
For i = 1 To MaxU1
  
    For j = 1 To MaxU1
    'resta delas matrices
    objTable1(i, j) = objTable1(i, j) - objTable2(i, j)
    Next j
    
Next i
If sVariable = "" Then
    Call t.Add(objTable1, "Resta")
    Load frmResults
    Call frmResults.ShowOpResults("Resta", t)
    
Else
    Call t.Add(objTable1, sVariable)
    Call objMatCol.Add(MaxU1, sVariable, sVariable, t)
End If


End Sub
Private Sub CalcMulty(sNamn1$, sNamn2$, Optional sVariable$)
Dim objTable1
Dim objTable2
Dim objTable3
Dim i%, j%, t As New Collection
Dim MaxU%, k%
'asignando MATRICES
 objTable1 = objMatCol.Item(sNamn1).MatArray(sNamn1)
 objTable2 = objMatCol.Item(sNamn2).MatArray(sNamn2)
 '--------------------------------------------------
 '--------------------------------------------------
 If DetectDim(sNamn1, sNamn2) Then
    MsgBox ("Inserte matrices con dimensiones iguales!")
    Exit Sub
 End If
 '--------------------------------------------------
 '--------------------------------------------------
 'dimension
'cualquiera, se usara como temp
objTable3 = objMatCol.Item(sNamn2).MatArray(sNamn2)
'-----------------------------------------
'-----------------------------------------
MaxU = UBound(objTable1)
For i = 1 To MaxU
  
    For j = 1 To MaxU
    'Multiplicacion delas matrices
    objTable3(i, j) = 0
        For k = 1 To MaxU
     
            objTable3(i, j) = objTable3(i, j) + (objTable1(i, k) * objTable2(k, j))
        Next k
    Next j
    
Next i
'-----------------------------------------
'-----------------------------------------
If sVariable = "" Then
    Call t.Add(objTable3, "Multi")
    Load frmResults
    Call frmResults.ShowOpResults("Multi", t)
    
Else
    Call t.Add(objTable3, sVariable)
    Call objMatCol.Add(MaxU, sVariable, sVariable, t)
End If


End Sub
Public Function CalcDeterminante(sTabName$) As Single
Dim objMatris As New Collection
Dim result As Single

'Matriz
Call objMatris.Add(objMatCol.Item(sTabName).MatArray(sTabName), sTabName)
'Llamar a Gauss
Call Gauss(sTabName, objMatris, result)
CalcDeterminante = result

End Function



Private Sub Gauss(sTabName$, objArray As Collection, sngDet As Single, Optional objMatr As Collection, Optional objValues As Collection)
Dim iPivot%, i%, j%, x%, y%, u%, v%
'Dim sngDet As Single
Dim Matr, r As Single, iDim%, tm As Single
Dim temp As Single
Dim eps As Single, eps2 As Single
'Dim objMatr As Collection



eps = 1#
Do
    eps = eps / 2#
Loop Until (1 + eps = 1)
'MsgBox (eps)
eps = eps * 2

eps2 = eps * 2

sngDet = 1
Matr = objArray.Item(sTabName)
iDim = UBound(Matr)

For i = 1 To iDim
    iPivot = i
    
    For j = i To iDim
        If (Abs(Matr(iPivot, i)) < Abs(Matr(j, i))) Then
            iPivot = j
        End If
    Next j
    
If (iPivot <> i) Then
    For y = 1 To iDim
        tm = Matr(i, y)
        Matr(i, y) = Matr(iPivot, y)
        Matr(iPivot, y) = tm
    Next y
    sngDet = -1 * sngDet
End If
If (Matr(i, i) <> 0) Then

    For x = i + 1 To iDim
        If (Matr(x, i) <> 0) Then
            r = Matr(x, i) / Matr(i, i)
            For v = i To iDim
                temp = Matr(x, v)
                Matr(x, v) = Matr(x, v) - r * Matr(i, v) 'calculo principal
                If (Abs(Matr(x, v)) < eps2 * temp) Then
                    Matr(x, v) = 0
                End If
            Next v
        End If
    Next x
Else
    MsgBox ("Singular!")
End If
Next i
For u = 1 To iDim
    sngDet = sngDet * Matr(u, u)
Next u


Set objMatr = New Collection
Call objMatr.Add(Matr, CStr(sTabName))

End Sub

Private Sub Jordan(sObjName$, objMat1 As Collection, objMat2 As Collection)
Dim iDim%, j%
Dim i%, y%, sngTemp As Single
Dim MatArray, xj%

MatArray = objMat1(sObjName)
iDim = UBound(MatArray)
xj = iDim

For y = 1 To iDim
'MatArray(iDim - 1, iDim) = MatArray(iDim - 1, iDim) / MatArray(iDim - 1, iDim - 1)
For i = (iDim) To 1 Step -1
sngTemp = MatArray(i, iDim)
    For j = i + 1 To iDim
        sngTemp = sngTemp - (MatArray(i, j) * MatArray(j, iDim)) '+1
    Next j
    MatArray(i, iDim) = sngTemp / MatArray(i, i)
Next i
iDim = iDim - 1
Next y

Set objMat2 = New Collection
Call objMat2.Add(MatArray, sObjName)


End Sub

Public Sub CalcGauss(sTabName$, Optional t As Integer, Optional tValues As Collection)
Dim objMatris As New Collection
Dim result As Single

'Matriz
Call objMatris.Add(objMatCol.Item(sTabName).MatArray(sTabName), sTabName)
If t = 2 Then
    Call GaussEsp(sTabName, tValues, result, tValues)
    Load frmResults
    frmResults.Show
    Call frmResults.ShowOpResultsEsp(sTabName, tValues)
Else
    Call Gauss(sTabName, objMatris, result, objMatris)
    Load frmResults
    Call frmResults.ShowOpResults(sTabName, objMatris)
End If

End Sub
Public Sub CalcJordan(sTabName$, Optional t As Integer, Optional tValues As Collection)
Dim objMatris As New Collection
Dim result As Single

'Matriz
Call objMatris.Add(objMatCol.Item(sTabName).MatArray(sTabName), sTabName)

If t = 2 Then
    Call GaussEsp(sTabName, tValues, result, tValues)
    Call JordanEsp(sTabName, tValues, tValues)
    Load frmResults
    frmResults.Show
    Call frmResults.ShowOpResultsEsp(sTabName, tValues)
Else
Call Gauss(sTabName, objMatris, result, objMatris)
Call Jordan(sTabName, objMatris, objMatris)

Load frmResults
Call frmResults.ShowOpResults(sTabName, objMatris)
End If

'Set objMatris = Nothing
End Sub
Private Sub GaussEsp(sTabName$, objArray As Collection, sngDet As Single, Optional objMatr As Collection, Optional objValues As Collection)
Dim iPivot%, i%, j%, x%, y%, u%, v%
'Dim sngDet As Single
Dim Matr, r As Single, iDim%, tm As Single
Dim temp As Single
Dim eps As Single, eps2 As Single
'Dim objMatr As Collection



eps = 1#
Do
    eps = eps / 2#
Loop Until (1 + eps = 1)
'MsgBox (eps)
eps = eps * 2

eps2 = eps * 2

sngDet = 1
Matr = objArray.Item(sTabName)
iDim = UBound(Matr)

For i = 1 To iDim
    iPivot = i
    
    For j = i To iDim
        If (Abs(Matr(iPivot, i)) < Abs(Matr(j, i))) Then
            iPivot = j
        End If
    Next j
    
If (iPivot <> i) Then
    For y = 1 To iDim + 1
        tm = Matr(i, y)
        Matr(i, y) = Matr(iPivot, y)
        Matr(iPivot, y) = tm
    Next y
    sngDet = -1 * sngDet
End If
If (Matr(i, i) <> 0) Then

    For x = i + 1 To iDim
        If (Matr(x, i) <> 0) Then
            r = Matr(x, i) / Matr(i, i)
            '------------------------
            For v = i To iDim + 1
                temp = Matr(x, v)
                Matr(x, v) = Matr(x, v) - r * Matr(i, v) 'calculo principal
                If (Abs(Matr(x, v)) < eps2 * temp) Then
                    Matr(x, v) = 0
                End If
            Next v
            '-------------------------
        End If
    Next x
Else
    MsgBox ("Singular!")
End If
Next i
For u = 1 To iDim
    sngDet = sngDet * Matr(u, u)
Next u


Set objMatr = New Collection
Call objMatr.Add(Matr, CStr(sTabName))

End Sub
Private Sub JordanEsp(sObjName$, objMat1 As Collection, objMat2 As Collection)
Dim iDim%, j%
Dim i%, y%, sngTemp As Single
Dim MatArray, xj%

MatArray = objMat1(sObjName)
iDim = UBound(MatArray)
xj = iDim

For y = 1 To iDim
'MatArray(iDim - 1, iDim) = MatArray(iDim - 1, iDim) / MatArray(iDim - 1, iDim - 1)
For i = (iDim) To 1 Step -1
sngTemp = MatArray(i, iDim + 1)
    For j = i + 1 To iDim
        sngTemp = sngTemp - (MatArray(i, j) * MatArray(j, iDim + 1)) '+1
    Next j
    MatArray(i, iDim + 1) = sngTemp / MatArray(i, i)
Next i
iDim = iDim - 1
Next y
'Mostrar bonito, ya know what I mean!....
For i = 1 To xj
    For j = 1 To xj
        If (i = j) Then
            MatArray(i, j) = 1
            Else
            MatArray(i, j) = 0
        End If
    Next j
Next i

Set objMat2 = New Collection
Call objMat2.Add(MatArray, sObjName)


End Sub

Private Function DetectDim(sUno1 As String, sDos2 As String) As Boolean
Dim objTable1, objTable2

objTable1 = objMatCol.Item(sUno1).MatArray(sUno1)
objTable2 = objMatCol.Item(sDos2).MatArray(sDos2)
 If UBound(objTable1) = UBound(objTable2) Then
    DetectDim = False
    Else
    DetectDim = True
 End If
End Function

Public Function GetMemMat(i%) As String
GetMemMat = objMatCol.Item(i).Key

End Function

Public Function GetMemMat2(n$) As Collection
Dim oTable, t As Collection
oTable = objMatCol(n).MatArray(n)
Set t = New Collection
Call t.Add(oTable, n)
Set GetMemMat2 = t
Set t = Nothing

End Function
Public Function CountMemMat() As Integer
CountMemMat = objMatCol.Count
End Function


Public Sub CalcSeidel(sTabName$, tValues As Collection, iTer%)
Dim tErr As New Collection
    Call Seidel(sTabName, tValues, tValues, iTer, tErr)
    Load frmResults
    frmResults.Show
    Call frmResults.ResultsXYZ(sTabName, tValues, tErr)
Set tErr = Nothing
End Sub
Public Sub CalcJacobi(sTabName$, tValues As Collection, iTer%)
Dim tErr As New Collection
    Call Jacobi(sTabName, tValues, tValues, iTer, tErr)
    Load frmResults
    frmResults.Show
    Call frmResults.ResultsXYZ(sTabName, tValues, tErr)
Set tErr = Nothing
End Sub
Private Sub Seidel(sNombre$, MatArray As Collection, MatResult As Collection, iIter%, objError As Collection)
Dim MVi() As Single, MError() As Single, objMVi As Collection
Dim Xi%, y%, iDimension
Dim iSum As Single, MatA, A%, B%
Dim one%, bDominant As Boolean, tm As Single
Dim iCheck%

'iIter = 30
iDimension = UBound(MatArray(sNombre))
ReDim MVi(iDimension)
ReDim MError(iDimension)
Set objMVi = New Collection
'''''''''''CODIGO DE DOMINANCIA CREADO POR ROGELIO MORRELL; RUSH MOLEKILLA
''''ME DIO PEREZA HACERLE UNA FUNCION SORRY Y'ALL
'--------------------------------------------------
MatA = MatArray(sNombre)
Do Until (bDominant)
For A = 1 To iDimension - 1

    For B = 1 To iDimension
        If (A <> B) Then
            iSum = iSum + MatA(A, B)
        End If
    Next B
    If (MatA(A, A) > iSum) Then
        'es dominante
        bDominant = True
        Exit For
    Else
        For one = 1 To iDimension + 1
            tm = MatA(A + 1, one)
            MatA(A + 1, one) = MatA(A, one)
            MatA(A, one) = tm
        Next one
    End If
Next A
iSum = 0
iCheck = iCheck + 1
If iCheck = 200 Then
    MsgBox ("Error, this is a security message fo Fast Matris, to much time!")
Exit Sub
End If
Loop
'asignar a MatArray
MatArray.Remove (sNombre)
Call MatArray.Add(MatA, sNombre)
'-----------------------------------------------------
    For Xi = 1 To iIter
    
    Call objMVi.Add(MVi, sNombre)
        For y = 1 To UBound(MVi)
            MError(y) = MVi(y)
            MVi(y) = fn(sNombre, MatArray, objMVi, y)
            MError(y) = Abs(MError(y) - MVi(y))
        Next y
        objMVi.Remove (sNombre)
    Next Xi
    MatResult.Remove (sNombre)
    Call MatResult.Add(MVi, sNombre)
    
    Call objError.Add(MError, sNombre)
End Sub
Private Sub Jacobi(sNombre$, MatArray As Collection, MatResult As Collection, iIter%, objError As Collection)
Dim MVi() As Single, MError() As Single, objMVi As Collection
Dim Xi%, y%, iDimension
Dim iSum As Single, MatA, A%, B%
Dim one%, bDominant As Boolean, tm As Single
Dim iCheck%

'iIter = 30
iDimension = UBound(MatArray(sNombre))
ReDim MVi(iDimension)
ReDim MError(iDimension)
Set objMVi = New Collection
'''''''''''CODIGO DE DOMINANCIA CREADO POR ROGELIO MORRELL; RUSH MOLEKILLA
''''ME DIO PEREZA HACERLE UNA FUNCION SORRY Y'ALL
'--------------------------------------------------
MatA = MatArray(sNombre)
Do Until (bDominant)
For A = 1 To iDimension - 1

    For B = 1 To iDimension
        If (A <> B) Then
            iSum = iSum + MatA(A, B)
        End If
    Next B
    If (MatA(A, A) > iSum) Then
        'es dominante
        bDominant = True
        Exit For
    Else
        For one = 1 To iDimension + 1
            tm = MatA(A + 1, one)
            MatA(A + 1, one) = MatA(A, one)
            MatA(A, one) = tm
        Next one
    End If
Next A
iSum = 0
iCheck = iCheck + 1
If iCheck = 200 Then
    MsgBox ("Error, this is a security message for Fast Matris, to much time!")
End If
Loop
'asignar a MatArray
MatArray.Remove (sNombre)
Call MatArray.Add(MatA, sNombre)
'-----------------------------------------------------
    For Xi = 1 To iIter
    
    Call objMVi.Add(MVi, sNombre)
        For y = 1 To UBound(MVi)
            MError(y) = MVi(y)
            MVi(y) = fn(sNombre, MatArray, objMVi, y)
            MError(y) = Abs(MError(y) - MVi(y))
            '---------------------------------
            objMVi.Remove (sNombre)
            Call objMVi.Add(MVi, sNombre)
            '---------------------------------
        Next y
        objMVi.Remove (sNombre)
    Next Xi
    MatResult.Remove (sNombre)
    Call MatResult.Add(MVi, sNombre)
    Call objError.Add(MError, sNombre)
End Sub

Private Function fn(sName$, objMat As Collection, Vi As Collection, iExcludeValue As Integer) As Single
Dim MatVi, MatA, iDim%, sngTemp As Single
Dim i%, j%

iDim = UBound(objMat(sName))
MatA = objMat(sName)
MatVi = Vi(sName)
sngTemp = MatA(iExcludeValue, iDim + 1)
    'For i = 1 To iDim
        For j = 1 To iDim
            If iExcludeValue <> j Then
                sngTemp = sngTemp - (MatA(iExcludeValue, j)) * MatVi(j)
            End If
            'Exit For
        Next j
    'Next i
    sngTemp = sngTemp / MatA(iExcludeValue, iExcludeValue)
    fn = sngTemp
End Function

Public Sub InvMat(sTabName$)
Dim iDim%, i%, MatA, j%
Dim objMatris As New Collection
Dim objMatris2 As New Collection
Dim sngNone As Single
'Matriz
Call objMatris.Add(objMatCol.Item(sTabName).MatArray(sTabName), sTabName)
iDim = UBound(objMatris(sTabName))
MatA = objMatris(sTabName)
ReDim Preserve MatA(iDim, iDim + iDim)



For i = 1 To (iDim)
    For j = (iDim + 1) To (iDim + iDim)
    
    If (i = (j - iDim)) Then
        MatA(i, j) = 1
        Else
        MatA(i, j) = 0
    End If

    Next j
Next i
objMatris.Remove (sTabName)
Call objMatris.Add(MatA, sTabName)
Call GaussInv(sTabName, objMatris, sngNone, objMatris2)
Call JordanInv(sTabName, objMatris2, objMatris2)
    Load frmResults
    frmResults.Show
    Call frmResults.ResultsInvMat(sTabName, objMatris2)
End Sub
Private Sub GaussInv(sTabName$, objArray As Collection, sngDet As Single, Optional objMatr As Collection, Optional objValues As Collection)
Dim iPivot%, i%, j%, x%, y%, u%, v%
'Dim sngDet As Single
Dim Matr, r As Single, iDim%, tm As Single
Dim temp As Single
Dim eps As Single, eps2 As Single
'Dim objMatr As Collection



eps = 1#
Do
    eps = eps / 2#
Loop Until (1 + eps = 1)
'MsgBox (eps)
eps = eps * 2

eps2 = eps * 2

sngDet = 1
Matr = objArray.Item(sTabName)
iDim = UBound(Matr)

For i = 1 To iDim
    iPivot = i
    
    For j = i To iDim
        If (Abs(Matr(iPivot, i)) < Abs(Matr(j, i))) Then
            iPivot = j
        End If
    Next j
    
If (iPivot <> i) Then
    For y = 1 To iDim * 2
        tm = Matr(i, y)
        Matr(i, y) = Matr(iPivot, y)
        Matr(iPivot, y) = tm
    Next y
    sngDet = -1 * sngDet
End If
If (Matr(i, i) <> 0) Then

    For x = i + 1 To iDim
        If (Matr(x, i) <> 0) Then
            r = Matr(x, i) / Matr(i, i)
            '------------------------
            For v = i To iDim * 2
                temp = Matr(x, v)
                Matr(x, v) = Matr(x, v) - r * Matr(i, v) 'calculo principal
                If (Abs(Matr(x, v)) < eps2 * temp) Then
                    Matr(x, v) = 0
                End If
            Next v
            '-------------------------
        End If
    Next x
Else
    MsgBox ("Singular!")
End If
Next i
'For u = 1 To iDim
'    sngDet = sngDet * Matr(u, u)
'Next u


Set objMatr = New Collection
Call objMatr.Add(Matr, CStr(sTabName))

End Sub

Private Sub JordanInv(sObjName$, objMat1 As Collection, objMat2 As Collection)
Dim iDim%, j%
Dim i%, y%, sngTemp As Single
Dim MatArray, xj%

MatArray = objMat1(sObjName)
iDim = UBound(MatArray)
xj = iDim * 2

For y = 1 To iDim * 2
'MatArray(iDim, y) = MatArray(iDim, y) / MatArray(iDim, iDim)
For i = (iDim) To 1 Step -1
If (iDim \ 2 = 0) Then
sngTemp = MatArray(i, xj)
Else
sngTemp = MatArray(i, xj)
End If
    For j = i + 1 To iDim
        sngTemp = sngTemp - (MatArray(i, j) * MatArray(j, xj)) '+1
    Next j
    MatArray(i, xj) = sngTemp / MatArray(i, i)
Next i
xj = xj - 1
'iDim = iDim - 1
Next y


Set objMat2 = New Collection
Call objMat2.Add(MatArray, sObjName)


End Sub


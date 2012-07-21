Attribute VB_Name = "AJCurvas"

Public Sub RegresionSup(iPolyGrade%, objArray As Collection, sngValues As Variant)
Dim iValueArray() As Single
Dim sngX As Variant, sngY As Variant
Dim jj%, m%, k%, j%, i%, JEX%, yy As Single
Dim objRegSup As Collection

ReDim iValueArray(iPolyGrade + 1, iPolyGrade + 2) '3 + 1 y de largo y 3 + 2 de x longitud

 sngY = objArray.Item(2)
 sngX = objArray.Item(1)
m = iPolyGrade + 1
For k = 1 To m
    For j = 1 To m + 1
        iValueArray(k, j) = 0#
    Next j
Next k
'iMaxPoints es el numero de puntos dados
iMaxPoints = UBound(sngX)
For k = 1 To m
    For i = 1 To iMaxPoints
        For j = 1 To m
            jj = k - 1 + j - 1
            yy = 1#
            If (jj <> 0) Then
                yy = sngX(i) ^ jj
            End If
                iValueArray(k, j) = iValueArray(k, j) + yy
        Next j
        JEX = k - 1
        yy = 1#
        If (JEX <> 0) Then
            yy = sngX(i) ^ JEX
        End If
        iValueArray(k, m + 1) = iValueArray(k, m + 1) + (sngY(i) * yy)
    Next i
Next k

Set objRegSup = New Collection
Call objRegSup.Add(iValueArray, "RegSup")

'Enviar matris a Gauss-Jordan para calcular valores de la ecuacion G(x)
Call GaussEsp("RegSup", objRegSup, yy, objRegSup, objRegSup)
Call JordanEsp("RegSup", objRegSup, objRegSup)
sngValues = objRegSup.Item("RegSup")

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


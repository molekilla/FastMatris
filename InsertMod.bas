Attribute VB_Name = "InsertMod"
Option Explicit


'Private Function CalcBasicDet(objMatColl As Collection) As Single
'Dim tDet
''a b
''c d
''=ad-cb
'
'tDet = objMatColl.Item(objMatColl.Count)
''remove from Pila
'objMatColl.Remove (objMatColl.Count)
''calculate
'
'CalcBasicDet = ((tDet(1, 1) * tDet(2, 2)) - (tDet(2, 1) * tDet(1, 2)))
'
'End Function
'Private Function CalcDet(objMat As Collection, sTabName$) As Single
'Static iStCont%, iALine%
'Dim j%, iContCopy As Single, x%
'Dim iDim%
'Dim MatArray, MatTemp
''iCont = 1
''Matriz
'
'MatArray = objMat.Item(sTabName)
''Dimension
'iDim = UBound(MatArray)
''Si la dimension es 2x2, calcula la det de 2x2
'    If iDim = 2 Then
'        If bSign Then
'            'POSITIVE
'            CalcDet = (-objPila(objPila.Count)) * CalcBasicDet(objMat)
'            Else
'            'NEGATIVE
'            CalcDet = (objPila(objPila.Count)) * CalcBasicDet(objMat)
'        End If
'         Exit Function
'    Else
'    End If
''Instruccion basica
''For x = 1 To UBound(MatTemp)
'    For j = 1 To iDim
'    'agregar a pila
'    Call objPila.Add(MatArray(1, j))
'    'räkna Inre matris
'    Call SetInarray(objMat, sTabName, 1, j, (iDim - 1))
'    If bSign Then
'        bSign = False
'        'USE FOR POSITIVE VALUES
'    Else
'        bSign = True
'        'USE FOR NEGATIVES VALUES
'    End If
'    Next j
'    'set to menos
'
''Next x   '!!!!!!!!!!!!!!!!!
'
''Cofactores
'If iALine = 0 Then
'    iALine = 1
'End If
'    Call objCofactor.Add(iCont, CStr(iStCont + 1))
'iStCont = iStCont + 1
''Call objCofactor.Add(objPila(iALine), CStr(iStCont + 1))
''iStCont = iStCont + 1
'iContCopy = iCont
'iALine = iALine + 4
''iStCont = iStCont + 1
'iCont = 0
'CalcDet = iContCopy
''iCont = 0
'End Function
'
'Private Sub SetInarray(objMatArray As Collection, sName$, iFila%, iColumna%, iDim%)
'Dim tArray2
'Dim x%, y%, i%, j%, sngMatris() As Single
'Dim objDetA As New Collection
'x = 1
'y = 1
''Matriz
'ReDim sngMatris(iDim, iDim)
'
'tArray2 = objMatArray.Item(sName)
'
'For i = 1 To iDim + 1
'    For j = 1 To iDim + 1
'      If (i <> iFila) Then
'        If (j <> iColumna) Then
'            sngMatris(x, y) = tArray2(i, j)
'            y = y + 1
'        End If
'      End If
'    Next j
'   x = i
'   y = 1
'Next i
''Añadir a colecion
'Call objDetA.Add(sngMatris, sName)
''SetInarray = sngMatris
'    'sumar
'iCont = iCont + CalcDet(objDetA, sName)
''iCont = 0
'End Sub

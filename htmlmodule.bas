Attribute VB_Name = "HtmlModule"
Option Explicit

Public Sub OpenMat(filnamn As String, objDOM As Object)
Dim iLedigt%, iDim%, rowPos%, i%, j%

Dim ii%, jj%
Dim tabMN As Object


rowPos = 0
Call CreateRow(3, rowPos, objDOM)
Call CreateTable(3, rowPos, objDOM)
Call ShowTable(filnamn, rowPos, 0, objDOM)

End Sub

Private Sub CreateRow(iCells%, rPos%, objDom1 As Object)
Dim iLastRow%, i%

objDom1.All.workspace.insertRow
iLastRow = GetLastRow(objDom1)
For i = 1 To iCells
objDom1.All.workspace.rows(iLastRow - 1).insertCell
objDom1.All.workspace.rows(iLastRow - 1).cells(i - 1).vAlign = "top"
Next i
rPos = iLastRow - 1
End Sub

Private Function GetLastRow(objDom2 As Object) As Integer
Dim valor$
Dim i%, t As Object
i = 0
Set t = objDom2.All.workspace.rows(i)
For Each t In objDom2.All.workspace.rows
    valor = t.rowIndex
i = i + 1
Set t = objDom2.All.workspace.rows(i)
Next t
GetLastRow = i
End Function

Private Sub CreateTable(maxCells%, rPos%, objDomm As Object)
Dim defTable$, i%
For i = 1 To maxCells
defTable = "<table valign='top' align='center' border='1' id='tab" & CStr(rPos) & CStr(i - 1) & "'></table>"

objDomm.All.workspace.rows(rPos).cells(i - 1).innerHTML = defTable
Next i


End Sub

Public Sub AddMat(filnamn As String, objDynDom As Object)
If (GetLastRow(objDynDom) = 1) Then
    Call OpenMat(filnamn, objDynDom)
    Else
    Call AddMatris(filnamn, objDynDom)
End If
End Sub

Private Sub AddMatris(sFile$, objDyDom As Object)
Dim CurrentRow%, tabFirstCell As Object, tabFirstCell2 As Object, tabObjRow As Object
Dim i%, y%, t As Object, rowPos%

CurrentRow = GetLastRow(objDyDom) - 1
'current row
Set tabObjRow = objDyDom.All.workspace.rows(CurrentRow).cells
'current row, first cell
Set tabFirstCell = objDyDom.All.workspace.rows(CurrentRow).cells(0)
'current row, first cell, current nested table
Set tabFirstCell2 = tabFirstCell.children(i).rows(0).cells(0)

For Each tabFirstCell2 In tabObjRow
'MsgBox (t.innerText)
If (tabFirstCell2.innerText = "") Then
    y = 1
    Exit For
End If
i = i + 1
Next tabFirstCell2

'-----------------------------------------------------
If y = 1 Then
        Call ShowTable(sFile, CurrentRow, i, objDyDom)
Else
    'MsgBox ("wait")
    rowPos = 0
    Call CreateRow(3, rowPos, objDyDom)
    Call CreateTable(3, rowPos, objDyDom)
    CurrentRow = GetLastRow(objDyDom) - 1
    Call ShowTable(sFile, CurrentRow, 0, objDyDom)
End If
'-----------------------------------------------------
End Sub



Private Sub ShowTable(filnamn$, rowPos%, colPos%, objDOM As Object)
Dim iLedigt%, iDim%, i%, j%, tabMN As Object, ii%, jj%
Dim xyValue As String, matName$
Dim tabName$
tabName = "tab" & CStr(rowPos) & CStr(colPos)
'-----------------------------------
'-----------------------------------
DoEvents
'Llenar nombre de MATRICES EN LISTAS
Call frmMatriz.ListItems(tabName)
'-----------------------------------
'-----------------------------------
Set tabMN = objDOM.All("tab" & CStr(rowPos) & CStr(colPos))
'table properties------------
'----------------------------
tabMN.border = 2
'----------------------------
iLedigt = FreeFile
Open filnamn For Input As iLedigt
'primera linea contiene dimension
Input #iLedigt, iDim
Input #iLedigt, matName
tabMN.insertRow
'FIRST CELL
tabMN.rows(0).insertCell

'properties--table-TITLE---------------------
tabMN.rows(0).cells(0).colSpan = iDim
tabMN.rows(0).cells(0).innerText = matName & " [" & tabName & "] Dim" & CStr(iDim) & "X" & CStr(iDim)
tabMN.rows(0).cells(0).Style.fontFamily = "Arial"
tabMN.rows(0).cells(0).Style.FontSize = "12"
tabMN.rows(0).cells(0).Style.fontWeight = "bold"
tabMN.rows(0).cells(0).Style.backgroundColor = "#d29e51"
'---------------------------------------------

For i = 1 To iDim
ii = i
    'insert TR ROW
    tabMN.insertRow
    For j = 1 To iDim
    jj = j - 1
    'insert TD CELL
    tabMN.rows(ii).insertCell
    'retrieve VALUE
    Input #iLedigt, xyValue
    tabMN.rows(ii).cells(jj).innerText = xyValue
    'properties----------------------------------------
    '--------------------------------------------------
    tabMN.rows(ii).cells(jj).Align = "center"
    tabMN.rows(ii).cells(jj).Style.fontFamily = "Arial"
    tabMN.rows(ii).cells(jj).Style.FontSize = "12"
    tabMN.rows(ii).cells(jj).Style.backgroundColor = "#8080c0"
    '----------------------------------------------------
    Next j
Next i

Close #iLedigt
'guardando MATRIZ EN COLLECION-----
'----------------------------------
'----------------------------------
Call MatOps.SetMemMat(tabName, objDOM)
'----------------------------------
'----------------------------------
End Sub


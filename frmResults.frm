VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmResults 
   Caption         =   "Resultado"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6900
   Icon            =   "frmResults.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser webResults 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      ExtentX         =   11880
      ExtentY         =   5953
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const TABLE_PROPERTIES As String = "<table border='0' align='center' valign='top' cellpadding='5' width='100%' id='result'></table>"
Private WithEvents domResults As HTMLDocument
Attribute domResults.VB_VarHelpID = -1

Private Sub Form_Load()

webResults.Navigate App.Path & "\results.htm"
frmResults.Show
End Sub
Public Sub ShowEuler(sName$, t As Collection)
Dim objTabOp
Dim i%, j%, ii%, jj%
Dim tabMN As Object, iDimX%, iDimY%
'IMPORTANTE PONERLO SI NO NO LLAMA A SUS EVENTOS DE INICIO!!!!!
DoEvents



'asignando MATRIZ
objTabOp = t(sName)
iDimX = UBound(objTabOp, 1)
iDimY = UBound(objTabOp, 2)


domResults.body.innerHTML = TABLE_PROPERTIES
Set tabMN = domResults.All("result")
tabMN.Style.fontFamily = "Arial"
tabMN.Style.backgroundColor = "white" 'new
tabMN.Style.FontSize = "12"
tabMN.insertRow
tabMN.rows(0).insertCell
tabMN.rows(0).cells(0).innerText = sName
tabMN.rows(0).cells(0).colSpan = CStr(iDimY)
tabMN.rows(0).cells(0).Align = "center"
'tabMN.rows().cells(0).FontSize = "bold"

tabMN.insertRow
tabMN.rows(1).insertCell
tabMN.rows(1).cells(0).innerHTML = "X<sub>n</sub>"
tabMN.rows(1).cells(0).Align = "center"

tabMN.rows(1).insertCell
tabMN.rows(1).cells(1).innerHTML = "Y<sub>n</sub>"
tabMN.rows(1).cells(1).Align = "center"

tabMN.rows(1).insertCell
tabMN.rows(1).cells(2).innerHTML = "Y'<sub>n</sub>"
tabMN.rows(1).cells(2).Align = "center"

tabMN.rows(1).insertCell
tabMN.rows(1).cells(3).innerHTML = "hY'<sub>n</sub>"
tabMN.rows(1).cells(3).Align = "center"

If sName = "EulerM" Then
    tabMN.rows(1).insertCell
    tabMN.rows(1).cells(4).innerHTML = "Y<sub>n+1</sub>"
    tabMN.rows(1).cells(4).Align = "center"
    
    tabMN.rows(1).insertCell
    tabMN.rows(1).cells(5).innerHTML = "Y'<sub>n+1</sub>"
    tabMN.rows(1).cells(5).Align = "center"
    
    tabMN.rows(1).insertCell
    tabMN.rows(1).cells(6).innerHTML = "Y'<sub>prom</sub>"
    tabMN.rows(1).cells(6).Align = "center"
    
    tabMN.rows(1).insertCell
    tabMN.rows(1).cells(7).innerHTML = "hY'<sub>prom</sub>"
    tabMN.rows(1).cells(7).Align = "center"

End If


For i = 1 To iDimX
ii = i + 1
    'insert TR ROW
    tabMN.insertRow
    For j = 1 To iDimY
    jj = j - 1
    'insert TD CELL
    tabMN.rows(ii).insertCell
    'set VALUE
    tabMN.rows(ii).cells(jj).innerText = objTabOp(i, j)
    tabMN.rows(ii).cells(jj).Style.backgroundColor = "#8080C0" 'new
    Next j
Next i
    frmResults.ZOrder

End Sub

Public Sub ShowOpResults(sName$, t As Collection)
Dim objTabOp
Dim i%, j%, ii%, jj%
Dim iDim%, tabMN As Object
'IMPORTANTE PONERLO SI NO NO LLAMA A SUS EVENTOS DE INICIO!!!!!
DoEvents



'asignando MATRIZ
objTabOp = t(sName)
iDim = UBound(objTabOp)

domResults.body.innerHTML = TABLE_PROPERTIES
Set tabMN = domResults.All("result")
tabMN.Style.fontFamily = "Arial"
tabMN.Style.backgroundColor = "white" 'new
tabMN.Style.FontSize = "12"
tabMN.insertRow
tabMN.rows(0).insertCell
tabMN.rows(0).cells(0).innerText = sName
tabMN.rows(0).cells(0).colSpan = CStr(iDim)
tabMN.rows(0).cells(0).Align = "center"
'tabMN.rows(0).cells(0).FontSize = "bold"
For i = 1 To iDim
ii = i
    'insert TR ROW
    tabMN.insertRow
    For j = 1 To iDim
    jj = j - 1
    'insert TD CELL
    tabMN.rows(ii).insertCell
    'set VALUE
    tabMN.rows(ii).cells(jj).innerText = objTabOp(i, j)
    tabMN.rows(ii).cells(jj).Style.backgroundColor = "#8080C0" 'new
    Next j
Next i
    frmResults.ZOrder

End Sub

Public Sub ResultsXYZ(sName$, t As Collection, tErr As Collection)
Dim objTabOp
Dim i%, j%, ii%, jj%
Dim iDim%, tabMN As Object, objError
'IMPORTANTE PONERLO SI NO NO LLAMA A SUS EVENTOS DE INICIO!!!!!
DoEvents



'asignando MATRIZ
objTabOp = t(sName)
objError = tErr(sName)
iDim = UBound(objTabOp)

domResults.body.innerHTML = TABLE_PROPERTIES
Set tabMN = domResults.All("result")
tabMN.Style.fontFamily = "Arial"
tabMN.Style.backgroundColor = "white" 'new
tabMN.Style.FontSize = "12"
tabMN.insertRow
tabMN.rows(0).insertCell
tabMN.rows(0).cells(0).innerText = sName
tabMN.rows(0).cells(0).colSpan = CStr(iDim + 1)
tabMN.rows(0).cells(0).Align = "center"
'tabMN.rows(0).cells(0).FontSize = "bold"
tabMN.insertRow
tabMN.rows(1).insertCell
tabMN.rows(1).insertCell
tabMN.rows(1).insertCell

tabMN.rows(1).cells(0).innerText = CStr(Time)
tabMN.rows(1).cells(0).Style.backgroundColor = "#00BFBF"
tabMN.rows(1).cells(1).innerText = "Valor"
tabMN.rows(1).cells(1).Style.backgroundColor = "#00BFBF"
tabMN.rows(1).cells(2).innerText = "Error"
tabMN.rows(1).cells(2).Style.backgroundColor = "#00BFBF"


For i = 2 To iDim + 1
ii = i
    'insert TR ROW
    tabMN.insertRow
    'insert TD CELL
    tabMN.rows(ii).insertCell
    tabMN.rows(ii).cells(0).innerText = "Valor" & CStr(i - 1)
    tabMN.rows(ii).cells(0).Style.backgroundColor = "#00BFBF"
    tabMN.rows(ii).insertCell
    tabMN.rows(ii).cells(1).innerText = objTabOp(i - 1)
    tabMN.rows(ii).cells(1).Style.backgroundColor = "#8080C0"
    tabMN.rows(ii).insertCell
    tabMN.rows(ii).cells(2).innerText = objError(i - 1)
    tabMN.rows(ii).cells(2).Style.backgroundColor = "#8080C0"
Next i
    frmResults.ZOrder

End Sub
Public Sub ShowOpResultsEsp(sName$, t As Collection)
Dim objTabOp
Dim i%, j%, ii%, jj%
Dim iDim%, tabMN As Object
'IMPORTANTE PONERLO SI NO NO LLAMA A SUS EVENTOS DE INICIO!!!!!
DoEvents



'asignando MATRIZ
objTabOp = t(sName)
iDim = UBound(objTabOp)

domResults.body.innerHTML = TABLE_PROPERTIES
Set tabMN = domResults.All("result")
tabMN.Style.fontFamily = "Arial"
tabMN.Style.backgroundColor = "white" 'new
tabMN.Style.FontSize = "12"
tabMN.insertRow
tabMN.rows(0).insertCell
tabMN.rows(0).cells(0).innerText = sName
tabMN.rows(0).cells(0).colSpan = CStr(iDim + 1)
tabMN.rows(0).cells(0).Align = "center"
'tabMN.rows(0).cells(0).FontSize = "bold"
For i = 1 To iDim
ii = i
    'insert TR ROW
    tabMN.insertRow
    For j = 1 To (iDim + 1)
    jj = j - 1
    'insert TD CELL
    tabMN.rows(ii).insertCell
    'set VALUE
    tabMN.rows(ii).cells(jj).innerText = objTabOp(i, j)
    tabMN.rows(ii).cells(jj).Style.backgroundColor = "#8080C0" ' new
    Next j
Next i
    frmResults.ZOrder

End Sub

Private Sub Form_Resize()
webResults.Width = frmResults.Width
webResults.Height = frmResults.Height
End Sub

Private Sub webResults_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If webResults.Document Is Nothing Then Exit Sub
Set domResults = webResults.Document


End Sub

Private Sub webResults_DownloadComplete()
If webResults.Document Is Nothing Then Exit Sub
Set domResults = webResults.Document
End Sub
Public Sub ResultsInvMat(sName$, t As Collection)
Dim objTabOp
Dim i%, j%, ii%, jj%
Dim iDim%, tabMN As Object
'IMPORTANTE PONERLO SI NO NO LLAMA A SUS EVENTOS DE INICIO!!!!!
DoEvents



'asignando MATRIZ
objTabOp = t(sName)
iDim = UBound(objTabOp)

domResults.body.innerHTML = TABLE_PROPERTIES
Set tabMN = domResults.All("result")
tabMN.Style.fontFamily = "Arial"
tabMN.Style.backgroundColor = "white" 'new
tabMN.Style.FontSize = "12"
tabMN.insertRow
tabMN.rows(0).insertCell
tabMN.rows(0).cells(0).innerText = sName
tabMN.rows(0).cells(0).colSpan = CStr(iDim * 2)
tabMN.rows(0).cells(0).Align = "center"
'tabMN.rows(0).cells(0).FontSize = "bold"
For i = 1 To iDim
ii = i
    'insert TR ROW
    tabMN.insertRow
    For j = 1 To (iDim * 2)
    jj = j - 1
    'insert TD CELL
    tabMN.rows(ii).insertCell
    'set VALUE
    tabMN.rows(ii).cells(jj).innerText = objTabOp(i, j)
    tabMN.rows(ii).cells(jj).Style.backgroundColor = "#8080C0" 'new
    Next j
Next i
    frmResults.ZOrder

End Sub


Public Sub ShowRootResults(objT As Collection)
Dim objTable As Variant
Dim i%, j%, ii%, jj%
Dim iDimY%, iDimX%, tabMN As Object
'IMPORTANTE PONERLO SI NO NO LLAMA A SUS EVENTOS DE INICIO!!!!!
DoEvents

'asignando MATRIZ
objTable = objT(1)
'Dimension
iDimY = UBound(objTable, 1)
iDimX = UBound(objTable, 2)

domResults.body.innerHTML = TABLE_PROPERTIES
Set tabMN = domResults.All("result")
tabMN.Style.fontFamily = "Arial"
tabMN.Style.FontSize = "12"
tabMN.Style.backgroundColor = "white"
tabMN.insertRow

If iDimX = 7 Then
'INTERVALO MEDIO y RF
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(0).innerText = "Iteración"
    tabMN.rows(0).cells(0).Align = "center"
    
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(1).innerText = "A"
    tabMN.rows(0).cells(1).Align = "center"
    
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(2).innerText = "B"
    tabMN.rows(0).cells(2).Align = "center"
    
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(3).innerText = "C"
    tabMN.rows(0).cells(3).Align = "center"
    
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(4).innerText = "F(A)"
    tabMN.rows(0).cells(4).Align = "center"
    
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(5).innerText = "F(B)"
    tabMN.rows(0).cells(5).Align = "center"
    
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(6).innerText = "F(C)"
    tabMN.rows(0).cells(6).Align = "center"
ElseIf iDimX = 3 Then
'SECANTE
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(0).innerText = "Iteración"
    tabMN.rows(0).cells(0).Align = "center"
    
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(1).innerHTML = "X<sub>n</sub>"
    tabMN.rows(0).cells(1).Align = "center"
    
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(2).innerHTML = "Y<sub>n</sub>"
    tabMN.rows(0).cells(2).Align = "center"
ElseIf iDimX = 4 Then
'NEWTON RAPHSON
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(0).innerText = "Iteración"
    tabMN.rows(0).cells(0).Align = "center"
    
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(1).innerHTML = "X<sub>n</sub>"
    tabMN.rows(0).cells(1).Align = "center"
    
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(2).innerHTML = "f(X<sub>n</sub>)"
    tabMN.rows(0).cells(2).Align = "center"
    
    tabMN.rows(0).insertCell
    tabMN.rows(0).cells(3).innerHTML = "f'(X<sub>n</sub>)"
    tabMN.rows(0).cells(3).Align = "center"
End If

For i = 1 To iDimY
DoEvents
ii = i
    'insert TR ROW
    tabMN.insertRow
    For j = 1 To iDimX
    jj = j - 1
    'insert TD CELL
    tabMN.rows(ii).insertCell
    'set VALUE
    tabMN.rows(ii).cells(jj).innerText = Format(objTable(i, j), "0.0000000000")
    
    'properties
    tabMN.rows(ii).cells(jj).Align = "right"
    tabMN.rows(ii).cells(jj).Style.backgroundColor = "#8080C0"
    Next j
Next i
    frmResults.ZOrder

End Sub

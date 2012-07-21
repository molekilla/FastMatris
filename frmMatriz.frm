VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMatriz 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fast Matris"
   ClientHeight    =   7935
   ClientLeft      =   750
   ClientTop       =   1635
   ClientWidth     =   8715
   Icon            =   "frmMatriz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   529
   ScaleMode       =   0  'User
   ScaleWidth      =   581
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser webBar 
      Height          =   7215
      Left            =   7200
      TabIndex        =   14
      Top             =   0
      Width           =   1575
      ExtentX         =   2778
      ExtentY         =   12726
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
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   ">"
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox txtMemInMem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   5520
      TabIndex        =   11
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox txtMemVar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   7
      Top             =   7560
      Width           =   1335
   End
   Begin VB.ComboBox List2 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   7560
      Width           =   1335
   End
   Begin VB.ComboBox List1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   7440
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cdlMainMat 
      Left            =   2520
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Abrir archivo FST"
      Filter          =   "FST (*.fst)|*.fst"
      InitDir         =   "c:\"
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      ExtentX         =   12726
      ExtentY         =   12726
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Matrices en Mem:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Matrices:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Matrices:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mem Variable:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6960
      TabIndex        =   8
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label lblOp1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "Click me!"
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Label2"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   7560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   -240
      Top             =   7320
      Width           =   9015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuOpenMat 
         Caption         =   "&Abrir Matriz"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnunAdd 
         Caption         =   "Aña&dir matris"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSaveMat 
         Caption         =   "&Guardar Matriz"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Terminar"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuInsertMat 
         Caption         =   "&Insertar Matrices"
      End
   End
   Begin VB.Menu mnuFunction 
      Caption         =   "&Funciones"
      Begin VB.Menu mnuGauss 
         Caption         =   "&Gauss"
      End
      Begin VB.Menu mnuGaussJordan 
         Caption         =   "Gauss-&Jordan"
      End
      Begin VB.Menu mnuSeidel 
         Caption         =   "Gauss-&Seidel"
      End
      Begin VB.Menu mnuJacobi 
         Caption         =   "Gauss-Jaco&bi"
      End
      Begin VB.Menu mnuInversa 
         Caption         =   "Gauss usando una in&versa"
      End
      Begin VB.Menu mnuDet 
         Caption         =   "&Determinante"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuMe 
         Caption         =   "&Sobre Fast Matris y su autor"
      End
      Begin VB.Menu mnuMainHelp 
         Caption         =   "Ayuda de &Fast Matris"
      End
   End
End
Attribute VB_Name = "frmMatriz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents dom2 As HTMLDocument
Attribute dom2.VB_VarHelpID = -1
Private WithEvents dom1 As HTMLDocument
Attribute dom1.VB_VarHelpID = -1

Private Sub cmdCalc_Click()
Dim sTab$, sTab2$, sTab3$
Dim sItem$, i%, sSigno$, tempMat As String


'Set Mem Matrices
sTab = List1.Text
sTab2 = List2.Text

tempMat = txtMemVar

   sSigno = lblOp1.Caption
   Call MatOps.CalcMat(sSigno, sTab, sTab2, dom1, tempMat)

If (txtMemInMem.Text = "") And (txtMemVar.Text <> "") Then
    txtMemInMem.Text = 0
End If
If (txtMemVar.Text <> "") Then
txtMemInMem.Text = CStr(txtMemInMem.Text + 1) & "+"
End If

End Sub



Private Sub Command1_Click()
'mem matrices
Call frmControls.ShowMemMat

End Sub

Private Sub Command2_Click()
WebBrowser1.Navigate App.Path & "\html\" & "graphpage.htm"
End Sub



Private Function dom1_onclick() As Boolean
Dim sImgID$

sImgID = dom1.parentWindow.event.srcElement.iD
Select Case sImgID
    '-----Para Graphpage---------
        Case "Clear": Call RootEquations.Clear(dom1)
End Select

End Function

Private Function dom1_ondblclick() As Boolean
Dim sString As String
Dim Leng As Integer, i%, y%, x%, iRows%, iCells%
Dim sName As String, iDim%
On Error GoTo NoNO
sString = dom1.parentWindow.event.srcElement.innerText
Leng = Len(sString)
'-------------------------------------------------
For i = 1 To Leng
    If Mid(sString, i, 1) = "[" Then
        y = i + 1
    ElseIf Mid(sString, i, 1) = "]" Then
        x = i - y
    End If
    If Mid(sString, i, 1) = "X" Then
        iDim = CInt(Mid(sString, i + 1, Leng))
    End If
Next i
'-------------------------------------------------
sName = Mid(sString, y, x)
iRows = dom1.All(CVar(sName)).rows.length
iCells = dom1.All(CVar(sName)).rows(1).cells.length
If (iRows <> iCells) Then
 Call AddValuesBoxes(sName, iDim)
    Else
    'contiene terminos simples
 Call RemoveValuesBoxes(sName, iDim)
    End If

NoNO:
If (Err.Number <> 0) Then
Select Case dom1.parentWindow.event.srcElement.iD
    Case "MoveUp": Exit Function
    Case "MoveDown": Exit Function
    Case "ScaleUp": Exit Function
    Case "ScaleDown": Exit Function
    Case "MoveLeft": Exit Function
    Case "MoveRight": Exit Function
    Case "Rotate180": Exit Function
    Case "Clear": Exit Function
    Case Else: MsgBox ("No hay no clikees hay!!!")
End Select

End If
End Function



Private Sub dom1_onmouseup()
Dim sImgID$

sImgID = dom1.parentWindow.event.srcElement.iD
Select Case sImgID
    Case "combo": Call GetFunc(dom1, dom2)
    Case "ra": WebBrowser1.Navigate App.Path & "\help02.htm"
    Case "ma": WebBrowser1.Navigate App.Path & "\help.htm"
End Select

End Sub

Private Function dom2_onclick() As Boolean
Dim sImgID$

sImgID = dom2.parentWindow.event.srcElement.iD
Select Case sImgID
    '-----Para Root Web bar---------
    Case "ShowGraph":
        Call GraphXY(dom1, dom2)
        Call InsertFunc(dom1, dom2)
    Case "SearchLeft":
        Call MoveLineLeft(dom1)
    Case "SearchRight":
        Call MoveLineRight(dom1)
    Case "Noll":
        Call RootEquations.Noll(dom1)
    Case "Clear":
        Call RootEquations.Clear(dom1)
    'Last Update 5 Dec. 1998
    Case "Bonus":
        Load frmBairstow
        frmBairstow.Show
        frmBairstow.ZOrder
    Case "Dif":
        Load frmEcDiferencial
        frmEcDiferencial.Show
        frmEcDiferencial.ZOrder
    Case "Integral":
        Load frmIntegral
        frmIntegral.Show
        frmIntegral.ZOrder
    Case "Ajuste":
        Load frmFlex
        frmFlex.Show
        frmFlex.ZOrder
    Case "RootWeb":
        webBar.Navigate2 App.Path & "\html\rootweb.htm"
        WebBrowser1.Navigate App.Path & "\html\" & "graphpage.htm"
        webBar_DownloadComplete
    Case "MatrisWeb":
        webBar.Navigate2 App.Path & "\html\webbar.htm"
        WebBrowser1.Navigate App.Path & "\main.htm"
        webBar_DownloadComplete
    End Select

End Function

Private Function dom2_ondblclick() As Boolean
Dim sImgID$

sImgID = dom2.parentWindow.event.srcElement.iD
Select Case sImgID
    '-----Para Matriz Web bar-------
    Case "Seidel": Call Seidel
    Case "Jacobi": Call Jacobi
    Case "Inversa": Call Inversa
    Case "GJ": Call GaussJordan
    Case "Gauss": Call Gauss
    Case "Det": Call Determinante
    '-----Para Root Web bar---------
    Case "IM": Call Bisection(dom2)
    Case "RF": Call RegulaFalsi(dom2)
    Case "Secante": Call Secante(dom2)
    Case "Newton": Call Newton(dom2)
    Case "RFMOD": Call RFMod(dom2)
End Select
End Function

Private Sub Form_Load()

WebBrowser1.Navigate App.Path & "\" & "main.htm"
webBar.Navigate App.Path & "\html\" & "webbar.htm"
Load frmControls
frmControls.Show
frmControls.Top = 0
frmControls.Left = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = 0
End
End Sub
Private Sub LblOp1_Click()
Call ChangeOperator(lblOp1)
End Sub
Private Sub mnuExit_Click()
Unload Me
End
End Sub

Private Sub mnuInsertMat_Click()
Load frmInsertMat
frmInsertMat.Show
End Sub

Private Sub mnuMainHelp_Click()
WebBrowser1.Navigate App.Path & "\" & "help.htm"
End Sub

Private Sub mnuMe_Click()
Load frmRush
frmRush.Show
End Sub

Private Sub mnunAdd_Click()
On Error Resume Next
Dim filename$
cdlMainMat.ShowOpen
filename = cdlMainMat.filename
Call OpenFile(filename)
End Sub

Private Sub mnuOpenMat_Click()
On Error GoTo Cancelar
Dim filename$

cdlMainMat.ShowOpen
filename = cdlMainMat.filename
Call HtmlModule.OpenMat(filename, dom1)

mnuOpenMat.Enabled = False
Cancelar:
Exit Sub
End Sub

Private Sub webBar_DownloadComplete()
If webBar.Document Is Nothing Then Exit Sub
Set dom2 = webBar.Document
'dom2.body.Style.backgroundColor = vbBlack
End Sub

Private Sub WebBrowser1_DownloadComplete()
If WebBrowser1.Document Is Nothing Then Exit Sub
Set dom1 = WebBrowser1.Document

End Sub



Public Sub ListItems(items$)
List1.AddItem items
List2.AddItem items


End Sub

Private Sub ChangeOperator(lblControl As Control)
Dim lblop As Label
Set lblop = lblControl

If (lblop.Caption = "+") Then
lblop.Caption = "-"
ElseIf (lblop.Caption = "-") Then
    lblop.Caption = "*"
    ElseIf (lblop.Caption = "*") Then
    lblop.Caption = "N/A"
        ElseIf (lblop.Caption = "N/A") Then
        lblop.Caption = "+"
         End If
    
End Sub

Public Sub OpenFile(sFileName As String)
On Error GoTo Cancelar


Call HtmlModule.AddMat(sFileName, dom1)
If (mnuOpenMat.Enabled = True) Then
    mnuOpenMat.Enabled = False
End If
Cancelar:
Exit Sub
End Sub


Private Sub RemoveValuesBoxes(sName$, iDim As Integer)
Dim i%
Const htmlINPUTBOX$ = "<input type='text' size='3' "
Const A% = 64
DoEvents
dom1.All(CVar(sName)).rows(0).cells(0).colSpan = (iDim)
For i = iDim To 1 Step -1
    dom1.All(CVar(sName)).rows(i).deleteCell (iDim)
Next i
End Sub
Private Sub AddValuesBoxes(sName$, iDim As Integer)
Dim i%
Const htmlINPUTBOX$ = "<input type='text' size='3' "
Const A% = 64
DoEvents
dom1.All(CVar(sName)).rows(0).cells(0).colSpan = (iDim + 1)
For i = 1 To iDim
    dom1.All(CVar(sName)).rows(i).insertCell
    dom1.All(CVar(sName)).rows(i).cells(iDim).innerHTML = htmlINPUTBOX & "id='" & Chr(A + i) & "'>"
    dom1.All(CVar(sName)).rows(i).cells(iDim).Align = "right"
Next i
End Sub

Private Sub AddValuesToMatr(sNamn$, t As Collection)
Dim Mat() As Single
Dim iDim%, i%, j%, iRows%, iCols%
iRows = dom1.All(CVar(sNamn)).rows.length - 1
iCols = iRows + 1
ReDim Mat(iRows, iCols)

Set t = New Collection
For i = 1 To iRows
    For j = 1 To (iCols - 1)
        
        Mat(i, j) = CSng(dom1.All(CVar(sNamn)).rows(i).cells(j - 1).innerText)
    Next j
    Mat(i, iCols) = CSng(dom1.All(CVar(Chr(64 + i))).Value)
Next i
Call t.Add(Mat, sNamn)
End Sub

Private Sub Seidel()
Dim sTabtitle$, iRows%, iCells%
Dim objT As Collection



sTabtitle = dom2.All.tabNum.Value
If sTabtitle <> "" And dom2.All.iTer.Value <> "" Then
'checkear por si tiene terminos libres
iRows = dom1.All(CVar(sTabtitle)).rows.length
iCells = dom1.All(CVar(sTabtitle)).rows(1).cells.length
    If (iRows <> iCells) Then
'        Call CalcGauss(sTabtitle)
    Else
    'contiene terminos simples
        Call AddValuesToMatr(sTabtitle, objT)
        Call CalcSeidel(sTabtitle, objT, CInt(dom2.All.iTer.Value))
    End If
End If

End Sub

Private Sub Jacobi()
Dim sTabtitle$, iRows%, iCells%
Dim objT As Collection


sTabtitle = dom2.All.tabNum.Value
If sTabtitle <> "" And dom2.All.iTer.Value <> "" Then
'checkear por si tiene terminos libres
iRows = dom1.All(CVar(sTabtitle)).rows.length
iCells = dom1.All(CVar(sTabtitle)).rows(1).cells.length
    If (iRows <> iCells) Then
'        Call CalcGauss(sTabtitle)
    Else
    'contiene terminos simples
        Call AddValuesToMatr(sTabtitle, objT)
        Call CalcJacobi(sTabtitle, objT, CInt(dom2.All.iTer.Value))
    End If
End If

End Sub

Private Sub Inversa()
Dim sTabtitle$, iRows%, iCells%



sTabtitle = dom2.All.tabNum.Value
If sTabtitle <> "" Then
'checkear por si tiene terminos libres
iRows = dom1.All(CVar(sTabtitle)).rows.length
iCells = dom1.All(CVar(sTabtitle)).rows(1).cells.length
    If (iRows <> iCells) Then
        Call InvMat(sTabtitle)
    Else
    'contiene terminos simples
        'Call AddValuesToMatr(sTabtitle, objT)
        'Call CalcGauss(sTabtitle, 2, objT)
    End If
End If
End Sub

Private Sub Gauss()
Dim sTabtitle$, iRows%, iCells%
Dim result As Single
Dim objT As Collection


sTabtitle = dom2.All.tabNum.Value
If sTabtitle <> "" Then
'checkear por si tiene terminos libres
iRows = dom1.All(CVar(sTabtitle)).rows.length
iCells = dom1.All(CVar(sTabtitle)).rows(1).cells.length
    If (iRows <> iCells) Then
        Call CalcGauss(sTabtitle)
    Else
    'contiene terminos simples
        Call AddValuesToMatr(sTabtitle, objT)
        Call CalcGauss(sTabtitle, 2, objT)
    End If
End If

End Sub

Private Sub GaussJordan()
Dim sTabtitle$
Dim result As Single, iRows%, iCells%
Dim objT As Collection

sTabtitle = dom2.All.tabNum.Value
If sTabtitle <> "" Then
'checkear por si tiene terminos libres
iRows = dom1.All(CVar(sTabtitle)).rows.length
iCells = dom1.All(CVar(sTabtitle)).rows(1).cells.length
    If (iRows <> iCells) Then
        Call CalcJordan(sTabtitle)
    Else
    'contiene terminos simples
        Call AddValuesToMatr(sTabtitle, objT)
        Call CalcJordan(sTabtitle, 2, objT)
    End If
End If

End Sub

Private Sub Determinante()
Dim sTabtitle$
sTabtitle = dom2.All.tabNum.Value
If sTabtitle <> "" Then

dom1.All.eventbar.innerHTML = "<span id='RD'>Resultado de la determinante " & Space(2) & sTabtitle & " :  </span><span id='ResDet'>" & CStr(MatOps.CalcDeterminante(sTabtitle)) & "</span>"
dom1.All.eventbar.Style.backgroundColor = "white"
dom1.All.eventbar.colSpan = 3
dom1.All.RD.Style.Color = "black"
dom1.All.RD.Style.fontFamily = "Arial"
dom1.All.ResDet.Style.Color = "red"
dom1.All.ResDet.Style.fontFamily = "Arial"


End If

End Sub



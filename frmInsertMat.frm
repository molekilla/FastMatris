VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmInsertMat 
   Caption         =   "Insertar Matrices"
   ClientHeight    =   3780
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   4770
   Icon            =   "frmInsertMat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlInsert 
      Left            =   3480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      CancelError     =   -1  'True
      DialogTitle     =   "Guardar como..."
      Filter          =   "FST(*.fst)|*.fst"
      InitDir         =   "c:\"
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   2040
      Max             =   0
      Min             =   100
      TabIndex        =   4
      Top             =   480
      Value           =   100
      Width           =   255
   End
   Begin SHDocVwCtl.WebBrowser web02 
      Height          =   2415
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   4695
      ExtentX         =   8281
      ExtentY         =   4260
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
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insertar"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtDim 
      Height          =   285
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Usar (,) o (.) dependiendo su sistema. Ingles=(.), Europeo=(,)"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3480
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "Matrices en disco:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WAIT..."
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   240
      Width           =   975
      Visible         =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Dimension:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Dimension"
      Top             =   480
      Width           =   855
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuSaveMatr 
         Caption         =   "&Guardar Matriz"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmInsertMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const htmlINPUT = "<input type='text' size='5' " 'plus ID
Dim WithEvents dom As HTMLDocument
Attribute dom.VB_VarHelpID = -1
Public Sub SaveMat(sFilnamn$, realname$)
Dim iFile%
Dim rcID$, indexValue As String
Dim i%, j%, iRows%
iFile = FreeFile
If dom Is Nothing Then Exit Sub
'sabemos el nombre de la tabla TAB01
iRows = MaxRows
'--------------------------------------------------------
'sabemos que es cuadratica, no hay necesidad de ver CELLS
'--------------------------------------------------------
Open sFilnamn For Output As iFile
Write #iFile, iRows
Write #iFile, realname
For i = 1 To iRows
'dimension de MATRIZ
    For j = 1 To iRows
        rcID = "m" & CStr(i) & "n" & CStr(j)
        indexValue = dom.All(rcID).Value
        Write #iFile, indexValue
    Next j
Next i

Close #iFile
End Sub
Private Sub cmdInsert_Click()
Dim sRows$, sTD$, i%, j%
Dim iDim%, ii%, jj%
Dim sID$
sTD = ""
sRows = ""
iDim = CInt(txtDim.Text)

'-----------------------------------------
dom.body.innerHTML = "<table width='100%' border='1' id='tab01'></table>"
For i = 1 To iDim
Label2.Visible = True
DoEvents

    'insert TABLE ROW <TR>
    dom.All.tab01.insertRow
    For j = 1 To iDim
    
        ii = i - 1
        jj = j - 1
        'insert TABLE DATA CELL <TD>
        dom.All.tab01.rows(ii).insertCell
        'set COLOR
        dom.All.tab01.rows(ii).cells(jj).bgColor = "Blue"
        'set TEST VALUE
        sID = ("m" & CStr(i) & "n" & CStr(j))
        dom.All.tab01.rows(ii).cells(jj).innerHTML = htmlINPUT & "id='" & sID & "'>"
        'dom.All.tab01.rows(ii).cells(jj).id = (CStr(i) & "c" & CStr(j))
            Label2.BackColor = vbRed
    Next j
   
Next i
Label2.Visible = False
'------------------------------------------

End Sub

Private Sub cmdInsert_KeyPress(KeyAscii As Integer)
Form_KeyPress (KeyAscii)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'eventos de teclado para LA FORMA
'--------------------------------
If (vbKeyD = KeyAscii) Then
    txtDim.SetFocus
End If

End Sub

Private Sub Form_Load()
web02.Navigate "c:\text.htm"

End Sub

Private Sub mnuExit_Click()
Unload frmInsertMat
End Sub

Private Sub mnuSaveMatr_Click()
On Error GoTo Cancelar
Dim filename$, rname$
cdlInsert.InitDir = App.Path & "\tables"
cdlInsert.ShowSave
filename = cdlInsert.filename
rname = cdlInsert.FileTitle
Call SaveMat(filename, rname)
txtQty.Text = CInt(txtQty.Text) + 1
Cancelar:
Exit Sub

End Sub

Private Sub txtDim_KeyPress(KeyAscii As Integer)
If ((KeyAscii < 48) Or (KeyAscii > 57)) Then
 txtDim.Text = ""
End If

End Sub

Private Sub VScroll1_Change()
txtDim.Text = VScroll1.Value
End Sub

Private Sub web02_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If web02.Document Is Nothing Then Exit Sub
Set dom = web02.Document

End Sub


Private Function MaxRows() As Integer
Dim val$
Dim i%, t As Object
i = 0
Set t = dom.All.tab01.rows(i)
For Each t In dom.All.tab01.rows
    val = t.rowIndex
i = i + 1
Set t = dom.All.tab01.rows(i)
Next t
MaxRows = i
End Function

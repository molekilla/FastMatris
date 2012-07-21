VERSION 5.00
Begin VB.Form frmControls 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mat/Mem List"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3930
   Icon            =   "frmControls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Matrices en memoria:"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Matrices Disponibles:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
Dim iResp%
iResp = MsgBox("Eliminar archivo actual?", vbOKCancel, "Eliminacón de archivo")
If iResp <> vbCancel Then
Kill File1.Path & "\" & File1.filename
File1.Refresh
End If
End Sub

Private Sub File1_DblClick()
frmMatriz.OpenFile (File1.Path & "\" & File1.filename)
End Sub

Private Sub Form_Load()


'file matrices
File1.Path = App.Path & "\tables"

End Sub

Public Sub ShowMemMat()
'mem matrices
Dim i%
   List2.Clear
For i = 1 To MatOps.CountMemMat
    List2.AddItem MatOps.GetMemMat(i)
Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = -1
End Sub

Private Sub List2_Click()
'Dim objDOM As Object
''get object
'Set objDOM = frmMatriz.GetDynDOM
SaveMM (List2.List(List2.ListIndex))
End Sub
Private Sub SaveMM(sName$)
Dim sFilnamn$
Dim t As Collection
Dim objTable
Dim iRows%, iFile%, i%, j%
iFile = FreeFile
sFilnamn = App.Path & "\tables\" & sName & ".fst"

Set t = MatOps.GetMemMat2(sName)
objTable = t(sName)
iRows = UBound(objTable)

Open sFilnamn For Output As iFile
Write #iFile, iRows
Write #iFile, sName
For i = 1 To iRows
'dimension de MATRIZ
    For j = 1 To iRows
        'rcID = "m" & CStr(i) & "n" & CStr(j)
        'indexValue = dom.All(rcID).Value
        Write #iFile, objTable(i, j)
    Next j
Next i
Close #iFile

frmMatriz.OpenFile (sFilnamn)

End Sub


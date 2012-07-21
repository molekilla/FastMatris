VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFlex 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste de curvas"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7650
   Icon            =   "frmFlex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLag 
      Caption         =   "Interpolacion de Lagrange"
      Height          =   495
      Left            =   5640
      TabIndex        =   12
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit / Terminar"
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtInter 
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Text            =   "0,73"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdIntNewton 
      Caption         =   "Interpolacion de Newton"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid flexInter 
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4048
      _Version        =   65541
   End
   Begin VB.TextBox txtPoly 
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Text            =   "1"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Regresion Superior"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid flxTabla 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3836
      _Version        =   65541
      Rows            =   3
      Cols            =   4
      AllowUserResizing=   1
   End
   Begin VB.Label Label8 
      Caption         =   "Tabla de valores iniciales"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label Label7 
      Caption         =   "Tabla de resultados"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2880
      Width           =   4815
   End
   Begin VB.Label Label6 
      Caption         =   "Nota: Cuando el numero a buscar esta por debajo de la mitad de la tabla, la tabla se invierte."
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   7455
   End
   Begin VB.Label Label5 
      Caption         =   "Resultado"
      Height          =   255
      Left            =   5640
      TabIndex        =   10
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Valor para interpolacion:"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Ajuste de curvas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de polinomios:"
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
End
Attribute VB_Name = "frmFlex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub IntNewton()
On Error Resume Next
Dim k%, NI%, i%, j%
Dim f() As Single, objArray As Variant
Dim sngInter As Single, sngRest As Single
Dim S As Single, iPos%, iNum%, sngResult As Single
Dim sTemp As Single, iRest As Integer

With flxTabla
NI = .rows - 2
ReDim f(NI, NI)
For x = 1 To NI
    f(x, 0) = CSng(.TextMatrix(x, 2))
Next x
'RESTAS SUCESIVAS, ->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
For k = 1 To NI
    j = NI - k
    For i = 1 To j
        f(i, k) = (f(i + 1, k - 1) - f(i, k - 1)) '/ (.TextMatrix(i, 1) - .TextMatrix(i + k, 1))
    Next i
Next k
End With
With flexInter
.cols = NI
.rows = NI
For k = 1 To NI
        j = NI - k
    For i = 1 To j
        .TextMatrix(i, k) = Format(f(i, k), "###0.0000000")
    Next i
Next k
End With
'----------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'para listas ascendentes
With flxTabla
sngInter = CSng(txtInter)
For i = 1 To NI
    If (sngInter > CSng(.TextMatrix(i, 1))) Then
            sngRest = CSng(.TextMatrix(i, 1))
            iPos = i
    End If
Next i
'para listas descendentes
If (.TextMatrix(1, 1) - .TextMatrix(2, 1)) > 0 Then
'MsgBox ("desc")
sngInter = CSng(txtInter)
For i = 1 To NI
    If (sngInter < CSng(.TextMatrix(i, 1))) Then
            sngRest = CSng(.TextMatrix(i, 1))
            iPos = i
    End If
Next i
End If

S = (sngInter - sngRest) / (.TextMatrix(2, 1) - .TextMatrix(1, 1))
'CALCULAR P '--------------------------------------->>>
k = 0
iNum = 1
For j = 1 To (NI)
If j = 1 Then
    sngResult = CSng(.TextMatrix(iPos, 2))
Else
    sTemp = CSng(flexInter.TextMatrix(iPos, k))

    For i = 1 To (j - 1)
        sTemp = sTemp * (S - (i - 1))
        iNum = iNum * i
    Next i
    sTemp = sTemp / iNum
    sngResult = sngResult + sTemp
End If
    
    iNum = 1
    k = k + 1
    'NI = NI - 1
Next j
'CALCULAR P END*----------------------------------->>>>>>
End With
Label4.Caption = CStr(sngResult)

End Sub
Private Sub cmdCalc_Click()
Call CalcRegSup
End Sub


Private Sub cmdIntNewton_Click()
IntNewton
End Sub

Private Sub cmdLag_Click()
IntLagrange
End Sub

Private Sub Command1_Click()
frmFlex.Hide
Unload frmFlex
End Sub

Private Sub flxTabla_Click()

With flxTabla
If ((.Col = 1) And (.Row = .rows - 1)) Then
    AddNewItem

End If

End With

End Sub

Private Sub flxTabla_EnterCell()
    With flxTabla
        If (.Col <> 0) Then
            .CellBackColor = &H9FAE33
        End If
    End With
End Sub

Private Sub flxTabla_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyDelete) Then
    flxTabla.Text = ""
End If
End Sub

Private Sub flxTabla_KeyPress(KeyAscii As Integer)
Dim sInput$

sInput = flxTabla.Text

If (KeyAscii = 8) Then
'borrar un caracter (backdelete)
    If Len(sInput) = 0 Then
        Exit Sub
    End If
    flxTabla.Text = Mid(sInput, 1, Len(sInput) - 1)
Else
    flxTabla.Text = flxTabla.Text & Chr(KeyAscii)
End If
End Sub
Private Sub flxTabla_LeaveCell()
With flxTabla
    If (.Col <> 0) Then
        .CellBackColor = vbWhite
    End If
End With
End Sub

Private Sub Form_Load()

With flxTabla

.TextMatrix(0, 0) = "i"
.TextMatrix(0, 1) = "x"
.TextMatrix(0, 2) = "y"
.TextMatrix(0, 3) = "G (x)"

.TextMatrix(1, 0) = "1"
.Row = 1
.Col = 0
.CellFontSize = 10


.TextMatrix(2, 0) = "*"
.Row = 2
.Col = 0
.CellAlignment = 7
.CellFontSize = 10

End With

End Sub

Private Sub AddNewItem()
flxTabla.AddItem ("")
With flxTabla
.TextMatrix(.rows - 2, 0) = .rows - 2
.Row = .rows - 2
.Col = 0
.CellFontSize = 10

.TextMatrix(.rows - 1, 0) = "*"
.Row = .rows - 1
.Col = 0
.CellAlignment = 7
.CellFontSize = 10
End With
End Sub

Private Sub CalcRegSup()
Dim iMax%
Dim objColl As Collection
Dim xArray() As Single
Dim yArray() As Single
Dim GXvalues As Variant

'Longitud de matris
iMax = flxTabla.rows - 2

'Redim matris
ReDim xArray(iMax)
ReDim yArray(iMax)

With flxTabla
    Set objColl = New Collection
    'Insertar a Xs
    For i = 1 To iMax
        xArray(i) = .TextMatrix(i, 1)
    Next i
    'Insertar Ys
    For j = 1 To iMax
        yArray(j) = .TextMatrix(j, 2)
    Next j
End With
Call objColl.Add(xArray)
Call objColl.Add(yArray)

'Llamar funcion regresion superior
iMax = CInt(txtPoly.Text)
Call RegresionSup(iMax, objColl, GXvalues)
Call ShowGXvalues(GXvalues)
Set objColl = Nothing

End Sub

Private Sub ShowGXvalues(sngGX As Variant)
Dim i%, iMaxArray%, iX As Single, iExp%
Dim sngGXval As Single

With flxTabla
'Get Max Array position
iMaxArray = UBound(sngGX)
For j = 1 To (.rows - 2)
    For i = 1 To iMaxArray
    
    iExp = i - 1
    iX = CSng(.TextMatrix(j, 1))
    'add value to G(x) column ''' g(x)=a+ax+ax^2+...+ax^n
    sngGXval = sngGXval + ((iX ^ iExp) * sngGX(i, iMaxArray + 1))
    Next i
    .TextMatrix(j, 3) = CStr(sngGXval)
    sngGXval = 0
Next j
End With
flexInter.cols = 2
flexInter.rows = (iMaxArray + 1)
For i = 1 To iMaxArray
    flexInter.TextMatrix(i, 1) = CStr(sngGX(i, iMaxArray + 1))
Next i

End Sub

Private Sub IntLagrange()
Dim iMax%
Dim sngYRES As Single
Dim sngWRES As Single
Dim Z As Single
Dim XA As Single
On Error GoTo ErrorMessage
XA = CSng(txtInter)


With flxTabla
iMax = .rows - 2   'cantidad de puntos

If (XA < .TextMatrix(1, 1)) Or (XA > .TextMatrix(iMax, 1)) Then
    MsgBox ("Advertencia: el valor se encuentra en el rango de extrapolación")
End If
sngYRES = 0
For i = 1 To iMax
    Z = 1#
    For j = 1 To iMax
        If (i <> j) Then
            Z = Z * (XA - CSng(.TextMatrix(j, 1))) / (CSng(.TextMatrix(i, 1)) - CSng(.TextMatrix(j, 1)))
        End If
    Next j
sngYRES = sngYRES + (Z * CSng(.TextMatrix(i, 2)))
Next i
End With
Label4.Caption = CStr(sngYRES)
ErrorMessage:
If (Err.Number >= 0) Then
    If (Err.Number = 11) Then
    MsgBox ("Division por cero!")
    End If
End If
End Sub

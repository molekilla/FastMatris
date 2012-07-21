VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBairstow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bonus - Metodo de Bairstow"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   6240
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtN 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Text            =   "4"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Metodo de Bairstow"
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8295
      Begin VB.CommandButton cmdCalc 
         Caption         =   "Calcular"
         Height          =   375
         Left            =   5520
         TabIndex        =   11
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit \ Terminar"
         Height          =   375
         Left            =   6840
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtQ 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1080
         TabIndex        =   9
         Text            =   "-1"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txtP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   8
         Text            =   "-1"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtFunction 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Text            =   "x^4-1,1x^3+2,3x^2+0,5x+3,3"
         Top             =   480
         Width           =   7095
      End
      Begin VB.Label Label2 
         Caption         =   $"frmBairstow.frx":0000
         Height          =   615
         Left            =   2040
         TabIndex        =   13
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "q:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "p:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "n:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Función:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxTabla 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3836
      _Version        =   65541
      Rows            =   3
      Cols            =   6
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmBairstow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public objTablas As Collection
Private Sub cmdCalc_Click()
Dim objV As Variant
Dim i%
Set objTablas = New Collection
List1.Clear
objV = Bairstow1(txtFunction(0), CInt(txtP(2)), CInt(txtQ(3)))
objTablas.Add (objV)
List1.AddItem "Tabla1"
For i = 1 To (CInt(txtN(1)) - 1)
    objV = Bairstow2(txtFunction(0), objV)
    objTablas.Add (objV)
    List1.AddItem ("Tabla" & CStr(i + 1))
Next i
End Sub

Private Sub cmdTable1_Click()
ShowTable (1)
End Sub

Private Sub cmdExit_Click()
Unload Me

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
.TextMatrix(0, 1) = "p"
.TextMatrix(0, 2) = "q"
.TextMatrix(0, 3) = "a"
.TextMatrix(0, 4) = "b"
.TextMatrix(0, 5) = "c"



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


Private Sub ShowTable(idx As Integer)
Dim i%, iMaxArray%, iX As Single, iExp%
Dim objTable As Variant, j%

objTable = objTablas(idx)
With flxTabla
    'Get Max Array position
    iMaxArray = UBound(objTablas(idx))
    .rows = iMaxArray + 1
    For j = 1 To iMaxArray
        For i = 1 To 5
            'iExp = i - 1
            .TextMatrix(j, i) = CStr(objTable(j, i))
        Next i
            '.TextMatrix(j, 3) = CStr(sngGXval)
            'sngGXval = 0
    Next j
End With
End Sub

Private Sub List1_Click()
ShowTable (List1.ListIndex + 1)
End Sub

VERSION 5.00
Begin VB.Form frmEcDiferencial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ecuaciones Diferenciales"
   ClientHeight    =   4065
   ClientLeft      =   2415
   ClientTop       =   1830
   ClientWidth     =   8280
   Icon            =   "frmEcDiferencial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   8280
   Begin VB.TextBox txtResult 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   6120
      TabIndex        =   21
      Top             =   2280
      Width           =   1815
   End
   Begin VB.OptionButton optRungeKutta2 
      Caption         =   "Runge - Kutta de II orden"
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   3480
      Width           =   3255
   End
   Begin VB.OptionButton optRungeKutta3 
      Caption         =   "Runge - Kutta de III orden"
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   3120
      Width           =   3255
   End
   Begin VB.OptionButton optRungeKutta4 
      Caption         =   "Runge - Kutta de IV orden"
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   2760
      Width           =   3255
   End
   Begin VB.OptionButton optEulerMod 
      Caption         =   "Euler Modificado / Predictor -Corrector"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   2400
      Width           =   3255
   End
   Begin VB.OptionButton optEuler 
      Caption         =   "Euler Normal"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   2040
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.TextBox txtSeekValue 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6120
      TabIndex        =   15
      Text            =   "0,1"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtH 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6840
      TabIndex        =   13
      Text            =   "0,02"
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valores iniciales"
      Height          =   1215
      Left            =   6120
      TabIndex        =   9
      Top             =   120
      Width           =   1935
      Begin VB.TextBox txtInitY 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "h="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   " Y (0)="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtY 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      TabIndex        =   6
      Text            =   "1"
      Top             =   1395
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   5
      Text            =   "1"
      Top             =   1395
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit / Terminar"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Resultado"
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Valor a buscar"
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   " f (x,y)="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "X  +"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Nota: Funciones exponenciales no son posibles."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "Ecuaciones Diferenciales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmEcDiferencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalc_Click()
    If optEuler = True Then
        txtResult = Euler(txtX, txtY, CSng(txtInitY), CSng(txtH), CSng(txtSeekValue))
    ElseIf optEulerMod = True Then
        txtResult = EulerModificado(txtX, txtY, CSng(txtInitY), CSng(txtH), CSng(txtSeekValue))
    ElseIf optRungeKutta4 = True Then
        txtResult = RungeKuttaIV(txtX, txtY, CSng(txtInitY), CSng(txtH), CSng(txtSeekValue))
    ElseIf optRungeKutta3 = True Then
        txtResult = RungeKuttaIII(txtX, txtY, CSng(txtInitY), CSng(txtH), CSng(txtSeekValue))
    ElseIf optRungeKutta2 = True Then
       txtResult = RungeKuttaII(txtX, txtY, CSng(txtInitY), CSng(txtH), CSng(txtSeekValue))
    End If
End Sub

Private Sub Command1_Click()
    Me.Hide
    Unload Me
End Sub

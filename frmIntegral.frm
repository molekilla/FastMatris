VERSION 5.00
Begin VB.Form frmIntegral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Integrales"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7665
   Icon            =   "frmIntegral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit / Terminar"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   3000
      Width           =   1935
   End
   Begin VB.OptionButton optTrapecio 
      Caption         =   "Formula del trapecio"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   2400
      Width           =   2775
   End
   Begin VB.OptionButton optSimpson 
      Caption         =   "Formula de Simpson"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   2040
      Value           =   -1  'True
      Width           =   3375
   End
   Begin VB.TextBox txtResult 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calcule integral"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   6
      Text            =   "1"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtFormula 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Text            =   "3x^3+5x-1"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Text            =   "1"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "0"
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Nota: Funciones exponenciales y trigonometricas no son posibles."
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Label Label6 
      Caption         =   "Integrales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Ecuación"
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Resultado"
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "N # de rectangulos:"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "dx  ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ò"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   48
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmIntegral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalc_Click()
Call Integral.ParseFunction(txtFormula)

If optTrapecio = True Then
    txtResult = Trapecio(CSng(txtB), CSng(txtA), CInt(txtQty))
Else
    If (CInt(txtQty) Mod 2) = 0 Then
        'par
        'simpson 1/3
        txtResult = SimpsonPar(CSng(txtB), CSng(txtA), CInt(txtQty))
    ElseIf (CInt(txtQty) Mod 3) = 0 Then
      'impar simpson 3/8
        txtResult = Simpson3(CSng(txtB), CSng(txtA), CInt(txtQty))
    Else
        txtResult = ""
    End If
End If
End Sub

Private Sub Command1_Click()
frmIntegral.Hide
Unload frmIntegral
End Sub

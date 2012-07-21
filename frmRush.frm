VERSION 5.00
Begin VB.Form frmRush 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre Fast Matris y su autor"
   ClientHeight    =   6120
   ClientLeft      =   285
   ClientTop       =   1110
   ClientWidth     =   9000
   Icon            =   "frmRush.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9000
   Begin VB.Frame Frame1 
      Caption         =   "The Genius Mob saga continues..."
      Height          =   5895
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   6615
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000005&
         Height          =   5535
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmRush.frx":030A
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      Picture         =   "frmRush.frx":08DA
      ScaleHeight     =   2145
      ScaleWidth      =   2025
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Quotes: There is nothing without your own will.           "
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Musical Mind: Rap, Hip-Hop for life made my life hype, cause every opportunity to rap was just passing me by."
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Nicknames: Rush Molekilla, Main Man, Molekilla, Method Man, Nigga Rush."
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Estudios: Rollingby High School back in Sweden"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre: Rogelio Morrell"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
End
Attribute VB_Name = "frmRush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture1_Click()
MsgBox ("Secret Words from an ex-Genius mob! Prepare to erase your C:\ drive!")
Text1.Text = "And it DON'T STOP..." & vbCrLf & "Hey, naci en Panama el 29 de Diciembre de 1977, tengo 21 años, y todo lo que quiero en este mundo es MONEY, MI GUIAL DE LOS SUEÑOS Y...BE HAPPY. Llegue a Swedenlove un 18 de Enero de 1989, un dia frio con nieve. Vivi 7 años con algunos mese en Swedenlove y aprendi que la vida es un tesoro perdido (titulo de mi primer libro, todavia no publicado). Los inviernos frios y depresivos me enseñaron las otras cualidades, que junto con el rap, me convirtieron en un man que todo lo que veia, gracias a mi educacion, debe tener una explicacion racional, en forma cientifica o cultural. De ahi salio todo my flow for the music I love, for the things I like to write and philosophy. Yo! Swedenlove was the shit. Made my hopes come thru, in a town there dreams won't come true. I represent sometimes all my niggas back in Swedenlove. For those whom didn't made it, I was there and for those who made I " & _
" was there too. Swedenlove era lo maximo. Dedico todo esto a ellos, a mi city de Åkersberga. Peace and is all in you." & vbCrLf & "-Dreams ain't always what you expect, but neither is the reality. Just go with the flow and don't let you fool by the things-" & vbCrLf & "    -Rush Molekilla '98-"
End Sub


Private Sub Text1_Click()

Picture1.SetFocus
End Sub

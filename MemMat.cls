VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MemMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'local variable(s) to hold property value(s)
Private mvarDimension As Integer 'local copy
Private mvarName As String 'local copy
Private mvarMatArray As Object 'local copy
Public Property Set MatArray(vData As Object)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MatArray = 5
   Set mvarMatArray = vData
End Property


Public Property Get MatArray() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MatArray
   Set MatArray = mvarMatArray
End Property



Public Property Let Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name = 5
    mvarName = vData
End Property


Public Property Get Name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name
    Name = mvarName
End Property



Public Property Let Dimension(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Dimension = 5
    mvarDimension = vData
End Property


Public Property Get Dimension() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Dimension
    Dimension = mvarDimension
End Property




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParseFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"FunctionX"
Option Explicit

'Public sngConstant As Single
'Public bVariable As Boolean
'Public sngExponent As Single
'Public sOperator As String
'local variable(s) to hold property value(s)
Private mvarsngConstant As Single 'local copy
Private mvarbVariable As Boolean 'local copy
Private mvarsngExponent As Single 'local copy
Private mvarsOperator As String 'local copy






Public Property Let sOperator(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sOperator = 5
    mvarsOperator = vData
End Property


Public Property Get sOperator() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sOperator
    sOperator = mvarsOperator
End Property



Public Property Let sngExponent(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sngExponent = 5
    mvarsngExponent = vData
End Property


Public Property Get sngExponent() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sngExponent
    sngExponent = mvarsngExponent
End Property



Public Property Let bVariable(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bVariable = 5
    mvarbVariable = vData
End Property


Public Property Get bVariable() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bVariable
    bVariable = mvarbVariable
End Property



Public Property Let sngConstant(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sngConstant = 5
    mvarsngConstant = vData
End Property


Public Property Get sngConstant() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sngConstant
    sngConstant = mvarsngConstant
End Property




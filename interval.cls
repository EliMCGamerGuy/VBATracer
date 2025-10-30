VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "interval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarmin As Double 'local copy
Private mvarmax As Double 'local copy



Public Property Let max(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.max = 5
    mvarmax = vData
End Property


Public Property Get max() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.max
    max = mvarmax
End Property



Public Property Let min(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.min = 5
    mvarmin = vData
End Property


Public Property Get min() As Double
Attribute min.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.min
    min = mvarmin
End Property



Private Sub Class_Initialize()
    Me.max = -infinity()
    Me.min = infinity()
End Sub



Public Function surrounds(x As Double) As Boolean
    surrounds = Me.min < x And x < Me.max
End Function



Public Function contains(x As Double) As Boolean
    surrounds = Me.min <= x And x <= Me.max
End Function



Public Function size()
    size = Me.max - Me.min
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarorigin As vec3 'local copy
Private mvardirection As vec3 'local copy
Public Function at(t) As vec3
    Set at = vecAddVec(Me.origin, vecMul(Me.direction, CDbl(t)))
End Function

Public Property Set direction(ByVal vData As vec3)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.direction = Form1
    Set mvardirection = vData
End Property


Public Property Get direction() As vec3
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.direction
    Set direction = mvardirection
End Property



Public Property Set origin(ByVal vData As vec3)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.origin = Form1
    Set mvarorigin = vData
End Property


Public Property Get origin() As vec3
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.origin
    Set origin = mvarorigin
End Property



' Return the larger value.
Private Function Max(val1, val2)
    If val1 > val2 Then
        Max = val1
    Else
        Max = val2
    End If
End Function



' Return the smaller value.
Private Function Min(val1, val2)
    If val1 < val2 Then
        Min = val1
    Else
        Min = val2
    End If
End Function



Private Function vecAddVec(v1 As vec3, v2 As vec3) As vec3
    Set vecAddVec = v1.addvec(v2)
End Function



Private Function vecSubVec(v1 As vec3, v2 As vec3) As vec3
    Set vecSubVec = v1.subvec(v2)
End Function



Private Function vecMul(v1 As vec3, t) As vec3
    Set vecMul = v1.mul(CDbl(t))
End Function



Private Function vecDiv(v1 As vec3, t) As vec3
    Set vecDiv = v1.div(CDbl(t))
End Function



Private Function vecDot(v1 As vec3, v2 As vec3) As Double
    vecDot = v1.dot(v2)
End Function



Private Function vecMulVec(v1 As vec3, v2 As vec3) As vec3
    Set vecMulVec = v1.mulvec(v2)
End Function




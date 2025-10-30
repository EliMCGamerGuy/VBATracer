VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hit_record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarp As vec3 'local copy
Private mvarnormal As vec3 'local copy
Private mvart As Double 'local copy
Private mvarfront_face As Boolean 'local copy
Public Property Let t(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.t = 5
    mvart = vData
End Property


Public Property Get t() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.t
    t = mvart
End Property



Public Property Set normal(ByVal vData As vec3)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.normal = Form1
    Set mvarnormal = vData
End Property


Public Property Get normal() As vec3
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.normal
    Set normal = mvarnormal
End Property



Public Property Set p(ByVal vData As vec3)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.p = Form1
    Set mvarp = vData
End Property


Public Property Get p() As vec3
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.p
    Set p = mvarp
End Property


Public Property Let front_face(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.front_face = True
    mvarfront_face = vData
End Property


Public Property Get front_face() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.front_face
    front_face = mvarfront_face
End Property


Public Sub set_face_normal(r As ray, outward_normal As vec3)
    ' Sets the hit record normal vector.
    ' NOTE: the parameter 'outward_normal' is assumed to be normalized.
    
    Me.front_face = vecDot(r.direction, outward_normal) < 0
    'If front_face Then Set Me.normal = outward_normal Else Set Me.normal = vecMul(outward_normal, -1)
    Set Me.normal = IIf(Me.front_face, outward_normal, vecMul(outward_normal, -1))
End Sub




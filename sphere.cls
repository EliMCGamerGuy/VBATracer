VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sphere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Implements hittable

'local variable(s) to hold property value(s)
Private mvarcenter As vec3 'local copy
Private mvarradius As Double 'local copy
Public Property Let radius(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.radius = 5
    mvarradius = vData
End Property


Public Property Get radius() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.radius
    radius = mvarradius
End Property



Public Property Set center(ByVal vData As vec3)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.center = Form1
    Set mvarcenter = vData
End Property


Public Property Get center() As vec3
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.center
    Set center = mvarcenter
End Property



Public Function hit(r As ray, ray_tmin As Double, ray_tmax As Double, ByRef rec As hit_record) As Boolean
    ' there's a point at which coding just turns into straight up wizardry
    Me.radius = CDbl(Max(0, Me.radius))
        
    Dim oc As New vec3
    Set oc = vecSubVec(center, r.origin)
    
    a = vecDot(r.direction, r.direction)
    h = vecDot(r.direction, oc)
    c = vecDot(oc, oc) - radius * radius
    discriminant = h * h - a * c
    
    ' return -1 break if discriminant is less than 0.
    If discriminant < 0 Then
        hit = False
        Exit Function
    End If
    
    sqrtd = Sqr(discriminant)
    
    ' Find the nearest root that lies in the acceptable range.
    Dim Root As Double
    Root = (h - sqrtd) / a
    If (Root <= ray_tmin) Or (ray_tmax <= Root) Then
        Root = (h + sqrtd) / a
        If (Root <= ray_tmin) Or (ray_tmax <= Root) Then
            hit = False
            Exit Function
        End If
    End If
    
    rec.t = Root
    Set rec.p = r.at(rec.t)
    Set rec.normal = vecDiv(vecSubVec(rec.p, Me.center), Me.radius)
    Dim outward_normal As New vec3
    Set outward_normal = vecDiv(vecSubVec(rec.p, Me.center), Me.radius)
    rec.set_face_normal r, outward_normal
    
    hit = True
    ' with sphere.hit(), i think we passed that point
End Function


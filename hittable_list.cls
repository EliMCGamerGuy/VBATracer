VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hittable_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Implements hittable

'local variable(s) to hold property value(s)
Private mvarobjects As New Dictionary 'local copy
Public Property Let objects(ByVal vData As Dictionary)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.objects = 5
    mvarobjects = vData
End Property


Public Property Set objects(ByVal vData As Dictionary)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objects = Form1
    Set mvarobjects = vData
End Property


Public Property Get objects() As Dictionary
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objects
    If IsObject(mvarobjects) Then
        Set objects = mvarobjects
    Else
        Set objects = mvarobjects
    End If
End Property



Public Function hit(r As ray, ray_tmin, ray_tmax, ByRef rec As hit_record) As Boolean
    Dim temp_rec As New hit_record
    Dim hit_anything
    hit_anything = False
    closest_so_far = CDbl(ray_tmax)
    
    For Each objKey In Me.objects.Keys
        If Me.objects(objKey).hit(r, CDbl(ray_tmin), CDbl(closest_so_far), temp_rec) Then
            hit_anything = True
            closest_so_far = temp_rec.t
            Set rec = temp_rec
        End If
    Next
    
    hit = hit_anything
End Function



Public Sub add(obj)
    Me.objects.add CStr(Int(Me.objects.Count)), obj
End Sub



Public Sub clear()
    Me.objects.RemoveAll
End Sub

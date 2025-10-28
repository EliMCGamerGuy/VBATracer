Attribute VB_Name = "VBATracerGlobals"
' Return the larger value.
Function Max(val1, val2)
    'If val1 > val2 Then
    '    Max = val1
    'Else
    '    Max = val2
    'End If
    
    Max = val2 * -(val1 < val2) + val1 * -(val2 <= val1)
End Function



' Return the smaller value.
Function Min(val1, val2)
    'If val1 < val2 Then
    '    Min = val1
    'Else
    '    Min = val2
    'End If
    
    Min = val1 * -(val1 < val2) + val2 * -(val2 <= val1)
End Function



Function vecAddVec(v1 As vec3, v2 As vec3) As vec3
    Set vecAddVec = v1.addvec(v2)
End Function



Function vecSubVec(v1 As vec3, v2 As vec3) As vec3
    Set vecSubVec = v1.subvec(v2)
End Function



Function vecMul(v1 As vec3, t) As vec3
    Set vecMul = v1.mul(CDbl(t))
End Function



Function vecDiv(v1 As vec3, t) As vec3
    Set vecDiv = v1.div(CDbl(t))
End Function



Function vecDot(v1 As vec3, v2 As vec3) As Double
    vecDot = v1.dot(v2)
End Function



Function vecMulVec(v1 As vec3, v2 As vec3) As vec3
    Set vecMulVec = v1.mulvec(v2)
End Function



Private Function gamma2(color As Double)
    'gamma2 = Sqr(color) ' Gamma 2.0? What? I dunno.
    
    ' what would happen if i just multiplied by a fraction?
End Function


' CONSTRUCTORS


Function vec3(x, y, z) As vec3
    Dim temp As New vec3
    temp.x = CDbl(x)
    temp.y = CDbl(y)
    temp.z = CDbl(z)
    Set vec3 = temp
End Function



Function point3(x, y, z) As vec3
    Dim temp As New vec3
    temp.x = CDbl(x)
    temp.y = CDbl(y)
    temp.z = CDbl(z)
    Set point3 = temp
End Function



Function ray(origin As vec3, dir As vec3) As ray
    Dim temp As New ray
    Set temp.origin = origin
    Set temp.direction = dir
    Set ray = temp
End Function



Function hit_record() As hit_record
    Dim temp As New hit_record
    Set hit_record = temp
End Function



Function sphere(center As vec3, radius) As sphere
    Dim temp As New sphere
    Set temp.center = center
    temp.radius = CDbl(radius)
    Set sphere = temp
End Function



' Return the larger value.
Function Max_BT(val1, val2)
    If val1 > val2 Then
        Max_BT = val1
    Else
        Max_BT = val2
    End If
    
    'Max = val2 * -(val1 < val2) + val1 * -(val2 <= val1)
End Function



' Return the smaller value.
Function Min_BT(val1, val2)
    If val1 < val2 Then
        Min_BT = val1
    Else
        Min_BT = val2
    End If
    
    'Min = val1 * -(val1 < val2) + val2 * -(val2 <= val1)
End Function



Function infinity()
    infinity = 1E+30
End Function



Function pi()
    pi = 3.14159265358979
End Function



Function degrees_to_radian(degrees)
    degrees_to_radian = CDbl(degrees) * pi() / 180#
End Function


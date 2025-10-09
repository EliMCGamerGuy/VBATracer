VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Function hit_sphere(center As vec3, radius, r As ray) As Boolean
    Dim oc As New vec3
    Set oc = vecSubVec(center, r.origin)
    
    a = vecDot(r.direction, r.direction)
    b = -2# * vecDot(r.direction, oc)
    c = vecDot(oc, oc) - radius * radius
    discriminant = b * b - 4 * a * c
    hit_sphere = (discriminant >= 0)
End Function



Private Function ray_color(r As ray) As vec3
    If (hit_sphere(vec3(0, 0, -1), 0.5, r)) Then
        Set ray_color = vec3(1, 0, 0)
        Exit Function
    End If
    
    ' Let's make a little gradient!
    Dim unit_dir As New vec3
    Set unit_dir = r.direction.normalize()
    
    'Debug.Print "r pos x" & CStr(r.origin.x) & " y" & CStr(r.origin.y) & " z" & CStr(r.origin.z)
    'Debug.Print "r dir x" & CStr(r.direction.x) & " y" & CStr(r.direction.y) & " z" & CStr(r.direction.z)
    'Debug.Print "unit dir x" & CStr(unit_dir.x) & " y" & CStr(unit_dir.y) & " z" & CStr(unit_dir.z)
    
    
    a = 0.5 * (unit_dir.y + 1)
    Set ray_color = vecAddVec(vecMul(vec3(1, 1, 1), (1 - a)), vecMul(vec3(0.5, 0.7, 1), a))
    ' vec3(1,1,1) * (1.0 - a) + vec3(0.5,0.7,1) * a
    ' it's a little sloppy but it works
End Function



' "We have int main() at home"
Private Sub Class_Initialize()
    aspect = 15 / 9
    image_width = 400
    
    ' Get image height based on aspect, and make sure it's >= 1
    Dim image_height As Integer
    image_height = Int(image_width / aspect)
    image_height = Max(image_height, 1)
    
    Window.paintBox.Width = image_width * 15
    Window.paintBox.Height = image_height * 15
    
    ' Makes the progress bar work, ignore.
    ' Requires image_width and image_height to work.
    Window.renderProgress.Max = (image_width + 1) * (image_height + 1)
    Window.renderProgress.Value = 0
    
    
    ' Camera Settings
    
    focal_length = 1
    viewport_height = 2
    viewport_width = viewport_height * (image_width / image_height)
    Dim camera_center As New vec3
    Set camera_center = point3(0, 0, 0)
    
    
    ' Calculate vectors across horizontal and down vertical viewport sides.
    Dim viewport_u As New vec3
    Dim viewport_v As New vec3
    Set viewport_u = vec3(viewport_width, 0, 0)
    Set viewport_v = vec3(0, -viewport_height, 0)
    
    
    ' Calculate the horizontal and vertical delta vectors from pix to pix.
    Dim pixel_delta_u As New vec3
    Dim pixel_delta_v As New vec3
    Set pixel_delta_u = vecDiv(viewport_u, image_width)
    Set pixel_delta_v = vecDiv(viewport_v, image_height)
    
    
    ' Calculate the point3 location of the upper left pix.
    ' Are you ready?
    Dim viewport_ul As New vec3
    Dim pix00_loc As New vec3
    Set viewport_ul = vecSubVec(vecSubVec(vecSubVec(camera_center, vec3(0, 0, focal_length)), vecDiv(viewport_u, 2)), vecDiv(viewport_v, 2))
    ' viewport_ul = camera_center - vec3(0,0,focal_length) - viewport_u/2 - viewport_v/2
    Set pix00_loc = vecAddVec(viewport_ul, vecMul(vecAddVec(pixel_delta_u, pixel_delta_v), 0.5))
    ' pix00_loc = viewport_ul + (pixel_delta_u + pixel_delta_v) * 0.5
    
    
    ' Render!
    
    ' Variable setup.
    Dim pixelColor As New vec3
    Dim pix_center As New vec3
    Dim ray_dir As New vec3
    Dim r As New ray
    
    
    For j = 0 To image_height - 1 ' For every pixel
        For i = 0 To image_width - 1
        
            Set pix_center = vecAddVec(pix00_loc, vecAddVec(vecMul(pixel_delta_u, i), vecMul(pixel_delta_v, j)))
            ' pix_center = pix00_loc + (pixel_delta_u * i) + (pixel_delta_v * j)
            
            ' ray direction is the pixel's center minus the camera center, giving us dir
            Set ray_dir = vecSubVec(pix_center, camera_center)
            Set r = ray(camera_center, ray_dir) ' MAKE! THE! RAY!
            
            ' get color from ray
            Set pixelColor = ray_color(r)
            
            ' This allows Windows to run required routines so that the
            ' window doesn't freeze during rendering. It also refreshes the
            ' render box and progressbar, so it's rather useful.
            If i Mod 128 = 0 Then
                DoEvents
            End If
            
            
            ' Paint the pixel to the render box.
            Window.drawPixel i, j, (pixelColor.convToRGB(True, Int(i), Int(j))) ' / 2
            
            ' Quality of life progressbar. Shows render progress.
            Window.renderProgress.Value = Window.renderProgress.Value + 1
        Next i
    Next j
    
    ' Visually show the progressbar as full when done rendering
    Window.renderProgress.Value = Window.renderProgress.Max
End Sub



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



Private Function vec3(x, y, z) As vec3
    Dim temp As New vec3
    temp.x = CDbl(x)
    temp.y = CDbl(y)
    temp.z = CDbl(z)
    Set vec3 = temp
End Function



Private Function point3(x As Double, y As Double, z As Double) As vec3
    Dim temp As New vec3
    temp.x = x
    temp.y = y
    temp.z = z
    Set point3 = temp
End Function



Private Function ray(origin As vec3, dir As vec3) As ray
    Dim temp As New ray
    Set temp.origin = origin
    Set temp.direction = dir
    Set ray = temp
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



Private Function gamma2(color As Double)
    'gamma2 = Sqr(color) ' Gamma 2.0? What? I dunno.
    
    ' what would happen if i just multiplied by a fraction?
    
    gamma2 = color * 0.75
End Function

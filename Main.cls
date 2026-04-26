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
Attribute VB_Ext_KEY = "Member0" ,"main"
'Private Function hit_sphere(center As vec3, radius, r As ray) As Double
'    Dim oc As New vec3
'    Set oc = vecSubVec(center, r.origin)
'
'    a = vecDot(r.direction, r.direction)
'    h = vecDot(r.direction, oc)
'    c = vecDot(oc, oc) - radius * radius
'    discriminant = h * h - a * c
'
'    ' WHAT ARE WE EVEN RETURNING ANYMORE
'    If discriminant < 0 Then
'        hit_sphere = -1#
'    Else
'        hit_sphere = (h - Sqr(CDbl(discriminant))) / a
'    End If
'End Function



Private Function ray_color(r As ray, ByVal world) As vec3
End Function



' "We have int main() at home"
Public Sub main()
    
    'aspect = 15 / 9
    'image_width = 400
    'image_height = Int(image_width / aspect)
    'image_height = max(image_height, 1)
    
    
    ' World Settings
    
    Dim world As New hittable_list
    
    world.add sphere(vec3(0, 0, -1), 0.5)
    world.add sphere(vec3(0, -100.5, -1), 100)
    world.add sphere(vec3(1, 0, -1.5), 0.5)
    world.add sphere(vec3(-1, 0, -1.5), 0.5)
    
    
    ' Create the camera
    Dim cam As New camera
    
    
    ' Camera Settings
    
    cam.aspect_ratio = 16 / 9
    cam.image_width = 400
    
    
    'focal_length = 1
    'viewport_height = 2
    'viewport_width = viewport_height * (image_width / image_height)
    'Dim camera_center As New vec3
    'Set camera_center = point3(0, 0, 0)
    
    
    ' Calculate vectors across horizontal and down vertical viewport sides.
    'Dim viewport_u As New vec3
    'Dim viewport_v As New vec3
    'Set viewport_u = vec3(viewport_width, 0, 0)
    'Set viewport_v = vec3(0, -viewport_height, 0)
    
    
    ' Calculate the horizontal and vertical delta vectors from pix to pix.
    'Dim pixel_delta_u As New vec3
    'Dim pixel_delta_v As New vec3
    'Set pixel_delta_u = vecDiv(viewport_u, image_width)
    'Set pixel_delta_v = vecDiv(viewport_v, image_height)
    
    
    ' Calculate the point3 location of the upper left pix.
    ' Are you ready?
    'Dim viewport_ul As New vec3
    'Dim pix00_loc As New vec3
    'Set viewport_ul = vecSubVec(vecSubVec(vecSubVec(camera_center, vec3(0, 0, focal_length)), vecDiv(viewport_u, 2)), vecDiv(viewport_v, 2))
    ' viewport_ul = camera_center - vec3(0,0,focal_length) - viewport_u/2 - viewport_v/2
    'Set pix00_loc = vecAddVec(viewport_ul, vecMul(vecAddVec(pixel_delta_u, pixel_delta_v), 0.5))
    ' pix00_loc = viewport_ul + (pixel_delta_u + pixel_delta_v) * 0.5
    
    
    ' Render!
    cam.render world
    
    
    ' Variable setup.
    'Dim pixelColor As New vec3
    'Dim pix_center As New vec3
    'Dim ray_dir As New vec3
    'Dim r As New ray
    
    'For j = 0 To image_height - 1 ' For every pixel
    '    For i = 0 To image_width - 1
    '
    '        Set pix_center = vecAddVec(pix00_loc, vecAddVec(vecMul(pixel_delta_u, i), vecMul(pixel_delta_v, j)))
    '        ' pix_center = pix00_loc + (pixel_delta_u * i) + (pixel_delta_v * j)
    '
    '        ' ray direction is the pixel's center minus the camera center, giving us dir
    '        Set ray_dir = vecSubVec(pix_center, camera_center)
    '        Set r = ray(camera_center, ray_dir) ' MAKE! THE! RAY!
    '
    '        ' get color from ray
    '        Set pixelColor = ray_color(r, world)
    '
    '        ' This allows Windows to run required routines so that the
    '        ' window doesn't freeze during rendering. It also refreshes the
    '        ' render box and progressbar, so it's rather useful.
    '        If i = 0 Then
    '            DoEvents
    '        End If
    '
    '
    '        ' Paint the pixel to the render box.
    '        Window.drawPixel i, j, (pixelColor.convToRGB(True, Int(i), Int(j))) ' / 2
    '
    '        ' Quality of life progressbar. Shows render progress.
    '        Window.renderProgress.Value = Window.renderProgress.Value + 1
    '    Next i
    'Next j
    '
    ' Visually show the progressbar as full when done rendering
    'Window.renderProgress.Value = Window.renderProgress.max
    '
    'Dim newCurrentTime
    'newCurrentTime = Now
    'Debug.Print (Format(newCurrentTime - currentTime, "hh:mm:ss"))
End Sub



Public Sub branchlessMaxProfiler()
    
    Randomize (Second(Now))
    
    
    Debug.Print ("Running profiler for standard test...")
    
    Dim tStart As Double, tEnd As Double
    Dim totalMax As Double
    Dim countMax As Long
    
    
    Dim TBT
    TBT = Now
    totalMax = 0: countMax = 0
    
    For i = 0 To 5000000
        tStart = PerfNow
    
        temp = Max_BT(Rnd() * 500, Rnd() * 500)
        
        tEnd = PerfNow
        totalMax = totalMax + (tEnd - tStart)
        countMax = countMax + 1
    Next
    
    Dim TAT
    TAT = Now
    Debug.Print ("Time for 5m Max")
    Debug.Print (Format(TAT - TBT, "hh:mm:ss"))
    Debug.Print "Max avg (ms):"; Format((totalMax / countMax) * 1000, "0.00 ms")
    Debug.Print "Avg (탎): " & Format((totalMax / countMax) * 1000000, "0.00 탎")
    
    
    Debug.Print ("Running profiler for branchless test...")
    
    
    TBT = Now
    totalMax = 0: countMax = 0
    
    For i = 0 To 5000000
        tStart = PerfNow
    
        temp = max(Rnd() * 500, Rnd() * 500)
        
        tEnd = PerfNow
        totalMax = totalMax + (tEnd - tStart)
        countMax = countMax + 1
    Next
    
    TAT = Now
    Debug.Print ("Time for 5m Branchless Max")
    Debug.Print (Format(TAT - TBT, "hh:mm:ss"))
    Debug.Print "Max avg (ms):"; Format((totalMax / countMax) * 1000, "0.00 ms")
    Debug.Print "Avg (탎): " & Format((totalMax / countMax) * 1000000, "0.00 탎")
    
End Sub

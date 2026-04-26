VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "camera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvaraspect_ratio As Double
Private mvarimage_width As Integer
Private mvarimage_height As Integer
Private mvarcenter As vec3
Private mvarpixel00_loc As vec3
Private mvarpixel_delta_u As vec3
Private mvarpixel_delta_v As vec3



Public Property Let aspect_ratio(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.aspect_ratio = 5
    mvaraspect_ratio = vData
End Property


Public Property Get aspect_ratio() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.aspect_ratio
    aspect_ratio = mvaraspect_ratio
End Property



Private Property Let image_height(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.image_width = 5
    mvarimage_width = vData
End Property


Private Property Get image_height() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.image_height
    image_width = mvarimage_width
End Property



Public Property Let image_width(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.image_width = 5
    mvarimage_width = vData
End Property


Public Property Get image_width() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.image_width
    image_width = mvarimage_width
End Property



Public Function ray_color(r As ray, world As hittable_list) As vec3
    Dim rec As New hit_record
    If world.hit(r, interval(0, infinity()), rec) Then
        Set ray_color = vecMul(vecAddVec(rec.normal, vec3(1, 1, 1)), 0.5)
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



Public Sub render(world As hittable)
    ' Start recording render time
    Dim currentTime
    currentTime = Now
    
    ' Initialize!
    Me.init
    
    ' Render!
    
    ' Variable setup.
    Dim pixelColor As New vec3
    Dim pix_center As New vec3
    Dim ray_dir As New vec3
    Dim r As New ray
    
    For j = 0 To mvarimage_height - 1 ' For every pixel
        For i = 0 To mvarimage_width - 1
        
            Set pix_center = vecAddVec(mvarpixel00_loc, vecAddVec(vecMul(mvarpixel_delta_u, i), vecMul(mvarpixel_delta_v, j)))
            ' pix_center = pix00_loc + (pixel_delta_u * i) + (pixel_delta_v * j)
            
            ' ray direction is the pixel's center minus the camera center, giving us dir
            Set ray_dir = vecSubVec(pix_center, mvarcenter)
            Set r = ray(mvarcenter, ray_dir) ' MAKE! THE! RAY!
            
            ' get color from ray
            Set pixelColor = Me.ray_color(r, world)
            
            ' This allows Windows to run required routines so that the
            ' window doesn't freeze during rendering. It also refreshes the
            ' render box and progressbar, so it's rather useful.
            If i = 0 Then
                DoEvents
            End If
            
            
            ' Paint the pixel to the render box.
            Window.drawPixel i, j, (pixelColor.convToRGB(True, Int(i), Int(j))) ' / 2
            
            ' Quality of life progressbar. Shows render progress.
            Window.renderProgress.Value = Window.renderProgress.Value + 1
        Next i
    Next j
    
    ' Visually show the progressbar as full when done rendering
    Window.renderProgress.Value = Window.renderProgress.max
    
    ' Report render time
    Dim newCurrentTime
    newCurrentTime = Now
    Debug.Print (Format(newCurrentTime - currentTime, "hh:mm:ss"))
End Sub



Public Sub init()
    
    ' Get image height based on aspect, and make sure it's >= 1
    'Dim image_height As Integer
    mvarimage_height = Int(mvarimage_width / mvaraspect_ratio)
    mvarimage_height = max(mvarimage_height, 1)
    
    Window.paintBox.Width = (mvarimage_width * 15) + 32
    Window.paintBox.Height = (mvarimage_height * 15) + 32
    
    ' Makes the progress bar work, ignore.
    ' Requires image_width and image_height to work.
    Window.renderProgress.max = (mvarimage_width + 1#) * (mvarimage_height + 1#)
    Window.renderProgress.Value = 0
    
    
    ' Camera Settings
    
    focal_length = 1
    viewport_height = 2
    viewport_width = viewport_height * (mvarimage_width / mvarimage_height)
    Set mvarcenter = point3(0, 0, 0)
    
    
    ' Calculate vectors across horizontal and down vertical viewport sides.
    Dim viewport_u As New vec3
    Dim viewport_v As New vec3
    Set viewport_u = vec3(viewport_width, 0, 0)
    Set viewport_v = vec3(0, -viewport_height, 0)
    
    
    ' Calculate the horizontal and vertical delta vectors from pix to pix.
    'Dim pixel_delta_u As New vec3
    'Dim pixel_delta_v As New vec3
    Set mvarpixel_delta_u = vecDiv(viewport_u, mvarimage_width)
    Set mvarpixel_delta_v = vecDiv(viewport_v, mvarimage_height)
    
    
    ' Calculate the point3 location of the upper left pix.
    ' Are you ready?
    Dim viewport_ul As New vec3
    'Dim pix00_loc As New vec3
    Set viewport_ul = vecSubVec(vecSubVec(vecSubVec(mvarcenter, vec3(0, 0, focal_length)), vecDiv(viewport_u, 2)), vecDiv(viewport_v, 2))
    ' viewport_ul = camera_center - vec3(0,0,focal_length) - viewport_u/2 - viewport_v/2
    Set mvarpixel00_loc = vecAddVec(viewport_ul, vecMul(vecAddVec(mvarpixel_delta_u, mvarpixel_delta_v), 0.5))
    ' pix00_loc = viewport_ul + (pixel_delta_u + pixel_delta_v) * 0.5
    
    
End Sub

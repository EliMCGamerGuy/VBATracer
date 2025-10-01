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
Private Function ray_color(r As ray) As vec3
    ' Let's make a little gradient!
    Dim unit_dir As New vec3
    Set unit_dir = r.direction.normalize()
    
    
    a = 0.5 * (unit_dir.y + 1#)
    Set ray_color = vec3(1#, 1#, 1#).mul(1# - a).addvec(vec3(0.5, 0.7, 1#).mul(CDbl(a)))
End Function



' "We have int main() at home"
Private Sub Class_Initialize()
    aspect = 16 / 9
    image_width = 128
    
    ' Get image height based on aspect, and make sure it's >0
    Dim image_height As Integer
    image_height = Int(image_width / aspect)
    image_height = Max(image_height, 1)
    
    ' Makes the progress bar work, ignore.
    ' Requires image_width and image_height to work.
    Window.renderProgress.Max = (image_width + 1) * (image_height + 1)
    Window.renderProgress.Value = 0
    
    
    ' Camera Settings
    
    focal_length = 1#
    viewport_width = 2#
    viewport_height = viewport_width * (image_width / image_height)
    Dim camera_center As New vec3
    Set camera_center = point3(0, 0, 0)
    
    
    ' Calculate vectors across horizontal and down vertical viewport sides.
    Dim viewport_u As New vec3
    Dim viewport_v As New vec3
    Set viewport_u = vec3(CDbl(viewport_width), 0, 0)
    Set viewport_v = vec3(0, CDbl(-viewport_height), 0)
    
    
    ' Calculate the horizontal and vertical delta vectors from pix to pix.
    Dim pixel_delta_u As New vec3
    Dim pixel_delta_v As New vec3
    Set pixel_delta_u = viewport_u.div(CDbl(image_width))
    Set pixel_delta_v = viewport_v.div(CDbl(image_height))
    
    
    ' Calculate the point3 location of the upper left pix.
    ' Are you ready?
    Dim viewport_ul As New vec3
    Dim pix00_loc As New vec3
    Set viewport_ul = camera_center.subvec(vec3(0, 0, CDbl(focal_lenth)).subvec(viewport_u.div(2)).subvec(viewport_v.div(2)))
    Set pix00_loc = viewport_ul.addvec(pixel_delta_u.addvec(pixel_delta_v).mul(0.5))
    ' WHAT EVEN
    ' THIS IS WORSE THAN SMILEBASIC
    ' I DO NOT UNDERSTAND THIS ENOUGH TO UNTANGLE THIS
    
    
    ' Render!
    
    ' Variable setup.
    Dim pixelColor As New vec3
    Dim pix_center As New vec3
    Dim ray_dir As New vec3
    Dim r As New ray
    
    
    For j = 0 To image_height ' For every pixel
        For i = 0 To image_width
            ' black magic i can't comment on
            Set pix_center = pix00_loc.addvec(pixel_delta_u.mul(CDbl(i)).addvec(pixel_delta_v.mul(CDbl(j))))
            
            ' ray direction is the pixel's center minus the camera center, giving us dir
            Set ray_dir = pix_center.subvec(camera_center)
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
            Window.drawPixel i, j, pixelColor.convToRGB(False, Int(i), Int(j))
            
            ' Quality of life progressbar. Shows render progress.
            Window.renderProgress.Value = Window.renderProgress.Value + 1
        Next i
    Next j
    
    
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



Private Function vec3(x As Double, y As Double, z As Double) As vec3
    Dim temp As New vec3
    temp.x = x
    temp.y = y
    temp.z = z
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

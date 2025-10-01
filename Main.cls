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
Private Sub Class_Initialize()
    image_width = 512
    image_height = 512
    
    
    Window.renderProgress.Max = (image_width + 1) * (image_height + 1)
    Window.renderProgress.Value = 0
    
    Dim pixelColor As vec3
    Set pixelColor = New vec3
    
    ' loop through vertically
    For j = 0 To image_height
        ' loop through horizontally
        For i = 0 To image_width
            
            pixelColor.x = i / (image_width - 1)
            pixelColor.y = j / (image_height - 1)
            pixelColor.z = 0
            
            ' make it not freeze
            If i Mod 128 = 0 Then
                DoEvents
            End If
            
            ' make it not do slow
            'If i Mod 999999 = 0 Then
            '    Debug.Print "color r" & CStr(pixelColor.x) & " color g" & CStr(pixelColor.y) & " color b" & CStr(pixelColor.z)
            'End If
            
            Dim rgbcolor As Long
            rgbcolor = pixelColor.convToRGB(False, Int(i), Int(j))
            
            ' make it not do slow
            'If i Mod 999999 = 0 Then
            '    Debug.Print "pixelcolor long " & CStr(rgbcolor)
            'End If
            
            
            Window.drawPixel i, j, rgbcolor
            Window.renderProgress.Value = Window.renderProgress.Value + 1
        Next i
    Next j
    
    
    Window.renderProgress.Value = Window.renderProgress.Max
End Sub

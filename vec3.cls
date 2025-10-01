VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vec3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A vector with 3 values. X, Y, and Z."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"vec3"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarx As Variant 'local copy
Private mvary As Variant 'local copy
Private mvarz As Variant 'local copy
Public Property Let z(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.z = 5
    mvarz = vData
End Property


Public Property Set z(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.z = Form1
    Set mvarz = vData
End Property


Public Property Get z() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.z
    If IsObject(mvarz) Then
        Set z = mvarz
    Else
        z = mvarz
    End If
End Property



Public Property Let y(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.y = 5
    mvary = vData
End Property


Public Property Set y(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.y = Form1
    Set mvary = vData
End Property


Public Property Get y() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.y
    If IsObject(mvary) Then
        Set y = mvary
    Else
        y = mvary
    End If
End Property



Public Property Let x(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.x = 5
    mvarx = vData
End Property


Public Property Set x(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.x = Form1
    Set mvarx = vData
End Property


Public Property Get x() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.x
    If IsObject(mvarx) Then
        Set x = mvarx
    Else
        x = mvarx
    End If
End Property

Public Function convToRGB(Optional enableDither As Variant = False, Optional scrnX As Integer = 0, Optional scrnY As Integer = 0) As Long
    tempr = Me.x
    tempg = Me.y
    tempb = Me.z
    
    If enableDither Then
        ' Get placement in lookup in relation to screen
        Dim indexX As Integer
        Dim indexY As Integer
        indexX = scrnX Mod 4
        indexY = scrnY Mod 4
        
        dither = bayer(indexX, indexY)
        
        ' Magic numbers!
        threshold = (dither - 0.5) / 128 ' / 45
        
        ' apply dither!
        tempr = tempr + threshold
        tempg = tempg + threshold
        tempb = tempb + threshold
        
    End If
    
    tempr = 255.999 * tempr
    tempg = 255.999 * tempg
    tempb = 255.999 * tempb
    
    tempr = Min(tempr, 255)
    tempg = Min(tempg, 255)
    tempb = Min(tempb, 255)
    
    tempr = Max(tempr, 0)
    tempg = Max(tempg, 0)
    tempb = Max(tempb, 0)
    
    convToRGB = RGB(tempr, tempg, tempb)
End Function

' For dithering.
Private Function bayer(lookupX As Integer, lookupY As Integer) As Integer
    Dim bayerlookup(4, 4) As Integer
    
    bayerlookup(0, 0) = 0
    bayerlookup(0, 1) = 8
    bayerlookup(0, 2) = 2
    bayerlookup(0, 3) = 10
    
    bayerlookup(1, 0) = 12
    bayerlookup(1, 1) = 4
    bayerlookup(1, 2) = 14
    bayerlookup(1, 3) = 6
    
    bayerlookup(1, 3) = 6
    
    bayerlookup(2, 0) = 3
    bayerlookup(2, 1) = 11
    bayerlookup(2, 2) = 1
    bayerlookup(2, 3) = 9
    
    bayerlookup(3, 0) = 15
    bayerlookup(3, 1) = 7
    bayerlookup(3, 2) = 13
    bayerlookup(3, 3) = 5
        
    bayer = bayerlookup(lookupX, lookupY)
End Function

' return bigger
Private Function Max(val1, val2)
    If val1 > val2 Then
        Max = val1
    Else
        Max = val2
    End If
End Function

' return smaller
Private Function Min(val1, val2)
    If val1 < val2 Then
        Min = val1
    Else
        Min = val2
    End If
End Function

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
Private mvarx As Double 'local copy
Private mvary As Double 'local copy
Private mvarz As Double 'local copy
Public Property Let z(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.z = 5
    mvarz = vData
End Property

Public Property Get z() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.z
    z = mvarz
End Property



Public Property Let y(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.y = 5
    mvary = vData
End Property


Public Property Get y() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.y
        y = mvary
End Property



Public Property Let x(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.x = 5
    mvarx = vData
End Property


Public Property Get x() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.x
    x = mvarx
End Property



' When used as a color, this converts to RGB with optional dither.
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
        threshold = (dither - 0.5) / 45 '128
        
        ' apply dither!
        tempr = tempr + threshold
        tempg = tempg + threshold
        tempb = tempb + threshold
        
    End If
    
    tempr = 255 * tempr
    tempg = 255 * tempg
    tempb = 255 * tempb
    
    tempr = Min(tempr, 255)
    tempg = Min(tempg, 255)
    tempb = Min(tempb, 255)
    
    tempr = Max(tempr, 0)
    tempg = Max(tempg, 0)
    tempb = Max(tempb, 0)
    
    convToRGB = RGB(tempr, tempg, tempb)
End Function



' Bayer dithering lookup function.
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
        
    bayer = bayerlookup(lookupX, lookupY) / 16
End Function



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



' Add two vec3s together.
Public Function addvec(v As vec3) As vec3
    Dim temp As New vec3
    temp.x = Me.x + v.x
    temp.y = Me.y + v.y
    temp.z = Me.z + v.z
    Set addvec = temp
End Function



' Subtract two vec3s.
Public Function subvec(v As vec3) As vec3
    Dim temp As New vec3
    temp.x = Me.x - v.x
    temp.y = Me.y - v.y
    temp.z = Me.z - v.z
    Set subvec = temp
End Function



' Multiply a vec3 by a single number.
Public Function mul(t As Double) As vec3
    Dim temp As New vec3
    temp.x = Me.x * t
    temp.y = Me.y * t
    temp.z = Me.z * t
    Set mul = temp
End Function



' Divide a vec3 by a single number.
Public Function div(t As Double) As vec3
    Dim temp As New vec3
    Set temp = Me.mul(1# / t)
    Set div = temp
End Function



' Return the dot product of two vec3s.
Public Function dot(v As vec3) As Double
    dot = (Me.x * v.x) + (Me.y * v.y) + (Me.z * v.z)
End Function



' Length^2 is literally just the dot product of the vector
' and itself.
Public Function length_squared() As Double
    length_squared = Me.dot(Me)
End Function



' Length is the square root of length squared. It's not hard to
' understand.
Public Function length() As Double
    length = Sqr(Me.length_squared())
End Function



' Me Div() on Me Length()!
' Normalize. Theoretically the longest-running funct here.
' Also known as unit_vector().
' This places all the vec3s values between -1 to 1.
Public Function normalize() As vec3
    Set normalize = Me.div(Me.length())
End Function



' Cross?
Public Function crossvec(v As vec3) As vec3
    Dim temp As New vec3
    temp.x = Me.y * v.z - Me.z * v.y
    temp.y = Me.z * v.x - Me.x * v.z
    temp.z = Me.x * v.y - Me.y * v.x
    Set crossvec = temp
End Function



' Multiply two vec3s.
Public Function mulvec(v As vec3) As vec3
    Dim temp As New vec3
    temp.x = Me.x * v.x
    temp.y = Me.y * v.y
    temp.z = Me.z * v.z
    Set mulvec = temp
End Function


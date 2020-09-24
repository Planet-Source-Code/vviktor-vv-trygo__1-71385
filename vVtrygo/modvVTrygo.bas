Attribute VB_Name = "modvVTrygo"
'Some essential constants
Private Const PI As Double = 3.14159265358979
Private Const PIM2 As Double = 2 * PI
Private Const PID2 As Double = PI / 2
Private Const PI32 As Double = PI * (3 / 2)
Private Const PI180 As Double = PI / 180

'Function to reduce angle values (not used but maybe it'll be helpful for you so you can uncomment it):
'Public Function TrimRad(dblRad As Double) As Double
'    Do While dblRad < 0
'        dblRad = dblRad + PIM2
'    Loop
'    Do While dblRad >= PIM2
'        dblRad = dblRad - PIM2
'    Loop
'    TrimRad = dblRad
'End Function

'Conversion from radians to degress
Public Function Rad2Deg(dblRad As Double) As Single
    Rad2Deg = dblRad / PI180
End Function

'This function finds angle between to points (returned value is in radians)
Public Function AngleBetween(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Double
    Dim X As Single, Y As Single
    X = X2 - X1
    Y = Y2 - Y1
    
    If Y = 0 Then
        If X1 < X2 Then
            AngleBetween = 0
        Else
            AngleBetween = PI
        End If
    Else
        If Y < 0 Then
            AngleBetween = Atn(X / Y) + PID2
        Else
            AngleBetween = Atn(X / Y) + PI32
        End If
        
    End If

End Function

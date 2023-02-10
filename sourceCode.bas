Attribute VB_Name = "Module1"
Global turn As Integer
Dim red(11, 11) As Boolean
Dim blue(11, 11) As Boolean
Dim path(11, 11) As Boolean

Sub playerTurnDisplay()
    Dim turnDisplayBox As shape
    Set turnDisplayBox = ActiveSheet.Shapes("Hex 122")
    
    If turn Mod 2 = 1 Then
        turnDisplayBox.Fill.ForeColor.RGB = RGB(0, 0, 255)
        turnDisplayBox.TextFrame.Characters.Text = "BLUE'S TURN"
    Else
        turnDisplayBox.Fill.ForeColor.RGB = RGB(255, 0, 0)
        turnDisplayBox.TextFrame.Characters.Text = "RED'S TURN"
    End If
End Sub

Sub selectHex()
    Dim i As Integer
    Dim j As Integer
    
    Dim currentHex As shape
    Dim fullNameHexSplitted() As String
    Dim fullNameHex As String
    Dim numHexStr As String
    Dim numHexInt As Integer
    
    Set currentHex = ActiveSheet.Shapes(Application.Caller)
    
    If currentHex.Fill.ForeColor.RGB <> RGB(255, 0, 0) And currentHex.Fill.ForeColor.RGB <> RGB(0, 0, 255) Then
        If turn Mod 2 = 1 Then
            currentHex.Fill.ForeColor.RGB = RGB(0, 0, 255)
            turn = turn + 1
            Call playerTurnDisplay
            
            fullNameHex = currentHex.name
            fullNameHexSplitted = Split(fullNameHex)
            numHexStr = fullNameHexSplitted(1)
            numHexInt = CInt(numHexStr)
            blue((numHexInt - 1) \ 11, (numHexInt - 1) Mod 11) = True
            
            For i = 0 To 10
                If blue(i, 0) = True Then
                    Call checkBlue(i, 0)
                End If
            Next i
            
            For i = 0 To 10
                For j = 0 To 10
                    path(i, j) = False
                Next j
            Next i
        Else
            currentHex.Fill.ForeColor.RGB = RGB(255, 0, 0)
            turn = turn + 1
            Call playerTurnDisplay
            
            fullNameHex = currentHex.name
            fullNameHexSplitted = Split(fullNameHex)
            numHexStr = fullNameHexSplitted(1)
            numHexInt = CInt(numHexStr)
            red((numHexInt - 1) \ 11, (numHexInt - 1) Mod 11) = True
            
            For i = 0 To 10
                If red(0, i) = True Then
                    Call checkRed(0, i)
                End If
            Next i
            
            For i = 0 To 10
                For j = 0 To 10
                    path(i, j) = False
                Next j
            Next i
        End If
    End If
End Sub

Sub startGame()
    Dim shape As shape
    Dim i As Integer
    Dim j As Integer
    
    For Each shape In ActiveSheet.Shapes
        If shape.Type = msoAutoShape And shape.AutoShapeType = msoShapeHexagon Then
            shape.Fill.ForeColor.RGB = RGB(200, 200, 200)
        End If
    Next shape
    
    For i = 0 To 10
        For j = 0 To 10
            blue(i, j) = False
            red(i, j) = False
            path(i, j) = False
        Next j
    Next i
End Sub

Sub checkRed(x As Integer, y As Integer)
    If path(x, y) = True Then
        Exit Sub
    Else
        If x = 10 And red(x, y) = True Then
            MsgBox ("RED HAS WON!")
            Exit Sub
        End If
        
        path(x, y) = True
    End If
    
    If x - 1 < 11 And x - 1 >= 0 Then
        If red(x - 1, y) = True Then
            Call checkRed(x - 1, y)
        End If
    End If
    
    If x - 1 < 11 And x - 1 >= 0 And y + 1 < 11 And y + 1 >= 0 Then
        If red(x - 1, y + 1) = True Then
            Call checkRed(x - 1, y + 1)
        End If
    End If
    
    If y + 1 < 11 And y + 1 >= 0 Then
        If red(x, y + 1) = True Then
            Call checkRed(x, y + 1)
        End If
    End If
    
    If x + 1 < 11 And x + 1 >= 0 Then
        If red(x + 1, y) = True Then
            Call checkRed(x + 1, y)
        End If
    End If
    
    If x + 1 < 11 And x + 1 >= 0 And y - 1 < 11 And y - 1 >= 0 Then
        If red(x + 1, y - 1) = True Then
            Call checkRed(x + 1, y - 1)
        End If
    End If
    
    If y - 1 < 11 And y - 1 >= 0 Then
        If red(x, y - 1) = True Then
            Call checkRed(x, y - 1)
        End If
    End If
End Sub

Sub checkBlue(x As Integer, y As Integer)
    If path(x, y) = True Then
        Exit Sub
    Else
        If y = 10 And blue(x, y) = True Then
            MsgBox ("BLUE HAS WON!")
            Exit Sub
        End If
        
        path(x, y) = True
    End If
    
    If x - 1 < 11 And x - 1 >= 0 Then
        If blue(x - 1, y) = True Then
            Call checkBlue(x - 1, y)
        End If
    End If
    
    If x - 1 < 11 And x - 1 >= 0 And y + 1 < 11 And y + 1 >= 0 Then
        If blue(x - 1, y + 1) = True Then
            Call checkBlue(x - 1, y + 1)
        End If
    End If
    
    If y + 1 < 11 And y + 1 >= 0 Then
        If blue(x, y + 1) = True Then
            Call checkBlue(x, y + 1)
        End If
    End If
    
    If x + 1 < 11 And x + 1 >= 0 Then
        If blue(x + 1, y) = True Then
            Call checkBlue(x + 1, y)
        End If
    End If
    
    If x + 1 < 11 And x + 1 >= 0 And y - 1 < 11 And y - 1 >= 0 Then
        If blue(x + 1, y - 1) = True Then
            Call checkBlue(x + 1, y - 1)
        End If
    End If
    
    If y - 1 < 11 And y - 1 >= 0 Then
        If blue(x, y - 1) = True Then
            Call checkBlue(x, y - 1)
        End If
    End If
End Sub










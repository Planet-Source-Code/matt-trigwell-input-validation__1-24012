Attribute VB_Name = "modInputValidation"
Option Explicit

'Types of input
Public Enum InputType
    Date_Slash_Input = 0
    Date_Dash_Input = 1
    Numeric_Input = 2
    Text_Input = 3
    Currency_Input = 4
End Enum

Public Function ValidateInput(KeyAscii As Integer, Format As InputType, Optional Uppercase As Boolean) As Integer
'Description: - Validates the
'Parameters:  - KeyAscii - The ANSI Keycode pressed
'             - Uppercase - Display text in uppercase if true
'Returnvalue: - KeyAscii or 0 if not allowed
'Programmer : - Matt Trigwell 11/06/01
'Comments:    - Auzzie Auzzie Auzzie oi oi oi!


If (Format = Date_Slash_Input) Then
'dd/mm/yy

'Keycodes
'48 = 0
'57 = 9
'8 = BackSpace
'47 = /
'0 = Cancel user input

    If (KeyAscii > 57 Or KeyAscii < 48) Then
        If (KeyAscii <> 8) Then
            If (KeyAscii <> 47) Then
                KeyAscii = 0
            End If
        End If
    End If

ElseIf (Format = Date_Dash_Input) Then
'dd-mm-yy

'Keycodes
'48 = 0
'57 = 9
'8 = BackSpace
'45 = -
'0 = Cancel user input

    If (KeyAscii > 57 Or KeyAscii < 48) Then
        If (KeyAscii <> 8) Then
            If (KeyAscii <> 45) Then
                KeyAscii = 0
            End If
        End If
    End If
    
ElseIf (Format = Numeric_Input) Then
'0-9

'Keycodes
'48 = 0
'57 = 9
'8 = BackSpace
'0 = Cancel user input
If (KeyAscii < 48 Or KeyAscii > 57) Then
    If (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End If
   
    
ElseIf (Format = Text_Input) Then
'A-Z a-z

'65 = A
'122 = Z
'32 = Space
'8 = BackSpace

If (KeyAscii >= 65 And KeyAscii <= 122 Or (KeyAscii = 32 Or KeyAscii = 8)) Then
    'Change to uppercase
    If (Uppercase = True) Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
Else
    KeyAscii = 0
End If


ElseIf (Format = Currency_Input) Then
'$0,000.00

'Keycodes
'48 = 0
'57 = 9
'8 = BackSpace
'36 = $
'44 = ,
'46 = .
'0 = Cancel user input

    If (KeyAscii > 57 Or KeyAscii < 48) Then
        If (KeyAscii <> 8 And KeyAscii <> 36 And KeyAscii <> 44 And KeyAscii <> 46) Then
            KeyAscii = 0
        End If
    End If
End If


'Return KeyAscii or 0 if value is not allowed
ValidateInput = KeyAscii

End Function

Attribute VB_Name = "udf_modulus_zero"
Option Explicit

Private Sub ModulusZero_Test()
    Debug.Print "ModulusZero(3.001, 3) -> " & ModulusZero(3.001, 3)
    Debug.Print "ModulusZero(2, 0.3) -> " & ModulusZero(2, 0.3)
    Debug.Print "ModulusZero(3, 0.3) -> " & ModulusZero(3, 0.3)
    Debug.Print "ModulusZero(6.3, 0.3) -> " & ModulusZero(6.3, 0.3)
    Debug.Print "ModulusZero(57, 0.4) -> " & ModulusZero(57, 0.4)
    Debug.Print "ModulusZero(6324692, 415) -> " & ModulusZero(6324692, 415)
    Debug.Print "ModulusZero(6324600, 415) -> " & ModulusZero(6324600, 415)
    Debug.Print "ModulusZero(6324600.99, 415) -> " & ModulusZero(6324600.99, 415)
End Sub

Public Function ModulusZero(ByVal dblInputValue As Single, ByVal dblDivisor As Single) As Boolean
    'This procedure checks to see if any integer or decimal is equally divisible by the
    'provided divisor. Note that the precision of the input value is evaluated at the same level
    'as the precision of the divisor.
    If dblInputValue = dblDivisor Then
        'If the numbers are the same, then they are equally divisble.
        ModulusZero = True
    ElseIf dblInputValue < dblDivisor Then
        'If the input value is smaller than the divisor, then it cannot be equally divisible.
        ModulusZero = False
    Else
        'Microsoft warns about comparisons of doubles and singles and that they may not always
        'evaluate how you exptect them to. Therefore, I treat their results of the maths as
        'strings.
        If InStr(CStr(dblInputValue), ".") = 0 And InStr(CStr(dblDivisor), ".") = 0 Then
            ModulusZero = IIf(InStr(CStr(CLng(dblInputValue) / CLng(dblDivisor)), ".") > 0, False, True)
        Else
            ModulusZero = IIf(InStr(dblInputValue / dblDivisor, ".") > 0, False, True)
        End If
    End If
End Function



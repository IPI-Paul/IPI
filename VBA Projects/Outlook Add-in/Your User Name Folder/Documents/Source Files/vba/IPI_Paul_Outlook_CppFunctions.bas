Attribute VB_Name = "IPI_Paul_Outlook_CppFunctions"
Private Declare PtrSafe Function multiplyBy Lib "C:\Users\Paul\Documents\Source Files\dll\multiplyBy x64.dll" (ByRef X As Double, ByRef Y As Double) As Double

Public Function cppMultiplyBy(ByRef X As Double, ByRef Y As Double)
    Dim result
    multiplyBy X, Y
    cppMultiplyBy = Y
End Function

Function normVal(val)
    For Each itm In Array(",", "'")
        val = Replace(val, itm, "")
    Next
    normVal = val
End Function

Sub retCppMultiplyBy()
    Dim X As Double, Y As Double
0:
    vals = InputBox("Please enter the two values to multiply separated by a space", "C++ multiplyBy Linked Function", "2.1 2,324.41")
    If vals > "" Then
        vals = Split(normVal(vals), " ")
        X = vals(0)
        Y = vals(1)
        If MsgBox(X & " * " & vals(1) & " = " & cppMultiplyBy(X, Y) & vbCrLf & vbCrLf & "Do you want to calculate another?", vbYesNo, "C++ multiplyBy Linked Function Result") = vbYes Then GoTo 0
    End If
End Sub

Sub viewForm_CppMultiplyBy()
    IPI_Paul_Outlook_CppMultiPlyBy.Show 0
End Sub

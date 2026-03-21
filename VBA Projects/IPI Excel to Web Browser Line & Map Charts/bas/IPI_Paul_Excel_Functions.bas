Attribute VB_Name = "IPI_Paul_Excel_Functions"
Option Explicit

Function allIsInArray(arr, iStr) As Boolean
    Dim aStr As Variant
    
    For Each aStr In arr
        If iStr = aStr Then
            allIsInArray = True
            Exit Function
        End If
    Next
End Function

Function allIsInFields(flds, arr) As Boolean
    Dim fld As Variant, aStr As Variant
    
    For Each fld In flds
        For Each aStr In arr
            If aStr = fld.name Then
                allIsInFields = True
                Exit Function
            End If
        Next
    Next
End Function

Function arrayIsInString(iStr, arr)
    Dim aStr As Variant
    
    For Each aStr In arr
        If InStr(iStr, aStr) Then
            arrayIsInString = True
            Exit Function
        End If
    Next
End Function

Function getClipboard()
    Dim objData As New MSForms.DataObject
    
    getClipboard = ""
    On Error Resume Next
    objData.GetFromClipboard
    getClipboard = objData.GetText()
End Function

Function isInArray(arr, iStr) As Boolean
    Dim aStr As Variant, itm As Variant
    
    For Each aStr In arr
        If iStr = aStr Then
            isInArray = True
            Exit Function
        End If
    Next
    For Each itm In Split(iStr, " ")
        For Each aStr In arr
            If itm = aStr Then
                isInArray = True
                Exit Function
            End If
        Next
    Next
End Function

Function match(arr As Variant, iStr As String) As Integer
    Dim i As Integer, itm As Variant
    
    i = 0
    For Each itm In arr
        If itm = iStr Then Exit For
        i = i + 1
    Next
    match = i
End Function

Function max(ParamArray arr() As Variant)
    Dim mn As Long, itm As Variant
    
    For Each itm In arr
        If min < itm Then min = itm
    Next
End Function

Function min(ParamArray arr() As Variant)
    Dim mn As Long, itm As Variant
    
    For Each itm In arr
        If min = 0 Or min > itm Then min = itm
    Next
End Function

Function nz(obj As Variant, Optional tp As Variant = "")
    On Error Resume Next
    
    nz = tp
    If Not IsNull(obj) Then
        nz = obj
    End If
End Function

Function waitTill(Optional dur As String = "00:00:01") As Boolean
    Dim tNow As Date
    
    tNow = Now() + TimeValue(dur)
    While Now() < tNow
        DoEvents
    Wend
End Function

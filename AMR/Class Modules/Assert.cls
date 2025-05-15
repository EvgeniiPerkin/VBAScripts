VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private IsError As Boolean
Private name_method As String
Property Get NameMethod() As String
    NameMethod = name_method
End Property
Public Property Let NameMethod(m_name As String)
    name_method = m_name
End Property
'---------------------------------------------------------
'                   A class for testing
'---------------------------------------------------------

'---------------------------------------------------------
' Usage example:
'
' Public Sub TestCalculate()
'   Dim Assert As New Assert
'   Assert.NameMethod = "TestCalculate"
'   Call Assert.Equal_Long(100, Calculate(10, 10))
'   Call Assert.Equal_Long(101, Calculate(10, 10))
'   Call Assert.ResultAssert
' End Sub
'
' Public Function Calculate(value1 As Long, value2 As Long) As Long
'   Calculate = value1 * value2
' End Function
'---------------------------------------------------------

Public Sub Equal_Date(expected As Date, actual As Date)
    If expected <> actual Then
        Debug.Print name_method & " error expected: " & expected & ", actual: " & actual
        IsError = True
    End If
End Sub
Public Sub Equal_String(expected As String, actual As String)
    If expected <> actual Then
        Debug.Print name_method & " error expected: " & expected & ", actual: " & actual
        IsError = True
    End If
End Sub
Public Sub Equal_Long(expected As Long, actual As Long)
    If expected <> actual Then
        Debug.Print name_method & " error expected: " & expected & ", actual: " & actual
        IsError = True
    End If
End Sub
Public Sub Equal_Boolean(expected As Boolean, actual As Boolean)
    If expected <> actual Then
        Debug.Print name_method & " error expected: " & expected & ", actual: " & actual
        IsError = True
    End If
End Sub
Public Sub Equal_Integer(expected As Integer, actual As Integer)
    If expected <> actual Then
        Debug.Print name_method & " error expected: " & expected & ", actual: " & actual
        IsError = True
    End If
End Sub
Public Sub Equal_Double(expected As Double, actual As Double, delta As Long)
    Dim expected_r As Double, actual_r As Double
    expected_r = Round(expected, delta)
    actual_r = Round(actual, delta)
    If expected_r <> actual_r Then
        Debug.Print name_method & " error expected: " & expected & ", actual: " & actual
        IsError = True
    End If
End Sub

Public Sub ResultAssert()
    If IsError Then
        MsgBox "Testing completed with errors!" & vbNewLine & "Open the Debug window", vbCritical, "Testing"
    Else
        MsgBox "Testing has been completed successfully!", vbInformation, "Testing"
    End If
End Sub

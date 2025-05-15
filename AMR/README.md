# AMR

## Class Modules\Assert.cls

The class contains a set of procedures and functions for testing VBA code.

Example usage:

```vbnet
Public Sub TestCalculate()
  Dim Assert As New Assert
  Assert.NameMethod = "TestCalculate"
  Call Assert.Equal_Long(100, Calculate(10, 10))
  Call Assert.Equal_Long(101, Calculate(10, 10))
  Call Assert.ResultAssert
End Sub

Public Function Calculate(value1 As Long, value2 As Long) As Long
  Calculate = value1 * value2
End Function
```

## Modules\ModAMR.bas

A module for working with the data of uploads (books) of arithmetic mean balances (CFT).
The main procedure for launching is BreakBookIntoParts

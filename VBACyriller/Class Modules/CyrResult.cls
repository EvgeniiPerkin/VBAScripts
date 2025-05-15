VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CyrResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************Description*******************************
'*********************************************************

'***************Fields'***********************************
Private Case1 As String
Private case2 As String
Private Case3 As String
Private Case4 As String
Private Case5 As String
Private Case6 As String
'*********************************************************

'***************Properties********************************
Property Get Case_1() As String
    Case_1 = Case1
End Property
Public Property Let Case_1(p_Case1 As String)
    Case1 = p_Case1
End Property

Property Get Case_2() As String
    Case_2 = case2
End Property
Public Property Let Case_2(p_Case2 As String)
    case2 = p_Case2
End Property

Property Get Case_3() As String
    Case_3 = Case3
End Property
Public Property Let Case_3(p_Case3 As String)
    Case3 = p_Case3
End Property

Property Get Case_4() As String
    Case_4 = Case4
End Property
Public Property Let Case_4(p_Case4 As String)
    Case4 = p_Case4
End Property

Property Get Case_5() As String
    Case_5 = Case5
End Property
Public Property Let Case_5(p_Case5 As String)
    Case5 = p_Case5
End Property

Property Get Case_6() As String
    Case_6 = Case6
End Property
Public Property Let Case_6(p_Case6 As String)
    Case6 = p_Case6
End Property
'*********************************************************

'***************Methods***********************************
Public Sub Constructor1(Txt As String)
    Case1 = Txt
    case2 = Txt
    Case3 = Txt
    Case4 = Txt
    Case5 = Txt
    Case6 = Txt
End Sub
Public Sub Constructor2(p_Case1 As String, p_Case2 As String, p_Case3 As String, p_Case4 As String, p_Case5 As String, p_Case6 As String)
    Case1 = p_Case1
    case2 = p_Case2
    Case3 = p_Case3
    Case4 = p_Case4
    Case5 = p_Case5
    Case6 = p_Case6
End Sub
Public Sub Constructor3(myCasesArr() As String)
    If UBound(myCasesArr) <= 6 Then
        Case1 = myCasesArr(0)
        case2 = myCasesArr(1)
        Case3 = myCasesArr(2)
        Case4 = myCasesArr(3)
        Case5 = myCasesArr(4)
        Case6 = myCasesArr(5)
    End If
End Sub

' Именительный, Кто? Что?   (есть)
Public Function Nominative() As String
    Nominative = Case1
End Function
' Родительный, Кого? Чего?  (нет)
Public Function Genitive() As String
    Genitive = case2
End Function
' Дательный, Кому? Чего?    (дам)
Public Function Dative() As String
    Dative = Case3
End Function
' Винительный, Кого? Что?   (вижу)
Public Function Accusative() As String
    Accusative = Case4
End Function
' Творительный, Кем? Чем?   (горжусь)
Public Function Instrumental() As String
    Instrumental = Case5
End Function
' Предложный, О ком? О Чем? (думаю)
Public Function Prepositional() As String
    Prepositional = Case6
End Function

' Именительный, Кто? Что?   (есть)
Public Function Именительный() As String
    Именительный = Case1
End Function
' Родительный, Кого? Чего?  (нет)
Public Function Родительный() As String
    Родительный = case2
End Function
' Дательный, Кому? Чего?    (дам)
Public Function Дательный() As String
    Дательный = Case3
End Function
' Винительный, Кого? Что?   (вижу)
Public Function Винительный() As String
    Винительный = Case4
End Function
' Творительный, Кем? Чем?   (горжусь)
Public Function Творительный() As String
    Творительный = Case5
End Function
' Предложный, О ком? О Чем? (думаю)
Public Function Предложный() As String
    Предложный = Case6
End Function

Public Function GetCase(CasesParam As CasesEnum) As String
    Dim Result As String
    
    Select Case CasesParam
        Case 1
            Result = Case1
        Case 2
            Result = case2
        Case 3
            Result = Case3
        Case 4
            Result = Case4
        Case 5
            Result = Case5
        Case 6
            Result = Case6
        Case Else
            Result = Case1
    End Select
    
    GetCase = Result
End Function

Public Function SetCase(CasesParam As CasesEnum, Value As String)
    Select Case CasesParam
        Case 1
            Case1 = Value
        Case 2
            case2 = Value
        Case 3
            Case3 = Value
        Case 4
            Case4 = Value
        Case 5
            Case5 = Value
        Case 6
            Case6 = Value
        Case Else
            Case1 = Value
    End Select
End Function

Public Sub add(Result As CyrResult, Optional separator As String = "-")
    Case1 = Case1 + separator + Result.Case_1
    case2 = case2 + separator + Result.Case_2
    Case3 = Case3 + separator + Result.Case_3
    Case4 = Case4 + separator + Result.Case_4
    Case5 = Case5 + separator + Result.Case_5
    Case6 = Case6 + separator + Result.Case_6
End Sub

Public Function ToList() As String()
    Dim Result() As String
    ReDim Result(5)
    Result(0) = Case1
    Result(1) = case2
    Result(2) = Case3
    Result(3) = Case4
    Result(4) = Case5
    Result(5) = Case6
    ToList = Result
End Function

Public Function ToArray() As String()
    Dim Result() As String
    ReDim Result(5)
    Result(0) = Case1
    Result(1) = case2
    Result(2) = Case3
    Result(3) = Case4
    Result(4) = Case5
    Result(5) = Case6
    ToArray = Result
End Function

Public Function ToDictionary() As CasesEnumAndString()
    Dim Result() As CasesEnumAndString
    
    ReDim Result(5)
    Result(0).Key = Nominative
    Result(0).Value = Case1
    
    Result(1).Key = Genitive
    Result(1).Value = case2
    
    Result(2).Key = Dative
    Result(2).Value = Case3
    
    Result(3).Key = Accusative
    Result(3).Value = Case4
    
    Result(4).Key = Instrumental
    Result(4).Value = Case5
    
    Result(5).Key = Prepositional
    Result(5).Value = Case6
    
    ToDictionary = Result
End Function

Public Function ToStringDictionary() As StringAndString()
    Dim Result() As StringAndString
    
    ReDim Result(5)
    Result(0).Key = "Nominative"
    Result(0).Value = Case1
    
    Result(1).Key = "Genitive"
    Result(1).Value = case2
    
    Result(2).Key = "Dative"
    Result(2).Value = Case3
    
    Result(3).Key = "Accusative"
    Result(3).Value = Case4
    
    Result(4).Key = "Instrumental"
    Result(4).Value = Case5
    
    Result(5).Key = "Prepositional"
    Result(5).Value = Case6
    
    ToStringDictionary = Result
End Function

Public Function ToRussianStringDictionary() As StringAndString()
    Dim Result() As StringAndString
    
    ReDim Result(5)
    Result(0).Key = "Именительный"
    Result(0).Value = Case1
    
    Result(1).Key = "Родительный"
    Result(1).Value = case2
    
    Result(2).Key = "Дательный"
    Result(2).Value = Case3
    
    Result(3).Key = "Винительный"
    Result(3).Value = Case4
    
    Result(4).Key = "Творительный"
    Result(4).Value = Case5
    
    Result(5).Key = "Предложный"
    Result(5).Value = Case6
    
    ToRussianStringDictionary = Result
End Function
'*********************************************************
